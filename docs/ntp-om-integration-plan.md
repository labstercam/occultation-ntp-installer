# NTP Analyser Integration into Occultation Manager

## Overview

Split `analyze_ntp_timing_accuracy.py` into a pure computation module (`ntp_analysis_core.py`)
and the existing GUI shell. Occultation Manager imports `ntp_analysis_core` via the monorepo
relative path, calls three functions to get offset + uncertainty at event time, and opens the
full interactive GUI by instantiating `AnalyzerForm().Show()` in the same IronPython process.

---

## Phase 1 — Extract `ntp_analysis_core.py` (this repo)

The script has a clean seam at line ~1699 where `class AnalyzerForm` begins. Everything before
it is pure computation with no Windows Forms dependency.

**Steps:**

1. Create `scripts/ntp_analysis_core.py` — move everything before `class AnalyzerForm` into it:
   all stdlib imports, constants (`MJD_EPOCH`, `_FIBRE_SPEED_KMS`), data classes
   (`LoopRecord`, `PeerRecord`, `DayOption`, `AnalysisResult`), and all computation functions
   (`parse_loopstats`, `parse_peerstats`, `estimate_offset_at_time`, `resolve_server_location`,
   `load_known_servers`, `_load_ip_location_cache`, etc.).
   **No `clr` / Windows Forms imports in this file.**

2. Reduce `analyze_ntp_timing_accuracy.py` to:
   ```python
   sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
   from ntp_analysis_core import *
   # ... class AnalyzerForm, main(), if __name__
   ```
   Same external behaviour; ~1 500 fewer lines in the GUI file.

---

## Phase 2 — OM Integration (Post-Recording Analysis panel, Sub-section A)

Monorepo layout:
```
occultation-tools/
  gps-timing-analysis/
    scripts/ntp_analysis_core.py        ← source of truth (never copied)
    resources/national_utc_ntp_servers.json
    resources/ip_location_cache.json    ← shared 90-day cache
  occultation-manager/
    python/<post_recording_panel>.py    ← integration point
```

**Steps:**

3. At OM module load, wire paths and import:
   ```python
   import os, sys
   _NTP_SCRIPTS   = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "..", "gps-timing-analysis", "scripts"))
   _NTP_RESOURCES = os.path.normpath(os.path.join(os.path.dirname(__file__), "..", "..", "gps-timing-analysis", "resources"))
   sys.path.insert(0, _NTP_SCRIPTS)
   import ntp_analysis_core as ntp
   ```

4. On panel init (no network — fast local file reads):
   ```python
   self._known_servers = ntp.load_known_servers(os.path.join(_NTP_RESOURCES, "national_utc_ntp_servers.json"))
   ntp._load_ip_location_cache(os.path.join(_NTP_RESOURCES, "ip_location_cache.json"))
   ```

5. **"Analyse NTP" button** — triggered when user opens Post-Recording Analysis with an event selected:
   ```python
   from datetime import date
   mjd = (event_date - date(1858, 11, 17)).days
   sec = event_hour * 3600.0 + event_minute * 60.0 + event_second

   loop_rows = ntp.parse_loopstats(loopstats_path, mjd)
   peer_rows  = ntp.parse_peerstats(peerstats_path, mjd)
   result = ntp.estimate_offset_at_time(
       mjd, sec, loop_rows, peer_rows,
       known_servers=self._known_servers,
       observer_lat=observer_lat,
       observer_lon=observer_lon,
   )
   ```

   Display in panel:

   | Result field | Multiply by | Display |
   |---|---|---|
   | `result['best_offset']` | × 1000 | `+2.3 ms` offset |
   | `result['u_expanded']`  | × 1000 | `± 1.1 ms` (95%) |
   | `result['gap_before_s']` | — | `data age: 4 min before event` |
   | `result['active_server_at_T']` | — | server info line |

6. **"Open NTP Analyser" button** — opens the full interactive tool as a peer window,
   non-blocking, in the same SharpCap IronPython process:
   ```python
   import analyze_ntp_timing_accuracy as ntp_gui
   ntp_gui.AnalyzerForm().Show()   # .Show() not Application.Run() — non-blocking
   ```

---

## Phase 3 — Loopstats Path Configuration

7. Add `loopstats_folder` to OM's config (Tools → Configuration → File Paths).
   Default: Meinberg standard log path (`C:\Program Files (x86)\NTP\etc\` or similar).
   Sub-section A pre-fills from config; user can override per session.

---

## Relevant Files

| File | Location | Action |
|---|---|---|
| `analyze_ntp_timing_accuracy.py` | `gps-timing-analysis/scripts/` | Reduce to GUI shell only (split at line ~1699) |
| `ntp_analysis_core.py` | `gps-timing-analysis/scripts/` | **New** — extracted computation module |
| `national_utc_ntp_servers.json` | `gps-timing-analysis/resources/` | No change — path passed explicitly |
| `ip_location_cache.json` | `gps-timing-analysis/resources/` | No change — shared 90-day cache |
| `<post_recording_panel>.py` | `occultation-manager/python/` | **New** (OM repo) — Sub-section A wiring |

---

## Key Result Fields from `estimate_offset_at_time`

```
result['best_offset']          float (s)   NTP clock offset at event time  → × 1000 for ms
result['u_expanded']           float (s)   expanded ~95% uncertainty        → × 1000 for ms
result['u_combined']           float (s)   combined k=1 standard uncertainty
result['gap_before_s']         float (s)   seconds between event and last loopstats record
result['active_server_at_T']   str|None    NTP server active at event time
result['server_location_note'] str         how the server location was resolved
result['note']                 str         full human-readable summary
```

---

## Decisions

- `ntp_analysis_core.py` is **not copied** to OM — imported directly via monorepo relative path.
  Single source of truth; updates to the core benefit both tools automatically.
- `ip_location_cache.json` lives in `gps-timing-analysis/resources/` so both the standalone tool
  and OM share the same 90-day IP → location cache.
- GUI open uses `.Show()` not `Application.Run()` — non-blocking, peer window, same process.
- `ntp_analysis_core.py` has **no Windows Forms imports** — safe to import in any IronPython context.
- Observer lat/lon for geographic b_asym tightening comes from OM's existing equipment/observer
  config, not re-entered in the NTP panel.
