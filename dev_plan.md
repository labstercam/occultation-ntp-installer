# Occultation Workflow Integration Plan

## Overview

This document captures the development plan for integrating NTP timing analysis, camera
acquisition delay calculation, and the full post-recording workflow into occultation-manager.

The full end-to-end workflow the plan targets:

1. Select events in Occultation Manager
2. Record using Occultation Manager sequences in SharpCap
3. Analyse NTP loopstats/peerstats offsets to determine the timing correction for the event window
4. Estimate the camera acquisition delay for the occulted star's Y line using prior LED line delay calibration
5. Enter the corrections into TANGRA and run the light curve analysis
6. Analyse the TANGRA light curve in AOTA and generate AOTA outputs
7. Generate the report in Occultation Manager

Steps 1, 2, and 7 are already implemented. Steps 3 and 4 are currently standalone tools with no
integration into the event workflow. This plan fills that gap.

---

## Background: What Already Exists

### Occultation Manager
- Event selection, sequence generation, and recording (Steps 1–2) are complete.
- Report generation (Step 7) reads AOTA XML/text and Tangra CSV and populates NA/TT/SODIS reports.
- `light_curves_iron.py` (IronPython-compatible Tangra CSV reader) is already in the
  `occultation-manager/python/` folder.
- Report generation reads `acquisition_delay` from the Tangra CSV — this creates a circular
  dependency (Tangra must be run before the delay is known, but the delay must be entered before
  running Tangra). The plan resolves this.
- `led_line_delay_calibration.py` will be **moved here** from `gps-timing-analysis` (see Part 2).

### gps-timing-analysis
- `python/ntp_analysis.py` — NTP loopstats/peerstats offset analysis (CPython).
- `scripts/analyze_ntp_timing_accuracy.py` — IronPython Windows Forms GUI wrapping `ntp_analysis.py`.
- `python/led_line_delay_calibration.py` — Rolling shutter line delay calibration via GPS LED.
  Currently saves only a Tangra CSV and displays results in a TextBox. Does **not** persist
  calibration coefficients to a structured file.
  **Decision: this module will be moved into occultation-manager** (see Part 2). Its only consumer
  is the acquisition delay workflow in OM, and keeping it in a separate SharpCap add-in would
  require hardcoded cross-add-in import paths. The NTP analysis scripts remain in gps-timing-analysis.

### Acquisition Delay Formula

The acquisition delay for a given event is calculated as:

```
acquisition_delay_ms = intercept + slope × Y_line
```

where `intercept` and `slope` are the coefficients from a prior LED calibration for the specific
camera, resolution, and binning combination, and `Y_line` is the pixel row of the target star
in the recorded frame (identified by the user in Tangra or SharpCap before analysis).

---

## Part 1: Calibration Storage

### Calibration JSON Format

Add a **Save Calibration** button to the `led_line_delay_calibration.py` results panel. On click,
write a JSON file with this structure:

```json
{
  "schema_version": 1,
  "calibrated_utc": "2026-01-15T09:32:00",
  "camera_name": "ASI174MM",
  "frame_width_px": 1920,
  "frame_height_px": 1216,
  "binning": 1,
  "slope_ms_per_px": 0.00812,
  "intercept_ms": -4.93,
  "r_squared": 0.9981,
  "n_measurements": 47,
  "notes": ""
}
```

The calibration is keyed by `camera_name` + `frame_width_px × frame_height_px` + `binning`. The
filename encodes these fields so the consumer can match without opening every file:

```
ASI174MM_1920x1216_bin1_20260115.json
ASI174MM_960x608_bin2_20260115.json
QHY174GPS_1920x1216_bin1_20260210.json
```

### Calibration Storage Location

Store calibrations in a **user-documents folder** outside the add-in installation:

```
Documents\occultation-tools\calibrations\line_delay\
```

Rationale:
- Both the calibration tool and the acquisition delay calculator now live inside occultation-manager,
  so there is no cross-add-in coupling problem. The user-documents path is still preferred so that
  calibrations survive an OM reinstall or upgrade.
- Users can maintain multiple calibrations per camera (different dates, settings) and choose which
  to apply.
- The path is configurable in Occultation Manager via Tools → Configuration → File Paths.

After saving, the GUI should display the equation prominently so the user can verify it:

```
Calibration saved.
Equation: delay = -4.930 + 0.00281 × Y  ms   (R² = 0.998)
```

### Changes Required in `led_line_delay_calibration.py`
- **Move file** from `gps-timing-analysis/python/` to `occultation-manager/python/`.
- Add **Save Calibration** button to `display_results()`.
- Default save folder: `Documents\occultation-tools\calibrations\line_delay\`, created on first use.
- Show the save path clearly after writing.

---

## Part 2: LED Line Delay Tool Placement in Occultation Manager

Calibration is a **one-time setup per camera/settings combination** and is not part of the
per-event workflow. It therefore belongs in the Tools menu, not in the main workflow panel.

**Placement:** Tools → Camera Calibration → LED Line Delay Calibration

`led_line_delay_calibration.py` is moved from `gps-timing-analysis` into
`occultation-manager/python/`. It is invoked as a module within OM — no cross-add-in import or
external launcher is needed. The existing IronPython Windows Forms UI is retained unchanged.

---

## Part 3: New Post-Recording Analysis Panel

A new panel is added between Observation Preparation and Report Generation:

```
[Events]  →  [Prepare]  →  [Post-Recording Analysis]  →  [Generate Report]
```

This panel contains three sub-sections.

### Sub-section A: NTP Offset Analysis

| Element | Detail |
|---|---|
| Input | Event selected in event grid (provides UTC event time) |
| Config | Meinberg loopstats/peerstats folder path (Tools → Configuration → File Paths) |
| Window | Configurable margin around event time (default ±30 min) |
| Logic | Calls existing `ntp_analysis.py` / `analyze_ntp_timing_accuracy.py` logic |
| Output | NTP offset correction in ms with uncertainty (e.g. `+2.3 ms ± 1.1 ms`) |

No new analysis code is needed — the logic already exists in `gps-timing-analysis`. The integration
work is wiring the event time window into the existing functions and surfacing the result in the UI.

### Sub-section B: Acquisition Delay Calculator

| Element | Detail |
|---|---|
| Input — camera | Automatically populated from active equipment profile |
| Input — Y line | User enters the pixel row of the target star (found in Tangra or SharpCap) |
| Calibration lookup | Scans `calibrations\line_delay\` for JSON matching camera + frame dims + binning |
| Multiple matches | Dropdown of available calibration dates; user picks |
| No match | Warning: "No calibration found for {camera} {W×H} bin{N} — run Tools → Camera Calibration" |
| Calculation | `delay = intercept + slope × Y_line` |
| Output | Acquisition delay in ms (e.g. `+8.7 ms`) |

### Sub-section C: TANGRA Setup Summary

A read-only display consolidating all three values the user must enter into TANGRA before running
the light curve:

```
── Values to configure in TANGRA ─────────────────────────────
  NTP offset correction:      +2.3 ms  ± 1.1 ms
  Acquisition delay (Y=480):  +8.7 ms
    (intercept = -4.93 ms, slope = 0.00281 ms/px, R² = 0.998)
──────────────────────────────────────────────────────────────
  [Copy to clipboard]   [Launch TANGRA]
```

**[Copy to clipboard]** copies a plain-text version for pasting into notes.  
**[Launch TANGRA]** opens TANGRA from a configurable path in Tools → Configuration.

The `acquisition_delay` field in the existing Report Generation dialog is pre-populated from this
panel's result, removing the previous circular dependency on parsing it from the Tangra CSV after
TANGRA has already been run.

---

## Part 4: Equipment Profile Changes

To enable automatic calibration lookup (without requiring the user to re-enter camera settings),
add the following fields to each camera entry in `occultation_config.json`:

```json
{
  "camera_name": "ASI174MM",
  "frame_width_px": 1920,
  "frame_height_px": 1216,
  "binning": 1
}
```

These values populate automatically from the active camera when a calibration is saved, and are
used as the lookup key when calculating acquisition delay per event.

---

## Part 5: Configuration Changes

Add to Tools → Configuration → File Paths:

| Setting | Default | Purpose |
|---|---|---|
| `calibrations_folder` | `Documents\occultation-tools\calibrations\line_delay\` | Where calibration JSONs are read and written |
| `loopstats_folder` | Meinberg default logs path | Where NTP loopstats/peerstats are read from |
| `tangra_exe_path` | (blank) | Path to Tangra executable for the Launch button |

---

## Summary of Changes by Repo

| Repo | File / Area | Change |
|---|---|---|
| `gps-timing-analysis` | `led_line_delay_calibration.py` | **Remove** — file moved to occultation-manager |
| `occultation-manager` | `python/led_line_delay_calibration.py` | **Add** (moved from gps-timing-analysis); add Save Calibration button; write JSON to calibrations folder |
| `occultation-manager` | `occultation_config.json` schema | Add `calibrations_folder`, `loopstats_folder`, `tangra_exe_path` to config |
| `occultation-manager` | Tools → Configuration → File Paths | Expose new paths in settings UI |
| `occultation-manager` | Equipment profile | Add `frame_width_px`, `frame_height_px`, `binning` per camera |
| `occultation-manager` | New panel | Post-Recording Analysis panel (NTP sub-section A, acquisition delay sub-section B, TANGRA summary sub-section C) |
| `occultation-manager` | Report Generation dialog | Pre-populate `acquisition_delay` from Post-Recording panel; keep CSV fallback if panel result not available |
| `occultation-manager` | Tools menu | Add Tools → Camera Calibration → LED Line Delay Calibration (module within OM, no external launcher) |
