# LED Line Delay Calibration — Repo Move Checklist

Move `led_line_delay_calibration.py` and its dependencies from `gps-timing-analysis` into
`occultation-manager` within the `labstercam/occultation-tools` monorepo.

---

## 1. Copy Files to occultation-manager

### Python modules
`gps-timing-analysis/python/` → `occultation-manager/python/`

- [x] Copy `led_line_delay_calibration.py`
- [x] Copy `adv_processing_iron.py`
- [x] Copy `adv_helper.py`

### Documentation
`gps-timing-analysis/` → `occultation-manager/`

- [x] Copy `LED_LINE_DELAY_QUICKSTART.md`
- [x] Copy `LED_LINE_DELAY_DEVELOPMENT_PLAN.md`

### ADV SDK DLLs
Copy ADV DLLs from `gps-timing-analysis/lib/` into the occultation-manager lib folder.
Check where OM currently stores its Openize DLLs (likely `occultation-manager/python/lib/`
or `occultation-manager/app/lib/`) and copy to the same location.

- [x] Identify the correct `lib/` folder location in `occultation-manager`
- [x] Copy ADV DLL files from `gps-timing-analysis/lib/` to that location
- [x] Verify `lib/unblock_dlls.ps1` (or equivalent) is also present for new users

---

## 2. Update led_line_delay_calibration.py After Move

- [x] Update any hardcoded `lib/` path for ADV DLL loading to match the new location in OM
- [ ] Test that ADV file replay mode still works from the new location
- [ ] Add **Save Calibration** button to `display_results()` (writes JSON to
  `Documents\occultation-tools\calibrations\line_delay\`)

---

## 3. Delete Source Files from gps-timing-analysis

Only delete after confirming the copies in occultation-manager are working.

- [ ] Delete `gps-timing-analysis/python/led_line_delay_calibration.py`
- [ ] Delete `gps-timing-analysis/python/adv_processing_iron.py`
- [ ] Delete `gps-timing-analysis/python/adv_helper.py`
- [ ] Delete `gps-timing-analysis/LED_LINE_DELAY_QUICKSTART.md`
- [ ] Delete `gps-timing-analysis/LED_LINE_DELAY_DEVELOPMENT_PLAN.md`

### Review before deleting (leave in gps-timing-analysis)
These files stay — do not delete:

- [ ] Confirm `example_adv_processing.py` is reference-only and can stay in gps-timing-analysis
- [ ] Confirm `test_adv_iron.py`, `test_adv_loading.py`, `test_timestamps.py` stay in gps-timing-analysis
- [ ] Check whether `gps-timing-analysis/python/light_curves_iron.py` is identical to
  `occultation-manager/python/light_curves_iron.py` — if so, delete the gps-timing-analysis copy

---

## 4. Update Documentation

### gps-timing-analysis/ReadMe.md
- [ ] Update the "LED Line Delay Calibration (SharpCap Add-in)" section to state the tool has
  moved to `occultation-manager`
- [ ] Update or remove the `execfile()` launch path example (was
  `C:\path\to\gps-timing-analysis\python\led_line_delay_calibration.py`)
- [ ] Add a pointer: "LED calibration tool is now part of Occultation Manager —
  see Tools → Camera Calibration → LED Line Delay Calibration"

### occultation-manager/python/ReadMe.md
- [ ] Add a section documenting the LED Line Delay Calibration module
- [ ] Document the `adv_processing_iron.py` and `adv_helper.py` dependencies
- [ ] Document the `lib/` DLL requirement and unblock step

### gps-timing-analysis/LED_LINE_DELAY_QUICKSTART.md (in new location)
- [ ] Update any paths that reference `gps-timing-analysis/` to reflect the new location in OM

---

## 5. Wire Up in Occultation Manager

- [ ] Add menu entry: Tools → Camera Calibration → LED Line Delay Calibration
- [ ] Entry invokes `led_line_delay_calibration.py` as a module within OM (not an external launch)
- [ ] Verify IronPython import works from `occultation-manager/python/`

---

## 6. Commit

- [ ] Commit changes to `gps-timing-analysis` (deletions + ReadMe update)
- [ ] Commit changes to `occultation-manager` (new files + ReadMe update + menu wiring)
- [ ] Consider a single combined commit message:
  `Move LED line delay calibration tool from gps-timing-analysis to occultation-manager`
