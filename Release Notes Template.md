# v1.0.0 - First Dedicated NTP Installer Release

## What This Is
Standalone Windows installer package for NTP guided setup, separated from analysis tools.

## Download
- `install_ntp_timing_guided.cmd` (recommended for most users)
- `install_ntp_timing_guided.ps1` (PowerShell entrypoint)

## Install Steps
1. Download `install_ntp_timing_guided.cmd`.
2. Double-click.
3. Click Yes on Administrator prompt.
4. The launcher downloads the latest guided PowerShell installer from GitHub.
5. If GitHub is unavailable and a previous local copy exists, you can choose to continue in offline mode.
6. Follow guided installer prompts.

## Notes
- This release is focused on installer usability for non-technical users.
- Legacy/testing script is still included under `scripts/legacy/`.
- Guided resource files are remote-first during normal runs and cached locally for offline reuse.
