# Occultation NTP Installer (Windows)

Simple guided installer for Windows NTP + GPS/PPS setup used in occultation timing workflows.

## For Most Users

1. Download `install_ntp_timing_bootstrap.cmd` from the latest Release.
2. Double-click it.
3. Accept Administrator prompt (UAC).
4. Follow the guided installer.

## Files

- `install_ntp_timing_bootstrap.cmd`:
  Beginner entry point (double-click).
- `install_ntp_timing_bootstrap.ps1`:
  Bootstrap engine/fallback.
- `scripts/install_ntp_timing_guided.ps1`:
  Main guided installer logic.
- `scripts/legacy/setup_ntp_timing.ps1`:
  Legacy/testing only.

## Troubleshooting

If Windows blocks launch:
- Right-click file -> Properties -> Unblock (if shown), then Run as administrator.
- See `docs/troubleshooting.md`.
