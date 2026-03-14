# Occultation NTP Installer (Windows)

Guided Windows installer for Meinberg NTP + GPS/PPS timing setup used in occultation workflows.

This repository includes a beginner CMD launcher, a guided PowerShell installer, and legacy/testing scripts.

## For Most Users

1. Download `install_ntp_timing_guided.cmd` from the latest release.
2. Double-click it.
3. Accept Administrator prompt (UAC).
4. Follow the guided installer prompts.

## Repository Files

- `install_ntp_timing_guided.cmd`:
  Beginner entry point (double-click, requests elevation automatically).
- `install_ntp_timing_guided.ps1`:
  Main guided installer logic.
- `scripts/legacy/setup_ntp_timing.ps1`:
  Legacy/testing automation script.
- `scripts/legacy/find_gps_com_port.ps1`:
  Legacy/testing helper for COM auto-detection.
- `config/ntp-country-servers.json`:
  Curated country profiles.
- `config/ntp.conf.template`:
  Base template used to generate managed config blocks.
- `resources/ntp_pool_zones.json`:
  Country-to-region pool mapping metadata.
- `resources/national_utc_ntp_servers.json`:
  National UTC/NTP authority metadata.

## Documentation

- `docs/automated-setup.md`:
  Detailed usage, automation flow, and advanced examples.

## Basic Troubleshooting

- If Windows blocks launch, right-click the file -> Properties -> Unblock (if shown), then run as Administrator.
- If PowerShell policy is restrictive, launch using:
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_ntp_timing_guided.ps1`
