# Occultation NTP Installer (Windows)

Guided Windows installer for Meinberg NTP + GPS/PPS timing setup used in occultation workflows. Designed to make installing and setting up NTP and GPS time as easy as it can be.

What it does:
1. Downloads and installs Meinberg NTP
2. Downloads and installs Meinberg NTP Time Server Monitor for Windows
3. Assists with setting up GPS receivers for PPS and NMEA time
4. Configures the NTP servers for specific countries to use their National Standard time server, and a set of good quality servers

## For Most Users

1. Download `install_ntp_timing_guided.cmd` from the latest release. [install_ntp_timing_guided.cmd](https://github.com/labstercam/occultation-ntp-installer/releases/download/v1.1.0/install_ntp_timing_guided.cmd)
2. Double-click it.
3. Accept Administrator prompt (UAC).
4. The launcher downloads the latest `install_ntp_timing_guided.ps1` from GitHub.
5. If GitHub is unavailable but a previously downloaded local copy exists, you can choose to continue in offline mode.
6. Follow the guided installer prompts.

## Repository Files

- `install_ntp_timing_guided.cmd`:
  Beginner entry point (double-click, requests elevation automatically, refreshes the main installer from GitHub, and offers offline fallback when GitHub is unavailable).
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
- `docs/release-instructions.md`:
  Release checklist, semantic versioning guidance, and GitHub release procedure.

## Basic Troubleshooting

- If Windows blocks launch, right-click the file -> Properties -> Unblock (if shown), then run as Administrator.
- If internet access is unavailable, the CMD launcher can continue with a previously downloaded local installer script after confirmation.
- Resource files used during country setup are fetched from GitHub when available and cached locally for later offline reruns.
- If PowerShell policy is restrictive, launch using:
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_ntp_timing_guided.ps1`
