# Automated NTP + PPS Setup (Windows)

This guide describes automation options for NTP timing setup in this repository.

Recommended entrypoint for most users:
- `install_ntp_timing_guided.cmd`

## Guided Installer (recommended)

Beginner launch (double-click):

```cmd
install_ntp_timing_guided.cmd
```

PowerShell launch (Administrator):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_ntp_timing_guided.ps1
```

GitHub-first one-liner (always pulls latest `main`):

```powershell
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $u="https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/install_ntp_timing_guided.ps1"; $p=Join-Path $env:TEMP "install_ntp_timing_guided.ps1"; Invoke-WebRequest -UseBasicParsing -Uri $u -OutFile $p; powershell.exe -NoProfile -ExecutionPolicy Bypass -File $p
```

Equivalent two-step command:

```powershell
$installer = Join-Path $env:TEMP 'install_ntp_timing_guided.ps1'
Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/install_ntp_timing_guided.ps1' -OutFile $installer
powershell.exe -NoProfile -ExecutionPolicy Bypass -File $installer
```

## What the guided installer automates

- Optional download/install of Meinberg NTP.
- Optional download/install of NTP Time Server Monitor.
- Optional GPS/PPS setup with COM auto-detection or manual COM entry.
- Country-based server setup using curated profiles and fallback logic.
- Marker-based managed updates to `ntp.conf` while preserving unrelated config.
- Restart prompt when configuration changes are detected.
- Transcript logging to repository `logs/` when run from a local clone.

## Files used by guided installer

- Guided script: `install_ntp_timing_guided.ps1`
- CMD launcher: `install_ntp_timing_guided.cmd`
- Country servers: `config/ntp-country-servers.json`
- NTP Pool zones resource: `resources/ntp_pool_zones.json`
- National UTC/NTP inventory resource: `resources/national_utc_ntp_servers.json`
- Template: `config/ntp.conf.template`

Resource notes:
- `ntp-country-servers.json` drives generated `ntp.conf` server lines.
- `ntp_pool_zones.json` is used for `-Country Other` country-to-region mapping.
- `national_utc_ntp_servers.json` is consulted for `-Country Other` when no curated country profile exists.
- The scripts attempt GitHub raw URLs first, then fall back to local files.

## Legacy/testing script

`scripts/legacy/setup_ntp_timing.ps1` is retained for testing and advanced workflows.
For normal installs, use `install_ntp_timing_guided.cmd` or `install_ntp_timing_guided.ps1`.

Typical legacy run (Administrator):

```powershell
# From repository root
$com = .\scripts\legacy\find_gps_com_port.ps1

.\scripts\legacy\setup_ntp_timing.ps1 `
  -Country NZ `
  -ComPort $com `
  -GpsMode 18 `
  -MeinbergInstallerSilentArgs "<SILENT_ARGS>" `
  -NtpMonitorInstallerSilentArgs "<SILENT_ARGS>"
```

Dry-run example:

```powershell
.\scripts\legacy\setup_ntp_timing.ps1 -Country NZ -ComPort 1 -WhatIf
```

## Install path behavior

When install root is not provided, scripts choose in this order:
- `%ProgramFiles(x86)%\NTP` (if present/available)
- `%ProgramFiles%\NTP`

## Safety behavior

- Requires Administrator privileges.
- Backs up existing `ntp.conf` before managed updates.
- Writes generated configuration in ASCII.
- Supports skip switches for incremental runs (legacy script and guided flow options).

## Validation checks after setup

- Confirm `NTP` service is running.
- Confirm `ntp.conf` exists under the selected NTP install root.
- Validate peers/runtime values with `ntpq -pn` and `ntpq -c rv`.
- Validate PPS lock/source behavior in NTP Time Server Monitor.
