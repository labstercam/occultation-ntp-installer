# Automated NTP + PPS Setup (Windows)

This guide describes automation options for NTP timing setup in this repository.

Recommended entrypoint for most users:
- `install_ntp_timing_guided.cmd`

## Guided Installer (recommended)

Recommended for most users because the CMD launcher bootstraps the latest guided PowerShell installer from GitHub and can fall back to an already-downloaded local copy when internet access is unavailable.

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
- Step 1 mode choice:
  - `Automatic install (recommended)` using `config/install.auto.template.ini` and Meinberg `/USE_FILE`.
  - `Guided install (manual screens)` interactive installer flow.
- Optional download/install of NTP Time Server Monitor.
- Optional GPS/PPS setup with COM auto-detection or manual COM entry.
- Country-based server setup using curated profiles and fallback logic.
- Marker-based managed updates to `ntp.conf` while preserving unrelated config.
- Static IP guidance during server setup, including suggested values and Windows 10/11 entry steps.
- Optional Windows QoS policy creation (DSCP 46) for NTP UDP port 123 traffic.
- Automatic registry backup of `HKLM\SYSTEM\CurrentControlSet\Services\NTP` before any changes.
- Restart prompt when configuration changes are detected.
- Transcript logging to repository `logs/` when run from a local clone.

## Step 1 Automatic Mode Details

- Template source: `config/install.auto.template.ini`
- Runtime-generated files (working folder):
  - generated INI passed to Meinberg
  - optional placeholder `ntp.conf` (when `UseConfigFile` is enabled)
  - Meinberg unattended log file
- Prompts:
  - Upgrade mode: `Upgrade` (recommended) vs `Reinstall`
  - Explicit warning: `Reinstall` can delete previous NTP configuration and servers
  - In `Upgrade`, prompt to import placeholder `UseConfigFile` (default: No)

## Standard-User Access Layout (Standalone-Friendly)

The guided installer can apply a layout that allows non-admin users to operate NTP tooling:

- Config path: `C:\ProgramData\NTP\etc\ntp.conf`
- Logs path: `C:\ProgramData\NTP\logs`
- Service config path (`-c`) is updated to ProgramData config path.
- ACLs are applied so standard users can:
  - edit `ntp.conf`
  - write log files
  - execute NTP `.cmd`/`.bat` scripts
  - start/stop/restart `NTP` service

Application behavior:
- Automatic Step 1: applied automatically.
- Guided/manual Step 1: prompted with recommendation.
- If any ACL/service-right change fails, warnings are shown with details.

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
- During a normal online run, the guided installer attempts GitHub raw URLs first and caches resource files locally.
- If the CMD launcher cannot reach GitHub and you choose to continue with the local installer, the guided installer switches to local-only mode for the rest of that run.
- In local-only mode, cached local resources are used when available and missing required files are reported clearly.

## Launcher behavior

- `install_ntp_timing_guided.cmd` always requests elevation first.
- It then attempts to download the latest `install_ntp_timing_guided.ps1` from GitHub.
- If the download fails and no local copy exists, the launcher exits with an error.
- If the download fails and a local copy exists, the launcher prompts once to continue with the local installer.
- When you continue with the local installer, the PowerShell script shows an `[OFFLINE MODE]` banner and does not attempt further remote downloads during that run.

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
  -GpsMode 17 `
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

If Program Files or TEMP/TMP environment paths are invalid/unavailable, the guided installer prompts for fallback absolute paths.

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
- For standard-user layout, verify:
  - non-admin user can edit `C:\ProgramData\NTP\etc\ntp.conf`
  - non-admin user can create files in `C:\ProgramData\NTP\logs`
  - non-admin user can run stop/start/restart scripts or equivalent service actions

## AU Server Selection UX Update

AU/NMI/University interactive prompts were clarified to reduce selection ambiguity:
- uses `Select up to 2 ... servers`
- explicitly says to enter server number(s) separated by comma
- retains fallback note: `Additional servers will be added in the next steps`

## GPS Mode Selection — NMEA Baud Rate Guidance

When **GPS NMEA-only** mode is selected in Step 3, the GPS mode prompt shows additional guidance:

```
For NMEA-only receivers, modes 1 (4800 baud) or 17 (9600 baud) are recommended.
NMEA data works most reliably at these lower baud rates.
Use recommended GPS mode 17 (9600 baud)? [Y/n]
```

- Default: **17** (9600 baud)
- Recommended alternative: **1** (4800 baud)
- Higher baud modes (33/49/65/81) are listed but not recommended for NMEA-only use
- **GPS PPS + NMEA** mode uses the same mode values but shows a generic prompt without the baud rate advisory

## GPS PPS — FTDI USB Serial Driver

When Step 3 is run in **GPS PPS + NMEA** mode, the installer prompts:

```
Have you already installed the FTDI USB serial driver for the GPS PPS device? [y/N]
```

- If **No**: downloads or locates `CDM212364_Setup.exe` and runs it
  - Local source: `resources\CDM212364_Setup.exe` (used if present)
  - Remote source: `https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/resources/CDM212364_Setup.exe`
  - After install, prompts user to connect the GPS PPS device to verify driver loaded
- If **Yes**: skips install and proceeds directly to COM detection
- GPS **NMEA-only** mode skips this step entirely

## Desktop Shortcut — Restart NTP

At the end of a successful install run the installer prompts whether to create a Desktop shortcut:

```
Do you want to add a Desktop shortcut for Restarting NTP (recommended)? [Y/n]
```

- Shortcut target: `<NTP install root>\bin\restartntp.bat`
- Shortcut location: all-users Desktop (`CommonDesktopDirectory`)
- Icon: `<NTP install root>\bin\restart.ico` if present, otherwise default
- If `restartntp.bat` is not found a warning is shown and the shortcut is skipped
- Requires the standard-user access layout to be applied so non-admin users can execute the script without elevation

## GPS Refclock Poll Interval

The GPS/PPS serial refclock lines in `config/ntp.conf.template` (and generated `ntp.conf`) use:

```
minpoll 6 maxpoll 7
```

This gives a 64–128 s adaptive polling range instead of the previous fixed 16 s (`minpoll 4 maxpoll 4`). A local GPS/PPS source is stable enough that faster polling adds no benefit and increases log and driver overhead.

## Windows QoS Priority for NTP (Step 5)

Step 5 creates two Windows Policy-based QoS rules using `New-NetQosPolicy`:

| Policy name | Match | DSCP |
|---|---|---|
| NTP Outbound Priority | UDP destination port 123 | 46 (EF) |
| NTP Inbound Priority | UDP source port 123 | 46 (EF) |

DSCP 46 is the Expedited Forwarding (EF) per-hop behaviour — the same class used for VoIP. The Windows kernel network scheduler de-queues EF-marked packets ahead of best-effort traffic. On managed networks with DiffServ-aware switches and routers the marking is also honoured by the infrastructure.

Behavior notes:
- Requires Windows 8 / Server 2012 or later.
- Any pre-existing policies with these names are removed before new ones are created.
- Step can be skipped and re-run independently at any time.
- On home/SOHO networks the benefit is limited to the local Windows scheduler; ISPs strip DSCP on internet traffic.
- Policies are written to `HKLM\SOFTWARE\Policies\Microsoft\Windows\QoS`.

## Registry Backup

Immediately after the user confirms they want to proceed, the installer runs `Backup-NtpRegistry`:

- Exports `HKLM\SYSTEM\CurrentControlSet\Services\NTP` using `reg export`.
- Saves to a timestamped `.reg` file in the user’s **Downloads** folder, e.g. `NTP_registry_backup_20260328_120000.reg`.
- Displays a Windows notification dialog confirming the backup path.
- If the NTP service registry key does not yet exist (first-time install before Meinberg has run) the backup is silently skipped with an info message.
- To restore: double-click the `.reg` file and confirm the import prompt.
