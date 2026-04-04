# Occultation NTP Installer (Windows)

Guided Windows installer for Meinberg NTP + GPS/PPS timing setup used in occultation workflows. Designed to make installing and setting up NTP and GPS time as easy as it can be.

What it does:
1. Downloads and installs Meinberg NTP
2. Downloads and installs Meinberg NTP Time Server Monitor for Windows
3. Configures the NTP servers for specific countries to use their National Standard time server, and a set of good quality servers
4. Optionally sets Windows QoS priority (DSCP 46) for NTP UDP port 123 traffic and disables WiFi power saving
5. Assists with setting up GPS receivers for PPS and NMEA time

## Step 1 Install Modes

Step 1 now supports two modes:
1. `Automatic install (recommended)`:
  Uses `config/install.auto.template.ini`, generates a local INI, and runs Meinberg with `/USE_FILE=...`.
2. `Guided install (manual screens)`:
  Launches Meinberg interactively.

Automatic mode details:
- Prompts for `Upgrade` or `Reinstall` mode.
- Clearly warns that `Reinstall` can delete previous NTP configuration/servers.
- In `Upgrade`, prompts whether to import placeholder `UseConfigFile`.

## Standard-User Access Model

To support standalone private users, the guided installer can apply a standard-user layout:
1. Uses `ProgramData\NTP\etc\ntp.conf` for writable config.
2. Uses `ProgramData\NTP\logs` for writable log output.
3. Grants standard users modify rights for config/log paths.
4. Grants standard users execute rights on NTP `.cmd`/`.bat` scripts.
5. Grants standard users NTP service control rights (start/stop/restart).

Behavior by Step 1 mode:
- Automatic mode: applies this layout automatically.
- Guided/manual mode: prompts and recommends applying it.

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

## GPS Mode — NMEA Baud Rate Guidance

When **GPS NMEA-only** mode is selected in Step 5, the GPS mode prompt recommends:

- **Mode 17** — 9600 baud (default, recommended)
- **Mode 1** — 4800 baud (alternative recommended)

NMEA sentences work most reliably at these lower baud rates. Higher baud modes (33, 49, 65, 81) are available but not recommended for NMEA-only setups.

## GPS PPS — FTDI USB Serial Driver

When configuring GPS PPS + NMEA mode (Step 5), the installer asks:

```
Have you already installed the FTDI USB serial driver for the GPS PPS device? [y/N]
```

If No, it downloads `CDM212364_Setup.exe` from the repository `resources/` folder (or uses the local copy if already present) and runs it silently. After installation you are prompted to plug in the GPS PPS device to verify the driver loaded before COM port detection begins. GPS NMEA-only mode skips this step.

## Desktop Shortcut — Restart NTP

At the end of Step 1 (Meinberg NTP install) the installer automatically creates **Restart NTP.lnk** on the all-users Desktop pointing to `restartntp.bat` in the Meinberg `bin` folder. If `restart.ico` is present in the same folder it is used as the icon. Requires the standard-user layout (Step 1) to be applied so non-admin users can execute the script.

## GPS Refclock Poll Interval

The GPS/PPS serial refclock poll interval is set differently per mode:
- **PPS + NMEA**: `minpoll 4 maxpoll 4` (fixed 16 s) — keeps the PPS discipline loop tight.
- **NMEA-only**: no poll parameters set (NTP default of 16 s applies).

## Windows QoS Priority for NTP (Step 4)

Step 5 creates two Windows Policy-based QoS rules that mark NTP packets with **DSCP 46** (Expedited Forwarding — the same class used for VoIP):

- **NTP Outbound Priority** — marks packets sent to UDP port 123
- **NTP Inbound Priority** — marks packets received from UDP port 123

This causes the Windows kernel network scheduler to de-queue NTP traffic ahead of best-effort traffic. On managed networks with DiffServ-aware switches and routers, the DSCP marking is also honoured by the network infrastructure. On home or SOHO networks the benefit is confined to the local Windows scheduler.

Step 5 is optional and can be skipped or re-run at any time.

## Registry Backup

At startup, after you confirm you want to proceed, the installer automatically exports the `HKLM\SYSTEM\CurrentControlSet\Services\NTP` registry key to a timestamped `.reg` file in your **Downloads** folder (e.g. `NTP_registry_backup_20260328_120000.reg`). The backup path is printed to the installer console output. Double-clicking the file restores the key if needed. If the NTP service key does not yet exist the backup step is silently skipped.

## Basic Troubleshooting

- If Windows blocks launch, right-click the file -> Properties -> Unblock (if shown), then run as Administrator.
- If internet access is unavailable, the CMD launcher can continue with a previously downloaded local installer script after confirmation.
- Resource files used during country setup are fetched from GitHub when available and cached locally for later offline reruns.
- If PowerShell policy is restrictive, launch using:
  `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install_ntp_timing_guided.ps1`
- If Program Files or TEMP/TMP environment paths are unavailable or invalid, the installer now prompts for fallback paths.
