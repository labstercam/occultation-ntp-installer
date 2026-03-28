# v1.4.0 - Windows QoS Priority for NTP, Registry Backup

## What Changed
- **Step 5 (new):** Optional Windows QoS step that creates two Policy-based QoS rules marking NTP UDP port 123 traffic with DSCP 46 (Expedited Forwarding). Policy names: `NTP Outbound Priority` (dst port 123) and `NTP Inbound Priority` (src port 123). Any pre-existing policies with these names are replaced. Step can be skipped and re-run independently at any time.
- **Registry backup:** At installer startup (after the user confirms they want to proceed), the installer automatically exports `HKLM\SYSTEM\CurrentControlSet\Services\NTP` to a timestamped `.reg` file in the user's Downloads folder. A Windows notification dialog confirms the backup path. If the NTP service registry key does not yet exist the backup is silently skipped.
- Welcome screen updated to list Step 5.

## Why This Release Matters
- NTP timing accuracy on Windows benefits from the kernel network scheduler giving NTP packets priority over best-effort traffic. DSCP 46 (EF) is the standard marking for latency-sensitive traffic and is honoured by managed network infrastructure.
- The automatic registry backup gives users a simple one-click restore point before the installer modifies any registry values, reducing risk for first-time and upgrade installs.

## Download
- `install_ntp_timing_guided.cmd` (recommended for most users)
- `install_ntp_timing_guided.ps1` (PowerShell entrypoint)

## Install Steps
1. Download `install_ntp_timing_guided.cmd`.
2. Double-click.
3. Click Yes on Administrator prompt.
4. The launcher downloads the latest guided PowerShell installer from GitHub.
5. If GitHub is unavailable and a previous local copy exists, choose whether to continue in offline mode.
6. Follow guided installer prompts.

## Notes
- QoS policy creation requires Windows 8 / Server 2012 or later (`New-NetQosPolicy` cmdlet).
- DSCP marking is most impactful on managed enterprise/lab networks with DiffServ-aware switches. On home/SOHO networks the benefit is limited to the local Windows scheduler.
- Registry backup file is a standard `.reg` file; double-click to restore if needed.

---

# v1.3.0 - Desktop Shortcut, GPS Poll Interval Tuning, FTDI Driver Install, NMEA Baud Guidance

## What Changed
- Step 3 (GPS PPS + NMEA mode only): installer now prompts whether the FTDI USB serial driver has already been installed. If not, it downloads `CDM212364_Setup.exe` from the repository (or uses the local copy in `resources/`) and runs it. After install, the user is prompted to plug in the GPS PPS device to verify the driver loaded before proceeding to COM port detection.
- Step 3 GPS mode selection: when **NMEA-only** mode is chosen, the installer now explicitly recommends mode **1** (4800 baud) or **17** (9600 baud) and explains that NMEA data works most reliably at these lower baud rates. The advanced mode list also notes that 4800/9600 baud are recommended for NMEA.
- After installation completes, the installer now prompts whether to create a Desktop shortcut **"Restart NTP"** pointing to `restartntp.bat` in the Meinberg `bin` folder (all-users Desktop). The shortcut uses `restart.ico` from the same folder if present. Default answer is Yes.
- GPS refclock poll interval changed from `minpoll 4 maxpoll 4` (16 s fixed) to `minpoll 6 maxpoll 7` (64–128 s adaptive). This reduces unnecessary polling load on the local serial driver while staying well within the NTP discipline window.

## Why This Release Matters
- FTDI driver setup is now guided in-installer for GPS PPS users instead of being a manual pre-requisite step.
- NMEA-only users are now steered toward the correct baud rate at setup time, reducing mis-configuration.
- The Desktop shortcut gives users a quick, no-admin-required way to restart NTP after a config edit (when the standard-user layout is applied).
- The adjusted poll interval better matches the stability characteristics of a local GPS/PPS refclock and reduces log churn.

## Download
- `install_ntp_timing_guided.cmd` (recommended for most users)
- `install_ntp_timing_guided.ps1` (PowerShell entrypoint)

## Install Steps
1. Download `install_ntp_timing_guided.cmd`.
2. Double-click.
3. Click Yes on Administrator prompt.
4. The launcher downloads the latest guided PowerShell installer from GitHub.
5. If GitHub is unavailable and a previous local copy exists, choose whether to continue in offline mode.
6. Follow guided installer prompts.

## Notes
- FTDI driver install is only offered in GPS **PPS + NMEA** mode; GPS NMEA-only mode skips this step.
- Desktop shortcut creation requires `restartntp.bat` to already exist in the Meinberg `bin` folder; if it is missing a warning is shown and the shortcut is skipped.
- GPS poll interval change affects `config/ntp.conf.template` and the generated `ntp.conf`; existing installations are updated on the next config-applying run.

---

# v1.2.0 - Installer Modes, Standard-User Access, And AU Flow Fixes

## What Changed
- Step 1 now supports two explicit modes:
	- `Automatic install (recommended)` via `config/install.auto.template.ini`
	- `Guided install (manual screens)`
- Automatic mode now prompts for `Upgrade` vs `Reinstall` and clearly warns that `Reinstall` can delete previous NTP config/servers.
- In automatic `Upgrade` mode, installer now prompts whether to import placeholder `UseConfigFile`.
- Added standard-user access layout support for standalone users:
	- writable config at `ProgramData\NTP\etc\ntp.conf`
	- writable logs at `ProgramData\NTP\logs`
	- grants standard users script execute rights and NTP service start/stop/restart rights
- Automatic mode applies standard-user layout unprompted.
- Guided/manual mode recommends standard-user layout and prompts before applying.
- Added resilient prompts when Program Files or TEMP/TMP environment paths are invalid/unavailable.
- AU server selection flow fixed and clarified:
	- fixed scalar `.Count` error path when selecting no university servers
	- improved prompt wording for NMI/University selection and comma-separated input
- Additional array-return hardening added in country/region flows to avoid scalar `.Count` failures.

## Why This Release Matters
- Step 1 behavior is clearer and safer for both new installs and reinstalls.
- Standalone non-admin users can now run daily operations more reliably (service control, logging, config edits) when the recommended layout is applied.
- AU interactive server selection is less error-prone and easier to understand.
- Installer is more resilient on systems with unusual filesystem environment variable layouts.

## Download
- `install_ntp_timing_guided.cmd` (recommended for most users)
- `install_ntp_timing_guided.ps1` (PowerShell entrypoint)

## Install Steps
1. Download `install_ntp_timing_guided.cmd`.
2. Double-click.
3. Click Yes on Administrator prompt.
4. The launcher downloads the latest guided PowerShell installer from GitHub.
5. If GitHub is unavailable and a previous local copy exists, choose whether to continue in offline mode.
6. Follow guided installer prompts.

## Notable Improvements
- New automatic install template file: `config/install.auto.template.ini`
- Improved reporting for permission/service-right application outcomes
- Startup now prints resolved install root and working folder
- Program Files / TEMP fallback prompts added for path resilience

## Notes
- Legacy/testing script is still included under `scripts/legacy/`.
- Current release behavior is latest-oriented: the launcher and remote resource URLs currently point to `main`.
- For fully immutable historical releases, see `docs/release-instructions.md`.
