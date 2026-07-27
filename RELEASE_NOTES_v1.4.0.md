# v1.4.0 - NTP Time Server Monitor Configuration & Australian Server Improvements

## What Changed
- **NTP Time Server Monitor configuration:** The installer now configures NTP Time Server Monitor to run as Administrator, start automatically on Windows Startup, enable DNS lookup for readable server names, and appear in the System Tray
- **Australian server improvements:** Comprehensive three-tier server selection for Australia:
  - NMI UTC(AUS) servers (National Measurement Institute traceable time servers)
  - University public NTP servers  
  - AU pool fallback (au.pool.ntp.org plus 0..4.au.pool.ntp.org)
- **Registry backup:** At installer startup, automatically exports the NTP service registry key to a timestamped `.reg` file in the user's Downloads folder
- **Windows QoS Priority for NTP:** Optional Step 5 creates two Policy-based QoS rules marking NTP UDP port 123 traffic with DSCP 46 (Expedited Forwarding)
- **Desktop Shortcut:** Automatically creates "Restart NTP.lnk" on the all-users Desktop pointing to restartntp.bat
- **GPS improvements:** FTDI driver installation guidance, NMEA baud rate recommendations, and GPS poll interval tuning
- **Standard-user access model:** Supports standalone users with writable config at `ProgramData\NTP\etc\ntp.conf` and logs at `ProgramData\NTP\logs`
- **Installer modes:** Step 1 now supports Automatic install (recommended) and Guided install (manual screens)

## Why This Release Matters
- NTP Time Server Monitor is now properly configured out-of-the-box with Administrator privileges, Startup execution, and DNS lookup enabled
- Australian users get a comprehensive, guided server selection process with clear registration instructions for NMI servers
- Automatic registry backup provides a safety net before installer modifications
- Windows QoS priority improves NTP timing accuracy on managed networks
- Standard-user access enables non-admin users to control NTP service and edit configuration

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
- NTP Time Server Monitor configuration requires locating the executable after installation; if not found, manual configuration may be needed
- NMI server registration requires emailing time@measurement.gov.au and obtaining a static IP address from your ISP
- QoS policy creation requires Windows 8 / Server 2012 or later (`New-NetQosPolicy` cmdlet)
- Registry backup file is a standard `.reg` file; double-click to restore if needed
- Standard-user layout is automatically applied in Automatic mode, recommended in Guided mode

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