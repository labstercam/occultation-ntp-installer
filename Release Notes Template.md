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
