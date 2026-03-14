# v1.1.0 - Improved Guided Installer Bootstrap And Offline Fallback

## What Changed
- `install_ntp_timing_guided.cmd` now always attempts to download the latest guided PowerShell installer from GitHub before running.
- If GitHub is unavailable and a previously downloaded local installer exists, the launcher now warns once and prompts whether to continue with the local copy.
- When local fallback is chosen, the guided installer now runs in an explicit `[OFFLINE MODE]` and avoids further remote downloads for that run.
- Guided installer resource handling is now consistent across Step 4 support files.
- Static IP guidance in Step 4 now shows suggested values directly inside the Windows 10/11 entry instructions.

## Why This Release Matters
- Single-file users can launch from the `.cmd` file more reliably.
- Reinstalls are safer because the launcher refreshes the main `.ps1` on each run when internet is available.
- Offline reruns are clearer and more predictable.
- Step 4 country server setup behaves better when files are not already present locally.

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
- Remote-first resource loading now behaves consistently for:
	- country server configuration
	- national UTC metadata
	- NTP pool metadata
	- `ntp.conf` template handling
- Successfully downloaded resource files are cached locally for later offline use.
- Misleading missing-local-file warnings during country summary/help flow were removed.
- Static IP help is clearer for Windows 10/11 users during Step 4.

## Notes
- Legacy/testing script is still included under `scripts/legacy/`.
- Current release behavior is latest-oriented: the launcher and remote resource URLs currently point to `main`.
- For fully immutable historical releases, see `docs/release-instructions.md`.
