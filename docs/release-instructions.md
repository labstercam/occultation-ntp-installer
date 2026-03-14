# Release Instructions

This document describes a practical release process for this repository, including semantic versioning guidance, pre-release checks, GitHub release steps, and the current limitations of the launcher-based distribution model.

## Current Release Model

This repository does not currently store a version number inside the scripts themselves. The release version is therefore defined by:

- the Git tag, such as `v1.1.0`
- the GitHub release title
- the release notes

Important limitation:

- `install_ntp_timing_guided.cmd` downloads `install_ntp_timing_guided.ps1` from the `main` branch
- the guided PowerShell installer also uses GitHub raw URLs that currently point to `main`

That means a GitHub release is not fully immutable today. A user who downloads an older release asset later may still run newer installer logic or newer resource files from `main`.

Practical meaning:

- the current release process is suitable for distributing the latest installer
- it is not ideal for preserving perfectly reproducible historical releases

If you want immutable releases later, see the section `Making Releases Fully Reproducible` below.

## Semantic Versioning

Use `vMAJOR.MINOR.PATCH`.

### Patch release

Use a patch bump for backward-compatible fixes and content updates.

Examples:

- bug fixes in installer flow
- wording and documentation updates
- improvements to prompts or warnings
- curated NTP server list updates
- metadata/resource updates that do not require users to change how they run the installer

Example:

- `v1.2.3` -> `v1.2.4`

### Minor release

Use a minor bump for backward-compatible feature additions.

Examples:

- adding a new guided step
- adding a new supported country workflow
- adding new installer behavior that users can benefit from without breaking existing usage
- adding new optional files, fallback behavior, or validation logic

Example:

- `v1.2.3` -> `v1.3.0`

### Major release

Use a major bump for breaking changes.

Examples:

- changing entrypoint filenames or expected usage
- changing generated config structure in a way that breaks prior assumptions
- removing a legacy workflow users may still depend on
- changing resource formats in a way older scripts cannot read

Example:

- `v1.2.3` -> `v2.0.0`

## Version Numbering Guidance For This Repo

Recommended default policy:

- bug fix only: patch
- documentation only: patch
- NTP server list or resource data updates only: patch
- new guided installer capability: minor
- breaking installer workflow or file layout change: major

If unsure between patch and minor:

- choose patch when users run the installer the same way and only get safer or better results
- choose minor when users gain a new capability or a visibly expanded workflow

## Pre-Release Checklist

Before creating a release:

1. Confirm the working tree is clean.
2. Confirm the intended changes are already committed to `main`.
3. Review `README.md`, `docs/automated-setup.md`, and release notes content for accuracy.
4. Test the guided installer in a normal online scenario.
5. Test the CMD launcher when the PowerShell script is missing.
6. Test offline fallback behavior by simulating GitHub unavailability after a previous successful download.
7. Test Step 4 country server configuration, including template/resource fallback behavior.
8. Confirm no file references or GitHub raw URLs still point to old repository names or paths.

Recommended smoke test scenarios:

- clean machine or clean folder: run `install_ntp_timing_guided.cmd`
- online install path: confirm latest `.ps1` is downloaded and runs
- offline fallback path: confirm single warning, user prompt, and `[OFFLINE MODE]` banner
- Step 4: confirm country server configuration completes successfully

## Preparing Release Notes

Use this template as the starting point:

- `Release Notes Template.md`

Update the following:

- version number in the heading
- short release title
- what changed
- install or upgrade notes if anything user-visible changed
- any known limitations

Good release notes should explain:

- what improved
- whether users need to do anything differently
- whether the release changes install behavior or fallback behavior

## Standard Release Procedure

### 1. Choose the version number

Pick the next semantic version using the rules above.

Examples:

- only bug fixes and docs: `v1.0.1`
- new fallback mode and resource caching behavior: `v1.1.0`

### 2. Finalize docs and release notes

Update:

- release notes text
- any README or docs affected by the change

### 3. Verify changes locally

Run your normal checks and smoke tests.

Suggested commands:

```powershell
git status
git log --oneline -5
```

### 4. Commit release-ready changes

Example:

```powershell
git add .
git commit -m "Prepare release v1.1.0"
```

### 5. Create the tag

Use an annotated tag:

```powershell
git tag -a v1.1.0 -m "Release v1.1.0"
```

### 6. Push branch and tag

```powershell
git push origin main
git push origin v1.1.0
```

### 7. Create the GitHub release

In GitHub:

1. Open the repository.
2. Open Releases.
3. Choose Draft a new release.
4. Select the tag, such as `v1.1.0`.
5. Set the release title, such as `v1.1.0 - Improved Guided Installer Fallbacks`.
6. Paste the prepared release notes.
7. Attach any additional packaged assets if desired.
8. Publish the release.

## Recommended Release Assets

Minimum:

- Git tag
- GitHub release notes
- GitHub auto-generated source archive

Useful additional asset:

- `install_ntp_timing_guided.cmd`

Optional but strongly recommended if you want easier offline distribution:

- a ZIP of the repository snapshot containing:
  - `install_ntp_timing_guided.cmd`
  - `install_ntp_timing_guided.ps1`
  - `config/`
  - `resources/`
  - docs

The ZIP makes the release more self-contained and reduces dependence on live GitHub raw downloads.

## Making Releases Fully Reproducible

The current setup is latest-oriented, not tag-pinned.

If you want a release to remain historically exact, choose one of these approaches before publishing:

### Option 1: Pin raw URLs to the release tag

For the release branch or release commit, change raw URLs from `main` to the specific tag.

Example:

- from `.../main/install_ntp_timing_guided.ps1`
- to `.../v1.1.0/install_ntp_timing_guided.ps1`

Apply the same idea to:

- `install_ntp_timing_guided.cmd`
- `config/ntp-country-servers.json`
- `config/ntp.conf.template`
- `resources/ntp_pool_zones.json`
- `resources/national_utc_ntp_servers.json`

This gives you an immutable release, but each release requires version-specific URL updates.

### Option 2: Publish a release ZIP and recommend local execution

Keep the current latest-oriented launcher for convenience, but also publish a ZIP containing the exact release snapshot.

This is usually the simplest operational approach.

### Option 3: Maintain a separate stable release channel

Use a stable branch or separate release asset process for users who need stronger reproducibility.

## Post-Release Checklist

After publishing:

1. Open the GitHub release page and verify title, tag, and notes.
2. Download the published asset as a user would.
3. Run `install_ntp_timing_guided.cmd` from a non-repo folder.
4. Confirm online bootstrap works.
5. Confirm offline fallback still behaves clearly after a prior successful run.
6. Confirm release links in README or docs still make sense.

## Suggested Release Title Format

Use:

- `v1.1.0 - Short human summary`

Examples:

- `v1.0.1 - Guided installer bug fixes`
- `v1.1.0 - Improved offline fallback and resource handling`

## Suggested Release Cadence

Use patch releases freely for:

- bug fixes
- documentation corrections
- server list maintenance

Use minor releases for grouped functional improvements rather than releasing every tiny feature separately.

## Quick Release Checklist

- Decide `vMAJOR.MINOR.PATCH`
- Update docs and release notes
- Test online bootstrap
- Test offline fallback
- Commit changes
- Create annotated tag
- Push `main` and tag
- Draft GitHub release
- Publish and smoke test the released asset