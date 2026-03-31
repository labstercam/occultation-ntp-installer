[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

function Write-Info([string]$Message) { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
function Write-WarnMsg([string]$Message) { Write-Host "[WARN] $Message" -ForegroundColor Yellow }
function Write-Ok([string]$Message) { Write-Host "[ OK ] $Message" -ForegroundColor Green }
function Write-Step([string]$Message) { Write-Host "`n=== $Message ===" -ForegroundColor Green }

function Backup-NtpRegistry {
    $ntpKeyPath = "HKLM\SYSTEM\CurrentControlSet\Services\NTP"
    $downloadsFolder = [System.IO.Path]::Combine([System.Environment]::GetFolderPath("UserProfile"), "Downloads")
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = [System.IO.Path]::Combine($downloadsFolder, "NTP_registry_backup_$timestamp.reg")

    if (-not (Test-Path -LiteralPath "HKLM:\SYSTEM\CurrentControlSet\Services\NTP")) {
        Write-Info "NTP service registry key not yet present; backup skipped."
        return
    }

    try {
        $null = reg export $ntpKeyPath $backupPath /y 2>&1
        Write-Ok "Registry backup saved to: $backupPath"
        Write-Host ""
        [System.Windows.Forms.MessageBox]::Show(
            "A backup of the NTP registry key has been saved to:`n`n$backupPath`n`nYou can restore it by double-clicking the .reg file if needed.",
            "Registry Backup Created",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
    }
    catch {
        Write-WarnMsg ("Registry backup failed: {0}" -f $_.Exception.Message)
        Write-WarnMsg "You may want to manually export HKLM\SYSTEM\CurrentControlSet\Services\NTP before continuing."
    }
}

$script:RemoteDownloadsDisabled = ($env:OCNTP_REMOTE_DISABLED -eq '1')
$script:RemoteOverrideNoticeShown = $false

function Show-RemoteOverrideNotice {
    if ($script:RemoteDownloadsDisabled -and -not $script:RemoteOverrideNoticeShown) {
        Write-WarnMsg "Remote downloads are disabled by launcher choice. Using local files only where available."
        $script:RemoteOverrideNoticeShown = $true
    }
}

function Wait-BeforeCloseIfNeeded {
    # Keep standalone ConsoleHost windows open so users can read output/errors.
    if ($Host.Name -ne 'ConsoleHost') {
        return
    }

    # VS Code integrated terminals should not be blocked.
    if ($env:TERM_PROGRAM -eq 'vscode') {
        return
    }

    try {
        [void](Read-Host "Press Enter to close")
    }
    catch {
        # Ignore prompt failures during shutdown.
    }
}

function Read-YesNo {
    param(
        [string]$Prompt,
        [bool]$DefaultYes = $true
    )

    while ($true) {
        $suffix = if ($DefaultYes) { "[Y/n]" } else { "[y/N]" }
        $reply = Read-Host "$Prompt $suffix"
        if ([string]::IsNullOrWhiteSpace($reply)) {
            return $DefaultYes
        }

        switch -Regex ($reply.Trim()) {
            '^(y|yes)$' { return $true }
            '^(n|no)$' { return $false }
            default { Write-WarnMsg "Please answer y or n." }
        }
    }
}

function Read-StepAction {
    param([string]$Title)

    while ($true) {
        $choice = Read-Host "Action for '$Title' (Enter = Install): [I]nstall / [S]kip / E[x]it"
        if ([string]::IsNullOrWhiteSpace($choice)) {
            return "Install"
        }

        switch -Regex ($choice.Trim().ToLowerInvariant()) {
            '^(i|install)$' { return "Install" }
            '^(s|skip)$' { return "Skip" }
            '^(x|e|exit|quit)$' { return "Exit" }
            default { Write-WarnMsg "Enter I, S, or X." }
        }
    }
}

function Confirm-Step {
    param(
        [string]$Title,
        [string[]]$Details
    )

    Write-Step $Title
    foreach ($line in $Details) {
        Write-Host (" - {0}" -f $line)
    }

    $action = Read-StepAction -Title $Title
    if ($action -eq "Exit") {
        throw "Installer exited by user at step: $Title"
    }

    if ($action -eq "Skip") {
        Write-WarnMsg "Step skipped: $Title"
        return $false
    }

    return $true
}

function Assert-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($id)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "This installer needs Administrator rights. Close this window, then right-click PowerShell and select 'Run as administrator', and run the installer again."
    }
}

function Test-DirectoryWritable {
    param([string]$Path)

    try {
        if ([string]::IsNullOrWhiteSpace($Path)) {
            return $false
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -ItemType Directory -Path $Path -Force | Out-Null
        }

        $probe = Join-Path $Path (".ocntp_write_test_{0}.tmp" -f ([guid]::NewGuid().ToString("N")))
        Set-Content -LiteralPath $probe -Value "ok" -Encoding ASCII
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        return $false
    }
}

function Read-PathOrExit {
    param([string]$Prompt)

    while ($true) {
        $raw = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-WarnMsg "Path cannot be blank."
            continue
        }

        $value = $raw.Trim()
        if ($value -match '^(x|exit|quit)$') {
            throw "Installer exited by user while entering required filesystem paths."
        }

        return $value
    }
}

function Resolve-DefaultInstallRoot {
    $pf86 = ${env:ProgramFiles(x86)}
    $pf64 = $env:ProgramFiles

    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($pf86)) { $candidates += (Join-Path $pf86 "NTP") }
    if (-not [string]::IsNullOrWhiteSpace($pf64)) { $candidates += (Join-Path $pf64 "NTP") }

    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) {
            return (Resolve-Path -LiteralPath $c).Path
        }
    }

    if ($candidates.Count -gt 0) {
        $preferred = $candidates[0]
        $preferredParent = Split-Path -Parent $preferred
        if (-not [string]::IsNullOrWhiteSpace($preferredParent) -and (Test-Path -LiteralPath $preferredParent)) {
            return $preferred
        }

        Write-WarnMsg "Program Files path appears invalid on this system."
        Write-WarnMsg ("Detected candidate was: {0}" -f $preferred)
    }

    Write-WarnMsg "Could not automatically determine a valid Program Files install location."
    Write-WarnMsg "Enter a writable install folder for NTP (for example D:\\Program Files\\NTP)."
    Write-Info "Type 'x' to exit."

    while ($true) {
        $manualRoot = Read-PathOrExit -Prompt "Enter NTP install root path"
        if (-not [System.IO.Path]::IsPathRooted($manualRoot)) {
            Write-WarnMsg "Please enter an absolute path."
            continue
        }

        $parent = Split-Path -Parent $manualRoot
        if ([string]::IsNullOrWhiteSpace($parent)) {
            Write-WarnMsg "Could not determine parent folder for that path."
            continue
        }

        if (-not (Test-DirectoryWritable -Path $parent)) {
            Write-WarnMsg ("Parent folder is not writable: {0}" -f $parent)
            continue
        }

        return $manualRoot
    }

    throw "Could not determine Program Files directory on this system."
}

function Resolve-WorkingTempDirectory {
    param([string]$SubFolder = "occultation-ntp-installer")

    $bases = @()
    if (-not [string]::IsNullOrWhiteSpace($env:TEMP)) { $bases += $env:TEMP }
    if (-not [string]::IsNullOrWhiteSpace($env:TMP)) { $bases += $env:TMP }
    $bases = @($bases | Select-Object -Unique)

    foreach ($base in $bases) {
        try {
            if (-not (Test-Path -LiteralPath $base)) {
                continue
            }
            $candidate = Join-Path $base $SubFolder
            if (Test-DirectoryWritable -Path $candidate) {
                return $candidate
            }
        }
        catch {
            continue
        }
    }

    Write-WarnMsg "Could not use TEMP/TMP for working files."
    Write-WarnMsg "Enter a writable working folder for downloads and generated installer files."
    Write-Info "Type 'x' to exit."

    while ($true) {
        $manualDir = Read-PathOrExit -Prompt "Enter working folder path"
        if (-not [System.IO.Path]::IsPathRooted($manualDir)) {
            Write-WarnMsg "Please enter an absolute path."
            continue
        }

        if (-not (Test-DirectoryWritable -Path $manualDir)) {
            Write-WarnMsg ("Folder is not writable: {0}" -f $manualDir)
            continue
        }

        return $manualDir
    }

    throw "Could not determine Program Files directory on this system."
}

function Convert-ToTextFromResponseContent {
    param([object]$Content)

    if ($Content -is [string]) {
        return $Content
    }

    if ($Content -is [System.Array]) {
        return (($Content | ForEach-Object { [char]$_ }) -join '')
    }

    return [string]$Content
}

function Get-ExpectedSha256FromUrl {
    param([string]$ShaUrl)

    if ([string]::IsNullOrWhiteSpace($ShaUrl)) {
        return ""
    }

    $resp = Invoke-WebRequest -Uri $ShaUrl -UseBasicParsing
    $text = Convert-ToTextFromResponseContent -Content $resp.Content
    $line = ($text -split "`n" | Select-Object -First 1).Trim()
    if ($line -match '^(?<hash>[A-Fa-f0-9]{64})\s+\*?.+$') {
        return $matches['hash'].ToLowerInvariant()
    }

    throw "Could not parse SHA256 checksum from $ShaUrl"
}

function Assert-FileSha256 {
    param(
        [string]$Path,
        [string]$ExpectedSha256,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($ExpectedSha256)) {
        Write-WarnMsg "No expected SHA256 provided for $Label. Skipping checksum validation."
        return
    }

    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    $expected = $ExpectedSha256.ToLowerInvariant()

    if ($actual -ne $expected) {
        throw "$Label checksum mismatch. Expected $expected but got $actual"
    }

    Write-Ok "$Label checksum validated (SHA256)"
}

function Invoke-InstallerDownload {
    param(
        [string]$Url,
        [string]$OutputPath,
        [string]$Label
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw "$Label URL is empty."
    }

    if ($script:RemoteDownloadsDisabled) {
        Show-RemoteOverrideNotice
        if (Test-Path -LiteralPath $OutputPath) {
            Write-WarnMsg "Using local cached $Label installer: $OutputPath"
            return
        }

        throw "$Label download skipped (remote disabled) and no local cached installer was found: $OutputPath"
    }

    Write-Info "Downloading $Label from $Url"
    Invoke-WebRequest -Uri $Url -OutFile $OutputPath -UseBasicParsing
    Write-Ok "Downloaded $Label to $OutputPath"
}

function Ensure-FileAvailableWithRemote {
    param(
        [string]$LocalPath,
        [string]$RemoteUrl,
        [string]$Label
    )

    if (Test-Path -LiteralPath $LocalPath) {
        return $true
    }

    if ($script:RemoteDownloadsDisabled) {
        Show-RemoteOverrideNotice
        Write-WarnMsg ("{0} not found locally and remote downloads are disabled: {1}" -f $Label, $LocalPath)
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($RemoteUrl)) {
        Write-WarnMsg ("{0} not found locally and no remote URL is configured: {1}" -f $Label, $LocalPath)
        return $false
    }

    try {
        $dir = Split-Path -Parent $LocalPath
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        Write-Info ("Downloading missing {0} from GitHub..." -f $Label)
        Invoke-WebRequest -Uri $RemoteUrl -OutFile $LocalPath -UseBasicParsing -TimeoutSec 20
        Write-Ok ("Downloaded {0}: {1}" -f $Label, $LocalPath)
        return $true
    }
    catch {
        Write-WarnMsg ("Could not download {0}: {1}" -f $Label, $_.Exception.Message)
        return $false
    }
}

function Refresh-RemoteFile {
    # Always attempts to download the latest version from GitHub.
    # If the download succeeds the local file is replaced.
    # If the download fails and a local copy already exists, the local copy is kept and $true is returned.
    # If the download fails and no local copy exists, $false is returned.
    param(
        [string]$LocalPath,
        [string]$RemoteUrl,
        [string]$Label
    )

    if ($script:RemoteDownloadsDisabled) {
        Show-RemoteOverrideNotice
        if (Test-Path -LiteralPath $LocalPath) {
            Write-Info ("{0}: using existing local copy (remote downloads disabled)." -f $Label)
            return $true
        }
        Write-WarnMsg ("{0} not found locally and remote downloads are disabled: {1}" -f $Label, $LocalPath)
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($RemoteUrl)) {
        Write-WarnMsg ("{0}: no remote URL configured, skipping refresh." -f $Label)
        return (Test-Path -LiteralPath $LocalPath)
    }

    try {
        $dir = Split-Path -Parent $LocalPath
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        Write-Info ("Refreshing {0} from GitHub..." -f $Label)
        Invoke-WebRequest -Uri $RemoteUrl -OutFile $LocalPath -UseBasicParsing -TimeoutSec 20
        Write-Ok ("Refreshed {0}: {1}" -f $Label, $LocalPath)
        return $true
    }
    catch {
        if (Test-Path -LiteralPath $LocalPath) {
            Write-WarnMsg ("Could not refresh {0} ({1}); using existing local copy." -f $Label, $_.Exception.Message)
            return $true
        }
        Write-WarnMsg ("Could not download {0} and no local copy exists: {1}" -f $Label, $_.Exception.Message)
        return $false
    }
}

function Check-SupportFileAvailability {
    param(
        [string]$TemplatePath,
        [string]$CountryConfigPath,
        [string]$NationalUtcPath,
        [string]$PoolZonesPath,
        [string]$AutoIniTemplatePath,
        [string]$TemplateRemoteUrl,
        [string]$CountryConfigRemoteUrl,
        [string]$NationalUtcRemoteUrl,
        [string]$PoolZonesRemoteUrl,
        [string]$AutoIniTemplateRemoteUrl
    )

    Write-Info "Checking local support file availability..."

    # Always refresh the template so config changes reach machines that already have a local copy.
    $templateOk = Refresh-RemoteFile -LocalPath $TemplatePath -RemoteUrl $TemplateRemoteUrl -Label "ntp.conf template"
    if (-not $templateOk) {
        Write-WarnMsg "ntp.conf template is currently unavailable. Related steps may require internet access or local files."
    }

    $checks = @(
        [PSCustomObject]@{ Label = "country server config";        Local = $CountryConfigPath;        Remote = $CountryConfigRemoteUrl },
        [PSCustomObject]@{ Label = "national UTC/NTP inventory";   Local = $NationalUtcPath;          Remote = $NationalUtcRemoteUrl },
        [PSCustomObject]@{ Label = "NTP pool zones";               Local = $PoolZonesPath;            Remote = $PoolZonesRemoteUrl },
        [PSCustomObject]@{ Label = "automatic install template";   Local = $AutoIniTemplatePath;      Remote = $AutoIniTemplateRemoteUrl }
    )

    foreach ($c in $checks) {
        $ok = Ensure-FileAvailableWithRemote -LocalPath $c.Local -RemoteUrl $c.Remote -Label $c.Label
        if (-not $ok) {
            Write-WarnMsg ("{0} is currently unavailable. Related steps may require internet access or local files." -f $c.Label)
        }
    }
}

function Test-MeinbergAutomaticInstallSuccess {
    param([string]$InstallerLogPath)

    if ([string]::IsNullOrWhiteSpace($InstallerLogPath)) {
        return $false
    }

    if (-not (Test-Path -LiteralPath $InstallerLogPath)) {
        return $false
    }

    try {
        $logText = Get-Content -LiteralPath $InstallerLogPath -Raw -ErrorAction Stop
        return ($logText -match '(?i)installation\s+successfully\s+completed')
    }
    catch {
        Write-WarnMsg ("Could not read Meinberg installer log for success check: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Install-Exe {
    param(
        [string]$InstallerPath,
        [string[]]$Arguments,
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $InstallerPath)) {
        throw "$Label installer not found: $InstallerPath"
    }

    $startArgs = @{
        FilePath = $InstallerPath
        PassThru = $true
        Wait     = $true
    }

    $argList = @($Arguments | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

    if ($argList.Count -eq 0) {
        Write-WarnMsg "$Label silent arguments are empty; starting installer without arguments (interactive mode)."
    }
    else {
        $startArgs.ArgumentList = $argList
        Write-Info ("{0} arguments: {1}" -f $Label, ($argList -join " "))
    }

    Write-Info "Starting installer: $Label"
    $proc = Start-Process @startArgs
    if ($proc.ExitCode -ne 0) {
        throw "$Label installer exited with code $($proc.ExitCode)"
    }

    Write-Ok "$Label installation completed"
}

function Set-IniKeyValue {
    param(
        [string[]]$Lines,
        [string]$Section,
        [string]$Key,
        [string]$Value
    )

    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($line in @($Lines)) {
        $list.Add([string]$line)
    }

    $sectionStart = -1
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim().Equals("[$Section]", [System.StringComparison]::OrdinalIgnoreCase)) {
            $sectionStart = $i
            break
        }
    }

    if ($sectionStart -lt 0) {
        if ($list.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($list[$list.Count - 1])) {
            $list.Add("")
        }
        $list.Add("[$Section]")
        $list.Add(("{0}={1}" -f $Key, $Value))
        return @($list)
    }

    $sectionEnd = $list.Count
    for ($i = $sectionStart + 1; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim().StartsWith("[") -and $list[$i].Trim().EndsWith("]")) {
            $sectionEnd = $i
            break
        }
    }

    $keyFound = $false
    $keyPattern = "^\s*{0}\s*=" -f [regex]::Escape($Key)
    for ($i = $sectionStart + 1; $i -lt $sectionEnd; $i++) {
        if ([regex]::IsMatch($list[$i], $keyPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $list[$i] = ("{0}={1}" -f $Key, $Value)
            $keyFound = $true
            break
        }
    }

    if (-not $keyFound) {
        $list.Insert($sectionEnd, ("{0}={1}" -f $Key, $Value))
    }

    return @($list)
}

function Remove-IniKey {
    param(
        [string[]]$Lines,
        [string]$Section,
        [string]$Key
    )

    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($line in @($Lines)) {
        $list.Add([string]$line)
    }

    $sectionStart = -1
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim().Equals("[$Section]", [System.StringComparison]::OrdinalIgnoreCase)) {
            $sectionStart = $i
            break
        }
    }

    if ($sectionStart -lt 0) {
        return @($list)
    }

    $sectionEnd = $list.Count
    for ($i = $sectionStart + 1; $i -lt $list.Count; $i++) {
        if ($list[$i].Trim().StartsWith("[") -and $list[$i].Trim().EndsWith("]")) {
            $sectionEnd = $i
            break
        }
    }

    $keyPattern = "^\s*{0}\s*=" -f [regex]::Escape($Key)
    for ($i = $sectionEnd - 1; $i -gt $sectionStart; $i--) {
        if ([regex]::IsMatch($list[$i], $keyPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $list.RemoveAt($i)
        }
    }

    return @($list)
}

function Prepare-MeinbergAutomaticInstallFiles {
    param(
        [string]$IniTemplatePath,
        [string]$TempDir,
        [string]$InstallRoot,
        [ValidateSet("Upgrade", "Reinstall")]
        [string]$UpgradeMode = "Upgrade",
        [bool]$ApplyUseConfigFile = $true
    )

    if (-not (Test-Path -LiteralPath $IniTemplatePath)) {
        throw "Automatic install template not found: $IniTemplatePath"
    }

    if (-not (Test-Path -LiteralPath $TempDir)) {
        New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $generatedIniPath = Join-Path $TempDir ("meinberg_install_{0}.ini" -f $stamp)
    $generatedConfPath = Join-Path $TempDir ("ntp_placeholder_{0}.conf" -f $stamp)
    $installerLogPath = Join-Path $TempDir ("meinberg_uam_{0}.log" -f $stamp)

    $placeholderConf = @(
        "# Minimal placeholder ntp.conf for Meinberg automatic install.",
        "# Internet servers are configured later by Step 4 of this guided installer.",
        ""
    )
    Set-Content -LiteralPath $generatedConfPath -Value $placeholderConf -Encoding ASCII

    $iniLines = @(Get-Content -LiteralPath $IniTemplatePath)
    $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Installer" -Key "InstallDir" -Value $InstallRoot)
    $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Installer" -Key "UpgradeMode" -Value $UpgradeMode)
    $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Installer" -Key "Logfile" -Value $installerLogPath)
    $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Installer" -Key "Silent" -Value "Yes")
    $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Service" -Key "ServiceAccount" -Value "@SYSTEM")
    if ($ApplyUseConfigFile) {
        $iniLines = @(Set-IniKeyValue -Lines $iniLines -Section "Configuration" -Key "UseConfigFile" -Value $generatedConfPath)
    }
    else {
        $iniLines = @(Remove-IniKey -Lines $iniLines -Section "Configuration" -Key "UseConfigFile")
    }

    Set-Content -LiteralPath $generatedIniPath -Value $iniLines -Encoding ASCII

    return [PSCustomObject]@{
        IniPath          = $generatedIniPath
        PlaceholderPath  = $generatedConfPath
        InstallerLogPath = $installerLogPath
    }
}

function Read-MeinbergAutomaticUpgradeMode {
    param(
        [string]$InstallRoot,
        [string]$NtpConfPath
    )

    $hasExistingInstall = (Test-Path -LiteralPath (Join-Path $InstallRoot "bin\ntpd.exe")) -or (Test-Path -LiteralPath $NtpConfPath)

    while ($true) {
        Write-Host ""
        Write-Host "Automatic install upgrade mode:" -ForegroundColor Cyan
        Write-Host "  1) Upgrade (recommended; tries to preserve existing configuration)"
        Write-Host "  2) Reinstall (clean install; deletes previous NTP configuration and servers)"

        if ($hasExistingInstall) {
            Write-WarnMsg "Existing NTP install/config was detected. Reinstall will delete previous NTP configuration and servers."
        }
        else {
            Write-Info "No existing NTP install/config detected."
        }

        $choice = Read-Host "Enter 1 or 2 (default 1)"
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice.Trim() -eq "1") {
            return "Upgrade"
        }
        if ($choice.Trim() -eq "2") {
            Write-WarnMsg "You selected Reinstall. This will delete previous NTP configuration and servers."
            if (Read-YesNo -Prompt "Continue with Reinstall mode?" -DefaultYes $false) {
                return "Reinstall"
            }
            continue
        }

        Write-WarnMsg "Enter 1 or 2."
    }
}

function Read-MeinbergAutomaticUseConfigChoice {
    param([string]$UpgradeMode)

    if ($UpgradeMode -eq "Reinstall") {
        return $true
    }

    Write-WarnMsg "Upgrade mode selected. Importing a placeholder UseConfigFile can overwrite existing NTP configuration and servers."
    return (Read-YesNo -Prompt "In Upgrade mode, import placeholder config via UseConfigFile?" -DefaultYes $false)
}

function Read-MeinbergInstallMode {
    param([string]$IniTemplatePath)

    while ($true) {
        Write-Host ""
        Write-Host "Choose Meinberg setup mode:" -ForegroundColor Cyan
        Write-Host "  1) Automatic install (recommended)"
        Write-Host "  2) Guided install (manual screens)"

        $choice = Read-Host "Enter 1 or 2 (default 1)"
        if ([string]::IsNullOrWhiteSpace($choice) -or $choice.Trim() -eq "1") {
            if (-not (Test-Path -LiteralPath $IniTemplatePath)) {
                Write-WarnMsg "Automatic install template was not found: $IniTemplatePath"
                Write-WarnMsg "Automatic install is unavailable."
                if (Read-YesNo -Prompt "Continue with guided install mode instead?" -DefaultYes $true) {
                    return "Guided"
                }
                continue
            }
            return "Automatic"
        }

        if ($choice.Trim() -eq "2") {
            return "Guided"
        }

        Write-WarnMsg "Enter 1 or 2."
    }
}

function Set-LoggingConfig {
    param(
        [string]$NtpConfPath,
        [string]$StatsDir
    )

    if (-not (Test-Path -LiteralPath (Split-Path -Parent $NtpConfPath))) {
        New-Item -ItemType Directory -Path (Split-Path -Parent $NtpConfPath) -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $StatsDir)) {
        New-Item -ItemType Directory -Path $StatsDir -Force | Out-Null
    }

    $content = ""
    if (Test-Path -LiteralPath $NtpConfPath) {
        $content = Get-Content -Raw -LiteralPath $NtpConfPath
    }

    if ([string]::IsNullOrWhiteSpace($content)) {
        $content = @(
            "# Placeholder ntp.conf created by guided installer.",
            "# Server selection is configured in the country setup step.",
            ""
        ) -join [Environment]::NewLine
    }

    $required = @(
        "enable stats",
        "statsdir `"$StatsDir`"",
        "statistics loopstats",
        "statistics peerstats"
    )

    foreach ($line in $required) {
        if ($content -notmatch [regex]::Escape($line)) {
            $content += [Environment]::NewLine + $line
        }
    }

    Set-Content -LiteralPath $NtpConfPath -Value $content -Encoding ASCII
    Write-Ok "Logging directives prepared in $NtpConfPath"
}

function Set-NtpServiceConfigPath {
    param([string]$ConfigPath)

    $svcRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\NTP"
    if (-not (Test-Path -LiteralPath $svcRegPath)) {
        Write-WarnMsg "NTP service registry path not found. Service config path was not updated."
        return $false
    }

    $currentImagePath = ""
    try {
        $currentImagePath = [string](Get-ItemProperty -LiteralPath $svcRegPath -Name "ImagePath" -ErrorAction Stop).ImagePath
    }
    catch {
        Write-WarnMsg "Could not read NTP service ImagePath."
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($currentImagePath)) {
        Write-WarnMsg "NTP service ImagePath is empty; could not update -c config path."
        return $false
    }

    $updatedImagePath = $currentImagePath
    if ($updatedImagePath -match '(?i)\s-c\s+') {
        $updatedImagePath = $updatedImagePath -replace '(?i)(\s-c\s+)("[^"]*"|\S+)', ('$1"' + $ConfigPath + '"')
    }
    else {
        $updatedImagePath = ($updatedImagePath.TrimEnd() + ' -c "' + $ConfigPath + '"')
    }

    if ($updatedImagePath -eq $currentImagePath) {
        return $true
    }

    try {
        Set-ItemProperty -LiteralPath $svcRegPath -Name "ImagePath" -Value $updatedImagePath -ErrorAction Stop
        Write-Ok ("Updated NTP service config path to: {0}" -f $ConfigPath)
        return $true
    }
    catch {
        Write-WarnMsg ("Failed to update NTP service ImagePath: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Grant-UsersNtpServiceControl {
    # Include query/read rights in addition to start/stop/restart so GUI tools
    # like NTP Time Server Monitor can open and manage the service reliably.
    $desiredAce = "(A;;CCLCSWRPWPDTLOCRRC;;;BU)"
    $legacyAce = "(A;;LCRPWPDTLO;;;BU)"

    $sdOutput = @(& sc.exe sdshow NTP 2>$null)
    if ($sdOutput.Count -eq 0) {
        Write-WarnMsg "Could not read NTP service security descriptor (sc sdshow)."
        return $false
    }

    $sddl = [string]($sdOutput | Where-Object { $_ -match '^D:' } | Select-Object -First 1)
    if ([string]::IsNullOrWhiteSpace($sddl)) {
        Write-WarnMsg "NTP service SDDL was not returned by sc sdshow."
        return $false
    }

    if ($sddl.Contains($desiredAce)) {
        return $true
    }

    $newSddl = ""
    $saclIndex = $sddl.IndexOf("S:")
    if ($saclIndex -gt 0) {
        $newSddl = $sddl.Substring(0, $saclIndex) + $desiredAce + $sddl.Substring($saclIndex)
    }
    else {
        $newSddl = $sddl + $desiredAce
    }

    $setOutput = @(& sc.exe sdset NTP $newSddl 2>&1)
    if ($LASTEXITCODE -ne 0) {
        Write-WarnMsg ("Failed to grant Users service control rights on NTP. {0}" -f (($setOutput | Out-String).Trim()))
        return $false
    }

    if ($sddl.Contains($legacyAce)) {
        Write-Ok "Upgraded NTP service ACL from legacy minimal rights to GUI-compatible standard-user rights."
    }
    else {
        Write-Ok "Granted GUI-compatible standard-user rights for NTP service control."
    }
    return $true
}

function Grant-UsersNtpFileAccess {
    param(
        [string]$InstallRoot,
        [string]$ConfigDir,
        [string]$LogsDir,
        [string]$ConfigFile
    )

    $usersSid = '*S-1-5-32-545'
    $allOk = $true

    foreach ($target in @($ConfigDir, $LogsDir)) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        if (-not (Test-Path -LiteralPath $target)) {
            New-Item -ItemType Directory -Path $target -Force | Out-Null
        }
        $out = @(& icacls.exe $target /grant "${usersSid}:(OI)(CI)(M)" /T /C 2>&1)
        if ($LASTEXITCODE -ne 0) {
            $allOk = $false
            Write-WarnMsg ("Failed to grant modify rights on {0}. {1}" -f $target, (($out | Out-String).Trim()))
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($ConfigFile) -and (Test-Path -LiteralPath $ConfigFile)) {
        $out = @(& icacls.exe $ConfigFile /grant "${usersSid}:(M)" /C 2>&1)
        if ($LASTEXITCODE -ne 0) {
            $allOk = $false
            Write-WarnMsg ("Failed to grant modify rights on {0}. {1}" -f $ConfigFile, (($out | Out-String).Trim()))
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($InstallRoot) -and (Test-Path -LiteralPath $InstallRoot)) {
        $scriptFiles = @(Get-ChildItem -LiteralPath $InstallRoot -Recurse -File -Include *.cmd, *.bat -ErrorAction SilentlyContinue)
        foreach ($f in $scriptFiles) {
            $out = @(& icacls.exe $f.FullName /grant "${usersSid}:(RX)" /C 2>&1)
            if ($LASTEXITCODE -ne 0) {
                $allOk = $false
                Write-WarnMsg ("Failed to grant execute rights on script {0}. {1}" -f $f.FullName, (($out | Out-String).Trim()))
            }
        }
    }

    if ($allOk) {
        Write-Ok "Granted standard-user access for ntp.conf/log files and script execution."
    }

    return $allOk
}

function Configure-StandaloneUserNtpAccess {
    param(
        [string]$InstallRoot,
        [string]$CurrentNtpConfPath
    )

    $programDataBase = if (-not [string]::IsNullOrWhiteSpace($env:ProgramData)) {
        $env:ProgramData
    }
    else {
        Join-Path $env:SystemDrive "ProgramData"
    }

    $targetRoot = Join-Path $programDataBase "NTP"
    $targetEtc = Join-Path $targetRoot "etc"
    $targetLogs = Join-Path $targetRoot "logs"
    $targetConf = Join-Path $targetEtc "ntp.conf"

    if (-not (Test-Path -LiteralPath $targetEtc)) {
        New-Item -ItemType Directory -Path $targetEtc -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $targetLogs)) {
        New-Item -ItemType Directory -Path $targetLogs -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $targetConf) -and (Test-Path -LiteralPath $CurrentNtpConfPath)) {
        Copy-Item -LiteralPath $CurrentNtpConfPath -Destination $targetConf -Force
        Write-Info ("Copied existing ntp.conf to user-writable path: {0}" -f $targetConf)
    }

    $serviceConfigUpdated = Set-NtpServiceConfigPath -ConfigPath $targetConf
    $fileAccessGranted = Grant-UsersNtpFileAccess -InstallRoot $InstallRoot -ConfigDir $targetEtc -LogsDir $targetLogs -ConfigFile $targetConf
    $serviceControlGranted = Grant-UsersNtpServiceControl
    $allApplied = ($serviceConfigUpdated -and $fileAccessGranted -and $serviceControlGranted)

    if (-not $allApplied) {
        Write-WarnMsg "One or more standard-user access changes did not apply successfully."
    }

    return [PSCustomObject]@{
        NtpConfPath          = $targetConf
        StatsDir             = $targetLogs
        ServiceConfigUpdated = $serviceConfigUpdated
        FileAccessGranted    = $fileAccessGranted
        ServiceControlGranted= $serviceControlGranted
        AllApplied           = $allApplied
    }
}

function Get-JsonResourceWithFallback {
    param(
        [string]$LocalPath,
        [string]$RemoteUrl,
        [string]$ResourceLabel
    )

    if ($script:RemoteDownloadsDisabled) {
        Show-RemoteOverrideNotice
    }
    elseif (-not [string]::IsNullOrWhiteSpace($RemoteUrl)) {
        try {
            Write-Info ("Loading {0} from GitHub: {1}" -f $ResourceLabel, $RemoteUrl)
            $resp = Invoke-WebRequest -Uri $RemoteUrl -UseBasicParsing -TimeoutSec 20
            $contentText = Convert-ToTextFromResponseContent -Content $resp.Content

            $localDir = Split-Path -Parent $LocalPath
            if (-not [string]::IsNullOrWhiteSpace($localDir) -and -not (Test-Path -LiteralPath $localDir)) {
                New-Item -ItemType Directory -Path $localDir -Force | Out-Null
            }

            Set-Content -LiteralPath $LocalPath -Value $contentText -Encoding UTF8
            Write-Info ("Cached {0} locally: {1}" -f $ResourceLabel, $LocalPath)

            return ($contentText | ConvertFrom-Json)
        }
        catch {
            Write-WarnMsg ("Failed to load {0} from GitHub: {1}" -f $ResourceLabel, $_.Exception.Message)
            Write-WarnMsg ("Falling back to local file: {0}" -f $LocalPath)
        }
    }

    if (-not (Test-Path -LiteralPath $LocalPath)) {
        throw ("{0} not found remotely or locally. Local path: {1}" -f $ResourceLabel, $LocalPath)
    }

    return (Get-Content -Raw -LiteralPath $LocalPath | ConvertFrom-Json)
}

function Load-CountryConfigResource {
    param([string]$ResourcePath)
    return Get-JsonResourceWithFallback -LocalPath $ResourcePath -RemoteUrl $script:CountryConfigRemoteUrl -ResourceLabel "Country server config"
}

function Load-NtpPoolZonesResource {
    param([string]$ResourcePath)
    return Get-JsonResourceWithFallback -LocalPath $ResourcePath -RemoteUrl $script:PoolZonesRemoteUrl -ResourceLabel "NTP pool zones resource"
}

function Load-NationalUtcInventoryResource {
    param([string]$ResourcePath)
    return Get-JsonResourceWithFallback -LocalPath $ResourcePath -RemoteUrl $script:NationalUtcRemoteUrl -ResourceLabel "National UTC/NTP inventory resource"
}

function Get-CountryConfigEntry {
    param([string]$CountryCode, [string]$ConfigPath)

    $json = Load-CountryConfigResource -ResourcePath $ConfigPath
    $prop = @($json.PSObject.Properties | Where-Object { $_.Name -ieq $CountryCode } | Select-Object -First 1)
    if ($prop.Count -eq 0) {
        return $null
    }

    return $prop[0].Value
}

function Get-CountryServers {
    param([string]$CountryCode, [string]$ConfigPath)

    $entry = Get-CountryConfigEntry -CountryCode $CountryCode -ConfigPath $ConfigPath
    if ($null -eq $entry) {
        throw "Country '$CountryCode' not found in $ConfigPath"
    }

    return @($entry.servers)
}

function Get-RegionPoolZoneForCountryCode {
    param([string]$CountryCode)

    $cc = $CountryCode.ToLowerInvariant()
    $europe = @('ad','al','am','at','ax','az','ba','be','bg','by','ch','cy','cz','de','dk','ee','es','fi','fo','fr','gb','ge','gg','gi','gr','hr','hu','ie','im','is','it','je','kg','kz','li','lt','lu','lv','mc','md','me','mk','mt','nl','no','pl','pt','ro','rs','ru','se','si','sj','sk','sm','tj','tm','tr','ua','uz','va')
    $northAmerica = @('ag','ai','aw','bb','bm','bs','bz','ca','cr','cu','dm','do','gd','gl','gt','hn','ht','jm','kn','ky','lc','mq','ms','mx','ni','pa','pm','pr','sv','sx','tc','tt','us','vc','vg','vi')
    $southAmerica = @('ar','bo','br','cl','co','ec','fk','gf','gy','pe','py','sr','uy','ve')
    $asia = @('ae','af','bd','bh','bn','bt','cn','hk','id','il','in','iq','ir','jo','jp','kh','kp','kr','kw','la','lb','lk','mm','mn','mo','mv','my','np','om','ph','pk','ps','qa','sa','sg','sy','th','tl','tw','vn','ye')
    $oceania = @('as','au','ck','fj','fm','gu','ki','mh','mp','nc','nf','nr','nu','nz','pg','pn','pw','sb','tk','to','tv','um','vu','wf','ws')
    $africa = @('ao','bf','bi','bj','bw','cd','cf','cg','ci','cm','cv','dj','dz','eg','eh','er','et','ga','gh','gm','gn','gq','gw','ke','km','lr','ls','ly','ma','mg','ml','mr','mu','mw','mz','na','ne','ng','re','rw','sc','sd','sh','sl','sn','so','ss','st','sz','td','tg','tn','tz','ug','yt','za','zm','zw')

    if ($europe -contains $cc) { return 'europe' }
    if ($northAmerica -contains $cc) { return 'north-america' }
    if ($southAmerica -contains $cc) { return 'south-america' }
    if ($asia -contains $cc) { return 'asia' }
    if ($oceania -contains $cc) { return 'oceania' }
    if ($africa -contains $cc) { return 'africa' }

    return ''
}

function Get-IndexedPoolServers {
    param(
        [string[]]$PoolHostnames,
        [int]$MaxCount
    )

    $indexed = @(
        $PoolHostnames |
            Where-Object { $_ -match '^([0-9]+)\..+\.pool\.ntp\.org$' } |
            Sort-Object { [int](($_ -split '\.')[0]) }
    )

    return @($indexed | Select-Object -First $MaxCount)
}

function Add-UniqueServers {
    param(
        [string[]]$Base,
        [string[]]$ToAdd
    )

    $result = @($Base)
    foreach ($item in $ToAdd) {
        if ($result -notcontains $item) {
            $result += $item
        }
    }
    return ,@($result)
}

function Get-ServerHostFromEntry {
    param([string]$Entry)

    if ([string]::IsNullOrWhiteSpace($Entry)) {
        return ""
    }

    return ($Entry.Trim() -split '\s+')[0].ToLowerInvariant()
}

function Select-ServersFromListInteractive {
    param(
        [string]$Header,
        [string]$Prompt,
        [string[]]$Servers,
        [int]$MaxCount
    )

    $choices = @($Servers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($choices.Count -eq 0) {
        return @()
    }

    while ($true) {
        Write-Host $Header -ForegroundColor Cyan
        for ($i = 0; $i -lt $choices.Count; $i++) {
            Write-Host ("  {0}) {1}" -f ($i + 1), $choices[$i])
        }

        $raw = Read-Host ($Prompt + " Enter server number(s) separated by comma (for example: 1,2), or press Enter for 0")
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return ,@()
        }

        $parts = @($raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        if ($parts.Count -eq 1 -and $parts[0] -eq "0") {
            return ,@()
        }

        if ($parts -contains "0") {
            Write-WarnMsg "Enter 0 by itself for no selection, or use server numbers such as 1,2."
            continue
        }

        if ($parts.Count -gt $MaxCount) {
            Write-WarnMsg ("Please select up to {0} servers." -f $MaxCount)
            continue
        }

        $selected = @()
        $valid = $true
        foreach ($p in $parts) {
            if ($p -notmatch '^\d+$') {
                $valid = $false
                break
            }

            $idx = [int]$p
            if ($idx -lt 1 -or $idx -gt $choices.Count) {
                $valid = $false
                break
            }

            $selected += $choices[$idx - 1]
        }

        if (-not $valid) {
            Write-WarnMsg "One or more selections were invalid."
            continue
        }

        return ,@($selected | Select-Object -Unique)
    }
}

function Convert-PrefixLengthToSubnetMask {
    param([int]$PrefixLength)

    if ($PrefixLength -lt 0 -or $PrefixLength -gt 32) {
        return ""
    }

    $mask = [uint32]0
    for ($i = 0; $i -lt $PrefixLength; $i++) {
        $mask = $mask -bor ([uint32]1 -shl (31 - $i))
    }

    $o1 = ($mask -shr 24) -band 255
    $o2 = ($mask -shr 16) -band 255
    $o3 = ($mask -shr 8) -band 255
    $o4 = $mask -band 255
    return ("{0}.{1}.{2}.{3}" -f $o1, $o2, $o3, $o4)
}

function Get-ActiveIpv4Configuration {
    $configs = @()
    try {
        $configs = @(
            Get-NetIPConfiguration -ErrorAction Stop |
                Where-Object { $null -ne $_.IPv4Address -and $null -ne $_.IPv4DefaultGateway }
        )
    }
    catch {
        return $null
    }

    if ($configs.Count -eq 0) {
        return $null
    }

    $selected = @($configs | Where-Object { $null -ne $_.NetAdapter -and $_.NetAdapter.Status -eq 'Up' } | Select-Object -First 1)
    if ($selected.Count -eq 0) {
        $selected = @($configs | Select-Object -First 1)
    }

    $cfg = $selected[0]
    $ipObj = @($cfg.IPv4Address | Select-Object -First 1)
    $ip = if ($ipObj.Count -gt 0) { [string]$ipObj[0].IPAddress } else { "" }
    $prefix = if ($ipObj.Count -gt 0) { [int]$ipObj[0].PrefixLength } else { 24 }
    $gateway = ""
    if ($null -ne $cfg.IPv4DefaultGateway) {
        $gateway = [string]$cfg.IPv4DefaultGateway.NextHop
    }

    $dns = @()
    if ($null -ne $cfg.DNSServer -and $null -ne $cfg.DNSServer.ServerAddresses) {
        $dns = @($cfg.DNSServer.ServerAddresses | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' })
    }

    $dhcpState = "Unknown"
    try {
        $iface = Get-NetIPInterface -InterfaceIndex $cfg.InterfaceIndex -AddressFamily IPv4 -ErrorAction Stop
        if ($iface.Dhcp -eq 'Enabled') { $dhcpState = "Enabled" }
        elseif ($iface.Dhcp -eq 'Disabled') { $dhcpState = "Disabled" }
    }
    catch {
        $dhcpState = "Unknown"
    }

    return [pscustomobject]@{
        AdapterAlias = [string]$cfg.InterfaceAlias
        IPAddress    = $ip
        PrefixLength = $prefix
        SubnetMask   = Convert-PrefixLengthToSubnetMask -PrefixLength $prefix
        Gateway      = $gateway
        DnsServers   = $dns
        DhcpState    = $dhcpState
    }
}

function Get-SuggestedStaticIpAddress {
    param(
        [string]$CurrentIp,
        [string]$Gateway
    )

    try {
        if ([string]::IsNullOrWhiteSpace($CurrentIp)) {
            return ""
        }

        $parts = @($CurrentIp -split '\.')
        if ($parts.Count -ne 4) {
            return $CurrentIp
        }

        $lastOctetText = [string]$parts[3]
        if ($lastOctetText -notmatch '^\d+$') {
            return $CurrentIp
        }

        $octet = [int]$lastOctetText
        if ($octet -lt 0 -or $octet -gt 255) {
            return $CurrentIp
        }

        foreach ($delta in @(20, 30, 10)) {
            $cand = $octet + [int]$delta
            if ($cand -lt 2 -or $cand -gt 254) {
                continue
            }

            $testIp = "{0}.{1}.{2}.{3}" -f $parts[0], $parts[1], $parts[2], $cand
            if ($testIp -ne $Gateway -and $testIp -ne $CurrentIp) {
                return $testIp
            }
        }

        return $CurrentIp
    }
    catch {
        return $CurrentIp
    }
}

function Show-StaticIpGuidance {
    try {
        Write-Host "Brief static IP guide for Windows 10/11:" -ForegroundColor Cyan

        $guidanceLines = @()

        $net = Get-ActiveIpv4Configuration
        if ($null -ne $net -and -not [string]::IsNullOrWhiteSpace([string]$net.IPAddress)) {
            $suggestedIp = Get-SuggestedStaticIpAddress -CurrentIp ([string]$net.IPAddress) -Gateway ([string]$net.Gateway)

            Write-Host ""
            Write-Host "Detected current active network settings:" -ForegroundColor Cyan
            Write-Host ("  Adapter: {0}" -f [string]$net.AdapterAlias)
            Write-Host ("  Current IPv4: {0}" -f [string]$net.IPAddress)
            Write-Host ("  Subnet mask: {0}" -f [string]$net.SubnetMask)
            Write-Host ("  Gateway: {0}" -f [string]$net.Gateway)

            $dnsList = @($net.DnsServers)
            if ($dnsList.Count -gt 0) {
                Write-Host ("  DNS: {0}" -f ($dnsList -join ', '))
            }
            Write-Host ("  DHCP: {0}" -f [string]$net.DhcpState)

            Write-Host ""
            Write-Host "Suggested static IPv4 values to enter:" -ForegroundColor Cyan
            Write-Host ("  IP address: {0}" -f $suggestedIp)
            Write-Host ("  Subnet mask: {0}" -f [string]$net.SubnetMask)
            Write-Host ("  Gateway: {0}" -f [string]$net.Gateway)

            $guidanceLines += ("  IP address: {0}" -f $suggestedIp)
            $guidanceLines += ("  Subnet mask: {0}" -f [string]$net.SubnetMask)
            $guidanceLines += ("  Gateway: {0}" -f [string]$net.Gateway)

            if ($dnsList.Count -gt 0) {
                Write-Host ("  Preferred DNS: {0}" -f [string]$dnsList[0])
                $guidanceLines += ("  Preferred DNS: {0}" -f [string]$dnsList[0])
                if ($dnsList.Count -gt 1) {
                    Write-Host ("  Alternate DNS: {0}" -f [string]$dnsList[1])
                    $guidanceLines += ("  Alternate DNS: {0}" -f [string]$dnsList[1])
                }
            }

            Write-WarnMsg "Important: use an IP address outside your router DHCP range (or reserve this IP in the router) to avoid conflicts."
        }
        else {
            Write-WarnMsg "Could not auto-detect current IPv4 settings."
            Write-Host "Use your current IPv4, subnet mask, gateway, and DNS as a starting point, then choose a nearby unused IP." -ForegroundColor Yellow
        }

        Write-Host ""
        Write-Host "How to enter on Windows 10/11:" -ForegroundColor Cyan
        Write-Host "  1) Open Settings > Network and Internet > Advanced network settings."
        Write-Host "  2) Open your active adapter (Ethernet/Wi-Fi) and choose View additional properties."
        Write-Host "  3) Under IP assignment, click Edit."
        Write-Host "  4) Select Manual, enable IPv4, and enter the suggested values."
        if ($guidanceLines.Count -gt 0) {
            Write-Host "     Suggested values for this PC:" -ForegroundColor Cyan
            foreach ($line in $guidanceLines) {
                Write-Host ("     {0}" -f $line.Trim())
            }
        }
        Write-Host "  5) Save and confirm internet access still works."
        Write-Host ""
        Write-Host "Simple step-by-step guide: https://support.microsoft.com/windows/change-tcp-ip-settings"
    }
    catch {
        Write-WarnMsg ("Static IP helper could not read network settings on this machine: {0}" -f $_.Exception.Message)
        Write-Host "Use your current IPv4, subnet mask, gateway, and DNS values from Windows network adapter details." -ForegroundColor Yellow
        Write-Host "Simple step-by-step guide: https://support.microsoft.com/windows/change-tcp-ip-settings" -ForegroundColor Yellow
    }
}

function Resolve-AuServersInteractive {
    param([string]$ConfigPath)

    $entry = Get-CountryConfigEntry -CountryCode "AU" -ConfigPath $ConfigPath
    if ($null -eq $entry) {
        throw "Country 'AU' not found in $ConfigPath"
    }

    $all = @($entry.servers)
    $nmiAll = @($all | Where-Object {
            $h = Get-ServerHostFromEntry -Entry $_
            $h -match '(^|\.)nmi\.gov\.au$'
        })

    $eduAll = @($all | Where-Object {
            $h = Get-ServerHostFromEntry -Entry $_
            $h -match '\.edu\.au$'
        })

    $poolNumbered = @($all | Where-Object {
            $h = Get-ServerHostFromEntry -Entry $_
            $h -match '^[0-3]\.au\.pool\.ntp\.org$'
        } | Select-Object -Unique)

    if ($poolNumbered.Count -lt 4) {
        $poolNumbered = @("0.au.pool.ntp.org iburst", "1.au.pool.ntp.org iburst", "2.au.pool.ntp.org iburst", "3.au.pool.ntp.org iburst")
    }

    $selected = @()
    $nmiUsed = $false

    Write-Info "Select the NTP servers to use."
    Write-Host "Where possible, choose servers that are near to you." -ForegroundColor Cyan
    Write-Host "Prefer servers in the same or neighbouring city/state/territory." -ForegroundColor Cyan

    Write-Host "" 
    Write-Host "Do you want to use NTP servers from the National Measurement Institute (NMI)?" -ForegroundColor Cyan
    Write-Host "These are the best servers and are traceable to UTC." -ForegroundColor Cyan
    Write-Host "However, you must register your computer and use a static IP address." -ForegroundColor Cyan
    Write-Host "Details will be provided." -ForegroundColor Cyan
    $useNmi = Read-YesNo -Prompt "Use NMI servers?" -DefaultYes $true
    if ($useNmi) {
        $primaryNmi = @($nmiAll | Where-Object { (Get-ServerHostFromEntry -Entry $_) -eq 'ntp.nmi.gov.au' } | Select-Object -First 1)
        if ($primaryNmi.Count -eq 0) {
            $primaryNmi = @("ntp.nmi.gov.au iburst")
        }

        $selected = Add-UniqueServers -Base $selected -ToAdd $primaryNmi
        $nmiUsed = $true

        $remainingNmi = @($nmiAll | Where-Object { (Get-ServerHostFromEntry -Entry $_) -ne 'ntp.nmi.gov.au' } | Select-Object -Unique)
        $chosenNmi = Select-ServersFromListInteractive -Header "Select up to 2 NMI servers.`nChoose ones nearest to your city/state/territory.`nAdditional servers will be added in the next steps." -Prompt "Select NMI servers." -Servers $remainingNmi -MaxCount 2
        $selected = Add-UniqueServers -Base $selected -ToAdd $chosenNmi
    }

    $chosenEdu = Select-ServersFromListInteractive -Header "Select up to 2 University servers.`nChoose ones nearest to your city/state/territory.`nEnter server number(s) separated by comma (for example: 1,2), or press Enter for 0.`nAdditional servers will be added in the next steps." -Prompt "Select University servers." -Servers $eduAll -MaxCount 2
    $selected = Add-UniqueServers -Base $selected -ToAdd $chosenEdu

    foreach ($pool in $poolNumbered) {
        if ($selected.Count -ge 5) {
            break
        }
        $selected = Add-UniqueServers -Base $selected -ToAdd @($pool)
    }

    if ($selected.Count -lt 5) {
        if ($selected.Count -eq $poolNumbered.Count -and $poolNumbered.Count -eq 4) {
            Write-Info "Using the 4 standard numbered AU pool servers. Add NMI or University servers if you want a fifth unique source."
        }
        else {
            Write-WarnMsg "Not enough numbered AU pool servers were available to reach 5 unique entries."
        }
    }

    Write-Step "AU server selection summary"
    for ($i = 0; $i -lt $selected.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $selected[$i])
    }

    if ($nmiUsed) {
        Write-WarnMsg "You must register to use NMI servers, and have a static IP address."
        Write-Host "To register, email time@measurement.gov.au" -ForegroundColor Yellow
        Write-Host "More details:" -ForegroundColor Yellow
        Write-Host "https://www.industry.gov.au/national-measurement-institute/nmi-services/physical-measurement-services/time-and-frequency-services" -ForegroundColor Yellow

        if (Read-YesNo -Prompt "Do you want more information about using a static IP address for your PC?" -DefaultYes $false) {
            Show-StaticIpGuidance
        }
    }

    return ,@($selected)
}

function Select-NationalServersInteractive {
    param([string[]]$Servers)

    $choices = @($Servers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($choices.Count -le 1) {
        return ,@($choices)
    }

    while ($true) {
        Write-Host "National servers found. Select up to TWO servers:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $choices.Count; $i++) {
            Write-Host ("  {0}) {1}" -f ($i + 1), $choices[$i])
        }

        $raw = Read-Host "Enter one or two numbers separated by comma (default: 1,2)"
        if ([string]::IsNullOrWhiteSpace($raw)) {
            return @($choices | Select-Object -First 2)
        }

        $parts = @($raw -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
        if ($parts.Count -eq 0 -or $parts.Count -gt 2) {
            Write-WarnMsg "Please select one or two server numbers."
            continue
        }

        $selected = @()
        $valid = $true
        foreach ($p in $parts) {
            if ($p -notmatch '^\d+$') { $valid = $false; break }
            $idx = [int]$p
            if ($idx -lt 1 -or $idx -gt $choices.Count) { $valid = $false; break }
            $selected += $choices[$idx - 1]
        }

        if (-not $valid) {
            Write-WarnMsg "One or more selections were invalid."
            continue
        }

        return @($selected | Select-Object -Unique)
    }
}

function Resolve-ServersForOtherCountry {
    param(
        [string]$CountryCode,
        [string]$PoolZonesPath,
        [switch]$UseRegionWhenCountryPoolSmallOnly
    )

    $cc = $CountryCode.ToLowerInvariant()
    $poolData = Load-NtpPoolZonesResource -ResourcePath $PoolZonesPath

    $countryEntry = @($poolData.countries | Where-Object { $_.zone -eq $cc } | Select-Object -First 1)
    $countryPoolHostnames = if ($countryEntry.Count -gt 0) { @($countryEntry[0].pool_hostnames) } else { @() }

    $listedActive = $null
    if ($countryEntry.Count -gt 0 -and $null -ne $countryEntry[0].counts -and $null -ne $countryEntry[0].counts.listed_active) {
        try { $listedActive = [int]$countryEntry[0].counts.listed_active } catch { $listedActive = $null }
    }

    $countryIndexed = @(Get-IndexedPoolServers -PoolHostnames $countryPoolHostnames -MaxCount 3)
    if (@($countryIndexed).Count -lt 3) {
        $countryIndexed = @("0.$cc.pool.ntp.org", "1.$cc.pool.ntp.org", "2.$cc.pool.ntp.org")
    }

    $servers = @($countryIndexed | ForEach-Object { "$_ iburst" })

    $region = $null
    if ($null -ne $poolData.country_to_region) {
        $regionProp = $poolData.country_to_region.PSObject.Properties[$cc]
        if ($null -ne $regionProp) {
            $region = [string]$regionProp.Value
        }
    }
    if ([string]::IsNullOrWhiteSpace([string]$region) -and $countryEntry.Count -gt 0) { $region = $countryEntry[0].region }
    if ([string]::IsNullOrWhiteSpace([string]$region)) { $region = Get-RegionPoolZoneForCountryCode -CountryCode $cc }

    $includeRegionServers = $true
    if ($UseRegionWhenCountryPoolSmallOnly -and $null -ne $listedActive -and $listedActive -ge 30) {
        $includeRegionServers = $false
        Write-Info "Country pool for '$cc' has $listedActive listed active servers; skipping regional pool fallback."
    }

    if ($includeRegionServers -and -not [string]::IsNullOrWhiteSpace([string]$region)) {
        $regionEntry = @($poolData.regions | Where-Object { $_.zone -eq $region } | Select-Object -First 1)
        $regionPoolHostnames = if ($regionEntry.Count -gt 0) { @($regionEntry[0].pool_hostnames) } else { @() }

        $regionIndexed = @(Get-IndexedPoolServers -PoolHostnames $regionPoolHostnames -MaxCount 2)
        if (@($regionIndexed).Count -lt 2) {
            $regionIndexed = @("0.$region.pool.ntp.org", "1.$region.pool.ntp.org")
        }

        $servers += @($regionIndexed | ForEach-Object { "$_ iburst" })
    }
    elseif ($includeRegionServers) {
        Write-WarnMsg "Could not determine continental region for '$cc'. Using global pool fallback for regional entries."
        $servers += @("0.pool.ntp.org iburst", "1.pool.ntp.org iburst")
    }

    return @($servers | Select-Object -Unique)
}

function Resolve-ServersForCountryViaNationalInventory {
    param(
        [string]$CountryCode,
        [string]$NationalUtcPath,
        [string]$PoolZonesPath
    )

    $selectedServers = @()
    $resource = Load-NationalUtcInventoryResource -ResourcePath $NationalUtcPath
    $entryResult = @($resource.entries | Where-Object { $_.country_code -ieq $CountryCode } | Select-Object -First 1)
    $entry = if ($entryResult.Count -gt 0) { $entryResult[0] } else { $null }

    if ($null -ne $entry) {
        $countryName = [string](Get-ComOptionalPropertyValue -Object $entry -Name "country_name" -Default $CountryCode.ToUpperInvariant())
        $authority = [string](Get-ComOptionalPropertyValue -Object $entry -Name "authority" -Default "")
        $status = [string](Get-ComOptionalPropertyValue -Object $entry -Name "status" -Default "")
        $usageNote = [string](Get-ComOptionalPropertyValue -Object $entry -Name "usage_note" -Default "")

        Write-Info ("National UTC/NTP entry for '{0}': {1}" -f $CountryCode.ToUpperInvariant(), $countryName)
        if (-not [string]::IsNullOrWhiteSpace($authority)) {
            Write-Host ("Authority: {0}" -f $authority)
        }
        if (-not [string]::IsNullOrWhiteSpace($status)) {
            Write-Host ("Status: {0}" -f $status)
        }
        if (-not [string]::IsNullOrWhiteSpace($usageNote)) {
            Write-Host ("Note: {0}" -f $usageNote)
        }

        $urls = @((Get-ComOptionalPropertyValue -Object $entry -Name "source_urls" -Default @()))
        if ($urls.Count -gt 0) {
            Write-Host "Source URLs:"
            foreach ($u in $urls) { Write-Host ("  - {0}" -f $u) }
        }

        $nationalHosts = @((Get-ComOptionalPropertyValue -Object $entry -Name "ntp_servers" -Default @()) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($nationalHosts.Count -gt 0) {
            $chosenNational = Select-NationalServersInteractive -Servers $nationalHosts
            $selectedServers = Add-UniqueServers -Base $selectedServers -ToAdd (@($chosenNational | ForEach-Object { "$_ iburst" }))
        }
        else {
            Write-WarnMsg "No national server hostnames listed for '$CountryCode'."
        }
    }
    else {
        Write-WarnMsg "No national UTC/NTP inventory entry found for '$CountryCode'."
    }

    $poolFallback = Resolve-ServersForOtherCountry -CountryCode $CountryCode -PoolZonesPath $PoolZonesPath -UseRegionWhenCountryPoolSmallOnly
    $selectedServers = Add-UniqueServers -Base $selectedServers -ToAdd $poolFallback
    return @($selectedServers)
}

function Resolve-ServersForCountry {
    param(
        [string]$CountryCode,
        [string]$ConfigPath,
        [string]$NationalUtcPath,
        [string]$PoolZonesPath,
        [string]$OtherCc
    )

    if ($CountryCode -eq "Other") {
        $resolvedCode = $OtherCc.ToUpperInvariant()
        $configEntry = Get-CountryConfigEntry -CountryCode $resolvedCode -ConfigPath $ConfigPath
        if ($null -ne $configEntry) {
            Write-Info "Country '$resolvedCode' is defined in config; using configured server list."
            return @($configEntry.servers)
        }

        Write-Info "Country '$resolvedCode' is not defined in config; using national UTC/NTP inventory plus pool fallback."
        return Resolve-ServersForCountryViaNationalInventory -CountryCode $resolvedCode -NationalUtcPath $NationalUtcPath -PoolZonesPath $PoolZonesPath
    }

    if ($CountryCode -eq "AU") {
        return Resolve-AuServersInteractive -ConfigPath $ConfigPath
    }

    return Get-CountryServers -CountryCode $CountryCode -ConfigPath $ConfigPath
}

function Show-InstalledServerList {
    param(
        [string[]]$Servers,
        [string]$Country,
        [string]$OtherCode
    )

    $label = if ($Country -eq "Other" -and -not [string]::IsNullOrWhiteSpace($OtherCode)) {
        ("Other/{0}" -f $OtherCode.ToUpperInvariant())
    }
    else {
        $Country
    }

    Write-Step ("Installed NTP servers ({0})" -f $label)

    $list = @($Servers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($list.Count -eq 0) {
        Write-WarnMsg "No servers were selected in this run."
        return
    }

    for ($i = 0; $i -lt $list.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $list[$i])
    }
}

function Update-NtpManagedSectionsFromTemplate {
    param(
        [string]$TemplatePath,
        [string[]]$Servers,
        [int]$Port,
        [int]$Mode,
        [string]$StatsFolder,
        [string]$DriftFile,
        [string]$OutputPath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) {
        $remoteTemplateUrl = $script:TemplateRemoteUrl
        if ($script:RemoteDownloadsDisabled) {
            Show-RemoteOverrideNotice
            throw "Template not found locally and remote downloads are disabled: $TemplatePath"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($remoteTemplateUrl)) {
            try {
                Write-Info "Template not found locally. Downloading from GitHub..."
                $templateDir = Split-Path -Parent $TemplatePath
                if (-not (Test-Path -LiteralPath $templateDir)) {
                    New-Item -ItemType Directory -Path $templateDir -Force | Out-Null
                }
                Invoke-WebRequest -Uri $remoteTemplateUrl -OutFile $TemplatePath -UseBasicParsing -TimeoutSec 20
                Write-Ok "Downloaded ntp.conf.template"
            }
            catch {
                throw "Template not found locally and could not be downloaded: $_"
            }
        }
        else {
            throw "Template not found: $TemplatePath"
        }
    }

    $template = Get-Content -Raw -LiteralPath $TemplatePath
    $serverLines = ($Servers | ForEach-Object { "server $_ minpoll 6 maxpoll 7" }) -join [Environment]::NewLine
    $rendered = $template.Replace("{{SERVER_LINES}}", $serverLines)
    $rendered = $rendered.Replace("{{COM_PORT}}", [string]$Port)
    $rendered = $rendered.Replace("{{MODE}}", [string]$Mode)
    $rendered = $rendered.Replace("{{STATSDIR}}", $StatsFolder)
    $rendered = $rendered.Replace("{{DRIFTFILE}}", $DriftFile)

    $headerStart = "# >>> NTP_GUIDED_MANAGED_HEADER_START"
    $headerEnd = "# <<< NTP_GUIDED_MANAGED_HEADER_END"
    $serverStart = "# >>> NTP_GUIDED_MANAGED_SERVERS_START"
    $serverEnd = "# <<< NTP_GUIDED_MANAGED_SERVERS_END"
    $loggingStart = "# >>> NTP_GUIDED_MANAGED_LOGGING_START"
    $loggingEnd = "# <<< NTP_GUIDED_MANAGED_LOGGING_END"

    $headerBlockPattern = "(?ms)^\s*" + [regex]::Escape($headerStart) + ".*?^\s*" + [regex]::Escape($headerEnd) + "\s*$"
    $serverBlockPattern = "(?ms)^\s*" + [regex]::Escape($serverStart) + ".*?^\s*" + [regex]::Escape($serverEnd) + "\s*$"
    $loggingBlockPattern = "(?ms)^\s*" + [regex]::Escape($loggingStart) + ".*?^\s*" + [regex]::Escape($loggingEnd) + "\s*$"

    $headerMatch = [regex]::Match($rendered, $headerBlockPattern)
    $serverMatch = [regex]::Match($rendered, $serverBlockPattern)
    $loggingMatch = [regex]::Match($rendered, $loggingBlockPattern)
    if (-not $headerMatch.Success -or -not $serverMatch.Success -or -not $loggingMatch.Success) {
        throw "Template must include managed section markers for HEADER, SERVERS and LOGGING."
    }

    $managedHeaderBlock = $headerMatch.Value.TrimEnd()
    $managedServersBlock = $serverMatch.Value.TrimEnd()
    $managedLoggingBlock = $loggingMatch.Value.TrimEnd()

    $outDir = Split-Path -Parent $OutputPath
    if (-not (Test-Path -LiteralPath $outDir)) {
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
    }

    $contentToWrite = $rendered
    if (Test-Path -LiteralPath $OutputPath) {
        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backup = "$OutputPath.bak_$stamp"
        Copy-Item -LiteralPath $OutputPath -Destination $backup -Force
        Write-Info "Backup created: $backup"

        $existing = Get-Content -Raw -LiteralPath $OutputPath
        $preserved = [regex]::Replace($existing, $headerBlockPattern + "\r?\n?", "")
        $preserved = [regex]::Replace($preserved, $serverBlockPattern + "\r?\n?", "")
        $preserved = [regex]::Replace($preserved, $loggingBlockPattern + "\r?\n?", "")
        $preserved = [regex]::Replace($preserved, '(?ms)^\s*# Internet NTP servers \(country-specific\)\r?\n(?:.*\r?\n)*?(?=^\s*# GPS PPS / NMEA source from serial COM port)', '')
        $preserved = [regex]::Replace($preserved, '(?m)^\s*# Enable monitoring logs\r?\n', '')
        $preserved = [regex]::Replace($preserved, '(?ms)^\s*enable stats\r?\nstatsdir "[^"]*"\r?\nstatistics loopstats\r?\nstatistics peerstats\r?\n?', '')
        $preserved = $preserved.TrimEnd()

        if ([string]::IsNullOrWhiteSpace($preserved)) {
            $contentToWrite = ($managedHeaderBlock + [Environment]::NewLine + [Environment]::NewLine + $managedServersBlock + [Environment]::NewLine + [Environment]::NewLine + $managedLoggingBlock)
        }
        else {
            $contentToWrite = ($managedHeaderBlock + [Environment]::NewLine + [Environment]::NewLine + $preserved + [Environment]::NewLine + [Environment]::NewLine + $managedServersBlock + [Environment]::NewLine + [Environment]::NewLine + $managedLoggingBlock)
        }

        Write-Info "Preserved existing ntp.conf settings outside managed sections."
    }

    Set-Content -LiteralPath $OutputPath -Value $contentToWrite -Encoding ASCII
    Write-Ok "Updated ntp.conf managed sections: $OutputPath"
}

function Read-CountrySelection {
    while ($true) {
        Write-Host "Select country profile for NTP servers:" -ForegroundColor Cyan
        Write-Host "  1) NZ"
        Write-Host "  2) AU"
        Write-Host "  3) US"
        Write-Host "  4) Other (enter 2-letter country code, for example FR, DE, JP)"

        $choice = Read-Host "Enter 1-4"
        if ($choice -eq "1") { return @{ Country = "NZ"; OtherCode = "" } }
        if ($choice -eq "2") { return @{ Country = "AU"; OtherCode = "" } }
        if ($choice -eq "3") { return @{ Country = "US"; OtherCode = "" } }
        if ($choice -eq "4") {
            $raw = Read-Host "Enter 2-letter country code (e.g. fr, de, jp)"
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $cc = $raw.Trim().ToLowerInvariant()
                if ($cc -match '^[a-z]{2}$') {
                    return @{ Country = "Other"; OtherCode = $cc }
                }
            }
            Write-WarnMsg "Please enter exactly two letters for the country code."
            continue
        }

        Write-WarnMsg "Invalid selection."
    }
}

function Read-GpsModeInteractive {
    param(
        [int]$CurrentMode = 17,
        [bool]$NmeaOnly = $false
    )

    Write-Info "GPS mode controls the serial baud rate and how serial GPS data is interpreted."

    if ($NmeaOnly) {
        Write-Info "For NMEA-only receivers, modes 1 (4800 baud) or 17 (9600 baud) are recommended."
        Write-Info "NMEA data works most reliably at these lower baud rates."
        if (Read-YesNo -Prompt "Use recommended GPS mode 17 (9600 baud)?" -DefaultYes $true) {
            return 17
        }
        Write-Host "Alternative recommended value: 1 (4800 baud)" -ForegroundColor Yellow
    }
    else {
        if (Read-YesNo -Prompt "Use recommended GPS mode 17 (9600 baud)?" -DefaultYes $true) {
            return 17
        }
    }

    Write-Host "All mode values (4800 or 9600 baud recommended for NMEA): 1, 17, 33, 49, 65, 81" -ForegroundColor Yellow
    Write-Host "Corresponding baud rates: 4800, 9600, 19200, 38400, 57600, 115200" -ForegroundColor Yellow
    Write-Host "If unsure, use 17." -ForegroundColor Yellow

    while ($true) {
        $modeRaw = Read-Host "Enter GPS mode value"
        if ($modeRaw -match '^(1|17|33|49|65|81)$') {
            return [int]$modeRaw
        }

        Write-WarnMsg "Unsupported GPS mode. Allowed values: 1, 17, 33, 49, 65, 81."
    }
}

function Get-ComOptionalPropertyValue {
    param(
        [object]$Object,
        [string]$Name,
        [object]$Default = ""
    )

    if ($null -eq $Object) {
        return $Default
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop) {
        return $Default
    }

    return $prop.Value
}

function Get-ComPortSnapshot {
    $hasGetPnpDevice = $null -ne (Get-Command -Name Get-PnpDevice -ErrorAction SilentlyContinue)
    $hasGetPnpDeviceProperty = $null -ne (Get-Command -Name Get-PnpDeviceProperty -ErrorAction SilentlyContinue)

    $ports = @()
    if ($hasGetPnpDevice) {
        try {
            $ports = @(Get-PnpDevice -Class Ports -PresentOnly -ErrorAction Stop)
        }
        catch {
            # Get-PnpDevice not available on this system; Win32_SerialPort fallback will be used silently
            $ports = @()
        }
    }

    $serialByInstanceId = @{}
    $serialRecords = @(Get-CimInstance Win32_SerialPort -ErrorAction SilentlyContinue)

    foreach ($serial in $serialRecords) {
        if (-not [string]::IsNullOrWhiteSpace($serial.PNPDeviceID)) {
            $serialByInstanceId[$serial.PNPDeviceID] = $serial
        }
    }

    if ($ports.Count -eq 0) {
        foreach ($serial in $serialRecords) {
            $friendly = [string](Get-ComOptionalPropertyValue -Object $serial -Name "Name")
            $comPort = [string](Get-ComOptionalPropertyValue -Object $serial -Name "DeviceID")
            if ([string]::IsNullOrWhiteSpace($comPort) -and $friendly -match '\((COM\d+)\)') {
                $comPort = $matches[1]
            }

            $ports += [pscustomobject]@{
                InstanceId   = [string](Get-ComOptionalPropertyValue -Object $serial -Name "PNPDeviceID")
                FriendlyName = $friendly
                Status       = [string](Get-ComOptionalPropertyValue -Object $serial -Name "Status")
                Manufacturer = [string](Get-ComOptionalPropertyValue -Object $serial -Name "Manufacturer")
                Description  = [string](Get-ComOptionalPropertyValue -Object $serial -Name "Description")
                Service      = [string](Get-ComOptionalPropertyValue -Object $serial -Name "Service")
                COMPort      = $comPort
                HardwareIds  = @()
            }
        }

        return @($ports | Sort-Object COMPort, FriendlyName)
    }

    $result = @()
    foreach ($port in $ports) {
        $friendly = [string]$port.FriendlyName
        $comPort = ""
        if ($friendly -match '\((COM\d+)\)') {
            $comPort = $matches[1]
        }

        $manufacturer = ""
        $hardwareIds = @()

        if ($hasGetPnpDeviceProperty) {
            try {
                $manufacturer = [string](Get-PnpDeviceProperty -InstanceId $port.InstanceId -KeyName 'DEVPKEY_Device_Manufacturer' -ErrorAction SilentlyContinue).Data
            }
            catch {
                $manufacturer = ""
            }

            try {
                $hardwareIds = @((Get-PnpDeviceProperty -InstanceId $port.InstanceId -KeyName 'DEVPKEY_Device_HardwareIds' -ErrorAction SilentlyContinue).Data)
            }
            catch {
                $hardwareIds = @()
            }
        }

        $serial = $null
        if ($serialByInstanceId.ContainsKey($port.InstanceId)) {
            $serial = $serialByInstanceId[$port.InstanceId]
        }

        $result += [pscustomobject]@{
            InstanceId   = [string]$port.InstanceId
            COMPort      = $comPort
            FriendlyName = $friendly
            Status       = [string]$port.Status
            Manufacturer = if ([string]::IsNullOrWhiteSpace($manufacturer)) { [string](Get-ComOptionalPropertyValue -Object $serial -Name "Manufacturer") } else { $manufacturer }
            Description  = if ($null -ne $serial) { [string](Get-ComOptionalPropertyValue -Object $serial -Name "Description") } else { "" }
            Service      = if ($null -ne $serial) { [string](Get-ComOptionalPropertyValue -Object $serial -Name "Service") } else { "" }
            HardwareIds  = $hardwareIds
        }
    }

    return @($result | Sort-Object COMPort, FriendlyName)
}

function Show-ComPorts {
    param(
        [array]$Snapshot,
        [string]$Title
    )

    Write-Info $Title
    $snapshotItems = @($Snapshot)
    if ($snapshotItems.Count -eq 0) {
        Write-Host "  (No active COM/Ports devices found.)"
        return
    }

    foreach ($d in $snapshotItems) {
        $comValue = [string](Get-ComOptionalPropertyValue -Object $d -Name "COMPort")
        $comText = "(none)"
        if (-not [string]::IsNullOrWhiteSpace($comValue)) {
            $comText = $comValue
        }

        Write-Host "------------------------------------------------------------"
        Write-Host ("COM Port      : {0}" -f $comText)
        Write-Host ("Friendly Name : {0}" -f [string](Get-ComOptionalPropertyValue -Object $d -Name "FriendlyName"))
        Write-Host ("Status        : {0}" -f [string](Get-ComOptionalPropertyValue -Object $d -Name "Status"))
        $manufacturerValue = [string](Get-ComOptionalPropertyValue -Object $d -Name "Manufacturer")
        if (-not [string]::IsNullOrWhiteSpace($manufacturerValue)) {
            Write-Host ("Manufacturer  : {0}" -f $manufacturerValue)
        }
        $descriptionValue = [string](Get-ComOptionalPropertyValue -Object $d -Name "Description")
        if (-not [string]::IsNullOrWhiteSpace($descriptionValue)) {
            Write-Host ("Description   : {0}" -f $descriptionValue)
        }
        $serviceValue = [string](Get-ComOptionalPropertyValue -Object $d -Name "Service")
        if (-not [string]::IsNullOrWhiteSpace($serviceValue)) {
            Write-Host ("Service       : {0}" -f $serviceValue)
        }
        Write-Host ("Instance ID   : {0}" -f [string](Get-ComOptionalPropertyValue -Object $d -Name "InstanceId"))
        $hardwareIds = @(Get-ComOptionalPropertyValue -Object $d -Name "HardwareIds" -Default @())
        if ($hardwareIds.Count -gt 0) {
            Write-Host ("Hardware IDs  : {0}" -f ($hardwareIds -join '; '))
        }
    }
    Write-Host "------------------------------------------------------------"
}

function Select-DetectedComDevice {
    param([array]$Candidates)

    $candidateItems = @($Candidates)
    if ($candidateItems.Count -eq 1) {
        return $candidateItems[0]
    }

    while ($true) {
        Write-Info "Multiple candidate COM devices found. Select the GPS/PPS device:"
        for ($i = 0; $i -lt $candidateItems.Count; $i++) {
            $c = $candidateItems[$i]
            $comValue = [string](Get-ComOptionalPropertyValue -Object $c -Name "COMPort")
            $comText = if ([string]::IsNullOrWhiteSpace($comValue)) { "(no COM tag)" } else { $comValue }
            Write-Host ("  {0}) {1}  [{2}]" -f ($i + 1), $c.FriendlyName, $comText)
        }

        $raw = Read-Host "Enter a number"
        if ($raw -match '^\d+$') {
            $n = [int]$raw
            if ($n -ge 1 -and $n -le $candidateItems.Count) {
                return $candidateItems[$n - 1]
            }
        }

        Write-WarnMsg "Invalid selection."
    }
}

function Install-FtdiDriverInteractive {
    param(
        [string]$LocalResourcePath,
        [string]$RemoteUrl,
        [string]$DownloadDir
    )

    if (Read-YesNo -Prompt "Have you already installed the FTDI USB serial driver for the GPS PPS device?" -DefaultYes $false) {
        Write-Ok "FTDI driver already installed. Skipping."
        return
    }

    Write-Info "Preparing FTDI USB serial driver installer (CDM212364_Setup.exe)..."

    $installerPath = $LocalResourcePath
    if (-not (Test-Path -LiteralPath $installerPath)) {
        $installerPath = Join-Path $DownloadDir "CDM212364_Setup.exe"
        Invoke-InstallerDownload -Url $RemoteUrl -OutputPath $installerPath -Label "FTDI USB serial driver"
    }
    else {
        Write-Info "Using local FTDI driver installer: $installerPath"
    }

    try {
        Install-Exe -InstallerPath $installerPath -Arguments @() -Label "FTDI USB serial driver"
        Write-Ok "FTDI driver installation completed."
    }
    catch {
        Write-WarnMsg ("FTDI driver installer finished with a warning: {0}" -f $_.Exception.Message)
        Write-WarnMsg "If the driver appears installed in Device Manager, you can continue."
    }

    Write-Host "Please plug in the GPS PPS device now to verify the driver has loaded correctly." -ForegroundColor Yellow
    [void](Read-Host "Press Enter when the GPS PPS device is connected")
}

function Find-GpsComPortInteractive {
    Write-Step "COM detection"
    Write-Host "Disconnect the GPS receiver/GPS PPS device before starting." -ForegroundColor Yellow

    while (-not (Read-YesNo -Prompt "Is the GPS/PPS device currently disconnected?" -DefaultYes $true)) {
        Write-WarnMsg "Please disconnect the device before continuing."
    }

    $baseline = Get-ComPortSnapshot
    Show-ComPorts -Snapshot $baseline -Title "Baseline active COM/Ports devices (device disconnected):"

    Write-Host ""
    Write-Host "Now plug in the GPS receiver or GPS PPS device." -ForegroundColor Yellow
    while (-not (Read-YesNo -Prompt "Has the GPS/PPS device been connected now?" -DefaultYes $true)) {
        Write-WarnMsg "Connect the device, then answer 'y' to continue."
    }

    Start-Sleep -Seconds 2
    $after = Get-ComPortSnapshot
    Show-ComPorts -Snapshot $after -Title "Active COM/Ports devices after connecting GPS/PPS device:"

    $newDevices = @()
    $baselineIds = @($baseline | ForEach-Object { $_.InstanceId })
    foreach ($d in $after) {
        if ($baselineIds -notcontains $d.InstanceId) {
            $newDevices += $d
        }
    }

    if ($newDevices.Count -eq 0) {
        Write-WarnMsg "No newly detected COM/Ports device found."
        Write-WarnMsg "Select the GPS/PPS device manually from current list."
        if ($after.Count -eq 0) {
            throw "No active COM/Ports devices available to select."
        }
        $newDevices = $after
    }

    $selected = Select-DetectedComDevice -Candidates $newDevices
    $selectedComText = "(not detected)"
    if (-not [string]::IsNullOrWhiteSpace($selected.COMPort)) {
        $selectedComText = $selected.COMPort
    }

    Write-Host ""
    Write-Info "Selected device candidate:"
    Write-Host ("  Friendly Name: {0}" -f $selected.FriendlyName)
    Write-Host ("  COM Port     : {0}" -f $selectedComText)
    Write-Host ("  Instance ID  : {0}" -f $selected.InstanceId)

    if (-not (Read-YesNo -Prompt "Is this the GPS/PPS device?" -DefaultYes $true)) {
        throw "User did not confirm selected device."
    }

    if ([string]::IsNullOrWhiteSpace($selected.COMPort)) {
        throw "Selected device does not expose a COMx name."
    }

    if ($selected.COMPort -notmatch '^COM(\d+)$') {
        throw "Unexpected COM port format: $($selected.COMPort)"
    }

    $comNumber = [int]$matches[1]
    Write-Ok ("Detected GPS/PPS COM port: {0} (number {1})" -f $selected.COMPort, $comNumber)
    return $comNumber
}

function Show-CountryInstallSummary {
    param(
        [string]$Country,
        [string]$OtherCode,
        [string]$CountryConfigPath,
        [string]$NationalUtcPath
    )

    try {
        $countryConfig = Load-CountryConfigResource -ResourcePath $CountryConfigPath
    }
    catch {
        Write-WarnMsg ("Country server config unavailable: {0}" -f $_.Exception.Message)
        return
    }

    if ($Country -ne "Other") {
        Write-Ok "Country profile '$Country' is curated in ntp-country-servers.json and expected to work well."
        return
    }

    $cc = $OtherCode.ToUpperInvariant()
    $prop = @($countryConfig.PSObject.Properties | Where-Object { $_.Name -ieq $cc } | Select-Object -First 1)
    if ($prop.Count -gt 0) {
        Write-Ok "Country profile '$cc' is curated in ntp-country-servers.json and expected to work well."
        return
    }

    Write-WarnMsg "Country '$cc' is not curated in ntp-country-servers.json. Pool-based servers should generally work."
    Write-Info "Country pool is usually preferred over region pool where available."

    try {
        $national = Load-NationalUtcInventoryResource -ResourcePath $NationalUtcPath
    }
    catch {
        Write-WarnMsg ("National UTC inventory unavailable: {0}" -f $_.Exception.Message)
        return
    }
    $entry = @($national.entries | Where-Object { $_.country_code -ieq $cc } | Select-Object -First 1)
    if ($entry.Count -eq 0) {
        Write-WarnMsg "No national UTC/NTP metadata found for '$cc'."
        return
    }

    $item = $entry[0]
    $countryName = [string](Get-ComOptionalPropertyValue -Object $item -Name "country_name" -Default $cc.ToUpperInvariant())
    $status = [string](Get-ComOptionalPropertyValue -Object $item -Name "status" -Default "")
    $authority = [string](Get-ComOptionalPropertyValue -Object $item -Name "authority" -Default "")
    $usageNote = [string](Get-ComOptionalPropertyValue -Object $item -Name "usage_note" -Default "")
    Write-WarnMsg "National standards servers have not been fully tested by this installer."
    Write-WarnMsg "These servers may be inaccessible, restricted, or require registration with the authority."
    Write-Host ("Country     : {0}" -f $countryName)
    if (-not [string]::IsNullOrWhiteSpace($status)) {
        Write-Host ("Status      : {0}" -f $status)
    }
    if (-not [string]::IsNullOrWhiteSpace($authority)) {
        Write-Host ("Authority   : {0}" -f $authority)
    }
    if (-not [string]::IsNullOrWhiteSpace($usageNote)) {
        Write-Host ("Usage Note  : {0}" -f $usageNote)
    }

    $urls = @((Get-ComOptionalPropertyValue -Object $item -Name "source_urls" -Default @()))
    if ($urls.Count -gt 0) {
        Write-Host "Authority / Source URL(s):"
        foreach ($u in $urls) {
            Write-Host (" - {0}" -f $u)
        }
    }
}

function Remove-GpsClockLines {
    param([string[]]$Lines)

    $cleaned = @()
    foreach ($line in @($Lines)) {
        if ($line -match '^\s*#\s*GPS serial source configured by install_ntp_timing_guided\.ps1\s*$') { continue }
        if ($line -match '^\s*#?\s*server\s+127\.127\.20\.') { continue }
        if ($line -match '^\s*#?\s*fudge\s+127\.127\.20\.') { continue }
        $cleaned += $line
    }

    return @($cleaned)
}

function Get-GpsClockEntries {
    param([string[]]$Lines)

    $entries = @()
    $headerComment = ""

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]

        if ($line -match '^\s*#\s*GPS serial source configured by install_ntp_timing_guided\.ps1\s*$') {
            $headerComment = $line
            continue
        }

        # Only parse active (uncommented) server 127.127.20.x lines
        if ($line -match '^\s*server\s+127\.127\.20\.(\d+)') {
            $comPort = [int]$matches[1]
            $serverLine = $line.Trim()
            $fudgeLine = ""

            for ($j = $i + 1; $j -lt [Math]::Min($i + 5, $Lines.Count); $j++) {
                if ($Lines[$j] -match ('^\s*fudge\s+127\.127\.20\.' + $comPort + '\b')) {
                    $fudgeLine = $Lines[$j].Trim()
                    $i = $j
                    break
                }
            }

            $isPPS = ($fudgeLine -match '\bflag1\s+1\b')

            $entries += [PSCustomObject]@{
                ComPort    = $comPort
                IsPPS      = $isPPS
                Header     = $headerComment
                ServerLine = $serverLine
                FudgeLine  = $fudgeLine
            }
            $headerComment = ""
            continue
        }

        if ($headerComment -ne "" -and $line -match '\S') {
            $headerComment = ""
        }
    }

    return $entries
}

function Select-GpsEntriesToKeep {
    param([PSCustomObject[]]$Entries)

    Write-Host ""
    Write-Host "GPS/PPS entries after this update:" -ForegroundColor Cyan
    Write-Host ""

    for ($i = 0; $i -lt $Entries.Count; $i++) {
        $e = $Entries[$i]
        $modeStr = if ($e.IsPPS) { "PPS+NMEA" } else { "NMEA only" }
        Write-Host ("  [{0}] COM{1}  ({2})" -f ($i + 1), $e.ComPort, $modeStr) -ForegroundColor White
        Write-Host ("       {0}" -f $e.ServerLine) -ForegroundColor DarkGray
        if (-not [string]::IsNullOrWhiteSpace($e.FudgeLine)) {
            Write-Host ("       {0}" -f $e.FudgeLine) -ForegroundColor DarkGray
        }
        Write-Host ""
    }

    $dupCom = @($Entries | Group-Object -Property ComPort | Where-Object { $_.Count -gt 1 })
    foreach ($dup in $dupCom) {
        Write-WarnMsg ("Conflict: COM{0} appears {1} times. Only one entry per COM port is valid." -f $dup.Name, $dup.Count)
    }

    $ppsCount = @($Entries | Where-Object { $_.IsPPS }).Count
    if ($ppsCount -gt 1) {
        Write-WarnMsg ("Multiple PPS entries detected ({0}). Only one PPS GPS receiver can be active at a time." -f $ppsCount)
    }

    while ($true) {
        Write-Host "Enter numbers to keep (e.g. 1,2), 'all', or 'none':" -ForegroundColor Cyan
        $reply = (Read-Host).Trim()

        if ($reply -ieq 'all') {
            $selected = @($Entries)
        }
        elseif ($reply -ieq 'none') {
            $selected = @()
        }
        else {
            $parts = @($reply -split '[,\s]+' | Where-Object { $_ -match '^\d+$' })
            if ($parts.Count -eq 0) {
                Write-WarnMsg "Enter numbers like '1,2', 'all', or 'none'."
                continue
            }
            $indices = @($parts | ForEach-Object { [int]$_ })
            $invalid = @($indices | Where-Object { $_ -lt 1 -or $_ -gt $Entries.Count })
            if ($invalid.Count -gt 0) {
                Write-WarnMsg ("Invalid numbers: {0}. Valid range is 1 to {1}." -f ($invalid -join ', '), $Entries.Count)
                continue
            }
            $selected = @($indices | ForEach-Object { $Entries[$_ - 1] })
        }

        $dupSel = @($selected | Group-Object -Property ComPort | Where-Object { $_.Count -gt 1 })
        if ($dupSel.Count -gt 0) {
            Write-WarnMsg ("Duplicate COM port in selection: COM{0}. Select only one entry per COM port." -f (($dupSel | ForEach-Object { $_.Name }) -join ', '))
            continue
        }

        $ppsSel = @($selected | Where-Object { $_.IsPPS })
        if ($ppsSel.Count -gt 1) {
            Write-WarnMsg "Multiple PPS entries selected. Only one PPS GPS receiver can be active at a time. Remove extra PPS entries."
            continue
        }

        return ,@($selected)
    }
}

function Write-GpsEntriesToConf {
    param(
        [string]$NtpConfPath,
        [PSCustomObject[]]$Entries
    )

    $lines = @(Get-Content -LiteralPath $NtpConfPath)
    $cleaned = @(Remove-GpsClockLines -Lines $lines)

    foreach ($e in $Entries) {
        $cleaned += ""
        $cleaned += "# GPS serial source configured by install_ntp_timing_guided.ps1"
        $cleaned += $e.ServerLine
        if (-not [string]::IsNullOrWhiteSpace($e.FudgeLine)) {
            $cleaned += $e.FudgeLine
        }
    }

    Set-Content -LiteralPath $NtpConfPath -Value $cleaned -Encoding ASCII
}

function Update-GpsLines {
    param(
        [string]$NtpConfPath,
        [int]$ComPort,
        [int]$GpsMode,
        [bool]$NmeaOnly
    )

    if (-not (Test-Path -LiteralPath $NtpConfPath)) {
        throw "ntp.conf not found at $NtpConfPath. Complete the NTP setup step first."
    }

    if ($NmeaOnly) {
        $newServerLine = "server 127.127.20.$ComPort mode $GpsMode minpoll 6 maxpoll 7 iburst"
        $newFudgeLine  = "fudge 127.127.20.$ComPort flag1 0 flag2 0 refid NMEA"
    }
    else {
        $newServerLine = "server 127.127.20.$ComPort mode $GpsMode minpoll 6 maxpoll 7 prefer"
        $newFudgeLine  = "fudge 127.127.20.$ComPort flag1 1 flag2 1 refid GPS"
    }

    $newEntry = [PSCustomObject]@{
        ComPort    = $ComPort
        IsPPS      = (-not $NmeaOnly)
        Header     = "# GPS serial source configured by install_ntp_timing_guided.ps1"
        ServerLine = $newServerLine
        FudgeLine  = $newFudgeLine
    }

    # Get existing entries excluding any for this COM port (new entry replaces them)
    $allLines = @(Get-Content -LiteralPath $NtpConfPath)
    $otherEntries = @(Get-GpsClockEntries -Lines $allLines | Where-Object { $_.ComPort -ne $ComPort })

    # Merge: other existing entries + new entry for this COM port
    $allEntries = @($otherEntries) + @($newEntry)

    if ($allEntries.Count -gt 1) {
        $modeStr = if ($NmeaOnly) { "NMEA-only" } else { "PPS+NMEA" }
        Write-Info ("New COM{0} ({1}) entry ready. Review all GPS entries below." -f $ComPort, $modeStr)
        $selected = @(Select-GpsEntriesToKeep -Entries $allEntries)
    }
    else {
        $selected = $allEntries
    }

    Write-GpsEntriesToConf -NtpConfPath $NtpConfPath -Entries $selected
    $modeStr = if ($NmeaOnly) { "NMEA-only" } else { "PPS+NMEA" }
    Write-Ok ("GPS entries updated in {0}: {1} kept ({2})" -f $NtpConfPath, @($selected).Count, $modeStr)
}

function Invoke-NtpQosStep {
    $policyNames = @(
        "NTP Outbound Priority",
        "NTP Inbound Priority"
    )

    # Remove any pre-existing policies with these names to avoid conflicts.
    foreach ($name in $policyNames) {
        try {
            $existing = Get-NetQosPolicy -Name $name -ErrorAction SilentlyContinue
            if ($null -ne $existing) {
                Remove-NetQosPolicy -Name $name -Confirm:$false -ErrorAction Stop
                Write-Info ("Removed existing QoS policy: {0}" -f $name)
            }
        }
        catch {
            Write-WarnMsg ("Could not remove existing QoS policy '{0}': {1}" -f $name, $_.Exception.Message)
        }
    }

    try {
        New-NetQosPolicy -Name "NTP Outbound Priority" `
            -IPProtocol UDP -IPDstPort 123 `
            -DSCPAction 46 -NetworkProfile All -ErrorAction Stop | Out-Null
        Write-Ok "Created QoS policy: NTP Outbound Priority (UDP dst 123, DSCP 46)"

        New-NetQosPolicy -Name "NTP Inbound Priority" `
            -IPProtocol UDP -IPSrcPort 123 `
            -DSCPAction 46 -NetworkProfile All -ErrorAction Stop | Out-Null
        Write-Ok "Created QoS policy: NTP Inbound Priority (UDP src 123, DSCP 46)"

        return $true
    }
    catch {
        Write-WarnMsg ("Failed to create QoS policy: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Invoke-WifiPowerSavingStep {
    # GUIDs are stable across Windows 7 / 8 / 10 / 11 (all editions).
    # Subgroup: Wireless Adapter Settings
    $wirelessSubgroup  = "19cbb8fa-5279-450e-9fac-8a3d5fedd0c1"
    # Setting: Power Saving Mode  (0 = Maximum Performance)
    $powerSavingSetting = "12bbebe6-58d6-4636-95bb-3217ef867c1a"

    try {
        $output = powercfg /setacvalueindex SCHEME_CURRENT $wirelessSubgroup $powerSavingSetting 0 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-WarnMsg ("powercfg reported an issue: {0}" -f ($output -join ' '))
            return $false
        }
        powercfg /setactive SCHEME_CURRENT 2>&1 | Out-Null
        Write-Ok "WiFi adapter set to Maximum Performance (power saving disabled)."
        Write-Info "This applies while the PC is plugged in to mains power."
        return $true
    }
    catch {
        Write-WarnMsg ("Failed to set WiFi power saving mode: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Set-PpsProviderRegistryValue {
    param([string]$DllPath)

    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Services\NTP"
    if (-not (Test-Path -LiteralPath $regPath)) {
        Write-WarnMsg "Registry path not found: $regPath"
        Write-WarnMsg "Install Meinberg NTP before enabling PPS provider."
        return $false
    }

    if (-not (Test-Path -LiteralPath $DllPath)) {
        Write-WarnMsg "PPS provider DLL not found: $DllPath"
        Write-WarnMsg "Install Meinberg NTP (or provide DLL path) before enabling PPS provider."
        return $false
    }

    $existing = $null
    try {
        $existing = (Get-ItemProperty -LiteralPath $regPath -Name "PPSProviders" -ErrorAction Stop).PPSProviders
    }
    catch { }

    if ($null -ne $existing) {
        $existingDisplay = if ($existing -is [array]) { $existing -join "; " } else { [string]$existing }
        Write-Host ""
        Write-Host "PPSProviders registry value already exists:" -ForegroundColor Cyan
        Write-Host ("  Current: {0}" -f $existingDisplay) -ForegroundColor White
        Write-Host ("  New:     {0}" -f $DllPath) -ForegroundColor White
        Write-Host ""

        if ($existingDisplay -eq $DllPath) {
            Write-Info "PPSProviders is already set to the correct value. No change needed."
            return $true
        }

        $overwrite = Read-YesNo -Prompt "Overwrite existing PPSProviders value with the new path?" -DefaultYes $false
        if (-not $overwrite) {
            Write-Info "Keeping existing PPSProviders value."
            return $true
        }
    }

    New-ItemProperty -Path $regPath -Name "PPSProviders" -PropertyType MultiString -Value $DllPath -Force | Out-Null
    Write-Ok ("Set registry PPSProviders = {0}" -f $DllPath)
    return $true
}

function Try-RestartNtpService {
    try {
        $svc = Get-Service -Name "NTP" -ErrorAction Stop
    }
    catch {
        Write-WarnMsg "NTP service is not installed or not registered. Restart skipped."
        return $false
    }

    try {
        Write-Info "Restarting NTP service..."
        Restart-Service -Name "NTP" -ErrorAction Stop
        Start-Sleep -Seconds 2
        $svc = Get-Service -Name "NTP" -ErrorAction Stop
        if ($svc.Status -ne "Running") {
            Write-WarnMsg ("NTP service status after restart: {0}" -f $svc.Status)
            return $false
        }

        Write-Ok "NTP service restarted successfully."
        return $true
    }
    catch {
        Write-WarnMsg ("NTP service restart failed: {0}" -f $_.Exception.Message)
        return $false
    }
}

function New-RestartNtpDesktopShortcut {
    param([string]$InstallRoot)

    $binDir     = Join-Path $InstallRoot "bin"
    $batPath    = Join-Path $binDir "restartntp.bat"
    $iconPath   = Join-Path $binDir "restart.ico"
    $shortcutPath = Join-Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) "Restart NTP.lnk"

    $create = Read-YesNo -Prompt "Do you want to add a Desktop shortcut for Restarting NTP (recommended)?" -DefaultYes $true
    if (-not $create) { return }

    if (-not (Test-Path -LiteralPath $batPath)) {
        Write-WarnMsg ("restartntp.bat not found at {0}. Desktop shortcut not created." -f $batPath)
        return
    }

    try {
        $shell   = New-Object -ComObject WScript.Shell
        $lnk     = $shell.CreateShortcut($shortcutPath)
        $lnk.TargetPath       = $batPath
        $lnk.WorkingDirectory = $binDir
        $lnk.Description      = "Restart the NTP service"
        if (Test-Path -LiteralPath $iconPath) {
            $lnk.IconLocation = "$iconPath,0"
        }
        $lnk.Save()
        Write-Ok ("Desktop shortcut created: {0}" -f $shortcutPath)
    }
    catch {
        Write-WarnMsg ("Could not create desktop shortcut: {0}" -f $_.Exception.Message)
    }
}

function Prompt-RestartIfNeeded {
    param([bool]$RestartNeeded)

    if (-not $RestartNeeded) {
        return $false
    }

    Write-WarnMsg "NTP configuration changed in this run."
    $doRestart = Read-YesNo -Prompt "Restart NTP service now to apply changes?" -DefaultYes $true
    if (-not $doRestart) {
        Write-WarnMsg "Please restart the NTP service manually before relying on this configuration."
        return $false
    }

    return (Try-RestartNtpService)
}

$scriptRoot = Split-Path -Parent $PSCommandPath
$projectRoot = $scriptRoot
$templatePath = Join-Path $projectRoot "config\ntp.conf.template"

$countryConfigPath = Join-Path $projectRoot "config\ntp-country-servers.json"
$nationalUtcPath = Join-Path $projectRoot "resources\national_utc_ntp_servers.json"
$poolZonesPath = Join-Path $projectRoot "resources\ntp_pool_zones.json"

$CountryConfigRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/config/ntp-country-servers.json"
$PoolZonesRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/resources/ntp_pool_zones.json"
$NationalUtcRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/resources/national_utc_ntp_servers.json"
$TemplateRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/config/ntp.conf.template"
$AutoInstallTemplateRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/config/install.auto.template.ini"

$script:CountryConfigRemoteUrl = $CountryConfigRemoteUrl
$script:PoolZonesRemoteUrl = $PoolZonesRemoteUrl
$script:NationalUtcRemoteUrl = $NationalUtcRemoteUrl
$script:TemplateRemoteUrl = $TemplateRemoteUrl
$script:AutoInstallTemplateRemoteUrl = $AutoInstallTemplateRemoteUrl

$meinbergInstallerUrl = "https://www.meinbergglobal.com/download/ntp/windows/ntp-4.2.8p18a2-win32-setup.exe"
$meinbergInstallerSha256 = "f933bc66ed987eb436f8345f6331de4ffad24e6ce5e5a6f5ce98109b7b29f164"
$meinbergInstallerSha256Url = "https://www.meinbergglobal.com/download/ntp/windows/ntp-4.2.8p18a2-win32-setup.exe.sha256sum"
$meinbergInstallerArgs = ""
$meinbergAutoIniTemplatePath = Join-Path $projectRoot "config\install.auto.template.ini"

$ntpMonitorInstallerUrl = "https://www.meinbergglobal.com/download/ntp/windows/time-server-monitor/ntp-time-server-monitor-104.exe"
$ntpMonitorInstallerArgs = ""

$ftdiDriverRemoteUrl = "https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/resources/CDM212364_Setup.exe"

$installRoot = Resolve-DefaultInstallRoot
$statsDir = Join-Path $installRoot "etc\"
$ntpConfPath = Join-Path $installRoot "etc\ntp.conf"
$driftFilePath = Join-Path $installRoot "etc\ntp.drift"
$ppsRegistryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\NTP"
$ppsDllPath = Join-Path $installRoot "bin\loopback-ppsapi-provider.dll"
$downloadDir = Resolve-WorkingTempDirectory -SubFolder "occultation-ntp-installer"

$logDir = Join-Path $projectRoot "logs"
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$logPath = Join-Path $logDir ("guided_ntp_installer_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$gpsConfigured = $false
$gpsNmeaOnly = $false
$gpsApplied = $false
$selectedComPort = 1
$selectedGpsMode = 17
$selectedCountry = "NZ"
$selectedOtherCode = ""
$transcriptStarted = $false
$restartRecommended = $false
$restartCompleted = $false

try {
    Start-Transcript -Path $logPath -Force | Out-Null
    $transcriptStarted = $true
}
catch {
    Write-WarnMsg ("Could not start transcript log at {0}: {1}" -f $logPath, $_.Exception.Message)
    Write-WarnMsg "Continuing without transcript logging for this run."
}

try {
    Assert-Admin

    Check-SupportFileAvailability `
        -TemplatePath $templatePath `
        -CountryConfigPath $countryConfigPath `
        -NationalUtcPath $nationalUtcPath `
        -PoolZonesPath $poolZonesPath `
        -AutoIniTemplatePath $meinbergAutoIniTemplatePath `
        -TemplateRemoteUrl $TemplateRemoteUrl `
        -CountryConfigRemoteUrl $CountryConfigRemoteUrl `
        -NationalUtcRemoteUrl $NationalUtcRemoteUrl `
        -PoolZonesRemoteUrl $PoolZonesRemoteUrl `
        -AutoIniTemplateRemoteUrl $AutoInstallTemplateRemoteUrl

    if ($script:RemoteDownloadsDisabled) {
        Write-Host "" 
        Write-Host "[OFFLINE MODE] Running with local files only (GitHub downloads disabled by user choice)." -ForegroundColor Yellow
        Write-Host "" 
    }

    Write-Step "Welcome to the NTP Installer"
    Write-Host "This guided installer can perform any or all of the following:" -ForegroundColor Cyan
    Write-Host " 1) Install Meinberg NTP and prepare logging"
    Write-Host " 2) Install NTP Time Server Monitor"
    Write-Host " 3) Configure optional GPS/PPS serial source"
    Write-Host " 4) Configure internet NTP servers by country"
    Write-Host " 5) Set Windows QoS priority for NTP traffic (UDP port 123)"
    Write-Host ""
    Write-Host "Estimated time: 10-20 minutes (depends on internet speed and installer prompts)." -ForegroundColor Cyan
    Write-Host "Internet access is needed for optional downloads." -ForegroundColor Cyan
    Write-Host "You can safely skip any step and run this installer again later." -ForegroundColor Cyan
    Write-Host "If Windows shows security/UAC prompts for trusted installers, choose Allow/Yes to continue." -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Each step is optional. At every step you can choose: Install, Skip, or Exit." -ForegroundColor Cyan
    Write-Host "You can also do all steps manually and follow project documentation when available." -ForegroundColor Yellow
    Write-Host ("Installer log: {0}" -f $logPath) -ForegroundColor DarkGray
    Write-Host ("Resolved NTP install root: {0}" -f $installRoot) -ForegroundColor DarkGray
    Write-Host ("Resolved working folder: {0}" -f $downloadDir) -ForegroundColor DarkGray

    if (-not (Read-YesNo -Prompt "Proceed with guided installer?" -DefaultYes $true)) {
        throw "Installer canceled by user at launch page."
    }

    Backup-NtpRegistry

    if (Confirm-Step -Title "Step 1: Install Meinberg NTP and prepare logging" -Details @(
            "Downloads and installs Meinberg NTP.",
            "Automatic install (recommended) uses config\\install.auto.template.ini and applies local paths.",
            "Guided install remains available for manual screen-by-screen setup.",
            "NTP internet server selection is done in a later step.",
            "Install using default. Do not add any predefined servers. They will be added later in Step 4.",
            ("Installer URL: {0}" -f $meinbergInstallerUrl),
            ("Automatic install template: {0}" -f $meinbergAutoIniTemplatePath),
            ("Install root: {0}" -f $installRoot),
            ("Config file: {0}" -f $ntpConfPath),
            ("Log folder: {0}" -f $statsDir),
            ("Advanced (automatic if PPS is enabled): registry value {0}\\PPSProviders" -f $ppsRegistryPath)
        )) {

        $meinbergInstallMode = Read-MeinbergInstallMode -IniTemplatePath $meinbergAutoIniTemplatePath
        if ($meinbergInstallMode -eq "Guided") {
            Write-WarnMsg "Install using default. Do not add any predefined servers. They will be added later in Step 4."
        }

        $meinbergInstallerPath = Join-Path $downloadDir "meinberg_installer.exe"
        Invoke-InstallerDownload -Url $meinbergInstallerUrl -OutputPath $meinbergInstallerPath -Label "Meinberg NTP"

        $effectiveMeinbergSha = $meinbergInstallerSha256
        if ([string]::IsNullOrWhiteSpace($effectiveMeinbergSha) -and -not [string]::IsNullOrWhiteSpace($meinbergInstallerSha256Url)) {
            $effectiveMeinbergSha = Get-ExpectedSha256FromUrl -ShaUrl $meinbergInstallerSha256Url
        }
        Assert-FileSha256 -Path $meinbergInstallerPath -ExpectedSha256 $effectiveMeinbergSha -Label "Meinberg NTP installer"

        if ($meinbergInstallMode -eq "Automatic") {
            $selectedUpgradeMode = Read-MeinbergAutomaticUpgradeMode -InstallRoot $installRoot -NtpConfPath $ntpConfPath
            $useConfigFileInAutomaticMode = Read-MeinbergAutomaticUseConfigChoice -UpgradeMode $selectedUpgradeMode
            $autoInstall = Prepare-MeinbergAutomaticInstallFiles -IniTemplatePath $meinbergAutoIniTemplatePath -TempDir $downloadDir -InstallRoot $installRoot -UpgradeMode $selectedUpgradeMode -ApplyUseConfigFile:$useConfigFileInAutomaticMode
            # Keep /USE_FILE unquoted to avoid passing embedded quote characters
            # to the NSIS option parser used by the Meinberg installer.
            $automaticArgs = @("/USE_FILE=$($autoInstall.IniPath)")
            Write-Host ""
            Write-Host "Automatic install settings:" -ForegroundColor Cyan
            Write-Host ("  InstallDir:   {0}" -f $installRoot)
            Write-Host ("  UpgradeMode:  {0}" -f $selectedUpgradeMode)
            Write-Host ("  Silent:       Yes")
            Write-Host ("  ServiceAccount: @SYSTEM")
            if ($useConfigFileInAutomaticMode) {
                Write-Host ("  UseConfigFile: {0}" -f $autoInstall.PlaceholderPath)
            }
            else {
                Write-Host ("  UseConfigFile: not set (preserve existing config if installer supports it)")
            }
            Write-Host ("  Installer log: {0}" -f $autoInstall.InstallerLogPath)
            Write-Host ""
            Write-Info ("Automatic install file: {0}" -f $autoInstall.IniPath)
            if ($useConfigFileInAutomaticMode) {
                Write-Info ("Automatic placeholder ntp.conf: {0}" -f $autoInstall.PlaceholderPath)
            }
            else {
                Write-Info "Automatic mode will not import a placeholder ntp.conf in this run."
            }
            Write-Info ("Automatic installer log: {0}" -f $autoInstall.InstallerLogPath)
            $automaticInstallRecovered = $false
            try {
                Install-Exe -InstallerPath $meinbergInstallerPath -Arguments $automaticArgs -Label "Meinberg NTP"
            }
            catch {
                Write-WarnMsg ("Automatic install failed: {0}" -f $_.Exception.Message)
                Write-WarnMsg ("Generated automatic INI: {0}" -f $autoInstall.IniPath)
                Write-WarnMsg ("Installer log path: {0}" -f $autoInstall.InstallerLogPath)

                $logIndicatesSuccess = Test-MeinbergAutomaticInstallSuccess -InstallerLogPath $autoInstall.InstallerLogPath
                if ($logIndicatesSuccess) {
                    Write-WarnMsg "Installer returned a non-zero code, but the Meinberg log reports successful completion."
                    Write-Ok "Continuing based on installer log success marker."
                    $automaticInstallRecovered = $true
                }

                if (Test-Path -LiteralPath $autoInstall.InstallerLogPath) {
                    Write-Info "Last lines from Meinberg automatic installer log:"
                    Get-Content -LiteralPath $autoInstall.InstallerLogPath -Tail 25 | ForEach-Object {
                        Write-Host ("  {0}" -f $_)
                    }
                }
                else {
                    Write-WarnMsg "Installer log file was not created by Meinberg installer."
                }

                if (-not $automaticInstallRecovered) {
                    if (Read-YesNo -Prompt "Retry Step 1 now using guided install mode?" -DefaultYes $true) {
                        Write-WarnMsg "Falling back to guided install mode for Step 1."
                        Write-WarnMsg "Install using default. Do not add any predefined servers. They will be added later in Step 4."
                        Install-Exe -InstallerPath $meinbergInstallerPath -Arguments @() -Label "Meinberg NTP"
                    }
                    else {
                        throw
                    }
                }
            }

            # Automatic mode: always apply recommended standard-user access layout.
            $layout = Configure-StandaloneUserNtpAccess -InstallRoot $installRoot -CurrentNtpConfPath $ntpConfPath
            $ntpConfPath = [string]$layout.NtpConfPath
            $statsDir = [string]$layout.StatsDir
            $driftFilePath = Join-Path (Split-Path -Parent $ntpConfPath) "ntp.drift"
            if ($layout.AllApplied) {
                Write-Ok "Applied standard-user access layout automatically for this install."
            }
            else {
                Write-WarnMsg "Standard-user access layout was attempted, but some required permission/service changes failed."
                Write-WarnMsg "Review warnings above to complete: ntp.conf edit rights, log write rights, and NTP service control rights."
            }
        }
        else {
            Install-Exe -InstallerPath $meinbergInstallerPath -Arguments @($meinbergInstallerArgs) -Label "Meinberg NTP"

            Write-WarnMsg "Recommended: move ntp.conf/logs to ProgramData and grant standard-user rights for scripts + NTP service control."
            if (Read-YesNo -Prompt "Apply recommended standard-user access changes now?" -DefaultYes $true) {
                $layout = Configure-StandaloneUserNtpAccess -InstallRoot $installRoot -CurrentNtpConfPath $ntpConfPath
                $ntpConfPath = [string]$layout.NtpConfPath
                $statsDir = [string]$layout.StatsDir
                $driftFilePath = Join-Path (Split-Path -Parent $ntpConfPath) "ntp.drift"
                if ($layout.AllApplied) {
                    Write-Ok "Applied standard-user access layout."
                }
                else {
                    Write-WarnMsg "Standard-user access layout was attempted, but some required permission/service changes failed."
                    Write-WarnMsg "Review warnings above to complete: ntp.conf edit rights, log write rights, and NTP service control rights."
                }
            }
            else {
                Write-WarnMsg "Keeping default Program Files layout. Standard users may be unable to edit ntp.conf, write logs, or control the NTP service."
            }
        }

        Set-LoggingConfig -NtpConfPath $ntpConfPath -StatsDir $statsDir
        $restartRecommended = $true
    }

    if (Confirm-Step -Title "Step 2: Install NTP Time Server Monitor" -Details @(
            "Downloads and installs NTP Time Server Monitor.",
            "Use this tool later to verify lock, offsets, and source selection.",
            ("Installer URL: {0}" -f $ntpMonitorInstallerUrl),
            "No official checksum URL is currently configured in this script."
        )) {

        $monitorInstallerPath = Join-Path $downloadDir "ntp_monitor_installer.exe"
        Invoke-InstallerDownload -Url $ntpMonitorInstallerUrl -OutputPath $monitorInstallerPath -Label "NTP Time Server Monitor"
        Install-Exe -InstallerPath $monitorInstallerPath -Arguments @($ntpMonitorInstallerArgs) -Label "NTP Time Server Monitor"
    }

    if (Confirm-Step -Title "Step 3: Optional GPS/PPS source setup" -Details @(
            "Optionally auto-detect COM port with built-in detection.",
            "Can configure either PPS+NMEA (GPS PPS) or NMEA-only receiver mode.",
            "Writes GPS server/fudge lines to ntp.conf.",
            ("ntp.conf target: {0}" -f $ntpConfPath),
            ("PPS provider DLL path (automatic if PPS mode is selected): {0}" -f $ppsDllPath),
            "No manual registry editing is required for normal setup.",
            ("Advanced detail: PPS mode updates {0}\\PPSProviders automatically" -f $ppsRegistryPath)
        )) {

        Write-WarnMsg "Installing GPS receivers has not been fully tested yet. You may need to manually alter the ntp.conf file and change some parameters to get it working."

        Write-Host "Select GPS mode:" -ForegroundColor Cyan
        Write-Host "  1) GPS PPS + NMEA (recommended when PPS available)"
        Write-Host "  2) GPS NMEA only (no PPS signal)"

        $gpsChoice = Read-Host "Enter 1 or 2 (default 1)"
        if ($gpsChoice -eq "2") {
            $gpsNmeaOnly = $true
            Write-Info "Selected mode: NMEA only"
        }
        else {
            $gpsNmeaOnly = $false
            Write-Info "Selected mode: PPS + NMEA"
        }

        if ($gpsNmeaOnly) {
            Write-Host ""
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host " NMEA-only GPS: prerequisites" -ForegroundColor Cyan
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "Before continuing, make sure:" -ForegroundColor White
            Write-Host "  1. Your GPS receiver is physically connected to a USB or serial port." -ForegroundColor White
            Write-Host "  2. Windows has installed a driver for it." -ForegroundColor White
            Write-Host "     - Most USB GPS receivers install automatically (Plug and Play)." -ForegroundColor White
            Write-Host "     - If yours came with a driver disc or OEM software, install that first." -ForegroundColor White
            Write-Host "  3. The receiver appears as a COM port in Device Manager." -ForegroundColor White
            Write-Host "     (Device Manager -> Ports (COM & LPT) -> look for your GPS device)" -ForegroundColor White
            Write-Host "  4. The baud rate your GPS outputs NMEA data on." -ForegroundColor White
            Write-Host "     Most receivers default to 4800 or 9600 baud." -ForegroundColor White
            Write-Host "     Check your receiver's manual or configuration software if unsure." -ForegroundColor White
            Write-Host "     The installer will let you set the matching baud rate in the next step." -ForegroundColor White
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host ""
            if (-not (Read-YesNo -Prompt "Is your GPS receiver connected, driver installed, and visible as a COM port?" -DefaultYes $false)) {
                Write-WarnMsg "Skipping GPS configuration. Re-run Step 3 once the receiver is connected and driver is installed."
                $gpsConfigured = $false
                $skipGpsSetup = $true
            }
            else {
                Write-Ok "GPS receiver confirmed ready."
                $skipGpsSetup = $false
            }
        }
        else {
            $skipGpsSetup = $false
        }

        if (-not $skipGpsSetup) {

        if (-not $gpsNmeaOnly) {
            $ftdiLocalPath = Join-Path $projectRoot "resources\CDM212364_Setup.exe"
            Install-FtdiDriverInteractive -LocalResourcePath $ftdiLocalPath -RemoteUrl $ftdiDriverRemoteUrl -DownloadDir $downloadDir
        }

        if (Read-YesNo -Prompt "Run built-in COM port detection now?" -DefaultYes $true) {
            $selectedComPort = Find-GpsComPortInteractive
            Write-Ok ("Using detected COM port: COM{0}" -f $selectedComPort)
        }
        else {
            while ($true) {
                $manualCom = Read-Host "Enter COM port number (example: 3 for COM3)"
                if ($manualCom -match '^\d+$') {
                    $comValue = [int]$manualCom
                    if ($comValue -ge 1 -and $comValue -le 256) {
                        $selectedComPort = $comValue
                        break
                    }
                }
                Write-WarnMsg "Enter a numeric COM port value between 1 and 256."
            }
        }

        $selectedGpsMode = Read-GpsModeInteractive -CurrentMode $selectedGpsMode -NmeaOnly:$gpsNmeaOnly

        if (-not (Test-Path -LiteralPath $ntpConfPath)) {
            Write-WarnMsg "ntp.conf not found yet. GPS lines will be applied after Step 4 completes."
        }
        else {
            Update-GpsLines -NtpConfPath $ntpConfPath -ComPort $selectedComPort -GpsMode $selectedGpsMode -NmeaOnly:$gpsNmeaOnly
            $gpsApplied = $true
            $restartRecommended = $true
        }

        if (-not $gpsNmeaOnly) {
            Write-Info "PPS mode selected: configuring PPSProviders registry value."
            $ppsConfigured = Set-PpsProviderRegistryValue -DllPath $ppsDllPath
            if (-not $ppsConfigured) {
                Write-WarnMsg "PPSProviders registry value was not set in this run."
            }
            else {
                $restartRecommended = $true
            }
        }

        if ($gpsNmeaOnly) {
            Write-Host ""
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host " GPS NMEA-only: USB delay adjustment required" -ForegroundColor Cyan
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host "GPS NMEA time delivered over USB has a propagation delay, typically 50-150 ms." -ForegroundColor Yellow
            Write-Host "This offset must be corrected in ntp.conf using the 'time2' fudge parameter." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "How to estimate and apply the correction:" -ForegroundColor Cyan
            Write-Host "  1. After NTP is running, open NTP Time Server Monitor and check the GPS source offset." -ForegroundColor White
            Write-Host "  2. Edit the fudge line in ntp.conf and set 'time2' to the negative of the observed offset." -ForegroundColor White
            Write-Host "  3. Restart NTP and re-check. Repeat until the GPS source offset is near zero." -ForegroundColor White
            Write-Host ""
            Write-Host "Example fudge line (replace X with your COM port number, adjust time2 value):" -ForegroundColor Cyan
            Write-Host ("  fudge 127.127.20.{0} time2 -0.100 refid GPS" -f $selectedComPort) -ForegroundColor White
            Write-Host ""
            Write-Host "If the offset shown in NTP Monitor is +120 ms, set time2 to -0.120 (seconds)." -ForegroundColor Yellow
            Write-Host "------------------------------------------------------------" -ForegroundColor Cyan
            Write-Host ""
        }

        $gpsConfigured = $true
        } # end if (-not $skipGpsSetup)
    }

    if ($gpsConfigured) {
        Write-Step "GPS reminder"
        Write-WarnMsg "GPS parameters may still require tuning for your hardware and serial settings."
        Write-Info "Test behavior in NTP Time Server Monitor and refer to project documentation when available."
    }

    if (Confirm-Step -Title "Step 4: Configure internet NTP servers by country" -Details @(
            "Uses guided installer built-in country logic.",
            "Country profiles in ntp-country-servers.json are curated.",
            "Other countries use pool logic and may include national UTC inventory data.",
            ("Template: {0}" -f $templatePath),
            ("Country config: {0}" -f $countryConfigPath),
            ("National UTC metadata: {0}" -f $nationalUtcPath),
            ("Pool zones metadata: {0}" -f $poolZonesPath),
            ("Target config file: {0}" -f $ntpConfPath)
        )) {

        $countryChoice = Read-CountrySelection
        $selectedCountry = [string]$countryChoice.Country
        $selectedOtherCode = [string]$countryChoice.OtherCode

        $servers = Resolve-ServersForCountry -CountryCode $selectedCountry -ConfigPath $countryConfigPath -NationalUtcPath $nationalUtcPath -PoolZonesPath $poolZonesPath -OtherCc $selectedOtherCode
        Update-NtpManagedSectionsFromTemplate -TemplatePath $templatePath -Servers $servers -Port $selectedComPort -Mode $selectedGpsMode -StatsFolder $statsDir -DriftFile $driftFilePath -OutputPath $ntpConfPath

        if ($gpsConfigured -and -not $gpsApplied) {
            Write-Info "Applying deferred GPS configuration to ntp.conf."
            Update-GpsLines -NtpConfPath $ntpConfPath -ComPort $selectedComPort -GpsMode $selectedGpsMode -NmeaOnly:$gpsNmeaOnly
            $gpsApplied = $true
        }

        Show-InstalledServerList -Servers $servers -Country $selectedCountry -OtherCode $selectedOtherCode
        $restartRecommended = $true

        Show-CountryInstallSummary -Country $selectedCountry -OtherCode $selectedOtherCode -CountryConfigPath $countryConfigPath -NationalUtcPath $nationalUtcPath
    }

    # If Step 4 was skipped but GPS was configured and ntp.conf already exists (e.g. Meinberg was
    # installed in a prior session), apply the deferred GPS lines now rather than silently losing them.
    if ($gpsConfigured -and -not $gpsApplied -and (Test-Path -LiteralPath $ntpConfPath)) {
        Write-Info "Step 4 was skipped but ntp.conf is present. Applying GPS configuration now."
        Update-GpsLines -NtpConfPath $ntpConfPath -ComPort $selectedComPort -GpsMode $selectedGpsMode -NmeaOnly:$gpsNmeaOnly
        $gpsApplied = $true
        $restartRecommended = $true
    }
    elseif ($gpsConfigured -and -not $gpsApplied) {
        Write-WarnMsg "GPS was configured but ntp.conf does not exist yet. Run Step 4 to create ntp.conf and apply the GPS settings."
    }

    if (Confirm-Step -Title "Step 5: Prioritise NTP traffic on this PC" -Details @(
            "Tells Windows to send NTP time-sync packets before other network traffic.",
            "This helps keep accurate time even when the connection is busy with downloads or streaming.",
            "Recommended for all installs."
        )) {
        Invoke-NtpQosStep | Out-Null
    }

    if (Confirm-Step -Title "Step 6: Set WiFi adapter to maximum performance" -Details @(
            "WiFi adapters can use a power-saving mode that lets the radio briefly sleep between packets.",
            "This can cause NTP time-sync packets to arrive late or unevenly, adding timing jitter of 10-50 ms.",
            "This step tells Windows to keep your WiFi radio fully awake at all times.",
            "Safe to apply on any PC - your internet connection and WiFi speed are not affected.",
            "You can skip this step if your PC uses a wired Ethernet connection instead of WiFi."
        )) {
        Invoke-WifiPowerSavingStep | Out-Null
    }

    if ($restartRecommended -and -not $restartCompleted) {
        $restartCompleted = Prompt-RestartIfNeeded -RestartNeeded $restartRecommended
    }

    New-RestartNtpDesktopShortcut -InstallRoot $installRoot

    Write-Step "Completed"
    Write-Ok "Guided installer finished."
    Write-Host "You can run this installer again and complete any skipped steps later." -ForegroundColor Cyan
    Write-Host "You can also install all components manually and use project documentation." -ForegroundColor Cyan
    Write-Host ("Log written to: {0}" -f $logPath) -ForegroundColor Cyan
}
catch {
    if ($_.Exception.Message -like "Installer exited by user at step:*") {
        Write-WarnMsg $_.Exception.Message
        if ($restartRecommended -and -not $restartCompleted) {
            Write-WarnMsg "You exited before final restart check."
            $restartCompleted = Prompt-RestartIfNeeded -RestartNeeded $restartRecommended
        }
        Write-Host ("Log written to: {0}" -f $logPath) -ForegroundColor Cyan
        return
    }

    if ($_.Exception.Message -eq "Installer canceled by user at launch page.") {
        Write-WarnMsg "Installer canceled by user at launch page."
        if ($restartRecommended -and -not $restartCompleted) {
            Write-WarnMsg "Configuration changed earlier in this run; restart may still be required."
            $restartCompleted = Prompt-RestartIfNeeded -RestartNeeded $restartRecommended
        }
        Write-Host ("Log written to: {0}" -f $logPath) -ForegroundColor Cyan
        return
    }

    Write-Host "" 
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "See installer log for full details:" -ForegroundColor Yellow
    Write-Host ("  {0}" -f $logPath) -ForegroundColor Yellow
    throw
}
finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch {}
    }

    Wait-BeforeCloseIfNeeded
}
