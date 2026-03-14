@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%install_ntp_timing_guided.ps1"
set "PS_SCRIPT_URL=https://raw.githubusercontent.com/labstercam/occultation-ntp-installer/main/install_ntp_timing_guided.ps1"
set "ELEVATED_FLAG=%~1"
set "OCNTP_REMOTE_DISABLED=0"

rem Ensure admin context first; if not elevated, relaunch this CMD via UAC once.
powershell.exe -NoProfile -Command "$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()); if ($p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { exit 0 } else { exit 1 }"
if not "%ERRORLEVEL%"=="0" (
  if /I "%ELEVATED_FLAG%"=="--elevated" (
    echo [ERROR] Could not obtain Administrator rights.
    echo Please right-click this file and choose "Run as administrator".
    pause
    exit /b 1
  )

  echo Requesting Administrator rights...
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -Verb RunAs -FilePath '%ComSpec%' -ArgumentList '/c """%~f0"" --elevated"'"
  exit /b %ERRORLEVEL%
)

echo [INFO] Downloading latest PowerShell script from GitHub...
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$ProgressPreference = 'SilentlyContinue'; Invoke-WebRequest -Uri '%PS_SCRIPT_URL%' -OutFile '%PS_SCRIPT%' -UseBasicParsing -ErrorAction Stop"
if not "%ERRORLEVEL%"=="0" (
  echo.
  echo [WARN] Could not download the latest PowerShell script from GitHub.
  echo [WARN] Internet may be unavailable, or GitHub may be unreachable.
  if not exist "%PS_SCRIPT%" (
    echo.
    echo [ERROR] No previously downloaded local script is available:
    echo   %PS_SCRIPT%
    echo.
    echo Please check internet access and try again.
    echo If needed, download the full installer package from:
    echo   https://github.com/labstercam/occultation-ntp-installer
    echo.
    pause
    exit /b 1
  )

  echo.
  echo A local previously downloaded installer script was found:
  echo   %PS_SCRIPT%
  choice /C YN /N /M "Continue with the local script version? [Y/N]: "
  if errorlevel 2 (
    echo.
    echo [INFO] Cancelled by user.
    exit /b 1
  )
  set "OCNTP_REMOTE_DISABLED=1"
  echo [INFO] Continuing with local script version.
  echo.
) else (
  if not exist "%PS_SCRIPT%" (
    echo.
    echo [ERROR] Download appeared to succeed but script is missing:
    echo   %PS_SCRIPT%
    echo.
    pause
    exit /b 1
  )
  echo [OK] Downloaded install_ntp_timing_guided.ps1
  echo.
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [ERROR] Guided installer exited with code %EXIT_CODE%.
  echo If it failed before opening, try running this CMD launcher as Administrator.
  pause
)

exit /b %EXIT_CODE%
