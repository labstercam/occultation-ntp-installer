@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%install_ntp_timing_guided.ps1"
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

echo [INFO] Using local PowerShell script...
if not exist "%PS_SCRIPT%" (
  echo.
  echo [ERROR] Local PowerShell script not found:
  echo   %PS_SCRIPT%
  echo.
  echo Please ensure the script exists in the installer directory.
  echo.
  pause
  exit /b 1
)

echo [INFO] Found local script: %PS_SCRIPT%
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
  echo.
  echo [ERROR] Guided installer exited with code %EXIT_CODE%.
  echo If it failed before opening, try running this CMD launcher as Administrator.
  pause
)

exit /b %EXIT_CODE%