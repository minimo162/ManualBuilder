@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0src\Start-ManualBuilderLauncher.ps1"
set "MB_EXIT=%ERRORLEVEL%"
if not "%MB_EXIT%"=="0" (
  echo.
  echo ManualBuilder could not start. Exit code: %MB_EXIT%
  pause
)
exit /b %MB_EXIT%
