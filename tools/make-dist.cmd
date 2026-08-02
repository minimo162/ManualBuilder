@echo off
setlocal
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0New-MbDistribution.ps1" %*
set "MB_EXIT=%ERRORLEVEL%"
echo.
pause
exit /b %MB_EXIT%
