@echo off
setlocal

echo ======================================================================
echo   ManualBuilder Phase 1 foundation tests
echo ======================================================================
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-Static.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-LocalStorage.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-LocalAppCache.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ProjectCatalog.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ProjectStore.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-CaptureStore.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-WebRender.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ExcelUtilities.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-WordUtilities.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-Server.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ProjectLibraryServer.ps1"
if errorlevel 1 goto :failed

echo.
echo All Phase 1 foundation tests passed.
pause
exit /b 0

:failed
echo.
echo One or more tests failed. Copy the console output and share it.
pause
exit /b 1
