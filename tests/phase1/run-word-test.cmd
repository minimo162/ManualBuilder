@echo off
setlocal

echo ======================================================================
echo   ManualBuilder Word COM export test
echo ======================================================================
echo.
echo Close every Word window before continuing.
echo This test creates a docx in a temporary folder and removes it afterward.
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-WordExport.ps1"
if errorlevel 1 goto :failed

echo.
echo Word export test passed.
pause
exit /b 0

:failed
echo.
echo Word export test failed. Copy the console output and share it.
pause
exit /b 1
