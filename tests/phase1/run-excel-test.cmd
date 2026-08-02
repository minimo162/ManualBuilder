@echo off
setlocal

echo ======================================================================
echo   ManualBuilder Excel COM export test
echo ======================================================================
echo.
echo Close every Excel window before continuing.
echo This test creates an xlsx in a temporary folder and removes it afterward.
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ExcelExport.ps1"
set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" echo Excel export test failed. Copy the console output and share it.
if "%RESULT%"=="0" echo Excel export test passed.
pause
exit /b %RESULT%
