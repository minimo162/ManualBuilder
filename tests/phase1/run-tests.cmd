@echo off
setlocal

echo ======================================================================
echo   ManualBuilder Phase 1 foundation tests
echo ======================================================================
echo.

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-Static.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-LocalDraft.ps1"
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

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-CaptureHeartbeat.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-VideoAttachment.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-VideoToManual.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-EvalFixtures.ps1"
if errorlevel 1 goto :failed

rem Evaluation scoring is implemented in dependency-free Python.
where python >nul 2>&1
if errorlevel 1 (
  echo [SKIP] Python not found. Skipping the evaluation scoring tests.
) else (
  python "%~dp0Test-RecorderCopilotRunEval.py"
  if errorlevel 1 goto :failed
  python "%~dp0Test-ProductProjectEval.py"
  if errorlevel 1 goto :failed
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-WebRender.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ExcelUtilities.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-OfficeLayoutFixture.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-WordUtilities.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-Server.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-ProjectLibraryServer.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-Recorder.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-RecorderCopilot.ps1"
if errorlevel 1 goto :failed

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-RecorderCopilotManualTools.ps1"
if errorlevel 1 goto :failed

rem The recorder monitor's C# is compiled at run time, so a syntax error would
rem only surface when a user starts recording. Compile it here instead.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Test-RecorderCompanionCompile.ps1"
if errorlevel 1 goto :failed

rem Scene splitting runs in the browser, so Node checks it. Skipped when Node is absent.
where node >nul 2>&1
if errorlevel 1 (
  echo [SKIP] Node not found. Skipping the video scene tests.
) else (
  node "%~dp0Test-VideoScenes.mjs"
  if errorlevel 1 goto :failed
  node "%~dp0Test-WebAssets.mjs"
  if errorlevel 1 goto :failed
)

echo.
echo [SKIP] Excel/Word COM export tests are not part of this suite.
echo        The primary output path is NOT covered here. Close Excel and Word,
echo        then run run-excel-test.cmd and run-word-test.cmd before releasing.

echo.
echo All Phase 1 foundation tests passed.
if not defined CI pause
exit /b 0

:failed
echo.
echo One or more tests failed. Copy the console output and share it.
if not defined CI pause
exit /b 1
