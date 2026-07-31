@echo off
rem =====================================================================
rem  run.cmd  --  ManualBuilder Phase 0 verification kit launcher
rem
rem  ASCII-only on purpose.
rem  This launcher does not rewrite or unblock the supplied scripts.
rem  Run the gate first so that MOTW and encoding can be inspected intact.
rem =====================================================================
setlocal enabledelayedexpansion
cd /d "%~dp0"
set "gate_ok=0"

echo.
echo ======================================================================
echo   ManualBuilder Phase 0 kit
echo ======================================================================
echo.
echo   No files have been changed. Run item 0 first.

:menu
echo.
echo ======================================================================
echo   ManualBuilder Phase 0 verification kit
echo ======================================================================
echo.
echo   [GATE]
echo    0. 00-gate-check.ps1     V-0 V-14 V-16 V-17   (30 sec) *run first
echo.
echo   [Environment]
echo    1. 01-env.ps1            V-15 and machine info    (30 sec)
echo    2. 02-known-folders.ps1  V-9 V-18  screenshot dir (1 min)
echo    3. 03-fonts.ps1          V-13 fonts               (10 sec)
echo.
echo   [Word safety / secondary output]
echo    4. supervised Word test  V-3 V-4 V-5 V-11         (close Word first)
echo    5.   -^> supervised x10 repeat (close Word / leak check)
echo    6.   -^> supervised cancel test  (-CancelAtStep 3)
echo    7.   -^> supervised abnormal branch  (-SimulateExisting)
echo.
echo   [Others]
echo    8. 05-capture-dpi.ps1    V-6  DPI / multi monitor (1 min)
echo    9. 06-mutex.ps1          V-12 single instance     (1 min)
echo   10. 07-html-size.ps1      V-19 HTML size           (2 min)
echo   11. 08-server-min.ps1     V-1 V-2 V-7 V-10         (10 min)
echo.
echo   [Excel primary output]
echo   12. supervised Excel test X-01/02/05-08/10/11/13    (close Excel first)
echo   13.   -^> supervised x10 repeat (close Excel / leak check)
echo   14.   -^> supervised cancel test  (-CancelAtStep 3)
echo   15.   -^> supervised abnormal branch  (-SimulateExisting)
echo.
echo   16. open the out folder
echo   17. open the product UI/UX preview
echo    Q. quit
echo.
set "sel="
set /p sel="Enter number: "

if "%sel%"=="0"  (
  call :run "00-gate-check.ps1" ""
  if errorlevel 1 ( set "gate_ok=0" ) else ( set "gate_ok=1" )
  goto menu
)
if "%sel%"=="1"  ( call :run "01-env.ps1" ""                               & goto menu )
if "%sel%"=="2"  ( call :run "02-known-folders.ps1" ""                     & goto menu )
if "%sel%"=="3"  ( call :run "03-fonts.ps1" ""                             & goto menu )
if "%sel%"=="4"  ( call :runsta "04-word-layout-supervisor.ps1" "-Runs 1"                  & goto menu )
if "%sel%"=="5"  ( call :runsta "04-word-layout-supervisor.ps1" "-Runs 10"                 & goto menu )
if "%sel%"=="6"  ( call :runsta "04-word-layout-supervisor.ps1" "-Runs 1 -CancelAtStep 3"  & goto menu )
if "%sel%"=="7"  ( call :runsta "04-word-layout-supervisor.ps1" "-Runs 1 -SimulateExisting" & goto menu )
if "%sel%"=="8"  ( call :run "05-capture-dpi.ps1" ""                       & goto menu )
if "%sel%"=="9"  ( call :run "06-mutex.ps1" ""                             & goto menu )
if "%sel%"=="10" ( call :run "07-html-size.ps1" ""                         & goto menu )
if "%sel%"=="11" ( call :runsta "08-server-min.ps1" ""                     & goto menu )
if "%sel%"=="12" if "%gate_ok%"=="0" goto needgate
if "%sel%"=="13" if "%gate_ok%"=="0" goto needgate
if "%sel%"=="14" if "%gate_ok%"=="0" goto needgate
if "%sel%"=="15" if "%gate_ok%"=="0" goto needgate
if "%sel%"=="12" ( call :runsta "09-excel-layout-supervisor.ps1" "-Runs 1"                   & goto menu )
if "%sel%"=="13" ( call :runsta "09-excel-layout-supervisor.ps1" "-Runs 10"                  & goto menu )
if "%sel%"=="14" ( call :runsta "09-excel-layout-supervisor.ps1" "-Runs 1 -CancelAtStep 3"   & goto menu )
if "%sel%"=="15" ( call :runsta "09-excel-layout-supervisor.ps1" "-Runs 1 -SimulateExisting" & goto menu )
if "%sel%"=="16" (
  if exist ".\out" ( start "" explorer "%~dp0out" ) else ( echo   The out folder does not exist yet. )
  goto menu
)
if "%sel%"=="17" (
  if exist ".\uiux-preview.html" ( start "" "%~dp0uiux-preview.html" ) else ( echo   uiux-preview.html was not found. )
  goto menu
)
if /i "%sel%"=="Q" goto end
goto menu

:needgate
echo.
echo   Run item 0 successfully before starting the Excel tests.
pause
goto menu

:run
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File ".\%~1" %~2
set "run_rc=!errorlevel!"
echo.
pause
exit /b !run_rc!

:runsta
echo.
powershell -NoProfile -STA -ExecutionPolicy Bypass -File ".\%~1" %~2
set "run_rc=!errorlevel!"
echo.
pause
exit /b !run_rc!

:end
endlocal
