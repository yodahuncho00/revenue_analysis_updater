@echo off
setlocal
cd /d "%~dp0\.."
if not exist data\logs mkdir data\logs
set "PYTHON=%CD%\.venv\Scripts\python.exe"
set "ALERT_FILE=%CD%\ATTENTION - Revenue update failed.txt"
set "SUMMARY_FILE=%CD%\data\logs\preflight_summary.json"
if /I "%~1"=="--test-alert" goto test_alert

if not "%REVENUE_ETL_REEXEC%"=="1" (
  if /I not "%~1"=="--no-update" (
    if exist scripts\auto_update.ps1 (
      powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0auto_update.ps1" >> data\logs\scheduler.log 2>&1
      if "%ERRORLEVEL%"=="10" (
        set "REVENUE_ETL_REEXEC=1"
        call "%~f0" %*
        exit /b %ERRORLEVEL%
      )
    )
  )
)

echo [%DATE% %TIME%] Starting daily Revenue Analysis ETL run... >> data\logs\scheduler.log
echo [%DATE% %TIME%] Project directory: %CD% >> data\logs\scheduler.log
if exist data\logs\latest_run.txt del data\logs\latest_run.txt
if exist data\logs\latest_run_backoffice.txt del data\logs\latest_run_backoffice.txt
if exist data\logs\latest_run_pos.txt del data\logs\latest_run_pos.txt
if exist "%SUMMARY_FILE%" del "%SUMMARY_FILE%"
echo [%DATE% %TIME%] Checking sessions and running Backoffice + POS pipelines... >> data\logs\scheduler.log
if not exist "%PYTHON%" (
  echo [%DATE% %TIME%] ERROR: Virtual-environment Python not found: %PYTHON% >> data\logs\scheduler.log
  set "PREFLIGHT_EXIT=3"
  goto preflight_failed
)
if not exist scripts\refresh_backoffice_session.py (
  echo [%DATE% %TIME%] ERROR: Session preflight script not found: %CD%\scripts\refresh_backoffice_session.py >> data\logs\scheduler.log
  set "PREFLIGHT_EXIT=3"
  goto preflight_failed
)
set "EXTRA_ARGS="
if /I "%~1"=="--resume" set "EXTRA_ARGS=--resume"
if /I "%~2"=="--resume" set "EXTRA_ARGS=--resume"
if /I "%~1"=="--force-visible" set "EXTRA_ARGS=%EXTRA_ARGS% --force-visible"
if /I "%~2"=="--force-visible" set "EXTRA_ARGS=%EXTRA_ARGS% --force-visible"
"%PYTHON%" scripts\refresh_backoffice_session.py --preflight %EXTRA_ARGS% >> data\logs\scheduler.log 2>&1
set "PREFLIGHT_EXIT=%ERRORLEVEL%"
if not "%PREFLIGHT_EXIT%"=="0" (
  set "ETL_EXIT=%PREFLIGHT_EXIT%"
  set "FAILED_STAGE=One or more source pipelines failed; see per-source status below"
  goto failed
)
echo [%DATE% %TIME%] Both sources completed successfully. >> data\logs\scheduler.log
set "ETL_EXIT=0"
if exist "%ALERT_FILE%" del "%ALERT_FILE%"
goto finished
:failed
echo [%DATE% %TIME%] FAILED at %FAILED_STAGE%; exit code %ETL_EXIT%. >> data\logs\scheduler.log
set "RUN_ID=unknown"
if exist data\logs\latest_run.txt set /p RUN_ID=<data\logs\latest_run.txt
(
  echo Revenue Analysis daily run needs attention.
  echo Failed stage: %FAILED_STAGE%.
  echo Exit code: %ETL_EXIT%.
  echo.
  echo Summary file: data\logs\preflight_summary.json
  echo Scheduler log: data\logs\scheduler.log
  echo Last Run ID: %RUN_ID%.
  echo To retry only failed/unfinished stages without re-scraping completed ones:
  echo   Double-click "Update and re-run" on your Desktop (or run scripts\run_daily_etl.bat --resume)
  echo If login or OTP is requested, run:
  echo   .venv\Scripts\python.exe scripts\refresh_backoffice_session.py
  echo Do not send the workbook, exports, browser profile, or storage_state.json.
) > "%ALERT_FILE%"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0show_failure_alert.ps1" -AlertFile "%ALERT_FILE%" -RunReport "%CD%\data\reports\%RUN_ID%.json" -SummaryFile "%SUMMARY_FILE%"
goto finished
:preflight_failed
echo [%DATE% %TIME%] FAILED before both source pipelines completed. Exit code %PREFLIGHT_EXIT%. >> data\logs\scheduler.log
(
  echo SOURCE RUN FAILED - Revenue Analysis daily run did not complete.
  echo The Python environment or login orchestration did not complete successfully.
  echo A source that passed login may already have completed its update.
  echo.
  echo Check data\logs\scheduler.log for details. Complete both logins in Chromium and keep it open until both are confirmed.
) > "%ALERT_FILE%"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0show_failure_alert.ps1" -AlertFile "%ALERT_FILE%"
exit /b %PREFLIGHT_EXIT%
:test_alert
echo [%DATE% %TIME%] Showing a test failure alert; ETL is not running. >> data\logs\scheduler.log
(
  echo TEST ALERT ONLY - the ETL and workbook were not run or changed.
  echo This is the failure popup test. Close this message to finish.
) > "%ALERT_FILE%"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0show_failure_alert.ps1" -AlertFile "%ALERT_FILE%"
del "%ALERT_FILE%"
exit /b 0
:finished
echo [%DATE% %TIME%] Daily job completed with exit code %ETL_EXIT% >> data\logs\scheduler.log
exit /b %ETL_EXIT%
