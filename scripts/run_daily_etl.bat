@echo off
setlocal
cd /d "%~dp0\.."
if not exist data\logs mkdir data\logs
set "PYTHON=%CD%\.venv\Scripts\python.exe"
set "ALERT_FILE=%CD%\ATTENTION - Revenue update failed.txt"
if /I "%~1"=="--test-alert" goto test_alert
echo [%DATE% %TIME%] Starting daily Revenue Analysis ETL run... >> data\logs\scheduler.log
if exist data\logs\latest_run.txt del data\logs\latest_run.txt
"%PYTHON%" main.py >> data\logs\scheduler.log 2>&1
set "ETL_EXIT=%ERRORLEVEL%"
if not "%ETL_EXIT%"=="0" (
  set "FAILED_STAGE=Revenue pipeline (workbook update or dashboard refresh)"
  goto failed
)
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
  echo Run ID: %RUN_ID%.
  echo Run log: data\logs\%RUN_ID%.log
  echo Run report: data\reports\%RUN_ID%.md
  echo Read the popup details and report before retrying.
  echo If it still fails, send the Run ID and those two files to your support person.
  echo If login or OTP is requested, use scripts\refresh_backoffice_session.py.
  echo Do not send the workbook, exports, browser profile, or storage_state.json.
) > "%ALERT_FILE%"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0show_failure_alert.ps1" -AlertFile "%ALERT_FILE%" -RunReport "%CD%\data\reports\%RUN_ID%.json"
goto finished
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
