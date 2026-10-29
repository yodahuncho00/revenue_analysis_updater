@echo off
setlocal
cd /d "%~dp0\.."
if not exist data\logs mkdir data\logs
set "PYTHON=%CD%\.venv\Scripts\python.exe"
set "ALERT_FILE=%CD%\ATTENTION - Revenue update failed.txt"
echo [%DATE% %TIME%] Starting daily Revenue Analysis ETL run... >> data\logs\scheduler.log
"%PYTHON%" main.py >> data\logs\scheduler.log 2>&1
set "ETL_EXIT=%ERRORLEVEL%"
if not "%ETL_EXIT%"=="0" (
  set "FAILED_STAGE=Workbook update"
  goto failed
)
echo [%DATE% %TIME%] ETL succeeded; refreshing HTML dashboard... >> data\logs\scheduler.log
"%PYTHON%" scripts\update_revenue_dashboard.py >> data\logs\scheduler.log 2>&1
set "DASHBOARD_EXIT=%ERRORLEVEL%"
if not "%DASHBOARD_EXIT%"=="0" (
  set "ETL_EXIT=%DASHBOARD_EXIT%"
  set "FAILED_STAGE=HTML dashboard refresh"
  goto failed
)
if exist "%ALERT_FILE%" del "%ALERT_FILE%"
goto finished
:failed
echo [%DATE% %TIME%] FAILED at %FAILED_STAGE%; exit code %ETL_EXIT%. >> data\logs\scheduler.log
(
  echo Revenue Analysis daily run needs attention.
  echo Failed stage: %FAILED_STAGE%.
  echo.
  echo If Backoffice asked for login or OTP, open PowerShell in this project folder and run:
  echo .venv\Scripts\python.exe scripts\refresh_backoffice_session.py
  echo Complete login and OTP in the browser, then press Enter in PowerShell.
  echo Then run scripts\run_daily_etl.bat to retry.
  echo.
  echo If it still fails, send this file and data\logs\scheduler.log to your support contact.
) > "%ALERT_FILE%"
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0show_failure_alert.ps1" -AlertFile "%ALERT_FILE%"
:finished
echo [%DATE% %TIME%] Daily job completed with exit code %ETL_EXIT% >> data\logs\scheduler.log
exit /b %ETL_EXIT%
