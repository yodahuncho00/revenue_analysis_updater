@echo off
setlocal
cd /d "%~dp0\.."
if not exist data\logs mkdir data\logs
set "PYTHON=%CD%\.venv\Scripts\python.exe"
echo [%DATE% %TIME%] Starting daily Revenue Analysis ETL run... >> data\logs\scheduler.log
"%PYTHON%" main.py >> data\logs\scheduler.log 2>&1
set "ETL_EXIT=%ERRORLEVEL%"
if not "%ETL_EXIT%"=="0" goto finished
echo [%DATE% %TIME%] ETL succeeded; refreshing HTML dashboard... >> data\logs\scheduler.log
"%PYTHON%" scripts\update_revenue_dashboard.py >> data\logs\scheduler.log 2>&1
set "DASHBOARD_EXIT=%ERRORLEVEL%"
if not "%DASHBOARD_EXIT%"=="0" set "ETL_EXIT=%DASHBOARD_EXIT%"
:finished
echo [%DATE% %TIME%] Daily job completed with exit code %ETL_EXIT% >> data\logs\scheduler.log
exit /b %ETL_EXIT%
