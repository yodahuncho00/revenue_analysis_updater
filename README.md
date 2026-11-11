# Revenue Analysis Updater

Windows ETL that independently updates the Gateway and POS revenue workbooks from Backoffice and POS Portal, then refreshes both standalone HTML dashboards.

## Install on Windows

1. Get access to the authorized company workbook and Backoffice account. The workbook is intentionally not stored in this public repository.
2. Install Python 3.11 or newer, Microsoft Excel desktop, and Git.
3. Clone this repository into a local folder (not a shared folder if another computer also runs the same workbook):
   ```powershell
   git clone https://github.com/yodahuncho00/revenue_analysis_updater.git
   cd revenue_analysis_updater
   ```
4. Copy the authorized `05. Revenue_analysis_Gateway.xlsx` and `05. Revenue_analysis_POS.xlsx` into the project root.
5. Install project dependencies and Chromium:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\install_windows.ps1
   ```
   The installer uses Windows' trusted certificate store for the Chromium download.
   If you already ran it and saw `unable to verify the first certificate`, retry with:
   ```powershell
   $env:NODE_USE_SYSTEM_CA = "1"
   $env:PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT = "120000"
   .venv\Scripts\python.exe -m playwright install chromium
   ```
6. Authenticate this computer interactively. Run the command below and complete login and OTP in both browser tabs:
   ```powershell
   .venv\Scripts\python.exe scripts\refresh_backoffice_session.py
   ```
   Tab 1 is Backoffice and tab 2 is POS Payments. Session state is saved per Windows user. Do not copy `data\storage_state.json` or the browser profile from another computer.
7. Register the daily task at 10:58 local time:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\setup_task_scheduler.ps1
   ```
   The task runs as the current interactive Windows user. That user must be logged in, and Excel must be available. The task can wake the computer from sleep if Windows power settings and hardware allow wake timers. Backoffice may request OTP again when its session expires; the daily run then waits for login before starting the ETL.

## Daily workflow

The scheduled task calls `scripts\run_daily_etl.bat`. Chromium opens Backoffice first and POS Payments second. Each source ETL is queued as soon as its login is confirmed, so one source can update while the other still needs login. Chromium closes as soon as both logins pass; the batch then finishes the queued ETLs one at a time. Each successful workbook refresh rebuilds its dashboard. Failures create `ATTENTION - Revenue update failed.txt` and are recorded in `data\logs\scheduler.log`. You can run the same sequence manually with:

```powershell
scripts\run_daily_etl.bat
```

To verify that Windows popups are visible without running the ETL or changing the workbook, run:
```powershell
scripts\run_daily_etl.bat --test-alert
```

The workbooks and dashboard outputs are local/private files and are ignored by Git. Sanitized dashboard templates are stored under `templates/`.

## Keep one runner per workbook

Only one computer should have the scheduled task enabled for a given shared workbook. Concurrent jobs can collide on workbook locks or overwrite each other's updates. This repository contains application code and a sanitized dashboard template; it does not contain company workbook data, exports, backups, or login sessions.

For operations and recovery details, see [RUNBOOK.md](RUNBOOK.md).
