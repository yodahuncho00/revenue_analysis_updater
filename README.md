# Revenue Analysis Updater

Windows ETL that updates the revenue analysis workbook from Backoffice and refreshes a standalone HTML dashboard. `main.py` rebuilds the dashboard after a successful workbook commit.

## Install on Windows

1. Get access to the authorized company workbook and Backoffice account. The workbook is intentionally not stored in this public repository.
2. Install Python 3.11 or newer, Microsoft Excel desktop, and Git.
3. Clone this repository into a local folder (not a shared folder if another computer also runs the same workbook):
   ```powershell
   git clone https://github.com/yodahuncho00/revenue_analysis_updater.git
   cd revenue_analysis_updater
   ```
4. Copy the authorized `05. Revenue_analysis.xlsx` into the project root.
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
6. Authenticate this computer interactively. Run the command below, sign in and complete OTP in the browser, then press Enter after the Backoffice Dashboard loads:
   ```powershell
   .venv\Scripts\python.exe scripts\refresh_backoffice_session.py
   ```
   Session state is saved per Windows user. Do not copy `data\storage_state.json` or the browser profile from another computer.
7. Register the daily task at 08:30 local time:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\setup_task_scheduler.ps1
   ```
   The task runs as the current interactive Windows user. That user must be logged in, and Excel must be available. The task can wake the computer from sleep if Windows power settings and hardware allow wake timers. Backoffice may request OTP again when its session expires.

## Daily workflow

The scheduled task calls `scripts\run_daily_etl.bat`, which runs `main.py`. After a successful workbook commit, `main.py` rebuilds `05. Revenue_dashboard.html` from the workbook. If the workbook update or dashboard refresh fails, the batch job creates `ATTENTION - Revenue update failed.txt` and displays a popup with the error, run ID, and paths to the detailed log and report. Login or OTP failures include the session refresh command. You can run the same sequence manually with:

```powershell
scripts\run_daily_etl.bat
```

To verify that Windows popups are visible without running the ETL or changing the workbook, run:
```powershell
scripts\run_daily_etl.bat --test-alert
```

The dashboard output and workbook are local generated/private files and are ignored by Git. The HTML template is `templates/05. Revenue_dashboard.template.html`.

## Keep one runner per workbook

Only one computer should have the scheduled task enabled for a given shared workbook. Concurrent jobs can collide on workbook locks or overwrite each other's updates. This repository contains application code and a sanitized dashboard template; it does not contain company workbook data, exports, backups, or login sessions.

For operations and recovery details, see [RUNBOOK.md](RUNBOOK.md).
