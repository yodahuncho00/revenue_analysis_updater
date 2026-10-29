# Operations Runbook: Autonomous Revenue Analysis ETL Pipeline

**System:** Revenue Analysis Daily Updater  
**Primary Workbook:** `05. Revenue_analysis.xlsx`  
**Configuration:** `config.yaml`  
**Timezone:** `UTC+3`

---

## 1. Quick Start & Common Commands

All commands are run from the project root using PowerShell or Command Prompt.

### 1.0 Set Up the Python Environment
Activate the project virtual environment, then install the declared dependencies and
Playwright's Chromium browser:
```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```
If PowerShell is currently in the `scripts` directory, use
`python -m pip install -r ..\requirements.txt` instead.

To install on another Windows computer, first copy the project and the authorized
workbook there, then run this from PowerShell in the project folder:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\install_windows.ps1
```
The installer creates a local `.venv`, installs the requirements and Chromium,
and creates the data directories. It does not copy Backoffice login state. On the
new computer, run `python scripts\refresh_backoffice_session.py` and complete that
computer's login and OTP. After authentication, register the daily task with
`powershell -ExecutionPolicy Bypass -File scripts\setup_task_scheduler.ps1`.
Use one scheduled runner per shared workbook to prevent concurrent updates.

### 1.1 Dry Run (Zero-Risk Validation)
Uses the live Playwright extractor by default and executes the full pipeline
(extraction, validation, precision write to working copy, Excel COM recalculation,
and post-write reconciliation) without touching the production workbook:
```bash
python main.py --dry-run
```

Fixture data is for tests only; select it explicitly with `--extractor fixture`.
Do not use the fixture extractor for a production update.

### 1.2 Refresh the Backoffice Login Session
The live ETL and the refresh utility share a persistent Chromium profile under
`%LOCALAPPDATA%\RevenueAnalysisUpdater\BackofficeProfile`. On first use, the
existing `data/storage_state.json` is imported into that profile once. If
Backoffice asks for login or OTP, refresh the profile interactively:
```powershell
python scripts\refresh_backoffice_session.py
```
Complete login and OTP in the visible browser. Press Enter in the terminal only
after the Dashboard has loaded. The session state is saved only after the script
verifies the Dashboard is ready. Run the ETL and refresh utility under the same
Windows account, and do not run them concurrently because Chromium locks its profile.
Treat both the profile and `data/storage_state.json` as passwords: they contain
authentication data and must not be shared or committed.

Merchant Balance pastes omit the table header row. Displayed monetary values
are written as numbers with space grouping and comma decimals; zero values
display without decimal places. Country rows in column A are left-aligned and
other rows are right-aligned. Merchant Balance rows remain height 13. The old
table range is cleared first, trailing blank rows are omitted, and borders are
removed from cleared tail rows to avoid stale data and horizontal lines.
Summary `SUMIFS` ranges begin at the first detail row and resize to include the
latest captured table rows.

Provider Balance is excluded from the ETL: it is not extracted, pasted, or
otherwise modified. Its worksheet remains in the workbook unchanged.

### 1.3 Normal Production Daily Run
Runs the complete ETL pipeline and atomically commits updates to `05. Revenue_analysis.xlsx`:
```bash
python main.py
```

### 1.4 Backfill or Run for a Specific Date
Re-runs or backfills the pipeline for an explicit as-of date (calculates `to_date = D - 1`):
```bash
python main.py --date 2026-10-27
```

### 1.5 Run a Specific Step Only
Runs only an isolated step against the working copy:
```bash
python main.py --step dashboard           # Step 1: Dashboard -> BO_extract -> Data
python main.py --step merchants_balance   # Step 2: Merchants Balance
python main.py --step settlements         # Step 4: Settlements (append-only)
python main.py --step topups              # Step 5: Top-ups (append-only)
python main.py --step control             # Step 6: Update Control sheet as-of date
```

### 1.6 Validate the Exact Pipeline Captures
Run the live pipeline in dry-run mode to validate the extracted data without committing
workbook changes:
```powershell
python scripts\extract_to_csv.py
```
Every successful extraction writes the captured data as CSVs under a unique
`data/exports/<run_id>/` directory. These are not a second browser scrape. Settlements
and Top-ups come from the pages' filtered **Export to csv** downloads; CSV IDs are
checked against the paginated API (state-3 rows omitted by the export are allowed).
When existing transaction rows have blank Merchant Order Id or Note cells, the ETL
downloads the date range from the earliest such row and fills only blank cells from
the CSV; existing nonblank details are preserved. The transaction CSVs contain those
downloaded rows as well as any new rows considered for append. The workbook displays
transaction dates as `dd.mm.yyyy` while preserving the full date-time value, including
time, in the cell. A failed extraction or validation does not report success or produce
a complete five-file set.

---

## 2. Self-Healing & Catch-Up Behavior

The pipeline is inherently self-healing after missed runs (e.g. holidays, downtime, network failures):
1. **Dashboard & Data Sheets (Step 1):** Always completely rebuilds the current reporting month from day 1 to $D - 1$. If a day or week was missed, the entire month is refreshed consistently with zero partial state.
2. **Settlements & Top-ups (Steps 4–5):** Inspects the highest existing timestamp $L$ in the sheet and queries from $[L + 1\text{ day}]$ through $D$, skipping weekends. All missed working days are automatically retrieved and appended in strictly ascending order by ID.
3. **Merchant Balance (Step 2):** Replaces its snapshot with the latest live status. Provider Balance is excluded and remains unchanged.

---

## 3. Disaster Recovery & Restoring Backups

Every single ETL run automatically creates a timestamped backup in `data/backups/` before any changes are made:
`data/backups/05. Revenue_analysis_YYYYMMDD_HHMMSS.xlsx`

### To restore a backup:
1. Close Microsoft Excel if open.
2. List available backups:
   ```powershell
   Get-ChildItem -Path "data\backups" | Sort-Object LastWriteTime -Descending
   ```
3. Copy the desired backup over the production workbook:
   ```powershell
   Copy-Item -Path "data\backups\05. Revenue_analysis_20261028_150913.xlsx" -Destination "05. Revenue_analysis.xlsx" -Force
   ```
4. Verify the file opens normally in Excel.

---

## 4. Alert & Failure Diagnostic Catalog

Treat all alerts as hard stop conditions. When an alert occurs, the production workbook is left 100% untouched, and the working copy is retained for debugging.

| Alert / Error | Meaning | Resolution Steps |
|---|---|---|
| `StructureValidationError: Target workbook does not exist` | File was renamed or moved. | Verify `config.yaml` points to the correct workbook path. |
| `StructureValidationError: Workbook is locked or currently open` | Excel has an exclusive lock on the file. | Close Excel or save your work and re-run. |
| `Backoffice authentication is required` / `login or verification form` | The saved session expired or OTP/login is required. | Run `python scripts\refresh_backoffice_session.py`, complete login and OTP, then retry. Production remains untouched on this failure. |
| `Dashboard did not become ready` | Backoffice did not show its expected controls and no login form was detected. | Check network/VPN/Backoffice availability; inspect the run log and retry after service is restored. |
| `CopyManagerError: Another ETL run is in progress (.lock exists)` | A previous run crashed or is currently executing. | Confirm no Python process is running, then remove `05. Revenue_analysis.xlsx.lock`. |
| `RawValidationError: Hierarchy sum mismatch` | Sum of child rows does not foot to parent row in Backoffice table. | Check `data/raw/<run_id>/dashboard.json` for corrupted data in source Backoffice. |
| `RawValidationError: Duplicate key found in raw leaf rows` | The Backoffice returned multiple rows with the exact same route key. | Inspect source Backoffice filters or route configuration. |
| `RawValidationError: Duplicate Settlements/Topups ID` | A fetched ID already exists in the sheet. | Verify Backoffice date filters and ID ordering. |
| `ReconciliationError: Row count mismatch between BO_extract and Data` | The number of `Filter == 1` rows does not match rows written to `Data`. | Check `data/reports/<run_id>.json` diagnostics. |
| `ReconciliationError: Control strict gate failed` | A core accounting identity on `Control` (rows 68, 74–78) failed to foot. | Inspect `Control` tab formula checks. |
| `Mapping notice at row 69: Add merchant to Merchant_master` | A new merchant appeared in Backoffice that has not been mapped yet. | Informational notice: open `Merchant_master` and map the new merchant name in Column B. |

---

## 5. Shadow Mode Verification Workflow

Before going fully live, run the autonomous pipeline in shadow mode:
1. Execute `python main.py --dry-run` daily alongside manual reporting.
2. Compare the output working copy against the manually updated workbook using:
   ```powershell
   python -c "
   import openpyxl
   wb1 = openpyxl.load_workbook('05. Revenue_analysis.xlsx', data_only=True)
   wb2 = openpyxl.load_workbook('05. Revenue_analysis.xlsx.working.xlsx', data_only=True)
   # Assert cell equality across key tabs
   "
   ```
3. Once satisfied with complete cell-by-cell equality over 3–5 days, activate the automated Windows Task Scheduler job.

### Task Scheduler Notes
Install or update the daily job at 08:30 local time (all seven days) from a
PowerShell prompt in the project directory:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_task_scheduler.ps1
```
The task runs the ETL and, only after a successful workbook update, regenerates
`05. Revenue_dashboard.html` from the updated workbook. Its log is
`data\logs\scheduler.log`. To run the same sequence manually, run
`scripts\run_daily_etl.bat`.

Before scheduling unattended daily runs:
1. Run `python scripts\refresh_backoffice_session.py` once under the Windows account
   that will own the scheduled task.
2. The supplied task uses that same interactive Windows account and the project
   root as its working directory, with `.venv\Scripts\python.exe` for both steps.
3. Keep the computer awake and the same user session available. This pipeline uses
   Excel COM for workbook paste/recalculation, so test the task under the intended
   logon state rather than assuming it works after signing out.
4. A scheduled job cannot type an OTP or refresh an expired session by itself. On
   expiry, it must fail without committing; an operator refreshes the session and
   reruns. Configure an external alert from the task's nonzero exit code or run
   report before relying on unattended operation.
