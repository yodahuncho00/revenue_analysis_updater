param(
    [Parameter(Mandatory = $true)]
    [string]$AlertFile,
    [string]$RunReport = "",
    [string]$SummaryFile = ""
)

Add-Type -AssemblyName System.Windows.Forms
$Message = Get-Content -LiteralPath $AlertFile -Raw
if ($Message -notmatch "TEST ALERT ONLY") {
    if ($SummaryFile -and (Test-Path -LiteralPath $SummaryFile)) {
        try {
            $Summary = Get-Content -LiteralPath $SummaryFile -Raw | ConvertFrom-Json
            $Message += "`r`n--- Per-Source Status ---`r`n"
            foreach ($SourceName in @("backoffice", "pos")) {
                $Info = $Summary.sources.$SourceName
                if (-not $Info) { continue }
                $Label = $SourceName.ToUpper()
                $Status = $Info.etl_status
                $Line = "${Label}: $Status"
                if ($Info.run_id) { $Line += " (run_id=$($Info.run_id))" }
                $Message += "$Line`r`n"
                if ($Info.error) {
                    $Message += "  Details: $($Info.error)`r`n"
                }
                if ($Info.report_path -and (Test-Path -LiteralPath $Info.report_path)) {
                    try {
                        $SrcReport = Get-Content -LiteralPath $Info.report_path -Raw | ConvertFrom-Json
                        if ($SrcReport.error_message) {
                            $Message += "  Error: $($SrcReport.error_message)`r`n"
                            if ($SrcReport.error_message -match "(?i)Google OAuth|authorize_google_sheets") {
                                $Message += "  Fix: Run .venv\Scripts\python.exe scripts\authorize_google_sheets.py, then scripts\run_daily_etl.bat --resume`r`n"
                            }
                        }
                        if ($SrcReport.checks_passed -contains "atomic_commit_successful") {
                            $Message += "  Note: Local workbook committed before a later sink failed. Re-run with --resume.`r`n"
                        }
                    } catch {
                        $Message += "  Report unreadable: $($Info.report_path)`r`n"
                    }
                }
                if ($Info.log_path) {
                    $Message += "  Log: $($Info.log_path)`r`n"
                }
            }
        } catch {
            $Message += "`r`nCould not read preflight summary file.`r`n"
        }
    } elseif ($Message -match "(?i)PRE-FLIGHT FAILED") {
        $Message += "`r`nThe daily ETL was stopped before it started. Check data\logs\scheduler.log for the exact preflight error. If Python is missing, rerun scripts\install_windows.ps1 from the project folder. If login/OTP is required, complete it in Chromium and rerun scripts\run_daily_etl.bat."
    } elseif ($RunReport -and (Test-Path -LiteralPath $RunReport)) {
        try {
            $Report = Get-Content -LiteralPath $RunReport -Raw | ConvertFrom-Json
            if ($Report.error_message) {
                $Message += "`r`nFailure details: $($Report.error_message)`r`n"
            }
            if ($Report.error_message -match "(?i)authentication is required|login or verification|OTP|one.time.code") {
                $Message += "`r`nBackoffice login/OTP is required. In PowerShell, run:`r`n.venv\Scripts\python.exe scripts\refresh_backoffice_session.py`r`nComplete sign-in and OTP in the browser, then wait for the Dashboard check to finish. Retry with scripts\run_daily_etl.bat."
            }
            if ($Report.checks_passed -contains "atomic_commit_successful") {
                $Message += "`r`nThe workbook commit succeeded before a later step failed. Re-run with --resume to finish remaining publish steps."
            } elseif ($Report.status -eq "FAILED") {
                $Message += "`r`nThe workbook was not committed; the last committed workbook remains in place."
            }
        } catch {
            $Message += "`r`nThe run report could not be read. Use the run log path above."
        }
    } else {
        $Message += "`r`nNo run report was written. The Python process may have crashed or stopped before its error handler; use the run log path above and check Task Scheduler history."
    }
}
[System.Windows.Forms.MessageBox]::Show(
    $Message,
    "Revenue Analysis updater needs attention",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Error
) | Out-Null
