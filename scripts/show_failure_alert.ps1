param(
    [Parameter(Mandatory = $true)]
    [string]$AlertFile,
    [string]$RunReport = ""
)

Add-Type -AssemblyName System.Windows.Forms
$Message = Get-Content -LiteralPath $AlertFile -Raw
if ($Message -notmatch "TEST ALERT ONLY") {
    if ($RunReport -and (Test-Path -LiteralPath $RunReport)) {
        try {
            $Report = Get-Content -LiteralPath $RunReport -Raw | ConvertFrom-Json
            if ($Report.error_message) {
                $Message += "`r`nFailure details: $($Report.error_message)`r`n"
            }
            if ($Report.error_message -match "(?i)authentication is required|login or verification|OTP|one.time.code") {
                $Message += "`r`nBackoffice login/OTP is required. In PowerShell, run:`r`n.venv\Scripts\python.exe scripts\refresh_backoffice_session.py`r`nComplete the sign-in and OTP in the browser, then press Enter. Retry with scripts\run_daily_etl.bat."
            }
            if ($Report.checks_passed -contains "atomic_commit_successful") {
                $Message += "`r`nThe workbook commit succeeded before a later step failed. Check whether the dashboard refresh completed before retrying."
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
