param(
    [switch]$RegisterDailyTask
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $ProjectDir

if (-not (Test-Path -LiteralPath (Join-Path $ProjectDir "05. Revenue_analysis.xlsx"))) {
    throw "05. Revenue_analysis.xlsx is missing. Copy the authorized workbook into the project folder first."
}

$PythonLauncher = Get-Command py -ErrorAction SilentlyContinue
if ($PythonLauncher) {
    & $PythonLauncher.Source -3 --version
    if ($LASTEXITCODE -ne 0) {
        throw "Python 3 was not found. Install Python 3.11 or newer, then rerun this setup."
    }
    & $PythonLauncher.Source -3 -m venv (Join-Path $ProjectDir ".venv")
} else {
    $PythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if (-not $PythonCommand) {
        throw "Install Python 3.11 or newer, then rerun this setup."
    }
    & $PythonCommand.Source -m venv (Join-Path $ProjectDir ".venv")
}
if ($LASTEXITCODE -ne 0) { throw "Could not create the project's Python virtual environment." }

$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
& $VenvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) { throw "Could not upgrade pip in the project environment." }
& $VenvPython -m pip install -r (Join-Path $ProjectDir "requirements.txt")
if ($LASTEXITCODE -ne 0) { throw "Could not install the Python requirements." }
& $VenvPython -m playwright install chromium
if ($LASTEXITCODE -ne 0) { throw "Could not install Playwright Chromium." }

foreach ($Directory in @("data\logs", "data\raw", "data\exports", "data\reports", "data\backups")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $ProjectDir $Directory) | Out-Null
}

$DashboardOutput = Join-Path $ProjectDir "05. Revenue_dashboard.html"
$DashboardTemplate = Join-Path $ProjectDir "templates\05. Revenue_dashboard.template.html"
if (-not (Test-Path -LiteralPath $DashboardOutput)) {
    Copy-Item -LiteralPath $DashboardTemplate -Destination $DashboardOutput
}

Write-Host "Dependencies installed in $VenvPython"
Write-Host "Next, authenticate this computer's own Backoffice session:"
Write-Host "  $VenvPython scripts\refresh_backoffice_session.py"
Write-Host "Complete login and OTP in the browser. Do not copy another user's storage_state.json or browser profile."

if ($RegisterDailyTask) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "setup_task_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Could not register the daily scheduled task." }
}
