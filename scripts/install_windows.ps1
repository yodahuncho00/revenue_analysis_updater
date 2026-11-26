param(
    [switch]$RegisterDailyTask,
    [switch]$EnableAutoUpdate,
    [string]$UpdateBranch = "release",
    [switch]$CreateDesktopShortcuts,
    [switch]$SkipWorkbookCheck
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $ProjectDir

$LocalConfigPath = Join-Path $ProjectDir "config.local.yaml"
if (-not $SkipWorkbookCheck -and -not (Test-Path -LiteralPath $LocalConfigPath)) {
    if (-not (Test-Path -LiteralPath (Join-Path $ProjectDir "05. Revenue_analysis_Gateway.xlsx"))) {
        throw "05. Revenue_analysis_Gateway.xlsx is missing. Copy the authorized workbook into the project folder first (or configure paths in config.local.yaml / pass -SkipWorkbookCheck)."
    }
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
$ReqPath = Join-Path $ProjectDir "requirements.txt"
& $VenvPython -m pip install -r $ReqPath
if ($LASTEXITCODE -ne 0) { throw "Could not install the Python requirements." }
# Record requirements.txt hash so auto_update.ps1 skips unnecessary pip installs
$ReqHash = (Get-FileHash -LiteralPath $ReqPath -Algorithm SHA256).Hash.ToLowerInvariant()
Set-Content -LiteralPath (Join-Path $ProjectDir ".venv\.requirements.sha256") -Value $ReqHash -Encoding ASCII

# On managed Windows networks, the system/browser may trust a proxy CA that
# Node's bundled CA list does not. Recent Node runtimes can use Windows' trusted
# certificate store; keep TLS verification enabled and allow slow downloads.
$env:NODE_USE_SYSTEM_CA = "1"
$env:PLAYWRIGHT_DOWNLOAD_CONNECTION_TIMEOUT = "120000"
& $VenvPython -m playwright install chromium
if ($LASTEXITCODE -ne 0) {
    throw "Could not install Playwright Chromium. Verify that Windows trusts the network certificate and retry with .venv\Scripts\python.exe -m playwright install chromium."
}

foreach ($Directory in @("data\logs", "data\raw", "data\exports", "data\reports", "data\backups")) {
    New-Item -ItemType Directory -Force -Path (Join-Path $ProjectDir $Directory) | Out-Null
}

$DashboardOutput = Join-Path $ProjectDir "05. Revenue_dashboard_Gateway.html"
$DashboardTemplate = Join-Path $ProjectDir "templates\05. Revenue_dashboard_Gateway.template.html"
if (-not (Test-Path -LiteralPath $DashboardOutput)) {
    Copy-Item -LiteralPath $DashboardTemplate -Destination $DashboardOutput
}
$PosDashboardOutput = Join-Path $ProjectDir "05. Revenue_dashboard_POS.html"
$PosDashboardTemplate = Join-Path $ProjectDir "templates\05. Revenue_dashboard_POS.template.html"
if (-not (Test-Path -LiteralPath $PosDashboardOutput)) {
    Copy-Item -LiteralPath $PosDashboardTemplate -Destination $PosDashboardOutput
}

if ($EnableAutoUpdate) {
    $pyUpdateCfg = @"
import os, sys, yaml
path, branch = sys.argv[1], sys.argv[2]
data = {}
if os.path.isfile(path):
    loaded = yaml.safe_load(open(path, encoding='utf-8'))
    if isinstance(loaded, dict):
        data = loaded
au = data.setdefault('auto_update', {})
au['enabled'] = True
au['branch'] = branch
au.setdefault('remote', 'origin')
au.setdefault('fetch_timeout_seconds', 20)
with open(path, 'w', encoding='utf-8') as f:
    yaml.safe_dump(data, f, sort_keys=False, allow_unicode=True)
"@
    & $VenvPython -c $pyUpdateCfg $LocalConfigPath $UpdateBranch
    if ($LASTEXITCODE -ne 0) { throw "Could not configure auto_update in config.local.yaml." }
    Write-Host "Enabled automatic pre-run updates from origin/$UpdateBranch in config.local.yaml."
}

if ($CreateDesktopShortcuts) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "create_desktop_shortcut.ps1") -ProjectDir $ProjectDir
    if ($LASTEXITCODE -ne 0) { throw "Could not create the Desktop shortcut." }
}

Write-Host "Dependencies installed in $VenvPython"
Write-Host "Next steps on this computer:"
Write-Host "  1. Authorize Google Sheets OAuth once:"
Write-Host "       $VenvPython scripts\authorize_google_sheets.py"
Write-Host "  2. Authenticate this computer's own Backoffice and POS Portal sessions:"
Write-Host "       $VenvPython scripts\refresh_backoffice_session.py"
Write-Host "Complete login and OTP in both tabs. Do not copy another user's storage_state.json or browser profile."

if ($RegisterDailyTask) {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "setup_task_scheduler.ps1")
    if ($LASTEXITCODE -ne 0) { throw "Could not register the daily scheduled task." }
}
