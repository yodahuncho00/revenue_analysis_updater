<#
.SYNOPSIS
Creates the "Update and re-run" Desktop shortcut with the ZOYK icon and removes legacy shortcuts.
#>
param(
    [string]$ProjectDir = ""
)

$ErrorActionPreference = "Stop"
if (-not $ProjectDir) {
    $ProjectDir = Split-Path -Parent $PSScriptRoot
}
$ProjectDir = (Resolve-Path -LiteralPath $ProjectDir).Path
$DesktopDir = [Environment]::GetFolderPath("Desktop")
$WshShell = New-Object -ComObject WScript.Shell
$BatPath = Join-Path $ProjectDir "scripts\run_daily_etl.bat"
$IconPath = Join-Path $ProjectDir "assets\zoyk_logo.ico"

foreach ($oldName in @("Run Revenue Update.lnk", "Retry Revenue Update.lnk")) {
    $oldPath = Join-Path $DesktopDir $oldName
    if (Test-Path -LiteralPath $oldPath) {
        Remove-Item -LiteralPath $oldPath -Force
        Write-Host "Removed old Desktop shortcut: $oldName"
    }
}

$shortcutName = "Update and re-run.lnk"
$shortcutPath = Join-Path $DesktopDir $shortcutName
$lnk = $WshShell.CreateShortcut($shortcutPath)
$lnk.TargetPath = "cmd.exe"
$lnk.Arguments = "/c `"$BatPath`" --resume"
$lnk.WorkingDirectory = $ProjectDir
$lnk.Description = "Pull latest fixes and run or resume today's Revenue Analysis update"
if (Test-Path -LiteralPath $IconPath) {
    $lnk.IconLocation = "$IconPath,0"
}
$lnk.Save()
Write-Host "Created Desktop shortcut: $shortcutPath (icon: $IconPath)"
