# Registers the daily Revenue Analysis job at 08:30 local time, every day.
$ErrorActionPreference = "Stop"

$TaskName = "RevenueAnalysisDailyETL"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$BatPath = Join-Path $ScriptDir "run_daily_etl.bat"

Write-Host "Registering Windows Scheduled Task: $TaskName"
Write-Host "Working Directory: $ProjectDir"
Write-Host "Batch Executable: $BatPath"

$Action = New-ScheduledTaskAction -Execute $BatPath -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -Daily -At "08:30"
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description "Daily ETL and Revenue Dashboard refresh" -Force -ErrorAction Stop

Write-Host "Scheduled Task '$TaskName' registered successfully!"
