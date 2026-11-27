# Registers the daily Revenue Analysis job at 09:11 local time, every day.
$ErrorActionPreference = "Stop"

$TaskName = "RevenueAnalysisDailyETL"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectDir = Split-Path -Parent $ScriptDir
$BatPath = Join-Path $ScriptDir "run_daily_etl.bat"

Write-Host "Registering Windows Scheduled Task: $TaskName"
Write-Host "Working Directory: $ProjectDir"
Write-Host "Batch Executable: $BatPath"

$Action = New-ScheduledTaskAction -Execute "$env:SystemRoot\System32\cmd.exe" -Argument "/c scripts\run_daily_etl.bat" -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -Daily -At "09:11"
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Hours 2)
$Principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal -Description "Daily ETL and Revenue Dashboard refresh" -Force -ErrorAction Stop

Write-Host "Scheduled Task '$TaskName' registered successfully!"
