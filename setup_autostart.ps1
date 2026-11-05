$ErrorActionPreference = 'Stop'

$workDir = 'C:\Users\AndronikLindgren\.gemini\antigravity\scratch\Projects\Trendcarpet_Interiors'
$scriptPath = Join-Path $workDir 'start_hub_with_tunnel.py'
$watchdogScript = Join-Path $workDir 'qa_watchdog.ps1'
$python = 'C:\Users\AndronikLindgren\miniconda3\python.exe'

Write-Host '[1/3] Registrerar schemalagd aktivitet i Windows Task Scheduler (Trendcarpet_QA_Hub)...'
$action = New-ScheduledTaskAction -Execute $python -Argument $scriptPath -WorkingDirectory $workDir
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Days 365) -MultipleInstances IgnoreNew
Register-ScheduledTask -TaskName 'Trendcarpet_QA_Hub' -Action $action -Trigger $trigger -Settings $settings -Description 'Trendcarpet Photo QA Hub Backend' -Force
Write-Host '  [OK] Schemalagd aktivitet skapad: Trendcarpet_QA_Hub (kor vid inloggning)'

Write-Host '[2/3] Registrerar schemalagd watchdog i Windows Task Scheduler (Trendcarpet_QA_Watchdog)...'
$watchdogAction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$watchdogScript`"" -WorkingDirectory $workDir
$watchdogTrigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 5)
$watchdogSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 3) -MultipleInstances IgnoreNew -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries
Register-ScheduledTask -TaskName 'Trendcarpet_QA_Watchdog' -Action $watchdogAction -Trigger $watchdogTrigger -Settings $watchdogSettings -Description 'Trendcarpet Photo QA Hub Watchdog (var 5:e minut)' -Force
Write-Host '  [OK] Schemalagd watchdog skapad: Trendcarpet_QA_Watchdog (kontrollerar var 5:e minut)'

Write-Host '[3/3] Uppdaterar autostart i Startup-mappen...'
$startupFolder = [System.Environment]::GetFolderPath('Startup')
$vbsPath = Join-Path $startupFolder 'Start_Foto_QA_Hub.vbs'
$v1 = 'Set WshShell = CreateObject("WScript.Shell")' + [Environment]::NewLine
$v2 = 'WshShell.CurrentDirectory = "' + $workDir + '"' + [Environment]::NewLine
$v3 = 'WshShell.Run """' + $python + '"" start_hub_with_tunnel.py", 0, False' + [Environment]::NewLine
[System.IO.File]::WriteAllText(
    $vbsPath,
    ($v1 + $v2 + $v3),
    [System.Text.Encoding]::ASCII
)
Write-Host ('  [OK] Uppdaterade: ' + $vbsPath)

Write-Host ''
Write-Host '=== Autostart & Watchdog ar nu 100% konfigurerat for Windows ==='
