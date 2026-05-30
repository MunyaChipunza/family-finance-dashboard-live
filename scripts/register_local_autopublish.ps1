[CmdletBinding()]
param(
    [string]$TaskName = "Family Finance Dashboard Auto Publish",
    [string]$DataPath = "",
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerExe = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }
$triggerTime = (Get-Date).AddMinutes(1)

if ($WorkbookPath -and -not $DataPath) {
    Write-Host "WorkbookPath is retired. The scheduled task will watch dashboard_data.json instead."
}

if ($DataPath) {
    $candidateData = (Resolve-Path $DataPath).Path
} else {
    $candidateData = (Resolve-Path (Join-Path $PSScriptRoot "..\dashboard_data.json")).Path
}

$taskArgs = '"' + $runnerScriptPath + '" --data "' + $candidateData + '"'
$action = New-ScheduledTaskAction -Execute $runnerExe -Argument $taskArgs -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Checks dashboard_data.json and publishes the live finance dashboard when the data changes." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Data: $candidateData"
Write-Host "  Checks every 5 minutes on this PC."
