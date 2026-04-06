[CmdletBinding()]
param(
    [string]$TaskName = "Family Finance Dashboard Auto Publish",
    [string]$WorkbookPath = ""
)

$ErrorActionPreference = "Stop"

$runnerScriptPath = (Resolve-Path (Join-Path $PSScriptRoot "run_local_autopublish.pyw")).Path
$pythonPath = (Get-Command python -ErrorAction Stop).Source
$pythonwPath = Join-Path (Split-Path $pythonPath) "pythonw.exe"
$runnerExe = if (Test-Path $pythonwPath) { $pythonwPath } else { $pythonPath }
$triggerTime = (Get-Date).AddMinutes(1)

if ($WorkbookPath) {
    $candidateWorkbook = (Resolve-Path $WorkbookPath).Path
} else {
    $candidateWorkbook = (Resolve-Path (Join-Path $PSScriptRoot "..\\..\\Finance.xlsx")).Path
}

$taskArgs = '"' + $runnerScriptPath + '" --workbook "' + $candidateWorkbook + '"'
$action = New-ScheduledTaskAction -Execute $runnerExe -Argument $taskArgs -WorkingDirectory $PSScriptRoot
$trigger = New-ScheduledTaskTrigger -Once -At $triggerTime -RepetitionInterval (New-TimeSpan -Minutes 2) -RepetitionDuration (New-TimeSpan -Days 3650)
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description "Checks the family finance workbook and publishes the live dashboard when the data changes." -Force | Out-Null

Write-Host "Scheduled task created:"
Write-Host "  Name: $TaskName"
Write-Host "  Workbook: $candidateWorkbook"
Write-Host "  Checks every 2 minutes on this PC."
