# Publish public weekly records from this PC (runs publish_records.py).
#   Publish-Records.bat              first run asks for the folder and token, shows the weeks, then asks before upload
#   Publish-Records.bat -DryRun      only show what would change
#   Publish-Records.bat -Schedule    register a daily task (default 21:00, e.g. -At 07:30) that uploads changed weeks without asking
#   Publish-Records.bat -Unschedule  remove the daily task
param([switch]$Yes, [switch]$DryRun, [switch]$Schedule, [switch]$Unschedule, [string]$At = '21:00')
$ErrorActionPreference = 'Stop'
$taskName = 'MacroNotesPublish'
if ($Unschedule) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task $taskName."
    exit 0
}
if ($Schedule) {
    $argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Yes"
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument -WorkingDirectory $PSScriptRoot
    $trigger = New-ScheduledTaskTrigger -Daily -At $At
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Description 'macro-notes-automation: publish changed weekly records (publish_records.py --yes)' -Force | Out-Null
    Write-Host "Registered scheduled task $taskName (daily at $At). Log: $HOME\.macro-notes\publish.log"
    exit 0
}
$pythonCandidates = @(
    (Join-Path $env:LOCALAPPDATA 'Programs/Python/Python313/python.exe')
)
$python = $pythonCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $python) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) { $python = $pythonCommand.Source }
}
if (-not $python) { throw 'Python 3.10+ is required. See README.md.' }
& $python -X utf8 -c 'import openpyxl'
if ($LASTEXITCODE -ne 0) { throw 'Install dependencies: python -m pip install -r requirements.txt' }
$arguments = @('-X', 'utf8', (Join-Path $PSScriptRoot 'publish_records.py'))
if ($Yes) { $arguments += '--yes' }
if ($DryRun) { $arguments += '--dry-run' }
& $python @arguments
exit $LASTEXITCODE
