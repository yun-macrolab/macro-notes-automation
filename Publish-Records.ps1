# Publish public weekly records from this PC (runs publish_records.py).
#   Publish-Records.bat              first run asks for the folder and token, shows the weeks, then asks before upload
#   Publish-Records.bat -DryRun      only show what would change
#   Publish-Records.bat -Schedule    check every 30 minutes from 13:10 (customize -At / -EveryMinutes)
#   Publish-Records.bat -Unschedule  remove the daily task
param([switch]$Yes, [switch]$DryRun, [switch]$Schedule, [switch]$Unschedule, [string]$At = '13:10', [int]$EveryMinutes = 30)
$ErrorActionPreference = 'Stop'
$taskName = 'MacroNotesPublish'
$logPath = Join-Path $HOME '.macro-notes\publish.log'
if ($Unschedule) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "Removed scheduled task $taskName."
    exit 0
}
if ($Schedule) {
    $argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`" -Yes"
    $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argument -WorkingDirectory $PSScriptRoot
    $trigger = New-ScheduledTaskTrigger -Daily -At $At
    if ($EveryMinutes -lt 5 -or $EveryMinutes -gt 1440) { throw 'EveryMinutes must be between 5 and 1440.' }
    $repeated = New-ScheduledTaskTrigger -Once -At $At -RepetitionInterval (New-TimeSpan -Minutes $EveryMinutes) -RepetitionDuration (New-TimeSpan -Days 1)
    $trigger.Repetition = $repeated.Repetition
    # Catch up on wake, run on battery, retry failures, stop a stuck run after 10 minutes.
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -MultipleInstances IgnoreNew -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
    Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description 'macro-notes-automation: publish changed weekly records (publish_records.py --yes)' -Force | Out-Null
    Write-Host "Registered scheduled task $taskName (every $EveryMinutes minutes, starting at $At). Log: $logPath"
    exit 0
}
# A modal dialog would block retries. Keep a local failure marker until a successful run.
function Show-Failure([string]$text) {
    if (-not $Yes) { return }
    $directory = Split-Path $logPath
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $directory 'publish-error.txt') -Value "$(Get-Date -Format o) $text" -Encoding UTF8
}
try {
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
    if ($Yes) {
        & $python @arguments 1>$null
    } else {
        & $python @arguments
    }
    $code = $LASTEXITCODE
} catch {
    Show-Failure "Publishing weekly records failed: $($_.Exception.Message)"
    throw
}
if ($code -ne 0) { Show-Failure "Publishing weekly records failed. See $logPath" }
elseif ($Yes) { Remove-Item -LiteralPath (Join-Path (Split-Path $logPath) 'publish-error.txt') -ErrorAction SilentlyContinue }
exit $code
