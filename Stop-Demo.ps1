$ErrorActionPreference = 'Stop'
$pidFile = Join-Path $PSScriptRoot 'runtime/server.pid'
if (-not (Test-Path -LiteralPath $pidFile)) { Write-Output 'No launcher PID found.'; exit 0 }
$demoPid = [int](Get-Content -LiteralPath $pidFile)
$processInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $demoPid"
$expectedApp = Join-Path $PSScriptRoot 'app.py'
if ($processInfo -and $processInfo.Name -eq 'python.exe' -and $processInfo.CommandLine.Contains($expectedApp)) {
    Stop-Process -Id $demoPid
    Write-Output 'Local demo server stopped.'
} elseif ($processInfo) {
    throw 'PID belongs to a different process; nothing was stopped.'
} else { Write-Output 'The demo server is already stopped.' }
