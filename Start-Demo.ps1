param([switch]$NoBrowser)
$ErrorActionPreference = 'Stop'
$demoRoot = $PSScriptRoot
$demoUrl = 'http://127.0.0.1:8765'
try {
    $existing = Invoke-RestMethod -Uri "$demoUrl/api/config" -TimeoutSec 2
    if ($existing.app -eq 'macro-career-lab-v1') {
        if (-not $NoBrowser) { Start-Process $demoUrl }
        exit 0
    }
    throw 'Port 8765 is already used by a different service.'
} catch {
    if ($_.Exception.Message -like 'Port 8765*') { throw }
}
$pythonCandidates = @(
    (Join-Path $env:LOCALAPPDATA 'Programs/Python/Python313/python.exe'),
    (Join-Path $env:USERPROFILE '.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe')
)
$demoPython = $pythonCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $demoPython) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCommand) { $demoPython = $pythonCommand.Source }
}
if (-not $demoPython) { throw 'Python 3.10+ is required. See README.md.' }
& $demoPython -X utf8 -c 'import openpyxl'
if ($LASTEXITCODE -ne 0) { throw 'Install dependencies: python -m pip install -r requirements.txt' }
$runtimeDir = Join-Path $demoRoot 'runtime'
New-Item -ItemType Directory -Path $runtimeDir -Force | Out-Null
$appFile = Join-Path $demoRoot 'app.py'
$process = Start-Process -FilePath $demoPython -ArgumentList @('-X','utf8', ('"' + $appFile + '"')) -WorkingDirectory $demoRoot -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runtimeDir 'server.out.log') -RedirectStandardError (Join-Path $runtimeDir 'server.err.log')
$process.Id | Set-Content -LiteralPath (Join-Path $runtimeDir 'server.pid')
for ($attempt = 0; $attempt -lt 25; $attempt++) {
    Start-Sleep -Milliseconds 400
    try {
        $ready = Invoke-RestMethod -Uri "$demoUrl/api/config" -TimeoutSec 1
        if ($ready.app -eq 'macro-career-lab-v1') { if (-not $NoBrowser) { Start-Process $demoUrl }; exit 0 }
    } catch { }
}
throw 'Server did not start. Inspect runtime/server.err.log. Port 8765 may be occupied.'
