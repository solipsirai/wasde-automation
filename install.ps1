$ErrorActionPreference = 'Stop'
$Root = $PSScriptRoot
$Python = Join-Path $Root '.venv\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $Python)) {
    & py -3.13 -m venv (Join-Path $Root '.venv')
    if ($LASTEXITCODE -ne 0) { throw 'Install Python 3.13 and the py launcher, then retry.' }
}
& $Python -m pip install --requirement (Join-Path $Root 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
& $Python -m pip check
if ($LASTEXITCODE -ne 0) { throw 'Python dependency check failed.' }

$Config = Join-Path $Root 'config.json'
if (-not (Test-Path -LiteralPath $Config)) {
    Copy-Item -LiteralPath (Join-Path $Root 'config.example.json') -Destination $Config
}
New-Item -ItemType Directory -Force -Path (Join-Path $Root 'workbooks') | Out-Null

foreach ($Test in 'test_wheat_corn.py', 'test_parser.py', 'test_monthly_snapshots.py') {
    & $Python (Join-Path $Root $Test)
    if ($LASTEXITCODE -ne 0) { throw "Regression test failed ($Test)." }
}
Write-Host 'Install complete. Copy the five workbooks and follow README.md.'
