param([string]$App)
$ErrorActionPreference = 'Stop'
$OutDir = 'C:\guide\ui\out'; New-Item -ItemType Directory $OutDir -Force | Out-Null
$appDir = Join-Path 'C:\guide' $App
try {
    . C:\guide\ui\guilib.ps1
    # Settings that the orchestrator passes in (site code, SMS provider).
    $Capture = Get-Content -LiteralPath C:\guide\ui\capture.json -Raw | ConvertFrom-Json
    . C:\guide\ui\steps.ps1
}
catch {
    # Log may not exist when guilib.ps1 failed, so write the line directly.
    ('{0}  ERROR: {1} @ {2}' -f (Get-Date -Format 'HH:mm:ss'), $_.Exception.Message, $_.ScriptStackTrace) | Add-Content -LiteralPath (Join-Path $OutDir 'run.log')
}
finally { 'finished' | Set-Content (Join-Path $OutDir 'finished.flag') }
