param([string]$GuideRoot)
$ErrorActionPreference = 'Stop'
$ui = Join-Path $GuideRoot 'ui'
$OutDir = Join-Path $ui 'out'; New-Item -ItemType Directory $OutDir -Force | Out-Null
$appDir = $GuideRoot
$status = 'finished'
try {
    . (Join-Path $ui 'guilib.ps1')
    # Settings that the orchestrator passes in (site code, SMS provider).
    $Capture = Get-Content -LiteralPath (Join-Path $ui 'capture.json') -Raw | ConvertFrom-Json
    . (Join-Path $ui 'steps.ps1')
}
catch {
    # Log may not exist when guilib.ps1 failed, so write the line directly.
    $status = 'failed: ' + $_.Exception.Message
    try { ('{0}  ERROR: {1} @ {2}' -f (Get-Date -Format 'HH:mm:ss'), $_.Exception.Message, $_.ScriptStackTrace) | Add-Content -LiteralPath (Join-Path $OutDir 'run.log') } catch { $null = $_ }
}
finally { $status | Set-Content (Join-Path $OutDir 'finished.flag') }
