# Test copy only: make the module path lookup return nothing, as on a computer without the console.
$f = Join-Path $appDir 'Lib\SuiteCommon\SuiteCommon.psm1'
$t = Get-Content -LiteralPath $f -Raw
$patched = [regex]::Replace($t, '(function Resolve-ConfigurationManagerModulePath \{)', '${1} return $null;')
if ($patched -eq $t) { throw 'The module lookup function was not found. The missing-console test cannot run.' }
Set-Content -LiteralPath $f -Value $patched -Encoding UTF8
Remove-Item Env:SMS_ADMIN_UI_PATH -ErrorAction SilentlyContinue
$win = Start-App 'start-maintenancewindowmgr.ps1' 'Maintenance Window Manager'
Size-Win $win 1440 900
$win = Find-Top 'Maintenance Window Manager' 10
Invoke-El (Get-El $win 'btnRefresh')
# Wait for the log to show the failed refresh. A fixed sleep is not evidence.
if (-not (Wait-Until { (Get-ElText (Get-El $win 'txtLog')) -match 'Refresh failed: Failed to connect' } 30)) { throw 'The message Refresh failed did not show in the log.' }
Start-Sleep -Seconds 1
Shot $win 'console-missing-refresh'
Stop-App 'start-maintenancewindowmgr.ps1'
