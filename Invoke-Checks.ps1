#Requires -Version 5.1
<#
.SYNOPSIS
    Runs the linter, the tests and the code coverage of Maintenance Window Manager, and
    fails below the coverage floor.

.DESCRIPTION
    One command for developers and for CI:
      1. PSScriptAnalyzer on the module, the script and the tests. Any finding of
         severity Error or Warning fails the run.
      2. Pester on Module\ and Tests\, with code coverage of the module files.
         A failed test fails the run. A run that finds no test fails too.
      3. The coverage of the module files must reach -CoverageFloor percent.
    The summary lists the test count as run (each data row counts as one test)
    and the coverage. Lab tests live in Tests\Lab and never run here.

.PARAMETER CoverageFloor
    Minimum percent of module commands that the tests run. Default 80.

.PARAMETER SkipWindow
    Leave out Tests\Window.Tests.ps1, which opens windows and needs a desktop.

.PARAMETER SkipLint
    Leave out PSScriptAnalyzer.

.EXAMPLE
    .\Invoke-Checks.ps1
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console harness: Write-Host is the output channel.')]
[CmdletBinding()]
param(
    [ValidateRange(0, 100)][int]$CoverageFloor = 80,
    [switch]$SkipWindow,
    [switch]$SkipLint
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$failed = $false

Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop

# ---- Window test guard -----------------------------------------------------
# Window tests open a window only when SUITE_ALLOW_WINDOW_TESTS is 1.
# When they are skipped, Pester reports them as Skipped. This script prints one line and passes.
# The coverage floor does not depend on them: they run no module code.
. (Join-Path $root 'Tests\WindowHarness.ps1')
$windowSkipReported = $false
$windowGuard = Test-WindowTestsAllowed
if (-not $SkipWindow -and -not $windowGuard.Allowed) {
    Write-Host 'Window tests skipped: SUITE_ALLOW_WINDOW_TESTS is not 1.'
    $windowSkipReported = $true
}

# ---- 1. Lint ---------------------------------------------------------------
if (-not $SkipLint) {
    Import-Module PSScriptAnalyzer -ErrorAction Stop
    $targets = @(
        Get-ChildItem -LiteralPath (Join-Path $root 'Module') -Include '*.psm1', '*.psd1', '*.ps1' -File -Recurse
        Get-ChildItem -LiteralPath $root -Filter 'start-*.ps1' -File
        Get-ChildItem -LiteralPath (Join-Path $root 'Tests') -Filter '*.ps1' -File -Recurse
        Get-Item -LiteralPath (Join-Path $root 'Invoke-Checks.ps1')
    )
    $settings = Join-Path $root 'Tests\PSScriptAnalyzerSettings.psd1'
    # Test files pass state to mock bodies through a global variable, and they define fake window controls that the code under test reads by name. Both are intended.
    $testOnlyExclude = 'PSAvoidGlobalVars', 'PSUseDeclaredVarsMoreThanAssignments'
    $findings = @(foreach ($t in $targets) {
        if ($t.FullName.Substring($root.Length + 1) -like 'Tests\*') { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings -ExcludeRule $testOnlyExclude }
        else { Invoke-ScriptAnalyzer -Path $t.FullName -Settings $settings }
    })
    $blocking = @($findings | Where-Object { $_.Severity -in 'Error', 'Warning' })
    Write-Host ('Lint: {0} files, {1} findings ({2} Error or Warning)' -f $targets.Count, $findings.Count, $blocking.Count)
    foreach ($f in $blocking) { Write-Host ('  {0}:{1} [{2}] {3}' -f (Split-Path $f.ScriptPath -Leaf), $f.Line, $f.RuleName, $f.Message) }
    if ($blocking.Count -gt 0) { $failed = $true }
}

# ---- 2. Tests and coverage -------------------------------------------------
$cfg = New-PesterConfiguration
$testFiles = @(Get-ChildItem -LiteralPath $root -Recurse -Filter '*.Tests.ps1' -File |
        Where-Object { $rel = $_.FullName.Substring($root.Length + 1); $rel -notlike 'Tests\Lab\*' -and $rel -notlike '.git\*' -and $rel -notlike '.claude\*' })
if ($testFiles.Count -eq 0) { Write-Host 'No test file found.'; exit 1 }
if ($SkipWindow) { $testFiles = @($testFiles | Where-Object { $_.Name -ne 'Window.Tests.ps1' }) }
$cfg.Run.Path = $testFiles.FullName
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'Minimal'
$cfg.CodeCoverage.Enabled = $true
$cfg.CodeCoverage.OutputPath = Join-Path ([System.IO.Path]::GetTempPath()) 'mwm-coverage.xml'
$cfg.CodeCoverage.Path = @(Get-ChildItem -LiteralPath (Join-Path $root 'Module') -Filter '*.psm1' -File | ForEach-Object FullName)
# Window.Tests.ps1 prints the same reason when it runs alone. The marker keeps it quiet during
# this run only, and the previous value comes back afterwards, so a later standalone run is not muted.
$previousMarker = $env:SUITE_WINDOW_GUARD_REPORTED
if ($windowSkipReported) { $env:SUITE_WINDOW_GUARD_REPORTED = '1' }
try { $result = Invoke-Pester -Configuration $cfg }
finally { $env:SUITE_WINDOW_GUARD_REPORTED = $previousMarker }

$cov = $result.CodeCoverage
# Keep the unrounded value for the floor: 79.96 must not round up to 80.
$pct = if ($cov -and $cov.CommandsAnalyzedCount -gt 0) { 100 * $cov.CommandsExecutedCount / $cov.CommandsAnalyzedCount } else { 0 }
Write-Host ('Tests: {0} run, {1} passed, {2} failed, {3} skipped' -f $result.TotalCount, $result.PassedCount, $result.FailedCount, $result.SkippedCount)
Write-Host ('Coverage: {0:N1}% of {1} module commands (floor {2}%)' -f $pct, $cov.CommandsAnalyzedCount, $CoverageFloor)

if ($result.TotalCount -eq 0) { Write-Host 'No test ran.'; $failed = $true }
if ($result.FailedCount -gt 0) { $failed = $true }
if ($result.Result -ne 'Passed') { Write-Host ('Pester result: {0}' -f $result.Result); $failed = $true }
if ($pct -lt $CoverageFloor) { Write-Host ('Coverage {0:N1}% is below the floor of {1}%.' -f $pct, $CoverageFloor); $failed = $true }

if ($failed) { Write-Host 'CHECKS FAILED'; exit 1 }
Write-Host 'CHECKS PASSED'
exit 0
