# Shared helpers for the Maintenance Window Manager tests. Dot-source from BeforeAll.

$script:AppRoot = Split-Path -Parent $PSScriptRoot

function Import-AppModule {
    Import-Module (Join-Path $script:AppRoot 'Module\MaintWindowMgrCommon.psd1') -Force -DisableNameChecking
}

# Defines a global stub function for each named ConfigMgr command, so Pester can
# mock it the same way on every machine. A function beats a cmdlet in command
# lookup, so the stub also shadows the real cmdlet on a machine that has the
# ConfigMgr console. Call Remove-CmStub afterwards so the stubs never outlive
# the test file. The stub declares the parameters that the code under test and
# the tests use; the real signatures are checked in CmCmdlet.Contract.Tests.ps1.
$script:CmStubDefinition = {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]$InputObject,
        $CollectionId, $Id, $Name, $NewName, $ResourceId, $RuleName, $QueryExpression,
        $IncludeCollectionId, $ExcludeCollectionId, $LimitingCollectionName, $RefreshType, $RefreshSchedule,
        $Comment, $Query, $Option, $Expression, $TargetClassName, $LimitToCollectionId, $QueryId,
        $PackageId, $DistributionPointName, $RecurInterval, $RecurCount, $Start, $End, $Duration,
        $DurationInterval, $DurationCount, $DayOfWeek, $WeekOrder, $DayOfMonth, $MaintenanceWindowName, $ApplyTo, $IsEnabled, $IsUtc, $Schedule, $ScheduleString, $CollectionName,
        [switch]$Nonrecurring,
        [switch]$Fast, [switch]$Force, [switch]$PassThru, [switch]$DisableWildcardHandling
    )
    process { }
}

function New-CmStub {
    param(
        [Parameter(Mandatory)][string[]]$Name,
        [scriptblock]$Definition = $script:CmStubDefinition
    )
    foreach ($n in $Name) { Set-Item -Path "function:global:$n" -Value $Definition }
    return $Name
}

function Remove-CmStub {
    param([string[]]$Name)
    foreach ($n in $Name) { Remove-Item -Path "function:global:$n" -ErrorAction SilentlyContinue }
}

function New-RollupTable {
    # DataTable with the column names the shell passes to the export functions.
    param([string[]]$Column, [object[][]]$Row)
    $t = New-Object System.Data.DataTable
    foreach ($c in $Column) { [void]$t.Columns.Add($c) }
    foreach ($r in $Row) { [void]$t.Rows.Add($r) }
    # The comma keeps PowerShell from unrolling the table into its rows.
    return , $t
}
