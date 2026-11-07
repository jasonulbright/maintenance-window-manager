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

# The parameter names of one command call: the ones written as -Name, and the keys of a splatted
# hashtable (@params) that the same function or script block builds with a literal or with
# $params['Key'] / $params.Key assignments.
function Get-CommandParameterName {
    param([Parameter(Mandatory)][System.Management.Automation.Language.CommandAst]$Command)
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($e in $Command.CommandElements) {
        if ($e -is [System.Management.Automation.Language.CommandParameterAst]) { $names.Add($e.ParameterName) }
        elseif ($e -is [System.Management.Automation.Language.VariableExpressionAst] -and $e.Splatted) {
            $var = $e.VariablePath.UserPath
            $scope = $Command.Parent
            while ($scope -and $scope -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and $scope -isnot [System.Management.Automation.Language.ScriptBlockAst]) { $scope = $scope.Parent }
            if (-not $scope) { continue }
            $assigns = $scope.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
            foreach ($a in $assigns) {
                $left = $a.Left
                if ($left -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.VariablePath.UserPath -eq $var -and $a.Right.Expression -is [System.Management.Automation.Language.HashtableAst]) {
                    foreach ($kv in $a.Right.Expression.KeyValuePairs) { if ($kv.Item1 -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $names.Add($kv.Item1.Value) } }
                }
                elseif ($left -is [System.Management.Automation.Language.IndexExpressionAst] -and $left.Target -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Target.VariablePath.UserPath -eq $var -and $left.Index -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Index.Value)
                }
                elseif ($left -is [System.Management.Automation.Language.MemberExpressionAst] -and $left.Expression -is [System.Management.Automation.Language.VariableExpressionAst] -and $left.Expression.VariablePath.UserPath -eq $var -and $left.Member -is [System.Management.Automation.Language.StringConstantExpressionAst]) {
                    $names.Add($left.Member.Value)
                }
            }
        }
    }
    return @($names | Select-Object -Unique)
}

# Lists the ConfigMgr cmdlet calls of a script file with the parameter names of each call.
function Get-CmCmdletCall {
    param([Parameter(Mandatory)][string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $calls = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -match '^(Get|Set|New|Remove|Start|Invoke|Add|Copy|Convert|Export|Import)-CM' -and $n.GetCommandName() -ne 'Get-CMConnectionInfo' }, $true)
    foreach ($c in $calls) {
        @{ Command = $c.GetCommandName(); Parameters = @(Get-CommandParameterName -Command $c); Line = $c.Extent.StartLineNumber }
    }
}
