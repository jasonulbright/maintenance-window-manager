# Realistic rows for the grid data tests. The rows come out of the real module functions, which read ConfigMgr through
# stubs of the ConfigMgr commands that return data in the shape of the real commands, and out of the row functions of
# start-maintenancewindowmgr.ps1 (cut from its syntax tree, as in Shell.Logic.Tests.ps1). No row is written by hand, so
# a property name that the code stops producing shows in the grid tests.
# Call it where Write-Log of the module is mocked and the module is imported.

function New-GridFixture {
    $root = Split-Path -Parent $PSScriptRoot
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'start-maintenancewindowmgr.ps1'), [ref]$null, [ref]$null)
    foreach ($n in 'Get-WindowStatusGlyph', 'Get-CoverageGlyph', 'ConvertTo-WindowGridRows', 'ConvertTo-CoverageGridRows', 'ConvertTo-TemplateGridRows') {
        $def = $ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
    $stubs = @('Invoke-CMWmiQuery', 'Get-CMMaintenanceWindow')
    try {
        Set-Item -Path 'function:global:Invoke-CMWmiQuery' -Value {
            [CmdletBinding()] param($Query, $Option)
            if ($Query -like '*SMS_CollectionSettings*') {
                [PSCustomObject]@{ CollectionID = 'MCM00001' }
                [PSCustomObject]@{ CollectionID = 'MCM00002' }
            }
            else {
                [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'Servers Production'; MemberCount = 120 }
                [PSCustomObject]@{ CollectionID = 'MCM00002'; Name = 'Workstations Pilot'; MemberCount = 450 }
                [PSCustomObject]@{ CollectionID = 'MCM00003'; Name = 'Kiosks'; MemberCount = 8 }
                [PSCustomObject]@{ CollectionID = 'SMS00001'; Name = 'All Systems'; MemberCount = 578 }
            }
        }
        Set-Item -Path 'function:global:Get-CMMaintenanceWindow' -Value {
            [CmdletBinding()] param($CollectionId)
            $start = (Get-Date).Date.AddDays(-30).AddHours(2)
            if ($CollectionId -eq 'MCM00001') {
                [PSCustomObject]@{ Name = 'Sunday night'; ServiceWindowID = 'W1'; ServiceWindowType = 4; RecurrenceType = 3; StartTime = $start; Duration = 270; IsEnabled = $true; IsGMT = $false; Description = 'Weekly patch window' }
                [PSCustomObject]@{ Name = 'Emergency'; ServiceWindowID = 'W2'; ServiceWindowType = 1; RecurrenceType = 2; StartTime = $start; Duration = 60; IsEnabled = $false; IsGMT = $true; Description = 'Daily' }
            }
            if ($CollectionId -eq 'MCM00002') {
                [PSCustomObject]@{ Name = 'Pilot ring'; ServiceWindowID = 'W3'; ServiceWindowType = 5; RecurrenceType = 2; StartTime = $start; Duration = 120; IsEnabled = $true; IsGMT = $false; Description = 'Task sequences' }
            }
        }
        $windows = @(Get-AllMaintenanceWindows)
        $collections = @(Get-DeviceCollectionSummary)
        $templates = @(Get-WindowTemplates -TemplatesPath (Join-Path $root 'Templates'))
        @{
            gridWindows   = @(ConvertTo-WindowGridRows -Windows $windows)
            gridCoverage  = @(ConvertTo-CoverageGridRows -Collections $collections -Windows $windows)
            gridTemplates = @(ConvertTo-TemplateGridRows -Templates $templates)
        }
    }
    finally { foreach ($s in $stubs) { Remove-Item -Path "function:global:$s" -ErrorAction SilentlyContinue } }
}

# Columns that a fixture cannot fill on any row, each with the reason. The grid test accepts only these.
function Get-GridEmptyAllowed {
    @{}
}
