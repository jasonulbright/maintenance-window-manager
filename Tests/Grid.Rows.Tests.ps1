#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Data behind each grid of the main window. A grid shows an empty column when its XAML binds to a property name that
# the rows do not have, or when the property is empty for real data. The rows here come out of the real module
# functions (GridFixtures.ps1). Each bound column must exist on every row, and a value must show on at least one row. No window is opened.

BeforeDiscovery {
    . "$PSScriptRoot\GridData.Helpers.ps1"
    $script:Columns = @(Get-XamlGridColumn -XamlPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'MainWindow.xaml'))
    $script:ColumnCases = @($script:Columns | ForEach-Object { @{ Grid = $_.Grid; Header = $_.Header; Path = $_.Path } })
    $script:GridCases = @($script:Columns | Select-Object -ExpandProperty Grid -Unique | ForEach-Object { @{ Grid = $_ } })
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\GridFixtures.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName MaintWindowMgrCommon
    $script:Fixture = New-GridFixture
    $script:EmptyAllowed = Get-GridEmptyAllowed
}

Describe 'Grid data' {
    It 'has rows for the grid <Grid>' -ForEach $script:GridCases {
        $script:Fixture.ContainsKey($Grid) | Should -BeTrue -Because 'each grid of the XAML needs rows from the real module functions in GridFixtures.ps1'
        @($script:Fixture[$Grid]).Count | Should -BeGreaterThan 0
    }

    It 'fills column <Header> (<Path>) of the grid <Grid>' -ForEach $script:ColumnCases {
        $rows = @($script:Fixture[$Grid])
        foreach ($row in $rows) {
            $row.PSObject.Properties.Name | Should -Contain $Path -Because "the grid $Grid binds a column to $Path"
        }
        if ($script:EmptyAllowed.ContainsKey("$Grid.$Path")) { return }
        @($rows | Where-Object { -not [string]::IsNullOrEmpty([string]$_.$Path) }).Count | Should -BeGreaterThan 0 -Because "the column $Path of $Grid would show only empty cells"
    }
}
