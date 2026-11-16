#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x test for the compact 1080p density of the main window (XAML only, no window is opened).
    The shared style values are asserted in one place. The window test Density.Window.Tests.ps1 measures them on screen.
#>

BeforeAll {
    $script:Xaml = [xml](Get-Content -LiteralPath (Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'MainWindow.xaml') -Raw)
    $script:Ns = New-Object System.Xml.XmlNamespaceManager($script:Xaml.NameTable)
    $script:Ns.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:Ns.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    function Get-StyleValue {
        param([string]$Key, [string]$Property)
        $style = $script:Xaml.SelectSingleNode("//w:Style[@x:Key='$Key']", $script:Ns)
        $style | Should -Not -BeNullOrEmpty -Because $Key
        $setter = $style.SelectSingleNode("w:Setter[@Property='$Property']", $script:Ns)
        if ($setter) { $setter.GetAttribute('Value') } else { $null }
    }
}

Describe 'Compact density at 1080p' {
    It 'opens at the design size of a maximized 1080p work area and keeps the 1024x640 minimum' {
        $w = $script:Xaml.DocumentElement
        [int]$w.GetAttribute('Width') | Should -Be 1920
        [int]$w.GetAttribute('Height') | Should -Be 1032
        [int]$w.GetAttribute('MinWidth') | Should -BeLessOrEqual 1024
        [int]$w.GetAttribute('MinHeight') | Should -BeLessOrEqual 640
    }

    It 'sizes the sidebar view buttons between 34 and 36 px' {
        [double](Get-StyleValue 'ViewButtonStyle' 'Height') | Should -BeIn 34, 35, 36
    }

    It 'sizes the inline buttons and the single-line boxes between 26 and 28 px' {
        [double](Get-StyleValue 'SquareInlineButtonStyle' 'Height') | Should -BeGreaterOrEqual 26
        [double](Get-StyleValue 'SquareInlineButtonStyle' 'Height') | Should -BeLessOrEqual 28
        foreach ($tb in @($script:Xaml.SelectNodes('//w:TextBox[@Height] | //w:ComboBox[@Height]', $script:Ns))) {
            [double]$tb.GetAttribute('Height') | Should -BeIn 26, 27, 28 -Because $tb.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }

    It 'gives every grid rows of 24 to 26 px and a 26 px column header from the one shared style' {
        [double](Get-StyleValue 'ViewDataGrid' 'RowHeight') | Should -BeGreaterOrEqual 24
        [double](Get-StyleValue 'ViewDataGrid' 'RowHeight') | Should -BeLessOrEqual 26
        [double](Get-StyleValue 'ViewDataGrid' 'ColumnHeaderHeight') | Should -Be 26
        foreach ($g in @($script:Xaml.SelectNodes('//w:DataGrid', $script:Ns))) {
            $g.GetAttribute('RowHeight') | Should -BeNullOrEmpty -Because 'the row height has one place: ViewDataGrid'
            $g.GetAttribute('Style') | Should -Match 'ViewDataGrid' -Because $g.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        }
    }

    It 'gives no pane that holds a list, grid or text box a fixed height (the free height is shared with * rows)' {
        # Row definitions of the view grids. The log row and the 5 px splitter rows may stay fixed.
        $fixed = @($script:Xaml.SelectNodes("//w:Grid[starts-with(@x:Name,'view')]/w:Grid.RowDefinitions/w:RowDefinition[@Height and not(contains(@Height,'*')) and @Height != 'Auto' and @Height != '5']", $script:Ns))
        $fixed.Count | Should -Be 0 -Because ('fixed-height rows in a view: ' + (($fixed | ForEach-Object { $_.GetAttribute('Height') }) -join ', '))
    }
}
