#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard for the load robustness (Q-47): the load can be cancelled. It parses the files and opens no window.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ScriptPath = Join-Path $script:Root 'start-maintenancewindowmgr.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ScriptPath, [ref]$null, [ref]$null)
    function Get-FunctionText {
        param([string]$Name)
        $fn = $script:Ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        if (-not $fn) { throw "Function $Name not found" }
        return $fn.Extent.Text
    }
    $script:CmCall = 'Get-CM[A-Z]\w*|Invoke-CM\w+|Get-CimInstance|Get-WmiObject|Connect-CMSite'
}
Describe 'The load can be cancelled' {
    It 'has a Cancel button on the progress overlay and the overlay takes the click' {
        $script:XamlText = Get-Content -LiteralPath (Join-Path $script:Root 'MainWindow.xaml') -Raw
        $script:XamlText | Should -Match '<Button x:Name="btnCancelWork" Content="Cancel" Style="\{StaticResource ThemedButtonStyle\}"'
        $script:XamlText | Should -Match '<Border x:Name="progressOverlay"\s+Visibility="Collapsed"\s+IsHitTestVisible="True"'
    }

    It 'stops the work, drops the runspace and gives the window back when the button is clicked' {
        $text = Get-Content -LiteralPath $script:ScriptPath -Raw
        $text | Should -Match '\$btnCancelWork\s+= \$window\.FindName\(''btnCancelWork''\)'
        $text | Should -Match '\$btnCancelWork\.Add_Click\(\{ Invoke-CancelWork \}\)'
        $fn = Get-FunctionText -Name 'Invoke-CancelWork'
        $fn | Should -Match "\`$script:BgState\['Cancel'\] = \`$true"
        $fn | Should -Match 'Dispose-BgWork'
        $fn | Should -Match 'Close-SuiteBgRunspace -Runspace \$script:BgRunspace'
        $fn | Should -Match '\$progressOverlay\.Visibility = \[System\.Windows\.Visibility\]::Collapsed'
        $fn | Should -Not -Match $script:CmCall
    }
}