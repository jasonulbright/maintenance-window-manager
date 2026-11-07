#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x tests for the main window layout (XAML only, no window is opened).
#>

BeforeAll {
    $script:strXamlPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'MainWindow.xaml'
    $script:xmlWindow = [xml](Get-Content -LiteralPath $script:strXamlPath -Raw)
    $script:nsWpf = New-Object System.Xml.XmlNamespaceManager($script:xmlWindow.NameTable)
    $script:nsWpf.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $script:nsWpf.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
}

Describe 'Main window tabs' {
    It 'gives every tab header an explicit theme foreground so unselected tabs keep AA contrast' {
        $arrTab = @($script:xmlWindow.SelectNodes('//w:TabItem', $script:nsWpf))
        $arrTab.Count | Should -BeGreaterThan 0
        foreach ($nodeTab in $arrTab) {
            $nodeText = $nodeTab.SelectSingleNode('w:TabItem.Header/w:TextBlock', $script:nsWpf)
            $nodeText | Should -Not -BeNullOrEmpty -Because 'a plain Header string takes the dim MahApps tab color'
            $nodeText.GetAttribute('Foreground') | Should -BeExactly '{DynamicResource MahApps.Brushes.ThemeForeground}'
            $nodeText.GetAttribute('Style') | Should -BeExactly '{StaticResource TabHeaderText}'
        }
    }

    It 'makes the selected tab header bold with a trigger on IsSelected' {
        $nodeStyle = $script:xmlWindow.SelectSingleNode("//w:Style[@x:Key='TabHeaderText']", $script:nsWpf)
        $nodeStyle | Should -Not -BeNullOrEmpty
        $nodeTrigger = $nodeStyle.SelectSingleNode('w:Style.Triggers/w:DataTrigger', $script:nsWpf)
        $nodeStyle.SelectSingleNode("w:Setter[@Property='FontSize']", $script:nsWpf).GetAttribute('Value') | Should -BeExactly '13'
        $nodeTrigger.GetAttribute('Binding') | Should -Match 'IsSelected'
        $nodeTrigger.GetAttribute('Binding') | Should -Match 'RelativeSource=\{RelativeSource AncestorType=TabItem\}'
        $nodeTrigger.GetAttribute('Value') | Should -BeExactly 'True'
        $nodeTrigger.SelectSingleNode("w:Setter[@Property='FontWeight']", $script:nsWpf).GetAttribute('Value') | Should -BeExactly 'Bold'
    }
}

Describe 'Page subtitle' {
    It 'starts with the subtitle of the Windows view, and that subtitle has no hint about Refresh that stays after the connection' {
        $strRoot = Split-Path -Path $script:strXamlPath -Parent
        $strScript = Get-Content -LiteralPath (Join-Path -Path $strRoot -ChildPath 'start-maintenancewindowmgr.ps1') -Raw
        $strWindows = [regex]::Match($strScript, "'Windows'\s*=\s*@\{ Title = 'Windows';\s*Subtitle = '([^']+)'").Groups[1].Value
        $strWindows | Should -Not -BeNullOrEmpty
        $strWindows | Should -Not -Match 'Configure Site'
        $nodeSubtitle = $script:xmlWindow.SelectSingleNode("//w:TextBlock[@x:Name='txtModuleSubtitle']", $script:nsWpf)
        $nodeSubtitle.GetAttribute('Text') | Should -BeExactly $strWindows
    }
}
