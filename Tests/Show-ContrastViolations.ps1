<#
.SYNOPSIS
    Prints the contrast violations of the main window in both themes, in the
    format of KnownContrastDefects.ps1.

.DESCRIPTION
    Run it when the window test reports a new or a fixed violation, to see the
    current list. Opens the window off screen at 1440x900, as Window.Tests.ps1 does.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console tool.')]
param()
. "$PSScriptRoot\WindowHarness.ps1"
. "$PSScriptRoot\WindowConfig.ps1"
try {
    foreach ($dark in $true, $false) {
        $r = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig } -Script {
            $cfg = $Arguments.Config
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            $out = @()
            foreach ($view in $cfg.Views) {
                foreach ($v in $cfg.Views) { $w.FindName($v.Container).Visibility = if ($v.Name -eq $view.Name) { 'Visible' } else { 'Collapsed' } }
                Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -Views $cfg.Views -OptionsButtonName $cfg.OptionsButton -LogLabelName $cfg.LogLabel -ActiveView $view.Name
                $out += [PSCustomObject]@{ Text = @(Get-TextContrast $w); Interactive = @(Get-InteractiveContrast $w); TitleBar = Get-TitleBarContrast $w }
            }
            $w.Close(); , $out
        }
        $keys = New-Object System.Collections.Generic.List[string]
        foreach ($v in $r) {
            foreach ($t in $v.Text) { $min = if ($t.Size -ge 24 -or ($t.Bold -and $t.Size -ge 18.66)) { 3.0 } else { 4.5 }; if ($t.Ratio -lt $min) { [void]$keys.Add(('text:{0}' -f $t.Text)) } }
            foreach ($c in $v.Interactive) {
                foreach ($m in 'EnabledText', 'DisabledText', 'EnabledBorder', 'DisabledBorder') {
                    $x = $c.$m; if ($null -eq $x) { continue }
                    $min = if ($m -like '*Text') { 4.5 } else { 3.0 }
                    if ($x -lt $min) { [void]$keys.Add(('{0}:{1}' -f $c.Name, $m)) }
                }
            }
        }
        if ($null -ne $r[0].TitleBar.Inactive -and $r[0].TitleBar.Inactive -lt 4.5) { [void]$keys.Add('titlebar:inactive') }
        Write-Host ("== {0}" -f $(if ($dark) { 'dark' } else { 'light' }))
        Write-Host (($keys | Sort-Object -Unique | ForEach-Object { "'" + ($_ -replace "'", "''") + "'" }) -join ', ')
    }
}
finally { Close-WpfSession }
