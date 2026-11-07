#Requires -Modules Pester
# Window tests. The main window and the dialogs that the script builds from XAML
# text are opened off screen on an STA thread, in both themes and at 1024x720 and
# 1440x900, and measured. The app-specific names live in WindowConfig.ps1.

BeforeDiscovery {
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Matrix = foreach ($dark in $true, $false) {
        foreach ($size in @(@(1024, 720), @(1440, 900))) {
            @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]) }
        }
    }
    $script:ViewCases = @($script:AppConfig.Views | ForEach-Object { @{ View = $_.Name } })
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $PSScriptRoot) $script:AppConfig.Shell), [ref]$null, [ref]$null)
    $script:NameCases = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq '$window' -and $n.Arguments.Count -eq 1
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique | ForEach-Object { @{ Name = $_ } })
    $script:DialogCases = @(for ($i = 0; $i -lt @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow' }, $true)).Count; $i++) { @{ Index = $i } })
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    . "$PSScriptRoot\WindowHarness.ps1"
    . "$PSScriptRoot\KnownContrastDefects.ps1"
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ShellPath = Join-Path $script:Root $script:AppConfig.Shell
    $script:XamlPath = Join-Path $script:Root $script:AppConfig.Xaml

    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $script:WindowNames = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Value -eq 'FindName' -and
                $n.Expression.Extent.Text -eq '$window' -and $n.Arguments.Count -eq 1
            }, $true) | ForEach-Object { $_.Arguments[0].Value } | Sort-Object -Unique)
    $script:DialogXaml = @($ast.FindAll({
                param($n)
                $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $n.Value -match '^\s*<Controls:MetroWindow'
            }, $true) | ForEach-Object { $_.Value })

    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = 1440; H = 900 },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = 1440; H = 900 })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $cfg = $Arguments.Config
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
            $out = @()
            foreach ($view in $cfg.Views) {
                foreach ($v in $cfg.Views) { $w.FindName($v.Container).Visibility = if ($v.Name -eq $view.Name) { 'Visible' } else { 'Collapsed' } }
                Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -Views $cfg.Views `
                    -OptionsButtonName $cfg.OptionsButton -LogLabelName $cfg.LogLabel -ActiveView $view.Name
                $out += [PSCustomObject]@{
                    View        = $view.Name
                    Text        = @(Get-TextContrast $w)
                    Clipped     = @(Get-ClippedText $w)
                    Offscreen   = @(Get-OffscreenControls $w)
                    Interactive = @(Get-InteractiveContrast $w)
                    TitleBar    = Get-TitleBarContrast $w
                }
            }
            $w.Close()
            , $out
        }
    }
}

AfterAll { Close-WpfSession }

Describe 'Main window XAML' {
    It 'loads as a MetroWindow with the app title and a minimum size that fits 1024x720' {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:AppConfig.Xaml } -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Xaml)
            $o = [PSCustomObject]@{ Type = $w.GetType().FullName; Title = $w.Title; MinW = $w.MinWidth; MinH = $w.MinHeight }
            $w.Close(); $o
        }
        $r.Type | Should -Be 'MahApps.Metro.Controls.MetroWindow'
        $r.Title | Should -Be $script:AppConfig.Title
        $r.MinW | Should -BeLessOrEqual 1024
        $r.MinH | Should -BeLessOrEqual 720
    }

    It 'declares each x:Name once' {
        $names = [regex]::Matches((Get-Content -LiteralPath $script:XamlPath -Raw), 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        ($names | Group-Object | Where-Object Count -gt 1 | ForEach-Object Name) | Should -BeNullOrEmpty
    }

    It 'finds the script look-ups of the window' {
        $script:WindowNames.Count | Should -BeGreaterThan 10
    }

    It 'has the control <Name> that the script looks up' -ForEach $script:NameCases {
        (Get-Content -LiteralPath $script:XamlPath -Raw) | Should -Match ('x:Name="{0}"' -f [regex]::Escape($Name))
    }

    It 'has each configured view container and sidebar button' {
        $xaml = Get-Content -LiteralPath $script:XamlPath -Raw
        foreach ($v in $script:AppConfig.Views) {
            $xaml | Should -Match ('x:Name="{0}"' -f $v.Container)
            $xaml | Should -Match ('x:Name="{0}"' -f $v.Button)
        }
    }

    It 'shows the first view and keeps the others hidden' {
        $r = Invoke-WpfSta -Arguments @{ Config = $script:AppConfig } -Script {
            $w = Open-TestWindow -XamlPath (Join-Path $Root $Arguments.Config.Xaml)
            $o = foreach ($v in $Arguments.Config.Views) { [PSCustomObject]@{ Name = $v.Name; Vis = [string]$w.FindName($v.Container).Visibility } }
            $w.Close(); , @($o)
        }
        $first = $script:AppConfig.Views[0].Name
        ($r | Where-Object Name -eq $first).Vis | Should -Be 'Visible'
        @($r | Where-Object { $_.Name -ne $first -and $_.Vis -ne 'Collapsed' }) | Should -BeNullOrEmpty
    }
}

Describe 'Dialog windows' {
    It 'finds the dialog XAML in the script' {
        $script:DialogXaml.Count | Should -BeGreaterThan 0
    }

    It 'parses dialog <Index> as a MetroWindow with named controls' -ForEach $script:DialogCases {
        $r = Invoke-WpfSta -Arguments @{ Xaml = $script:DialogXaml[$Index] } -Script {
            [xml]$x = $Arguments.Xaml
            $w = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
            $names = [regex]::Matches($Arguments.Xaml, 'x:Name="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
            [PSCustomObject]@{ Type = $w.GetType().Name; Missing = @($names | Where-Object { -not $w.FindName($_) }); Count = @($names).Count }
        }
        $r.Type | Should -Be 'MetroWindow'
        $r.Missing | Should -BeNullOrEmpty
        $r.Count | Should -BeGreaterThan 0
    }

    It 'gives every dialog a title that differs from the main window title' {
        $titles = @([regex]::Matches((Get-Content -LiteralPath $script:ShellPath -Raw), "-Title\s+'([^']+)'") | ForEach-Object { $_.Groups[1].Value })
        $titles += @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        $titles.Count | Should -BeGreaterThan 1
        $titles | Should -Not -Contain $script:AppConfig.Title
        foreach ($t in $titles) { $t.Trim() | Should -Not -BeNullOrEmpty }
    }

    It 'gives the dialogs in XAML different titles' {
        $titles = @($script:DialogXaml | ForEach-Object { if ($_ -match '\sTitle="([^"]*)"') { $Matches[1] } } | Where-Object { $_ })
        ($titles | Select-Object -Unique).Count | Should -Be $titles.Count
    }
}

Describe 'Layout <Theme> <Size>' -ForEach $script:Matrix {
    It 'has no clipped text in the <View> view' -ForEach $script:ViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m.Clipped | ForEach-Object { '{0}: needs {1}px, has {2}px' -f $_.Text, $_.Needed, $_.Actual } | Should -BeNullOrEmpty
    }

    It 'keeps every button, box and list inside the window in the <View> view' -ForEach $script:ViewCases {
        $m = ($script:Measured['{0}|{1}' -f $Theme, $Size] | Where-Object View -eq $View)
        $m.Offscreen | ForEach-Object { '{0} at {1},{2} size {3}x{4}' -f $_.Name, $_.X, $_.Y, $_.W, $_.H } | Should -BeNullOrEmpty
    }
}

Describe 'Contrast <Theme>' -ForEach @(@{ Theme = 'dark' }, @{ Theme = 'light' }) {
    It 'shows every text at 4.5:1 (3:1 for large text) in the <View> view, except the listed known defects' -ForEach $script:ViewCases {
        $m = ($script:Measured['{0}|1440x900' -f $Theme] | Where-Object View -eq $View)
        $m.Text.Count | Should -BeGreaterThan 3
        $known = @(Get-KnownContrastDefect -Theme $Theme)
        $bad = $m.Text | Where-Object { $min = if ($_.Size -ge 24 -or ($_.Bold -and $_.Size -ge 18.66)) { 3.0 } else { 4.5 }; $_.Ratio -lt $min -and ($known -notcontains ('text:{0}' -f $_.Text)) }
        $bad | ForEach-Object { '{0} "{1}" {2} on {3} = {4}' -f $_.Name, $_.Text, $_.Foreground, $_.Backdrop, $_.Ratio } | Should -BeNullOrEmpty
    }
    It 'keeps the title text readable on the active title bar' {
        $tb = ($script:Measured['{0}|1440x900' -f $Theme] | Select-Object -First 1).TitleBar
        $tb.Active | Should -BeGreaterOrEqual 4.5
    }

    It 'meets the standard on every control, except the listed known defects' {
        $known = Get-KnownContrastDefect -Theme $Theme
        $violations = New-Object System.Collections.Generic.List[string]
        foreach ($v in ($script:Measured['{0}|1440x900' -f $Theme])) {
            foreach ($c in $v.Interactive) {
                foreach ($metric in 'EnabledText', 'DisabledText', 'EnabledBorder', 'DisabledBorder') {
                    $value = $c.$metric
                    if ($null -eq $value) { continue }
                    $min = if ($metric -like '*Text') { 4.5 } else { 3.0 }
                    if ($value -lt $min) { [void]$violations.Add(('{0}:{1}' -f $c.Name, $metric)) }
                }
            }
        }
        foreach ($v in ($script:Measured['{0}|1440x900' -f $Theme])) {
            foreach ($r in $v.Text) {
                $min = if ($r.Size -ge 24 -or ($r.Bold -and $r.Size -ge 18.66)) { 3.0 } else { 4.5 }
                if ($r.Ratio -lt $min) { [void]$violations.Add(('text:{0}' -f $r.Text)) }
            }
        }
        $tbi = ($script:Measured['{0}|1440x900' -f $Theme] | Select-Object -First 1).TitleBar.Inactive
        if ($null -ne $tbi -and $tbi -lt 4.5) { [void]$violations.Add('titlebar:inactive') }
        $found = @($violations | Sort-Object -Unique)
        $new = @($found | Where-Object { $known -notcontains $_ })
        $fixed = @($known | Where-Object { $found -notcontains $_ })
        $new | Should -BeNullOrEmpty -Because 'a new contrast violation appeared'
        $fixed | Should -BeNullOrEmpty -Because 'a known defect is fixed: remove it from Get-KnownContrastDefect'
    }
}
