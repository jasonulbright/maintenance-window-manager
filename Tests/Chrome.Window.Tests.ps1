#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Window test for the chrome and the buttons of the main window and of every dialog, in both themes.
# Every dialog of the app is measured:
#   - the dialogs that the start script builds from XAML text (here-strings in the functions Show-...Dialog) are loaded,
#     given the theme the way the script gives it (Set-DialogTheme) and shown
#   - the dialogs that the start script takes from SuiteCommon (Show-ConfirmDialog, Show-CollectionPickerDialog) are
#     opened by the real function with the main window as owner, and measured and closed by a timer while they are open
# Rules:
#   - the dialog title bar brush (active and inactive) is the title bar brush of the main window
#   - the title bar text of the dialog is readable (4.5:1) on that brush
#   - every button has bold text (SemiBold or heavier, 600; a normal weight is a defect)
#   - no button has a border thicker than 1 px, and no button border has the text color of the theme

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:RepoDir = Split-Path -Parent $PSScriptRoot
    $shellPath = Join-Path $script:RepoDir $script:AppConfig.Shell
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($shellPath, [ref]$null, [ref]$null)

    # A string that holds the XAML of a dialog: a constant here-string or an expandable one ($titleText, $okText).
    $isDialogString = {
        param($n)
        ($n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) -and $n.Value -match '^\s*<Controls:MetroWindow'
    }
    $titles = @{ 'Show-NewWindowDialog' = 'New Maintenance Window'; 'Show-TemplateEditorDialog' = 'New Template' }
    $script:OwnDialogFunction = @()
    $dialogXaml = @()
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        foreach ($s in @($fn.Body.FindAll($isDialogString, $true))) {
            $script:OwnDialogFunction += $fn.Name
            $text = $s.Value
            if ($titles.ContainsKey($fn.Name)) { $text = $text.Replace('$titleText', $titles[$fn.Name]).Replace('$okText', 'Create') }
            $dialogXaml += $text.Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
        }
    }

    # The dialogs of SuiteCommon that the start script calls.
    $suiteAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoDir 'Lib\SuiteCommon\SuiteCommon.psm1'), [ref]$null, [ref]$null)
    $suiteDialog = @($suiteAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and @($n.Body.FindAll($isDialogString, $true)).Count -gt 0 }, $true) | ForEach-Object { $_.Name })
    $called = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Sort-Object -Unique)
    $script:SharedDialogFunction = @($suiteDialog | Where-Object { $called -contains $_ } | Sort-Object -Unique)

    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig; Dialogs = $dialogXaml; Shared = $script:SharedDialogFunction } -Script {
            $cfg = $Arguments.Config
            $main = $null
            try {
            $main = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellTheme -Window $main -Dark $Arguments.Dark -Views $cfg.Views -OptionsButtonName $cfg.OptionsButton -LogLabelName $cfg.LogLabel -ActiveView $cfg.Views[0].Name
            $mainChrome = Get-SuiteWindowChrome -Window $main -Label 'main'
            $mainBody = Get-SuiteWindowBody -Window $main -Label 'main'
            $out = [pscustomobject]@{
                ThemeText = $mainChrome.ThemeText
                MainChrome = $mainChrome; MainTitle = $mainChrome.Active; MainTitleInactive = $mainChrome.Inactive
                Main = @($mainChrome.Buttons); MainBody = $mainBody; Dialogs = @(); Unmeasured = @()
            }
            $left = if ($env:SUITE_VISIBLE_WINDOWS -eq '1') { 120 } else { -32000 }
            $i = 0
            foreach ($xaml in $Arguments.Dialogs) {
                $i++
                [xml]$x = $xaml
                $d = $null
                try {
                    $d = [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $x))
                    $d.WindowStartupLocation = 'Manual'
                    $d.Left = $left; $d.Top = $left
                    $d.ShowInTaskbar = $false
                    # The way the script gives the theme: no -IsDark, the theme context of the app answers.
                    Set-DialogTheme -Dialog $d
                    $d.Show(); Wait-Layout $d; Start-Sleep -Milliseconds 400; Wait-Layout $d
                    if ($i -eq 1) { $out | Add-Member -NotePropertyName MainInactive -NotePropertyValue (Get-SuiteWindowChrome -Window $main -Label 'main behind a dialog') }
                    $c = Get-SuiteWindowChrome -Window $d -Label ([string]$d.Title)
                    $body = Get-SuiteWindowBody -Window $d -Label ([string]$d.Title)
                    $out.Dialogs += [pscustomobject]@{ Body = $body; Title = $c.Title; Index = $i; Source = 'script'; Active = $c.Active; Inactive = $c.Inactive; TitleRatio = $c.TitleRatio; Buttons = @($c.Buttons) }
                }
                finally { if ($d) { $d.Close() } }
            }

            # SuiteCommon dialogs: the real function opens the dialog modal; a timer measures and closes it.
            $collections = @(
                [pscustomobject]@{ CollectionID = 'MCM00001'; Name = 'All Systems'; FolderID = 0; MemberCount = 3; IsBuiltIn = $false }
                [pscustomobject]@{ CollectionID = 'MCM00002'; Name = 'Servers'; FolderID = 1; MemberCount = 2; IsBuiltIn = $false }
            )
            $folders = @([pscustomobject]@{ FolderID = 1; ParentID = 0; Name = 'Maintenance' })
            foreach ($name in $Arguments.Shared) {
                $holder = @{ Chrome = $null; Body = $null; Error = $null }
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(700)
                $mainWindow = $main
                $timer.Add_Tick({
                        $timer.Stop()
                        try {
                            $d = @($mainWindow.OwnedWindows) | Where-Object { $_ -is [MahApps.Metro.Controls.MetroWindow] } | Select-Object -First 1
                            Wait-Layout $d
                            $holder.Chrome = Get-SuiteWindowChrome -Window $d -Label ([string]$d.Title)
                            $holder.Body = Get-SuiteWindowBody -Window $d -Label ([string]$d.Title)
                            $d.Close()
                        }
                        catch { $holder.Error = $_.ToString(); foreach ($o in @($mainWindow.OwnedWindows)) { $o.Close() } }
                    }.GetNewClosure())
                $watchdog = New-Object System.Windows.Threading.DispatcherTimer
                $watchdog.Interval = [TimeSpan]::FromSeconds(10)
                $watchdog.Add_Tick({ $holder.Error = "watchdog closed the dialog"; foreach ($o in @($mainWindow.OwnedWindows)) { $o.Close() }
                    }.GetNewClosure())
                $timer.Start(); $watchdog.Start()
                try {
                    switch ($name) {
                        'Show-ConfirmDialog' { [void](Show-ConfirmDialog -Title 'Remove Maintenance Window' -Message 'Permanently delete the window?' -Owner $main) }
                        'Show-CollectionPickerDialog' { [void](Show-CollectionPickerDialog -Owner $main -Collections $collections -Folders $folders -Title 'Pick Target Collection' -IncludeBuiltIn $false) }
                        default { $timer.Stop(); $holder.Error = 'no recipe for ' + $name }
                    }
                }
                finally { $timer.Stop(); $watchdog.Stop() }
                if ($holder.Chrome) {
                    $c = $holder.Chrome
                    $out.Dialogs += [pscustomobject]@{ Body = $holder.Body; Title = $c.Title; Index = 0; Source = $name; Active = $c.Active; Inactive = $c.Inactive; TitleRatio = $c.TitleRatio; Buttons = @($c.Buttons) }
                }
                else { $out.Unmeasured += ('{0}: {1}' -f $name, $holder.Error) }
            }
            $out
            }
            finally { if ($main) { $main.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Chrome <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:M = $script:Measured[$Theme] }

    It 'finds every dialog of the app: the dialogs of the script and the dialogs of SuiteCommon that it calls' {
        # One dialog per function Show-...Dialog of the script that builds a window.
        $funcs = @([regex]::Matches((Get-Content -LiteralPath (Join-Path $script:RepoDir $script:AppConfig.Shell) -Raw), '(?m)^function Show-\w*Dialog\b'))
        $funcs.Count | Should -BeGreaterThan 3
        @($script:OwnDialogFunction | Sort-Object -Unique).Count | Should -Be $funcs.Count
        @($script:M.Dialogs | Where-Object { $_.Source -eq 'script' }).Count | Should -Be $funcs.Count
        $script:SharedDialogFunction.Count | Should -BeGreaterThan 0
        @($script:M.Dialogs | Where-Object { $_.Source -ne 'script' }).Count | Should -Be $script:SharedDialogFunction.Count
        @($script:M.Unmeasured) | Should -BeNullOrEmpty
    }

    It 'gives every dialog the title bar brush of the main window' {
        $bad = @($script:M.Dialogs | Where-Object { $_.Active -ne $script:M.MainTitle -or $_.Inactive -ne $script:M.MainTitleInactive } |
                ForEach-Object { '{0}: active {1} inactive {2}, main {3} {4}' -f $_.Title, $_.Active, $_.Inactive, $script:M.MainTitle, $script:M.MainTitleInactive })
        $bad | Should -BeNullOrEmpty
    }

    It 'lists the body colors of the main window and of every dialog' {
        Write-Host ('{0} main: surface {1} text {2}' -f $Theme, $script:M.MainBody.BackLuminance, $script:M.MainBody.TextRatio)
        foreach ($d in $script:M.Dialogs) {
            Write-Host ('{0} {1}: surface {2} text {3}; inputs: {4}' -f $Theme, $d.Title, $d.Body.BackLuminance, $d.Body.TextRatio, ((@($d.Body.Inputs) | ForEach-Object { '{0} {1} lum {2} text {3} border {4}' -f $_.Kind, $_.Name, $_.BackLuminance, $_.TextRatio, $_.BorderRatio }) -join ' | '))
        }
        @($script:M.Dialogs).Count | Should -BeGreaterThan 0
    }

    It 'gives the body of every dialog the surface of the theme (light in light, dark in dark)' {
        $light = ($Theme -eq 'light')
        $bad = @($script:M.Dialogs | Where-Object { $null -ne $_.Body.BackLuminance -and ((($light -and $_.Body.BackLuminance -lt 0.5) -or (-not $light -and $_.Body.BackLuminance -gt 0.25))) } |
                ForEach-Object { '{0}: surface luminance {1} in the {2} theme' -f $_.Title, $_.Body.BackLuminance, $Theme })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives the body of every dialog the same surface as the main window' {
        $bad = @($script:M.Dialogs | Where-Object { $null -ne $_.Body.BackLuminance -and $null -ne $script:M.MainBody.BackLuminance -and [math]::Abs($_.Body.BackLuminance - $script:M.MainBody.BackLuminance) -gt 0.1 } |
                ForEach-Object { '{0}: surface {1}, main {2}' -f $_.Title, $_.Body.BackLuminance, $script:M.MainBody.BackLuminance })
        $bad | Should -BeNullOrEmpty
    }

    It 'keeps 4.5:1 for the text of every dialog body and every input' {
        $bad = @()
        foreach ($d in $script:M.Dialogs) {
            if ($null -ne $d.Body.TextRatio -and $d.Body.TextRatio -lt 4.5) { $bad += ('{0}: window text {1}' -f $d.Title, $d.Body.TextRatio) }
            foreach ($i in @($d.Body.Inputs)) { if ($null -ne $i.TextRatio -and $i.TextRatio -lt 4.5) { $bad += ('{0}: {1} {2} text {3}' -f $d.Title, $i.Kind, $i.Name, $i.TextRatio) } }
        }
        $bad | Should -BeNullOrEmpty
    }

    It 'gives every input of every dialog the surface side of the theme' {
        $light = ($Theme -eq 'light')
        $bad = @()
        foreach ($d in $script:M.Dialogs) { foreach ($i in @($d.Body.Inputs)) { if ($null -ne $i.BackLuminance -and ((($light -and $i.BackLuminance -lt 0.4) -or (-not $light -and $i.BackLuminance -gt 0.3)))) { $bad += ('{0}: {1} {2} luminance {3}' -f $d.Title, $i.Kind, $i.Name, $i.BackLuminance) } } }
        $bad | Should -BeNullOrEmpty
    }

    It 'shows the caption of the main window readable (4.5:1) on the bar of its state, active and behind a dialog' {
        $script:M.MainChrome.TitleRatio | Should -BeGreaterOrEqual 4.5
        $script:M.MainInactive.IsActive | Should -BeFalse -Because 'the dialog took the focus'
        $script:M.MainInactive.TitleRatio | Should -BeGreaterOrEqual 4.5
    }

    It 'gives the inactive light title bar blue with dark caption text, and the dark one white text' {
        if ($Theme -eq 'light') {
            $script:M.MainInactive.TitleBrushNow | Should -Be '#FF4BA3E0'
            $script:M.MainInactive.TitleText | Should -Be '#FF333333'
        }
        else { $script:M.MainInactive.TitleText | Should -Not -Be '#FF333333' }
        $script:M.MainChrome.TitleText | Should -Not -Be '#FF333333' -Because 'the active caption is white'
    }

    It 'keeps the title text of every dialog readable at 4.5:1' {
        $bad = @($script:M.Dialogs | Where-Object { $_.TitleRatio -lt 4.5 } | ForEach-Object { '{0}: {1}' -f $_.Title, $_.TitleRatio })
        $bad | Should -BeNullOrEmpty
    }

    It 'shows bold text on every button of the main window and of every dialog' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $all.Count | Should -BeGreaterThan 5
        $bad = @($all | Where-Object { $_.Weight -lt 600 } | ForEach-Object { '{0} "{1}" weight {2}' -f $_.Window, $_.Text, $_.Weight })
        $bad | Should -BeNullOrEmpty
    }

    It 'gives no button a border thicker than 1 px or in the text color of the theme' {
        $all = @($script:M.Main) + @($script:M.Dialogs | ForEach-Object { $_.Buttons })
        $bad = @($all | Where-Object { $_.BorderThickness -gt 1 -or ($_.BorderThickness -gt 0 -and $_.BorderColor -eq $script:M.ThemeText) } |
                ForEach-Object { '{0} "{1}" border {2}px {3}' -f $_.Window, $_.Text, $_.BorderThickness, $_.BorderColor })
        $bad | Should -BeNullOrEmpty
    }
}
