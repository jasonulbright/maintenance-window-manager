#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the theme rules. They parse the files and create no WPF control, so they run on
# Windows PowerShell 5.1 and PowerShell 7 and open no window.
#   (a) Every blue background in XAML and PowerShell has a white foreground.
#   (b) No font size, weight, style or text decoration depends on the theme.
#   (c) No banned word is in a tooltip, a UI string, the README or the user guide.
#   (d) The site code and SMS Provider titles are bold 14 pt, the values are not.
#   (e) Every button takes the shared themed style, and no style holds a color literal.
#   (f) No raw exception text reaches the UI.

BeforeDiscovery {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:PsFilesD = @(
        'start-maintenancewindowmgr.ps1', 'Module\MaintWindowMgrCommon.psm1', 'Lib\SuiteCommon\SuiteCommon.psm1'
    ) | ForEach-Object { @{ Name = $_; Path = (Join-Path $script:RepoRoot $_) } }
    $script:XamlDocsD = @(Get-XamlDocument)
}

BeforeAll {
    . "$PSScriptRoot\ThemeGuardHelpers.ps1"
    $script:XamlDocs = @(Get-XamlDocument)
    $script:ShellName = 'start-maintenancewindowmgr.ps1'
}

Describe 'Blue backgrounds have white text (a)' {
    It 'finds the XAML of the main window, the four dialogs and the schedule editor' {
        $script:XamlDocs.Count | Should -BeGreaterOrEqual 6
        @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' }).Count | Should -Be 1
        @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" }).Count | Should -BeGreaterOrEqual 5
    }

    It 'sets a white foreground on every blue background in <_.Source>' -ForEach $script:XamlDocsD {
        $bad = @(Get-BlueWithoutWhite -Text $_.Text -Source $_.Source)
        $bad | Should -BeNullOrEmpty
    }

    It 'recognizes blue and white brushes' {
        Test-BlueBrush '{DynamicResource MahApps.Brushes.Accent}' | Should -BeTrue
        Test-BlueBrush '#0078D4' | Should -BeTrue
        Test-BlueBrush '#1E1E1E' | Should -BeFalse
        Test-BlueBrush '#3A3A3A' | Should -BeFalse
        Test-WhiteBrush '{DynamicResource MahApps.Brushes.IdealForeground}' | Should -BeTrue
        Test-WhiteBrush 'Black' | Should -BeFalse
    }

    It 'flags a blue background that has no white foreground' {
        $xaml = '<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Setter Property="Background" Value="#0078D4"/></Style>'
        @(Get-BlueWithoutWhite -Text $xaml -Source 'sample').Count | Should -Be 1
        $ok = '<Style xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" TargetType="Button"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/></Style>'
        @(Get-BlueWithoutWhite -Text $ok -Source 'sample').Count | Should -Be 0
        $attr = '<Button xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Background="{DynamicResource MahApps.Brushes.Accent}"/>'
        @(Get-BlueWithoutWhite -Text $attr -Source 'sample').Count | Should -Be 1
    }

    It 'sets white text on the selected state of <Target> in the main window' -ForEach @(
        @{ Target = 'ViewDataGridRow'; Kind = 'key' }
        @{ Target = 'GridCellStyle'; Kind = 'key' }
        @{ Target = 'ComboBoxItem'; Kind = 'type' }
    ) {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        $styles = @($x.SelectNodes("//*[local-name()='Style']") | Where-Object {
                if ($Kind -eq 'key') { $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') -eq $Target }
                else { $_.GetAttribute('TargetType') -eq $Target -and -not $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') }
            })
        $styles.Count | Should -BeGreaterThan 0
        foreach ($st in $styles) {
            $trigger = @($st.SelectNodes("*[local-name()='Style.Triggers']/*[local-name()='Trigger']") | Where-Object { $_.GetAttribute('Property') -in 'IsSelected', 'IsHighlighted' })
            $trigger.Count | Should -BeGreaterThan 0
            $fg = @($trigger[0].SelectNodes("*[local-name()='Setter']") | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' })
            $fg.Count | Should -Be 1
            Test-WhiteBrush (Get-SetterValue $fg[0]) | Should -BeTrue
        }
    }

    It 'sets white text on the selected state of the target list and on the highlighted combo items in each dialog' {
        $dialogs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' })
        $dialogs.Count | Should -BeGreaterOrEqual 4
        $listStyles = 0
        foreach ($d in $dialogs) {
            [xml]$x = $d.Text
            foreach ($type in 'ComboBoxItem', 'ListBoxItem') {
                $uses = ($type -eq 'ComboBoxItem' -and $x.SelectNodes("//*[local-name()='ComboBox']").Count -gt 0) -or ($type -eq 'ListBoxItem' -and $x.SelectNodes("//*[local-name()='ListBox']").Count -gt 0)
                if (-not $uses) { continue }
                $st = @($x.SelectNodes("//*[local-name()='Style']") | Where-Object { $_.GetAttribute('TargetType') -eq $type -and -not $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml') })
                $st.Count | Should -Be 1 -Because ('{0} needs a {1} style' -f $d.Source, $type)
                if ($type -eq 'ListBoxItem') { $listStyles++ }
                $setters = @($st[0].SelectNodes('.//*[local-name()="Trigger"]/*[local-name()="Setter"]') | Where-Object { $_.GetAttribute('Property') -eq 'Foreground' })
                $setters.Count | Should -Be 1
                Test-WhiteBrush (Get-SetterValue $setters[0]) | Should -BeTrue
            }
        }
        $listStyles | Should -BeGreaterOrEqual 1
    }

    It 'has no blue brush literal in <Name>, and each CSS rule with a blue background has white text' -ForEach @(
        @{ Name = 'start-maintenancewindowmgr.ps1'; Path = (Join-Path $script:RepoRoot 'start-maintenancewindowmgr.ps1') }
        @{ Name = 'MaintWindowMgrCommon.psm1'; Path = (Join-Path $script:RepoRoot 'Module\MaintWindowMgrCommon.psm1') }
    ) {
        $bad = foreach ($s in @(Get-ScriptString -Path $Path)) {
            if ($s -notmatch '<' -and (Test-BlueBrush $s)) { 'literal ' + $s }
            foreach ($m in [regex]::Matches($s, '\{([^}]*)\}')) {
                $rule = $m.Groups[1].Value
                $bg = [regex]::Match($rule, '(?i)(?<![-\w])background(?:-color)?\s*:\s*(#[0-9A-F]{6,8})\b')
                if ($bg.Success -and (Test-BlueBrush $bg.Groups[1].Value) -and $rule -notmatch '(?i)(?<![-\w])color\s*:\s*(#fff(fff)?|white)\b') { 'css ' + $rule }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'paints the light theme buttons of SuiteCommon with white text' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match "OnBlue\s*=\s*\`$bc\.ConvertFrom\('#FFFFFF'\)"
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Update-SidebarButtonTheme', 'Set-ButtonTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Match 'Set-ThemeForeground' -Because ('{0} sets the foreground with each background' -f $name)
        }
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Set-ThemeForeground' }, $true)
        $fn.Extent.Text | Should -Match '\.Foreground\s*=\s*\$OnBlue'
    }
}

Describe 'Fonts do not depend on the theme (b)' {
    It 'has no font property in a theme branch of <Name>' -ForEach $script:PsFilesD {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true)
        $bad = foreach ($a in $assigns) {
            $member = [string]$a.Left.Member.Value
            if ($script:ThemeProperty -notcontains $member) { continue }
            $p = $a.Parent
            while ($p) {
                $condText = ''
                if ($p -is [System.Management.Automation.Language.IfStatementAst]) { $condText = (@($p.Clauses | ForEach-Object { $_.Item1.Extent.Text }) -join ' ') }
                elseif ($p -is [System.Management.Automation.Language.SwitchStatementAst]) { $condText = $p.Condition.Extent.Text }
                if ($condText -match '(?i)IsDark|DarkMode|Light|Theme|IsOn') { '{0}:{1} sets {2} in a theme branch' -f $Name, $a.Extent.StartLineNumber, $member; break }
                $p = $p.Parent
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'detects a font property in a theme branch (the check itself)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('if ($isDark) { $b.Background = 1; $b.FontWeight = 2 } else { $b.Background = 3 }', [ref]$null, [ref]$null)
        $assigns = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] -and $script:ThemeProperty -contains [string]$n.Left.Member.Value }, $true))
        $assigns.Count | Should -Be 1
        $assigns[0].Parent.Parent | Should -BeOfType [System.Management.Automation.Language.IfStatementAst]
    }

    It 'sets no font property in the theme functions of SuiteCommon except the connection boxes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$null, [ref]$null)
        foreach ($name in 'Set-ThemeForeground', 'Update-SidebarButtonTheme', 'Set-ButtonTheme', 'Set-DialogTheme', 'Update-TitleBarBrushes', 'Initialize-SuiteTheme') {
            $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
            $fn | Should -Not -BeNullOrEmpty -Because $name
            $fn.Extent.Text | Should -Not -Match '\.(FontSize|FontWeight|FontStyle|TextDecorations)\s*=' -Because $name
        }
    }

    It 'makes the field titles bold and 14 pt or larger the same way in both themes' {
        $path = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $text = Get-Content -LiteralPath $path -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match '\.FontWeight\s*=\s*\[System\.Windows\.FontWeights\]::Bold'
    }

    It 'uses no theme trigger on a font property in <_.Source>' -ForEach $script:XamlDocsD {
        @(Get-ThemeFontTrigger -Text $_.Text) | Should -BeNullOrEmpty
    }

    It 'flags a trigger that sets a font for a theme, and not one that sets it for the selected state (the check itself)' {
        $ns = 'xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"'
        $theme = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsDark}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        $state = "<Style $ns TargetType=`"TextBlock`"><Style.Triggers><DataTrigger Binding=`"{Binding IsSelected, RelativeSource={RelativeSource AncestorType=TabItem}}`" Value=`"True`"><Setter Property=`"FontWeight`" Value=`"Bold`"/></DataTrigger></Style.Triggers></Style>"
        @(Get-ThemeFontTrigger -Text $theme).Count | Should -Be 1
        @(Get-ThemeFontTrigger -Text $state).Count | Should -Be 0
    }
}

Describe 'No banned word in the UI text and the documents (c)' {
    It 'finds the banned word <Word>' -ForEach @(@{ Word = 'simply' }, @{ Word = 'Helps you' }, @{ Word = 'streamlines' }) {
        $Word | Should -Match $script:BannedPattern
    }

    It 'has no banned word in the XAML text of <_.Source>' -ForEach $script:XamlDocsD {
        $bad = @(Get-XamlUiText -Text $_.Text | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the strings of <Name>' -ForEach $script:PsFilesD {
        $bad = @(Get-ScriptString -Path $Path | Where-Object { $_ -match $script:BannedPattern })
        $bad | Should -BeNullOrEmpty
    }

    It 'has no banned word in the template files' {
        $files = @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'Templates') -Filter '*.json' -File)
        $files.Count | Should -BeGreaterThan 0
        $bad = foreach ($f in $files) { if ((Get-Content -LiteralPath $f.FullName -Raw) -match $script:BannedPattern) { '{0}: {1}' -f $f.Name, $Matches[0] } }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has no banned word in the README or the user guide' {
        $files = @(Get-Item -LiteralPath (Join-Path $script:RepoRoot 'README.md')) + @(Get-ChildItem -LiteralPath (Join-Path $script:RepoRoot 'docs') -Filter '*.md' -File -Recurse)
        $files.Count | Should -BeGreaterThan 5
        $bad = foreach ($f in $files) {
            $n = 0
            foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
                $n++
                if ($line -match $script:BannedPattern) { '{0}:{1}: {2}' -f $f.Name, $n, $Matches[0] }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Theme helpers of SuiteCommon behave and are used (a, b, d)' {
    BeforeAll {
        $script:SuitePath = Join-Path $script:RepoRoot 'Lib\SuiteCommon\SuiteCommon.psm1'
        $script:SuiteAst = [System.Management.Automation.Language.Parser]::ParseFile($script:SuitePath, [ref]$null, [ref]$null)
        function Get-SuiteFunction {
            param([string]$Name)
            $script:SuiteAst.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        }
        function Get-CommandCount {
            param($Ast, [string]$Name)
            @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $Name }, $true)).Count
        }
        . ([scriptblock]::Create((Get-SuiteFunction 'Set-ThemeForeground').Extent.Text))
    }

    It 'Set-ThemeForeground sets the on-blue brush in light mode and clears it in dark mode' {
        $c = [pscustomobject]@{ Foreground = $null }
        Set-ThemeForeground -Control $c -IsDark $false -OnBlue 'White'
        $c.Foreground | Should -Be 'White'
        Set-ThemeForeground -Control $c -IsDark $true -OnBlue 'White'
        $c.Foreground | Should -BeNullOrEmpty
    }

    It '<Name> calls Set-ThemeForeground' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ThemeForeground') | Should -BeGreaterThan 0
    }

    It '<Name> calls Set-ConnectionTextTheme' -ForEach @(@{ Name = 'Set-ButtonTheme' }, @{ Name = 'Update-SidebarButtonTheme' }) {
        (Get-CommandCount (Get-SuiteFunction $Name) 'Set-ConnectionTextTheme') | Should -BeGreaterThan 0
    }

    It 'the script starts the theme and every dialog calls Set-DialogTheme' {
        $shell = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:RepoRoot $script:ShellName), [ref]$null, [ref]$null)
        (Get-CommandCount $shell 'Initialize-SuiteTheme') | Should -BeGreaterThan 0
        (Get-CommandCount $shell 'Update-SidebarButtonTheme') | Should -BeGreaterThan 0
        $dialogs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' }).Count
        $dialogs | Should -BeGreaterOrEqual 4
        (Get-CommandCount $shell 'Set-DialogTheme') | Should -BeGreaterOrEqual $dialogs
        (Get-SuiteFunction 'Set-DialogTheme') | Should -Not -BeNullOrEmpty
    }

    It 'sets a white foreground in every style based on an accent style, in <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $bad = foreach ($st in @($x.SelectNodes("//*[local-name()='Style']"))) {
            if ($st.GetAttribute('BasedOn') -match 'Accent') {
                $fg = $st.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) { 'Style {0} has no white Foreground' -f $st.GetAttribute('Key') }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has the site code and SMS Provider boxes with a bold 14 pt title and a normal value' {
        foreach ($name in 'txtSiteCode', 'txtSmsProvider') {
            $hits = @()
            foreach ($d in $script:XamlDocs) {
                [xml]$x = $d.Text
                $el = $x.SelectSingleNode(("//*[@*[local-name()='Name']='{0}']" -f $name))
                if ($el) { $hits += [pscustomobject]@{ Value = $el; Title = (Get-FieldTitleElement $el); Source = $d.Source } }
            }
            $hits.Count | Should -BeGreaterOrEqual 2 -Because ('{0} must exist in the main window and in the Options dialog' -f $name)
            foreach ($h in $hits) {
                $h.Title | Should -Not -BeNullOrEmpty -Because ('{0} needs a title TextBlock before it' -f $name)
                [double]$h.Title.GetAttribute('FontSize') | Should -BeGreaterOrEqual 14 -Because ('title of {0} in {1}' -f $name, $h.Source)
                $h.Title.GetAttribute('FontWeight') | Should -Be 'Bold' -Because ('title of {0} in {1}' -f $name, $h.Source)
                $h.Value.GetAttribute('FontWeight') | Should -BeIn @('', 'Normal') -Because ('value {0}' -f $name)
                if ($h.Value.GetAttribute('FontSize')) { [double]$h.Value.GetAttribute('FontSize') | Should -BeLessThan 14 -Because ('value {0}' -f $name) }
            }
        }
    }

    It 'SuiteCommon bolds the titles with 14 pt and not the values' {
        $text = Get-Content -LiteralPath $script:SuitePath -Raw
        $text | Should -Match 'FieldTitleFontSize\s*=\s*14'
        $text | Should -Match 'Find-FieldTitle'
        $text | Should -Not -Match '\$box\.FontWeight'
    }

    It 'bolds the field titles of the dialogs and the main window' {
        # A field title is a TextBlock with FontSize 11 and the gray brush that sits above or before a box. It is bold.
        $bad = foreach ($d in $script:XamlDocs) {
            [xml]$x = $d.Text
            foreach ($t in @($x.SelectNodes("//*[local-name()='TextBlock'][@FontSize='11'][contains(@Foreground,'Gray1')][@Text]"))) {
                if ($t.GetAttribute('TextWrapping')) { continue }
                if ($t.GetAttribute('FontWeight') -ne 'Bold' -and $t.GetAttribute('Text') -notmatch '^(Author:|\()') { '{0}: {1}' -f $d.Source, $t.GetAttribute('Text') }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }
}

Describe 'Every button takes the shared themed style (e)' {
    BeforeAll {
        function Get-StyleChain {
            param($Doc, [string]$Key)
            $guard = 0
            while ($Key -and $guard -lt 10) {
                $guard++
                $st = $Doc.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $Key))
                if (-not $st) { return $false }
                $based = $st.GetAttribute('BasedOn')
                if ($based -match 'MahApps\.Styles\.Button') { return $true }
                $Key = if ($based -match 'StaticResource\s+([^\}\s]+)') { $Matches[1] } else { '' }
            }
            return $false
        }
        $script:ShellText = Get-Content -LiteralPath (Join-Path $script:RepoRoot $script:ShellName) -Raw
    }

    It 'gives each Button in <_.Source> a style that is based on a MahApps button style' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $bad = foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
            $style = $b.GetAttribute('Style')
            $key = if ($style -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { $Matches[1] } else { '' }
            if (-not $key -or -not (Get-StyleChain -Doc $x -Key $key)) { 'Button {0} has no shared style' -f $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'detects a button without a style (the check itself)' {
        [xml]$x = '<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Button x:Name="bare" Content="x"/></Grid>'
        $b = $x.SelectSingleNode("//*[local-name()='Button']")
        $b.GetAttribute('Style') | Should -BeNullOrEmpty
    }

    It 'creates no button in code' {
        foreach ($f in $script:ShellName, 'Module\MaintWindowMgrCommon.psm1') {
            (Get-Content -LiteralPath (Join-Path $script:RepoRoot $f) -Raw) | Should -Not -Match 'Controls\.Button\]|New-Object\s+(-TypeName\s+)?System\.Windows\.Controls\.Button' -Because $f
        }
    }

    It 'themes every button of the main window through SuiteCommon' {
        [xml]$x = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
        $names = @($x.SelectNodes("//*[local-name()='Button']") | ForEach-Object { $_.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') })
        $names.Count | Should -BeGreaterThan 10
        $wf = [regex]::Match($script:ShellText, '\$script:WorkflowButtons\s*=\s*@\(([^)]*)\)').Groups[1].Value
        $views = ([regex]::Matches($script:ShellText, 'Button\s*=\s*\$(\w+)') | ForEach-Object { $_.Groups[1].Value }) -join ' '
        $opt = [regex]::Match($script:ShellText, '-OptionsButton\s+\$(\w+)').Groups[1].Value
        $themed = ($wf + ' ' + $views + ' ' + $opt)
        $unthemed = $names | Where-Object { $themed -notmatch ('\b{0}\b' -f $_) }
        @($unthemed) | Should -BeNullOrEmpty
        $script:ShellText | Should -Match '-WorkflowButtons\s+\$script:WorkflowButtons'
        $script:ShellText | Should -Match '(?m)^Set-ButtonTheme'
        $script:ShellText | Should -Match '(?m)^\s+Set-ButtonTheme'
    }

    It 'themes every button of each dialog: an accent style or Set-DialogButtonTheme' {
        $docs = @($script:XamlDocs | Where-Object { $_.Source -like "$($script:ShellName) here-string*" -and $_.Text -match 'MetroWindow' })
        $docs.Count | Should -BeGreaterOrEqual 4
        $called = ([regex]::Matches($script:ShellText, 'Set-DialogButtonTheme\s+-Dialog\s+\$dlg\s+-Name\s+([^\r\n]+)') | ForEach-Object { $_.Groups[1].Value }) -join ' '
        $bad = foreach ($d in $docs) {
            [xml]$x = $d.Text
            foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
                $name = $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
                $style = $b.GetAttribute('Style')
                $key = if ($style -match 'StaticResource\s+([^\}\s]+)') { $Matches[1] } else { '' }
                $st = $x.SelectSingleNode(("//*[local-name()='Style'][@*[local-name()='Key']='{0}']" -f $key))
                $accent = $st -and $st.GetAttribute('BasedOn') -match 'Accent'
                if (-not $accent -and $called -notmatch ("'{0}'" -f $name)) { '{0}: {1} is not themed' -f $d.Source, $name }
            }
        }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'sets no color literal in any Style or ControlTemplate of <_.Source>' -ForEach $script:XamlDocsD {
        [xml]$x = $_.Text
        $nodes = @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]'))
        @(Get-HardColor -Nodes $nodes -Label $_.Source) | Should -BeNullOrEmpty
    }

    It 'scans the styles and templates of the main window (the scan is not empty)' {
        $doc = @($script:XamlDocs | Where-Object { $_.Source -eq 'MainWindow.xaml' })[0]
        [xml]$x = $doc.Text
        @($x.SelectNodes('//*[local-name()="Style"]')).Count | Should -BeGreaterThan 3
        @($x.SelectNodes('//*[local-name()="ControlTemplate"]')).Count | Should -BeGreaterThan 0
    }

    It 'detects a color literal in a style, a template and an attribute (the check itself)' {
        [xml]$x = '<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"><Style x:Key="S" TargetType="Button"><Setter Property="Foreground" Value="White"/><Setter Property="Background"><Setter.Value><SolidColorBrush Color="#1E1E1E"/></Setter.Value></Setter><Setter Property="BorderBrush" Value="{DynamicResource MahApps.Brushes.Accent}"/></Style><ControlTemplate x:Key="T" TargetType="Button"><Border BorderBrush="#555555" Background="{TemplateBinding Background}"/></ControlTemplate></ResourceDictionary>'
        $nodes = @($x.SelectNodes('//*[local-name()="Style" or local-name()="ControlTemplate"]'))
        $found = @(Get-HardColor -Nodes $nodes -Label 'sample')
        $found.Count | Should -Be 3
        ($found -join ' ') | Should -Match 'White'
        ($found -join ' ') | Should -Match '#1E1E1E'
        ($found -join ' ') | Should -Match '#555555'
    }
}

Describe 'No raw exception text in the UI (f)' {
    It 'flags raw exception text in a UI sink' {
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Add-LogLine ("x {0}" -f $_.Exception.Message) }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $row.Detail = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { $State.ErrorMsg = $_.Exception.Message }').Count | Should -Be 1
        @(Get-RawExceptionToUi -ScriptText 'try { 1 } catch { Write-Log ("x {0}" -f $_.Exception.Message) -Level ERROR -Quiet; Add-LogLine "Failed. See the log file." }').Count | Should -Be 0
    }

    It 'sends no raw exception text to a UI sink in <Name>' -ForEach $script:PsFilesD {
        @(Get-RawExceptionToUi -Path $Path) | Should -BeNullOrEmpty
    }
}
