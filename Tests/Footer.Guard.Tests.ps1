#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the footer. They parse the XAML and the script code. They create no WPF
# control and open no window, so they run on Windows PowerShell 5.1 and on PowerShell 7.
#
# Rules:
#   1. The window has one footer status element, in a Border or a DockPanel.
#   2. No footer element sets a font property. SuiteCommon sets the one shared footer style
#      (Initialize-SuiteFooter). Only the colors differ between the themes.
#   3. The app calls Initialize-SuiteFooter once, and Update-SuiteFooter after the window loads,
#      after the Options dialog saves and after a connect.
#   4. The site code and the site server are edited only in the Options dialog (and in a first-run
#      window). Any other window shows them as read-only text, or does not show them.


BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlNs = 'http://schemas.microsoft.com/winfx/2006/xaml'

    # ---- Per-app configuration ------------------------------------------------
    # The script of the app that builds the window.
    $script:AppScript = 'start-maintenancewindowmgr.ps1'
    # The name of the footer status element. SuiteCommon adds the site connection text next to it.
    $script:StatusName = 'txtStatus'
    $script:FooterSiteName = 'txtFooterSite'
    # The member that carries the window-load script block.
    $script:LoadedMember = 'Add_Loaded'
    # The function that shows the Options dialog, and the functions that connect to the site.
    $script:OptionsFunction = 'Show-OptionsDialog'
    $script:ConnectFunction = @('Invoke-Refresh')
    # A window whose title is in this list, or whose label matches the pattern, may edit the site fields.
    $script:AllowedWindowTitle = @('Options')
    $script:AllowedWindowLabelPattern = '(?i)first.?run|wizard|welcome'
    # The names of the site code and the site server fields.
    $script:SiteFieldPattern = '(?i)^(txt|cbo|cmb|tb|edit)?(wiz|opt|setup)?(site(code|server|name)?|sms(provider)?|provider(name|fqdn)?)$'

    $script:FontAttribute = '^(FontFamily|FontSize|FontWeight|FontStyle|FontStretch|TextDecorations|TextElement\.Font\w+|Typography\.\w+|TextBlock\.Font\w+|Control\.Font\w+)$'
    $script:FontProperty = 'FontFamily|FontSize|FontWeight|FontStyle|FontStretch|TextDecorations'
    $script:InputName = @('TextBox', 'ComboBox', 'PasswordBox')


    function Get-XName {
        param($Node)
        $Node.GetAttribute('Name', $script:XamlNs)
    }

    # Every XAML document of the app: each .xaml file and each XAML here-string of a script or module.
    function Get-XamlDocument {
        $skip = '\\(Tests|tools|docs|Lib|\.git|\.github|\.claude|TestResults)\\'
        $files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.xaml', '*.ps1', '*.psm1' |
                Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch $skip })
        $docs = New-Object System.Collections.Generic.List[object]
        foreach ($f in $files) {
            $rel = $f.FullName.Substring($script:Root.Length).TrimStart('\')
            $text = [IO.File]::ReadAllText($f.FullName)
            $chunks = New-Object System.Collections.Generic.List[string]
            if ($f.Extension -eq '.xaml') { $chunks.Add($text) }
            else {
                foreach ($m in [regex]::Matches($text, "(?ms)@(?:'|`")[ \t]*\r?\n(\s*<.*?)\r?\n(?:'|`")@")) { $chunks.Add($m.Groups[1].Value) }
            }
            $i = 0
            foreach ($c in $chunks) {
                $i++
                $x = New-Object System.Xml.XmlDocument
                $err = ''
                try { $x.LoadXml($c) }
                catch {
                    $wrapped = '<Wrap xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" xmlns:Controls="http://metro.mahapps.com/winfx/xaml/controls">' + ($c -replace '^\s*<\?xml[^>]*\?>', '') + '</Wrap>'
                    try { $x.LoadXml($wrapped) } catch { $err = $_.Exception.Message }
                }
                $label = if ($f.Extension -eq '.xaml') { $rel } else { '{0} xaml {1}' -f $rel, $i }
                $docs.Add([pscustomobject]@{ Label = $label; Xml = $x; ParseError = $err })
            }
        }
        , $docs.ToArray()
    }

    function Get-WindowTitle {
        param($Document)
        $root = $Document.Xml.DocumentElement
        if ($null -eq $root) { return '' }
        if ($root.LocalName -eq 'Wrap') { $root = @($root.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] })[0] }
        if ($null -eq $root) { return '' }
        $root.GetAttribute('Title')
    }

    function Test-AllowedWindow {
        param($Document)
        if ($script:AllowedWindowTitle -contains (Get-WindowTitle $Document)) { return $true }
        $Document.Label -match $script:AllowedWindowLabelPattern
    }

    # The style keys of a document that set a font property.
    function Get-FontStyleKey {
        param($Document)
        $keys = New-Object System.Collections.Generic.List[string]
        foreach ($style in @($Document.Xml.SelectNodes('//*[local-name()="Style"]'))) {
            $key = $style.GetAttribute('Key', $script:XamlNs)
            if (-not $key) { continue }
            if (@($style.SelectNodes('.//*[local-name()="Setter"][@Property]') | Where-Object { $_.GetAttribute('Property') -replace '^.*\.', '' -match ('^(' + $script:FontProperty + ')$') }).Count -gt 0) { $keys.Add($key) }
        }
        $keys.ToArray()
    }

    function Get-FooterElement {
        param($Document, [string[]]$Name)
        $found = New-Object System.Collections.Generic.List[object]
        foreach ($n in $Name) {
            foreach ($el in @($Document.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $n + '"]'))) { $found.Add($el) }
        }
        $found.ToArray()
    }

    # Font settings on the footer elements: the element, the bar that holds it, every element of the
    # bar, the inline elements inside it, and a style that sets a font property.
    function Get-FooterFontViolation {
        param($Document, [string[]]$Name)
        $out = New-Object System.Collections.Generic.List[string]
        $styleKeys = @(Get-FontStyleKey -Document $Document)
        foreach ($status in @(Get-FooterElement -Document $Document -Name $Name)) {
            $bar = $status.ParentNode
            $scope = New-Object System.Collections.Generic.List[object]
            $scope.Add($status)
            if ($bar -is [System.Xml.XmlElement]) {
                $scope.Add($bar)
                foreach ($d in @($bar.SelectNodes('.//*'))) { $scope.Add($d) }
            }
            foreach ($d in @($status.SelectNodes('.//*'))) { $scope.Add($d) }
            foreach ($el in @($scope | Select-Object -Unique)) {
                foreach ($attr in @($el.Attributes)) {
                    if ($attr.Name -match $script:FontAttribute) {
                        $out.Add(('{0}: {1} sets {2}' -f $Document.Label, $el.LocalName, $attr.Name))
                    }
                }
                $style = $el.GetAttribute('Style')
                foreach ($k in $styleKeys) {
                    if ($style -match ('StaticResource\s+' + [regex]::Escape($k) + '\b')) { $out.Add(('{0}: {1} takes the style {2} that sets a font property' -f $Document.Label, $el.LocalName, $k)) }
                }
                if ($el.LocalName -match '^(Run|Bold|Italic|Underline|Hyperlink|Span)$') {
                    $out.Add(('{0}: the footer holds an inline {1} element' -f $Document.Label, $el.LocalName))
                }
            }
        }
        $out | Select-Object -Unique
    }

    # Site code and site server fields that a user can edit, outside the allowed windows.
    function Get-SiteFieldViolation {
        param($Document)
        if (Test-AllowedWindow -Document $Document) { return }
        foreach ($el in @($Document.Xml.SelectNodes('//*[@*[local-name()="Name"]]'))) {
            if ($script:InputName -notcontains $el.LocalName) { continue }
            $name = Get-XName $el
            if ($name -notmatch $script:SiteFieldPattern) { continue }
            if ($el.GetAttribute('IsReadOnly') -eq 'True') { continue }
            if ($el.GetAttribute('IsEnabled') -eq 'False') { continue }
            '{0}: {1} {2} is an editable site field outside the Options dialog' -f $Document.Label, $el.LocalName, $name
        }
    }

    function Get-SiteFieldInAllowedWindow {
        param($Document)
        if (-not (Test-AllowedWindow -Document $Document)) { return }
        foreach ($el in @($Document.Xml.SelectNodes('//*[@*[local-name()="Name"]]'))) {
            if ($script:InputName -notcontains $el.LocalName) { continue }
            if ((Get-XName $el) -match $script:SiteFieldPattern) { Get-XName $el }
        }
    }

    function Get-ScriptAst {
        param([string]$Path)
        $tokens = $null; $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
    }

    # Code that sets a font property of a footer element at run time.
    function Get-FooterRuntimeFontViolation {
        param($Ast, [string]$Label, [string[]]$Name)
        $out = New-Object System.Collections.Generic.List[string]
        $namePattern = '(' + (($Name | ForEach-Object { [regex]::Escape($_) }) -join '|') + ')'
        $bound = New-Object System.Collections.Generic.List[string]
        foreach ($a in @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true))) {
            if ($a.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and $a.Right.Extent.Text -match ('FindName\(\s*[''"]' + $namePattern + '[''"]\s*\)')) {
                $bound.Add($a.Left.VariablePath.UserPath)
            }
        }
        $isFooterTarget = {
            param([string]$t)
            foreach ($b in $bound) { if ($t -match ('^\$(script:|global:)?' + [regex]::Escape($b) + '$')) { return $true } }
            ($t -match ('FindName\(\s*[''"]' + $namePattern + '[''"]\s*\)')) -or ($t -match '(?i)status|footer')
        }
        foreach ($a in @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true))) {
            $member = $a.Left.Member.Extent.Text.Trim('''"')
            if ($member -notmatch ('^(' + $script:FontProperty + ')$')) { continue }
            if (& $isFooterTarget $a.Left.Expression.Extent.Text) {
                $out.Add(('{0} line {1}: sets {2} of a footer element' -f $Label, $a.Extent.StartLineNumber, $member))
            }
        }
        foreach ($c in @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] }, $true))) {
            if ($c.Extent.Text -notmatch ('(?:' + $script:FontProperty + ')Property')) { continue }
            if (& $isFooterTarget $c.Expression.Extent.Text) {
                $out.Add(('{0} line {1}: sets a font dependency property of a footer element' -f $Label, $c.Extent.StartLineNumber))
            }
        }
        $out.ToArray()
    }

    function Get-FunctionTable {
        param($Ast)
        $t = @{}
        foreach ($f in @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))) { $t[$f.Name] = $f }
        $t
    }

    function Get-CommandName {
        param($Ast)
        @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
    }

    # True when the node calls Update-SuiteFooter, directly or through a function of the same script.
    function Test-CallsFooterUpdate {
        param($Node, $FunctionTable, [int]$Depth = 3)
        if ($null -eq $Node) { return $false }
        $names = @(Get-CommandName -Ast $Node)
        if ($names -contains 'Update-SuiteFooter') { return $true }
        if ($Depth -le 0) { return $false }
        foreach ($n in ($names | Select-Object -Unique)) {
            if ($FunctionTable.ContainsKey($n) -and (Test-CallsFooterUpdate -Node $FunctionTable[$n] -FunctionTable $FunctionTable -Depth ($Depth - 1))) { return $true }
        }
        $false
    }

    # The script block that the app passes to the load event of the window.
    function Get-LoadedBlock {
        param($Ast, [string]$Member)
        $found = @($Ast.FindAll({ param($n)
                    $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -eq $Member
                }, $true))
        foreach ($f in $found) { foreach ($arg in @($f.Arguments)) { if ($arg -is [System.Management.Automation.Language.ScriptBlockExpressionAst]) { $arg } } }
    }

    function Get-FooterCallProblem {
        param($Ast)
        $problems = New-Object System.Collections.Generic.List[string]
        $fns = Get-FunctionTable -Ast $Ast
        $loaded = @(Get-LoadedBlock -Ast $Ast -Member $script:LoadedMember)
        if ($loaded.Count -eq 0) { $problems.Add("no $($script:LoadedMember) block found") }
        else {
            $init = @($loaded | Where-Object { @(Get-CommandName -Ast $_) -contains 'Initialize-SuiteFooter' })
            if ($init.Count -eq 0) { $problems.Add('Initialize-SuiteFooter is not called after the window loads') }
            if (-not (@($loaded | Where-Object { Test-CallsFooterUpdate -Node $_ -FunctionTable $fns }).Count)) { $problems.Add('Update-SuiteFooter is not called after the window loads') }
        }
        $initCalls = @(Get-CommandName -Ast $Ast | Where-Object { $_ -eq 'Initialize-SuiteFooter' })
        if ($initCalls.Count -ne 1) { $problems.Add(('Initialize-SuiteFooter is called {0} times, not once' -f $initCalls.Count)) }
        if (-not $fns.ContainsKey($script:OptionsFunction)) { $problems.Add("function $($script:OptionsFunction) not found") }
        elseif (-not (Test-CallsFooterUpdate -Node $fns[$script:OptionsFunction] -FunctionTable $fns)) { $problems.Add('Update-SuiteFooter is not called after the Options dialog saves') }
        foreach ($cf in $script:ConnectFunction) {
            if (-not $fns.ContainsKey($cf)) { $problems.Add("function $cf not found") }
            elseif (-not (Test-CallsFooterUpdate -Node $fns[$cf] -FunctionTable $fns)) { $problems.Add("Update-SuiteFooter is not called after $cf") }
        }
        $problems.ToArray()
    }

    function New-XmlSample {
        param([string]$Body, [string]$Title = '')
        $x = New-Object System.Xml.XmlDocument
        $t = if ($Title) { ' Title="' + $Title + '"' } else { '' }
        $x.LoadXml('<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"' + $t + '>' + $Body + '</Window>')
        [pscustomobject]@{ Label = 'sample'; Xml = $x; ParseError = '' }
    }

    function New-AstSample {
        param([string]$Code)
        [System.Management.Automation.Language.Parser]::ParseInput($Code, [ref]$null, [ref]$null)
    }

    $script:Documents = Get-XamlDocument
    $script:ScriptFiles = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.ps1', '*.psm1' |
            Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch '\\(Tests|tools|docs|Lib|\.git|\.github|\.claude|TestResults)\\' })
    $script:AppAst = Get-ScriptAst -Path (Join-Path $script:Root $script:AppScript)
}

Describe 'The footer of the window' {
    It 'finds the XAML and the script of the app (the scan is not empty)' {
        @($script:Documents).Count | Should -BeGreaterThan 1
        @($script:Documents | Where-Object { $_.Label -eq 'MainWindow.xaml' }).Count | Should -Be 1
        @($script:ScriptFiles).Count | Should -BeGreaterThan 0
        @($script:AppAst.FindAll({ $true }, $true)).Count | Should -BeGreaterThan 100
    }

    It 'parses every XAML document' {
        @($script:Documents | Where-Object { $_.ParseError } | ForEach-Object { '{0}: {1}' -f $_.Label, $_.ParseError }) | Should -BeNullOrEmpty
    }

    It 'has the footer status element in the main window, in a Border or a DockPanel' {
        $main = $script:Documents | Where-Object { $_.Label -eq 'MainWindow.xaml' }
        $status = @(Get-FooterElement -Document $main -Name @($script:StatusName))
        $status.Count | Should -Be 1
        $status[0].LocalName | Should -Be 'TextBlock'
        $status[0].ParentNode.LocalName | Should -BeIn @('Border', 'DockPanel')
    }

    It 'has the vendored SuiteCommon footer helper' {
        $psm1 = Get-Content -LiteralPath (Join-Path $script:Root 'Lib\SuiteCommon\SuiteCommon.psm1') -Raw
        foreach ($fn in 'Get-SuiteFooterStyle', 'Format-SuiteFooterSite', 'Initialize-SuiteFooter', 'Update-SuiteFooter') {
            $psm1 | Should -Match ('function\s+' + $fn + '\b')
        }
    }
}

Describe 'The footer font' {
    It 'sets no font property on a footer element in XAML' {
        $bad = foreach ($d in $script:Documents) { Get-FooterFontViolation -Document $d -Name @($script:StatusName, $script:FooterSiteName) }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'sets no font property on a footer element in code' {
        $bad = foreach ($f in $script:ScriptFiles) { Get-FooterRuntimeFontViolation -Ast (Get-ScriptAst -Path $f.FullName) -Label $f.Name -Name @($script:StatusName, $script:FooterSiteName) }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'scans the code that writes the status text (the scan is not empty)' {
        $fns = Get-FunctionTable -Ast $script:AppAst
        @($fns.Keys | Where-Object { $_ -match 'Status' }).Count | Should -BeGreaterThan 0
        (Get-Content -LiteralPath (Join-Path $script:Root $script:AppScript) -Raw) | Should -Match ([regex]::Escape($script:StatusName))
    }
}

Describe 'The footer calls' {
    It 'calls Initialize-SuiteFooter once and Update-SuiteFooter after load, after the Options save and after a connect' {
        @(Get-FooterCallProblem -Ast $script:AppAst) | Should -BeNullOrEmpty
    }
}

Describe 'The site fields' {
    It 'has no editable site code or site server field outside the Options dialog and the first-run windows' {
        $bad = foreach ($d in $script:Documents) { Get-SiteFieldViolation -Document $d }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'finds the site fields in the Options dialog (the scan is not empty)' {
        $names = foreach ($d in $script:Documents) { Get-SiteFieldInAllowedWindow -Document $d }
        @($names).Count | Should -BeGreaterThan 0
    }
}

Describe 'The footer check itself' {
    # The samples below use the names of the sample code, not the names of the app.
    BeforeEach {
        $script:SavedConfig = @{ Loaded = $script:LoadedMember; Options = $script:OptionsFunction; Connect = $script:ConnectFunction; Titles = $script:AllowedWindowTitle }
        $script:LoadedMember = 'Add_Loaded'
        $script:OptionsFunction = 'Show-OptionsDialog'
        $script:ConnectFunction = @('Invoke-ConnectSite')
        $script:AllowedWindowTitle = @('Options')
    }
    AfterEach {
        $script:LoadedMember = $script:SavedConfig.Loaded
        $script:OptionsFunction = $script:SavedConfig.Options
        $script:ConnectFunction = $script:SavedConfig.Connect
        $script:AllowedWindowTitle = $script:SavedConfig.Titles
    }

    It 'fails a font size on the status element (bad sample)' {
        $d = New-XmlSample '<Border><TextBlock x:Name="txtStatus" FontSize="12"/></Border>'
        @(Get-FooterFontViolation -Document $d -Name @('txtStatus')).Count | Should -Be 1
    }

    It 'fails a font weight, a style and a text decoration on a footer element (bad sample)' {
        $d = New-XmlSample '<DockPanel><TextBlock x:Name="txtFooterSite" FontWeight="Bold" TextDecorations="Underline"/><TextBlock x:Name="txtStatus"/></DockPanel>'
        @(Get-FooterFontViolation -Document $d -Name @('txtStatus', 'txtFooterSite')).Count | Should -Be 2
    }

    It 'fails a font setting on the bar, on a sibling and on an inline element (bad sample)' {
        $d = New-XmlSample '<Border TextElement.FontSize="9"><StackPanel><TextBlock x:Name="txtStatus"><Run FontFamily="Consolas" Text="x"/></TextBlock></StackPanel></Border>'
        @(Get-FooterFontViolation -Document $d -Name @('txtStatus')).Count | Should -BeGreaterThan 1
    }

    It 'fails a status element that takes a style with a font setter (bad sample)' {
        $d = New-XmlSample '<Border><Border.Resources><Style x:Key="Small" TargetType="TextBlock"><Setter Property="FontSize" Value="9"/></Style></Border.Resources><TextBlock x:Name="txtStatus" Style="{StaticResource Small}"/></Border>'
        @(Get-FooterFontViolation -Document $d -Name @('txtStatus')).Count | Should -Be 1
    }

    It 'passes a footer without font settings and ignores other elements' {
        $d = New-XmlSample '<Grid><TextBlock Text="Title" FontSize="18" FontWeight="Bold"/><Border><TextBlock x:Name="txtStatus" Text="Ready."/></Border></Grid>'
        @(Get-FooterFontViolation -Document $d -Name @('txtStatus')).Count | Should -Be 0
    }

    It 'fails a font setting on the status element at run time (bad sample)' {
        $ast = New-AstSample '$txtStatus = $window.FindName(''txtStatus'')
$txtStatus.FontSize = 12'
        @(Get-FooterRuntimeFontViolation -Ast $ast -Label 'sample' -Name @('txtStatus')).Count | Should -Be 1
        $ast = New-AstSample '$window.FindName(''txtStatus'').FontWeight = [System.Windows.FontWeights]::Bold'
        @(Get-FooterRuntimeFontViolation -Ast $ast -Label 'sample' -Name @('txtStatus')).Count | Should -Be 1
        $ast = New-AstSample 'function Set-StatusText { param($t) $script:status.TextDecorations = $null }'
        @(Get-FooterRuntimeFontViolation -Ast $ast -Label 'sample' -Name @('txtStatus')).Count | Should -Be 1
    }

    It 'passes code that sets text and fonts of other elements' {
        $ast = New-AstSample '$txtStatus = $window.FindName(''txtStatus'')
$txtStatus.Text = ''x''
$node.FontWeight = [System.Windows.FontWeights]::SemiBold'
        @(Get-FooterRuntimeFontViolation -Ast $ast -Label 'sample' -Name @('txtStatus')).Count | Should -Be 0
    }

    It 'fails an editable site field outside the Options dialog (bad sample)' {
        $d = New-XmlSample '<StackPanel><TextBox x:Name="txtSiteServer"/><ComboBox x:Name="cboSite"/><TextBox x:Name="txtProvider"/><TextBox x:Name="txtWizSiteCode"/></StackPanel>' 'Main'
        @(Get-SiteFieldViolation -Document $d).Count | Should -Be 4
    }

    It 'passes a read-only site field, other fields and the Options dialog' {
        $d = New-XmlSample '<StackPanel><TextBox x:Name="txtSiteServer" IsReadOnly="True"/><TextBox x:Name="txtSiteNotes"/><TextBlock x:Name="txtSiteCode"/></StackPanel>' 'Main'
        @(Get-SiteFieldViolation -Document $d).Count | Should -Be 0
        $o = New-XmlSample '<StackPanel><TextBox x:Name="txtSiteCode"/><TextBox x:Name="txtSmsProvider"/></StackPanel>' 'Options'
        @(Get-SiteFieldViolation -Document $o).Count | Should -Be 0
        @(Get-SiteFieldInAllowedWindow -Document $o).Count | Should -Be 2
    }

    It 'fails an app that does not call the footer helpers (bad sample)' {
        $ast = New-AstSample 'function Show-OptionsDialog { }
function Invoke-ConnectSite { }
$window.Add_Loaded({ Update-StatusBarSummary })'
        @(Get-FooterCallProblem -Ast $ast).Count | Should -BeGreaterThan 3
    }

    It 'passes an app that calls the footer helpers directly and through a function' {
        $ast = New-AstSample 'function Sync-FooterSite { Update-SuiteFooter -Window $window -SiteCode $a -SiteServer $b }
function Show-OptionsDialog { Sync-FooterSite }
function Invoke-ConnectSite { $t.Add_Tick({ Update-SuiteFooter -Window $window }) }
$window.Add_Loaded({ [void](Initialize-SuiteFooter -Window $window); Sync-FooterSite })'
        @(Get-FooterCallProblem -Ast $ast).Count | Should -Be 0
    }

    It 'fails an app that calls Initialize-SuiteFooter twice (bad sample)' {
        $ast = New-AstSample 'function Sync-FooterSite { Update-SuiteFooter -Window $window }
function Show-OptionsDialog { Sync-FooterSite }
function Invoke-ConnectSite { Sync-FooterSite; Initialize-SuiteFooter -Window $window }
$window.Add_Loaded({ Initialize-SuiteFooter -Window $window; Sync-FooterSite })'
        @(Get-FooterCallProblem -Ast $ast).Count | Should -Be 1
    }
}
