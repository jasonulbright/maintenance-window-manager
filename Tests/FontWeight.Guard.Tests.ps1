#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for the font weight. They parse the files. They create no WPF control
# and open no window, so they run on Windows PowerShell 5.1 and on PowerShell 7.
#
# Rule: bold text is for two things only.
#   1. A field title: the text label of an input or of a value field.
#   2. The app title in the sidebar (txtAppTitle).
# A page header, a section header, a status line, the About text and a progress
# title keep their normal weight (SemiBold or Normal). A theme change must never
# make them bold.
#
# (A button in a container is skipped when the test looks for the first control.)
# Field title, as the test defines it (an allow-list entry is the text, or Text@FontSize to name one size):
#   - a TextBlock or Label whose next control in the same panel is an input
#     (TextBox, PasswordBox, ComboBox, CheckBox, RadioButton, DatePicker, Slider,
#     ToggleSwitch, NumericUpDown, TimePicker), or the one in the cell to its
#     right or below in the same Grid; or
#   - a TextBlock or Label with the exact text of the short allow-list below
#     (it names a value field that the window shows as read-only text); or
#   - a TextBlock or Label that takes a style of the allow-list of title styles.
# A button keeps its own bold text (the sidebar buttons). A control other than a
# button, a TextBlock or a Label is never bold.


BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlNs = 'http://schemas.microsoft.com/winfx/2006/xaml'

    # ---- Per-app configuration ------------------------------------------------
    # The exact text of a label that names a value field shown as read-only text.
    $script:AllowedTitleText = @(
        'Description', 'Schedule@11', 'Target Collection', 'Site Code', 'SMS Provider'
    )
    # The key of a style that only field titles take.
    $script:AllowedTitleStyleKey = @()
    # A style key that makes the selected tab header bold (it was bold before the theme change).
    $script:AllowedStyleKey = @('TabHeaderText')
    # The functions that may set a bold weight at run time: the site code and SMS provider titles.
    $script:AllowedRuntimeFunction = @('Set-ConnectionTextTheme')
    # Names that must never be bold. A name that is not in the app is skipped; at least one must exist.
    $script:NeverBoldName = @(
        'txtStatus', 'txtModuleTitle', 'txtModuleHeader', 'txtAboutVersion', 'txtDeployTitle',
        'txtEditHeader', 'txtCveTitle', 'txtProgressTitle', 'txtProgressStep', 'txtTemplateParamsHeader',
        'lblLogOutput', 'txtSelectedColl', 'txtVersion', 'txtCollName'
    )
    $script:AppTitleName = 'txtAppTitle'

    $script:BoldWord = '^(Bold|ExtraBold|UltraBold|Black|Heavy|ExtraBlack|UltraBlack|[7-9]\d\d|1000)$'
    $script:InputName = @('TextBox', 'PasswordBox', 'ComboBox', 'CheckBox', 'RadioButton', 'DatePicker', 'Slider', 'ToggleSwitch', 'NumericUpDown', 'TimePicker')
    $script:ButtonName = @('Button', 'ToggleButton', 'RepeatButton')
    $script:LabelName = @('TextBlock', 'Label')


    function Get-XName {
        param($Node)
        $Node.GetAttribute('Name', $script:XamlNs)
    }

    function Get-StyleKey {
        param($Node)
        $Node.GetAttribute('Key', $script:XamlNs)
    }

    # Every XAML document of the app: each .xaml file and each XAML here-string of a script or module.
    function Get-XamlDocument {
        $skip = '\\(Tests|tools|docs|\.git|\.github|\.claude|TestResults)\\'
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

    function Test-InputNode {
        param($Node)
        $Node -is [System.Xml.XmlElement] -and ($script:InputName -contains $Node.LocalName)
    }

    # True when the first control found in the node (the node itself, then its descendants) is an input.
    function Test-FirstControlIsInput {
        param($Node)
        if ($null -eq $Node -or $Node -isnot [System.Xml.XmlElement]) { return $false }
        $layout = 'Border|Grid|StackPanel|DockPanel|WrapPanel|Canvas|ScrollViewer|Viewbox|UniformGrid'
        if ($Node.LocalName -match ('^(' + $layout + ')$')) {
            foreach ($child in @($Node.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.LocalName -notmatch '\.' -and $_.LocalName -ne 'Button' })) {
                return (Test-FirstControlIsInput -Node $child)
            }
            return $false
        }
        Test-InputNode -Node $Node
    }

    # The control that a label titles: the cell to the right or below in a Grid, else the next sibling.
    function Get-TitledControl {
        param($Label)
        $parent = $Label.ParentNode
        if ($parent -isnot [System.Xml.XmlElement]) { return $null }
        $sibs = @($parent.ChildNodes | Where-Object { $_ -is [System.Xml.XmlElement] -and $_.LocalName -notmatch '\.' })
        if ($parent.LocalName -eq 'Grid') {
            $row = 0; $col = 0
            [void][int]::TryParse($Label.GetAttribute('Grid.Row'), [ref]$row)
            [void][int]::TryParse($Label.GetAttribute('Grid.Column'), [ref]$col)
            $best = $null; $bestD = 1000
            foreach ($s in $sibs) {
                if ([object]::ReferenceEquals($s, $Label)) { continue }
                $r = 0; $k = 0
                [void][int]::TryParse($s.GetAttribute('Grid.Row'), [ref]$r)
                [void][int]::TryParse($s.GetAttribute('Grid.Column'), [ref]$k)
                $d = $null
                if ($r -eq $row -and $k -gt $col) { $d = $k - $col }
                elseif ($k -eq $col -and $r -gt $row) { $d = 100 + ($r - $row) }
                if ($null -ne $d -and $d -lt $bestD) { $best = $s; $bestD = $d }
            }
            return $best
        }
        for ($n = 0; $n -lt $sibs.Count; $n++) {
            if ([object]::ReferenceEquals($sibs[$n], $Label)) {
                if ($n + 1 -lt $sibs.Count) { return $sibs[$n + 1] }
                return $null
            }
        }
        $null
    }

    function Test-FieldTitle {
        param($Node)
        if ($script:LabelName -notcontains $Node.LocalName) { return $false }
        $caption = $Node.GetAttribute('Text'); if (-not $caption) { $caption = $Node.GetAttribute('Content') }
        if ($caption -and ($script:AllowedTitleText -contains $caption -or $script:AllowedTitleText -contains ('{0}@{1}' -f $caption, $Node.GetAttribute('FontSize')))) { return $true }
        $style = $Node.GetAttribute('Style')
        foreach ($k in $script:AllowedTitleStyleKey) { if ($style -match ('StaticResource\s+' + [regex]::Escape($k) + '\b')) { return $true } }
        Test-FirstControlIsInput -Node (Get-TitledControl -Label $Node)
    }

    function Get-Describe {
        param($Node)
        $name = Get-XName $Node
        $text = $Node.GetAttribute('Text'); if (-not $text) { $text = $Node.GetAttribute('Content') }
        '{0}{1}{2}' -f $Node.LocalName, $(if ($name) { " $name" } else { '' }), $(if ($text) { " ""$text""" } else { '' })
    }

    # Lists every element, style setter and inline Bold of a document that has a bold weight
    # and is not a field title, the sidebar app title or a button.
    function Get-BoldViolation {
        param($Document)
        $found = New-Object System.Collections.Generic.List[string]
        $xml = $Document.Xml
        foreach ($el in @($xml.SelectNodes('//*[@FontWeight]'))) {
            if ($el.GetAttribute('FontWeight') -notmatch $script:BoldWord) { continue }
            $name = Get-XName $el
            $kind = $el.LocalName
            if ($name -eq $script:AppTitleName) { continue }
            if ($script:ButtonName -contains $kind) { continue }
            if ($script:LabelName -contains $kind) {
                if ($name -and $script:NeverBoldName -contains $name) { $found.Add(('{0}: {1} is a header or status text and must not be bold' -f $Document.Label, (Get-Describe $el))); continue }
                if (Test-FieldTitle -Node $el) { continue }
                $found.Add(('{0}: {1} is bold and is not a field title' -f $Document.Label, (Get-Describe $el)))
                continue
            }
            if ($kind -eq 'Setter') { continue }
            $found.Add(('{0}: {1} is bold and is not a field title' -f $Document.Label, (Get-Describe $el)))
        }
        # <Bold> inline elements make bold text.
        foreach ($el in @($xml.SelectNodes('//*[local-name()="Bold"]'))) {
            $found.Add(('{0}: a Bold inline element holds bold text' -f $Document.Label))
        }
        # Style setters: a Setter of FontWeight with a bold value, in a style or a trigger.
        foreach ($setter in @($xml.SelectNodes('//*[local-name()="Setter"][@Property="FontWeight"]'))) {
            if ($setter.GetAttribute('Value') -notmatch $script:BoldWord) { continue }
            $style = $setter.SelectSingleNode('ancestor::*[local-name()="Style"][1]')
            if ($null -eq $style) { $found.Add(('{0}: a FontWeight setter outside a style is bold' -f $Document.Label)); continue }
            $key = Get-StyleKey $style
            $target = ($style.GetAttribute('TargetType') -replace '^.*:', '')
            if ($key -and $script:AllowedStyleKey -contains $key) { continue }
            if ($key -and $script:AllowedTitleStyleKey -contains $key) { continue }
            if ($script:ButtonName -contains $target) { continue }
            $found.Add(('{0}: style {1} (target {2}) sets a bold weight' -f $Document.Label, $(if ($key) { $key } else { '(implicit)' }), $target))
        }
        $found.ToArray()
    }

    # Names of the elements of the app that must never be bold: returns the bold ones.
    function Get-NeverBoldViolation {
        param($Document)
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($name in $script:NeverBoldName) {
            foreach ($el in @($Document.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $name + '"]'))) {
                if ($el.GetAttribute('FontWeight') -match $script:BoldWord) { $out.Add(('{0}: {1} is bold' -f $Document.Label, $name)) }
            }
        }
        $out.ToArray()
    }

    # Code that sets a bold weight at run time, outside the allowed functions.
    function Get-RuntimeBoldViolation {
        param([string]$Path, [string]$Label)
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
        Get-RuntimeBoldViolationFromAst -Ast $ast -Label $Label
    }

    function Get-RuntimeBoldViolationFromAst {
        param($Ast, [string]$Label)
        $out = New-Object System.Collections.Generic.List[string]
        $isBoldText = { param($t) $t -match '(?i)FontWeights\]?::(Bold|ExtraBold|UltraBold|Black|Heavy|ExtraBlack|UltraBlack)\b' -or $t -match '(?i)^[''"]?(Bold|ExtraBold|UltraBold|Black|Heavy)[''"]?$' }
        $hits = @($Ast.FindAll({ param($n)
                    ($n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -match '(?i)FontWeight') -or
                    ($n -is [System.Management.Automation.Language.MemberExpressionAst] -and $n.Extent.Text -match '(?i)FontWeights\]?::(Bold|ExtraBold|UltraBold|Black|Heavy|ExtraBlack|UltraBlack)\b') -or
                    ($n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Extent.Text -match '(?i)FontWeightProperty' -and $n.Extent.Text -match '(?i)Bold|Heavy|Black')
                }, $true))
        foreach ($h in $hits) {
            if ($h -is [System.Management.Automation.Language.AssignmentStatementAst]) {
                if (-not (& $isBoldText $h.Right.Extent.Text.Trim())) { continue }
            }
            $fn = $h.Parent
            while ($null -ne $fn -and $fn -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
            $fname = if ($fn) { $fn.Name } else { '(script)' }
            if ($script:AllowedRuntimeFunction -contains $fname) { continue }
            $out.Add(('{0} line {1} in {2}: sets a bold weight at run time' -f $Label, $h.Extent.StartLineNumber, $fname))
        }
        $out | Select-Object -Unique
    }

    function New-XmlSample {
        param([string]$Body)
        $x = New-Object System.Xml.XmlDocument
        $x.LoadXml('<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">' + $Body + '</Grid>')
        [pscustomobject]@{ Label = 'sample'; Xml = $x; ParseError = '' }
    }

    $script:Documents = Get-XamlDocument
    $script:ScriptFiles = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.ps1', '*.psm1' |
            Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch '\\(Tests|tools|docs|\.git|\.github|\.claude|TestResults)\\' })
}

Describe 'Bold text is for field titles and the sidebar app title only' {
    It 'finds the XAML of the app (the scan is not empty)' {
        @($script:Documents).Count | Should -BeGreaterThan 1
        @($script:Documents | Where-Object { $_.Label -eq 'MainWindow.xaml' }).Count | Should -Be 1
        $elements = 0
        foreach ($d in $script:Documents) { $elements += @($d.Xml.SelectNodes('//*[local-name()="TextBlock" or local-name()="Label"]')).Count }
        $elements | Should -BeGreaterThan 20
    }

    It 'parses every XAML document' {
        @($script:Documents | Where-Object { $_.ParseError } | ForEach-Object { '{0}: {1}' -f $_.Label, $_.ParseError }) | Should -BeNullOrEmpty
    }

    It 'finds the sidebar app title as the one bold title that is no field title' {
        $main = $script:Documents | Where-Object { $_.Label -eq 'MainWindow.xaml' }
        $title = $main.Xml.SelectSingleNode('//*[@*[local-name()="Name"]="' + $script:AppTitleName + '"]')
        $title | Should -Not -BeNullOrEmpty
        $title.GetAttribute('FontWeight') | Should -Match $script:BoldWord
    }

    It 'has no bold element that is not a field title, the app title or a button' {
        $bad = foreach ($d in $script:Documents) { Get-BoldViolation -Document $d }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'keeps the named header and status elements out of bold' {
        $bad = foreach ($d in $script:Documents) { Get-NeverBoldViolation -Document $d }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'has the named header and status elements in the app (the name list is not stale)' {
        $present = 0
        foreach ($d in $script:Documents) {
            foreach ($name in $script:NeverBoldName) { $present += @($d.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $name + '"]')).Count }
        }
        $present | Should -BeGreaterThan 0
    }

    It 'sets no bold weight at run time outside the field titles of the connection boxes' {
        $bad = foreach ($f in $script:ScriptFiles) { Get-RuntimeBoldViolation -Path $f.FullName -Label $f.Name }
        @($bad) | Should -BeNullOrEmpty
    }

    It 'scans script files for run-time font weights (the scan is not empty)' {
        @($script:ScriptFiles).Count | Should -BeGreaterThan 1
        $withSet = @($script:ScriptFiles | Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'FontWeight' })
        $withSet.Count | Should -BeGreaterThan 0
    }
}

Describe 'The font weight check itself' {
    It 'fails a bold page header (bad sample)' {
        $d = New-XmlSample '<TextBlock x:Name="txtModuleTitle" Text="Collections" FontSize="18" FontWeight="Bold"/>'
        @(Get-BoldViolation -Document $d).Count | Should -BeGreaterThan 0
        @(Get-NeverBoldViolation -Document $d).Count | Should -Be 1
    }

    It 'fails a bold label that titles no input (bad sample)' {
        $d = New-XmlSample '<StackPanel><TextBlock Text="Overview" FontWeight="Bold"/><ListBox/></StackPanel>'
        @(Get-BoldViolation -Document $d).Count | Should -Be 1
    }

    It 'fails a bold status text and a bold About text (bad sample)' {
        $d = New-XmlSample '<StackPanel><TextBlock x:Name="txtStatus" FontWeight="Bold"/><TextBlock x:Name="txtAboutVersion" FontWeight="Bold"/></StackPanel>'
        @(Get-NeverBoldViolation -Document $d).Count | Should -Be 2
        @(Get-BoldViolation -Document $d).Count | Should -Be 2
    }

    It 'fails a bold header that has a name of the never-bold list even when an input follows (bad sample)' {
        $d = New-XmlSample '<StackPanel><TextBlock x:Name="txtEditHeader" FontWeight="Bold"/><TextBox/></StackPanel>'
        @(Get-BoldViolation -Document $d).Count | Should -Be 1
    }

    It 'fails a bold weight in a text style and a Bold inline element (bad sample)' {
        $d = New-XmlSample '<Style x:Key="Plain" TargetType="TextBlock"><Setter Property="FontWeight" Value="Bold"/></Style><TextBlock><Bold>x</Bold></TextBlock>'
        @(Get-BoldViolation -Document $d).Count | Should -Be 2
    }

    It 'fails a bold weight set at run time outside the allowed function (bad sample)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('function Set-Header { param($t) $t.FontWeight = [System.Windows.FontWeights]::Bold }', [ref]$null, [ref]$null)
        @(Get-RuntimeBoldViolationFromAst -Ast $ast -Label 'sample').Count | Should -Be 1
    }

    It 'passes the allowed cases' {
        $d = New-XmlSample '<StackPanel><TextBlock x:Name="txtAppTitle" FontWeight="Bold"/><TextBlock Text="Name" FontWeight="Bold"/><TextBox/><TextBlock x:Name="txtModuleTitle" FontWeight="SemiBold"/><Button FontWeight="Bold"/></StackPanel>'
        @(Get-BoldViolation -Document $d).Count | Should -Be 0
        @(Get-NeverBoldViolation -Document $d).Count | Should -Be 0
        $ast = [System.Management.Automation.Language.Parser]::ParseInput('function Set-ConnectionTextTheme { param($t) $t.FontWeight = [System.Windows.FontWeights]::Bold } function Other { param($t) $t.FontWeight = [System.Windows.FontWeights]::SemiBold }', [ref]$null, [ref]$null)
        @(Get-RuntimeBoldViolationFromAst -Ast $ast -Label 'sample').Count | Should -Be 0
    }

    It 'takes a label above an input and a label left of an input in a Grid as field titles' {
        $d = New-XmlSample '<Grid><TextBlock Grid.Row="0" Grid.Column="0" Text="A" FontWeight="Bold"/><ComboBox Grid.Row="0" Grid.Column="1"/><TextBlock Grid.Row="1" Text="B" FontWeight="Bold"/><TextBox Grid.Row="2"/></Grid>'
        @(Get-BoldViolation -Document $d).Count | Should -Be 0
    }
}

Describe 'The allow-list of the font weight check' {
    It 'takes a label with the text of the allow-list as a field title, and no other label' {
        $saved = $script:AllowedTitleText
        try {
            $script:AllowedTitleText = @('Value Label')
            $d = New-XmlSample '<StackPanel><TextBlock Text="Value Label" FontWeight="Bold"/><TextBlock Text="Other Label" FontWeight="Bold"/></StackPanel>'
            $bad = @(Get-BoldViolation -Document $d)
            $bad.Count | Should -Be 1
            $bad[0] | Should -Match 'Other Label'
        }
        finally { $script:AllowedTitleText = $saved }
    }
}
