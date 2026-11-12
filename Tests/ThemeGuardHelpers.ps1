# Helpers of Theme.Guard.Tests.ps1. Dot-source from BeforeDiscovery and from BeforeAll.
# They parse files. They create no WPF control and open no window.

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:ThemeProperty = @('FontSize', 'FontWeight', 'FontStyle', 'TextDecorations', 'FontFamily', 'FontStretch')
$script:BannedPattern = '(?i)\b(easily|simply|seamless(ly)?|powerful|robust|streamline[sd]?|streamlining|quickly|just|effortless(ly)?|intuitive(ly)?|handy|helps you|allows you to)\b'
$script:XamlNamespace = 'http://schemas.microsoft.com/winfx/2006/xaml/presentation'

# The path of a file relative to the repository root. A worktree lives under .claude, so a filter
# on the full path would match every file. A filter on the relative path does not.
function Get-RelativeRepoPath {
    param([string]$FullName)
    return $FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
}

# True when a brush value names an accent or a blue color.
function Test-BlueBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    if ($Value -match 'MahApps\.(Brushes|Colors)\.(Accent\d*|AccentBase|Highlight|Button\.Flat\.Background\.Pressed)\b') { return $true }
    if ($Value -match '(?i)^\s*(blue|dodgerblue|royalblue|steelblue|navy|mediumblue|darkblue|cornflowerblue)\s*$') { return $true }
    if ($Value -match '^\s*#(?:[0-9A-Fa-f]{2})?([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})([0-9A-Fa-f]{2})\s*$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0
        $g = [Convert]::ToInt32($Matches[2], 16) / 255.0
        $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $max = [Math]::Max($r, [Math]::Max($g, $b)); $min = [Math]::Min($r, [Math]::Min($g, $b))
        $d = $max - $min
        if ($d -lt 0.2 -or $max -eq 0) { return $false }
        if ($max -ne $b) { return $false }
        $hue = 60.0 * ((4.0 + (($r - $g) / $d)))
        return ($hue -ge 190 -and $hue -le 260)
    }
    return $false
}

function Test-WhiteBrush {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    return ($Value -match '(?i)^\s*(white|#FFFFFF|#FFFFFFFF)\s*$' -or $Value -match 'MahApps\.Brushes\.(Ideal|ThemeIdeal)Foreground')
}

# The value of a Setter: the Value attribute, or the Color of a nested brush.
function Get-SetterValue {
    param($Setter)
    if ($Setter.HasAttribute('Value')) { return $Setter.GetAttribute('Value') }
    $node = $Setter.SelectSingleNode('.//*[@Color]')
    if ($node) { return $node.GetAttribute('Color') }
    return ''
}

# The script files of the app: the start script and the module. The tests, the tools and the vendored library are outside.
function Get-AppScriptFile {
    $files = @(Get-ChildItem -LiteralPath $script:RepoRoot -File -Recurse | Where-Object {
            $rel = Get-RelativeRepoPath $_.FullName
            $_.Extension -in '.ps1', '.psm1' -and $_.Name -notlike '*.Tests.ps1' -and $rel -notmatch '^(\.claude|\.git|Tests|tools|Lib|Logs)([\\/]|$)'
        })
    return $files
}

# The styles that the dialogs take from MainWindow.xaml. The start script holds the same list (Get-SharedStyleXaml).
$script:SharedStyleMarker = '<!--SHARED_STYLES-->'
$script:SharedStyleKey = @('ThemedButtonTemplate', 'ThemedButtonStyle', 'DialogAccentButton', 'DialogButton')
function Get-SharedStyleXaml {
    [xml]$main = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'MainWindow.xaml') -Raw
    $nodes = @($main.SelectNodes("//*[local-name()='ResourceDictionary']/*[local-name()='ControlTemplate' or local-name()='Style']") | Where-Object {
            $key = $_.GetAttribute('Key', 'http://schemas.microsoft.com/winfx/2006/xaml')
            ($key -in $script:SharedStyleKey) -or (-not $key -and $_.GetAttribute('TargetType') -in 'ComboBoxItem', 'ListBoxItem')
        })
    return (($nodes | ForEach-Object { $_.OuterXml }) -join [Environment]::NewLine)
}

# Every XAML document of the app: the XAML files and the XAML here-strings of the scripts.
# A here-string counts as XAML when its first tag declares the WPF namespace.
function Get-XamlDocument {
    $docs = @()
    foreach ($f in @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.xaml' -File -Recurse | Where-Object { (Get-RelativeRepoPath $_.FullName) -notmatch '^(\.claude|\.git)([\\/]|$)' })) {
        $docs += [PSCustomObject]@{ Source = $f.Name; Text = (Get-Content -LiteralPath $f.FullName -Raw) }
    }
    foreach ($f in @(Get-AppScriptFile)) {
        $raw = Get-Content -LiteralPath $f.FullName -Raw
        $i = 0
        foreach ($m in [regex]::Matches($raw, "(?s)@['""]\r?\n(\s*<[^\r\n]*\r?\n(?:[^\r\n]*\r?\n)*?)['""]@")) {
            $text = $m.Groups[1].Value
            if ($text -notmatch ('^\s*<[^>]*xmlns="' + [regex]::Escape($script:XamlNamespace) + '"')) { continue }
            $i++
            # The dialogs take the shared styles from MainWindow.xaml, as the start script does at run time.
            $text = $text.Replace($script:SharedStyleMarker, (Get-SharedStyleXaml))
            $docs += [PSCustomObject]@{ Source = ('{0} here-string {1}' -f $f.Name, $i); Text = $text }
        }
    }
    return $docs
}

# Returns one text per blue-background rule that has no white foreground.
function Get-BlueWithoutWhite {
    param([string]$Text, [string]$Source)
    [xml]$x = $Text
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($el in @($x.SelectNodes('//*'))) {
        if ($el.LocalName -eq 'Setter' -and $el.GetAttribute('Property') -eq 'Background') {
            if (Test-BlueBrush (Get-SetterValue $el)) {
                $fg = $el.ParentNode.ChildNodes | Where-Object { $_.LocalName -eq 'Setter' -and $_.GetAttribute('Property') -eq 'Foreground' } | Select-Object -First 1
                if (-not $fg -or -not (Test-WhiteBrush (Get-SetterValue $fg))) {
                    [void]$found.Add(('{0}: Setter Background {1} in <{2}> has no white Foreground' -f $Source, (Get-SetterValue $el), $el.ParentNode.LocalName))
                }
            }
        }
        elseif ($el.HasAttribute('Background') -and $el.LocalName -ne 'Setter') {
            if (Test-BlueBrush $el.GetAttribute('Background')) {
                if (-not (Test-WhiteBrush $el.GetAttribute('Foreground'))) {
                    [void]$found.Add(('{0}: <{1}> Background {2} has no white Foreground' -f $Source, $el.LocalName, $el.GetAttribute('Background')))
                }
            }
        }
    }
    return $found.ToArray()
}

# Every string of a script: the constant strings and the text parts of expandable strings.
function Get-ScriptString {
    param([string]$Path)
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $nodes = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.StringConstantExpressionAst] -or $n -is [System.Management.Automation.Language.ExpandableStringExpressionAst] }, $true)
    foreach ($n in $nodes) {
        # An attribute such as SuppressMessage holds a justification for the analyzer, not a UI string.
        $p = $n.Parent; $inAttribute = $false
        while ($p) { if ($p -is [System.Management.Automation.Language.AttributeAst]) { $inAttribute = $true; break }; $p = $p.Parent }
        if (-not $inAttribute) { [string]$n.Value }
    }
}

# The text a user can read in a XAML document: tooltips and text attributes.
function Get-XamlUiText {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($el in @($x.SelectNodes('//*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.LocalName -match '^(ToolTip|Text|Content|Header|Title|Watermark|OnContent|OffContent|ToolTipService.*)$' -or $a.Name -match 'Watermark$') {
                if ($a.Value -notmatch '^\{') { $a.Value }
            }
        }
    }
}

# One text per catch block that sends raw exception text to a UI sink.
# UI sinks: Add-LogLine, Set-StatusText, the themed dialogs, and a UI property such as Text or Detail.
# The log file (Write-Log) and a state field that the UI does not show are allowed.
function Get-RawExceptionToUi {
    param([string]$Path, [string]$ScriptText)
    $ast = if ($ScriptText) { [System.Management.Automation.Language.Parser]::ParseInput($ScriptText, [ref]$null, [ref]$null) } else { [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null) }
    $rawPattern = '\$_(?![A-Za-z0-9_])|\$PSItem|\$Error\b|\.Exception\b|\$ex\b|\$err\b'
    $uiProperty = @('Text', 'Content', 'Detail', 'ToolTip', 'Title', 'Header', 'ErrorMsg')
    $sinkCommand = @('Add-LogLine', 'Set-StatusText', 'Show-ThemedMessage', 'Show-ConfirmDialog', 'Show-Dialog')
    $found = @()
    foreach ($c in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CatchClauseAst] }, $true))) {
        foreach ($cmd in @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true))) {
            $name = [string]$cmd.GetCommandName()
            if ($sinkCommand -contains $name -and $cmd.Extent.Text -match $rawPattern) { $found += ('line {0}: {1}' -f $cmd.Extent.StartLineNumber, $cmd.Extent.Text) }
        }
        foreach ($as in @($c.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left -is [System.Management.Automation.Language.MemberExpressionAst] }, $true))) {
            $member = [string]$as.Left.Member.Value
            if ($uiProperty -contains $member -and $as.Right.Extent.Text -match $rawPattern) { $found += ('line {0}: {1}' -f $as.Extent.StartLineNumber, $as.Extent.Text) }
        }
    }
    return $found
}

# The title of a value box: the TextBlock that comes before it in the same panel (the rule of SuiteCommon).
function Get-FieldTitleElement {
    param($Box)
    $prev = $Box.PreviousSibling
    while ($prev -and $prev.NodeType -ne 'Element') { $prev = $prev.PreviousSibling }
    if ($prev -and $prev.LocalName -in 'TextBlock', 'Label') { return $prev }
    return $null
}

# Lists every hard-coded color in the Setters, the attributes and the nested brushes of the given nodes.
# Only a DynamicResource or a TemplateBinding is allowed. A property is a brush property when its name ends in
# Brush, Background, Foreground, Fill or Stroke, also as Border.Background or as an attached name.
# Any Color attribute that is not a DynamicResource or a TemplateBinding is a literal, wherever it sits.
function Get-HardColor {
    param($Nodes, [string]$Label)
    $brushProp = '(^|[.:])\w*(Brush|Background|Foreground|Fill|Stroke)$'
    $allowed = '^\{(DynamicResource|TemplateBinding)\b'
    foreach ($n in $Nodes) {
        foreach ($setter in @($n.SelectNodes('descendant-or-self::*[local-name()="Setter"]'))) {
            if ($setter.GetAttribute('Property') -notmatch $brushProp) { continue }
            $v = $setter.GetAttribute('Value')
            if ($v -and $v -notmatch $allowed) { '{0}: Setter {1}="{2}"' -f $Label, $setter.GetAttribute('Property'), $v }
        }
        foreach ($el in @($n.SelectNodes('descendant-or-self::*[not(local-name()="Setter")]'))) {
            foreach ($attr in @($el.Attributes)) {
                if ($attr.Name -match '^xmlns' -or $attr.NamespaceURI -eq 'http://www.w3.org/2000/xmlns/') { continue }
                if ($attr.LocalName -match $brushProp -and $attr.Value -notmatch $allowed) { '{0}: {1} {2}="{3}"' -f $Label, $el.LocalName, $attr.LocalName, $attr.Value }
                elseif ($attr.LocalName -eq 'Color' -and $attr.Value -notmatch $allowed) { '{0}: {1} holds the color {2}' -f $Label, $el.LocalName, $attr.Value }
            }
        }
    }
}

# One text per trigger that sets a font property and whose condition names a theme (Dark, Light, Theme).
# A trigger on a control state, such as the selected tab, sets the same font in both themes and is allowed.
function Get-ThemeFontTrigger {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($trigger in @($x.SelectNodes("//*[local-name()='Trigger' or local-name()='DataTrigger' or local-name()='MultiTrigger' or local-name()='MultiDataTrigger']"))) {
        $sets = @($trigger.SelectNodes("*[local-name()='Setter']") | Where-Object { $script:ThemeProperty -contains $_.GetAttribute('Property') })
        if ($sets.Count -eq 0) { continue }
        $condition = (@($trigger.Attributes | ForEach-Object { $_.Value }) + @($trigger.SelectNodes(".//*[local-name()='Condition']") | ForEach-Object { @($_.Attributes | ForEach-Object { $_.Value }) })) -join ' '
        if ($condition -match '(?i)IsDark|DarkMode|IsLight|LightMode|ThemeMode|\bTheme\b|\bDark\b') { $trigger.OuterXml.Substring(0, [Math]::Min(100, $trigger.OuterXml.Length)) }
    }
}

# True when the style with this key is based, directly or through other styles of the document, on a MahApps button style.
function Test-StyleBasedOnMahAppsButton {
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

# One text per Button that has no Style, or a Style that is not based on a MahApps button style.
function Get-ButtonWithoutSharedStyle {
    param([string]$Text)
    [xml]$x = $Text
    foreach ($b in @($x.SelectNodes("//*[local-name()='Button']"))) {
        $style = $b.GetAttribute('Style')
        $key = if ($style -match '(?:Static|Dynamic)Resource\s+([^\}\s]+)') { $Matches[1] } else { '' }
        if (-not $key -or -not (Test-StyleBasedOnMahAppsButton -Doc $x -Key $key)) { 'Button {0} has no shared style' -f $b.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml') }
    }
}
