#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Guard tests for tooltips (Q-103). They parse the XAML of the main window and the XAML here-strings of the
# script. They create no WPF control and open no window, so they run on Windows PowerShell 5.1 and on
# PowerShell 7.
#
# Rules:
#   1. Each control in the audited list has a ToolTip.
#   2. A tooltip is not empty, is plain ASCII, ends in a period and is at most 200 characters.
#   3. A tooltip holds no double quote, ampersand or angle bracket (the here-strings stay valid).

BeforeDiscovery {
    # The audited controls: the x:Name of each control that must explain itself.
    $script:Audited = @(
        'toggleTheme', 'btnViewWindows', 'btnViewCoverage', 'btnViewTemplates', 'btnOptions', 'btnRefresh', 'btnExportCsv', 'btnExportHtml', 'cboStatusFilter', 'btnCancelWork', 'txtFilter',
        'btnNewWindow', 'btnEditWindow', 'btnToggleEnabled', 'btnRemoveWindow',
        'btnApplyTemplateBulk', 'btnNewTemplate', 'btnEditTemplate', 'btnDeleteTemplate', 'btnApplyTemplateOne',
        'cboScheduleRecurrence', 'dpScheduleStartDate', 'txtScheduleHour', 'txtScheduleMinute',
        'txtDurationHours', 'txtDurationMinutes', 'chkScheduleUtc', 'cboDayOfWeek', 'cboWeekOrder',
        'txtDayOfMonth', 'txtPatchOffset',
        'btnPickTargetColl', 'txtWindowName', 'cboWindowType', 'chkWindowEnabled',
        'txtTemplateName', 'txtTemplateDesc', 'cboTplWindowType',
        'cboApplyTemplate', 'txtApplyWindowName', 'lstApplyTargets', 'btnAddTarget', 'btnRemoveTarget',
        'btnCatConnection', 'btnCatAbout', 'txtSiteCode', 'txtSmsProvider'
    ) | ForEach-Object { @{ Name = $_ } }
}

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:XamlNs = 'http://schemas.microsoft.com/winfx/2006/xaml'
    $script:MaxLength = 200

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
        $docs.ToArray()
    }

    # The tooltip of every control that has the x:Name, in every document.
    function Get-ControlTooltip {
        param([string]$Name)
        $found = New-Object System.Collections.Generic.List[object]
        foreach ($d in $script:Docs) {
            if ($d.ParseError) { continue }
            foreach ($el in @($d.Xml.SelectNodes('//*[@*[local-name()="Name"]="' + $Name + '"]'))) {
                if ((Get-XName $el) -ne $Name) { continue }
                $found.Add([pscustomobject]@{ Label = $d.Label; Tip = $el.GetAttribute('ToolTip') })
            }
        }
        $found.ToArray()
    }

    $script:Docs = @(Get-XamlDocument)
}

Describe 'XAML documents' {
    It 'finds the main window and the dialogs of the script' {
        @($script:Docs).Count | Should -BeGreaterThan 1
    }

    It 'parses every XAML document' {
        foreach ($d in $script:Docs) { $d.ParseError | Should -BeNullOrEmpty -Because $d.Label }
    }
}

Describe 'Tooltips of the audited controls' {
    It '<Name> exists and has a tooltip that is not empty' -ForEach $script:Audited {
        $hits = @(Get-ControlTooltip -Name $Name)
        $hits.Count | Should -BeGreaterThan 0 -Because "$Name must exist in a XAML document"
        foreach ($h in $hits) { $h.Tip | Should -Not -BeNullOrEmpty -Because ('{0}: {1} has no ToolTip' -f $h.Label, $Name) }
    }

    It '<Name> has a tooltip that is ASCII, ends in a period and is short' -ForEach $script:Audited {
        foreach ($h in @(Get-ControlTooltip -Name $Name)) {
            $h.Tip | Should -Match '^[\x20-\x7E]+$' -Because ('{0}: the tooltip of {1} is plain ASCII' -f $h.Label, $Name)
            $h.Tip | Should -Match '\.$' -Because ('{0}: the tooltip of {1} ends in a period' -f $h.Label, $Name)
            $h.Tip.Length | Should -BeLessOrEqual $script:MaxLength -Because ('{0}: the tooltip of {1} is at most {2} characters' -f $h.Label, $Name, $script:MaxLength)
            $h.Tip | Should -Not -Match '["&<>]' -Because ('{0}: the tooltip of {1} holds no quote, ampersand or angle bracket' -f $h.Label, $Name)
        }
    }
}
