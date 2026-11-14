# Contrast violations that exist today and are not fixed yet. Each entry is
# "<control>:<metric>" or "text:<visible text>". The window test fails when a
# violation is not listed (a new defect) and when a listed one no longer happens
# (remove it here). Run Show-ContrastViolations.ps1 to print the current list.
#
# Text must reach 4.5:1 and a control border 3:1, enabled and disabled, in both
# themes. The causes are shared by every suite tool: the SuiteCommon palette
# (the disabled-button text), the MahApps disabled
# style, the light text box border and the unselected tab header. The fix
# belongs in app-packager-suite, then a sync of the vendored copy, not here.

function Get-KnownContrastDefect {
    param([Parameter(Mandatory)][ValidateSet('dark', 'light')][string]$Theme)
    switch ($Theme) {
        'dark' {
            @(
                'btnViewCoverage:DisabledText', 'btnViewTemplates:DisabledText', 'btnViewWindows:DisabledText',
                'txtTemplateScheduleSummary:DisabledBorder', 'txtTemplateScheduleSummary:EnabledBorder'
            )
        }
        'light' {
            @(
                'btnApplyTemplateBulk:DisabledText', 'btnApplyTemplateOne:DisabledText', 'btnDeleteTemplate:DisabledText',
                'btnEditTemplate:DisabledText', 'btnEditWindow:DisabledText', 'btnExportCsv:DisabledText',
                'btnExportHtml:DisabledText', 'btnNewTemplate:DisabledText', 'btnNewWindow:DisabledText',
                'btnOptions:DisabledText', 'btnRefresh:DisabledText', 'btnRemoveWindow:DisabledText',
                'btnToggleEnabled:DisabledText', 'btnViewCoverage:DisabledText', 'btnViewTemplates:DisabledText',
                'btnViewWindows:DisabledText', 'cboStatusFilter:DisabledText', 'cboStatusFilter:EnabledText',
                'txtTemplateScheduleSummary:DisabledBorder', 'txtTemplateScheduleSummary:EnabledBorder'
            )
        }
    }
}