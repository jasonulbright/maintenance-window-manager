#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x window test for the compact 1080p density and for vacant space. The main window is opened off screen, in
    both themes, at the work area of the session (1920x1032 on a 1080p screen at 100 percent) and at 1024x720, and
    measured: sidebar buttons, inline buttons, single-line boxes, grid rows and headers, the width that the content
    uses, and the free area below and right of the content of each view (at most 120 px high and 160 px wide).
    The visible run prints the work area of the session (SystemParameters.WorkArea).
#>

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed -and -not $env:SUITE_WINDOW_GUARD_REPORTED) { Write-Host ('Window tests skipped: {0}.' -f $script:WindowGuard.Reason) }
    $script:Cases = @(foreach ($dark in $true, $false) {
            $design = Get-DesignSize
            foreach ($size in @(@(1024, 720), @($design.W, $design.H))) {
                @{ Dark = $dark; Theme = $(if ($dark) { 'dark' } else { 'light' }); W = $size[0]; H = $size[1]; Size = ('{0}x{1}' -f $size[0], $size[1]); Design = ($size[0] -ne 1024) }
            }
        })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $script:WindowGuard = Test-WindowTestsAllowed
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Design = Get-DesignSize
    Add-Type -AssemblyName PresentationFramework
    $wa = [System.Windows.SystemParameters]::WorkArea
    Write-Host ('Work area of this session: {0}x{1}' -f $wa.Width, $wa.Height)
    $script:Measured = @{}
    foreach ($case in @(
            @{ Dark = $true; W = 1024; H = 720 }, @{ Dark = $true; W = $script:Design.W; H = $script:Design.H },
            @{ Dark = $false; W = 1024; H = 720 }, @{ Dark = $false; W = $script:Design.W; H = $script:Design.H })) {
        $key = '{0}|{1}x{2}' -f $(if ($case.Dark) { 'dark' } else { 'light' }), $case.W, $case.H
        $case.Config = $script:AppConfig
        $script:Measured[$key] = Invoke-WpfSta -Arguments $case -Script {
            $dsCfg = $Arguments.Config
            $dsW = $null
            try {
                $dsW = Open-TestWindow -XamlPath (Join-Path $Root $dsCfg.Xaml) -Dark $Arguments.Dark -Width $Arguments.W -Height $Arguments.H
                Initialize-ShellTheme -Window $dsW -Dark $Arguments.Dark -Views $dsCfg.Views -OptionsButtonName $dsCfg.OptionsButton -LogLabelName $dsCfg.LogLabel -ActiveView $dsCfg.Views[0].Name
                Wait-Layout $dsW
                $dsSidebar = @(foreach ($dsV in $dsCfg.Views) { $dsB = $dsW.FindName($dsV.Button); $dsP = $dsB.TransformToAncestor($dsW).Transform([System.Windows.Point]::new(0, 0)); [pscustomobject]@{ Name = $dsV.Button; Top = $dsP.Y; Bottom = $dsP.Y + $dsB.ActualHeight; Height = $dsB.ActualHeight } })
                $dsViews = @(foreach ($dsView in $dsCfg.Views) {
                        foreach ($dsV in $dsCfg.Views) { $dsW.FindName($dsV.Container).Visibility = if ($dsV.Name -eq $dsView.Name) { 'Visible' } else { 'Collapsed' } }
                        Wait-Layout $dsW
                        $dsC = $dsW.FindName($dsView.Container)
                        $dsInline = @(foreach ($dsEl in (Get-VisualDescendants $dsC)) { if ($dsEl -is [System.Windows.Controls.Button] -and $dsEl.Style -and (Test-VisibleInTree $dsEl) -and $dsEl.Name) { [pscustomobject]@{ Name = $dsEl.Name; Height = $dsEl.ActualHeight } } })
                        $dsBoxes = @(foreach ($dsEl in (Get-VisualDescendants $dsW)) { if ($dsEl -is [System.Windows.Controls.TextBox] -and $dsEl.Name -eq 'txtFilter' -and (Test-VisibleInTree $dsEl) -and $dsView.Name -eq 'Windows') { [pscustomobject]@{ Name = $dsEl.Name; Height = $dsEl.ActualHeight } } })
                        $dsGrids = @(foreach ($dsEl in (Get-VisualDescendants $dsC)) {
                                if ($dsEl -is [System.Windows.Controls.DataGrid] -and (Test-VisibleInTree $dsEl)) {
                                    $dsHdr = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.Primitives.DataGridColumnHeadersPresenter]) { $dsHdr = $dsX; break } }
                                    $dsRow = $null; foreach ($dsX in (Get-VisualDescendants $dsEl)) { if ($dsX -is [System.Windows.Controls.DataGridRow]) { $dsRow = $dsX; break } }
                                    [pscustomobject]@{ Name = $dsEl.Name; Width = $dsEl.ActualWidth; HeaderHeight = $(if ($dsHdr) { $dsHdr.ActualHeight } else { $null }); RowHeight = $(if ($dsRow) { $dsRow.ActualHeight } else { $null }) }
                                }
                            })
                        # The free area of the view: the content elements are the visible grids, boxes, lists, tabs and buttons.
                        $dsRight = 0.0; $dsBottom = 0.0
                        foreach ($dsEl in (Get-VisualDescendants $dsC)) {
                            if (($dsEl -is [System.Windows.Controls.DataGrid] -or $dsEl -is [System.Windows.Controls.TextBox] -or $dsEl -is [System.Windows.Controls.TabControl] -or $dsEl -is [System.Windows.Controls.ListBox] -or $dsEl -is [System.Windows.Controls.Button]) -and $dsEl.ActualWidth -gt 0 -and (Test-VisibleInTree $dsEl)) {
                                $dsPt = $dsEl.TransformToAncestor($dsC).Transform([System.Windows.Point]::new($dsEl.ActualWidth, $dsEl.ActualHeight))
                                if ($dsPt.X -gt $dsRight) { $dsRight = $dsPt.X }; if ($dsPt.Y -gt $dsBottom) { $dsBottom = $dsPt.Y }
                            }
                        }
                        [pscustomobject]@{ View = $dsView.Name; Width = $dsC.ActualWidth; Height = $dsC.ActualHeight; FreeRight = ($dsC.ActualWidth - $dsRight); FreeBottom = ($dsC.ActualHeight - $dsBottom); Buttons = $dsInline; Boxes = $dsBoxes; Grids = $dsGrids }
                    })
                [pscustomobject]@{ Width = $dsW.ActualWidth; Height = $dsW.ActualHeight; Sidebar = $dsSidebar; Views = $dsViews }
            }
            finally { if ($dsW) { $dsW.Close() } }
        }
    }
}

AfterAll {
    if ($script:WindowGuard.Allowed) { Close-WpfSession }
}

Describe 'Density <Theme> <Size>' -ForEach $script:Cases -Skip:(-not $script:WindowGuard.Allowed) {
    It 'sizes every sidebar button between 34 and 36 px and fits all of them without scrolling' {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size]
        foreach ($b in $m.Sidebar) { $b.Height | Should -BeIn 34, 35, 36 -Because $b.Name }
        ($m.Sidebar | Select-Object -Last 1).Bottom | Should -BeLessThan $m.Height
    }

    It 'sizes the inline buttons and boxes in the <View> view' -ForEach @(@{ View = 'Windows' }, @{ View = 'Coverage' }, @{ View = 'Templates' }) {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($b in $m.Buttons | Where-Object { $_.Name -notmatch '^btnView|^btnOptions' }) { $b.Height | Should -BeLessOrEqual 28 -Because $b.Name }
        foreach ($b in $m.Boxes) { $b.Height | Should -BeIn 26, 27, 28 -Because $b.Name }
    }

    It 'gives grid rows 24 to 26 px and headers 26 px in the <View> view' -ForEach @(@{ View = 'Windows' }, @{ View = 'Coverage' }, @{ View = 'Templates' }) {
        $m = $script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object View -eq $View
        foreach ($g in $m.Grids) {
            if ($null -ne $g.HeaderHeight) { $g.HeaderHeight | Should -BeLessOrEqual 26.5 -Because $g.Name }
            if ($null -ne $g.RowHeight) { [math]::Round($g.RowHeight) | Should -BeIn 24, 25, 26 -Because $g.Name }
        }
    }

    It 'lets the content fill the width in the <View> view' -ForEach @(@{ View = 'Windows' }, @{ View = 'Coverage' }, @{ View = 'Templates' }) {
        $mm = $script:Measured['{0}|{1}' -f $Theme, $Size]
        $m = $mm.Views | Where-Object View -eq $View
        $m.Width | Should -BeGreaterOrEqual ($mm.Width - 300)
        foreach ($g in $m.Grids) { $g.Width | Should -BeGreaterOrEqual ($m.Width * 0.3) -Because $g.Name }
    }

    It 'leaves no vacant area in the views (at most 120 px high and 160 px wide free)' -Skip:(-not $Design) {
        $bad = @($script:Measured['{0}|{1}' -f $Theme, $Size].Views | Where-Object { $_.FreeBottom -gt 120 -or $_.FreeRight -gt 160 } |
                ForEach-Object { '{0}: free {1} px high, {2} px wide' -f $_.View, [math]::Round($_.FreeBottom), [math]::Round($_.FreeRight) })
        $bad | Should -BeNullOrEmpty -Because ('vacant area in: ' + ($bad -join '; '))
    }
}
