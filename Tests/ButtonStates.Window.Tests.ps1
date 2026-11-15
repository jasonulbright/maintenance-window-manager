#Requires -Modules Pester
# Window test for the hover, focus and pressed states of the buttons of the main window, in both themes.
# It needs the real mouse, so it runs only in a visible run (SUITE_VISIBLE_WINDOWS=1, the labadmin session of CLIENT01).
# The measurement is Get-SuiteButtonStateReport of SuiteCommon. For each sampled button and each state it requires:
#   - the button text has 4.5:1 contrast on the button fill (the backdrop of the button where the fill is transparent)
#   - the button is told apart from its backdrop: its border or its fill has 3:1 contrast on it

BeforeDiscovery {
    . "$PSScriptRoot\WindowHarness.ps1"
    $guard = Test-WindowTestsAllowed
    $script:Visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    if ($guard.Allowed -and -not $script:Visible) { Write-Host 'Button state tests skipped: SUITE_VISIBLE_WINDOWS is not 1 (they need the real mouse).' }
    # The window guard of this file also needs the visible run, because the test moves the real mouse.
    $script:WindowGuard = [pscustomobject]@{ Allowed = ($guard.Allowed -and $script:Visible); Reason = $guard.Reason }
    $script:ThemeCases = @(@{ Dark = $true; Theme = 'dark' }, @{ Dark = $false; Theme = 'light' })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    $guard = Test-WindowTestsAllowed
    $script:Visible = ($env:SUITE_VISIBLE_WINDOWS -eq '1')
    $script:WindowGuard = [pscustomobject]@{ Allowed = ($guard.Allowed -and $script:Visible); Reason = $guard.Reason }
    if (-not $script:WindowGuard.Allowed) { return }
    . "$PSScriptRoot\WindowConfig.ps1"
    $script:Measured = @{}
    foreach ($dark in $true, $false) {
        $key = if ($dark) { 'dark' } else { 'light' }
        $script:Measured[$key] = Invoke-WpfSta -Arguments @{ Dark = $dark; Config = $script:AppConfig } -Script {
            $cfg = $Arguments.Config
            $w = Open-TestWindow -XamlPath (Join-Path $Root $cfg.Xaml) -Dark $Arguments.Dark
            Initialize-ShellTheme -Window $w -Dark $Arguments.Dark -Views $cfg.Views -OptionsButtonName $cfg.OptionsButton -LogLabelName $cfg.LogLabel -ActiveView $cfg.Views[0].Name
            foreach ($v in $cfg.Views) { $w.FindName($v.Container).Visibility = if ($v.Name -eq $cfg.Views[0].Name) { 'Visible' } else { 'Collapsed' } }
            Wait-Layout $w
            try { Get-SuiteButtonStateReport -Window $w -Dark $Arguments.Dark }
            finally { $w.Close() }
        }
    }
}

Describe 'Button states <Theme>' -ForEach $script:ThemeCases -Skip:(-not $script:WindowGuard.Allowed) {
    BeforeAll { $script:Rows = @($script:Measured[$Theme] | ForEach-Object { $_ } | Where-Object { $_.PSObject.Properties['States'] }) }

    It 'samples at least 3 buttons and reaches the focus, hover and pressed states' {
        $script:Rows.Count | Should -BeGreaterOrEqual 3
        $bad = @($script:Rows | Where-Object { -not $_.Hover -or -not $_.Pressed -or -not $_.Focused } | ForEach-Object { '{0}: hover {1} pressed {2} focused {3}' -f $_.Name, $_.Hover, $_.Pressed, $_.Focused })
        $bad | Should -BeNullOrEmpty -Because 'the mouse and the keyboard focus must reach the button, or the states are not measured'
    }

    It 'keeps 4.5:1 text contrast in every state' {
        $bad = @()
        foreach ($r in $script:Rows) { foreach ($s in $r.States.Keys) { if ($r.States[$s].Text -lt 4.5) { $bad += ('{0} {1}: {2}' -f $r.Name, $s, $r.States[$s].Text) } } }
        $bad | Should -BeNullOrEmpty
    }

    It 'tells the button from its backdrop (3:1) in every state' {
        $bad = @()
        foreach ($r in $script:Rows) { foreach ($s in $r.States.Keys) { if ($r.States[$s].Edge -lt 3.0) { $bad += ('{0} {1}: {2}' -f $r.Name, $s, $r.States[$s].Edge) } } }
        $bad | Should -BeNullOrEmpty
    }

    It 'lists the measured values' {
        foreach ($r in $script:Rows) { Write-Host ('{0} {1}: ' -f $Theme, $r.Name) -NoNewline; Write-Host (($r.States.Keys | ForEach-Object { '{0} text {1} edge {2}' -f $_, $r.States[$_].Text, $r.States[$_].Edge }) -join ' | ') }
        $script:Rows.Count | Should -BeGreaterThan 0
    }
}
