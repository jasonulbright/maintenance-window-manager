#Requires -Modules Pester
# Logic of start-maintenancewindowmgr.ps1. The script cannot be dot-sourced (it opens
# the window), so each function is cut out of its syntax tree and defined here,
# with plain objects in place of the window controls.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:ShellPath = Join-Path $script:AppRoot 'start-maintenancewindowmgr.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    $wanted = 'Get-MwmPreferences', 'Save-MwmPreferences', 'Get-WindowStatusGlyph', 'Get-CoverageGlyph', 'Get-StatusFilterValue',
        'Update-Filter', 'Set-StatusText', 'Update-StatusBarSummary', 'Get-ScheduleRecurrenceTypeInt', 'Get-PreviewStartDateTime'
    foreach ($n in $wanted) {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
}

Describe 'Glyphs' {
    It 'Get-WindowStatusGlyph gives a check for an enabled window, a dot row for a disabled one, and nothing for null' {
        [int](Get-WindowStatusGlyph -Window ([PSCustomObject]@{ IsEnabled = $true })) | Should -Be 0x2713
        [int](Get-WindowStatusGlyph -Window ([PSCustomObject]@{ IsEnabled = $false })) | Should -Be 0x22EF
        Get-WindowStatusGlyph -Window $null | Should -Be ''
    }

    It 'Get-CoverageGlyph gives a warning for no window, a check otherwise, and nothing for null' {
        [int](Get-CoverageGlyph -Row ([PSCustomObject]@{ WindowCount = 0 })) | Should -Be 0x26A0
        [int](Get-CoverageGlyph -Row ([PSCustomObject]@{ WindowCount = 3 })) | Should -Be 0x2713
        [int](Get-CoverageGlyph -Row ([PSCustomObject]@{ WindowCount = $null })) | Should -Be 0x26A0
        Get-CoverageGlyph -Row $null | Should -Be ''
    }
}

Describe 'Get-ScheduleRecurrenceTypeInt' {
    It 'maps <Name> to <Value>' -ForEach @(
        @{ Name = 'One-time'; Value = 1 }, @{ Name = 'Daily'; Value = 2 }, @{ Name = 'Weekly'; Value = 3 }, @{ Name = 'MonthlyByWeekday'; Value = 4 },
        @{ Name = 'MonthlyByDate'; Value = 5 }, @{ Name = 'PatchTuesday'; Value = 4 }, @{ Name = 'Unknown'; Value = 3 }, @{ Name = ''; Value = 3 }
    ) {
        Get-ScheduleRecurrenceTypeInt -Recurrence $Name | Should -Be $Value
    }
}

Describe 'Get-StatusFilterValue' {
    It 'returns All when nothing is selected, the item content for a combo box item, and a string as it is' {
        Add-Type -AssemblyName PresentationFramework
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = $null }
        Get-StatusFilterValue | Should -Be 'All'
        $item = New-Object System.Windows.Controls.ComboBoxItem
        $item.Content = 'Enabled only'
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = $item }
        Get-StatusFilterValue | Should -Be 'Enabled only'
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = 'Disabled only' }
        Get-StatusFilterValue | Should -Be 'Disabled only'
    }
}

Describe 'Update-Filter' {
    BeforeEach {
        $script:ActiveView = 'Windows'
        $txtFilter = [PSCustomObject]@{ Text = '' }
        $cboStatusFilter = [PSCustomObject]@{ SelectedItem = 'All' }
        $gridWindows = [PSCustomObject]@{ ItemsSource = $null }
        $gridCoverage = [PSCustomObject]@{ ItemsSource = $null }
        $gridTemplates = [PSCustomObject]@{ ItemsSource = $null }
        $script:AllWindows = @(
            [PSCustomObject]@{ CollectionName = 'Servers'; WindowName = 'Night'; IsEnabled = $true;  Type = 'Software Updates' },
            [PSCustomObject]@{ CollectionName = 'Lab';     WindowName = 'Off';   IsEnabled = $false; Type = 'General' },
            [PSCustomObject]@{ CollectionName = 'Imaging'; WindowName = 'OSD';   IsEnabled = $true;  Type = 'Task Sequences' },
            [PSCustomObject]@{ CollectionName = $null;     WindowName = $null;   IsEnabled = $true;  Type = 'General' })
        $script:CoverageRows = @(
            [PSCustomObject]@{ Name = 'Servers'; CollectionID = 'MCM00001'; WindowCount = 2; IsBuiltIn = $false },
            [PSCustomObject]@{ Name = 'All Systems'; CollectionID = 'SMS00001'; WindowCount = 0; IsBuiltIn = $true },
            [PSCustomObject]@{ Name = 'Lab'; CollectionID = 'MCM00002'; WindowCount = 0; IsBuiltIn = $false })
        $script:Templates = @([PSCustomObject]@{ Name = 'Patch Tuesday' }, [PSCustomObject]@{ Name = 'Weekend' })
    }

    It 'Windows view: text "<Text>" and filter "<Filter>" show <Expected>' -ForEach @(
        @{ Text = '';        Filter = 'All';                 Expected = 'Night,Off,OSD,' }
        @{ Text = 'SERVERS'; Filter = 'All';                 Expected = 'Night' }
        @{ Text = ' osd ';   Filter = 'All';                 Expected = 'OSD' }
        @{ Text = 'zzz';     Filter = 'All';                 Expected = '' }
        @{ Text = '';        Filter = 'Enabled only';        Expected = 'Night,OSD,' }
        @{ Text = '';        Filter = 'Disabled only';       Expected = 'Off' }
        @{ Text = '';        Filter = 'SoftwareUpdates only'; Expected = 'Night' }
        @{ Text = '';        Filter = 'TaskSequences only';  Expected = 'OSD' }
        @{ Text = 'lab';     Filter = 'Disabled only';       Expected = 'Off' }
        @{ Text = '';        Filter = 'Other';               Expected = 'Night,Off,OSD,' }
    ) {
        $txtFilter.Text = $Text
        $cboStatusFilter.SelectedItem = $Filter
        Update-Filter
        (@($gridWindows.ItemsSource) | ForEach-Object WindowName) -join ',' | Should -Be $Expected
    }

    It 'Coverage view: text "<Text>" and filter "<Filter>" show <Expected>' -ForEach @(
        @{ Text = '';         Filter = 'All';             Expected = 'MCM00001,SMS00001,MCM00002' }
        @{ Text = 'sms';      Filter = 'All';             Expected = 'SMS00001' }
        @{ Text = 'lab';      Filter = 'All';             Expected = 'MCM00002' }
        @{ Text = '';         Filter = 'Without windows'; Expected = 'SMS00001,MCM00002' }
        @{ Text = '';         Filter = 'With windows';    Expected = 'MCM00001' }
        @{ Text = '';         Filter = 'Custom only';     Expected = 'MCM00001,MCM00002' }
        @{ Text = '';         Filter = 'Built-in only';   Expected = 'SMS00001' }
        @{ Text = 'mcm';      Filter = 'Without windows'; Expected = 'MCM00002' }
    ) {
        $script:ActiveView = 'Coverage'
        $txtFilter.Text = $Text
        $cboStatusFilter.SelectedItem = $Filter
        Update-Filter
        (@($gridCoverage.ItemsSource) | ForEach-Object CollectionID) -join ',' | Should -Be $Expected
    }

    It 'Templates view: text "<Text>" shows <Count> templates' -ForEach @(@{ Text = ''; Count = 2 }, @{ Text = 'patch'; Count = 1 }, @{ Text = 'END'; Count = 1 }, @{ Text = 'x'; Count = 0 }) {
        $script:ActiveView = 'Templates'
        $txtFilter.Text = $Text
        Update-Filter
        @($gridTemplates.ItemsSource).Count | Should -Be $Count
    }

    It 'treats the filter text <Text> as plain text' -ForEach @(@{ Text = '.*' }, @{ Text = '[' }, @{ Text = '(' }, @{ Text = '%' }, @{ Text = "'" }) {
        $txtFilter.Text = $Text
        { Update-Filter } | Should -Not -Throw
        @($gridWindows.ItemsSource).Count | Should -Be 0
    }
}

Describe 'Update-StatusBarSummary' {
    BeforeEach {
        function Set-StatusText { param([string]$Text) $script:StatusText = $Text }
        $global:Prefs = @{ SiteCode = 'MCM'; SMSProvider = 'cm01' }
        $script:IsConnectedFromBg = $false
        $script:Collections = @(); $script:RawWindows = @(); $script:Templates = @(); $script:LastRefreshTime = $null
        $script:StatusText = $null
    }

    AfterAll { Remove-Variable -Name Prefs -Scope Global -ErrorAction SilentlyContinue }

    It 'says what to do when <Case>' -ForEach @(
        @{ Case = 'the site code is missing'; Site = ''; Provider = 'cm01'; Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'the provider is missing'; Site = 'MCM'; Provider = ''; Expected = 'Open Options to configure site code and SMS provider' }
        @{ Case = 'both are set'; Site = 'MCM'; Provider = 'cm01'; Expected = 'Ready. Click Refresh.' }
    ) {
        $global:Prefs = @{ SiteCode = $Site; SMSProvider = $Provider }
        Update-StatusBarSummary
        $script:StatusText | Should -Be $Expected
    }

    It 'shows the counts and the refresh time when connected' {
        $script:IsConnectedFromBg = $true
        $script:Collections = @(1, 2); $script:RawWindows = @(1, 2, 3); $script:Templates = @(1)
        $script:LastRefreshTime = [datetime]'2026-11-06T07:06:05'
        Update-StatusBarSummary
        $script:StatusText | Should -Be 'Connected to MCM   |   2 collections   |   3 windows   |   1 templates   |   last refresh 07:06:05'
    }
}

Describe 'Preferences file' {
    BeforeEach { $global:PrefsPath = Join-Path $TestDrive ('prefs-{0}.json' -f [guid]::NewGuid().ToString('N')) }
    AfterAll { Remove-Variable -Name PrefsPath -Scope Global -ErrorAction SilentlyContinue }

    It 'gives the defaults on first run' {
        $p = Get-MwmPreferences
        $p.DarkMode | Should -BeTrue
        $p.SiteCode | Should -Be ''
        $p.SMSProvider | Should -Be ''
    }

    It 'gives the defaults for a <Case> file' -ForEach @(@{ Case = 'corrupt'; Content = '{ not json' }, @{ Case = 'empty'; Content = '' }, @{ Case = 'array'; Content = '[1,2,3]' }) {
        Set-Content -LiteralPath $global:PrefsPath -Value $Content
        (Get-MwmPreferences).SiteCode | Should -Be ''
    }

    It 'ignores a key it does not know' {
        Set-Content -LiteralPath $global:PrefsPath -Value '{ "SiteCode": "MCM", "Evil": "x" }'
        $p = Get-MwmPreferences
        $p.SiteCode | Should -Be 'MCM'
        $p.ContainsKey('Evil') | Should -BeFalse
    }

    It 'round-trips the value <Value>' -ForEach @(@{ Value = 'MCM' }, @{ Value = 'Ünï-日本' }, @{ Value = 'a "q" \ b' }, @{ Value = ('x' * 5000) }) {
        Save-MwmPreferences -Prefs @{ DarkMode = $false; SiteCode = $Value; SMSProvider = 'cm01' }
        $p = Get-MwmPreferences
        $p.SiteCode | Should -Be $Value
        $p.DarkMode | Should -BeFalse
    }

    It 'does not throw when the file cannot be written' {
        $global:PrefsPath = Join-Path $TestDrive 'dir-in-the-way'
        New-Item -ItemType Directory -Path $global:PrefsPath | Out-Null
        { Save-MwmPreferences -Prefs @{ DarkMode = $true; SiteCode = 'X'; SMSProvider = '' } } | Should -Not -Throw
    }
}

Describe 'Script contract' {
    It 'calls only commands that exist' {
        $own = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object Name)
        $called = @($script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] }, $true) |
                ForEach-Object { $_.GetCommandName() } | Where-Object { $_ } | Sort-Object -Unique)
        Import-Module (Join-Path $script:AppRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking -ErrorAction SilentlyContinue
        $missing = @($called | Where-Object { $_ -notin $own -and -not (Get-Command $_ -ErrorAction SilentlyContinue) })
        $missing | Where-Object { $_ -notmatch '^(Get|Set|New|Remove|Start|Invoke|Add|Copy|Convert)-CM' } | Should -BeNullOrEmpty
    }
}

Describe 'Get-PreviewStartDateTime' {
    BeforeAll {
        function New-State { param($Recurrence, $Dow = 'Tuesday', $Wo = 'Second', $Dom = 1) @{ Recurrence = $Recurrence; StartDate = [datetime]'2026-11-07'; Hour = 2; Minute = 30; DayOfWeek = $Dow; WeekOrder = $Wo; DayOfMonth = $Dom } }
    }

    It 'moves a weekly schedule to the selected day of week, not the weekday of the start date' {
        # 2026-11-07 is a Saturday. The editor shows Tuesday.
        (Get-PreviewStartDateTime -State (New-State 'Weekly' 'Tuesday')).ToString('yyyy-MM-dd ddd HH:mm') | Should -Be '2026-11-10 Tue 02:30'
    }

    It 'keeps the start date when it already has the selected day of week' {
        (Get-PreviewStartDateTime -State (New-State 'Weekly' 'Saturday')).ToString('yyyy-MM-dd') | Should -Be '2026-11-07'
    }

    It 'uses the selected day of the month for a monthly by date schedule' {
        (Get-PreviewStartDateTime -State (New-State 'MonthlyByDate' -Dom 15)).ToString('yyyy-MM-dd') | Should -Be '2026-11-15'
    }

    It 'cuts the day of the month to the last day of a short month' {
        (Get-PreviewStartDateTime -State (New-State 'MonthlyByDate' -Dom 31)).ToString('yyyy-MM-dd') | Should -Be '2026-11-30'
    }

    It 'uses the selected week order and day of week for a monthly by weekday schedule' {
        (Get-PreviewStartDateTime -State (New-State 'MonthlyByWeekday' 'Tuesday' 'Second')).ToString('yyyy-MM-dd') | Should -Be '2026-11-10'
        (Get-PreviewStartDateTime -State (New-State 'MonthlyByWeekday' 'Friday' 'Last')).ToString('yyyy-MM-dd') | Should -Be '2026-11-27'
    }

    It 'leaves a daily schedule on the start date' {
        (Get-PreviewStartDateTime -State (New-State 'Daily')).ToString('yyyy-MM-dd HH:mm') | Should -Be '2026-11-07 02:30'
    }
}