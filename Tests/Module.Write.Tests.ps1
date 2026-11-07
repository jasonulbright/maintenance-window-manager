#Requires -Modules Pester
# Write actions: window create, update, remove, template apply, copy, bulk enable.
# Templates and CSV import. ConfigMgr cmdlets are stubs and mocks. Covers success,
# failure, partial failure, -WhatIf and hostile input.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name New-CMMaintenanceWindow, Set-CMMaintenanceWindow, Remove-CMMaintenanceWindow, Get-CMMaintenanceWindow, Convert-CMSchedule, New-CMSchedule
    $global:TestLogs = [System.Collections.ArrayList]::new()
    Mock Write-Log { [void]$global:TestLogs.Add(@{ Message = $Message; Level = $Level }) } -ModuleName MaintWindowMgrCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx, TestLogs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'New-ManagedMaintenanceWindow' {
    BeforeEach {
        $global:TestLogs.Clear()
        Mock New-CMMaintenanceWindow { [PSCustomObject]@{ Name = $Name } } -ModuleName MaintWindowMgrCommon
    }

    It 'creates the window and returns it' {
        $r = New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name 'Night' -Schedule 'sched' -ApplyTo SoftwareUpdatesOnly -IsEnabled $false -IsUtc $true
        $r.Name | Should -Be 'Night'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 1 -Exactly -ParameterFilter {
            $CollectionId -eq 'MCM00001' -and $Name -eq 'Night' -and $Schedule -eq 'sched' -and $ApplyTo -eq 'SoftwareUpdatesOnly' -and $IsEnabled -eq $false -and $IsUtc -eq $true
        }
    }

    It 'defaults to any deployment, enabled, local time' {
        $null = New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name 'N' -Schedule 's'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $ApplyTo -eq 'Any' -and $IsEnabled -eq $true -and $IsUtc -eq $false }
    }

    It 'returns null for the blank name "<Name>" and does not call the site' -ForEach @(@{ Name = ' ' }, @{ Name = "`t" }, @{ Name = '   ' }) {
        New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name $Name -Schedule 's' | Should -BeNullOrEmpty
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
        @($global:TestLogs | Where-Object Level -eq 'ERROR').Count | Should -Be 1
    }

    It 'returns null and logs the message when the site refuses' {
        Mock New-CMMaintenanceWindow { throw 'name already exists' } -ModuleName MaintWindowMgrCommon
        New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name 'N' -Schedule 's' | Should -BeNullOrEmpty
        @($global:TestLogs | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -like '*name already exists*' }).Count | Should -Be 1
    }

    It 'passes the hostile name <Name> through unchanged' -ForEach @(
        @{ Name = "O'Brien ""Q"" <b>" }, @{ Name = ('n' * 400) }, @{ Name = 'Ünï 日本語' }, @{ Name = '..\x' }
    ) {
        $global:Fx = $Name
        $null = New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name $Name -Schedule 's'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $Name -ceq $global:Fx }
    }

    It 'rejects an unknown ApplyTo value' {
        { New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name 'N' -Schedule 's' -ApplyTo Everything } | Should -Throw
    }

    It 'creates nothing under -WhatIf' {
        New-ManagedMaintenanceWindow -CollectionId 'MCM00001' -Name 'N' -Schedule 's' -WhatIf | Should -BeNullOrEmpty
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }
}

Describe 'Set-ManagedMaintenanceWindow' {
    BeforeEach { Mock Set-CMMaintenanceWindow { } -ModuleName MaintWindowMgrCommon }

    It 'sends only the properties that are given' {
        Set-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' | Should -BeTrue
        Should -Invoke Set-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $null -eq $Schedule -and $null -eq $ApplyTo -and $null -eq $IsEnabled -and $null -eq $IsUtc }
    }

    It 'sends a schedule, an apply-to value and both flags' {
        Set-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' -Schedule 's' -ApplyTo TaskSequencesOnly -IsEnabled $false -IsUtc 1 | Should -BeTrue
        Should -Invoke Set-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $Schedule -eq 's' -and $ApplyTo -eq 'TaskSequencesOnly' -and $IsEnabled -eq $false -and $IsUtc -eq $true }
    }

    It 'returns false and logs an error when the site refuses' {
        Mock Set-CMMaintenanceWindow { throw 'window not found' } -ModuleName MaintWindowMgrCommon
        $global:TestLogs.Clear()
        Set-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' -IsEnabled $true | Should -BeFalse
        @($global:TestLogs | Where-Object { $_.Level -eq 'ERROR' -and $_.Message -like '*window not found*' }).Count | Should -Be 1
    }

    It 'changes nothing under -WhatIf' {
        Set-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' -IsEnabled $true -WhatIf | Should -BeFalse
        Should -Invoke Set-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }
}

Describe 'Remove-ManagedMaintenanceWindow' {
    BeforeEach { Mock Remove-CMMaintenanceWindow { } -ModuleName MaintWindowMgrCommon }

    It 'removes the window with -Force' {
        Remove-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' | Should -BeTrue
        Should -Invoke Remove-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 1 -Exactly -ParameterFilter { $CollectionId -eq 'MCM00001' -and $MaintenanceWindowName -eq 'W' -and $Force }
    }

    It 'returns false when the site refuses' {
        Mock Remove-CMMaintenanceWindow { throw 'no such window' } -ModuleName MaintWindowMgrCommon
        Remove-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' | Should -BeFalse
    }

    It 'removes nothing under -WhatIf' {
        Remove-ManagedMaintenanceWindow -CollectionId 'MCM00001' -MaintenanceWindowName 'W' -WhatIf | Should -BeFalse
        Should -Invoke Remove-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }

    It 'rejects an empty <Field>' -ForEach @(
        @{ Field = 'collection id'; P = @{ CollectionId = ''; MaintenanceWindowName = 'W' } }
        @{ Field = 'window name'; P = @{ CollectionId = 'MCM00001'; MaintenanceWindowName = '' } }
    ) {
        { Remove-ManagedMaintenanceWindow @P } | Should -Throw
    }
}

Describe 'Window template files' {
    BeforeEach { $script:dir = Join-Path $TestDrive ('tpl-{0}' -f [guid]::NewGuid().ToString('N')) }

    It 'saves a template as JSON and reads it back' {
        $path = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch Tuesday + 7' -RecurrenceType MonthlyByWeekday -DurationHours 3 -DurationMinutes 30 -WindowType SoftwareUpdatesOnly -IsUtc $true -Description 'd' -StartHour 5 -StartMinute 15 -DayOfWeek Tuesday -WeekOrder Second -DayOfMonth 9 -RecurCount 2 -PatchTuesdayOffset 7
        $path | Should -Be (Join-Path $script:dir 'patch-tuesday---7.json')
        $t = @(Get-WindowTemplates -TemplatesPath $script:dir)
        $t.Count | Should -Be 1
        $t[0].Name | Should -Be 'Patch Tuesday + 7'
        $t[0].RecurrenceType | Should -Be 'MonthlyByWeekday'
        $t[0].DurationMinutes | Should -Be 30
        $t[0].DayOfWeek | Should -Be 'Tuesday'
        $t[0].IsUtc | Should -BeTrue
        $t[0].PatchTuesdayOffset | Should -Be 7
        $t[0].FileName | Should -Be 'patch-tuesday---7.json'
        $t[0].FilePath | Should -Be $path
    }

    It 'creates the folder when it does not exist' {
        $null = Save-WindowTemplate -TemplatesPath (Join-Path $script:dir 'a\b') -Name 'X' -RecurrenceType Daily
        Test-Path -LiteralPath (Join-Path $script:dir 'a\b\x.json') | Should -BeTrue
    }

    It 'replaces the file when the same name is saved again' {
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Same' -RecurrenceType Daily -DurationHours 1
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Same' -RecurrenceType Daily -DurationHours 6
        $t = @(Get-WindowTemplates -TemplatesPath $script:dir)
        $t.Count | Should -Be 1
        $t[0].DurationHours | Should -Be 6
    }

    It 'keeps two names that give the same file name as two templates' {
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch + 7' -RecurrenceType Daily -DurationHours 1
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch - 7' -RecurrenceType Daily -DurationHours 2
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch . 7' -RecurrenceType Daily -DurationHours 3
        $t = @(Get-WindowTemplates -TemplatesPath $script:dir)
        $t.Count | Should -Be 3
        foreach ($n in 'Patch + 7', 'Patch - 7', 'Patch . 7') { $t.Name | Should -Contain $n }
    }

    It 'saves the device-like name <Name> under a safe file name' -ForEach @(@{ Name = 'CON' }, @{ Name = 'nul' }, @{ Name = 'COM1' }, @{ Name = 'Lpt9' }) {
        $path = Save-WindowTemplate -TemplatesPath $script:dir -Name $Name -RecurrenceType Daily
        (Split-Path $path -Leaf) | Should -Be ('_{0}.json' -f $Name.ToLower())
        @(Get-WindowTemplates -TemplatesPath $script:dir).Count | Should -Be 1
    }

    It 'keeps the name of a hostile template verbatim and makes a safe file name: <Name>' -ForEach @(
        @{ Name = '..\..\evil' }, @{ Name = 'a/b:c*d?e' }, @{ Name = 'Ünï 日本' }, @{ Name = ('x' * 100) }, @{ Name = '!!!' }
    ) {
        $path = Save-WindowTemplate -TemplatesPath $script:dir -Name $Name -RecurrenceType Daily
        (Split-Path $path -Parent) | Should -Be $script:dir
        (Split-Path $path -Leaf) | Should -Match '^[\w\-]+\.json$'
        (@(Get-WindowTemplates -TemplatesPath $script:dir))[0].Name | Should -Be $Name
    }

    It 'saves nothing under -WhatIf' {
        Save-WindowTemplate -TemplatesPath $script:dir -Name 'N' -RecurrenceType Daily -WhatIf | Should -BeNullOrEmpty
        Test-Path -LiteralPath $script:dir | Should -BeFalse
    }

    It 'returns no template for a folder that does not exist' {
        @(Get-WindowTemplates -TemplatesPath (Join-Path $script:dir 'nope')).Count | Should -Be 0
    }

    It 'skips a corrupt or empty file and logs a warning' {
        New-Item -ItemType Directory -Path $script:dir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:dir 'bad.json') -Value '{ not json'
        Set-Content -LiteralPath (Join-Path $script:dir 'empty.json') -Value ''
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Good' -RecurrenceType Daily
        $global:TestLogs.Clear()
        @(Get-WindowTemplates -TemplatesPath $script:dir).Count | Should -Be 1
        @($global:TestLogs | Where-Object Level -eq 'WARN').Count | Should -Be 2
    }

    It 'reads unicode from a file without a byte order mark' {
        New-Item -ItemType Directory -Path $script:dir -Force | Out-Null
        [IO.File]::WriteAllText((Join-Path $script:dir 'u.json'), '{"Name":"Ünï 日本","RecurrenceType":"Daily"}', (New-Object Text.UTF8Encoding $false))
        (@(Get-WindowTemplates -TemplatesPath $script:dir))[0].Name | Should -Be 'Ünï 日本'
    }

    It 'ignores files that are not JSON' {
        New-Item -ItemType Directory -Path $script:dir -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $script:dir 'readme.txt') -Value 'x'
        @(Get-WindowTemplates -TemplatesPath $script:dir).Count | Should -Be 0
    }

    It 'removes a template file and returns true, then false for a missing file' {
        $path = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Gone' -RecurrenceType Daily
        Remove-WindowTemplate -FilePath $path | Should -BeTrue
        Test-Path -LiteralPath $path | Should -BeFalse
        Remove-WindowTemplate -FilePath $path | Should -BeFalse
    }

    It 'refuses to remove a file that is not a .json file: <Name>' -ForEach @(@{ Name = 'app.dll' }, @{ Name = 'notes.txt' }, @{ Name = 'noextension' }) {
        New-Item -ItemType Directory -Path $script:dir -Force | Out-Null
        $f = Join-Path $script:dir $Name
        Set-Content -LiteralPath $f -Value 'x'
        Remove-WindowTemplate -FilePath $f | Should -BeFalse
        Test-Path -LiteralPath $f | Should -BeTrue
    }

    It 'replaces the numbered file of a template when its base file is gone' {
        $null = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch + 7' -RecurrenceType Daily -DurationHours 1
        $second = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch - 7' -RecurrenceType Daily -DurationHours 2
        Remove-Item -LiteralPath (Join-Path $script:dir 'patch---7.json') -Force
        $again = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Patch - 7' -RecurrenceType Daily -DurationHours 9
        $again | Should -Be $second
        @(Get-ChildItem -LiteralPath $script:dir -Filter '*.json').Count | Should -Be 1
        (Get-Content -LiteralPath $again -Raw | ConvertFrom-Json).DurationHours | Should -Be 9
    }

    It 'removes nothing under -WhatIf' {
        $path = Save-WindowTemplate -TemplatesPath $script:dir -Name 'Keep' -RecurrenceType Daily
        Remove-WindowTemplate -FilePath $path -WhatIf | Should -BeFalse
        Test-Path -LiteralPath $path | Should -BeTrue
    }
}

Describe 'New-WindowFromTemplate' {
    BeforeEach {
        $global:TestLogs.Clear()
        $script:tpl = [PSCustomObject]@{ Name = 'Night'; RecurrenceType = 'Weekly'; StartHour = 2; StartMinute = 30; DurationHours = 3; DurationMinutes = 0; DayOfWeek = 'Friday'; WeekOrder = 'First'; DayOfMonth = 1; RecurCount = 1; WindowType = 'SoftwareUpdatesOnly'; IsUtc = $false }
        Mock New-CMSchedule { 'schedule' } -ModuleName MaintWindowMgrCommon
        Mock New-CMMaintenanceWindow { [PSCustomObject]@{ Name = $Name } } -ModuleName MaintWindowMgrCommon
    }

    It 'creates the window on each target and reports success per collection' {
        $r = @(New-WindowFromTemplate -Template $script:tpl -TargetCollections @([PSCustomObject]@{ CollectionID = 'A'; Name = 'Alpha' }, 'B'))
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,True'
        ($r | ForEach-Object CollectionName) -join ',' | Should -Be 'Alpha,B'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 2 -Exactly -ParameterFilter { $Name -eq 'Night' -and $ApplyTo -eq 'SoftwareUpdatesOnly' -and $Schedule -eq 'schedule' }
    }

    It 'does not prompt a second time for each collection' {
        Mock New-ManagedMaintenanceWindow { [PSCustomObject]@{ Name = $Name } } -ModuleName MaintWindowMgrCommon
        $null = New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A', 'B')
        Should -Invoke New-ManagedMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 2 -Exactly -ParameterFilter { $PesterBoundParameters.ContainsKey('Confirm') -and -not $Confirm }
    }

    It 'uses the name override' {
        $null = New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A') -WindowNameOverride 'Special'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $Name -eq 'Special' }
    }

    It 'maps window type <Type> to apply-to <Expected>' -ForEach @(
        @{ Type = 'SoftwareUpdatesOnly'; Expected = 'SoftwareUpdatesOnly' }, @{ Type = 'TaskSequencesOnly'; Expected = 'TaskSequencesOnly' }, @{ Type = 'Any'; Expected = 'Any' }, @{ Type = 'Other'; Expected = 'Any' }, @{ Type = $null; Expected = 'Any' }
    ) {
        $script:tpl.WindowType = $Type
        $null = New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A')
        $global:ExpectedApplyTo = $Expected
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 1 -Exactly -ParameterFilter { $ApplyTo -ceq $global:ExpectedApplyTo }
    }

    It 'reports a failure for the collection that the site refuses and continues (partial failure)' {
        Mock New-CMMaintenanceWindow { if ($CollectionId -eq 'B') { throw 'refused' } else { [PSCustomObject]@{ Name = $Name } } } -ModuleName MaintWindowMgrCommon
        $r = @(New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A', 'B', 'C'))
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,False,True'
    }

    It 'returns no result when the schedule cannot be made' {
        $script:tpl.DurationHours = 0
        @(New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A')).Count | Should -Be 0
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }

    It 'creates nothing under -WhatIf' {
        @(New-WindowFromTemplate -Template $script:tpl -TargetCollections @('A', 'B') -WhatIf).Count | Should -Be 0
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }
}

Describe 'Import-MaintenanceWindowsCsv' {
    BeforeEach { $script:csv = Join-Path $TestDrive ('w-{0}.csv' -f [guid]::NewGuid().ToString('N')) }

    It 'returns nothing for a missing file' {
        @(Import-MaintenanceWindowsCsv -CsvPath (Join-Path $TestDrive 'missing.csv')).Count | Should -Be 0
    }

    It 'reads a complete valid row' {
        Set-Content -LiteralPath $script:csv -Value @(
            'CollectionID,CollectionName,WindowName,RecurrenceType,StartHour,StartMinute,DurationHours,DurationMinutes,WindowType,DayOfWeek,WeekOrder,DayOfMonth,RecurCount,IsUtc'
            'MCM00001,Servers,Night,MonthlyByWeekday,3,15,2,30,SoftwareUpdatesOnly,Tuesday,Second,9,2,true')
        $r = @(Import-MaintenanceWindowsCsv -CsvPath $script:csv)
        $r.Count | Should -Be 1
        $r[0].Valid | Should -BeTrue
        $r[0].Error | Should -Be ''
        $r[0].CollectionID | Should -Be 'MCM00001'
        $r[0].RecurrenceType | Should -Be 'MonthlyByWeekday'
        $r[0].StartHour | Should -Be 3
        $r[0].DurationMinutes | Should -Be 30
        $r[0].DayOfWeek | Should -Be 'Tuesday'
        $r[0].IsUtc | Should -BeTrue
    }

    It 'fills the defaults for a row with only a collection id' {
        Set-Content -LiteralPath $script:csv -Value @('CollectionID', 'MCM00001')
        $r = (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0]
        $r.Valid | Should -BeTrue
        $r.WindowName | Should -Be 'Maintenance Window'
        $r.RecurrenceType | Should -Be 'Weekly'
        $r.StartHour | Should -Be 2
        $r.DurationHours | Should -Be 4
        $r.DayOfWeek | Should -Be 'Sunday'
        $r.IsUtc | Should -BeFalse
    }

    It 'marks the row invalid for <Case>' -ForEach @(
        @{ Case = 'no collection';            Header = 'WindowName';                  Row = 'x';                       Fragment = 'CollectionID or CollectionName is required' }
        @{ Case = 'a text start hour';        Header = 'CollectionID,StartHour';      Row = 'C1,abc';                  Fragment = "StartHour 'abc' is not a number" }
        @{ Case = 'a start hour of 25';       Header = 'CollectionID,StartHour';      Row = 'C1,25';                   Fragment = 'StartHour 25 is not between 0 and 23' }
        @{ Case = 'a negative start minute';  Header = 'CollectionID,StartMinute';    Row = 'C1,-5';                   Fragment = 'StartMinute -5 is not between 0 and 59' }
        @{ Case = 'a start minute of 60';     Header = 'CollectionID,StartMinute';    Row = 'C1,60';                   Fragment = 'StartMinute 60 is not between 0 and 59' }
        @{ Case = 'a zero duration';          Header = 'CollectionID,DurationHours,DurationMinutes'; Row = 'C1,0,0';  Fragment = 'The duration must be at least 1 minute' }
        @{ Case = 'a duration of 24 hours';   Header = 'CollectionID,DurationHours';  Row = 'C1,24';                   Fragment = 'DurationHours 24 is not between 0 and 23' }
        @{ Case = 'an unknown recurrence';    Header = 'CollectionID,RecurrenceType'; Row = 'C1,Hourly';               Fragment = "RecurrenceType 'Hourly' is not one of" }
        @{ Case = 'an unknown day';           Header = 'CollectionID,DayOfWeek';      Row = 'C1,Funday';               Fragment = "DayOfWeek 'Funday' is not one of" }
        @{ Case = 'an unknown week order';    Header = 'CollectionID,WeekOrder';      Row = 'C1,Fifth';                Fragment = "WeekOrder 'Fifth' is not one of" }
        @{ Case = 'an unknown window type';   Header = 'CollectionID,WindowType';     Row = 'C1,Everything';           Fragment = "WindowType 'Everything' is not one of" }
        @{ Case = 'a day of month of 32';     Header = 'CollectionID,DayOfMonth';     Row = 'C1,32';                   Fragment = 'DayOfMonth 32 is not between 1 and 31' }
        @{ Case = 'a day of month of 0';      Header = 'CollectionID,DayOfMonth';     Row = 'C1,0';                    Fragment = 'DayOfMonth 0 is not between 1 and 31' }
    ) {
        Set-Content -LiteralPath $script:csv -Value @($Header, $Row)
        $r = (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0]
        $r.Valid | Should -BeFalse
        $r.Error | Should -BeLike "*$Fragment*"
    }

    It 'matches names without regard to case and returns the canonical spelling' {
        Set-Content -LiteralPath $script:csv -Value @('CollectionID,RecurrenceType,DayOfWeek,WeekOrder,WindowType', 'C1,weekly,sUnDaY,LAST,taskSEQUENCESonly')
        $r = (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0]
        $r.Valid | Should -BeTrue
        $r.RecurrenceType | Should -Be 'Weekly'
        $r.DayOfWeek | Should -Be 'Sunday'
        $r.WeekOrder | Should -Be 'Last'
        $r.WindowType | Should -Be 'TaskSequencesOnly'
    }

    It 'lists every problem of a row' {
        Set-Content -LiteralPath $script:csv -Value @('StartHour,DayOfWeek', '99,Nope')
        $e = (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0].Error
        $e | Should -BeLike '*required*'
        $e | Should -BeLike '*StartHour*'
        $e | Should -BeLike '*DayOfWeek*'
    }

    It 'checks every row separately and does not throw for a bad row' {
        Set-Content -LiteralPath $script:csv -Value @('CollectionID,StartHour', 'C1,3', 'C2,xx', 'C3,23')
        $r = @(Import-MaintenanceWindowsCsv -CsvPath $script:csv)
        ($r | ForEach-Object Valid) -join ',' | Should -Be 'True,False,True'
    }

    It 'reads IsUtc value <Raw> as <Expected>' -ForEach @(@{ Raw = 'true'; Expected = $true }, @{ Raw = '1'; Expected = $true }, @{ Raw = 'false'; Expected = $false }, @{ Raw = '0'; Expected = $false }, @{ Raw = 'yes'; Expected = $false }) {
        Set-Content -LiteralPath $script:csv -Value @('CollectionID,IsUtc', "C1,$Raw")
        (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0].IsUtc | Should -Be $Expected
    }

    It 'reads unicode and quoted values' {
        [IO.File]::WriteAllText($script:csv, "CollectionName,WindowName`r`n`"Ünï, `"`"q`"`"`",日本`r`n", (New-Object Text.UTF8Encoding $false))
        $r = (@(Import-MaintenanceWindowsCsv -CsvPath $script:csv))[0]
        $r.CollectionName | Should -Be 'Ünï, "q"'
        $r.WindowName | Should -Be '日本'
    }

    It 'reads 2000 rows' {
        $lines = @('CollectionID,WindowName') + @(1..2000 | ForEach-Object { "C$_,W$_" })
        Set-Content -LiteralPath $script:csv -Value $lines
        @(Import-MaintenanceWindowsCsv -CsvPath $script:csv).Count | Should -Be 2000
    }
}

Describe 'Copy-MaintenanceWindowToCollections' {
    BeforeEach {
        $script:src = [PSCustomObject]@{ Name = 'Night'; ServiceWindowSchedules = 'abc'; ServiceWindowType = 4; IsEnabled = $true; IsGMT = $true }
        Mock Get-CMMaintenanceWindow { $script:src } -ModuleName MaintWindowMgrCommon
        $global:Fx = $script:src
        Mock Get-CMMaintenanceWindow { $global:Fx } -ModuleName MaintWindowMgrCommon
        Mock Convert-CMSchedule { 'schedule' } -ModuleName MaintWindowMgrCommon
        Mock New-CMMaintenanceWindow { } -ModuleName MaintWindowMgrCommon
    }

    It 'copies the window to each target with its type, state and time zone' {
        $r = @(Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'Night' -TargetCollectionIds @('A', 'B'))
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,True'
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 2 -Exactly -ParameterFilter { $Name -eq 'Night' -and $ApplyTo -eq 'SoftwareUpdatesOnly' -and $IsEnabled -eq $true -and $IsUtc -eq $true -and $Schedule -eq 'schedule' }
    }

    It 'maps window type <Type> to <Expected>' -ForEach @(@{ Type = 4; Expected = 'SoftwareUpdatesOnly' }, @{ Type = 5; Expected = 'TaskSequencesOnly' }, @{ Type = 1; Expected = 'Any' }) {
        $global:Fx = [PSCustomObject]@{ Name = 'N'; ServiceWindowSchedules = 'x'; ServiceWindowType = $Type; IsEnabled = $false; IsGMT = $false }
        $null = Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'N' -TargetCollectionIds @('A')
        $global:ExpectedApplyTo = $Expected
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 1 -Exactly -ParameterFilter { $ApplyTo -ceq $global:ExpectedApplyTo -and $IsEnabled -eq $false }
    }

    It 'returns nothing and logs an error when the source window does not exist' {
        $global:Fx = $null
        @(Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'Nope' -TargetCollectionIds @('A')).Count | Should -Be 0
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }

    It 'continues after a refused target and reports it (partial failure)' {
        Mock New-CMMaintenanceWindow { if ($CollectionId -eq 'B') { throw 'refused' } } -ModuleName MaintWindowMgrCommon
        $r = @(Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'Night' -TargetCollectionIds @('A', 'B', 'C'))
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,False,True'
        $r[1].Error | Should -BeLike '*refused*'
    }

    It 'lets a source lookup failure reach the caller' {
        Mock Get-CMMaintenanceWindow { throw 'provider down' } -ModuleName MaintWindowMgrCommon
        { Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'Night' -TargetCollectionIds @('A') } | Should -Throw '*provider down*'
    }

    It 'copies nothing under -WhatIf' {
        @(Copy-MaintenanceWindowToCollections -SourceCollectionId 'S' -SourceWindowName 'Night' -TargetCollectionIds @('A', 'B') -WhatIf).Count | Should -Be 0
        Should -Invoke New-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }
}

Describe 'Set-BulkWindowEnabled' {
    BeforeEach { Mock Set-CMMaintenanceWindow { } -ModuleName MaintWindowMgrCommon }

    It 'sets the state of every window and reports each' {
        $recs = @([PSCustomObject]@{ CollectionID = 'A'; WindowName = 'W1' }, [PSCustomObject]@{ CollectionID = 'B'; WindowName = 'W2' })
        $r = @(Set-BulkWindowEnabled -WindowRecords $recs -Enabled $false)
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,True'
        Should -Invoke Set-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 2 -Exactly -ParameterFilter { $IsEnabled -eq $false }
    }

    It 'continues after a failure (partial failure)' {
        Mock Set-CMMaintenanceWindow { if ($CollectionId -eq 'B') { throw 'locked' } } -ModuleName MaintWindowMgrCommon
        $recs = @('A', 'B', 'C') | ForEach-Object { [PSCustomObject]@{ CollectionID = $_; WindowName = "W$_" } }
        $r = @(Set-BulkWindowEnabled -WindowRecords $recs -Enabled $true)
        ($r | ForEach-Object Success) -join ',' | Should -Be 'True,False,True'
        $r[1].Error | Should -BeLike '*locked*'
    }

    It 'accepts an empty list' {
        @(Set-BulkWindowEnabled -WindowRecords @() -Enabled $true).Count | Should -Be 0
    }

    It 'changes nothing under -WhatIf' {
        @(Set-BulkWindowEnabled -WindowRecords @([PSCustomObject]@{ CollectionID = 'A'; WindowName = 'W' }) -Enabled $true -WhatIf).Count | Should -Be 0
        Should -Invoke Set-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
    }
}
