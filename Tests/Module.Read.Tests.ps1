#Requires -Modules Pester
# Read side: folders, windows, summaries, human schedule text and next occurrences.
# ConfigMgr cmdlets are stubs and mocks. Fixture values reach the mock bodies through
# $global:Fx, because a mock body runs in the module scope.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name Get-CMDeviceCollection, Get-CMMaintenanceWindow
    Mock Write-Log {} -ModuleName MaintWindowMgrCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Folder tree and map' {
    It 'maps the container node columns to integers and a name' {
        Mock Get-CimInstance { [PSCustomObject]@{ ContainerNodeID = '7'; Name = 'Servers'; ParentContainerNodeID = '0' } } -ModuleName MaintWindowMgrCommon
        $r = @(Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM')
        $r[0].FolderID | Should -Be 7
        $r[0].Name | Should -Be 'Servers'
        $r[0].ParentID | Should -Be 0
        Should -Invoke Get-CimInstance -ModuleName MaintWindowMgrCommon -ParameterFilter { $Namespace -eq 'root\sms\site_MCM' -and $ClassName -eq 'SMS_ObjectContainerNode' -and $Filter -eq 'ObjectType = 5000' }
    }

    It 'returns no folders for a site without folders' {
        Mock Get-CimInstance { } -ModuleName MaintWindowMgrCommon
        $r = Get-CMCollectionFolderTree -SMSProvider 'cm01' -SiteCode 'MCM'
        @($r).Count | Should -Be 0
    }

    It 'maps each collection id to its folder' {
        Mock Get-CimInstance { [PSCustomObject]@{ InstanceKey = 'MCM00001'; ContainerNodeID = 3 }; [PSCustomObject]@{ InstanceKey = 'MCM00002'; ContainerNodeID = '4' } } -ModuleName MaintWindowMgrCommon
        $m = Get-CMCollectionFolderMap -SMSProvider 'cm01' -SiteCode 'MCM'
        $m['MCM00001'] | Should -Be 3
        $m['MCM00002'] | Should -Be 4
    }

    It '<Fn> lets a provider failure reach the caller' -ForEach @(@{ Fn = 'Get-CMCollectionFolderTree' }, @{ Fn = 'Get-CMCollectionFolderMap' }) {
        Mock Get-CimInstance { throw 'RPC server unavailable' } -ModuleName MaintWindowMgrCommon
        { & $Fn -SMSProvider 'cm01' -SiteCode 'MCM' } | Should -Throw '*RPC server unavailable*'
    }
}

Describe 'Get-AllMaintenanceWindows' {
    BeforeEach {
        $global:Fx = @{
            Collections = @([PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'Servers' })
            Windows     = @{ MCM00001 = @([PSCustomObject]@{ Name = 'Sunday night'; ServiceWindowID = 'W1'; ServiceWindowType = 4; RecurrenceType = 3; StartTime = [datetime]'2026-01-04T02:00:00'; Duration = 270; IsEnabled = $true; IsGMT = $false; Description = 'd' }) }
        }
        Mock Get-CMDeviceCollection { $global:Fx.Collections } -ModuleName MaintWindowMgrCommon
        Mock Get-CMMaintenanceWindow { $global:Fx.Windows[$CollectionId] } -ModuleName MaintWindowMgrCommon
    }

    It 'flattens a window with the collection columns' {
        $r = @(Get-AllMaintenanceWindows)
        $r.Count | Should -Be 1
        $r[0].CollectionName | Should -Be 'Servers'
        $r[0].CollectionID | Should -Be 'MCM00001'
        $r[0].WindowName | Should -Be 'Sunday night'
        $r[0].WindowID | Should -Be 'W1'
        $r[0].Type | Should -Be 'Software Updates'
        $r[0].Recurrence | Should -Be 'Weekly'
        $r[0].Duration | Should -Be '4h 30m'
        $r[0].DurationMinutes | Should -Be 270
        $r[0].IsEnabled | Should -BeTrue
        $r[0].IsUTC | Should -BeFalse
        $r[0].Schedule | Should -Match '^Every Sunday at 2:00 AM \(4h 30m\)$'
    }

    It 'names window type <Type> as <Name>' -ForEach @(
        @{ Type = 1; Name = 'General' }, @{ Type = 4; Name = 'Software Updates' }, @{ Type = 5; Name = 'Task Sequences' }, @{ Type = 9; Name = 'Unknown (9)' }, @{ Type = 0; Name = 'Unknown (0)' }
    ) {
        $global:Fx.Windows.MCM00001[0].ServiceWindowType = $Type
        (@(Get-AllMaintenanceWindows))[0].Type | Should -Be $Name
    }

    It 'names recurrence <Type> as <Name>' -ForEach @(
        @{ Type = 1; Name = 'None (One-time)' }, @{ Type = 2; Name = 'Daily' }, @{ Type = 3; Name = 'Weekly' }, @{ Type = 4; Name = 'Monthly by Weekday' },
        @{ Type = 5; Name = 'Monthly by Date' }, @{ Type = 7; Name = 'Unknown (7)' }
    ) {
        $global:Fx.Windows.MCM00001[0].RecurrenceType = $Type
        (@(Get-AllMaintenanceWindows))[0].Recurrence | Should -Be $Name
    }

    It 'shows N/A as the next occurrence of a one-time window in the past' {
        $w = $global:Fx.Windows.MCM00001[0]; $w.RecurrenceType = 1; $w.StartTime = [datetime]'2001-01-01T02:00:00'
        (@(Get-AllMaintenanceWindows))[0].NextOccurrence | Should -Be 'N/A'
    }

    It 'shows a date as the next occurrence of a daily window' {
        $global:Fx.Windows.MCM00001[0].RecurrenceType = 2
        (@(Get-AllMaintenanceWindows))[0].NextOccurrence | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$'
    }

    It 'skips a collection without windows' {
        $global:Fx.Collections = @([PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'A' }, [PSCustomObject]@{ CollectionID = 'MCM00002'; Name = 'B' })
        $global:Fx.Windows.MCM00002 = $null
        @(Get-AllMaintenanceWindows).Count | Should -Be 1
    }

    It 'returns nothing for a site without collections' {
        $global:Fx.Collections = @()
        @(Get-AllMaintenanceWindows).Count | Should -Be 0
    }

    It 'lists every window of a collection that has several' {
        $w = $global:Fx.Windows.MCM00001[0]
        $global:Fx.Windows.MCM00001 = @($w, $w, $w)
        @(Get-AllMaintenanceWindows).Count | Should -Be 3
    }

    It 'reads 500 collections' {
        $global:Fx.Collections = @(1..500 | ForEach-Object { [PSCustomObject]@{ CollectionID = "MCM$_"; Name = "C$_" } })
        @(Get-AllMaintenanceWindows).Count | Should -Be 0
        Should -Invoke Get-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -Times 500 -Exactly
    }

    It 'lets a collection query failure reach the caller' {
        Mock Get-CMDeviceCollection { throw 'not connected' } -ModuleName MaintWindowMgrCommon
        { Get-AllMaintenanceWindows } | Should -Throw '*not connected*'
    }

    It 'reports a window with a null duration as 0h 0m' {
        $global:Fx.Windows.MCM00001[0].Duration = $null
        (@(Get-AllMaintenanceWindows))[0].Duration | Should -Be '0h 0m'
    }
}

Describe 'Get-CollectionMaintenanceWindows' {
    It 'returns the windows of the collection' {
        Mock Get-CMMaintenanceWindow { [PSCustomObject]@{ Name = 'W' } } -ModuleName MaintWindowMgrCommon
        (@(Get-CollectionMaintenanceWindows -CollectionId 'MCM00001'))[0].Name | Should -Be 'W'
        Should -Invoke Get-CMMaintenanceWindow -ModuleName MaintWindowMgrCommon -ParameterFilter { $CollectionId -eq 'MCM00001' }
    }

    It 'returns null and logs an error when the query fails' {
        Mock Get-CMMaintenanceWindow { throw 'denied' } -ModuleName MaintWindowMgrCommon
        Get-CollectionMaintenanceWindows -CollectionId 'MCM00001' | Should -BeNullOrEmpty
        Should -Invoke Write-Log -ModuleName MaintWindowMgrCommon -ParameterFilter { $Level -eq 'ERROR' } -Times 1 -Exactly
    }
}

Describe 'Get-DeviceCollectionSummary' {
    It 'returns id, name, member count and the built-in flag' {
        Mock Get-CMDeviceCollection {
            [PSCustomObject]@{ CollectionID = 'SMS00001'; Name = 'All Systems'; MemberCount = 9 }
            [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'Mine'; MemberCount = $null }
        } -ModuleName MaintWindowMgrCommon
        $r = @(Get-DeviceCollectionSummary)
        $r[0].IsBuiltIn | Should -BeTrue
        $r[1].IsBuiltIn | Should -BeFalse
        $r[1].MemberCount | Should -Be 0
    }

    It 'returns nothing when there are no collections' {
        Mock Get-CMDeviceCollection { } -ModuleName MaintWindowMgrCommon
        @(Get-DeviceCollectionSummary).Count | Should -Be 0
    }
}

Describe 'Get-CollectionsWithoutWindows' {
    BeforeAll {
        $script:colls = @([PSCustomObject]@{ CollectionID = 'A' }, [PSCustomObject]@{ CollectionID = 'B' }, [PSCustomObject]@{ CollectionID = 'C' })
    }

    It 'returns the collections that have no window' {
        $r = @(Get-CollectionsWithoutWindows -AllCollections $script:colls -AllWindows @([PSCustomObject]@{ CollectionID = 'B' }, [PSCustomObject]@{ CollectionID = 'B' }))
        ($r.CollectionID -join ',') | Should -Be 'A,C'
    }

    It 'returns every collection when there is no window' {
        @(Get-CollectionsWithoutWindows -AllCollections $script:colls -AllWindows @()).Count | Should -Be 3
    }

    It 'returns nothing when there is no collection' {
        @(Get-CollectionsWithoutWindows -AllCollections @() -AllWindows @([PSCustomObject]@{ CollectionID = 'A' })).Count | Should -Be 0
    }

    It 'ignores a null window and a window without a collection id' {
        @(Get-CollectionsWithoutWindows -AllCollections $script:colls -AllWindows @($null, [PSCustomObject]@{ CollectionID = '' })).Count | Should -Be 3
    }

    It 'returns nothing when every collection has a window' {
        @(Get-CollectionsWithoutWindows -AllCollections $script:colls -AllWindows @($script:colls)).Count | Should -Be 0
    }
}

Describe 'ConvertTo-HumanSchedule' {
    It 'writes <Case>' -ForEach @(
        @{ Case = 'a one-time window';      Rec = 1; Start = '2026-03-04T14:30:00'; Dur = 90;  Expected = 'One-time: 2026-03-04 2:30 PM (1h 30m)' }
        @{ Case = 'a daily window';         Rec = 2; Start = '2026-03-04T02:00:00'; Dur = 240; Expected = 'Daily at 2:00 AM (4h)' }
        @{ Case = 'a weekly window';        Rec = 3; Start = '2026-03-04T23:15:00'; Dur = 45;  Expected = 'Every Wednesday at 11:15 PM (0h 45m)' }
        @{ Case = 'a monthly weekday one';  Rec = 4; Start = '2026-03-10T03:00:00'; Dur = 180; Expected = 'Monthly (Tuesday) at 3:00 AM (3h)' }
        @{ Case = 'a monthly date one';     Rec = 5; Start = '2026-03-15T01:05:00'; Dur = 61;  Expected = 'Monthly (day 15) at 1:05 AM (1h 1m)' }
        @{ Case = 'an unknown recurrence';  Rec = 9; Start = '2026-03-15T12:00:00'; Dur = 60;  Expected = 'Schedule at 12:00 PM (1h)' }
        @{ Case = 'midnight';               Rec = 2; Start = '2026-03-15T00:00:00'; Dur = 60;  Expected = 'Daily at 12:00 AM (1h)' }
    ) {
        ConvertTo-HumanSchedule -Window ([PSCustomObject]@{ StartTime = [datetime]$Start; Duration = $Dur; RecurrenceType = $Rec }) | Should -Be $Expected
    }

    It 'shows ? for a window without a start time' -ForEach @(
        @{ Rec = 1; Expected = 'One-time: ? (2h)' }, @{ Rec = 2; Expected = 'Daily at ? (2h)' }, @{ Rec = 3; Expected = 'Every ? at ? (2h)' },
        @{ Rec = 4; Expected = 'Monthly (?) at ? (2h)' }, @{ Rec = 5; Expected = 'Monthly (day ?) at ? (2h)' }
    ) {
        ConvertTo-HumanSchedule -Window ([PSCustomObject]@{ StartTime = $null; Duration = 120; RecurrenceType = $Rec }) | Should -Be $Expected
    }
}

Describe 'Get-NextOccurrences' {
    It 'returns nothing for a window without a start time' {
        @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = $null; RecurrenceType = 2 })).Count | Should -Be 0
    }

    It 'returns nothing for an unknown recurrence' {
        @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = (Get-Date).AddDays(1); RecurrenceType = 99 })).Count | Should -Be 0
    }

    It 'returns a future one-time window and not a past one' {
        @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = (Get-Date).AddDays(2); RecurrenceType = 1 })).Count | Should -Be 1
        @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = (Get-Date).AddDays(-2); RecurrenceType = 1 })).Count | Should -Be 0
    }

    It 'recurrence <Rec> gives <Count> future dates, in order' -ForEach @(
        @{ Rec = 2; Count = 1 }, @{ Rec = 2; Count = 5 }, @{ Rec = 3; Count = 5 }, @{ Rec = 4; Count = 5 }, @{ Rec = 5; Count = 5 }, @{ Rec = 5; Count = 13 }
    ) {
        $now = Get-Date
        $r = @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = $now.AddDays(-100); RecurrenceType = $Rec }) -Count $Count)
        $r.Count | Should -Be $Count
        $r[0] | Should -BeGreaterThan $now
        for ($i = 1; $i -lt $r.Count; $i++) { $r[$i] | Should -BeGreaterThan $r[$i - 1] }
    }

    It 'keeps the time of day for daily and weekly windows' -ForEach @(@{ Rec = 2 }, @{ Rec = 3 }) {
        $start = (Get-Date).Date.AddDays(-30).AddHours(3).AddMinutes(17)
        $r = @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = $start; RecurrenceType = $Rec }) -Count 3)
        foreach ($d in $r) { $d.TimeOfDay | Should -Be $start.TimeOfDay }
    }

    It 'steps a weekly window by seven days and keeps the weekday' {
        $start = (Get-Date).AddDays(-20)
        $r = @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = $start; RecurrenceType = 3 }) -Count 4)
        foreach ($d in $r) { $d.DayOfWeek | Should -Be $start.DayOfWeek }
        ($r[1] - $r[0]).TotalDays | Should -Be 7
    }

    It 'keeps the weekday of a monthly weekday window' {
        $start = (Get-Date).AddMonths(-3)
        $r = @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = $start; RecurrenceType = 4 }) -Count 4)
        foreach ($d in $r) { $d.DayOfWeek | Should -Be $start.DayOfWeek }
    }

    It 'returns no date when Count is 0' {
        @(Get-NextOccurrences -Window ([PSCustomObject]@{ StartTime = (Get-Date).AddDays(-3); RecurrenceType = 2 }) -Count 0).Count | Should -Be 0
    }
}
