#Requires -Modules Pester
# Schedule helpers: Get-PatchTuesday, New-WindowSchedule, New-PatchTuesdaySchedule.
# New-CMSchedule is a stub; the mock keeps the parameters it received.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:Stubs = New-CmStub -Name New-CMSchedule
    Mock Write-Log {} -ModuleName MaintWindowMgrCommon
}

AfterAll {
    Remove-CmStub -Name $script:Stubs
    Remove-Variable -Name Fx -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Get-PatchTuesday' {
    It 'gives the second Tuesday of <Year>-<Month>: <Day>' -ForEach @(
        @{ Year = 2026; Month = 1;  Day = 13 }
        @{ Year = 2026; Month = 2;  Day = 10 }
        @{ Year = 2026; Month = 3;  Day = 10 }
        @{ Year = 2026; Month = 11; Day = 10 }
        @{ Year = 2026; Month = 12; Day = 8 }
        @{ Year = 2027; Month = 2;  Day = 9 }
        @{ Year = 2025; Month = 12; Day = 9 }
        @{ Year = 2024; Month = 2;  Day = 13 }
        @{ Year = 2028; Month = 2;  Day = 8 }
    ) {
        $d = Get-PatchTuesday -Year $Year -Month $Month
        $d.Day | Should -Be $Day
        $d.Month | Should -Be $Month
        $d.Year | Should -Be $Year
        $d.DayOfWeek | Should -Be 'Tuesday'
    }

    It 'gives a Tuesday between the 8th and the 14th for every month from 2020 to 2040' {
        $bad = foreach ($y in 2020..2040) { foreach ($m in 1..12) { $d = Get-PatchTuesday -Year $y -Month $m; if ($d.DayOfWeek -ne 'Tuesday' -or $d.Day -lt 8 -or $d.Day -gt 14 -or $d.Month -ne $m) { "$y-$m -> $d" } } }
        $bad | Should -BeNullOrEmpty
    }

    It 'uses the current month when no month is given' {
        $now = Get-Date
        $d = Get-PatchTuesday
        $d.Month | Should -Be $now.Month
        $d.Year | Should -Be $now.Year
    }

    It 'sets the time to midnight' {
        (Get-PatchTuesday -Year 2026 -Month 5).TimeOfDay.TotalSeconds | Should -BeLessThan 1
    }
}

Describe 'New-WindowSchedule' {
    BeforeEach {
        $global:Fx = $null
        Mock New-CMSchedule { $global:Fx = $PesterBoundParameters; $global:Sched = [PSCustomObject]@{ HourDuration = 0; MinuteDuration = 0 }; $global:Sched } -ModuleName MaintWindowMgrCommon
        $script:start = [datetime]'2026-11-10T02:00:00'
    }

    It 'passes <Type> to the cmdlet as <Expect>' -ForEach @(
        @{ Type = 'OneTime';          Extra = @{};                                              Expect = 'Nonrecurring' }
        @{ Type = 'Daily';            Extra = @{ RecurCount = 3 };                              Expect = 'RecurInterval Days, RecurCount 3' }
        @{ Type = 'Weekly';           Extra = @{ DayOfWeek = 'Friday'; RecurCount = 2 };        Expect = 'DayOfWeek Friday, RecurCount 2' }
        @{ Type = 'MonthlyByDate';    Extra = @{ DayOfMonth = 15 };                             Expect = 'DayOfMonth 15' }
        @{ Type = 'MonthlyByWeekday'; Extra = @{ DayOfWeek = 'Tuesday'; WeekOrder = 'Second' }; Expect = 'DayOfWeek Tuesday, WeekOrder Second' }
    ) {
        $r = New-WindowSchedule -RecurrenceType $Type -StartTime $script:start @Extra
        $r | Should -Be $global:Sched
        $p = $global:Fx
        $p.Start | Should -Be $script:start
        switch ($Type) {
            'OneTime' { $p.Nonrecurring | Should -BeTrue }
            'Daily' { $p.RecurInterval | Should -Be 'Days'; $p.RecurCount | Should -Be 3 }
            'Weekly' { $p.DayOfWeek | Should -Be 'Friday'; $p.RecurCount | Should -Be 2 }
            'MonthlyByDate' { $p.DayOfMonth | Should -Be 15 }
            'MonthlyByWeekday' { $p.DayOfWeek | Should -Be 'Tuesday'; $p.WeekOrder | Should -Be 'Second' }
        }
    }

    It 'turns a duration of <H>h <M>m into <Interval> <Count> and <Extra> extra minutes' -ForEach @(
        @{ H = 4;  M = 0;  Interval = 'Hours';   Count = 4;  Extra = 0 }
        @{ H = 23; M = 0;  Interval = 'Hours';   Count = 23; Extra = 0 }
        @{ H = 0;  M = 45; Interval = 'Minutes'; Count = 45; Extra = 0 }
        @{ H = 1;  M = 30; Interval = 'Hours';   Count = 1;  Extra = 30 }
        @{ H = 23; M = 59; Interval = 'Hours';   Count = 23; Extra = 59 }
        @{ H = 0;  M = 1;  Interval = 'Minutes'; Count = 1;  Extra = 0 }
        @{ H = 1;  M = 0;  Interval = 'Hours';   Count = 1;  Extra = 0 }
        @{ H = 0;  M = 90; Interval = 'Hours';   Count = 1;  Extra = 30 }
        @{ H = 2;  M = 75; Interval = 'Hours';   Count = 3;  Extra = 15 }
    ) {
        $r = New-WindowSchedule -RecurrenceType Daily -StartTime $script:start -DurationHours $H -DurationMinutes $M
        $global:Fx.DurationInterval | Should -Be $Interval
        $global:Fx.DurationCount | Should -Be $Count
        $r.MinuteDuration | Should -Be $Extra
    }
    It 'returns null and logs an error for a duration of <H>h <M>m' -ForEach @(
        @{ H = 0;  M = 0 }, @{ H = 24; M = 0 }, @{ H = 23; M = 60 }, @{ H = 0; M = 1440 }, @{ H = -1; M = 0 }, @{ H = 5; M = -10 }, @{ H = 100; M = 0 }
    ) {
        New-WindowSchedule -RecurrenceType Daily -StartTime $script:start -DurationHours $H -DurationMinutes $M | Should -BeNullOrEmpty
        Should -Invoke New-CMSchedule -ModuleName MaintWindowMgrCommon -Times 0 -Exactly
        Should -Invoke Write-Log -ModuleName MaintWindowMgrCommon -ParameterFilter { $Level -eq 'ERROR' } -Times 1 -Exactly
    }

    It 'returns null and logs an error when the cmdlet fails' {
        Mock New-CMSchedule { throw 'bad schedule' } -ModuleName MaintWindowMgrCommon
        New-WindowSchedule -RecurrenceType Daily -StartTime $script:start | Should -BeNullOrEmpty
        Should -Invoke Write-Log -ModuleName MaintWindowMgrCommon -ParameterFilter { $Level -eq 'ERROR' -and $Message -like '*bad schedule*' }
    }

    It 'rejects an unknown recurrence type' {
        { New-WindowSchedule -RecurrenceType 'Hourly' -StartTime $script:start } | Should -Throw
    }
}

Describe 'New-PatchTuesdaySchedule' {
    BeforeEach {
        $global:Fx = $null
        Mock New-CMSchedule { $global:Fx = $PesterBoundParameters; [PSCustomObject]@{ HourDuration = 0; MinuteDuration = 0 } } -ModuleName MaintWindowMgrCommon
    }

    It 'schedules the second Tuesday for offset 0' {
        New-PatchTuesdaySchedule -OffsetDays 0 -StartHour 3 -StartMinute 30 -DurationHours 5 | Should -Not -BeNullOrEmpty
        $global:Fx.DayOfWeek | Should -Be 'Tuesday'
        $global:Fx.WeekOrder | Should -Be 'Second'
        $global:Fx.Start.Hour | Should -Be 3
        $global:Fx.Start.Minute | Should -Be 30
        $global:Fx.DurationCount | Should -Be 5
    }

    It 'schedules a day of the month for offset <Offset>' -ForEach @(@{ Offset = 1 }, @{ Offset = 3 }, @{ Offset = 7 }, @{ Offset = 20 }) {
        New-PatchTuesdaySchedule -OffsetDays $Offset -StartHour 2 | Should -Not -BeNullOrEmpty
        $global:Fx.DayOfMonth | Should -BeIn (1..31)
        $global:Fx.Start | Should -BeGreaterThan (Get-Date).AddDays(-1)
        $global:Fx.Start.Day | Should -Be $global:Fx.DayOfMonth
        $global:Fx.Start.Hour | Should -Be 2
    }

    It 'returns null when the duration is not valid' {
        New-PatchTuesdaySchedule -OffsetDays 0 -DurationHours 0 -DurationMinutes 0 | Should -BeNullOrEmpty
    }
}
