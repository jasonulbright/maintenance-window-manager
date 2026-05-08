#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x baseline for the MaintWindowMgrCommon shared module.

.DESCRIPTION
    Covers pure-logic exports: logging, schedule helpers (Patch Tuesday math),
    coverage calculations, schedule formatting, occurrence prediction, and
    template I/O. CM-cmdlet integration (Connect-CMSite, Get-AllMaintenanceWindows,
    New-ManagedMaintenanceWindow, etc.) requires a live MECM site and is
    verified end-to-end on a CM-console-equipped client (CLIENT01) rather
    than mocked here.

.EXAMPLE
    Invoke-Pester .\MaintWindowMgrCommon.Tests.ps1
#>

BeforeAll {
    Import-Module "$PSScriptRoot\MaintWindowMgrCommon.psd1" -Force -DisableNameChecking
}

# ============================================================================
# Write-Log / Initialize-Logging
# ============================================================================

Describe 'Write-Log' {
    It 'writes formatted message to log file' {
        $logFile = Join-Path $TestDrive 'test.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Hello world' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] Hello world'
    }

    It 'tags WARN messages correctly' {
        $logFile = Join-Path $TestDrive 'warn.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Something odd' -Level WARN -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[WARN \] Something odd'
    }

    It 'tags ERROR messages correctly' {
        $logFile = Join-Path $TestDrive 'error.log'
        Initialize-Logging -LogPath $logFile
        Write-Log 'Failure' -Level ERROR -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[ERROR\] Failure'
    }

    It 'accepts empty string message' {
        $logFile = Join-Path $TestDrive 'empty.log'
        Initialize-Logging -LogPath $logFile
        { Write-Log '' -Quiet } | Should -Not -Throw
    }
}

Describe 'Initialize-Logging' {
    It 'creates log file with header line' {
        $logFile = Join-Path $TestDrive 'init.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match '\[INFO \] === Log initialized ==='
    }

    It 'creates parent directories if missing' {
        $logFile = Join-Path $TestDrive 'sub\dir\deep.log'
        Initialize-Logging -LogPath $logFile
        Test-Path -LiteralPath $logFile | Should -BeTrue
    }

    It '-Attach preserves an externally-created log file' {
        $logFile = Join-Path $TestDrive 'attach.log'
        $sentinel = "[2026-05-02 00:00:00] [INFO ] Shell-managed header"
        Set-Content -LiteralPath $logFile -Value $sentinel -Encoding UTF8
        Initialize-Logging -LogPath $logFile -Attach
        Write-Log 'Module appended line' -Quiet
        $content = Get-Content -LiteralPath $logFile -Raw
        $content | Should -Match 'Shell-managed header'
        $content | Should -Match 'Module appended line'
    }
}

# ============================================================================
# Get-PatchTuesday (pure date math)
# ============================================================================

Describe 'Get-PatchTuesday' {
    It 'returns the second Tuesday of May 2026 (2026-05-12)' {
        $pt = Get-PatchTuesday -Year 2026 -Month 5
        $pt.Year   | Should -Be 2026
        $pt.Month  | Should -Be 5
        $pt.Day    | Should -Be 12
        $pt.DayOfWeek | Should -Be ([System.DayOfWeek]::Tuesday)
    }

    It 'returns the second Tuesday of January 2026 (2026-01-13)' {
        $pt = Get-PatchTuesday -Year 2026 -Month 1
        $pt.Day    | Should -Be 13
        $pt.DayOfWeek | Should -Be ([System.DayOfWeek]::Tuesday)
    }

    It 'returns the second Tuesday of November 2026 (2026-11-10)' {
        $pt = Get-PatchTuesday -Year 2026 -Month 11
        $pt.Day    | Should -Be 10
        $pt.DayOfWeek | Should -Be ([System.DayOfWeek]::Tuesday)
    }

    It 'always returns a Tuesday for every month of a year' {
        for ($m = 1; $m -le 12; $m++) {
            $pt = Get-PatchTuesday -Year 2026 -Month $m
            $pt.DayOfWeek | Should -Be ([System.DayOfWeek]::Tuesday)
            # Second Tuesday is always between day 8 and day 14 inclusive
            $pt.Day | Should -BeGreaterOrEqual 8
            $pt.Day | Should -BeLessOrEqual 14
        }
    }

    It 'defaults Year/Month to current month' {
        $now = Get-Date
        $pt = Get-PatchTuesday
        $pt.Year  | Should -Be $now.Year
        $pt.Month | Should -Be $now.Month
    }
}

# ============================================================================
# Get-CollectionsWithoutWindows (pure set difference)
# ============================================================================

Describe 'Get-CollectionsWithoutWindows' {
    BeforeAll {
        $script:AllColls = @(
            [PSCustomObject]@{ CollectionID = 'MCM00001'; Name = 'Coll A' }
            [PSCustomObject]@{ CollectionID = 'MCM00002'; Name = 'Coll B' }
            [PSCustomObject]@{ CollectionID = 'MCM00003'; Name = 'Coll C' }
        )
    }

    It 'returns collections not present in the windows array' {
        $windows = @(
            [PSCustomObject]@{ CollectionID = 'MCM00001'; WindowName = 'w1' }
        )
        $without = @(Get-CollectionsWithoutWindows -AllCollections $script:AllColls -AllWindows $windows)
        $without.Count | Should -Be 2
        $without.CollectionID | Should -Contain 'MCM00002'
        $without.CollectionID | Should -Contain 'MCM00003'
    }

    It 'returns empty when every collection has at least one window' {
        $windows = @(
            [PSCustomObject]@{ CollectionID = 'MCM00001' }
            [PSCustomObject]@{ CollectionID = 'MCM00002' }
            [PSCustomObject]@{ CollectionID = 'MCM00003' }
        )
        $without = @(Get-CollectionsWithoutWindows -AllCollections $script:AllColls -AllWindows $windows)
        $without.Count | Should -Be 0
    }

    It 'returns every collection when no windows exist' {
        $without = @(Get-CollectionsWithoutWindows -AllCollections $script:AllColls -AllWindows @())
        $without.Count | Should -Be 3
    }
}

# ============================================================================
# ConvertTo-HumanSchedule (string formatting from window object)
# ============================================================================

Describe 'ConvertTo-HumanSchedule' {
    It 'formats one-time recurrence as One-time: <date>' {
        $w = [PSCustomObject]@{
            RecurrenceType = 1
            StartTime = (Get-Date '2026-05-12 02:00')
            Duration  = 240
        }
        $s = ConvertTo-HumanSchedule -Window $w
        $s | Should -Match 'One-time'
        $s | Should -Match '2026-05-12'
        $s | Should -Match '4h'
    }

    It 'formats daily recurrence as Daily at <time>' {
        $w = [PSCustomObject]@{
            RecurrenceType = 2
            StartTime = (Get-Date '2026-05-12 23:00')
            Duration  = 360
        }
        $s = ConvertTo-HumanSchedule -Window $w
        $s | Should -Match 'Daily at'
        $s | Should -Match '11:00 PM'
        $s | Should -Match '6h'
    }

    It 'formats weekly recurrence with the day name' {
        $w = [PSCustomObject]@{
            RecurrenceType = 3
            StartTime = (Get-Date '2026-05-10 02:00')   # Sunday
            Duration  = 240
        }
        $s = ConvertTo-HumanSchedule -Window $w
        $s | Should -Match 'Every Sunday'
    }

    It 'formats monthly-by-date with the day number' {
        $w = [PSCustomObject]@{
            RecurrenceType = 5
            StartTime = (Get-Date '2026-05-15 02:00')
            Duration  = 240
        }
        $s = ConvertTo-HumanSchedule -Window $w
        $s | Should -Match 'Monthly \(day 15\)'
    }

    It 'shows minutes when non-zero' {
        $w = [PSCustomObject]@{
            RecurrenceType = 2
            StartTime = (Get-Date '2026-05-12 03:00')
            Duration  = 90
        }
        $s = ConvertTo-HumanSchedule -Window $w
        $s | Should -Match '1h 30m'
    }
}

# ============================================================================
# Get-NextOccurrences (date-math projection)
# ============================================================================

Describe 'Get-NextOccurrences' {
    It 'returns 1 occurrence for a future one-time window' {
        $w = [PSCustomObject]@{
            RecurrenceType = 1
            StartTime = (Get-Date).AddDays(7)
        }
        $occ = @(Get-NextOccurrences -Window $w -Count 5)
        $occ.Count | Should -Be 1
    }

    It 'returns 0 occurrences for a past one-time window' {
        $w = [PSCustomObject]@{
            RecurrenceType = 1
            StartTime = (Get-Date).AddDays(-7)
        }
        $occ = @(Get-NextOccurrences -Window $w -Count 5)
        $occ.Count | Should -Be 0
    }

    It 'projects N daily occurrences spaced 1 day apart' {
        $w = [PSCustomObject]@{
            RecurrenceType = 2
            StartTime = (Get-Date '2026-05-01 02:00')
        }
        $occ = @(Get-NextOccurrences -Window $w -Count 3)
        $occ.Count | Should -Be 3
        ($occ[1] - $occ[0]).TotalDays | Should -Be 1
        ($occ[2] - $occ[1]).TotalDays | Should -Be 1
    }

    It 'projects N weekly occurrences spaced 7 days apart' {
        $w = [PSCustomObject]@{
            RecurrenceType = 3
            StartTime = (Get-Date '2026-05-01 02:00')
        }
        $occ = @(Get-NextOccurrences -Window $w -Count 4)
        $occ.Count | Should -Be 4
        ($occ[1] - $occ[0]).TotalDays | Should -Be 7
        ($occ[3] - $occ[2]).TotalDays | Should -Be 7
    }

    It 'returns occurrences that are all in the future' {
        $now = Get-Date
        $w = [PSCustomObject]@{
            RecurrenceType = 2
            StartTime = $now.AddDays(-365)
        }
        $occ = @(Get-NextOccurrences -Window $w -Count 5)
        foreach ($o in $occ) {
            $o | Should -BeGreaterThan $now
        }
    }
}

# ============================================================================
# Get-WindowTemplates / Save-WindowTemplate / Remove-WindowTemplate
# ============================================================================

Describe 'Get-WindowTemplates' {
    It 'loads bundled templates from the Templates folder' {
        $tplPath = Join-Path $PSScriptRoot '..\Templates'
        $templates = @(Get-WindowTemplates -TemplatesPath $tplPath)
        $templates.Count | Should -BeGreaterThan 0
    }

    It 'each loaded template has the required properties' {
        $tplPath = Join-Path $PSScriptRoot '..\Templates'
        $templates = @(Get-WindowTemplates -TemplatesPath $tplPath)
        foreach ($t in $templates) {
            $t.Name           | Should -Not -BeNullOrEmpty
            $t.RecurrenceType | Should -Not -BeNullOrEmpty
            $t.PSObject.Properties['DurationHours']     | Should -Not -BeNullOrEmpty
            $t.PSObject.Properties['StartHour']         | Should -Not -BeNullOrEmpty
            $t.PSObject.Properties['PatchTuesdayOffset'] | Should -Not -BeNullOrEmpty
        }
    }

    It 'returns empty array for missing folder' {
        $bogus = Join-Path $TestDrive 'no-such-templates-folder'
        $templates = @(Get-WindowTemplates -TemplatesPath $bogus)
        $templates.Count | Should -Be 0
    }

    It 'attaches FileName + FilePath to each template' {
        $tplPath = Join-Path $PSScriptRoot '..\Templates'
        $templates = @(Get-WindowTemplates -TemplatesPath $tplPath)
        $templates[0].FileName | Should -Not -BeNullOrEmpty
        $templates[0].FilePath | Should -Match '\.json$'
    }
}

Describe 'Save-WindowTemplate' {
    It 'writes a JSON file under the templates folder' {
        $tplPath = Join-Path $TestDrive 'tpl'
        $path = Save-WindowTemplate -TemplatesPath $tplPath -Name 'My Window' `
            -RecurrenceType 'Daily' -DurationHours 4 -StartHour 23
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'slugifies the file name from the template name' {
        $tplPath = Join-Path $TestDrive 'slug'
        $path = Save-WindowTemplate -TemplatesPath $tplPath -Name 'Patch Tuesday + 7' `
            -RecurrenceType 'MonthlyByDate' -DurationHours 4 -StartHour 2
        $path | Should -Match 'patch-tuesday---7\.json$'
    }

    It 'creates the templates folder if missing' {
        $tplPath = Join-Path $TestDrive 'created\sub\folder'
        $null = Save-WindowTemplate -TemplatesPath $tplPath -Name 'Created' `
            -RecurrenceType 'Daily' -DurationHours 1 -StartHour 0
        Test-Path -LiteralPath $tplPath | Should -BeTrue
    }

    It 'round-trips: saved template loads back identically' {
        $tplPath = Join-Path $TestDrive 'roundtrip'
        $null = Save-WindowTemplate -TemplatesPath $tplPath -Name 'RT' `
            -RecurrenceType 'Weekly' -DurationHours 6 -DurationMinutes 30 `
            -StartHour 22 -StartMinute 15 -DayOfWeek Saturday `
            -Description 'Round-trip test' -PatchTuesdayOffset 0
        $loaded = @(Get-WindowTemplates -TemplatesPath $tplPath)
        $loaded.Count                     | Should -Be 1
        $loaded[0].Name                   | Should -Be 'RT'
        $loaded[0].RecurrenceType         | Should -Be 'Weekly'
        $loaded[0].DurationHours          | Should -Be 6
        $loaded[0].DurationMinutes        | Should -Be 30
        $loaded[0].StartHour              | Should -Be 22
        $loaded[0].StartMinute            | Should -Be 15
        $loaded[0].DayOfWeek              | Should -Be 'Saturday'
        $loaded[0].PatchTuesdayOffset     | Should -Be 0
    }
}

Describe 'Remove-WindowTemplate' {
    It 'deletes an existing template file' {
        $tplPath = Join-Path $TestDrive 'remove'
        $path = Save-WindowTemplate -TemplatesPath $tplPath -Name 'To Delete' `
            -RecurrenceType 'Daily' -DurationHours 1 -StartHour 0
        Test-Path -LiteralPath $path | Should -BeTrue
        $ok = Remove-WindowTemplate -FilePath $path
        $ok | Should -BeTrue
        Test-Path -LiteralPath $path | Should -BeFalse
    }

    It 'returns false when the file does not exist' {
        $bogus = Join-Path $TestDrive 'nonexistent.json'
        $ok = Remove-WindowTemplate -FilePath $bogus
        $ok | Should -BeFalse
    }
}
