#Requires -Modules Pester
# Export-MaintenanceWindowsCsv and Export-MaintenanceWindowsHtml, from a data table
# and from window objects (the Export buttons).

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    Mock Write-Log {} -ModuleName MaintWindowMgrCommon
    $script:rows = @(
        [PSCustomObject]@{ CollectionName = 'Servers'; CollectionID = 'MCM00001'; WindowName = 'Night'; Type = 'Software Updates'; Recurrence = 'Weekly'; Schedule = 'Every Sunday'; Duration = '4h 0m'; NextOccurrence = '2026-11-08 02:00'; IsEnabled = $true; Internal = 'x' },
        [PSCustomObject]@{ CollectionName = 'Lab <b> & "q"'; CollectionID = 'MCM00002'; WindowName = 'Off'; Type = 'General'; Recurrence = 'Daily'; Schedule = 'Daily'; Duration = '1h 0m'; NextOccurrence = 'N/A'; IsEnabled = $false },
        [PSCustomObject]@{ CollectionName = 'Imaging'; CollectionID = 'MCM00003'; WindowName = 'OSD'; Type = 'Task Sequences'; Recurrence = 'Daily'; Schedule = 'Daily'; Duration = '2h 0m'; NextOccurrence = 'N/A'; IsEnabled = $true })
}

Describe 'Export-MaintenanceWindowsCsv' {
    It 'writes the columns in table order and one line per row (table)' {
        $t = New-RollupTable -Column CollectionName, WindowName, Enabled -Row @(@('A', 'W1', 'True'), @('B', 'W2', 'False'))
        $path = Join-Path $TestDrive 'a.csv'
        Export-MaintenanceWindowsCsv -DataTable $t -OutputPath $path
        $r = @(Import-Csv -LiteralPath $path)
        $r.Count | Should -Be 2
        ($r[0].PSObject.Properties.Name -join ',') | Should -Be 'CollectionName,WindowName,Enabled'
        $r[1].Enabled | Should -Be 'False'
    }

    It 'writes the grid columns for window objects when no columns are given' {
        $path = Join-Path $TestDrive 'o.csv'
        Export-MaintenanceWindowsCsv -Windows $script:rows -OutputPath $path
        $r = @(Import-Csv -LiteralPath $path)
        $r.Count | Should -Be 3
        ($r[0].PSObject.Properties.Name -join ',') | Should -Be 'CollectionName,CollectionID,WindowName,Type,Recurrence,Schedule,Duration,NextOccurrence,Enabled'
        $r[0].Enabled | Should -Be 'True'
        $r[1].Enabled | Should -Be 'False'
        $r[1].CollectionName | Should -Be 'Lab <b> & "q"'
    }

    It 'writes only the requested columns' {
        $path = Join-Path $TestDrive 'c.csv'
        Export-MaintenanceWindowsCsv -Windows $script:rows -Columns 'WindowName', 'Missing' -OutputPath $path
        $r = @(Import-Csv -LiteralPath $path)
        ($r[0].PSObject.Properties.Name -join ',') | Should -Be 'WindowName,Missing'
        $r[0].Missing | Should -Be ''
    }

    It 'writes a file for an empty list' {
        $path = Join-Path $TestDrive 'e.csv'
        { Export-MaintenanceWindowsCsv -Windows @() -OutputPath $path } | Should -Not -Throw
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'creates the missing parent folders' {
        $path = Join-Path $TestDrive 'x\y\z\a.csv'
        Export-MaintenanceWindowsCsv -Windows $script:rows -OutputPath $path
        Test-Path -LiteralPath $path | Should -BeTrue
    }

    It 'round-trips the hostile value <Name>' -ForEach @(
        @{ Name = 'comma'; Value = 'a,b' }, @{ Name = 'quote'; Value = 'say "hi"' }, @{ Name = 'newline'; Value = "l1`r`nl2" }, @{ Name = 'unicode'; Value = 'Ünï 日本語' }, @{ Name = 'long'; Value = ('x' * 20000) }
    ) {
        $path = Join-Path $TestDrive "h-$Name.csv"
        Export-MaintenanceWindowsCsv -Windows @([PSCustomObject]@{ WindowName = $Value }) -Columns 'WindowName' -OutputPath $path
        (Import-Csv -LiteralPath $path)[0].WindowName | Should -Be $Value
    }

    It 'throws on an invalid path character' {
        { Export-MaintenanceWindowsCsv -Windows $script:rows -OutputPath (Join-Path $TestDrive 'bad|name?.csv') } | Should -Throw
    }

    It 'rejects a data table together with window objects' {
        { Export-MaintenanceWindowsCsv -DataTable (New-RollupTable -Column A -Row @(, @('x'))) -Windows $script:rows -OutputPath (Join-Path $TestDrive 'x.csv') } | Should -Throw
    }
}

Describe 'Export-MaintenanceWindowsHtml' {
    It 'writes one header cell per column and one row per window' {
        $path = Join-Path $TestDrive 'a.html'
        Export-MaintenanceWindowsHtml -Windows $script:rows -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        ([regex]::Matches($html, '<th>')).Count | Should -Be 9
        ([regex]::Matches($html, '<tr>')).Count | Should -Be 4
        $html | Should -Match 'Windows: 3'
        $html | Should -Match '<title>\s*Maintenance Windows Report\s*</title>'
    }

    It 'encodes the value <Value>' -ForEach @(
        @{ Value = '<script>alert(1)</script>'; Expected = '&lt;script&gt;alert(1)&lt;/script&gt;' }
        @{ Value = 'A & B'; Expected = 'A &amp; B' }
        @{ Value = '"q"'; Expected = '&quot;q&quot;' }
        @{ Value = '</td><td>x'; Expected = '&lt;/td&gt;&lt;td&gt;x' }
    ) {
        $path = Join-Path $TestDrive 'enc.html'
        Export-MaintenanceWindowsHtml -Windows @([PSCustomObject]@{ WindowName = $Value }) -Columns 'WindowName' -OutputPath $path
        (Get-Content -LiteralPath $path -Raw) | Should -Match ([regex]::Escape("<td>$Expected</td>"))
    }

    It 'encodes a column name and the report title' {
        $path = Join-Path $TestDrive 'enc2.html'
        Export-MaintenanceWindowsHtml -Windows $script:rows -Columns '<b>C</b>' -OutputPath $path -ReportTitle 'R&D <x>'
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<th>&lt;b&gt;C&lt;/b&gt;</th>'
        $html | Should -Match '<title>\s*R&amp;D &lt;x&gt;\s*</title>'
        $html | Should -Match '<h1>R&amp;D &lt;x&gt;</h1>'
    }

    It 'colors the type and marks a disabled window' {
        $path = Join-Path $TestDrive 'cls.html'
        Export-MaintenanceWindowsHtml -Windows $script:rows -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<td class="type-updates">Software Updates</td>'
        $html | Should -Match '<td class="type-general">General</td>'
        $html | Should -Match '<td class="type-osd">Task Sequences</td>'
        $html | Should -Match '<td class="disabled">False</td>'
        ([regex]::Matches($html, 'class="disabled"')).Count | Should -Be 1
    }

    It 'writes a valid page for an empty list' {
        $path = Join-Path $TestDrive 'e.html'
        Export-MaintenanceWindowsHtml -Windows @() -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match 'Windows: 0'
        $html | Should -Match '</html>\s*$'
    }

    It 'writes a data table without change' {
        $path = Join-Path $TestDrive 't.html'
        Export-MaintenanceWindowsHtml -DataTable (New-RollupTable -Column Type, Enabled -Row @(, @('General', 'False'))) -OutputPath $path
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<td class="type-general">General</td><td class="disabled">False</td>'
    }
}
