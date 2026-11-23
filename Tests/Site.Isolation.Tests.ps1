#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }

<#
.SYNOPSIS
    Two-site isolation of the per-site data of Maintenance Window Manager (Q-112f).

.DESCRIPTION
    Each test seeds an old-layout tool folder in TestDrive (prefs, reports, a log, the window state), points
    APP_PACKAGER_DATA_ROOT at TestDrive, and starts the tool logic for site AB1, then CD2, then AB1 again.
    The tool logic is the real code: Resolve-SuiteToolSite from the vendored SuiteCommon, Get-MwmPreferences and
    Save-MwmPreferences of the start script (cut out of its syntax tree, as in Shell.Logic.Tests.ps1), and
    Resolve-MwmReportsPath of the module. Site codes are fictional (AB1, CD2). Hosts are example.com.
    No window opens and no ConfigMgr connection is needed.
#>

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    $script:ShellPath = Join-Path $script:AppRoot 'start-maintenancewindowmgr.ps1'
    $script:ShellText = Get-Content -LiteralPath $script:ShellPath -Raw
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:ShellPath, [ref]$null, [ref]$null)
    foreach ($n in 'Get-MwmPreferences', 'Save-MwmPreferences') {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true)
        if (-not $def) { throw "function $n not found in the script" }
        . ([scriptblock]::Create($def.Extent.Text))
    }
    $script:EnvNames = 'SUITE_CM_SITECODE', 'APP_PACKAGER_SITE', 'APP_PACKAGER_DATA_ROOT', 'SUITE_CM_PROVIDER'

    function Write-Seed {
        param([string]$Root, [string]$Relative, [string]$Text)
        $p = Join-Path $Root $Relative
        $d = Split-Path $p -Parent
        if (-not (Test-Path -LiteralPath $d)) { [void](New-Item -ItemType Directory -Path $d -Force) }
        [IO.File]::WriteAllText($p, $Text, (New-Object Text.UTF8Encoding($false)))
    }
    function Get-TreeHash {
        param([string]$Root)
        $h = @{}
        if (Test-Path -LiteralPath $Root) { foreach ($f in Get-ChildItem -LiteralPath $Root -Recurse -File -Force) { $h[$f.FullName] = (Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256).Hash } }
        return $h
    }
    # An old tool folder beside the script, as the tool wrote it before the per-site layout.
    function New-OldToolFolder {
        param([string]$Suite, [switch]$NoSiteCode)
        $t = Join-Path $Suite 'maintenance-window-manager'
        $prefs = [ordered]@{ DarkMode = $false }
        if (-not $NoSiteCode) { $prefs['SiteCode'] = 'AB1'; $prefs['SMSProvider'] = 'cm01.example.com' }
        Write-Seed $t 'MaintWindowMgr.prefs.json' ($prefs | ConvertTo-Json)
        Write-Seed $t 'MaintWindowMgr.windowstate.json' '{"ActiveView":"Windows"}'
        Write-Seed $t 'Reports\MWM-Windows-20261120-100000.csv' ('"Name","CollectionID"' + "`n" + '"AB1 windows","AB100001"' + "`n")
        Write-Seed $t 'Reports\MWM-Windows-20261120-100000.html' '<html><body>AB1 windows</body></html>'
        Write-Seed $t 'Logs\MWM-20261120-100000.log' 'global log'
        Write-Seed $t 'Templates\operational-collections.json' '[]'
        return $t
    }
    # What the start script does at start: resolve the site, read the prefs, resolve the reports folder.
    function Start-Tool {
        param([string]$Folder)
        if ($Folder) { [Environment]::SetEnvironmentVariable('APP_PACKAGER_SITE', $Folder) } else { [Environment]::SetEnvironmentVariable('APP_PACKAGER_SITE', $null) }
        $global:PrefsPath = Join-Path $script:ToolRoot 'MaintWindowMgr.prefs.json'
        $script:Site = Resolve-SuiteToolSite -Tool 'maintenance-window-manager' -ToolRoot $script:ToolRoot -LegacyPrefsPath $global:PrefsPath
        $global:Prefs = Get-MwmPreferences
        return @{
            Site    = $script:Site
            Prefs   = $global:Prefs
            Reports = Resolve-MwmReportsPath -ToolRoot $script:ToolRoot -Site $script:Site -Create
        }
    }
}

AfterAll {
    Remove-Variable -Name PrefsPath, Prefs -Scope Global -ErrorAction SilentlyContinue
}

Describe 'Per-site data of Maintenance Window Manager' {
    BeforeAll {
        $script:Saved = @{}
        foreach ($n in $script:EnvNames) { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n); [Environment]::SetEnvironmentVariable($n, $null) }
        # Short names: the path limit of 260 characters in Windows PowerShell 5.1 counts the data root, the SID and the .migrating copy.
        $temp = Join-Path $TestDrive 'a'
        $script:Suite = Join-Path $temp 'Suite'
        $script:Data = Join-Path $temp 'D'
        [void](New-Item -ItemType Directory -Path $script:Suite, $script:Data -Force)
        [Environment]::SetEnvironmentVariable('APP_PACKAGER_DATA_ROOT', $script:Data)
        $script:ToolRoot = New-OldToolFolder -Suite $script:Suite
        $script:OldHash = Get-TreeHash -Root $script:ToolRoot
        $script:DataRoot = Get-SuiteDataRoot

        # Start for AB1, save a changed provider and write one report.
        $script:Ab1 = Start-Tool -Folder 'AB1'
        $script:Ab1Prefs = $script:Ab1.Prefs.Clone()
        $global:Prefs['SMSProvider'] = 'cm02.example.com'
        Save-MwmPreferences -Prefs $global:Prefs
        Write-Seed $script:Ab1.Reports 'MWM-Windows-20261121-090000.csv' '"Name"'
        $script:Ab1After = Get-TreeHash -Root (Join-Path $script:DataRoot 'sites\AB1')

        # Start for CD2 and write there.
        $script:Cd2 = Start-Tool -Folder 'CD2'
        $script:Cd2Prefs = $script:Cd2.Prefs.Clone()
        Write-Seed $script:Cd2.Reports 'MWM-Windows-20261121-100000.csv' '"CD2 windows"'
        $global:Prefs['SMSProvider'] = 'cm03.example.com'
        Save-MwmPreferences -Prefs $global:Prefs

        # Back to AB1. A report that the first start copied is deleted first: a later start must not copy it again.
        $script:DeletedCopy = Join-Path $script:DataRoot 'sites\AB1\tools\maintenance-window-manager\reports\MWM-Windows-20261120-100000.html'
        $script:HtmlCopied = Test-Path -LiteralPath $script:DeletedCopy
        Remove-Item -LiteralPath $script:DeletedCopy -Force
        $script:Ab1Again = Start-Tool -Folder 'AB1'
    }

    AfterAll {
        foreach ($n in $script:EnvNames) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
    }

    It 'keeps every old file unchanged (SHA-256), except the prefs file, which only gets global keys' {
        $now = Get-TreeHash -Root $script:ToolRoot
        foreach ($k in $script:OldHash.Keys) {
            $now.ContainsKey($k) | Should -BeTrue -Because "$k must still exist"
            if ($k -like '*MaintWindowMgr.prefs.json') { continue }
            $now[$k] | Should -Be $script:OldHash[$k] -Because "$k must keep its content"
        }
        $oldPrefs = Get-Content -LiteralPath (Join-Path $script:ToolRoot 'MaintWindowMgr.prefs.json') -Raw | ConvertFrom-Json
        $oldPrefs.SMSProvider | Should -Be 'cm01.example.com' -Because 'the old file keeps the site keys it already had'
        @(Get-ChildItem -LiteralPath (Join-Path $script:ToolRoot 'Reports') -File).Count | Should -Be 2 -Because 'no new report went to the old folder'
    }

    It 'copies the AB1 reports once into sites\AB1\tools\maintenance-window-manager' {
        $tool = Join-Path $script:DataRoot 'sites\AB1\tools\maintenance-window-manager'
        $script:Ab1.Site.Active | Should -BeTrue -Because $script:Ab1.Site.Reason
        $script:Ab1.Site.ToolFolder | Should -Be $tool
        $script:Ab1.Reports | Should -Be (Join-Path $tool 'reports')
        (Get-Content -LiteralPath (Join-Path $tool 'reports\MWM-Windows-20261120-100000.csv') -Raw) | Should -Match 'AB1 windows'
        $script:HtmlCopied | Should -BeTrue
        (Get-Content -LiteralPath (Join-Path $tool 'migrated.json') -Raw | ConvertFrom-Json).tool | Should -Be 'maintenance-window-manager'
    }

    It 'reads the AB1 site keys from the old prefs and keeps the other keys in the old file' {
        $script:Ab1Prefs['SiteCode'] | Should -Be 'AB1'
        $script:Ab1Prefs['SMSProvider'] | Should -Be 'cm01.example.com'
        $script:Ab1Prefs['DarkMode'] | Should -BeFalse
        $siteFile = Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\AB1\tools\maintenance-window-manager\prefs.json') -Raw | ConvertFrom-Json
        $siteFile.SiteCode | Should -Be 'AB1'
        $siteFile.PSObject.Properties.Name | Should -Not -Contain 'DarkMode'
    }

    It 'shows none of the AB1 values or files for CD2' {
        $script:Cd2.Site.SiteFolder | Should -Be 'CD2'
        $script:Cd2Prefs['SiteCode'] | Should -Be 'CD2'
        $script:Cd2Prefs['SMSProvider'] | Should -Be ''
        $script:Cd2Prefs['SMSProvider'] | Should -Not -Match 'example.com'
        $script:Cd2Prefs['DarkMode'] | Should -BeFalse -Because 'the theme is a global key'
        $tool = Join-Path $script:DataRoot 'sites\CD2\tools\maintenance-window-manager'
        $script:Cd2.Reports | Should -Be (Join-Path $tool 'reports')
        Test-Path -LiteralPath (Join-Path $tool 'reports\MWM-Windows-20261120-100000.csv') | Should -BeFalse
        @(Get-ChildItem -LiteralPath $tool -Recurse -File | Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'AB1' -and $_.Name -ne 'migrated.json' -and $_.Name -ne 'prefs.json' }) | Should -BeNullOrEmpty
    }

    It 'writes the CD2 data only under sites\CD2 and leaves AB1 and the old folder alone' {
        $cd2Root = Join-Path $script:DataRoot 'sites\CD2'
        Test-Path -LiteralPath (Join-Path $cd2Root 'tools\maintenance-window-manager\reports\MWM-Windows-20261121-100000.csv') | Should -BeTrue
        $ab1Now = Get-TreeHash -Root (Join-Path $script:DataRoot 'sites\AB1')
        foreach ($k in $script:Ab1After.Keys) {
            if ($k -like '*prefs.json' -or $k -eq $script:DeletedCopy) { continue }
            $ab1Now[$k] | Should -Be $script:Ab1After[$k] -Because "$k of AB1 must not change while CD2 runs"
        }
        foreach ($k in $script:OldHash.Keys) {
            if ($k -like '*MaintWindowMgr.prefs.json') { continue }
            (Get-FileHash -LiteralPath $k -Algorithm SHA256).Hash | Should -Be $script:OldHash[$k]
        }
    }

    It 'gives the AB1 values back when the tool starts for AB1 again' {
        $script:Ab1Again.Site.SiteFolder | Should -Be 'AB1'
        $script:Ab1Again.Prefs['SiteCode'] | Should -Be 'AB1'
        $script:Ab1Again.Prefs['SMSProvider'] | Should -Be 'cm02.example.com' -Because 'AB1 saved this provider before CD2 started'
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.Reports 'MWM-Windows-20261121-090000.csv') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.Reports 'MWM-Windows-20261121-100000.csv') | Should -BeFalse
    }

    It 'copies the old data only at the first start for a site' {
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.Reports 'MWM-Windows-20261120-100000.html') | Should -BeFalse -Because 'the copy runs once, so a deleted file is not restored'
        Test-Path -LiteralPath (Join-Path $script:Ab1Again.Reports 'MWM-Windows-20261120-100000.csv') | Should -BeTrue
    }

    It 'keeps a different provider for each site' {
        (Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\CD2\tools\maintenance-window-manager\prefs.json') -Raw | ConvertFrom-Json).SMSProvider | Should -Be 'cm03.example.com'
        (Get-Content -LiteralPath (Join-Path $script:DataRoot 'sites\AB1\tools\maintenance-window-manager\prefs.json') -Raw | ConvertFrom-Json).SMSProvider | Should -Be 'cm02.example.com'
    }
}

Describe 'Maintenance Window Manager with no site chosen' {
    BeforeAll {
        $script:Saved = @{}
        foreach ($n in $script:EnvNames) { $script:Saved[$n] = [Environment]::GetEnvironmentVariable($n); [Environment]::SetEnvironmentVariable($n, $null) }
        # Short names: the path limit of 260 characters in Windows PowerShell 5.1 counts the data root, the SID and the .migrating copy.
        $temp = Join-Path $TestDrive 'b'
        $script:Suite = Join-Path $temp 'Suite'
        $script:Data = Join-Path $temp 'D'
        [void](New-Item -ItemType Directory -Path $script:Suite, $script:Data -Force)
        [Environment]::SetEnvironmentVariable('APP_PACKAGER_DATA_ROOT', $script:Data)
        $script:ToolRoot = New-OldToolFolder -Suite $script:Suite -NoSiteCode
        $script:OldHash = Get-TreeHash -Root $script:ToolRoot
        $script:None = Start-Tool -Folder ''
    }
    AfterAll {
        foreach ($n in $script:EnvNames) { [Environment]::SetEnvironmentVariable($n, $script:Saved[$n]) }
    }

    It 'is not active and uses the Reports folder beside the script' {
        $script:None.Site.Active | Should -BeFalse
        $script:None.Reports | Should -Be (Join-Path $script:ToolRoot 'Reports')
    }

    It 'creates no site folder and changes no old file' {
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
        $now = Get-TreeHash -Root $script:ToolRoot
        foreach ($k in $script:OldHash.Keys) { $now[$k] | Should -Be $script:OldHash[$k] }
    }

    It 'saves the site keys and the other keys to the old prefs file' {
        $global:Prefs['SiteCode'] = 'EF3'
        $global:Prefs['SMSProvider'] = 'cm04.example.com'
        Save-MwmPreferences -Prefs $global:Prefs
        $p = Get-Content -LiteralPath (Join-Path $script:ToolRoot 'MaintWindowMgr.prefs.json') -Raw | ConvertFrom-Json
        $p.SiteCode | Should -Be 'EF3'
        $p.SMSProvider | Should -Be 'cm04.example.com'
        Test-Path -LiteralPath (Join-Path (Get-SuiteDataRoot) 'sites') | Should -BeFalse
    }
}

Describe 'Reports path' {
    BeforeAll {
        $script:Root = Join-Path $TestDrive 'old-tool'
        $script:SiteObj = [pscustomobject]@{ Active = $true; ToolFolder = (Join-Path $TestDrive 'sites\AB1\tools\maintenance-window-manager') }
    }

    It 'uses the reports store of the site and the old folder with no site' {
        Resolve-MwmReportsPath -ToolRoot $script:Root -Site $script:SiteObj | Should -Be (Join-Path $script:SiteObj.ToolFolder 'reports')
        Resolve-MwmReportsPath -ToolRoot $script:Root -Site $null | Should -Be (Join-Path $script:Root 'Reports')
        Resolve-MwmReportsPath -ToolRoot $script:Root -Site ([pscustomobject]@{ Active = $false }) | Should -Be (Join-Path $script:Root 'Reports')
    }

    It 'creates the folder only with -Create' {
        $p = Resolve-MwmReportsPath -ToolRoot $script:Root -Site $script:SiteObj
        Test-Path -LiteralPath $p | Should -BeFalse
        $q = Resolve-MwmReportsPath -ToolRoot $script:Root -Site $script:SiteObj -Create
        Test-Path -LiteralPath $q | Should -BeTrue
        $o = Resolve-MwmReportsPath -ToolRoot $script:Root -Site $null -Create
        Test-Path -LiteralPath $o | Should -BeTrue
    }
}

Describe 'Static guard: no site store is built from the script folder' {
    It 'builds no reports path from $PSScriptRoot in the start script' {
        $hits = [regex]::Matches($script:ShellText, "PSScriptRoot\s+'Reports")
        @($hits | ForEach-Object { $_.Value }) | Should -BeNullOrEmpty
    }

    It 'resolves the reports folder only through Resolve-MwmReportsPath (syntax tree)' {
        $calls = $script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.CommandAst] -and $x.GetCommandName() -eq 'Resolve-MwmReportsPath' }, $true)
        @($calls).Count | Should -Be 2 -Because 'Export CSV and Export HTML'
        $joins = $script:Ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.StringConstantExpressionAst] -and $x.Value -eq 'Reports' }, $true)
        @($joins).Count | Should -Be 0
    }

    It 'stores no absolute path in the prefs defaults' {
        $def = $script:Ast.Find({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq 'Get-MwmPreferences' }, $true)
        $def.Extent.Text | Should -Not -Match 'Join-Path|PSScriptRoot'
    }

    It 'keeps the Site Code box read-only and the site code unchanged in Options when a site is active' {
        $script:ShellText | Should -Match 'if \(\$script:Site\.Active\) \{[^}]*\$txtSiteCode\.IsReadOnly = \$true'
        $script:ShellText | Should -Match 'The site is chosen in the AppPackager Suite launcher\. Restart this tool after you change it\.'
        $script:ShellText | Should -Match 'if \(\$script:Site\.Active\) \{ \$newSite = \[string\]\$global:Prefs\.SiteCode \}'
    }

    It 'uses the launcher variables only when no site is active and passes the reports folder to the dialogs' {
        $script:ShellText | Should -Match 'if \(-not \$script:Site\.Active\) \{\s+if \(-not \$global:Prefs\.SiteCode\s+-and \$env:SUITE_CM_SITECODE\)'
        @([regex]::Matches($script:ShellText, 'InitialDirectory = Resolve-MwmReportsPath -ToolRoot \$PSScriptRoot -Site \$script:Site -Create')).Count | Should -Be 2
    }

    It 'reads and saves the prefs through the site prefs functions' {
        $script:ShellText | Should -Match 'Read-SuiteToolPrefs'
        $script:ShellText | Should -Match 'Save-SuiteToolPrefs'
        $script:ShellText | Should -Not -Match 'Read-SuiteSettings'
        $script:ShellText | Should -Not -Match 'Save-SuiteSettings'
    }
}
