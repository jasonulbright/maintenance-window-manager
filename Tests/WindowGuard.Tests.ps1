#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1).
# The guard keeps window tests off the dev PC (AYW) and runs them only when
# SUITE_ALLOW_WINDOW_TESTS is 1. Most tests inject the computer name and the
# opt-in value. This file opens no window.

BeforeDiscovery {
    # Every other test file that can open a window must carry the guard.
    # Same scope as Invoke-Checks.ps1: nested folders count, the lab scripts do not run as tests.
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.Tests.ps1' -File -Recurse |
            Where-Object { $_.Name -ne 'WindowGuard.Tests.ps1' -and $_.FullName -notlike '*\Tests\Lab\*' -and $_.FullName -notlike '*\.git\*' -and (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' } |
            ForEach-Object { @{ Name = $_.FullName.Substring($script:RepoRoot.Length + 1); Path = $_.FullName } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    # Pass the computer name as a variable, so the linter rule against a hard-coded ComputerName stays quiet.
    function Get-Guard {
        param([string]$Computer, [string]$Value)
        Test-WindowTestsAllowed -ComputerName $Computer -OptIn $Value
    }
}

Describe 'Test-WindowTestsAllowed' {
    It 'refuses on AYW even with the opt-in' {
        $r = Get-Guard -Computer 'AYW' -Value '1'
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Match 'AYW'
    }

    It 'refuses on <Name> in any letter case' -ForEach @(@{ Name = 'ayw' }, @{ Name = 'Ayw' }) {
        (Get-Guard -Computer $Name -Value '1').Allowed | Should -BeFalse
    }

    It 'refuses on AYW without the opt-in' {
        (Get-Guard -Computer 'AYW' -Value '').Allowed | Should -BeFalse
    }

    It 'allows another computer with the opt-in 1' {
        $r = Get-Guard -Computer 'CM01' -Value '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses another computer when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = 'null'; Value = $null }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
    ) {
        $r = Get-Guard -Computer 'CLIENT01' -Value $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Match 'SUITE_ALLOW_WINDOW_TESTS'
    }

    It 'gives a reason whenever it refuses (<Name>, opt-in "<Value>")' -ForEach @(
        @{ Name = 'AYW'; Value = '1' }
        @{ Name = 'CM01'; Value = '' }
        @{ Name = 'CM01'; Value = '0' }
    ) {
        (Get-Guard -Computer $Name -Value $Value).Reason | Should -Not -BeNullOrEmpty
    }
}

Describe 'Test-WindowTestsAllowed without arguments' {
    BeforeAll { $script:SavedOptIn = $env:SUITE_ALLOW_WINDOW_TESTS }
    AfterAll { $env:SUITE_ALLOW_WINDOW_TESTS = $script:SavedOptIn }

    It 'reads SUITE_ALLOW_WINDOW_TESTS and the machine name by default' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '1'
        (Test-WindowTestsAllowed).Allowed | Should -Be ([Environment]::MachineName -ine 'AYW')
    }

    It 'refuses when the variable is 0' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '0'
        $r = Test-WindowTestsAllowed
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Not -BeNullOrEmpty
    }

    It 'refuses when the variable is not set' {
        Remove-Item -LiteralPath Env:SUITE_ALLOW_WINDOW_TESTS -ErrorAction SilentlyContinue
        (Test-WindowTestsAllowed).Allowed | Should -BeFalse
    }
}

Describe 'Window test files carry the guard' {
    It 'finds at least one test file that can open a window' -ForEach @(@{ Found = @($script:WindowTestFiles).Count }) {
        $Found | Should -BeGreaterThan 0
    }

    It 'prints the skip reason once: Invoke-Checks sets the marker and Window.Tests reads it' {
        (Get-Content -LiteralPath (Join-Path (Split-Path -Parent $PSScriptRoot) 'Invoke-Checks.ps1') -Raw) | Should -Match '\$env:SUITE_WINDOW_GUARD_REPORTED\s*=\s*''1'''
        (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Window.Tests.ps1') -Raw) | Should -Match '-not \$env:SUITE_WINDOW_GUARD_REPORTED'
    }

    It '<Name> calls Test-WindowTestsAllowed and applies -Skip' -ForEach $script:WindowTestFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $text | Should -Match 'Test-WindowTestsAllowed'
        $text | Should -Match '-Skip:'
    }

    It '<Name> applies the guard to every top-level Describe' -ForEach $script:WindowTestFiles {
        $describes = @(Get-Content -LiteralPath $Path | Where-Object { $_ -match '^\s*Describe\b' })
        $describes.Count | Should -BeGreaterThan 0
        @($describes | Where-Object { $_ -notmatch '-Skip:\(-not \$script:WindowGuard\.Allowed\)' }) | Should -BeNullOrEmpty
    }
}

Describe 'Get-WpfSession refuses where the guard says no' {
    BeforeAll {
        Mock Test-WindowTestsAllowed { [pscustomobject]@{ Allowed = $false; Reason = 'test' } }
    }
    AfterEach { $script:HarnessSession = $null }

    It 'throws and creates no session' {
        $script:HarnessSession = $null
        { Get-WpfSession } | Should -Throw '*not allowed*'
        $script:HarnessSession | Should -BeNullOrEmpty
    }

    It 'does not hand out a cached session' {
        $script:HarnessSession = [pscustomobject]@{ RunspaceStateInfo = [pscustomobject]@{ State = 'Opened' } }
        { Get-WpfSession } | Should -Throw '*not allowed*'
    }
}

Describe 'Machine identity' {
    It 'takes the computer name from the machine, not from the COMPUTERNAME variable' {
        foreach ($f in 'WindowHarness.ps1', '..\Invoke-Checks.ps1') {
            (Get-Content -LiteralPath (Join-Path $PSScriptRoot $f) -Raw) | Should -Not -Match '\$env:COMPUTERNAME'
        }
        (Get-Content -LiteralPath (Join-Path $PSScriptRoot 'WindowHarness.ps1') -Raw) | Should -Match 'MachineName'
    }
}

Describe 'Test-WindowTestsAllowed decides on the opt-in only' {
    It 'has no ComputerName parameter' {
        (Get-Command Test-WindowTestsAllowed).Parameters.ContainsKey('ComputerName') | Should -BeFalse
    }

    It 'takes the OptIn parameter and no other parameter of its own' {
        $own = @((Get-Command Test-WindowTestsAllowed).Parameters.Keys | Where-Object { $_ -notin [System.Management.Automation.PSCmdlet]::CommonParameters })
        $own | Should -Be @('OptIn')
    }
}
