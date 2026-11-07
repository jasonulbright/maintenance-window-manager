#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1). The guard gets
# an injected computer name and opt-in value. This file opens no window.

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    # Wrapper with variables: the linter rejects a literal after -ComputerName.
    function Get-Guard {
        param([string]$Computer, [string]$Value)
        Test-WindowTestsAllowed -ComputerName $Computer -OptIn $Value
    }
}

Describe 'Test-WindowTestsAllowed' {
    It 'refuses on the dev PC even with the opt-in' {
        $r = Get-Guard -Computer 'AYW' -Value '1'
        $r.Allowed | Should -BeFalse
    }

    It 'compares the dev PC name without regard to case: <Name>' -ForEach @(@{ Name = 'ayw' }, @{ Name = 'Ayw' }, @{ Name = 'AYW' }) {
        (Get-Guard -Computer $Name -Value '1').Allowed | Should -BeFalse
    }

    It 'allows another computer with the opt-in' {
        $r = Get-Guard -Computer 'CM01' -Value '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses another computer when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
        @{ Label = 'missing'; Value = $null }
    ) {
        $r = Get-Guard -Computer 'CLIENT01' -Value $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Match 'SUITE_ALLOW_WINDOW_TESTS'
    }

    It 'gives a reason text whenever it refuses: <Name> with <Value>' -ForEach @(
        @{ Name = 'AYW'; Value = '1' }
        @{ Name = 'AYW'; Value = '' }
        @{ Name = 'CM01'; Value = '' }
        @{ Name = 'CM01'; Value = '0' }
    ) {
        $r = Get-Guard -Computer $Name -Value $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Not -BeNullOrEmpty
    }
}

Describe 'Window test files' {
    BeforeAll {
        $script:OpensWindows = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' -File |
                Where-Object { $_.Name -ne 'WindowGuard.Tests.ps1' } |
                Where-Object { Select-String -LiteralPath $_.FullName -Pattern 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' -Quiet })
    }

    It 'finds the window test file' {
        $script:OpensWindows.Name | Should -Contain 'Window.Tests.ps1'
    }

    It 'has the guard in every test file that opens a window' {
        $missing = @($script:OpensWindows | Where-Object { -not (Select-String -LiteralPath $_.FullName -Pattern 'Test-WindowTestsAllowed' -Quiet) } | ForEach-Object Name)
        $missing | Should -BeNullOrEmpty
    }

    It 'skips with the guard in every test file that opens a window' {
        $missing = @($script:OpensWindows | Where-Object { -not (Select-String -LiteralPath $_.FullName -Pattern '-Skip:\(-not \$script:WindowGuard\.Allowed\)' -Quiet) } | ForEach-Object Name)
        $missing | Should -BeNullOrEmpty
    }
}
