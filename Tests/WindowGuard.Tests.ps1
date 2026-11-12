#Requires -Modules Pester
# Tests of the window-test guard (Test-WindowTestsAllowed in WindowHarness.ps1).
# The guard runs window tests only when SUITE_ALLOW_WINDOW_TESTS is 1. Most tests
# inject the opt-in value. This file opens no window.

BeforeDiscovery {
    # Every other test file that can open a window must carry the guard.
    # Same scope as Invoke-Checks.ps1: nested folders count, the lab scripts do not run as tests.
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:WindowTestFiles = @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.Tests.ps1' -File -Recurse |
            Where-Object { $rel = $_.FullName.Substring($script:RepoRoot.Length + 1); $_.Name -ne 'WindowGuard.Tests.ps1' -and $rel -notlike 'Tests\Lab\*' -and $rel -notlike '.git\*' -and $rel -notlike '.claude\*' -and (Get-Content -LiteralPath $_.FullName -Raw) -match 'Invoke-WpfSta|New-WpfWindow|ShowDialog|\.Show\(\)' } |
            ForEach-Object { @{ Name = $_.FullName.Substring($script:RepoRoot.Length + 1); Path = $_.FullName } })
}

BeforeAll {
    . "$PSScriptRoot\WindowHarness.ps1"
    function Get-Guard {
        param([string]$Value)
        Test-WindowTestsAllowed -OptIn $Value
    }
    # Reads the outermost Describe blocks of a script from its syntax tree, so a comment or a
    # continuation line cannot fool the check. A Describe inside a control statement (if, foreach)
    # counts. A Describe nested inside another Describe needs no guard of its own.
    function Get-TopLevelDescribe {
        param([string]$Text)
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($Text, [ref]$null, [ref]$null)
        $isDescribe = { param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Describe' }
        foreach ($command in $ast.FindAll($isDescribe, $true)) {
            $nested = $false
            for ($p = $command.Parent; $p; $p = $p.Parent) {
                if (& $isDescribe $p) { $nested = $true; break }
            }
            if ($nested) { continue }
            $elements = $command.CommandElements
            $guarded = $false
            for ($i = 1; $i -lt $elements.Count; $i++) {
                $e = $elements[$i]
                if ($e -isnot [System.Management.Automation.Language.CommandParameterAst] -or $e.ParameterName -ne 'Skip') { continue }
                $arg = if ($e.Argument) { $e.Argument } elseif ($i + 1 -lt $elements.Count) { $elements[$i + 1] } else { $null }
                if ($arg -and ($arg.Extent.Text -replace '\s', '') -eq '(-not$script:WindowGuard.Allowed)') { $guarded = $true }
            }
            $name = if ($elements.Count -gt 1 -and $elements[1] -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $elements[1].Value } else { '(unnamed)' }
            [pscustomobject]@{ Name = $name; Guarded = $guarded }
        }
    }
    function Get-UnguardedDescribe {
        param([string]$Text)
        Get-TopLevelDescribe -Text $Text | Where-Object { -not $_.Guarded } | ForEach-Object Name
    }
}

Describe 'Test-WindowTestsAllowed' {
    It 'allows the opt-in 1' {
        $r = Get-Guard -Value '1'
        $r.Allowed | Should -BeTrue
        $r.Reason | Should -BeNullOrEmpty
    }

    It 'refuses when the opt-in is <Label>, and names the opt-in' -ForEach @(
        @{ Label = 'empty'; Value = '' }
        @{ Label = 'null'; Value = $null }
        @{ Label = '0'; Value = '0' }
        @{ Label = 'true'; Value = 'true' }
        @{ Label = 'yes'; Value = 'yes' }
    ) {
        $r = Get-Guard -Value $Value
        $r.Allowed | Should -BeFalse
        $r.Reason | Should -Be 'SUITE_ALLOW_WINDOW_TESTS is not 1'
    }
}

Describe 'Test-WindowTestsAllowed without arguments' {
    BeforeAll { $script:SavedOptIn = $env:SUITE_ALLOW_WINDOW_TESTS }
    AfterAll { $env:SUITE_ALLOW_WINDOW_TESTS = $script:SavedOptIn }

    It 'allows when the variable is 1' {
        $env:SUITE_ALLOW_WINDOW_TESTS = '1'
        (Test-WindowTestsAllowed).Allowed | Should -BeTrue
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

Describe 'Test-WindowTestsAllowed decides on the opt-in only' {
    It 'has no ComputerName parameter' {
        (Get-Command Test-WindowTestsAllowed).Parameters.ContainsKey('ComputerName') | Should -BeFalse
    }

    It 'takes the OptIn parameter and no other parameter of its own' {
        $own = @((Get-Command Test-WindowTestsAllowed).Parameters.Keys | Where-Object { $_ -notin [System.Management.Automation.PSCmdlet]::CommonParameters })
        $own | Should -Be @('OptIn')
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
        $text = Get-Content -LiteralPath $Path -Raw
        @(Get-TopLevelDescribe -Text $text).Count | Should -BeGreaterThan 0
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }
}

Describe 'Get-UnguardedDescribe reads the syntax tree' {
    It 'accepts a guard on a continuation line' {
        $text = "Describe 'a' ``" + [Environment]::NewLine + "    -Skip:(-not `$script:WindowGuard.Allowed) { It 'x' { } }"
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }

    It 'accepts the Skip parameter with a space instead of a colon' {
        $text = "Describe 'a' -Skip (-not `$script:WindowGuard.Allowed) { }"
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }

    It 'reports a Describe whose guard text is only in a comment' {
        $text = "# -Skip:(-not `$script:WindowGuard.Allowed)" + [Environment]::NewLine + "Describe 'a' { }"
        @(Get-UnguardedDescribe -Text $text) | Should -Be @('a')
    }

    It 'reports a Describe with another Skip value' {
        @(Get-UnguardedDescribe -Text "Describe 'a' -Skip:`$false { }") | Should -Be @('a')
    }

    It 'does not require a nested Describe to repeat the guard' {
        $text = "Describe 'a' -Skip:(-not `$script:WindowGuard.Allowed) { Describe 'b' { } }"
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }

    It 'reports an unguarded Describe inside an if block' {
        $text = "Describe 'a' -Skip:(-not `$script:WindowGuard.Allowed) { }" + [Environment]::NewLine + "if (`$true) { Describe 'b' { } }"
        @(Get-UnguardedDescribe -Text $text) | Should -Be @('b')
    }

    It 'reports an unguarded Describe inside a foreach block' {
        $text = "foreach (`$i in 1..2) { Describe 'c' { } }"
        @(Get-UnguardedDescribe -Text $text) | Should -Be @('c')
    }

    It 'accepts a guarded Describe inside an if block' {
        $text = "if (`$true) { Describe 'd' -Skip:(-not `$script:WindowGuard.Allowed) { } }"
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }

    It 'does not require a Describe in an if block inside a Describe to repeat the guard' {
        $text = "Describe 'a' -Skip:(-not `$script:WindowGuard.Allowed) { if (`$true) { Describe 'b' { } } }"
        @(Get-UnguardedDescribe -Text $text) | Should -BeNullOrEmpty
    }

    It 'reports only the unguarded one of two top-level Describe blocks' {
        $text = "Describe 'a' -Skip:(-not `$script:WindowGuard.Allowed) { }" + [Environment]::NewLine + "Describe 'b' { }"
        @(Get-UnguardedDescribe -Text $text) | Should -Be @('b')
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
