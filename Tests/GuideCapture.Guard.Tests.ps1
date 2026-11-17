#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.1.0' }
# Parse-only tests of the capture tool: the lab lock (D-16). Nothing runs and no window opens.
# One capture runs at a time on a lab computer. C:\guide\capture.lock is the shared lock.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:CapturePath = Join-Path $script:Root 'tools\guide-capture\Invoke-GuideCapture.ps1'
    $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($script:CapturePath, [ref]$null, [ref]$null)
    $script:Code = [regex]::Replace((Get-Content -LiteralPath $script:CapturePath -Raw), '(?s)<#.*?#>', '')
    $script:Try = $script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.TryStatementAst] -and $null -ne $n.Finally }, $false) | Select-Object -First 1
}

Describe 'The capture tool takes the lab lock' {
    It 'parses without an error' {
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($script:CapturePath, [ref]$null, [ref]$errors)
        @($errors) | Should -BeNullOrEmpty
    }

    It 'uses C:\guide\capture.lock' {
        $script:Code | Should -Match "\`$lockFile = 'C:\\guide\\capture\.lock'"
    }

    It 'creates the lock with CreateNew and treats an IOException as busy' {
        $script:Code | Should -Match 'FileMode\]::CreateNew'
        $script:Code | Should -Match 'catch \[System\.IO\.IOException\]'
        $script:Code | Should -Match "return 'busy'"
    }

    It 'removes a lock that is older than 60 minutes and waits at most 60 minutes' {
        $script:Code | Should -Match 'TotalMinutes -gt 60'
        $script:Code | Should -Match '\(Get-Date\)\.AddMinutes\(60\)'
        $script:Code | Should -Match 'Another capture holds'
    }

    It 'takes the lock before it registers the scheduled task' {
        $body = $script:Try.Body.Extent.Text
        $lock = $body.IndexOf('$haveLock = $true')
        $register = $body.IndexOf('Register-ScheduledTask')
        $lock | Should -BeGreaterThan -1
        $register | Should -BeGreaterThan $lock
    }

    It 'releases the lock in the finally, only when this capture took it' {
        $finally = $script:Try.Finally.Extent.Text
        $finally | Should -Match 'if \(\$haveLock\)'
        $finally | Should -Match 'Remove-Item \$l -Force'
        $finally.IndexOf('Remove-Item $l -Force') | Should -BeGreaterThan $finally.IndexOf('Unregister-ScheduledTask')
    }

    It 'never deletes the C:\guide folder' {
        $script:Code | Should -Not -Match 'Remove-Item\s+[''"]?C:\\guide[''"]?(\s|$)'
    }
}