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

    It 'uses C:\guide\capture.lock and the text owner|pid|start' {
        $script:Code | Should -Match "\`$lockPath = 'C:\\guide\\capture\.lock'"
        $script:Code | Should -Match "\`$lockText = '\{0\}\|\{1\}\|\{2\}' -f \`$app, \`$PID"
    }

    It 'creates the lock with New-Item and treats an IOException as busy' {
        $script:Code | Should -Match 'New-Item -ItemType File -Path \$path -Value \$text -ErrorAction Stop'
        $script:Code | Should -Match 'catch \[System\.IO\.IOException\]'
        $script:Code | Should -Match "'held: \{0\}\|\{1\}\|\{2\}'"
    }

    It 'treats a lock as stale only after 30 minutes with no running owner, and reports the removal (D-16)' {
        $script:Code | Should -Match '\[int\]\$age -gt 30 -and \$aliveLab -ne ''True'' -and -not \(Get-Process -Id \$ownerPid'
        $script:Code | Should -Match 'The stale lab lock .* was removed'
        $script:Code | Should -Match '''unknown: '' \+ \$held'
        $script:Code | Should -Not -Match 'TotalMinutes\)? -gt 60'
        $script:Code | Should -Match '\(Get-Date\)\.AddMinutes\(\$LockWaitMinutes\)'
    }

    It 'refreshes the lock in the poll loop' {
        $script:Code | Should -Match 'LastWriteTime = Get-Date'
    }

    It 'takes the lock before it registers the scheduled task' {
        $body = $script:Try.Body.Extent.Text
        $lock = $body.IndexOf('$haveLock = $true')
        $register = $body.IndexOf('Register-ScheduledTask')
        $lock | Should -BeGreaterThan -1
        $register | Should -BeGreaterThan $lock
    }

    It 'releases the lock in the finally, only when this capture took it and the text is still its own' {
        $finally = $script:Try.Finally.Extent.Text
        $finally | Should -Match 'if \(\$haveLock\)'
        $finally | Should -Match '-eq \$text\)'
        $finally.IndexOf('Remove-Item -LiteralPath $path -Force') | Should -BeGreaterThan $finally.IndexOf('Unregister-ScheduledTask')
    }

    It 'never deletes the C:\guide folder' {
        $script:Code | Should -Not -Match 'Remove-Item\s+[''"]?C:\\guide[''"]?(\s|$)'
    }
}
