#Requires -Modules Pester
# Every ConfigMgr cmdlet call in the module names only parameters that the real
# cmdlet has. Mocks accept any parameter name the test stub declares, so a wrong
# parameter name passes a mocked test and fails on a site. This test checks against
# the real cmdlet signature, and is skipped on a machine without the ConfigMgr
# console. It reads the parameters written as -Name, and the keys of a splatted
# hashtable (@params) that the function builds with a literal or with $params['Key'].

BeforeDiscovery {
    . "$PSScriptRoot\TestHelpers.ps1"
    $script:CmAvailable = $false
    $candidates = @('C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1',
        'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1')
    if (Get-Command Get-CMApplication -ErrorAction SilentlyContinue) { $script:CmAvailable = $true }
    else {
        foreach ($c in $candidates) {
            if (Test-Path -LiteralPath $c) {
                try { Import-Module $c -DisableNameChecking -ErrorAction Stop -WarningAction SilentlyContinue; $script:CmAvailable = [bool](Get-Command Get-CMApplication -ErrorAction SilentlyContinue); break } catch { $null = $_ }
            }
        }
    }
    $script:Calls = @(Get-CmCmdletCall -Path (Join-Path $script:AppRoot 'Module\MaintWindowMgrCommon.psm1'))
}

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    $script:Found = @(Get-CmCmdletCall -Path (Join-Path $script:AppRoot 'Module\MaintWindowMgrCommon.psm1'))
}

Describe 'ConfigMgr cmdlet calls' {
    It 'finds the ConfigMgr cmdlet calls of the module' {
        $script:Found.Count | Should -BeGreaterThan 8
    }

    It 'reads the keys of a splatted hashtable and the named parameters' {
        $f = Join-Path $TestDrive 'sample.ps1'
        Set-Content -LiteralPath $f -Value @(
            'function Demo {'
            '    $p = @{ A = 1; ErrorAction = ''Stop'' }'
            '    $p[''B''] = 2'
            '    $p.C = 3'
            '    New-CMThing @p -D 4'
            '    Copy-CMOther -E 5'
            '}')
        $calls = @(Get-CmCmdletCall -Path $f)
        $calls.Count | Should -Be 2
        ($calls | Where-Object Command -eq 'New-CMThing').Parameters | Should -Be @('A', 'ErrorAction', 'B', 'C', 'D')
        ($calls | Where-Object Command -eq 'Copy-CMOther').Parameters | Should -Be @('E')
    }

    It '<Command> at line <Line> uses only real parameters (<Parameters>)' -ForEach $script:Calls -Skip:(-not $script:CmAvailable) {
        $cmd = Get-Command $Command -ErrorAction Stop
        foreach ($p in $Parameters) {
            $exact = @($cmd.Parameters.Values | Where-Object { $_.Name -eq $p -or $_.Aliases -contains $p })
            if ($exact.Count -eq 0) {
                # PowerShell binds a prefix that names exactly one parameter.
                $exact = @($cmd.Parameters.Keys | Where-Object { $_ -like "$p*" })
                $exact.Count | Should -Be 1 -Because "-$p must name exactly one parameter of $Command (found $($exact.Count))"
            }
        }
    }
}
