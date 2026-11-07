#Requires -Modules Pester
# Every ConfigMgr cmdlet call in the module names only parameters that the real
# cmdlet has. Mocks accept any parameter name the test stub declares, so a wrong
# parameter name (the old "Invoke-CMContentValidation -Id") passes a mocked test
# and fails on a site. This test checks against the real cmdlet signature, and is
# skipped on a machine without the ConfigMgr console.

BeforeDiscovery {
    $script:Root = Split-Path -Parent $PSScriptRoot
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
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Module\MaintWindowMgrCommon.psm1'), [ref]$null, [ref]$null)
    $script:Calls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -match '^(Get|Set|New|Remove|Start|Invoke|Add)-CM' -and $n.GetCommandName() -ne 'Get-CMConnectionInfo' }, $true) | ForEach-Object {
            $named = @($_.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } | ForEach-Object { $_.ParameterName })
            @{ Command = $_.GetCommandName(); Parameters = $named; Line = $_.Extent.StartLineNumber }
        })
}

Describe 'ConfigMgr cmdlet calls' {
    It 'finds the ConfigMgr cmdlet calls of the module' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path (Split-Path -Parent $PSScriptRoot) 'Module\MaintWindowMgrCommon.psm1'), [ref]$null, [ref]$null)
        $calls = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -match '^(Get|Set|New|Remove|Start|Invoke|Add)-CM' -and $n.GetCommandName() -ne 'Get-CMConnectionInfo' }, $true)
        @($calls).Count | Should -BeGreaterThan 8
    }

    It '<Command> at line <Line> uses only real parameters (<Parameters>)' -ForEach $script:Calls -Skip:(-not $script:CmAvailable) {
        $cmd = Get-Command $Command -ErrorAction Stop
        foreach ($p in $Parameters) {
            $known = $cmd.Parameters.Values | Where-Object { $_.Name -eq $p -or $_.Aliases -contains $p }
            # A prefix that is not ambiguous also binds, as PowerShell allows.
            if (-not $known) { $known = $cmd.Parameters.Keys | Where-Object { $_ -like "$p*" } }
            $known | Should -Not -BeNullOrEmpty -Because "$Command has no parameter -$p"
        }
    }
}
