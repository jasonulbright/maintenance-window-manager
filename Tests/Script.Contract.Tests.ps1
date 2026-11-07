#Requires -Modules Pester
# The start script and the module call functions of the app module and of
# SuiteCommon by name. A call with a parameter that the function does not have
# fails only when the user clicks the button. This test checks every call with
# named parameters against the real function signature.

BeforeAll {
    . "$PSScriptRoot\TestHelpers.ps1"
    Import-AppModule
    # SuiteCommon parameter types name MahApps classes; load them so the signatures resolve.
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase -ErrorAction SilentlyContinue
    foreach ($dll in 'Microsoft.Xaml.Behaviors.dll', 'ControlzEx.dll', 'MahApps.Metro.dll') { [void][System.Reflection.Assembly]::LoadFrom((Join-Path $script:AppRoot "Lib\$dll")) }
    Import-Module (Join-Path $script:AppRoot 'Lib\SuiteCommon\SuiteCommon.psd1') -Global -DisableNameChecking -ErrorAction SilentlyContinue
    $script:Files = @(
        Get-ChildItem -LiteralPath $script:AppRoot -Filter 'start-*.ps1' -File
        Get-ChildItem -LiteralPath (Join-Path $script:AppRoot 'Module') -Filter '*.psm1' -File
    )
    $script:Problems = New-Object System.Collections.Generic.List[string]
    $script:Checked = 0
    $own = @{}
    foreach ($f in $script:Files) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
        foreach ($d in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $own[$d.Name] = $d }
    }
    foreach ($f in $script:Files) {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
        foreach ($c in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
            $name = $c.GetCommandName()
            if (-not $name) { continue }
            $named = @($c.CommandElements | Where-Object { $_ -is [System.Management.Automation.Language.CommandParameterAst] } | ForEach-Object { $_.ParameterName })
            if ($named.Count -eq 0) { continue }
            $params = $null
            if ($own.ContainsKey($name)) {
                # A function defined in the script or the module: read its parameter block from the syntax tree.
                $def = $own[$name]
                $block = $def.Body.ParamBlock
                if (-not $block -and $def.Parameters) { $params = @($def.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath }) }
                elseif ($block) {
                    $params = @($block.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
                    foreach ($p in $block.Parameters) { foreach ($a in $p.Attributes) { if ($a.TypeName.Name -eq 'Alias') { $params += @($a.PositionalArguments | ForEach-Object { $_.Value }) } } }
                }
                else { $params = @() }
                $common = 'Verbose', 'Debug', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'WhatIf', 'Confirm'
                $isAdvanced = [bool]($def.Body.ParamBlock -and ($def.Body.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }))
                if ($isAdvanced) { $params += $common }
            }
            else {
                $cmd = Get-Command $name -ErrorAction SilentlyContinue
                # Only functions of SuiteCommon are checked here. Cmdlets are checked by the ConfigMgr contract test.
                if (-not $cmd -or $cmd.CommandType -ne 'Function' -or $cmd.Source -ne 'SuiteCommon') { continue }
                $params = @($cmd.Parameters.Values | ForEach-Object { $_.Name; $_.Aliases })
            }
            foreach ($p in $named) {
                $script:Checked++
                $hit = $params | Where-Object { $_ -ieq $p }
                if (-not $hit) { $hit = $params | Where-Object { $_ -like "$p*" } }
                if (-not $hit) { [void]$script:Problems.Add(('{0}:{1} {2} has no parameter -{3}' -f $f.Name, $c.Extent.StartLineNumber, $name, $p)) }
            }
        }
    }
}

Describe 'Calls into app functions and SuiteCommon' {
    It 'checks a useful number of named parameters' {
        $script:Checked | Should -BeGreaterThan 100
    }

    It 'uses only parameters that the called function has' {
        $script:Problems | Should -BeNullOrEmpty
    }
}
