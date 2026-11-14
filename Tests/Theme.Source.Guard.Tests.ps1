#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x source guards for theme use (no window is opened):
    1. No text uses the MahApps brushes Gray2 to Gray7 as a Foreground, because they stay under 4.5:1
       on one of the two themes. Muted text uses MahApps.Brushes.Gray.MouseOver (4.5:1 or more in both themes).
       Gray2 stays allowed as a BorderBrush (3:1 is the limit for a control border).
    2. Every dialog that the script builds from XAML text calls Set-DialogTheme in the same function,
       so that it gets the shared palette and the shared implicit styles.
#>

BeforeDiscovery {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:OwnFiles = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.xaml', '*.ps1', '*.psm1' |
            Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch'[\\/](Lib|Tests|\.git|\.claude|docs|tools|node_modules)[\\/]' } |
            ForEach-Object { @{ Path = $_.FullName; Name = $_.FullName.Substring($script:Root.Length + 1) } })
    $script:ScriptFiles = @($script:OwnFiles | Where-Object { $_.Name -like '*.ps1' -or $_.Name -like '*.psm1' })
}

Describe 'Muted and secondary text brushes' {
    It 'does not use Gray2 to Gray7 as a Foreground in <Name>' -ForEach $script:OwnFiles {
        $text = Get-Content -LiteralPath $Path -Raw
        $hits = [regex]::Matches($text, 'Foreground\s*=\s*"\{(?:DynamicResource|StaticResource)\s+MahApps\.Brushes\.Gray[2-7]\}"')
        $hits.Count | Should -Be 0 -Because 'use MahApps.Brushes.Gray.MouseOver for muted text'
        $setter = [regex]::Matches($text, '<Setter\s+Property\s*=\s*"Foreground"\s+Value\s*=\s*"\{(?:DynamicResource|StaticResource)\s+MahApps\.Brushes\.Gray[2-7]\}"')
        $setter.Count | Should -Be 0 -Because 'use MahApps.Brushes.Gray.MouseOver for muted text'
    }
}

Describe 'Dialog theme' {
    It 'applies Set-DialogTheme to every dialog built from XAML text in <Name>' -ForEach $script:ScriptFiles {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
        $funcs = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true))
        $loads = @($ast.FindAll({
                    param($n)
                    $n -is [System.Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Static -and
                    $n.Member.Value -eq 'Load' -and $n.Expression.Extent.Text -match 'XamlReader'
                }, $true))
        $missing = @()
        foreach ($l in $loads) {
            # A dialog is a load whose result goes into a variable named like dlg, dialog or win and whose text has a MetroWindow root.
            $assign = $l.Parent
            while ($assign -and $assign -isnot [System.Management.Automation.Language.AssignmentStatementAst]) { $assign = $assign.Parent }
            if (-not $assign) { continue }
            $var = $assign.Left.Extent.Text
            if ($var -notmatch '(?i)dlg|dialog|popup') { continue }
            $owner = $funcs | Where-Object { $_.Extent.StartOffset -le $l.Extent.StartOffset -and $_.Extent.EndOffset -ge $l.Extent.EndOffset } |
                Sort-Object { $_.Extent.EndOffset - $_.Extent.StartOffset } | Select-Object -First 1
            $scopeAst = if ($owner) { $owner } else { $ast }
            $themed = @($scopeAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Set-DialogTheme' -and $n.Extent.Text -match [regex]::Escape($var) }, $true))
            if ($themed.Count -eq 0) { $missing += ('{0} (line {1})' -f $var, $l.Extent.StartLineNumber) }
        }
        $missing | Should -BeNullOrEmpty -Because 'call Set-DialogTheme instead of ChangeTheme on the dialog'
    }
}
