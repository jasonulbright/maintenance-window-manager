#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x source guard for button styles (no window is opened).
    A button style that sets a border must not draw it in the text color of the theme and must not be thicker than 1 px:
    - a style based on MahApps.Styles.Button.Square (or on the app's ThemedButtonStyle) sets BorderThickness 1 or less
      and a BorderBrush that is not ThemeForeground (Gray2 for a normal button, AccentBase for an accent button);
    - no style, at the top level, sets BorderBrush to ThemeForeground or BorderThickness above 1.
    The hover and focus triggers may use ThemeForeground, because they mark the state.
    Button text stays bold: no button style sets FontWeight below SemiBold.
#>

BeforeDiscovery {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:Files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File |
            Where-Object { $_.Extension -in '.xaml', '.ps1' } |
            Where-Object { $_.FullName.Substring($script:Root.Length) -notmatch '^[\\/]Lib[\\/]|[\\/](Tests|\.git|\.claude|node_modules|TestResults|docs|tools)[\\/]' } |
            Where-Object { (Get-Content -LiteralPath $_.FullName -Raw) -match 'TargetType="Button"' } |
            ForEach-Object { @{ Path = $_.FullName; Name = $_.FullName.Substring($script:Root.Length + 1) } })
}

Describe 'Button styles' {
    It 'finds the files with button styles' {
        $root = Split-Path -Path $PSScriptRoot -Parent
        @(Get-ChildItem -LiteralPath $root -Recurse -File -Filter MainWindow.xaml).Count | Should -BeGreaterThan 0
    }

    It 'draws no heavy or text-colored border in <Name>' -ForEach $script:Files {
        $text = Get-Content -LiteralPath $Path -Raw
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($m in [regex]::Matches($text, '(?s)<Style\s+x:Key="(?<key>[^"]+)"\s+TargetType="Button"(?<attrs>[^>/]*)(?<self>)>(?<body>.*?)</Style>')) {
            if ($m.Groups['self'].Value -eq '/') { continue }
            $key = $m.Groups['key'].Value
            $attrs = $m.Groups['attrs'].Value
            $top = ($m.Groups['body'].Value -split '<Style\.Triggers>')[0]
            $isAccent = $key -match 'Accent|Primary' -or $attrs -match 'Accent'
            $thick = [regex]::Match($top, 'Property="BorderThickness"\s+Value="(?<v>[^"]+)"')
            if ($thick.Success -and (($thick.Groups['v'].Value -split '[,\s]+' | ForEach-Object { [double]$_ } | Measure-Object -Maximum).Maximum -gt 1)) { $bad.Add(('{0}: BorderThickness {1}' -f $key, $thick.Groups['v'].Value)) }
            if ($top -match 'Property="BorderBrush"\s+Value="\{DynamicResource MahApps\.Brushes\.ThemeForeground\}"') { $bad.Add(('{0}: BorderBrush is ThemeForeground' -f $key)) }
            $base = $attrs -match 'MahApps\.Styles\.Button\.Square' -and $attrs -notmatch 'ThemedButtonStyle'
            if ($key -eq 'ThemedButtonStyle' -or $base) {
                if ($top -notmatch '(?m)^\s*<Setter Property="BorderBrush"') { if (-not $isAccent) { $bad.Add(('{0}: no BorderBrush (the MahApps default is heavy and text-colored)' -f $key)) } }
                if ($base -and $top -notmatch 'Property="BorderThickness"') { $bad.Add(('{0}: no BorderThickness (the MahApps default is 2)' -f $key)) }
            }
            $w = [regex]::Match($top, 'Property="FontWeight"\s+Value="(?<v>[^"]+)"')
            if ($w.Success -and $w.Groups['v'].Value -match '^(Normal|Regular|Light|Thin|ExtraLight|UltraLight|[1-5]\d\d)$') { $bad.Add(('{0}: FontWeight {1}' -f $key, $w.Groups['v'].Value)) }
        }
        $bad | Should -BeNullOrEmpty
    }
}
