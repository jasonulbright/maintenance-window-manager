#Requires -Modules Pester
# Guard for the images of the README. It reads files and compares bytes. It opens no window.
# Every image that the README shows is under docs/images or is a copy of a docs/images file that
# tools/guide-capture/readme-images.json names. A mapped README image is byte-identical to its source.

BeforeDiscovery {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:Map = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'tools\guide-capture\readme-images.json') -Raw | ConvertFrom-Json
    $script:MapCases = @($script:Map.PSObject.Properties | ForEach-Object { @{ Readme = $_.Name; Source = $_.Value } })
}

BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:Map = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'tools\guide-capture\readme-images.json') -Raw | ConvertFrom-Json
    $readme = Get-Content -LiteralPath (Join-Path $script:RepoRoot 'README.md') -Raw
    $script:Images = @([regex]::Matches($readme, '!\[[^\]]*\]\(([^)\s]+)\)') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -notmatch '^https?://' })
}

Describe 'README images' {
    It 'finds at least one image in the README' {
        $script:Images.Count | Should -BeGreaterThan 0
    }

    It 'maps every README image that is not under docs/images' {
        $unmapped = @($script:Images | Where-Object { $_ -notmatch '^docs/images/' -and -not $script:Map.PSObject.Properties[$_] })
        $unmapped | Should -BeNullOrEmpty
    }

    It 'copies <Source> to <Readme>, byte for byte' -ForEach $script:MapCases {
        $source = Join-Path $script:RepoRoot ($Source -replace '/', '\')
        $target = Join-Path $script:RepoRoot ($Readme -replace '/', '\')
        $source | Should -Exist
        $target | Should -Exist
        (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash | Should -Be (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    }

    It 'maps only images that the README shows' {
        $extra = @($script:Map.PSObject.Properties.Name | Where-Object { $script:Images -notcontains $_ })
        $extra | Should -BeNullOrEmpty
    }
}
