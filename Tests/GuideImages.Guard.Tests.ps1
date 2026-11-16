#Requires -Modules Pester

<#
.SYNOPSIS
    Pester 5.x guard (headless, reads files only): every guide image under docs\images was produced by a capture run, and
    every image that the last capture run produced is in the guide. The capture run (D-21) writes
    tools\guide-capture\produced-images.json. A guide image that no check run produced is listed by name.
#>

BeforeAll {
    $script:Root = Split-Path -Path $PSScriptRoot -Parent
    $script:Produced = @((Get-Content -LiteralPath (Join-Path $script:Root 'tools\guide-capture\produced-images.json') -Raw | ConvertFrom-Json).Images)
    $script:Guide = @(Get-ChildItem -LiteralPath (Join-Path $script:Root 'docs\images') -Filter '*.png' | ForEach-Object { $_.Name })
}

Describe 'Guide images and the capture run' {
    It 'has a produced-images list with at least one image' {
        $script:Produced.Count | Should -BeGreaterThan 0
    }

    It 'has a producing capture run for every guide image' {
        $none = @($script:Guide | Where-Object { $script:Produced -notcontains $_ })
        $none | Should -BeNullOrEmpty -Because ('no capture run produced: ' + ($none -join ', '))
    }

    It 'has every produced image in the guide, so that a retired image leaves the list' {
        $gone = @($script:Produced | Where-Object { $script:Guide -notcontains $_ })
        $gone | Should -BeNullOrEmpty -Because ('produced but not in docs\images (run the capture with -AddNew once): ' + ($gone -join ', '))
    }
}
