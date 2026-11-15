<#
.SYNOPSIS
    Copies the guide images that the README shows over their README paths.

.DESCRIPTION
    readme-images.json maps each README image path to the file of docs\images that it is a copy of.
    Run this script after Invoke-GuideCapture.ps1, so the README shows the new images.
    The script fails when a source image is missing.

.PARAMETER RepoRoot
    The repository folder. Default: the repository that holds this script.

.EXAMPLE
    .\tools\guide-capture\Sync-ReadmeImages.ps1
#>
[CmdletBinding()]
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
)
$ErrorActionPreference = 'Stop'
$map = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'readme-images.json') -Raw | ConvertFrom-Json
foreach ($entry in $map.PSObject.Properties) {
    $source = Join-Path $RepoRoot ($entry.Value -replace '/', '\')
    $target = Join-Path $RepoRoot ($entry.Name -replace '/', '\')
    if (-not (Test-Path -LiteralPath $source)) { throw ('The source image is missing: {0}' -f $entry.Value) }
    Copy-Item -LiteralPath $source -Destination $target -Force
    Write-Host ('{0} <- {1}' -f $entry.Name, $entry.Value)
}
