<#
.SYNOPSIS
    Lab test for the Maintenance Window Manager write actions. Creates, checks and removes
    a ZZTEST- collection with ZZTEST- maintenance windows on the lab site server.

.DESCRIPTION
    Copies the module to the site server and runs MaintenanceWindow.Lab.Remote.ps1 there as
    a one-time scheduled task under the lab account. A task is needed because the
    module queries the SMS Provider over CIM with the caller's credential, and a
    WinRM session cannot pass that credential on to a second hop. The script then
    reads the results, removes the task and the copy, prints one line per step, and
    exits 1 when a step failed. It refuses to start when a ZZTEST- package already
    exists. The lab credential comes from the file in -CredentialFile, which
    defines $LabCredential. The script never prints it.

.EXAMPLE
    .\Tests\Lab\Invoke-MaintenanceWindowLab.ps1
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification='Console harness.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingPlainTextForPassword', '', Justification='CredentialFile is a file path, not a secret.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseUsingScopeModifierInNewRunspaces', '', Justification='Each remote script block declares its values in a param block and receives them with -ArgumentList.')]
[CmdletBinding()]
param(
    [string]$CredentialFile = 'C:\projects\app-packager\docs\lab-credentials.local.ps1',
    [string]$ComputerName = 'CM01',
    [string]$SiteCode = 'MCM',
    [string]$Provider = 'CM01.CONTOSO.COM',
    [int]$TimeoutMinutes = 12
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
. $CredentialFile
if (-not $LabCredential) { throw "The credential file did not define `$LabCredential." }

$session = New-PSSession -ComputerName $ComputerName -Credential $LabCredential
$remoteRoot = $null
$taskName = 'ZZTEST-MWM-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$results = @()
try {
    $remoteRoot = Invoke-Command -Session $session -ScriptBlock {
        $p = Join-Path $env:SystemRoot ('Temp\ZZTEST-mwm-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        $p
    }
    Copy-Item -ToSession $session -Path (Join-Path $repo 'Module') -Destination $remoteRoot -Recurse
    Copy-Item -ToSession $session -Path (Join-Path $repo 'Lib') -Destination $remoteRoot -Recurse
    Copy-Item -ToSession $session -Path (Join-Path $PSScriptRoot 'MaintenanceWindow.Lab.Remote.ps1') -Destination $remoteRoot

    $plain = $LabCredential.GetNetworkCredential().Password
    $user = $LabCredential.UserName
    Invoke-Command -Session $session -ArgumentList $taskName, $remoteRoot, $SiteCode, $Provider, $user, $plain -ScriptBlock {
        param($task, $root, $site, $prov, $user, $pw)
        $wrapper = Join-Path $root 'run.ps1'
        @(
            '$ErrorActionPreference = ''Continue'''
            "`$r = & '$root\MaintenanceWindow.Lab.Remote.ps1' -ModuleRoot '$root' -SiteCode '$site' -Provider '$prov'"
            "`$r | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath '$root\results.json' -Encoding UTF8"
        ) | Set-Content -LiteralPath $wrapper -Encoding UTF8
        $action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $wrapper)
        Register-ScheduledTask -TaskName $task -Action $action -User $user -Password $pw -RunLevel Highest -Force | Out-Null
        Start-ScheduledTask -TaskName $task
    }

    # The wrapper writes results.json as its last step; wait for that file.
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    do {
        Start-Sleep -Seconds 5
        $done = Invoke-Command -Session $session -ArgumentList $remoteRoot -ScriptBlock { param($root) Test-Path -LiteralPath (Join-Path $root 'results.json') }
    } while (-not $done -and (Get-Date) -lt $deadline)
    if (-not $done) { Write-Warning 'The lab task wrote no result before the timeout.' }
    $json = Invoke-Command -Session $session -ArgumentList $remoteRoot -ScriptBlock {
        param($root)
        $f = Join-Path $root 'results.json'
        if (Test-Path -LiteralPath $f) { Get-Content -LiteralPath $f -Raw -Encoding UTF8 }
    }
    if ($json) { $results = @($json | ConvertFrom-Json) }
    else { $results = @([PSCustomObject]@{ Step = 'run'; Ok = $false; Detail = 'The lab task wrote no result.' }) }
}
finally {
    Invoke-Command -Session $session -ArgumentList $taskName, $remoteRoot -ScriptBlock {
        param($t, $p)
        if ($t -like 'ZZTEST-MWM-*') { Stop-ScheduledTask -TaskName $t -ErrorAction SilentlyContinue; Unregister-ScheduledTask -TaskName $t -Confirm:$false -ErrorAction SilentlyContinue }
        if ($p -like "$env:SystemRoot\Temp\ZZTEST-mwm-*") { Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue }
    }
    Remove-PSSession $session
}

foreach ($r in $results) { Write-Host ('{0} {1}  {2}' -f $(if ($r.Ok) { '[PASS]' } else { '[FAIL]' }), $r.Step, $r.Detail) }
$bad = @($results | Where-Object { -not $_.Ok })
Write-Host ('{0} steps, {1} failed' -f @($results).Count, $bad.Count)
if ($bad.Count -gt 0) { exit 1 }
