<#
.SYNOPSIS
    Regenerates every image of the user guide with one command.

.DESCRIPTION
    Copies this app to a lab computer, runs the capture steps in an interactive
    session of a lab account (scheduled task, interactive logon), copies the PNG
    files back, and removes the task and the copy. Only the images that already
    exist in docs\images are replaced. A failed step stops the command before any
    image is replaced.

    The lab computer needs a Configuration Manager console, a reachable SMS
    Provider and PowerShell remoting for the account in the credential script.
    Read the steps before you run them: they open dialogs of buttons that change
    the site, and they answer No or Cancel. They never confirm a change.

.PARAMETER CredentialScript
    A script that sets $LabCredential (a PSCredential). Keep it outside the repository.

.PARAMETER Target
    The lab computer name or address.

.PARAMETER LogonUser
    The domain account that has the interactive session on the lab computer.

.PARAMETER SiteCode
    The site code that the capture types in Options.

.PARAMETER SmsProvider
    The SMS Provider name that the capture types in Options.

.PARAMETER OutDir
    A folder that receives a copy of every captured image, also the ones that are not in docs\images. Use it to look at a new state.

.PARAMETER AddNew
    Also copy images that are not in docs\images yet. Use it for the first capture of a new page.

.EXAMPLE
    .\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript C:\lab\credentials.local.ps1 -Target CLIENT01 -LogonUser contoso\labadmin -SmsProvider cm01.contoso.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CredentialScript,
    [Parameter(Mandatory)][string]$Target,
    [Parameter(Mandatory)][string]$LogonUser,
    [string]$SiteCode = 'MCM',
    [Parameter(Mandatory)][string]$SmsProvider,
    [int]$Minutes = 15,
    [int]$LockWaitMinutes = 90,
    [string]$OutDir,
    [switch]$AddNew
)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
# Stops the processes that earlier runs of this tool recorded in the ledger on the lab VM (process id and start time), and nothing else.
# No process is stopped by a command-line pattern: another session may own a process that matches (vault: testing-policy, 'Process rules on a lab VM').
$stopRecorded = @'
param($ledger)
if (Test-Path -LiteralPath $ledger) {
    foreach ($line in @(Get-Content -LiteralPath $ledger)) {
        $id, $ticks = $line -split '\|', 2
        $p = Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue
        if ($p -and $p.Id -ne $PID -and $p.StartTime.Ticks -eq [long]$ticks) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
    }
    Remove-Item -LiteralPath $ledger -Force -ErrorAction SilentlyContinue
}
'@
$repo = (Resolve-Path (Join-Path $here '..\..')).Path
# The lab names are fixed, so a capture from a worktree uses the same folder and task name as one from the main folder.
$app = 'maintenance-window-manager'
$images = Join-Path $repo 'docs\images'
. $CredentialScript
$prefsFile = 'MaintWindowMgr.prefs.json'
# Own folder and own task name per app, so that two captures in the lab never meet. The C:\guide folder itself stays.
$root = "C:\guide\$app"
$taskName = "GuideShots-$app"
$lockPath = 'C:\guide\capture.lock'
$lockText = '{0}|{1}|{2}' -f $app, $PID, (Get-Date).ToString('o')
$haveLock = $false
$settings = @{ SiteCode = $SiteCode; SMSProvider = $SmsProvider } | ConvertTo-Json
$failed = New-Object System.Collections.Generic.List[string]
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('guide-capture-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
$session = $null
try {
    $session = New-PSSession -ComputerName $Target -Credential $LabCredential
    # The lock (D-16): C:\guide\capture.lock, text owner|pid|start. New-Item fails with an IOException when the file exists: busy.
    # A lock is stale only when its last write is older than 30 minutes AND its owner pid runs neither on the lab computer nor on
    # this computer. The lock of a live owner is never removed. A lock in a format this tool does not know blocks and is reported.
    # A removed lock is reported with a warning. The holder refreshes the lock in its poll loop.
    $lockEnd = (Get-Date).AddMinutes($LockWaitMinutes)
    $unknownReported = $false
    while (-not $haveLock) {
        $got = Invoke-Command $session {
            param($path, $text)
            New-Item -ItemType Directory -Path (Split-Path -Path $path -Parent) -Force | Out-Null
            try { $null = New-Item -ItemType File -Path $path -Value $text -ErrorAction Stop; 'taken' }
            catch [System.IO.IOException] {
                $item = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                if (-not $item) { 'gone' }
                else {
                    $held = ([string](Get-Content -LiteralPath $path -Raw -ErrorAction SilentlyContinue)).Trim()
                    if ($held -notmatch '^[^|\s][^|]*\|\d+\|\d{4}-\d{2}-\d{2}T') { 'unknown: ' + $held }
                    else { 'held: {0}|{1}|{2}' -f [int]((Get-Date) - $item.LastWriteTime).TotalMinutes, [bool](Get-Process -Id ([int]($held -split '\|')[1]) -ErrorAction SilentlyContinue), $held }
                }
            }
        } -ArgumentList $lockPath, $lockText
        if ($got -eq 'taken') { $haveLock = $true; continue }
        if ($got -eq 'gone') { continue }
        if ($got -like 'unknown:*') {
            if (-not $unknownReported) { Write-Warning "The lab lock $lockPath has a format this tool does not know and blocks the capture ($got). Waiting; the lock is not removed."; $unknownReported = $true }
        }
        elseif ($got -like 'held:*') {
            $age, $aliveLab, $heldText = $got.Substring(6) -split '\|', 3
            $ownerPid = [int]($heldText -split '\|')[1]
            if ([int]$age -gt 30 -and $aliveLab -ne 'True' -and -not (Get-Process -Id $ownerPid -ErrorAction SilentlyContinue)) {
                $removed = Invoke-Command $session { param($path, $text) if (([string](Get-Content -LiteralPath $path -Raw -ErrorAction SilentlyContinue)).Trim() -eq $text) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue; $true } else { $false } } -ArgumentList $lockPath, $heldText
                if ($removed) { Write-Warning "The stale lab lock $lockPath was removed ($heldText; last write $age minutes ago; owner pid $ownerPid is not running)."; continue }
            }
        }
        if ((Get-Date) -gt $lockEnd) { throw "Another capture holds $lockPath. The wait of $LockWaitMinutes minutes is over." }
        Start-Sleep -Seconds 15
    }
    foreach ($run in @(@{ Steps = 'steps-guide.ps1'; Prefs = $false }, @{ Steps = 'steps-console-missing.ps1'; Prefs = $true })) {
        $dest = $root
        # Stop only the processes that an earlier capture recorded in its ledger, by process id.
        Invoke-Command $session {
            param($d)
            & ([scriptblock]::Create($using:stopRecorded)) (Join-Path $d 'ui\run.pids')
            Start-Sleep -Seconds 2
            if (Test-Path $d) { Remove-Item $d -Recurse -Force }
            New-Item -ItemType Directory $d, (Join-Path $d 'ui') -Force | Out-Null
        } -ArgumentList $dest
        foreach ($item in Get-ChildItem -LiteralPath $repo -Force | Where-Object { $_.Name -notin '.git', '.claude', '.github', '.vs', '.vscode', '.idea', 'node_modules', 'TestOutput', 'Logs', 'History', 'Reports', 'screenshots', 'TestResults', 'docs', 'Tests', 'tools' -and $_.Name -notlike '*.prefs.json' -and $_.Name -notlike '*.windowstate.json' -and $_.Name -notlike '*.lastscan.json' -and $_.Name -notlike '*.suppressions.json' -and $_.Name -notlike '*.local.*' -and $_.Name -notlike '*.png' }) {
            Copy-Item -LiteralPath $item.FullName -Destination "$dest\$($item.Name)" -ToSession $session -Recurse -Force
        }
        Copy-Item (Join-Path $here 'guilib.ps1') (Join-Path $root 'ui\guilib.ps1') -ToSession $session -Force
        Copy-Item (Join-Path $here 'main.ps1') (Join-Path $root 'ui\main.ps1') -ToSession $session -Force
        Copy-Item (Join-Path $here $run.Steps) (Join-Path $root 'ui\steps.ps1') -ToSession $session -Force
        Invoke-Command $session { param($j, $r) Set-Content (Join-Path $r 'ui\capture.json') $j -Encoding UTF8 } -ArgumentList $settings, $root
        Invoke-Command $session -ScriptBlock {
            param($a, $user, $json, $prefsName, $r, $task)
            if ($json) { Set-Content (Join-Path $r $prefsName) $json -Encoding UTF8 }
            Remove-Item (Join-Path $r 'ui\out') -Recurse -Force -ErrorAction SilentlyContinue
            $act = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$r\ui\main.ps1`" -GuideRoot `"$r`""
            $pr = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Highest
            Register-ScheduledTask -TaskName $task -Action $act -Principal $pr -Force | Out-Null
            Start-ScheduledTask $task
        } -ArgumentList $app, $LogonUser, $(if ($run.Prefs) { $settings } else { $null }), $prefsFile, $root, $taskName
        $end = (Get-Date).AddMinutes($Minutes)
        $done = $false
        do {
            Start-Sleep 5
            Invoke-Command $session { param($p) (Get-Item -LiteralPath $p -ErrorAction SilentlyContinue).LastWriteTime = Get-Date } -ArgumentList $lockPath
            $state = Invoke-Command $session { param($r, $t) [pscustomobject]@{ Done = (Test-Path (Join-Path $r 'ui\out\finished.flag')); Task = (Get-ScheduledTask $t -ErrorAction SilentlyContinue).State; Result = (Get-ScheduledTaskInfo $t -ErrorAction SilentlyContinue).LastTaskResult } } -ArgumentList $root, $taskName
            $done = $state.Done
            # The task is not running any more, no finished flag exists and the last result is an error: stop now.
            if (-not $done -and $state.Task -eq 'Ready' -and $state.Result -notin 0, 267009, 267011) { throw ("The capture task stopped with result {0}: {1}" -f $state.Result, $run.Steps) }
        } while (-not $done -and (Get-Date) -lt $end)
        if (-not $done) { throw "The capture did not finish: $($run.Steps)" }
        $out = Join-Path $tmp ([IO.Path]::GetFileNameWithoutExtension($run.Steps))
        New-Item -ItemType Directory $out | Out-Null
        Copy-Item (Join-Path $root 'ui\out\*') $out -FromSession $session -Recurse -Force
        # The flag file holds inished or ailed: <reason>. A failure is reported even when the log could not be written.
        $flagText = (Get-Content (Join-Path $out 'finished.flag') -Raw -ErrorAction SilentlyContinue)
        if ($flagText -and $flagText.Trim() -like 'failed:*') { $failed.Add($flagText.Trim()) }
        $log = Join-Path $out 'run.log'
        if (-not (Test-Path $log)) { throw "The capture wrote no run.log: $($run.Steps)" }
        Get-Content $log | Where-Object { $_ -match '^\d\d:\d\d:\d\d\s+ERROR' } | ForEach-Object { Write-Warning $_; $failed.Add($_) }
        Get-Content $log | Where-Object { $_ -match 'TOOLTIP still shown' } | ForEach-Object { Write-Warning $_ }
    }
    # A failed step stops the capture before any guide image is replaced.
    if ($failed.Count) { throw ("The capture failed. No image was replaced.`n" + ($failed -join "`n")) }
    $captured = @(Get-ChildItem $tmp -Recurse -Filter *.png)
    $capturedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $duplicate = New-Object System.Collections.Generic.List[string]
    foreach ($png in $captured) { if (-not $capturedNames.Add($png.Name)) { $duplicate.Add($png.Name) } }
    if ($duplicate.Count) { throw ("Two runs produced the same image. No image was replaced.`n" + ($duplicate -join "`n")) }
    $missing = @(Get-ChildItem $images -Filter *.png | Where-Object { -not $capturedNames.Contains($_.Name) } | ForEach-Object { $_.Name })
    if ($missing.Count) { throw ("The capture did not produce these guide images. No image was replaced.`n" + ($missing -join "`n")) }
    foreach ($png in $captured) {
        if ($AddNew -or (Test-Path (Join-Path $images $png.Name))) { Copy-Item $png.FullName (Join-Path $images $png.Name) -Force }
    }
    # One pass is the check run and the capture run (D-21): the steps sample the colors of the windows they shoot, and the
    # guide image test reads the manifest of the images that this run produced.
    $sampleCount = 0; $warnCount = 0
    foreach ($runLog in @(Get-ChildItem $tmp -Recurse -Filter run.log)) {
        $lines = @(Get-Content $runLog.FullName)
        $sampleCount += @($lines | Where-Object { $_ -match '\sSAMPLE\s' }).Count
        $warnCount += @($lines | Where-Object { $_ -match '\sWARN\s' }).Count
    }
    $names = @($captured | ForEach-Object { $_.Name } | Sort-Object -Unique)
    $manifest = [ordered]@{ Count = $names.Count; Images = $names } | ConvertTo-Json
    [IO.File]::WriteAllText((Join-Path $here 'produced-images.json'), ($manifest -replace "`r?`n", "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    Write-Output ('Capture verdict: PASS. Images produced {0}, probe samples {1}, warnings {2}, errors 0.' -f $names.Count, $sampleCount, $warnCount)
}
finally {
    # Each cleanup runs on its own, so that a failed cleanup cannot replace the original error.
    if ($session) {
        if ($haveLock) {
            try { Invoke-Command $session { param($t, $r) Stop-ScheduledTask $t -ErrorAction SilentlyContinue; & ([scriptblock]::Create($using:stopRecorded)) (Join-Path $r 'ui\run.pids') } -ArgumentList $taskName, $root } catch { Write-Warning "Stopping the capture processes failed: $($_.Exception.Message)" }
            try { Invoke-Command $session { param($t) Unregister-ScheduledTask $t -Confirm:$false -ErrorAction SilentlyContinue } -ArgumentList $taskName } catch { Write-Warning "Cleanup of the scheduled task failed: $($_.Exception.Message)" }
            try { Invoke-Command $session { param($r) Remove-Item $r -Recurse -Force -ErrorAction SilentlyContinue } -ArgumentList $root } catch { Write-Warning "Cleanup of $root failed: $($_.Exception.Message)" }
            try { Invoke-Command $session { param($path, $text) if ((Test-Path -LiteralPath $path) -and (([string](Get-Content -LiteralPath $path -Raw -ErrorAction SilentlyContinue)).Trim() -eq $text)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } } -ArgumentList $lockPath, $lockText } catch { Write-Warning "Removing $lockPath failed: $($_.Exception.Message)" }
        }
        try { Remove-PSSession $session } catch { Write-Warning "Closing the session failed: $($_.Exception.Message)" }
    }
    # -OutDir receives every captured image and every run log, also when a step failed or threw.
    if ($OutDir) {
        try {
            New-Item -ItemType Directory $OutDir -Force | Out-Null
            foreach ($png in @(Get-ChildItem $tmp -Recurse -Filter *.png)) { Copy-Item $png.FullName (Join-Path $OutDir $png.Name) -Force }
            foreach ($runLog in @(Get-ChildItem $tmp -Recurse -Filter run.log)) { Copy-Item $runLog.FullName (Join-Path $OutDir ($runLog.Directory.Name + '.run.log')) -Force }
        } catch { Write-Warning "Copy to $OutDir failed: $($_.Exception.Message)" }
    }
    try { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue } catch { Write-Warning "Cleanup of $tmp failed: $($_.Exception.Message)" }
}
