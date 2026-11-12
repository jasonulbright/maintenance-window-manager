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
    [string]$OutDir,
    [switch]$AddNew
)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $here '..\..')).Path
# The lab names are fixed, so a capture from a worktree uses the same folder and task name as one from the main folder.
$app = 'maintenance-window-manager'
$images = Join-Path $repo 'docs\images'
. $CredentialScript
$prefsFile = 'MaintWindowMgr.prefs.json'
# Own folder and own task name per app, so that two captures in the lab never meet. The C:\guide folder itself stays.
$root = "C:\guide\$app"
$taskName = "GuideShots-$app"
$lockFile = 'C:\guide\capture.lock'
$haveLock = $false
$settings = @{ SiteCode = $SiteCode; SMSProvider = $SmsProvider } | ConvertTo-Json
$failed = New-Object System.Collections.Generic.List[string]
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('guide-capture-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $tmp | Out-Null
$session = $null
try {
    $session = New-PSSession -ComputerName $Target -Credential $LabCredential
    # One capture at a time per lab computer. New-Item fails when the lock file exists, so the capture waits.
    $lockEnd = (Get-Date).AddMinutes(60)
    do {
        try {
            Invoke-Command $session { param($l) New-Item -ItemType Directory 'C:\guide' -Force | Out-Null; New-Item -ItemType File $l -ErrorAction Stop | Out-Null } -ArgumentList $lockFile -ErrorAction Stop
            $haveLock = $true
        }
        catch {
            # Only an existing lock file means that another capture runs. Any other error stops here.
            if ($_.Exception.Message -notmatch 'already exists') { throw }
            Start-Sleep -Seconds 30
        }
    } until ($haveLock -or (Get-Date) -gt $lockEnd)
    if (-not $haveLock) { throw "Another capture holds $lockFile. This capture waited 60 minutes." }
    foreach ($run in @(@{ Steps = 'steps-guide.ps1'; Prefs = $false }, @{ Steps = 'steps-console-missing.ps1'; Prefs = $true })) {
        $dest = $root
        # Stop only the leftovers of an earlier capture: the capture runner and the copy of this app.
        Invoke-Command $session {
            param($d, $appName)
            Get-CimInstance Win32_Process -Filter "Name='powershell.exe'" | Where-Object { $_.ProcessId -ne $PID -and $_.CommandLine -like "*$d\*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
            Start-Sleep -Seconds 2
            if (Test-Path $d) { Remove-Item $d -Recurse -Force }
            New-Item -ItemType Directory $d, (Join-Path $d 'ui') -Force | Out-Null
        } -ArgumentList $dest, $app
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
}
finally {
    # Each cleanup runs on its own, so that a failed cleanup cannot replace the original error.
    if ($session) {
        if ($haveLock) {
            try { Invoke-Command $session { param($t) Unregister-ScheduledTask $t -Confirm:$false -ErrorAction SilentlyContinue } -ArgumentList $taskName } catch { Write-Warning "Cleanup of the scheduled task failed: $($_.Exception.Message)" }
            try { Invoke-Command $session { param($r) Remove-Item $r -Recurse -Force -ErrorAction SilentlyContinue } -ArgumentList $root } catch { Write-Warning "Cleanup of $root failed: $($_.Exception.Message)" }
            try { Invoke-Command $session { param($l) Remove-Item $l -Force -ErrorAction SilentlyContinue } -ArgumentList $lockFile } catch { Write-Warning "Removing $lockFile failed: $($_.Exception.Message)" }
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
