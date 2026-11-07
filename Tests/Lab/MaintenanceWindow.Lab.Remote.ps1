# Runs ON the lab site server (Windows PowerShell 5.1). Started by Invoke-MaintenanceWindowLab.ps1.
#
# Creates two ZZTEST- device collections, runs the module's maintenance window actions
# on them (schedules, create, update, disable and enable in bulk, copy, template, CSV
# preview, remove, WhatIf), and removes everything it created, also when a step fails.
# It compares every other maintenance window with a snapshot taken before the first
# change. Only objects named ZZTEST-* change.
param(
    [Parameter(Mandatory)][string]$ModuleRoot,
    [Parameter(Mandatory)][string]$SiteCode,
    [Parameter(Mandatory)][string]$Provider
)

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[object]]::new()
function Add-Result { param([string]$Step, [bool]$Ok, [string]$Detail = '') $results.Add([PSCustomObject]@{ Step = $Step; Ok = $Ok; Detail = $Detail }) }

$tag = [guid]::NewGuid().ToString('N').Substring(0, 8)
$ns = "root\SMS\site_$SiteCode"
$tplDir = Join-Path $ModuleRoot 'tpl'

function Get-WindowSnapshot {
    # SMS_ServiceWindow is an embedded class: read the windows from each collection's settings
    # (a lazy property that Get-CimInstance -InputObject loads).
    $lines = foreach ($s in @(Get-CimInstance -Namespace $ns -ClassName SMS_CollectionSettings)) {
        $full = Get-CimInstance -InputObject $s
        foreach ($w in @($full.ServiceWindows)) {
            if (-not $w -or $w.Name -like 'ZZTEST-*') { continue }
            '{0}|{1}|{2}|{3}|{4}|{5}' -f $s.CollectionID, $w.ServiceWindowID, $w.Name, $w.ServiceWindowType, $w.IsEnabled, $w.Duration
        }
    }
    @($lines | Sort-Object)
}

try {
    Import-Module (Join-Path $ModuleRoot 'Module\MaintWindowMgrCommon.psd1') -Force -DisableNameChecking
    Initialize-Logging -LogPath (Join-Path $ModuleRoot 'lab.log')

    $others = @(Get-CimInstance -Namespace $ns -ClassName SMS_Collection -Filter "Name LIKE 'ZZTEST-%'")
    if ($others.Count -gt 0) { throw 'ZZTEST- objects already exist. Another test may be running.' }

    Add-Result 'connect' ([bool](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider)) 'Connect-CMSite'
    $snapshot = Get-WindowSnapshot

    $a = New-CMDeviceCollection -Name "ZZTEST-MWM-$tag-A" -LimitingCollectionName 'All Systems' -RefreshType 1
    $b = New-CMDeviceCollection -Name "ZZTEST-MWM-$tag-B" -LimitingCollectionName 'All Systems' -RefreshType 1
    Add-Result 'create two ZZTEST- collections' ([bool]($a.CollectionID -and $b.CollectionID)) "$($a.CollectionID) $($b.CollectionID)"

    $start = (Get-Date).Date.AddDays(1).AddHours(2)

    # --- schedules -------------------------------------------------------------------------------------
    $sWeekly = New-WindowSchedule -RecurrenceType Weekly -StartTime $start -DurationHours 3 -DayOfWeek Friday
    Add-Result 'weekly schedule' ([bool]$sWeekly) ''
    $sMixed = New-WindowSchedule -RecurrenceType Daily -StartTime $start -DurationHours 1 -DurationMinutes 30
    Add-Result 'daily schedule of 1 hour 30 minutes' ([bool]$sMixed) ''
    $sMonthly = New-WindowSchedule -RecurrenceType MonthlyByWeekday -StartTime $start -DurationHours 4 -DayOfWeek Tuesday -WeekOrder Second
    Add-Result 'monthly weekday schedule' ([bool]$sMonthly) ''
    $sDate = New-WindowSchedule -RecurrenceType MonthlyByDate -StartTime $start -DurationHours 2 -DayOfMonth 15
    Add-Result 'monthly date schedule' ([bool]$sDate) ''
    $sOnce = New-WindowSchedule -RecurrenceType OneTime -StartTime $start -DurationHours 1
    Add-Result 'one-time schedule' ([bool]$sOnce) ''
    $sPatch = New-PatchTuesdaySchedule -OffsetDays 0 -StartHour 3 -DurationHours 5
    Add-Result 'Patch Tuesday schedule' ([bool]$sPatch) ''

    # --- create ----------------------------------------------------------------------------------------
    $w1 = New-ManagedMaintenanceWindow -CollectionId $a.CollectionID -Name 'ZZTEST-weekly' -Schedule $sWeekly -ApplyTo SoftwareUpdatesOnly
    Add-Result 'create weekly window' ([bool]$w1) ''
    $w2 = New-ManagedMaintenanceWindow -CollectionId $a.CollectionID -Name 'ZZTEST-mixed' -Schedule $sMixed -IsEnabled $false
    Add-Result 'create window of 1h 30m' ([bool]$w2) ''
    $got = @(Get-CollectionMaintenanceWindows -CollectionId $a.CollectionID)
    Add-Result 'both windows are listed' ($got.Count -eq 2) "count=$($got.Count)"
    $mixed = $got | Where-Object Name -eq 'ZZTEST-mixed'
    Add-Result 'the 1h 30m window has a duration of 90 minutes' ($mixed -and [int]$mixed.Duration -eq 90) "Duration=$($mixed.Duration)"
    $dup = New-ManagedMaintenanceWindow -CollectionId $a.CollectionID -Name 'ZZTEST-weekly' -Schedule $sWeekly
    Add-Result 'a second window with the same name is refused' (-not $dup) "result=[$([bool]$dup)]"
    $whatif = New-ManagedMaintenanceWindow -CollectionId $b.CollectionID -Name 'ZZTEST-whatif' -Schedule $sOnce -WhatIf
    Add-Result 'WhatIf creates nothing' (-not $whatif -and @(Get-CollectionMaintenanceWindows -CollectionId $b.CollectionID).Count -eq 0) ''

    # --- all windows view ------------------------------------------------------------------------------
    $all = @(Get-AllMaintenanceWindows | Where-Object CollectionID -in $a.CollectionID, $b.CollectionID)
    Add-Result 'schedule text and type are filled' (@($all | Where-Object { $_.Schedule -and $_.Type -and $_.Recurrence }).Count -eq $all.Count) (($all | ForEach-Object { '{0}: {1}' -f $_.WindowName, $_.Schedule }) -join '; ')

    # --- update, bulk, copy ---------------------------------------------------------------------------
    $okSet = Set-ManagedMaintenanceWindow -CollectionId $a.CollectionID -MaintenanceWindowName 'ZZTEST-weekly' -IsEnabled $false -ApplyTo TaskSequencesOnly
    Add-Result 'update window' ($okSet -eq $true) ''
    $weekly = Get-CollectionMaintenanceWindows -CollectionId $a.CollectionID | Where-Object Name -eq 'ZZTEST-weekly'
    Add-Result 'the update took effect' ($weekly -and -not $weekly.IsEnabled -and [int]$weekly.ServiceWindowType -eq 5) "IsEnabled=$($weekly.IsEnabled) Type=$($weekly.ServiceWindowType)"

    $recs = @(Get-AllMaintenanceWindows | Where-Object { $_.CollectionID -eq $a.CollectionID })
    $bulk = @(Set-BulkWindowEnabled -WindowRecords $recs -Enabled $true)
    Add-Result 'bulk enable' ($bulk.Count -eq 2 -and @($bulk | Where-Object { -not $_.Success }).Count -eq 0) (($bulk | ForEach-Object Error) -join ';')
    $enabled = @(Get-CollectionMaintenanceWindows -CollectionId $a.CollectionID | Where-Object IsEnabled).Count
    Add-Result 'both windows are enabled' ($enabled -eq 2) "enabled=$enabled"
    $bulkWhatIf = @(Set-BulkWindowEnabled -WindowRecords $recs -Enabled $false -WhatIf)
    Add-Result 'bulk WhatIf changes nothing' ($bulkWhatIf.Count -eq 0 -and @(Get-CollectionMaintenanceWindows -CollectionId $a.CollectionID | Where-Object IsEnabled).Count -eq 2) ''

    $copy = @(Copy-MaintenanceWindowToCollections -SourceCollectionId $a.CollectionID -SourceWindowName 'ZZTEST-mixed' -TargetCollectionIds @($b.CollectionID))
    Add-Result 'copy window to another collection' ($copy.Count -eq 1 -and $copy[0].Success) ($copy | ForEach-Object Error)
    $copied = Get-CollectionMaintenanceWindows -CollectionId $b.CollectionID | Where-Object Name -eq 'ZZTEST-mixed'
    Add-Result 'the copy keeps the 90 minute duration' ($copied -and [int]$copied.Duration -eq 90) "Duration=$($copied.Duration)"

    # --- templates, CSV --------------------------------------------------------------------------------
    $tp = Save-WindowTemplate -TemplatesPath $tplDir -Name 'ZZTEST-template' -RecurrenceType Weekly -DurationHours 2 -DurationMinutes 15 -DayOfWeek Saturday
    Add-Result 'save template' (Test-Path -LiteralPath $tp) ''
    $tpl = @(Get-WindowTemplates -TemplatesPath $tplDir)[0]
    $applied = @(New-WindowFromTemplate -Template $tpl -TargetCollections @($b.CollectionID) -WindowNameOverride 'ZZTEST-from-template')
    Add-Result 'apply the template to a collection' ($applied.Count -eq 1 -and $applied[0].Success) ($applied | ForEach-Object Error)
    $fromTpl = Get-CollectionMaintenanceWindows -CollectionId $b.CollectionID | Where-Object Name -eq 'ZZTEST-from-template'
    Add-Result 'the template window has a duration of 135 minutes' ($fromTpl -and [int]$fromTpl.Duration -eq 135) "Duration=$($fromTpl.Duration)"
    Add-Result 'remove the template file' ((Remove-WindowTemplate -FilePath $tp) -eq $true) ''

    $csv = Join-Path $ModuleRoot 'w.csv'
    Set-Content -LiteralPath $csv -Value @('CollectionID,WindowName,StartHour', "$($a.CollectionID),ZZTEST-csv,3", 'C1,bad,99')
    $pv = @(Import-MaintenanceWindowsCsv -CsvPath $csv)
    Add-Result 'CSV preview marks the bad row' ($pv.Count -eq 2 -and $pv[0].Valid -and -not $pv[1].Valid) ($pv | ForEach-Object { "Valid=$($_.Valid)" })

    # --- remove ----------------------------------------------------------------------------------------
    $rw = Remove-ManagedMaintenanceWindow -CollectionId $a.CollectionID -MaintenanceWindowName 'ZZTEST-mixed' -WhatIf
    Add-Result 'remove under WhatIf keeps the window' ($rw -eq $false -and [bool](Get-CollectionMaintenanceWindows -CollectionId $a.CollectionID | Where-Object Name -eq 'ZZTEST-mixed')) ''
    foreach ($pair in @(@($a.CollectionID, 'ZZTEST-weekly'), @($a.CollectionID, 'ZZTEST-mixed'), @($b.CollectionID, 'ZZTEST-mixed'), @($b.CollectionID, 'ZZTEST-from-template'))) {
        $r = Remove-ManagedMaintenanceWindow -CollectionId $pair[0] -MaintenanceWindowName $pair[1]
        Add-Result "remove window $($pair[1]) from $($pair[0])" ($r -eq $true) ''
    }
    $gone = Remove-ManagedMaintenanceWindow -CollectionId $a.CollectionID -MaintenanceWindowName 'ZZTEST-no-such-window'
    Add-Result 'remove of an unknown window returns false' ($gone -eq $false) "result=$gone"
}
catch {
    Add-Result 'run' $false $_.Exception.Message
}
finally {
    try {
        if (-not (Test-CMConnection)) { [void](Connect-CMSite -SiteCode $SiteCode -SMSProvider $Provider) }
        foreach ($coll in @(Get-CMDeviceCollection -Name "ZZTEST-MWM-$tag-*" -ErrorAction SilentlyContinue)) {
            if ($coll.Name -notlike "ZZTEST-MWM-$tag-*") { continue }
            foreach ($w in @(Get-CMMaintenanceWindow -CollectionId $coll.CollectionID -ErrorAction SilentlyContinue)) {
                try { Remove-CMMaintenanceWindow -CollectionId $coll.CollectionID -MaintenanceWindowName $w.Name -Force -ErrorAction Stop } catch { $null = $_ }
            }
            try { Remove-CMDeviceCollection -CollectionId $coll.CollectionID -Force -ErrorAction Stop } catch { Add-Result "cleanup collection $($coll.Name)" $false $_.Exception.Message }
        }
    } catch { Add-Result 'cleanup' $false $_.Exception.Message }
    try {
        $left = @(Get-CimInstance -Namespace $ns -ClassName SMS_Collection -Filter "Name LIKE 'ZZTEST-MWM-$tag-%'")
        Add-Result 'no ZZTEST- collection of this run remains' ($left.Count -eq 0) ((@($left | ForEach-Object Name)) -join ',')
        $after = @(Get-WindowSnapshot)
        Add-Result 'other maintenance windows are unchanged' ((@($snapshot) -join "`n") -ceq ($after -join "`n")) ("$(@($snapshot).Count) windows before, $($after.Count) after")
    } catch { Add-Result 'verify cleanup' $false $_.Exception.Message }
}

$results
