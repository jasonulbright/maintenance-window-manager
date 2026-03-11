<#
.SYNOPSIS
    Core module for MECM Maintenance Window Manager.

.DESCRIPTION
    Import this module to get:
      - Structured logging (Initialize-Logging, Write-Log)
      - CM site connection management (Connect-CMSite, Disconnect-CMSite, Test-CMConnection)
      - Maintenance window data retrieval and CRUD operations
      - Schedule creation helpers (including Patch Tuesday)
      - Template management (load, save, apply)
      - Bulk operations (import CSV, copy, enable/disable)
      - Export to CSV and HTML

.EXAMPLE
    Import-Module "$PSScriptRoot\Module\MaintWindowMgrCommon.psd1" -Force
    Initialize-Logging -LogPath "C:\temp\mwm.log"
    Connect-CMSite -SiteCode 'MCM' -SMSProvider 'sccm01.contoso.com'
#>

# ---------------------------------------------------------------------------
# Module-scoped state
# ---------------------------------------------------------------------------

$script:__MWLogPath            = $null
$script:OriginalLocation       = $null
$script:ConnectedSiteCode      = $null
$script:ConnectedSMSProvider   = $null

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

function Initialize-Logging {
    param([string]$LogPath)

    $script:__MWLogPath = $LogPath

    if ($LogPath) {
        $parentDir = Split-Path -Path $LogPath -Parent
        if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
            New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
        }

        $header = "[{0}] [INFO ] === Log initialized ===" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Set-Content -LiteralPath $LogPath -Value $header -Encoding UTF8
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped, severity-tagged log message.
    #>
    param(
        [AllowEmptyString()]
        [Parameter(Mandatory, Position = 0)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO',

        [switch]$Quiet
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $formatted = "[{0}] [{1,-5}] {2}" -f $timestamp, $Level, $Message

    if (-not $Quiet) {
        Write-Host $formatted
        if ($Level -eq 'ERROR') {
            $host.UI.WriteErrorLine($formatted)
        }
    }

    if ($script:__MWLogPath) {
        Add-Content -LiteralPath $script:__MWLogPath -Value $formatted -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# CM Connection
# ---------------------------------------------------------------------------

function Connect-CMSite {
    <#
    .SYNOPSIS
        Imports the ConfigurationManager module, creates a PSDrive, and sets location.
    .DESCRIPTION
        Returns $true on success, $false on failure.
    #>
    param(
        [Parameter(Mandatory)][string]$SiteCode,
        [Parameter(Mandatory)][string]$SMSProvider
    )

    $script:OriginalLocation = Get-Location

    if (-not (Get-Module ConfigurationManager -ErrorAction SilentlyContinue)) {
        $cmModulePath = $null
        if ($env:SMS_ADMIN_UI_PATH) {
            $cmModulePath = Join-Path $env:SMS_ADMIN_UI_PATH '..\ConfigurationManager.psd1'
        }

        if (-not $cmModulePath -or -not (Test-Path -LiteralPath $cmModulePath)) {
            Write-Log "ConfigurationManager module not found. Ensure the CM console is installed." -Level ERROR
            return $false
        }

        try {
            Import-Module $cmModulePath -ErrorAction Stop
            Write-Log "Imported ConfigurationManager module"
        }
        catch {
            Write-Log "Failed to import ConfigurationManager module: $_" -Level ERROR
            return $false
        }
    }

    if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
        try {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SMSProvider -ErrorAction Stop | Out-Null
            Write-Log "Created PSDrive for site $SiteCode"
        }
        catch {
            Write-Log "Failed to create PSDrive for site $SiteCode : $_" -Level ERROR
            return $false
        }
    }

    try {
        Set-Location "${SiteCode}:" -ErrorAction Stop
        $site = Get-CMSite -SiteCode $SiteCode -ErrorAction Stop
        Write-Log "Connected to site $SiteCode ($($site.SiteName))"
        $script:ConnectedSiteCode    = $SiteCode
        $script:ConnectedSMSProvider = $SMSProvider
        return $true
    }
    catch {
        Write-Log "Failed to connect to site $SiteCode : $_" -Level ERROR
        return $false
    }
}

function Disconnect-CMSite {
    <#
    .SYNOPSIS
        Restores the original location before CM connection.
    #>
    if ($script:OriginalLocation) {
        try { Set-Location $script:OriginalLocation -ErrorAction SilentlyContinue } catch { }
    }
    $script:ConnectedSiteCode    = $null
    $script:ConnectedSMSProvider = $null
    Write-Log "Disconnected from CM site"
}

function Test-CMConnection {
    <#
    .SYNOPSIS
        Returns $true if currently connected to a CM site.
    #>
    if (-not $script:ConnectedSiteCode) { return $false }

    try {
        $drive = Get-PSDrive -Name $script:ConnectedSiteCode -PSProvider CMSite -ErrorAction Stop
        return ($null -ne $drive)
    }
    catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# Data Retrieval
# ---------------------------------------------------------------------------

function Get-AllMaintenanceWindows {
    <#
    .SYNOPSIS
        Retrieves all maintenance windows across all device collections.
    .DESCRIPTION
        Iterates all device collections via Get-CMDeviceCollection and queries
        each for maintenance windows via Get-CMMaintenanceWindow. Returns flat
        PSCustomObject array.
    #>
    Write-Log "Querying all maintenance windows via CM cmdlets..."

    # Get all device collections
    $allCollections = @(Get-CMDeviceCollection -ErrorAction Stop)
    Write-Log "Checking $($allCollections.Count) device collections for maintenance windows..."

    $results = [System.Collections.ArrayList]::new()

    foreach ($coll in $allCollections) {
        $collId = $coll.CollectionID
        $collName = $coll.Name

        try {
            $windows = Get-CMMaintenanceWindow -CollectionId $collId -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log "Failed to get windows for $collId : $_" -Level WARN
            continue
        }

        if (-not $windows) { continue }

        foreach ($w in $windows) {
            $typeStr = switch ([int]$w.ServiceWindowType) {
                1 { 'General' }
                4 { 'Software Updates' }
                5 { 'Task Sequences' }
                default { "Unknown ($([int]$w.ServiceWindowType))" }
            }

            $recurrenceStr = switch ([int]$w.RecurrenceType) {
                1 { 'None (One-time)' }
                2 { 'Daily' }
                3 { 'Weekly' }
                4 { 'Monthly by Weekday' }
                5 { 'Monthly by Date' }
                default { "Unknown ($([int]$w.RecurrenceType))" }
            }

            $humanSchedule = ConvertTo-HumanSchedule -Window $w
            $durationMinutes = [int]$w.Duration
            $durationStr = "{0}h {1}m" -f [math]::Floor($durationMinutes / 60), ($durationMinutes % 60)

            $nextOccurrences = Get-NextOccurrences -Window $w -Count 1
            $nextStr = if ($nextOccurrences -and $nextOccurrences.Count -gt 0) {
                $nextOccurrences[0].ToString('yyyy-MM-dd HH:mm')
            } else { 'N/A' }

            [void]$results.Add([PSCustomObject]@{
                CollectionName   = $collName
                CollectionID     = $collId
                WindowName       = $w.Name
                WindowID         = $w.ServiceWindowID
                Type             = $typeStr
                Recurrence       = $recurrenceStr
                Schedule         = $humanSchedule
                Duration         = $durationStr
                DurationMinutes  = $durationMinutes
                NextOccurrence   = $nextStr
                StartTime        = $w.StartTime
                IsEnabled        = [bool]$w.IsEnabled
                IsUTC            = [bool]$w.IsGMT
                Description      = $w.Description
                ServiceWindowType = [int]$w.ServiceWindowType
                RecurrenceType   = [int]$w.RecurrenceType
            })
        }
    }

    Write-Log "Found $($results.Count) maintenance windows across all collections"
    return $results
}

function Get-CollectionMaintenanceWindows {
    <#
    .SYNOPSIS
        Returns maintenance windows for a specific collection.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId
    )

    Write-Log "Getting maintenance windows for collection $CollectionId..." -Quiet
    try {
        $windows = Get-CMMaintenanceWindow -CollectionId $CollectionId -ErrorAction Stop
        return $windows
    }
    catch {
        Write-Log "Failed to get maintenance windows for $CollectionId : $_" -Level ERROR
        return $null
    }
}

function Get-DeviceCollectionSummary {
    <#
    .SYNOPSIS
        Returns all device collections with Name, CollectionID, MemberCount.
    #>
    Write-Log "Querying device collection summary..."

    $collections = Get-CMDeviceCollection -ErrorAction Stop

    $results = foreach ($c in $collections) {
        [PSCustomObject]@{
            CollectionID = $c.CollectionID
            Name         = $c.Name
            MemberCount  = [int]$c.MemberCount
            IsBuiltIn    = $c.CollectionID -like 'SMS*'
        }
    }

    Write-Log "Found $(@($results).Count) device collections"
    return $results
}

function Get-CollectionsWithoutWindows {
    <#
    .SYNOPSIS
        Returns device collections that have zero maintenance windows.
    .PARAMETER AllCollections
        Array of collection summary objects from Get-DeviceCollectionSummary.
    .PARAMETER AllWindows
        Array of window objects from Get-AllMaintenanceWindows.
    #>
    param(
        [Parameter(Mandatory)][array]$AllCollections,
        [Parameter(Mandatory)][array]$AllWindows
    )

    $withWindows = @{}
    foreach ($w in $AllWindows) {
        $withWindows[$w.CollectionID] = $true
    }

    $without = foreach ($c in $AllCollections) {
        if (-not $withWindows.ContainsKey($c.CollectionID)) {
            $c
        }
    }

    return $without
}

function ConvertTo-HumanSchedule {
    <#
    .SYNOPSIS
        Converts a maintenance window object to a human-readable schedule string.
    #>
    param(
        [Parameter(Mandatory)]$Window
    )

    $startTime = if ($Window.StartTime) {
        ([datetime]$Window.StartTime).ToString('h:mm tt')
    } else { '?' }

    $durationMin = [int]$Window.Duration
    $durHrs = [math]::Floor($durationMin / 60)
    $durMin = $durationMin % 60
    $durStr = if ($durMin -gt 0) { "{0}h {1}m" -f $durHrs, $durMin } else { "{0}h" -f $durHrs }

    switch ([int]$Window.RecurrenceType) {
        1 {
            # One-time
            $dateStr = if ($Window.StartTime) {
                ([datetime]$Window.StartTime).ToString('yyyy-MM-dd h:mm tt')
            } else { '?' }
            return "One-time: $dateStr ($durStr)"
        }
        2 {
            # Daily
            return "Daily at $startTime ($durStr)"
        }
        3 {
            # Weekly
            $dayOfWeek = if ($Window.StartTime) {
                ([datetime]$Window.StartTime).DayOfWeek
            } else { '?' }
            return "Every $dayOfWeek at $startTime ($durStr)"
        }
        4 {
            # Monthly by weekday
            $dayOfWeek = if ($Window.StartTime) {
                ([datetime]$Window.StartTime).DayOfWeek
            } else { '?' }
            return "Monthly ($dayOfWeek) at $startTime ($durStr)"
        }
        5 {
            # Monthly by date
            $dayOfMonth = if ($Window.StartTime) {
                ([datetime]$Window.StartTime).Day
            } else { '?' }
            return "Monthly (day $dayOfMonth) at $startTime ($durStr)"
        }
        default {
            return "Schedule at $startTime ($durStr)"
        }
    }
}

function Get-NextOccurrences {
    <#
    .SYNOPSIS
        Calculates the next N occurrence datetimes for a maintenance window.
    .PARAMETER Window
        A maintenance window object from Get-CMMaintenanceWindow.
    .PARAMETER Count
        Number of future occurrences to calculate.
    #>
    param(
        [Parameter(Mandatory)]$Window,
        [int]$Count = 5
    )

    $now = Get-Date
    $startTime = if ($Window.StartTime) { [datetime]$Window.StartTime } else { return @() }
    $recurrenceType = [int]$Window.RecurrenceType

    $occurrences = [System.Collections.ArrayList]::new()

    switch ($recurrenceType) {
        1 {
            # One-time: only return if in the future
            if ($startTime -gt $now) {
                [void]$occurrences.Add($startTime)
            }
        }
        2 {
            # Daily
            $candidate = $startTime
            while ($candidate -le $now) { $candidate = $candidate.AddDays(1) }
            for ($i = 0; $i -lt $Count; $i++) {
                [void]$occurrences.Add($candidate)
                $candidate = $candidate.AddDays(1)
            }
        }
        3 {
            # Weekly
            $candidate = $startTime
            while ($candidate -le $now) { $candidate = $candidate.AddDays(7) }
            for ($i = 0; $i -lt $Count; $i++) {
                [void]$occurrences.Add($candidate)
                $candidate = $candidate.AddDays(7)
            }
        }
        4 {
            # Monthly by weekday - approximate: advance by months, find matching weekday
            $targetDayOfWeek = $startTime.DayOfWeek
            $weekOfMonth = [math]::Ceiling($startTime.Day / 7)
            $candidate = $startTime
            while ($candidate -le $now) {
                $candidate = $candidate.AddMonths(1)
                # Recalculate to the Nth weekday of that month
                $firstOfMonth = Get-Date -Year $candidate.Year -Month $candidate.Month -Day 1 -Hour $startTime.Hour -Minute $startTime.Minute -Second 0
                $firstTargetDay = $firstOfMonth
                while ($firstTargetDay.DayOfWeek -ne $targetDayOfWeek) {
                    $firstTargetDay = $firstTargetDay.AddDays(1)
                }
                $candidate = $firstTargetDay.AddDays(7 * ($weekOfMonth - 1))
            }

            for ($i = 0; $i -lt $Count; $i++) {
                [void]$occurrences.Add($candidate)
                $nextMonth = $candidate.AddMonths(1)
                $firstOfMonth = Get-Date -Year $nextMonth.Year -Month $nextMonth.Month -Day 1 -Hour $startTime.Hour -Minute $startTime.Minute -Second 0
                $firstTargetDay = $firstOfMonth
                while ($firstTargetDay.DayOfWeek -ne $targetDayOfWeek) {
                    $firstTargetDay = $firstTargetDay.AddDays(1)
                }
                $candidate = $firstTargetDay.AddDays(7 * ($weekOfMonth - 1))
            }
        }
        5 {
            # Monthly by date
            $targetDay = $startTime.Day
            $candidate = $startTime
            while ($candidate -le $now) { $candidate = $candidate.AddMonths(1) }
            for ($i = 0; $i -lt $Count; $i++) {
                [void]$occurrences.Add($candidate)
                $candidate = $candidate.AddMonths(1)
            }
        }
        default {
            # Unknown recurrence, return empty
        }
    }

    return $occurrences.ToArray()
}

# ---------------------------------------------------------------------------
# CRUD
# ---------------------------------------------------------------------------

function New-ManagedMaintenanceWindow {
    <#
    .SYNOPSIS
        Creates a maintenance window on a collection with validation and logging.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Schedule,
        [ValidateSet('Any', 'SoftwareUpdatesOnly', 'TaskSequencesOnly')]
        [string]$ApplyTo = 'Any',
        [bool]$IsEnabled = $true,
        [bool]$IsUtc = $false
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        Write-Log "Maintenance window name cannot be empty" -Level ERROR
        return $null
    }

    Write-Log "Creating maintenance window '$Name' on collection $CollectionId (ApplyTo=$ApplyTo, UTC=$IsUtc)"

    try {
        $params = @{
            CollectionId = $CollectionId
            Name         = $Name
            Schedule     = $Schedule
            ApplyTo      = $ApplyTo
            IsEnabled    = $IsEnabled
            IsUtc        = $IsUtc
            ErrorAction  = 'Stop'
        }

        $result = New-CMMaintenanceWindow @params
        Write-Log "Created maintenance window '$Name' on $CollectionId"
        return $result
    }
    catch {
        Write-Log "Failed to create maintenance window '$Name' on $CollectionId : $_" -Level ERROR
        return $null
    }
}

function Set-ManagedMaintenanceWindow {
    <#
    .SYNOPSIS
        Updates an existing maintenance window with before/after logging.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$MaintenanceWindowName,
        $Schedule,
        [string]$ApplyTo,
        [object]$IsEnabled,
        [object]$IsUtc
    )

    Write-Log "Updating maintenance window '$MaintenanceWindowName' on collection $CollectionId"

    try {
        $params = @{
            CollectionId          = $CollectionId
            MaintenanceWindowName = $MaintenanceWindowName
            ErrorAction           = 'Stop'
        }

        if ($Schedule)                     { $params['Schedule']  = $Schedule }
        if ($ApplyTo)                      { $params['ApplyTo']   = $ApplyTo }
        if ($null -ne $IsEnabled)          { $params['IsEnabled'] = [bool]$IsEnabled }
        if ($null -ne $IsUtc)              { $params['IsUtc']     = [bool]$IsUtc }

        Set-CMMaintenanceWindow @params
        Write-Log "Updated maintenance window '$MaintenanceWindowName' on $CollectionId"
        return $true
    }
    catch {
        Write-Log "Failed to update maintenance window '$MaintenanceWindowName' on $CollectionId : $_" -Level ERROR
        return $false
    }
}

function Remove-ManagedMaintenanceWindow {
    <#
    .SYNOPSIS
        Removes a maintenance window with confirmation logging.
    #>
    param(
        [Parameter(Mandatory)][string]$CollectionId,
        [Parameter(Mandatory)][string]$MaintenanceWindowName
    )

    Write-Log "Removing maintenance window '$MaintenanceWindowName' from collection $CollectionId"

    try {
        Remove-CMMaintenanceWindow -CollectionId $CollectionId -MaintenanceWindowName $MaintenanceWindowName `
            -Force -ErrorAction Stop
        Write-Log "Removed maintenance window '$MaintenanceWindowName' from $CollectionId"
        return $true
    }
    catch {
        Write-Log "Failed to remove maintenance window '$MaintenanceWindowName' from $CollectionId : $_" -Level ERROR
        return $false
    }
}

# ---------------------------------------------------------------------------
# Schedule Helpers
# ---------------------------------------------------------------------------

function New-WindowSchedule {
    <#
    .SYNOPSIS
        Wraps New-CMSchedule with friendlier parameter names for maintenance window creation.
    .PARAMETER RecurrenceType
        One of: OneTime, Daily, Weekly, MonthlyByDate, MonthlyByWeekday
    .PARAMETER StartTime
        Start datetime for the schedule.
    .PARAMETER DurationHours
        Duration in hours (0-23).
    .PARAMETER DurationMinutes
        Duration in minutes (0-59).
    .PARAMETER DayOfWeek
        Day of week for Weekly or MonthlyByWeekday (Sunday..Saturday).
    .PARAMETER WeekOrder
        Week order for MonthlyByWeekday (First, Second, Third, Fourth, Last).
    .PARAMETER DayOfMonth
        Day of month for MonthlyByDate (1-31).
    .PARAMETER RecurCount
        Recurrence interval (every N days/weeks). Default 1.
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('OneTime', 'Daily', 'Weekly', 'MonthlyByDate', 'MonthlyByWeekday')]
        [string]$RecurrenceType,

        [Parameter(Mandatory)][datetime]$StartTime,

        [int]$DurationHours = 4,
        [int]$DurationMinutes = 0,

        [System.DayOfWeek]$DayOfWeek = [System.DayOfWeek]::Sunday,
        [string]$WeekOrder = 'First',
        [int]$DayOfMonth = 1,
        [int]$RecurCount = 1
    )

    $totalMinutes = ($DurationHours * 60) + $DurationMinutes
    if ($totalMinutes -le 0 -or $totalMinutes -ge 1440) {
        Write-Log "Duration must be between 1 minute and 23 hours 59 minutes" -Level ERROR
        return $null
    }

    $baseParams = @{
        Start             = $StartTime
        DurationInterval  = 'Hours'
        DurationCount     = $DurationHours
        ErrorAction       = 'Stop'
    }

    # Handle minutes-only or mixed duration
    if ($DurationMinutes -gt 0 -and $DurationHours -eq 0) {
        $baseParams['DurationInterval'] = 'Minutes'
        $baseParams['DurationCount']    = $DurationMinutes
    }
    elseif ($DurationMinutes -gt 0) {
        # CM schedule duration is limited; set hours and note minutes are approximate
        $baseParams['DurationCount'] = $DurationHours
    }

    try {
        switch ($RecurrenceType) {
            'OneTime' {
                $schedule = New-CMSchedule @baseParams -Nonrecurring
            }
            'Daily' {
                $schedule = New-CMSchedule @baseParams -RecurInterval Days -RecurCount $RecurCount
            }
            'Weekly' {
                $schedule = New-CMSchedule @baseParams -DayOfWeek $DayOfWeek -RecurCount $RecurCount
            }
            'MonthlyByDate' {
                $schedule = New-CMSchedule @baseParams -DayOfMonth $DayOfMonth
            }
            'MonthlyByWeekday' {
                $schedule = New-CMSchedule @baseParams -DayOfWeek $DayOfWeek -WeekOrder $WeekOrder
            }
        }

        Write-Log "Created schedule: $RecurrenceType, Start=$($StartTime.ToString('yyyy-MM-dd HH:mm')), Duration=${DurationHours}h${DurationMinutes}m" -Quiet
        return $schedule
    }
    catch {
        Write-Log "Failed to create schedule: $_" -Level ERROR
        return $null
    }
}

function Get-PatchTuesday {
    <#
    .SYNOPSIS
        Returns the Patch Tuesday (second Tuesday) date for a given month.
    #>
    param(
        [int]$Year  = (Get-Date).Year,
        [int]$Month = (Get-Date).Month
    )

    $firstOfMonth = Get-Date -Year $Year -Month $Month -Day 1 -Hour 0 -Minute 0 -Second 0
    $firstTuesday = $firstOfMonth
    while ($firstTuesday.DayOfWeek -ne [System.DayOfWeek]::Tuesday) {
        $firstTuesday = $firstTuesday.AddDays(1)
    }
    $patchTuesday = $firstTuesday.AddDays(7) # second Tuesday
    return $patchTuesday
}

function New-PatchTuesdaySchedule {
    <#
    .SYNOPSIS
        Creates a monthly schedule for Patch Tuesday + optional day offset.
    .PARAMETER OffsetDays
        Days after Patch Tuesday (0 = Patch Tuesday itself, 3 = Friday after, 7 = following Tuesday).
    .PARAMETER StartHour
        Hour to start the window (0-23).
    .PARAMETER StartMinute
        Minute to start the window (0-59).
    .PARAMETER DurationHours
        Duration in hours.
    #>
    param(
        [int]$OffsetDays = 0,
        [int]$StartHour = 2,
        [int]$StartMinute = 0,
        [int]$DurationHours = 4,
        [int]$DurationMinutes = 0
    )

    # Calculate the target day: Patch Tuesday + offset
    # If offset is 0, target is Second Tuesday
    # If offset > 0, we need a different approach since CM doesn't natively support offsets

    if ($OffsetDays -eq 0) {
        # Exactly Patch Tuesday: second Tuesday of the month
        $startTime = Get-Date -Year (Get-Date).Year -Month (Get-Date).Month -Day 1 `
            -Hour $StartHour -Minute $StartMinute -Second 0
        return New-WindowSchedule -RecurrenceType MonthlyByWeekday `
            -StartTime $startTime -DayOfWeek Tuesday -WeekOrder Second `
            -DurationHours $DurationHours -DurationMinutes $DurationMinutes
    }
    else {
        # Offset from Patch Tuesday: calculate next occurrence and create schedule
        # CM doesn't support "second Tuesday + N days" natively, so we calculate
        # the actual date for the current/next month and create a monthly-by-date schedule
        $pt = Get-PatchTuesday
        $targetDate = $pt.AddDays($OffsetDays)

        # If target date has passed this month, use next month
        if ($targetDate -lt (Get-Date)) {
            $pt = Get-PatchTuesday -Year (Get-Date).AddMonths(1).Year -Month (Get-Date).AddMonths(1).Month
            $targetDate = $pt.AddDays($OffsetDays)
        }

        $startTime = Get-Date -Year $targetDate.Year -Month $targetDate.Month -Day $targetDate.Day `
            -Hour $StartHour -Minute $StartMinute -Second 0

        Write-Log "Patch Tuesday + $OffsetDays days = day $($targetDate.Day) of month (approximate, varies monthly)" -Quiet

        return New-WindowSchedule -RecurrenceType MonthlyByDate `
            -StartTime $startTime -DayOfMonth $targetDate.Day `
            -DurationHours $DurationHours -DurationMinutes $DurationMinutes
    }
}

# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------

function Get-WindowTemplates {
    <#
    .SYNOPSIS
        Loads all maintenance window templates from the Templates folder.
    .PARAMETER TemplatesPath
        Path to the Templates folder.
    #>
    param(
        [Parameter(Mandatory)][string]$TemplatesPath
    )

    if (-not (Test-Path -LiteralPath $TemplatesPath)) { return @() }

    $templates = @()
    $jsonFiles = Get-ChildItem -LiteralPath $TemplatesPath -Filter '*.json' -File -ErrorAction SilentlyContinue

    foreach ($f in $jsonFiles) {
        try {
            $tmpl = Get-Content -LiteralPath $f.FullName -Raw | ConvertFrom-Json
            $tmpl | Add-Member -NotePropertyName 'FileName' -NotePropertyValue $f.Name -Force
            $tmpl | Add-Member -NotePropertyName 'FilePath' -NotePropertyValue $f.FullName -Force
            $templates += $tmpl
        }
        catch {
            Write-Log "Failed to load template $($f.Name): $_" -Level WARN
        }
    }

    return $templates
}

function Save-WindowTemplate {
    <#
    .SYNOPSIS
        Saves a maintenance window template to JSON.
    #>
    param(
        [Parameter(Mandatory)][string]$TemplatesPath,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$RecurrenceType,
        [int]$DurationHours = 4,
        [int]$DurationMinutes = 0,
        [string]$WindowType = 'Any',
        [bool]$IsUtc = $false,
        [string]$Description = '',
        [int]$StartHour = 2,
        [int]$StartMinute = 0,
        [System.DayOfWeek]$DayOfWeek = [System.DayOfWeek]::Sunday,
        [string]$WeekOrder = 'First',
        [int]$DayOfMonth = 1,
        [int]$RecurCount = 1,
        [int]$PatchTuesdayOffset = -1
    )

    if (-not (Test-Path -LiteralPath $TemplatesPath)) {
        New-Item -ItemType Directory -Path $TemplatesPath -Force | Out-Null
    }

    $template = [ordered]@{
        Name              = $Name
        Description       = $Description
        RecurrenceType    = $RecurrenceType
        WindowType        = $WindowType
        DurationHours     = $DurationHours
        DurationMinutes   = $DurationMinutes
        StartHour         = $StartHour
        StartMinute       = $StartMinute
        IsUtc             = $IsUtc
        DayOfWeek         = $DayOfWeek.ToString()
        WeekOrder         = $WeekOrder
        DayOfMonth        = $DayOfMonth
        RecurCount        = $RecurCount
        PatchTuesdayOffset = $PatchTuesdayOffset
    }

    $fileName = ($Name -replace '[^\w\-]', '-').ToLower() + '.json'
    $filePath = Join-Path $TemplatesPath $fileName

    $template | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $filePath -Encoding UTF8
    Write-Log "Saved template '$Name' to $fileName"
    return $filePath
}

function Remove-WindowTemplate {
    <#
    .SYNOPSIS
        Deletes a template JSON file.
    #>
    param(
        [Parameter(Mandatory)][string]$FilePath
    )

    if (Test-Path -LiteralPath $FilePath) {
        Remove-Item -LiteralPath $FilePath -Force
        Write-Log "Removed template: $FilePath"
        return $true
    }
    return $false
}

function Apply-WindowTemplate {
    <#
    .SYNOPSIS
        Creates a maintenance window on each target collection from a template.
    .RETURNS
        Array of PSCustomObjects with CollectionId, CollectionName, Success, Error properties.
    #>
    param(
        [Parameter(Mandatory)]$Template,
        [Parameter(Mandatory)][array]$TargetCollections,
        [string]$WindowNameOverride
    )

    $results = [System.Collections.ArrayList]::new()

    # Build schedule from template
    $startTime = Get-Date -Hour $Template.StartHour -Minute $Template.StartMinute -Second 0

    $scheduleParams = @{
        RecurrenceType  = $Template.RecurrenceType
        StartTime       = $startTime
        DurationHours   = [int]$Template.DurationHours
        DurationMinutes = [int]$Template.DurationMinutes
    }

    if ($Template.RecurrenceType -eq 'Weekly' -or $Template.RecurrenceType -eq 'MonthlyByWeekday') {
        $scheduleParams['DayOfWeek'] = [System.DayOfWeek]$Template.DayOfWeek
    }
    if ($Template.RecurrenceType -eq 'MonthlyByWeekday') {
        $scheduleParams['WeekOrder'] = $Template.WeekOrder
    }
    if ($Template.RecurrenceType -eq 'MonthlyByDate') {
        $scheduleParams['DayOfMonth'] = [int]$Template.DayOfMonth
    }
    if ($Template.RecurCount) {
        $scheduleParams['RecurCount'] = [int]$Template.RecurCount
    }

    $schedule = New-WindowSchedule @scheduleParams
    if (-not $schedule) {
        Write-Log "Failed to create schedule from template" -Level ERROR
        return $results
    }

    $windowName = if ($WindowNameOverride) { $WindowNameOverride } else { $Template.Name }
    $applyTo = switch ($Template.WindowType) {
        'SoftwareUpdatesOnly' { 'SoftwareUpdatesOnly' }
        'TaskSequencesOnly'   { 'TaskSequencesOnly' }
        default               { 'Any' }
    }

    foreach ($coll in $TargetCollections) {
        $collId   = if ($coll.CollectionID) { $coll.CollectionID } else { $coll }
        $collName = if ($coll.Name) { $coll.Name } else { $collId }

        try {
            $result = New-ManagedMaintenanceWindow -CollectionId $collId -Name $windowName `
                -Schedule $schedule -ApplyTo $applyTo -IsEnabled $true -IsUtc ([bool]$Template.IsUtc)

            [void]$results.Add([PSCustomObject]@{
                CollectionID   = $collId
                CollectionName = $collName
                Success        = ($null -ne $result)
                Error          = ''
            })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{
                CollectionID   = $collId
                CollectionName = $collName
                Success        = $false
                Error          = $_.Exception.Message
            })
        }
    }

    $successCount = @($results | Where-Object { $_.Success }).Count
    Write-Log "Applied template '$windowName' to $successCount of $($TargetCollections.Count) collections"
    return $results
}

# ---------------------------------------------------------------------------
# Bulk Operations
# ---------------------------------------------------------------------------

function Import-MaintenanceWindowsCsv {
    <#
    .SYNOPSIS
        Parses a CSV file and returns preview objects for maintenance window creation.
    .DESCRIPTION
        Expected CSV columns: CollectionID (or CollectionName), WindowName, RecurrenceType,
        StartHour, StartMinute, DurationHours, DurationMinutes, WindowType, DayOfWeek,
        WeekOrder, DayOfMonth, RecurCount, IsUtc
    #>
    param(
        [Parameter(Mandatory)][string]$CsvPath
    )

    if (-not (Test-Path -LiteralPath $CsvPath)) {
        Write-Log "CSV file not found: $CsvPath" -Level ERROR
        return @()
    }

    $rows = Import-Csv -LiteralPath $CsvPath -ErrorAction Stop
    Write-Log "Parsed $(@($rows).Count) rows from CSV"

    $preview = foreach ($row in $rows) {
        [PSCustomObject]@{
            CollectionID    = if ($row.CollectionID) { $row.CollectionID } else { '' }
            CollectionName  = if ($row.CollectionName) { $row.CollectionName } else { '' }
            WindowName      = if ($row.WindowName) { $row.WindowName } else { 'Maintenance Window' }
            RecurrenceType  = if ($row.RecurrenceType) { $row.RecurrenceType } else { 'Weekly' }
            StartHour       = if ($row.StartHour) { [int]$row.StartHour } else { 2 }
            StartMinute     = if ($row.StartMinute) { [int]$row.StartMinute } else { 0 }
            DurationHours   = if ($row.DurationHours) { [int]$row.DurationHours } else { 4 }
            DurationMinutes = if ($row.DurationMinutes) { [int]$row.DurationMinutes } else { 0 }
            WindowType      = if ($row.WindowType) { $row.WindowType } else { 'Any' }
            DayOfWeek       = if ($row.DayOfWeek) { $row.DayOfWeek } else { 'Sunday' }
            WeekOrder       = if ($row.WeekOrder) { $row.WeekOrder } else { 'First' }
            DayOfMonth      = if ($row.DayOfMonth) { [int]$row.DayOfMonth } else { 1 }
            RecurCount      = if ($row.RecurCount) { [int]$row.RecurCount } else { 1 }
            IsUtc           = if ($row.IsUtc) { [bool]($row.IsUtc -eq 'true' -or $row.IsUtc -eq '1') } else { $false }
            Valid           = $true
            Error           = ''
        }
    }

    return $preview
}

function Copy-MaintenanceWindowToCollections {
    <#
    .SYNOPSIS
        Copies a maintenance window from one collection to multiple target collections.
    #>
    param(
        [Parameter(Mandatory)][string]$SourceCollectionId,
        [Parameter(Mandatory)][string]$SourceWindowName,
        [Parameter(Mandatory)][array]$TargetCollectionIds
    )

    Write-Log "Copying window '$SourceWindowName' from $SourceCollectionId to $($TargetCollectionIds.Count) collections"

    $sourceWindows = Get-CMMaintenanceWindow -CollectionId $SourceCollectionId -MaintenanceWindowName $SourceWindowName -ErrorAction Stop
    if (-not $sourceWindows) {
        Write-Log "Source window '$SourceWindowName' not found on $SourceCollectionId" -Level ERROR
        return @()
    }

    $srcWindow = $sourceWindows | Select-Object -First 1

    $results = [System.Collections.ArrayList]::new()

    foreach ($targetId in $TargetCollectionIds) {
        try {
            # Re-create the schedule from the source window's schedule string
            $schedule = Convert-CMSchedule -ScheduleString $srcWindow.ServiceWindowSchedules -ErrorAction Stop

            $params = @{
                CollectionId = $targetId
                Name         = $srcWindow.Name
                Schedule     = $schedule
                ApplyTo      = switch ([int]$srcWindow.ServiceWindowType) {
                    4 { 'SoftwareUpdatesOnly' }
                    5 { 'TaskSequencesOnly' }
                    default { 'Any' }
                }
                IsEnabled    = [bool]$srcWindow.IsEnabled
                IsUtc        = [bool]$srcWindow.IsGMT
                ErrorAction  = 'Stop'
            }

            New-CMMaintenanceWindow @params | Out-Null

            [void]$results.Add([PSCustomObject]@{
                CollectionID = $targetId
                Success      = $true
                Error        = ''
            })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{
                CollectionID = $targetId
                Success      = $false
                Error        = $_.Exception.Message
            })
        }
    }

    $successCount = @($results | Where-Object { $_.Success }).Count
    Write-Log "Copied window to $successCount of $($TargetCollectionIds.Count) target collections"
    return $results
}

function Set-BulkWindowEnabled {
    <#
    .SYNOPSIS
        Enables or disables maintenance windows matching criteria across collections.
    #>
    param(
        [Parameter(Mandatory)][array]$WindowRecords,
        [Parameter(Mandatory)][bool]$Enabled
    )

    $stateStr = if ($Enabled) { 'Enabling' } else { 'Disabling' }
    Write-Log "$stateStr $($WindowRecords.Count) maintenance windows..."

    $results = [System.Collections.ArrayList]::new()

    foreach ($rec in $WindowRecords) {
        try {
            Set-CMMaintenanceWindow -CollectionId $rec.CollectionID `
                -MaintenanceWindowName $rec.WindowName -IsEnabled $Enabled -ErrorAction Stop

            [void]$results.Add([PSCustomObject]@{
                CollectionID = $rec.CollectionID
                WindowName   = $rec.WindowName
                Success      = $true
                Error        = ''
            })
        }
        catch {
            [void]$results.Add([PSCustomObject]@{
                CollectionID = $rec.CollectionID
                WindowName   = $rec.WindowName
                Success      = $false
                Error        = $_.Exception.Message
            })
        }
    }

    $successCount = @($results | Where-Object { $_.Success }).Count
    Write-Log "$stateStr complete: $successCount of $($WindowRecords.Count) succeeded"
    return $results
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

function Export-MaintenanceWindowsCsv {
    <#
    .SYNOPSIS
        Exports a DataTable of maintenance windows to CSV.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $rows = @()
    foreach ($row in $DataTable.Rows) {
        $obj = [ordered]@{}
        foreach ($col in $DataTable.Columns) {
            $obj[$col.ColumnName] = $row[$col.ColumnName]
        }
        $rows += [PSCustomObject]$obj
    }

    $rows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Log "Exported CSV to $OutputPath"
}

function Export-MaintenanceWindowsHtml {
    <#
    .SYNOPSIS
        Exports a DataTable of maintenance windows to a styled HTML report.
    #>
    param(
        [Parameter(Mandatory)][System.Data.DataTable]$DataTable,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$ReportTitle = 'Maintenance Windows Report'
    )

    $parentDir = Split-Path -Path $OutputPath -Parent
    if ($parentDir -and -not (Test-Path -LiteralPath $parentDir)) {
        New-Item -ItemType Directory -Path $parentDir -Force | Out-Null
    }

    $css = @(
        '<style>',
        'body { font-family: "Segoe UI", Arial, sans-serif; margin: 20px; background: #fafafa; }',
        'h1 { color: #0078D4; margin-bottom: 4px; }',
        '.summary { color: #666; margin-bottom: 12px; font-size: 0.9em; }',
        'table { border-collapse: collapse; width: 100%; margin-top: 12px; }',
        'th { background: #0078D4; color: #fff; padding: 8px 12px; text-align: left; font-size: 13px; }',
        'td { padding: 6px 12px; border-bottom: 1px solid #e0e0e0; font-size: 13px; }',
        'tr:nth-child(even) { background: #f5f5f5; }',
        'tr:hover { background: #e8f0fe; }',
        '.type-general { color: #0078D4; font-weight: bold; }',
        '.type-updates { color: #107C10; font-weight: bold; }',
        '.type-osd { color: #CA5010; font-weight: bold; }',
        '.disabled { color: #999; font-style: italic; }',
        '</style>'
    ) -join "`r`n"

    $headerRow = ($DataTable.Columns | ForEach-Object { "<th>$($_.ColumnName)</th>" }) -join ''

    $bodyRows = foreach ($row in $DataTable.Rows) {
        $cells = foreach ($col in $DataTable.Columns) {
            $val = [string]$row[$col.ColumnName]
            $class = ''

            if ($col.ColumnName -eq 'Type') {
                $class = switch -Wildcard ($val) {
                    'General'           { ' class="type-general"' }
                    'Software Updates'  { ' class="type-updates"' }
                    'Task Sequences'    { ' class="type-osd"' }
                    default             { '' }
                }
            }
            elseif ($col.ColumnName -eq 'Enabled' -and $val -eq 'False') {
                $class = ' class="disabled"'
            }

            "<td$class>$val</td>"
        }
        "<tr>$($cells -join '')</tr>"
    }

    $html = @(
        '<!DOCTYPE html>',
        '<html><head><meta charset="utf-8"><title>' + $ReportTitle + '</title>',
        $css,
        '</head><body>',
        "<h1>$ReportTitle</h1>",
        "<div class='summary'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | Windows: $($DataTable.Rows.Count)</div>",
        "<table><thead><tr>$headerRow</tr></thead>",
        "<tbody>$($bodyRows -join "`r`n")</tbody></table>",
        '</body></html>'
    ) -join "`r`n"

    Set-Content -LiteralPath $OutputPath -Value $html -Encoding UTF8
    Write-Log "Exported HTML to $OutputPath"
}
