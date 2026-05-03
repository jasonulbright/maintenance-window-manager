<#
.SYNOPSIS
    MahApps.Metro WPF shell for the MECM Maintenance Window Manager.

.DESCRIPTION
    Sidebar navigation across three views (Windows, Coverage, Templates),
    inline action bar (Refresh, filter, status filter, exports), modal
    dialogs for Options, New / Edit Window, Apply Template, and Template
    Editor. Schedule editor is inline inside the window dialogs and
    supports One-time / Daily / Weekly / MonthlyByDate / MonthlyByWeekday
    / Patch Tuesday recurrence with a live next-5-occurrences preview.

    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.7.2+
      - MahApps.Metro DLLs in .\Lib\
      - MaintWindowMgrCommon module under .\Module\
      - ConfigurationManager console (Get-CMDeviceCollection, Get-CMMaintenanceWindow, etc.)

.NOTES
    ScriptName : start-maintenancewindowmgr.ps1
    Version    : 1.0.0
    Updated    : 2026-05-02
#>

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification='Per feedback_ps_wpf_handler_rules.md and PS51-WPF-001..003: flat-.ps1 GetNewClosure strips $script: scope. $global: survives closure scope-strip and keeps shared mutable state reachable from closure-captured handlers.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification='WPF event handler scriptblocks bind positional sender/args ($s, $e). The sender is required to fulfill the signature even when the handler body does not read it.')]
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# =============================================================================
# Startup transcript.
# =============================================================================
$__txDir = Join-Path $PSScriptRoot 'Logs'
try {
    if (-not (Test-Path -LiteralPath $__txDir)) { New-Item -ItemType Directory -Path $__txDir -Force | Out-Null }
    $__tx = Join-Path $__txDir ('MWM-startup-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    Start-Transcript -LiteralPath $__tx -Force | Out-Null
} catch { $null = $_ }

# =============================================================================
# STA guard. WPF requires STA. PS51-WPF-009.
# =============================================================================
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $psExe = (Get-Process -Id $PID).Path
    $fwd   = @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',$PSCommandPath)
    Start-Process -FilePath $psExe -ArgumentList $fwd | Out-Null
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    exit 0
}

# =============================================================================
# Assemblies.
# =============================================================================
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$libDir = Join-Path $PSScriptRoot 'Lib'
if (-not (Test-Path -LiteralPath $libDir)) {
    throw "Lib/ directory not found at: $libDir. Re-extract the release zip."
}

Get-ChildItem -LiteralPath $libDir -File -ErrorAction SilentlyContinue |
    Unblock-File -ErrorAction SilentlyContinue

[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'Microsoft.Xaml.Behaviors.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'ControlzEx.dll'))
[void][System.Reflection.Assembly]::LoadFrom((Join-Path $libDir 'MahApps.Metro.dll'))

# =============================================================================
# Module import.
# =============================================================================
$__modulePath = Join-Path $PSScriptRoot 'Module\MaintWindowMgrCommon.psd1'
if (-not (Test-Path -LiteralPath $__modulePath)) {
    throw "Shared module not found at: $__modulePath"
}
Import-Module -Name $__modulePath -Force -DisableNameChecking
if (-not (Get-Command Initialize-Logging -ErrorAction SilentlyContinue)) {
    throw "MaintWindowMgrCommon imported but Initialize-Logging is not exported."
}

# =============================================================================
# Preferences.
# =============================================================================
$global:PrefsPath = Join-Path $PSScriptRoot 'MaintWindowMgr.prefs.json'

function Get-MwmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Returns the full preferences hashtable by design.')]
    param()
    $defaults = @{
        DarkMode    = $true
        SiteCode    = ''
        SMSProvider = ''
    }
    if (Test-Path -LiteralPath $global:PrefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $global:PrefsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            foreach ($k in @($defaults.Keys)) {
                $val = $loaded.$k
                if ($null -ne $val) { $defaults[$k] = $val }
            }
        } catch { $null = $_ }
    }
    return $defaults
}

function Save-MwmPreferences {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Writes the full preferences hashtable by design.')]
    param([Parameter(Mandatory)][hashtable]$Prefs)
    try {
        $Prefs | ConvertTo-Json | Set-Content -LiteralPath $global:PrefsPath -Encoding UTF8
    } catch { $null = $_ }
}

$global:Prefs = Get-MwmPreferences

# =============================================================================
# Tool log.
# =============================================================================
$script:ToolLogPath = Join-Path $__txDir ('MWM-{0}.log' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $script:ToolLogPath

# =============================================================================
# Load XAML and resolve named elements.
# =============================================================================
$xamlPath = Join-Path $PSScriptRoot 'MainWindow.xaml'
[xml]$xaml = Get-Content -LiteralPath $xamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$txtAppTitle        = $window.FindName('txtAppTitle')
$txtVersion         = $window.FindName('txtVersion')
$txtThemeLabel      = $window.FindName('txtThemeLabel')
$toggleTheme        = $window.FindName('toggleTheme')

$btnViewWindows   = $window.FindName('btnViewWindows')
$btnViewCoverage  = $window.FindName('btnViewCoverage')
$btnViewTemplates = $window.FindName('btnViewTemplates')
$btnOptions       = $window.FindName('btnOptions')

$txtModuleTitle    = $window.FindName('txtModuleTitle')
$txtModuleSubtitle = $window.FindName('txtModuleSubtitle')

$btnRefresh      = $window.FindName('btnRefresh')
$txtFilter       = $window.FindName('txtFilter')
$cboStatusFilter = $window.FindName('cboStatusFilter')
$btnExportCsv    = $window.FindName('btnExportCsv')
$btnExportHtml   = $window.FindName('btnExportHtml')

$viewWindows   = $window.FindName('viewWindows')
$viewCoverage  = $window.FindName('viewCoverage')
$viewTemplates = $window.FindName('viewTemplates')

$gridWindows         = $window.FindName('gridWindows')
$txtWindowProperties = $window.FindName('txtWindowProperties')
$txtNextOccurrences  = $window.FindName('txtNextOccurrences')

$btnNewWindow     = $window.FindName('btnNewWindow')
$btnEditWindow    = $window.FindName('btnEditWindow')
$btnToggleEnabled = $window.FindName('btnToggleEnabled')
$btnRemoveWindow  = $window.FindName('btnRemoveWindow')

$gridCoverage         = $window.FindName('gridCoverage')
$btnApplyTemplateBulk = $window.FindName('btnApplyTemplateBulk')

$gridTemplates                = $window.FindName('gridTemplates')
$txtTemplateDescription       = $window.FindName('txtTemplateDescription')
$txtTemplateScheduleSummary   = $window.FindName('txtTemplateScheduleSummary')
$btnNewTemplate     = $window.FindName('btnNewTemplate')
$btnEditTemplate    = $window.FindName('btnEditTemplate')
$btnDeleteTemplate  = $window.FindName('btnDeleteTemplate')
$btnApplyTemplateOne = $window.FindName('btnApplyTemplateOne')

$progressOverlay  = $window.FindName('progressOverlay')
$txtProgressTitle = $window.FindName('txtProgressTitle')
$txtProgressStep  = $window.FindName('txtProgressStep')

$lblLogOutput = $window.FindName('lblLogOutput')
$txtLog       = $window.FindName('txtLog')
$txtStatus    = $window.FindName('txtStatus')

$null = $txtAppTitle, $txtVersion

# =============================================================================
# Helpers: log drawer + status bar.
# =============================================================================
function Add-LogLine {
    param([Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $line = '{0}  {1}' -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($txtLog.Text)) {
        $txtLog.Text = $line
    } else {
        $txtLog.AppendText([Environment]::NewLine + $line)
    }
    $txtLog.ScrollToEnd()
}

function Set-StatusText {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param([Parameter(Mandatory)][string]$Text)
    $txtStatus.Text = $Text
}

# =============================================================================
# Title-bar drag fallback (PS51-WPF-033).
# =============================================================================
$script:TitleBarHitTestWindows = @{}
$script:TitleBarHitTestHooks   = @{}

function Get-TitleBarDragHeight {
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try { $h = [double]$Window.TitleBarHeight; if ($h -gt 0 -and -not [double]::IsNaN($h)) { return $h } } catch { $null = $_ }
    return 30.0
}

function Get-InputAncestors {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Visual-tree helper yields an ancestor chain.')]
    param([System.Windows.DependencyObject]$Start)
    $cur = $Start
    while ($cur) {
        $cur
        $parent = $null
        if ($cur -is [System.Windows.Media.Visual] -or $cur -is [System.Windows.Media.Media3D.Visual3D]) {
            try { $parent = [System.Windows.Media.VisualTreeHelper]::GetParent($cur) } catch { $parent = $null }
        }
        if (-not $parent -and $cur -is [System.Windows.FrameworkElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.FrameworkContentElement]) { $parent = $cur.Parent }
        if (-not $parent -and $cur -is [System.Windows.ContentElement]) {
            try { $parent = [System.Windows.ContentOperations]::GetParent($cur) } catch { $parent = $null }
        }
        $cur = $parent
    }
}

function Test-IsWindowCommandPoint {
    param([MahApps.Metro.Controls.MetroWindow]$Window, [System.Windows.Point]$Point)
    try {
        [void]$Window.ApplyTemplate()
        $commands = $Window.Template.FindName('PART_WindowButtonCommands', $Window)
        if ($commands -and $commands.IsVisible -and $commands.ActualWidth -gt 0 -and $commands.ActualHeight -gt 0) {
            $origin = $commands.TransformToAncestor($Window).Transform([System.Windows.Point]::new(0, 0))
            if ($Point.X -ge $origin.X -and $Point.X -le ($origin.X + $commands.ActualWidth) -and
                $Point.Y -ge $origin.Y -and $Point.Y -le ($origin.Y + $commands.ActualHeight)) {
                return $true
            }
        }
    } catch { $null = $_ }
    return ($Window.ActualWidth -gt 150 -and $Point.X -ge ($Window.ActualWidth - 150))
}

function Add-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Installs an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
        if (-not $source) { return }
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) { return }
        $script:TitleBarHitTestWindows[$key] = $Window
        $hook = [System.Windows.Interop.HwndSourceHook]{
            param([IntPtr]$hwnd, [int]$msg, [IntPtr]$wParam, [IntPtr]$lParam, [ref]$handled)
            $WM_NCHITTEST = 0x0084; $HTCAPTION = 2
            if ($msg -ne $WM_NCHITTEST) { return [IntPtr]::Zero }
            try {
                $target = $script:TitleBarHitTestWindows[$hwnd.ToInt64().ToString()]
                if (-not $target) { return [IntPtr]::Zero }
                $raw = $lParam.ToInt64()
                $screenX = [int]($raw -band 0xffff); if ($screenX -ge 0x8000) { $screenX -= 0x10000 }
                $screenY = [int](($raw -shr 16) -band 0xffff); if ($screenY -ge 0x8000) { $screenY -= 0x10000 }
                $pt = $target.PointFromScreen([System.Windows.Point]::new($screenX, $screenY))
                $titleBarH = Get-TitleBarDragHeight -Window $target
                if ($pt.X -lt 0 -or $pt.X -gt $target.ActualWidth) { return [IntPtr]::Zero }
                if ($pt.Y -lt 4 -or $pt.Y -gt $titleBarH) { return [IntPtr]::Zero }
                if (Test-IsWindowCommandPoint -Window $target -Point $pt) { return [IntPtr]::Zero }
                $handled.Value = $true
                return [IntPtr]$HTCAPTION
            } catch { return [IntPtr]::Zero }
        }
        $script:TitleBarHitTestHooks[$key] = $hook
        $source.AddHook($hook)
    } catch { $null = $_ }
}

function Remove-NativeTitleBarHitTestHook {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Removes an in-process HWND hook for this WPF window only.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    try {
        $helper = [System.Windows.Interop.WindowInteropHelper]::new($Window)
        $key = $helper.Handle.ToInt64().ToString()
        if ($script:TitleBarHitTestHooks.ContainsKey($key)) {
            $source = [System.Windows.Interop.HwndSource]::FromHwnd($helper.Handle)
            if ($source) { $source.RemoveHook($script:TitleBarHitTestHooks[$key]) }
            $script:TitleBarHitTestHooks.Remove($key)
        }
        if ($script:TitleBarHitTestWindows.ContainsKey($key)) { $script:TitleBarHitTestWindows.Remove($key) }
    } catch { $null = $_ }
}

function Install-TitleBarDragFallback {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Registers window-local WPF event handlers for title-bar drag fallback.')]
    param([MahApps.Metro.Controls.MetroWindow]$Window)
    $Window.Add_SourceInitialized({ param($s, $e) Add-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_Closed({ param($s, $e) Remove-NativeTitleBarHitTestHook -Window $s })
    $Window.Add_PreviewMouseLeftButtonDown({
        param($s, $e)
        try {
            if ($s.WindowState -eq [System.Windows.WindowState]::Maximized) { return }
            $titleBarH = Get-TitleBarDragHeight -Window $s
            $pos = $e.GetPosition($s)
            if ($pos.Y -lt 4 -or $pos.Y -gt $titleBarH) { return }
            if (Test-IsWindowCommandPoint -Window $s -Point $pos) { return }
            foreach ($ancestor in Get-InputAncestors -Start ($e.OriginalSource -as [System.Windows.DependencyObject])) {
                if ($ancestor -is [System.Windows.Controls.Primitives.ButtonBase]) { return }
            }
            $s.DragMove()
            $e.Handled = $true
        } catch { $null = $_ }
    })
}

Install-TitleBarDragFallback -Window $window

# =============================================================================
# Theme setup.
# =============================================================================
[void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')

$script:DarkButtonBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#1E1E1E')
$script:DarkButtonBorder = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#555555')
$script:DarkActiveBg     = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#3A3A3A')
$script:LightWfBg        = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:LightWfBorder    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#006CBE')
$script:LightActiveBg    = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#005A9E')

$script:TitleBarBlue         = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#0078D4')
$script:TitleBarBlueInactive = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#4BA3E0')

$script:LogLabelDark  = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#B0B0B0')
$script:LogLabelLight = [System.Windows.Media.BrushConverter]::new().ConvertFrom('#595959')

$script:ViewButtons = @(
    @{ Name = 'Windows';   Button = $btnViewWindows   },
    @{ Name = 'Coverage';  Button = $btnViewCoverage  },
    @{ Name = 'Templates'; Button = $btnViewTemplates }
)
$script:ActiveView = 'Windows'

function Update-SidebarButtonTheme {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    param()
    $isDark   = [bool]$global:Prefs['DarkMode']
    $idleBg   = if ($isDark) { $script:DarkButtonBg }     else { $script:LightWfBg }
    $activeBg = if ($isDark) { $script:DarkActiveBg }     else { $script:LightActiveBg }
    $border   = if ($isDark) { $script:DarkButtonBorder } else { $script:LightWfBorder }
    $thickness = [System.Windows.Thickness]::new(1)

    foreach ($v in $script:ViewButtons) {
        if (-not $v.Button) { continue }
        $isActive = ($v.Name -eq $script:ActiveView)
        $v.Button.Background      = if ($isActive) { $activeBg } else { $idleBg }
        $v.Button.BorderBrush     = $border
        $v.Button.BorderThickness = $thickness
    }
    if ($btnOptions) {
        $btnOptions.Background      = $idleBg
        $btnOptions.BorderBrush     = $border
        $btnOptions.BorderThickness = $thickness
    }
    if ($lblLogOutput) {
        $lblLogOutput.Foreground = if ($isDark) { $script:LogLabelDark } else { $script:LogLabelLight }
    }
}

function Update-TitleBarBrushes {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates in-window brush properties only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Sets both active and non-active title brushes per theme.')]
    param()
    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::WindowTitleBrushProperty)
        $window.ClearValue([MahApps.Metro.Controls.MetroWindow]::NonActiveWindowTitleBrushProperty)
    } else {
        $window.WindowTitleBrush          = $script:TitleBarBlue
        $window.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

$__startIsDark = [bool]$global:Prefs['DarkMode']
$toggleTheme.IsOn = $__startIsDark
$txtThemeLabel.Text = if ($__startIsDark) { 'Dark Theme' } else { 'Light Theme' }
Update-SidebarButtonTheme

$toggleTheme.Add_Toggled({
    $isDark = [bool]$toggleTheme.IsOn
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Dark.Steel')
        $txtThemeLabel.Text = 'Dark Theme'
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
        $txtThemeLabel.Text = 'Light Theme'
    }
    $global:Prefs['DarkMode'] = $isDark
    Save-MwmPreferences -Prefs $global:Prefs
    Update-SidebarButtonTheme
    Update-TitleBarBrushes
    Add-LogLine ('Theme: {0}' -f $(if ($isDark) { 'dark' } else { 'light' }))
})

# =============================================================================
# View switching.
# =============================================================================
$script:ViewMeta = @{
    'Windows'   = @{ Title = 'Windows';   Subtitle = 'Every maintenance window across every device collection. Configure Site / Provider in Options, then click Refresh.' }
    'Coverage'  = @{ Title = 'Coverage';  Subtitle = 'Which collections have / lack maintenance windows -- gap remediation. Multi-select rows for bulk apply.' }
    'Templates' = @{ Title = 'Templates'; Subtitle = 'Saved schedule templates. Apply one to a target collection or to many collections at once from the Coverage view.' }
}

function Set-ActiveView {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates in-window Visibility + header text only.')]
    param([Parameter(Mandatory)][ValidateSet('Windows','Coverage','Templates')][string]$View)

    $script:ActiveView = $View

    $viewWindows.Visibility   = if ($View -eq 'Windows')   { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewCoverage.Visibility  = if ($View -eq 'Coverage')  { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }
    $viewTemplates.Visibility = if ($View -eq 'Templates') { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Collapsed }

    $meta = $script:ViewMeta[$View]
    if ($meta) {
        $txtModuleTitle.Text    = $meta.Title
        $txtModuleSubtitle.Text = $meta.Subtitle
    }

    Update-SidebarButtonTheme
    Update-ActionBarVisibility
    Update-Filter
    Update-StatusBarSummary
}

$btnViewWindows.Add_Click({   Set-ActiveView -View 'Windows'   })
$btnViewCoverage.Add_Click({  Set-ActiveView -View 'Coverage'  })
$btnViewTemplates.Add_Click({ Set-ActiveView -View 'Templates' })

# =============================================================================
# Crash handlers (PS51-WPF-010, PS51-WPF-011, PS51-WPF-025).
# =============================================================================
$global:__crashLog = Join-Path $__txDir ('MWM-crash-{0}.txt' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$global:__writeCrash = {
    param($Source, $Exception)
    try {
        $lines = @()
        $lines += ('=== ' + $Source + ' @ ' + (Get-Date -Format 'o') + ' ===')
        $lines += ('Type   : ' + $Exception.GetType().FullName)
        $lines += ('Message: ' + $Exception.Message)
        $lines += ('Stack  :')
        $lines += ([string]$Exception.StackTrace).Split([Environment]::NewLine)
        $inner = $Exception.InnerException
        $depth = 1
        while ($inner) {
            $lines += ('--- InnerException depth ' + $depth + ' ---')
            $lines += ('Type   : ' + $inner.GetType().FullName)
            $lines += ('Message: ' + $inner.Message)
            $inner = $inner.InnerException
            $depth++
        }
        [System.IO.File]::AppendAllText($global:__crashLog, (($lines -join [Environment]::NewLine) + [Environment]::NewLine))
    } catch { $null = $_ }
}

$window.Dispatcher.Add_UnhandledException({ param($s, $e) & $global:__writeCrash 'DispatcherUnhandledException' $e.Exception; $e.Handled = $false })
[AppDomain]::CurrentDomain.Add_UnhandledException({ param($s, $e) & $global:__writeCrash 'AppDomainUnhandledException' ([Exception]$e.ExceptionObject) })

# =============================================================================
# State.
# =============================================================================
$script:AllWindows         = @()
$script:RawWindows         = @()
$script:Collections        = @()
$script:CoverageRows       = @()
$script:Templates          = @()
$script:LastRefreshTime    = $null
$script:IsConnectedFromBg  = $false

# =============================================================================
# Glyph helpers.
# =============================================================================
function Get-WindowStatusGlyph {
    param($Window)
    if (-not $Window) { return '' }
    if (-not $Window.IsEnabled) { return [char]0x22EF }
    return [char]0x2713
}

function Get-CoverageGlyph {
    param($Row)
    if (-not $Row) { return '' }
    if ([int]$Row.WindowCount -eq 0) { return [char]0x26A0 }
    return [char]0x2713
}

# =============================================================================
# Action bar visibility + status bar summary.
# =============================================================================
function Update-ActionBarVisibility {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Toggles in-window Visibility only.')]
    param()
    switch ($script:ActiveView) {
        'Windows' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnRefresh.Content         = 'Refresh'
            $txtFilter.Tag              = 'Filter by collection or window name...'
            $cboStatusFilter.Items.Clear()
            foreach ($it in @('All','Enabled only','Disabled only','SoftwareUpdates only','TaskSequences only')) {
                $cboItem = New-Object System.Windows.Controls.ComboBoxItem
                $cboItem.Content = $it
                if ($it -eq 'All') { $cboItem.IsSelected = $true }
                [void]$cboStatusFilter.Items.Add($cboItem)
            }
        }
        'Coverage' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Visible
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Visible
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Visible
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnRefresh.Content         = 'Refresh'
            $txtFilter.Tag              = 'Filter by collection name or ID...'
            $cboStatusFilter.Items.Clear()
            foreach ($it in @('All','Without windows','With windows','Custom only','Built-in only')) {
                $cboItem = New-Object System.Windows.Controls.ComboBoxItem
                $cboItem.Content = $it
                if ($it -eq 'All') { $cboItem.IsSelected = $true }
                [void]$cboStatusFilter.Items.Add($cboItem)
            }
        }
        'Templates' {
            $cboStatusFilter.Visibility = [System.Windows.Visibility]::Collapsed
            $btnExportCsv.Visibility    = [System.Windows.Visibility]::Collapsed
            $btnExportHtml.Visibility   = [System.Windows.Visibility]::Collapsed
            $txtFilter.Visibility       = [System.Windows.Visibility]::Visible
            $btnRefresh.Content         = 'Reload Templates'
            $txtFilter.Tag              = 'Filter by template name...'
        }
    }
    if ($txtFilter.Visibility -eq [System.Windows.Visibility]::Visible) {
        [MahApps.Metro.Controls.TextBoxHelper]::SetWatermark($txtFilter, [string]$txtFilter.Tag)
    }
}

function Update-StatusBarSummary {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates an in-window TextBlock only.')]
    param()
    $parts = @()
    if ($script:IsConnectedFromBg -and $global:Prefs.SiteCode) {
        $parts += "Connected to $($global:Prefs.SiteCode)"
    } elseif (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        $parts += 'Open Options to configure site code and SMS provider'
    } else {
        $parts += 'Ready. Click Refresh.'
    }
    if (@($script:Collections).Count -gt 0) { $parts += ('{0} collections' -f @($script:Collections).Count) }
    if (@($script:RawWindows).Count -gt 0)  { $parts += ('{0} windows'     -f @($script:RawWindows).Count) }
    if (@($script:Templates).Count -gt 0)   { $parts += ('{0} templates'   -f @($script:Templates).Count) }
    if ($script:LastRefreshTime) {
        $parts += ('last refresh {0}' -f $script:LastRefreshTime.ToString('HH:mm:ss'))
    }
    Set-StatusText ($parts -join '   |   ')
}

# =============================================================================
# Filter + detail panel wiring.
# =============================================================================
function Get-StatusFilterValue {
    if (-not $cboStatusFilter.SelectedItem) { return 'All' }
    $item = $cboStatusFilter.SelectedItem
    if ($item -is [System.Windows.Controls.ComboBoxItem]) { return [string]$item.Content }
    return [string]$item
}

function Update-Filter {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Recomputes ItemsSource on the active grid only.')]
    param()
    $needle = ([string]$txtFilter.Text).Trim().ToLowerInvariant()
    $statusFilter = Get-StatusFilterValue

    switch ($script:ActiveView) {
        'Windows' {
            $rows = $script:AllWindows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.CollectionName).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.WindowName).ToLowerInvariant().Contains($needle)
                })
            }
            switch ($statusFilter) {
                'Enabled only'        { $rows = @($rows | Where-Object { $_.IsEnabled }) }
                'Disabled only'       { $rows = @($rows | Where-Object { -not $_.IsEnabled }) }
                'SoftwareUpdates only' { $rows = @($rows | Where-Object { $_.Type -like '*Software Updates*' }) }
                'TaskSequences only'  { $rows = @($rows | Where-Object { $_.Type -like '*Task Sequences*' }) }
            }
            $gridWindows.ItemsSource = $rows
        }
        'Coverage' {
            $rows = $script:CoverageRows
            if ($needle) {
                $rows = @($rows | Where-Object {
                    ([string]$_.Name).ToLowerInvariant().Contains($needle) -or
                    ([string]$_.CollectionID).ToLowerInvariant().Contains($needle)
                })
            }
            switch ($statusFilter) {
                'Without windows' { $rows = @($rows | Where-Object { [int]$_.WindowCount -eq 0 }) }
                'With windows'    { $rows = @($rows | Where-Object { [int]$_.WindowCount -gt 0 }) }
                'Custom only'     { $rows = @($rows | Where-Object { -not $_.IsBuiltIn }) }
                'Built-in only'   { $rows = @($rows | Where-Object {       $_.IsBuiltIn }) }
            }
            $gridCoverage.ItemsSource = $rows
        }
        'Templates' {
            $rows = $script:Templates
            if ($needle) {
                $rows = @($rows | Where-Object { ([string]$_.Name).ToLowerInvariant().Contains($needle) })
            }
            $gridTemplates.ItemsSource = $rows
        }
    }
}

$txtFilter.Add_TextChanged({ Update-Filter })
$cboStatusFilter.Add_SelectionChanged({ Update-Filter })

$gridWindows.Add_SelectionChanged({
    $row = $gridWindows.SelectedItem
    if (-not $row) {
        $txtWindowProperties.Text = 'Select a maintenance window to see its properties.'
        $txtNextOccurrences.Text  = 'Select a maintenance window to project its next occurrences.'
        return
    }

    $lines = @(
        ('Collection:    {0}  ({1})' -f $row.CollectionName, $row.CollectionID),
        ('Window:        {0}' -f $row.WindowName),
        ('Type:          {0}' -f $row.Type),
        ('Recurrence:    {0}' -f $row.Recurrence),
        ('Schedule:      {0}' -f $row.Schedule),
        ('Duration:      {0}' -f $row.Duration),
        ('Next:          {0}' -f $row.NextOccurrence),
        ('Enabled:       {0}' -f $row.IsEnabled),
        ('UTC:           {0}' -f $row.IsUTC),
        '',
        'Description:',
        $row.Description
    )
    $txtWindowProperties.Text = $lines -join [Environment]::NewLine

    if ($row.PSObject.Properties['Raw'] -and $row.Raw) {
        try {
            $next = @(Get-NextOccurrences -Window $row.Raw -Count 5)
            if (@($next).Count -eq 0) {
                $txtNextOccurrences.Text = '(no future occurrences -- one-time window in the past, or unsupported recurrence type)'
            } else {
                $txtNextOccurrences.Text = (($next | ForEach-Object { $_.ToString('yyyy-MM-dd  ddd  HH:mm') }) -join [Environment]::NewLine)
            }
        } catch {
            $txtNextOccurrences.Text = ('Could not project occurrences: {0}' -f $_.Exception.Message)
        }
    } else {
        $txtNextOccurrences.Text = '(occurrence projection unavailable for this row)'
    }
})

$gridTemplates.Add_SelectionChanged({
    $row = $gridTemplates.SelectedItem
    if (-not $row) {
        $txtTemplateDescription.Text = ''
        $txtTemplateScheduleSummary.Text = ''
        return
    }
    $txtTemplateDescription.Text = [string]$row.Description
    $lines = @(
        ('Recurrence:   {0}' -f $row.RecurrenceType),
        ('Window Type:  {0}' -f $row.WindowType),
        ('Start Time:   {0:00}:{1:00}' -f [int]$row.StartHour, [int]$row.StartMinute),
        ('Duration:     {0}h {1}m' -f [int]$row.DurationHours, [int]$row.DurationMinutes)
    )
    if ($row.RecurrenceType -in @('Weekly','MonthlyByWeekday')) { $lines += ('Day of Week:  {0}' -f $row.DayOfWeek) }
    if ($row.RecurrenceType -eq 'MonthlyByWeekday')             { $lines += ('Week Order:   {0}' -f $row.WeekOrder) }
    if ($row.RecurrenceType -eq 'MonthlyByDate')                { $lines += ('Day of Month: {0}' -f [int]$row.DayOfMonth) }
    if ([int]$row.PatchTuesdayOffset -ge 0)                     { $lines += ('Patch Tuesday Offset: {0} days' -f [int]$row.PatchTuesdayOffset) }
    $lines += ('UTC:          {0}' -f [bool]$row.IsUtc)
    $txtTemplateScheduleSummary.Text = $lines -join [Environment]::NewLine
})

# =============================================================================
# Background runspace for refresh.
# =============================================================================
$script:BgRunspace     = $null
$script:BgPowerShell   = $null
$script:BgInvokeHandle = $null
$script:BgState        = $null
$script:BgTimer        = $null

function Initialize-BgRunspace {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Lazy-init; idempotent.')]
    param()
    if ($script:BgRunspace -and $script:BgRunspace.RunspaceStateInfo.State -eq 'Opened') { return }
    $script:BgRunspace = [runspacefactory]::CreateRunspace()
    $script:BgRunspace.ApartmentState = 'STA'
    $script:BgRunspace.ThreadOptions  = 'ReuseThread'
    $script:BgRunspace.Open()
    $modulePath = Join-Path $PSScriptRoot 'Module\MaintWindowMgrCommon.psd1'
    $initPS = [powershell]::Create()
    $initPS.Runspace = $script:BgRunspace
    [void]$initPS.AddScript({
        param($ModulePath, $LogPath)
        Import-Module -Name $ModulePath -Force -DisableNameChecking
        if ($LogPath) { Initialize-Logging -LogPath $LogPath -Attach }
    }).AddArgument($modulePath).AddArgument($script:ToolLogPath)
    [void]$initPS.Invoke()
    $initPS.Dispose()
}

function Dispose-BgWork {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Dispose semantics intentional.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Tears down ephemeral runspace plumbing only.')]
    param()
    if ($script:BgTimer) { try { $script:BgTimer.Stop() } catch { $null = $_ } ; $script:BgTimer = $null }
    if ($script:BgPowerShell) {
        try { [void]$script:BgPowerShell.Stop() } catch { $null = $_ }
        try { $script:BgPowerShell.Dispose() }   catch { $null = $_ }
        $script:BgPowerShell = $null
    }
    $script:BgInvokeHandle = $null
}

function Invoke-Refresh {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Posts work to the background runspace and arms a DispatcherTimer.')]
    param()

    if ($script:ActiveView -eq 'Templates') {
        Invoke-LoadTemplates
        return
    }

    if (-not $global:Prefs.SiteCode -or -not $global:Prefs.SMSProvider) {
        Add-LogLine 'Refresh: site code and SMS provider must be set in Options first.'
        Set-StatusText 'Open Options to configure site code and SMS provider, then refresh.'
        return
    }

    Initialize-BgRunspace
    Dispose-BgWork

    $script:BgState = [hashtable]::Synchronized(@{
        Step     = 'Connecting...'
        Done     = $false
        Result   = $null
        ErrorMsg = $null
    })

    $btnRefresh.IsEnabled = $false
    $txtProgressTitle.Text = 'Loading maintenance windows'
    $txtProgressStep.Text  = 'Connecting...'
    $progressOverlay.Visibility = [System.Windows.Visibility]::Visible
    Add-LogLine ('Refresh: site={0} provider={1}' -f $global:Prefs.SiteCode, $global:Prefs.SMSProvider)
    Set-StatusText 'Refreshing...'

    $siteCode    = [string]$global:Prefs.SiteCode
    $smsProvider = [string]$global:Prefs.SMSProvider

    $script:BgPowerShell = [powershell]::Create()
    $script:BgPowerShell.Runspace = $script:BgRunspace
    [void]$script:BgPowerShell.AddScript({
        param($SiteCode, $SMSProvider, $State)
        try {
            if (-not (Test-CMConnection)) {
                $State.Step = "Connecting to $SiteCode..."
                $ok = Connect-CMSite -SiteCode $SiteCode -SMSProvider $SMSProvider
                if (-not $ok) {
                    $State.ErrorMsg = "Failed to connect to site $SiteCode (provider $SMSProvider)."
                    return
                }
            }

            $State.Step = 'Loading folder hierarchy...'
            $folders   = @()
            $folderMap = @{}
            try {
                $folders   = @(Get-CMCollectionFolderTree -SMSProvider $SMSProvider -SiteCode $SiteCode)
                $folderMap = Get-CMCollectionFolderMap -SMSProvider $SMSProvider -SiteCode $SiteCode
            } catch {
                # Non-fatal: picker degrades to a flat root list when CIM is blocked.
                $folders = @()
                $folderMap = @{}
            }

            $State.Step = 'Loading device collection summary...'
            $collections = @(Get-DeviceCollectionSummary)

            # Annotate each collection with its FolderID (0 = tree root).
            $collections = @($collections | ForEach-Object {
                $fid = if ($folderMap.ContainsKey([string]$_.CollectionID)) { [int]$folderMap[[string]$_.CollectionID] } else { 0 }
                $_ | Add-Member -MemberType NoteProperty -Name FolderID -Value $fid -Force -PassThru
            })

            $State.Step = 'Loading maintenance windows across all collections...'
            $windows = @(Get-AllMaintenanceWindows)

            $State.Result = [PSCustomObject]@{
                Collections = $collections
                Windows     = $windows
                Folders     = $folders
            }
        }
        catch {
            $State.ErrorMsg = $_.Exception.Message
        }
        finally {
            $State.Done = $true
        }
    }).AddArgument($siteCode).AddArgument($smsProvider).AddArgument($script:BgState)

    $script:BgInvokeHandle = $script:BgPowerShell.BeginInvoke()

    $script:BgTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:BgTimer.Interval = [TimeSpan]::FromMilliseconds(150)
    $script:BgTimer.Add_Tick({
        if ($script:BgState) {
            $current = [string]$script:BgState.Step
            if ($txtProgressStep.Text -ne $current) { $txtProgressStep.Text = $current }
        }
        if ($script:BgState -and $script:BgState.Done) {
            $script:BgTimer.Stop()
            try { [void]$script:BgPowerShell.EndInvoke($script:BgInvokeHandle) } catch { $null = $_ }
            try { $script:BgPowerShell.Dispose() } catch { $null = $_ }
            $script:BgPowerShell   = $null
            $script:BgInvokeHandle = $null

            if ($script:BgState.ErrorMsg) {
                $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
                $btnRefresh.IsEnabled = $true
                $script:IsConnectedFromBg = $false
                Add-LogLine ('Refresh failed: {0}' -f $script:BgState.ErrorMsg)
                Set-StatusText 'Refresh failed.'
                return
            }

            $script:IsConnectedFromBg = $true
            $r = $script:BgState.Result
            $script:Collections = @($r.Collections)
            $script:RawWindows  = @($r.Windows)
            $script:Folders     = @($r.Folders)
            $script:LastRefreshTime = Get-Date

            $script:AllWindows = @($script:RawWindows | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name StatusGlyph -Value (Get-WindowStatusGlyph -Window $_) -Force -PassThru |
                     Add-Member -MemberType NoteProperty -Name Raw         -Value $_                                 -Force -PassThru
            })

            $windowCountByColl = @{}
            foreach ($w in $script:RawWindows) {
                $collId = [string]$w.CollectionID
                if (-not $windowCountByColl.ContainsKey($collId)) { $windowCountByColl[$collId] = 0 }
                $windowCountByColl[$collId] += 1
            }
            $script:CoverageRows = @($script:Collections | ForEach-Object {
                $collId = [string]$_.CollectionID
                $cnt    = if ($windowCountByColl.ContainsKey($collId)) { [int]$windowCountByColl[$collId] } else { 0 }
                [PSCustomObject]@{
                    CoverageGlyph = Get-CoverageGlyph -Row @{ WindowCount = $cnt }
                    Name          = $_.Name
                    CollectionID  = $_.CollectionID
                    MemberCount   = $_.MemberCount
                    WindowCount   = $cnt
                    IsBuiltIn     = $_.IsBuiltIn
                }
            })

            $gridWindows.ItemsSource  = $script:AllWindows
            $gridCoverage.ItemsSource = $script:CoverageRows

            try {
                [void](Connect-CMSite -SiteCode $global:Prefs.SiteCode -SMSProvider $global:Prefs.SMSProvider)
            } catch {
                Add-LogLine ('UI-thread CM connect: {0}' -f $_.Exception.Message)
            }

            Update-Filter
            Update-StatusBarSummary
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled = $true

            Add-LogLine ('Refresh complete: {0} windows across {1} collections.' -f @($script:RawWindows).Count, @($script:Collections).Count)
        }
    })
    $script:BgTimer.Start()
}

function Invoke-LoadTemplates {
    Add-LogLine 'Loading templates from disk...'
    try {
        $tplPath = Join-Path $PSScriptRoot 'Templates'
        $raw = @(Get-WindowTemplates -TemplatesPath $tplPath)
        $script:Templates = @($raw | ForEach-Object {
            $durDisplay = '{0}h {1}m' -f [int]$_.DurationHours, [int]$_.DurationMinutes
            $_ | Add-Member -MemberType NoteProperty -Name DurationDisplay -Value $durDisplay -Force -PassThru
        })
        $gridTemplates.ItemsSource = $script:Templates
        Add-LogLine ('Loaded {0} templates.' -f @($script:Templates).Count)
        Update-StatusBarSummary
    } catch {
        Add-LogLine ('Template load failed: {0}' -f $_.Exception.Message)
    }
}

$btnRefresh.Add_Click({ Invoke-Refresh })

# =============================================================================
# Tree builder (folder hierarchy) used by both the modal collection picker
# and any other pickers in the per-action modals.
# =============================================================================
function Build-CollectionTree {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Mutates the in-window TreeView only.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Build is the natural verb for tree assembly.')]
    param(
        [Parameter(Mandatory)][System.Windows.Controls.TreeView]$TreeView,
        [Parameter(Mandatory)]$AllCollections,
        $AllFolders,
        [string]$Needle = ''
    )

    $colls   = @($AllCollections)
    $folders = @($AllFolders)

    $needleLower = ([string]$Needle).Trim().ToLowerInvariant()
    $hasFilter   = -not [string]::IsNullOrWhiteSpace($needleLower)

    $foldersByParent = @{}
    foreach ($f in $folders) {
        $parentId = [int]$f.ParentID
        if (-not $foldersByParent.ContainsKey($parentId)) { $foldersByParent[$parentId] = @() }
        $foldersByParent[$parentId] += $f
    }
    $collectionsByFolder = @{}
    foreach ($c in $colls) {
        $fid = [int]$c.FolderID
        if (-not $collectionsByFolder.ContainsKey($fid)) { $collectionsByFolder[$fid] = @() }
        $collectionsByFolder[$fid] += $c
    }

    $TreeView.Items.Clear()
    $script:__TreeLeafCount = 0

    $matchesNeedle = {
        param($Coll)
        if (-not $hasFilter) { return $true }
        $name = ([string]$Coll.Name).ToLowerInvariant()
        $id   = ([string]$Coll.CollectionID).ToLowerInvariant()
        return ($name.Contains($needleLower) -or $id.Contains($needleLower))
    }

    $populate = {
        param($ParentNode, [int]$FolderID)
        $any = $false

        $childFolders = if ($foldersByParent.ContainsKey($FolderID)) { @($foldersByParent[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($f in $childFolders) {
            $folderNode = New-Object System.Windows.Controls.TreeViewItem
            $folderNode.Header = ('[+] {0}' -f $f.Name)
            $folderNode.Tag = @{ Type = 'Folder'; Object = $f }
            $folderNode.FontWeight = [System.Windows.FontWeights]::SemiBold
            if ($hasFilter) { $folderNode.IsExpanded = $true }
            $hadAny = & $populate $folderNode ([int]$f.FolderID)
            if ($hadAny -or -not $hasFilter) {
                [void]$ParentNode.Items.Add($folderNode)
                $any = $true
            }
        }

        $childColls = if ($collectionsByFolder.ContainsKey($FolderID)) { @($collectionsByFolder[$FolderID] | Sort-Object Name) } else { @() }
        foreach ($c in $childColls) {
            if (-not (& $matchesNeedle $c)) { continue }
            $collNode = New-Object System.Windows.Controls.TreeViewItem
            $collNode.Header = ('{0}  ({1}, {2} members)' -f $c.Name, $c.CollectionID, $c.MemberCount)
            $collNode.Tag = @{ Type = 'Collection'; Object = $c }
            [void]$ParentNode.Items.Add($collNode)
            $script:__TreeLeafCount++
            $any = $true
        }
        return $any
    }

    & $populate $TreeView 0
    return $script:__TreeLeafCount
}

function Show-CollectionPickerDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        [string]$Title = 'Pick Collection',
        [bool]$IncludeBuiltIn = $true
    )

    if (-not $script:Collections -or @($script:Collections).Count -eq 0) {
        Add-LogLine 'Picker: refresh first to load collections.'
        return $null
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="720" Height="640" MinWidth="540" MinHeight="420"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBox x:Name="txtPickerFilter" Grid.Row="0" FontSize="12" Padding="6,4,6,4" Margin="0,4,0,8"
                 Controls:TextBoxHelper.Watermark="Filter by collection name or ID..."/>
        <Border Grid.Row="1" BorderThickness="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                Background="{DynamicResource MahApps.Brushes.ThemeBackground}">
            <TreeView x:Name="treePicker" FontSize="12"
                      VirtualizingStackPanel.IsVirtualizing="True"
                      VirtualizingStackPanel.VirtualizationMode="Recycling"
                      Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                      Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                      BorderThickness="0"/>
        </Border>
        <TextBlock x:Name="txtPickerStatus" Grid.Row="2" FontSize="11" Margin="0,8,0,0"
                   Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                   Text="Pick a collection (folders are not selectable)."/>
        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True" IsEnabled="False"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $txtPickerFilter = $dlg.FindName('txtPickerFilter')
    $treePicker      = $dlg.FindName('treePicker')
    $txtPickerStatus = $dlg.FindName('txtPickerStatus')
    $btnOk           = $dlg.FindName('btnOk')
    $btnCancel       = $dlg.FindName('btnCancel')

    $allCollections = if ($IncludeBuiltIn) { $script:Collections } else { $script:Collections | Where-Object { -not $_.IsBuiltIn } }
    $allCollections = @($allCollections)

    $rebuildTree = {
        param([string]$Needle)
        $count = Build-CollectionTree -TreeView $treePicker `
            -AllCollections $allCollections `
            -AllFolders     $script:Folders `
            -Needle         $Needle
        if ([string]::IsNullOrWhiteSpace($Needle)) {
            $totalColls = @($allCollections).Count
            $totalFolders = @($script:Folders).Count
            $txtPickerStatus.Text = ('{0} collections across {1} folders. Pick one (folders are not selectable).' -f $totalColls, $totalFolders)
        } else {
            $txtPickerStatus.Text = ('{0} collections match "{1}".' -f $count, $Needle.Trim())
        }
    }
    & $rebuildTree ''

    $script:PickerResult = $null
    $treePicker.Add_SelectedItemChanged({
        $node = $treePicker.SelectedItem
        if (-not $node -or -not $node.Tag -or $node.Tag.Type -ne 'Collection') {
            $btnOk.IsEnabled = $false
            $script:PickerResult = $null
            return
        }
        $btnOk.IsEnabled = $true
        $script:PickerResult = $node.Tag.Object
    })
    $txtPickerFilter.Add_TextChanged({ & $rebuildTree ([string]$txtPickerFilter.Text) })
    $btnOk.Add_Click({ if ($script:PickerResult) { $dlg.DialogResult = $true; $dlg.Close() } })
    $btnCancel.Add_Click({ $script:PickerResult = $null; $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:PickerResult
}

# =============================================================================
# Schedule Editor: inline reusable XAML fragment + read/write helpers.
# Embedded inside Show-NewWindowDialog and Show-TemplateEditorDialog.
# =============================================================================
$script:ScheduleEditorXamlFragment = @'
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
      xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
      xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro">
    <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <TextBlock Grid.Row="0" Text="Recurrence" FontSize="11" Margin="0,4,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
    <ComboBox x:Name="cboScheduleRecurrence" Grid.Row="1" FontSize="12">
        <ComboBoxItem Content="One-time"/>
        <ComboBoxItem Content="Daily"/>
        <ComboBoxItem Content="Weekly" IsSelected="True"/>
        <ComboBoxItem Content="MonthlyByDate"/>
        <ComboBoxItem Content="MonthlyByWeekday"/>
        <ComboBoxItem Content="PatchTuesday"/>
    </ComboBox>

    <Grid Grid.Row="2" Margin="0,12,0,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
            <TextBlock Text="Start Date" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <DatePicker x:Name="dpScheduleStartDate" FontSize="12" HorizontalAlignment="Left" Width="180"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="12,0,0,0">
            <TextBlock Text="Hour" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtScheduleHour" FontSize="12" Width="60" Padding="6,4,6,4"/>
        </StackPanel>
        <StackPanel Grid.Column="2" Margin="8,0,0,0">
            <TextBlock Text="Minute" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtScheduleMinute" FontSize="12" Width="60" Padding="6,4,6,4"/>
        </StackPanel>
    </Grid>

    <Grid Grid.Row="3" Margin="0,12,0,0">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0">
            <TextBlock Text="Duration Hours" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtDurationHours" FontSize="12" Width="60" Padding="6,4,6,4"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="8,0,0,0">
            <TextBlock Text="Duration Min" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtDurationMinutes" FontSize="12" Width="60" Padding="6,4,6,4"/>
        </StackPanel>
        <CheckBox x:Name="chkScheduleUtc" Grid.Column="2" Content="UTC time" VerticalAlignment="Bottom" Margin="16,0,0,4" FontSize="12"/>
    </Grid>

    <StackPanel x:Name="panelDayOfWeek" Grid.Row="4" Orientation="Horizontal" Margin="0,12,0,0" Visibility="Visible">
        <StackPanel>
            <TextBlock Text="Day of Week" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <ComboBox x:Name="cboDayOfWeek" FontSize="12" Width="140">
                <ComboBoxItem Content="Sunday"/>
                <ComboBoxItem Content="Monday"/>
                <ComboBoxItem Content="Tuesday" IsSelected="True"/>
                <ComboBoxItem Content="Wednesday"/>
                <ComboBoxItem Content="Thursday"/>
                <ComboBoxItem Content="Friday"/>
                <ComboBoxItem Content="Saturday"/>
            </ComboBox>
        </StackPanel>
    </StackPanel>

    <StackPanel x:Name="panelWeekOrder" Grid.Row="5" Margin="0,12,0,0" Visibility="Collapsed">
        <TextBlock Text="Week Order" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        <ComboBox x:Name="cboWeekOrder" FontSize="12" Width="140" HorizontalAlignment="Left">
            <ComboBoxItem Content="First"/>
            <ComboBoxItem Content="Second" IsSelected="True"/>
            <ComboBoxItem Content="Third"/>
            <ComboBoxItem Content="Fourth"/>
            <ComboBoxItem Content="Last"/>
        </ComboBox>
    </StackPanel>

    <StackPanel x:Name="panelMonthlyDate" Grid.Row="6" Margin="0,12,0,0" Visibility="Collapsed">
        <TextBlock Text="Day of Month (1-31)" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        <TextBox x:Name="txtDayOfMonth" FontSize="12" Width="60" Padding="6,4,6,4" HorizontalAlignment="Left"/>
    </StackPanel>

    <StackPanel x:Name="panelPatchTuesday" Grid.Row="7" Margin="0,12,0,0" Visibility="Collapsed">
        <TextBlock Text="Days After Patch Tuesday (0-30)" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        <TextBox x:Name="txtPatchOffset" FontSize="12" Width="60" Padding="6,4,6,4" HorizontalAlignment="Left"/>
    </StackPanel>

    <StackPanel Grid.Row="8" Margin="0,16,0,0">
        <TextBlock Text="Next 5 occurrences" FontSize="11" FontWeight="SemiBold" Margin="0,0,0,4"/>
        <Border BorderThickness="1" BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                Background="{DynamicResource MahApps.Brushes.ThemeBackground}" Padding="8" MinHeight="100">
            <TextBlock x:Name="txtSchedulePreview" FontSize="11" FontFamily="Cascadia Code, Consolas, Courier New"
                       Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                       TextWrapping="Wrap" Text="(set values to preview)"/>
        </Border>
    </StackPanel>
</Grid>
'@

function Get-ScheduleRecurrenceTypeInt {
    param([string]$Recurrence)
    switch ($Recurrence) {
        'One-time'         { return 1 }
        'Daily'            { return 2 }
        'Weekly'           { return 3 }
        'MonthlyByWeekday' { return 4 }
        'MonthlyByDate'    { return 5 }
        'PatchTuesday'     { return 4 }   # Patch Tuesday is monthly-by-weekday at the schedule layer
        default            { return 3 }
    }
}

function Update-ScheduleEditorVisibility {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Toggles in-window Visibility only.')]
    param([System.Windows.Window]$Dialog)
    if (-not $Dialog) { $Dialog = $script:__CurrentScheduleDialog }
    if (-not $Dialog) { return }
    $cbo = $Dialog.FindName('cboScheduleRecurrence')
    if (-not $cbo -or -not $cbo.SelectedItem) { return }
    $rec = [string]$cbo.SelectedItem.Content

    $panelDow = $Dialog.FindName('panelDayOfWeek')
    $panelWO  = $Dialog.FindName('panelWeekOrder')
    $panelMD  = $Dialog.FindName('panelMonthlyDate')
    $panelPT  = $Dialog.FindName('panelPatchTuesday')

    $vis = [System.Windows.Visibility]
    $panelDow.Visibility = if ($rec -in @('Weekly','MonthlyByWeekday')) { $vis::Visible } else { $vis::Collapsed }
    $panelWO.Visibility  = if ($rec -eq 'MonthlyByWeekday')             { $vis::Visible } else { $vis::Collapsed }
    $panelMD.Visibility  = if ($rec -eq 'MonthlyByDate')                { $vis::Visible } else { $vis::Collapsed }
    $panelPT.Visibility  = if ($rec -eq 'PatchTuesday')                 { $vis::Visible } else { $vis::Collapsed }
}

function Get-ScheduleEditorState {
    param([Parameter(Mandatory)][System.Windows.Window]$Dialog)
    $cboRec = $Dialog.FindName('cboScheduleRecurrence')
    $rec = if ($cboRec.SelectedItem) { [string]$cboRec.SelectedItem.Content } else { 'Weekly' }

    $cboDow = $Dialog.FindName('cboDayOfWeek')
    $dow = if ($cboDow.SelectedItem) { [string]$cboDow.SelectedItem.Content } else { 'Tuesday' }

    $cboWO = $Dialog.FindName('cboWeekOrder')
    $wo = if ($cboWO.SelectedItem) { [string]$cboWO.SelectedItem.Content } else { 'Second' }

    $hour     = [int]([string]$Dialog.FindName('txtScheduleHour').Text)
    $minute   = [int]([string]$Dialog.FindName('txtScheduleMinute').Text)
    $durHours = [int]([string]$Dialog.FindName('txtDurationHours').Text)
    $durMins  = [int]([string]$Dialog.FindName('txtDurationMinutes').Text)
    $isUtc    = [bool]$Dialog.FindName('chkScheduleUtc').IsChecked
    $startDate = $Dialog.FindName('dpScheduleStartDate').SelectedDate
    if (-not $startDate) { $startDate = (Get-Date).Date }

    $domTxt = [string]$Dialog.FindName('txtDayOfMonth').Text
    $dom = if ($domTxt -match '^\d+$') { [int]$domTxt } else { 1 }
    $offTxt = [string]$Dialog.FindName('txtPatchOffset').Text
    $off = if ($offTxt -match '^\d+$') { [int]$offTxt } else { 0 }

    return @{
        Recurrence       = $rec
        StartDate        = $startDate
        Hour             = $hour
        Minute           = $minute
        DurationHours    = $durHours
        DurationMinutes  = $durMins
        IsUtc            = $isUtc
        DayOfWeek        = $dow
        WeekOrder        = $wo
        DayOfMonth       = $dom
        PatchTuesdayOffset = $off
    }
}

function Set-ScheduleEditorState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Populates in-window controls only.')]
    param(
        [Parameter(Mandatory)][System.Windows.Window]$Dialog,
        [hashtable]$State = @{}
    )
    $defaults = @{
        Recurrence       = 'Weekly'
        StartDate        = (Get-Date).Date
        Hour             = 2
        Minute           = 0
        DurationHours    = 4
        DurationMinutes  = 0
        IsUtc            = $false
        DayOfWeek        = 'Tuesday'
        WeekOrder        = 'Second'
        DayOfMonth       = 1
        PatchTuesdayOffset = 0
    }
    foreach ($k in $defaults.Keys) {
        if (-not $State.ContainsKey($k)) { $State[$k] = $defaults[$k] }
    }

    $cboRec = $Dialog.FindName('cboScheduleRecurrence')
    foreach ($i in $cboRec.Items) { if ($i.Content -eq $State.Recurrence) { $cboRec.SelectedItem = $i; break } }

    $Dialog.FindName('dpScheduleStartDate').SelectedDate = [datetime]$State.StartDate
    $Dialog.FindName('txtScheduleHour').Text       = ([int]$State.Hour).ToString()
    $Dialog.FindName('txtScheduleMinute').Text     = ([int]$State.Minute).ToString()
    $Dialog.FindName('txtDurationHours').Text      = ([int]$State.DurationHours).ToString()
    $Dialog.FindName('txtDurationMinutes').Text    = ([int]$State.DurationMinutes).ToString()
    $Dialog.FindName('chkScheduleUtc').IsChecked   = [bool]$State.IsUtc

    $cboDow = $Dialog.FindName('cboDayOfWeek')
    foreach ($i in $cboDow.Items) { if ($i.Content -eq $State.DayOfWeek) { $cboDow.SelectedItem = $i; break } }
    $cboWO  = $Dialog.FindName('cboWeekOrder')
    foreach ($i in $cboWO.Items)  { if ($i.Content -eq $State.WeekOrder) { $cboWO.SelectedItem = $i; break } }

    $Dialog.FindName('txtDayOfMonth').Text  = ([int]$State.DayOfMonth).ToString()
    $Dialog.FindName('txtPatchOffset').Text = ([int]$State.PatchTuesdayOffset).ToString()

    Update-ScheduleEditorVisibility -Dialog $Dialog
}

function Update-SchedulePreview {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Updates the in-window preview TextBlock only.')]
    param([System.Windows.Window]$Dialog)
    if (-not $Dialog) { $Dialog = $script:__CurrentScheduleDialog }
    if (-not $Dialog) { return }
    $tx = $Dialog.FindName('txtSchedulePreview')
    if (-not $tx) { return }
    try {
        $st = Get-ScheduleEditorState -Dialog $Dialog
        $startDt = [datetime]::new($st.StartDate.Year, $st.StartDate.Month, $st.StartDate.Day, $st.Hour, $st.Minute, 0)

        # For PatchTuesday we shift the start date to the actual second-Tuesday + offset of this month.
        if ($st.Recurrence -eq 'PatchTuesday') {
            $pt = Get-PatchTuesday -Year $st.StartDate.Year -Month $st.StartDate.Month
            $startDt = [datetime]::new($pt.Year, $pt.Month, $pt.Day, $st.Hour, $st.Minute, 0).AddDays($st.PatchTuesdayOffset)
        }

        $rtInt = Get-ScheduleRecurrenceTypeInt -Recurrence $st.Recurrence
        $duration = ($st.DurationHours * 60) + $st.DurationMinutes
        $syn = [PSCustomObject]@{
            RecurrenceType = $rtInt
            StartTime      = $startDt
            Duration       = $duration
        }
        $next = @(Get-NextOccurrences -Window $syn -Count 5)
        if (@($next).Count -eq 0) {
            $tx.Text = '(no future occurrences -- check that the start date is in the future for one-time, or that recurrence is set)'
        } else {
            $tx.Text = (($next | ForEach-Object { $_.ToString('yyyy-MM-dd  ddd  HH:mm') }) -join [Environment]::NewLine)
        }
    } catch {
        $tx.Text = ('Preview unavailable: {0}' -f $_.Exception.Message)
    }
}

function Wire-ScheduleEditor {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Registers WPF event handlers for live preview.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Wire is the natural verb for hooking handlers.')]
    param([Parameter(Mandatory)][System.Windows.Window]$Dialog)

    # PS 5.1 scriptblock closures DO NOT reliably capture function-local
    # parameters across a WPF dispatcher boundary -- the handler fires later
    # in a different scope frame and $Dialog binds to $null. Stash the active
    # editor's dialog in script scope so the helpers can resolve it at fire
    # time. The Add_Closed clears the reference so a stale dialog never
    # leaks into the next modal.
    $script:__CurrentScheduleDialog = $Dialog
    $Dialog.Add_Closed({ $script:__CurrentScheduleDialog = $null })

    $Dialog.FindName('cboScheduleRecurrence').Add_SelectionChanged({
        Update-ScheduleEditorVisibility -Dialog $script:__CurrentScheduleDialog
        Update-SchedulePreview          -Dialog $script:__CurrentScheduleDialog
    })
    foreach ($n in @('dpScheduleStartDate')) {
        $Dialog.FindName($n).Add_SelectedDateChanged({ Update-SchedulePreview -Dialog $script:__CurrentScheduleDialog })
    }
    foreach ($n in @('cboDayOfWeek','cboWeekOrder')) {
        $Dialog.FindName($n).Add_SelectionChanged({ Update-SchedulePreview -Dialog $script:__CurrentScheduleDialog })
    }
    foreach ($n in @('txtScheduleHour','txtScheduleMinute','txtDurationHours','txtDurationMinutes','txtDayOfMonth','txtPatchOffset')) {
        $Dialog.FindName($n).Add_TextChanged({ Update-SchedulePreview -Dialog $script:__CurrentScheduleDialog })
    }
}

# =============================================================================
# New / Edit Maintenance Window dialog. Same XAML, mode-aware.
# =============================================================================
function Show-NewWindowDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        $Existing = $null         # optional: pass an AllWindows row to enter edit mode
    )

    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'New / Edit Window: refresh first to establish a CM connection.'
        return $false
    }

    $isEdit = $null -ne $Existing
    $titleText = if ($isEdit) { 'Edit Maintenance Window' } else { 'New Maintenance Window' }
    $okText    = if ($isEdit) { 'Save' } else { 'Create' }

    $dlgXaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="$titleText"
    Width="640" Height="780" MinWidth="540" MinHeight="700"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <Grid Margin="20,16,20,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <StackPanel Grid.Row="0">
                <TextBlock Text="Target Collection" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" BorderThickness="1"
                            BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                            Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                            Padding="8,5,8,5" Margin="0,0,8,0">
                        <TextBlock x:Name="txtTargetColl" FontSize="12" Text="(no collection selected)"
                                   Foreground="{DynamicResource MahApps.Brushes.Gray1}"
                                   VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                    </Border>
                    <Button x:Name="btnPickTargetColl" Grid.Column="1" Content="Browse..."
                            Style="{StaticResource DialogButton}" MinWidth="90" Margin="0"/>
                </Grid>
            </StackPanel>

            <StackPanel Grid.Row="1" Margin="0,12,0,0">
                <TextBlock Text="Window Name" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtWindowName" FontSize="12" Padding="6,4,6,4"/>
            </StackPanel>

            <StackPanel Grid.Row="2" Margin="0,12,0,0">
                <TextBlock Text="Window Type" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <ComboBox x:Name="cboWindowType" FontSize="12" Width="280" HorizontalAlignment="Left">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="SoftwareUpdatesOnly"/>
                    <ComboBoxItem Content="TaskSequencesOnly"/>
                </ComboBox>
            </StackPanel>

            <CheckBox x:Name="chkWindowEnabled" Grid.Row="3" Margin="0,12,0,0" Content="Enabled" IsChecked="True" FontSize="12"/>

            <TextBlock Grid.Row="4" Text="Schedule" FontSize="13" FontWeight="SemiBold" Margin="0,16,0,4"/>

            <ContentControl x:Name="scheduleHost" Grid.Row="5"/>

            <StackPanel Grid.Row="6" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                <Button x:Name="btnOk"     Content="$okText" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel"  Style="{StaticResource DialogButton}"        IsCancel="True"/>
            </StackPanel>
        </Grid>
    </ScrollViewer>
</Controls:MetroWindow>
"@

    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    # Inject the schedule editor fragment into the placeholder ContentControl.
    [xml]$scheduleXaml = $script:ScheduleEditorXamlFragment
    $sReader = New-Object System.Xml.XmlNodeReader $scheduleXaml
    $scheduleGrid = [System.Windows.Markup.XamlReader]::Load($sReader)
    $dlg.FindName('scheduleHost').Content = $scheduleGrid
    # FindName on the dialog after content injection: WPF NameScope does not
    # propagate from a dynamically-loaded subtree, so register names manually.
    $script:ScheduleNamedElements = @('cboScheduleRecurrence','dpScheduleStartDate','txtScheduleHour','txtScheduleMinute','txtDurationHours','txtDurationMinutes','chkScheduleUtc','panelDayOfWeek','cboDayOfWeek','panelWeekOrder','cboWeekOrder','panelMonthlyDate','txtDayOfMonth','panelPatchTuesday','txtPatchOffset','txtSchedulePreview')
    foreach ($n in $script:ScheduleNamedElements) {
        $el = $scheduleGrid.FindName($n)
        if ($el) { [void]$dlg.RegisterName($n, $el) }
    }

    $txtTargetColl    = $dlg.FindName('txtTargetColl')
    $btnPickTargetColl = $dlg.FindName('btnPickTargetColl')
    $txtWindowName    = $dlg.FindName('txtWindowName')
    $cboWindowType    = $dlg.FindName('cboWindowType')
    $chkWindowEnabled = $dlg.FindName('chkWindowEnabled')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    $script:NewWinTarget    = $null
    $script:NewWinOriginalName = $null
    if ($isEdit) {
        $script:NewWinTarget = $script:Collections | Where-Object { $_.CollectionID -eq $Existing.CollectionID } | Select-Object -First 1
        if ($script:NewWinTarget) {
            $txtTargetColl.Text = ('{0}  ({1})' -f $script:NewWinTarget.Name, $script:NewWinTarget.CollectionID)
        }
        $txtWindowName.Text = [string]$Existing.WindowName
        # Name IS editable in edit mode -- rename triggers a delete + recreate
        # because MECM keys maintenance windows by name. Original name kept so
        # we can detect the rename on Save.
        $script:NewWinOriginalName = [string]$Existing.WindowName
        $btnPickTargetColl.IsEnabled = $false
        $chkWindowEnabled.IsChecked = [bool]$Existing.IsEnabled
        # Best-effort populate from the existing window's properties.
        $startDt = if ($Existing.StartTime) { [datetime]$Existing.StartTime } else { (Get-Date) }
        $durMin  = [int]$Existing.DurationMinutes
        $durHrs  = [math]::Floor($durMin / 60)
        $durMin2 = $durMin % 60
        $rec = switch ([int]$Existing.RecurrenceType) {
            1 { 'One-time' }
            2 { 'Daily' }
            3 { 'Weekly' }
            4 { 'MonthlyByWeekday' }
            5 { 'MonthlyByDate' }
            default { 'Weekly' }
        }
        $existingType = switch ([string]$Existing.Type) {
            'Software Updates' { 'SoftwareUpdatesOnly' }
            'Task Sequences'   { 'TaskSequencesOnly' }
            default            { 'Any' }
        }
        foreach ($i in $cboWindowType.Items) { if ($i.Content -eq $existingType) { $cboWindowType.SelectedItem = $i; break } }

        Set-ScheduleEditorState -Dialog $dlg -State @{
            Recurrence      = $rec
            StartDate       = $startDt.Date
            Hour            = $startDt.Hour
            Minute          = $startDt.Minute
            DurationHours   = $durHrs
            DurationMinutes = $durMin2
            IsUtc           = [bool]$Existing.IsUTC
            DayOfWeek       = if ($startDt.DayOfWeek) { [string]$startDt.DayOfWeek } else { 'Tuesday' }
            DayOfMonth      = $startDt.Day
        }
    } else {
        Set-ScheduleEditorState -Dialog $dlg -State @{}
    }

    Wire-ScheduleEditor -Dialog $dlg
    Update-SchedulePreview -Dialog $dlg

    $btnPickTargetColl.Add_Click({
        $picked = Show-CollectionPickerDialog -Title 'Pick Target Collection' -IncludeBuiltIn $false
        if ($picked) {
            $script:NewWinTarget = $picked
            $txtTargetColl.Text = ('{0}  ({1})' -f $picked.Name, $picked.CollectionID)
        }
    })

    $script:NewWinResult = $false
    $btnOk.Add_Click({
        $coll = $script:NewWinTarget
        $name = ([string]$txtWindowName.Text).Trim()
        if (-not $coll) { Add-LogLine 'Validation: target collection required.'; return }
        if (-not $name) { Add-LogLine 'Validation: window name required.'; return }

        $st = Get-ScheduleEditorState -Dialog $dlg
        $startDt = [datetime]::new($st.StartDate.Year, $st.StartDate.Month, $st.StartDate.Day, $st.Hour, $st.Minute, 0)
        $applyTo = if ($cboWindowType.SelectedItem) { [string]$cboWindowType.SelectedItem.Content } else { 'Any' }
        $isEnabled = [bool]$chkWindowEnabled.IsChecked

        try {
            if ($st.Recurrence -eq 'PatchTuesday') {
                $schedule = New-PatchTuesdaySchedule -OffsetDays $st.PatchTuesdayOffset `
                    -StartHour $st.Hour -StartMinute $st.Minute `
                    -DurationHours $st.DurationHours -DurationMinutes $st.DurationMinutes
            } else {
                $schedule = New-WindowSchedule -RecurrenceType $st.Recurrence `
                    -StartTime $startDt `
                    -DurationHours $st.DurationHours -DurationMinutes $st.DurationMinutes `
                    -DayOfWeek ([System.DayOfWeek]$st.DayOfWeek) `
                    -WeekOrder $st.WeekOrder `
                    -DayOfMonth $st.DayOfMonth
            }
            if (-not $schedule) { Add-LogLine 'Schedule creation returned null.'; return }

            if ($isEdit) {
                $original = [string]$script:NewWinOriginalName
                if ($original -and $original -ne $name) {
                    # Rename: MECM keys windows by name, so a rename is delete + recreate.
                    Add-LogLine ('Renaming "{0}" -> "{1}" on {2} (delete + recreate)' -f $original, $name, $coll.Name)
                    $removed = Remove-ManagedMaintenanceWindow -CollectionId $coll.CollectionID -MaintenanceWindowName $original
                    if (-not $removed) { Add-LogLine 'Rename aborted: failed to remove the old window.'; return }
                    $created = New-ManagedMaintenanceWindow -CollectionId $coll.CollectionID -Name $name `
                        -Schedule $schedule -ApplyTo $applyTo -IsEnabled $isEnabled -IsUtc $st.IsUtc
                    if ($null -eq $created) { Add-LogLine 'Rename failed: old window deleted but new one did not create. Check log.'; return }
                    Add-LogLine ('Renamed and updated "{0}" on {1}' -f $name, $coll.Name)
                } else {
                    Set-ManagedMaintenanceWindow -CollectionId $coll.CollectionID -MaintenanceWindowName $name `
                        -Schedule $schedule -ApplyTo $applyTo -IsEnabled $isEnabled -IsUtc $st.IsUtc
                    Add-LogLine ('Updated maintenance window "{0}" on {1}' -f $name, $coll.Name)
                }
            } else {
                $result = New-ManagedMaintenanceWindow -CollectionId $coll.CollectionID -Name $name `
                    -Schedule $schedule -ApplyTo $applyTo -IsEnabled $isEnabled -IsUtc $st.IsUtc
                if ($null -eq $result) { Add-LogLine 'Create returned null -- check log for details.'; return }
                Add-LogLine ('Created maintenance window "{0}" on {1}' -f $name, $coll.Name)
            }
            $script:NewWinResult = $true
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('New / Edit Window failed: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:NewWinResult
}

# =============================================================================
# Template Editor dialog (New + Edit). Reuses the Schedule Editor fragment.
# =============================================================================
function Show-TemplateEditorDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        $Existing = $null
    )

    $isEdit = $null -ne $Existing
    $titleText = if ($isEdit) { 'Edit Template' } else { 'New Template' }
    $okText    = if ($isEdit) { 'Save' } else { 'Create' }

    $dlgXaml = @"
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="$titleText" Width="640" Height="780" MinWidth="540" MinHeight="700"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <Grid Margin="20,16,20,12">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <StackPanel Grid.Row="0">
                <TextBlock Text="Template Name" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtTemplateName" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. Patch Tuesday + 7"/>
            </StackPanel>

            <StackPanel Grid.Row="1" Margin="0,12,0,0">
                <TextBlock Text="Description" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtTemplateDesc" FontSize="12" Padding="6,4,6,4"
                         AcceptsReturn="True" TextWrapping="Wrap" Height="60"
                         VerticalScrollBarVisibility="Auto"/>
            </StackPanel>

            <StackPanel Grid.Row="2" Margin="0,12,0,0">
                <TextBlock Text="Window Type" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <ComboBox x:Name="cboTplWindowType" FontSize="12" Width="280" HorizontalAlignment="Left">
                    <ComboBoxItem Content="Any" IsSelected="True"/>
                    <ComboBoxItem Content="SoftwareUpdatesOnly"/>
                    <ComboBoxItem Content="TaskSequencesOnly"/>
                </ComboBox>
            </StackPanel>

            <TextBlock Grid.Row="3" Text="Schedule" FontSize="13" FontWeight="SemiBold" Margin="0,16,0,4"/>
            <ContentControl x:Name="scheduleHost" Grid.Row="4"/>

            <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                <Button x:Name="btnOk"     Content="$okText" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel"  Style="{StaticResource DialogButton}"        IsCancel="True"/>
            </StackPanel>
        </Grid>
    </ScrollViewer>
</Controls:MetroWindow>
"@

    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    [xml]$scheduleXaml = $script:ScheduleEditorXamlFragment
    $sReader = New-Object System.Xml.XmlNodeReader $scheduleXaml
    $scheduleGrid = [System.Windows.Markup.XamlReader]::Load($sReader)
    $dlg.FindName('scheduleHost').Content = $scheduleGrid
    foreach ($n in @('cboScheduleRecurrence','dpScheduleStartDate','txtScheduleHour','txtScheduleMinute','txtDurationHours','txtDurationMinutes','chkScheduleUtc','panelDayOfWeek','cboDayOfWeek','panelWeekOrder','cboWeekOrder','panelMonthlyDate','txtDayOfMonth','panelPatchTuesday','txtPatchOffset','txtSchedulePreview')) {
        $el = $scheduleGrid.FindName($n)
        if ($el) { [void]$dlg.RegisterName($n, $el) }
    }

    $txtTemplateName  = $dlg.FindName('txtTemplateName')
    $txtTemplateDesc  = $dlg.FindName('txtTemplateDesc')
    $cboTplWindowType = $dlg.FindName('cboTplWindowType')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    $script:TplEditOriginalPath = $null
    if ($isEdit) {
        $txtTemplateName.Text = [string]$Existing.Name
        $txtTemplateDesc.Text = [string]$Existing.Description
        $script:TplEditOriginalPath = [string]$Existing.FilePath
        $existingType = if ($Existing.WindowType) { [string]$Existing.WindowType } else { 'Any' }
        foreach ($i in $cboTplWindowType.Items) { if ($i.Content -eq $existingType) { $cboTplWindowType.SelectedItem = $i; break } }

        Set-ScheduleEditorState -Dialog $dlg -State @{
            Recurrence       = [string]$Existing.RecurrenceType
            StartDate        = (Get-Date).Date
            Hour             = [int]$Existing.StartHour
            Minute           = [int]$Existing.StartMinute
            DurationHours    = [int]$Existing.DurationHours
            DurationMinutes  = [int]$Existing.DurationMinutes
            IsUtc            = [bool]$Existing.IsUtc
            DayOfWeek        = [string]$Existing.DayOfWeek
            WeekOrder        = [string]$Existing.WeekOrder
            DayOfMonth       = [int]$Existing.DayOfMonth
            PatchTuesdayOffset = if ([int]$Existing.PatchTuesdayOffset -ge 0) { [int]$Existing.PatchTuesdayOffset } else { 0 }
        }
    } else {
        Set-ScheduleEditorState -Dialog $dlg -State @{}
    }

    Wire-ScheduleEditor -Dialog $dlg
    Update-SchedulePreview -Dialog $dlg

    $script:TplEditResult = $false
    $btnOk.Add_Click({
        $name = ([string]$txtTemplateName.Text).Trim()
        if (-not $name) { Add-LogLine 'Template: name required.'; return }
        $st = Get-ScheduleEditorState -Dialog $dlg
        $tplPath = Join-Path $PSScriptRoot 'Templates'
        $winType = if ($cboTplWindowType.SelectedItem) { [string]$cboTplWindowType.SelectedItem.Content } else { 'Any' }

        try {
            $params = @{
                TemplatesPath      = $tplPath
                Name               = $name
                Description        = ([string]$txtTemplateDesc.Text)
                RecurrenceType     = $st.Recurrence
                WindowType         = $winType
                DurationHours      = $st.DurationHours
                DurationMinutes    = $st.DurationMinutes
                StartHour          = $st.Hour
                StartMinute        = $st.Minute
                IsUtc              = $st.IsUtc
                DayOfWeek          = ([System.DayOfWeek]$st.DayOfWeek)
                WeekOrder          = $st.WeekOrder
                DayOfMonth         = $st.DayOfMonth
                PatchTuesdayOffset = if ($st.Recurrence -eq 'PatchTuesday') { $st.PatchTuesdayOffset } else { -1 }
            }
            $newPath = Save-WindowTemplate @params

            # If editing and the slugified file name changed, drop the original.
            if ($script:TplEditOriginalPath -and $script:TplEditOriginalPath -ne $newPath) {
                if (Test-Path -LiteralPath $script:TplEditOriginalPath) {
                    Remove-WindowTemplate -FilePath $script:TplEditOriginalPath | Out-Null
                    Add-LogLine ('Template: removed old file after rename ({0})' -f ([System.IO.Path]::GetFileName($script:TplEditOriginalPath)))
                }
            }

            Add-LogLine ('Template saved: {0} ({1})' -f $name, ([System.IO.Path]::GetFileName($newPath)))
            $script:TplEditResult = $true
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('Save Template failed: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:TplEditResult
}

# =============================================================================
# Apply Template dialog. Single from Templates view, bulk from Coverage view.
# =============================================================================
function Show-ApplyTemplateDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param(
        $InitialTemplate = $null,
        $InitialTargets  = @()
    )

    if (-not $script:IsConnectedFromBg) {
        Add-LogLine 'Apply Template: refresh first to establish a CM connection.'
        return $false
    }
    if (-not $script:Templates -or @($script:Templates).Count -eq 0) {
        Add-LogLine 'Apply Template: no templates loaded. Reload templates first.'
        return $false
    }

    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Apply Template" Width="640" Height="640" MinWidth="540" MinHeight="540"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="20,16,20,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0">
            <TextBlock Text="Template" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <ComboBox x:Name="cboApplyTemplate" FontSize="12" DisplayMemberPath="Name"/>
        </StackPanel>

        <StackPanel Grid.Row="1" Margin="0,12,0,0">
            <TextBlock Text="Window Name (override -- blank uses template name)" FontSize="11" Margin="0,0,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            <TextBox x:Name="txtApplyWindowName" FontSize="12" Padding="6,4,6,4"
                     Controls:TextBoxHelper.Watermark="(blank = use template name)"/>
        </StackPanel>

        <TextBlock Grid.Row="2" Text="Target Collections" FontSize="11" FontWeight="SemiBold" Margin="0,16,0,4"/>
        <Border Grid.Row="3" BorderThickness="1"
                BorderBrush="{DynamicResource MahApps.Brushes.Gray8}"
                Background="{DynamicResource MahApps.Brushes.ThemeBackground}">
            <ListBox x:Name="lstApplyTargets" FontSize="12" BorderThickness="0"
                     Background="{DynamicResource MahApps.Brushes.ThemeBackground}"
                     Foreground="{DynamicResource MahApps.Brushes.ThemeForeground}"
                     SelectionMode="Extended"/>
        </Border>

        <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,8,0,0">
            <Button x:Name="btnAddTarget"    Content="Add Target..."   Style="{StaticResource DialogButton}" MinWidth="130"/>
            <Button x:Name="btnRemoveTarget" Content="Remove Selected" Style="{StaticResource DialogButton}" MinWidth="140"/>
            <TextBlock x:Name="txtTargetCount" FontSize="11" VerticalAlignment="Center" Margin="8,0,0,0"
                       Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
        </StackPanel>

        <StackPanel Grid.Row="5" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
            <Button x:Name="btnOk"     Content="Apply"  Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $cboApplyTemplate  = $dlg.FindName('cboApplyTemplate')
    $txtApplyWindowName = $dlg.FindName('txtApplyWindowName')
    $lstApplyTargets   = $dlg.FindName('lstApplyTargets')
    $btnAddTarget      = $dlg.FindName('btnAddTarget')
    $btnRemoveTarget   = $dlg.FindName('btnRemoveTarget')
    $txtTargetCount    = $dlg.FindName('txtTargetCount')
    $btnOk             = $dlg.FindName('btnOk')
    $btnCancel         = $dlg.FindName('btnCancel')

    $cboApplyTemplate.ItemsSource = $script:Templates
    if ($InitialTemplate) {
        $match = $script:Templates | Where-Object { $_.FilePath -eq $InitialTemplate.FilePath } | Select-Object -First 1
        if ($match) { $cboApplyTemplate.SelectedItem = $match }
    } elseif (@($script:Templates).Count -gt 0) {
        $cboApplyTemplate.SelectedIndex = 0
    }

    $script:ApplyTargets = New-Object System.Collections.ObjectModel.ObservableCollection[psobject]
    foreach ($t in @($InitialTargets)) {
        # Map from Coverage-row PSCustomObject (Name + CollectionID) to a real Collection
        $resolved = $script:Collections | Where-Object { $_.CollectionID -eq $t.CollectionID } | Select-Object -First 1
        if ($resolved -and -not $resolved.IsBuiltIn) { $script:ApplyTargets.Add($resolved) }
    }
    $lstApplyTargets.ItemsSource = $script:ApplyTargets
    $lstApplyTargets.DisplayMemberPath = 'Name'
    $txtTargetCount.Text = ('{0} target(s)' -f $script:ApplyTargets.Count)

    $btnAddTarget.Add_Click({
        $picked = Show-CollectionPickerDialog -Title 'Add Target Collection' -IncludeBuiltIn $false
        if ($picked) {
            $already = $script:ApplyTargets | Where-Object { $_.CollectionID -eq $picked.CollectionID }
            if (-not $already) { $script:ApplyTargets.Add($picked) }
            $txtTargetCount.Text = ('{0} target(s)' -f $script:ApplyTargets.Count)
        }
    })
    $btnRemoveTarget.Add_Click({
        $sel = @($lstApplyTargets.SelectedItems)
        foreach ($s in $sel) { [void]$script:ApplyTargets.Remove($s) }
        $txtTargetCount.Text = ('{0} target(s)' -f $script:ApplyTargets.Count)
    })

    $script:ApplyResult = $false
    $btnOk.Add_Click({
        $tpl = $cboApplyTemplate.SelectedItem
        if (-not $tpl) { Add-LogLine 'Apply Template: pick a template.'; return }
        if ($script:ApplyTargets.Count -eq 0) { Add-LogLine 'Apply Template: pick at least one target collection.'; return }
        $override = ([string]$txtApplyWindowName.Text).Trim()

        try {
            $applyParams = @{
                Template          = $tpl
                TargetCollections = @($script:ApplyTargets)
            }
            if ($override) { $applyParams['WindowNameOverride'] = $override }
            $results = @(Apply-WindowTemplate @applyParams)
            $okCount  = @($results | Where-Object { $_.Success }).Count
            $failCount = @($results | Where-Object { -not $_.Success }).Count
            Add-LogLine ('Apply Template "{0}": {1} succeeded, {2} failed.' -f $tpl.Name, $okCount, $failCount)
            foreach ($r in $results | Where-Object { -not $_.Success }) {
                Add-LogLine ('  failed on {0}: {1}' -f $r.CollectionName, $r.Error)
            }
            $script:ApplyResult = ($okCount -gt 0)
            $dlg.DialogResult = $true
            $dlg.Close()
        } catch {
            Add-LogLine ('Apply Template threw: {0}' -f $_.Exception.Message)
        }
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })

    [void]$dlg.ShowDialog()
    return $script:ApplyResult
}

# =============================================================================
# Wired action button handlers.
# =============================================================================
$btnNewWindow.Add_Click({
    if (Show-NewWindowDialog) { Invoke-Refresh }
})
$btnEditWindow.Add_Click({
    $row = $gridWindows.SelectedItem
    if (-not $row) { Add-LogLine 'Edit: pick a window first.'; return }
    if (Show-NewWindowDialog -Existing $row) { Invoke-Refresh }
})
$btnRemoveWindow.Add_Click({
    $row = $gridWindows.SelectedItem
    if (-not $row) { Add-LogLine 'Remove: pick a window first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Remove: refresh first to establish a CM connection.'; return }
    if (-not (Show-ConfirmDialog -Title 'Remove Maintenance Window' -Message ("Permanently delete window '$($row.WindowName)' from $($row.CollectionName)?"))) {
        return
    }
    try {
        $ok = Remove-ManagedMaintenanceWindow -CollectionId $row.CollectionID -MaintenanceWindowName $row.WindowName
        if ($ok) {
            Add-LogLine ('Removed window "{0}" from {1}' -f $row.WindowName, $row.CollectionName)
            Invoke-Refresh
        }
    } catch {
        Add-LogLine ('Remove failed: {0}' -f $_.Exception.Message)
    }
})
$btnApplyTemplateBulk.Add_Click({
    $selected = @($gridCoverage.SelectedItems)
    if ($selected.Count -eq 0) { Add-LogLine 'Apply Template (bulk): multi-select rows on the Coverage view first.'; return }
    if (Show-ApplyTemplateDialog -InitialTargets $selected) { Invoke-Refresh }
})
$btnNewTemplate.Add_Click({
    if (Show-TemplateEditorDialog) { Invoke-LoadTemplates }
})
$btnEditTemplate.Add_Click({
    $row = $gridTemplates.SelectedItem
    if (-not $row) { Add-LogLine 'Edit Template: pick a template first.'; return }
    if (Show-TemplateEditorDialog -Existing $row) { Invoke-LoadTemplates }
})
$btnDeleteTemplate.Add_Click({
    $row = $gridTemplates.SelectedItem
    if (-not $row) { Add-LogLine 'Delete Template: pick a template first.'; return }
    if (-not (Show-ConfirmDialog -Title 'Delete Template' -Message ("Permanently delete template '$($row.Name)' ($([System.IO.Path]::GetFileName($row.FilePath)))?"))) {
        return
    }
    try {
        $ok = Remove-WindowTemplate -FilePath $row.FilePath
        if ($ok) {
            Add-LogLine ('Deleted template: {0}' -f $row.Name)
            Invoke-LoadTemplates
        }
    } catch {
        Add-LogLine ('Delete Template failed: {0}' -f $_.Exception.Message)
    }
})
$btnApplyTemplateOne.Add_Click({
    $tpl = $gridTemplates.SelectedItem
    if (-not $tpl) { Add-LogLine 'Apply Template: pick a template first.'; return }
    if (Show-ApplyTemplateDialog -InitialTemplate $tpl) { Invoke-Refresh }
})

$btnToggleEnabled.Add_Click({
    $row = $gridWindows.SelectedItem
    if (-not $row) { Add-LogLine 'Toggle: pick a window first.'; return }
    if (-not $script:IsConnectedFromBg) { Add-LogLine 'Toggle: refresh first to establish a CM connection.'; return }
    $newState = -not $row.IsEnabled
    try {
        Set-ManagedMaintenanceWindow -CollectionId $row.CollectionID -MaintenanceWindowName $row.WindowName -IsEnabled $newState
        Add-LogLine ('Toggled "{0}" on {1} -> {2}' -f $row.WindowName, $row.CollectionName, $(if ($newState) { 'Enabled' } else { 'Disabled' }))
        Invoke-Refresh
    } catch {
        Add-LogLine ('Toggle failed: {0}' -f $_.Exception.Message)
    }
})

# =============================================================================
# Export buttons.
# =============================================================================
$btnExportCsv.Add_Click({
    if (@($script:RawWindows).Count -eq 0) { Add-LogLine 'Export CSV: nothing to export.'; return }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'CSV files (*.csv)|*.csv'
    $sfd.FileName = ('MWM-Windows-{0}.csv' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        Export-MaintenanceWindowsCsv -Windows $script:RawWindows -OutputPath $sfd.FileName
        Add-LogLine ('Exported CSV: {0}' -f $sfd.FileName)
    }
})

$btnExportHtml.Add_Click({
    if (@($script:RawWindows).Count -eq 0) { Add-LogLine 'Export HTML: nothing to export.'; return }
    $sfd = New-Object Microsoft.Win32.SaveFileDialog
    $sfd.Filter = 'HTML files (*.html)|*.html'
    $sfd.FileName = ('MWM-Windows-{0}.html' -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $reportsDir = Join-Path $PSScriptRoot 'Reports'
    if (-not (Test-Path -LiteralPath $reportsDir)) { New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null }
    $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq $true) {
        Export-MaintenanceWindowsHtml -Windows $script:RawWindows -OutputPath $sfd.FileName
        Add-LogLine ('Exported HTML: {0}' -f $sfd.FileName)
    }
})

# =============================================================================
# Themed Yes/No confirmation dialog.
# =============================================================================
function Set-DialogTheme {
    param([Parameter(Mandatory)][System.Windows.Window]$Dialog)
    $isDark = [bool]$global:Prefs['DarkMode']
    if ($isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Dark.Steel')
    } else {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($Dialog, 'Light.Blue')
        $Dialog.WindowTitleBrush          = $script:TitleBarBlue
        $Dialog.NonActiveWindowTitleBrush = $script:TitleBarBlueInactive
    }
}

function Show-ConfirmDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Message)
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="" Width="480" SizeToContent="Height" MinWidth="380"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ResizeMode="NoResize" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid Margin="16,12,16,12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="txtMsg" Grid.Row="0" TextWrapping="Wrap" FontSize="13" Margin="0,8,0,16"/>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnYes" Content="Yes" Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
            <Button x:Name="btnNo"  Content="No"  Style="{StaticResource DialogButton}"        IsCancel="True"/>
        </StackPanel>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    $dlg.Title = $Title
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg
    $dlg.FindName('txtMsg').Text = $Message
    $dlg.FindName('btnYes').Add_Click({ $dlg.DialogResult = $true;  $dlg.Close() })
    $dlg.FindName('btnNo').Add_Click({  $dlg.DialogResult = $false; $dlg.Close() })
    return [bool]$dlg.ShowDialog()
}

# =============================================================================
# Options dialog.
# =============================================================================
function Show-OptionsDialog {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseSingularNouns', '', Justification='Modal dialog show / dispose; reads as a single action.')]
    param()
    $dlgXaml = @'
<Controls:MetroWindow
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
    Title="Options" Width="640" Height="380"
    MinWidth="560" MinHeight="380"
    WindowStartupLocation="CenterOwner" TitleCharacterCasing="Normal"
    GlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    NonActiveGlowBrush="{DynamicResource MahApps.Brushes.Accent}"
    BorderThickness="1" ShowIconOnTitleBar="False">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Themes/Dark.Steel.xaml" />
            </ResourceDictionary.MergedDictionaries>
            <Style x:Key="CategoryRowStyle" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="Height" Value="36"/>
                <Setter Property="HorizontalContentAlignment" Value="Left"/>
                <Setter Property="Padding" Value="14,0,14,0"/>
                <Setter Property="FontSize" Value="13"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
                <Setter Property="Margin" Value="0"/>
            </Style>
            <Style x:Key="DialogButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
            <Style x:Key="DialogAccentButton" TargetType="Button" BasedOn="{StaticResource MahApps.Styles.Button.Square.Accent}">
                <Setter Property="MinWidth" Value="90"/><Setter Property="Height" Value="32"/>
                <Setter Property="Margin" Value="0,0,8,0"/>
                <Setter Property="Controls:ControlsHelper.ContentCharacterCasing" Value="Normal"/>
            </Style>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="180"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Border Grid.Column="0" Grid.Row="0" Padding="6,12,0,12">
            <StackPanel>
                <Button x:Name="btnCatConnection" Content="Connection" Style="{StaticResource CategoryRowStyle}"/>
                <Button x:Name="btnCatAbout"      Content="About"      Style="{StaticResource CategoryRowStyle}"/>
            </StackPanel>
        </Border>
        <Border Grid.Column="1" Grid.Row="0" Background="{DynamicResource MahApps.Brushes.Gray8}"/>
        <Grid Grid.Column="2" Grid.Row="0" Margin="20,16,20,16">
            <StackPanel x:Name="paneConnection" Visibility="Visible">
                <TextBlock Text="MECM Connection" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="Site Code" FontSize="11" Margin="0,4,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSiteCode" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. MCM" Width="120" HorizontalAlignment="Left"/>
                <TextBlock Text="SMS Provider FQDN" FontSize="11" Margin="0,12,0,2" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
                <TextBox x:Name="txtSmsProvider" FontSize="12" Padding="6,4,6,4"
                         Controls:TextBoxHelper.Watermark="e.g. cm01.contoso.com"/>
                <TextBlock Text="Used for the CM PSDrive root. Maintenance Window Manager performs both reads (Get-CMMaintenanceWindow) and writes (New / Set / Remove) -- the account running this app needs the matching CM permissions."
                           FontSize="11" TextWrapping="Wrap" Margin="0,16,0,0"
                           Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
            <StackPanel x:Name="paneAbout" Visibility="Collapsed">
                <TextBlock Text="About" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,10"/>
                <TextBlock Text="Maintenance Window Manager v1.0.0" FontSize="13" FontWeight="SemiBold"/>
                <TextBlock Text="Browse, create, edit, toggle, and bulk-apply MECM maintenance windows across every device collection. Schedule editor supports One-time / Daily / Weekly / Monthly-by-Date / Monthly-by-Weekday / Patch Tuesday +N days, with a live next-5-occurrences preview."
                           FontSize="12" TextWrapping="Wrap" Margin="0,8,0,0"/>
                <TextBlock Text="Author: Jason Ulbright. License: MIT."
                           FontSize="11" Margin="0,16,0,0" Foreground="{DynamicResource MahApps.Brushes.Gray1}"/>
            </StackPanel>
        </Grid>
        <Border Grid.Row="1" Grid.ColumnSpan="3" Padding="16,12,16,12">
            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="btnOk"     Content="OK"     Style="{StaticResource DialogAccentButton}" IsDefault="True"/>
                <Button x:Name="btnCancel" Content="Cancel" Style="{StaticResource DialogButton}"        IsCancel="True"/>
            </StackPanel>
        </Border>
    </Grid>
</Controls:MetroWindow>
'@
    [xml]$dx = $dlgXaml
    $reader2 = New-Object System.Xml.XmlNodeReader $dx
    $dlg = [System.Windows.Markup.XamlReader]::Load($reader2)
    $dlg.Owner = $window
    Install-TitleBarDragFallback -Window $dlg
    Set-DialogTheme -Dialog $dlg

    $btnCatConnection = $dlg.FindName('btnCatConnection')
    $btnCatAbout      = $dlg.FindName('btnCatAbout')
    $paneConnection   = $dlg.FindName('paneConnection')
    $paneAbout        = $dlg.FindName('paneAbout')
    $txtSiteCode      = $dlg.FindName('txtSiteCode')
    $txtSmsProvider   = $dlg.FindName('txtSmsProvider')
    $btnOk            = $dlg.FindName('btnOk')
    $btnCancel        = $dlg.FindName('btnCancel')

    $txtSiteCode.Text    = [string]$global:Prefs.SiteCode
    $txtSmsProvider.Text = [string]$global:Prefs.SMSProvider

    $btnCatConnection.Add_Click({ $paneConnection.Visibility = [System.Windows.Visibility]::Visible; $paneAbout.Visibility = [System.Windows.Visibility]::Collapsed })
    $btnCatAbout.Add_Click({      $paneConnection.Visibility = [System.Windows.Visibility]::Collapsed; $paneAbout.Visibility = [System.Windows.Visibility]::Visible })

    $btnOk.Add_Click({
        $newSite     = ([string]$txtSiteCode.Text).Trim()
        $newProvider = ([string]$txtSmsProvider.Text).Trim()
        $changed = ($newSite -ne [string]$global:Prefs.SiteCode) -or ($newProvider -ne [string]$global:Prefs.SMSProvider)
        $global:Prefs.SiteCode    = $newSite
        $global:Prefs.SMSProvider = $newProvider
        Save-MwmPreferences -Prefs $global:Prefs
        if ($changed) {
            Dispose-BgWork
            if ($script:BgRunspace) {
                try { $script:BgRunspace.Close() }   catch { $null = $_ }
                try { $script:BgRunspace.Dispose() } catch { $null = $_ }
                $script:BgRunspace = $null
            }
            $script:BgState           = $null
            $script:IsConnectedFromBg = $false
            $progressOverlay.Visibility = [System.Windows.Visibility]::Collapsed
            $btnRefresh.IsEnabled       = $true
        }
        $dlg.DialogResult = $true; $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = $false; $dlg.Close() })
    [void]$dlg.ShowDialog()
    Update-StatusBarSummary
}

$btnOptions.Add_Click({ Show-OptionsDialog })

# =============================================================================
# Window state persistence.
# =============================================================================
$global:WindowStatePath = Join-Path $PSScriptRoot 'MaintWindowMgr.windowstate.json'

function Save-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Writes a small JSON state file; idempotent.')]
    param()
    try {
        $state = @{
            Left       = [int]$window.Left
            Top        = [int]$window.Top
            Width      = [int]$window.Width
            Height     = [int]$window.Height
            Maximized  = ($window.WindowState -eq [System.Windows.WindowState]::Maximized)
            ActiveView = $script:ActiveView
        }
        $state | ConvertTo-Json | Set-Content -LiteralPath $global:WindowStatePath -Encoding UTF8
    } catch { $null = $_ }
}

function Restore-WindowState {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification='Reads the JSON state file and applies geometry; idempotent.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseApprovedVerbs', '', Justification='Restore is intentional and reads as a single action.')]
    param()
    if (-not (Test-Path -LiteralPath $global:WindowStatePath)) { return }
    try {
        $s = Get-Content -LiteralPath $global:WindowStatePath -Raw | ConvertFrom-Json -ErrorAction Stop
        $left = if ($null -ne $s.Left) { [int]$s.Left } elseif ($null -ne $s.X) { [int]$s.X } else { $null }
        $top  = if ($null -ne $s.Top)  { [int]$s.Top  } elseif ($null -ne $s.Y) { [int]$s.Y } else { $null }
        $w    = if ($null -ne $s.Width)  { [int]$s.Width  } else { $null }
        $h    = if ($null -ne $s.Height) { [int]$s.Height } else { $null }

        if ($s.Maximized) {
            $window.WindowState = [System.Windows.WindowState]::Maximized
        } elseif ($null -ne $left -and $null -ne $top -and $null -ne $w -and $null -ne $h) {
            $screen = [System.Windows.Forms.Screen]::FromPoint([System.Drawing.Point]::new($left, $top))
            $bounds = $screen.WorkingArea
            $left = [Math]::Max($bounds.X, [Math]::Min($left, $bounds.Right - 200))
            $top  = [Math]::Max($bounds.Y, [Math]::Min($top,  $bounds.Bottom - 100))
            $window.Left   = $left
            $window.Top    = $top
            $window.Width  = [Math]::Max($window.MinWidth,  $w)
            $window.Height = [Math]::Max($window.MinHeight, $h)
        }
        if ($s.ActiveView -in @('Windows','Coverage','Templates')) {
            Set-ActiveView -View ([string]$s.ActiveView)
        }
    } catch { $null = $_ }
}

$window.Add_Closing({
    Save-WindowState
    Dispose-BgWork
    if ($script:BgRunspace) {
        try { $script:BgRunspace.Close() }  catch { $null = $_ }
        try { $script:BgRunspace.Dispose() } catch { $null = $_ }
    }
})

$window.Add_Loaded({
    Restore-WindowState
    $isDark = [bool]$global:Prefs['DarkMode']
    if (-not $isDark) {
        [void][ControlzEx.Theming.ThemeManager]::Current.ChangeTheme($window, 'Light.Blue')
    }
    Update-TitleBarBrushes
    Update-ActionBarVisibility
    Update-StatusBarSummary
    Add-LogLine 'Maintenance Window Manager ready. Configure Site / Provider in Options, then click Refresh.'
    Invoke-LoadTemplates
})

[void]$window.ShowDialog()
try { Stop-Transcript | Out-Null } catch { $null = $_ }
