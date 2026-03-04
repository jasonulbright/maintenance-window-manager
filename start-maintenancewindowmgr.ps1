<#
.SYNOPSIS
    WinForms front-end for MECM Maintenance Window Manager.

.DESCRIPTION
    Provides a GUI for viewing, creating, editing, and bulk-managing
    maintenance windows across all MECM device collections.

.EXAMPLE
    .\start-maintenancewindowmgr.ps1

.NOTES
    Requirements:
      - PowerShell 5.1
      - .NET Framework 4.8+
      - Windows Forms (System.Windows.Forms)
      - Configuration Manager console installed

    ScriptName : start-maintenancewindowmgr.ps1
    Purpose    : WinForms front-end for MECM maintenance window management
    Version    : 1.0.0
    Updated    : 2026-03-03
#>

param()

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { }

$moduleRoot = Join-Path $PSScriptRoot "Module"
Import-Module (Join-Path $moduleRoot "MaintWindowMgrCommon.psd1") -Force -DisableNameChecking

$toolLogFolder = Join-Path $PSScriptRoot "Logs"
if (-not (Test-Path -LiteralPath $toolLogFolder)) {
    New-Item -ItemType Directory -Path $toolLogFolder -Force | Out-Null
}
$toolLogPath = Join-Path $toolLogFolder ("MaintWinMgr-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
Initialize-Logging -LogPath $toolLogPath

$reportsDir = Join-Path $PSScriptRoot "Reports"
if (-not (Test-Path -LiteralPath $reportsDir)) {
    New-Item -ItemType Directory -Path $reportsDir -Force | Out-Null
}
$templatesDir = Join-Path $PSScriptRoot "Templates"

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Set-ModernButtonStyle {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Button]$Button,
        [Parameter(Mandatory)][System.Drawing.Color]$BackColor
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.BackColor = $BackColor
    $Button.ForeColor = [System.Drawing.Color]::White
    $Button.UseVisualStyleBackColor = $false
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $hover = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 18), [Math]::Max(0, $BackColor.G - 18), [Math]::Max(0, $BackColor.B - 18))
    $down  = [System.Drawing.Color]::FromArgb([Math]::Max(0, $BackColor.R - 36), [Math]::Max(0, $BackColor.G - 36), [Math]::Max(0, $BackColor.B - 36))
    $Button.FlatAppearance.MouseOverBackColor = $hover
    $Button.FlatAppearance.MouseDownBackColor = $down
}

function Enable-DoubleBuffer {
    param([Parameter(Mandatory)][System.Windows.Forms.Control]$Control)
    $prop = $Control.GetType().GetProperty("DoubleBuffered", [System.Reflection.BindingFlags] "Instance,NonPublic")
    if ($prop) { $prop.SetValue($Control, $true, $null) | Out-Null }
}

function Add-LogLine {
    param([Parameter(Mandatory)][System.Windows.Forms.TextBox]$TextBox, [Parameter(Mandatory)][string]$Message)
    $ts = (Get-Date).ToString("HH:mm:ss")
    $line = "{0}  {1}" -f $ts, $Message
    if ([string]::IsNullOrWhiteSpace($TextBox.Text)) { $TextBox.Text = $line }
    else { $TextBox.AppendText([Environment]::NewLine + $line) }
    $TextBox.SelectionStart = $TextBox.TextLength
    $TextBox.ScrollToCaret()
}

# ---------------------------------------------------------------------------
# Window state persistence
# ---------------------------------------------------------------------------

function Save-WindowState {
    $statePath = Join-Path $PSScriptRoot "MaintWindowMgr.windowstate.json"
    $state = @{
        X = $form.Location.X; Y = $form.Location.Y
        Width = $form.Size.Width; Height = $form.Size.Height
        Maximized = ($form.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized)
        ActiveTab = $tabMain.SelectedIndex
    }
    $state | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Restore-WindowState {
    $statePath = Join-Path $PSScriptRoot "MaintWindowMgr.windowstate.json"
    if (-not (Test-Path -LiteralPath $statePath)) { return }
    try {
        $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
        if ($state.Maximized) { $form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
        else {
            $screen = [System.Windows.Forms.Screen]::FromPoint((New-Object System.Drawing.Point($state.X, $state.Y)))
            $bounds = $screen.WorkingArea
            $x = [Math]::Max($bounds.X, [Math]::Min($state.X, $bounds.Right - 200))
            $y = [Math]::Max($bounds.Y, [Math]::Min($state.Y, $bounds.Bottom - 100))
            $form.Location = New-Object System.Drawing.Point($x, $y)
            $form.Size = New-Object System.Drawing.Size([Math]::Max($form.MinimumSize.Width, $state.Width), [Math]::Max($form.MinimumSize.Height, $state.Height))
        }
        if ($null -ne $state.ActiveTab -and $state.ActiveTab -ge 0 -and $state.ActiveTab -lt $tabMain.TabCount) {
            $tabMain.SelectedIndex = [int]$state.ActiveTab
        }
    } catch { }
}

# ---------------------------------------------------------------------------
# Preferences
# ---------------------------------------------------------------------------

function Get-MwPreferences {
    $prefsPath = Join-Path $PSScriptRoot "MaintWindowMgr.prefs.json"
    $defaults = @{ DarkMode = $false; SiteCode = ''; SMSProvider = '' }
    if (Test-Path -LiteralPath $prefsPath) {
        try {
            $loaded = Get-Content -LiteralPath $prefsPath -Raw | ConvertFrom-Json
            if ($null -ne $loaded.DarkMode) { $defaults.DarkMode = [bool]$loaded.DarkMode }
            if ($loaded.SiteCode)           { $defaults.SiteCode = $loaded.SiteCode }
            if ($loaded.SMSProvider)         { $defaults.SMSProvider = $loaded.SMSProvider }
        } catch { }
    }
    return $defaults
}

function Save-MwPreferences {
    param([hashtable]$Prefs)
    $prefsPath = Join-Path $PSScriptRoot "MaintWindowMgr.prefs.json"
    $Prefs | ConvertTo-Json | Set-Content -LiteralPath $prefsPath -Encoding UTF8
}

$script:Prefs = Get-MwPreferences

# ---------------------------------------------------------------------------
# Theme colors
# ---------------------------------------------------------------------------

$clrAccent = [System.Drawing.Color]::FromArgb(0, 120, 212)
$clrDanger = [System.Drawing.Color]::FromArgb(200, 50, 50)

if ($script:Prefs.DarkMode) {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $clrPanelBg  = [System.Drawing.Color]::FromArgb(40, 40, 40)
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(180, 200, 220)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(48, 48, 48)
    $clrGridLine = [System.Drawing.Color]::FromArgb(60, 60, 60)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(45, 45, 45)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(55, 55, 55)
    $clrLogBg    = [System.Drawing.Color]::FromArgb(35, 35, 35)
    $clrLogFg    = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $clrText     = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrGridText = [System.Drawing.Color]::FromArgb(220, 220, 220)
    $clrErrText  = [System.Drawing.Color]::FromArgb(255, 100, 100)
    $clrWarnText = [System.Drawing.Color]::FromArgb(255, 200, 80)
    $clrOkText   = [System.Drawing.Color]::FromArgb(80, 200, 80)
    $clrInfoText = [System.Drawing.Color]::FromArgb(100, 180, 255)
    $clrCardBlue = [System.Drawing.Color]::FromArgb(25, 40, 60)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(20, 50, 30)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(50, 45, 20)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(55, 25, 25)
} else {
    $clrFormBg   = [System.Drawing.Color]::FromArgb(245, 246, 248)
    $clrPanelBg  = [System.Drawing.Color]::White
    $clrHint     = [System.Drawing.Color]::FromArgb(140, 140, 140)
    $clrSubtitle = [System.Drawing.Color]::FromArgb(220, 230, 245)
    $clrGridAlt  = [System.Drawing.Color]::FromArgb(248, 250, 252)
    $clrGridLine = [System.Drawing.Color]::FromArgb(230, 230, 230)
    $clrDetailBg = [System.Drawing.Color]::FromArgb(250, 250, 250)
    $clrSepLine  = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $clrLogBg    = [System.Drawing.Color]::White
    $clrLogFg    = [System.Drawing.Color]::Black
    $clrText     = [System.Drawing.Color]::Black
    $clrGridText = [System.Drawing.Color]::Black
    $clrErrText  = [System.Drawing.Color]::FromArgb(180, 0, 0)
    $clrWarnText = [System.Drawing.Color]::FromArgb(180, 120, 0)
    $clrOkText   = [System.Drawing.Color]::FromArgb(34, 139, 34)
    $clrInfoText = [System.Drawing.Color]::FromArgb(0, 100, 180)
    $clrCardBlue = [System.Drawing.Color]::FromArgb(220, 235, 255)
    $clrCardGreen  = [System.Drawing.Color]::FromArgb(225, 245, 230)
    $clrCardYellow = [System.Drawing.Color]::FromArgb(255, 248, 220)
    $clrCardRed    = [System.Drawing.Color]::FromArgb(255, 230, 230)
}

if ($script:Prefs.DarkMode) {
    if (-not ('DarkToolStripRenderer' -as [type])) {
        $rendererCs = @(
            'using System.Drawing;', 'using System.Windows.Forms;',
            'public class DarkToolStripRenderer : ToolStripProfessionalRenderer {',
            '    private Color _bg;',
            '    public DarkToolStripRenderer(Color bg) : base() { _bg = bg; }',
            '    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e) { }',
            '    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) { if (e.Item.Selected || e.Item.Pressed) { using (var b = new SolidBrush(Color.FromArgb(60, 60, 60))) { e.Graphics.FillRectangle(b, new Rectangle(Point.Empty, e.Item.Size)); } } }',
            '    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e) { int y = e.Item.Height / 2; using (var p = new Pen(Color.FromArgb(70, 70, 70))) { e.Graphics.DrawLine(p, 0, y, e.Item.Width, y); } }',
            '    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { using (var b = new SolidBrush(_bg)) { e.Graphics.FillRectangle(b, e.AffectedBounds); } }',
            '}'
        ) -join "`r`n"
        Add-Type -ReferencedAssemblies System.Windows.Forms, System.Drawing -TypeDefinition $rendererCs
    }
    $script:DarkRenderer = New-Object DarkToolStripRenderer($clrPanelBg)
}

# ---------------------------------------------------------------------------
# Shared grid + card helpers
# ---------------------------------------------------------------------------

function New-ThemedGrid { param([switch]$MultiSelect)
    $g = New-Object System.Windows.Forms.DataGridView; $g.Dock = [System.Windows.Forms.DockStyle]::Fill
    $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false; $g.AllowUserToResizeRows = $false
    $g.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect; $g.MultiSelect = [bool]$MultiSelect
    $g.AutoGenerateColumns = $false; $g.RowHeadersVisible = $false; $g.BackgroundColor = $clrPanelBg
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $g.GridColor = $clrGridLine; $g.ColumnHeadersDefaultCellStyle.BackColor = $clrAccent
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $g.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $g.ColumnHeadersDefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(4)
    $g.ColumnHeadersDefaultCellStyle.WrapMode = [System.Windows.Forms.DataGridViewTriState]::False
    $g.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
    $g.ColumnHeadersHeight = 32; $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.EnableHeadersVisualStyles = $false; $g.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $g.DefaultCellStyle.ForeColor = $clrGridText; $g.DefaultCellStyle.BackColor = $clrPanelBg
    $g.DefaultCellStyle.Padding = New-Object System.Windows.Forms.Padding(2)
    $g.DefaultCellStyle.SelectionBackColor = if ($script:Prefs.DarkMode) { [System.Drawing.Color]::FromArgb(38, 79, 120) } else { [System.Drawing.Color]::FromArgb(0, 120, 215) }
    $g.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
    $g.RowTemplate.Height = 26; $g.AlternatingRowsDefaultCellStyle.BackColor = $clrGridAlt
    Enable-DoubleBuffer -Control $g; return $g
}

function New-SummaryCard { param([string]$Title, [int]$TabIndex)
    $card = New-Object System.Windows.Forms.Panel; $card.Size = New-Object System.Drawing.Size(200, 44)
    $card.Margin = New-Object System.Windows.Forms.Padding(4, 2, 4, 2); $card.BackColor = $clrPanelBg
    $card.BorderStyle = [System.Windows.Forms.BorderStyle]::None; $card.Cursor = [System.Windows.Forms.Cursors]::Hand; $card.Tag = $TabIndex
    $bar = New-Object System.Windows.Forms.Panel; $bar.Dock = [System.Windows.Forms.DockStyle]::Left; $bar.Width = 4; $bar.BackColor = $clrHint; $card.Controls.Add($bar)
    $lt = New-Object System.Windows.Forms.Label; $lt.Text = $Title; $lt.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $lt.ForeColor = $clrText; $lt.AutoSize = $true; $lt.Location = New-Object System.Drawing.Point(10, 4); $lt.BackColor = [System.Drawing.Color]::Transparent; $card.Controls.Add($lt)
    $lv = New-Object System.Windows.Forms.Label; $lv.Text = "--"; $lv.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lv.ForeColor = $clrHint; $lv.AutoSize = $true; $lv.Location = New-Object System.Drawing.Point(10, 22); $lv.BackColor = [System.Drawing.Color]::Transparent; $lv.Tag = "value"; $card.Controls.Add($lv)
    $ch = { $tabMain.SelectedIndex = [int]$this.Parent.Tag }; $cch = { $tabMain.SelectedIndex = [int]$this.Tag }
    $card.Add_Click($cch); $lt.Add_Click($ch); $lv.Add_Click($ch)
    return $card
}

function Update-Card { param([System.Windows.Forms.Panel]$Card, [string]$ValueText, [string]$Severity)
    $bar = $Card.Controls[0]; $vl = $Card.Controls | Where-Object { $_.Tag -eq 'value' }
    switch ($Severity) {
        'ok'       { $bar.BackColor = $clrOkText;   $Card.BackColor = $clrCardGreen;  if ($vl) { $vl.ForeColor = $clrOkText } }
        'warn'     { $bar.BackColor = $clrWarnText;  $Card.BackColor = $clrCardYellow; if ($vl) { $vl.ForeColor = $clrWarnText } }
        'critical' { $bar.BackColor = $clrErrText;   $Card.BackColor = $clrCardRed;    if ($vl) { $vl.ForeColor = $clrErrText } }
        'info'     { $bar.BackColor = $clrInfoText;  $Card.BackColor = $clrCardBlue;   if ($vl) { $vl.ForeColor = $clrInfoText } }
        default    { $bar.BackColor = $clrHint;      $Card.BackColor = $clrPanelBg;    if ($vl) { $vl.ForeColor = $clrHint } }
    }
    if ($vl) { $vl.Text = $ValueText }
}

# ---------------------------------------------------------------------------
# Module-level data stores
# ---------------------------------------------------------------------------

$script:AllWindows     = @()
$script:AllCollections = @()

# ---------------------------------------------------------------------------
# Form shell
# ---------------------------------------------------------------------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Maintenance Window Manager"
$form.Size = New-Object System.Drawing.Size(1200, 750)
$form.MinimumSize = New-Object System.Drawing.Size(900, 550)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
$form.BackColor = $clrFormBg
$form.Icon = [System.Drawing.SystemIcons]::Application
Enable-DoubleBuffer -Control $form

# ---------------------------------------------------------------------------
# StatusStrip
# ---------------------------------------------------------------------------

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $clrPanelBg; $statusStrip.ForeColor = $clrText; $statusStrip.SizingGrip = $false
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $statusStrip.Renderer = $script:DarkRenderer }
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusLabel.Text = "Disconnected"
$statusLabel.Spring = $true; $statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft; $statusLabel.ForeColor = $clrText
$statusStrip.Items.Add($statusLabel) | Out-Null
$statusRowCount = New-Object System.Windows.Forms.ToolStripStatusLabel; $statusRowCount.Text = ""; $statusRowCount.ForeColor = $clrHint
$statusStrip.Items.Add($statusRowCount) | Out-Null
$form.Controls.Add($statusStrip)

# ---------------------------------------------------------------------------
# Log panel (bottom dock)
# ---------------------------------------------------------------------------

$pnlLog = New-Object System.Windows.Forms.Panel; $pnlLog.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlLog.Height = 90; $pnlLog.BackColor = $clrLogBg; $pnlLog.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$form.Controls.Add($pnlLog)

$txtLog = New-Object System.Windows.Forms.TextBox; $txtLog.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtLog.Font = New-Object System.Drawing.Font("Consolas", 9); $txtLog.BackColor = $clrLogBg; $txtLog.ForeColor = $clrLogFg
$txtLog.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlLog.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
# Preferences dialog
# ---------------------------------------------------------------------------

function Show-PreferencesDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Preferences"; $dlg.Size = New-Object System.Drawing.Size(440, 300)
    $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = $dlg.Size
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

    $grpApp = New-Object System.Windows.Forms.GroupBox
    $grpApp.Text = "Appearance"; $grpApp.SetBounds(16, 12, 392, 60)
    $grpApp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpApp.ForeColor = $clrText; $grpApp.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpApp.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpApp.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpApp)

    $chkDark = New-Object System.Windows.Forms.CheckBox
    $chkDark.Text = "Enable dark mode (requires restart)"; $chkDark.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $chkDark.AutoSize = $true; $chkDark.Location = New-Object System.Drawing.Point(14, 24)
    $chkDark.Checked = $script:Prefs.DarkMode; $chkDark.ForeColor = $clrText; $chkDark.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $chkDark.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $chkDark.ForeColor = [System.Drawing.Color]::FromArgb(170, 170, 170) }
    $grpApp.Controls.Add($chkDark)

    $grpConn = New-Object System.Windows.Forms.GroupBox
    $grpConn.Text = "MECM Connection"; $grpConn.SetBounds(16, 82, 392, 110)
    $grpConn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $grpConn.ForeColor = $clrText; $grpConn.BackColor = $clrFormBg
    if ($script:Prefs.DarkMode) { $grpConn.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $grpConn.ForeColor = $clrSepLine }
    $dlg.Controls.Add($grpConn)

    $lblSC = New-Object System.Windows.Forms.Label; $lblSC.Text = "Site Code:"; $lblSC.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSC.Location = New-Object System.Drawing.Point(14, 30); $lblSC.AutoSize = $true; $lblSC.ForeColor = $clrText
    $grpConn.Controls.Add($lblSC)
    $txtSC = New-Object System.Windows.Forms.TextBox; $txtSC.SetBounds(130, 27, 80, 24); $txtSC.Text = $script:Prefs.SiteCode
    $txtSC.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtSC.BackColor = $clrDetailBg; $txtSC.ForeColor = $clrText
    $grpConn.Controls.Add($txtSC)

    $lblSP = New-Object System.Windows.Forms.Label; $lblSP.Text = "SMS Provider:"; $lblSP.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSP.Location = New-Object System.Drawing.Point(14, 64); $lblSP.AutoSize = $true; $lblSP.ForeColor = $clrText
    $grpConn.Controls.Add($lblSP)
    $txtSP = New-Object System.Windows.Forms.TextBox; $txtSP.SetBounds(130, 61, 240, 24); $txtSP.Text = $script:Prefs.SMSProvider
    $txtSP.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtSP.BackColor = $clrDetailBg; $txtSP.ForeColor = $clrText
    $grpConn.Controls.Add($txtSP)

    $btnSave = New-Object System.Windows.Forms.Button; $btnSave.Text = "Save"; $btnSave.SetBounds(220, 210, 90, 32)
    $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnSave -BackColor $clrAccent
    $dlg.Controls.Add($btnSave)
    $btnCancel = New-Object System.Windows.Forms.Button; $btnCancel.Text = "Cancel"; $btnCancel.SetBounds(318, 210, 90, 32)
    $btnCancel.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $clrText; $btnCancel.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancel)

    $btnSave.Add_Click({
        $script:Prefs.DarkMode = $chkDark.Checked; $script:Prefs.SiteCode = $txtSC.Text.Trim(); $script:Prefs.SMSProvider = $txtSP.Text.Trim()
        Save-MwPreferences -Prefs $script:Prefs
        $lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
        $lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
        $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close()
    })
    $btnCancel.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.AcceptButton = $btnSave; $dlg.CancelButton = $btnCancel
    $dlg.ShowDialog($form) | Out-Null; $dlg.Dispose()
}

# ---------------------------------------------------------------------------
# MenuStrip
# ---------------------------------------------------------------------------

$menuStrip = New-Object System.Windows.Forms.MenuStrip; $menuStrip.BackColor = $clrPanelBg; $menuStrip.ForeColor = $clrText
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $menuStrip.Renderer = $script:DarkRenderer }

$mnuFile = New-Object System.Windows.Forms.ToolStripMenuItem("&File"); $mnuFile.ForeColor = $clrText
$mnuPrefs = New-Object System.Windows.Forms.ToolStripMenuItem("&Preferences...")
$mnuPrefs.ShortcutKeys = [System.Windows.Forms.Keys]::Control -bor [System.Windows.Forms.Keys]::Oemcomma
$mnuPrefs.ForeColor = $clrText; $mnuPrefs.Add_Click({ Show-PreferencesDialog })
$mnuFile.DropDownItems.Add($mnuPrefs) | Out-Null
$mnuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$mnuExportCsv = New-Object System.Windows.Forms.ToolStripMenuItem("Export to &CSV..."); $mnuExportCsv.ForeColor = $clrText
$mnuExportCsv.Add_Click({
    if ($dtAllWindows.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", "OK", "Information") | Out-Null; return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "CSV files (*.csv)|*.csv"
    $sfd.FileName = "MaintenanceWindows-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"; $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-MaintenanceWindowsCsv -DataTable $dtAllWindows -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported CSV: $($sfd.FileName)"
    }
})
$mnuFile.DropDownItems.Add($mnuExportCsv) | Out-Null

$mnuExportHtml = New-Object System.Windows.Forms.ToolStripMenuItem("Export to &HTML..."); $mnuExportHtml.ForeColor = $clrText
$mnuExportHtml.Add_Click({
    if ($dtAllWindows.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No data to export.", "Export", "OK", "Information") | Out-Null; return }
    $sfd = New-Object System.Windows.Forms.SaveFileDialog; $sfd.Filter = "HTML files (*.html)|*.html"
    $sfd.FileName = "MaintenanceWindows-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"; $sfd.InitialDirectory = $reportsDir
    if ($sfd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Export-MaintenanceWindowsHtml -DataTable $dtAllWindows -OutputPath $sfd.FileName
        Add-LogLine -TextBox $txtLog -Message "Exported HTML: $($sfd.FileName)"
    }
})
$mnuFile.DropDownItems.Add($mnuExportHtml) | Out-Null
$mnuFile.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$mnuExit = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit"); $mnuExit.ForeColor = $clrText
$mnuExit.Add_Click({ $form.Close() }); $mnuFile.DropDownItems.Add($mnuExit) | Out-Null
$menuStrip.Items.Add($mnuFile) | Out-Null

$mnuView = New-Object System.Windows.Forms.ToolStripMenuItem("&View"); $mnuView.ForeColor = $clrText
$tabNames = @('All Windows', 'By Collection', 'Templates', 'Bulk Operations')
for ($idx = 0; $idx -lt $tabNames.Count; $idx++) {
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem($tabNames[$idx]); $mi.ForeColor = $clrText; $mi.Tag = $idx
    $mi.Add_Click({ $tabMain.SelectedIndex = [int]$this.Tag }.GetNewClosure())
    $mnuView.DropDownItems.Add($mi) | Out-Null
}
$menuStrip.Items.Add($mnuView) | Out-Null

$mnuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("&Help"); $mnuHelp.ForeColor = $clrText
$mnuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("&About"); $mnuAbout.ForeColor = $clrText
$mnuAbout.Add_Click({
    [System.Windows.Forms.MessageBox]::Show("Maintenance Window Manager v1.0.0`r`n`r`nView, create, edit, and bulk-manage MECM maintenance windows across all device collections.`r`n`r`nRequires: ConfigMgr console.", "About", "OK", "Information") | Out-Null
})
$mnuHelp.DropDownItems.Add($mnuAbout) | Out-Null; $menuStrip.Items.Add($mnuHelp) | Out-Null

# ---------------------------------------------------------------------------
# Connection bar
# ---------------------------------------------------------------------------

$pnlConnBar = New-Object System.Windows.Forms.Panel; $pnlConnBar.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlConnBar.Height = 40; $pnlConnBar.BackColor = $clrPanelBg
$pnlConnBar.Padding = New-Object System.Windows.Forms.Padding(12, 6, 12, 6); $form.Controls.Add($pnlConnBar)
$flowConn = New-Object System.Windows.Forms.FlowLayoutPanel; $flowConn.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowConn.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowConn.WrapContents = $false
$flowConn.BackColor = $clrPanelBg; $pnlConnBar.Controls.Add($flowConn)

$lblSite = New-Object System.Windows.Forms.Label; $lblSite.Text = "Site:"; $lblSite.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblSite.AutoSize = $true; $lblSite.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0); $lblSite.ForeColor = $clrText; $lblSite.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSite)
$lblSiteVal = New-Object System.Windows.Forms.Label; $lblSiteVal.Text = if ($script:Prefs.SiteCode) { $script:Prefs.SiteCode } else { "(not set)" }
$lblSiteVal.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblSiteVal.AutoSize = $true
$lblSiteVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 16, 0); $lblSiteVal.ForeColor = $clrHint; $lblSiteVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblSiteVal)
$lblProvider = New-Object System.Windows.Forms.Label; $lblProvider.Text = "Provider:"; $lblProvider.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblProvider.AutoSize = $true; $lblProvider.Margin = New-Object System.Windows.Forms.Padding(0, 5, 2, 0); $lblProvider.ForeColor = $clrText; $lblProvider.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProvider)
$lblProviderVal = New-Object System.Windows.Forms.Label; $lblProviderVal.Text = if ($script:Prefs.SMSProvider) { $script:Prefs.SMSProvider } else { "(not set)" }
$lblProviderVal.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblProviderVal.AutoSize = $true
$lblProviderVal.Margin = New-Object System.Windows.Forms.Padding(0, 5, 24, 0); $lblProviderVal.ForeColor = $clrHint; $lblProviderVal.BackColor = $clrPanelBg
$flowConn.Controls.Add($lblProviderVal)

$btnLoad = New-Object System.Windows.Forms.Button; $btnLoad.Text = "Load Windows"
$btnLoad.Size = New-Object System.Drawing.Size(130, 26); $btnLoad.Font = New-Object System.Drawing.Font("Segoe UI", 9)
Set-ModernButtonStyle -Button $btnLoad -BackColor $clrAccent; $flowConn.Controls.Add($btnLoad)

$btnRefresh = New-Object System.Windows.Forms.Button; $btnRefresh.Text = "Refresh"
$btnRefresh.Size = New-Object System.Drawing.Size(80, 26); $btnRefresh.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnRefresh.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0); $btnRefresh.Enabled = $false
Set-ModernButtonStyle -Button $btnRefresh -BackColor $clrAccent; $flowConn.Controls.Add($btnRefresh)

# ---------------------------------------------------------------------------
# Summary cards panel
# ---------------------------------------------------------------------------

$pnlCards = New-Object System.Windows.Forms.Panel; $pnlCards.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlCards.Height = 54; $pnlCards.BackColor = $clrFormBg; $form.Controls.Add($pnlCards)

$flowCards = New-Object System.Windows.Forms.FlowLayoutPanel; $flowCards.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowCards.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowCards.WrapContents = $false
$flowCards.BackColor = $clrFormBg; $flowCards.Padding = New-Object System.Windows.Forms.Padding(8, 2, 8, 2)
$pnlCards.Controls.Add($flowCards)

$cardTotal       = New-SummaryCard -Title "Total Windows"     -TabIndex 0
$cardNoWindows   = New-SummaryCard -Title "No Windows"        -TabIndex 1
$cardDisabled    = New-SummaryCard -Title "Disabled"           -TabIndex 0
$cardUpcoming    = New-SummaryCard -Title "Upcoming (7 days)" -TabIndex 0
$flowCards.Controls.Add($cardTotal); $flowCards.Controls.Add($cardNoWindows)
$flowCards.Controls.Add($cardDisabled); $flowCards.Controls.Add($cardUpcoming)

# ---------------------------------------------------------------------------
# Tab control
# ---------------------------------------------------------------------------

$tabMain = New-Object System.Windows.Forms.TabControl; $tabMain.Dock = [System.Windows.Forms.DockStyle]::Fill
$tabMain.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$tabMain.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabMain.ItemSize = New-Object System.Drawing.Size(130, 30)
$tabMain.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabMain.Add_DrawItem({
    param($s, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $tab = $s.TabPages[$e.Index]; $sel = ($s.SelectedIndex -eq $e.Index)
    $bg = if ($script:Prefs.DarkMode) { if ($sel) { $clrAccent } else { $clrPanelBg } } else { if ($sel) { $clrAccent } else { [System.Drawing.Color]::FromArgb(240, 240, 240) } }
    $fg = if ($sel) { [System.Drawing.Color]::White } else { $clrText }
    $bb = New-Object System.Drawing.SolidBrush($bg); $e.Graphics.FillRectangle($bb, $e.Bounds)
    $ft = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $sf = New-Object System.Drawing.StringFormat; $sf.Alignment = [System.Drawing.StringAlignment]::Near
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Far; $sf.FormatFlags = [System.Drawing.StringFormatFlags]::NoWrap
    $tr = New-Object System.Drawing.RectangleF(($e.Bounds.X + 8), $e.Bounds.Y, ($e.Bounds.Width - 12), ($e.Bounds.Height - 3))
    $tb = New-Object System.Drawing.SolidBrush($fg); $e.Graphics.DrawString($tab.Text, $ft, $tb, $tr, $sf)
    $bb.Dispose(); $tb.Dispose(); $ft.Dispose(); $sf.Dispose()
})

$tabAllWindows    = New-Object System.Windows.Forms.TabPage("All Windows");     $tabAllWindows.BackColor = $clrFormBg
$tabByCollection  = New-Object System.Windows.Forms.TabPage("By Collection");   $tabByCollection.BackColor = $clrFormBg
$tabTemplates     = New-Object System.Windows.Forms.TabPage("Templates");       $tabTemplates.BackColor = $clrFormBg
$tabBulk          = New-Object System.Windows.Forms.TabPage("Bulk Operations"); $tabBulk.BackColor = $clrFormBg
$tabMain.TabPages.AddRange([System.Windows.Forms.TabPage[]]@($tabAllWindows, $tabByCollection, $tabTemplates, $tabBulk))

# ===========================================================================================
# TAB 1: ALL WINDOWS
# ===========================================================================================

# -- DataTable
$dtAllWindows = New-Object System.Data.DataTable
[void]$dtAllWindows.Columns.Add("Collection", [string])
[void]$dtAllWindows.Columns.Add("CollectionID", [string])
[void]$dtAllWindows.Columns.Add("Window Name", [string])
[void]$dtAllWindows.Columns.Add("Type", [string])
[void]$dtAllWindows.Columns.Add("Schedule", [string])
[void]$dtAllWindows.Columns.Add("Duration", [string])
[void]$dtAllWindows.Columns.Add("Next Occurrence", [string])
[void]$dtAllWindows.Columns.Add("UTC", [string])
[void]$dtAllWindows.Columns.Add("Enabled", [string])

# -- Filter bar
$pnlT1Filter = New-Object System.Windows.Forms.Panel; $pnlT1Filter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlT1Filter.Height = 36; $pnlT1Filter.BackColor = $clrFormBg; $tabAllWindows.Controls.Add($pnlT1Filter)

$flowT1Filter = New-Object System.Windows.Forms.FlowLayoutPanel; $flowT1Filter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowT1Filter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowT1Filter.WrapContents = $false
$flowT1Filter.BackColor = $clrFormBg; $flowT1Filter.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$pnlT1Filter.Controls.Add($flowT1Filter)

$lblT1Search = New-Object System.Windows.Forms.Label; $lblT1Search.Text = "Filter:"; $lblT1Search.AutoSize = $true
$lblT1Search.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblT1Search.ForeColor = $clrText
$lblT1Search.Margin = New-Object System.Windows.Forms.Padding(2, 5, 4, 0); $flowT1Filter.Controls.Add($lblT1Search)

$txtT1Search = New-Object System.Windows.Forms.TextBox; $txtT1Search.Width = 200; $txtT1Search.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtT1Search.BackColor = $clrDetailBg; $txtT1Search.ForeColor = $clrText; $flowT1Filter.Controls.Add($txtT1Search)

$cmbT1Type = New-Object System.Windows.Forms.ComboBox; $cmbT1Type.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbT1Type.Width = 150; $cmbT1Type.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbT1Type.BackColor = $clrDetailBg; $cmbT1Type.ForeColor = $clrText
$cmbT1Type.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$cmbT1Type.Items.AddRange(@("All Types", "General", "Software Updates", "Task Sequences")); $cmbT1Type.SelectedIndex = 0
$flowT1Filter.Controls.Add($cmbT1Type)

$chkT1ShowDisabled = New-Object System.Windows.Forms.CheckBox; $chkT1ShowDisabled.Text = "Show disabled"
$chkT1ShowDisabled.AutoSize = $true; $chkT1ShowDisabled.Checked = $true; $chkT1ShowDisabled.ForeColor = $clrText
$chkT1ShowDisabled.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$chkT1ShowDisabled.Margin = New-Object System.Windows.Forms.Padding(12, 4, 0, 0)
if ($script:Prefs.DarkMode) { $chkT1ShowDisabled.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
$flowT1Filter.Controls.Add($chkT1ShowDisabled)

# -- SplitContainer: grid top, detail bottom
$splitT1 = New-Object System.Windows.Forms.SplitContainer; $splitT1.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitT1.Orientation = [System.Windows.Forms.Orientation]::Horizontal; $splitT1.SplitterDistance = 350
$splitT1.SplitterWidth = 6; $splitT1.BackColor = $clrSepLine; $splitT1.Panel1MinSize = 100; $splitT1.Panel2MinSize = 80
$tabAllWindows.Controls.Add($splitT1); $splitT1.BringToFront()

# -- Grid
$gridT1 = New-ThemedGrid
$colT1Coll     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Coll.HeaderText = "Collection"; $colT1Coll.DataPropertyName = "Collection"; $colT1Coll.Width = 200
$colT1CollId   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1CollId.HeaderText = "Collection ID"; $colT1CollId.DataPropertyName = "CollectionID"; $colT1CollId.Width = 100
$colT1Name     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Name.HeaderText = "Window Name"; $colT1Name.DataPropertyName = "Window Name"; $colT1Name.Width = 180
$colT1Type     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Type.HeaderText = "Type"; $colT1Type.DataPropertyName = "Type"; $colT1Type.Width = 120
$colT1Sched    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Sched.HeaderText = "Schedule"; $colT1Sched.DataPropertyName = "Schedule"; $colT1Sched.Width = 200
$colT1Dur      = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Dur.HeaderText = "Duration"; $colT1Dur.DataPropertyName = "Duration"; $colT1Dur.Width = 70
$colT1Next     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Next.HeaderText = "Next Occurrence"; $colT1Next.DataPropertyName = "Next Occurrence"; $colT1Next.Width = 130
$colT1Utc      = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Utc.HeaderText = "UTC"; $colT1Utc.DataPropertyName = "UTC"; $colT1Utc.Width = 45
$colT1Enabled  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT1Enabled.HeaderText = "Enabled"; $colT1Enabled.DataPropertyName = "Enabled"; $colT1Enabled.Width = 60
$gridT1.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colT1Coll, $colT1CollId, $colT1Name, $colT1Type, $colT1Sched, $colT1Dur, $colT1Next, $colT1Utc, $colT1Enabled))
$gridT1.DataSource = $dtAllWindows
$splitT1.Panel1.Controls.Add($gridT1)

# -- Detail panel
$txtT1Detail = New-Object System.Windows.Forms.TextBox; $txtT1Detail.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtT1Detail.Multiline = $true; $txtT1Detail.ReadOnly = $true; $txtT1Detail.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtT1Detail.Font = New-Object System.Drawing.Font("Consolas", 9.5); $txtT1Detail.BackColor = $clrDetailBg; $txtT1Detail.ForeColor = $clrText
$txtT1Detail.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitT1.Panel2.Controls.Add($txtT1Detail)

# -- Context menu for Tab 1 grid
$ctxT1 = New-Object System.Windows.Forms.ContextMenuStrip
if ($script:Prefs.DarkMode -and $script:DarkRenderer) { $ctxT1.Renderer = $script:DarkRenderer }
$ctxT1Edit = New-Object System.Windows.Forms.ToolStripMenuItem("Edit Window..."); $ctxT1Edit.ForeColor = $clrText
$ctxT1Delete = New-Object System.Windows.Forms.ToolStripMenuItem("Delete Window"); $ctxT1Delete.ForeColor = $clrErrText
$ctxT1Toggle = New-Object System.Windows.Forms.ToolStripMenuItem("Enable/Disable"); $ctxT1Toggle.ForeColor = $clrText
$ctxT1Clone = New-Object System.Windows.Forms.ToolStripMenuItem("Clone to Collections..."); $ctxT1Clone.ForeColor = $clrText
$ctxT1SaveTemplate = New-Object System.Windows.Forms.ToolStripMenuItem("Save as Template..."); $ctxT1SaveTemplate.ForeColor = $clrText
$ctxT1.Items.AddRange(@($ctxT1Edit, $ctxT1Toggle, (New-Object System.Windows.Forms.ToolStripSeparator), $ctxT1Clone, $ctxT1SaveTemplate, (New-Object System.Windows.Forms.ToolStripSeparator), $ctxT1Delete))

$gridT1.Add_CellMouseClick({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0) {
        $gridT1.ClearSelection(); $gridT1.Rows[$e.RowIndex].Selected = $true
        $ctxT1.Show($gridT1, $gridT1.PointToClient([System.Windows.Forms.Cursor]::Position))
    }
})

# -- Tab 1 filter logic
$script:ApplyT1Filter = {
    $searchText = $txtT1Search.Text.Trim().ToLower()
    $typeFilter = $cmbT1Type.SelectedItem.ToString()
    $showDisabled = $chkT1ShowDisabled.Checked

    $parts = [System.Collections.ArrayList]::new()
    if ($searchText) {
        # Filter across collection name and window name
        $escaped = $searchText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
        [void]$parts.Add("(Collection LIKE '%$escaped%' OR [Window Name] LIKE '%$escaped%' OR CollectionID LIKE '%$escaped%')")
    }
    if ($typeFilter -ne 'All Types') {
        [void]$parts.Add("[Type] = '$typeFilter'")
    }
    if (-not $showDisabled) {
        [void]$parts.Add("[Enabled] = 'True'")
    }

    $dtAllWindows.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
    $statusRowCount.Text = "$($dtAllWindows.DefaultView.Count) of $($dtAllWindows.Rows.Count) windows"
}

$txtT1Search.Add_TextChanged($script:ApplyT1Filter)
$cmbT1Type.Add_SelectedIndexChanged($script:ApplyT1Filter)
$chkT1ShowDisabled.Add_CheckedChanged($script:ApplyT1Filter)

# -- Tab 1 selection changed -> detail panel
$gridT1.Add_SelectionChanged({
    if ($gridT1.SelectedRows.Count -eq 0) { $txtT1Detail.Text = ''; return }
    $row = $gridT1.SelectedRows[0]
    $collName = $row.Cells["Collection"].Value
    $collId = $row.Cells["CollectionID"].Value
    $winName = $row.Cells["Window Name"].Value

    # Find the full record in $script:AllWindows
    $rec = $script:AllWindows | Where-Object { $_.CollectionID -eq $collId -and $_.WindowName -eq $winName } | Select-Object -First 1

    $lines = [System.Collections.ArrayList]::new()
    [void]$lines.Add("Window:        $winName")
    [void]$lines.Add("Collection:    $collName ($collId)")
    [void]$lines.Add("Type:          $($row.Cells['Type'].Value)")
    [void]$lines.Add("Schedule:      $($row.Cells['Schedule'].Value)")
    [void]$lines.Add("Duration:      $($row.Cells['Duration'].Value)")
    [void]$lines.Add("UTC:           $($row.Cells['UTC'].Value)")
    [void]$lines.Add("Enabled:       $($row.Cells['Enabled'].Value)")
    [void]$lines.Add("")

    if ($rec) {
        [void]$lines.Add("Description:   $($rec.Description)")
        [void]$lines.Add("Recurrence:    $($rec.Recurrence)")
        [void]$lines.Add("Start Time:    $($rec.StartTime)")
        [void]$lines.Add("Window ID:     $($rec.WindowID)")
        [void]$lines.Add("")

        $nextOccs = if ($rec.IsEnabled) {
            # Recalculate next 5 occurrences from the source window
            $srcWindow = $null
            try { $srcWindow = Get-CMMaintenanceWindow -CollectionId $collId -MaintenanceWindowName $winName -ErrorAction SilentlyContinue | Select-Object -First 1 } catch { }
            if ($srcWindow) { Get-NextOccurrences -Window $srcWindow -Count 5 } else { @() }
        } else { @() }

        if ($nextOccs.Count -gt 0) {
            [void]$lines.Add("Next 5 Occurrences:")
            foreach ($occ in $nextOccs) {
                [void]$lines.Add("  $($occ.ToString('yyyy-MM-dd HH:mm (dddd)'))")
            }
        }
    }

    $txtT1Detail.Text = $lines -join "`r`n"
})

# ===========================================================================================
# TAB 2: BY COLLECTION
# ===========================================================================================

$dtCollections = New-Object System.Data.DataTable
[void]$dtCollections.Columns.Add("Name", [string])
[void]$dtCollections.Columns.Add("CollectionID", [string])
[void]$dtCollections.Columns.Add("Windows", [int])
[void]$dtCollections.Columns.Add("Members", [int])

$dtCollWindows = New-Object System.Data.DataTable
[void]$dtCollWindows.Columns.Add("Window Name", [string])
[void]$dtCollWindows.Columns.Add("Type", [string])
[void]$dtCollWindows.Columns.Add("Schedule", [string])
[void]$dtCollWindows.Columns.Add("Duration", [string])
[void]$dtCollWindows.Columns.Add("Enabled", [string])
[void]$dtCollWindows.Columns.Add("UTC", [string])

# -- Filter bar for collection tab
$pnlT2Filter = New-Object System.Windows.Forms.Panel; $pnlT2Filter.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlT2Filter.Height = 36; $pnlT2Filter.BackColor = $clrFormBg; $tabByCollection.Controls.Add($pnlT2Filter)

$flowT2Filter = New-Object System.Windows.Forms.FlowLayoutPanel; $flowT2Filter.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowT2Filter.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowT2Filter.WrapContents = $false
$flowT2Filter.BackColor = $clrFormBg; $flowT2Filter.Padding = New-Object System.Windows.Forms.Padding(4, 4, 4, 4)
$pnlT2Filter.Controls.Add($flowT2Filter)

$lblT2Search = New-Object System.Windows.Forms.Label; $lblT2Search.Text = "Filter:"; $lblT2Search.AutoSize = $true
$lblT2Search.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblT2Search.ForeColor = $clrText
$lblT2Search.Margin = New-Object System.Windows.Forms.Padding(2, 5, 4, 0); $flowT2Filter.Controls.Add($lblT2Search)

$txtT2Search = New-Object System.Windows.Forms.TextBox; $txtT2Search.Width = 200; $txtT2Search.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$txtT2Search.BackColor = $clrDetailBg; $txtT2Search.ForeColor = $clrText; $flowT2Filter.Controls.Add($txtT2Search)

$btnT2NoWindows = New-Object System.Windows.Forms.Button; $btnT2NoWindows.Text = "Without Windows Only"
$btnT2NoWindows.AutoSize = $true; $btnT2NoWindows.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$btnT2NoWindows.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$btnT2NoWindows.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnT2NoWindows.ForeColor = $clrText; $btnT2NoWindows.BackColor = $clrFormBg
$btnT2NoWindows.Tag = $false
$flowT2Filter.Controls.Add($btnT2NoWindows)

# -- Vertical split: collection list (left) / windows (right)
$splitT2 = New-Object System.Windows.Forms.SplitContainer; $splitT2.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitT2.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitT2.SplitterWidth = 6; $splitT2.BackColor = $clrSepLine; $splitT2.Panel1MinSize = 100; $splitT2.Panel2MinSize = 100
$tabByCollection.Controls.Add($splitT2); $splitT2.BringToFront()
$splitT2.SplitterDistance = [Math]::Max($splitT2.Panel1MinSize, [int]($splitT2.Width * 0.4))

# -- Left: collection grid
$gridT2Colls = New-ThemedGrid
$colT2Name   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2Name.HeaderText = "Collection"; $colT2Name.DataPropertyName = "Name"; $colT2Name.Width = 200
$colT2Id     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2Id.HeaderText = "ID"; $colT2Id.DataPropertyName = "CollectionID"; $colT2Id.Width = 90
$colT2WinCt  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WinCt.HeaderText = "Windows"; $colT2WinCt.DataPropertyName = "Windows"; $colT2WinCt.Width = 65
$colT2Mem    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2Mem.HeaderText = "Members"; $colT2Mem.DataPropertyName = "Members"; $colT2Mem.Width = 65
$gridT2Colls.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colT2Name, $colT2Id, $colT2WinCt, $colT2Mem))
$gridT2Colls.DataSource = $dtCollections
$splitT2.Panel1.Controls.Add($gridT2Colls)

# -- Right: windows grid + buttons
$pnlT2Buttons = New-Object System.Windows.Forms.Panel; $pnlT2Buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlT2Buttons.Height = 40; $pnlT2Buttons.BackColor = $clrFormBg
$pnlT2Buttons.Padding = New-Object System.Windows.Forms.Padding(4, 6, 4, 6)
$splitT2.Panel2.Controls.Add($pnlT2Buttons)

$flowT2Btns = New-Object System.Windows.Forms.FlowLayoutPanel; $flowT2Btns.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowT2Btns.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowT2Btns.WrapContents = $false
$flowT2Btns.BackColor = $clrFormBg; $pnlT2Buttons.Controls.Add($flowT2Btns)

$btnT2New = New-Object System.Windows.Forms.Button; $btnT2New.Text = "New Window"; $btnT2New.Size = New-Object System.Drawing.Size(110, 28)
$btnT2New.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnT2New -BackColor $clrAccent
$flowT2Btns.Controls.Add($btnT2New)

$btnT2Edit = New-Object System.Windows.Forms.Button; $btnT2Edit.Text = "Edit"; $btnT2Edit.Size = New-Object System.Drawing.Size(70, 28)
$btnT2Edit.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnT2Edit.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
Set-ModernButtonStyle -Button $btnT2Edit -BackColor $clrAccent; $flowT2Btns.Controls.Add($btnT2Edit)

$btnT2Delete = New-Object System.Windows.Forms.Button; $btnT2Delete.Text = "Delete"; $btnT2Delete.Size = New-Object System.Drawing.Size(70, 28)
$btnT2Delete.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnT2Delete.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
Set-ModernButtonStyle -Button $btnT2Delete -BackColor $clrDanger; $flowT2Btns.Controls.Add($btnT2Delete)

$gridT2Windows = New-ThemedGrid
$colT2WName  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WName.HeaderText = "Window Name"; $colT2WName.DataPropertyName = "Window Name"; $colT2WName.Width = 180
$colT2WType  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WType.HeaderText = "Type"; $colT2WType.DataPropertyName = "Type"; $colT2WType.Width = 120
$colT2WSched = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WSched.HeaderText = "Schedule"; $colT2WSched.DataPropertyName = "Schedule"; $colT2WSched.Width = 200
$colT2WDur   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WDur.HeaderText = "Duration"; $colT2WDur.DataPropertyName = "Duration"; $colT2WDur.Width = 70
$colT2WEn    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WEn.HeaderText = "Enabled"; $colT2WEn.DataPropertyName = "Enabled"; $colT2WEn.Width = 60
$colT2WUtc   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT2WUtc.HeaderText = "UTC"; $colT2WUtc.DataPropertyName = "UTC"; $colT2WUtc.Width = 45
$gridT2Windows.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colT2WName, $colT2WType, $colT2WSched, $colT2WDur, $colT2WEn, $colT2WUtc))
$gridT2Windows.DataSource = $dtCollWindows
$splitT2.Panel2.Controls.Add($gridT2Windows); $gridT2Windows.BringToFront()

# -- Highlight zero-window rows in red
$gridT2Colls.Add_CellFormatting({
    param($s, $e)
    if ($e.ColumnIndex -eq 2 -and $null -ne $e.Value -and [int]$e.Value -eq 0) {
        $e.CellStyle.ForeColor = $clrErrText
        $e.CellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    }
})

# -- Collection filter
$script:ApplyT2Filter = {
    $searchText = $txtT2Search.Text.Trim().ToLower()
    $noWindowsOnly = [bool]$btnT2NoWindows.Tag

    $parts = [System.Collections.ArrayList]::new()
    if ($searchText) {
        $escaped = $searchText.Replace("'", "''").Replace("[", "[[]").Replace("%", "[%]").Replace("*", "[*]")
        [void]$parts.Add("(Name LIKE '%$escaped%' OR CollectionID LIKE '%$escaped%')")
    }
    if ($noWindowsOnly) {
        [void]$parts.Add("Windows = 0")
    }
    $dtCollections.DefaultView.RowFilter = if ($parts.Count -gt 0) { $parts -join ' AND ' } else { '' }
}

$txtT2Search.Add_TextChanged($script:ApplyT2Filter)
$btnT2NoWindows.Add_Click({
    $btnT2NoWindows.Tag = -not [bool]$btnT2NoWindows.Tag
    if ([bool]$btnT2NoWindows.Tag) {
        Set-ModernButtonStyle -Button $btnT2NoWindows -BackColor $clrAccent
    } else {
        $btnT2NoWindows.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnT2NoWindows.ForeColor = $clrText; $btnT2NoWindows.BackColor = $clrFormBg
    }
    & $script:ApplyT2Filter
})

# -- Collection selection -> load windows
$gridT2Colls.Add_SelectionChanged({
    $dtCollWindows.Clear()
    if ($gridT2Colls.SelectedRows.Count -eq 0) { return }
    $selCollId = $gridT2Colls.SelectedRows[0].Cells["CollectionID"].Value
    $collWindows = $script:AllWindows | Where-Object { $_.CollectionID -eq $selCollId }
    foreach ($w in $collWindows) {
        [void]$dtCollWindows.Rows.Add($w.WindowName, $w.Type, $w.Schedule, $w.Duration, $w.IsEnabled.ToString(), $w.IsUTC.ToString())
    }
})

# ===========================================================================================
# TAB 3: TEMPLATES
# ===========================================================================================

$dtTemplates = New-Object System.Data.DataTable
[void]$dtTemplates.Columns.Add("Name", [string])
[void]$dtTemplates.Columns.Add("Type", [string])
[void]$dtTemplates.Columns.Add("Recurrence", [string])
[void]$dtTemplates.Columns.Add("Duration", [string])
[void]$dtTemplates.Columns.Add("Start Time", [string])

$splitT3 = New-Object System.Windows.Forms.SplitContainer; $splitT3.Dock = [System.Windows.Forms.DockStyle]::Fill
$splitT3.Orientation = [System.Windows.Forms.Orientation]::Vertical
$splitT3.SplitterWidth = 6; $splitT3.BackColor = $clrSepLine; $splitT3.Panel1MinSize = 100; $splitT3.Panel2MinSize = 100
$tabTemplates.Controls.Add($splitT3)
$splitT3.SplitterDistance = [Math]::Max($splitT3.Panel1MinSize, [int]($splitT3.Width * 0.4))

# -- Left: template grid + buttons
$pnlT3Buttons = New-Object System.Windows.Forms.Panel; $pnlT3Buttons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlT3Buttons.Height = 40; $pnlT3Buttons.BackColor = $clrFormBg
$pnlT3Buttons.Padding = New-Object System.Windows.Forms.Padding(4, 6, 4, 6)
$splitT3.Panel1.Controls.Add($pnlT3Buttons)

$flowT3Btns = New-Object System.Windows.Forms.FlowLayoutPanel; $flowT3Btns.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowT3Btns.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowT3Btns.WrapContents = $false
$flowT3Btns.BackColor = $clrFormBg; $pnlT3Buttons.Controls.Add($flowT3Btns)

$btnT3New = New-Object System.Windows.Forms.Button; $btnT3New.Text = "New Template"; $btnT3New.Size = New-Object System.Drawing.Size(120, 28)
$btnT3New.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnT3New -BackColor $clrAccent
$flowT3Btns.Controls.Add($btnT3New)

$btnT3Delete = New-Object System.Windows.Forms.Button; $btnT3Delete.Text = "Delete"; $btnT3Delete.Size = New-Object System.Drawing.Size(70, 28)
$btnT3Delete.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnT3Delete.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
Set-ModernButtonStyle -Button $btnT3Delete -BackColor $clrDanger; $flowT3Btns.Controls.Add($btnT3Delete)

$btnT3Apply = New-Object System.Windows.Forms.Button; $btnT3Apply.Text = "Apply to Collections..."; $btnT3Apply.Size = New-Object System.Drawing.Size(160, 28)
$btnT3Apply.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnT3Apply.Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
Set-ModernButtonStyle -Button $btnT3Apply -BackColor $clrAccent; $flowT3Btns.Controls.Add($btnT3Apply)

$gridT3 = New-ThemedGrid
$colT3Name = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT3Name.HeaderText = "Name"; $colT3Name.DataPropertyName = "Name"; $colT3Name.Width = 180
$colT3Type = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT3Type.HeaderText = "Type"; $colT3Type.DataPropertyName = "Type"; $colT3Type.Width = 120
$colT3Rec  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT3Rec.HeaderText = "Recurrence"; $colT3Rec.DataPropertyName = "Recurrence"; $colT3Rec.Width = 140
$colT3Dur  = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT3Dur.HeaderText = "Duration"; $colT3Dur.DataPropertyName = "Duration"; $colT3Dur.Width = 80
$colT3Time = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT3Time.HeaderText = "Start Time"; $colT3Time.DataPropertyName = "Start Time"; $colT3Time.Width = 80
$gridT3.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colT3Name, $colT3Type, $colT3Rec, $colT3Dur, $colT3Time))
$gridT3.DataSource = $dtTemplates
$splitT3.Panel1.Controls.Add($gridT3); $gridT3.BringToFront()

# -- Right: template preview
$txtT3Preview = New-Object System.Windows.Forms.TextBox; $txtT3Preview.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtT3Preview.Multiline = $true; $txtT3Preview.ReadOnly = $true; $txtT3Preview.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtT3Preview.Font = New-Object System.Drawing.Font("Consolas", 9.5); $txtT3Preview.BackColor = $clrDetailBg; $txtT3Preview.ForeColor = $clrText
$txtT3Preview.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$splitT3.Panel2.Controls.Add($txtT3Preview)

# -- Template selection -> preview
$script:LoadedTemplates = @()

function Refresh-TemplateGrid {
    $script:LoadedTemplates = Get-WindowTemplates -TemplatesPath $templatesDir
    $dtTemplates.Clear()
    foreach ($t in $script:LoadedTemplates) {
        $durStr = "{0}h {1}m" -f [int]$t.DurationHours, [int]$t.DurationMinutes
        $timeStr = "{0}:{1:D2}" -f [int]$t.StartHour, [int]$t.StartMinute
        [void]$dtTemplates.Rows.Add($t.Name, $t.WindowType, $t.RecurrenceType, $durStr, $timeStr)
    }
}

$gridT3.Add_SelectionChanged({
    if ($gridT3.SelectedRows.Count -eq 0) { $txtT3Preview.Text = ''; return }
    $idx = $gridT3.SelectedRows[0].Index
    if ($idx -lt 0 -or $idx -ge $script:LoadedTemplates.Count) { return }
    $t = $script:LoadedTemplates[$idx]

    $lines = [System.Collections.ArrayList]::new()
    [void]$lines.Add("Template:      $($t.Name)")
    [void]$lines.Add("Description:   $($t.Description)")
    [void]$lines.Add("Window Type:   $($t.WindowType)")
    [void]$lines.Add("Recurrence:    $($t.RecurrenceType)")
    [void]$lines.Add("Start Time:    $($t.StartHour):$($t.StartMinute.ToString('D2'))")
    [void]$lines.Add("Duration:      $($t.DurationHours)h $($t.DurationMinutes)m")
    [void]$lines.Add("UTC:           $($t.IsUtc)")
    [void]$lines.Add("")

    if ($t.RecurrenceType -eq 'Weekly' -or $t.RecurrenceType -eq 'MonthlyByWeekday') {
        [void]$lines.Add("Day of Week:   $($t.DayOfWeek)")
    }
    if ($t.RecurrenceType -eq 'MonthlyByWeekday') {
        [void]$lines.Add("Week Order:    $($t.WeekOrder)")
    }
    if ($t.RecurrenceType -eq 'MonthlyByDate') {
        [void]$lines.Add("Day of Month:  $($t.DayOfMonth)")
    }
    if ([int]$t.PatchTuesdayOffset -ge 0) {
        [void]$lines.Add("Patch Tuesday: +$($t.PatchTuesdayOffset) days")
    }
    [void]$lines.Add("")
    [void]$lines.Add("File: $($t.FileName)")

    $txtT3Preview.Text = $lines -join "`r`n"
})

$btnT3Delete.Add_Click({
    if ($gridT3.SelectedRows.Count -eq 0) { return }
    $idx = $gridT3.SelectedRows[0].Index
    if ($idx -lt 0 -or $idx -ge $script:LoadedTemplates.Count) { return }
    $t = $script:LoadedTemplates[$idx]
    $confirm = [System.Windows.Forms.MessageBox]::Show("Delete template '$($t.Name)'?", "Confirm Delete", "YesNo", "Question")
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-WindowTemplate -FilePath $t.FilePath
        Add-LogLine -TextBox $txtLog -Message "Deleted template: $($t.Name)"
        Refresh-TemplateGrid
    }
})

# ===========================================================================================
# TAB 4: BULK OPERATIONS
# ===========================================================================================

$pnlT4Top = New-Object System.Windows.Forms.Panel; $pnlT4Top.Dock = [System.Windows.Forms.DockStyle]::Top
$pnlT4Top.Height = 44; $pnlT4Top.BackColor = $clrFormBg
$pnlT4Top.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 4); $tabBulk.Controls.Add($pnlT4Top)

$flowT4Top = New-Object System.Windows.Forms.FlowLayoutPanel; $flowT4Top.Dock = [System.Windows.Forms.DockStyle]::Fill
$flowT4Top.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight; $flowT4Top.WrapContents = $false
$flowT4Top.BackColor = $clrFormBg; $pnlT4Top.Controls.Add($flowT4Top)

$lblT4Op = New-Object System.Windows.Forms.Label; $lblT4Op.Text = "Operation:"; $lblT4Op.AutoSize = $true
$lblT4Op.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $lblT4Op.ForeColor = $clrText
$lblT4Op.Margin = New-Object System.Windows.Forms.Padding(0, 5, 6, 0); $flowT4Top.Controls.Add($lblT4Op)

$cmbT4Op = New-Object System.Windows.Forms.ComboBox; $cmbT4Op.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$cmbT4Op.Width = 200; $cmbT4Op.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$cmbT4Op.BackColor = $clrDetailBg; $cmbT4Op.ForeColor = $clrText
$cmbT4Op.Items.AddRange(@("Import from CSV", "Copy Window", "Bulk Enable/Disable", "Bulk Delete")); $cmbT4Op.SelectedIndex = 0
$flowT4Top.Controls.Add($cmbT4Op)

$btnT4Execute = New-Object System.Windows.Forms.Button; $btnT4Execute.Text = "Execute"; $btnT4Execute.Size = New-Object System.Drawing.Size(100, 28)
$btnT4Execute.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnT4Execute.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
Set-ModernButtonStyle -Button $btnT4Execute -BackColor $clrAccent; $flowT4Top.Controls.Add($btnT4Execute)

# -- Preview grid
$dtT4Preview = New-Object System.Data.DataTable
[void]$dtT4Preview.Columns.Add("Collection", [string])
[void]$dtT4Preview.Columns.Add("CollectionID", [string])
[void]$dtT4Preview.Columns.Add("Action", [string])
[void]$dtT4Preview.Columns.Add("Detail", [string])
[void]$dtT4Preview.Columns.Add("Result", [string])

$gridT4 = New-ThemedGrid -MultiSelect
$colT4Coll   = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT4Coll.HeaderText = "Collection"; $colT4Coll.DataPropertyName = "Collection"; $colT4Coll.Width = 200
$colT4Id     = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT4Id.HeaderText = "Collection ID"; $colT4Id.DataPropertyName = "CollectionID"; $colT4Id.Width = 100
$colT4Act    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT4Act.HeaderText = "Action"; $colT4Act.DataPropertyName = "Action"; $colT4Act.Width = 130
$colT4Det    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT4Det.HeaderText = "Detail"; $colT4Det.DataPropertyName = "Detail"; $colT4Det.Width = 250
$colT4Res    = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colT4Res.HeaderText = "Result"; $colT4Res.DataPropertyName = "Result"; $colT4Res.Width = 100
$gridT4.Columns.AddRange([System.Windows.Forms.DataGridViewColumn[]]@($colT4Coll, $colT4Id, $colT4Act, $colT4Det, $colT4Res))
$gridT4.DataSource = $dtT4Preview

# -- Results log at bottom
$pnlT4Log = New-Object System.Windows.Forms.Panel; $pnlT4Log.Dock = [System.Windows.Forms.DockStyle]::Bottom
$pnlT4Log.Height = 100; $pnlT4Log.BackColor = $clrLogBg
$pnlT4Log.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4); $tabBulk.Controls.Add($pnlT4Log)

$txtT4Log = New-Object System.Windows.Forms.TextBox; $txtT4Log.Dock = [System.Windows.Forms.DockStyle]::Fill
$txtT4Log.Multiline = $true; $txtT4Log.ReadOnly = $true; $txtT4Log.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$txtT4Log.Font = New-Object System.Drawing.Font("Consolas", 9); $txtT4Log.BackColor = $clrLogBg; $txtT4Log.ForeColor = $clrLogFg
$txtT4Log.BorderStyle = [System.Windows.Forms.BorderStyle]::None
$pnlT4Log.Controls.Add($txtT4Log)

$tabBulk.Controls.Add($gridT4); $gridT4.BringToFront()

# -- Color-code Result column
$gridT4.Add_CellFormatting({
    param($s, $e)
    if ($e.ColumnIndex -eq 4 -and $null -ne $e.Value) {
        $val = $e.Value.ToString()
        if ($val -eq 'Success') { $e.CellStyle.ForeColor = $clrOkText }
        elseif ($val -eq 'Failed') { $e.CellStyle.ForeColor = $clrErrText }
        elseif ($val -eq 'Pending') { $e.CellStyle.ForeColor = $clrWarnText }
    }
})

# -- Import CSV handler
$btnT4Execute.Add_Click({
    $op = $cmbT4Op.SelectedItem.ToString()

    switch ($op) {
        "Import from CSV" {
            $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = "CSV files (*.csv)|*.csv"
            if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $preview = Import-MaintenanceWindowsCsv -CsvPath $ofd.FileName
            $dtT4Preview.Clear()
            foreach ($row in $preview) {
                $collDisplay = if ($row.CollectionName) { $row.CollectionName } else { $row.CollectionID }
                $detail = "$($row.WindowName) - $($row.RecurrenceType) $($row.DurationHours)h"
                [void]$dtT4Preview.Rows.Add($collDisplay, $row.CollectionID, "Create Window", $detail, "Pending")
            }
            Add-LogLine -TextBox $txtT4Log -Message "Loaded $(@($preview).Count) rows from CSV. Click Execute again to apply."
        }
        "Bulk Enable/Disable" {
            if ($script:AllWindows.Count -eq 0) {
                Add-LogLine -TextBox $txtT4Log -Message "No windows loaded. Load windows first."
                return
            }
            $confirm = [System.Windows.Forms.MessageBox]::Show("Toggle enabled state for all currently loaded windows?", "Bulk Toggle", "YesNo", "Question")
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

            $dtT4Preview.Clear()
            foreach ($w in $script:AllWindows) {
                $newState = -not $w.IsEnabled
                $action = if ($newState) { "Enable" } else { "Disable" }
                [void]$dtT4Preview.Rows.Add($w.CollectionName, $w.CollectionID, $action, $w.WindowName, "Pending")
            }
            Add-LogLine -TextBox $txtT4Log -Message "Prepared $($script:AllWindows.Count) toggle operations. Review preview."
        }
        default {
            Add-LogLine -TextBox $txtT4Log -Message "Operation '$op' - use the appropriate controls to set up the operation first."
        }
    }
})

# ===========================================================================================
# SCHEDULE BUILDER DIALOG
# ===========================================================================================

function Show-ScheduleBuilderDialog {
    <#
    .SYNOPSIS
        Modal dialog for creating/editing a maintenance window schedule.
    .PARAMETER Mode
        'New' or 'Edit'
    .PARAMETER CollectionId
        Collection to create/edit the window on.
    .PARAMETER ExistingWindowName
        For edit mode, the name of the existing window.
    .RETURNS
        $true if a window was created/modified, $false otherwise.
    #>
    param(
        [string]$Mode = 'New',
        [string]$CollectionId = '',
        [string]$ExistingWindowName = ''
    )

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = if ($Mode -eq 'Edit') { "Edit Maintenance Window" } else { "New Maintenance Window" }
    $dlg.Size = New-Object System.Drawing.Size(520, 520)
    $dlg.MinimumSize = $dlg.Size; $dlg.MaximumSize = New-Object System.Drawing.Size(520, 600)
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false; $dlg.ShowInTaskbar = $false
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9.5); $dlg.BackColor = $clrFormBg

    $y = 14

    # Name
    $lblName = New-Object System.Windows.Forms.Label; $lblName.Text = "Window Name:"; $lblName.SetBounds(16, $y, 120, 20)
    $lblName.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblName.ForeColor = $clrText; $dlg.Controls.Add($lblName)
    $txtName = New-Object System.Windows.Forms.TextBox; $txtName.SetBounds(140, $y, 350, 24)
    $txtName.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtName.BackColor = $clrDetailBg; $txtName.ForeColor = $clrText; $dlg.Controls.Add($txtName)
    $y += 32

    # Description
    $lblDesc = New-Object System.Windows.Forms.Label; $lblDesc.Text = "Description:"; $lblDesc.SetBounds(16, $y, 120, 20)
    $lblDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblDesc.ForeColor = $clrText; $dlg.Controls.Add($lblDesc)
    $txtDesc = New-Object System.Windows.Forms.TextBox; $txtDesc.SetBounds(140, $y, 350, 24)
    $txtDesc.Font = New-Object System.Drawing.Font("Segoe UI", 9); $txtDesc.BackColor = $clrDetailBg; $txtDesc.ForeColor = $clrText; $dlg.Controls.Add($txtDesc)
    $y += 32

    # Recurrence type
    $lblRec = New-Object System.Windows.Forms.Label; $lblRec.Text = "Recurrence:"; $lblRec.SetBounds(16, $y, 120, 20)
    $lblRec.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblRec.ForeColor = $clrText; $dlg.Controls.Add($lblRec)
    $cmbRec = New-Object System.Windows.Forms.ComboBox; $cmbRec.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbRec.SetBounds(140, $y, 200, 24); $cmbRec.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbRec.BackColor = $clrDetailBg; $cmbRec.ForeColor = $clrText
    $cmbRec.Items.AddRange(@("OneTime", "Daily", "Weekly", "MonthlyByDate", "MonthlyByWeekday")); $cmbRec.SelectedIndex = 2
    $dlg.Controls.Add($cmbRec)
    $y += 32

    # Day of week
    $lblDow = New-Object System.Windows.Forms.Label; $lblDow.Text = "Day of Week:"; $lblDow.SetBounds(16, $y, 120, 20)
    $lblDow.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblDow.ForeColor = $clrText; $dlg.Controls.Add($lblDow)
    $cmbDow = New-Object System.Windows.Forms.ComboBox; $cmbDow.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbDow.SetBounds(140, $y, 140, 24); $cmbDow.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbDow.BackColor = $clrDetailBg; $cmbDow.ForeColor = $clrText
    $cmbDow.Items.AddRange(@("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")); $cmbDow.SelectedIndex = 0
    $dlg.Controls.Add($cmbDow)
    $y += 32

    # Week order (for MonthlyByWeekday)
    $lblWo = New-Object System.Windows.Forms.Label; $lblWo.Text = "Week Order:"; $lblWo.SetBounds(16, $y, 120, 20)
    $lblWo.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblWo.ForeColor = $clrText; $dlg.Controls.Add($lblWo)
    $cmbWo = New-Object System.Windows.Forms.ComboBox; $cmbWo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbWo.SetBounds(140, $y, 140, 24); $cmbWo.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbWo.BackColor = $clrDetailBg; $cmbWo.ForeColor = $clrText
    $cmbWo.Items.AddRange(@("First", "Second", "Third", "Fourth", "Last")); $cmbWo.SelectedIndex = 0
    $dlg.Controls.Add($cmbWo)
    $y += 32

    # Day of month (for MonthlyByDate)
    $lblDom = New-Object System.Windows.Forms.Label; $lblDom.Text = "Day of Month:"; $lblDom.SetBounds(16, $y, 120, 20)
    $lblDom.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblDom.ForeColor = $clrText; $dlg.Controls.Add($lblDom)
    $nudDom = New-Object System.Windows.Forms.NumericUpDown; $nudDom.SetBounds(140, $y, 70, 24)
    $nudDom.Font = New-Object System.Drawing.Font("Segoe UI", 9); $nudDom.Minimum = 1; $nudDom.Maximum = 31; $nudDom.Value = 1
    $nudDom.BackColor = $clrDetailBg; $nudDom.ForeColor = $clrText; $dlg.Controls.Add($nudDom)
    $y += 32

    # Start date/time
    $lblStart = New-Object System.Windows.Forms.Label; $lblStart.Text = "Start Date/Time:"; $lblStart.SetBounds(16, $y, 120, 20)
    $lblStart.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblStart.ForeColor = $clrText; $dlg.Controls.Add($lblStart)
    $dtpStart = New-Object System.Windows.Forms.DateTimePicker; $dtpStart.SetBounds(140, $y, 200, 24)
    $dtpStart.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $dtpStart.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
    $dtpStart.CustomFormat = "yyyy-MM-dd HH:mm"; $dtpStart.ShowUpDown = $true
    $dtpStart.Value = (Get-Date -Hour 2 -Minute 0 -Second 0)
    $dlg.Controls.Add($dtpStart)
    $y += 32

    # Duration
    $lblDur = New-Object System.Windows.Forms.Label; $lblDur.Text = "Duration:"; $lblDur.SetBounds(16, $y, 120, 20)
    $lblDur.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblDur.ForeColor = $clrText; $dlg.Controls.Add($lblDur)
    $nudDurH = New-Object System.Windows.Forms.NumericUpDown; $nudDurH.SetBounds(140, $y, 60, 24)
    $nudDurH.Font = New-Object System.Drawing.Font("Segoe UI", 9); $nudDurH.Minimum = 0; $nudDurH.Maximum = 23; $nudDurH.Value = 4
    $nudDurH.BackColor = $clrDetailBg; $nudDurH.ForeColor = $clrText; $dlg.Controls.Add($nudDurH)
    $lblH = New-Object System.Windows.Forms.Label; $lblH.Text = "h"; $lblH.SetBounds(204, $y, 16, 20)
    $lblH.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblH.ForeColor = $clrText; $dlg.Controls.Add($lblH)
    $nudDurM = New-Object System.Windows.Forms.NumericUpDown; $nudDurM.SetBounds(224, $y, 60, 24)
    $nudDurM.Font = New-Object System.Drawing.Font("Segoe UI", 9); $nudDurM.Minimum = 0; $nudDurM.Maximum = 59; $nudDurM.Value = 0
    $nudDurM.BackColor = $clrDetailBg; $nudDurM.ForeColor = $clrText; $dlg.Controls.Add($nudDurM)
    $lblM = New-Object System.Windows.Forms.Label; $lblM.Text = "m"; $lblM.SetBounds(288, $y, 16, 20)
    $lblM.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblM.ForeColor = $clrText; $dlg.Controls.Add($lblM)
    $y += 32

    # Window type
    $lblWType = New-Object System.Windows.Forms.Label; $lblWType.Text = "Window Type:"; $lblWType.SetBounds(16, $y, 120, 20)
    $lblWType.Font = New-Object System.Drawing.Font("Segoe UI", 9); $lblWType.ForeColor = $clrText; $dlg.Controls.Add($lblWType)
    $cmbWType = New-Object System.Windows.Forms.ComboBox; $cmbWType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $cmbWType.SetBounds(140, $y, 200, 24); $cmbWType.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $cmbWType.BackColor = $clrDetailBg; $cmbWType.ForeColor = $clrText
    $cmbWType.Items.AddRange(@("Any", "SoftwareUpdatesOnly", "TaskSequencesOnly")); $cmbWType.SelectedIndex = 0
    $dlg.Controls.Add($cmbWType)
    $y += 32

    # UTC + Enabled checkboxes
    $chkUtc = New-Object System.Windows.Forms.CheckBox; $chkUtc.Text = "UTC"; $chkUtc.SetBounds(140, $y, 70, 24)
    $chkUtc.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkUtc.ForeColor = $clrText
    if ($script:Prefs.DarkMode) { $chkUtc.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
    $dlg.Controls.Add($chkUtc)

    $chkEnabled = New-Object System.Windows.Forms.CheckBox; $chkEnabled.Text = "Enabled"; $chkEnabled.SetBounds(220, $y, 100, 24)
    $chkEnabled.Checked = $true; $chkEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 9); $chkEnabled.ForeColor = $clrText
    if ($script:Prefs.DarkMode) { $chkEnabled.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat }
    $dlg.Controls.Add($chkEnabled)
    $y += 36

    # Show/hide controls based on recurrence type
    $updateVisibility = {
        $rec = $cmbRec.SelectedItem.ToString()
        $lblDow.Visible = $rec -eq 'Weekly' -or $rec -eq 'MonthlyByWeekday'
        $cmbDow.Visible = $rec -eq 'Weekly' -or $rec -eq 'MonthlyByWeekday'
        $lblWo.Visible  = $rec -eq 'MonthlyByWeekday'
        $cmbWo.Visible  = $rec -eq 'MonthlyByWeekday'
        $lblDom.Visible = $rec -eq 'MonthlyByDate'
        $nudDom.Visible = $rec -eq 'MonthlyByDate'
    }
    $cmbRec.Add_SelectedIndexChanged($updateVisibility)
    & $updateVisibility

    # OK / Cancel
    $btnOk = New-Object System.Windows.Forms.Button; $btnOk.Text = "OK"; $btnOk.SetBounds(310, ($y), 90, 32)
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9); Set-ModernButtonStyle -Button $btnOk -BackColor $clrAccent
    $dlg.Controls.Add($btnOk)
    $btnCancelDlg = New-Object System.Windows.Forms.Button; $btnCancelDlg.Text = "Cancel"; $btnCancelDlg.SetBounds(408, ($y), 90, 32)
    $btnCancelDlg.Font = New-Object System.Drawing.Font("Segoe UI", 9); $btnCancelDlg.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancelDlg.ForeColor = $clrText; $btnCancelDlg.BackColor = $clrFormBg
    $dlg.Controls.Add($btnCancelDlg)

    $script:ScheduleDialogResult = $false

    $btnOk.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtName.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Window name is required.", "Validation", "OK", "Warning") | Out-Null
            return
        }

        $totalMin = ([int]$nudDurH.Value * 60) + [int]$nudDurM.Value
        if ($totalMin -le 0 -or $totalMin -ge 1440) {
            [System.Windows.Forms.MessageBox]::Show("Duration must be between 1 minute and 23 hours 59 minutes.", "Validation", "OK", "Warning") | Out-Null
            return
        }

        $schedParams = @{
            RecurrenceType  = $cmbRec.SelectedItem.ToString()
            StartTime       = $dtpStart.Value
            DurationHours   = [int]$nudDurH.Value
            DurationMinutes = [int]$nudDurM.Value
        }
        if ($cmbDow.Visible) { $schedParams['DayOfWeek'] = [System.DayOfWeek]$cmbDow.SelectedItem.ToString() }
        if ($cmbWo.Visible)  { $schedParams['WeekOrder'] = $cmbWo.SelectedItem.ToString() }
        if ($nudDom.Visible) { $schedParams['DayOfMonth'] = [int]$nudDom.Value }

        $schedule = New-WindowSchedule @schedParams
        if (-not $schedule) {
            [System.Windows.Forms.MessageBox]::Show("Failed to create schedule. Check log for details.", "Error", "OK", "Error") | Out-Null
            return
        }

        if ($Mode -eq 'Edit' -and $ExistingWindowName) {
            # Delete old, create new (CM doesn't support renaming)
            Remove-ManagedMaintenanceWindow -CollectionId $CollectionId -MaintenanceWindowName $ExistingWindowName | Out-Null
        }

        $result = New-ManagedMaintenanceWindow -CollectionId $CollectionId -Name $txtName.Text.Trim() `
            -Schedule $schedule -ApplyTo $cmbWType.SelectedItem.ToString() `
            -IsEnabled $chkEnabled.Checked -IsUtc $chkUtc.Checked

        if ($result) {
            $script:ScheduleDialogResult = $true
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK; $dlg.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to create maintenance window. Check log.", "Error", "OK", "Error") | Out-Null
        }
    })

    $btnCancelDlg.Add_Click({ $dlg.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $dlg.Close() })
    $dlg.AcceptButton = $btnOk; $dlg.CancelButton = $btnCancelDlg
    $dlg.ShowDialog($form) | Out-Null; $dlg.Dispose()

    return $script:ScheduleDialogResult
}

# ===========================================================================================
# Wire up Tab 2 CRUD buttons
# ===========================================================================================

$btnT2New.Add_Click({
    if ($gridT2Colls.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select a collection first.", "New Window", "OK", "Information") | Out-Null; return
    }
    $selCollId = $gridT2Colls.SelectedRows[0].Cells["CollectionID"].Value
    $changed = Show-ScheduleBuilderDialog -Mode 'New' -CollectionId $selCollId
    if ($changed) {
        Add-LogLine -TextBox $txtLog -Message "Created maintenance window on $selCollId"
        & $script:RefreshAllData
    }
})

$btnT2Edit.Add_Click({
    if ($gridT2Colls.SelectedRows.Count -eq 0 -or $gridT2Windows.SelectedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Select a collection and window to edit.", "Edit Window", "OK", "Information") | Out-Null; return
    }
    $selCollId = $gridT2Colls.SelectedRows[0].Cells["CollectionID"].Value
    $selWinName = $gridT2Windows.SelectedRows[0].Cells["Window Name"].Value
    $changed = Show-ScheduleBuilderDialog -Mode 'Edit' -CollectionId $selCollId -ExistingWindowName $selWinName
    if ($changed) {
        Add-LogLine -TextBox $txtLog -Message "Updated maintenance window '$selWinName' on $selCollId"
        & $script:RefreshAllData
    }
})

$btnT2Delete.Add_Click({
    if ($gridT2Colls.SelectedRows.Count -eq 0 -or $gridT2Windows.SelectedRows.Count -eq 0) { return }
    $selCollId = $gridT2Colls.SelectedRows[0].Cells["CollectionID"].Value
    $selWinName = $gridT2Windows.SelectedRows[0].Cells["Window Name"].Value
    $confirm = [System.Windows.Forms.MessageBox]::Show("Delete window '$selWinName' from $selCollId ?", "Confirm Delete", "YesNo", "Question")
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        $ok = Remove-ManagedMaintenanceWindow -CollectionId $selCollId -MaintenanceWindowName $selWinName
        if ($ok) { Add-LogLine -TextBox $txtLog -Message "Deleted '$selWinName' from $selCollId" }
        & $script:RefreshAllData
    }
})

# Wire Tab 1 context menu
$ctxT1Edit.Add_Click({
    if ($gridT1.SelectedRows.Count -eq 0) { return }
    $collId = $gridT1.SelectedRows[0].Cells["CollectionID"].Value
    $winName = $gridT1.SelectedRows[0].Cells["Window Name"].Value
    $changed = Show-ScheduleBuilderDialog -Mode 'Edit' -CollectionId $collId -ExistingWindowName $winName
    if ($changed) { & $script:RefreshAllData }
})

$ctxT1Delete.Add_Click({
    if ($gridT1.SelectedRows.Count -eq 0) { return }
    $collId = $gridT1.SelectedRows[0].Cells["CollectionID"].Value
    $winName = $gridT1.SelectedRows[0].Cells["Window Name"].Value
    $confirm = [System.Windows.Forms.MessageBox]::Show("Delete window '$winName'?", "Confirm Delete", "YesNo", "Question")
    if ($confirm -eq [System.Windows.Forms.DialogResult]::Yes) {
        Remove-ManagedMaintenanceWindow -CollectionId $collId -MaintenanceWindowName $winName | Out-Null
        & $script:RefreshAllData
    }
})

$ctxT1Toggle.Add_Click({
    if ($gridT1.SelectedRows.Count -eq 0) { return }
    $collId = $gridT1.SelectedRows[0].Cells["CollectionID"].Value
    $winName = $gridT1.SelectedRows[0].Cells["Window Name"].Value
    $currentEnabled = $gridT1.SelectedRows[0].Cells["Enabled"].Value -eq 'True'
    Set-ManagedMaintenanceWindow -CollectionId $collId -MaintenanceWindowName $winName -IsEnabled (-not $currentEnabled) | Out-Null
    & $script:RefreshAllData
})

# ===========================================================================================
# Data loading
# ===========================================================================================

$script:RefreshAllData = {
    $statusLabel.Text = "Loading maintenance windows..."
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $script:AllWindows = @(Get-AllMaintenanceWindows)
        $script:AllCollections = @(Get-DeviceCollectionSummary)

        # Build window count per collection
        $winCountMap = @{}
        foreach ($w in $script:AllWindows) {
            if (-not $winCountMap.ContainsKey($w.CollectionID)) { $winCountMap[$w.CollectionID] = 0 }
            $winCountMap[$w.CollectionID]++
        }

        # Populate Tab 1 DataTable
        $dtAllWindows.Clear()
        foreach ($w in $script:AllWindows) {
            [void]$dtAllWindows.Rows.Add(
                $w.CollectionName, $w.CollectionID, $w.WindowName, $w.Type,
                $w.Schedule, $w.Duration, $w.NextOccurrence, $w.IsUTC.ToString(), $w.IsEnabled.ToString()
            )
        }

        # Populate Tab 2 collection DataTable
        $dtCollections.Clear()
        foreach ($c in $script:AllCollections) {
            $wCount = if ($winCountMap.ContainsKey($c.CollectionID)) { $winCountMap[$c.CollectionID] } else { 0 }
            [void]$dtCollections.Rows.Add($c.Name, $c.CollectionID, $wCount, $c.MemberCount)
        }

        # Update summary cards
        $totalWindows = $script:AllWindows.Count
        $disabledCount = @($script:AllWindows | Where-Object { -not $_.IsEnabled }).Count
        $noWindowCount = @($script:AllCollections | Where-Object { -not $winCountMap.ContainsKey($_.CollectionID) }).Count

        $now = Get-Date
        $weekFromNow = $now.AddDays(7)
        $upcomingCount = @($script:AllWindows | Where-Object {
            $_.NextOccurrence -ne 'N/A' -and $_.IsEnabled -and
            [datetime]::TryParse($_.NextOccurrence, [ref]$null) -and
            [datetime]$_.NextOccurrence -le $weekFromNow
        }).Count

        Update-Card -Card $cardTotal -ValueText $totalWindows.ToString() -Severity $(if ($totalWindows -gt 0) { 'info' } else { 'warn' })
        Update-Card -Card $cardNoWindows -ValueText $noWindowCount.ToString() -Severity $(if ($noWindowCount -eq 0) { 'ok' } elseif ($noWindowCount -gt 50) { 'critical' } else { 'warn' })
        Update-Card -Card $cardDisabled -ValueText $disabledCount.ToString() -Severity $(if ($disabledCount -eq 0) { 'ok' } else { 'warn' })
        Update-Card -Card $cardUpcoming -ValueText $upcomingCount.ToString() -Severity 'info'

        # Apply filters
        & $script:ApplyT1Filter
        & $script:ApplyT2Filter

        $statusLabel.Text = "Connected - $totalWindows windows across $($script:AllCollections.Count) collections"
        Add-LogLine -TextBox $txtLog -Message "Loaded $totalWindows maintenance windows from $($script:AllCollections.Count) collections"
    }
    catch {
        $statusLabel.Text = "Error loading data"
        Add-LogLine -TextBox $txtLog -Message "ERROR: $_"
        Write-Log "Data load error: $_" -Level ERROR
    }
}

$btnLoad.Add_Click({
    if (-not $script:Prefs.SiteCode -or -not $script:Prefs.SMSProvider) {
        [System.Windows.Forms.MessageBox]::Show("Set Site Code and SMS Provider in Preferences first.", "Connection", "OK", "Warning") | Out-Null
        Show-PreferencesDialog
        return
    }

    $statusLabel.Text = "Connecting to $($script:Prefs.SiteCode)..."
    [System.Windows.Forms.Application]::DoEvents()

    $connected = Connect-CMSite -SiteCode $script:Prefs.SiteCode -SMSProvider $script:Prefs.SMSProvider
    if (-not $connected) {
        $statusLabel.Text = "Connection failed"
        Add-LogLine -TextBox $txtLog -Message "Failed to connect to site $($script:Prefs.SiteCode)"
        return
    }

    Add-LogLine -TextBox $txtLog -Message "Connected to site $($script:Prefs.SiteCode)"
    $btnRefresh.Enabled = $true

    & $script:RefreshAllData

    # Load templates
    Refresh-TemplateGrid
})

$btnRefresh.Add_Click({
    if (-not (Test-CMConnection)) {
        Add-LogLine -TextBox $txtLog -Message "Not connected. Click Load Windows first."
        return
    }
    & $script:RefreshAllData
    Refresh-TemplateGrid
})

# ===========================================================================================
# Dock ordering + finalize
# ===========================================================================================

$form.Controls.Add($tabMain)
$form.Controls.Add($menuStrip)

# Dock Z-order: back to front
$menuStrip.SendToBack()
$pnlConnBar.BringToFront()
$pnlCards.BringToFront()
$tabMain.BringToFront()

# ---------------------------------------------------------------------------
# Form events
# ---------------------------------------------------------------------------

$form.Add_Shown({ Restore-WindowState })
$form.Add_FormClosing({
    Save-WindowState
    if (Test-CMConnection) { Disconnect-CMSite }
})

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

Add-LogLine -TextBox $txtLog -Message "Maintenance Window Manager v1.0.0 started"

[System.Windows.Forms.Application]::Run($form)
