# Lab-side UI Automation helpers. Dot-sourced by the runner inside the interactive session.
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing, System.Windows.Forms
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
public class NativeUi {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint f, int dx, int dy, uint d, UIntPtr e);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int ht, bool r);
}
'@
[void][NativeUi]::SetProcessDPIAware()
$AE = [System.Windows.Automation.AutomationElement]
$root = $AE::RootElement
function Log([string]$m) { ('{0}  {1}' -f (Get-Date -Format 'HH:mm:ss'), $m) | Add-Content -LiteralPath (Join-Path $OutDir 'run.log') }
function Find-Top([string]$Name, [int]$Seconds = 30) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    $end = (Get-Date).AddSeconds($Seconds)
    do { $el = $root.FindFirst('Descendants', $cond); if ($el) { return $el }; Start-Sleep -Milliseconds 400 } while ((Get-Date) -lt $end)
    $names = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition) | ForEach-Object { $_.Current.Name + '|' + $_.Current.ClassName }) -join ' ; '
    Log ("top windows: " + $names); Shot-Screen ('fail-' + ($Name -replace '\W', '_')); throw "window not found: $Name"
}
function Get-El($Scope, [string]$Id, [switch]$Quiet) {
    $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::AutomationIdProperty, $Id)))
    if (-not $el) { $el = $Scope.FindFirst('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Id))) }
    if (-not $el -and -not $Quiet) { throw "control not found: $Id" }
    $el
}
function Invoke-El($El) { [void][NativeUi]::SetCursorPos(1900, 1060); $El.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke() }
function Set-Text($El, [string]$Text) { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).SetValue($Text) }
function Toggle-El($El) { $El.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern).Toggle() }
function Select-El($El) { $El.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select() }
function Expand-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Expand() }
function Collapse-El($El) { $El.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse() }
function Get-ElText($El) { try { $El.GetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern).Current.Value } catch { $El.Current.Name } }
function Wait-Until([scriptblock]$Test, [int]$Seconds = 120) { $end = (Get-Date).AddSeconds($Seconds); while ((Get-Date) -lt $end) { if (& $Test) { return $true }; Start-Sleep -Milliseconds 500 }; $false }
function Click-At([int]$x, [int]$y) { [void][NativeUi]::SetCursorPos($x, $y); Start-Sleep -Milliseconds 150; [NativeUi]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 250 }
function Click-El($El) { $r = $El.Current.BoundingRectangle; Click-At ([int]($r.X + $r.Width / 2)) ([int]($r.Y + $r.Height / 2)) }
function Move-To($El) { $r = $El.Current.BoundingRectangle; [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)) }
function Send-Keys([string]$k) { [System.Windows.Forms.SendKeys]::SendWait($k); Start-Sleep -Milliseconds 400 }
function Focus-Win($Win) { try { [void][NativeUi]::SetForegroundWindow([IntPtr]$Win.Current.NativeWindowHandle) } catch { $null = $_ }; Start-Sleep -Milliseconds 300 }
function Size-Win($Win, [int]$W = 1440, [int]$H = 900) {
    $hh = [IntPtr]$Win.Current.NativeWindowHandle
    $sw = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $w2 = [Math]::Min($W, $sw.Width); $h2 = [Math]::Min($H, $sw.Height - 40)
    [void][NativeUi]::MoveWindow($hh, 0, 0, $w2, $h2, $true); Start-Sleep -Milliseconds 900
}
function Maximize-Win($Win) {
    # The design size is the work area of a maximized window on a 1080p screen (D-20): the window fills it, nothing is hardcoded.
    $Win.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern).SetWindowVisualState([System.Windows.Automation.WindowVisualState]::Maximized)
    Start-Sleep -Milliseconds 1500
    $wa = [System.Windows.Forms.SystemInformation]::WorkingArea
    Log ('work area {0}x{1}' -f $wa.Width, $wa.Height)
}
# Evidence helpers: sample the screen colors of a control. A sample is part of the verdict of the capture run (D-21).
function Get-Px([int]$x, [int]$y) {
    $b = New-Object System.Drawing.Bitmap 1, 1; $g = [System.Drawing.Graphics]::FromImage($b)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size 1, 1)); $g.Dispose()
    $c = $b.GetPixel(0, 0); $b.Dispose(); $c
}
function Get-Lum($c) {
    $f = { param($v) $s = $v / 255.0; if ($s -le 0.03928) { $s / 12.92 } else { [Math]::Pow(($s + 0.055) / 1.055, 2.4) } }
    0.2126 * (& $f $c.R) + 0.7152 * (& $f $c.G) + 0.0722 * (& $f $c.B)
}
function Get-Contrast($a, $b) { $l1 = Get-Lum $a; $l2 = Get-Lum $b; ([Math]::Max($l1, $l2) + 0.05) / ([Math]::Min($l1, $l2) + 0.05) }
function Sample-El($El, [string]$Label) {
    # The background is the most frequent color along a strip through the middle of the control. The text color is the
    # color that differs most from it. The strip skips the border and the first pixels.
    $r = $El.Current.BoundingRectangle
    $y = [int]($r.Y + $r.Height / 2); $count = @{}; $px = New-Object System.Collections.Generic.List[object]
    $span = [Math]::Min(160, [int]($r.Width) - 24)
    foreach ($dx in 0..([Math]::Max(1, $span / 2))) {
        foreach ($dy in -3..3) { $c = Get-Px ([int]($r.X + 12 + $dx * 2)) ($y + $dy); $px.Add($c); $k = $c.ToArgb(); $count[$k] = 1 + [int]$count[$k] }
    }
    $bgKey = ($count.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    $bg = [System.Drawing.Color]::FromArgb($bgKey); $best = $bg; $bestD = -1
    foreach ($c in $px) { $d = [Math]::Abs($c.R - $bg.R) + [Math]::Abs($c.G - $bg.G) + [Math]::Abs($c.B - $bg.B); if ($d -gt $bestD) { $bestD = $d; $best = $c } }
    Log ('SAMPLE {0}: bg=#{1:X2}{2:X2}{3:X2} text=#{4:X2}{5:X2}{6:X2} contrast={7:N1}' -f $Label, $bg.R, $bg.G, $bg.B, $best.R, $best.G, $best.B, (Get-Contrast $bg $best))
}
function Extra([string]$Name, [scriptblock]$Body) {
    # A step that only gathers evidence. A failure logs WARN, so that it never stops the capture.
    try { & $Body } catch { Log ('WARN extra ' + $Name + ': ' + $_.Exception.Message) }
}
# A real mouse move, not only SetCursorPos: WPF raises MouseLeave and closes a tooltip only after a real move. Then wait until no tooltip shows.
# A tooltip that stays (the keyboard focus of a clicked control can hold one) goes away when the mouse touches the title bar of the window and leaves again.
# The mouse parks on an empty part of the taskbar: at the clock in the corner it would show the date tooltip of the taskbar in the image.
function Move-MouseAway($Win) {
    [void][NativeUi]::SetCursorPos(1500, 1065); Start-Sleep -Milliseconds 100
    [NativeUi]::mouse_event(1, 3, 3, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 100
    [NativeUi]::mouse_event(1, -3, -3, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 300
    $tipCond = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::ToolTip)
    $end = (Get-Date).AddSeconds(4)
    while ((Get-Date) -lt $end -and $root.FindFirst('Descendants', $tipCond)) { Start-Sleep -Milliseconds 400 }
    if ($Win -and $root.FindFirst('Descendants', $tipCond)) {
        $r = $Win.Current.BoundingRectangle
        # A window that is closing has an empty rectangle: skip the touch then.
        if (-not [double]::IsInfinity($r.X) -and -not [double]::IsNaN($r.X) -and $r.Width -gt 0) {
            [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + 14)); Start-Sleep -Milliseconds 300
            [NativeUi]::mouse_event(1, 2, 2, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 400
        }
        [void][NativeUi]::SetCursorPos(1500, 1065); Start-Sleep -Milliseconds 100
        [NativeUi]::mouse_event(1, 3, 3, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 100
        [NativeUi]::mouse_event(1, -3, -3, 0, [UIntPtr]::Zero); Start-Sleep -Seconds 2
    }
    if ($Win -and ($Win.Current.Name -eq 'Options' -or $root.FindFirst('Descendants', $tipCond))) {
        # The keyboard focus of a text box holds its tooltip. The Options window is always treated so, because UI Automation does not always list that tooltip. The focus moves to the Cancel or OK button of the window, which has none.
        foreach ($n in 'Cancel', 'OK') {
            $b = $Win.FindFirst('Descendants', (New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $n)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Button)))))
            if ($b) { try { $b.SetFocus() } catch { $null = $_ }; Start-Sleep -Seconds 2; break }
        }
        [void][NativeUi]::SetCursorPos(1500, 1065); Start-Sleep -Milliseconds 300
    }
    if ($root.FindFirst('Descendants', $tipCond)) { Log 'TOOLTIP still shown after the mouse moved away; the next shot holds it' }
}
function Shot($Win, [string]$Name, [switch]$Hover) {
    if (-not $Hover) { Move-MouseAway $Win; if ($Win.Current.ClassName -eq 'MetroWindow') { try { (Get-El $Win 'txtLog' -Quiet).SetFocus() } catch { $null = $_ } } }
    Start-Sleep -Milliseconds 1300
    $r = $Win.Current.BoundingRectangle
    $x = [int][Math]::Max(0, $r.X); $y = [int][Math]::Max(0, $r.Y)
    $w = [int][Math]::Min($r.Width, [System.Windows.Forms.SystemInformation]::VirtualScreen.Width - $x)
    $h = [int][Math]::Min($r.Height, [System.Windows.Forms.SystemInformation]::VirtualScreen.Height - $y)
    $bmp = New-Object System.Drawing.Bitmap $w, $h; $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size $w, $h)); $g.Dispose()
    $bmp.Save((Join-Path $OutDir ($Name + '.png')), [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    Log ('shot {0} {1}x{2}' -f $Name, $w, $h)
}
function Shot-Screen([string]$Name) {
    $s = [System.Windows.Forms.SystemInformation]::VirtualScreen
    $b = New-Object System.Drawing.Bitmap $s.Width, $s.Height; $g = [System.Drawing.Graphics]::FromImage($b); $g.CopyFromScreen(0, 0, 0, 0, $b.Size); $g.Dispose()
    $b.Save((Join-Path $OutDir ($Name + '.png'))); $b.Dispose(); Log ('screen ' + $Name)
}
function Dump-Tree($Scope, [string]$Name) {
    $out = foreach ($e in $Scope.FindAll('Descendants', [System.Windows.Automation.Condition]::TrueCondition)) {
        $c = $e.Current; if ($c.Name -or $c.AutomationId) { '{0}|{1}|{2}|{3}|{4}' -f $c.ControlType.ProgrammaticName.Replace('ControlType.', ''), $c.AutomationId, $c.Name, $c.IsEnabled, [math]::Round([double]$c.BoundingRectangle.Width) }
    }
    $out | Set-Content (Join-Path $OutDir ($Name + '.tree.txt'))
}
# The capture stops only the processes that it started, by process id (vault: testing-policy, "Process rules on a lab VM").
# Each run records its processes in a ledger in its own run folder. No process is ever stopped by a command-line pattern.
$script:RunLedger = Join-Path $PSScriptRoot 'run.pids'
$script:AppProcs = @{}
function Register-RunProcess($Process, [string]$Key) {
    $script:AppProcs[$Key] = $Process
    ('{0}|{1}' -f $Process.Id, $Process.StartTime.Ticks) | Add-Content -LiteralPath $script:RunLedger
}
function Start-App([string]$Script, [string]$Title) {
    [void][NativeUi]::SetCursorPos(1900, 1060)
    $script:proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', ('"' + (Join-Path $appDir $Script) + '"')) -PassThru
    Register-RunProcess $script:proc $Script
    [void](Find-Top $Title 60); Start-Sleep -Seconds 3; Find-Top $Title 10
}
function Stop-App([string]$Script) {
    # Stop only the process that Start-App started for this script, by its process id.
    $p = $script:AppProcs[$Script]
    if ($p -and -not $p.HasExited) { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue }
}
function Section([string]$Name, [scriptblock]$Body) { Log ('== ' + $Name); try { & $Body } catch { Log ('ERROR in ' + $Name + ': ' + $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber) } }
function Find-All([string]$Name) {
    $cond = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)), (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)))
    @($root.FindAll('Descendants', $cond))
}
function Find-Last([string]$Name, [int]$Seconds = 15) {
    $end = (Get-Date).AddSeconds($Seconds); do { $a = Find-All $Name; if ($a.Count) { return $a[$a.Count - 1] }; Start-Sleep -Milliseconds 400 } while ((Get-Date) -lt $end); throw "no window $Name"
}
function Find-Role($Scope, [string]$Type, [string]$Name) {
    $ct = [System.Windows.Automation.ControlType]::($Type)
    $c = New-Object System.Windows.Automation.AndCondition((New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, $ct)), (New-Object System.Windows.Automation.PropertyCondition($AE::NameProperty, $Name)))
    $Scope.FindFirst('Descendants', $c)
}
function RightClick-El($El) { $r = $El.Current.BoundingRectangle; [void][NativeUi]::SetCursorPos([int]($r.X + $r.Width / 2), [int]($r.Y + $r.Height / 2)); Start-Sleep -Milliseconds 200; [NativeUi]::mouse_event(8, 0, 0, 0, [UIntPtr]::Zero); [NativeUi]::mouse_event(16, 0, 0, 0, [UIntPtr]::Zero); Start-Sleep -Milliseconds 600 }
function File-Dialog([string]$Path, [string]$Shot) {
    # Wait for the Save As window. Stop the capture when it does not show,
    # so that an export that was never saved cannot count as a success.
    $d = $null; $end = (Get-Date).AddSeconds(20)
    while (-not $d -and (Get-Date) -lt $end) {
        $d = @(Find-All 'Save As') | Select-Object -First 1
        if (-not $d) { $d = @($root.FindAll('Children', [System.Windows.Automation.Condition]::TrueCondition)) | Where-Object { $_.Current.ClassName -eq '#32770' } | Select-Object -First 1 }
        if (-not $d) { Start-Sleep -Milliseconds 400 }
    }
    if (-not $d) {
        $names = @($root.FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window))) | ForEach-Object { $_.Current.Name + '|' + $_.Current.ClassName }) -join ' ; '
        Log ('windows: ' + $names)
        throw 'The Save As dialog did not show.'
    }
    Start-Sleep -Milliseconds 800
    if ($Shot) { Shot $d $Shot }
    Focus-Win $d
    if ($Path) { Send-Keys ($Path -replace '([+^%~{}\[\]()])', '{$1}') }
    Send-Keys '{ENTER}'; Start-Sleep -Seconds 2
}
