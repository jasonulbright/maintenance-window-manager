$win = Start-App 'start-maintenancewindowmgr.ps1' 'Maintenance Window Manager'
Size-Win $win 1440 900
$win = Find-Top 'Maintenance Window Manager' 10
function Idle([string]$id = 'btnRefresh', [int]$sec = 300) { Start-Sleep -Seconds 2; if (-not (Wait-Until { $b = Get-El $win $id -Quiet; $b -and $b.Current.IsEnabled } $sec)) { throw ("Timeout: $id did not become enabled in $sec seconds.") }; Start-Sleep -Seconds 2 }
function Rows($grid) { @((Get-El $win $grid).FindAll('Descendants', (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::DataItem)))) }
function Tab([string]$name) { Select-El (Find-Role $win 'TabItem' $name); Start-Sleep -Milliseconds 900 }
function WinNames { $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window); (@($root.FindAll('Descendants', $c)) | ForEach-Object { $_.Current.Name }) -join ' ; ' }
# Open a dialog with a button, capture the dialog, dump its tree and close it with its Cancel button.
# The step never answers Create, Apply or Save. When the dialog has no Cancel button, the step fails.
function Dlg([string]$btn, [string]$shot) {
  Invoke-El (Get-El $win $btn); Start-Sleep -Seconds 3
  Log ('after ' + $btn + ' windows: ' + (WinNames))
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Window)
  $d = @($root.FindAll('Descendants', $c)) | Where-Object { $_.Current.Name -notin 'Maintenance Window Manager', 'Program Manager' -and $_.Current.Name -ne '' -and $_.Current.Name -notlike 'Administrator*' -and $_.Current.Name -notlike 'GlowWindow*' } | Select-Object -First 1
  if (-not $d) { throw ('No dialog after ' + $btn) }
  Shot $d $shot; Dump-Tree $d $shot
  $cancel = Get-El $d 'btnCancel' -Quiet
  if (-not $cancel) { throw ('No Cancel button in the dialog of ' + $btn + '. Nothing was answered.') }
  Invoke-El $cancel; Start-Sleep -Seconds 1
}
Section 'first run' {
  Shot $win 'start-first-run'
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 2; Shot $win 'refresh-without-connection'
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'options-empty'; Dump-Tree $d 'options'
  Set-Text (Get-El $d 'txtSiteCode') $Capture.SiteCode; Set-Text (Get-El $d 'txtSmsProvider') $Capture.SMSProvider; Shot $d 'options-connection'
  Invoke-El (Get-El $d 'btnCatAbout'); Start-Sleep -Milliseconds 600; Shot $d 'options-about'
  Invoke-El (Get-El $d 'btnCatConnection'); Invoke-El (Get-El $d 'btnOk'); Start-Sleep -Seconds 1
}
Section 'refresh' {
  Invoke-El (Get-El $win 'btnRefresh'); Start-Sleep -Seconds 3; Shot $win 'refresh-running'; Idle; Shot $win 'windows'
}
Section 'window detail' {
  $r = Rows 'gridWindows'; Log ('window rows: ' + $r.Count)
  if ($r.Count) { Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'window-properties'; Tab 'Next 5 Occurrences'; Shot $win 'window-occurrences'; Tab 'Properties' }
  Set-Text (Get-El $win 'txtFilter') 'a'; Start-Sleep -Seconds 1; Shot $win 'windows-filter-text'; Set-Text (Get-El $win 'txtFilter') ''
  Expand-El (Get-El $win 'cboStatusFilter'); Start-Sleep -Milliseconds 500; Shot $win 'windows-filter-list' -Hover; Collapse-El (Get-El $win 'cboStatusFilter')
}
Section 'exports' {
  Invoke-El (Get-El $win 'btnExportCsv'); File-Dialog $null $null; Shot $win 'windows-exported'
  Invoke-El (Get-El $win 'btnExportHtml'); File-Dialog $null $null; Start-Sleep -Seconds 1
}
Section 'window dialogs' {
  Dlg 'btnNewWindow' 'dialog-new-window'
  $r = Rows 'gridWindows'; if ($r.Count) { Select-El $r[0]; Start-Sleep -Milliseconds 700; Dlg 'btnEditWindow' 'dialog-edit-window' }
}
Section 'coverage' {
  Invoke-El (Get-El $win 'btnViewCoverage'); Start-Sleep -Seconds 3; Shot $win 'coverage-before'
  Invoke-El (Get-El $win 'btnRefresh'); Idle; Shot $win 'coverage'
  $r = Rows 'gridCoverage'; Log ('coverage rows: ' + $r.Count)
  if ($r.Count) { Select-El $r[0]; Start-Sleep -Milliseconds 700; Shot $win 'coverage-selected'; Dlg 'btnApplyTemplateBulk' 'dialog-apply-template-bulk' }
}
Section 'templates' {
  Invoke-El (Get-El $win 'btnViewTemplates'); Start-Sleep -Seconds 3; Shot $win 'templates'
  $r = Rows 'gridTemplates'; Log ('template rows: ' + $r.Count)
  if ($r.Count) { Select-El $r[0]; Start-Sleep -Seconds 1; Shot $win 'template-selected'; Dlg 'btnEditTemplate' 'dialog-edit-template'; Dlg 'btnApplyTemplateOne' 'dialog-apply-template-one' }
  Dlg 'btnNewTemplate' 'dialog-new-template'
}
Section 'tab order' { Invoke-El (Get-El $win 'btnViewWindows'); Focus-Win $win; Click-El (Get-El $win 'btnViewWindows'); 1..14 | ForEach-Object { Send-Keys '{TAB}'; $f = [System.Windows.Automation.AutomationElement]::FocusedElement; Log ('tab -> ' + $f.Current.AutomationId + ' / ' + $f.Current.Name) } }
Section 'light' {
  Toggle-El (Get-El $win 'toggleTheme'); Start-Sleep -Seconds 1
  foreach ($v in @(@('Windows', 'light-windows'), @('Coverage', 'light-coverage'), @('Templates', 'light-templates'))) { Invoke-El (Get-El $win ('btnView' + $v[0])); Start-Sleep -Seconds 2; Shot $win $v[1] }
  Invoke-El (Get-El $win 'btnOptions'); $d = Find-Last 'Options' 10; Shot $d 'light-options'; Focus-Win $d; Send-Keys '{ESC}'; Start-Sleep -Seconds 1
  Toggle-El (Get-El $win 'toggleTheme')
}
Section 'persist' { Invoke-El (Get-El $win 'btnViewTemplates'); Start-Sleep -Milliseconds 600; Invoke-El (Get-El $win 'Close'); Start-Sleep -Seconds 8; $win = Start-App 'start-maintenancewindowmgr.ps1' 'Maintenance Window Manager'; Shot $win 'start-reopen' }
Stop-App 'start-maintenancewindowmgr.ps1'
