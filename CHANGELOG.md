# Changelog

All notable changes to Maintenance Window Manager are documented in this
file.

## [1.2.2] - 2026-09-04

### Changed

- **Vendored `SuiteCommon` 0.4.3.** The module repairs the process
  PSModulePath at import: a Windows PowerShell process launched from
  PowerShell 7 inherits the 7.x module directories, and the background
  runspace opened later autoloaded a Microsoft.PowerShell.Utility without
  Get-FileHash or ConvertFrom-Json, so background operations failed with
  an unrelated "term not recognized". A background runspace whose module
  import fails is disposed and the original error thrown instead of being
  returned as an opened but unusable worker.

## [1.2.1] - 2026-08-16

### Changed

- **Vendored `SuiteCommon` 0.3.2.** Window restore applies the saved
  geometry before maximizing, so un-maximizing returns to the saved size
  instead of the XAML defaults; background runspace bootstrap failures
  are named in the log instead of surfacing later as an unrelated
  "term not recognized".

## [1.2.0] - 2026-08-16

### Changed

- **Window chrome, theming, dialogs, background-work, and the collection
  widgets now come from the vendored `SuiteCommon` module** (0.3.0).
  The confirm dialog and collection picker/tree are the shared
  implementations (confirm call sites pass their owner explicitly).
  Behavior gains: hook state no longer leaks on window close, a
  maximized close persists the pre-maximize geometry, an off-screen
  saved position clamps into the nearest monitor, background teardown no
  longer blocks the UI thread, and the collection tree's filtered-count
  status line no longer receives a stray boolean ahead of the count.

## [1.1.0] - 2026-08-14

### Changed

- **Shared plumbing moved to the vendored `SuiteCommon` module.** Logging
  (`Initialize-Logging`, `Write-Log`) and CM site connection
  (`Connect-CMSite`, `Disconnect-CMSite`, `Test-CMConnection`) now load
  from `Lib\SuiteCommon\`, shared across the tool suite and synced from
  the suite-core repository instead of hand-edited per repo. Behavior is
  unchanged — same log format, same connection flow — and the connection
  gains provider rebind when the configured SMS Provider changes plus
  rebuild of a stale drive whose provider connection died.

## [1.0.0] - 2026-05-02

Maintenance Window Manager is a MahApps.Metro WPF desktop tool for
auditing, creating, editing, and bulk-applying MECM device collection
maintenance windows. Extract the zip and run `start-maintenancewindowmgr.ps1`.

### Features

- **Windows view** -- master-detail grid of every maintenance window
  across every device collection: glyph status (enabled vs muted),
  Collection, Window, Type, Recurrence, human-readable Schedule,
  Duration, Next Occurrence. Filter by collection or window name.
  Status filter (All / Enabled only / Disabled only / SoftwareUpdates
  only / TaskSequences only). Detail panel tabs: Properties, projected
  Next 5 occurrences. Action bar: New Window..., Edit..., Toggle
  Enabled, Remove...
- **Coverage view** -- per-collection gap analysis grid: Name,
  Collection ID, Members, Window count, Built-in flag. Glyph: check
  for collections with at least one window, warn for those with zero.
  Multi-select rows + "Apply Template..." for bulk schedule rollout.
- **Templates view** -- ships 5 baseline templates (Daily After Hours,
  Monthly First Saturday, Patch Tuesday +3 days, Patch Tuesday +7
  days, Weekly Sunday 2 AM). New Template / Edit / Delete / Apply.
- **Schedule Editor** -- inline reusable component embedded in New
  Window, Edit Window, and Template Editor dialogs. Recurrence:
  One-time / Daily / Weekly / Monthly-by-Date / Monthly-by-Weekday /
  Patch Tuesday +N days. Auto-toggling subpanels per recurrence. Live
  next-5-occurrences preview that updates on every keystroke. UTC
  flag.
- **Tree-picker collection selection** -- folder hierarchy from
  `SMS_ObjectContainerNode` so the picker scales to thousands of
  collections nested under console folders. Used in New Window,
  Apply Template (single + bulk), and any other collection-target
  control.
- **Edit Window with rename** -- editing a window's name re-keys it
  via delete + recreate (MECM keys windows by name). Same-name edits
  use the cheaper Set path.
- **Modal dialogs** -- New / Edit Maintenance Window, New / Edit
  Template, Apply Template (single + bulk), Remove confirmations,
  Options (Connection, About). All MetroWindow inline-XAML, theme-
  honoring, drag-fallback installed.
- **Core module** `MaintWindowMgrCommon.psm1` with 30 exported
  functions covering logging, CM connection, data retrieval, CRUD,
  schedule helpers (including Patch Tuesday calculation and offset
  scheduling), template management, bulk operations, folder-tree
  enumeration, and CSV / HTML export.
- **WPF brand alignment** -- MahApps.Metro shell with sidebar
  navigation, glyph status (no red/green for state), animated
  ProgressRing during refresh, log drawer, status bar, dark and
  light themes with runtime toggle, window-state persistence
  including legacy WinForms schema bridge.

### Stack

- PowerShell 5.1 + .NET Framework 4.7.2+
- WPF + MahApps.Metro 2.4.10 (vendored under `Lib/`)
- ConfigurationManager PowerShell module (CM console required on
  the host machine running this app)
