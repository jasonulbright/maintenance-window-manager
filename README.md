# Maintenance Window Manager

A MahApps.Metro WPF GUI for auditing, creating, editing, and bulk-applying MECM device collection maintenance windows. The MECM console buries maintenance windows deep in individual collection properties; this tool gives you a single-pane view of every window in the environment with full CRUD, schedule editing, templates, and gap analysis.

![Maintenance Window Manager](screenshot.png)

## Requirements

- Windows 10 / 11 or Server 2016+
- PowerShell 5.1
- .NET Framework 4.7.2+
- Configuration Manager console installed (provides the `ConfigurationManager` PowerShell module)
- MECM RBAC rights to read and modify collections and maintenance windows

## Quick Start

1. Download the release zip and extract it to a working folder.
2. Right-click `start-maintenancewindowmgr.ps1` -> **Run with PowerShell**, or from a PowerShell prompt:

   ```powershell
   powershell -ExecutionPolicy Bypass -File start-maintenancewindowmgr.ps1
   ```
3. Click **Options** on the sidebar and set Site Code and SMS Provider.
4. Click **Refresh** to load every collection and every maintenance window.

## Layout

The shell uses a sidebar layout with three views and an Options modal:

- **Windows** -- master-detail grid of every maintenance window across every collection. Glyph status (enabled vs muted), Collection, Window, Type, Recurrence, human-readable Schedule, Duration, Next Occurrence. Filter by collection or window name + status filter (Enabled / Disabled / SoftwareUpdates / TaskSequences). Detail tabs: Properties, projected Next 5 occurrences. Action bar: New Window, Edit, Toggle Enabled, Remove.
- **Coverage** -- per-collection gap analysis. Glyph: check for collections with at least one window, warn for those with zero. Multi-select rows + **Apply Template...** rolls a saved schedule out to many collections in one shot.
- **Templates** -- 5 baseline schedule templates ship with the app (Daily After Hours, Monthly First Saturday, Patch Tuesday +3 days, Patch Tuesday +7 days, Weekly Sunday 2 AM). Edit any of them, or author your own with **New Template**.

## Schedule Editor

Embedded in New Window, Edit Window, and the Template Editor. Recurrence options:

- **One-time** -- single date + time
- **Daily** -- every N days at a given time
- **Weekly** -- pick day-of-week
- **Monthly by Date** -- pick day-of-month (1-31)
- **Monthly by Weekday** -- pick week order (First / Second / Third / Fourth / Last) + day-of-week
- **Patch Tuesday +N days** -- shifts each month's second-Tuesday by an offset

Subpanels appear/disappear with the recurrence choice. The **Next 5 occurrences** preview re-projects on every keystroke so you can see exactly what the schedule will fire as before you commit.

## Workflows

### Create a window on one collection

1. **Windows** view -> **New Window...**
2. **Browse...** the folder tree to pick a target collection (built-ins are filtered out).
3. Set name, window type (Any / SoftwareUpdatesOnly / TaskSequencesOnly), enabled flag.
4. Configure the schedule. Watch the preview update.
5. **Create**.

### Edit or rename a window

1. Pick the window in the **Windows** grid -> **Edit...**
2. Adjust name, type, enabled, or any schedule field.
3. If you rename, the app issues delete + recreate (MECM keys windows by name); same-name edits use the cheaper Set path.

### Roll a template out to many collections

1. **Coverage** view, multi-select target collections (anything missing a window is glyphed).
2. **Apply Template...**
3. Pick the template from the dropdown. Optionally set a window-name override.
4. **Apply**. Per-collection success/failure prints to the log drawer.

### Author a custom template

1. **Templates** view -> **New Template...**
2. Name + description + window type + schedule.
3. **Create**. Template lands as JSON under `Templates/`.

## Project Structure

```
maintenance-window-manager/
+- start-maintenancewindowmgr.ps1            # WPF shell
+- MainWindow.xaml                           # Main window layout
+- Lib/                                      # Vendored MahApps.Metro 2.4.10
+- Module/
|  +- MaintWindowMgrCommon.psd1              # Module manifest
|  \- MaintWindowMgrCommon.psm1              # Business logic (30 functions)
+- Templates/                                # Bundled schedule templates (JSON)
+- Logs/                                     # Session logs (per-run)
+- Reports/                                  # CSV / HTML exports
+- CHANGELOG.md
+- LICENSE
\- README.md
```

## Safety

- Built-in collections (SMS prefix) are filtered out of target pickers so you can't accidentally apply a window to system collections.
- Confirmation dialogs before every destructive operation (window removal, template deletion).
- Rename = delete + recreate. The app logs the original name before deletion so you can recover the spec from the session log if the recreate fails.
- All MECM data access uses supported ConfigurationManager PowerShell cmdlets (Get-CMDeviceCollection, Get-CMMaintenanceWindow, New-CMMaintenanceWindow, Set-CMMaintenanceWindow, Remove-CMMaintenanceWindow). No direct WMI queries.

## License

This project is licensed under the [MIT License](LICENSE).

## Author

Jason Ulbright
