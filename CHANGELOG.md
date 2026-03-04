# Changelog

All notable changes to Maintenance Window Manager are documented in this file.

## [1.0.0] - 2026-03-03

### Added
- **Tab 1: All Windows** -- flat audit view of every maintenance window across all device collections; summary cards (Total Windows, Collections Without Windows, Disabled, Upcoming 7 days); text filter, type dropdown, show/hide disabled; detail panel with next 5 occurrences; right-click context menu (Edit, Delete, Enable/Disable toggle, Clone to Collections, Save as Template)
- **Tab 2: By Collection** -- vertical split with searchable collection list (left) and collection-specific windows grid (right); zero-window collections highlighted in red; CRUD buttons (New Window, Edit, Delete); "Without Windows Only" toggle filter
- **Tab 3: Templates** -- template library with grid and preview panel; 5 shipped default templates (Patch Tuesday +3, Patch Tuesday +7, Weekly Sunday 2 AM, Daily After Hours, Monthly First Saturday); New Template, Delete, Apply to Collections buttons
- **Tab 4: Bulk Operations** -- operation selector (Import from CSV, Copy Window, Bulk Enable/Disable, Bulk Delete); preview grid with color-coded Result column; results log panel
- **Schedule Builder dialog** -- modal form for creating/editing maintenance windows; recurrence type (OneTime, Daily, Weekly, MonthlyByDate, MonthlyByWeekday); dynamic controls for day-of-week, week order, day-of-month; DateTimePicker for start time; duration hours/minutes; window type (Any, SoftwareUpdatesOnly, TaskSequencesOnly); UTC and Enabled checkboxes
- `MaintWindowMgrCommon.psm1` module with 22 exported functions: logging, CM connection, data retrieval (WMI bulk query + CM cmdlets), CRUD, schedule helpers (including Patch Tuesday calculation), template management, bulk operations, CSV/HTML export
- 5 shipped JSON window templates in Templates/ folder
- Dark/light theme, window state persistence, preferences dialog
