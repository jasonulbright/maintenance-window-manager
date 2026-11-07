# Maintenance Window Manager: test plan

Measured 2026-11-06 with `Invoke-Checks.ps1` (Pester 5.7.1, PSScriptAnalyzer 1.25.0). Counts are tests as run; each data row counts as one test. Coverage is the share of module commands that the tests run.

| | Before | After |
|---|---|---|
| Tests | 35 | 410 on Windows PowerShell 5.1 and on PowerShell 7 (16 contract tests skip on a machine without the console) |
| Module coverage | 20.4% (114 of 560 commands) | 97.6% (638 of 654 commands) |
| Lab steps | none | 34 (`Tests/Lab`) |

Run: `.\Invoke-Checks.ps1` (lint, tests, coverage floor 80%). It runs in CI on both hosts (`.github/workflows/checks.yml`).

Window guard: `Window.Tests.ps1` opens windows only when the environment variable `SUITE_ALLOW_WINDOW_TESTS` is `1`, and never on the dev PC (computer name `AYW`), even with the variable set. Otherwise Pester reports those tests as Skipped, not as Failed. Pester does not store the reason, so `Invoke-Checks.ps1` (and `Window.Tests.ps1`, when it runs alone) prints it to the console, once. `Test-WindowTestsAllowed` in `Tests/WindowHarness.ps1` makes the decision, and `Get-WpfSession` refuses to create a runspace when it says no, so no window can open. CI sets the variable in both jobs. The coverage floor stays in force: the window tests run no module code, so skipping them does not lower module coverage. `WindowGuard.Tests.ps1` tests the guard with injected values and with the real environment, and checks that every test file that can open a window, in any folder, applies the guard to each top-level `Describe`. It opens no window. The test counts above were measured before the guard and are not measured again here.

## What is covered

| Area | File | Contents |
|---|---|---|
| Read side | `Module.Read.Tests.ps1` | Folders, all windows (type and recurrence names, N/A, null duration, 500 collections), summaries, coverage gaps, human schedule text for every recurrence, next occurrences (order, time of day, weekday) |
| Schedules | `Module.Schedule.Tests.ps1` | Patch Tuesday for 252 months, `New-WindowSchedule` for every recurrence and duration rule, Patch Tuesday schedules |
| Write actions | `Module.Write.Tests.ps1` | Create, update, remove, template save, load and apply, copy, bulk enable, CSV import with per-row checks. Success, failure, partial failure, `-WhatIf`, hostile names |
| Export | `Module.Export.Tests.ps1` | CSV and HTML from tables and from window objects, encoding, classes, empty list |
| Script logic | `Shell.Logic.Tests.ps1` | Glyphs, filters of the three views, status bar text, preferences file |
| Script calls | `Script.Contract.Tests.ps1` | Every call from the script and the module into an app function or SuiteCommon names only parameters that exist |
| Real cmdlet signatures | `CmCmdlet.Contract.Tests.ps1` | Every ConfigMgr call names only parameters that the real cmdlet has. Runs where the console is installed |
| Start without the console | `Startup.Tests.ps1` | Plain message, `$false`, no raw exception |
| Windows | `Window.Tests.ps1` | Main window and dialogs: names, titles, no clipped text and nothing off screen at 1024x720 and 1440x900 in both themes and all views, text and control contrast |
| Window guard | `WindowGuard.Tests.ps1` | The guard with injected computer name and opt-in (dev PC refused, case-insensitive, opt-in needed, reason text) and with the real environment, the session refusal, and the guard on every top-level `Describe` of every window test file. Opens no window |
| Lab | `Lab/Invoke-MaintenanceWindowLab.ps1` | Two ZZTEST- collections on CM01: every schedule type, create, update, bulk enable, copy, template, CSV preview, remove, WhatIf. Checks that no other maintenance window changed |

## Defects found

| # | What | Fix |
|---|---|---|
| 1 | A duration of 1 hour 30 minutes was saved as 1 hour. A count of minutes above 59 is stored wrong by the site (90 became 60, found in the lab) | Whole hours in the cmdlet call; the minutes are set on the schedule object |
| 2 | Export CSV and HTML called the export functions with `-Windows`, which they did not have | Both functions accept `-Windows` and `-Columns` |
| 3 | The HTML export did not encode names | HTML encoding of values, columns and title |
| 4 | `Import-MaintenanceWindowsCsv` threw a raw error for a text number and passed an hour of 99 or a day named Funday as valid | Per-row checks; `Valid` and `Error` are set |
| 5 | Two template names with the same file name replaced each other (`Patch + 7` and `Patch - 7`); a name like CON made a device file name | Numbered file names, safe names |
| 6 | `Remove-WindowTemplate` deleted any file that was passed | Only `.json` files |
| 7 | Templates read without an encoding; an empty template file passed silently | UTF-8; a warning for an empty file |
| 8 | Templates buttons Delete and Apply were off screen at 1024 pixels | `WrapPanel` |
| 9 | `Get-CollectionsWithoutWindows` rejected a null window that its code handles | `AllowNull` |

## Not covered, and why

- Contrast of disabled controls, the sidebar button border, the light text boxes and the unselected tab headers misses the standard (text 4.5:1, borders 3:1). The cause is the shared SuiteCommon palette and the MahApps styles. `Tests/KnownContrastDefects.ps1` lists each violation. The window test fails on a new one and on a listed one that is fixed. The fix belongs in `app-packager-suite`.
- The dialogs are not shown with data, so their layout and contrast are not measured. The title bar of an inactive window misses 4.5:1 (listed as `titlebar:inactive`). The border of a disabled sidebar button is read from the button, not from the template border at half opacity.
- The schedule editor in the New Window and Template dialogs reads numeric text boxes with `[int]`. An empty box throws. The dialogs build inside the script and need a live window; not driven here.
- The refresh handler, the busy overlay and the dialogs' button logic run inside the window with a site. The lab run covers the module calls that they make, not the clicks.
- `Get-NextOccurrences` for a monthly-by-weekday window that starts in the fifth week approximates the date. The tests check order, count and weekday only.
- The settings save in SuiteCommon is not atomic. The fix belongs in `app-packager-suite`.
