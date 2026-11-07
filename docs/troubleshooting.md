# Troubleshooting

This page lists the most common problems of Maintenance Window Manager with cause and fix. See [Error messages](reference/error-messages.md) for every message.

## What it does

It gives a short fix for each common problem.

## When to use it

Use it when a refresh, a change or an export does not work.

## Requirements

None.

## Problems

| Message or symptom | Cause | Fix |
|---|---|---|
| Refresh does nothing and the log asks you to open Options. | The site code or the provider is empty. | Set the connection. |
| `Refresh failed: Failed to connect to site <code> (provider <name>).` | The console is not installed on this computer, or the provider name is wrong, or `SMS_ADMIN_UI_PATH` does not point to the `bin\i386` folder of the console. | Check the provider name. Install the console. If it is installed, set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder and start the tool again. |
| The Windows, Coverage or Templates view is empty. | No refresh has run, or no template file loaded. | Select **Refresh**. For templates, check the `Templates` folder. |
| The preview shows `no future occurrences`. | A one-time start is in the past. | Set a start date in the future. |
| **Edit...**, **Toggle Enabled** or **Remove...** does nothing. | No row is selected. | Select a window. |
| A rename left the collection without the window. | The new window was refused after the old one was deleted. | Create the window again. |
| The settings are lost. | The folder is not writable. | Move the tool to a writable folder. |

## Steps

1. Read the log pane. It names the cause of most problems.
2. Read the session log in `Logs`.
3. Do the fix of the table.

## Limits

None.

## Links

- [Error messages](reference/error-messages.md)
- [Files and logs](reference/files-and-logs.md)
