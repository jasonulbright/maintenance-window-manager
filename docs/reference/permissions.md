# Permissions

This page lists the rights that Maintenance Window Manager needs.

## What it does

A refresh reads the site. The buttons that change a window use the rights of the account that started the tool.

## When to use it

Use it to set up an account, or to find out why a read or a change is refused.

## Requirements

| Action | Right |
|---|---|
| Start the tool | Read and write in the folder of the script. |
| Read and reload the templates | Read access to the `Templates` folder. No site access is needed. |
| Create, edit and delete a template | Read and write access to the `Templates` folder. No site access is needed. |
| **Refresh** and all views | The Configuration Manager console on the computer, and read access to the SMS Provider and to collections under your Configuration Manager role. |
| **New Window...**, **Edit...**, **Toggle Enabled**, **Remove...** | The right to modify the collection, which includes its maintenance windows. |
| **Apply Template...** and **Apply...** | The same right on every target collection. |

## Steps

1. Start the tool as the account that has the Configuration Manager role.
2. Select **Refresh** to test the read access.

## Limits

- The tool does not check the rights before a change. A refused change shows in the log as a failure.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Refresh failed:` | The account has no access to the provider, or the console is missing. | Check the role, the provider name and the console. |
| `<action> failed. The log file has the details.` in the log | The action failed. | Open the log file and read the reason. Check the permissions only when it shows an access failure. |

## Links

- [Network](network.md)
- [Windows](../core/windows.md)
