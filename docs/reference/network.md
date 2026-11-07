# Network

This page lists every outbound call of Maintenance Window Manager.

## What it does

The tool talks to the SMS Provider through the console module. When it is installed by the suite installer, it also checks GitHub once a day for a newer suite version. The WQL queries that you send to the SMS Provider can contain values that you typed, for example a device name. The GitHub update check sends no user data and no telemetry.

## When to use it

Use it to open the firewall for the tool, or to check the traffic.

## Requirements

| Call | Target | When | Content |
|---|---|---|---|
| ConfigurationManager module | The SMS Provider in **Options > Connection**, through the site drive that the console module opens. The ports are those of the console. | Every refresh, and every button that changes the site. | `Get-CM*` reads and the changes of maintenance windows. |
| `https://api.github.com/repos/<repository>/releases/latest` | HTTPS to GitHub. | At the start of the window, once a day, only when the tool is installed by the suite installer. | The check for a newer suite version. The result goes to the log. |

## Steps

None.

## Limits

- A plain extract of the zip (no `suite-manifest.json`) makes no call to GitHub. A failed update check is ignored. The tool does not change a proxy setting.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `Refresh failed:` | The provider is not reachable, or the console is missing. | Check the name, the firewall and the role. For a missing console see [Error messages](error-messages.md). |

## Links

- [Permissions](permissions.md)
- [Options dialog](../core/options.md)
