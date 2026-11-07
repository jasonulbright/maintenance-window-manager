# Options dialog

The Options dialog sets the connection.

## What it does

The dialog has two pages: **Connection** and **About**. **OK** saves. **Cancel** or Esc closes the dialog and keeps the old values.

## When to use it

Use it at the first start and when the site changes.

## Requirements

None.

## Steps

### Connection

Select **Options** in the sidebar.

| Field | Meaning |
|---|---|
| **Site Code** | The site code, for example `MCM`. |
| **SMS Provider FQDN** | The server with the SMS Provider role. The tool uses it for the PSDrive root. |

![The Connection page](../images/options-connection.png)

### About

The page shows the version and a short description.

![The About page](../images/options-about.png)

### Theme

The **Dark Theme** switch at the bottom of the sidebar changes the theme. The tool keeps the choice.

![The Windows view in the light theme](../images/light-windows.png)

![The Coverage view in the light theme](../images/light-coverage.png)

![The Templates view in the light theme](../images/light-templates.png)

![The Options dialog in the light theme](../images/light-options.png)

### Start again

The tool saves the open view and the window place. After a restart the window shows the same view.

![The window after a restart](../images/start-reopen.png)

## Options and settings of this page

See the table above. [Settings](../reference/settings.md) lists the file keys.

## Limits

- **Refresh** needs both the site code and the provider.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The status bar asks you to open Options. | The site code or the provider is empty. | Fill both fields. |

## Links

- [Settings](../reference/settings.md)
- [Getting started](../getting-started.md)
