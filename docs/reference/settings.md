# Settings

This page lists what Maintenance Window Manager saves and where.

## What it does

The tool saves its settings in `MaintWindowMgr.prefs.json` in the folder of the script. The Options dialog writes the file when you select **OK**. The window state is in `MaintWindowMgr.windowstate.json`.

## When to use it

Use it to copy the settings to another computer or to reset them.

## Requirements

None.

## Settings

| Key | Option | Default |
|---|---|---|
| `SiteCode` | **Site Code** | Empty. |
| `SMSProvider` | **SMS Provider** | Empty. |
| `DarkMode` | **Dark Theme** switch | `true`. |

`MaintWindowMgr.windowstate.json` keeps the size, the place and the open view of the window.

## Steps

### Reset the settings

1. Close the tool.
2. Delete `MaintWindowMgr.prefs.json` and `MaintWindowMgr.windowstate.json`.
3. If the suite launcher started the tool, clear the environment variables `SUITE_CM_SITECODE` and `SUITE_CM_PROVIDER` too. The tool takes the site code and the provider from them when the settings are empty.
4. Start the tool. The window of the first start shows.

## Limits

- The files are in the folder of the script. Extract the tool to a folder that you can write to.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The settings do not stick. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Options dialog](../core/options.md)
- [Files and logs](files-and-logs.md)
