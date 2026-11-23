# Settings

This page lists what Maintenance Window Manager saves and where.

## What it does

The tool saves its settings in two places. The global setting (`DarkMode`) is in `MaintWindowMgr.prefs.json` in the folder of the script. When the launcher chose a site, `SiteCode` and `SMSProvider` are in `%LOCALAPPDATA%\AppPackagerData\sites\<SITE>\tools\maintenance-window-manager\prefs.json`, for example `sites\AB1\tools\maintenance-window-manager\prefs.json`. The Options dialog writes the files when you select **OK**. The window state is in `MaintWindowMgr.windowstate.json`.

## When to use it

Use it to copy the settings to another computer or to reset them.

## Requirements

None.

## Settings

| Key | Option | Default |
|---|---|---|
| `SiteCode` | **Site Code** | Empty. With a site, the value is the site that the launcher chose, and the field is read-only. |
| `SMSProvider` | **SMS Provider** | Empty. Each site keeps its own value. |
| `DarkMode` | **Dark Theme** switch | `true`. |

`MaintWindowMgr.windowstate.json` keeps the size, the place and the open view of the window.

## Steps

### Reset the settings

1. Close the tool.
2. Delete `MaintWindowMgr.prefs.json` and `MaintWindowMgr.windowstate.json`. To reset the connection of one site, delete `prefs.json` in the folder of that site too.
3. If the suite launcher started the tool, clear the environment variables `SUITE_CM_SITECODE`, `APP_PACKAGER_SITE` and `SUITE_CM_PROVIDER` too. A valid site code in `SUITE_CM_SITECODE` or `APP_PACKAGER_SITE` chooses the site. The variable `SUITE_CM_PROVIDER` fills an empty provider only when no site is chosen. With a site, the provider comes from the site.
4. Start the tool. The window of the first start shows.

## Limits

- The global files are in the folder of the script. Extract the tool to a folder that you can write to.
- When no launcher chose a site, the `SiteCode` in `MaintWindowMgr.prefs.json` chooses the site. To change it, edit that file, or choose the site in the launcher.
- The tool never changes the site code. Choose the site in the AppPackager Suite launcher and restart the tool.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The settings do not stick. | The folder is not writable. | Move the tool to a writable folder. |

## Links

- [Options dialog](../core/options.md)
- [Files and logs](files-and-logs.md)
