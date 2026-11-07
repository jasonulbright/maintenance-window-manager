# Capture the images again

One command regenerates every image of this guide.

## What it does

`tools\guide-capture\Invoke-GuideCapture.ps1` copies the tool to a lab computer, starts it in an interactive session of a lab account, walks every view, dialog and setting with Windows UI Automation, takes the screenshots at 1440 by 900 pixels, copies the PNG files back to `docs\images` and removes its files from the lab computer. It replaces only the images that exist already.

## When to use it

Use it after a change that alters the window, and before a release.

## Requirements

- A lab computer with the Configuration Manager console, an SMS Provider that you can read, and PowerShell remoting for an account.
- A script, outside the repository, that sets `$LabCredential` to the credential of that account.
- A domain account that has an interactive session on the lab computer. The capture runs as a scheduled task with an interactive logon.
- Lab data only. The capture reads the site. It opens the dialogs of New Window, Edit, Apply Template, Edit Template, New Template and Apply, and selects **Cancel**. It never creates, changes or removes a window or a template.

## Steps

1. Run this command from the repository folder:

```powershell
.\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript C:\lab\credentials.local.ps1 -Target CLIENT01 -LogonUser contoso\labadmin -SmsProvider cm01.contoso.com
```

2. Wait for the command to end. It prints a warning for each step that failed.
3. Open every changed image and check it. Look for clipped text, empty grids and stale version strings.
4. Run the documentation test: `Invoke-Pester .\Tests\Documentation.Tests.ps1`.

A second run inside the command does the missing-console test on a copy of the tool whose module lookup returns nothing.

## Limits

- The capture needs an interactive session on the lab computer. A locked session breaks it.
- The images show the data of the lab site.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `The capture did not finish` | The session is locked or a window did not show. | Unlock the session and run again. |
| A step shows `control not found`. | The window changed. | Update the step in `steps-guide.ps1`. |

## Links

- [Index](../index.md)
