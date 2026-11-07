# Schedule editor

The schedule editor is part of the New Window, Edit Window and Template dialogs.

## What it does

It sets when a maintenance window starts and how long it lasts. The **Next 5 occurrences** box shows the next five start times for the settings that you chose. The box updates when you change a setting.

## When to use it

Use it in each dialog that creates or changes a schedule.

## Requirements

None. The editor needs no connection.

## Steps

### Choose the recurrence

Select a value in **Recurrence**:

| Recurrence | Settings |
|---|---|
| **One-time** | A date and a time. |
| **Daily** | A start date and a time. The window repeats every day. |
| **Weekly** | The **Day of Week**. |
| **MonthlyByDate** | The **Day of Month**, 1 to 31. |
| **MonthlyByWeekday** | The **Week Order** (First, Second, Third, Fourth or Last) and the **Day of Week**. |
| **PatchTuesday** | An offset in days after the second Tuesday of each month. |

The settings that do not belong to the recurrence are hidden.

### Set the start and the duration

1. Select the **Start Date** and type the **Hour** and the **Minute**.
2. Type the **Duration Hours** and the **Duration Min**.
3. Select **UTC time** if the times are UTC.

### Read the preview

The **Next 5 occurrences** box shows the dates and times, for example `2026-11-10  Tue  02:00`. For a weekly schedule, the first line is the first selected weekday on or after the start date. For a monthly schedule, it is the selected day in the month of the start date.

![A weekly schedule that starts on a Saturday and fires on Tuesdays](../images/dialog-new-window.png)

For a one-time schedule with a start date in the past, the box shows a hint that the start date must be in the future.

## Options and settings of this page

| Control | Effect |
|---|---|
| **Recurrence** | Chooses the kind of schedule. |
| **Start Date**, **Hour**, **Minute** | Set the first start. |
| **Duration Hours**, **Duration Min** | Set the length of the window. |
| **UTC time** | Reads the times as UTC. |
| **Day of Week**, **Week Order**, **Day of Month**, offset | Set the day, by recurrence. |

## Limits

- The preview comes from the settings of the editor. The site computes the real schedule.
- A day of month above the days of the month in the start month is cut to the last day.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| The preview shows `no future occurrences`. | A one-time start is in the past, or the recurrence is not set. | Set a start date in the future. |

## Links

- [Windows](windows.md)
- [Templates](templates.md)
