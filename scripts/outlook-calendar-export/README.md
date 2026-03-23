# Outlook Calendar Export

Exports calendar entries from the Outlook Desktop application to a JSON file for syncing to another calendar (e.g., Google Calendar). Uses the Outlook COM automation interface — no API or cloud access required.

## What It Does

- Extracts calendar entries from a configurable date range (default: 15 days back, 90 days forward)
- Outputs a structured JSON file with scheduling metadata
- Handles both one-time and recurring events, including full recurrence pattern details
- Provides verbose, color-coded console output and timestamped log files

## What It Does NOT Export

To protect sensitive information, the following are **excluded** from the export:

- Meeting body/notes content
- Attendee/participant lists
- Organizer name and email address (only the email **domain** is kept, e.g., `google.com`)
- Attachments
- HTML or RTF content
- Sensitivity details beyond the busy status

Only scheduling metadata needed for calendar sync is exported.

## Prerequisites

| Requirement | Details |
|---|---|
| Operating System | Windows 10 or Windows 11 |
| Outlook | Microsoft Outlook Desktop app, installed and **running** with a configured account |
| PowerShell | 5.1 or later (ships with Windows — verify with `$PSVersionTable.PSVersion`) |
| Admin Rights | Not required |
| Additional Modules | None — uses only built-in PowerShell and Outlook COM |

## How to Run

### Step 1: Get the scripts

Clone or download this repository, then open PowerShell and navigate to the script directory:

```powershell
cd path\to\script-sink\scripts\outlook-calendar-export
```

### Step 2: Check execution policy

If you've never run PowerShell scripts before, you may need to allow script execution:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### Step 3: Run the diagnostic test first

Before running the full export, verify that Outlook COM access works:

```powershell
.\Test-OutlookAccess.ps1
```

This will:
- Test the connection to Outlook
- Extract today's calendar entries
- Display them in a table on screen
- Save a CSV to `output\test-calendar.csv` (EntryID, OrganizerDomain, Start, End, TimeZone, ResponseStatus)

If this fails, check the [Troubleshooting](#troubleshooting) section.

### Step 4: Run the full export

```powershell
.\Export-OutlookCalendar.ps1
```

The script will display verbose, step-by-step progress on screen and write a log file to the `logs\` directory.

Output is saved to `output\calendar-export.json` by default.

### CLI Parameter Examples

Override any default or config.json value via CLI parameters:

```powershell
# Export 30 days back and 120 days forward
.\Export-OutlookCalendar.ps1 -DaysBack 30 -DaysForward 120

# Custom output path
.\Export-OutlookCalendar.ps1 -OutputPath "C:\Exports\my-calendar.json"

# Use a different config file
.\Export-OutlookCalendar.ps1 -ConfigPath "C:\config\my-config.json"
```

## Configuration

Settings are resolved in this priority order: **CLI parameters > config.json > built-in defaults**.

The `config.json` file is optional. If not present, the script uses built-in defaults and logs a warning.

| Key | Type | Default | Description |
|---|---|---|---|
| `daysBack` | integer | `15` | Number of days in the past to include |
| `daysForward` | integer | `90` | Number of days in the future to include |
| `outputPath` | string | `./output/calendar-export.json` | Path for the JSON output file (relative to script dir, or absolute) |
| `logPath` | string | `./logs/` | Directory for log files (relative to script dir, or absolute) |

Example `config.json`:

```json
{
  "daysBack": 15,
  "daysForward": 90,
  "outputPath": "./output/calendar-export.json",
  "logPath": "./logs/"
}
```

## Output Format

The exported JSON has this structure:

```json
{
  "exportDate": "2026-03-23T10:00:00Z",
  "rangeStart": "2026-03-08",
  "rangeEnd": "2026-06-21",
  "itemCount": 142,
  "entries": [
    {
      "entryId": "00000000...",
      "lastModified": "2026-03-20T14:30:00Z",
      "subject": "Weekly Standup",
      "start": "2026-03-24T09:00:00",
      "startTimeZone": "Eastern Standard Time",
      "end": "2026-03-24T09:30:00",
      "endTimeZone": "Eastern Standard Time",
      "location": "Teams Meeting",
      "organizerDomain": "google.com",
      "busyStatus": "Busy",
      "responseStatus": "Accepted",
      "isAllDay": false,
      "isRecurring": true,
      "recurrencePattern": {
        "type": "Weekly",
        "interval": 1,
        "daysOfWeek": ["Monday", "Wednesday", "Friday"],
        "dayOfMonth": 0,
        "monthOfYear": 0,
        "instance": 0,
        "patternStart": "2026-01-06",
        "patternEnd": null,
        "occurrences": 0
      }
    }
  ]
}
```

### Field Reference

| Field | Description |
|---|---|
| `entryId` | Outlook unique identifier — use this as the sync key |
| `lastModified` | UTC timestamp of last modification — use to detect updates |
| `subject` | Meeting/event title |
| `start` / `end` | Local date-time of the occurrence |
| `startTimeZone` / `endTimeZone` | Windows timezone ID (e.g., `Eastern Standard Time`, `Pacific Standard Time`) — falls back to system local timezone if unavailable |
| `location` | Meeting location (room, link, etc.) or null |
| `organizerDomain` | Email domain of the organizer (e.g., `google.com`) — no names or full emails are stored |
| `busyStatus` | One of: `Free`, `Tentative`, `Busy`, `OutOfOffice`, `WorkingElsewhere` |
| `responseStatus` | One of: `None`, `Organized`, `Tentative`, `Accepted`, `Declined`, `NotResponded` |
| `isAllDay` | `true` if this is an all-day event |
| `isRecurring` | `true` if this is an occurrence of a recurring event |
| `recurrencePattern` | Recurrence details (null for non-recurring events) |

### Recurrence Pattern Fields

| Field | Description |
|---|---|
| `type` | `Daily`, `Weekly`, `Monthly`, `MonthlyNth`, `Yearly`, `YearlyNth` |
| `interval` | Frequency (e.g., `2` = every 2 weeks for a weekly recurrence) |
| `daysOfWeek` | Array of day names (for weekly recurrences) |
| `dayOfMonth` | Day of month (for monthly/yearly recurrences) |
| `monthOfYear` | Month number (for yearly recurrences) |
| `instance` | Week instance for MonthlyNth/YearlyNth (1=first, 2=second, ..., 5=last) |
| `patternStart` | Date the recurrence pattern begins |
| `patternEnd` | Date the recurrence pattern ends (null if no end date) |
| `occurrences` | Total occurrences if bounded (0 if unbounded) |

## Scheduling as a Windows Task

To run this export automatically:

1. Open **Task Scheduler** (`taskschd.msc`)
2. Create a new task with these settings:
   - **Action**: Start a program
   - **Program**: `powershell.exe`
   - **Arguments**: `-ExecutionPolicy Bypass -File "C:\path\to\Export-OutlookCalendar.ps1"`
   - **Start in**: `C:\path\to\scripts\outlook-calendar-export`
3. Set your desired trigger (e.g., every morning at 7:00 AM)
4. Under **Conditions**, uncheck "Start only if the computer is on AC power" if on a laptop

Note: Outlook Desktop must be running when the scheduled task executes.

## Troubleshooting

| Problem | Solution |
|---|---|
| "Could not connect to Outlook" | Make sure Outlook Desktop is open and running (not just in the system tray minimized to nothing) |
| Execution policy error | Run `Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned` |
| "Could not open calendar folder" | Your default calendar may not be set. Open Outlook, go to Calendar, and ensure a default account is configured |
| Empty results but calendar has items | Check that the date range covers the expected period. Try `-DaysBack 0 -DaysForward 1` to get just today |
| COM error on specific items | Some corrupted or restricted items may fail. The script logs these and continues. Check the log file for details |
| Script runs but Outlook prompts for security | Outlook may show a security dialog about programmatic access. Click "Allow" or configure the Trust Center to allow COM access |
