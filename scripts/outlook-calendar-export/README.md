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
- Attendee and organizer names and email addresses (only email **domains** are kept, e.g., `google.com`)
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

### Step 5: Detect changes (incremental sync)

After running the full export, run the change detection script:

```powershell
.\Detect-CalendarChanges.ps1
```

This will:
- Load the export JSON and compare against the previous run
- Output only new, modified, and deleted entries to `output\calendar-changes.json`
- Save the current state to `output\last-run.json` for the next comparison
- Append a row to `output\run-history.csv` (viewable in Excel, max 500 rows)
- Write a log to `logs\changes-YYYYMMDD-HHmmss.log`

On the first run (or if you delete `last-run.json`), all entries are reported as "new".

To force a full run at any time:

```powershell
Remove-Item .\output\last-run.json
.\Detect-CalendarChanges.ps1
```

### Step 6: Upload changes to Sink API

After detecting changes, upload them to the Sink API for Google Calendar sync:

```powershell
# Using a CLI token
.\Upload-CalendarChanges.ps1 -Token "your-sink-token"

# Using the SINK_TOKEN environment variable
$env:SINK_TOKEN = "your-sink-token"
.\Upload-CalendarChanges.ps1
```

This will:
- Load the changes JSON from `output\calendar-changes.json`
- POST it to the Sink calendar upload API endpoint
- Log the server response (upload ID, processed/created/updated/deleted counts)
- Append a row to `output\upload-history.csv` (viewable in Excel, max 500 rows)
- Write a log to `logs\upload-YYYYMMDD-HHmmss.log`
- Skip the upload if there are zero changes (still records in history CSV)

Token resolution order: `-Token` CLI parameter > `SINK_TOKEN` environment variable > error and exit.

### CLI Parameter Examples

Override any default or config.json value via CLI parameters:

```powershell
# Export: 30 days back and 120 days forward
.\Export-OutlookCalendar.ps1 -DaysBack 30 -DaysForward 120

# Export: custom output path
.\Export-OutlookCalendar.ps1 -OutputPath "C:\Exports\my-calendar.json"

# Export: use a different config file
.\Export-OutlookCalendar.ps1 -ConfigPath "C:\config\my-config.json"

# Changes: custom paths
.\Detect-CalendarChanges.ps1 -ExportPath "C:\Exports\my-calendar.json" -ChangesOutputPath "C:\Exports\changes.json"

# Upload: explicit token
.\Upload-CalendarChanges.ps1 -Token "my-secret-token"

# Upload: custom endpoint and changes file
.\Upload-CalendarChanges.ps1 -Endpoint "https://custom-host/api/calendar/entries/upload" -ChangesPath "C:\Exports\changes.json"
```

## Configuration

Settings are resolved in this priority order: **CLI parameters > config.json > built-in defaults**.

The `config.json` file is optional. If not present, the script uses built-in defaults and logs a warning.

| Key | Type | Default | Used By | Description |
|---|---|---|---|---|
| `daysBack` | integer | `15` | Export | Number of days in the past to include |
| `daysForward` | integer | `90` | Export | Number of days in the future to include |
| `outputPath` | string | `./output/calendar-export.json` | Export | Path for the JSON export file |
| `logPath` | string | `./logs/` | Both | Directory for log files |
| `changesOutputPath` | string | `./output/calendar-changes.json` | Changes | Path for the changes-only JSON |
| `lastRunPath` | string | `./output/last-run.json` | Changes | Path to the last-run state file (delete to force full run) |
| `runHistoryPath` | string | `./output/run-history.csv` | Changes | Path to the run history CSV |
| `runHistoryMaxRows` | integer | `500` | Changes | Maximum rows to keep in the run history CSV |
| `uploadEndpoint` | string | `https://sink.marin.cr/api/calendar/entries/upload` | Upload | Sink API endpoint for calendar uploads |
| `uploadHistoryPath` | string | `./output/upload-history.csv` | Upload | Path to the upload history CSV |
| `uploadHistoryMaxRows` | integer | `500` | Upload | Maximum rows to keep in the upload history CSV |

All paths can be relative (resolved against the script directory) or absolute.

Example `config.json`:

```json
{
  "daysBack": 15,
  "daysForward": 90,
  "outputPath": "./output/calendar-export.json",
  "logPath": "./logs/",
  "changesOutputPath": "./output/calendar-changes.json",
  "lastRunPath": "./output/last-run.json",
  "runHistoryPath": "./output/run-history.csv",
  "runHistoryMaxRows": 500
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
      "attendeeCount": 5,
      "attendeeDomains": ["contoso.com", "google.com"],
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
| `attendeeCount` | Total number of meeting attendees |
| `attendeeDomains` | Sorted array of unique attendee email domains (e.g., `["contoso.com", "google.com"]`) — no names or full emails |
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

## Changes Output Format

The `Detect-CalendarChanges.ps1` script outputs a changes-only JSON (`calendar-changes.json`):

```json
{
  "detectedAt": "2026-03-24T10:05:00Z",
  "previousRunDate": "2026-03-24T10:00:00Z",
  "isFullRun": false,
  "summary": { "new": 2, "modified": 3, "deleted": 1, "unchanged": 94 },
  "changes": [
    { "changeType": "new", "entry": { "...full entry object..." } },
    { "changeType": "modified", "entry": { "...full entry object..." } },
    {
      "changeType": "deleted",
      "entryId": "00000000...",
      "lastKnownSubject": "Cancelled Meeting",
      "lastKnownStart": "2026-03-25T14:00:00"
    }
  ]
}
```

| Field | Description |
|---|---|
| `detectedAt` | UTC timestamp of when change detection ran |
| `previousRunDate` | UTC timestamp of the previous run (null on first run) |
| `isFullRun` | `true` if no previous state existed — all entries are "new" |
| `summary` | Counts of new, modified, deleted, and unchanged entries |
| `changes[].changeType` | One of: `new`, `modified`, `deleted` |
| `changes[].entry` | Full entry object for new/modified (same schema as export entries) |
| `changes[].entryId` | Entry ID for deleted items |
| `changes[].lastKnownSubject` | Last known title for deleted items (may be null) |
| `changes[].lastKnownStart` | Last known start time for deleted items (may be null) |

### Run History CSV

Each run appends a row to `output/run-history.csv` (max 500 rows, oldest trimmed automatically):

| Column | Description |
|---|---|
| `RunDate` | When this run executed |
| `PreviousRunDate` | When the previous run executed (or "N/A") |
| `IsFullRun` | Whether this was a full run |
| `TotalEntries` | Total entries in the current export |
| `New` / `Modified` / `Deleted` / `Unchanged` | Change counts |
| `Duration` | How long the detection took |
| `Status` | "Success" or error status |

### Upload History CSV

Each upload run appends a row to `output/upload-history.csv` (max 500 rows, oldest trimmed automatically):

| Column | Description |
|---|---|
| `RunDate` | When this upload executed |
| `Endpoint` | API endpoint URL used |
| `IsFullRun` | Whether the changes came from a full run |
| `New` / `Modified` / `Deleted` | Change counts from the input file |
| `HttpStatus` | HTTP response status code (or "N/A" if skipped) |
| `UploadId` | Server-assigned upload ID |
| `Processed` / `Created` / `Updated` / `ServerDeleted` | Server-side processing counts |
| `Duration` | How long the upload took |
| `Status` | "Success", "Failed", or "Skipped-NoChanges" |
| `ErrorMessage` | Error details if failed (empty otherwise) |

### Persistence: last-run.json

The `output/last-run.json` file stores the state needed for change detection:
- `lastRunDate` — when the last successful run occurred
- `entryIndex` — map of entryId to lastModified timestamp
- `entrySummary` — map of entryId to subject/start (used for deleted entry context)

**Delete this file to force a full run.** You can also edit `lastRunDate` to replay changes from a specific point in time.

## Scheduling as a Windows Task

To run the export and change detection automatically:

1. Open **Task Scheduler** (`taskschd.msc`)
2. Create a new task with these settings:
   - **Action**: Start a program
   - **Program**: `powershell.exe`
   - **Arguments**: `-ExecutionPolicy Bypass -Command "& '.\Export-OutlookCalendar.ps1'; & '.\Detect-CalendarChanges.ps1'; & '.\Upload-CalendarChanges.ps1'"`
   - **Start in**: `C:\path\to\scripts\outlook-calendar-export`
3. Set your desired trigger (e.g., every morning at 7:00 AM)
4. Under **Conditions**, uncheck "Start only if the computer is on AC power" if on a laptop

5. For the upload script, set the `SINK_TOKEN` environment variable in the task's **Environment** section, or use `-Token` in the arguments
6. Ensure the machine has network access to the Sink API endpoint

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
| Upload: "No API token provided" | Set `SINK_TOKEN` environment variable or pass `-Token` parameter. See the [CALENDAR-SYNC.md](https://github.com/marinoscar/sink/blob/main/docs/CALENDAR-SYNC.md) guide for creating a Personal Access Token |
| Upload: HTTP 401 Unauthorized | Your token has expired or is invalid. Generate a new PAT via the Sink API |
| Upload: HTTP 413 or timeout | The changes file may be too large. Try running more frequently to reduce the number of changes per upload |
| Upload: Network/connection error | Verify network access to the Sink API endpoint. Check firewall, proxy, and DNS settings |
