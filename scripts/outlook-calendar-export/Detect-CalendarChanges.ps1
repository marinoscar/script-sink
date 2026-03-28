<#
.SYNOPSIS
    Detects changes in the Outlook calendar export by comparing against the previous run.

.DESCRIPTION
    Reads the full calendar export JSON (produced by Export-OutlookCalendar.ps1) and compares
    it against a stored index from the previous run. Outputs a changes-only JSON file containing
    only new, modified, and deleted entries — suitable for incremental sync to Google Calendar.

    Change detection uses the entryId as the sync key and lastModified as the change indicator:
    - New:       entryId exists in current export but not in the previous index
    - Modified:  entryId exists in both but lastModified timestamp differs
    - Deleted:   entryId exists in the previous index but not in the current export

    On first run (or if last-run.json is deleted), all entries are reported as "new" and
    isFullRun is set to true. Delete last-run.json at any time to force a full run.

    Each run is recorded in a CSV history file (max 500 rows) for diagnostics in Excel.

.PARAMETER ExportPath
    Path to the calendar export JSON. Overrides config.json value.
    Default: ./output/calendar-export.json

.PARAMETER ChangesOutputPath
    Path for the changes-only JSON output. Overrides config.json value.
    Default: ./output/calendar-changes.json

.PARAMETER LastRunPath
    Path to the last-run state file. Overrides config.json value.
    Default: ./output/last-run.json

.PARAMETER RunHistoryPath
    Path to the run history CSV. Overrides config.json value.
    Default: ./output/run-history.csv

.PARAMETER LogPath
    Directory for log files. Overrides config.json value.
    Default: ./logs/

.PARAMETER ConfigPath
    Path to the configuration file. Default: ./config.json

.EXAMPLE
    .\Detect-CalendarChanges.ps1
    Runs with defaults (or config.json values if present).

.EXAMPLE
    .\Detect-CalendarChanges.ps1 -ExportPath "C:\Exports\calendar-export.json"
    Uses a custom export path.

.NOTES
    Prerequisites:
    - Run Export-OutlookCalendar.ps1 first to generate the calendar export JSON
    - PowerShell 5.1+ (ships with Windows)
    - No admin rights required
    - No additional modules required

    To force a full run, delete the last-run.json file:
    Remove-Item .\output\last-run.json
#>

param(
    [string]$ExportPath,
    [string]$ChangesOutputPath,
    [string]$LastRunPath,
    [string]$RunHistoryPath,
    [string]$LogPath,
    [string]$ConfigPath
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$startTime = Get-Date

# ==================================================
# Built-in defaults
# ==================================================
$defaults = @{
    ExportPath        = Join-Path $scriptDir "output\calendar-export.json"
    ChangesOutputPath = Join-Path $scriptDir "output\calendar-changes.json"
    LastRunPath       = Join-Path $scriptDir "output\last-run.json"
    RunHistoryPath    = Join-Path $scriptDir "output\run-history.csv"
    RunHistoryMaxRows = 500
    LogPath           = Join-Path $scriptDir "logs"
}

# ==================================================
# Logging
# ==================================================
$script:logFile = $null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"
    $color = switch ($Level) {
        "Info"    { "White" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
    }
    Write-Host $logLine -ForegroundColor $color
    if ($script:logFile) {
        Add-Content -Path $script:logFile -Value $logLine -Encoding UTF8
    }
}

# ==================================================
# Banner
# ==================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Calendar Change Detection" -ForegroundColor Cyan
Write-Host "  Started: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# ==================================================
# Load configuration
# ==================================================
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $scriptDir "config.json"
}

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading configuration from: $ConfigPath"
    try {
        $configRaw = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        $configRaw.PSObject.Properties | ForEach-Object {
            $config[$_.Name] = $_.Value
        }
        Write-Log "Configuration loaded successfully." -Level Success
    } catch {
        Write-Log "WARNING: Failed to parse config file: $($_.Exception.Message)" -Level Warning
        Write-Log "Falling back to default values." -Level Warning
        $config = @{}
    }
} else {
    Write-Log "Config file not found at: $ConfigPath" -Level Warning
    Write-Log "Using default values." -Level Warning
}

# Helper to resolve a path setting: CLI param > config key > default
function Resolve-Setting {
    param(
        [string]$ParamName,
        [string]$ConfigKey,
        [string]$DefaultValue
    )
    # Check CLI parameter
    if ($PSBoundParameters.ContainsKey($ParamName) -and $PSBoundParameters[$ParamName]) {
        $val = $PSBoundParameters[$ParamName]
        Write-Log "$ParamName overridden by CLI parameter: $val"
        return $val
    }
    # Check config
    if ($config.ContainsKey($ConfigKey) -and $config[$ConfigKey]) {
        $val = $config[$ConfigKey]
        if (-not [System.IO.Path]::IsPathRooted($val)) {
            $val = Join-Path $scriptDir $val
        }
        return $val
    }
    return $DefaultValue
}

# Resolve all settings. Note: we pass PSBoundParameters from the script scope explicitly.
$scriptBound = $PSBoundParameters

$finalExportPath = if ($scriptBound.ContainsKey('ExportPath') -and $scriptBound['ExportPath']) {
    Write-Log "ExportPath overridden by CLI parameter: $($scriptBound['ExportPath'])"
    $scriptBound['ExportPath']
} elseif ($config.ContainsKey('outputPath') -and $config['outputPath']) {
    $p = $config['outputPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.ExportPath }

$finalChangesOutputPath = if ($scriptBound.ContainsKey('ChangesOutputPath') -and $scriptBound['ChangesOutputPath']) {
    Write-Log "ChangesOutputPath overridden by CLI parameter: $($scriptBound['ChangesOutputPath'])"
    $scriptBound['ChangesOutputPath']
} elseif ($config.ContainsKey('changesOutputPath') -and $config['changesOutputPath']) {
    $p = $config['changesOutputPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.ChangesOutputPath }

$finalLastRunPath = if ($scriptBound.ContainsKey('LastRunPath') -and $scriptBound['LastRunPath']) {
    Write-Log "LastRunPath overridden by CLI parameter: $($scriptBound['LastRunPath'])"
    $scriptBound['LastRunPath']
} elseif ($config.ContainsKey('lastRunPath') -and $config['lastRunPath']) {
    $p = $config['lastRunPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.LastRunPath }

$finalRunHistoryPath = if ($scriptBound.ContainsKey('RunHistoryPath') -and $scriptBound['RunHistoryPath']) {
    Write-Log "RunHistoryPath overridden by CLI parameter: $($scriptBound['RunHistoryPath'])"
    $scriptBound['RunHistoryPath']
} elseif ($config.ContainsKey('runHistoryPath') -and $config['runHistoryPath']) {
    $p = $config['runHistoryPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.RunHistoryPath }

$finalRunHistoryMaxRows = if ($config.ContainsKey('runHistoryMaxRows')) {
    [int]$config['runHistoryMaxRows']
} else { $defaults.RunHistoryMaxRows }

$finalLogPath = if ($scriptBound.ContainsKey('LogPath') -and $scriptBound['LogPath']) {
    Write-Log "LogPath overridden by CLI parameter: $($scriptBound['LogPath'])"
    $scriptBound['LogPath']
} elseif ($config.ContainsKey('logPath') -and $config['logPath']) {
    $p = $config['logPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.LogPath }

# ==================================================
# Initialize log file
# ==================================================
if (-not (Test-Path $finalLogPath)) {
    New-Item -ItemType Directory -Path $finalLogPath -Force | Out-Null
}
$logFileName = "changes-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$script:logFile = Join-Path $finalLogPath $logFileName
Write-Log "Log file initialized: $($script:logFile)"

# Ensure output directories exist
foreach ($path in @($finalChangesOutputPath, $finalLastRunPath, $finalRunHistoryPath)) {
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Log "Created directory: $dir"
    }
}

# Log resolved configuration
Write-Log "--- Resolved Configuration ---"
Write-Log "  Export Path:        $finalExportPath"
Write-Log "  Changes Output:     $finalChangesOutputPath"
Write-Log "  Last Run State:     $finalLastRunPath"
Write-Log "  Run History CSV:    $finalRunHistoryPath"
Write-Log "  Run History Max:    $finalRunHistoryMaxRows rows"
Write-Log "  Log Path:           $finalLogPath"
Write-Log "-------------------------------"

# ==================================================
# Step 1: Load the calendar export
# ==================================================
Write-Log "Loading calendar export from: $finalExportPath"

if (-not (Test-Path $finalExportPath)) {
    Write-Log "FATAL: Calendar export file not found: $finalExportPath" -Level Error
    Write-Log "Run Export-OutlookCalendar.ps1 first to generate the export." -Level Error
    exit 1
}

try {
    $exportRaw = Get-Content -Path $finalExportPath -Raw | ConvertFrom-Json
    $currentEntries = $exportRaw.entries
    Write-Log "Loaded $($currentEntries.Count) entries from export (exported: $($exportRaw.exportDate))." -Level Success
} catch {
    Write-Log "FATAL: Failed to parse export file: $($_.Exception.Message)" -Level Error
    exit 1
}

# Build a lookup hashtable: entryId -> entry object
$currentIndex = @{}
foreach ($entry in $currentEntries) {
    $currentIndex[$entry.entryId] = $entry
}
Write-Log "Built current entry index with $($currentIndex.Count) entries."

# ==================================================
# Step 2: Load the last-run state
# ==================================================
$isFullRun = $false
$previousRunDate = $null
$previousIndex = @{}

if (Test-Path $finalLastRunPath) {
    Write-Log "Loading last-run state from: $finalLastRunPath"
    try {
        $lastRunRaw = Get-Content -Path $finalLastRunPath -Raw | ConvertFrom-Json
        $previousRunDate = $lastRunRaw.lastRunDate
        Write-Log "Previous run date: $previousRunDate"

        # Build previous index: entryId -> lastModified
        $lastRunRaw.entryIndex.PSObject.Properties | ForEach-Object {
            $previousIndex[$_.Name] = $_.Value
        }
        Write-Log "Previous index contains $($previousIndex.Count) entries." -Level Success
    } catch {
        Write-Log "WARNING: Failed to parse last-run state: $($_.Exception.Message)" -Level Warning
        Write-Log "Treating this as a full run." -Level Warning
        $isFullRun = $true
        $previousIndex = @{}
    }
} else {
    Write-Log "No last-run state found at: $finalLastRunPath" -Level Warning
    Write-Log "This is a full run — all entries will be reported as new." -Level Warning
    $isFullRun = $true
}

# ==================================================
# Step 3: Compare entries and detect changes
# ==================================================
Write-Log "Comparing current entries against previous index..."

$changes = @()
$newCount = 0
$modifiedCount = 0
$deletedCount = 0
$unchangedCount = 0

# Check each current entry against the previous index
foreach ($entry in $currentEntries) {
    $id = $entry.entryId
    $currentLastMod = $entry.lastModified

    if (-not $previousIndex.ContainsKey($id)) {
        # New entry
        $newCount++
        $changes += [ordered]@{
            changeType = "new"
            entry      = $entry
        }
        Write-Log "  [NEW] $($entry.subject) ($($entry.start))"
    } elseif ($previousIndex[$id] -ne $currentLastMod) {
        # Modified entry — lastModified timestamp differs
        $modifiedCount++
        $changes += [ordered]@{
            changeType = "modified"
            entry      = $entry
        }
        Write-Log "  [MODIFIED] $($entry.subject) ($($entry.start)) — was: $($previousIndex[$id]), now: $currentLastMod"
    } else {
        # Unchanged
        $unchangedCount++
    }
}

# Check for deleted entries (in previous index but not in current export)
# We also store the previous index's subject/start for deleted items so the
# consumer knows what was removed. We read these from the last-run state's
# entrySummary if available, otherwise just report the entryId.
$previousSummary = @{}
if (-not $isFullRun -and $lastRunRaw.PSObject.Properties.Name -contains 'entrySummary') {
    $lastRunRaw.entrySummary.PSObject.Properties | ForEach-Object {
        $previousSummary[$_.Name] = $_.Value
    }
}

foreach ($prevId in $previousIndex.Keys) {
    if (-not $currentIndex.ContainsKey($prevId)) {
        $deletedCount++
        $deletedEntry = [ordered]@{
            changeType       = "deleted"
            entryId          = $prevId
            lastKnownSubject = $null
            lastKnownStart   = $null
        }
        # Try to fill in subject/start from the stored summary
        if ($previousSummary.ContainsKey($prevId)) {
            $summary = $previousSummary[$prevId]
            $deletedEntry["lastKnownSubject"] = $summary.subject
            $deletedEntry["lastKnownStart"] = $summary.start
            Write-Log "  [DELETED] $($summary.subject) ($($summary.start))"
        } else {
            Write-Log "  [DELETED] entryId=$prevId (no prior summary available)"
        }
        $changes += $deletedEntry
    }
}

Write-Log "--- Change Summary ---" -Level Success
Write-Log "  New:       $newCount"
Write-Log "  Modified:  $modifiedCount"
Write-Log "  Deleted:   $deletedCount"
Write-Log "  Unchanged: $unchangedCount"
Write-Log "  Total:     $($currentEntries.Count) current entries"
Write-Log "----------------------"

# ==================================================
# Step 4: Write changes JSON
# ==================================================
Write-Log "Writing changes to: $finalChangesOutputPath"

$changesOutput = [ordered]@{
    detectedAt      = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    previousRunDate = $previousRunDate
    isFullRun       = $isFullRun
    summary         = [ordered]@{
        new       = $newCount
        modified  = $modifiedCount
        deleted   = $deletedCount
        unchanged = $unchangedCount
    }
    changes         = $changes
}

$changesJson = $changesOutput | ConvertTo-Json -Depth 10 -Compress:$false
$changesJson | Out-File -FilePath $finalChangesOutputPath -Encoding UTF8
$changesSizeKB = [math]::Round((Get-Item $finalChangesOutputPath).Length / 1024, 1)
Write-Log "Changes file written (${changesSizeKB} KB)." -Level Success

# ==================================================
# Step 5: Update last-run state
# ==================================================
Write-Log "Updating last-run state: $finalLastRunPath"

# Build the new entry index (entryId -> lastModified) and a summary (entryId -> {subject, start})
# for use in detecting deleted entries on the next run.
$newEntryIndex = [ordered]@{}
$newEntrySummary = [ordered]@{}
foreach ($entry in $currentEntries) {
    $newEntryIndex[$entry.entryId] = $entry.lastModified
    $newEntrySummary[$entry.entryId] = [ordered]@{
        subject = $entry.subject
        start   = $entry.start
    }
}

$lastRunState = [ordered]@{
    lastRunDate  = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    entryCount   = $currentEntries.Count
    entryIndex   = $newEntryIndex
    entrySummary = $newEntrySummary
}

$lastRunState | ConvertTo-Json -Depth 5 -Compress:$false | Out-File -FilePath $finalLastRunPath -Encoding UTF8
Write-Log "Last-run state saved with $($currentEntries.Count) entries." -Level Success

# ==================================================
# Step 6: Append to run history CSV
# ==================================================
$elapsed = (Get-Date) - $startTime
$runStatus = "Success"

Write-Log "Appending to run history CSV: $finalRunHistoryPath"

$historyRow = [PSCustomObject]@{
    RunDate         = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    PreviousRunDate = if ($previousRunDate) { $previousRunDate } else { "N/A" }
    IsFullRun       = $isFullRun
    TotalEntries    = $currentEntries.Count
    New             = $newCount
    Modified        = $modifiedCount
    Deleted         = $deletedCount
    Unchanged       = $unchangedCount
    Duration        = $elapsed.ToString('hh\:mm\:ss\.ff')
    Status          = $runStatus
}

# If CSV doesn't exist, create it with header
if (-not (Test-Path $finalRunHistoryPath)) {
    $historyRow | Export-Csv -Path $finalRunHistoryPath -NoTypeInformation -Encoding UTF8
    Write-Log "Run history CSV created (1 of ${finalRunHistoryMaxRows} max rows)." -Level Success
} else {
    # Append the new row
    $historyRow | Export-Csv -Path $finalRunHistoryPath -NoTypeInformation -Encoding UTF8 -Append

    # Enforce max row limit: read all rows, keep only the newest N
    $allRows = Import-Csv -Path $finalRunHistoryPath
    $rowCount = $allRows.Count
    if ($rowCount -gt $finalRunHistoryMaxRows) {
        $trimmed = $allRows | Select-Object -Last $finalRunHistoryMaxRows
        $trimmed | Export-Csv -Path $finalRunHistoryPath -NoTypeInformation -Encoding UTF8
        Write-Log "Run history trimmed from $rowCount to $finalRunHistoryMaxRows rows (oldest rows removed)." -Level Warning
    }
    $displayRows = [math]::Min($rowCount, $finalRunHistoryMaxRows)
    Write-Log "Run history CSV updated (${displayRows} of ${finalRunHistoryMaxRows} max rows)." -Level Success
}

# ==================================================
# Summary
# ==================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Change Detection Complete" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "Full run:        $isFullRun"
Write-Log "Changes found:   $($newCount + $modifiedCount + $deletedCount) ($newCount new, $modifiedCount modified, $deletedCount deleted)"
Write-Log "Unchanged:       $unchangedCount"
Write-Log "Changes file:    $finalChangesOutputPath"
Write-Log "Last-run state:  $finalLastRunPath"
Write-Log "Run history:     $finalRunHistoryPath"
Write-Log "Log file:        $($script:logFile)"
Write-Log "Elapsed time:    $($elapsed.ToString('hh\:mm\:ss\.ff'))"
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""