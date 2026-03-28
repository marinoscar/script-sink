<#
.SYNOPSIS
    Uploads calendar changes to the Sink API for Google Calendar sync processing.

.DESCRIPTION
    Reads the changes-only JSON file produced by Detect-CalendarChanges.ps1 and POSTs it
    to the Sink calendar upload API endpoint. The Sink service processes the upload by
    upserting new and modified entries and soft-deleting removed ones.

    If the changes file contains zero changes (0 new, 0 modified, 0 deleted), the upload
    is skipped and the run is recorded as "Skipped-NoChanges" in the history CSV.

    Token resolution order:
    1. -Token CLI parameter
    2. SINK_TOKEN environment variable
    3. If neither is available, the script exits with an error

    Configuration is resolved in priority order: CLI parameters > config.json > built-in defaults.

    Each run is recorded in a CSV history file (max 500 rows) for diagnostics in Excel.

.PARAMETER Token
    Bearer token for API authentication. Overrides the SINK_TOKEN environment variable.
    Tokens are never stored in config.json for security reasons.

.PARAMETER ChangesPath
    Path to the changes JSON file (output of Detect-CalendarChanges.ps1).
    Overrides config.json value. Default: ./output/calendar-changes.json

.PARAMETER Endpoint
    API endpoint URL for the calendar upload. Overrides config.json value.
    Default: https://sink.marin.cr/api/calendar/entries/upload

.PARAMETER LogPath
    Directory for log files. Overrides config.json value.
    Default: ./logs/

.PARAMETER ConfigPath
    Path to the configuration file. Default: ./config.json

.EXAMPLE
    .\Upload-CalendarChanges.ps1 -Token "my-secret-token"
    Uploads changes using an explicit token.

.EXAMPLE
    .\Upload-CalendarChanges.ps1
    Uploads changes using the SINK_TOKEN environment variable.

.EXAMPLE
    $env:SINK_TOKEN = "my-secret-token"
    .\Upload-CalendarChanges.ps1 -Endpoint "https://custom-host/api/calendar/entries/upload"
    Uploads to a custom endpoint using the environment variable token.

.NOTES
    Prerequisites:
    - Run Export-OutlookCalendar.ps1 and Detect-CalendarChanges.ps1 first
    - PowerShell 5.1+ (ships with Windows)
    - No admin rights required
    - No additional modules required
    - Network access to the Sink API endpoint
#>

param(
    [string]$Token,
    [string]$ChangesPath,
    [string]$Endpoint,
    [string]$LogPath,
    [string]$ConfigPath
)

$scriptVersion = "1.1.0"
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$startTime = Get-Date
Write-Host "Upload-CalendarChanges v$scriptVersion" -ForegroundColor Cyan

# ==================================================
# Built-in defaults
# ==================================================
$defaults = @{
    Endpoint             = "https://sink.marin.cr/api/calendar/entries/upload"
    ChangesPath          = Join-Path $scriptDir "output\calendar-changes.json"
    UploadHistoryPath    = Join-Path $scriptDir "output\upload-history.csv"
    UploadHistoryMaxRows = 500
    LogPath              = Join-Path $scriptDir "logs"
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
Write-Host "  Calendar Changes Upload" -ForegroundColor Cyan
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

# ==================================================
# Resolve settings: CLI params > config.json > defaults
# ==================================================
$scriptBound = $PSBoundParameters

$finalEndpoint = if ($scriptBound.ContainsKey('Endpoint') -and $scriptBound['Endpoint']) {
    Write-Log "Endpoint overridden by CLI parameter: $($scriptBound['Endpoint'])"
    $scriptBound['Endpoint']
} elseif ($config.ContainsKey('uploadEndpoint') -and $config['uploadEndpoint']) {
    $config['uploadEndpoint']
} else { $defaults.Endpoint }

$finalChangesPath = if ($scriptBound.ContainsKey('ChangesPath') -and $scriptBound['ChangesPath']) {
    Write-Log "ChangesPath overridden by CLI parameter: $($scriptBound['ChangesPath'])"
    $scriptBound['ChangesPath']
} elseif ($config.ContainsKey('changesOutputPath') -and $config['changesOutputPath']) {
    $p = $config['changesOutputPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.ChangesPath }

$finalUploadHistoryPath = if ($config.ContainsKey('uploadHistoryPath') -and $config['uploadHistoryPath']) {
    $p = $config['uploadHistoryPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else { $defaults.UploadHistoryPath }

$finalUploadHistoryMaxRows = if ($config.ContainsKey('uploadHistoryMaxRows')) {
    [int]$config['uploadHistoryMaxRows']
} else { $defaults.UploadHistoryMaxRows }

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
$logFileName = "upload-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$script:logFile = Join-Path $finalLogPath $logFileName
Write-Log "Log file initialized: $($script:logFile)"
Write-Log "Script version: $scriptVersion"

# Ensure output directories exist
foreach ($path in @($finalUploadHistoryPath)) {
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Log "Created directory: $dir"
    }
}

# ==================================================
# Resolve token: CLI param > env var > error
# ==================================================
$finalToken = if ($scriptBound.ContainsKey('Token') -and $scriptBound['Token']) {
    Write-Log "Token provided via CLI parameter."
    $scriptBound['Token']
} elseif ($env:SINK_TOKEN) {
    Write-Log "Token loaded from SINK_TOKEN environment variable."
    $env:SINK_TOKEN
} else {
    $null
}

if (-not $finalToken) {
    Write-Log "FATAL: No API token provided." -Level Error
    Write-Log "Provide a token using the -Token parameter or set the SINK_TOKEN environment variable." -Level Error
    Write-Log "Example: .\Upload-CalendarChanges.ps1 -Token ""your-token-here""" -Level Error
    Write-Log "Example: `$env:SINK_TOKEN = ""your-token-here""; .\Upload-CalendarChanges.ps1" -Level Error
    exit 1
}

# Mask token for logging (show first 4 chars only)
$tokenMask = if ($finalToken.Length -gt 4) { $finalToken.Substring(0, 4) + "..." } else { "***" }

# Log resolved configuration
Write-Log "--- Resolved Configuration ---"
Write-Log "  Endpoint:          $finalEndpoint"
Write-Log "  Changes Path:      $finalChangesPath"
Write-Log "  Upload History:    $finalUploadHistoryPath"
Write-Log "  History Max Rows:  $finalUploadHistoryMaxRows"
Write-Log "  Log Path:          $finalLogPath"
Write-Log "  Token:             $tokenMask"
Write-Log "-------------------------------"

# ==================================================
# Step 1: Load the changes JSON
# ==================================================
Write-Log "Loading changes file from: $finalChangesPath"

if (-not (Test-Path $finalChangesPath)) {
    Write-Log "FATAL: Changes file not found: $finalChangesPath" -Level Error
    Write-Log "Run Detect-CalendarChanges.ps1 first to generate the changes file." -Level Error
    exit 1
}

try {
    $changesRaw = Get-Content -Path $finalChangesPath -Raw
    $changesData = $changesRaw | ConvertFrom-Json
    $newCount = [int]$changesData.summary.new
    $modifiedCount = [int]$changesData.summary.modified
    $deletedCount = [int]$changesData.summary.deleted
    $unchangedCount = [int]$changesData.summary.unchanged
    $isFullRun = [bool]$changesData.isFullRun
    $totalChanges = $newCount + $modifiedCount + $deletedCount

    Write-Log 'Changes file loaded.' -Level Success
    Write-Log "  Full run:    $isFullRun"
    Write-Log "  New:         $newCount"
    Write-Log "  Modified:    $modifiedCount"
    Write-Log "  Deleted:     $deletedCount"
    Write-Log "  Unchanged:   $unchangedCount"
    Write-Log "  Total changes: $totalChanges"
} catch {
    Write-Log "FATAL: Failed to parse changes file: $($_.Exception.Message)" -Level Error
    exit 1
}

# ==================================================
# Transform payload for API format
# ==================================================
Write-Log "Transforming payload to API format (merging new + modified into entries)..."

$entriesList = @()
if ($changesData.new) { $entriesList += @($changesData.new) }
if ($changesData.modified) { $entriesList += @($changesData.modified) }

$apiPayloadObj = [PSCustomObject]@{
    exportDate = $changesData.exportDate
    rangeStart = $changesData.rangeStart
    rangeEnd   = $changesData.rangeEnd
    itemCount  = $entriesList.Count
    entries    = $entriesList
}

$apiPayload = $apiPayloadObj | ConvertTo-Json -Depth 20 -Compress
$payloadSizeKB = [math]::Round($apiPayload.Length / 1KB, 1)

Write-Log ('API payload built: {0} entries ({1} new + {2} modified), {3} KB' -f $entriesList.Count, $newCount, $modifiedCount, $payloadSizeKB) -Level Success

# ==================================================
# Step 2: Check if upload is needed
# ==================================================
$httpStatus = $null
$uploadId = $null
$serverProcessed = $null
$serverCreated = $null
$serverUpdated = $null
$serverDeleted = $null
$errorMessage = $null
$runStatus = $null

if ($totalChanges -eq 0) {
    Write-Log 'No changes to upload (0 new, 0 modified, 0 deleted). Skipping API call.' -Level Warning
    $runStatus = "Skipped-NoChanges"
} else {
    # ==================================================
    # Step 3: Upload to Sink API
    # ==================================================

    # Ensure TLS 1.2 is available (required for HTTPS on PowerShell 5.1)
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {
        Write-Log "WARNING: Could not set TLS 1.2. HTTPS calls may fail on older systems." -Level Warning
    }

    Write-Log "Uploading $totalChanges changes to: $finalEndpoint"
    Write-Log "  Payload size: $payloadSizeKB KB"

    $headers = @{
        "Authorization" = "Bearer $finalToken"
        "Content-Type"  = "application/json"
    }

    try {
        $response = Invoke-WebRequest -Uri $finalEndpoint -Method POST -Headers $headers -Body $apiPayload -UseBasicParsing

        $httpStatus = $response.StatusCode
        Write-Log "HTTP $httpStatus received." -Level Success

        # Parse the response body
        try {
            $responseObj = $response.Content | ConvertFrom-Json
            $uploadId        = $responseObj.uploadId
            $serverProcessed = $responseObj.processed
            $serverCreated   = $responseObj.created
            $serverUpdated   = $responseObj.updated
            $serverDeleted   = $responseObj.deleted

            Write-Log "  Upload ID:       $uploadId"
            Write-Log "  Server processed: $serverProcessed"
            Write-Log "  Server created:   $serverCreated"
            Write-Log "  Server updated:   $serverUpdated"
            Write-Log "  Server deleted:   $serverDeleted"
        } catch {
            Write-Log "WARNING: Could not parse response body: $($_.Exception.Message)" -Level Warning
            Write-Log "Raw response: $($response.Content)" -Level Warning
        }

        $runStatus = "Success"

    } catch {
        $runStatus = "Failed"
        $errorMessage = $_.Exception.Message

        # Try to extract HTTP status and response body from the exception
        if ($_.Exception.Response) {
            $httpStatus = [int]$_.Exception.Response.StatusCode

            # PowerShell 7+ exposes error details directly
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $responseBody = $_.ErrorDetails.Message
            } else {
                # PowerShell 5.1 requires reading the response stream
                try {
                    $stream = $_.Exception.Response.GetResponseStream()
                    $reader = New-Object System.IO.StreamReader($stream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Close()
                    $stream.Close()
                } catch {
                    $responseBody = "(could not read response body)"
                }
            }

            Write-Log "FATAL: API call failed with HTTP $httpStatus" -Level Error
            Write-Log "Response: $responseBody" -Level Error
        } else {
            Write-Log "FATAL: API call failed: $errorMessage" -Level Error
        }
    }
}

# ==================================================
# Step 4: Append to upload history CSV
# ==================================================
$elapsed = (Get-Date) - $startTime

Write-Log "Appending to upload history CSV: $finalUploadHistoryPath"

$historyRow = [PSCustomObject]@{
    RunDate       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Endpoint      = $finalEndpoint
    IsFullRun     = $isFullRun
    New           = $newCount
    Modified      = $modifiedCount
    Deleted       = $deletedCount
    HttpStatus    = if ($httpStatus) { $httpStatus } else { "N/A" }
    UploadId      = if ($uploadId) { $uploadId } else { "N/A" }
    Processed     = if ($serverProcessed -ne $null) { $serverProcessed } else { "N/A" }
    Created       = if ($serverCreated -ne $null) { $serverCreated } else { "N/A" }
    Updated       = if ($serverUpdated -ne $null) { $serverUpdated } else { "N/A" }
    ServerDeleted = if ($serverDeleted -ne $null) { $serverDeleted } else { "N/A" }
    Duration      = $elapsed.ToString('hh\:mm\:ss\.ff')
    Status        = $runStatus
    ErrorMessage  = if ($errorMessage) { $errorMessage } else { "" }
}

if (-not (Test-Path $finalUploadHistoryPath)) {
    $historyRow | Export-Csv -Path $finalUploadHistoryPath -NoTypeInformation -Encoding UTF8
    Write-Log ('Upload history CSV created (1/{0} rows).' -f $finalUploadHistoryMaxRows) -Level Success
} else {
    $historyRow | Export-Csv -Path $finalUploadHistoryPath -NoTypeInformation -Encoding UTF8 -Append

    # Enforce max row limit: read all rows, keep only the newest N
    $allRows = Import-Csv -Path $finalUploadHistoryPath
    $rowCount = $allRows.Count
    if ($rowCount -gt $finalUploadHistoryMaxRows) {
        $trimmed = $allRows | Select-Object -Last $finalUploadHistoryMaxRows
        $trimmed | Export-Csv -Path $finalUploadHistoryPath -NoTypeInformation -Encoding UTF8
        Write-Log ('Upload history trimmed from {0} to {1} rows (oldest rows removed).' -f $rowCount, $finalUploadHistoryMaxRows) -Level Warning
    }
    Write-Log ('Upload history CSV updated ({0}/{1} rows).' -f [math]::Min($rowCount, $finalUploadHistoryMaxRows), $finalUploadHistoryMaxRows) -Level Success
}

# ==================================================
# Summary
# ==================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Calendar Changes Upload Complete" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

if ($runStatus -eq "Success") {
    Write-Log "Status:          $runStatus"
    Write-Log ('Changes sent:    {0} ({1} new, {2} modified, {3} deleted)' -f $totalChanges, $newCount, $modifiedCount, $deletedCount)
    Write-Log "HTTP Status:     $httpStatus"
    Write-Log "Upload ID:       $uploadId"
    Write-Log "Server stats:    $serverProcessed processed, $serverCreated created, $serverUpdated updated, $serverDeleted deleted"
} elseif ($runStatus -eq "Skipped-NoChanges") {
    Write-Log "Status:          Skipped - no changes to upload"
    Write-Log "Unchanged:       $unchangedCount entries"
} else {
    Write-Log "Status:          FAILED" -Level Error
    Write-Log "HTTP Status:     $(if ($httpStatus) { $httpStatus } else { 'N/A' })" -Level Error
    Write-Log "Error:           $errorMessage" -Level Error
}

Write-Log "History CSV:     $finalUploadHistoryPath"
Write-Log "Log file:        $($script:logFile)"
Write-Log "Elapsed time:    $($elapsed.ToString('hh\:mm\:ss\.ff'))"
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
