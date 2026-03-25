<#
.SYNOPSIS
    Exports Outlook Desktop calendar entries to a JSON file for syncing to another calendar.

.DESCRIPTION
    Connects to the Outlook Desktop application via COM automation and extracts calendar
    entries within a configurable date range (default: 15 days back, 90 days forward).

    The exported JSON includes scheduling metadata needed for calendar sync:
    - Unique entry ID and last-modified timestamp for sync integrity
    - Title, start/end times, location, organizer
    - Busy status (Free/Tentative/Busy/OutOfOffice/WorkingElsewhere)
    - Response status (Accepted/Tentative/Declined/etc.)
    - Recurrence pattern details for recurring events

    Sensitive information is explicitly excluded: no body text, no attendee lists,
    no attachments, no HTML/RTF content.

    Configuration can be provided via config.json (in the script directory) or CLI parameters.
    CLI parameters override config file values, which override built-in defaults.

.PARAMETER DaysBack
    Number of days in the past to include. Overrides config.json value. Default: 15

.PARAMETER DaysForward
    Number of days in the future to include. Overrides config.json value. Default: 90

.PARAMETER OutputPath
    Path for the JSON output file. Overrides config.json value. Default: ./output/calendar-export.json

.PARAMETER LogPath
    Directory for log files. Overrides config.json value. Default: ./logs/

.PARAMETER ConfigPath
    Path to the configuration file. Default: ./config.json

.EXAMPLE
    .\Export-OutlookCalendar.ps1
    Runs with defaults (or config.json values if present).

.EXAMPLE
    .\Export-OutlookCalendar.ps1 -DaysBack 30 -DaysForward 120
    Exports 30 days of history and 120 days into the future.

.EXAMPLE
    .\Export-OutlookCalendar.ps1 -OutputPath "C:\Exports\calendar.json"
    Exports to a custom output path.

.NOTES
    Prerequisites:
    - Windows 10/11
    - Microsoft Outlook Desktop installed and running with a configured account
    - PowerShell 5.1+ (ships with Windows)
    - No admin rights required
    - No additional modules required

    If you get an execution policy error, run:
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
#>

param(
    [int]$DaysBack,
    [int]$DaysForward,
    [string]$OutputPath,
    [string]$LogPath,
    [string]$ConfigPath,
    [string]$CompanyDomain
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$startTime = Get-Date

# ==================================================
# Built-in defaults
# ==================================================
$defaults = @{
    DaysBack      = 15
    DaysForward   = 90
    OutputPath    = Join-Path $scriptDir "output\calendar-export.json"
    LogPath       = Join-Path $scriptDir "logs"
    CompanyDomain = "mycompany.com"
}

# ==================================================
# Logging setup (bootstrap with defaults until config is resolved)
# ==================================================
$script:logFile = $null

function Write-Log {
    <#
    .SYNOPSIS
        Writes a timestamped message to both console and log file.
    .PARAMETER Message
        The message to log.
    .PARAMETER Level
        Log level: Info, Warning, Error, Success. Controls console color.
    #>
    param(
        [string]$Message,
        [ValidateSet("Info", "Warning", "Error", "Success")]
        [string]$Level = "Info"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logLine = "[$timestamp] [$Level] $Message"

    # Console output with color coding
    $color = switch ($Level) {
        "Info"    { "White" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
    }
    Write-Host $logLine -ForegroundColor $color

    # Append to log file if available
    if ($script:logFile) {
        Add-Content -Path $script:logFile -Value $logLine -Encoding UTF8
    }
}

# ==================================================
# Load configuration
# ==================================================
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Outlook Calendar Export" -ForegroundColor Cyan
Write-Host "  Started: $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Resolve config file path
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $scriptDir "config.json"
}

$config = @{}
if (Test-Path $ConfigPath) {
    Write-Log "Loading configuration from: $ConfigPath"
    try {
        $configRaw = Get-Content -Path $ConfigPath -Raw | ConvertFrom-Json
        # Convert PSObject properties to hashtable
        $configRaw.PSObject.Properties | ForEach-Object {
            $config[$_.Name] = $_.Value
        }
        Write-Log "  daysBack    = $($config['daysBack'])" -Level Info
        Write-Log "  daysForward = $($config['daysForward'])" -Level Info
        Write-Log "  outputPath  = $($config['outputPath'])" -Level Info
        Write-Log "  logPath     = $($config['logPath'])" -Level Info
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

# Resolve final settings: CLI params > config.json > defaults
# DaysBack
$finalDaysBack = if ($PSBoundParameters.ContainsKey('DaysBack')) {
    Write-Log "DaysBack overridden by CLI parameter: $DaysBack"
    $DaysBack
} elseif ($config.ContainsKey('daysBack')) {
    $config['daysBack']
} else {
    $defaults.DaysBack
}

# DaysForward
$finalDaysForward = if ($PSBoundParameters.ContainsKey('DaysForward')) {
    Write-Log "DaysForward overridden by CLI parameter: $DaysForward"
    $DaysForward
} elseif ($config.ContainsKey('daysForward')) {
    $config['daysForward']
} else {
    $defaults.DaysForward
}

# OutputPath
$finalOutputPath = if ($PSBoundParameters.ContainsKey('OutputPath')) {
    Write-Log "OutputPath overridden by CLI parameter: $OutputPath"
    $OutputPath
} elseif ($config.ContainsKey('outputPath')) {
    # Resolve relative paths against script directory
    $p = $config['outputPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else {
    $defaults.OutputPath
}

# LogPath
$finalLogPath = if ($PSBoundParameters.ContainsKey('LogPath')) {
    Write-Log "LogPath overridden by CLI parameter: $LogPath"
    $LogPath
} elseif ($config.ContainsKey('logPath')) {
    $p = $config['logPath']
    if (-not [System.IO.Path]::IsPathRooted($p)) { Join-Path $scriptDir $p } else { $p }
} else {
    $defaults.LogPath
}

# CompanyDomain
$finalCompanyDomain = if ($PSBoundParameters.ContainsKey('CompanyDomain')) {
    Write-Log "CompanyDomain overridden by CLI parameter: $CompanyDomain"
    $CompanyDomain
} elseif ($config.ContainsKey('companyDomain')) {
    $config['companyDomain']
} else {
    $defaults.CompanyDomain
}

# ==================================================
# Initialize log file
# ==================================================
if (-not (Test-Path $finalLogPath)) {
    New-Item -ItemType Directory -Path $finalLogPath -Force | Out-Null
}
$logFileName = "export-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$script:logFile = Join-Path $finalLogPath $logFileName
Write-Log "Log file initialized: $($script:logFile)"

# Ensure output directory exists
$outputDir = Split-Path -Parent $finalOutputPath
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Log "Created output directory: $outputDir"
}

# Log resolved configuration
Write-Log "--- Resolved Configuration ---"
Write-Log "  Days Back:    $finalDaysBack"
Write-Log "  Days Forward: $finalDaysForward"
Write-Log "  Output Path:  $finalOutputPath"
Write-Log "  Log Path:     $finalLogPath"
Write-Log "  Company Dom:  $finalCompanyDomain"
Write-Log "-------------------------------"

# ==================================================
# Lookup tables for Outlook enum values
# ==================================================
$busyStatusMap = @{
    0 = "Free"
    1 = "Tentative"
    2 = "Busy"
    3 = "OutOfOffice"
    4 = "WorkingElsewhere"
}

$responseStatusMap = @{
    0 = "None"
    1 = "Organized"
    2 = "Tentative"
    3 = "Accepted"
    4 = "Declined"
    5 = "NotResponded"
}

$recurrenceTypeMap = @{
    0 = "Daily"
    1 = "Weekly"
    2 = "Monthly"
    3 = "MonthlyNth"
    5 = "Yearly"
    6 = "YearlyNth"
}

function Resolve-SmtpAddress {
    <#
    .SYNOPSIS
        Resolves an SMTP email address from an Outlook AddressEntry or display name.
        Handles EX-type (Exchange internal X500) addresses where GetExchangeUser()
        may fail with E_ABORT and PropertyAccessor may be null.

        Resolution order:
        1. ExchangeUser → PrimarySmtpAddress
        2. Direct Address if SMTP type or contains @
        3. PR_SMTP_ADDRESS via PropertyAccessor (0x39FE001F)
        4. Namespace.CreateRecipient() re-resolve by display name
        5. Namespace.CreateRecipient() re-resolve by X500 address
    #>
    param(
        [object]$AddressEntry,
        [string]$DisplayName,
        [object]$Namespace
    )

    $ErrorActionPreference = "Continue"

    # Try 1: ExchangeUser → PrimarySmtpAddress
    if ($AddressEntry) {
        try {
            $exchUser = $AddressEntry.GetExchangeUser()
            if ($exchUser -and $exchUser.PrimarySmtpAddress) {
                return $exchUser.PrimarySmtpAddress
            }
        } catch {}

        # Try 2: Direct address if SMTP type or contains @
        try {
            if ($AddressEntry.Address) {
                if ($AddressEntry.Type -eq "SMTP") {
                    return $AddressEntry.Address
                } elseif ($AddressEntry.Address -match "@") {
                    return $AddressEntry.Address
                }
            }
        } catch {}

        # Try 3: PropertyAccessor PR_SMTP_ADDRESS
        try {
            if ($AddressEntry.PropertyAccessor) {
                $prop = $AddressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                if ($prop) { return $prop }
            }
        } catch {}
    }

    # Try 4: Re-resolve by display name via Namespace.CreateRecipient()
    # This forces Outlook to do a fresh address book lookup
    if ($Namespace -and $DisplayName) {
        try {
            $resolvedRecip = $Namespace.CreateRecipient($DisplayName)
            $resolvedRecip.Resolve()
            if ($resolvedRecip.Resolved) {
                $ae = $resolvedRecip.AddressEntry
                if ($ae) {
                    try {
                        $eu = $ae.GetExchangeUser()
                        if ($eu -and $eu.PrimarySmtpAddress) {
                            return $eu.PrimarySmtpAddress
                        }
                    } catch {}
                    try {
                        if ($ae.PropertyAccessor) {
                            $prop = $ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                            if ($prop) { return $prop }
                        }
                    } catch {}
                    # Last resort: if the resolved entry is SMTP type
                    try {
                        if ($ae.Type -eq "SMTP" -and $ae.Address) {
                            return $ae.Address
                        }
                    } catch {}
                }
            }
        } catch {}
    }

    # Try 5: Re-resolve by the X500 address itself via CreateRecipient()
    if ($Namespace -and $AddressEntry -and $AddressEntry.Type -eq "EX") {
        try {
            $x500Addr = $AddressEntry.Address
            if ($x500Addr) {
                $resolvedRecip = $Namespace.CreateRecipient($x500Addr)
                $resolvedRecip.Resolve()
                if ($resolvedRecip.Resolved) {
                    $ae = $resolvedRecip.AddressEntry
                    if ($ae) {
                        try {
                            $eu = $ae.GetExchangeUser()
                            if ($eu -and $eu.PrimarySmtpAddress) {
                                return $eu.PrimarySmtpAddress
                            }
                        } catch {}
                        try {
                            if ($ae.PropertyAccessor) {
                                $prop = $ae.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F")
                                if ($prop) { return $prop }
                            }
                        } catch {}
                    }
                }
            }
        } catch {}
    }

    return $null
}

function Get-OrganizerDomain {
    <#
    .SYNOPSIS
        Extracts only the email domain of the meeting organizer (e.g., "google.com").
        No names or email usernames are returned — only the domain portion,
        to avoid persisting sensitive/personal information.

        Uses Resolve-SmtpAddress for address resolution, then falls back to
        MAPI properties, Organizer/SenderEmailAddress strings, and Recipients
        collection lookup.
    #>
    param($Item, $Namespace, [string]$CompanyDomain)

    $ErrorActionPreference = "Continue"

    $smtpAddress = $null
    $diagInfo = @{}

    # Try 1: GetOrganizer() → Resolve-SmtpAddress
    try {
        $addressEntry = $Item.GetOrganizer()
        if ($addressEntry) {
            $diagInfo["GetOrganizer.Type"] = $addressEntry.Type
            $diagInfo["GetOrganizer.Address"] = $addressEntry.Address
            $diagInfo["GetOrganizer.Name"] = $addressEntry.Name

            $smtpAddress = Resolve-SmtpAddress -AddressEntry $addressEntry -DisplayName $addressEntry.Name -Namespace $Namespace
            if ($smtpAddress) { $diagInfo["ResolvedVia"] = "GetOrganizer.Resolve" }
        } else {
            $diagInfo["GetOrganizer"] = "returned null"
        }
    } catch {
        $diagInfo["GetOrganizer.Error"] = $_.Exception.Message
    }

    # Try 2: MAPI PR_SENT_REPRESENTING_SMTP_ADDRESS on the item itself
    if (-not $smtpAddress) {
        try {
            $prop = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
            if ($prop) {
                $smtpAddress = $prop
                $diagInfo["ResolvedVia"] = "MAPI.SentRepresentingSMTP"
            }
        } catch {
            $diagInfo["MAPI.0x5D01.Error"] = $_.Exception.Message
        }
    }

    # Try 3: MAPI PR_SENT_REPRESENTING_EMAIL_ADDRESS (may be X500 or SMTP)
    if (-not $smtpAddress) {
        try {
            $addr = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F")
            $diagInfo["MAPI.SentRepresentingEmail"] = $addr
            if ($addr -and $addr -match "@") {
                $smtpAddress = $addr
                $diagInfo["ResolvedVia"] = "MAPI.SentRepresentingEmail"
            }
        } catch {
            $diagInfo["MAPI.0x0065.Error"] = $_.Exception.Message
        }
    }

    # Try 4: Organizer string property (may contain email)
    if (-not $smtpAddress) {
        try {
            $organizer = $Item.Organizer
            $diagInfo["Item.Organizer"] = $organizer
            if ($organizer -and $organizer -match "@") {
                $smtpAddress = $organizer
                $diagInfo["ResolvedVia"] = "Item.Organizer"
            }
        } catch {
            $diagInfo["Item.Organizer.Error"] = $_.Exception.Message
        }
    } else {
        try { $diagInfo["Item.Organizer"] = $Item.Organizer } catch {}
    }

    # Try 5: SenderEmailAddress property
    if (-not $smtpAddress) {
        try {
            $senderEmail = $Item.SenderEmailAddress
            $diagInfo["Item.SenderEmailAddress"] = $senderEmail
            if ($senderEmail -and $senderEmail -match "@") {
                $smtpAddress = $senderEmail
                $diagInfo["ResolvedVia"] = "Item.SenderEmailAddress"
            }
        } catch {
            $diagInfo["SenderEmail.Error"] = $_.Exception.Message
        }
    }

    # Try 6: Resolve organizer via Namespace.CreateRecipient() using display name
    # (handles EX-type addresses where GetOrganizer's AddressEntry is stale)
    if (-not $smtpAddress -and $Namespace) {
        $displayName = $diagInfo["GetOrganizer.Name"]
        if (-not $displayName) { try { $displayName = $Item.Organizer } catch {} }
        if ($displayName) {
            $smtpAddress = Resolve-SmtpAddress -AddressEntry $null -DisplayName $displayName -Namespace $Namespace
            if ($smtpAddress) { $diagInfo["ResolvedVia"] = "Namespace.CreateRecipient(Name)" }
        }
        if (-not $smtpAddress) { $diagInfo["CreateRecipient.Name"] = "failed for '$displayName'" }
    }

    # Try 7: Match organizer in Recipients collection and re-resolve
    if (-not $smtpAddress) {
        try {
            $organizer = $Item.Organizer
            if (-not $organizer) { $organizer = $diagInfo["GetOrganizer.Name"] }
            $recipients = $Item.Recipients
            if ($organizer -and $recipients) {
                for ($i = 1; $i -le $recipients.Count; $i++) {
                    $recip = $recipients.Item($i)
                    if ($recip.Name -eq $organizer) {
                        $smtpAddress = Resolve-SmtpAddress -AddressEntry $recip.AddressEntry -DisplayName $recip.Name -Namespace $Namespace
                        if ($smtpAddress) {
                            $diagInfo["ResolvedVia"] = "Recipients.Resolve"
                            break
                        }
                    }
                }
            }
        } catch {
            $diagInfo["Recipients.Error"] = $_.Exception.Message
        }
    }

    # Extract domain from the SMTP address
    if ($smtpAddress -and $smtpAddress -match "@(.+)$") {
        return $Matches[1].ToLower()
    }

    # If organizer is EX-type (internal Exchange), use configured company domain
    $isInternal = $diagInfo["GetOrganizer.Type"] -eq "EX"
    if ($isInternal -and $CompanyDomain) {
        $organizerName = $diagInfo["GetOrganizer.Name"]
        Write-Log "  -> Organizer '$organizerName' is internal (EX-type), using company domain: $CompanyDomain" -Level Warning
        return $CompanyDomain.ToLower()
    }

    # Log all diagnostic info so we can see what Outlook actually returned
    $diagParts = @()
    foreach ($key in $diagInfo.Keys | Sort-Object) {
        $diagParts += "${key}='$($diagInfo[$key])'"
    }
    $diagString = $diagParts -join ", "
    Write-Log "  -> WARNING: Could not resolve organizer domain. Diagnostics: $diagString" -Level Warning
    return $null
}

function Get-AttendeeDomains {
    <#
    .SYNOPSIS
        Returns attendee count and domains for a calendar item.

        NOTE: Attendee data (Recipients, RequiredAttendees, OptionalAttendees,
        PropertyAccessor) is blocked by Outlook's Object Model Guard security
        policy when controlled by admin group policy. This function returns
        empty results when the data is inaccessible.
    #>
    param($Item)

    # Attendee data is blocked by Outlook Object Model Guard (admin policy).
    # Recipients, RequiredAttendees/OptionalAttendees, and PropertyAccessor
    # all return empty/null. Return empty until policy is changed.
    return @{ count = 0; domains = @() }
}

# DayOfWeekMask is a bitmask: Sun=1, Mon=2, Tue=4, Wed=8, Thu=16, Fri=32, Sat=64
function Convert-DayOfWeekMask {
    <#
    .SYNOPSIS
        Converts the Outlook DayOfWeekMask bitmask into an array of day names.
    #>
    param([int]$Mask)
    $days = @()
    if ($Mask -band 1)  { $days += "Sunday" }
    if ($Mask -band 2)  { $days += "Monday" }
    if ($Mask -band 4)  { $days += "Tuesday" }
    if ($Mask -band 8)  { $days += "Wednesday" }
    if ($Mask -band 16) { $days += "Thursday" }
    if ($Mask -band 32) { $days += "Friday" }
    if ($Mask -band 64) { $days += "Saturday" }
    return $days
}

# ==================================================
# Step 1: Connect to Outlook COM
# ==================================================
Write-Log "Connecting to Outlook COM object..."
try {
    $outlook = New-Object -ComObject Outlook.Application
    Write-Log "Connected to Outlook successfully." -Level Success
} catch {
    Write-Log "FATAL: Could not connect to Outlook Desktop." -Level Error
    Write-Log "Error: $($_.Exception.Message)" -Level Error
    Write-Log "Make sure Outlook Desktop is running and try again." -Level Error
    exit 1
}

# ==================================================
# Step 2: Access MAPI namespace
# ==================================================
Write-Log "Accessing MAPI namespace..."
try {
    $namespace = $outlook.GetNamespace("MAPI")
    Write-Log "MAPI namespace accessed." -Level Success
} catch {
    Write-Log "FATAL: Could not access MAPI namespace." -Level Error
    Write-Log "Error: $($_.Exception.Message)" -Level Error
    exit 1
}

# ==================================================
# Step 3: Open default calendar folder
# ==================================================
Write-Log "Opening default calendar folder..."
try {
    # olFolderCalendar = 9
    $calendarFolder = $namespace.GetDefaultFolder(9)
    Write-Log "Calendar folder opened: '$($calendarFolder.Name)'" -Level Success
} catch {
    Write-Log "FATAL: Could not open default calendar folder." -Level Error
    Write-Log "Error: $($_.Exception.Message)" -Level Error
    exit 1
}

# ==================================================
# Step 4: Calculate date range and set up filter
# ==================================================
$rangeStart = (Get-Date).Date.AddDays(-$finalDaysBack)
$rangeEnd = (Get-Date).Date.AddDays($finalDaysForward + 1)  # +1 to include the last day fully

Write-Log "Calculating date range..."
Write-Log "  Range start: $($rangeStart.ToString('yyyy-MM-dd'))"
Write-Log "  Range end:   $($rangeEnd.ToString('yyyy-MM-dd'))"

$filter = "[Start] >= '$($rangeStart.ToString("MM/dd/yyyy HH:mm"))' AND [Start] < '$($rangeEnd.ToString("MM/dd/yyyy HH:mm"))'"
Write-Log "  Outlook filter: $filter"

# IMPORTANT: The order of operations matters for Outlook COM recurring item expansion.
# You MUST: (1) Sort by [Start], (2) Set IncludeRecurrences = $true, (3) Apply Restrict().
# If you change this order, recurring items will not be expanded into individual occurrences.
Write-Log "Setting up item collection with IncludeRecurrences..."
$items = $calendarFolder.Items
$items.Sort("[Start]")
$items.IncludeRecurrences = $true
$filteredItems = $items.Restrict($filter)
Write-Log "Item filter applied. Beginning enumeration..." -Level Success

# ==================================================
# Step 5: Extract calendar entries
# ==================================================
$entries = @()
$itemCount = 0
$errorCount = 0

$item = $filteredItems.GetFirst()
while ($item -ne $null) {
    $itemCount++

    try {
        # Extract the subject safely (may be null for some items)
        $subject = if ($item.Subject) { $item.Subject } else { "(No Subject)" }

        Write-Log "Processing item [$itemCount]: $subject ($($item.Start.ToString('yyyy-MM-dd HH:mm')))"

        # Map enum values to readable strings
        $busyText = $busyStatusMap[[int]$item.BusyStatus]
        if (-not $busyText) { $busyText = "Unknown($($item.BusyStatus))" }

        $responseText = $responseStatusMap[[int]$item.ResponseStatus]
        if (-not $responseText) { $responseText = "Unknown($($item.ResponseStatus))" }

        Write-Log "  -> Recurring: $($item.IsRecurring), BusyStatus: $busyText, Response: $responseText"

        # Resolve time zones — Outlook exposes StartTimeZone and EndTimeZone (Outlook 2007+).
        # These return a TimeZone object with an ID property (IANA-style Windows timezone ID).
        # Fall back to the system local timezone if the property is unavailable.
        $startTz = $null
        $endTz = $null
        try { $startTz = $item.StartTimeZone.ID } catch {}
        try { $endTz = $item.EndTimeZone.ID } catch {}
        if (-not $startTz) { $startTz = [System.TimeZoneInfo]::Local.Id }
        if (-not $endTz) { $endTz = [System.TimeZoneInfo]::Local.Id }

        # Build the entry object
        $entry = [ordered]@{
            entryId         = $item.EntryID
            lastModified    = $item.LastModificationTime.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
            subject         = $subject
            start           = $item.Start.ToString("yyyy-MM-ddTHH:mm:ss")
            startTimeZone   = $startTz
            end             = $item.End.ToString("yyyy-MM-ddTHH:mm:ss")
            endTimeZone     = $endTz
            location        = if ($item.Location) { $item.Location } else { $null }
            organizerDomain  = Get-OrganizerDomain -Item $item -Namespace $namespace -CompanyDomain $finalCompanyDomain
            attendeeCount    = 0
            attendeeDomains  = @()
            busyStatus       = $busyText
            responseStatus   = $responseText
            isAllDay         = [bool]$item.AllDayEvent
            isRecurring      = [bool]$item.IsRecurring
        }

        # Extract attendee domains (unique, no names or emails — just domains)
        $attendeeInfo = Get-AttendeeDomains -Item $item
        $entry["attendeeCount"] = $attendeeInfo.count
        $entry["attendeeDomains"] = $attendeeInfo.domains
        if ($attendeeInfo.count -gt 0) {
            Write-Log "  -> Attendees: $($attendeeInfo.count) total, domains: $($attendeeInfo.domains -join ', ')"
        }

        # Extract recurrence pattern for recurring items
        if ($item.IsRecurring) {
            try {
                $pattern = $item.GetRecurrencePattern()

                $recTypeText = $recurrenceTypeMap[[int]$pattern.RecurrenceType]
                if (-not $recTypeText) { $recTypeText = "Unknown($($pattern.RecurrenceType))" }

                $daysOfWeek = Convert-DayOfWeekMask -Mask $pattern.DayOfWeekMask

                $patternEnd = $null
                try {
                    # PatternEndDate throws if the recurrence has no end date
                    if ($pattern.NoEndDate -eq $false) {
                        $patternEnd = $pattern.PatternEndDate.ToString("yyyy-MM-dd")
                    }
                } catch {
                    # No end date set — leave as null
                }

                $entry["recurrencePattern"] = [ordered]@{
                    type         = $recTypeText
                    interval     = $pattern.Interval
                    daysOfWeek   = $daysOfWeek
                    dayOfMonth   = $pattern.DayOfMonth
                    monthOfYear  = $pattern.MonthOfYear
                    instance     = $pattern.Instance
                    patternStart = $pattern.PatternStartDate.ToString("yyyy-MM-dd")
                    patternEnd   = $patternEnd
                    occurrences  = $pattern.Occurrences
                }

                $daysText = if ($daysOfWeek.Count -gt 0) { $daysOfWeek -join "," } else { "N/A" }
                Write-Log "  -> Recurrence: $recTypeText, every $($pattern.Interval) interval(s), days: $daysText"
            } catch {
                # Some occurrence items may not cleanly expose the parent recurrence pattern.
                # This is expected for certain exception occurrences (modified single instances).
                Write-Log "  -> WARNING: Could not read recurrence pattern: $($_.Exception.Message)" -Level Warning
                $entry["recurrencePattern"] = $null
            }
        } else {
            $entry["recurrencePattern"] = $null
        }

        $entries += $entry
    } catch {
        $errorCount++
        Write-Log "  ERROR processing item [$itemCount]: $($_.Exception.Message) - skipping" -Level Error
    }

    $item = $filteredItems.GetNext()
}

Write-Log "Enumeration complete: $($entries.Count) items extracted, $errorCount error(s)." -Level Success

# ==================================================
# Step 6: Build and write JSON output
# ==================================================
Write-Log "Building JSON output..."

$output = [ordered]@{
    exportDate = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    rangeStart = $rangeStart.ToString("yyyy-MM-dd")
    rangeEnd   = $rangeEnd.AddDays(-1).ToString("yyyy-MM-dd")  # Adjust back since we added +1 for the filter
    itemCount  = $entries.Count
    entries    = $entries
}

# ConvertTo-Json with sufficient depth to capture nested recurrencePattern
$json = $output | ConvertTo-Json -Depth 5 -Compress:$false

Write-Log "Writing JSON to: $finalOutputPath"
$json | Out-File -FilePath $finalOutputPath -Encoding UTF8
Write-Log "JSON file written successfully ($([math]::Round((Get-Item $finalOutputPath).Length / 1KB, 1)) KB)." -Level Success

# ==================================================
# Summary
# ==================================================
$elapsed = (Get-Date) - $startTime
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Export Complete" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Log "Items exported:  $($entries.Count)"
Write-Log "Errors:          $errorCount"
Write-Log "Output file:     $finalOutputPath"
Write-Log "Log file:        $($script:logFile)"
Write-Log "Elapsed time:    $($elapsed.ToString('hh\:mm\:ss\.ff'))"
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
