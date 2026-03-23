<#
.SYNOPSIS
    Diagnostic script to test Outlook Desktop COM access and extract today's calendar entries.

.DESCRIPTION
    This is a lightweight smoke test that verifies:
    1. The Outlook COM object can be created
    2. The MAPI namespace is accessible
    3. The default calendar folder can be opened
    4. Calendar items for today can be read

    Outputs a CSV file with minimal fields: EntryID, Organizer, Start, End, ResponseStatus.
    Use this script to validate connectivity before running the full Export-OutlookCalendar.ps1.

.EXAMPLE
    .\Test-OutlookAccess.ps1

    Extracts today's calendar entries and saves to output\test-calendar.csv

.NOTES
    Prerequisites: Windows 10/11, Outlook Desktop running with a configured account, PowerShell 5.1+
#>

# --------------------------------------------------
# Setup
# --------------------------------------------------
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$outputDir = Join-Path $scriptDir "output"
$outputFile = Join-Path $outputDir "test-calendar.csv"

# Ensure output directory exists
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Outlook Calendar Access - Smoke Test" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# --------------------------------------------------
# Step 1: Connect to Outlook COM
# --------------------------------------------------
Write-Host "[1/5] Connecting to Outlook COM object..." -ForegroundColor White
try {
    $outlook = New-Object -ComObject Outlook.Application
    Write-Host "       Connected successfully." -ForegroundColor Green
} catch {
    Write-Host "       FAILED: Could not connect to Outlook." -ForegroundColor Red
    Write-Host "       Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "       Make sure Outlook Desktop is running and try again." -ForegroundColor Yellow
    exit 1
}

# --------------------------------------------------
# Step 2: Access MAPI namespace
# --------------------------------------------------
Write-Host "[2/5] Accessing MAPI namespace..." -ForegroundColor White
try {
    $namespace = $outlook.GetNamespace("MAPI")
    Write-Host "       MAPI namespace accessed." -ForegroundColor Green
} catch {
    Write-Host "       FAILED: Could not access MAPI namespace." -ForegroundColor Red
    Write-Host "       Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --------------------------------------------------
# Step 3: Open default calendar folder
# --------------------------------------------------
Write-Host "[3/5] Opening default calendar folder..." -ForegroundColor White
try {
    # olFolderCalendar = 9
    $calendarFolder = $namespace.GetDefaultFolder(9)
    Write-Host "       Calendar folder opened: $($calendarFolder.Name)" -ForegroundColor Green
} catch {
    Write-Host "       FAILED: Could not open calendar folder." -ForegroundColor Red
    Write-Host "       Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --------------------------------------------------
# Step 4: Query today's entries
# --------------------------------------------------
Write-Host "[4/5] Querying today's calendar entries..." -ForegroundColor White

$todayStart = (Get-Date).Date
$todayEnd = $todayStart.AddDays(1)

$filter = "[Start] >= '$($todayStart.ToString("MM/dd/yyyy HH:mm"))' AND [Start] < '$($todayEnd.ToString("MM/dd/yyyy HH:mm"))'"
Write-Host "       Date range: $($todayStart.ToString('yyyy-MM-dd')) to $($todayEnd.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
Write-Host "       Filter: $filter" -ForegroundColor Gray

try {
    $items = $calendarFolder.Items
    # IMPORTANT: Sort by Start BEFORE setting IncludeRecurrences, then apply filter
    $items.Sort("[Start]")
    $items.IncludeRecurrences = $true
    $filteredItems = $items.Restrict($filter)

    # Collect results into an array
    $results = @()
    $responseStatusMap = @{
        0 = "None"
        1 = "Organized"
        2 = "Tentative"
        3 = "Accepted"
        4 = "Declined"
        5 = "NotResponded"
    }

    # Extract only the organizer's email domain (e.g., "google.com") to avoid
    # persisting sensitive personal information like names or full email addresses.
    function Get-OrganizerDomain {
        param($Item)
        $smtpAddress = $null
        # Try 1: GetOrganizer() → ExchangeUser → PrimarySmtpAddress
        try {
            $addressEntry = $Item.GetOrganizer()
            if ($addressEntry) {
                try {
                    $exchUser = $addressEntry.GetExchangeUser()
                    if ($exchUser -and $exchUser.PrimarySmtpAddress) { $smtpAddress = $exchUser.PrimarySmtpAddress }
                } catch {}
                if (-not $smtpAddress -and $addressEntry.Type -eq "SMTP") { $smtpAddress = $addressEntry.Address }
                if (-not $smtpAddress) {
                    try { $smtpAddress = $addressEntry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F") } catch {}
                }
            }
        } catch {}
        # Try 2: MAPI PR_SENT_REPRESENTING_SMTP_ADDRESS
        if (-not $smtpAddress) {
            try { $smtpAddress = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001F") } catch {}
        }
        # Try 3: MAPI PR_SENT_REPRESENTING_EMAIL_ADDRESS (only if it looks like SMTP)
        if (-not $smtpAddress) {
            try {
                $addr = $Item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x0065001F")
                if ($addr -and $addr -match "@") { $smtpAddress = $addr }
            } catch {}
        }
        if ($smtpAddress -and $smtpAddress -match "@(.+)$") { return $Matches[1].ToLower() }
        return $null
    }

    $item = $filteredItems.GetFirst()
    $count = 0
    while ($item -ne $null) {
        $count++
        try {
            $responseText = $responseStatusMap[[int]$item.ResponseStatus]
            if (-not $responseText) { $responseText = "Unknown($($item.ResponseStatus))" }

            # Resolve timezone — fall back to system local if unavailable
            $tz = $null
            try { $tz = $item.StartTimeZone.ID } catch {}
            if (-not $tz) { $tz = [System.TimeZoneInfo]::Local.Id }

            $results += [PSCustomObject]@{
                EntryID         = $item.EntryID
                OrganizerDomain = Get-OrganizerDomain -Item $item
                Start           = $item.Start.ToString("yyyy-MM-dd HH:mm")
                End             = $item.End.ToString("yyyy-MM-dd HH:mm")
                TimeZone        = $tz
                ResponseStatus  = $responseText
            }
        } catch {
            Write-Host "       WARNING: Could not read item $count - $($_.Exception.Message)" -ForegroundColor Yellow
        }
        $item = $filteredItems.GetNext()
    }

    Write-Host "       Found $count item(s) for today." -ForegroundColor Green
} catch {
    Write-Host "       FAILED: Could not query calendar items." -ForegroundColor Red
    Write-Host "       Error: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --------------------------------------------------
# Step 5: Output results
# --------------------------------------------------
Write-Host "[5/5] Writing results..." -ForegroundColor White

if ($results.Count -eq 0) {
    Write-Host "       No calendar entries found for today." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "       This could mean:" -ForegroundColor Gray
    Write-Host "         - Your calendar is empty for today" -ForegroundColor Gray
    Write-Host "         - The default calendar is not the one you expected" -ForegroundColor Gray
    Write-Host ""
} else {
    # Display on console as a table
    Write-Host ""
    $results | Format-Table -AutoSize | Out-String | Write-Host

    # Export to CSV
    $results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
    Write-Host "       CSV saved to: $outputFile" -ForegroundColor Green
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Smoke test complete." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
