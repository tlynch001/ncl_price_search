<#
.SYNOPSIS
    Searches the Norwegian Cruise Line (NCL) vacations site and lists each
    matching vacation along with its length (in days) and price, flagging
    any vacation priced below a configurable "deal" threshold.

.DESCRIPTION
    The public-facing page at https://www.ncl.com/vacations is a single-page
    app that renders its results client-side by calling NCL's own JSON
    search API at https://www.ncl.com/api/v2/vacations/search using the
    exact same query-string parameters found on the page URL (embPorts,
    dates, guests, etc.). This script calls that same API directly, pages
    through all of the results, and prints a simple report.

    Each result also shows the actual cruise (sail) date, resolved via one
    extra call per unique itinerary to NCL's per-itinerary "sailings"
    endpoint (the search API itself only returns the lowest price found
    across the whole requested date range, not which date it applies to).

    For any vacation whose price is below -AlertThreshold, an extra
    attention-grabbing "***ALERT***" line is printed immediately below it.

    Every run is also appended (not overwritten) as rows to a CSV file, with
    a timestamp column, so that repeated runs (e.g. via Task Scheduler) build
    up a price history you can chart in Excel over time. By default the CSV
    is written into a "NCL Price Tracking" folder inside your OneDrive
    folder (auto-detected via the $env:OneDriveConsumer / $env:OneDrive
    environment variables set by the OneDrive desktop app), so it syncs to
    your Microsoft 365 account automatically. Use -CsvPath to point
    somewhere else, or -NoCsv to skip writing the CSV entirely.

.PARAMETER EmbPorts
    One or more embarkation port codes to search from (e.g. JAX for
    Jacksonville, FL). Matches the "embPorts" parameter on ncl.com.

.PARAMETER Dates
    One or more month/year tokens to search (formatted like "Nov-2026").
    Matches the "dates" parameter on ncl.com.

.PARAMETER Guests
    Number of guests to price the vacation for. Matches the "guests"
    parameter on ncl.com.

.PARAMETER Url
    Optional: instead of building the query from -EmbPorts/-Dates/-Guests,
    supply a full ncl.com/vacations URL (exactly like the one you'd copy out
    of your browser's address bar) and its query string will be reused as-is.
    When provided, -EmbPorts/-Dates/-Guests are ignored.

.PARAMETER AlertThreshold
    Any vacation priced strictly below this dollar amount gets a
    "***ALERT***" callout line. Defaults to 320.

.PARAMETER PageSize
    How many results to request per page while paging through the API.
    Defaults to 50, which is large enough to cover most single-port
    searches in one request.

.PARAMETER SortByPrice
    If specified, results are sorted from cheapest to most expensive before
    being printed (the API's default order is NCL's own "Featured" order).

.PARAMETER SkipCruiseDate
    If specified, skips the extra per-itinerary lookup used to resolve each
    result's actual cruise (sail) date, leaving the "Cruise Date" column
    blank/"Unknown". Use this to speed up large, unfiltered searches (it
    makes one additional API call per unique itinerary in the results).

.PARAMETER CsvPath
    Path to the CSV file that results are appended to (the file and any
    parent folders are created automatically on first run; a header row is
    written once). Defaults to "NCL Price Tracking\NCL-Vacation-Price-History.csv"
    inside your OneDrive folder, auto-detected from $env:OneDriveConsumer
    (personal/family Microsoft accounts) or $env:OneDrive. Falls back to the
    script's own folder if no OneDrive folder can be detected (for example
    on non-Windows hosts, or when OneDrive isn't installed/signed in).

.PARAMETER NoCsv
    If specified, skips writing/appending to the CSV file entirely.

.EXAMPLE
    .\Search-NclVacations.ps1

    Runs the exact search from the NCL URL:
    https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026,Jan-2027,Feb-2027,Mar-2027&guests=2
    and appends the results to the default CSV in your OneDrive folder.

.EXAMPLE
    .\Search-NclVacations.ps1 -EmbPorts MIA -Dates Jun-2027,Jul-2027 -Guests 4 -AlertThreshold 500 -SortByPrice

.EXAMPLE
    .\Search-NclVacations.ps1 -Url "https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026&guests=2"

.EXAMPLE
    .\Search-NclVacations.ps1 -CsvPath "$env:OneDriveConsumer\Documents\ncl-prices.csv"

.EXAMPLE
    .\Search-NclVacations.ps1 -NoCsv
#>

[CmdletBinding(DefaultParameterSetName = 'Discrete')]
param(
    [Parameter(ParameterSetName = 'Discrete')]
    [string[]]$EmbPorts = @('JAX'),

    [Parameter(ParameterSetName = 'Discrete')]
    [string[]]$Dates = @('Nov-2026', 'Dec-2026', 'Jan-2027', 'Feb-2027', 'Mar-2027'),

    [Parameter(ParameterSetName = 'Discrete')]
    [int]$Guests = 2,

    [Parameter(ParameterSetName = 'Url', Mandatory = $true)]
    [string]$Url,

    [double]$AlertThreshold = 320,

    [ValidateRange(1, 200)]
    [int]$PageSize = 50,

    [switch]$SortByPrice,

    [switch]$SkipCruiseDate,

    [string]$CsvPath,

    [switch]$NoCsv
)

$ErrorActionPreference = 'Stop'

# Ensure TLS 1.2 is used on older Windows PowerShell hosts (no-op on PS 6+/pwsh).
if ($PSVersionTable.PSVersion.Major -lt 6) {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

$ApiRoot = 'https://www.ncl.com/api/v2/vacations/search'
$SailingsApiRoot = 'https://www.ncl.com/api/vacations/sailings'
$UserAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

function Get-NclQueryString {
    <#
        Builds (or extracts) the query string that will be forwarded to the
        API. NCL's own front-end forwards the page's query string verbatim
        to /api/v2/vacations/search, so we do the same thing here.
    #>
    param(
        [string]$Url,
        [string[]]$EmbPorts,
        [string[]]$Dates,
        [int]$Guests
    )

    if ($Url) {
        $uri = [System.Uri]$Url
        $query = $uri.Query.TrimStart('?')
        if ([string]::IsNullOrWhiteSpace($query)) {
            throw "The supplied -Url does not contain any query parameters (embPorts/dates/guests)."
        }
        return $query
    }

    $parts = @()
    if ($EmbPorts -and $EmbPorts.Count -gt 0) {
        $parts += 'embPorts=' + ($EmbPorts -join ',')
    }
    if ($Dates -and $Dates.Count -gt 0) {
        $parts += 'dates=' + ($Dates -join ',')
    }
    $parts += "guests=$Guests"

    return ($parts -join '&')
}

function Get-NclVacations {
    <#
        Pages through the NCL vacations search API and returns every
        matching itinerary as a flat array of custom objects.
    #>
    param(
        [string]$QueryString,
        [int]$PageSize
    )

    $results = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $total = [int]::MaxValue

    while ($offset -lt $total) {
        $pageUrl = "{0}?{1}&limit={2}&offset={3}" -f $ApiRoot, $QueryString, $PageSize, $offset

        try {
            $response = Invoke-RestMethod -Uri $pageUrl -Method Get -Headers @{ Accept = 'application/json' } -UserAgent $UserAgent
        }
        catch {
            throw "Failed to query the NCL vacations API at '$pageUrl'. $_"
        }

        $total = [int]$response.total
        $itineraries = @($response.itineraries)

        if ($itineraries.Count -eq 0) {
            break
        }

        foreach ($item in $itineraries) {
            $results.Add([PSCustomObject]@{
                ItineraryCode = $item.code
                PackageId     = $item.packageId
                Title         = $item.title
                Days          = $item.duration.days
                Price         = [double]$item.combinedPrice
                Currency      = $item.currencyCode
                Ship          = $item.ship.title
                CruiseDate    = $null
            })
        }

        $offset += $itineraries.Count
    }

    return $results
}

function Resolve-NclCruiseDates {
    <#
        The search API only returns each itinerary's lowest price across
        the whole requested date range, not which specific sail date that
        price applies to. To resolve an actual cruise date, this calls
        NCL's per-itinerary "sailings" endpoint (the same one the site's
        "Choose Your Sail Dates" panel uses) once per unique itinerary code
        in the results, then matches each vacation's PackageId to the
        corresponding sailStartDate and fills in its CruiseDate property.
    #>
    param(
        [System.Collections.Generic.List[object]]$Vacations,
        [string]$QueryString
    )

    $sailDatesByKey = @{}
    $itineraryCodes = @($Vacations | Select-Object -ExpandProperty ItineraryCode -Unique)
    $current = 0

    foreach ($code in $itineraryCodes) {
        $current++
        Write-Progress -Activity 'Looking up cruise dates' -Status $code -PercentComplete (100 * $current / [Math]::Max(1, $itineraryCodes.Count))

        $sailingsUrl = "{0}/{1}?{2}" -f $SailingsApiRoot, $code, $QueryString
        try {
            $response = Invoke-RestMethod -Uri $sailingsUrl -Method Get -Headers @{ Accept = 'application/json' } -UserAgent $UserAgent
        }
        catch {
            Write-Warning "Could not look up the cruise date for itinerary '$code': $_"
            continue
        }

        foreach ($room in @($response.pricingStateRooms)) {
            $key = "$code|$($room.packageId)"
            if (-not $sailDatesByKey.ContainsKey($key) -and $room.sailStartDate) {
                $sailDatesByKey[$key] = [datetime]$room.sailStartDate
            }
        }
    }

    Write-Progress -Activity 'Looking up cruise dates' -Completed

    foreach ($vacation in $Vacations) {
        $key = "$($vacation.ItineraryCode)|$($vacation.PackageId)"
        if ($sailDatesByKey.ContainsKey($key)) {
            $vacation.CruiseDate = $sailDatesByKey[$key]
        }
    }
}

function Write-VacationReport {
    param(
        [System.Collections.Generic.List[object]]$Vacations,
        [double]$AlertThreshold
    )

    if ($Vacations.Count -eq 0) {
        Write-Host "No vacations were found for this search." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "Found $($Vacations.Count) vacation(s):" -ForegroundColor Cyan
    Write-Host ("=" * 60)

    foreach ($vacation in $Vacations) {
        # Format explicitly as USD regardless of the host's locale, since
        # NCL prices its US site in dollars (see $vacation.Currency).
        $formattedPrice = '${0:N2}' -f $vacation.Price

        $formattedCruiseDate = if ($vacation.CruiseDate) { $vacation.CruiseDate.ToString('MMMM d, yyyy') } else { 'Unknown' }

        Write-Host ""
        Write-Host "Vacation:    $($vacation.Title)"
        Write-Host "Ship:        $($vacation.Ship)"
        Write-Host "Cruise Date: $formattedCruiseDate"
        Write-Host "Days:        $($vacation.Days)"
        Write-Host "Price:       $formattedPrice per person"

        if ($vacation.Price -lt $AlertThreshold) {
            Write-Host "*** ALERT *** Price $formattedPrice is under `$$AlertThreshold! Grab this deal! *** ALERT ***" -ForegroundColor Red -BackgroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host ("=" * 60)

    $dealCount = ($Vacations | Where-Object { $_.Price -lt $AlertThreshold }).Count
    if ($dealCount -gt 0) {
        Write-Host "$dealCount of $($Vacations.Count) vacation(s) are under `$$AlertThreshold." -ForegroundColor Red
    }
}

function Get-DefaultCsvPath {
    <#
        Auto-detects the local OneDrive sync folder so results land somewhere
        that syncs to a Microsoft 365 account without any extra setup.
        $env:OneDriveConsumer is set by the OneDrive desktop app when signed
        into a personal/family Microsoft account; $env:OneDrive points at
        whichever OneDrive account was configured first (personal or work).
    #>
    $oneDriveRoot = $env:OneDriveConsumer
    if (-not $oneDriveRoot) {
        $oneDriveRoot = $env:OneDrive
    }

    if ($oneDriveRoot -and (Test-Path -LiteralPath $oneDriveRoot)) {
        return Join-Path -Path (Join-Path -Path $oneDriveRoot -ChildPath 'NCL Price Tracking') -ChildPath 'NCL-Vacation-Price-History.csv'
    }

    Write-Warning "Could not find a OneDrive folder (checked `$env:OneDriveConsumer and `$env:OneDrive). Writing the CSV next to the script instead. Pass -CsvPath to choose a specific location, e.g. your OneDrive folder."
    $scriptFolder = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    return Join-Path -Path $scriptFolder -ChildPath 'NCL-Vacation-Price-History.csv'
}

function Export-VacationHistory {
    <#
        Appends one row per vacation to a CSV file, tagged with a timestamp
        and the search that produced it, so repeated runs build up a price
        history over time. Creates the file (with header) and any parent
        folders on first use; subsequent runs only append data rows.
    #>
    param(
        [System.Collections.Generic.List[object]]$Vacations,
        [string]$CsvPath,
        [string]$SearchQuery,
        [double]$AlertThreshold,
        [datetime]$Timestamp
    )

    if ($Vacations.Count -eq 0) {
        return
    }

    $folder = Split-Path -Path $CsvPath -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $columnNames = @('Timestamp', 'SearchQuery', 'ItineraryCode', 'PackageId', 'Title', 'Ship', 'CruiseDate', 'Days', 'Price', 'Currency', 'IsDeal', 'AlertThreshold')

    if (Test-Path -LiteralPath $CsvPath) {
        $existingHeader = Get-Content -LiteralPath $CsvPath -TotalCount 1 -ErrorAction SilentlyContinue
        $expectedHeader = ($columnNames | ForEach-Object { '"' + $_ + '"' }) -join ','
        if ($existingHeader -and $existingHeader -ne $expectedHeader) {
            # PowerShell's Export-Csv -Append silently keeps whatever columns
            # are already in the file's header and drops anything that isn't
            # there (e.g. a newly added CruiseDate column) rather than
            # erroring -- so warn loudly instead of losing data quietly.
            Write-Warning "'$CsvPath' already has a different set of columns than this version of the script writes. PowerShell will silently DROP any new columns (like CruiseDate) when appending -- it won't error. To capture every column going forward, either rename/archive the old file so a fresh one is created, or manually update its header row to:`n$expectedHeader"
        }
    }

    $rows = foreach ($vacation in $Vacations) {
        [PSCustomObject]@{
            Timestamp      = $Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
            SearchQuery    = $SearchQuery
            ItineraryCode  = $vacation.ItineraryCode
            PackageId      = $vacation.PackageId
            Title          = $vacation.Title
            Ship           = $vacation.Ship
            CruiseDate     = if ($vacation.CruiseDate) { $vacation.CruiseDate.ToString('yyyy-MM-dd') } else { '' }
            Days           = $vacation.Days
            Price          = $vacation.Price
            Currency       = $vacation.Currency
            IsDeal         = $vacation.Price -lt $AlertThreshold
            AlertThreshold = $AlertThreshold
        }
    }

    try {
        $rows | Export-Csv -Path $CsvPath -NoTypeInformation -Append -Encoding UTF8
        Write-Host ""
        Write-Host "Appended $($rows.Count) row(s) to '$CsvPath'." -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not write to '$CsvPath': $_ (Is the file open in Excel, or is OneDrive still syncing/signed out? If you have an older CSV from before the 'CruiseDate' column was added, either rename/archive it so a fresh one can be created, or manually add a 'CruiseDate' column to its header row.)"
    }
}

$queryString = Get-NclQueryString -Url $Url -EmbPorts $EmbPorts -Dates $Dates -Guests $Guests

Write-Host "Searching NCL vacations (https://www.ncl.com/vacations?$queryString) ..." -ForegroundColor Cyan

$vacations = Get-NclVacations -QueryString $queryString -PageSize $PageSize

if (-not $SkipCruiseDate) {
    Resolve-NclCruiseDates -Vacations $vacations -QueryString $queryString
}

if ($SortByPrice) {
    $vacations = [System.Collections.Generic.List[object]]($vacations | Sort-Object -Property Price)
}

Write-VacationReport -Vacations $vacations -AlertThreshold $AlertThreshold

if (-not $NoCsv) {
    if (-not $CsvPath) {
        $CsvPath = Get-DefaultCsvPath
    }
    Export-VacationHistory -Vacations $vacations -CsvPath $CsvPath -SearchQuery $queryString -AlertThreshold $AlertThreshold -Timestamp (Get-Date)
}
