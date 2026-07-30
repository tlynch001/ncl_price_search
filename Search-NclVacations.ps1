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

    For any vacation whose price is below -AlertThreshold, an extra
    attention-grabbing "***ALERT***" line is printed immediately below it.

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

.EXAMPLE
    .\Search-NclVacations.ps1

    Runs the exact search from the NCL URL:
    https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026,Jan-2027,Feb-2027,Mar-2027&guests=2

.EXAMPLE
    .\Search-NclVacations.ps1 -EmbPorts MIA -Dates Jun-2027,Jul-2027 -Guests 4 -AlertThreshold 500 -SortByPrice

.EXAMPLE
    .\Search-NclVacations.ps1 -Url "https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026&guests=2"
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

    [switch]$SortByPrice
)

$ErrorActionPreference = 'Stop'

# Ensure TLS 1.2 is used on older Windows PowerShell hosts (no-op on PS 6+/pwsh).
if ($PSVersionTable.PSVersion.Major -lt 6) {
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
}

$ApiRoot = 'https://www.ncl.com/api/v2/vacations/search'
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
                Title    = $item.title
                Days     = $item.duration.days
                Price    = [double]$item.combinedPrice
                Currency = $item.currencyCode
                Ship     = $item.ship.title
            })
        }

        $offset += $itineraries.Count
    }

    return $results
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

        Write-Host ""
        Write-Host "Vacation: $($vacation.Title)"
        Write-Host "Ship:     $($vacation.Ship)"
        Write-Host "Days:     $($vacation.Days)"
        Write-Host "Price:    $formattedPrice per person"

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

$queryString = Get-NclQueryString -Url $Url -EmbPorts $EmbPorts -Dates $Dates -Guests $Guests

Write-Host "Searching NCL vacations (https://www.ncl.com/vacations?$queryString) ..." -ForegroundColor Cyan

$vacations = Get-NclVacations -QueryString $queryString -PageSize $PageSize

if ($SortByPrice) {
    $vacations = [System.Collections.Generic.List[object]]($vacations | Sort-Object -Property Price)
}

Write-VacationReport -Vacations $vacations -AlertThreshold $AlertThreshold
