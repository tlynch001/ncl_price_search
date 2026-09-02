<#
.SYNOPSIS
    Raspberry-Pi-friendly NCL price monitor with history-aware deal alerts
    and optional Fastmail SMTP notifications.

.DESCRIPTION
    Calls NCL's JSON vacation search API directly, resolves sail dates,
    appends observations to the same CSV schema used by Search-NclVacations.ps1,
    evaluates price movement against prior history, and sends one aggregate
    email when a meaningful deal event occurs.

    Alerts are event-driven, not "still cheap" reminders. A sailing can trigger
    when it crosses below an absolute threshold, drops by a configured percent
    from the previous observation, sets a new recent low, or sets a new all-time
    low. If the same price persists on the next run, it will not alert again.

    Fastmail credentials are read from environment variables by default:
      NCL_SMTP_USERNAME
      NCL_SMTP_PASSWORD   (Fastmail app password, not account password)
      NCL_ALERT_FROM
      NCL_ALERT_TO
#>

[CmdletBinding()]
param(
    [string[]]$EmbPorts = @('JAX'),
    [string[]]$Dates = @('Nov-2026', 'Dec-2026', 'Jan-2027', 'Feb-2027', 'Mar-2027'),
    [int]$Guests = 2,
    [string]$Url,
    [ValidateRange(1, 200)][int]$PageSize = 50,
    [switch]$SortByPrice,
    [switch]$SkipCruiseDate,
    [string]$CsvPath,

    [double]$AbsoluteAlertThreshold = 350,
    [ValidateRange(0.1, 100.0)][double]$DropPercentThreshold = 10,
    [ValidateRange(1, 3650)][int]$RecentLowDays = 30,

    [switch]$EmailAlerts,
    [string]$EmailTo = $env:NCL_ALERT_TO,
    [string]$EmailFrom = $env:NCL_ALERT_FROM,
    [string]$SmtpUsername = $env:NCL_SMTP_USERNAME,
    [string]$SmtpPassword = $env:NCL_SMTP_PASSWORD,
    [string]$SmtpHost = 'smtp.fastmail.com',
    [int]$SmtpPort = 587
)

$ErrorActionPreference = 'Stop'
$ApiRoot = 'https://www.ncl.com/api/v2/vacations/search'
$SailingsApiRoot = 'https://www.ncl.com/api/vacations/sailings'
$UserAgent = 'Mozilla/5.0 (X11; Linux aarch64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36'

function Get-NclQueryString {
    if ($Url) {
        $uri = [System.Uri]$Url
        $query = $uri.Query.TrimStart('?')
        if ([string]::IsNullOrWhiteSpace($query)) {
            throw 'The supplied -Url has no query parameters.'
        }
        return $query
    }

    $parts = @()
    if ($EmbPorts.Count -gt 0) { $parts += 'embPorts=' + ($EmbPorts -join ',') }
    if ($Dates.Count -gt 0) { $parts += 'dates=' + ($Dates -join ',') }
    $parts += "guests=$Guests"
    return ($parts -join '&')
}

function Get-NclVacations {
    param([string]$QueryString)

    $results = New-Object System.Collections.Generic.List[object]
    $offset = 0
    $total = [int]::MaxValue

    while ($offset -lt $total) {
        $pageUrl = "{0}?{1}&limit={2}&offset={3}" -f $ApiRoot, $QueryString, $PageSize, $offset
        $response = Invoke-RestMethod -Uri $pageUrl -Method Get -Headers @{ Accept = 'application/json' } -UserAgent $UserAgent
        $total = [int]$response.total
        $itineraries = @($response.itineraries)
        if ($itineraries.Count -eq 0) { break }

        foreach ($item in $itineraries) {
            $results.Add([PSCustomObject]@{
                ItineraryCode = [string]$item.code
                PackageId     = [string]$item.packageId
                Title         = [string]$item.title
                Days          = [int]$item.duration.days
                Price         = [double]$item.combinedPrice
                Currency      = [string]$item.currencyCode
                Ship          = [string]$item.ship.title
                CruiseDate    = $null
            })
        }
        $offset += $itineraries.Count
    }
    return $results
}

function Resolve-NclCruiseDates {
    param(
        [System.Collections.Generic.List[object]]$Vacations,
        [string]$QueryString
    )

    $datesByKey = @{}
    $codes = @($Vacations | Select-Object -ExpandProperty ItineraryCode -Unique)

    foreach ($code in $codes) {
        $sailingsUrl = "{0}/{1}?{2}" -f $SailingsApiRoot, $code, $QueryString
        try {
            $response = Invoke-RestMethod -Uri $sailingsUrl -Method Get -Headers @{ Accept = 'application/json' } -UserAgent $UserAgent
        }
        catch {
            Write-Warning "Could not resolve sail date for '$code': $_"
            continue
        }

        foreach ($room in @($response.pricingStateRooms)) {
            if (-not $room.sailStartDate) { continue }
            $key = "$code|$($room.packageId)"
            if (-not $datesByKey.ContainsKey($key)) {
                $datesByKey[$key] = [datetime]$room.sailStartDate
            }
        }
    }

    foreach ($vacation in $Vacations) {
        $key = "$($vacation.ItineraryCode)|$($vacation.PackageId)"
        if ($datesByKey.ContainsKey($key)) {
            $vacation.CruiseDate = $datesByKey[$key]
        }
    }
}

function Get-DefaultCsvPath {
    $dataDir = Join-Path $PSScriptRoot 'data'
    return Join-Path $dataDir 'NCL-Vacation-Price-History.csv'
}

function Read-History {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    try { return @(Import-Csv -LiteralPath $Path) }
    catch {
        Write-Warning "Could not read history '$Path': $_"
        return @()
    }
}

function Get-HistoryForVacation {
    param(
        [object[]]$History,
        [object]$Vacation
    )

    return @($History | Where-Object {
        $_.ItineraryCode -eq $Vacation.ItineraryCode -and
        $_.PackageId -eq $Vacation.PackageId
    } | ForEach-Object {
        $parsedTime = [datetime]::MinValue
        [datetime]::TryParse($_.Timestamp, [ref]$parsedTime) | Out-Null
        [PSCustomObject]@{
            Timestamp = $parsedTime
            Price = [double]$_.Price
        }
    } | Where-Object { $_.Timestamp -ne [datetime]::MinValue } | Sort-Object Timestamp)
}

function Test-DealEvents {
    param(
        [object]$Vacation,
        [object[]]$History,
        [datetime]$Now
    )

    $prior = @(Get-HistoryForVacation -History $History -Vacation $Vacation)
    $previous = if ($prior.Count -gt 0) { $prior[-1] } else { $null }
    $reasons = New-Object System.Collections.Generic.List[string]

    $previousPrice = $null
    $dropPercent = $null
    if ($previous) {
        $previousPrice = [double]$previous.Price
        if ($previousPrice -gt 0 -and $Vacation.Price -lt $previousPrice) {
            $dropPercent = (($previousPrice - $Vacation.Price) / $previousPrice) * 100.0
            if ($dropPercent -ge $DropPercentThreshold) {
                $reasons.Add(('dropped {0:N1}% since last check (${1:N2} -> ${2:N2})' -f $dropPercent, $previousPrice, $Vacation.Price))
            }
        }
    }

    $crossedAbsolute = $Vacation.Price -lt $AbsoluteAlertThreshold -and (
        -not $previous -or $previousPrice -ge $AbsoluteAlertThreshold
    )
    if ($crossedAbsolute) {
        $reasons.Add(('crossed below ${0:N0}' -f $AbsoluteAlertThreshold))
    }

    if ($prior.Count -gt 0) {
        $allTimeMin = ($prior | Measure-Object Price -Minimum).Minimum
        if ($Vacation.Price -lt $allTimeMin) {
            $reasons.Add(('new all-time low (prior low ${0:N2})' -f $allTimeMin))
        }

        $cutoff = $Now.AddDays(-$RecentLowDays)
        $recent = @($prior | Where-Object { $_.Timestamp -ge $cutoff })
        if ($recent.Count -gt 0) {
            $recentMin = ($recent | Measure-Object Price -Minimum).Minimum
            if ($Vacation.Price -lt $recentMin -and $Vacation.Price -ge $allTimeMin) {
                $reasons.Add(('new {0}-day low (prior low ${1:N2})' -f $RecentLowDays, $recentMin))
            }
        }
    }

    return [PSCustomObject]@{
        Vacation = $Vacation
        PreviousPrice = $previousPrice
        DropPercent = $dropPercent
        Reasons = @($reasons)
        ShouldAlert = $reasons.Count -gt 0
    }
}

function Send-FastmailAlert {
    param([object[]]$Events)

    if (-not $EmailAlerts -or $Events.Count -eq 0) { return }

    $missing = @()
    if ([string]::IsNullOrWhiteSpace($EmailTo)) { $missing += 'NCL_ALERT_TO/-EmailTo' }
    if ([string]::IsNullOrWhiteSpace($EmailFrom)) { $missing += 'NCL_ALERT_FROM/-EmailFrom' }
    if ([string]::IsNullOrWhiteSpace($SmtpUsername)) { $missing += 'NCL_SMTP_USERNAME/-SmtpUsername' }
    if ([string]::IsNullOrWhiteSpace($SmtpPassword)) { $missing += 'NCL_SMTP_PASSWORD/-SmtpPassword' }
    if ($missing.Count -gt 0) {
        throw 'Email alerts requested but these settings are missing: ' + ($missing -join ', ')
    }

    $top = $Events | Sort-Object { $_.Vacation.Price } | Select-Object -First 1
    $subject = 'NCL DEAL ALERT: {0} {1} at ${2:N2}' -f $top.Vacation.Ship, $(if ($top.Vacation.CruiseDate) { $top.Vacation.CruiseDate.ToString('MMM d') } else { $top.Vacation.Title }), $top.Vacation.Price

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("NCL price monitor found $($Events.Count) meaningful deal event(s).")
    $lines.Add('')

    foreach ($event in ($Events | Sort-Object { $_.Vacation.Price })) {
        $v = $event.Vacation
        $dateText = if ($v.CruiseDate) { $v.CruiseDate.ToString('MMMM d, yyyy') } else { 'date unresolved' }
        $lines.Add("$($v.Ship) — $dateText")
        $lines.Add("$($v.Title)")
        $lines.Add(('Current price: ${0:N2} per person' -f $v.Price))
        foreach ($reason in $event.Reasons) { $lines.Add("- $reason") }
        $lines.Add('')
    }

    $lines.Add("Search: https://www.ncl.com/vacations?$(Get-NclQueryString)")
    $body = $lines -join [Environment]::NewLine

    $message = [System.Net.Mail.MailMessage]::new($EmailFrom, $EmailTo, $subject, $body)
    $client = [System.Net.Mail.SmtpClient]::new($SmtpHost, $SmtpPort)
    $client.EnableSsl = $true
    $client.Credentials = [System.Net.NetworkCredential]::new($SmtpUsername, $SmtpPassword)

    try {
        $client.Send($message)
        Write-Host "Sent deal alert email to $EmailTo." -ForegroundColor Green
    }
    finally {
        $message.Dispose()
        $client.Dispose()
    }
}

function Export-History {
    param(
        [System.Collections.Generic.List[object]]$Vacations,
        [string]$Path,
        [string]$SearchQuery,
        [datetime]$Now
    )

    $folder = Split-Path -Path $Path -Parent
    if ($folder -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $rows = foreach ($vacation in $Vacations) {
        [PSCustomObject]@{
            Timestamp      = $Now.ToString('yyyy-MM-dd HH:mm:ss')
            SearchQuery    = $SearchQuery
            ItineraryCode  = $vacation.ItineraryCode
            PackageId      = $vacation.PackageId
            Title          = $vacation.Title
            Ship           = $vacation.Ship
            CruiseDate     = if ($vacation.CruiseDate) { $vacation.CruiseDate.ToString('yyyy-MM-dd') } else { '' }
            Days           = $vacation.Days
            Price          = $vacation.Price
            Currency       = $vacation.Currency
            IsDeal         = $vacation.Price -lt $AbsoluteAlertThreshold
            AlertThreshold = $AbsoluteAlertThreshold
        }
    }

    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Append -Encoding UTF8
}

$queryString = Get-NclQueryString
if (-not $CsvPath) { $CsvPath = Get-DefaultCsvPath }
$now = Get-Date

Write-Host "Searching NCL: https://www.ncl.com/vacations?$queryString" -ForegroundColor Cyan
$vacations = Get-NclVacations -QueryString $queryString
if (-not $SkipCruiseDate) { Resolve-NclCruiseDates -Vacations $vacations -QueryString $queryString }
if ($SortByPrice) { $vacations = [System.Collections.Generic.List[object]]($vacations | Sort-Object Price) }

$history = @(Read-History -Path $CsvPath)
$events = New-Object System.Collections.Generic.List[object]

foreach ($vacation in $vacations) {
    $dateText = if ($vacation.CruiseDate) { $vacation.CruiseDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
    Write-Host ('{0,-20} {1}  {2,8:C2}  {3}' -f $vacation.Ship, $dateText, $vacation.Price, $vacation.Title)

    $event = Test-DealEvents -Vacation $vacation -History $history -Now $now
    if ($event.ShouldAlert) {
        $events.Add($event)
        foreach ($reason in $event.Reasons) {
            Write-Host "  *** ALERT: $reason" -ForegroundColor Yellow
        }
    }
}

# Send before appending current observations so a transient SMTP failure does not
# make the current prices look like already-processed history on a retry.
Send-FastmailAlert -Events @($events)
Export-History -Vacations $vacations -Path $CsvPath -SearchQuery $queryString -Now $now

Write-Host "Saved $($vacations.Count) observation(s) to $CsvPath." -ForegroundColor Green
Write-Host "Meaningful alerts this run: $($events.Count)." -ForegroundColor Cyan
