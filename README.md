# ncl_price_search

A PowerShell script that searches Norwegian Cruise Line's (NCL) vacation
finder for a given embarkation port, sailing months, and guest count, then
lists each matching vacation with its length (in days) and price. Any
vacation priced under a configurable threshold (default $320) gets an extra
attention-grabbing `***ALERT***` line.

Under the hood, [ncl.com/vacations](https://www.ncl.com/vacations) is a
single-page app that fetches its results from NCL's own JSON API at
`https://www.ncl.com/api/v2/vacations/search` using the same query
parameters found in the page URL. This script calls that API directly (no
browser automation required) and pages through all of the results.

## Usage

Run with the defaults, which reproduce this search exactly:

```
https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026,Jan-2027,Feb-2027,Mar-2027&guests=2
```

```powershell
.\Search-NclVacations.ps1
```

Customize the search with typed parameters:

```powershell
.\Search-NclVacations.ps1 -EmbPorts MIA -Dates Jun-2027,Jul-2027 -Guests 4 -AlertThreshold 500 -SortByPrice
```

Or reuse any full `ncl.com/vacations` URL copied from your browser:

```powershell
.\Search-NclVacations.ps1 -Url "https://www.ncl.com/vacations?embPorts=JAX&dates=Nov-2026,Dec-2026&guests=2"
```

See `Get-Help .\Search-NclVacations.ps1 -Full` for all parameters
(`-EmbPorts`, `-Dates`, `-Guests`, `-Url`, `-AlertThreshold`, `-PageSize`,
`-SortByPrice`).

Requires PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+ (`pwsh`), and
internet access to `www.ncl.com`.
