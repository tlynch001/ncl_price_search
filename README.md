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
`-SortByPrice`, `-CsvPath`, `-NoCsv`).

## Price history (CSV)

Every run also appends a row per vacation (with a timestamp and the search
that produced it) to a CSV file, so repeated runs — for example on a
schedule via Windows Task Scheduler — build up a price history you can open
in Excel and chart over time.

By default, the CSV is written to `NCL Price Tracking\NCL-Vacation-Price-History.csv`
inside your OneDrive folder, auto-detected from the `OneDriveConsumer` (used
for personal/family Microsoft accounts) or `OneDrive` environment variables
that the OneDrive desktop app sets. As long as you're signed into OneDrive
with your Microsoft 365 family account, that folder syncs to the cloud
automatically — no extra upload step required. If no OneDrive folder can be
found (e.g. on macOS/Linux, or OneDrive isn't installed/signed in), the
script falls back to writing the CSV next to itself and prints a warning.

Point it at a specific file instead:

```powershell
.\Search-NclVacations.ps1 -CsvPath "$env:OneDriveConsumer\Documents\ncl-prices.csv"
```

Or skip the CSV entirely:

```powershell
.\Search-NclVacations.ps1 -NoCsv
```

The CSV columns are: `Timestamp`, `SearchQuery`, `ItineraryCode`,
`PackageId`, `Title`, `Ship`, `Days`, `Price`, `Currency`, `IsDeal`,
`AlertThreshold`.

Requires PowerShell 5.1+ (Windows PowerShell) or PowerShell 7+ (`pwsh`), and
internet access to `www.ncl.com`.

## Running twice a day (8 AM / 8 PM)

### Windows (Task Scheduler)

`Register-NclVacationsSchedule.ps1` registers a Windows Scheduled Task that
runs `Search-NclVacations.ps1` automatically at 8:00 AM and 8:00 PM every
day:

```powershell
.\Register-NclVacationsSchedule.ps1
```

Since a scheduled task has no visible console, each run's output (including
any `***ALERT***` lines) is appended to a log file — by default
`NCL Price Tracking\NCL-Vacation-Search-Log.txt` inside your OneDrive
folder, right next to the CSV. The price history CSV itself is still
controlled by `Search-NclVacations.ps1`'s own `-CsvPath`/`-NoCsv` options;
forward them through `-ScriptArguments` if you want to customize the search:

```powershell
.\Register-NclVacationsSchedule.ps1 -ScriptArguments "-EmbPorts JAX -AlertThreshold 320"
```

Other options:

```powershell
# Different times (24-hour "HH:mm")
.\Register-NclVacationsSchedule.ps1 -Times 08:00,14:00,20:00

# Run even when you're signed out (prompts once for your Windows password)
.\Register-NclVacationsSchedule.ps1 -RunWhetherLoggedOnOrNot

# Remove the scheduled task
.\Register-NclVacationsSchedule.ps1 -Unregister
```

By default the task only runs while you're signed in to Windows (no
password required) — that's normally fine as long as your PC is on and
you're logged in around 8 AM/8 PM. Use `-RunWhetherLoggedOnOrNot` if you
want it to run even when signed out.

See `Get-Help .\Register-NclVacationsSchedule.ps1 -Full` for all parameters.

### macOS/Linux (cron)

Task Scheduler isn't available outside Windows, but `pwsh` and `cron` work
the same way. Add a crontab entry that runs at 8 AM and 8 PM:

```bash
crontab -e
```

```cron
0 8,20 * * * /usr/bin/pwsh -NoProfile -File "/path/to/Search-NclVacations.ps1" *>> "/path/to/NCL-Vacation-Search-Log.txt"
```

(On macOS, point `-CsvPath` at your local iCloud/OneDrive-synced folder,
e.g. `~/OneDrive/NCL Price Tracking/NCL-Vacation-Price-History.csv`.)
