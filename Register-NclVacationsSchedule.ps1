<#
.SYNOPSIS
    Registers (or removes) a Windows Scheduled Task that runs
    Search-NclVacations.ps1 twice a day.

.DESCRIPTION
    Creates a Task Scheduler task (default name "NCL Vacation Price Check")
    with daily triggers at 8:00 AM and 8:00 PM (customizable via -Times)
    that launch PowerShell to run Search-NclVacations.ps1 in the background.
    Each run's console output (including any ***ALERT*** lines) is appended
    to a log file, since a scheduled task has no visible console of its own
    -- the CSV price history and this log are how you'll see what happened.

    If a task with the same name (see -TaskName) already exists -- for
    example one registered by an older version of this script -- it is
    removed first and replaced with the one described by the parameters
    you pass this time. Re-running this script is therefore the supported
    way to update an existing schedule (change the times, the search
    parameters, etc.): it always leaves exactly one task in place, matching
    your latest settings.

    Any parameters you'd normally pass to Search-NclVacations.ps1 (e.g.
    -EmbPorts, -Dates, -AlertThreshold, -CsvPath) can be supplied here via
    -ScriptArguments as a single string and are forwarded through untouched.

    This script only works on Windows (it uses the built-in ScheduledTasks
    module). See the README for a cron-based alternative on macOS/Linux.

.PARAMETER ScriptPath
    Path to Search-NclVacations.ps1. Defaults to the copy next to this
    script.

.PARAMETER ScriptArguments
    Extra arguments to forward to Search-NclVacations.ps1 on every run, e.g.
    "-EmbPorts JAX -AlertThreshold 320". Defaults to none (Search-NclVacations.ps1's
    own defaults are used, which reproduce the original NCL search).

.PARAMETER TaskName
    Name of the scheduled task. Defaults to "NCL Vacation Price Check".

.PARAMETER Times
    One or more times of day (24-hour "HH:mm") to run at. Defaults to
    08:00 and 20:00 (8 AM and 8 PM).

.PARAMETER LogPath
    Path to the text file that each run's console output is appended to.
    Defaults to "NCL Price Tracking\NCL-Vacation-Search-Log.txt" inside your
    OneDrive folder (same auto-detection as Search-NclVacations.ps1's CSV),
    falling back to a file next to this script if no OneDrive folder can be
    found.

.PARAMETER NoLog
    If specified, runs are not redirected to a log file (output is simply
    discarded, since a scheduled task has no console to show it on).

.PARAMETER RunWhetherLoggedOnOrNot
    If specified, the task runs even when you're signed out (you'll be
    prompted once for your Windows password so Task Scheduler can store it
    securely). Without this switch, the task only runs while you're signed
    in -- no password needed, but it won't fire if you're logged off at
    8 AM/8 PM.

.PARAMETER Unregister
    Removes the scheduled task instead of creating/updating it.

.EXAMPLE
    .\Register-NclVacationsSchedule.ps1

    Registers the task to run Search-NclVacations.ps1 (with its own
    defaults) at 8 AM and 8 PM daily, only while you're logged on.

.EXAMPLE
    .\Register-NclVacationsSchedule.ps1 -ScriptArguments "-EmbPorts JAX -AlertThreshold 320"

.EXAMPLE
    .\Register-NclVacationsSchedule.ps1 -Times 08:00,14:00,20:00

.EXAMPLE
    .\Register-NclVacationsSchedule.ps1 -RunWhetherLoggedOnOrNot

.EXAMPLE
    .\Register-NclVacationsSchedule.ps1 -Unregister
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ScriptPath = (Join-Path -Path $PSScriptRoot -ChildPath 'Search-NclVacations.ps1'),

    [string]$ScriptArguments = '',

    [string]$TaskName = 'NCL Vacation Price Check',

    [string[]]$Times = @('08:00', '20:00'),

    [string]$LogPath,

    [switch]$NoLog,

    [switch]$RunWhetherLoggedOnOrNot,

    [switch]$Unregister
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
    throw "This script requires Windows Task Scheduler (the built-in ScheduledTasks module), so it only works on Windows. See the README for a cron-based alternative on macOS/Linux."
}

if ($Unregister) {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed scheduled task '$TaskName'." -ForegroundColor Green
    }
    else {
        Write-Host "No scheduled task named '$TaskName' was found." -ForegroundColor Yellow
    }
    return
}

if (-not (Test-Path -LiteralPath $ScriptPath)) {
    throw "Could not find Search-NclVacations.ps1 at '$ScriptPath'. Pass -ScriptPath explicitly."
}

function Get-DefaultOneDriveFolder {
    $oneDriveRoot = $env:OneDriveConsumer
    if (-not $oneDriveRoot) {
        $oneDriveRoot = $env:OneDrive
    }
    if ($oneDriveRoot -and (Test-Path -LiteralPath $oneDriveRoot)) {
        return Join-Path -Path $oneDriveRoot -ChildPath 'NCL Price Tracking'
    }
    return $null
}

if (-not $NoLog -and -not $LogPath) {
    $folder = Get-DefaultOneDriveFolder
    if (-not $folder) {
        Write-Warning "Could not find a OneDrive folder (checked `$env:OneDriveConsumer and `$env:OneDrive). Writing the log next to this script instead. Pass -LogPath to choose a specific location."
        $folder = $PSScriptRoot
    }
    $LogPath = Join-Path -Path $folder -ChildPath 'NCL-Vacation-Search-Log.txt'
}

if (-not $NoLog) {
    $logFolder = Split-Path -Path $LogPath -Parent
    if ($logFolder -and -not (Test-Path -LiteralPath $logFolder)) {
        New-Item -ItemType Directory -Path $logFolder -Force | Out-Null
    }
}

$pwshCommand = Get-Command pwsh -ErrorAction SilentlyContinue
$exePath = if ($pwshCommand) { $pwshCommand.Source } else { (Get-Command powershell).Source }

# Run through -Command (rather than -File) so we can use PowerShell's own
# "*>>" redirection to append every output stream -- including the
# ***ALERT*** lines written via Write-Host -- to the log file. A scheduled
# task has no console of its own, so without this the output would just be
# discarded.
$innerCommand = "& '$ScriptPath' $ScriptArguments".Trim()
if (-not $NoLog) {
    $innerCommand += " *>> '$LogPath'"
}

$argumentList = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command `"$innerCommand`""

$action = New-ScheduledTaskAction -Execute $exePath -Argument $argumentList

$triggers = foreach ($time in $Times) {
    $parsedTime = [datetime]::ParseExact($time, 'HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
    New-ScheduledTaskTrigger -Daily -At $parsedTime
}

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

$existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

if ($PSCmdlet.ShouldProcess($TaskName, "Replace scheduled task to run at $($Times -join ', ')")) {
    if ($existingTask) {
        # Remove any existing task with this name first (e.g. one left over
        # from an older version of this script) rather than relying solely
        # on Register-ScheduledTask's -Force, so it's obvious exactly one
        # up-to-date task is left behind afterwards.
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed existing scheduled task '$TaskName'." -ForegroundColor Yellow
    }

    if ($RunWhetherLoggedOnOrNot) {
        $credential = Get-Credential -UserName "$env:USERDOMAIN\$env:USERNAME" -Message "Enter your Windows password so this task can run even when you're signed out"
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings `
            -User $credential.UserName -Password $credential.GetNetworkCredential().Password -RunLevel Limited -Force | Out-Null
    }
    else {
        $principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $triggers -Settings $settings -Principal $principal -Force | Out-Null
    }

    Write-Host ""
    $verb = if ($existingTask) { 'replaced with a new one' } else { 'registered' }
    Write-Host "Scheduled task '$TaskName' $verb to run daily at: $($Times -join ', ')" -ForegroundColor Green
    if (-not $NoLog) {
        Write-Host "Each run's output will be appended to: $LogPath"
    }
    Write-Host "Price history CSV location is controlled separately by Search-NclVacations.ps1 (see -CsvPath)."
    Write-Host ""
    Write-Host "To remove this task later, run:"
    Write-Host "  .\Register-NclVacationsSchedule.ps1 -TaskName '$TaskName' -Unregister"
}
