param(
    [string]$BaseDir = "C:\Arbeitszeit"
)

$ErrorActionPreference = "Stop"

$StatePath = Join-Path $BaseDir "state.json"
$ControlPath = Join-Path $BaseDir "control.json"
$CsvPath = Join-Path $BaseDir "Arbeitszeiten.csv"
$ActivityCsvPath = Join-Path $BaseDir "Arbeitszeit_Taetigkeiten.csv"
$ActivityJsonPath = Join-Path $BaseDir "taetigkeiten.json"
$IconPath = Join-Path $BaseDir "Arbeitszeit.ico"
$NotificationPath = Join-Path $BaseDir "notification.json"
$LogPath = Join-Path $BaseDir "ArbeitszeitAnzeige.log"
$SharedPath = Join-Path $PSScriptRoot "ArbeitszeitSettings.ps1"

if (!(Test-Path $BaseDir)) {
    New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null
}

if (!(Test-Path $SharedPath)) {
    throw "Gemeinsame Settings-Datei nicht gefunden: $SharedPath"
}

. $SharedPath

$mutexCreated = $false
$mutex = [System.Threading.Mutex]::new($true, "ArbeitszeitAnzeigeLocal", [ref]$mutexCreated)

if (-not $mutexCreated) {
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ("ArbeitszeitWindowHelper" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class ArbeitszeitWindowHelper
{
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
}
"@
}

$consoleWindow = [ArbeitszeitWindowHelper]::GetConsoleWindow()

if ($consoleWindow -ne [IntPtr]::Zero) {
    [ArbeitszeitWindowHelper]::ShowWindow($consoleWindow, 0) | Out-Null
}

function Convert-FromXaml {
    param(
        [string]$Xaml
    )

    $xml = [xml]$Xaml
    $reader = New-Object System.Xml.XmlNodeReader $xml
    return [System.Windows.Markup.XamlReader]::Load($reader)
}

function Write-AppLog {
    param(
        [string]$Message
    )

    try {
        Add-Content -LiteralPath $LogPath -Value ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + " " + $Message) -Encoding UTF8
    }
    catch {}
}

function Format-Duration {
    param(
        [double]$Seconds
    )

    if ($Seconds -lt 0) {
        $Seconds = 0
    }

    $secondsRounded = [math]::Floor($Seconds)
    $hours = [math]::Floor($secondsRounded / 3600)
    $minutes = [math]::Floor(($secondsRounded % 3600) / 60)
    $remainingSeconds = [math]::Floor($secondsRounded % 60)

    return "{0:00}:{1:00}:{2:00}" -f $hours, $minutes, $remainingSeconds
}

function Format-CompactDuration {
    param(
        [double]$Seconds
    )

    $sign = ""

    if ($Seconds -lt 0) {
        $sign = "-"
        $Seconds = [math]::Abs($Seconds)
    }

    $hours = [math]::Floor($Seconds / 3600)
    $minutes = [math]::Floor(($Seconds % 3600) / 60)

    return "{0}{1:00}:{2:00}" -f $sign, $hours, $minutes
}

function Convert-DurationTextToSeconds {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0.0
    }

    $parts = $Text.Trim().Split(":")

    if ($parts.Count -lt 2) {
        return 0.0
    }

    try {
        $hours = [int]$parts[0]
        $minutes = [int]$parts[1]
        $seconds = 0

        if ($parts.Count -ge 3) {
            $seconds = [int]$parts[2]
        }

        return [double](($hours * 3600) + ($minutes * 60) + $seconds)
    }
    catch {
        return 0.0
    }
}

function Get-WeekStart {
    param(
        [datetime]$Date = (Get-Date)
    )

    $offset = ([int]$Date.DayOfWeek + 6) % 7
    return $Date.Date.AddDays(-1 * $offset)
}

function Get-DayNameShort {
    param(
        [datetime]$Date
    )

    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
    return $culture.DateTimeFormat.GetAbbreviatedDayName($Date.DayOfWeek)
}

function Convert-ToTimeText {
    param(
        [string]$TimeText
    )

    if ([string]::IsNullOrWhiteSpace($TimeText)) {
        return $null
    }

    $formats = @("HH:mm:ss", "HH:mm")
    $parsed = [datetime]::MinValue

    foreach ($format in $formats) {
        if ([datetime]::TryParseExact(
            $TimeText.Trim(),
            $format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$parsed
        )) {
            return $parsed.ToString("HH:mm:ss")
        }
    }

    return $null
}

function Get-WorkDate {
    param(
        [string]$DateText
    )

    if ([string]::IsNullOrWhiteSpace($DateText)) {
        return $null
    }

    $date = [datetime]::MinValue

    foreach ($format in @("yyyy-MM-dd", "dd.MM.yyyy")) {
        if ([datetime]::TryParseExact(
            $DateText.Trim(),
            $format,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::None,
            [ref]$date
        )) {
            return $date.Date
        }
    }

    return $null
}

function Convert-ToDouble {
    param(
        [string]$Text,
        [double]$DefaultValue = 0
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $DefaultValue
    }

    $normalized = $Text.Trim().Replace(",", ".")
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $value = 0.0

    if ([double]::TryParse($normalized, [System.Globalization.NumberStyles]::Float, $culture, [ref]$value)) {
        return $value
    }

    return $DefaultValue
}

function Convert-ToWorkEntryHours {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return 0.0
    }

    $normalized = $Text.Trim().Replace(",", ".")
    $hours = 0.0

    if ([double]::TryParse(
        $normalized,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$hours
    )) {
        if ($hours -gt 0) {
            return $hours
        }
    }

    return 0.0
}

function Read-WorkEntries {
    if (!(Test-Path $ActivityJsonPath)) {
        return @()
    }

    try {
        $loaded = Get-Content -LiteralPath $ActivityJsonPath -Raw | ConvertFrom-Json

        if ($loaded -is [System.Array]) {
            foreach ($entry in $loaded) {
                $entry
            }
        }
        elseif ($null -ne $loaded) {
            $loaded
        }
    }
    catch {
        return @()
    }
}

function Save-WorkEntries {
    param(
        [object[]]$Entries
    )

    $json = "[]"

    if (@($Entries).Count -eq 0) {
        $json = "[]"
    }
    else {
        $json = @($Entries) | ConvertTo-Json -Depth 5
    }

    $tmpPath = $ActivityJsonPath + ".tmp"
    $backupPath = $ActivityJsonPath + ".bak"

    if (Test-Path $ActivityJsonPath) {
        Copy-Item -LiteralPath $ActivityJsonPath -Destination $backupPath -Force -ErrorAction SilentlyContinue
    }

    $json | Set-Content -LiteralPath $tmpPath -Encoding UTF8
    Move-Item -LiteralPath $tmpPath -Destination $ActivityJsonPath -Force
}

function Write-WorkEntriesCsv {
    param(
        [object[]]$Entries
    )

    $columns = @("Id", "Datum", "ErfasstAm", "Stunden", "Projekt", "Beschreibung")
    $normalized = @()
    $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

    foreach ($entry in @($Entries)) {
        $normalized += [PSCustomObject][ordered]@{
            Id           = [string]$entry.Id
            Datum        = [string]$entry.Datum
            ErfasstAm    = [string]$entry.ErfasstAm
            Stunden      = ([double]$entry.Stunden).ToString("N2", $deCulture)
            Projekt      = [string]$entry.Projekt
            Beschreibung = [string]$entry.Beschreibung
        }
    }

    if ($normalized.Count -eq 0) {
        [PSCustomObject][ordered]@{
            Id           = ""
            Datum        = ""
            ErfasstAm    = ""
            Stunden      = ""
            Projekt      = ""
            Beschreibung = ""
        } |
            Select-Object $columns |
            Export-Csv -LiteralPath $ActivityCsvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
        return
    }

    $normalized |
        Sort-Object Datum, ErfasstAm |
        Select-Object $columns |
        Export-Csv -LiteralPath $ActivityCsvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
}

function Get-WorkEntriesForDate {
    param(
        [string]$DateText
    )

    return @(Read-WorkEntries | Where-Object { [string]$_.Datum -eq $DateText })
}

function Get-WorkEntrySummaryForDate {
    param(
        [string]$DateText
    )

    $entries = @(Get-WorkEntriesForDate -DateText $DateText)
    $hours = 0.0
    $parts = @()
    $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")

    foreach ($entry in $entries) {
        $entryHours = 0.0

        try {
            $entryHours = [double]$entry.Stunden
        }
        catch {}

        $hours += $entryHours

        $project = ([string]$entry.Projekt).Trim()
        $description = ([string]$entry.Beschreibung).Trim()
        $duration = $entryHours.ToString("N2", $deCulture) + "h"

        if ([string]::IsNullOrWhiteSpace($project)) {
            $parts += "$description ($duration)"
        }
        else {
            $parts += "${project}: $description ($duration)"
        }
    }

    return [PSCustomObject][ordered]@{
        Count = $entries.Count
        Hours = $hours
        Details = ($parts -join " | ")
    }
}

function Convert-ToCsvRow {
    param(
        $Row,
        [string[]]$Columns
    )

    $normalized = [ordered]@{}
    $rowColumns = @()

    if ($null -ne $Row) {
        $rowColumns = $Row.PSObject.Properties.Name
    }

    foreach ($column in $Columns) {
        if ($rowColumns -contains $column) {
            $normalized[$column] = $Row.$column
        }
        else {
            $normalized[$column] = ""
        }
    }

    return [PSCustomObject]$normalized
}

function Update-WorkCsvActivityColumns {
    if (!(Test-Path $CsvPath)) {
        return
    }

    try {
        $rows = @(Import-Csv -LiteralPath $CsvPath -Delimiter ";")

        if ($rows.Count -eq 0) {
            return
        }

        $columns = @()

        foreach ($row in $rows) {
            foreach ($column in @($row.PSObject.Properties.Name)) {
                if ($columns -notcontains $column) {
                    $columns += $column
                }
            }
        }

        foreach ($column in @("Taetigkeiten_Stunden", "Taetigkeiten_Anzahl", "Taetigkeiten_Details")) {
            if ($columns -notcontains $column) {
                $columns += $column
            }
        }

        $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
        $normalized = @()

        foreach ($row in $rows) {
            $csvRow = Convert-ToCsvRow -Row $row -Columns $columns
            $rowDate = Get-WorkDate -DateText ([string]$csvRow.Datum)

            if ($null -ne $rowDate) {
                $summary = Get-WorkEntrySummaryForDate -DateText $rowDate.ToString("yyyy-MM-dd")
                $csvRow.Taetigkeiten_Stunden = ([double]$summary.Hours).ToString("N2", $deCulture)
                $csvRow.Taetigkeiten_Anzahl = [string]$summary.Count
                $csvRow.Taetigkeiten_Details = [string]$summary.Details
            }

            $normalized += $csvRow
        }

        $tmpPath = $CsvPath + ".tmp"
        $backupPath = $CsvPath + ".bak"
        Copy-Item -LiteralPath $CsvPath -Destination $backupPath -Force -ErrorAction SilentlyContinue

        $normalized |
            Export-Csv -LiteralPath $tmpPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

        Move-Item -LiteralPath $tmpPath -Destination $CsvPath -Force
    }
    catch {
        Write-AppLog ("CSV-Taetigkeiten-Migration-Fehler: " + $_.Exception.ToString())
    }
}

function Add-WorkEntryLocal {
    param(
        [double]$Hours,
        [string]$Description,
        [string]$Project
    )

    $now = Get-Date
    $entry = [PSCustomObject][ordered]@{
        Id           = [guid]::NewGuid().ToString()
        Datum        = $now.ToString("yyyy-MM-dd")
        ErfasstAm    = $now.ToString("o")
        Stunden      = [math]::Round($Hours, 2)
        Projekt      = $Project.Trim()
        Beschreibung = $Description.Trim()
    }

    $entries = @(Read-WorkEntries)
    $entries += $entry

    Save-WorkEntries -Entries $entries
    Write-WorkEntriesCsv -Entries $entries
    Update-WorkCsvActivityColumns
}

function Update-WorkEntryLocal {
    param(
        [string]$Id,
        [double]$Hours,
        [string]$Description,
        [string]$Project
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return
    }

    $entries = @(Read-WorkEntries)

    foreach ($entry in $entries) {
        if ([string]$entry.Id -eq $Id) {
            $entry.Stunden = [math]::Round($Hours, 2)
            $entry.Projekt = $Project.Trim()
            $entry.Beschreibung = $Description.Trim()
            break
        }
    }

    Save-WorkEntries -Entries $entries
    Write-WorkEntriesCsv -Entries $entries
    Update-WorkCsvActivityColumns
}

function Remove-WorkEntryLocal {
    param(
        [string]$Id
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return
    }

    $entries = @(Read-WorkEntries | Where-Object { [string]$_.Id -ne $Id })
    Save-WorkEntries -Entries $entries
    Write-WorkEntriesCsv -Entries $entries
    Update-WorkCsvActivityColumns
}

function Get-TodayWorkEntries {
    $today = (Get-Date).ToString("yyyy-MM-dd")
    return @(Read-WorkEntries | Where-Object { [string]$_.Datum -eq $today } | Sort-Object ErfasstAm)
}

function Get-TodayWorkEntrySummary {
    $today = (Get-Date).ToString("yyyy-MM-dd")
    $entries = @(Read-WorkEntries | Where-Object { [string]$_.Datum -eq $today })
    $hours = 0.0

    foreach ($entry in $entries) {
        try {
            $hours += [double]$entry.Stunden
        }
        catch {}
    }

    return [PSCustomObject]@{
        Count = $entries.Count
        Hours = $hours
    }
}

function Recover-DisplayStateFromCsv {
    if (!(Test-Path $CsvPath)) {
        return $null
    }

    try {
        $today = (Get-Date).ToString("yyyy-MM-dd")
        $row = Import-Csv -LiteralPath $CsvPath -Delimiter ";" |
            Where-Object {
                $rowDate = Get-WorkDate -DateText ([string]$_.Datum)
                $null -ne $rowDate -and $rowDate.ToString("yyyy-MM-dd") -eq $today
            } |
            Select-Object -Last 1

        if ($null -eq $row) {
            return $null
        }

        $now = Get-Date

        return [PSCustomObject][ordered]@{
            Date                     = $today
            StartTime                = [string]$row.Start
            EndTime                  = [string]$row.Ende
            StartPending             = $false
            StartCandidateAt         = ""
            GrossSeconds             = Convert-DurationTextToSeconds ([string]$row.Brutto)
            PauseMorningSeconds      = Convert-DurationTextToSeconds ([string]$row.Pause_08_55_09_35)
            PauseNoonSeconds         = Convert-DurationTextToSeconds ([string]$row.Pause_11_55_12_45)
            ManualPauseSeconds       = Convert-DurationTextToSeconds ([string]$row.Pause_Manuell)
            PauseMorningCountedUntil = ""
            PauseNoonCountedUntil    = ""
            ManualPauseActive        = $false
            ManualPauseStartedAt     = ""
            ManualPauseCountedUntil  = ""
            Note                     = [string]$row.Notiz
            LastControlId            = ""
            LastTimestamp            = $now.ToString("o")
        }
    }
    catch {
        Write-AppLog ("CSV-State-Recovery-Fehler: " + $_.Exception.ToString())
        return $null
    }
}

function Read-State {
    if (!(Test-Path $StatePath)) {
        return Recover-DisplayStateFromCsv
    }

    try {
        $item = Get-Item -LiteralPath $StatePath -ErrorAction SilentlyContinue

        if ($null -eq $item -or $item.Length -le 0) {
            return Recover-DisplayStateFromCsv
        }

        return Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
    }
    catch {
        return Recover-DisplayStateFromCsv
    }
}

function Get-StateValue {
    param(
        $State,
        [string]$Name,
        $DefaultValue
    )

    if ($null -eq $State -or $State.PSObject.Properties.Name -notcontains $Name) {
        return $DefaultValue
    }

    return $State.PSObject.Properties[$Name].Value
}

function Get-StateDouble {
    param(
        $State,
        [string]$Name
    )

    try {
        return [double](Get-StateValue -State $State -Name $Name -DefaultValue 0)
    }
    catch {
        return 0
    }
}

function Get-StateBool {
    param(
        $State,
        [string]$Name
    )

    try {
        return [System.Convert]::ToBoolean((Get-StateValue -State $State -Name $Name -DefaultValue $false))
    }
    catch {
        return $false
    }
}

function Get-UiSettings {
    return Read-ArbeitszeitSettings -BaseDir $BaseDir
}

function Get-LiveValues {
    $state = Read-State
    $now = Get-Date
    $settings = Get-UiSettings
    $targetSeconds = [double]$settings.TargetNetHours * 3600

    if ($null -eq $state) {
        return [PSCustomObject]@{
            HasState           = $false
            Date               = $now.ToString("yyyy-MM-dd")
            StartTime          = "--:--:--"
            EndTime            = "--:--:--"
            GrossSeconds       = 0
            AutoPauseSeconds   = 0
            ManualPauseSeconds = 0
            TotalPauseSeconds  = 0
            NetSeconds         = 0
            ManualPauseActive  = $false
            StatusText         = "Noch keine Daten"
            StatusKind         = "Missing"
            UpdatedText        = "-"
            Note               = ""
            TargetSeconds      = $targetSeconds
            RemainingSeconds   = $targetSeconds
            ForecastTime       = "-"
            Progress           = 0
            Settings           = $settings
        }
    }

    $grossSeconds = Get-StateDouble -State $state -Name "GrossSeconds"
    $autoPause = 0.0

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $settings -IncludeDisabled)) {
        $autoPause += Get-StateDouble -State $state -Name $pause.SecondsProperty
    }

    $manualPause = Get-StateDouble -State $state -Name "ManualPauseSeconds"
    $manualActive = Get-StateBool -State $state -Name "ManualPauseActive"
    $startPending = Get-StateBool -State $state -Name "StartPending"
    $lastTimestamp = Get-StateValue -State $state -Name "LastTimestamp" -DefaultValue ""
    $updatedText = "-"
    $statusText = "Arbeitszeit läuft"
    $statusKind = "Work"
    $isStale = $false

    if (-not [string]::IsNullOrWhiteSpace($lastTimestamp)) {
        try {
            $last = [datetime]::Parse($lastTimestamp)
            $updatedText = $last.ToString("HH:mm:ss")
            $liveDelta = ($now - $last).TotalSeconds

            if (-not $startPending -and $liveDelta -ge 0 -and $liveDelta -le 30 -and $state.Date -eq $now.ToString("yyyy-MM-dd")) {
                $grossSeconds += $liveDelta

                if ($manualActive) {
                    $countedUntilText = Get-StateValue -State $state -Name "ManualPauseCountedUntil" -DefaultValue ""

                    if (-not [string]::IsNullOrWhiteSpace($countedUntilText)) {
                        $countedUntil = [datetime]::Parse($countedUntilText)
                        $manualDelta = ($now - $countedUntil).TotalSeconds

                        if ($manualDelta -ge 0 -and $manualDelta -le 30) {
                            $manualPause += $manualDelta
                        }
                    }
                }
            }
            elseif ($liveDelta -gt 30) {
                $isStale = $true
            }
        }
        catch {
            $isStale = $true
        }
    }

    if ($isStale) {
        $statusText = "Tracker nicht aktuell"
        $statusKind = "Stale"
    }
    elseif ($startPending) {
        $statusText = "Wartet auf Aktivität"
        $statusKind = "Missing"
    }
    elseif ($manualActive) {
        $statusText = "Manuelle Pause"
        $statusKind = "Pause"
    }

    $totalPause = $autoPause + $manualPause
    $netSeconds = $grossSeconds - $totalPause

    if ($netSeconds -lt 0) {
        $netSeconds = 0
    }

    $remainingSeconds = $targetSeconds - $netSeconds

    if ($remainingSeconds -lt 0) {
        $remainingSeconds = 0
    }

    $progress = 0

    if ($targetSeconds -gt 0) {
        $progress = $netSeconds / $targetSeconds
    }

    $forecastTime = "Jetzt"

    if ($remainingSeconds -gt 0) {
        $forecastTime = $now.AddSeconds($remainingSeconds).ToString("HH:mm")
    }

    return [PSCustomObject]@{
        HasState           = $true
        Date               = (Get-StateValue -State $state -Name "Date" -DefaultValue $now.ToString("yyyy-MM-dd"))
        StartTime          = if ($startPending) { "--:--:--" } else { (Get-StateValue -State $state -Name "StartTime" -DefaultValue "--:--:--") }
        EndTime            = if ($startPending) { "--:--:--" } else { (Get-StateValue -State $state -Name "EndTime" -DefaultValue "--:--:--") }
        GrossSeconds       = $grossSeconds
        AutoPauseSeconds   = $autoPause
        ManualPauseSeconds = $manualPause
        TotalPauseSeconds  = $totalPause
        NetSeconds         = $netSeconds
        ManualPauseActive  = $manualActive
        StartPending       = $startPending
        StatusText         = $statusText
        StatusKind         = $statusKind
        UpdatedText        = $updatedText
        Note               = (Get-StateValue -State $state -Name "Note" -DefaultValue "")
        TargetSeconds      = $targetSeconds
        RemainingSeconds   = $remainingSeconds
        ForecastTime       = $forecastTime
        Progress           = $progress
        Settings           = $settings
    }
}

function Write-ControlCommand {
    param(
        [string]$Action,
        $Values = $null
    )

    $id = [guid]::NewGuid().ToString()
    $command = [PSCustomObject][ordered]@{
        Id        = $id
        Action    = $Action
        CreatedAt = (Get-Date).ToString("o")
        Values    = $Values
    }

    $tmpPath = Join-Path $BaseDir ("control.{0}.tmp" -f $id)

    $command |
        ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $tmpPath -Encoding UTF8

    Move-Item -LiteralPath $tmpPath -Destination $ControlPath -Force
}

function Show-Info {
    param(
        [string]$Text
    )

    [System.Windows.MessageBox]::Show($Text, "Arbeitszeit", "OK", "Information") | Out-Null
}

function Show-Warning {
    param(
        [string]$Text
    )

    [System.Windows.MessageBox]::Show($Text, "Arbeitszeit", "OK", "Warning") | Out-Null
}

function Read-WorkCsvRows {
    if (!(Test-Path $CsvPath)) {
        return @()
    }

    try {
        return @(Import-Csv -LiteralPath $CsvPath -Delimiter ";")
    }
    catch {
        return @()
    }
}

function Get-WeekReportData {
    param(
        $Values = (Get-LiveValues)
    )

    $settings = $Values.Settings
    $today = Get-Date
    $weekStart = Get-WeekStart -Date $today
    $weekEnd = $weekStart.AddDays(6)
    $dayMap = @{}

    foreach ($row in (Read-WorkCsvRows)) {
        $date = Get-WorkDate -DateText ([string]$row.Datum)

        if ($null -eq $date) {
            continue
        }

        if ($date -lt $weekStart -or $date -gt $weekEnd) {
            continue
        }

        $key = $date.ToString("yyyy-MM-dd")
        $activitySummary = Get-WorkEntrySummaryForDate -DateText $key
        $dayMap[$key] = [PSCustomObject][ordered]@{
            Date          = $date
            NetSeconds    = Convert-DurationTextToSeconds ([string]$row.Netto)
            GrossSeconds  = Convert-DurationTextToSeconds ([string]$row.Brutto)
            PauseSeconds  = Convert-DurationTextToSeconds ([string]$row.Pause_Gesamt)
            ActivityHours = [double]$activitySummary.Hours
            ActivityCount = [int]$activitySummary.Count
            ActivityText  = [string]$activitySummary.Details
            Status        = [string]$row.Status
        }
    }

    if ($Values.HasState) {
        $liveDate = [datetime]::ParseExact(
            [string]$Values.Date,
            "yyyy-MM-dd",
            [System.Globalization.CultureInfo]::InvariantCulture
        )

        if ($liveDate -ge $weekStart -and $liveDate -le $weekEnd) {
            $activitySummary = Get-WorkEntrySummaryForDate -DateText $liveDate.ToString("yyyy-MM-dd")
            $dayMap[$liveDate.ToString("yyyy-MM-dd")] = [PSCustomObject][ordered]@{
                Date          = $liveDate
                NetSeconds    = [double]$Values.NetSeconds
                GrossSeconds  = [double]$Values.GrossSeconds
                PauseSeconds  = [double]$Values.TotalPauseSeconds
                ActivityHours = [double]$activitySummary.Hours
                ActivityCount = [int]$activitySummary.Count
                ActivityText  = [string]$activitySummary.Details
                Status        = if ($Values.ManualPauseActive) { "Manuelle Pause" } else { "Arbeit" }
            }
        }
    }

    $days = @()

    for ($i = 0; $i -lt 7; $i++) {
        $date = $weekStart.AddDays($i)
        $key = $date.ToString("yyyy-MM-dd")

        if ($dayMap.ContainsKey($key)) {
            $days += $dayMap[$key]
        }
        else {
            $days += [PSCustomObject][ordered]@{
                Date          = $date
                NetSeconds    = 0.0
                GrossSeconds  = 0.0
                PauseSeconds  = 0.0
                ActivityHours = [double](Get-WorkEntrySummaryForDate -DateText $key).Hours
                ActivityCount = [int](Get-WorkEntrySummaryForDate -DateText $key).Count
                ActivityText  = [string](Get-WorkEntrySummaryForDate -DateText $key).Details
                Status        = ""
            }
        }
    }

    $totalNet = [double](@($days | Measure-Object -Property NetSeconds -Sum).Sum)
    $totalGross = [double](@($days | Measure-Object -Property GrossSeconds -Sum).Sum)
    $totalPause = [double](@($days | Measure-Object -Property PauseSeconds -Sum).Sum)
    $totalActivityHours = [double](@($days | Measure-Object -Property ActivityHours -Sum).Sum)
    $workedDays = @($days | Where-Object { $_.NetSeconds -gt 0 })
    $average = 0.0

    if ($workedDays.Count -gt 0) {
        $average = $totalNet / $workedDays.Count
    }

    $pauseRatio = 0.0

    if ($totalGross -gt 0) {
        $pauseRatio = ($totalPause / $totalGross) * 100
    }

    $bestDay = $workedDays | Sort-Object NetSeconds -Descending | Select-Object -First 1
    $weekTargetSeconds = [double]$settings.WeekTargetHours * 3600

    return [PSCustomObject][ordered]@{
        WeekStart         = $weekStart
        WeekEnd           = $weekEnd
        Days              = $days
        TotalNetSeconds   = $totalNet
        TotalGrossSeconds = $totalGross
        TotalPauseSeconds = $totalPause
        TotalActivityHours = $totalActivityHours
        WeekTargetSeconds = $weekTargetSeconds
        BalanceSeconds    = $totalNet - $weekTargetSeconds
        AverageSeconds    = $average
        PauseRatio        = $pauseRatio
        BestDay           = $bestDay
        Values            = $Values
    }
}

function ConvertTo-ReportHtml {
    param(
        $Data
    )

    $rows = foreach ($day in @($Data.Days)) {
        $name = Get-DayNameShort -Date $day.Date
        $date = $day.Date.ToString("dd.MM.yyyy")
        $activity = "{0:N2} h" -f [double]$day.ActivityHours
        $activityText = [System.Net.WebUtility]::HtmlEncode([string]$day.ActivityText)
        "<tr><td>$name</td><td>$date</td><td>$(Format-CompactDuration $day.NetSeconds)</td><td>$activity</td><td>$(Format-CompactDuration $day.PauseSeconds)</td><td>$([System.Net.WebUtility]::HtmlEncode($day.Status))</td><td>$activityText</td></tr>"
    }

    $best = "-"

    if ($null -ne $Data.BestDay) {
        $best = "{0}, {1}" -f (Get-DayNameShort -Date $Data.BestDay.Date), (Format-CompactDuration $Data.BestDay.NetSeconds)
    }

    $balanceClass = if ($Data.BalanceSeconds -ge 0) { "good" } else { "bad" }

    return @"
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<title>Arbeitszeit Wochenbericht</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 34px; color: #1c1c1e; }
h1 { margin: 0 0 4px 0; font-size: 28px; }
.sub { color: #6e6e73; margin-bottom: 26px; }
.cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 24px; }
.card { border: 1px solid #e5e5ea; border-radius: 14px; padding: 14px; }
.label { color: #6e6e73; font-size: 12px; text-transform: uppercase; font-weight: 700; }
.value { font-size: 22px; font-weight: 700; margin-top: 8px; }
.good { color: #248a3d; }
.bad { color: #d70015; }
table { width: 100%; border-collapse: collapse; margin-top: 14px; }
th, td { text-align: left; padding: 10px 8px; border-bottom: 1px solid #e5e5ea; }
th { color: #6e6e73; font-size: 12px; text-transform: uppercase; }
</style>
</head>
<body>
<h1>Wochenbericht Arbeitszeit</h1>
<div class="sub">$($Data.WeekStart.ToString("dd.MM.yyyy")) bis $($Data.WeekEnd.ToString("dd.MM.yyyy"))</div>
<div class="cards">
  <div class="card"><div class="label">Netto</div><div class="value">$(Format-CompactDuration $Data.TotalNetSeconds)</div></div>
  <div class="card"><div class="label">Ziel</div><div class="value">$(Format-CompactDuration $Data.WeekTargetSeconds)</div></div>
  <div class="card"><div class="label">Saldo</div><div class="value $balanceClass">$(Format-CompactDuration $Data.BalanceSeconds)</div></div>
  <div class="card"><div class="label">Tätigkeiten</div><div class="value">$("{0:N2}" -f $Data.TotalActivityHours) h</div></div>
</div>
<div class="cards">
  <div class="card"><div class="label">Produktivster Tag</div><div class="value">$best</div></div>
  <div class="card"><div class="label">Ø Arbeitstag</div><div class="value">$(Format-CompactDuration $Data.AverageSeconds)</div></div>
  <div class="card"><div class="label">Pause gesamt</div><div class="value">$(Format-CompactDuration $Data.TotalPauseSeconds)</div></div>
  <div class="card"><div class="label">Pausenquote</div><div class="value">$("{0:N1}" -f $Data.PauseRatio)%</div></div>
</div>
<table>
<thead><tr><th>Tag</th><th>Datum</th><th>Netto</th><th>Tätigkeiten</th><th>Pause</th><th>Status</th><th>Details</th></tr></thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>
</body>
</html>
"@
}

function Get-BrowserPath {
    $candidates = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Export-WeekReport {
    param(
        $Owner
    )

    $data = Get-WeekReportData
    $reportDir = Join-Path $BaseDir "reports"

    if (!(Test-Path $reportDir)) {
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    }

    $baseName = "Wochenbericht_{0}" -f $data.WeekStart.ToString("yyyy-MM-dd")
    $htmlPath = Join-Path $reportDir "$baseName.html"
    $pdfPath = Join-Path $reportDir "$baseName.pdf"

    ConvertTo-ReportHtml -Data $data |
        Set-Content -LiteralPath $htmlPath -Encoding UTF8

    $browser = Get-BrowserPath

    if ($browser) {
        $htmlUri = ([Uri](Resolve-Path $htmlPath).Path).AbsoluteUri
        $arguments = @(
            "--headless",
            "--disable-gpu",
            "--print-to-pdf=$pdfPath",
            $htmlUri
        )

        Start-Process -FilePath $browser -ArgumentList $arguments -Wait -WindowStyle Hidden
    }

    if (Test-Path $pdfPath) {
        Start-Process -FilePath $pdfPath
        Show-Info "Wochenbericht wurde als PDF erstellt."
    }
    else {
        Start-Process -FilePath $htmlPath
        Show-Info "PDF konnte nicht automatisch erstellt werden. Der HTML-Wochenbericht wurde geöffnet."
    }
}

function Get-ThemePalette {
    param(
        [string]$Theme
    )

    if ($Theme -eq "Dark") {
        return [PSCustomObject][ordered]@{
            Window    = "#1C1C1E"
            Card      = "#2C2C2E"
            Soft      = "#3A3A3C"
            Border    = "#48484A"
            Primary   = "#F5F5F7"
            Secondary = "#C7C7CC"
            Muted     = "#8E8E93"
            Accent    = "#0A84FF"
        }
    }

    return [PSCustomObject][ordered]@{
        Window    = "#F5F5F7"
        Card      = "#FFFFFF"
        Soft      = "#F2F2F7"
        Border    = "#ECECF0"
        Primary   = "#1C1C1E"
        Secondary = "#6E6E73"
        Muted     = "#8E8E93"
        Accent    = "#007AFF"
    }
}

function New-Brush {
    param(
        [string]$Color
    )

    return [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
}

function Get-BrushText {
    param($Brush)

    if ($Brush -is [System.Windows.Media.SolidColorBrush]) {
        return $Brush.Color.ToString()
    }

    return ""
}

function Apply-ThemeRecursive {
    param(
        $Element,
        $Palette
    )

    if ($null -eq $Element) {
        return
    }

    if ($Element -is [System.Windows.Controls.Border] -and $Element.Name -ne "StatusPill") {
        $background = Get-BrushText $Element.Background

        if ($background -in @("#FFFFFFFF", "#FFF5F5F7", "#FFF2F2F7", "#FF2C2C2E", "#FF3A3A3C")) {
            $Element.Background = New-Brush $Palette.Card
        }

        $Element.BorderBrush = New-Brush $Palette.Border
    }

    if ($Element -is [System.Windows.Controls.TextBlock] -and $Element.Name -ne "StatusText") {
        $foreground = Get-BrushText $Element.Foreground

        if ($foreground -in @("#FF1C1C1E", "#FFF5F5F7")) {
            $Element.Foreground = New-Brush $Palette.Primary
        }
        elseif ($foreground -in @("#FF6E6E73", "#FF8E8E93", "#FFC7C7CC")) {
            $Element.Foreground = New-Brush $Palette.Secondary
        }
    }

    if ($Element -is [System.Windows.Controls.CheckBox]) {
        $Element.Foreground = New-Brush $Palette.Secondary
    }

    if ($Element -is [System.Windows.Controls.TextBox]) {
        $Element.Background = New-Brush $Palette.Soft
        $Element.Foreground = New-Brush $Palette.Primary
        $Element.BorderBrush = New-Brush $Palette.Border
        $Element.CaretBrush = New-Brush $Palette.Primary
    }

    if ($Element -is [System.Windows.Controls.ListView]) {
        $Element.Background = New-Brush $Palette.Soft
        $Element.Foreground = New-Brush $Palette.Primary
        $Element.BorderBrush = New-Brush $Palette.Border
    }

    $count = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($Element)

    for ($i = 0; $i -lt $count; $i++) {
        Apply-ThemeRecursive -Element ([System.Windows.Media.VisualTreeHelper]::GetChild($Element, $i)) -Palette $Palette
    }
}

function Apply-MainTheme {
    param(
        $Settings
    )

    if ($null -eq $mainWindow) {
        return
    }

    $palette = Get-ThemePalette -Theme ([string]$Settings.Theme)
    $mainWindow.Background = New-Brush $palette.Window

    if ($null -ne $rootGrid) {
        Apply-ThemeRecursive -Element $rootGrid -Palette $palette
    }

    foreach ($element in @($netText, $startText, $endText, $grossText, $autoPauseText, $manualPauseText, $updatedText)) {
        if ($null -ne $element) {
            $element.Foreground = New-Brush $palette.Primary
        }
    }

    foreach ($element in @($remainingText, $forecastText, $activityStatusText)) {
        if ($null -ne $element) {
            $element.Foreground = New-Brush $palette.Secondary
        }
    }

    if ($null -ne $targetProgress) {
        $targetProgress.Background = New-Brush $palette.Border
        $targetProgress.Foreground = New-Brush $palette.Accent
    }

    if ($null -ne $topMostBox) {
        $topMostBox.Foreground = New-Brush $palette.Secondary
    }

    if ($null -ne $setupButton) {
        $setupButton.Background = New-Brush $palette.Soft
        $setupButton.Foreground = New-Brush $palette.Primary
    }

    if ($null -ne $themeButton) {
        $themeButton.Content = if ($Settings.Theme -eq "Dark") { "Light" } else { "Dark" }
    }
}

$script:NotificationIcons = New-Object System.Collections.ArrayList
$script:NotificationCleanupTimers = New-Object System.Collections.ArrayList

function Read-NotificationState {
    if (!(Test-Path $NotificationPath)) {
        return [PSCustomObject]@{ TargetReachedDate = "" }
    }

    try {
        return Get-Content -LiteralPath $NotificationPath -Raw | ConvertFrom-Json
    }
    catch {
        return [PSCustomObject]@{ TargetReachedDate = "" }
    }
}

function Save-NotificationState {
    param($State)

    $State |
        ConvertTo-Json -Depth 3 |
        Set-Content -LiteralPath $NotificationPath -Encoding UTF8
}

function Show-TargetReachedNotification {
    param(
        $Values
    )

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon

    if (Test-Path $IconPath) {
        $notifyIcon.Icon = New-Object System.Drawing.Icon -ArgumentList $IconPath
    }
    else {
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Information
    }

    $notifyIcon.Visible = $true
    $notifyIcon.BalloonTipTitle = "Arbeitszeit"
    $notifyIcon.BalloonTipText = "Tagesziel erreicht: $(Format-CompactDuration $Values.NetSeconds) netto."
    $notifyIcon.ShowBalloonTip(6000)
    $script:NotificationIcons.Add($notifyIcon) | Out-Null

    $cleanupTimer = New-Object System.Windows.Threading.DispatcherTimer
    $cleanupTimer.Interval = [TimeSpan]::FromSeconds(8)
    $cleanupTimer.Tag = $notifyIcon
    $cleanupTimer.Add_Tick({
        param($sender, $eventArgs)

        try {
            if ($null -ne $sender) {
                $sender.Stop()
                $icon = $sender.Tag

                if ($null -ne $icon) {
                    $icon.Visible = $false
                    $icon.Dispose()
                    [void]$script:NotificationIcons.Remove($icon)
                }

                [void]$script:NotificationCleanupTimers.Remove($sender)
            }
        }
        catch {
            Write-AppLog ("Benachrichtigungs-Cleanup-Fehler: " + $_.Exception.ToString())
        }
    })
    [void]$script:NotificationCleanupTimers.Add($cleanupTimer)
    $cleanupTimer.Start()
}

function Test-TargetNotification {
    param(
        $Values
    )

    if (-not $Values.HasState -or -not [bool]$Values.Settings.NotifyTargetReached) {
        return
    }

    if ($Values.TargetSeconds -le 0 -or $Values.NetSeconds -lt $Values.TargetSeconds) {
        return
    }

    $state = Read-NotificationState
    $date = [string]$Values.Date

    if ([string]$state.TargetReachedDate -eq $date) {
        return
    }

    Show-TargetReachedNotification -Values $Values
    Save-NotificationState -State ([PSCustomObject]@{ TargetReachedDate = $date })
}

$dialogStyles = @"
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
                    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Style x:Key="DialogButton" TargetType="Button">
        <Setter Property="Height" Value="36"/>
        <Setter Property="MinWidth" Value="98"/>
        <Setter Property="Padding" Value="18,0"/>
        <Setter Property="FontWeight" Value="SemiBold"/>
        <Setter Property="Foreground" Value="White"/>
        <Setter Property="Background" Value="#007AFF"/>
        <Setter Property="BorderThickness" Value="0"/>
        <Setter Property="Template">
            <Setter.Value>
                <ControlTemplate TargetType="Button">
                    <Border Background="{TemplateBinding Background}" CornerRadius="18">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                </ControlTemplate>
            </Setter.Value>
        </Setter>
    </Style>
    <Style TargetType="TextBox">
        <Setter Property="Height" Value="30"/>
        <Setter Property="Padding" Value="9,3"/>
        <Setter Property="BorderBrush" Value="#D1D1D6"/>
        <Setter Property="BorderThickness" Value="1"/>
        <Setter Property="Background" Value="White"/>
        <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="TextBlock">
        <Setter Property="FontFamily" Value="Segoe UI"/>
    </Style>
</ResourceDictionary>
"@

function Add-DialogStyles {
    param(
        $Window
    )

    $resources = Convert-FromXaml $dialogStyles
    $Window.Resources.MergedDictionaries.Add($resources) | Out-Null
}

function Apply-DialogButtonStyles {
    param(
        $Window
    )

    $style = $Window.Resources["DialogButton"]

    foreach ($name in @("SaveButton", "CancelButton")) {
        $button = $Window.FindName($name)

        if ($null -ne $button) {
            $button.Style = $style
        }
    }
}

function Open-CorrectionWindow {
    param(
        $Owner
    )

    $values = Get-LiveValues

    if (-not $values.HasState) {
        Show-Info "Es gibt noch keinen Tagesstand. Starte den Tracker einmal oder warte kurz."
        return
    }

    $state = Read-State
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Heute korrigieren"
        Width="430"
        Height="430"
        ResizeMode="NoResize"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F7"
        FontFamily="Segoe UI">
    <Border Margin="16" Padding="20" Background="White" CornerRadius="26">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="18"/>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="28"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>

            <StackPanel>
                <TextBlock Text="Heute korrigieren" FontSize="20" FontWeight="SemiBold" Foreground="#1C1C1E"/>
                <TextBlock Text="Ändert nur den aktuellen Tagesstand." Margin="0,4,0,0" Foreground="#6E6E73"/>
            </StackPanel>

            <Grid Grid.Row="2">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="170"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="40"/>
                    <RowDefinition Height="40"/>
                    <RowDefinition Height="40"/>
                    <RowDefinition Height="40"/>
                    <RowDefinition Height="40"/>
                </Grid.RowDefinitions>

                <TextBlock Text="Startzeit" Grid.Row="0" Foreground="#6E6E73" VerticalAlignment="Center"/>
                <TextBox x:Name="StartBox" Grid.Row="0" Grid.Column="1"/>

                <TextBlock Text="Pause morgens (min)" Grid.Row="1" Foreground="#6E6E73" VerticalAlignment="Center"/>
                <TextBox x:Name="MorningBox" Grid.Row="1" Grid.Column="1"/>

                <TextBlock Text="Pause mittags (min)" Grid.Row="2" Foreground="#6E6E73" VerticalAlignment="Center"/>
                <TextBox x:Name="NoonBox" Grid.Row="2" Grid.Column="1"/>

                <TextBlock Text="Pause manuell (min)" Grid.Row="3" Foreground="#6E6E73" VerticalAlignment="Center"/>
                <TextBox x:Name="ManualBox" Grid.Row="3" Grid.Column="1"/>

                <TextBlock Text="Notiz" Grid.Row="4" Foreground="#6E6E73" VerticalAlignment="Center"/>
                <TextBox x:Name="NoteBox" Grid.Row="4" Grid.Column="1"/>
            </Grid>

            <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="CancelButton" Content="Abbrechen" Background="#8E8E93" Margin="0,0,10,0"/>
                <Button x:Name="SaveButton" Content="Speichern"/>
            </StackPanel>
        </Grid>
    </Border>
</Window>
"@

    $dialog = Convert-FromXaml $xaml
    Add-DialogStyles $dialog
    Apply-DialogButtonStyles $dialog
    $dialog.Owner = $Owner

    if (Test-Path $IconPath) {
        try {
            $dialog.Icon = New-Object System.Windows.Media.Imaging.BitmapImage([Uri]$IconPath)
        }
        catch {}
    }

    $dialog.FindName("StartBox").Text = $values.StartTime
    $dialog.FindName("MorningBox").Text = ([math]::Round((Get-StateDouble -State $state -Name "PauseMorningSeconds") / 60, 1)).ToString("0.0").Replace(".", ",")
    $dialog.FindName("NoonBox").Text = ([math]::Round((Get-StateDouble -State $state -Name "PauseNoonSeconds") / 60, 1)).ToString("0.0").Replace(".", ",")
    $dialog.FindName("ManualBox").Text = ([math]::Round($values.ManualPauseSeconds / 60, 1)).ToString("0.0").Replace(".", ",")
    $dialog.FindName("NoteBox").Text = [string]$values.Note

    $dialog.FindName("CancelButton").Add_Click({
        $dialog.Close()
    })

    $dialog.FindName("SaveButton").Add_Click({
        $startTime = Convert-ToTimeText $dialog.FindName("StartBox").Text

        if ([string]::IsNullOrWhiteSpace($startTime)) {
            Show-Warning "Bitte die Startzeit als HH:mm oder HH:mm:ss eingeben."
            return
        }

        $commandValues = [PSCustomObject][ordered]@{
            StartTime           = $startTime
            PauseMorningSeconds = (Convert-ToDouble $dialog.FindName("MorningBox").Text) * 60
            PauseNoonSeconds    = (Convert-ToDouble $dialog.FindName("NoonBox").Text) * 60
            ManualPauseSeconds  = (Convert-ToDouble $dialog.FindName("ManualBox").Text) * 60
            Note                = $dialog.FindName("NoteBox").Text
        }

        Write-ControlCommand -Action "UpdateToday" -Values $commandValues
        $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
}

function Open-SettingsWindow {
    param(
        $Owner
    )

    $settings = Get-UiSettings

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Setup"
        Width="510"
        Height="690"
        MinHeight="610"
        ResizeMode="CanResize"
        WindowStartupLocation="CenterOwner"
        Background="#F5F5F7"
        FontFamily="Segoe UI">
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="16">
            <Border Padding="20" Background="White" CornerRadius="26">
                <StackPanel>
                    <TextBlock Text="Setup" FontSize="22" FontWeight="SemiBold" Foreground="#1C1C1E"/>
                    <TextBlock Text="Alle Regeln für Erfassung und Anzeige." Margin="0,4,0,18" Foreground="#6E6E73"/>

                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="210"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="42"/>
                            <RowDefinition Height="42"/>
                            <RowDefinition Height="42"/>
                            <RowDefinition Height="42"/>
                            <RowDefinition Height="42"/>
                            <RowDefinition Height="42"/>
                        </Grid.RowDefinitions>

                        <TextBlock Text="Zielzeit netto (h)" Grid.Row="0" Foreground="#6E6E73" VerticalAlignment="Center"/>
                        <TextBox x:Name="TargetBox" Grid.Row="0" Grid.Column="1"/>

                        <TextBlock Text="Wochenziel netto (h)" Grid.Row="1" Foreground="#6E6E73" VerticalAlignment="Center"/>
                        <TextBox x:Name="WeekTargetBox" Grid.Row="1" Grid.Column="1"/>

                        <TextBlock Text="Messintervall (Sek.)" Grid.Row="2" Foreground="#6E6E73" VerticalAlignment="Center"/>
                        <TextBox x:Name="IntervalBox" Grid.Row="2" Grid.Column="1"/>

                        <TextBlock Text="Idle-Schwelle (Sek.)" Grid.Row="3" Foreground="#6E6E73" VerticalAlignment="Center"/>
                        <TextBox x:Name="IdleBox" Grid.Row="3" Grid.Column="1"/>

                        <CheckBox x:Name="NotifyTargetBox" Grid.Row="4" Grid.ColumnSpan="2" Content="Benachrichtigung bei Tagesziel" VerticalAlignment="Center"/>

                        <CheckBox x:Name="TopMostBox" Grid.Row="5" Grid.ColumnSpan="2" Content="Anzeige immer oben halten" VerticalAlignment="Center"/>
                    </Grid>
                </StackPanel>
            </Border>

            <Border Margin="0,14,0,0" Padding="20" Background="White" CornerRadius="26">
                <StackPanel>
                    <TextBlock Text="Pausenfenster" FontSize="18" FontWeight="SemiBold" Foreground="#1C1C1E"/>
                    <TextBlock Text="Nur in aktiven Fenstern wird Inaktivität als Pause gezählt." Margin="0,4,0,18" Foreground="#6E6E73"/>

                    <Grid Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="56"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="86"/>
                            <ColumnDefinition Width="86"/>
                            <ColumnDefinition Width="44"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Text="Aktiv" Grid.Column="0" Foreground="#6E6E73"/>
                        <TextBlock Text="Name" Grid.Column="1" Foreground="#6E6E73"/>
                        <TextBlock Text="Start" Grid.Column="2" Foreground="#6E6E73"/>
                        <TextBlock Text="Ende" Grid.Column="3" Foreground="#6E6E73"/>
                    </Grid>

                    <StackPanel x:Name="PauseListPanel"/>

                    <Button x:Name="AddPauseButton" Content="+ Pause" Margin="0,12,0,0" HorizontalAlignment="Left"/>
                </StackPanel>
            </Border>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                <Button x:Name="CancelButton" Content="Abbrechen" Background="#8E8E93" Margin="0,0,10,0"/>
                <Button x:Name="SaveButton" Content="Speichern"/>
            </StackPanel>
        </StackPanel>
    </ScrollViewer>
</Window>
"@

    $dialog = Convert-FromXaml $xaml
    Add-DialogStyles $dialog
    Apply-DialogButtonStyles $dialog
    $dialog.Owner = $Owner

    if (Test-Path $IconPath) {
        try {
            $dialog.Icon = New-Object System.Windows.Media.Imaging.BitmapImage([Uri]$IconPath)
        }
        catch {}
    }

    $pauseRows = New-Object System.Collections.ArrayList
    $pauseListPanel = $dialog.FindName("PauseListPanel")
    $buttonStyle = $dialog.Resources["DialogButton"]

    function Add-PauseRow {
        param(
            $Pause,
            [bool]$CanRemove
        )

        $grid = New-Object System.Windows.Controls.Grid
        $grid.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 0, 10

        $columns = @(
            (New-Object System.Windows.Controls.ColumnDefinition),
            (New-Object System.Windows.Controls.ColumnDefinition),
            (New-Object System.Windows.Controls.ColumnDefinition),
            (New-Object System.Windows.Controls.ColumnDefinition),
            (New-Object System.Windows.Controls.ColumnDefinition)
        )

        $columns[0].Width = New-Object System.Windows.GridLength -ArgumentList 56
        $columns[1].Width = New-Object System.Windows.GridLength -ArgumentList 1, ([System.Windows.GridUnitType]::Star)
        $columns[2].Width = New-Object System.Windows.GridLength -ArgumentList 86
        $columns[3].Width = New-Object System.Windows.GridLength -ArgumentList 86
        $columns[4].Width = New-Object System.Windows.GridLength -ArgumentList 44

        foreach ($column in $columns) {
            $grid.ColumnDefinitions.Add($column) | Out-Null
        }

        $enabledBox = New-Object System.Windows.Controls.CheckBox
        $enabledBox.IsChecked = [bool]$Pause.Enabled
        $enabledBox.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        [System.Windows.Controls.Grid]::SetColumn($enabledBox, 0)
        $grid.Children.Add($enabledBox) | Out-Null

        $labelBox = New-Object System.Windows.Controls.TextBox
        $labelBox.Text = [string]$Pause.Label
        $labelBox.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 10, 0
        [System.Windows.Controls.Grid]::SetColumn($labelBox, 1)
        $grid.Children.Add($labelBox) | Out-Null

        $startBox = New-Object System.Windows.Controls.TextBox
        $startBox.Text = [string]$Pause.Start
        $startBox.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 10, 0
        [System.Windows.Controls.Grid]::SetColumn($startBox, 2)
        $grid.Children.Add($startBox) | Out-Null

        $endBox = New-Object System.Windows.Controls.TextBox
        $endBox.Text = [string]$Pause.End
        $endBox.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 8, 0
        [System.Windows.Controls.Grid]::SetColumn($endBox, 3)
        $grid.Children.Add($endBox) | Out-Null

        $removeButton = New-Object System.Windows.Controls.Button
        $removeButton.Content = "X"
        $removeButton.Height = 30
        $removeButton.MinWidth = 36
        $removeButton.Background = New-Object System.Windows.Media.SolidColorBrush -ArgumentList ([System.Windows.Media.Color]::FromRgb(142, 142, 147))
        $removeButton.Foreground = [System.Windows.Media.Brushes]::White

        if ($null -ne $buttonStyle) {
            $removeButton.Style = $buttonStyle
        }

        $removeButton.Visibility = if ($CanRemove) { [System.Windows.Visibility]::Visible } else { [System.Windows.Visibility]::Hidden }
        [System.Windows.Controls.Grid]::SetColumn($removeButton, 4)
        $grid.Children.Add($removeButton) | Out-Null

        $rowInfo = [PSCustomObject]@{
            Grid         = $grid
            Key          = [string]$Pause.Key
            EnabledBox   = $enabledBox
            LabelBox     = $labelBox
            StartBox     = $startBox
            EndBox       = $endBox
            RemoveButton = $removeButton
        }

        $removeButton.Add_Click({
            $pauseListPanel.Children.Remove($rowInfo.Grid)
            $pauseRows.Remove($rowInfo)
        })

        $pauseRows.Add($rowInfo) | Out-Null
        $pauseListPanel.Children.Add($grid) | Out-Null
    }

    foreach ($pause in @($settings.PauseWindows)) {
        Add-PauseRow -Pause $pause -CanRemove ($pause.Key -ne "Morning" -and $pause.Key -ne "Noon")
    }

    $addPauseButton = $dialog.FindName("AddPauseButton")
    $addPauseButton.Style = $buttonStyle
    $addPauseButton.Add_Click({
        $newPause = [PSCustomObject][ordered]@{
            Key     = ""
            Label   = "Neue Pause"
            Start   = "12:00"
            End     = "12:15"
            Enabled = $true
        }

        Add-PauseRow -Pause $newPause -CanRemove $true
    })

    $dialog.FindName("TargetBox").Text = ([double]$settings.TargetNetHours).ToString("0.##").Replace(".", ",")
    $dialog.FindName("WeekTargetBox").Text = ([double]$settings.WeekTargetHours).ToString("0.##").Replace(".", ",")
    $dialog.FindName("IntervalBox").Text = [string]$settings.IntervalSeconds
    $dialog.FindName("IdleBox").Text = [string]$settings.IdleThresholdSeconds
    $dialog.FindName("NotifyTargetBox").IsChecked = [bool]$settings.NotifyTargetReached
    $dialog.FindName("TopMostBox").IsChecked = [bool]$settings.AlwaysOnTop

    $dialog.FindName("CancelButton").Add_Click({
        $dialog.Close()
    })

    $dialog.FindName("SaveButton").Add_Click({
        $targetHours = Convert-ToDouble $dialog.FindName("TargetBox").Text 8
        $weekTargetHours = Convert-ToDouble $dialog.FindName("WeekTargetBox").Text 40
        $intervalSeconds = [int](Convert-ToDouble $dialog.FindName("IntervalBox").Text 5)
        $idleSeconds = [int](Convert-ToDouble $dialog.FindName("IdleBox").Text 20)
        $pauseWindows = @()
        $usedKeys = @{}

        foreach ($row in @($pauseRows)) {
            $label = [string]$row.LabelBox.Text
            $start = Convert-ArbeitszeitTimeText $row.StartBox.Text
            $end = Convert-ArbeitszeitTimeText $row.EndBox.Text

            if ([string]::IsNullOrWhiteSpace($label)) {
                Show-Warning "Bitte jeder Pause einen Namen geben."
                return
            }

            if ([string]::IsNullOrWhiteSpace($start) -or [string]::IsNullOrWhiteSpace($end)) {
                Show-Warning "Bitte alle Zeiten im Format HH:mm eingeben."
                return
            }

            $key = Convert-ArbeitszeitPauseKey -Key ([string]$row.Key) -Fallback $label
            $baseKey = $key
            $counter = 2

            while ($usedKeys.ContainsKey($key)) {
                $key = "$baseKey$counter"
                $counter++
            }

            $usedKeys[$key] = $true
            $pauseWindows += New-ArbeitszeitPauseWindow `
                -Key $key `
                -Label $label `
                -Start $start `
                -End $end `
                -Enabled ([bool]$row.EnabledBox.IsChecked)
        }

        $newSettings = [PSCustomObject][ordered]@{
            IntervalSeconds     = $intervalSeconds
            IdleThresholdSeconds = $idleSeconds
            TargetNetHours      = $targetHours
            WeekTargetHours     = $weekTargetHours
            AlwaysOnTop         = [bool]$dialog.FindName("TopMostBox").IsChecked
            NotifyTargetReached = [bool]$dialog.FindName("NotifyTargetBox").IsChecked
            Theme               = [string]$settings.Theme
            PauseWindows        = $pauseWindows
        }

        Save-ArbeitszeitSettings -BaseDir $BaseDir -Settings $newSettings
        $script:SuppressTopMostSave = $true
        $mainWindow.FindName("TopMostBox").IsChecked = [bool]$newSettings.AlwaysOnTop
        $script:SuppressTopMostSave = $false
        Update-Ui
        $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
}

function Open-WeekWindow {
    param(
        $Owner
    )

    $data = Get-WeekReportData
    $palette = Get-ThemePalette -Theme ([string]$data.Values.Settings.Theme)
    $bestText = "-"

    if ($null -ne $data.BestDay) {
        $bestText = "{0} {1}" -f (Get-DayNameShort -Date $data.BestDay.Date), (Format-CompactDuration $data.BestDay.NetSeconds)
    }

    $balanceColor = if ($data.BalanceSeconds -ge 0) { "#34C759" } else { "#FF3B30" }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Woche"
        Width="760"
        Height="760"
        MinWidth="680"
        MinHeight="620"
        ResizeMode="CanResize"
        WindowStartupLocation="CenterOwner"
        Background="$($palette.Window)"
        FontFamily="Segoe UI">
    <ScrollViewer VerticalScrollBarVisibility="Auto">
        <StackPanel Margin="18">
            <Border Padding="22" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="24">
                <StackPanel>
                    <TextBlock Text="Woche" FontSize="24" FontWeight="SemiBold" Foreground="$($palette.Primary)"/>
                    <TextBlock Text="$($data.WeekStart.ToString("dd.MM.yyyy")) bis $($data.WeekEnd.ToString("dd.MM.yyyy"))" Margin="0,4,0,18" Foreground="$($palette.Secondary)"/>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0">
                            <TextBlock Text="Netto" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                            <TextBlock Text="$(Format-CompactDuration $data.TotalNetSeconds)" Foreground="$($palette.Primary)" FontSize="26" FontWeight="SemiBold"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1">
                            <TextBlock Text="Wochenziel" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                            <TextBlock Text="$(Format-CompactDuration $data.WeekTargetSeconds)" Foreground="$($palette.Primary)" FontSize="26" FontWeight="SemiBold"/>
                        </StackPanel>
                        <StackPanel Grid.Column="2">
                            <TextBlock Text="Saldo" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                            <TextBlock Text="$(Format-CompactDuration $data.BalanceSeconds)" Foreground="$balanceColor" FontSize="26" FontWeight="SemiBold"/>
                        </StackPanel>
                        <StackPanel Grid.Column="3">
                            <TextBlock Text="Tätigkeiten" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                            <TextBlock Text="$("{0:N2} h" -f $data.TotalActivityHours)" Foreground="$($palette.Primary)" FontSize="26" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Grid>
                </StackPanel>
            </Border>

            <Grid Margin="0,14,0,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="12"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border Grid.Column="0" Padding="18" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="20">
                    <StackPanel>
                        <TextBlock Text="Produktivster Tag" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                        <TextBlock Text="$bestText" Foreground="$($palette.Primary)" FontSize="19" FontWeight="SemiBold" Margin="0,8,0,0"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="2" Padding="18" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="20">
                    <StackPanel>
                        <TextBlock Text="Ø Arbeitstag" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                        <TextBlock Text="$(Format-CompactDuration $data.AverageSeconds)" Foreground="$($palette.Primary)" FontSize="19" FontWeight="SemiBold" Margin="0,8,0,0"/>
                    </StackPanel>
                </Border>
                <Border Grid.Column="4" Padding="18" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="20">
                    <StackPanel>
                        <TextBlock Text="Pausenquote" Foreground="$($palette.Secondary)" FontSize="12" FontWeight="SemiBold"/>
                        <TextBlock Text="$("{0:N1}" -f $data.PauseRatio)%" Foreground="$($palette.Primary)" FontSize="19" FontWeight="SemiBold" Margin="0,8,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>

            <Border Margin="0,14,0,0" Padding="20" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="22">
                <StackPanel>
                    <TextBlock Text="Heute" Foreground="$($palette.Primary)" FontSize="18" FontWeight="SemiBold"/>
                    <Grid x:Name="TimelineBar" Height="18" Margin="0,14,0,12"/>
                    <StackPanel Orientation="Horizontal">
                        <TextBlock Text="Arbeit" Foreground="#0A84FF" Margin="0,0,18,0"/>
                        <TextBlock Text="Auto-Pause" Foreground="#FF9500" Margin="0,0,18,0"/>
                        <TextBlock Text="Manuell" Foreground="#AF52DE"/>
                    </StackPanel>
                </StackPanel>
            </Border>

            <Border Margin="0,14,0,0" Padding="20" Background="$($palette.Card)" BorderBrush="$($palette.Border)" BorderThickness="1" CornerRadius="22">
                <StackPanel>
                    <TextBlock Text="Tage" Foreground="$($palette.Primary)" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,10"/>
                    <StackPanel x:Name="DayListPanel"/>
                </StackPanel>
            </Border>

            <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0">
                <Button x:Name="ExportButton" Content="PDF" Margin="0,0,10,0"/>
                <Button x:Name="CloseButton" Content="Schließen" Background="#8E8E93"/>
            </StackPanel>
        </StackPanel>
    </ScrollViewer>
</Window>
"@

    $dialog = Convert-FromXaml $xaml
    Add-DialogStyles $dialog
    $dialog.Owner = $Owner

    if (Test-Path $IconPath) {
        try {
            $dialog.Icon = New-Object System.Windows.Media.Imaging.BitmapImage([Uri]$IconPath)
        }
        catch {}
    }

    $buttonStyle = $dialog.Resources["DialogButton"]
    $dialog.FindName("ExportButton").Style = $buttonStyle
    $dialog.FindName("CloseButton").Style = $buttonStyle

    $timelineBar = $dialog.FindName("TimelineBar")
    $segments = @(
        [PSCustomObject]@{ Seconds = [math]::Max(0, $data.Values.GrossSeconds - $data.Values.TotalPauseSeconds); Color = "#0A84FF" },
        [PSCustomObject]@{ Seconds = [math]::Max(0, $data.Values.AutoPauseSeconds); Color = "#FF9500" },
        [PSCustomObject]@{ Seconds = [math]::Max(0, $data.Values.ManualPauseSeconds); Color = "#AF52DE" }
    )

    $totalSegmentSeconds = [double](@($segments | Measure-Object -Property Seconds -Sum).Sum)

    if ($totalSegmentSeconds -le 0) {
        $segments = @([PSCustomObject]@{ Seconds = 1; Color = $palette.Soft })
    }

    for ($i = 0; $i -lt $segments.Count; $i++) {
        $timelineBar.ColumnDefinitions.Add((New-Object System.Windows.Controls.ColumnDefinition)) | Out-Null
        $timelineBar.ColumnDefinitions[$i].Width = New-Object System.Windows.GridLength -ArgumentList ([math]::Max(0.01, [double]$segments[$i].Seconds)), ([System.Windows.GridUnitType]::Star)

        $segment = New-Object System.Windows.Controls.Border
        $segment.Background = New-Brush $segments[$i].Color
        $segment.CornerRadius = New-Object System.Windows.CornerRadius -ArgumentList 9
        $segment.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 3, 0
        [System.Windows.Controls.Grid]::SetColumn($segment, $i)
        $timelineBar.Children.Add($segment) | Out-Null
    }

    $dayListPanel = $dialog.FindName("DayListPanel")

    foreach ($day in @($data.Days)) {
        $row = New-Object System.Windows.Controls.Grid
        $row.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 0, 8

        foreach ($width in @(90, 1, 90, 90, 100)) {
            $column = New-Object System.Windows.Controls.ColumnDefinition
            if ($width -eq 1) {
                $column.Width = New-Object System.Windows.GridLength -ArgumentList 1, ([System.Windows.GridUnitType]::Star)
            }
            else {
                $column.Width = New-Object System.Windows.GridLength -ArgumentList $width
            }
            $row.ColumnDefinitions.Add($column) | Out-Null
        }

        $cells = @(
            "{0} {1}" -f (Get-DayNameShort -Date $day.Date), $day.Date.ToString("dd.MM."),
            $day.Status,
            (Format-CompactDuration $day.NetSeconds),
            (Format-CompactDuration $day.PauseSeconds),
            ("{0:N2} h" -f [double]$day.ActivityHours)
        )

        for ($i = 0; $i -lt $cells.Count; $i++) {
            $text = New-Object System.Windows.Controls.TextBlock
            $text.Text = $cells[$i]
            $text.Foreground = New-Brush $(if ($i -eq 2) { $palette.Primary } else { $palette.Secondary })
            $text.FontWeight = if ($i -eq 2) { [System.Windows.FontWeights]::SemiBold } else { [System.Windows.FontWeights]::Normal }
            $text.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
            [System.Windows.Controls.Grid]::SetColumn($text, $i)
            $row.Children.Add($text) | Out-Null
        }

        $dayListPanel.Children.Add($row) | Out-Null
    }

    $dialog.FindName("ExportButton").Add_Click({
        Export-WeekReport -Owner $dialog
    })

    $dialog.FindName("CloseButton").Add_Click({
        $dialog.Close()
    })

    $dialog.ShowDialog() | Out-Null
}

$mainXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Arbeitszeit"
        Width="560"
        Height="820"
        MinWidth="520"
        MinHeight="760"
        ResizeMode="CanResize"
        WindowStartupLocation="Manual"
        Background="#F5F5F7"
        FontFamily="Segoe UI">
    <Window.Resources>
        <Style x:Key="Card" TargetType="Border">
            <Setter Property="Background" Value="White"/>
            <Setter Property="CornerRadius" Value="28"/>
            <Setter Property="Padding" Value="22"/>
            <Setter Property="BorderBrush" Value="#ECECF0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Effect">
                <Setter.Value>
                    <DropShadowEffect BlurRadius="18" ShadowDepth="1" Direction="270" Opacity="0.10"/>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PillButton" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Background" Value="#007AFF"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Root" Background="{TemplateBinding Background}" CornerRadius="20">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Root" Property="Opacity" Value="0.88"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Root" Property="Opacity" Value="0.74"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Root" Property="Background" Value="#E5E5EA"/>
                                <Setter Property="Foreground" Value="#8E8E93"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <ScrollViewer VerticalScrollBarVisibility="Auto">
    <Grid x:Name="RootGrid" Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="16"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="16"/>
            <RowDefinition Height="48"/>
            <RowDefinition Height="12"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="14"/>
            <RowDefinition Height="44"/>
        </Grid.RowDefinitions>

        <Border Style="{StaticResource Card}">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="64"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="22"/>
                    <RowDefinition Height="18"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="NETTO HEUTE" Foreground="#6E6E73" FontSize="12" FontWeight="SemiBold"/>
                <Button x:Name="SetupButton" Grid.Column="1" Content="Setup" Style="{StaticResource PillButton}" Background="#F2F2F7" Foreground="#1C1C1E" Height="34"/>

                <TextBlock x:Name="NetText" Grid.Row="1" Grid.ColumnSpan="2" Text="00:00:00" FontSize="44" FontWeight="SemiBold" Foreground="#1C1C1E" VerticalAlignment="Center"/>

                <Border x:Name="StatusPill" Grid.Row="2" Background="#34C759" CornerRadius="17" Padding="14,7" HorizontalAlignment="Left">
                    <TextBlock x:Name="StatusText" Text="Arbeitszeit läuft" Foreground="White" FontSize="13" FontWeight="SemiBold"/>
                </Border>
                <TextBlock x:Name="RemainingText" Grid.Row="2" Grid.Column="1" Text="Noch --:--:--" Foreground="#6E6E73" FontSize="13" VerticalAlignment="Center" HorizontalAlignment="Right"/>

                <TextBlock x:Name="ForecastText" Grid.Row="3" Grid.ColumnSpan="2" Text="Ohne weitere Pause: --:--" Foreground="#6E6E73" FontSize="13" VerticalAlignment="Center"/>

                <ProgressBar x:Name="TargetProgress" Grid.Row="4" Grid.ColumnSpan="2" Height="8" Minimum="0" Maximum="100" Value="0" Foreground="#007AFF" Background="#E5E5EA" BorderThickness="0" VerticalAlignment="Bottom"/>
            </Grid>
        </Border>

        <Border Grid.Row="2" Style="{StaticResource Card}" Padding="22,18">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="36"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBlock Text="Start" Grid.Row="0" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="StartText" Grid.Row="0" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
                <TextBlock Text="Ende" Grid.Row="1" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="EndText" Grid.Row="1" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
                <TextBlock Text="Brutto" Grid.Row="2" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="GrossText" Grid.Row="2" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
                <TextBlock Text="Pause auto" Grid.Row="3" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="AutoPauseText" Grid.Row="3" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
                <TextBlock Text="Pause manuell" Grid.Row="4" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="ManualPauseText" Grid.Row="4" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
                <TextBlock Text="Aktualisiert" Grid.Row="5" Foreground="#6E6E73" FontSize="14"/>
                <TextBlock x:Name="UpdatedText" Grid.Row="5" Grid.Column="1" Foreground="#1C1C1E" FontSize="15" FontWeight="SemiBold"/>
            </Grid>
        </Border>

        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="1.35*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="0.9*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="8"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <Button x:Name="PauseButton" Grid.Column="0" Content="Pause" Style="{StaticResource PillButton}" Background="#007AFF"/>
            <Button x:Name="ResumeButton" Grid.Column="2" Content="Weiter" Style="{StaticResource PillButton}" Background="#34C759"/>
            <Button x:Name="EditButton" Grid.Column="4" Content="Korrigieren" Style="{StaticResource PillButton}" Background="#3A3A3C"/>
            <Button x:Name="CsvButton" Grid.Column="6" Content="CSV" Style="{StaticResource PillButton}" Background="#3A3A3C"/>
            <Button x:Name="WeekButton" Grid.Column="8" Content="Woche" Style="{StaticResource PillButton}" Background="#5856D6"/>
            <Button x:Name="ThemeButton" Grid.Column="10" Content="Dark" Style="{StaticResource PillButton}" Background="#3A3A3C"/>
        </Grid>

        <Border Grid.Row="6" Style="{StaticResource Card}" Padding="18">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="10"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="6"/>
                    <RowDefinition Height="36"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="8"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="96"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="70"/>
                    <ColumnDefinition Width="8"/>
                    <ColumnDefinition Width="86"/>
                </Grid.ColumnDefinitions>

                <Grid Grid.ColumnSpan="7">
                    <TextBlock Text="TÄTIGKEITEN" Foreground="#6E6E73" FontSize="12" FontWeight="SemiBold"/>
                    <TextBlock x:Name="ActivityStatusText" Text="0 Einträge · 0,00 h" HorizontalAlignment="Right" Foreground="#8E8E93" FontSize="12"/>
                </Grid>

                <TextBlock Grid.Row="2" Grid.Column="0" Text="Projektnummer" Foreground="#8E8E93" FontSize="11" FontWeight="SemiBold"/>
                <TextBlock Grid.Row="2" Grid.Column="2" Text="Tätigkeit" Foreground="#8E8E93" FontSize="11" FontWeight="SemiBold"/>
                <TextBlock Grid.Row="2" Grid.Column="4" Text="Dauer" Foreground="#8E8E93" FontSize="11" FontWeight="SemiBold"/>

                <TextBox x:Name="ActivityProjectBox" Grid.Row="4" Grid.Column="0" ToolTip="Projektnummer" FontSize="13" Padding="10,7" BorderThickness="1"/>
                <TextBox x:Name="ActivityDescriptionBox" Grid.Row="4" Grid.Column="2" ToolTip="Tätigkeit" FontSize="13" Padding="10,7" BorderThickness="1"/>
                <TextBox x:Name="ActivityHoursBox" Grid.Row="4" Grid.Column="4" ToolTip="Dauer in Stunden, z.B. 1 oder 1,5" FontSize="13" Padding="10,7" BorderThickness="1"/>
                <Button x:Name="ActivitySaveButton" Grid.Row="4" Grid.Column="6" Content="Speichern" Style="{StaticResource PillButton}" Background="#007AFF" Height="36" Padding="10,0"/>

                <ScrollViewer Grid.Row="8" Grid.ColumnSpan="7" MaxHeight="94" VerticalScrollBarVisibility="Auto">
                    <StackPanel x:Name="ActivityListPanel"/>
                </ScrollViewer>
            </Grid>
        </Border>

        <Border Grid.Row="8" Background="White" CornerRadius="21" BorderBrush="#ECECF0" BorderThickness="1" Padding="16,0">
            <Grid>
                <CheckBox x:Name="TopMostBox" Content="Immer oben" VerticalAlignment="Center" Foreground="#6E6E73"/>
                <TextBlock Text="Schließen minimiert in den Infobereich" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="#8E8E93" FontSize="12"/>
            </Grid>
        </Border>
    </Grid>
    </ScrollViewer>
</Window>
"@

$mainWindow = Convert-FromXaml $mainXaml

if (Test-Path $IconPath) {
    try {
        $mainWindow.Icon = New-Object System.Windows.Media.Imaging.BitmapImage([Uri]$IconPath)
    }
    catch {}
}

$workArea = [System.Windows.SystemParameters]::WorkArea
$mainWindow.Left = $workArea.Right - $mainWindow.Width - 18
$mainWindow.Top = $workArea.Bottom - $mainWindow.Height - 18

$rootGrid = $mainWindow.FindName("RootGrid")
$netText = $mainWindow.FindName("NetText")
$statusText = $mainWindow.FindName("StatusText")
$statusPill = $mainWindow.FindName("StatusPill")
$remainingText = $mainWindow.FindName("RemainingText")
$forecastText = $mainWindow.FindName("ForecastText")
$targetProgress = $mainWindow.FindName("TargetProgress")
$startText = $mainWindow.FindName("StartText")
$endText = $mainWindow.FindName("EndText")
$grossText = $mainWindow.FindName("GrossText")
$autoPauseText = $mainWindow.FindName("AutoPauseText")
$manualPauseText = $mainWindow.FindName("ManualPauseText")
$updatedText = $mainWindow.FindName("UpdatedText")
$pauseButton = $mainWindow.FindName("PauseButton")
$resumeButton = $mainWindow.FindName("ResumeButton")
$editButton = $mainWindow.FindName("EditButton")
$csvButton = $mainWindow.FindName("CsvButton")
$weekButton = $mainWindow.FindName("WeekButton")
$themeButton = $mainWindow.FindName("ThemeButton")
$setupButton = $mainWindow.FindName("SetupButton")
$activityHoursBox = $mainWindow.FindName("ActivityHoursBox")
$activityDescriptionBox = $mainWindow.FindName("ActivityDescriptionBox")
$activityProjectBox = $mainWindow.FindName("ActivityProjectBox")
$activitySaveButton = $mainWindow.FindName("ActivitySaveButton")
$activityListPanel = $mainWindow.FindName("ActivityListPanel")
$activityStatusText = $mainWindow.FindName("ActivityStatusText")
$topMostBox = $mainWindow.FindName("TopMostBox")

$script:SuppressTopMostSave = $false
$script:AllowWindowClose = $false
$script:TrayHintShown = $false
$script:TrayIcon = $null
$script:EditingActivityId = ""

function Set-Brush {
    param(
        $Element,
        [string]$Color
    )

    $Element.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
}

function Show-MainWindow {
    try {
        $mainWindow.Show()
        $mainWindow.WindowState = [System.Windows.WindowState]::Normal
        $mainWindow.Activate() | Out-Null
    }
    catch {
        Write-AppLog ("Fenster-Anzeigen-Fehler: " + $_.Exception.ToString())
    }
}

function Hide-MainWindowToTray {
    try {
        $mainWindow.Hide()

        if (-not $script:TrayHintShown -and $null -ne $script:TrayIcon) {
            $script:TrayIcon.BalloonTipTitle = "Arbeitszeit läuft weiter"
            $script:TrayIcon.BalloonTipText = "Die Anzeige ist unten im Infobereich erreichbar."
            $script:TrayIcon.ShowBalloonTip(2500)
            $script:TrayHintShown = $true
        }
    }
    catch {
        Write-AppLog ("Fenster-Ausblenden-Fehler: " + $_.Exception.ToString())
    }
}

function Initialize-TrayIcon {
    if ($null -ne $script:TrayIcon) {
        return
    }

    $trayIcon = New-Object System.Windows.Forms.NotifyIcon

    if (Test-Path $IconPath) {
        $trayIcon.Icon = New-Object System.Drawing.Icon -ArgumentList $IconPath
    }
    else {
        $trayIcon.Icon = [System.Drawing.SystemIcons]::Application
    }

    $trayIcon.Text = "Arbeitszeit"
    $trayIcon.Visible = $true

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $openItem = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList "Öffnen"
    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem -ArgumentList "Beenden"

    [void]$menu.Items.Add($openItem)
    [void]$menu.Items.Add($exitItem)
    $trayIcon.ContextMenuStrip = $menu

    $openAction = {
        try {
            $mainWindow.Dispatcher.Invoke([Action]{ Show-MainWindow }) | Out-Null
        }
        catch {
            Write-AppLog ("Tray-Oeffnen-Fehler: " + $_.Exception.ToString())
        }
    }

    $openItem.Add_Click($openAction)
    $trayIcon.Add_DoubleClick($openAction)

    $exitItem.Add_Click({
        try {
            $mainWindow.Dispatcher.Invoke([Action]{
                $script:AllowWindowClose = $true
                $mainWindow.Close()
            }) | Out-Null
        }
        catch {
            Write-AppLog ("Tray-Beenden-Fehler: " + $_.Exception.ToString())
        }
    })

    $script:TrayIcon = $trayIcon
}

function Reset-ActivityEditor {
    $script:EditingActivityId = ""
    $activityProjectBox.Text = ""
    $activityDescriptionBox.Text = ""
    $activityHoursBox.Text = ""
    $activitySaveButton.Content = "Speichern"
}

function Refresh-ActivityList {
    $entries = @(Get-TodayWorkEntries)
    $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
    $palette = Get-ThemePalette -Theme ([string](Get-UiSettings).Theme)

    $activityListPanel.Children.Clear()

    if ($entries.Count -eq 0) {
        $emptyText = New-Object System.Windows.Controls.TextBlock
        $emptyText.Text = "Noch keine Tätigkeiten erfasst"
        $emptyText.Foreground = New-Brush $palette.Secondary
        $emptyText.FontSize = 12
        $emptyText.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 2, 0, 0
        $activityListPanel.Children.Add($emptyText) | Out-Null
        return
    }

    foreach ($entry in $entries) {
        $hours = 0.0

        try {
            $hours = [double]$entry.Stunden
        }
        catch {}

        $rowBorder = New-Object System.Windows.Controls.Border
        $rowBorder.Background = New-Brush $palette.Soft
        $rowBorder.BorderBrush = New-Brush $palette.Border
        $rowBorder.BorderThickness = New-Object System.Windows.Thickness -ArgumentList 1
        $rowBorder.CornerRadius = New-Object System.Windows.CornerRadius -ArgumentList 16
        $rowBorder.Padding = New-Object System.Windows.Thickness -ArgumentList 12, 7, 8, 7
        $rowBorder.Margin = New-Object System.Windows.Thickness -ArgumentList 0, 0, 0, 7

        $row = New-Object System.Windows.Controls.Grid
        foreach ($width in @(76, 1, 72, 54, 50)) {
            $column = New-Object System.Windows.Controls.ColumnDefinition

            if ($width -eq 1) {
                $column.Width = New-Object System.Windows.GridLength -ArgumentList 1, ([System.Windows.GridUnitType]::Star)
            }
            else {
                $column.Width = New-Object System.Windows.GridLength -ArgumentList $width
            }

            $row.ColumnDefinitions.Add($column) | Out-Null
        }

        $projectText = New-Object System.Windows.Controls.TextBlock
        $projectText.Text = if ([string]::IsNullOrWhiteSpace([string]$entry.Projekt)) { "-" } else { [string]$entry.Projekt }
        $projectText.Foreground = New-Brush $palette.Secondary
        $projectText.FontWeight = [System.Windows.FontWeights]::SemiBold
        $projectText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $projectText.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
        [System.Windows.Controls.Grid]::SetColumn($projectText, 0)
        $row.Children.Add($projectText) | Out-Null

        $descriptionText = New-Object System.Windows.Controls.TextBlock
        $descriptionText.Text = [string]$entry.Beschreibung
        $descriptionText.Foreground = New-Brush $palette.Primary
        $descriptionText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $descriptionText.TextTrimming = [System.Windows.TextTrimming]::CharacterEllipsis
        $descriptionText.Margin = New-Object System.Windows.Thickness -ArgumentList 8, 0, 8, 0
        [System.Windows.Controls.Grid]::SetColumn($descriptionText, 1)
        $row.Children.Add($descriptionText) | Out-Null

        $durationText = New-Object System.Windows.Controls.TextBlock
        $durationText.Text = $hours.ToString("N2", $deCulture) + " h"
        $durationText.Foreground = New-Brush $palette.Primary
        $durationText.FontWeight = [System.Windows.FontWeights]::SemiBold
        $durationText.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
        $durationText.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
        [System.Windows.Controls.Grid]::SetColumn($durationText, 2)
        $row.Children.Add($durationText) | Out-Null

        $editButton = New-Object System.Windows.Controls.Button
        $editButton.Content = "Bearb."
        $editButton.Tag = [string]$entry.Id
        $editButton.Height = 28
        $editButton.Padding = New-Object System.Windows.Thickness -ArgumentList 8, 0, 8, 0
        $editButton.Margin = New-Object System.Windows.Thickness -ArgumentList 8, 0, 0, 0
        $editButton.Background = New-Brush "#3A3A3C"
        $editButton.Foreground = New-Brush "#FFFFFF"
        $editButton.BorderThickness = New-Object System.Windows.Thickness -ArgumentList 0
        $editButton.Style = $mainWindow.Resources["PillButton"]
        [System.Windows.Controls.Grid]::SetColumn($editButton, 3)
        $row.Children.Add($editButton) | Out-Null

        $deleteButton = New-Object System.Windows.Controls.Button
        $deleteButton.Content = "Entf."
        $deleteButton.Tag = [string]$entry.Id
        $deleteButton.Height = 28
        $deleteButton.Padding = New-Object System.Windows.Thickness -ArgumentList 8, 0, 8, 0
        $deleteButton.Margin = New-Object System.Windows.Thickness -ArgumentList 6, 0, 0, 0
        $deleteButton.Background = New-Brush "#FF3B30"
        $deleteButton.Foreground = New-Brush "#FFFFFF"
        $deleteButton.BorderThickness = New-Object System.Windows.Thickness -ArgumentList 0
        $deleteButton.Style = $mainWindow.Resources["PillButton"]
        [System.Windows.Controls.Grid]::SetColumn($deleteButton, 4)
        $row.Children.Add($deleteButton) | Out-Null

        $editButton.Add_Click({
            param($sender, $eventArgs)
            Load-ActivityIntoEditorById -Id ([string]$sender.Tag)
        })

        $deleteButton.Add_Click({
            param($sender, $eventArgs)
            Remove-ActivityWithConfirm -Id ([string]$sender.Tag)
        })

        $rowBorder.Child = $row
        $activityListPanel.Children.Add($rowBorder) | Out-Null
    }
}

function Load-ActivityIntoEditorById {
    param(
        [string]$Id
    )

    $entry = @(Read-WorkEntries | Where-Object { [string]$_.Id -eq $Id } | Select-Object -First 1)

    if ($entry.Count -eq 0) {
        return
    }

    $item = $entry[0]
    $script:EditingActivityId = [string]$item.Id
    $activityProjectBox.Text = [string]$item.Projekt
    $activityDescriptionBox.Text = [string]$item.Beschreibung
    $activityHoursBox.Text = ([double]$item.Stunden).ToString("0.##", [System.Globalization.CultureInfo]::GetCultureInfo("de-DE"))
    $activitySaveButton.Content = "Aktualisieren"
}

function Remove-ActivityWithConfirm {
    param(
        [string]$Id
    )

    if ([string]::IsNullOrWhiteSpace($Id)) {
        return
    }

    $result = [System.Windows.MessageBox]::Show(
        "Diese Tätigkeit löschen?",
        "Arbeitszeit",
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )

    if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
        return
    }

    Remove-WorkEntryLocal -Id $Id
    Reset-ActivityEditor
    Refresh-ActivityList
    Update-Ui
}

function Update-Ui {
    $values = Get-LiveValues

    $netText.Text = Format-Duration $values.NetSeconds
    $statusText.Text = $values.StatusText
    $remainingText.Text = "Noch " + (Format-Duration $values.RemainingSeconds)
    $forecastText.Text = "Ohne weitere Pause: " + $values.ForecastTime
    $targetProgress.Value = [math]::Max(0, [math]::Min(100, [double]$values.Progress * 100))
    $startText.Text = $values.StartTime
    $endText.Text = $values.EndTime
    $grossText.Text = Format-Duration $values.GrossSeconds
    $autoPauseText.Text = Format-Duration $values.AutoPauseSeconds
    $manualPauseText.Text = Format-Duration $values.ManualPauseSeconds
    $updatedText.Text = $values.UpdatedText
    $activitySummary = Get-TodayWorkEntrySummary
    $activityStatusText.Text = "{0} Einträge · {1:N2} h" -f $activitySummary.Count, $activitySummary.Hours
    if ([string]::IsNullOrWhiteSpace($script:EditingActivityId)) {
        Refresh-ActivityList
    }

    switch ($values.StatusKind) {
        "Pause" { Set-Brush $statusPill "#FF9500" }
        "Stale" { Set-Brush $statusPill "#FF3B30" }
        "Missing" { Set-Brush $statusPill "#8E8E93" }
        default { Set-Brush $statusPill "#34C759" }
    }

    $pauseButton.IsEnabled = $values.HasState -and -not $values.ManualPauseActive
    $resumeButton.IsEnabled = $values.HasState -and $values.ManualPauseActive
    $editButton.IsEnabled = $values.HasState
    $mainWindow.Topmost = [bool]$values.Settings.AlwaysOnTop

    if ($topMostBox.IsChecked -ne [bool]$values.Settings.AlwaysOnTop) {
        $script:SuppressTopMostSave = $true
        $topMostBox.IsChecked = [bool]$values.Settings.AlwaysOnTop
        $script:SuppressTopMostSave = $false
    }

    Apply-MainTheme -Settings $values.Settings
    Test-TargetNotification -Values $values
}

$activitySaveButton.Add_Click({
    try {
        $hours = Convert-ToWorkEntryHours -Text ([string]$activityHoursBox.Text)
        $description = ([string]$activityDescriptionBox.Text).Trim()
        $project = ([string]$activityProjectBox.Text).Trim()

        if ($hours -le 0) {
            Show-Warning "Bitte Stunden eintragen, z.B. 1 oder 1,5."
            return
        }

        if ([string]::IsNullOrWhiteSpace($description)) {
            Show-Warning "Bitte eine kurze Beschreibung eintragen."
            return
        }

        $activitySaveButton.IsEnabled = $false
        $activityStatusText.Text = "Tätigkeit wird gespeichert..."

        if ([string]::IsNullOrWhiteSpace($script:EditingActivityId)) {
            Add-WorkEntryLocal -Hours $hours -Description $description -Project $project
        }
        else {
            Update-WorkEntryLocal `
                -Id $script:EditingActivityId `
                -Hours $hours `
                -Description $description `
                -Project $project
        }

        Reset-ActivityEditor
        Refresh-ActivityList
        Update-Ui
    }
    catch {
        Write-AppLog ("Taetigkeit-Speichern-Fehler: " + $_.Exception.ToString())
        Show-Warning "Die Tätigkeit konnte nicht gespeichert werden. Details stehen im Log."
    }
    finally {
        $activitySaveButton.IsEnabled = $true
    }
})

$pauseButton.Add_Click({
    Write-ControlCommand -Action "StartManualPause"
    Start-Sleep -Milliseconds 200
    Update-Ui
})

$resumeButton.Add_Click({
    Write-ControlCommand -Action "StopManualPause"
    Start-Sleep -Milliseconds 200
    Update-Ui
})

$editButton.Add_Click({
    Open-CorrectionWindow -Owner $mainWindow
})

$setupButton.Add_Click({
    Open-SettingsWindow -Owner $mainWindow
})

$csvButton.Add_Click({
    if (Test-Path $CsvPath) {
        Start-Process -FilePath $CsvPath
    }
    else {
        Show-Info "Die CSV-Datei existiert noch nicht."
    }
})

$weekButton.Add_Click({
    try {
        Open-WeekWindow -Owner $mainWindow
    }
    catch {
        Write-AppLog ("Woche-Fehler: " + $_.Exception.ToString())
        Show-Warning "Die Wochenübersicht konnte nicht geöffnet werden. Details stehen im Log."
    }
})

$themeButton.Add_Click({
    try {
        $settings = Get-UiSettings
        $settings.Theme = if ($settings.Theme -eq "Dark") { "Light" } else { "Dark" }
        Save-ArbeitszeitSettings -BaseDir $BaseDir -Settings $settings
        Update-Ui
    }
    catch {
        Write-AppLog ("Theme-Fehler: " + $_.Exception.ToString())
    }
})

$topMostBox.Add_Click({
    if ($script:SuppressTopMostSave) {
        return
    }

    $settings = Get-UiSettings
    $settings.AlwaysOnTop = [bool]$topMostBox.IsChecked
    Save-ArbeitszeitSettings -BaseDir $BaseDir -Settings $settings
    $mainWindow.Topmost = [bool]$topMostBox.IsChecked
})

$timer = New-Object System.Windows.Threading.DispatcherTimer
$timer.Interval = [TimeSpan]::FromSeconds(1)
$timer.Add_Tick({
    try {
        Update-Ui
    }
    catch {
        Write-AppLog ("UI-Timer-Fehler: " + $_.Exception.ToString())
    }
})

$mainWindow.Add_Loaded({
    try {
        Initialize-TrayIcon
        Update-WorkCsvActivityColumns
        Update-Ui
    }
    catch {
        Write-AppLog ("Initialer UI-Fehler: " + $_.Exception.ToString())
    }

    $timer.Start()

    $helper = New-Object System.Windows.Interop.WindowInteropHelper($mainWindow)
    [ArbeitszeitWindowHelper]::ShowWindow($helper.Handle, 5) | Out-Null
    [ArbeitszeitWindowHelper]::SetForegroundWindow($helper.Handle) | Out-Null
    $mainWindow.Activate() | Out-Null
})

$mainWindow.Add_Closing({
    param($sender, $eventArgs)

    if (-not $script:AllowWindowClose) {
        $eventArgs.Cancel = $true
        Hide-MainWindowToTray
    }
})

$mainWindow.Add_Closed({
    $timer.Stop()

    if ($null -ne $script:TrayIcon) {
        $script:TrayIcon.Visible = $false
        $script:TrayIcon.Dispose()
        $script:TrayIcon = $null
    }
})

try {
    [System.Windows.Application]::new().Run($mainWindow) | Out-Null
}
finally {
    if ($null -ne $mutex) {
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}
