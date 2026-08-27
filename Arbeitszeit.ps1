param(
    [string]$BaseDir = "C:\Arbeitszeit",
    [int]$IntervalSeconds = 5,
    [int]$IdleThresholdSeconds = 20,
    [switch]$Once
)

# Arbeitszeit-Erfassung fuer Windows
# Ziel: C:\Arbeitszeit\Arbeitszeit.ps1

$ErrorActionPreference = "Stop"

$CsvPath = Join-Path $BaseDir "Arbeitszeiten.csv"
$ActivityCsvPath = Join-Path $BaseDir "Arbeitszeit_Taetigkeiten.csv"
$ActivityJsonPath = Join-Path $BaseDir "taetigkeiten.json"
$StatePath = Join-Path $BaseDir "state.json"
$ControlPath = Join-Path $BaseDir "control.json"
$SettingsPath = Join-Path $BaseDir "settings.json"
$TrackerLogPath = Join-Path $BaseDir "ArbeitszeitTracker.log"
$SharedPath = Join-Path $PSScriptRoot "ArbeitszeitSettings.ps1"

if (!(Test-Path $BaseDir)) {
    New-Item -ItemType Directory -Path $BaseDir | Out-Null
}

if (!(Test-Path $SharedPath)) {
    throw "Gemeinsame Settings-Datei nicht gefunden: $SharedPath"
}

. $SharedPath

function Write-TrackerLog {
    param(
        [string]$Message
    )

    try {
        Add-Content -LiteralPath $TrackerLogPath -Value ((Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + " " + $Message) -Encoding UTF8
    }
    catch {}
}

# Verhindert, dass der Tracker parallel doppelt laeuft.
$mutexCreated = $false
$mutex = [System.Threading.Mutex]::new($true, "ArbeitszeitTrackerLocal", [ref]$mutexCreated)

if (-not $mutexCreated) {
    exit
}

if (-not ("IdleTimeHelper" -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public static class IdleTimeHelper
{
    [StructLayout(LayoutKind.Sequential)]
    struct LASTINPUTINFO
    {
        public uint cbSize;
        public uint dwTime;
    }

    [DllImport("user32.dll")]
    static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    public static uint GetIdleTimeMilliseconds()
    {
        LASTINPUTINFO lii = new LASTINPUTINFO();
        lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));

        if (!GetLastInputInfo(ref lii))
        {
            return 0;
        }

        return ((uint)Environment.TickCount - lii.dwTime);
    }

    public static uint GetIdleTimeSeconds()
    {
        return GetIdleTimeMilliseconds() / 1000;
    }
}
"@
}

function Get-IdleMilliseconds {
    return [IdleTimeHelper]::GetIdleTimeMilliseconds()
}

function Get-IdleSeconds {
    return [IdleTimeHelper]::GetIdleTimeMilliseconds() / 1000.0
}

function Format-Duration {
    param(
        [double]$Seconds
    )

    if ($Seconds -lt 0) {
        $Seconds = 0
    }

    $ts = [TimeSpan]::FromSeconds($Seconds)
    $hours = [math]::Floor($ts.TotalHours)

    return "{0:00}:{1:00}:{2:00}" -f $hours, $ts.Minutes, $ts.Seconds
}

function Get-MinDate {
    param(
        [datetime[]]$Dates
    )

    $sorted = @($Dates | Sort-Object)
    return $sorted[0]
}

function Get-MaxDate {
    param(
        [datetime[]]$Dates
    )

    $sorted = @($Dates | Sort-Object)
    return $sorted[$sorted.Count - 1]
}

function Get-TodayDateTime {
    param(
        [datetime]$Now,
        [string]$TimeText
    )

    $dateText = $Now.ToString("yyyy-MM-dd")

    return [datetime]::ParseExact(
        "$dateText $TimeText",
        "yyyy-MM-dd HH:mm",
        [System.Globalization.CultureInfo]::InvariantCulture
    )
}

function Get-StateDateTime {
    param(
        $State,
        [string]$TimeText
    )

    $normalized = Convert-ToTimeText -TimeText $TimeText

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return $null
    }

    return [datetime]::ParseExact(
        "$($State.Date) $normalized",
        "yyyy-MM-dd HH:mm:ss",
        [System.Globalization.CultureInfo]::InvariantCulture
    )
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

function New-DayState {
    param(
        [datetime]$Now
    )

    return [PSCustomObject]@{
        Date                     = $Now.ToString("yyyy-MM-dd")
        StartTime                = ""
        EndTime                  = ""
        StartPending             = $true
        StartCandidateAt         = ""

        GrossSeconds             = 0

        PauseMorningSeconds      = 0
        PauseNoonSeconds         = 0
        ManualPauseSeconds       = 0

        PauseMorningCountedUntil = ""
        PauseNoonCountedUntil    = ""

        ManualPauseActive        = $false
        ManualPauseStartedAt     = ""
        ManualPauseCountedUntil  = ""

        Note                     = ""
        LastControlId            = ""
        LastTimestamp            = $Now.ToString("o")
    }
}

function Ensure-Property {
    param(
        $Object,
        [string]$Name,
        $Value
    )

    if ($Object.PSObject.Properties.Name -notcontains $Name) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }

    return $Object
}

function Ensure-StateProperties {
    param($State)

    $State = Ensure-Property -Object $State -Name "PauseMorningSeconds" -Value 0
    $State = Ensure-Property -Object $State -Name "PauseNoonSeconds" -Value 0
    $State = Ensure-Property -Object $State -Name "ManualPauseSeconds" -Value 0
    $State = Ensure-Property -Object $State -Name "PauseMorningCountedUntil" -Value ""
    $State = Ensure-Property -Object $State -Name "PauseNoonCountedUntil" -Value ""
    $State = Ensure-Property -Object $State -Name "ManualPauseActive" -Value $false
    $State = Ensure-Property -Object $State -Name "ManualPauseStartedAt" -Value ""
    $State = Ensure-Property -Object $State -Name "ManualPauseCountedUntil" -Value ""
    $State = Ensure-Property -Object $State -Name "StartPending" -Value $false
    $State = Ensure-Property -Object $State -Name "StartCandidateAt" -Value ""
    $State = Ensure-Property -Object $State -Name "Note" -Value ""
    $State = Ensure-Property -Object $State -Name "LastControlId" -Value ""

    $State.ManualPauseActive = [System.Convert]::ToBoolean($State.ManualPauseActive)
    $State.StartPending = [System.Convert]::ToBoolean($State.StartPending)

    return $State
}

function Ensure-PauseStateProperties {
    param(
        $State,
        $Settings
    )

    if ($null -eq $Settings) {
        return $State
    }

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
        $State = Ensure-Property -Object $State -Name $pause.SecondsProperty -Value 0
        $State = Ensure-Property -Object $State -Name $pause.CountedProperty -Value ""
    }

    return $State
}

function Repair-PauseStateForCurrentDay {
    param(
        $State,
        $Settings
    )

    if ($null -eq $Settings) {
        return $State
    }

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
        $secondsProperty = [string]$pause.SecondsProperty
        $countedProperty = [string]$pause.CountedProperty

        $State = Ensure-Property -Object $State -Name $secondsProperty -Value 0
        $State = Ensure-Property -Object $State -Name $countedProperty -Value ""

        $seconds = 0.0

        try {
            $seconds = [double]$State.$secondsProperty
        }
        catch {
            $seconds = 0
        }

        if ($seconds -gt 0 -and [string]::IsNullOrWhiteSpace([string]$State.$countedProperty)) {
            $State.$secondsProperty = 0
        }
    }

    return $State
}

function Get-AutoPauseSeconds {
    param(
        $State,
        $Settings
    )

    $seconds = 0.0

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
        $property = [string]$pause.SecondsProperty

        if ($State.PSObject.Properties.Name -contains $property) {
            $seconds += [double]$State.$property
        }
    }

    return $seconds
}

function Get-PauseCsvColumnName {
    param($Pause)

    if ($Pause.Key -eq "Morning") {
        return "Pause_08_55_09_35"
    }

    if ($Pause.Key -eq "Noon") {
        return "Pause_11_55_12_45"
    }

    $key = Convert-ArbeitszeitPauseKey -Key ([string]$Pause.Key) -Fallback ([string]$Pause.Label)
    return "Pause_$key"
}

function Save-State {
    param($State)

    $tmpPath = $StatePath + ".tmp"
    $backupPath = $StatePath + ".bak"
    $json = $State | ConvertTo-Json -Depth 4

    if (Test-Path $StatePath) {
        $stateItem = Get-Item -LiteralPath $StatePath -ErrorAction SilentlyContinue

        if ($null -ne $stateItem -and $stateItem.Length -gt 0) {
            Copy-Item -LiteralPath $StatePath -Destination $backupPath -Force -ErrorAction SilentlyContinue
        }
    }

    $json | Set-Content -LiteralPath $tmpPath -Encoding UTF8
    Copy-Item -LiteralPath $tmpPath -Destination $StatePath -Force
    Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
}

function Recover-StateFromCsv {
    param(
        [datetime]$Now
    )

    if (!(Test-Path $CsvPath)) {
        return $null
    }

    try {
        $today = $Now.ToString("yyyy-MM-dd")
        $row = Import-Csv -LiteralPath $CsvPath -Delimiter ";" |
            Where-Object {
                $rowDate = Get-WorkDate -DateText ([string]$_.Datum)
                $null -ne $rowDate -and $rowDate.ToString("yyyy-MM-dd") -eq $today
            } |
            Select-Object -Last 1

        if ($null -eq $row) {
            return $null
        }

        $state = New-DayState -Now $Now
        $state.Date = $today
        $state.StartTime = if ([string]::IsNullOrWhiteSpace([string]$row.Start)) { $Now.ToString("HH:mm:ss") } else { [string]$row.Start }
        $state.EndTime = $Now.ToString("HH:mm:ss")
        $state.StartPending = $false
        $state.StartCandidateAt = ""
        $state.GrossSeconds = Convert-DurationTextToSeconds ([string]$row.Brutto)
        $state.PauseMorningSeconds = Convert-DurationTextToSeconds ([string]$row.Pause_08_55_09_35)
        $state.PauseNoonSeconds = Convert-DurationTextToSeconds ([string]$row.Pause_11_55_12_45)
        $state.ManualPauseSeconds = Convert-DurationTextToSeconds ([string]$row.Pause_Manuell)
        $state.Note = [string]$row.Notiz
        $state.LastTimestamp = $Now.ToString("o")

        return Ensure-StateProperties $state
    }
    catch {
        return $null
    }
}

function Load-State {
    if (Test-Path $StatePath) {
        try {
            $item = Get-Item -LiteralPath $StatePath -ErrorAction SilentlyContinue

            if ($null -eq $item -or $item.Length -le 0) {
                return Recover-StateFromCsv -Now (Get-Date)
            }

            $loaded = Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json
            return Ensure-StateProperties $loaded
        }
        catch {
            $recovered = Recover-StateFromCsv -Now (Get-Date)

            if ($null -ne $recovered) {
                return $recovered
            }

            return $null
        }
    }

    return Recover-StateFromCsv -Now (Get-Date)
}

function Load-ControlCommand {
    if (!(Test-Path $ControlPath)) {
        return $null
    }

    try {
        return Get-Content $ControlPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Clear-ControlCommand {
    try {
        Remove-Item -LiteralPath $ControlPath -Force -ErrorAction SilentlyContinue
    }
    catch {}
}

function Get-ControlTime {
    param(
        $Command,
        [datetime]$Now
    )

    if ($null -ne $Command.CreatedAt -and -not [string]::IsNullOrWhiteSpace($Command.CreatedAt)) {
        try {
            $commandTime = [datetime]::Parse($Command.CreatedAt)

            if ($commandTime -le $Now) {
                return $commandTime
            }
        }
        catch {
            return $Now
        }
    }

    return $Now
}

function Add-PauseOverlap {
    param(
        $State,
        [datetime]$Now,
        [datetime]$IdleStart,
        [datetime]$MeasurementStart,
        [bool]$LimitToMeasurementStart,
        $Pause
    )

    $windowStart = Get-TodayDateTime -Now $Now -TimeText $Pause.Start
    $windowEnd = Get-TodayDateTime -Now $Now -TimeText $Pause.End

    if ($Now -le $windowStart) {
        return $State
    }

    if ($IdleStart -ge $windowEnd) {
        return $State
    }

    $secondsProperty = $Pause.SecondsProperty
    $countedProperty = $Pause.CountedProperty
    $State = Ensure-Property -Object $State -Name $secondsProperty -Value 0
    $State = Ensure-Property -Object $State -Name $countedProperty -Value ""
    $countedUntilText = $State.$countedProperty

    if ([string]::IsNullOrWhiteSpace($countedUntilText)) {
        $countedUntil = $windowStart
    }
    else {
        $countedUntil = [datetime]::Parse($countedUntilText)
    }

    # Normalfall: Sobald die Idle-Schwelle erreicht ist, wird rueckwirkend
    # ab Beginn der Inaktivitaet gezaehlt. Bei Standby-/Ruhezustand-Luecken
    # begrenzen wir weiterhin auf den wirklich gemessenen Zeitraum.
    if ($LimitToMeasurementStart) {
        $from = Get-MaxDate @($IdleStart, $windowStart, $countedUntil, $MeasurementStart)
    }
    else {
        $from = Get-MaxDate @($IdleStart, $windowStart, $countedUntil)
    }

    $to = Get-MinDate @($Now, $windowEnd)

    if ($to -gt $from) {
        $addSeconds = ($to - $from).TotalSeconds

        $State.$secondsProperty = [double]$State.$secondsProperty + $addSeconds
        $State.$countedProperty = $to.ToString("o")
    }

    return $State
}

function Add-ManualPauseUntil {
    param(
        $State,
        [datetime]$Until,
        [datetime]$MeasurementStart
    )

    if (-not [bool]$State.ManualPauseActive) {
        return $State
    }

    if ([string]::IsNullOrWhiteSpace($State.ManualPauseStartedAt)) {
        $State.ManualPauseStartedAt = $Until.ToString("o")
        $State.ManualPauseCountedUntil = $Until.ToString("o")
        return $State
    }

    $startedAt = [datetime]::Parse($State.ManualPauseStartedAt)

    if ([string]::IsNullOrWhiteSpace($State.ManualPauseCountedUntil)) {
        $countedUntil = $startedAt
    }
    else {
        $countedUntil = [datetime]::Parse($State.ManualPauseCountedUntil)
    }

    $from = Get-MaxDate @($startedAt, $countedUntil, $MeasurementStart)
    $to = $Until

    if ($to -gt $from) {
        $State.ManualPauseSeconds = [double]$State.ManualPauseSeconds + ($to - $from).TotalSeconds
        $State.ManualPauseCountedUntil = $to.ToString("o")
    }

    return $State
}

function Get-EditedPauseCountedUntil {
    param(
        [datetime]$Now,
        $Pause
    )

    $windowStart = Get-TodayDateTime -Now $Now -TimeText $Pause.Start
    $windowEnd = Get-TodayDateTime -Now $Now -TimeText $Pause.End

    if ($Now -lt $windowStart) {
        return ""
    }

    if ($Now -gt $windowEnd) {
        return $windowEnd.ToString("o")
    }

    return $Now.ToString("o")
}

function Get-NumberValue {
    param(
        $Values,
        [string]$Name,
        [double]$DefaultValue
    )

    if ($null -eq $Values -or $Values.PSObject.Properties.Name -notcontains $Name) {
        return $DefaultValue
    }

    try {
        $value = [double]$Values.$Name

        if ($value -lt 0) {
            return 0
        }

        return $value
    }
    catch {
        return $DefaultValue
    }
}

function Get-TextValue {
    param(
        $Values,
        [string]$Name
    )

    if ($null -eq $Values -or $Values.PSObject.Properties.Name -notcontains $Name) {
        return ""
    }

    return ([string]$Values.$Name).Trim()
}

function Convert-ToWorkEntryHours {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return 0.0
    }

    $text = ([string]$Value).Trim().Replace(",", ".")
    $hours = 0.0

    if ([double]::TryParse(
        $text,
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

function Add-WorkEntry {
    param(
        [datetime]$Now,
        $Values
    )

    $hours = Convert-ToWorkEntryHours -Value (Get-TextValue -Values $Values -Name "Hours")
    $description = Get-TextValue -Values $Values -Name "Description"
    $project = Get-TextValue -Values $Values -Name "Project"

    if ($hours -le 0 -or [string]::IsNullOrWhiteSpace($description)) {
        return
    }

    $entry = [PSCustomObject][ordered]@{
        Id           = [guid]::NewGuid().ToString()
        Datum        = $Now.ToString("yyyy-MM-dd")
        ErfasstAm    = $Now.ToString("o")
        Stunden      = [math]::Round($hours, 2)
        Projekt      = $project
        Beschreibung = $description
    }

    $entries = @(Read-WorkEntries)
    $entries += $entry

    Save-WorkEntries -Entries $entries
    Write-WorkEntriesCsv -Entries $entries
}

function Apply-ControlCommand {
    param(
        $State,
        [datetime]$Now,
        [datetime]$MeasurementStart,
        $Settings
    )

    $command = Load-ControlCommand

    if ($null -eq $command -or [string]::IsNullOrWhiteSpace($command.Id)) {
        return $State
    }

    if ($State.LastControlId -eq $command.Id) {
        Clear-ControlCommand
        return $State
    }

    $commandTime = Get-ControlTime -Command $command -Now $Now

    if ($commandTime.Date -ne $Now.Date) {
        Clear-ControlCommand
        return $State
    }

    switch ($command.Action) {
        "StartManualPause" {
            if (-not [bool]$State.ManualPauseActive) {
                $State.ManualPauseActive = $true
                $State.ManualPauseStartedAt = $commandTime.ToString("o")
                $State.ManualPauseCountedUntil = $commandTime.ToString("o")
            }
        }

        "StopManualPause" {
            if ([bool]$State.ManualPauseActive) {
                $State = Add-ManualPauseUntil -State $State -Until $commandTime -MeasurementStart $MeasurementStart
                $State.ManualPauseActive = $false
                $State.ManualPauseStartedAt = ""
                $State.ManualPauseCountedUntil = ""
            }
        }

        "UpdateToday" {
            $values = $command.Values

            if ($null -ne $values) {
                if ($values.PSObject.Properties.Name -contains "StartTime") {
                    $startTime = Convert-ToTimeText -TimeText ([string]$values.StartTime)

                    if (-not [string]::IsNullOrWhiteSpace($startTime)) {
                        $State.StartTime = $startTime
                        $State.StartPending = $false
                        $startDateTime = Get-StateDateTime -State $State -TimeText $startTime
                        $grossSeconds = ($Now - $startDateTime).TotalSeconds

                        if ($grossSeconds -lt 0) {
                            $grossSeconds = 0
                        }

                        $State.GrossSeconds = $grossSeconds
                    }
                }

                $State.PauseMorningSeconds = Get-NumberValue -Values $values -Name "PauseMorningSeconds" -DefaultValue ([double]$State.PauseMorningSeconds)
                $State.PauseNoonSeconds = Get-NumberValue -Values $values -Name "PauseNoonSeconds" -DefaultValue ([double]$State.PauseNoonSeconds)
                $State.ManualPauseSeconds = Get-NumberValue -Values $values -Name "ManualPauseSeconds" -DefaultValue ([double]$State.ManualPauseSeconds)

                if ($values.PSObject.Properties.Name -contains "Note") {
                    $State.Note = [string]$values.Note
                }

                foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
                    $State = Ensure-Property -Object $State -Name $pause.SecondsProperty -Value 0
                    $State = Ensure-Property -Object $State -Name $pause.CountedProperty -Value ""
                    $State.($pause.CountedProperty) = Get-EditedPauseCountedUntil -Now $Now -Pause $pause
                }

                if ([bool]$State.ManualPauseActive) {
                    $State.ManualPauseCountedUntil = $Now.ToString("o")
                }
            }
        }

        "AddWorkEntry" {
            Add-WorkEntry -Now $commandTime -Values $command.Values
        }
    }

    $State.LastControlId = [string]$command.Id
    Clear-ControlCommand

    return $State
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

function Write-DayToCsv {
    param(
        $State,
        $Settings
    )

    $State = Ensure-PauseStateProperties -State $State -Settings $Settings
    $pauseAuto = Get-AutoPauseSeconds -State $State -Settings $Settings
    $pauseManual = [double]$State.ManualPauseSeconds
    $pauseTotal = $pauseAuto + $pauseManual

    $grossSeconds = [double]$State.GrossSeconds
    $netSeconds = $grossSeconds - $pauseTotal

    if ($netSeconds -lt 0) {
        $netSeconds = 0
    }

    $status = "Arbeit"

    if ([bool]$State.StartPending) {
        $status = "Wartet auf Aktivität"
    }
    elseif ([bool]$State.ManualPauseActive) {
        $status = "Manuelle Pause"
    }


    $netDecimal = $netSeconds / 3600
    $deCulture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
    $activitySummary = Get-WorkEntrySummaryForDate -DateText ([string]$State.Date)
    $pauseColumns = @()

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
        $pauseColumns += Get-PauseCsvColumnName -Pause $pause
    }

    $columns = @(
        "Datum",
        "Start",
        "Ende",
        "Brutto"
    ) + $pauseColumns + @(
        "Pause_Manuell",
        "Pause_Gesamt",
        "Netto",
        "Netto_Stunden_Dezimal",
        "Taetigkeiten_Stunden",
        "Taetigkeiten_Anzahl",
        "Taetigkeiten_Details",
        "Status",
        "Notiz"
    )

    $rowValues = [ordered]@{
        Datum                 = $State.Date
        Start                 = $State.StartTime
        Ende                  = $State.EndTime

        Brutto                = Format-Duration $grossSeconds
    }

    foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings -IncludeDisabled)) {
        $column = Get-PauseCsvColumnName -Pause $pause
        $rowValues[$column] = Format-Duration ([double]$State.($pause.SecondsProperty))
    }

    $rowValues["Pause_Manuell"] = Format-Duration $pauseManual
    $rowValues["Pause_Gesamt"] = Format-Duration $pauseTotal
    $rowValues["Netto"] = Format-Duration $netSeconds
    $rowValues["Netto_Stunden_Dezimal"] = $netDecimal.ToString("N2", $deCulture)
    $rowValues["Taetigkeiten_Stunden"] = ([double]$activitySummary.Hours).ToString("N2", $deCulture)
    $rowValues["Taetigkeiten_Anzahl"] = [string]$activitySummary.Count
    $rowValues["Taetigkeiten_Details"] = [string]$activitySummary.Details
    $rowValues["Status"] = $status
    $rowValues["Notiz"] = [string]$State.Note

    $row = [PSCustomObject]$rowValues

    try {
        $existing = @()

        if (Test-Path $CsvPath) {
            $existing = Import-Csv -Path $CsvPath -Delimiter ";"
            foreach ($existingRow in @($existing)) {
                foreach ($existingColumn in @($existingRow.PSObject.Properties.Name)) {
                    if ($columns -notcontains $existingColumn) {
                        $columns += $existingColumn
                    }
                }
            }
            $existing = $existing | Where-Object {
                $rowDate = Get-WorkDate -DateText ([string]$_.Datum)
                $null -eq $rowDate -or $rowDate.ToString("yyyy-MM-dd") -ne $State.Date
            }
            $existing = @($existing | ForEach-Object { Convert-ToCsvRow -Row $_ -Columns $columns })
        }

        $all = @($existing) + @($row)
        $dedupedByDate = [ordered]@{}
        $rowsWithoutDate = @()

        foreach ($csvRow in $all) {
            $rowDate = Get-WorkDate -DateText ([string]$csvRow.Datum)

            if ($null -eq $rowDate) {
                $rowsWithoutDate += $csvRow
            }
            else {
                $dedupedByDate[$rowDate.ToString("yyyy-MM-dd")] = $csvRow
            }
        }

        $all = @($rowsWithoutDate) + @($dedupedByDate.Values)
        $all = @($all | ForEach-Object {
            $csvRow = Convert-ToCsvRow -Row $_ -Columns $columns
            $rowDate = Get-WorkDate -DateText ([string]$csvRow.Datum)

            if ($null -ne $rowDate) {
                $summary = Get-WorkEntrySummaryForDate -DateText $rowDate.ToString("yyyy-MM-dd")
                $csvRow.Taetigkeiten_Stunden = ([double]$summary.Hours).ToString("N2", $deCulture)
                $csvRow.Taetigkeiten_Anzahl = [string]$summary.Count
                $csvRow.Taetigkeiten_Details = [string]$summary.Details
            }

            $csvRow
        })

        $all |
            Sort-Object {
                $rowDate = Get-WorkDate -DateText ([string]$_.Datum)

                if ($null -eq $rowDate) {
                    return [datetime]::MaxValue
                }

                return $rowDate
            } |
            Export-Csv -Path $CsvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }
    catch {
        Write-TrackerLog ("CSV-Schreibfehler: " + $_.Exception.ToString())
    }
}

function Initialize-State {
    $now = Get-Date
    $loadedState = Load-State

    if ($null -eq $loadedState -or $loadedState.Date -ne $now.ToString("yyyy-MM-dd")) {
        return New-DayState -Now $now
    }

    # Ausgeschaltete Zeit wird bei Neustart nicht nachtraeglich als Arbeit gezaehlt.
    $loadedState.LastTimestamp = $now.ToString("o")

    if ([bool]$loadedState.StartPending) {
        $loadedState.EndTime = ""
    }
    else {
        $loadedState.EndTime = $now.ToString("HH:mm:ss")
    }

    if ([bool]$loadedState.ManualPauseActive) {
        if ([string]::IsNullOrWhiteSpace($loadedState.ManualPauseStartedAt)) {
            $loadedState.ManualPauseStartedAt = $now.ToString("o")
        }

        $loadedState.ManualPauseCountedUntil = $now.ToString("o")
    }

    return $loadedState
}

function Update-StateOnce {
    param(
        $State,
        $Settings
    )

    $now = Get-Date
    $State = Ensure-PauseStateProperties -State $State -Settings $Settings
    $State = Repair-PauseStateForCurrentDay -State $State -Settings $Settings
    $effectiveIntervalSeconds = [math]::Max(1, [int]$Settings.IntervalSeconds)
    $effectiveIdleThresholdSeconds = [math]::Max(1, [int]$Settings.IdleThresholdSeconds)
    $startConfirmSeconds = [math]::Min($effectiveIdleThresholdSeconds, [math]::Max(5, $effectiveIntervalSeconds * 2))
    $idleMilliseconds = Get-IdleMilliseconds
    $idleSeconds = $idleMilliseconds / 1000.0

    if ($State.Date -ne $now.ToString("yyyy-MM-dd")) {
        $wasManualPauseActive = [bool]$State.ManualPauseActive

        Write-DayToCsv -State $State -Settings $Settings

        $newState = New-DayState -Now $now

        if ($wasManualPauseActive) {
            $newState.ManualPauseActive = $true
            $newState.ManualPauseStartedAt = $now.ToString("o")
            $newState.ManualPauseCountedUntil = $now.ToString("o")
        }

        return Ensure-PauseStateProperties -State $newState -Settings $Settings
    }

    if ([bool]$State.StartPending) {
        $State = Apply-ControlCommand -State $State -Now $now -MeasurementStart $now -Settings $Settings

        $hasStartCandidate = -not [string]::IsNullOrWhiteSpace([string]$State.StartCandidateAt)

        if ($idleSeconds -ge $effectiveIdleThresholdSeconds -and -not $hasStartCandidate) {
            $State.StartCandidateAt = ""
            $State.EndTime = ""
            $State.LastTimestamp = $now.ToString("o")
            return $State
        }

        $startCandidate = $now.AddMilliseconds(-1 * [math]::Max(0, [double]$idleMilliseconds))

        if ($startCandidate.Date -ne $now.Date) {
            $startCandidate = $now
        }

        if ([string]::IsNullOrWhiteSpace([string]$State.StartCandidateAt)) {
            $State.StartCandidateAt = $startCandidate.ToString("o")
            $State.EndTime = ""
            $State.LastTimestamp = $now.ToString("o")
            return $State
        }

        try {
            $storedCandidate = [datetime]::Parse([string]$State.StartCandidateAt)

            # Fuer den Tagesstart zaehlt der erste echte Input nach der Idle-Phase.
            # Spaetere Maus-/Tastaturereignisse duerfen diesen Zeitpunkt nicht nach vorne schieben.
            $startCandidate = $storedCandidate
        }
        catch {
            $State.StartCandidateAt = $startCandidate.ToString("o")
            $State.EndTime = ""
            $State.LastTimestamp = $now.ToString("o")
            return $State
        }

        # Tagesstart schnell bestaetigen; die lange Idle-Schwelle gilt nur fuer Pausen.
        if (($now - $startCandidate).TotalSeconds -lt $startConfirmSeconds) {
            $State.EndTime = ""
            $State.LastTimestamp = $now.ToString("o")
            return $State
        }

        $State.StartPending = $false
        $State.StartCandidateAt = ""
        $State.StartTime = $startCandidate.ToString("HH:mm:ss")
        $State.EndTime = $now.ToString("HH:mm:ss")
        $State.GrossSeconds = [math]::Max(0, ($now - $startCandidate).TotalSeconds)
        $State.LastTimestamp = $now.ToString("o")
        return $State
    }

    $last = [datetime]::Parse($State.LastTimestamp)
    $rawDeltaSeconds = ($now - $last).TotalSeconds
    $measurementStart = $last
    $deltaSeconds = $rawDeltaSeconds
    $limitPauseToMeasurementStart = $false

    # Schutz gegen Standby/Ruhezustand: lange Luecken werden nicht als Arbeitszeit
    # und auch nicht als Pause nachgezaehlt.
    if ($rawDeltaSeconds -gt ($effectiveIntervalSeconds * 4)) {
        $deltaSeconds = $effectiveIntervalSeconds
        $measurementStart = $now.AddSeconds(-1 * $effectiveIntervalSeconds)
        $limitPauseToMeasurementStart = $true
    }

    if ($deltaSeconds -lt 0) {
        $deltaSeconds = 0
    }

    $State = Apply-ControlCommand -State $State -Now $now -MeasurementStart $measurementStart -Settings $Settings

    $State.GrossSeconds = [double]$State.GrossSeconds + $deltaSeconds

    if ($idleSeconds -ge $effectiveIdleThresholdSeconds) {
        $idleStart = $now.AddSeconds(-1 * $idleSeconds)

        foreach ($pause in (Get-ArbeitszeitPauseWindows -Settings $Settings)) {
            $State = Add-PauseOverlap `
                -State $State `
                -Now $now `
                -IdleStart $idleStart `
                -MeasurementStart $measurementStart `
                -LimitToMeasurementStart $limitPauseToMeasurementStart `
                -Pause $pause
        }
    }

    if ([bool]$State.ManualPauseActive) {
        $State = Add-ManualPauseUntil -State $State -Until $now -MeasurementStart $measurementStart
    }

    $State.EndTime = $now.ToString("HH:mm:ss")
    $State.LastTimestamp = $now.ToString("o")

    return $State
}

try {
    $settings = Read-ArbeitszeitSettings `
        -BaseDir $BaseDir `
        -IntervalSeconds $IntervalSeconds `
        -IdleThresholdSeconds $IdleThresholdSeconds

    $state = Initialize-State
    $state = Ensure-PauseStateProperties -State $state -Settings $settings

    Save-State $state
    Write-DayToCsv -State $state -Settings $settings

    if ($Once) {
        $settings = Read-ArbeitszeitSettings `
            -BaseDir $BaseDir `
            -IntervalSeconds $IntervalSeconds `
            -IdleThresholdSeconds $IdleThresholdSeconds

        $state = Update-StateOnce -State $state -Settings $settings
        Save-State $state
        Write-DayToCsv -State $state -Settings $settings
        exit
    }

    while ($true) {
        try {
            $settings = Read-ArbeitszeitSettings `
                -BaseDir $BaseDir `
                -IntervalSeconds $IntervalSeconds `
                -IdleThresholdSeconds $IdleThresholdSeconds

            Start-Sleep -Seconds ([math]::Max(1, [int]$settings.IntervalSeconds))

            $settings = Read-ArbeitszeitSettings `
                -BaseDir $BaseDir `
                -IntervalSeconds $IntervalSeconds `
                -IdleThresholdSeconds $IdleThresholdSeconds

            $state = Update-StateOnce -State $state -Settings $settings
            Save-State $state
            Write-DayToCsv -State $state -Settings $settings
        }
        catch {
            Write-TrackerLog ("Tracker-Loop-Fehler: " + $_.Exception.ToString())
            Start-Sleep -Seconds ([math]::Max(1, [int]$IntervalSeconds))

            if ($null -eq $state) {
                $state = Initialize-State
            }
        }
    }
}
catch {
    Write-TrackerLog ("Tracker-Fatal-Fehler: " + $_.Exception.ToString())
    throw
}
finally {
    if ($null -ne $mutex) {
        $mutex.ReleaseMutex()
        $mutex.Dispose()
    }
}
