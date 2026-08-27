$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$trackerPath = Join-Path $repoRoot "Arbeitszeit.ps1"
$displayPath = Join-Path $repoRoot "ArbeitszeitAnzeige.ps1"
$settingsPath = Join-Path $repoRoot "ArbeitszeitSettings.ps1"
$script:Assertions = 0

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )

    $script:Assertions++

    if (-not $Condition) {
        throw "Assertion fehlgeschlagen: $Message"
    }
}

function Assert-Equal {
    param(
        $Expected,
        $Actual,
        [string]$Message
    )

    Assert-True -Condition ($Expected -eq $Actual) -Message ("$Message (erwartet: '$Expected', tatsächlich: '$Actual')")
}

function Import-FunctionsFromFile {
    param(
        [string]$Path,
        [string[]]$Names
    )

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $Path,
        [ref]$tokens,
        [ref]$errors
    )

    if ($errors.Count -gt 0) {
        throw "Syntaxfehler in $Path"
    }

    foreach ($name in $Names) {
        $definition = $ast.FindAll(
            {
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq $name
            },
            $true
        ) | Select-Object -First 1

        if ($null -eq $definition) {
            throw "Testfunktion nicht gefunden: $name"
        }

        Invoke-Expression ("function global:" + $name + " " + $definition.Body.Extent.Text)
    }
}

. $settingsPath

$legacySettings = [PSCustomObject]@{
    IntervalSeconds      = 7
    IdleThresholdSeconds = 30
    TargetNetHours       = 8
    WeekTargetHours      = 40
    AlwaysOnTop          = $false
    NotifyTargetReached  = $true
    Theme                = "Light"
    PauseWindows         = @()
}
$normalizedSettings = Ensure-ArbeitszeitSettings -Settings $legacySettings
Assert-Equal 0 $normalizedSettings.StartOffsetMinutes "Alte Settings erhalten den kompatiblen Offset 0"
Assert-Equal 40 ([double]$normalizedSettings.WeekTargetHours) "Die Regelarbeitszeit bleibt erhalten"

$legacySettings | Add-Member -NotePropertyName StartOffsetMinutes -NotePropertyValue 3
$normalizedSettings = Ensure-ArbeitszeitSettings -Settings $legacySettings
Assert-Equal 3 $normalizedSettings.StartOffsetMinutes "Start-Offset wird aus Settings übernommen"

Import-FunctionsFromFile -Path $trackerPath -Names @(
    "Convert-ToTimeText",
    "Get-OffsetStartDateTime",
    "Ensure-Property",
    "Get-PauseIntervalDateTime",
    "Add-PauseInterval",
    "Format-PauseIntervals",
    "ConvertFrom-PauseIntervalsText",
    "Format-Duration",
    "Get-WorkDate",
    "Get-TodayDateTime",
    "Ensure-PauseStateProperties",
    "Get-AutoPauseSeconds",
    "Get-PauseCsvColumnName",
    "Get-EditedPauseCountedUntil",
    "Apply-CorrectedPauseIntervals",
    "Convert-ToCsvRow",
    "Write-DayToCsv"
)

$candidate = [datetime]"2026-08-27T07:05:00"
$offsetStart = Get-OffsetStartDateTime -StartCandidate $candidate -Now $candidate -StartOffsetMinutes 3
Assert-Equal "07:02:00" $offsetStart.ToString("HH:mm:ss") "Drei Minuten Start-Offset werden abgezogen"

$midnightCandidate = [datetime]"2026-08-27T00:01:00"
$clampedStart = Get-OffsetStartDateTime -StartCandidate $midnightCandidate -Now $midnightCandidate -StartOffsetMinutes 3
Assert-Equal "00:00:00" $clampedStart.ToString("HH:mm:ss") "Offset wird an der Tagesgrenze begrenzt"

$state = [PSCustomObject]@{ PauseIntervals = @() }
$state = Add-PauseInterval -State $state -Kind "Auto" -Key "Morning" -Label "Frühstück" -From ([datetime]"2026-08-27T09:01:00") -To ([datetime]"2026-08-27T09:05:00")
$state = Add-PauseInterval -State $state -Kind "Auto" -Key "Morning" -Label "Frühstück" -From ([datetime]"2026-08-27T09:05:00") -To ([datetime]"2026-08-27T09:12:00")
Assert-Equal 1 @($state.PauseIntervals).Count "Zusammenhängende Messungen werden zu einer Pause verbunden"
$pauseText = Format-PauseIntervals -Intervals $state.PauseIntervals
Assert-Equal "09:01-09:12 (Frühstück)" $pauseText "Absolute Pausenzeit wird lesbar formatiert"
$importedIntervals = @(ConvertFrom-PauseIntervalsText -Text "9:01-9:12 (Frühstück)" -Date ([datetime]"2026-08-27"))
Assert-Equal 1 $importedIntervals.Count "Pausenzeiten mit einstelliger Stunde sind lesbar"

$correctionState = [PSCustomObject]@{
    Date                     = "2026-08-27"
    PauseMorningSeconds      = 1200
    PauseNoonSeconds         = 1800
    ManualPauseSeconds       = 600
    PauseMorningCountedUntil = ""
    PauseNoonCountedUntil    = ""
    ManualPauseActive        = $false
    ManualPauseStartedAt     = ""
    ManualPauseCountedUntil  = ""
    PauseIntervals           = @()
}
$corrected = Apply-CorrectedPauseIntervals `
    -State $correctionState `
    -Intervals @(
        [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Label = "Frühstück"; Start = "09:01"; End = "09:12" },
        [PSCustomObject]@{ Kind = "Manual"; Key = "Manual"; Label = "Manuell"; Start = "12:05"; End = "12:20" }
    ) `
    -Now ([datetime]"2026-08-27T16:00:00") `
    -Settings (New-ArbeitszeitDefaultSettings)
Assert-Equal 660 ([double]$corrected.PauseMorningSeconds) "Automatische Pausendauer wird aus Von-bis berechnet"
Assert-Equal 0 ([double]$corrected.PauseNoonSeconds) "Entfernte Pausenzeiträume werden auf null gesetzt"
Assert-Equal 900 ([double]$corrected.ManualPauseSeconds) "Manuelle Pausendauer wird aus Von-bis berechnet"
Assert-Equal 2 @($corrected.PauseIntervals).Count "Korrigierte Von-bis-Zeiträume werden gespeichert"

$overlapState = [PSCustomObject]@{
    Date                     = "2026-08-27"
    PauseMorningSeconds      = 321
    PauseNoonSeconds         = 0
    ManualPauseSeconds       = 0
    PauseMorningCountedUntil = ""
    PauseNoonCountedUntil    = ""
    ManualPauseActive        = $false
    PauseIntervals           = @()
}
$overlapResult = Apply-CorrectedPauseIntervals `
    -State $overlapState `
    -Intervals @(
        [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Start = "09:00"; End = "09:20" },
        [PSCustomObject]@{ Kind = "Manual"; Key = "Manual"; Start = "09:10"; End = "09:30" }
    ) `
    -Now ([datetime]"2026-08-27T16:00:00") `
    -Settings (New-ArbeitszeitDefaultSettings)
Assert-Equal 321 ([double]$overlapResult.PauseMorningSeconds) "Überlappende Korrekturen verändern den Zustand nicht"

$testDir = Join-Path ([System.IO.Path]::GetTempPath()) ("ArbeitszeitTests_" + [guid]::NewGuid().ToString("N"))
$testDir = [System.IO.Path]::GetFullPath($testDir)
$tempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())

if (-not $testDir.StartsWith($tempRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsicherer Testpfad: $testDir"
}

New-Item -ItemType Directory -Path $testDir -Force | Out-Null

try {
    $script:CsvPath = Join-Path $testDir "Arbeitszeiten.csv"
    $script:TrackerLogPath = Join-Path $testDir "tracker.log"

    function Write-TrackerLog { param([string]$Message) }
    function Get-WorkEntrySummaryForDate {
        param([string]$DateText)
        return [PSCustomObject]@{ Hours = 0.0; Count = 0; Details = "" }
    }

    [PSCustomObject][ordered]@{
        Datum                 = "2026-08-25"
        Start                 = "08:00:00"
        Ende                  = "16:30:00"
        Brutto                = "08:30:00"
        Pause_08_55_09_35     = "00:10:00"
        Pause_11_55_12_45     = "00:20:00"
        Pause_Manuell         = "00:00:00"
        Pause_Gesamt          = "00:30:00"
        Netto                 = "08:00:00"
        Netto_Stunden_Dezimal = "8,00"
        Status                = "Arbeit"
        Notiz                 = "Altbestand"
        Alte_Spalte           = "bleibt"
    } | Export-Csv -LiteralPath $script:CsvPath -Delimiter ";" -NoTypeInformation -Encoding UTF8

    $csvState = [PSCustomObject]@{
        Date                     = "2026-08-27"
        StartTime                = "07:02:00"
        EndTime                  = "16:00:00"
        StartPending             = $false
        GrossSeconds             = 32280
        PauseMorningSeconds      = 660
        PauseNoonSeconds         = 0
        ManualPauseSeconds       = 0
        PauseMorningCountedUntil = "2026-08-27T09:12:00"
        PauseNoonCountedUntil    = ""
        ManualPauseActive        = $false
        PauseIntervals           = @($state.PauseIntervals)
        Note                     = ""
    }

    Write-DayToCsv -State $csvState -Settings (New-ArbeitszeitDefaultSettings)
    $rows = @(Import-Csv -LiteralPath $script:CsvPath -Delimiter ";")
    Assert-Equal 2 $rows.Count "Alte und neue CSV-Zeile bleiben gemeinsam erhalten"
    Assert-Equal "bleibt" ([string]$rows[0].Alte_Spalte) "Unbekannte alte Spalten bleiben erhalten"
    Assert-Equal "" ([string]$rows[0].Pausen_Zeitraeume) "Neue Spalte wird bei Altdaten leer ergänzt"
    Assert-Equal "09:01-09:12 (Frühstück)" ([string]$rows[1].Pausen_Zeitraeume) "Neue CSV-Zeile enthält absolute Pausenzeit"
}
finally {
    if (Test-Path -LiteralPath $testDir) {
        Remove-Item -LiteralPath $testDir -Recurse -Force
    }
}

Import-FunctionsFromFile -Path $displayPath -Names @(
    "Convert-DurationTextToSeconds",
    "Get-WorkDate",
    "Get-ExpectedWorkSeconds",
    "Get-MonthReportData",
    "Format-CompactDuration",
    "Get-DayNameShort",
    "ConvertTo-ReportHtml"
)

$today = (Get-Date).Date
$monthStart = [datetime]::new($today.Year, $today.Month, 1)
$script:TestRows = @(
    [PSCustomObject]@{
        Datum         = $monthStart.ToString("yyyy-MM-dd")
        Netto         = "08:00:00"
        Brutto        = "08:30:00"
        Pause_Gesamt  = "00:30:00"
        Status        = "Arbeit"
    }
)

function Read-WorkCsvRows { return @($script:TestRows) }
function Get-WorkEntrySummaryForDate {
    param([string]$DateText)
    return [PSCustomObject]@{ Hours = 0.0; Count = 0; Details = "" }
}

$values = [PSCustomObject]@{
    HasState = $false
    Settings = [PSCustomObject]@{ WeekTargetHours = 40.0 }
}
$monthData = Get-MonthReportData -Values $values
Assert-Equal 28800 ([double]$monthData.TotalNetSeconds) "Legacy-Nettozeit fließt in die Monatsauswertung ein"
Assert-Equal $monthStart.ToString("yyyy-MM-dd") $monthData.OverallStart.ToString("yyyy-MM-dd") "Gesamtsaldo beginnt beim ersten Datensatz"
Assert-True -Condition ($monthData.TargetToDateSeconds -ge 0) -Message "Monatssoll wird berechnet"

$weekData = [PSCustomObject]@{
    WeekStart          = $today
    WeekEnd            = $today.AddDays(6)
    Days               = @([PSCustomObject]@{ Date = $today; NetSeconds = 28800; ActivityHours = 0; PauseSeconds = 1800; PauseRanges = "09:01-09:12 (Frühstück)"; Status = "Arbeit"; ActivityText = "" })
    TotalNetSeconds    = 28800
    WeekTargetSeconds = 144000
    BalanceSeconds     = -115200
    TotalActivityHours = 0
    BestDay            = $null
    AverageSeconds     = 28800
    TotalPauseSeconds  = 1800
    PauseRatio         = 5.0
}
$html = ConvertTo-ReportHtml -Data $weekData -MonthData $monthData
Assert-True -Condition $html.Contains("Pausenzeiten") -Message "Bericht enthält absolute Pausenzeiten"
Assert-True -Condition $html.Contains("berstunden gesamt") -Message "Bericht enthält Gesamtüberstunden"
Assert-True -Condition $html.Contains($monthData.MonthName) -Message "Bericht enthält die Monatsübersicht"

Import-FunctionsFromFile -Path $displayPath -Names @("Open-WeekWindow", "Open-CorrectionWindow")

function Get-WeekReportData {
    return [PSCustomObject]@{
        Values = [PSCustomObject]@{
            Settings = [PSCustomObject]@{ Theme = "Light" }
            GrossSeconds = 0
            TotalPauseSeconds = 0
            AutoPauseSeconds = 0
            ManualPauseSeconds = 0
        }
        BestDay = $null
        BalanceSeconds = 0
        WeekStart = $today
        WeekEnd = $today.AddDays(6)
        TotalNetSeconds = 0
        WeekTargetSeconds = 144000
        TotalActivityHours = 0
        AverageSeconds = 0
        PauseRatio = 0
        Days = @()
    }
}
function Get-MonthReportData {
    param($Values)
    return [PSCustomObject]@{
        MonthName = "August 2026"
        PeriodEnd = $today
        WeekTargetHours = 40
        TotalNetSeconds = 0
        TargetToDateSeconds = 0
        BalanceSeconds = 0
        OverallBalanceSeconds = 0
    }
}
function Get-ThemePalette {
    param([string]$Theme)
    return [PSCustomObject]@{ Window = "#FFFFFF"; Card = "#FFFFFF"; Border = "#CCCCCC"; Primary = "#111111"; Secondary = "#666666"; Soft = "#EEEEEE" }
}
function Convert-FromXaml {
    param([string]$Xaml)
    [xml]$Xaml | Out-Null
    throw "__XAML_VALID__"
}

try {
    Open-WeekWindow -Owner $null
    throw "Berichts-XAML wurde nicht geprüft."
}
catch {
    if ($_.Exception.Message -ne "__XAML_VALID__") {
        throw
    }
}

function Get-LiveValues {
    return [PSCustomObject]@{
        HasState = $true
        Date = $today.ToString("yyyy-MM-dd")
        StartTime = "07:02:00"
        PauseIntervalsText = "09:01-09:12 (Frühstück)"
        TotalPauseSeconds = 660
        Note = ""
        Settings = New-ArbeitszeitDefaultSettings
    }
}
function Read-State {
    return [PSCustomObject]@{
        PauseIntervals = @(
            [PSCustomObject]@{
                Kind = "Auto"
                Key = "Morning"
                Label = "Frühstück"
                Start = $today.AddHours(9).AddMinutes(1).ToString("o")
                End = $today.AddHours(9).AddMinutes(12).ToString("o")
            }
        )
    }
}

try {
    Open-CorrectionWindow -Owner $null
    throw "Korrektur-XAML wurde nicht geprüft."
}
catch {
    if ($_.Exception.Message -ne "__XAML_VALID__") {
        throw
    }
}

$tokens = $null
$errors = $null
$displayAst = [System.Management.Automation.Language.Parser]::ParseFile($displayPath, [ref]$tokens, [ref]$errors)
$mainXamlAssignment = $displayAst.FindAll(
    {
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -eq "mainXaml"
    },
    $true
) | Select-Object -First 1

if ($null -eq $mainXamlAssignment) {
    throw "Hauptfenster-XAML wurde nicht gefunden."
}

Invoke-Expression $mainXamlAssignment.Extent.Text
[xml]$mainXaml | Out-Null
Assert-True -Condition $true -Message "Hauptfenster-XAML ist gültig"
Assert-True -Condition $mainXaml.Contains('TextElement.Foreground="{TemplateBinding Foreground}"') -Message "Hauptfenster-Buttons übernehmen ihre Textfarbe sichtbar"

$dialogStylesAssignment = $displayAst.FindAll(
    {
        param($node)
        $node -is [System.Management.Automation.Language.AssignmentStatementAst] -and
            $node.Left -is [System.Management.Automation.Language.VariableExpressionAst] -and
            $node.Left.VariablePath.UserPath -eq "dialogStyles"
    },
    $true
) | Select-Object -First 1

if ($null -eq $dialogStylesAssignment) {
    throw "Dialog-Styles wurden nicht gefunden."
}

Invoke-Expression $dialogStylesAssignment.Extent.Text
[xml]$dialogStylesXml = $dialogStyles
Assert-True -Condition $dialogStyles.Contains('TextElement.Foreground="{TemplateBinding Foreground}"') -Message "Dialog-Buttons übernehmen ihre Textfarbe sichtbar"
Assert-True -Condition $dialogStyles.Contains('x:Key="DialogSecondaryButton"') -Message "Sekundäre Lightmode-Buttons besitzen einen kontrastreichen Stil"
Assert-True -Condition $dialogStyles.Contains('x:Key="DialogIconButton"') -Message "Löschaktionen besitzen einen eindeutigen Icon-Stil"

Add-Type -AssemblyName PresentationFramework
$dialogStylesReader = New-Object System.Xml.XmlNodeReader $dialogStylesXml
[System.Windows.Markup.XamlReader]::Load($dialogStylesReader) | Out-Null
Assert-True -Condition $true -Message "Dialog-Styles lassen sich von WPF laden"

$displaySource = Get-Content -LiteralPath $displayPath -Raw
Assert-True -Condition $displaySource.Contains('x:Name="PauseSummaryText"') -Message "Korrekturfenster zeigt die live berechnete Pausensumme"
Assert-True -Condition $displaySource.Contains('x:Name="EmptyPausePanel"') -Message "Korrekturfenster besitzt einen klaren Leerzustand"

Write-Host ("Alle Feature-Tests erfolgreich: {0} Assertions" -f $script:Assertions)
