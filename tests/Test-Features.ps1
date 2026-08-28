$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$trackerPath = Join-Path $repoRoot "Arbeitszeit.ps1"
$displayPath = Join-Path $repoRoot "ArbeitszeitAnzeige.ps1"
$legacyDisplayPath = Join-Path $repoRoot "ArbeitszeitAnzeige.hta"
$settingsPath = Join-Path $repoRoot "ArbeitszeitSettings.ps1"
$oneClickInstallerPath = Join-Path $repoRoot "Install-Arbeitszeit-OneClick.ps1"
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

function Get-RelativeLuminance {
    param([string]$HexColor)

    $hex = $HexColor.TrimStart("#")
    $red = [Convert]::ToInt32($hex.Substring(0, 2), 16) / 255.0
    $green = [Convert]::ToInt32($hex.Substring(2, 2), 16) / 255.0
    $blue = [Convert]::ToInt32($hex.Substring(4, 2), 16) / 255.0
    $linear = @($red, $green, $blue | ForEach-Object {
        if ($_ -le 0.04045) { $_ / 12.92 } else { [math]::Pow(($_ + 0.055) / 1.055, 2.4) }
    })

    return (0.2126 * $linear[0]) + (0.7152 * $linear[1]) + (0.0722 * $linear[2])
}

function Get-ContrastRatio {
    param(
        [string]$FirstColor,
        [string]$SecondColor
    )

    $first = Get-RelativeLuminance $FirstColor
    $second = Get-RelativeLuminance $SecondColor
    return ([math]::Max($first, $second) + 0.05) / ([math]::Min($first, $second) + 0.05)
}

function Convert-BrushToHex {
    param($Brush)

    if ($Brush -isnot [System.Windows.Media.SolidColorBrush]) {
        throw "Für den Kontrasttest wird eine Volltonfarbe benötigt."
    }

    return "#{0:X2}{1:X2}{2:X2}" -f $Brush.Color.R, $Brush.Color.G, $Brush.Color.B
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

Import-FunctionsFromFile -Path $oneClickInstallerPath -Names @(
    "Get-InstallDataScore",
    "Resolve-ArbeitszeitInstallDir"
)

$installerTestRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("ArbeitszeitInstallerTests_" + [guid]::NewGuid().ToString("N"))
$installerTestRoot = [System.IO.Path]::GetFullPath($installerTestRoot)
$installerTempRoot = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath())

if (-not $installerTestRoot.StartsWith($installerTempRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsicherer Installer-Testpfad: $installerTestRoot"
}

$preferredInstall = Join-Path $installerTestRoot "Preferred"
$legacyInstall = Join-Path $installerTestRoot "Legacy"
New-Item -ItemType Directory -Path $preferredInstall, $legacyInstall -Force | Out-Null

try {
    [PSCustomObject]@{ Datum = "2026-08-27" } |
        Export-Csv -LiteralPath (Join-Path $preferredInstall "Arbeitszeiten.csv") -Delimiter ";" -NoTypeInformation
    @(
        [PSCustomObject]@{ Datum = "2026-08-26" },
        [PSCustomObject]@{ Datum = "2026-08-27" }
    ) | Export-Csv -LiteralPath (Join-Path $legacyInstall "Arbeitszeiten.csv") -Delimiter ";" -NoTypeInformation

    $selectedInstall = Resolve-ArbeitszeitInstallDir -RequestedPath "" -PreferredPath $preferredInstall -LegacyPath $legacyInstall
    Assert-Equal ([System.IO.Path]::GetFullPath($legacyInstall)) $selectedInstall "OneClick aktualisiert die Installation mit der größeren Datenhistorie"

    $explicitInstall = Resolve-ArbeitszeitInstallDir -RequestedPath $preferredInstall -PreferredPath $preferredInstall -LegacyPath $legacyInstall
    Assert-Equal ([System.IO.Path]::GetFullPath($preferredInstall)) $explicitInstall "Explizit gewählter Installationsordner hat Vorrang"
}
finally {
    if (Test-Path -LiteralPath $installerTestRoot) {
        Remove-Item -LiteralPath $installerTestRoot -Recurse -Force
    }
}

Import-FunctionsFromFile -Path $trackerPath -Names @(
    "Convert-ToTimeText",
    "Get-OffsetStartDateTime",
    "Ensure-Property",
    "Get-PauseIntervalDateTime",
    "Merge-PauseIntervals",
    "Add-PauseInterval",
    "Format-PauseIntervals",
    "ConvertFrom-PauseIntervalsText",
    "Format-Duration",
    "Get-WorkDate",
    "Get-TodayDateTime",
    "Ensure-PauseStateProperties",
    "Repair-PauseStateForCurrentDay",
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

$duplicateIntervals = @(
    [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Label = "Frühstück"; Start = "2026-08-27T09:06:00"; End = "2026-08-27T09:15:00" },
    [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Label = "Frühstück"; Start = "2026-08-27T09:15:00"; End = "2026-08-27T09:28:00" }
)
$mergedDuplicates = @(Merge-PauseIntervals -Intervals $duplicateIntervals)
Assert-Equal 1 $mergedDuplicates.Count "Direkt angrenzende Pausen derselben Kategorie werden zusammengeführt"
Assert-Equal "09:06" ([datetime]$mergedDuplicates[0].Start).ToString("HH:mm") "Zusammengeführte Pause behält den ersten Beginn"
Assert-Equal "09:28" ([datetime]$mergedDuplicates[0].End).ToString("HH:mm") "Zusammengeführte Pause behält das letzte Ende"

$differentCategories = @(
    $duplicateIntervals[0],
    [PSCustomObject]@{ Kind = "Manual"; Key = "Manual"; Label = "Manuell"; Start = "2026-08-27T09:15:00"; End = "2026-08-27T09:28:00" }
)
Assert-Equal 2 @(Merge-PauseIntervals -Intervals $differentCategories).Count "Angrenzende Pausen verschiedener Kategorien bleiben getrennt"

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

$adjacentCorrectionState = [PSCustomObject]@{
    Date                     = "2026-08-27"
    PauseMorningSeconds      = 1320
    PauseNoonSeconds         = 0
    ManualPauseSeconds       = 0
    PauseMorningCountedUntil = "2026-08-27T09:28:00"
    PauseNoonCountedUntil    = ""
    ManualPauseActive        = $false
    ManualPauseStartedAt     = ""
    ManualPauseCountedUntil  = ""
    PauseIntervals           = @($duplicateIntervals)
}
$repairedAdjacentState = Repair-PauseStateForCurrentDay -State $adjacentCorrectionState -Settings (New-ArbeitszeitDefaultSettings)
Assert-Equal 1 @($repairedAdjacentState.PauseIntervals).Count "Bestehende doppelte Zeitpunkte werden automatisch repariert"

$savedAdjacentState = Apply-CorrectedPauseIntervals `
    -State $adjacentCorrectionState `
    -Intervals @(
        [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Label = "Frühstück"; Start = "09:06"; End = "09:15" },
        [PSCustomObject]@{ Kind = "Auto"; Key = "Morning"; Label = "Frühstück"; Start = "09:15"; End = "09:28" }
    ) `
    -Now ([datetime]"2026-08-27T16:00:00") `
    -Settings (New-ArbeitszeitDefaultSettings)
Assert-Equal 1 @($savedAdjacentState.PauseIntervals).Count "Angrenzende Pausen werden auch beim Korrigieren als ein Zeitraum gespeichert"
Assert-Equal 1320 ([double]$savedAdjacentState.PauseMorningSeconds) "Zusammenführen verändert die Pausendauer nicht"

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
Assert-True -Condition (-not $dialogStyles.Contains('Property="Opacity"')) -Message "Dialog-Hover macht die Beschriftung nicht transparent"

Add-Type -AssemblyName PresentationFramework
Import-FunctionsFromFile -Path $displayPath -Names @(
    "Get-ThemePalette",
    "New-Brush",
    "Get-BrushText",
    "Apply-ThemeRecursive",
    "Merge-PauseIntervals",
    "Set-CorrectionInputAppearance"
)

$displayMergedDuplicates = @(Merge-PauseIntervals -Intervals $duplicateIntervals)
Assert-Equal 1 $displayMergedDuplicates.Count "Anzeige führt mehrere angrenzende Pausen ohne Datentypfehler zusammen"
Assert-Equal "09:06" ([datetime]$displayMergedDuplicates[0].Start).ToString("HH:mm") "Anzeige behält beim Zusammenführen den ersten Beginn"
Assert-Equal "09:28" ([datetime]$displayMergedDuplicates[0].End).ToString("HH:mm") "Anzeige behält beim Zusammenführen das letzte Ende"

$correctionTimeBox = New-Object System.Windows.Controls.TextBox
Set-CorrectionInputAppearance -Control $correctionTimeBox -TimeMaxLength 5
Assert-True -Condition (-not $correctionTimeBox.IsReadOnly) -Message "Zeitfelder im Korrekturfenster sind tatsächlich editierbar"
Assert-Equal 5 $correctionTimeBox.MaxLength "Pausenzeiten akzeptieren das Format HH:mm"
Assert-Equal "#FFFFFF" (Convert-BrushToHex $correctionTimeBox.Background) "Zeitfelder besitzen eine eindeutig erkennbare Eingabefläche"
Assert-Equal "#1C1C1E" (Convert-BrushToHex $correctionTimeBox.Foreground) "Zeitfelder besitzen dunkle Schrift"
Assert-Equal "#1C1C1E" (Convert-BrushToHex $correctionTimeBox.CaretBrush) "Der Cursor im Zeitfeld bleibt sichtbar"

$correctionCategoryBox = New-Object System.Windows.Controls.ComboBox
Set-CorrectionInputAppearance -Control $correctionCategoryBox
$categoryRatio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $correctionCategoryBox.Background) -SecondColor (Convert-BrushToHex $correctionCategoryBox.Foreground)
Assert-True -Condition ($categoryRatio -ge 4.5) -Message "Kategorie-Auswahl besitzt auch im Darkmode ausreichenden Kontrast"
Assert-True -Condition ($null -ne $correctionCategoryBox.ItemContainerStyle) -Message "Einträge der Kategorie-Auswahl erhalten explizite lesbare Farben"

$dialogStylesReader = New-Object System.Xml.XmlNodeReader $dialogStylesXml
$dialogResources = [System.Windows.Markup.XamlReader]::Load($dialogStylesReader)
Assert-True -Condition $true -Message "Dialog-Styles lassen sich von WPF laden"

foreach ($styleName in @("DialogButton", "DialogSecondaryButton", "DialogIconButton")) {
    $button = New-Object System.Windows.Controls.Button
    $button.Style = $dialogResources[$styleName]
    $ratio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $button.Background) -SecondColor (Convert-BrushToHex $button.Foreground)
    Assert-True -Condition ($ratio -ge 4.5) -Message "$styleName besitzt zur Laufzeit ausreichend Kontrast"
}

$disabledDialogButton = New-Object System.Windows.Controls.Button
$disabledDialogButton.Style = $dialogResources["DialogButton"]
$disabledDialogButton.IsEnabled = $false
$disabledDialogButton.ApplyTemplate() | Out-Null
$disabledDialogRoot = $disabledDialogButton.Template.FindName("Root", $disabledDialogButton)
$disabledDialogRatio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $disabledDialogRoot.Background) -SecondColor (Convert-BrushToHex $disabledDialogButton.Foreground)
Assert-True -Condition ($disabledDialogRatio -ge 4.5) -Message "Deaktivierte Dialogbuttons besitzen zur Laufzeit ausreichend Kontrast"

$displaySource = Get-Content -LiteralPath $displayPath -Raw
Assert-True -Condition $displaySource.Contains('x:Name="PauseSummaryText"') -Message "Korrekturfenster zeigt die live berechnete Pausensumme"
Assert-True -Condition $displaySource.Contains('x:Name="EmptyPausePanel"') -Message "Korrekturfenster besitzt einen klaren Leerzustand"

$contrastPairs = @(
    @("#005BBB", "#FFFFFF", "Primärbuttons"),
    @("#1F7A3D", "#FFFFFF", "Weiter-Button"),
    @("#3A3A3C", "#FFFFFF", "Dunkle Buttons"),
    @("#5145CD", "#FFFFFF", "Bericht-Button"),
    @("#E4E4EA", "#1C1C1E", "Sekundärbuttons"),
    @("#B42318", "#FFFFFF", "Löschbuttons"),
    @("#D7D7DC", "#4A4A4F", "Deaktivierte Buttons")
)

foreach ($pair in $contrastPairs) {
    $ratio = Get-ContrastRatio -FirstColor $pair[0] -SecondColor $pair[1]
    Assert-True -Condition ($ratio -ge 4.5) -Message ("{0} erfüllen WCAG AA ({1:N2}:1)" -f $pair[2], $ratio)
}

$mainXamlReader = New-Object System.Xml.XmlNodeReader ([xml]$mainXaml)
$mainWindowForContrastTest = [System.Windows.Markup.XamlReader]::Load($mainXamlReader)

foreach ($buttonName in @("SetupButton", "PauseButton", "ResumeButton", "EditButton", "CsvButton", "WeekButton", "ThemeButton", "ActivitySaveButton")) {
    $button = $mainWindowForContrastTest.FindName($buttonName)
    Assert-True -Condition ($null -ne $button) -Message "$buttonName wurde für den Kontrasttest gefunden"
    $ratio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $button.Background) -SecondColor (Convert-BrushToHex $button.Foreground)
    Assert-True -Condition ($ratio -ge 4.5) -Message "$buttonName besitzt zur Laufzeit ausreichend Kontrast"
}

# Die echte Theme-Rekursion ausführen: Sie darf die internen Root-Border der
# Button-Templates nicht wie normale Karten umfärben.
foreach ($buttonName in @("SetupButton", "PauseButton", "ResumeButton", "EditButton", "CsvButton", "WeekButton", "ThemeButton", "ActivitySaveButton")) {
    $mainWindowForContrastTest.FindName($buttonName).ApplyTemplate() | Out-Null
}

Apply-ThemeRecursive `
    -Element $mainWindowForContrastTest.FindName("RootGrid") `
    -Palette (Get-ThemePalette -Theme "Light")

foreach ($buttonName in @("SetupButton", "PauseButton", "ResumeButton", "EditButton", "CsvButton", "WeekButton", "ThemeButton", "ActivitySaveButton")) {
    $button = $mainWindowForContrastTest.FindName($buttonName)
    $buttonRoot = $button.Template.FindName("Root", $button)
    $ratio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $buttonRoot.Background) -SecondColor (Convert-BrushToHex $button.Foreground)
    Assert-True -Condition ($ratio -ge 4.5) -Message "$buttonName bleibt nach vollständigem Lightmode-Theming lesbar"
}

Assert-Equal "#3A3A3C" (Convert-BrushToHex $mainWindowForContrastTest.FindName("EditButton").Template.FindName("Root", $mainWindowForContrastTest.FindName("EditButton")).Background) "Korrigieren-Button bleibt im Lightmode dunkel"
Assert-Equal "#3A3A3C" (Convert-BrushToHex $mainWindowForContrastTest.FindName("CsvButton").Template.FindName("Root", $mainWindowForContrastTest.FindName("CsvButton")).Background) "CSV-Button bleibt im Lightmode dunkel"
Assert-Equal "#3A3A3C" (Convert-BrushToHex $mainWindowForContrastTest.FindName("ThemeButton").Template.FindName("Root", $mainWindowForContrastTest.FindName("ThemeButton")).Background) "Theme-Button bleibt im Lightmode dunkel"

$disabledMainButton = $mainWindowForContrastTest.FindName("PauseButton")
$disabledMainButton.IsEnabled = $false
$disabledMainButton.ApplyTemplate() | Out-Null
$disabledMainRoot = $disabledMainButton.Template.FindName("Root", $disabledMainButton)
$disabledMainRatio = Get-ContrastRatio -FirstColor (Convert-BrushToHex $disabledMainRoot.Background) -SecondColor (Convert-BrushToHex $disabledMainButton.Foreground)
Assert-True -Condition ($disabledMainRatio -ge 4.5) -Message "Deaktivierte Hauptbuttons besitzen zur Laufzeit ausreichend Kontrast"

Assert-True -Condition (-not $mainXaml.Contains('Property="Opacity"')) -Message "Hauptfenster-Hover macht die Beschriftung nicht transparent"
Assert-True -Condition $mainXaml.Contains('Style="{StaticResource PillSuccessButton}"') -Message "Weiter-Button verwendet eine kontrastgeprüfte Rolle"
Assert-True -Condition $mainXaml.Contains('Style="{StaticResource PillSurfaceButton}"') -Message "Heller Setup-Button verwendet dunkle Schrift"
Assert-True -Condition $displaySource.Contains('$dialog.FindName("CloseButton").Style = $dialog.Resources["DialogSecondaryButton"]') -Message "Bericht-Schließen-Button verwendet dunkle Schrift auf hellem Grund"
Assert-True -Condition $displaySource.Contains('$dialog.FindName("ExportButton").Style = $dialog.Resources["DialogButton"]') -Message "Bericht-Export-Button verwendet den kontrastgeprüften Primärstil"
Assert-True -Condition $displaySource.Contains('$addPauseButton.Style = $secondaryButtonStyle') -Message "Setup-Pause-hinzufügen-Button verwendet den Sekundärstil"
Assert-True -Condition $displaySource.Contains('$addPauseButton.Style = $dialog.Resources["DialogSecondaryButton"]') -Message "Korrektur-Pause-hinzufügen-Button verwendet den Sekundärstil"
Assert-Equal 2 ([regex]::Matches($displaySource, [regex]::Escape('$removeButton.Style = $iconButtonStyle')).Count) "Beide dynamischen Löschbuttons verwenden den kontrastgeprüften Icon-Stil"

$legacyDisplaySource = Get-Content -LiteralPath $legacyDisplayPath -Raw
Assert-True -Condition $legacyDisplaySource.Contains('background: #005bbb;') -Message "Legacy-Pause-Button verwendet die kontrastgeprüfte Primärfarbe"
Assert-True -Condition $legacyDisplaySource.Contains('button:disabled') -Message "Legacy-Buttons besitzen einen lesbaren Disabled-Zustand"

Write-Host ("Alle Feature-Tests erfolgreich: {0} Assertions" -f $script:Assertions)
