function Convert-ArbeitszeitTimeText {
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
            return $parsed.ToString("HH:mm")
        }
    }

    return $null
}

function Convert-ArbeitszeitPauseKey {
    param(
        [string]$Key,
        [string]$Fallback = "Pause"
    )

    $source = $Key

    if ([string]::IsNullOrWhiteSpace($source)) {
        $source = $Fallback
    }

    $source = $source.Replace("ß", "ss").Normalize([System.Text.NormalizationForm]::FormD)
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $source.ToCharArray()) {
        $category = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)

        if ($category -eq [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            continue
        }

        if ([char]::IsLetterOrDigit($char)) {
            [void]$builder.Append($char)
        }
    }

    $normalized = $builder.ToString()

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = "Pause"
    }

    if ([char]::IsDigit($normalized[0])) {
        $normalized = "Pause$normalized"
    }

    return $normalized
}

function New-ArbeitszeitPauseWindow {
    param(
        [string]$Key,
        [string]$Label,
        [string]$Start = "12:00",
        [string]$End = "12:15",
        [bool]$Enabled = $true
    )

    $normalizedKey = Convert-ArbeitszeitPauseKey -Key $Key -Fallback $Label

    return [PSCustomObject][ordered]@{
        Key             = $normalizedKey
        Label           = $Label
        Start           = $Start
        End             = $End
        Enabled         = $Enabled
        SecondsProperty = "Pause${normalizedKey}Seconds"
        CountedProperty = "Pause${normalizedKey}CountedUntil"
    }
}

function New-ArbeitszeitDefaultSettings {
    param(
        [int]$IntervalSeconds = 5,
        [int]$IdleThresholdSeconds = 20
    )

    return [PSCustomObject][ordered]@{
        IntervalSeconds     = $IntervalSeconds
        IdleThresholdSeconds = $IdleThresholdSeconds
        TargetNetHours      = 8.0
        WeekTargetHours     = 40.0
        AlwaysOnTop         = $true
        NotifyTargetReached = $true
        Theme               = "Light"
        PauseWindows        = @(
            [PSCustomObject][ordered]@{
                Key             = "Morning"
                Label           = "Frühstück"
                Start           = "08:55"
                End             = "09:35"
                Enabled         = $true
                SecondsProperty = "PauseMorningSeconds"
                CountedProperty = "PauseMorningCountedUntil"
            },
            [PSCustomObject][ordered]@{
                Key             = "Noon"
                Label           = "Mittag"
                Start           = "11:55"
                End             = "12:45"
                Enabled         = $true
                SecondsProperty = "PauseNoonSeconds"
                CountedProperty = "PauseNoonCountedUntil"
            }
        )
    }
}

function Get-ArbeitszeitSettingNumber {
    param(
        $Settings,
        [string]$Name,
        [double]$DefaultValue,
        [double]$MinValue,
        [double]$MaxValue
    )

    if ($null -eq $Settings -or $Settings.PSObject.Properties.Name -notcontains $Name) {
        return $DefaultValue
    }

    if ($null -eq $Settings.$Name -or [string]::IsNullOrWhiteSpace([string]$Settings.$Name)) {
        return $DefaultValue
    }

    try {
        $value = [double]$Settings.$Name
    }
    catch {
        $value = $DefaultValue
    }

    if ($value -lt $MinValue) {
        return $MinValue
    }

    if ($value -gt $MaxValue) {
        return $MaxValue
    }

    return $value
}

function Get-ArbeitszeitPauseWindow {
    param(
        $Source,
        $Fallback
    )

    if ($null -eq $Source) {
        $Source = $Fallback
    }

    $start = Convert-ArbeitszeitTimeText ([string]$Source.Start)
    $end = Convert-ArbeitszeitTimeText ([string]$Source.End)

    if ([string]::IsNullOrWhiteSpace($start)) {
        $start = $Fallback.Start
    }

    if ([string]::IsNullOrWhiteSpace($end)) {
        $end = $Fallback.End
    }

    $enabled = $true

    try {
        $enabled = [System.Convert]::ToBoolean($Source.Enabled)
    }
    catch {
        $enabled = $true
    }

    $label = [string]$Source.Label

    if ([string]::IsNullOrWhiteSpace($label)) {
        $label = $Fallback.Label
    }

    if ($label -eq "Fruehstueck") {
        $label = "Frühstück"
    }

    $key = Convert-ArbeitszeitPauseKey -Key ([string]$Source.Key) -Fallback ([string]$Fallback.Key)
    $secondsProperty = [string]$Source.SecondsProperty
    $countedProperty = [string]$Source.CountedProperty

    if ([string]::IsNullOrWhiteSpace($secondsProperty)) {
        $secondsProperty = [string]$Fallback.SecondsProperty
    }

    if ([string]::IsNullOrWhiteSpace($countedProperty)) {
        $countedProperty = [string]$Fallback.CountedProperty
    }

    if ([string]::IsNullOrWhiteSpace($secondsProperty)) {
        $secondsProperty = "Pause${key}Seconds"
    }

    if ([string]::IsNullOrWhiteSpace($countedProperty)) {
        $countedProperty = "Pause${key}CountedUntil"
    }

    return [PSCustomObject][ordered]@{
        Key             = $key
        Label           = $label
        Start           = $start
        End             = $end
        Enabled         = $enabled
        SecondsProperty = $secondsProperty
        CountedProperty = $countedProperty
    }
}

function Ensure-ArbeitszeitSettings {
    param(
        $Settings,
        [int]$IntervalSeconds = 5,
        [int]$IdleThresholdSeconds = 20
    )

    $defaults = New-ArbeitszeitDefaultSettings `
        -IntervalSeconds $IntervalSeconds `
        -IdleThresholdSeconds $IdleThresholdSeconds

    if ($null -eq $Settings) {
        return $defaults
    }

    $pauseSources = @($Settings.PauseWindows)
    $pauseWindows = @()

    foreach ($fallback in @($defaults.PauseWindows)) {
        $source = $pauseSources | Where-Object { $_.Key -eq $fallback.Key } | Select-Object -First 1
        $pauseWindows += Get-ArbeitszeitPauseWindow -Source $source -Fallback $fallback
    }

    foreach ($source in $pauseSources) {
        if ($null -eq $source -or $null -eq $source.Key) {
            continue
        }

        $key = Convert-ArbeitszeitPauseKey -Key ([string]$source.Key)

        if (@($defaults.PauseWindows | Where-Object { $_.Key -eq $key }).Count -gt 0) {
            continue
        }

        $fallback = New-ArbeitszeitPauseWindow `
            -Key $key `
            -Label ([string]$source.Label) `
            -Start "12:00" `
            -End "12:15" `
            -Enabled $true

        $pauseWindows += Get-ArbeitszeitPauseWindow -Source $source -Fallback $fallback
    }

    $alwaysOnTop = $defaults.AlwaysOnTop

    if ($Settings.PSObject.Properties.Name -contains "AlwaysOnTop") {
        try {
            $alwaysOnTop = [System.Convert]::ToBoolean($Settings.AlwaysOnTop)
        }
        catch {
            $alwaysOnTop = $defaults.AlwaysOnTop
        }
    }

    $notifyTargetReached = $defaults.NotifyTargetReached

    if ($Settings.PSObject.Properties.Name -contains "NotifyTargetReached") {
        try {
            $notifyTargetReached = [System.Convert]::ToBoolean($Settings.NotifyTargetReached)
        }
        catch {
            $notifyTargetReached = $defaults.NotifyTargetReached
        }
    }

    $theme = [string]$Settings.Theme

    if ($theme -notin @("Light", "Dark")) {
        $theme = $defaults.Theme
    }

    return [PSCustomObject][ordered]@{
        IntervalSeconds     = [int](Get-ArbeitszeitSettingNumber -Settings $Settings -Name "IntervalSeconds" -DefaultValue $defaults.IntervalSeconds -MinValue 1 -MaxValue 60)
        IdleThresholdSeconds = [int](Get-ArbeitszeitSettingNumber -Settings $Settings -Name "IdleThresholdSeconds" -DefaultValue $defaults.IdleThresholdSeconds -MinValue 1 -MaxValue 600)
        TargetNetHours      = [double](Get-ArbeitszeitSettingNumber -Settings $Settings -Name "TargetNetHours" -DefaultValue $defaults.TargetNetHours -MinValue 0.5 -MaxValue 24)
        WeekTargetHours     = [double](Get-ArbeitszeitSettingNumber -Settings $Settings -Name "WeekTargetHours" -DefaultValue $defaults.WeekTargetHours -MinValue 1 -MaxValue 100)
        AlwaysOnTop         = $alwaysOnTop
        NotifyTargetReached = $notifyTargetReached
        Theme               = $theme
        PauseWindows        = $pauseWindows
    }
}

function Read-ArbeitszeitSettings {
    param(
        [string]$BaseDir,
        [int]$IntervalSeconds = 5,
        [int]$IdleThresholdSeconds = 20
    )

    $settingsPath = Join-Path $BaseDir "settings.json"
    $settings = $null

    if (Test-Path $settingsPath) {
        try {
            $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
        }
        catch {
            $settings = $null
        }
    }

    $settings = Ensure-ArbeitszeitSettings `
        -Settings $settings `
        -IntervalSeconds $IntervalSeconds `
        -IdleThresholdSeconds $IdleThresholdSeconds

    if (!(Test-Path $settingsPath)) {
        Save-ArbeitszeitSettings -BaseDir $BaseDir -Settings $settings
    }

    return $settings
}

function Save-ArbeitszeitSettings {
    param(
        [string]$BaseDir,
        $Settings
    )

    if (!(Test-Path $BaseDir)) {
        New-Item -ItemType Directory -Path $BaseDir -Force | Out-Null
    }

    $settingsPath = Join-Path $BaseDir "settings.json"
    $normalized = Ensure-ArbeitszeitSettings -Settings $Settings

    $normalized |
        ConvertTo-Json -Depth 6 |
        Set-Content -LiteralPath $settingsPath -Encoding UTF8
}

function Get-ArbeitszeitPauseWindows {
    param(
        $Settings,
        [switch]$IncludeDisabled
    )

    $windows = @($Settings.PauseWindows)

    if ($IncludeDisabled) {
        return $windows
    }

    return @($windows | Where-Object { [System.Convert]::ToBoolean($_.Enabled) })
}
