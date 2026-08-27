param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA "Arbeitszeit"),
    [switch]$NoDesktopShortcut
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Windows.Forms

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)
$logPath = Join-Path $env:TEMP "Arbeitszeit-Setup.log"

function Show-SetupMessage {
    param(
        [string]$Text,
        [string]$Title = "Arbeitszeit Setup",
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    ) | Out-Null
}

function New-Shortcut {
    param(
        [string]$Path,
        [string]$TargetPath,
        [string]$Arguments = "",
        [string]$WorkingDirectory = "",
        [string]$IconLocation = ""
    )

    $folder = Split-Path -Parent $Path

    if (!(Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($Path)
    $shortcut.TargetPath = $TargetPath
    $shortcut.Arguments = $Arguments
    $shortcut.WorkingDirectory = $WorkingDirectory

    if (-not [string]::IsNullOrWhiteSpace($IconLocation)) {
        $shortcut.IconLocation = $IconLocation
    }

    $shortcut.Save()
}

try {
    Start-Transcript -LiteralPath $logPath -Force | Out-Null

    $installer = Join-Path $sourceDir "Install-Arbeitszeit.ps1"

    if (!(Test-Path $installer)) {
        throw "Install-Arbeitszeit.ps1 wurde nicht gefunden."
    }

    Write-Host "Installationsquelle: $sourceDir"
    Write-Host "Zielordner: $InstallDir"

    & powershell.exe `
        -NoProfile `
        -ExecutionPolicy RemoteSigned `
        -File $installer `
        -InstallDir $InstallDir `
        -Quiet

    if ($LASTEXITCODE -ne 0) {
        throw "Basis-Installer wurde mit Fehlercode $LASTEXITCODE beendet."
    }

    $iconPath = Join-Path $InstallDir "Arbeitszeit.ico"
    $anzeigeExe = Join-Path $InstallDir "ArbeitszeitAnzeige.exe"
    $csvPath = Join-Path $InstallDir "Arbeitszeiten.csv"
    $startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Arbeitszeit"
    $desktopDir = [Environment]::GetFolderPath("DesktopDirectory")

    if (!(Test-Path $anzeigeExe)) {
        throw "ArbeitszeitAnzeige.exe wurde nicht erstellt."
    }

    $anzeigeTarget = $anzeigeExe
    $anzeigeArguments = '-BaseDir "' + $InstallDir + '"'

    New-Shortcut `
        -Path (Join-Path $startMenuDir "Arbeitszeit.lnk") `
        -TargetPath $anzeigeTarget `
        -Arguments $anzeigeArguments `
        -WorkingDirectory $InstallDir `
        -IconLocation $iconPath

    New-Shortcut `
        -Path (Join-Path $startMenuDir "Arbeitszeit Ordner.lnk") `
        -TargetPath "explorer.exe" `
        -Arguments ('"' + $InstallDir + '"') `
        -WorkingDirectory $InstallDir `
        -IconLocation $iconPath

    if (-not $NoDesktopShortcut) {
        New-Shortcut `
            -Path (Join-Path $desktopDir "Arbeitszeit.lnk") `
            -TargetPath $anzeigeTarget `
            -Arguments $anzeigeArguments `
            -WorkingDirectory $InstallDir `
            -IconLocation $iconPath
    }

    $message = @(
        "Arbeitszeit wurde erfolgreich installiert.",
        "",
        "Autostart ist aktiv.",
        "Zielordner: $InstallDir",
        "CSV: $csvPath",
        "",
        "Die Anzeige wurde gestartet und ist im Startmenue verfuegbar."
    ) -join [Environment]::NewLine

    Show-SetupMessage -Text $message
    exit 0
}
catch {
    $message = @(
        "Installation fehlgeschlagen.",
        "",
        $_.Exception.Message,
        "",
        "Logdatei:",
        $logPath
    ) -join [Environment]::NewLine

    Show-SetupMessage `
        -Text $message `
        -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)

    Write-Error $_
    exit 1
}
finally {
    try {
        Stop-Transcript | Out-Null
    }
    catch {}
}
