param(
    [string]$OutputDir = (Join-Path $PSScriptRoot "dist"),
    [string]$PackageName = "Arbeitszeit-OneClick"
)

$ErrorActionPreference = "Stop"

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageDir = Join-Path $OutputDir $PackageName
$zipPath = Join-Path $OutputDir ($PackageName + ".zip")

if (Test-Path $packageDir) {
    Remove-Item -LiteralPath $packageDir -Recurse -Force
}

if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

New-Item -ItemType Directory -Path $packageDir -Force | Out-Null

$files = @(
    "Arbeitszeit-Setup.cmd",
    "Install-Arbeitszeit-OneClick.ps1",
    "Install-Arbeitszeit.ps1",
    "Install-ArbeitszeitTask.ps1",
    "Arbeitszeit.ps1",
    "ArbeitszeitAnzeige.ps1",
    "ArbeitszeitTrackerHost.cs",
    "ArbeitszeitAnzeigeHost.cs",
    "ArbeitszeitSettings.ps1",
    "Arbeitszeit.ico",
    "README.md"
)

foreach ($file in $files) {
    $sourcePath = Join-Path $sourceDir $file

    if (!(Test-Path $sourcePath)) {
        throw "Paketdatei fehlt: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $packageDir $file) -Force
}

if (Test-Path $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

Compress-Archive -Path (Join-Path $packageDir "*") -DestinationPath $zipPath -Force

Write-Host "Mitarbeiter-Paket erstellt:"
Write-Host $zipPath
