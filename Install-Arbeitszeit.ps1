param(
    [string]$InstallDir = "C:\Arbeitszeit",
    [switch]$NoStart,
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

$sourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$InstallDir = [System.IO.Path]::GetFullPath($InstallDir)

if (!(Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}

$trackerScriptForStop = Join-Path $InstallDir "Arbeitszeit.ps1"
$anzeigeScriptForStop = Join-Path $InstallDir "ArbeitszeitAnzeige.ps1"
$anzeigeHtaForStop = Join-Path $InstallDir "ArbeitszeitAnzeige.hta"
$trackerExeForStop = Join-Path $InstallDir "ArbeitszeitTracker.exe"
$anzeigeExeForStop = Join-Path $InstallDir "ArbeitszeitAnzeige.exe"

Get-CimInstance Win32_Process -Filter "Name='powershell.exe' OR Name='mshta.exe' OR Name='ArbeitszeitAnzeige.exe' OR Name='ArbeitszeitTracker.exe'" |
    Where-Object {
        ($_.Name -eq "ArbeitszeitTracker.exe" -and $_.ExecutablePath -eq $trackerExeForStop) -or
        ($_.Name -eq "ArbeitszeitAnzeige.exe" -and $_.ExecutablePath -eq $anzeigeExeForStop) -or
        ($_.CommandLine -like ('*-File "' + $trackerScriptForStop + '"*')) -or
        ($_.CommandLine -like ('*-File "' + $anzeigeScriptForStop + '"*')) -or
        ($_.CommandLine -like ('*"' + $anzeigeHtaForStop + '"*'))
    } |
    ForEach-Object {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
    }

Start-Sleep -Milliseconds 500

$files = @(
    "Arbeitszeit.ps1",
    "ArbeitszeitAnzeige.ps1",
    "ArbeitszeitTrackerHost.cs",
    "ArbeitszeitAnzeigeHost.cs",
    "ArbeitszeitSettings.ps1",
    "Install-Arbeitszeit.ps1",
    "Install-ArbeitszeitTask.ps1",
    "Arbeitszeit.ico",
    "README.md"
)

foreach ($file in $files) {
    $sourcePath = Join-Path $sourceDir $file

    if (!(Test-Path $sourcePath)) {
        throw "Installationsdatei fehlt: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path $InstallDir $file) -Force
}

function Get-CSharpCompilerPath {
    $candidates = @(
        (Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"),
        (Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe")
    )

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    return $null
}

function Get-PowerShellAutomationAssemblyPath {
    $gacRoot = Join-Path $env:WINDIR "Microsoft.NET\assembly\GAC_MSIL\System.Management.Automation"

    if (!(Test-Path $gacRoot)) {
        return $null
    }

    $assembly = Get-ChildItem $gacRoot -Recurse -Filter "System.Management.Automation.dll" -ErrorAction SilentlyContinue |
        Select-Object -First 1

    if ($assembly) {
        return $assembly.FullName
    }

    return $null
}

function Build-PowerShellHost {
    param(
        [string]$InstallDir,
        [string]$SourceFile,
        [string]$OutputFile
    )

    $compilerPath = Get-CSharpCompilerPath
    $automationAssembly = Get-PowerShellAutomationAssemblyPath
    $sourcePath = Join-Path $InstallDir $SourceFile
    $outputPath = Join-Path $InstallDir $OutputFile
    $iconPath = Join-Path $InstallDir "Arbeitszeit.ico"

    if ([string]::IsNullOrWhiteSpace($compilerPath) -or [string]::IsNullOrWhiteSpace($automationAssembly)) {
        throw "C#-Compiler oder System.Management.Automation.dll wurde nicht gefunden. EXE-Host kann nicht gebaut werden."
    }

    $arguments = @(
        "/nologo",
        "/target:winexe",
        "/platform:anycpu",
        "/out:$outputPath",
        "/reference:$automationAssembly",
        "/reference:System.Windows.Forms.dll"
    )

    if (Test-Path $iconPath) {
        $arguments += "/win32icon:$iconPath"
    }

    $arguments += $sourcePath

    & $compilerPath @arguments

    if ($LASTEXITCODE -eq 0 -and (Test-Path $outputPath)) {
        return $outputPath
    }

    throw "EXE-Host konnte nicht gebaut werden: $OutputFile"
}

$trackerScript = Join-Path $InstallDir "Arbeitszeit.ps1"
$anzeigeScript = Join-Path $InstallDir "ArbeitszeitAnzeige.ps1"
$trackerExe = Build-PowerShellHost -InstallDir $InstallDir -SourceFile "ArbeitszeitTrackerHost.cs" -OutputFile "ArbeitszeitTracker.exe"
$anzeigeExe = Build-PowerShellHost -InstallDir $InstallDir -SourceFile "ArbeitszeitAnzeigeHost.cs" -OutputFile "ArbeitszeitAnzeige.exe"
$taskInstaller = Join-Path $InstallDir "Install-ArbeitszeitTask.ps1"

Remove-Item -LiteralPath (Join-Path $InstallDir "Start-Arbeitszeit.vbs") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $InstallDir "Start-ArbeitszeitAnzeige.vbs") -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath (Join-Path $InstallDir "ArbeitszeitAnzeige.hta") -Force -ErrorAction SilentlyContinue

& powershell.exe `
    -NoProfile `
    -ExecutionPolicy RemoteSigned `
    -File $taskInstaller `
    -TrackerExecute $trackerExe `
    -TrackerArgument ('-BaseDir "' + $InstallDir + '"') `
    -AnzeigeExecute $anzeigeExe `
    -AnzeigeArgument ('-BaseDir "' + $InstallDir + '"')

if (-not $NoStart) {
    Start-ScheduledTask -TaskName "Arbeitszeit-Erfassung"
    Start-ScheduledTask -TaskName "Arbeitszeit-Anzeige"
}

if (-not $Quiet) {
    Write-Host ""
    Write-Host "Arbeitszeit wurde installiert nach: $InstallDir"
    Write-Host "Autostart aktiv: Arbeitszeit-Erfassung und Arbeitszeit-Anzeige"
    Write-Host "CSV: $(Join-Path $InstallDir 'Arbeitszeiten.csv')"
}
