param(
    [string]$TrackerExecute = "C:\Arbeitszeit\ArbeitszeitTracker.exe",
    [string]$TrackerArgument = "-BaseDir `"C:\Arbeitszeit`"",
    [string]$AnzeigeExecute = "C:\Arbeitszeit\ArbeitszeitAnzeige.exe",
    [string]$AnzeigeArgument = "-BaseDir `"C:\Arbeitszeit`"",
    [string]$TrackerTaskName = "Arbeitszeit-Erfassung",
    [string]$AnzeigeTaskName = "Arbeitszeit-Anzeige"
)

$ErrorActionPreference = "Stop"

if (!(Test-Path $TrackerExecute)) {
    throw "Tracker-Starter nicht gefunden: $TrackerExecute"
}

if (!(Test-Path $AnzeigeExecute)) {
    throw "Anzeige-Starter nicht gefunden: $AnzeigeExecute"
}

$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

function Register-LogonTask {
    param(
        [string]$TaskName,
        [string]$Execute,
        [string]$Argument,
        [string]$Description
    )

    Stop-ScheduledTask `
        -TaskName $TaskName `
        -ErrorAction SilentlyContinue

    Unregister-ScheduledTask `
        -TaskName $TaskName `
        -Confirm:$false `
        -ErrorAction SilentlyContinue

    $action = New-ScheduledTaskAction `
        -Execute $Execute `
        -Argument $Argument

    $trigger = New-ScheduledTaskTrigger `
        -AtLogOn `
        -User $currentUser

    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable

    Register-ScheduledTask `
        -TaskName $TaskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Description $Description |
        Out-Null
}

Register-LogonTask `
    -TaskName $TrackerTaskName `
    -Execute $TrackerExecute `
    -Argument $TrackerArgument `
    -Description "Erfasst Arbeitszeit anhand Windows-Anmeldung und Pausenfenstern. Startet direkt als ArbeitszeitTracker.exe."

Register-LogonTask `
    -TaskName $AnzeigeTaskName `
    -Execute $AnzeigeExecute `
    -Argument $AnzeigeArgument `
    -Description "Zeigt Arbeitszeit live an und erlaubt manuelle Pausen/Korrekturen. Startet direkt als ArbeitszeitAnzeige.exe."

Write-Host "Geplante Aufgabe wurde angelegt: $TrackerTaskName"
Write-Host "Geplante Aufgabe wurde angelegt: $AnzeigeTaskName"
