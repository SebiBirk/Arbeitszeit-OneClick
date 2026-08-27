# Arbeitszeit OneClick

Arbeitszeit OneClick ist eine lokale Arbeitszeiterfassung für Windows. Das Tool
zeichnet Arbeitsbeginn, Arbeitsende und Nettoarbeitszeit auf, berücksichtigt
Inaktivität innerhalb frei konfigurierbarer Pausenfenster und stellt den
aktuellen Tages- und Wochenstand in einer Desktop-Anzeige dar.

Alle Arbeitszeitdaten bleiben lokal auf dem Rechner. Die Anwendung benötigt
keinen Cloud-Dienst und kein Benutzerkonto.

## Funktionen

- automatische Erfassung nach der Windows-Anmeldung
- konfigurierbare Tages- und Wochenziele
- automatische Pausenerkennung in definierbaren Zeitfenstern
- manuelle Pausen und nachträgliche Tageskorrekturen
- Erfassung von Tätigkeiten und Projektbuchungen
- Tagesübersicht und Wochenstatistik
- Export eines Wochenberichts als HTML beziehungsweise PDF
- Benachrichtigung beim Erreichen des Tagesziels
- Autostart über die Windows-Aufgabenplanung
- Installation pro Benutzer, normalerweise ohne Administratorrechte

## Schnellinstallation

1. Unter **Releases** die Datei `Arbeitszeit-OneClick.zip` herunterladen.
2. Die ZIP-Datei vollständig entpacken.
3. `Arbeitszeit-Setup.cmd` aus dem entpackten Ordner starten.
4. Eventuelle Windows-Sicherheitsabfragen prüfen und bestätigen.

Die OneClick-Installation verwendet standardmäßig diesen Ordner:

```text
%LOCALAPPDATA%\Arbeitszeit
```

Sie richtet den Autostart für Tracker und Anzeige ein, erstellt Verknüpfungen
im Startmenü sowie auf dem Desktop und startet die Anwendung direkt.

## Bedienung

Die Anzeige zeigt die aktuelle Nettoarbeitszeit, Pausen und den Fortschritt zum
Tagesziel. Über die Oberfläche lassen sich:

- manuelle Pausen starten und beenden,
- der aktuelle Tag korrigieren,
- Tätigkeiten erfassen,
- Tages- und Wochenziele ändern,
- Messintervall, Idle-Schwelle und Pausenfenster konfigurieren sowie
- Wochenberichte erzeugen.

## Lokale Daten

Die wichtigsten Dateien im Installationsordner sind:

| Datei | Inhalt |
| --- | --- |
| `Arbeitszeiten.csv` | tägliche Arbeitszeiten und Pausen |
| `Arbeitszeit_Taetigkeiten.csv` | erfasste Tätigkeiten und Projektzeiten |
| `taetigkeiten.json` | Tätigkeitsdaten für die Anzeige |
| `settings.json` | persönliche Einstellungen |
| `state.json` | aktueller Zustand des Trackers |

Die CSV-Dateien sollten in Excel nur kurz geöffnet und danach wieder
geschlossen werden, da Excel sie während der Bearbeitung sperren kann.

## Voraussetzungen

- Windows
- Windows PowerShell 5.1
- .NET Framework einschließlich C#-Compiler

Der Installer erstellt zwei kleine EXE-Hosts für Tracker und Anzeige. Die
eigentliche Programmlogik bleibt in den enthaltenen PowerShell-Skripten
nachvollziehbar.

## Installation aus dem Quellcode

Für eine technische Installation in `C:\Arbeitszeit`:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ".\Install-Arbeitszeit.ps1"
```

Ein anderer Zielordner kann übergeben werden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ".\Install-Arbeitszeit.ps1" -InstallDir "$env:LOCALAPPDATA\Arbeitszeit"
```

## Paket bauen

Das verteilbare Paket wird mit folgendem Befehl erstellt:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ".\Build-MitarbeiterPackage.ps1"
```

Das Ergebnis liegt anschließend unter:

```text
dist\Arbeitszeit-OneClick.zip
```

## Technischer Testlauf

Nach einer manuellen Installation kann der Tracker einmalig ausgeführt werden:

```powershell
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File "C:\Arbeitszeit\Arbeitszeit.ps1" -Once
```

