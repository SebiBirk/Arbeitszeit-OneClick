# Arbeitszeit OneClick

Arbeitszeit OneClick ist eine lokale Arbeitszeiterfassung für Windows. Das Tool
zeichnet Arbeitsbeginn, Arbeitsende und Nettoarbeitszeit auf, berücksichtigt
Inaktivität innerhalb frei konfigurierbarer Pausenfenster und stellt den
aktuellen Tages-, Wochen- und Monatsstand in einer Desktop-Anzeige dar.

Alle Arbeitszeitdaten bleiben lokal auf dem Rechner. Die Anwendung benötigt
keinen Cloud-Dienst und kein Benutzerkonto.

## Funktionen

- automatische Erfassung nach der Windows-Anmeldung
- frei einstellbarer Start-Offset vor der ersten Aktivität des Tages
- konfigurierbare Tages- und Wochenziele
- automatische Pausenerkennung in definierbaren Zeitfenstern
- absolute Pausenintervalle, zum Beispiel `09:01-09:12 (Frühstück)`
- manuelle Pausen und nachträgliche Tageskorrekturen mit genauen Von-bis-Zeiten
- Erfassung von Tätigkeiten und Projektbuchungen
- Tages-, Wochen- und Monatsstatistik
- Monats- und Gesamtüberstundensaldo auf Basis der Wochenarbeitszeit
- Export eines erweiterten Arbeitszeitberichts als HTML beziehungsweise PDF
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

Bei einer Aktualisierung erkennt der Installer vorhandene Installationen unter
`%LOCALAPPDATA%\Arbeitszeit` und `C:\Arbeitszeit`. Falls beide existieren, wird
automatisch die Installation mit der größeren CSV-Datenhistorie aktualisiert;
parallel laufende alte Programmprozesse werden beendet, die alten Datendateien
jedoch nicht gelöscht.

Sie richtet den Autostart für Tracker und Anzeige ein, erstellt Verknüpfungen
im Startmenü sowie auf dem Desktop und startet die Anwendung direkt.

## Bedienung

Die Anzeige zeigt die aktuelle Nettoarbeitszeit, Pausen und den Fortschritt zum
Tagesziel. Über die Oberfläche lassen sich:

- manuelle Pausen starten und beenden,
- der Arbeitsbeginn und einzelne Pausenintervalle des aktuellen Tages korrigieren,
- Tätigkeiten erfassen,
- Tages- und Wochenziele ändern,
- Start-Offset, Messintervall, Idle-Schwelle und Pausenfenster konfigurieren sowie
- Wochen- und Monatsberichte erzeugen.

Der Start-Offset wird beim ersten erkannten Arbeitsbeginn eines Tages
abgezogen. Bei einem Offset von drei Minuten und einer ersten Aktivität um
`07:05` beginnt die erfasste Zeit somit um `07:02`.

In der Tageskorrektur werden Pausen nicht als Minutensumme bearbeitet. Jede
Pause besitzt eine Kategorie sowie eine genaue Von- und Bis-Zeit. Mehrere
Pausen können hinzugefügt oder einzeln entfernt werden; überlappende und in der
Zukunft endende Intervalle werden abgewiesen. Die Dauerfelder in der CSV werden
anschließend automatisch aus diesen Zeiträumen berechnet.

## Berichte und Überstunden

Der Arbeitszeitbericht enthält die aktuelle Woche und den laufenden Monat. Er
zeigt Nettozeit, Sollzeit, Monats- beziehungsweise Wochensaldo, Tätigkeiten,
Pausen und die absoluten Pausenintervalle. Zusätzlich wird der Gesamtsaldo ab
dem ersten vorhandenen Datensatz berechnet.

Die Regelarbeitszeit beträgt standardmäßig 40 Stunden pro Woche und kann im
Setup geändert werden. Für Sollzeit und Überstunden wird sie gleichmäßig auf
Montag bis Freitag verteilt. Feiertage und Abwesenheiten werden nicht
automatisch abgezogen.

## Lokale Daten

Die wichtigsten Dateien im Installationsordner sind:

| Datei | Inhalt |
| --- | --- |
| `Arbeitszeiten.csv` | tägliche Arbeitszeiten und Pausen |
| `Arbeitszeit_Taetigkeiten.csv` | erfasste Tätigkeiten und Projektzeiten |
| `taetigkeiten.json` | Tätigkeitsdaten für die Anzeige |
| `settings.json` | persönliche Einstellungen |
| `state.json` | aktueller Zustand des Trackers |

Die bestehende Struktur von `Arbeitszeiten.csv` bleibt erhalten. Neue Versionen
ergänzen am Ende die Spalte `Pausen_Zeitraeume`; vorhandene Zeilen und auch
unbekannte zusätzliche Spalten werden beim nächsten Speichern übernommen.
Alte CSV- und Settings-Dateien bleiben damit lesbar.

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

Die automatischen Kompatibilitäts- und Feature-Tests lassen sich aus dem
Projektordner starten:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\tests\Test-Features.ps1"
```
