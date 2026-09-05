# PC-Konfigurator – Technische Dokumentation

## Versionierung

Das Projekt verwendet seit Release 2.0 die semantische Versionierung `major.minor`.

Beispiele:

- `2.0` = erste Version der neuen Major-Version
- `2.1` = Minor-Update mit Erweiterungen oder Korrekturen
- Release-Dateien tragen das Muster `PC-Konfigurator-v2.0.zip`

## Ziel

Diese Dokumentation beschreibt den technischen Aufbau, die Ausführungspfade und die wichtigsten Diagnosepunkte des PC-Konfigurators.

## Start und Betriebsmodi

- Empfohlener Einstieg: `PC-Konfigurator.bat` (mit erhöhten Rechten)
- Hauptskript: `PC-Konfigurator.ps1`
- Zielpfade:
  - **Laufwerk-Modus (L):** `X:\` (durch Anwender gewählt)
  - **Documents-Modus (D):** `%USERPROFILE%\Documents`

Bei der Initialisierung werden Word und Excel für jeweils kurze Zeit mit `WindowStyle Minimized` gestartet und danach beendet. Outlook wird bewusst nicht gestartet: Es kann seine Explorer-Oberfläche unabhängig von `WindowStyle` sichtbar öffnen; die folgenden Outlook-Anpassungen benötigen keinen gestarteten Outlook-Prozess.

## Funktionsgruppen

| Bereich | Typische Funktionen | Zweck |
| --- | --- | --- |
| Systemprüfung | `Test-SystemRequirements` | Prüfung von OS-/Office-Voraussetzungen |
| Synchronisation | `sync()`, Robocopy-Fallback | Kopieren von Vorlagen und Fonts |
| Outlook-Signaturen | `Sync-OutlookSignatures` | Sicherung nach `<Zielpfad>\Signaturen` und Rücksicherung bei Bedarf |
| Office-Konfiguration | `Set-WordCustomizer`, `Set-ExcelCustomizer`, `Sync-OfficeQuickAccessToolbarTemplates` | Defaults, Vorlagen und Schnellzugriffsvorlagen je Anwendung |
| Registry | `Set-OfficeRegistrySettings`, `Set-FontRegistrySettings`, `Set-ExcelQuickAccessToolbar` | Persistente Office-/Font-/QAT-Parameter |
| Logging | `Write-Log`, `Clear-OldLogs` | Nachvollziehbarkeit und Wartung |
| Windows | `Set-TaskbarSettings` | optionale Desktop-/Explorer-Anpassungen |

## Office-Konfiguration im Detail

### Excel

- `Mappe.xltx` wird nur bei fehlender Zieldatei nach `%APPDATA%\Microsoft\Excel\XLSTART` kopiert.
- Die Arbeitsmappenvorlagen werden per COM auf die gewählte Standardschrift und -größe angepasst; dabei werden die Formatvorlage **Normal**, das erste Arbeitsblatt und die Zellen berücksichtigt.
- Die Einstellung `Application.AutoCorrect.CorrectSentenceCap` wird per Excel-COM auf `False` gesetzt und ausgelesen. Die Registry-Werte `CorrectSentenceCap = 0` unter den vorhandenen Excel-Optionspfaden dienen als Fallback.
- Die QAT-Konfiguration schreibt `Command1` bis `Command5` für Speichern, Rückgängig, Wiederholen, Seitenlayoutansicht sowie Seitenansicht/Drucken. Zusätzlich wird `Excel.officeUI` nach `%LOCALAPPDATA%\Microsoft\Office` kopiert.
- Weitere Excel-Optionen konfigurieren Entwicklertools, Datei- und Vorlagenpfade, einen alternativen Startpfad, AutoSpeichern nach fünf Minuten und den Office-Startbildschirm.

### Word

- `Normal.dotm` wird bei fehlender Zieldatei nach `%APPDATA%\Microsoft\Templates` kopiert und per COM angepasst.
- Die Formatvorlagen **Standard** und **Normal** erhalten die gewählte Schrift und Größe, keinen Abstand nach dem Absatz, 1,1-fachen Zeilenabstand und einen Standardtabulator von 1 cm.
- Für Überschriften 1 bis 4 werden abgestufte Größen von 16, 14, 12 und 11 pt gesetzt; Überschrift 1 und 2 erhalten zusätzlich eine untere Linie.
- Die Lernsituationsvorlage `Datei-Vorlagen\Duisdorfer BüroKonzept KG\Lernsituationen\Lernsituationen DBK.dotx` wird bei verfügbarer Word-COM-Automatisierung entsprechend formatiert.
- Die Word-Autokorrektur und -Autoformatierung wird über Registry und, sofern verfügbar, COM konfiguriert.

### Corporate Design und Outlook

- Office-Themes werden abhängig von der Auswahl für INN-tegrativ, Duisdorfer BüroKonzept oder Careli bereitgestellt.
- Outlook erhält ohne Prozessstart Schriften für neue Nachrichten sowie Antworten/Weiterleitungen, die Anzeige von Kalenderwochen und die Vorlage `NormalEmail.dotm`.

## Schriftartauswahl

Die interaktive Auswahl ordnet die Werte `1` bis `8` diesen Schriftfamilien zu: Aptos, Aptos Narrow, Arial, Calibri, Futura, PT Sans, Roboto und Segoe UI. Ohne bestätigte Auswahl setzt das Skript Aptos sowie 11 pt für Word/Outlook und 10 pt für Excel.

## Signatur-Synchronisation (Outlook)

- Lokaler Quellpfad: `%APPDATA%\Microsoft\Signatures`
- Backup-Ziel: `<gewählter Zielpfad>\Signaturen`
- Ablauf pro Lauf:
  1. Lokale Signaturen werden (falls vorhanden) in den Backup-Ordner synchronisiert.
  2. Existiert ein Backup, wird geprüft, ob Inhalte nach `%APPDATA%\Microsoft\Signatures` zurückkopiert werden müssen.
  3. Falls lokal nichts vorhanden ist, erfolgt eine Wiederherstellung aus dem Backup.
  4. Falls lokal bereits Inhalte existieren, werden nur fehlende Backup-Inhalte ergänzt.

Hinweis für den Betrieb: Nachdem Anwender eigene Outlook-Signaturen erstellt oder geändert haben, sollte der PC-Konfigurator am gleichen Rechner erneut ausgeführt werden, damit das Backup aktualisiert wird.

## Systemanforderungen

- Windows 10 (Build 18362+) oder Windows 11
- Microsoft Office 2013 oder neuer
- Lokale Administratorrechte empfohlen (für vollständige Anwendung aller Änderungen)

## Hinweis zu erhöhten Rechten

Der PC-Konfigurator kann grundsätzlich im Benutzerkontext laufen. Ohne erhöhte Rechte können jedoch einzelne systemnahe Schritte eingeschränkt sein oder übersprungen werden, zum Beispiel:

- Beenden bestimmter Prozesse (z. B. `OfficeClickToRun`)
- Änderungen in geschützten Zielpfaden
- Sofortige Anwendung einzelner Shell-/Explorer-bezogener Anpassungen

Benutzerbezogene Konfigurationen (z. B. `HKCU`, `%LOCALAPPDATA%`, `%USERPROFILE%\Documents`) werden weiterhin regulär ausgeführt.

## Logging

```text
%USERPROFILE%\Documents\PC-Konfigurator\Logs\
├── Log_YYYYMMDD_HHMMSS.log
└── Robocopy_YYYYMMDD_HHMMSS.log
```

- Log-Level: `INFO`, `WARN`, `ERROR`
- Bereinigung älterer Logs ist integriert.

### Robocopy-Ausgabe

Die Synchronisationsjobs verwenden `/UNILOG+`, jedoch kein `/TEE`. Deshalb erscheinen keine detaillierten Robocopy-Ausgaben in der Konsole; sie werden ausschließlich in den Dateien unter `%USERPROFILE%\Documents\PC-Konfigurator\Robocopy-Logs\` protokolliert. Die Konsole zeigt weiterhin den Fortschritt und zusammengefasste Ergebnisse.

## Fehlerdiagnose (Kurzübersicht)

### Office-Erkennung prüfen

```powershell
Get-ItemProperty -Path "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration" -Name "ProductVersion"
```

### Windows-Version prüfen

```powershell
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Version, BuildNumber
```

### Typische Ursachen bei Installationsproblemen

- fehlende erhöhte Rechte bei systemnahen Schritten
- gesperrte Dateien (z. B. durch Office/OneDrive)
- unvollständige Office-Installation

## Synchronisierte Inhalte

- `Datei-Vorlagen/` → Office-Vorlagen
- `Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff\` → `%APPDATA%\Microsoft\Office\Excel.officeUI` und `Word.officeUI` → Schnellzugriffsvorlagen für Word/Excel
- `Fonts/` → Schriftarten
- `%APPDATA%\Microsoft\Signatures` ↔ `<Zielpfad>\Signaturen` → Outlook-Signaturen (Backup/Rücksicherung)
- Registry-Anpassungen für Office-Defaults
- `docs/PC-Konfigurator - Einstellungen.pdf` → ausführliche fachliche Übersicht der Office- und Windows-Einstellungen; im Release enthalten
- `docs/PC-Konfigurator - Einstellungen.docx` → Arbeitsdatei der Übersicht; vom Release ausgeschlossen

## Weitere Dokumente

- Anwenderdokumentation: `DOKUMENTATION_ANWENDER.md`
- Registry-Übersicht: `Registry-Einstellungen.md`
- Detaillierte Einstellungsübersicht: `PC-Konfigurator - Einstellungen.pdf`
- Projektüberblick: `../README.md`
- Release-Historie: `../Release/CHANGELOG.md`
- Web-Anleitung: `https://share.eu.articulate.com/d15vUSkhGBZUcTHq-gI4t`
