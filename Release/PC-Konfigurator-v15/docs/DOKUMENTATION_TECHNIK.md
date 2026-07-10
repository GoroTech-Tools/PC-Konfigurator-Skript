# PC-Konfigurator – Technische Dokumentation

## Ziel

Diese Dokumentation beschreibt den technischen Aufbau, die Ausführungspfade und die wichtigsten Diagnosepunkte des PC-Konfigurators.

## Start und Betriebsmodi

- Empfohlener Einstieg: `PC-Konfigurator.bat` (mit erhöhten Rechten)
- Hauptskript: `PC-Konfigurator.ps1`
- Zielpfade:
  - **Laufwerk-Modus (L):** `X:\` (durch Anwender gewählt)
  - **Documents-Modus (D):** `%USERPROFILE%\Documents`

## Funktionsgruppen

| Bereich | Typische Funktionen | Zweck |
| --- | --- | --- |
| Systemprüfung | `Test-SystemRequirements` | Prüfung von OS-/Office-Voraussetzungen |
| Synchronisation | `sync()`, Robocopy-Fallback | Kopieren von Vorlagen und Fonts |
| Outlook-Signaturen | `Sync-OutlookSignatures` | Sicherung nach `<Zielpfad>\Signaturen` und Rücksicherung bei Bedarf |
| Office-Konfiguration | `Set-WordCustomizer`, `Set-ExcelCustomizer` | Defaults und Vorlagen je Anwendung |
| Registry | `Set-OfficeRegistrySettings`, `Set-FontRegistrySettings` | Persistente Office-/Font-Parameter |
| Logging | `Write-Log`, `Clear-OldLogs` | Nachvollziehbarkeit und Wartung |
| Windows | `Set-TaskbarSettings` | optionale Desktop-/Explorer-Anpassungen |

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
- `Fonts/` → Schriftarten
- `%APPDATA%\Microsoft\Signatures` ↔ `<Zielpfad>\Signaturen` → Outlook-Signaturen (Backup/Rücksicherung)
- Registry-Anpassungen für Office-Defaults

## Weitere Dokumente

- Anwenderdokumentation: `DOKUMENTATION_ANWENDER.md`
- Projektüberblick: `../README.md`
- Release-Historie: `../../CHANGELOG.md`
- Web-Anleitung: `https://share.eu.articulate.com/d15vUSkhGBZUcTHq-gI4t`
