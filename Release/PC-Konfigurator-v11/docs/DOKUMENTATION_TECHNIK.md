# PC-Konfigurator – Technische Dokumentation (Release v11)

## Ziel

Diese Dokumentation beschreibt den technischen Aufbau, die Ausführungspfade und die wichtigsten Diagnosepunkte des PC-Konfigurators in der Release-Version 11.

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
| Office-Konfiguration | `Set-WordCustomizer`, `Set-ExcelCustomizer` | Defaults und Vorlagen je Anwendung |
| Registry | `Set-OfficeRegistrySettings`, `Set-FontRegistrySettings` | Persistente Office-/Font-Parameter |
| Logging | `Write-Log`, `Clear-OldLogs` | Nachvollziehbarkeit und Wartung |
| Windows | `Set-TaskbarSettings` | optionale Desktop-/Explorer-Anpassungen |

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
- Registry-Anpassungen für Office-Defaults

## Weitere Dokumente

- Anwenderdokumentation: `DOKUMENTATION_ANWENDER.md`
- Release-README: `../README.md`
- Changelog: `../../CHANGELOG.md`
