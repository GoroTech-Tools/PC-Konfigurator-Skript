# PC-Konfigurator - Technische Referenz

## Schnellstart

```powershell
# 1. Als Administrator ausführen
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# 2. Skript starten
.\PC-Konfigurator.bat
```

## Funktionsübersicht

| Kategorie | Funktionen | Beschreibung |
|-----------|------------|-------------|
| **System** | `Test-SystemRequirements` | Prüft Windows 10+ und Office 2013+ |
| **Sync** | `sync()` | Hauptsynchronisation mit Robocopy |
| **Office** | `CopyExcelTemplate`, `CopyWordTemplate`, `CopyOutlookTemplate` | Template-Installation |
| **Registry** | `Set-OfficeRegistrySettings`, `Set-WordAutoCorrectRegistry` | Office-Konfiguration |
| **Logging** | `Write-Log`, `Clear-OldLogs` | Protokollierung |

## Zielpfade

- **Laufwerk-Modus (L):** `X:\` (Benutzer wählt Laufwerksbuchstaben)
- **Documents-Modus (D):** `%USERPROFILE%\Documents`

## Systemanforderungen

- Windows 10 Build 18362+ oder Windows 11
- Office 2013 oder neuer
- Administrator-Rechte

## Log-Verzeichnis

```text
%USERPROFILE%\Documents\PC-Konfigurator\Logs\
├── Log_YYYYMMDD_HHMMSS.log          # Hauptprotokoll
└── Robocopy_YYYYMMDD_HHMMSS.log     # Sync-Details
```

## Fehlerbehebung

### Office nicht erkannt

```powershell
# Registry-Pfade prüfen
Get-ItemProperty -Path "HKLM:\Software\Microsoft\Office\ClickToRun\Configuration" -Name "ProductVersion"
```

### System nicht unterstützt

```powershell
# Windows-Version prüfen
Get-CimInstance -ClassName Win32_OperatingSystem | Select Version,BuildNumber
```

## Wichtige Registry-Schlüssel

```registry
# Office ClickToRun
HKLM:\Software\Microsoft\Office\ClickToRun\Configuration

# Office MSI
HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot

# Word Templates
HKCU:\Software\Microsoft\Office\16.0\Word\Options
```

## Synchronisierte Inhalte

- **Datei-Vorlagen/** → Office-Templates (Word, Excel, Outlook)
- **Fonts/** → Windows-Schriftarten
- Registry-Einstellungen für optimale Office-Konfiguration
