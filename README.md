# PC-Konfigurator

**Version:** 2.0  
**Datum:** September 2025  
**Autor:** Thomas Gorontzy  
**Kompatibilität:** Windows 10 (Build 18362+) / Windows 11, Office 2013+  

## Übersicht

Der PC-Konfigurator ist ein PowerShell-Skript zur automatischen Konfiguration von Windows 11 und Microsoft Office-Programmen. Das Skript synchronisiert Vorlagen, Schriftarten und konfiguriert Registry-Einstellungen für eine optimale Arbeitsumgebung.

## 🚀 Hauptfunktionen

### ✅ **Datei-Synchronisation**

- Synchronisation von Datei-Vorlagen (Word, Excel, Outlook)
- Installation und Verwaltung von Schriftarten
- Robocopy-basierte Synchronisation mit detailliertem Logging

### ✅ **Office-Konfiguration**

- Registry-Einstellungen für Word, Excel und Outlook
- Template-Installation für alle Office-Anwendungen
- Outlook-Anpassungen (Signaturen, Autokorrektur)
- Kompatibilität mit Office 2013 bis Office 365

### ✅ **System-Verwaltung**

- Automatische Systemanforderungsprüfung
- Umfassendes Logging-System
- Bereinigung alter Log-Dateien
- Windows Explorer-Neustart nach Änderungen

## 📋 Systemanforderungen

| Komponente | Mindestanforderung |
|------------|-------------------|
| **Windows** | Windows 10 (Build 18362) oder Windows 11 |
| **Office** | Microsoft Office 2013 oder neuer |
| **PowerShell** | Windows PowerShell 5.1+ |
| **Berechtigung** | Lokale Administrator-Rechte |

## 🎯 Zielverzeichnisse

Das Skript bietet zwei Installationsmodi:

### **Modus L (Laufwerk)**

- Installation auf ein externes Laufwerk (USB, Netzwerk)
- Benutzer wählt Laufwerksbuchstaben
- Ideal für portable Installationen

### **Modus D (Documents)**

- Installation im Documents-Verzeichnis des aktuellen Benutzers
- Pfad: `%USERPROFILE%\Documents`
- Ideal für persönliche Konfigurationen

## 📁 Verzeichnisstruktur

```text
PC-Konfigurator/
├── PC-Konfigurator.ps1          # Hauptskript
├── PC-Konfigurator.bat          # Windows-Starter
├── README.md                    # Diese Dokumentation
├── Datei-Vorlagen/             # Office-Vorlagen
│   ├── Bewerbungen/
│   ├── Praktikum/
│   ├── Privat/
│   └── Sonstiges/
├── Fonts/                      # Schriftarten
│   ├── Aptos/
│   ├── Font Awesome 6/
│   ├── Montserrat/
│   └── weitere...
└── _Entwicklung/               # Entwicklerdateien
    ├── Anleitung/
    └── Dokumentation
```

## ⚙️ Funktionsübersicht

### **Haupt-Funktionen**

| Funktion | Beschreibung |
|----------|-------------|
| `Confirm-OfficeClosure` | Überprüft, ob Office-Anwendungen geschlossen sind |
| `Test-SystemRequirements` | Prüft Windows- und Office-Versionen |
| `sync()` | Hauptfunktion für Datei-Synchronisation |
| `Write-Log` | Zentrales Logging-System |
| `Clear-OldLogs` | Bereinigung alter Log-Dateien |

### **Office-Template Funktionen**

| Funktion | Zweck |
|----------|-------|
| `CopyExcelTemplate` | Installiert Excel-Vorlagen |
| `CopyWordTemplate` | Installiert Word-Vorlagen |  
| `CopyOutlookTemplate` | Installiert Outlook-Signaturen |

### **Registry-Konfiguration**

| Funktion | Anwendung |
|----------|-----------|
| `Set-OfficeRegistrySettings` | Allgemeine Office-Einstellungen |
| `Set-WordAutoCorrectRegistry` | Word-Autokorrektur-Einstellungen |
| `Set-OutlookCustomizer` | Outlook-spezifische Konfiguration |

## 🔧 Installation und Verwendung

### **1. Vorbereitung**

```powershell
# Als Administrator ausführen
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### **2. Ausführung**

```powershell
# Über Batch-Datei (empfohlen)
.\PC-Konfigurator.bat

# Direkt über PowerShell
.\PC-Konfigurator.ps1
```

### **3. Interaktive Auswahl**

1. Office-Programme schließen und bestätigen
2. Zielpfad auswählen:
   - **L** für Laufwerk (externe Installation)
   - **D** für Documents (persönliche Installation)
3. Bei Laufwerk-Modus: Laufwerksbuchstaben eingeben
4. Automatische Konfiguration läuft

## 📊 Logging-System

### **Log-Verzeichnis**

```text
%USERPROFILE%\Documents\PC-Konfigurator\Logs\
```

### **Log-Dateien**

- `Log_YYYYMMDD_HHMMSS.log` - Hauptprotokoll
- `Robocopy_YYYYMMDD_HHMMSS.log` - Synchronisations-Details
- Automatische Bereinigung nach 30 Tagen

### **Log-Level**

| Level | Beschreibung |
|-------|-------------|
| `INFO` | Normale Vorgänge |
| `WARN` | Warnungen |
| `ERROR` | Fehler |

## 🛡️ Sicherheitsfeatures

### **System-Validierung**

- Automatische Windows-Versionsprüfung
- Office-Installations-Erkennung (ClickToRun, MSI, Executable)
- Berechtigungs-Validierung

### **Fehlerbehandlung**

- Try-Catch-Blöcke für alle kritischen Operationen
- Fallback-Mechanismen bei Registry-Fehlern
- Fortsetzung bei Teilfehlern

### **Blacklist-System**

Geschützte Verzeichnisse werden nicht synchronisiert:

- `System32`
- `SysWOW64`  
- `Windows`
- `Program Files`
- `AppData`

## 🔍 Office-Erkennung

Das Skript erkennt Office-Installationen über drei Methoden:

### **1. ClickToRun-Installation** (Office 365, 2019+)

```registry
HKLM:\Software\Microsoft\Office\ClickToRun\Configuration
HKLM:\Software\WOW6432Node\Microsoft\Office\ClickToRun\Configuration
```

### **2. MSI-Installation** (Office 2013-2016)

```registry
HKLM:\Software\Microsoft\Office\16.0\Common\InstallRoot
HKLM:\Software\Microsoft\Office\15.0\Common\InstallRoot
```

### **3. Executable-Erkennung** (Fallback)

```text
%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE
%ProgramFiles(x86)%\Microsoft Office\Office16\WINWORD.EXE
```

## 📝 Registry-Änderungen

### **Word-Einstellungen**

- Autokorrektur-Konfiguration
- Template-Pfade
- Dokumenten-Standards

### **Excel-Einstellungen**

- Arbeitsmappen-Vorlagen
- Berechnungsoptionen
- Darstellungseinstellungen

### **Outlook-Einstellungen**

- Signatur-Konfiguration
- E-Mail-Templates
- Anzeige-Optionen

## 🚨 Problembehandlung

### **Häufige Probleme**

#### **"Systemanforderungen nicht erfüllt"**

```powershell
# Manuelle Prüfung
Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Version,BuildNumber,Caption
```

#### **Office nicht erkannt**

- Office-Installation über `appwiz.cpl` prüfen
- Registry-Schlüssel manuell überprüfen
- Office reparieren über Systemsteuerung

#### **Berechtigungsfehler**

```powershell
# Als Administrator ausführen
Start-Process PowerShell -Verb RunAs
```

### **Log-Analyse**

```powershell
# Aktuelles Log anzeigen
Get-Content "$env:USERPROFILE\Documents\PC-Konfigurator\Logs\Log_*.log" | Select-Object -Last 50
```

## 🔄 Wartung

### **Automatische Bereinigung**

- Log-Dateien: 30 Tage Aufbewahrung
- Robocopy-Logs: 30 Tage Aufbewahrung
- Temporäre Dateien werden gelöscht

### **Manuelle Wartung**

```powershell
# Logs manuell bereinigen
Remove-Item "$env:USERPROFILE\Documents\PC-Konfigurator\Logs\*" -Older (Get-Date).AddDays(-30)
```

## 📞 Support

### **Entwickler-Kontakt**

- **Name:** Thomas Gorontzy
- **Repository:** <https://github.com/TomGorontzy/PC-Konfigurator>

### **Dokumentation**

- Vollständige Anleitung: `_Entwicklung/Anleitung/`
- Technische Details: `_Entwicklung/PC-Konfigurator (Beschreibung).pdf`

## 📄 Lizenz

Dieses Projekt ist für den internen Gebrauch entwickelt. Alle Rechte vorbehalten.

---

## 🔄 Changelog

### **Version 2.0** (September 2025)

- ✅ Documents-Verzeichnis Support hinzugefügt
- ✅ Erweiterte Office-Erkennung (ClickToRun, MSI, Executable)
- ✅ Verbesserte Systemanforderungsprüfung
- ✅ Outlook 2024+ Unterstützung
- ✅ Konsolidiertes Logging-System
- ✅ Robuste Fehlerbehandlung

### **Version 1.x** (2024)

- ✅ Grundfunktionalität
- ✅ Office-Template-Synchronisation
- ✅ Registry-Konfiguration
- ✅ Basis-Logging
