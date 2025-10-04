# PC-Konfigurator - Benutzeranleitung

## 🎯 Was macht der PC-Konfigurator?

Der PC-Konfigurator richtet Ihren Windows-PC mit Office-Vorlagen, Schriftarten und optimierten Einstellungen ein.

## ⚡ Schnellstart (5 Minuten)

### Schritt 1: Vorbereitung

1. **Office schließen** - Speichern Sie alle Dateien und schließen Word, Excel, Outlook
2. **Als Administrator ausführen** - Rechtsklick auf `PC-Konfigurator.bat` → "Als Administrator ausführen"

### Schritt 2: Installation

1. **Zielpfad wählen:**
   - **L** drücken = Installation auf externes Laufwerk (USB-Stick, Netzlaufwerk)
   - **D** drücken = Installation in Ihren Documents-Ordner

2. **Bei Laufwerk (L):** Laufwerksbuchstaben eingeben (z.B. "F" für USB-Stick)

3. **Warten** - Das Skript kopiert automatisch alle Dateien und konfiguriert Office

### Schritt 3: Fertig

- Der PC startet sich automatisch neu
- Alle Office-Vorlagen sind verfügbar
- Neue Schriftarten sind installiert

## 📁 Was wird installiert?

| Bereich | Inhalt |
|---------|---------|
| **Word** | Bewerbungsvorlagen, Briefvorlagen, Praktikumsunterlagen |
| **Excel** | Kalkulationsvorlagen, Übungsblätter |
| **Outlook** | E-Mail-Signaturen, Templates |
| **Schriftarten** | Aptos, Montserrat, Font Awesome, weitere Fonts |
| **Einstellungen** | Optimierte Office-Konfiguration |

## 🚨 Fehlerbehebung

### "Systemanforderungen nicht erfüllt"

- **Problem:** Ihr Windows oder Office ist zu alt
- **Lösung:** Windows 10 (aktuell) oder Windows 11 + Office 2013 oder neuer benötigt

### "Office nicht erkannt"

- **Problem:** Office-Installation wird nicht gefunden
- **Lösung:** Office über Microsoft Store oder Microsoft.com installieren/reparieren

### "Berechtigung verweigert"

- **Problem:** Skript läuft nicht als Administrator
- **Lösung:** Rechtsklick → "Als Administrator ausführen"

## 📞 Hilfe benötigt?

- **Logs finden:** `C:\Users\[IhrName]\Documents\PC-Konfigurator\Logs\`
- **Entwickler:** Thomas Gorontzy
- **Support:** Wenden Sie sich an Ihren IT-Administrator

## ⚙️ Erweiterte Optionen

### Manuelle Ausführung

```cmd
powershell.exe -ExecutionPolicy Bypass -File "PC-Konfigurator.ps1"
```

### Nur bestimmte Bereiche synchronisieren

Das Skript erkennt automatisch, welche Office-Programme installiert sind und konfiguriert nur diese.

---
**Tipp:** Führen Sie den PC-Konfigurator regelmäßig aus, um neue Vorlagen und Updates zu erhalten.
