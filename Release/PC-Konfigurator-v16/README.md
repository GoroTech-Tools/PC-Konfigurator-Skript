# PC-Konfigurator

**Stand:** 12.06.2026
**Autor:** Thomas Gorontzy
**Plattform:** Windows 10 (Build 18362+) / Windows 11, Office 2013+

## Überblick

Der PC-Konfigurator richtet Windows- und Office-Umgebungen automatisiert ein. Dabei werden Vorlagen und Schriftarten synchronisiert, Office- und Windows-Einstellungen gesetzt sowie Abläufe protokolliert.

## Dokumentation

- Anwenderdokumentation: `docs/DOKUMENTATION_ANWENDER.md`
- Technische Dokumentation: `docs/DOKUMENTATION_TECHNIK.md`
- Registry-Übersicht: `docs/Registry-Einstellungen.md`
- Release-Changelog: `Release/CHANGELOG.md`
- Web-Anleitung: `https://share.eu.articulate.com/d15vUSkhGBZUcTHq-gI4t`

## Kernfunktionen

- Synchronisation von `Datei-Vorlagen/` und `Fonts/`
- Sicherung und bedarfsorientierte Rücksicherung von Outlook-Signaturen (`%APPDATA%\Microsoft\Signatures` ↔ `<Zielpfad>\Signaturen`)
- Fallback-Logik bei Kopierproblemen (z. B. Robocopy)
- Interaktive Auswahl von Installationsmodus und Schriftoptionen
- Office-Anpassungen (Word, Excel, Outlook)
- Protokollierung und Bereinigung älterer Logs

## Schnellstart

1. `PC-Konfigurator.bat` per Rechtsklick mit erhöhten Rechten starten (**empfohlen**).
2. Installationsmodus wählen:
   - **L** = Laufwerk
   - **D** = `%USERPROFILE%\Documents`
3. Optional Schriftart und Schriftgrößen anpassen.
4. Nach Abschluss Office neu öffnen.

Hinweis zu Outlook-Signaturen: Nachdem Sie eigene Signaturen für Microsoft Outlook erstellt oder geändert haben, führen Sie den PC-Konfigurator am gleichen Rechner ein weiteres Mal aus.

Hinweis: Der PC-Konfigurator läuft grundsätzlich auch ohne erhöhte Rechte; einzelne systemnahe Schritte können dann jedoch nur eingeschränkt angewendet werden.

## Support

- Log-Verzeichnis: `%USERPROFILE%\Documents\PC-Konfigurator\Logs\`
- Bei Fehlern bitte die neueste Log-Datei beilegen.
