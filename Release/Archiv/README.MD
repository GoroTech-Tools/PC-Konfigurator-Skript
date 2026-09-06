# PC-Konfigurator-GUI

**PC-Konfigurator-GUI** ist der grafische Nachfolger des konsolenbasierten
Automatisierungs-Tools **PC-Konfigurator**. Die GUI ist als PowerShell-
Anwendung auf Basis der Windows Presentation Foundation umgesetzt und wird
mit PS2EXE als eigenständige EXE kompiliert.

Die Windows Presentation Foundation stellt die interaktive grafische
Benutzeroberfläche bereit. Sie führt durch die Konfiguration, macht Auswahl-
und Bestätigungsschritte verständlich zugänglich und zeigt den Fortschritt
sowie Protokollmeldungen während der Ausführung an.

Autor: Thomas Gorontzy · Version 1.0.0.0
Plattform: Windows 10 (Build 18362+) / Windows 11, Office 2013+

## Herkunft und Einsatzbereich

Das Vorhaben baut auf dem konsolenbasierten Projekt
**PC-Konfigurator** auf. Die GUI bildet dessen Konfigurationsabläufe in einem
Windows-Presentation-Foundation-Assistenten ab und ergänzt sie um geführte Auswahl, Systemprüfung,
Live-Protokollierung und einen automatischen Benutzer-Installationsablauf.

Die Anwendung richtet Arbeitsumgebungen mit vorbereiteten Office-Vorlagen,
Schriftarten und Corporate-Design-Einstellungen ein. Sie konfiguriert Word,
Excel und klassisches Outlook, synchronisiert bei Bedarf Outlook-Signaturen
und nimmt ausgewählte Windows-Explorer-Einstellungen vor. Vor dem Ersetzen
persönlicher Vorlagen werden Sicherungskopien angelegt.

## Kernfunktionen

- Synchronisation der Datei-Vorlagen (`Datei-Vorlagen\`) und Schriftarten (`Fonts\`)
- Auswahl von Zielort, Corporate Design, Schriftart und Schriftgrößen
- Vorbereitete Word-, Excel- und Outlook-Vorlagen ohne COM-Schriftbearbeitung
- Outlook-Signatur-Synchronisation und klassische Outlook-MailSettings
- Office-Registry, Themes, Autokorrektur und Schnellzugriffsleisten
- Windows-Explorer-Anpassungen und ausführliches Live-Logging

## Start

1. Release-ZIP entpacken.
2. Office-Dateien speichern und Word, Excel sowie Outlook schließen.
3. `PC-Konfigurator-GUI.exe` direkt aus dem Releaseordner starten.
4. Dem Windows-Presentation-Foundation-Assistenten folgen.

Beim ersten Start kopiert die EXE sich mit allen Laufzeitdateien automatisch nach:

`%LOCALAPPDATA%\PC-Konfigurator-GUI`

Während dieser Vorbereitung zeigt ein eigenes Fenster den Hinweis
„Zur Vorbereitung wird der PC-Konfigurator in Ihrem System hinterlegt. Es geht
gleich weiter.“ sowie den aktuellen Einrichtungsschritt an.

Anschließend startet sie aus diesem Benutzerordner. Ein manuelles
Installationsskript ist nicht erforderlich.

## Build und Release

Build:

```powershell
.\build.ps1
```

Jeder erfolgreiche Build aktualisiert die lokale Release-EXE und das
Release-ZIP unter `release\PC-Konfigurator-GUI-v<Version>\`.

Vollständiges Release-ZIP:

```powershell
.\create-release.ps1
```

Das ZIP liegt unter `release\` und enthält die EXE sowie `README.MD`. Die
Datei-Vorlagen, Fonts, Dokumentationen und `src\` sind vollständig in der EXE
eingebettet und werden beim Start nach LocalAppData extrahiert.

## Voraussetzungen

- Windows 10 (Build 18362+) oder Windows 11
- Windows PowerShell 5.1 oder höher
- Microsoft Office 2013 oder höher

## Logs

- `%USERPROFILE%\Documents\PC-Konfigurator-GUI\Logs\`
- `%USERPROFILE%\Documents\PC-Konfigurator-GUI\Robocopy-Logs\`

Weitere Informationen stehen unter `docs\DOKUMENTATION_ANWENDER.md` und
`docs\DOKUMENTATION_TECHNIK.md`.
