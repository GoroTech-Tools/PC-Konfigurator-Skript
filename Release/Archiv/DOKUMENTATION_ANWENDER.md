# PC-Konfigurator-GUI – Anwenderdokumentation

## Version

GUI-Version: 1.0.0.0

Release-Dateien verwenden das Format `PC-Konfigurator-GUI-v1.0.0.0.zip`.

## Zweck

Die PC-Konfigurator-GUI richtet Windows- und Office-Arbeitsumgebungen mit vorbereiteten Vorlagen, Schriftarten, Corporate Design und empfohlenen Einstellungen ein.

## Grafische Benutzeroberfläche

Die Anwendung verwendet die **Windows Presentation Foundation** für ihre
grafische Benutzeroberfläche. Sie ersetzt die Eingaben der ursprünglichen
Konsolen-Variante durch einen geführten Assistenten, in dem Auswahlen,
Bestätigungen, Fortschritt und Protokollmeldungen übersichtlich angezeigt
werden.

## Installation und Schnellstart

1. Entpacken Sie das Release-ZIP.
2. Speichern Sie offene Dokumente und schließen Sie Word, Excel und Outlook.
3. Starten Sie `PC-Konfigurator-GUI.exe` direkt aus dem entpackten Releaseordner.
4. Folgen Sie dem Windows-Presentation-Foundation-Assistenten.

Die EXE stellt sich beim ersten Start automatisch mit allen Laufzeitdateien im
Benutzerprofil unter `%LOCALAPPDATA%\PC-Konfigurator-GUI` bereit und startet
sich von dort erneut. Ein manueller Installeraufruf ist nicht erforderlich.

Während die Anwendung vorbereitet wird, erscheint ein eigenes Fenster mit dem
Hinweis „Zur Vorbereitung wird der PC-Konfigurator in Ihrem System hinterlegt.
Es geht gleich weiter.“ Der jeweils aktuelle Vorbereitungsschritt wird darin
angezeigt; eine Konsolenausgabe ist hierfür nicht erforderlich.

## GUI-Assistent

Der Assistent führt durch:

- Office-Schließbestätigung und Systemprüfung
- Zielauswahl für Datei-Vorlagen (Laufwerk oder Dokumente)
- Corporate Design (INN-tegrativ, DBK oder Careli)
- Schriftart und Schriftgrößen
- Taskleisten-Ausrichtung: zentriert (Windows-Standard) oder linksbündig
- Ausführung mit Live-Log
- optionalen Explorer-Neustart

## Was wird eingerichtet?

| Bereich | Inhalt |
| --- | --- |
| Word | vorbereitete `Normal.dotm`, Formatvorlagen, Schriftart und Schnellzugriff |
| Excel | vorbereitete `Mappe.xltx`, Standardschrift, Autokorrektur und Schnellzugriff |
| Outlook | vorbereitete `NormalEmail.dotm`, klassische MailSettings und Signatur-Synchronisation |
| Schriftarten | Installation mitgelieferter Fonts und Auswahl aus acht Office-Schriftarten |
| Windows | Explorer-Datenschutz, Suche, Taskleiste und Desktop-Schnellzugriff |

## Vorbereitete Vorlagen

Die GUI verändert die Standardvorlagen nicht per COM. Sie wählt die passende, bereits vorbereitete Datei nach dem Schema aus:

- `Normal-<Schrift>-<Größe>.dotm` → `%APPDATA%\Microsoft\Templates\Normal.dotm`
- `NormalEmail-<Schrift>-<Größe>.dotm` → `%APPDATA%\Microsoft\Templates\NormalEmail.dotm`
- `Mappe-<Schrift>-<Größe>.xltx` → `%APPDATA%\Microsoft\Excel\XLSTART\Mappe.xltx`

Vorhandene Benutzerdateien werden vor dem Überschreiben unter `Dokumente\PC-Konfigurator-GUI\Backups` gesichert.

## Schriftartauswahl

Zur Auswahl stehen Aptos, Aptos Narrow, Arial, Calibri, Futura, PT Sans, Roboto und Segoe UI. Standardmäßig gelten Aptos, 11 pt für Word/Outlook und 10 pt für Excel.

## Outlook-Hinweis

Die Registrywerte und `NormalEmail.dotm` gelten für **klassisches Outlook**. Die neue Outlook-App verwendet eigene Microsoft-365-Einstellungen. Dort muss die Schrift unter **Einstellungen → Mail → Verfassen und Antworten** gesetzt werden. Die GUI protokolliert, wenn die neue Outlook-App erkannt wird.

## Signaturen und Logs

Eigene Outlook-Signaturen werden in den gewählten Vorlagenzielordner unter `Signaturen` gesichert und bei Bedarf nach `%APPDATA%\Microsoft\Signatures` zurückgespielt.

Logs liegen unter:

- `%USERPROFILE%\Documents\PC-Konfigurator-GUI\Logs\`
- `%USERPROFILE%\Documents\PC-Konfigurator-GUI\Robocopy-Logs\`

## Fehlerdiagnose

- Vor dem Lauf alle Office-Programme vollständig schließen.
- Prüfen, ob `%LOCALAPPDATA%\PC-Konfigurator-GUI\Datei-Vorlagen` und `Fonts` vorhanden sind.
- Bei Vorlagenproblemen das Backup unter `Dokumente\PC-Konfigurator-GUI\Backups` verwenden.
- Bei neuer Outlook-App die Schrift in den Microsoft-365-/Outlook-Einstellungen setzen.

## Verwandte Dokumente

- `DOKUMENTATION_TECHNIK.md`
- `Registry-Einstellungen.md`
- `PC-Konfigurator - Einstellungen.pdf`
- `README.MD`
