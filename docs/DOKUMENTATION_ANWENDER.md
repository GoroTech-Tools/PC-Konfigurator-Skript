# PC-Konfigurator – Anwenderdokumentation

## Version

Aktuelle Hauptversion: 2.0

Release-Dateien verwenden das Format `PC-Konfigurator-v2.0.zip` und folgen der Versionierung `major.minor`.

## Zweck

Der PC-Konfigurator richtet Ihren Windows-PC automatisch mit Office-Vorlagen, Schriftarten und empfohlenen Einstellungen ein.

## Schnellstart

1. Speichern Sie offene Dokumente und schließen Sie Word, Excel und Outlook.
2. Starten Sie `PC-Konfigurator.bat` per Rechtsklick mit erhöhten Rechten (**empfohlen**).
3. Wählen Sie den Installationsmodus:
   - **L** = Ziel-Laufwerk (z. B. USB/Netzwerk)
   - **D** = `%USERPROFILE%\Documents`
4. Optional: Schriftart und Schriftgrößen individuell festlegen.
5. Warten Sie bis zum Abschlussdialog.

Während der Initialisierung startet der PC-Konfigurator Word und Excel kurz minimiert und beendet sie anschließend wieder. Outlook wird nicht geöffnet.

## Konsolenausgaben

Die Meldungen im Konsolenfenster sind farblich eingeordnet:

- **Weiß:** Bestätigungen und erfolgreiche Abschlüsse
- **Grün:** Fragen und Eingabeaufforderungen
- **Cyan:** Informationen, Fortschritt, Auswahlwerte und nicht kritische Hinweise
- **Rot:** Fehler, Abbrüche und ungültige Eingaben

## Was wird eingerichtet?

| Bereich | Inhalt |
| --- | --- |
| Word | Vorlagen (z. B. Bewerbungen, Schriftverkehr), Standardformatierung und Schnellzugriffsvorlagen |
| Excel | Vorlagen, Standardschrift, Autokorrektur, Schnellzugriffsvorlagen und automatische Speicherung |
| Outlook | Signaturen-Sicherung/-Wiederherstellung, E-Mail-Schriften und Kalenderwochen |
| Schriftarten | Installation mitgelieferter Schriftarten und Auswahl aus acht konfigurierbaren Office-Schriftarten |
| System | Ergänzende Windows-/Office-Einstellungen, Explorer-Datenschutz und Schnellzugriff |

## Office-Einstellungen im Überblick

### Excel

- Die Standardschrift und -größe für neue Arbeitsmappen werden auf Ihre Auswahl gesetzt.
- `Mappe.xltx` wird, falls noch nicht vorhanden, nach `%APPDATA%\Microsoft\Excel\XLSTART` kopiert.
- Die Standardformatierung von Vorlagen und Arbeitsblättern wird an die gewählte Schrift angepasst.
- Die Autokorrektur **„Jeden Satz mit einem Großbuchstaben beginnen“** wird deaktiviert.
- Die Schnellzugriffsleiste enthält unter anderem **Speichern**, **Rückgängig**, **Wiederholen**, **Seitenlayoutansicht** sowie **Seitenansicht und Drucken**.
- Entwicklertools, Datei- und Vorlagenpfade sowie ein AutoSpeichern-Intervall von fünf Minuten werden eingerichtet.

### Schriftartauswahl

Bei der individuellen Auswahl stehen diese Schriftarten zur Verfügung:

1. Aptos *(Windows-Standard)*
2. Aptos Narrow *(speziell)*
3. Arial *(veraltet, jedoch AP 1-Vorgabe)*
4. Calibri *(veralteter Windows-Standard)*
5. Futura *(speziell)*
6. PT Sans *(INN-tegrativ-Standard)*
7. Roboto *(speziell)*
8. Segoe UI *(speziell)*

Ohne bestätigte individuelle Auswahl gelten Aptos, 11 pt für Word/Outlook und 10 pt für Excel.

### Word

- Die Standardschrift und -größe sowie die Vorlagen `Normal.dotm` und die DBK-Lernsituationsvorlage werden angepasst.
- Die Formatvorlagen **Standard** und **Normal** erhalten keinen Abstand nach Absätzen, einen 1,1-fachen Zeilenabstand und einen Standardtabulator von 1 cm.
- Überschriften 1 bis 4 werden einheitlich formatiert.
- Entwicklertools, Lineal und Formatierungszeichen werden aktiviert.
- Automatische Listen, Tabellenzellenkorrekturen und unerwünschte automatische Großschreibungen werden eingeschränkt.

### Outlook und Corporate Design

- Schrift und Schriftgröße für neue E-Mails sowie für Antworten und Weiterleitungen werden gesetzt.
- Die Kalenderwoche wird eingeblendet und `NormalEmail.dotm` bereitgestellt.
- Je nach Auswahl werden Farbschemata für INN-tegrativ, Duisdorfer BüroKonzept oder Careli bereitgestellt.

## Hinweise zur Bedienung

- Führen Sie die Installation möglichst bei geschlossenem Office durch.
- Bei Laufwerksinstallation den korrekten Laufwerksbuchstaben angeben.
- Für Änderungen an Vorlagen/Fonts kann ein Neustart einzelner Office-Apps nötig sein.
- Eigene Outlook-Signaturen werden in einen Ordner `Signaturen` im gewählten Zielpfad kopiert und bei Bedarf nach `%APPDATA%\Microsoft\Signatures` zurückgespielt.
- Vorlagen für Schnellzugriffe aus `Datei-Vorlagen\Sonstiges\Symbolleiste Schnellzugriff` werden in den Office-Benutzerpfad übernommen, damit Word- und Excel-Symbolleisten konsistent gesetzt werden.
- Schließen Sie Excel vollständig, bevor Sie es nach der Installation erstmals wieder starten. Nur dann kann Excel die neue Schnellzugriffsleiste zuverlässig einlesen.
- Die Synchronisation läuft mit einer Fortschrittsanzeige; die ausführlichen Robocopy-Details stehen in den Logdateien statt im Konsolenfenster.
- **Wichtiger Hinweis:** Nachdem Sie eigene Signaturen für Microsoft Outlook erstellt oder geändert haben, führen Sie den PC-Konfigurator am gleichen Rechner ein weiteres Mal aus.

## Windows, Vorlagen und Datenquellen

- Zuletzt verwendete Dateien und häufig verwendete Ordner werden im Explorer ausgeblendet; der Verlauf wird bereinigt.
- Der Desktop wird im Explorer-Schnellzugriff hinterlegt.
- Die Taskleiste wird links ausgerichtet, die Suche als Symbol angezeigt und Widgets werden ausgeblendet.
- Das Paket enthält mehr als 200 Excel- und Word-Vorlagen für Lernsituationen, Schriftverkehr, Bewerbungen, Formulare und Übungen.
- Die DBK-Datenquelle kann als Verknüpfung nach `%USERPROFILE%\Meine Datenquellen` eingerichtet werden.

## Häufige Probleme

### FAQ: Warum soll ich nach Signatur-Änderungen erneut ausführen?

- Der PC-Konfigurator sichert Outlook-Signaturen in den Zielordner `Signaturen`.
- Wenn Sie Signaturen neu erstellen oder ändern, wird dieses Backup erst beim nächsten Lauf aktualisiert.
- Durch den erneuten Lauf stellen Sie sicher, dass Ihre neuesten Signaturen auch für eine spätere Wiederherstellung verfügbar sind.

### Fehlende erhöhte Rechte

- Symptom: Einzelne systemnahe Schritte werden übersprungen oder schlagen fehl.
- Lösung: Starter erneut per Rechtsklick **Als Administrator ausführen**.

### Office wird nicht erkannt

- Symptom: Office-spezifische Schritte werden nicht ausgeführt.
- Lösung: Office-Installation prüfen/reparieren und erneut starten.

### Wo finde ich Logs?

- `%USERPROFILE%\Documents\PC-Konfigurator\Logs\`

## Verwandte Dokumente

- Technische Dokumentation: `DOKUMENTATION_TECHNIK.md`
- Registry-Übersicht: `Registry-Einstellungen.md`
- Detaillierte Einstellungsübersicht: `PC-Konfigurator - Einstellungen.pdf`
- Release-Historie: `../Release/CHANGELOG.md`
