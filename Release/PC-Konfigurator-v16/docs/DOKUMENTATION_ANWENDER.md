# PC-Konfigurator – Anwenderdokumentation

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

## Was wird eingerichtet?

| Bereich | Inhalt |
| --- | --- |
| Word | Vorlagen (z. B. Bewerbungen, Schriftverkehr) |
| Excel | Vorlagen und Standardkonfiguration |
| Outlook | Signaturen-Sicherung/-Wiederherstellung und unterstützende Einstellungen |
| Schriftarten | Installation mitgelieferter Fonts |
| System | Ergänzende Windows-/Office-Einstellungen |

## Hinweise zur Bedienung

- Führen Sie die Installation möglichst bei geschlossenem Office durch.
- Bei Laufwerksinstallation den korrekten Laufwerksbuchstaben angeben.
- Für Änderungen an Vorlagen/Fonts kann ein Neustart einzelner Office-Apps nötig sein.
- Eigene Outlook-Signaturen werden in einen Ordner `Signaturen` im gewählten Zielpfad kopiert und bei Bedarf nach `%APPDATA%\Microsoft\Signatures` zurückgespielt.
- **Wichtiger Hinweis:** Nachdem Sie eigene Signaturen für Microsoft Outlook erstellt oder geändert haben, führen Sie den PC-Konfigurator am gleichen Rechner ein weiteres Mal aus.

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
- Release-Historie: `../Release/CHANGELOG.md`
