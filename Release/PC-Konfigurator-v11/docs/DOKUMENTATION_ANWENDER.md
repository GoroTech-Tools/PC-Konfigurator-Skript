# PC-Konfigurator – Anwenderdokumentation (Release v11)

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
| Outlook | Signaturen und unterstützende Einstellungen |
| Schriftarten | Installation mitgelieferter Fonts |
| System | Ergänzende Windows-/Office-Einstellungen |

## Häufige Probleme

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
- Changelog: `../../CHANGELOG.md`
