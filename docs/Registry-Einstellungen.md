# Verwendete Registry-Einstellungen

**Quelle:** `src/PC-Konfigurator.ps1`

Diese Datei fasst die im Projekt verwendeten Registry-Einstellungen für Windows, Word, Excel und Outlook zusammen.
Die meisten Werte werden für den **aktuellen Benutzer** unter `HKCU` gesetzt.

## Windows / Explorer

### `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `HideFileExt` | `DWord` | `0` | Dateiendungen im Explorer anzeigen. |
| `Hidden` | `DWord` | `1` | Versteckte Dateien anzeigen. |
| `ShowSuperHidden` | `DWord` | `1` | Geschützte Systemdateien anzeigen. |
| `ShowRecent` | `DWord` | `0` | **Zuletzt verwendete Dateien** im Schnellzugriff/Explorer deaktivieren. |
| `ShowFrequent` | `DWord` | `0` | **Häufig verwendete Ordner** im Schnellzugriff/Explorer deaktivieren. |

### `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Search`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `SearchSystemDirs` | `DWord` | `1` | Systemverzeichnisse in die Suche einbeziehen. |
| `SearchCompressedFiles` | `DWord` | `1` | Komprimierte Dateien in die Suche einbeziehen. |
| `SearchAlways` | `DWord` | `1` | Immer Dateinamen und Inhalte suchen. |

## Word

### `HKCU:\Software\Microsoft\Office\16.0\Word\Options`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `DeveloperTools` | `DWord` | `1` | Entwicklertools aktivieren. |
| `Ruler` | `DWord` | `1` | Lineal einblenden. |
| `ShowAllFormatting` | `DWord` | `1` | Alle Formatierungszeichen anzeigen. |
| `VisiDrawTableDrs` | `DWord` | `1` | Tabellenränder/Zeichenhilfen anzeigen. |
| `DOC-PATH` | `String` / `ExpandString` | `Z:\` oder `$driveRoot` | Standardpfad für Dokumente. |
| `PersonalTemplates` | `String` / `ExpandString` | `$BackupTargetPath` | Pfad für persönliche Vorlagen. |
| `DisableBootToOfficeStart` | `DWord` | `1` | Office-Startbildschirm deaktivieren. |
| `DisableBackstageOpenKeyShortcuts` | `DWord` | `1` | Tastenkürzel im Backstage-Bereich deaktivieren. |
| `DefaultFont` | `String` | z. B. `Aptos` | Standardschrift für neue Dokumente. |
| `DefaultFontSize` | `DWord` | z. B. `11` | Standardschriftgröße für neue Dokumente. |
| `CorrectCapsLock` | `DWord` | `0` | Korrektur bei Caps Lock deaktivieren. |
| `AutoFormatApplyBulletedLists` | `DWord` | `0` | Automatische Aufzählungslisten deaktivieren. |
| `AutoFormatApplyNumberedLists` | `DWord` | `0` | Automatische Nummerierungslisten deaktivieren. |
| `AutoFormatCapitalizeTableCells` | `DWord` | `0` | Automatische Großschreibung in Tabellenzellen deaktivieren. |
| `CorrectTableCells` | `DWord` | `0` | Tabellenzellenkorrekturen deaktivieren. |
| `PictureInsertLayout` | `DWord` | `1` | Bildlayout beim Einfügen steuern. |
| `Font` | `String` | z. B. `Aptos` | Schriftname für die Standardkonfiguration. |
| `Fontsubstitutes` | `String` | leer | Schrift-Ersatzliste leeren. |

### `HKCU:\Software\Microsoft\Office\16.0\Word\Options`, `17.0`, `18.0`

Diese Werte werden zusätzlich für mehrere Word-Versionen gesetzt:

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `AutoFormatAsYouTypeApplyNumberedLists` | `DWord` | `0` | Nummerierte Listen während der Eingabe deaktivieren. |
| `AutoFormatAsYouTypeApplyBulletedLists` | `DWord` | `0` | Aufzählungslisten während der Eingabe deaktivieren. |
| `CorrectSentenceCaps` | `DWord` | `1` | Satzanfang automatisch korrigieren. |
| `AutoFormatAsYouTypeReplaceHyperlinks` | `DWord` | `0` | Hyperlink-Ersetzung während der Eingabe deaktivieren. |
| `CorrectInitialCaps` | `DWord` | `0` | Automatische Großschreibung am Wortanfang deaktivieren. |
| `AutoFormatAsYouTypeReplaceQuotes` | `DWord` | `1` | Anführungszeichen automatisch umwandeln. |
| `AutoFormatAsYouTypeReplaceSymbols` | `DWord` | `1` | Symbole automatisch umwandeln. |
| `PasteFormattingOtherApp` | `DWord` | `2` | Einfügeverhalten aus anderen Programmen steuern. |
| `PasteFormattingTwoDocumentsNoStyles` | `DWord` | `1` | Einfügeverhalten zwischen Dokumenten ohne Formatvorlagen. |

### Zusätzliche Schrift-Einstellungen für Word

#### `HKCU:\Software\Microsoft\Office\16.0\Word\Options` und `Common\LanguageResources`

#### außerdem für die Versionen `14.0`, `15.0`, `16.0`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `DefaultFont` | `String` | z. B. `Aptos` | Standardschrift beibehalten/setzen. |
| `DefaultFontSize` | `DWord` | z. B. `11` | Standardschriftgröße beibehalten/setzen. |

## Excel

### `HKCU:\Software\Microsoft\Office\16.0\Excel\Options`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `DeveloperTools` | `DWord` | `1` | Entwicklertools aktivieren. |
| `DefaultPath` | `String` / `ExpandString` | `Z:\` oder `$driveRoot` | Standardpfad für Arbeitsmappen/Dateien. |
| `PersonalTemplates` | `String` / `ExpandString` | `$BackupTargetPath` | Pfad für persönliche Vorlagen. |
| `DisableBootToOfficeStart` | `DWord` | `1` | Office-Startbildschirm deaktivieren. |
| `StandardFont` | `String` | z. B. `Aptos` | Standardschrift für neue Arbeitsmappen. |
| `StandardFontSize` | `DWord` | z. B. `10` | Standardschriftgröße für neue Arbeitsmappen. |
| `Font` | `String` | z. B. `Aptos,10` | Schrift-/Größenkombination für die Standardkonfiguration. |
| `AltStartupPath` | `String` / `ExpandString` | `$BackupTargetPath` | Zusätzlicher Startpfad für Excel. |
| `AutoSaveInterval` | `DWord` | `5` | AutoSpeichern-Intervall in Minuten. |

### `HKCU:\Software\Microsoft\Office\16.0\Excel\Options`, `15.0`, `14.0`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `StandardFont` | `String` | z. B. `Aptos` | Standardschrift für neue Arbeitsmappen. |
| `StandardFontSize` | `DWord` | z. B. `10` | Standardschriftgröße für neue Arbeitsmappen. |

### `HKCU:\Software\Microsoft\Office\16.0\Excel\Options` – Zusatz für Autokorrektur

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `CorrectSentenceCap` | `DWord` | `0` | Automatische Satzanfangskorrektur in Excel deaktivieren. |

## Outlook

### `HKCU:\Software\Microsoft\Office\$version\Outlook\Options\Calendar`

> Im Skript werden die Versionen `16.0`, `15.0` und `14.0` durchlaufen.

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `WeekNum` | `DWord` | `1` | Kalender zeigt Kalenderwochen an. |

### `HKCU:\Software\Microsoft\Office\$version\Outlook\Options`

| Wert | Typ | Standard im Skript | Kurzbeschreibung |
|---|---:|---:|---|
| `NewMailFont` | `String` | z. B. `Aptos` | Standardschrift für neue E-Mails. |
| `NewMailFontSize` | `DWord` | z. B. `11` | Schriftgröße für neue E-Mails. |
| `ReplyForwardFont` | `String` | z. B. `Aptos` | Standardschrift für Antworten und Weiterleitungen. |
| `ReplyForwardFontSize` | `DWord` | z. B. `11` | Schriftgröße für Antworten und Weiterleitungen. |
| `DefaultMailFont` | `String` | z. B. `Aptos` | Standard-Mail-Schrift für Outlook. |

## Hinweise

- Die Werte werden überwiegend unter `HKCU` gesetzt, also benutzerspezifisch.
- Einige Pfade werden nur gesetzt, wenn sie existieren; andere werden bei Bedarf angelegt.
- Schriftarten und Größen können je nach Auswahl im Skript variieren.
- Die Tabelle bildet den Stand aus `src/PC-Konfigurator.ps1` ab und kann sich mit künftigen Änderungen im Skript ändern.
