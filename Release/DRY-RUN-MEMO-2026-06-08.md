# Dry-Run Memo (read-only) – 2026-06-08

## Zweck
Dieser Report dokumentiert den ursprünglichen Registry-Soll/Ist-Abgleich vom **2026-06-08** (ohne Änderungen am System) sowie die Fortschreibungen vom **2026-06-12**, damit der Stand vor einem produktiven Lauf nachvollziehbar bleibt.

## Executive Summary
- **Historischer Referenzstand (2026-06-08):** 69 Sollwerte geprüft, 22 Abweichungen.
- **Aktueller Stand (2026-06-12):** Follow-up- und Vollabgleich dokumentiert; die 6 Office-16-DIFFs wurden nachkorrigiert und als **OK** verifiziert.
- **Verbleibender Restbestand:** überwiegend erwartbare `MISSING_PATH`-Treffer für nicht installierte Zukunftsversionen (Office 17/18).

## Zusammenfassung
- Geprüfte feste Sollwerte gesamt: **69**
- Abweichungen gesamt: **22**

Aufteilung:
1. **Windows/Explorer**: 22 geprüft, 2 Abweichungen
2. **Office 16**: 27 geprüft, 0 Abweichungen
3. **Multi-Version (Office 17/18 + Outlook 14/15)**: 20 geprüft, 20 Abweichungen

## Konkrete echte Wertabweichungen
1. `HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\SearchboxTaskbarMode`
   - Erwartet: `1`
   - Ist: `0`
2. `HKCU:\Software\Microsoft\Windows\CurrentVersion\Search\SearchboxTaskbarModeCache`
   - Erwartet: `1`
   - Ist: `0`

## Multi-Version-Abweichungen (Einordnung)
- **19x MISSING_PATH** (typisch bei nicht installierten Office-Hauptversionen 17/18 bzw. Outlook 14)
- **1x MISSING_VALUE**:
  - `HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Calendar\WeekNum` fehlt

Zusatzprüfung:
- `HKCU:\Software\Microsoft\Office\14.0\Outlook\Options\Calendar` → Pfad fehlt
- `HKCU:\Software\Microsoft\Office\15.0\Outlook\Options\Calendar` → Pfad vorhanden, `WeekNum` fehlt

## Wichtiger Hinweis
Dynamische Laufzeitwerte wurden absichtlich **nicht** als feste Sollwerte bewertet, z. B.:
- `DOC-PATH`, `DefaultPath`, `PersonalTemplates`, `AltStartupPath`
- Schriftart-/Schriftgrößenwerte abhängig von Nutzerauswahl
- Outlook-Font-Keys abhängig von Laufparametern

## Empfehlung für produktiven Lauf
Vor dem produktiven Lauf optional nachziehen:
- `SearchboxTaskbarMode = 1`
- `SearchboxTaskbarModeCache = 1`
- optional (nur falls Office 15 aktiv genutzt): `WeekNum = 1` unter Outlook 15 Kalenderpfad

## Fortschreibung (Stand 2026-06-12)

### Einordnung
- In dieser Fortschreibung wurden ein **gezielter Read-only-Folgecheck** sowie ein **vollständiger 69er Read-only-Vollabgleich** dokumentiert.
- Die obere **Zusammenfassung** bleibt als historischer Referenzstand vom **2026-06-08** unverändert erhalten.

### Seitdem umgesetzte Skript-Anpassungen (relevant für Folgeläufe)
- Outlook-Font-Setzung erfolgt nun **nach** der Benutzerabfrage und wird explizit mit denselben Werten wie Word aufgerufen:
   - `Set-OutlookRegistry -FontName $ActualFontName -FontSize $FontSizeWord`
- Zusätzlicher Logeintrag ergänzt:
   - `Outlook-Schrift wurde mit Word synchronisiert: <Font> / <Size> pt`
- Fallback-Standardwerte (bei übersprungener Auswahl) sind jetzt:
   - **Schriftart:** `Aptos`
   - **Word/Outlook:** `11 pt`
   - **Excel:** `10 pt`

### Erwartung für den nächsten Dry-Run
- Die Registry-Prüfung sollte weiterhin dynamische Font-/Size-Werte nur kontextabhängig bewerten.
- Bei produktivem Lauf mit Standards (Aptos/11 für Word+Outlook, Excel 10) ist die Konfiguration konsistent zum aktuellen Skriptstand.

## Read-only Folgecheck (2026-06-12)

### Umfang
- Gezielt geprüfte Sollwerte: **11** (Windows-Suche, Outlook/Word/Excel 16.0, Outlook 15 Kalender)

### Ergebnis
- **OK:** 7
- **Abweichung (DIFF):** 4
- **MISSING_PATH:** 0
- **MISSING_VALUE:** 0

Konkrete DIFFs:
1. `HKCU:\Software\Microsoft\Office\16.0\Word\Options\DefaultFont`
   - Erwartet: `Aptos`
   - Ist: leer
2. `HKCU:\Software\Microsoft\Office\16.0\Word\Options\DefaultFontSize`
   - Erwartet: `11`
   - Ist: `0`
3. `HKCU:\Software\Microsoft\Office\16.0\Excel\Options\StandardFont`
   - Erwartet: `Aptos`
   - Ist: leer
4. `HKCU:\Software\Microsoft\Office\16.0\Excel\Options\StandardFontSize`
   - Erwartet: `10`
   - Ist: `0`

Bestätigte OK-Werte (Auszug):
- `SearchboxTaskbarMode = 1`
- `SearchboxTaskbarModeCache = 1`
- Outlook 16 (`NewMailFont`, `NewMailFontSize`, `ReplyForwardFont`, `ReplyForwardFontSize`) entspricht `Aptos` / `11`
- Outlook 15 Kalender: `WeekNum = 1`

### Hinweis zur Interpretation
- Dieser Folgecheck ist ein **gezielter Spot-Check** und ersetzt nicht den vollständigen 69er Abgleich vom 2026-06-08.
- Die Word-/Excel-DIFFs zeigen einen aktuellen Registry-Iststand, der von den im Skript definierten Fallback-Sollwerten abweicht.

## Vollabgleich (69) – 2026-06-12

### Ergebnis gesamt
- Geprüfte feste Sollwerte: **69**
- **OK:** 45
- **DIFF:** 6
- **MISSING_PATH:** 18
- **MISSING_VALUE:** 0

### Aufteilung nach Bereichen
1. **Windows/Explorer**: 22 geprüft, **0 Abweichungen**
2. **Office 16**: 27 geprüft, **6 Abweichungen (DIFF)**
3. **Multi-Version (Office 17/18 + Outlook 14/15)**: 20 geprüft, **18 Abweichungen (alle MISSING_PATH)**

### Konkrete DIFFs (echte Wertabweichungen)
1. `HKCU:\Software\Microsoft\Office\16.0\Word\Options\DefaultFont`
   - Erwartet: `Aptos`
   - Ist: leer
2. `HKCU:\Software\Microsoft\Office\16.0\Word\Options\DefaultFontSize`
   - Erwartet: `11`
   - Ist: `0`
3. `HKCU:\Software\Microsoft\Office\16.0\Word\Options\Font`
   - Erwartet: `Aptos`
   - Ist: leer
4. `HKCU:\Software\Microsoft\Office\16.0\Excel\Options\StandardFont`
   - Erwartet: `Aptos`
   - Ist: leer
5. `HKCU:\Software\Microsoft\Office\16.0\Excel\Options\StandardFontSize`
   - Erwartet: `10`
   - Ist: `0`
6. `HKCU:\Software\Microsoft\Office\16.0\Excel\Options\Font`
   - Erwartet: `Aptos,10`
   - Ist: `Aptos Narrow,10`

### Multi-Version-Einordnung (MISSING_PATH)
- Es fehlen die beiden Word-Pfade:
   - `HKCU:\Software\Microsoft\Office\17.0\Word\Options`
   - `HKCU:\Software\Microsoft\Office\18.0\Word\Options`
- Dadurch ergeben sich dort insgesamt **18x MISSING_PATH** (je 9 erwartete Werte pro Pfad).
- **Fachliche Bewertung:** Diese `MISSING_PATH`-Treffer sind hier **erwartbar** und stellen aktuell **keinen Konfigurationsfehler** dar, solange Office 17/18 auf dem System nicht installiert ist (Zukunftsversionen).
- **Vorbereitung ist vorhanden:** Die Sollwerte bleiben bewusst im Prüfkatalog enthalten. Sobald eine entsprechende Office-Version installiert ist und die Pfade existieren, werden Abweichungen sofort sichtbar und können ohne Doku-Änderung nachgezogen werden.
- Für Outlook 14/15 Kalender im 69er Satz ist in diesem Lauf **kein zusätzlicher fehlender Wert** aufgetreten.

### Vergleich zum Stand 2026-06-08
- Gesamtabweichungen im Vollabgleich vom 2026-06-12: **22 → 24** (formal +2)
- Gleichzeitig haben sich die damaligen Windows-Abweichungen (`SearchboxTaskbarMode`, `SearchboxTaskbarModeCache`) auf **OK** verbessert.
- Die Erhöhung kommt aus dem Office-16-Block (Font-/Size-Werte), nicht aus Windows/Explorer.
- Interpretation bleibt konsistent: fehlende 17/18-Pfade sind installationsabhängige Strukturabweichungen; die fachlich relevanten Unterschiede lagen in diesem Lauf bei den Office-16-Fontwerten.
- Nach der dokumentierten Nachkorrektur der 6 Office-16-DIFFs verbleiben als erwartbarer Restbestand vor allem die `MISSING_PATH`-Treffer für nicht installierte Zukunftsversionen (Office 17/18).

## Nachkorrektur der konkreten DIFFs (2026-06-12)

Die sechs zuvor dokumentierten Office-16-DIFFs wurden gezielt nachgezogen und anschließend read-only verifiziert.

### Verifikationsstatus
- `DefaultFont` (Word 16.0): **OK** (`Aptos`)
- `DefaultFontSize` (Word 16.0): **OK** (`11`)
- `Font` (Word 16.0): **OK** (`Aptos`)
- `StandardFont` (Excel 16.0): **OK** (`Aptos`)
- `StandardFontSize` (Excel 16.0): **OK** (`10`)
- `Font` (Excel 16.0): **OK** (`Aptos,10`)

Zusätzlich wurde im Skript ein Ursachenfix umgesetzt: Die Office-/Font-Registry-Setzung erfolgt jetzt erst **nach** der Font-Auswahl (gewählt oder Default), damit keine leeren bzw. `0`-Werte geschrieben werden.
