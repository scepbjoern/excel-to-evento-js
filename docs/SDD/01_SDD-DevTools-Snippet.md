# Solution Definition Document
## Variante E: DevTools Snippet

**Version:** 1.0  
**Datum:** 30. Januar 2026  
**Bezug:** PRD Excel-to-Evento Note Import Tool

---

## 1. Übersicht

### 1.1 Lösungsbeschreibung
Ein JavaScript-Snippet, das in der Browser-Konsole (Chrome DevTools) ausgeführt wird. Das Snippet lädt die SheetJS-Bibliothek dynamisch, erstellt einen File-Input für die Excel-Datei und führt das Matching sowie die DOM-Manipulation durch.

### 1.2 Zielgruppe
- Technisch versierte Benutzer
- Entwickler zum Testen und Prototyping

### 1.3 Voraussetzungen
- Google Chrome Browser
- Zugriff auf die EventoWeb-Qualifikationsseite
- Excel-Datei (.xlsx) mit korrekter Struktur

---

## 2. Architektur

### 2.1 Komponenten-Übersicht

```
┌─────────────────────────────────────────────────────────────┐
│                    DevTools Console                          │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────────┐    ┌─────────────────┐                │
│  │  SheetJS Loader │───▶│  Excel Parser   │                │
│  └─────────────────┘    └────────┬────────┘                │
│                                  │                          │
│                                  ▼                          │
│  ┌─────────────────┐    ┌─────────────────┐                │
│  │  DOM Scanner    │◀───│  Name Matcher   │                │
│  └────────┬────────┘    └────────┬────────┘                │
│           │                      │                          │
│           ▼                      ▼                          │
│  ┌─────────────────┐    ┌─────────────────┐                │
│  │  Grade Setter   │    │  Result Logger  │                │
│  └─────────────────┘    └─────────────────┘                │
│                                  │                          │
│                                  ▼                          │
│                         ┌─────────────────┐                │
│                         │  Excel Exporter │                │
│                         └─────────────────┘                │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### 2.2 Datenfluss

```
Excel-Datei (.xlsx)
       │
       ▼ [FileReader API]
┌──────────────────┐
│ ArrayBuffer      │
└────────┬─────────┘
         │
         ▼ [SheetJS XLSX.read()]
┌──────────────────┐
│ Workbook Object  │
└────────┬─────────┘
         │
         ▼ [Sheet "evento"]
┌──────────────────┐
│ Array of Rows    │
│ [{Kombi, Note}]  │
└────────┬─────────┘
         │
         ▼ [Matching]
┌──────────────────┐     ┌──────────────────┐
│ EventoWeb DOM    │────▶│ Matched Pairs    │
│ (Namen + Selects)│     │ + Not Found List │
└──────────────────┘     └────────┬─────────┘
                                  │
         ┌────────────────────────┴────────────────────────┐
         │                                                 │
         ▼                                                 ▼
┌──────────────────┐                            ┌──────────────────┐
│ DOM Manipulation │                            │ Excel mit        │
│ (Select Values)  │                            │ "Resultat"-Spalte│
└──────────────────┘                            └──────────────────┘
```

---

## 3. Detaillierte Spezifikation

### 3.1 Modul: SheetJS Loader

**Zweck:** Dynamisches Laden der SheetJS-Bibliothek

**Input:** Keine

**Output:** Promise, das resolved wenn XLSX global verfügbar ist

**Logik:**
```
1. Prüfe ob XLSX bereits geladen ist
2. Falls nein:
   a. Erstelle <script>-Element
   b. Setze src auf SheetJS CDN URL
   c. Füge zu document.head hinzu
   d. Warte auf load-Event
3. Resolve Promise
```

**CDN URL:** `https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js`

### 3.2 Modul: Excel Parser

**Zweck:** Einlesen und Parsen der Excel-Datei

**Input:** File-Objekt von File-Input

**Output:** Array von Objekten `[{name: string, grade: string, rowIndex: number}]`

**Logik:**
```
1. Lese Datei als ArrayBuffer (FileReader)
2. Parse mit XLSX.read(data, {type: 'array'})
3. Hole Sheet "evento" aus workbook.Sheets
4. Konvertiere zu JSON mit XLSX.utils.sheet_to_json()
5. Extrahiere Spalten "Kombi" und "Note"
6. Validiere Noten (muss zwischen 1.0 und 6.0 sein)
7. Gib Array zurück
```

**Fehlerbehandlung:**
- Sheet "evento" nicht gefunden → Fehlermeldung
- Spalte "Kombi" oder "Note" fehlt → Fehlermeldung
- Ungültige Note → Warnung, Zeile überspringen

### 3.3 Modul: DOM Scanner

**Zweck:** Extrahieren aller Studierenden-Namen und zugehörigen Select-Elemente aus der Seite

**Input:** Keine (arbeitet auf document)

**Output:** Map `{name: string → selectElement: HTMLSelectElement}`

**Logik:**
```
1. Finde alle <tr>-Elemente in der Tabelle
2. Für jede Zeile:
   a. Finde <a> mit href enthält "Brn_PersonDetailMA"
   b. Falls gefunden: Extrahiere textContent (= Name)
   c. Finde <select> in derselben Zeile
   d. Falls beides vorhanden: Speichere in Map
3. Gib Map zurück
```

**Selektor für Namen:**
```javascript
a[href*="Brn_PersonDetailMA"]
```

**Selektor für Select:**
```javascript
select[name][id]  // mit numerischer ID
```

### 3.4 Modul: Name Matcher

**Zweck:** Abgleich der Excel-Namen mit den DOM-Namen

**Input:** 
- Excel-Daten: `[{name, grade, rowIndex}]`
- DOM-Map: `{name → selectElement}`

**Output:** 
- Matches: `[{excelName, grade, selectElement, rowIndex}]`
- NotFound: `[{name, grade, rowIndex}]`

**Logik:**
```
1. Für jeden Excel-Eintrag:
   a. Suche exakten Match in DOM-Map (case-sensitive)
   b. Falls gefunden: Zu Matches hinzufügen
   c. Falls nicht gefunden: Zu NotFound hinzufügen
2. Gib beide Listen zurück
```

**Wichtig:** Kein Fuzzy-Matching! Nur exakte Übereinstimmung.

### 3.5 Modul: Grade Setter

**Zweck:** Setzen der Noten in den Select-Dropdowns

**Input:** Matches-Array

**Output:** Anzahl erfolgreich gesetzter Noten

**Noten-Mapping:**
```javascript
const gradeToValue = {
  '1':   '10', '1.0': '10',
  '1.5': '11',
  '2':   '12', '2.0': '12',
  '2.5': '16',
  '3':   '17', '3.0': '17',
  '3.5': '18',
  '4':   '19', '4.0': '19',
  '4.5': '20',
  '5':   '21', '5.0': '21',
  '5.5': '22',
  '6':   '23', '6.0': '23'
};
```

**Logik:**
```
1. Für jeden Match:
   a. Normalisiere Note (z.B. "5" → "5.0")
   b. Hole Value aus gradeToValue
   c. Setze selectElement.value = value
   d. Triggere 'change'-Event (für ASP.NET)
2. Zähle erfolgreiche Setzungen
```

### 3.6 Modul: Result Logger

**Zweck:** Anzeige der Ergebnisse in der Konsole

**Input:** Matches, NotFound, Statistiken

**Output:** Console-Ausgabe (formatiert)

**Format:**
```
═══════════════════════════════════════
  EVENTO NOTE IMPORT - ERGEBNIS
═══════════════════════════════════════
✓ Erfolgreich zugeordnet: 25
✗ Nicht gefunden: 3

Nicht gefundene Namen:
  - Müller Hans (Note: 5.0)
  - Meier Anna (Note: 4.5)
  - Schmidt Peter (Note: 3.5)
═══════════════════════════════════════
```

### 3.7 Modul: Excel Exporter

**Zweck:** Erstellen einer neuen Excel-Datei mit Resultat-Spalte

**Input:** Original-Workbook, Matches, NotFound

**Output:** Download der aktualisierten Excel-Datei

**Logik:**
```
1. Klone Original-Sheet
2. Füge Spalte "Resultat" hinzu (Spalte C)
3. Für jede Zeile:
   a. Falls in Matches: Setze "Matched"
   b. Falls in NotFound: Setze "Not Found"
4. Erstelle neuen Workbook
5. Konvertiere zu Blob
6. Triggere Download als "evento_result.xlsx"
```

---

## 4. User Interface

### 4.1 Interaktion

Da es sich um ein DevTools-Snippet handelt, ist das UI minimal:

1. **Start:** Benutzer führt Snippet in Konsole aus
2. **Dateiauswahl:** Ein File-Input-Dialog öffnet sich
3. **Feedback:** Ergebnisse werden in Konsole geloggt
4. **Download:** Aktualisierte Excel-Datei wird automatisch heruntergeladen

### 4.2 Visuelles Feedback (optional)

Für bessere Sichtbarkeit können gesetzte Noten visuell markiert werden:
```javascript
selectElement.style.backgroundColor = '#90EE90'; // Light green
```

---

## 5. Implementierungsplan für LLM

### 5.1 Schritt-für-Schritt Implementierung

**Phase 1: Setup und Bibliothek laden**
```
Aufgabe: Erstelle eine async Funktion loadSheetJS()
- Prüfe ob window.XLSX existiert
- Falls nicht: Erstelle script-Element mit CDN-URL
- Warte auf load-Event mit Promise
- Gib true zurück wenn geladen

CDN: https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js
```

**Phase 2: File-Input und Excel-Parsing**
```
Aufgabe: Erstelle Funktion createFileInput() und parseExcel(file)
- createFileInput(): Erstellt unsichtbaren input[type=file], .accept=".xlsx"
- parseExcel(file): 
  - Nutze FileReader.readAsArrayBuffer
  - Parse mit XLSX.read(arrayBuffer, {type: 'array'})
  - Hole Sheet "evento"
  - Konvertiere mit XLSX.utils.sheet_to_json(sheet)
  - Extrahiere "Kombi" und "Note" Spalten
  - Gib Array zurück: [{name, grade, rowIndex}]
```

**Phase 3: DOM-Scanning**
```
Aufgabe: Erstelle Funktion scanEventoPage()
- Finde alle <a> mit href enthält "Brn_PersonDetailMA"
- Für jedes a: 
  - Hole Name aus textContent (trim)
  - Finde parent <tr>
  - Finde <select> in dieser tr
  - Speichere in Map: name → selectElement
- Gib Map zurück
```

**Phase 4: Matching**
```
Aufgabe: Erstelle Funktion matchNames(excelData, domMap)
- Iteriere über excelData
- Für jeden Eintrag: Prüfe ob name exakt in domMap existiert
- Erstelle matches[] und notFound[] Arrays
- Gib {matches, notFound} zurück
```

**Phase 5: Noten setzen**
```
Aufgabe: Erstelle Funktion setGrades(matches)
- Definiere gradeToValue Mapping (siehe 3.5)
- Für jeden Match:
  - Normalisiere Note (z.B. "5" → "5.0", "5,0" → "5.0")
  - Hole value aus Mapping
  - Setze select.value = value
  - Dispatch 'change' Event
- Optional: Setze grünen Hintergrund
- Gib Anzahl erfolgreicher Setzungen zurück
```

**Phase 6: Excel-Export**
```
Aufgabe: Erstelle Funktion exportResultExcel(workbook, matches, notFound)
- Hole Sheet "evento"
- Bestimme letzte Spalte, füge "Resultat" Header hinzu
- Für jede Datenzeile: Setze "Matched" oder "Not Found"
- Erstelle Blob mit XLSX.write(workbook, {type: 'array'})
- Erstelle Download-Link und klicke ihn
- Dateiname: "evento_result.xlsx"
```

**Phase 7: Hauptfunktion**
```
Aufgabe: Erstelle async Funktion main()
1. await loadSheetJS()
2. Erstelle FileInput und warte auf Dateiauswahl
3. const excelData = await parseExcel(file)
4. const domMap = scanEventoPage()
5. const {matches, notFound} = matchNames(excelData, domMap)
6. const count = setGrades(matches)
7. logResults(matches, notFound)
8. exportResultExcel(workbook, matches, notFound)
```

**Phase 8: Self-executing wrapper**
```
Aufgabe: Wrap alles in IIFE
(async function() {
  // ... gesamter Code ...
})();
```

### 5.2 Code-Struktur

```javascript
(async function() {
  'use strict';
  
  // ========== KONFIGURATION ==========
  const CONFIG = {
    SHEETJS_CDN: 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js',
    SHEET_NAME: 'evento',
    COLUMN_NAME: 'Kombi',
    COLUMN_GRADE: 'Note'
  };
  
  const GRADE_MAP = { /* ... */ };
  
  // ========== HILFSFUNKTIONEN ==========
  async function loadSheetJS() { /* ... */ }
  function createFileInput() { /* ... */ }
  async function parseExcel(file) { /* ... */ }
  function scanEventoPage() { /* ... */ }
  function matchNames(excelData, domMap) { /* ... */ }
  function setGrades(matches) { /* ... */ }
  function logResults(matches, notFound) { /* ... */ }
  function exportResultExcel(workbook, matches, notFound) { /* ... */ }
  
  // ========== HAUPTPROGRAMM ==========
  async function main() { /* ... */ }
  
  // Start
  main().catch(console.error);
})();
```

---

## 6. Testing-Anleitung

### 6.1 Voraussetzungen

1. **Browser:** Google Chrome (aktuelle Version)
2. **Zugang:** EventoWeb ZHAW mit gültiger Anmeldung
3. **Testdaten:** Excel-Datei mit Tabellenblatt "evento"

### 6.2 Test-Excel erstellen

Erstelle eine Datei `test_noten.xlsx` mit Tabellenblatt "evento":

| Kombi | Note |
|-------|------|
| [Exakter Name aus EventoWeb] | 5.0 |
| [Exakter Name aus EventoWeb] | 4.5 |
| Test Nicht Vorhanden | 3.0 |

**Wichtig:** Kopiere die Namen exakt aus EventoWeb (inklusive Gross-/Kleinschreibung).

### 6.3 Schritt-für-Schritt Test

#### Test 1: Snippet laden und SheetJS laden

**Schritte:**
1. Öffne EventoWeb Qualifikationsseite
2. Öffne DevTools (F12)
3. Gehe zum Tab "Console"
4. Füge das Snippet ein und drücke Enter

**Erwartetes Ergebnis:**
- Meldung "SheetJS wird geladen..."
- Meldung "SheetJS geladen!"
- File-Dialog öffnet sich

#### Test 2: Excel einlesen

**Schritte:**
1. Wähle die Test-Excel-Datei
2. Bestätige Auswahl

**Erwartetes Ergebnis:**
- Meldung "Excel-Datei eingelesen: X Einträge"
- Keine Fehlermeldungen

#### Test 3: Matching prüfen

**Erwartetes Ergebnis:**
- Meldung zeigt Anzahl Matches und Not Found
- Konsolenausgabe listet nicht gefundene Namen auf
- "Test Nicht Vorhanden" erscheint in Not Found

#### Test 4: Noten in Dropdowns prüfen

**Schritte:**
1. Scrolle durch die EventoWeb-Seite
2. Prüfe die Select-Dropdowns der gematchten Namen

**Erwartetes Ergebnis:**
- Die korrekten Noten sind ausgewählt
- Optional: Grüner Hintergrund bei geänderten Selects

#### Test 5: Excel-Export prüfen

**Schritte:**
1. Öffne die heruntergeladene "evento_result.xlsx"
2. Prüfe die Spalte "Resultat"

**Erwartetes Ergebnis:**
- Spalte "Resultat" existiert
- Gematchte Zeilen haben "Matched"
- "Test Nicht Vorhanden" hat "Not Found"

### 6.4 Fehlerszenarien testen

#### Test E1: Falsches Sheet

**Schritte:**
- Excel ohne Sheet "evento" verwenden

**Erwartetes Ergebnis:**
- Fehlermeldung: "Sheet 'evento' nicht gefunden"

#### Test E2: Fehlende Spalten

**Schritte:**
- Excel ohne Spalte "Kombi" oder "Note"

**Erwartetes Ergebnis:**
- Fehlermeldung: "Spalte 'Kombi' nicht gefunden" oder "Spalte 'Note' nicht gefunden"

#### Test E3: Ungültige Note

**Schritte:**
- Excel mit Note "7.0" oder "abc"

**Erwartetes Ergebnis:**
- Warnung in Konsole
- Zeile wird übersprungen oder als Fehler markiert

### 6.5 Checkliste

| # | Test | Status |
|---|------|--------|
| 1 | SheetJS lädt korrekt | ☐ |
| 2 | File-Dialog öffnet sich | ☐ |
| 3 | Excel wird eingelesen | ☐ |
| 4 | Namen werden korrekt extrahiert | ☐ |
| 5 | DOM-Scanning findet alle Studierenden | ☐ |
| 6 | Exakte Matches werden erkannt | ☐ |
| 7 | Nicht gefundene Namen werden gelistet | ☐ |
| 8 | Noten werden korrekt gesetzt | ☐ |
| 9 | Noten-Mapping ist korrekt (1.0→10, ..., 6.0→23) | ☐ |
| 10 | Excel-Export funktioniert | ☐ |
| 11 | Resultat-Spalte ist korrekt | ☐ |
| 12 | Fehlerbehandlung funktioniert | ☐ |

---

## 7. Bekannte Einschränkungen

1. **Keine Persistenz:** Snippet muss bei jedem Besuch neu eingefügt werden
2. **Manueller Prozess:** DevTools müssen geöffnet werden
3. **Technisches Know-how:** Benutzer muss mit DevTools umgehen können
4. **Kein UI:** Nur Konsolenausgabe, kein grafisches Interface

---

## 8. Erweiterungsmöglichkeiten

1. **Als Chrome Snippet speichern:** DevTools → Sources → Snippets
2. **Besseres Feedback:** Alert-Dialoge statt nur Konsole
3. **Undo-Funktion:** Ursprüngliche Select-Werte speichern
4. **Dry-Run Modus:** Erst Preview, dann Ausführung
