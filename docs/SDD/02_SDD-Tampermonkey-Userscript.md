# Solution Definition Document
## Variante B: Tampermonkey Userscript

**Version:** 1.0  
**Datum:** 30. Januar 2026  
**Bezug:** PRD Excel-to-Evento Note Import Tool

---

## 1. Übersicht

### 1.1 Lösungsbeschreibung
Ein Tampermonkey Userscript, das automatisch auf der EventoWeb-Qualifikationsseite aktiviert wird und ein grafisches UI-Panel einblendet. Das Panel ermöglicht das Hochladen einer Excel-Datei und zeigt den Import-Status übersichtlich an.

### 1.2 Zielgruppe
- Dozenten der ZHAW
- Regelmässige Nutzer des Notenimports
- Benutzer ohne tiefe technische Kenntnisse

### 1.3 Voraussetzungen
- Google Chrome oder Firefox
- Tampermonkey Extension installiert
- Zugriff auf die EventoWeb-Qualifikationsseite
- Excel-Datei (.xlsx) mit korrekter Struktur

---

## 2. Architektur

### 2.1 Komponenten-Übersicht

```
┌─────────────────────────────────────────────────────────────┐
│                  Tampermonkey Userscript                    │
├─────────────────────────────────────────────────────────────┤
│  // ==UserScript==                                          │
│  // @name         Evento Note Importer                      │
│  // @match        *://eventoweb.zhaw.ch/*                   │
│  // @require      SheetJS CDN                               │
│  // ==/UserScript==                                         │
├─────────────────────────────────────────────────────────────┤
│                                                              │
│  ┌─────────────────────────────────────────────────────┐   │
│  │                    UI Module                         │   │
│  │  ┌─────────┐  ┌─────────┐  ┌─────────────────────┐  │   │
│  │  │ Toggle  │  │ Upload  │  │ Status/Result Panel │  │   │
│  │  │ Button  │  │ Area    │  │                     │  │   │
│  │  └─────────┘  └─────────┘  └─────────────────────┘  │   │
│  └─────────────────────────────────────────────────────┘   │
│                           │                                 │
│                           ▼                                 │
│  ┌─────────────────────────────────────────────────────┐   │
│  │                  Core Modules                        │   │
│  │  ┌────────────┐  ┌────────────┐  ┌────────────┐    │   │
│  │  │ Excel      │  │ DOM        │  │ Name       │    │   │
│  │  │ Parser     │  │ Scanner    │  │ Matcher    │    │   │
│  │  └────────────┘  └────────────┘  └────────────┘    │   │
│  │  ┌────────────┐  ┌────────────┐                    │   │
│  │  │ Grade      │  │ Excel      │                    │   │
│  │  │ Setter     │  │ Exporter   │                    │   │
│  │  └────────────┘  └────────────┘                    │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### 2.2 Userscript Header

```javascript
// ==UserScript==
// @name         Evento Note Importer
// @namespace    https://zhaw.ch/
// @version      1.0
// @description  Importiert Noten aus Excel in EventoWeb
// @author       ZHAW
// @match        *://eventoweb.zhaw.ch/cst_pages/brn_qualifikationdurchdozenten.aspx*
// @require      https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js
// @grant        none
// ==/UserScript==
```

### 2.3 URL-Matching

Das Script aktiviert sich nur auf der Qualifikationsseite:
- Pattern: `*://eventoweb.zhaw.ch/cst_pages/brn_qualifikationdurchdozenten.aspx*`
- Damit wird sichergestellt, dass das UI nur dort erscheint

---

## 3. User Interface Spezifikation

### 3.1 UI-Elemente

#### 3.1.1 Toggle-Button (minimiert)

```
┌─────────────────────┐
│ 📋 Noten Import     │
└─────────────────────┘
```

- **Position:** Rechts unten, fixed
- **Verhalten:** Klick öffnet das Hauptpanel
- **Styling:** Blauer Button, weisse Schrift, leichter Schatten

#### 3.1.2 Hauptpanel (expandiert)

```
┌────────────────────────────────────────────────┐
│ 📋 Evento Note Importer                    [X] │
├────────────────────────────────────────────────┤
│                                                │
│  ┌──────────────────────────────────────────┐ │
│  │                                          │ │
│  │     📁 Excel-Datei hier ablegen          │ │
│  │        oder klicken zum Auswählen        │ │
│  │                                          │ │
│  └──────────────────────────────────────────┘ │
│                                                │
│  ──────────── Status ────────────             │
│                                                │
│  Bereit. Bitte Excel-Datei hochladen.         │
│                                                │
│  ──────────── Ergebnis ────────────           │
│                                                │
│  (noch keine Daten)                           │
│                                                │
│  ┌────────────┐  ┌────────────────────────┐  │
│  │  Anwenden  │  │  Excel herunterladen   │  │
│  └────────────┘  └────────────────────────┘  │
│                                                │
└────────────────────────────────────────────────┘
```

- **Position:** Rechts unten, fixed
- **Breite:** 400px
- **Höhe:** Auto (max 80vh mit Scroll)
- **z-index:** 10000 (über allem anderen)

#### 3.1.3 Ergebnis-Ansicht (nach Import)

```
┌────────────────────────────────────────────────┐
│ 📋 Evento Note Importer                    [X] │
├────────────────────────────────────────────────┤
│                                                │
│  ✅ Datei geladen: noten.xlsx                  │
│                                                │
│  ──────────── Ergebnis ────────────           │
│                                                │
│  ✓ Gefunden: 25                               │
│  ✗ In Excel, nicht in Evento: 2               │
│  ✗ In Evento, nicht in Excel: 1               │
│  ⚠ Ungültige Noten: 1                          │
│                                                │
│  In Excel, nicht in Evento:                   │
│  ┌──────────────────────────────────────────┐ │
│  │ • Müller Hans (5.0)                      │ │
│  │ • Meier Anna (4.5)                       │ │
│  └──────────────────────────────────────────┘ │
│                                                │
│  In Evento, nicht in Excel:                     │
│  ┌──────────────────────────────────────────┐ │
│  │ • Weber Lisa                             │ │
│  └──────────────────────────────────────────┘ │
│                                                │
│  Ungültige Noten (nicht gesetzt):             │
│  ┌──────────────────────────────────────────┐ │
│  │ • Schmidt Karl: "2.3" (Zeile 8)          │ │
│  └──────────────────────────────────────────┘ │
│                                                │
│  ┌────────────┐  ┌────────────────────────┐  │
│  │  Anwenden  │  │  Excel herunterladen   │  │
│  └────────────┘  └────────────────────────┘  │
│                                                │
└────────────────────────────────────────────────┘
```

**Anzeige-Kategorien:**
- **Gefunden**: Namen die sowohl in Excel als auch in Evento vorhanden sind (Noten werden gesetzt)
- **In Excel, nicht in Evento**: Namen aus Excel die nicht in EventoWeb gefunden wurden
- **In Evento, nicht in Excel**: Namen aus EventoWeb die nicht in der Excel-Datei vorhanden sind
- **Ungültige Noten**: Einträge mit ungültigen Noten (z.B. 2.3) - nur halbe Noten erlaubt

### 3.2 Farbschema

| Element | Farbe | Hex |
|---------|-------|-----|
| Panel Hintergrund | Weiss | #FFFFFF |
| Header Hintergrund | ZHAW Blau | #0064A3 |
| Header Text | Weiss | #FFFFFF |
| Erfolg | Grün | #28A745 |
| Fehler/Warnung | Rot | #DC3545 |
| Button Primary | ZHAW Blau | #0064A3 |
| Button Secondary | Grau | #6C757D |
| Border | Hellgrau | #DEE2E6 |

### 3.3 Workflow-Zustände

```
                    ┌─────────────┐
                    │   Bereit    │
                    └──────┬──────┘
                           │ Datei hochladen
                           ▼
                    ┌─────────────┐
                    │   Parsing   │
                    └──────┬──────┘
                           │
              ┌────────────┴────────────┐
              │                         │
              ▼                         ▼
       ┌─────────────┐          ┌─────────────┐
       │   Fehler    │          │  Analysiert │
       └─────────────┘          └──────┬──────┘
                                       │ "Anwenden" klicken
                                       ▼
                                ┌─────────────┐
                                │ Angewendet  │
                                └──────┬──────┘
                                       │
                                       ▼
                                ┌─────────────┐
                                │  Download   │
                                └─────────────┘
```

---

## 4. Detaillierte Spezifikation

### 4.1 Modul: UI Builder

**Zweck:** Erstellen und Einfügen des UI in die Seite

**Funktionen:**
- `createPanel()` - Erstellt das Hauptpanel
- `createToggleButton()` - Erstellt den Toggle-Button
- `createDropZone()` - Erstellt den Drag&Drop-Bereich
- `createStatusArea()` - Erstellt den Status-Bereich
- `createResultArea()` - Erstellt den Ergebnis-Bereich
- `createButtons()` - Erstellt die Action-Buttons

**CSS-Injection:**
```javascript
const styles = `
  #evento-importer-panel {
    position: fixed;
    bottom: 20px;
    right: 20px;
    width: 400px;
    /* ... weitere Styles ... */
  }
`;
GM_addStyle(styles); // oder style-Element einfügen
```

### 4.2 Modul: Drag & Drop Handler

**Zweck:** Verarbeiten von Drag&Drop und File-Input

**Events:**
- `dragover` - Visuelles Feedback
- `dragleave` - Feedback entfernen
- `drop` - Datei verarbeiten
- `change` (input) - Datei aus Dialog

**Validierung:**
- Nur .xlsx-Dateien akzeptieren
- Dateigrösse prüfen (max 10MB)

### 4.3 Modul: Excel Parser

**Zweck:** Einlesen und Parsen der Excel-Datei

**Input:** File-Objekt

**Output:** 
```javascript
{
  success: boolean,
  data: [{name: string, grade: string, rowIndex: number}],
  invalidGrades: [{name: string, grade: string, rowIndex: number, reason: string}],
  workbook: XLSX.WorkBook,
  error: string | null
}
```

**Validierung:**
- Sheet "evento" muss existieren
- Spalte "Kombi" muss existieren
- Spalte "Note" muss existieren
- Noten müssen gültig sein (nur exakte halbe Noten: 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0)
- **Keine Rundung:** Werte wie 2.3 werden als ungültig markiert

### 4.4 Modul: DOM Scanner

**Zweck:** Extrahieren aller Studierenden aus der Seite

**Output:**
```javascript
Map<string, {
  name: string,
  selectElement: HTMLSelectElement,
  row: HTMLTableRowElement
}>
```

**Selektoren:**
```javascript
// Name finden
const nameLinks = document.querySelectorAll('a[href*="Brn_PersonDetailMA"]');

// Select finden (in gleicher Zeile)
const row = nameLink.closest('tr');
const select = row.querySelector('select');
```

### 4.5 Modul: Name Matcher

**Zweck:** Exakter Abgleich der Namen

**Algorithmus:**
```
1. Für jeden Excel-Eintrag:
   a. Trimme den Namen
   b. Suche exakten Match in DOM-Map
   c. Kategorisiere als "matched" oder "notFound"
2. Finde DOM-Einträge die nicht in Excel sind:
   a. Iteriere über alle DOM-Namen
   b. Prüfe ob Name in Excel-Daten existiert
   c. Falls nicht: Kategorisiere als "notInExcel"
3. Gib Ergebnis zurück
```

**Output:**
```javascript
{
  matched: [{excelName, grade, selectElement, rowIndex}],
  notFound: [{name, grade, rowIndex}],
  notInExcel: [{name, selectElement}]
}
```

**Kein Fuzzy-Matching!** Nur exakte String-Übereinstimmung.

### 4.6 Modul: Grade Setter

**Zweck:** Setzen der Noten in den Dropdowns

**Noten-Mapping:**
```javascript
const GRADE_TO_VALUE = {
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

**Ablauf:**
```javascript
function setGrade(selectElement, grade) {
  const normalizedGrade = normalizeGrade(grade);
  const value = GRADE_TO_VALUE[normalizedGrade];
  
  if (value && selectElement) {
    selectElement.value = value;
    selectElement.dispatchEvent(new Event('change', { bubbles: true }));
    highlightElement(selectElement); // visuelles Feedback
    return true;
  }
  return false;
}
```

### 4.7 Modul: Excel Exporter

**Zweck:** Erstellen der aktualisierten Excel-Datei

**Ablauf:**
```javascript
function exportResultExcel(workbook, results, invalidGrades, notInExcel) {
  const sheet = workbook.Sheets['evento'];
  
  // Finde/erstelle Spalte "Resultat" im Original-Sheet
  const resultCol = findOrCreateResultColumn(sheet);
  
  // Setze Werte im Original-Sheet
  results.matched.forEach(item => {
    setCellValue(sheet, resultCol, item.rowIndex, 'Matched');
  });
  
  results.notFound.forEach(item => {
    setCellValue(sheet, resultCol, item.rowIndex, 'Not in Evento');
  });
  
  invalidGrades.forEach(item => {
    setCellValue(sheet, resultCol, item.rowIndex, 'Invalid Grade');
  });
  
  // Erstelle zusätzliches Sheet für "Not in Excel"
  if (notInExcel.length > 0) {
    const notInExcelSheet = createSheet('Not in Excel', 
      [['Name', 'Bemerkung'], ...notInExcel.map(e => [e.name, 'In Evento, nicht in Excel'])]);
    workbook.SheetNames.push('Not in Excel');
    workbook.Sheets['Not in Excel'] = notInExcelSheet;
  }
  
  // Erstelle zusätzliches Sheet für "Invalid Grades"
  if (invalidGrades.length > 0) {
    const invalidSheet = createSheet('Invalid Grades',
      [['Name', 'Note', 'Zeile', 'Bemerkung'], ...invalidGrades.map(e => [e.name, e.grade, e.rowIndex, 'Ungültige Note'])])
    workbook.SheetNames.push('Invalid Grades');
    workbook.Sheets['Invalid Grades'] = invalidSheet;
  }
  
  // Download
  const blob = new Blob([XLSX.write(workbook, {type: 'array'})]);
  downloadBlob(blob, 'evento_result.xlsx');
}
```

**Resultat-Werte:**
- **Matched**: Name wurde gefunden und Note wurde gesetzt
- **Not in Evento**: Name aus Excel wurde in EventoWeb nicht gefunden
- **Invalid Grade**: Ungültige Note (z.B. 2.3) - nur halbe Noten erlaubt

---

## 5. Implementierungsplan für LLM

### 5.1 Dateistruktur

Eine einzelne JavaScript-Datei mit folgendem Aufbau:

```javascript
// ==UserScript==
// ... Header ...
// ==/UserScript==

(function() {
  'use strict';
  
  // ============ KONFIGURATION ============
  const CONFIG = { ... };
  const GRADE_MAP = { ... };
  
  // ============ STYLES ============
  const STYLES = `...`;
  
  // ============ UI MODULE ============
  const UI = {
    init() { ... },
    createPanel() { ... },
    updateStatus(msg) { ... },
    showResults(results) { ... },
    // ...
  };
  
  // ============ EXCEL MODULE ============
  const Excel = {
    parse(file) { ... },
    export(workbook, results) { ... },
  };
  
  // ============ EVENTO MODULE ============
  const Evento = {
    scanPage() { ... },
    setGrade(select, grade) { ... },
  };
  
  // ============ MATCHER MODULE ============
  const Matcher = {
    match(excelData, domMap) { ... },
  };
  
  // ============ MAIN ============
  function main() {
    UI.init();
  }
  
  // Start when DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', main);
  } else {
    main();
  }
})();
```

### 5.2 Schritt-für-Schritt Implementierung

**Phase 1: Userscript-Grundgerüst**
```
Aufgabe: Erstelle das Userscript mit Header
- @name: Evento Note Importer
- @match: *://eventoweb.zhaw.ch/cst_pages/brn_qualifikationdurchdozenten.aspx*
- @require: SheetJS CDN (xlsx-0.20.1)
- @grant: none
- Wrap Code in IIFE mit 'use strict'
```

**Phase 2: Konfiguration und Konstanten**
```
Aufgabe: Definiere CONFIG und GRADE_MAP
- CONFIG.SHEET_NAME = 'evento'
- CONFIG.COL_NAME = 'Kombi'
- CONFIG.COL_GRADE = 'Note'
- GRADE_MAP mit allen Noten (1.0-6.0)
```

**Phase 3: CSS Styles**
```
Aufgabe: Erstelle CSS für das UI
- Panel-Styles (position fixed, rechts unten)
- Header-Styles (ZHAW Blau #0064A3)
- Drop-Zone-Styles (gestrichelte Border, hover-Effekt)
- Button-Styles (primary und secondary)
- Liste-Styles für nicht gefundene Namen
- Highlight-Style für geänderte Selects
```

**Phase 4: UI Modul - Panel erstellen**
```
Aufgabe: Implementiere UI.createPanel()
- Erstelle Container-Div mit ID "evento-importer-panel"
- Erstelle Header mit Titel und Close-Button
- Erstelle Drop-Zone für Datei-Upload
- Erstelle Status-Bereich
- Erstelle Ergebnis-Bereich (initial versteckt)
- Erstelle Button-Bereich mit "Anwenden" und "Download"
- Füge alles in document.body ein
```

**Phase 5: UI Modul - Event Handler**
```
Aufgabe: Implementiere Event-Handler
- Drop-Zone: dragover, dragleave, drop, click
- File-Input: change
- Close-Button: click (Panel minimieren)
- Anwenden-Button: click (Noten setzen)
- Download-Button: click (Excel exportieren)
```

**Phase 6: Excel Modul - Parser**
```
Aufgabe: Implementiere Excel.parse(file)
- Nutze FileReader.readAsArrayBuffer
- Parse mit XLSX.read()
- Validiere Sheet "evento"
- Extrahiere Daten mit XLSX.utils.sheet_to_json()
- Validiere Spalten "Kombi" und "Note"
- **WICHTIG:** Prüfe Noten auf exakte halbe Noten (keine Rundung!)
  - Gültig: 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6
  - Ungültig: 2.3, 4.7, etc. → invalidGrades[]
- Gib Objekt zurück: {success, data, invalidGrades, workbook, error}
```

**Phase 7: Evento Modul - DOM Scanner**
```
Aufgabe: Implementiere Evento.scanPage()
- Finde alle Links mit href "Brn_PersonDetailMA"
- Für jeden Link:
  - Extrahiere Name aus textContent
  - Finde parent <tr>
  - Finde <select> in der Zeile
  - Speichere in Map
- Gib Map zurück
```

**Phase 8: Matcher Modul**
```
Aufgabe: Implementiere Matcher.match(excelData, domMap)
- Iteriere über excelData
- Prüfe exakten Match in domMap
- Erstelle matched[] und notFound[] Arrays
- **NEU:** Finde DOM-Einträge die nicht in Excel sind
  - Iteriere über domMap
  - Prüfe ob Name in excelData existiert
  - Falls nicht: notInExcel[] hinzufügen
- Gib {matched, notFound, notInExcel} zurück
```

**Phase 9: Evento Modul - Grade Setter**
```
Aufgabe: Implementiere Evento.setGrade(select, grade)
- Normalisiere Note (z.B. "5" → "5.0")
- Hole Value aus GRADE_MAP
- Setze select.value
- Dispatch 'change' Event
- Setze visuelles Feedback (grüner Hintergrund)
```

**Phase 10: Excel Modul - Exporter**
```
Aufgabe: Implementiere Excel.export(workbook, results, invalidGrades, notInExcel)
- Hole Sheet "evento"
- Finde/erstelle Spalte "Resultat"
- Setze Werte: "Matched", "Not in Evento", "Invalid Grade"
- **NEU:** Erstelle Sheet "Not in Excel" für Einträge nur in Evento
- **NEU:** Erstelle Sheet "Invalid Grades" für ungültige Noten
- Erstelle Blob mit XLSX.write()
- Triggere Download als "evento_result.xlsx"
```

**Phase 11: Main und Integration**
```
Aufgabe: Implementiere main() und verbinde alles
1. Prüfe ob auf korrekter Seite
2. Initialisiere UI
3. Verbinde Event-Handler:
   - Datei-Upload → Excel.parse → UI.showResults (inkl. invalidGrades)
   - Anwenden-Button → Evento.setGrade für alle Matches
   - Download-Button → Excel.export (mit notInExcel und invalidGrades)
```

**Phase 12: Fehlerbehandlung**
```
Aufgabe: Füge robuste Fehlerbehandlung hinzu
- Try-catch um alle async Operationen
- Benutzerfreundliche Fehlermeldungen im UI
- Console.error für Debugging
```

### 5.3 Code-Template

```javascript
// ==UserScript==
// @name         Evento Note Importer
// @namespace    https://zhaw.ch/
// @version      1.0
// @description  Importiert Noten aus Excel in EventoWeb
// @author       ZHAW
// @match        *://eventoweb.zhaw.ch/cst_pages/brn_qualifikationdurchdozenten.aspx*
// @require      https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js
// @grant        none
// ==/UserScript==

(function() {
  'use strict';
  
  // ============ KONFIGURATION ============
  const CONFIG = {
    SHEET_NAME: 'evento',
    COL_NAME: 'Kombi',
    COL_GRADE: 'Note'
  };
  
  const GRADE_MAP = {
    '1': '10', '1.0': '10', '1.5': '11',
    '2': '12', '2.0': '12', '2.5': '16',
    '3': '17', '3.0': '17', '3.5': '18',
    '4': '19', '4.0': '19', '4.5': '20',
    '5': '21', '5.0': '21', '5.5': '22',
    '6': '23', '6.0': '23'
  };
  
  // ============ STYLES ============
  function injectStyles() {
    // TODO: Implementieren
  }
  
  // ============ UI MODULE ============
  const UI = {
    panel: null,
    statusEl: null,
    resultEl: null,
    
    init() {
      // TODO: Implementieren
    },
    
    createPanel() {
      // TODO: Implementieren
    },
    
    updateStatus(message, type = 'info') {
      // TODO: Implementieren
    },
    
    showResults(results) {
      // TODO: Implementieren
    }
  };
  
  // ============ EXCEL MODULE ============
  const Excel = {
    workbook: null,
    
    async parse(file) {
      // TODO: Implementieren
    },
    
    export(results) {
      // TODO: Implementieren
    }
  };
  
  // ============ EVENTO MODULE ============
  const Evento = {
    domMap: null,
    
    scanPage() {
      // TODO: Implementieren
    },
    
    setGrade(selectElement, grade) {
      // TODO: Implementieren
    },
    
    applyGrades(matches) {
      // TODO: Implementieren
    }
  };
  
  // ============ MATCHER MODULE ============
  const Matcher = {
    match(excelData, domMap) {
      // TODO: Implementieren
    }
  };
  
  // ============ MAIN ============
  function main() {
    console.log('Evento Note Importer geladen');
    injectStyles();
    UI.init();
    Evento.scanPage();
  }
  
  // Start
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', main);
  } else {
    main();
  }
})();
```

---

## 6. Testing-Anleitung

### 6.1 Voraussetzungen

1. **Browser:** Google Chrome oder Firefox (aktuell)
2. **Extension:** Tampermonkey installiert
3. **Zugang:** EventoWeb ZHAW mit gültiger Anmeldung
4. **Testdaten:** Excel-Datei mit Tabellenblatt "evento"

### 6.2 Userscript Installation

1. Klicke auf Tampermonkey-Icon im Browser
2. Wähle "Neues Skript erstellen"
3. Lösche den Template-Code
4. Füge das Userscript ein
5. Speichere (Ctrl+S)
6. Schliesse den Editor

### 6.3 Test-Excel erstellen

Erstelle eine Datei `test_noten.xlsx` mit Tabellenblatt "evento":

| Kombi | Note |
|-------|------|
| [Exakter Name 1 aus EventoWeb] | 5.0 |
| [Exakter Name 2 aus EventoWeb] | 4.5 |
| [Exakter Name 3 aus EventoWeb] | 6.0 |
| Test Nicht Vorhanden | 3.0 |

### 6.4 Schritt-für-Schritt Tests

#### Test 1: Script-Aktivierung

**Schritte:**
1. Öffne EventoWeb
2. Navigiere zur Qualifikationsseite
3. Prüfe ob das Import-Panel erscheint

**Erwartetes Ergebnis:**
- Panel/Button erscheint rechts unten
- Tampermonkey zeigt "1" (aktives Script)

#### Test 2: Panel-Interaktion

**Schritte:**
1. Klicke auf den Toggle-Button (falls minimiert)
2. Panel sollte sich öffnen
3. Klicke auf [X]
4. Panel sollte sich schliessen

**Erwartetes Ergebnis:**
- Panel öffnet und schliesst korrekt
- Animation flüssig

#### Test 3: Datei-Upload per Klick

**Schritte:**
1. Öffne das Panel
2. Klicke auf die Drop-Zone
3. Wähle die Test-Excel-Datei

**Erwartetes Ergebnis:**
- File-Dialog öffnet sich
- Nach Auswahl: Status ändert sich
- Ergebnisse werden angezeigt

#### Test 4: Datei-Upload per Drag & Drop

**Schritte:**
1. Öffne das Panel
2. Ziehe die Excel-Datei auf die Drop-Zone

**Erwartetes Ergebnis:**
- Drop-Zone zeigt Hover-Effekt
- Nach Drop: Datei wird verarbeitet

#### Test 5: Ergebnis-Anzeige

**Erwartetes Ergebnis:**
- Anzahl Matches wird angezeigt
- Anzahl "In Excel, nicht in Evento" wird angezeigt
- Anzahl "In Evento, nicht in Excel" wird angezeigt
- Anzahl "Ungültige Noten" wird angezeigt
- Liste der nicht gefundenen Namen erscheint
- Liste der Namen nur in Evento erscheint
- Liste der ungültigen Noten erscheint
- "Test Nicht Vorhanden" ist in der Liste "In Excel, nicht in Evento"

#### Test 6: Noten anwenden

**Schritte:**
1. Klicke auf "Anwenden"
2. Scrolle durch die Seite

**Erwartetes Ergebnis:**
- Noten sind in den Dropdowns gesetzt
- Geänderte Felder sind visuell markiert (grün)
- Status zeigt "X Noten gesetzt"

#### Test 7: Noten-Korrektheit prüfen

**Schritte:**
1. Prüfe für jeden gematchten Namen das Dropdown
2. Vergleiche mit Excel

**Erwartetes Ergebnis:**
- Note 5.0 → Dropdown zeigt "5.00"
- Note 4.5 → Dropdown zeigt "4.50"
- etc.

#### Test 8: Excel-Download

**Schritte:**
1. Klicke auf "Excel herunterladen"
2. Öffne die heruntergeladene Datei

**Erwartetes Ergebnis:**
- Download startet
- Datei heisst "evento_result.xlsx"
- Sheet "evento" enthält Spalte "Resultat"
- Werte sind "Matched", "Not in Evento", oder "Invalid Grade"
- **NEU:** Sheet "Not in Excel" vorhanden (falls Namen nur in Evento)
- **NEU:** Sheet "Invalid Grades" vorhanden (falls ungültige Noten)

### 6.5 Fehlerszenarien testen

#### Test E1: Falsche Dateiendung

**Schritte:**
- Versuche .csv oder .xls hochzuladen

**Erwartetes Ergebnis:**
- Fehlermeldung im UI
- Keine Verarbeitung

#### Test E2: Falsches Sheet

**Schritte:**
- Excel ohne Sheet "evento" hochladen

**Erwartetes Ergebnis:**
- Fehlermeldung: "Sheet 'evento' nicht gefunden"

#### Test E3: Fehlende Spalten

**Schritte:**
- Excel ohne "Kombi" oder "Note" Spalte

**Erwartetes Ergebnis:**
- Fehlermeldung mit Hinweis auf fehlende Spalte

#### Test E4: Script auf falscher Seite

**Schritte:**
- Öffne EventoWeb Startseite (nicht Qualifikation)

**Erwartetes Ergebnis:**
- Kein Panel erscheint
- Script ist inaktiv

#### Test E5: Ungültige Note

**Schritte:**
- Excel mit Note "2.3" oder "4.7" verwenden

**Erwartetes Ergebnis:**
- Warnung im UI: "Ungültige Noten: 1"
- Note wird **nicht** gerundet
- Eintrag erscheint in "Invalid Grades"-Liste
- In Excel-Export: Status "Invalid Grade"
- Separates Sheet "Invalid Grades" wird erstellt

### 6.6 Checkliste

| # | Test | Status |
|---|------|--------|
| 1 | Script aktiviert sich auf Qualifikationsseite | ☐ |
| 2 | Script aktiviert sich NICHT auf anderen Seiten | ☐ |
| 3 | Panel erscheint korrekt | ☐ |
| 4 | Panel öffnet/schliesst | ☐ |
| 5 | Datei-Upload per Klick funktioniert | ☐ |
| 6 | Datei-Upload per Drag&Drop funktioniert | ☐ |
| 7 | Excel wird korrekt geparst | ☐ |
| 8 | Matches werden korrekt erkannt | ☐ |
| 9 | Not Found werden korrekt erkannt | ☐ |
| 10 | Ergebnis-Anzeige ist korrekt | ☐ |
| 11 | Noten werden korrekt gesetzt | ☐ |
| 12 | Visuelles Feedback bei geänderten Selects | ☐ |
| 13 | Excel-Export funktioniert | ☐ |
| 14 | Resultat-Spalte ist korrekt | ☐ |
| 15 | Sheet "Not in Excel" wird erstellt | ☐ |
| 16 | Sheet "Invalid Grades" wird erstellt | ☐ |
| 17 | Ungültige Noten werden nicht gerundet | ☐ |
| 18 | Fehlerbehandlung funktioniert | ☐ |
| 19 | UI ist responsive und bedienbar | ☐ |

---

## 7. Unterschiede zur DevTools-Variante

| Aspekt | DevTools Snippet | Tampermonkey |
|--------|------------------|--------------|
| Installation | Keine | Einmalig |
| Aktivierung | Manuell (Copy-Paste) | Automatisch |
| UI | Minimal (Konsole) | Grafisches Panel |
| Benutzerfreundlichkeit | Niedrig | Hoch |
| Persistenz | Keine | Ja |
| SheetJS laden | Dynamisch im Code | Via @require |
| Debugging | Direkt in Konsole | Via Tampermonkey-Editor |

---

## 8. Bekannte Einschränkungen

1. **Browser-Abhängigkeit:** Nur mit Tampermonkey/Greasemonkey
2. **URL-Änderungen:** Bei URL-Änderungen muss @match angepasst werden
3. **XLSX-Grösse:** Sehr grosse Dateien können langsam sein
4. **Keine Offline-Funktion:** SheetJS wird von CDN geladen

---

## 9. Erweiterungsmöglichkeiten

1. **Undo-Funktion:** Ursprüngliche Werte speichern und wiederherstellen
2. **Preview-Modus:** Änderungen anzeigen vor dem Anwenden
3. **Mehrere Dateien:** Batch-Import unterstützen
4. **Einstellungen:** Konfigurierbare Spalten-Namen
5. **Dark Mode:** UI-Theme anpassen
