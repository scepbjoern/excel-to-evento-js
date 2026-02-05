# Solution Design Variants
## Excel-to-Evento Note Import Tool

**Version:** 1.0  
**Datum:** 30. Januar 2026  
**Status:** Entwurf

---

## 1. Übersicht der Lösungsvarianten

| Variante | Technologie | Komplexität | Installationsaufwand | Wartbarkeit |
|----------|-------------|-------------|----------------------|-------------|
| A | Chrome Extension | Mittel | Mittel | Gut |
| B | Userscript (Tampermonkey/Greasemonkey) | Niedrig | Niedrig | Gut |
| C | Bookmarklet | Niedrig | Sehr niedrig | Mittel |
| D | Standalone Web-App mit Browser-Automation | Hoch | Hoch | Mittel |
| E | Browser DevTools Snippet | Sehr niedrig | Keine | Gering |

---

## 2. Variante A: Chrome Extension

### 2.1 Beschreibung
Eine dedizierte Chrome Extension, die auf der EventoWeb-Qualifikationsseite ein UI-Panel einblendet, über das die Excel-Datei hochgeladen werden kann.

### 2.2 Architektur
```
┌─────────────────────────────────────────────────┐
│                Chrome Extension                  │
├─────────────────────────────────────────────────┤
│  manifest.json (Manifest V3)                    │
│  ├── popup.html/js (UI für Dateiauswahl)        │
│  ├── content.js (DOM-Manipulation)              │
│  ├── background.js (Service Worker)             │
│  └── libs/xlsx.min.js (SheetJS Library)         │
└─────────────────────────────────────────────────┘
```

### 2.3 Technische Umsetzung
- **Excel-Parsing:** SheetJS (xlsx) Library für clientseitiges Lesen von .xlsx-Dateien
- **DOM-Zugriff:** Content Script mit Zugriff auf die EventoWeb-Seite
- **Datei-Upload:** HTML5 File API im Popup oder Content Script
- **Excel-Export:** SheetJS zum Schreiben der aktualisierten Datei mit "Resultat"-Spalte

### 2.4 Vorteile
- ✅ Professionelle Lösung mit eigenem UI
- ✅ Persistente Installation
- ✅ Einfache Updates über Chrome Web Store oder lokales Laden
- ✅ Volle Kontrolle über Berechtigungen
- ✅ Kann auf spezifische URLs beschränkt werden

### 2.5 Nachteile
- ❌ Erfordert Extension-Entwicklung (manifest.json, etc.)
- ❌ Installation komplexer als Bookmarklet
- ❌ Bei Veröffentlichung im Chrome Web Store: Review-Prozess
- ❌ Manifest V3 hat Einschränkungen bei dynamischem Code

### 2.6 Aufwandsschätzung
- Entwicklung: **8-12 Stunden**
- Installation: **5 Minuten** (unpacked extension laden)

---

## 3. Variante B: Userscript (Tampermonkey/Greasemonkey)

### 3.1 Beschreibung
Ein Userscript, das über Tampermonkey (Chrome) oder Greasemonkey (Firefox) installiert wird und automatisch auf EventoWeb-Qualifikationsseiten aktiv wird.

### 3.2 Architektur
```
┌─────────────────────────────────────────────────┐
│            Tampermonkey Userscript              │
├─────────────────────────────────────────────────┤
│  // ==UserScript==                              │
│  // @name         Evento Note Importer          │
│  // @match        *://eventoweb.zhaw.ch/*       │
│  // @require      https://cdn.sheetjs.com/...   │
│  // ==/UserScript==                             │
│                                                 │
│  - Injiziert Upload-Button in Seite             │
│  - Liest Excel via SheetJS                      │
│  - Manipuliert Select-Elemente                  │
│  - Generiert Download für aktualisierte Excel   │
└─────────────────────────────────────────────────┘
```

### 3.3 Technische Umsetzung
- **Excel-Parsing:** SheetJS via @require CDN
- **DOM-Zugriff:** Direkter Zugriff (Userscript läuft im Seitenkontext)
- **UI:** Dynamisch eingefügter Button/Panel
- **Excel-Export:** Blob-Download für aktualisierte Datei

### 3.4 Vorteile
- ✅ Einfache Installation (Copy-Paste in Tampermonkey)
- ✅ Keine Extension-Entwicklung nötig
- ✅ Schnelle Iteration während Entwicklung
- ✅ Kann externe Libraries via @require laden
- ✅ Funktioniert in Chrome, Firefox, Edge

### 3.5 Nachteile
- ❌ Erfordert Tampermonkey/Greasemonkey als Voraussetzung
- ❌ Weniger "offiziell" als Chrome Extension
- ❌ @require von lokalen Dateien problematisch

### 3.6 Aufwandsschätzung
- Entwicklung: **4-6 Stunden**
- Installation: **10 Minuten** (Tampermonkey + Script)

---

## 4. Variante C: Bookmarklet

### 4.1 Beschreibung
Ein JavaScript-Snippet, das als Lesezeichen im Browser gespeichert wird und per Klick auf der EventoWeb-Seite ausgeführt wird.

### 4.2 Architektur
```
┌─────────────────────────────────────────────────┐
│                  Bookmarklet                    │
├─────────────────────────────────────────────────┤
│  javascript:(function(){                        │
│    // Lädt SheetJS dynamisch                    │
│    // Erstellt File-Input                       │
│    // Verarbeitet Excel                         │
│    // Manipuliert DOM                           │
│  })();                                          │
└─────────────────────────────────────────────────┘
```

### 4.3 Technische Umsetzung
- **Excel-Parsing:** Dynamisches Laden von SheetJS via Script-Injection
- **DOM-Zugriff:** Direkter Zugriff nach Klick auf Bookmarklet
- **UI:** Alert/Prompt oder dynamisch eingefügtes Modal
- **Einschränkung:** Code-Länge begrenzt, daher oft externe Datei nötig

### 4.4 Vorteile
- ✅ Keine Installation von Extensions nötig
- ✅ Extrem portabel (nur URL kopieren)
- ✅ Funktioniert in allen Browsern

### 4.5 Nachteile
- ❌ Längenbeschränkung für Bookmarklet-Code
- ❌ Muss externes Script laden (CSP kann blocken)
- ❌ Keine persistente Speicherung
- ❌ Weniger elegant für komplexe Funktionalität
- ❌ Content Security Policy (CSP) kann Ausführung verhindern

### 4.6 Aufwandsschätzung
- Entwicklung: **3-5 Stunden**
- Installation: **2 Minuten**

---

## 5. Variante D: Standalone Web-App mit Browser-Automation

### 5.1 Beschreibung
Eine separate Webanwendung oder Desktop-App, die Selenium/Playwright verwendet, um den Browser zu steuern.

### 5.2 Architektur
```
┌─────────────────────────────────────────────────┐
│           Standalone Application                │
├─────────────────────────────────────────────────┤
│  Node.js / Python Application                   │
│  ├── Excel Reader (xlsx / openpyxl)             │
│  ├── Browser Automation (Playwright/Selenium)  │
│  └── UI (Electron / Web Interface)             │
└─────────────────────────────────────────────────┘
```

### 5.3 Technische Umsetzung
- **Excel-Parsing:** Node.js xlsx oder Python openpyxl
- **Browser-Automation:** Playwright oder Selenium WebDriver
- **Authentifizierung:** Muss Shibboleth-Login handhaben
- **UI:** Electron für Desktop-App oder Web-Interface

### 5.4 Vorteile
- ✅ Volle Kontrolle über den Prozess
- ✅ Kann auch komplexere Workflows automatisieren
- ✅ Unabhängig von Browser-Extensions

### 5.5 Nachteile
- ❌ Hohe Komplexität
- ❌ Shibboleth-Authentifizierung komplex zu automatisieren
- ❌ Erfordert lokale Installation von Node.js/Python
- ❌ Browser-Automation kann durch Updates brechen
- ❌ Overhead für einfachen Use Case

### 5.6 Aufwandsschätzung
- Entwicklung: **20-40 Stunden**
- Installation: **30+ Minuten**

---

## 6. Variante E: Browser DevTools Snippet

### 6.1 Beschreibung
Ein JavaScript-Snippet, das manuell in der Browser-Konsole (DevTools) ausgeführt wird.

### 6.2 Architektur
```
┌─────────────────────────────────────────────────┐
│            DevTools Console Snippet             │
├─────────────────────────────────────────────────┤
│  // In Browser Console einfügen:                │
│  const script = document.createElement('script');│
│  script.src = 'https://cdn.sheetjs.com/...';    │
│  document.head.appendChild(script);             │
│  // ... weiterer Code ...                       │
└─────────────────────────────────────────────────┘
```

### 6.3 Technische Umsetzung
- Manuelles Einfügen in die Konsole
- Kann als "Snippet" in Chrome DevTools gespeichert werden
- File-Input wird dynamisch erstellt

### 6.4 Vorteile
- ✅ Keine Installation nötig
- ✅ Volle Flexibilität
- ✅ Gut für Debugging und Entwicklung

### 6.5 Nachteile
- ❌ Manueller Prozess bei jeder Nutzung
- ❌ Technisches Know-how erforderlich
- ❌ Nicht benutzerfreundlich für Endanwender

### 6.6 Aufwandsschätzung
- Entwicklung: **2-4 Stunden**
- Installation: **Keine** (aber manuelle Ausführung)

---

## 7. Vergleichsmatrix

| Kriterium | A: Extension | B: Userscript | C: Bookmarklet | D: Standalone | E: DevTools |
|-----------|:------------:|:-------------:|:--------------:|:-------------:|:-----------:|
| Benutzerfreundlichkeit | ⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ | ⭐⭐⭐⭐ | ⭐ |
| Installationsaufwand | ⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐ | ⭐⭐⭐⭐⭐ |
| Entwicklungsaufwand | ⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐ | ⭐⭐⭐⭐ |
| Wartbarkeit | ⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐ | ⭐⭐ | ⭐⭐ |
| Portabilität | ⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐ | ⭐⭐⭐⭐⭐ |
| Robustheit | ⭐⭐⭐⭐ | ⭐⭐⭐ | ⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐ |
| Datenschutz | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ | ⭐⭐⭐⭐⭐ |

*Legende: ⭐ = niedrig/schlecht, ⭐⭐⭐⭐⭐ = hoch/gut*

---

## 8. Empfehlung

### 8.1 Für Einzelnutzer / Schnelle Lösung
**➡️ Variante B: Userscript (Tampermonkey)**

**Begründung:**
- Gutes Verhältnis zwischen Aufwand und Benutzerfreundlichkeit
- Schnelle Entwicklung und einfache Updates
- Tampermonkey ist weit verbreitet und vertrauenswürdig
- Keine eigene Extension-Entwicklung nötig

### 8.2 Für Team / Professioneller Einsatz
**➡️ Variante A: Chrome Extension**

**Begründung:**
- Professionellere Lösung mit eigenem UI
- Kann intern verteilt werden
- Bessere Kontrolle über Berechtigungen
- Langfristig wartbarer

### 8.3 Für Prototyping / Entwicklung
**➡️ Variante E: DevTools Snippet**

**Begründung:**
- Sofort einsatzbereit
- Ideal zum Testen der DOM-Manipulation
- Kann später in Userscript/Extension überführt werden

---

## 9. Technische Kernkomponenten (für alle Varianten)

### 9.1 Excel-Parsing mit SheetJS
```javascript
// Beispiel: Excel einlesen
const workbook = XLSX.read(data, { type: 'array' });
const sheet = workbook.Sheets['evento'];
const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
```

### 9.2 Name-Matching
```javascript
// Normalisierung für Vergleich
function normalizeName(name) {
  return name.trim().toLowerCase()
    .replace(/\s+/g, ' ')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, ''); // Umlaute
}
```

### 9.3 Noten-Mapping
```javascript
const gradeToValue = {
  '1.0': '10', '1.5': '11', '2.0': '12', '2.5': '16',
  '3.0': '17', '3.5': '18', '4.0': '19', '4.5': '20',
  '5.0': '21', '5.5': '22', '6.0': '23'
};
```

### 9.4 DOM-Selektion der Namen und Select-Elemente
```javascript
// Namen finden (in <a>-Tags)
const rows = document.querySelectorAll('tr');
rows.forEach(row => {
  const link = row.querySelector('a[href*="Brn_PersonDetailMA"]');
  const select = row.querySelector('select');
  if (link && select) {
    const name = link.textContent.trim();
    // ...
  }
});
```

---

## 10. Nächste Schritte

1. **Entscheidung** für eine Variante basierend auf Anforderungen
2. **Prototyp** als DevTools Snippet erstellen
3. **Iterative Entwicklung** mit Tests auf echter EventoWeb-Seite
4. **Dokumentation** für Endanwender erstellen
5. **Deployment** je nach gewählter Variante

---

## 11. Anhang: Relevante Libraries

| Library | Zweck | URL |
|---------|-------|-----|
| SheetJS (xlsx) | Excel lesen/schreiben | https://sheetjs.com/ |
| FileSaver.js | Datei-Download im Browser | https://github.com/eligrey/FileSaver.js |
| Fuse.js | Fuzzy-Matching für Namen | https://fusejs.io/ |
