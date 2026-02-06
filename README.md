# Excel-to-Evento Note Import Tool

Automatischer Import von Noten aus Excel in die ZHAW EventoWeb-Qualifikationsseite.

Sorry für die schlechte Tonqualität. Ich musste das Video mit dem eingebauten Mikrofon aufnehmen.

[![YouTube: Ausgangslage und Demo](https://img.youtube.com/vi/c4WZSbJMt-U/0.jpg)](https://www.youtube.com/watch?v=c4WZSbJMt-U)

## Überblick

Dieses Projekt bietet zwei Lösungen, um Noten aus einer Excel-Datei automatisiert in die EventoWeb-Qualifikationsseite der ZHAW zu übertragen:

- **DevTools Snippet** – Schnell und ohne Installation (für Entwickler)
- **Tampermonkey Userscript** – Komfortabel mit grafischer Oberfläche (für Dozenten)

## Funktionsweise

1. Excel-Datei mit Namen und Noten einlesen
2. Namen mit der EventoWeb-Seite abgleichen (exakter Match)
3. Noten automatisch in die Dropdowns eintragen
4. Ergebnisprotokoll in neuer Excel-Spalte "Resultat"

## Varianten

### Variante 1: DevTools Snippet

Schnelle Lösung für technisch versierte Benutzer.

```
📁 Verzeichnis: devtools-snippet/
📄 Datei: evento-note-import.js
```

**Voraussetzungen:** Chrome DevTools

**Nutzung:**
1. EventoWeb Qualifikationsseite öffnen
2. DevTools mit F12 öffnen → Tab "Console"
3. Snippet-Code einfügen und ausführen

📺 **Video-Anleitung:** 
[![DevTools Snippet Nutzung](https://img.youtube.com/vi/kHErDvA99og/0.jpg)](https://www.youtube.com/watch?v=kHErDvA99og)

[→ Detaillierte Anleitung](docs/Manuals/01_Bedienungsanleitung-DevTools-Snippet.md)

---

### Variante 2: Tampermonkey Userscript

Komfortable Lösung mit grafischem UI für regelmässige Nutzung.

```
📁 Verzeichnis: tampermonkey-userscript/
📄 Datei: evento-note-importer.user.js
```

**Voraussetzungen:**
- [Tampermonkey Extension](https://www.tampermonkey.net/)
- Chrome oder Firefox

**Installation:**
1. Tampermonkey im Browser installieren
2. Script in Tampermonkey → "Neues Skript erstellen" einfügen
3. Speichern (Ctrl+S)

**Nutzung:**
1. EventoWeb Qualifikationsseite öffnen
2. Panel "Evento Note Importer" erscheint automatisch rechts unten
3. Excel-Datei hochladen und Noten anwenden

📺 **Video-Anleitung:**
[![Tampermonkey Userscript Nutzung](https://img.youtube.com/vi/uUeG2u394oI/0.jpg)](https://www.youtube.com/watch?v=uUeG2u394oI)

[→ Detaillierte Anleitung](docs/Manuals/02_Bedienungsanleitung-Tampermonkey-Userscript.md)

## Projektstruktur

```
.
├── devtools-snippet/
│   └── evento-note-import.js       # DevTools Snippet Code
├── tampermonkey-userscript/
│   └── evento-note-importer.user.js # Tampermonkey Userscript
├── docs/
│   ├── Manuals/
│   │   ├── 01_Bedienungsanleitung-DevTools-Snippet.md
│   │   └── 02_Bedienungsanleitung-Tampermonkey-Userscript.md
│   ├── PRD/
│   │   └── PRD.md                   # Product Requirements
│   └── SDD/
│       ├── SDD-DevTools-Snippet.md  # Solution Design
│       └── SDD-Tampermonkey-Userscript.md
└── README.md
```

## Excel-Format

Das Excel-Datei benötigt ein Tabellenblatt mit dem Namen `evento`:

| Kombi            | Note |
|------------------|------|
| Muster Hans      | 5.0  |
| Beispiel Maria   | 4.5  |

**Wichtig:**
- Spalte A: `Kombi` (Name im Format "NACHNAME VORNAME")
- Spalte B: `Note` (1.0, 1.5, 2.0, ..., 6.0)
- Namen müssen exakt mit EventoWeb übereinstimmen

Nach dem Import enthält die Spalte `Resultat`:
- `Matched` – Name gefunden und Note gesetzt
- `Not Found` – Name nicht gefunden

## Noten-Mapping

| Note | Dropdown-Value |
|------|----------------|
| 1.0  | 10             |
| 1.5  | 11             |
| 2.0  | 12             |
| 2.5  | 16             |
| 3.0  | 17             |
| 3.5  | 18             |
| 4.0  | 19             |
| 4.5  | 20             |
| 5.0  | 21             |
| 5.5  | 22             |
| 6.0  | 23             |

## Systemanforderungen

- **Browser:** Google Chrome oder Firefox
- **Zugang:** ZHAW EventoWeb mit gültiger Anmeldung
- **Excel-Datei:** Format `.xlsx` mit Sheet "evento"

## Datenschutz

- Alle Daten verbleiben lokal im Browser
- Keine Daten werden an externe Server gesendet
- Keine Cloud-Anbindung

## Mitwirken und Making Of

Dieses Projekt wurde für Dozenten der ZHAW entwickelt. Wenn Du selbst daran weiter entwickeln möchtest oder mehr darüber erfahren möchtest, wie der Code entstanden ist, dann schau' Dir das folgende Making Of-Video an.

📺 **Making Of-Video:**
[![Making Of-Video](https://img.youtube.com/vi/r5GQHO5wXP8/0.jpg)](https://www.youtube.com/watch?v=r5GQHO5wXP8)

## Lizenz

[MIT License](LICENSE)

---

**ZHAW EventoWeb Note Import Tool** – Entwickelt für die Zürcher Hochschule für Angewandte Wissenschaften
