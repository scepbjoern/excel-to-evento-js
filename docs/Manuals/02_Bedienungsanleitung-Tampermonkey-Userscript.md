# Bedienungsanleitung: Evento Note Importer (Tampermonkey Userscript)

**Version:** 1.0  
**Datum:** 5. Februar 2026

---

## 1. Übersicht

Der **Evento Note Importer** ist ein Tampermonkey Userscript, das den Import von Noten aus einer Excel-Datei in EventoWeb automatisiert. Das Script fügt ein benutzerfreundliches Panel direkt in die EventoWeb-Qualifikationsseite ein.

### Funktionen
- Excel-Datei per Drag & Drop oder Klick hochladen
- Automatischer Abgleich der Namen zwischen Excel und EventoWeb
- Noten werden automatisch in die Dropdowns eingetragen
- Export einer Ergebnis-Excel mit Statusangaben
- Übersichtliche Anzeige von Diskrepanzen

---

## 2. Voraussetzungen

### 2.1 Browser
- **Google Chrome** (empfohlen) oder **Mozilla Firefox**
- Aktuelle Version

### 2.2 Tampermonkey Extension
1. Öffnen Sie den Chrome Web Store oder Firefox Add-ons
2. Suchen Sie nach "Tampermonkey"
3. Installieren Sie die Extension
4. Das Tampermonkey-Icon erscheint in der Browser-Toolbar

### 2.3 Zugang zu EventoWeb
- Gültige ZHAW-Anmeldedaten
- Zugriff auf die Qualifikationsseite mit Noteneinträgen

### 2.4 Excel-Datei
Die Excel-Datei muss folgende Struktur haben:
- **Dateiformat:** `.xlsx`
- **Tabellenblatt:** muss `evento` heissen
- **Erforderliche Spalten:**
  - `Kombi` - Enthält die Namen der Studierenden (exakt wie in EventoWeb)
  - `Note` - Enthält die Noten

**Beispiel:**

| Kombi | Note |
|-------|------|
| Müller Hans | 5.0 |
| Meier Anna | 4.5 |
| Weber Lisa | 6.0 |

### 2.5 Gültige Noten
Es werden Viertelnoten akzeptiert (0.25-Schritte zwischen 1.0 und 6.0):
- z.B. 1.0, 1.25, 1.5, 1.75, 2.0, ..., 5.25, 5.5, 5.75, 6.0

Welche dieser Noten in EventoWeb tatsächlich gesetzt werden können, hängt vom jeweiligen Dropdown ab (Bachelor bietet i.d.R. nur halbe Noten, Master auch Viertelnoten). Das Tool sucht die passende Option automatisch im Dropdown; existiert sie dort nicht, bleibt der Eintrag unverändert und wird in der Ergebnis-Excel vermerkt.

**Wichtig:** Ungültige Noten wie 2.3 oder 4.7 (kein Viertel-Schritt) werden **nicht** gerundet, sondern als Fehler markiert.

---

## 3. Installation des Userscripts

### Schritt 1: Tampermonkey öffnen
1. Klicken Sie auf das Tampermonkey-Icon in der Browser-Toolbar
2. Wählen Sie **"Neues Skript erstellen..."**

### Schritt 2: Script einfügen
1. Löschen Sie den gesamten Template-Code im Editor
2. Öffnen Sie die Datei `evento-note-importer.user.js`
3. Kopieren Sie den gesamten Inhalt
4. Fügen Sie den Code in den Tampermonkey-Editor ein

### Schritt 3: Speichern
1. Drücken Sie `Ctrl+S` (oder `Cmd+S` auf Mac)
2. Oder klicken Sie auf **Datei → Speichern**
3. Schliessen Sie den Editor

### Schritt 4: Aktivierung prüfen
1. Klicken Sie auf das Tampermonkey-Icon
2. Das Script **"Evento Note Importer"** sollte in der Liste erscheinen
3. Stellen Sie sicher, dass es aktiviert ist (Häkchen)

---

## 4. Bedienung

### 4.1 Script starten

1. Öffnen Sie EventoWeb und melden Sie sich an
2. Navigieren Sie zur **Qualifikationsseite** (Notenerfassung)
3. Das Script aktiviert sich automatisch
4. Rechts unten erscheint der Button **"📋 Noten Import"**

### 4.2 Panel öffnen

Klicken Sie auf den Button **"📋 Noten Import"**, um das Hauptpanel zu öffnen.

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
│  Status: Bereit. Bitte Excel-Datei hochladen. │
│                                                │
└────────────────────────────────────────────────┘
```

### 4.3 Excel-Datei hochladen

**Option A: Drag & Drop**
1. Öffnen Sie den Datei-Explorer
2. Ziehen Sie die Excel-Datei auf den Upload-Bereich
3. Lassen Sie die Datei los

**Option B: Dateiauswahl**
1. Klicken Sie auf den Upload-Bereich
2. Wählen Sie die Excel-Datei im Dialog aus
3. Klicken Sie auf "Öffnen"

### 4.4 Ergebnisse prüfen

Nach dem Upload werden die Ergebnisse angezeigt:

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
│  [Details werden aufgelistet]                  │
│                                                │
│  ┌────────────┐  ┌────────────────────────┐  │
│  │  Anwenden  │  │  Excel herunterladen   │  │
│  └────────────┘  └────────────────────────┘  │
│                                                │
└────────────────────────────────────────────────┘
```

**Kategorien:**
- **✓ Gefunden:** Namen, die in Excel UND EventoWeb vorhanden sind
- **✗ In Excel, nicht in Evento:** Namen aus Excel, die nicht gefunden wurden
- **✗ In Evento, nicht in Excel:** Studierende in EventoWeb ohne Eintrag in Excel
- **⚠ Ungültige Noten:** Einträge mit ungültigen Notenwerten

### 4.5 Noten anwenden

1. Prüfen Sie die Ergebnisse im Panel
2. Klicken Sie auf **"Anwenden"**
3. Die Noten werden in die Dropdowns eingetragen
4. Geänderte Felder werden kurz grün markiert
5. Status zeigt: "✅ X Noten wurden gesetzt."

**Wichtig:** Die Noten sind noch nicht gespeichert! Sie müssen anschliessend den **"Speichern"-Button** von EventoWeb klicken.

### 4.6 Ergebnis-Excel herunterladen

1. Klicken Sie auf **"Excel herunterladen"**
2. Die Datei `evento_result.xlsx` wird heruntergeladen

**Inhalt der Ergebnis-Datei:**

| Tabellenblatt | Inhalt |
|---------------|--------|
| `evento` | Originaldaten mit zusätzlicher Spalte "Resultat" |
| `Not in Excel` | Studierende, die nur in EventoWeb vorhanden sind |
| `Invalid Grades` | Einträge mit ungültigen Noten |

**Resultat-Werte:**
- `Matched` - Name gefunden und Note gesetzt
- `Not in Evento` - Name nicht in EventoWeb gefunden
- `Invalid Grade` - Ungültige Note (nicht gesetzt)

### 4.7 Panel schliessen

Klicken Sie auf das **[X]** oben rechts im Panel, um es zu minimieren. Der Button "📋 Noten Import" erscheint wieder.

---

## 5. Workflow-Empfehlung

### Schritt-für-Schritt Anleitung

1. **Excel vorbereiten**
   - Erstellen Sie ein Tabellenblatt namens `evento`
   - Fügen Sie Spalten `Kombi` und `Note` hinzu
   - Tragen Sie die Namen exakt wie in EventoWeb ein
   - Verwenden Sie nur gültige Notenwerte

2. **EventoWeb öffnen**
   - Melden Sie sich an
   - Navigieren Sie zur Qualifikationsseite
   - Das Import-Panel erscheint automatisch

3. **Excel hochladen**
   - Öffnen Sie das Panel
   - Laden Sie die Excel-Datei hoch
   - Prüfen Sie die Ergebnisse

4. **Diskrepanzen klären**
   - Prüfen Sie "In Excel, nicht in Evento"
     - Tippfehler im Namen?
     - Student nicht eingeschrieben?
   - Prüfen Sie "In Evento, nicht in Excel"
     - Fehlender Eintrag in Excel?
   - Prüfen Sie "Ungültige Noten"
     - Korrigieren Sie auf halbe Noten

5. **Noten anwenden**
   - Klicken Sie "Anwenden"
   - Verifizieren Sie die eingetragenen Noten

6. **In EventoWeb speichern**
   - Klicken Sie auf den Speichern-Button von EventoWeb
   - Bestätigen Sie die Speicherung

7. **Dokumentation**
   - Laden Sie die Ergebnis-Excel herunter
   - Archivieren Sie diese für Ihre Unterlagen

---

## 6. Fehlerbehebung

### Problem: Script erscheint nicht

**Mögliche Ursachen:**
- Sie befinden sich nicht auf der Qualifikationsseite
- Tampermonkey ist deaktiviert
- Das Script ist nicht installiert oder deaktiviert

**Lösung:**
1. Prüfen Sie die URL: muss `brn_qualifikationdurchdozenten.aspx` enthalten
2. Klicken Sie auf das Tampermonkey-Icon
3. Prüfen Sie, ob "Evento Note Importer" aktiviert ist

### Problem: Datei wird nicht akzeptiert

**Fehlermeldung:** "Nur .xlsx Dateien werden unterstützt"

**Lösung:**
- Speichern Sie die Datei als Excel-Arbeitsmappe (.xlsx)
- Ältere Formate (.xls) werden nicht unterstützt

### Problem: Sheet 'evento' nicht gefunden

**Lösung:**
1. Öffnen Sie die Excel-Datei
2. Prüfen Sie den Namen des Tabellenblatts
3. Benennen Sie es exakt in `evento` um (Kleinschreibung!)

### Problem: Spalte nicht gefunden

**Lösung:**
- Prüfen Sie, dass die Spalten exakt `Kombi` und `Note` heissen
- Gross-/Kleinschreibung beachten
- Keine Leerzeichen vor/nach dem Namen

### Problem: Namen werden nicht gefunden

**Mögliche Ursachen:**
- Namen stimmen nicht exakt überein
- Unterschiedliche Schreibweisen
- Zusätzliche Leerzeichen

**Lösung:**
- Kopieren Sie die Namen direkt aus EventoWeb
- Prüfen Sie auf Sonderzeichen und Umlaute
- Das Script verwendet **exakten** Namensabgleich

### Problem: Ungültige Noten

**Fehlermeldung:** "Ungültige Noten: X"

**Lösung:**
- Nur halbe Noten sind erlaubt: 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0
- Korrigieren Sie Werte wie 2.3 oder 4.7
- Noten werden **nicht** automatisch gerundet

---

## 7. Sicherheitshinweise

1. **Prüfen Sie die Ergebnisse** vor dem Anwenden
2. **Speichern Sie erst**, nachdem Sie alle Noten verifiziert haben
3. **Archivieren Sie die Ergebnis-Excel** als Dokumentation
4. Das Script **speichert keine Daten** - alle Verarbeitung erfolgt lokal im Browser
5. **Keine Netzwerkanfragen** werden vom Script durchgeführt

---

## 8. Technische Informationen

| Eigenschaft | Wert |
|-------------|------|
| Script-Version | 1.0 |
| Unterstützte Browser | Chrome, Firefox |
| Benötigte Extension | Tampermonkey |
| URL-Pattern | `*://eventoweb.zhaw.ch/cst_pages/brn_qualifikationdurchdozenten.aspx*` |
| Externe Bibliothek | SheetJS (xlsx-0.20.1) |
| Maximale Dateigrösse | 10 MB |

---

## 9. Support

Bei Problemen oder Fragen wenden Sie sich an die IT-Abteilung der ZHAW.

---

*Diese Anleitung bezieht sich auf Version 1.0 des Evento Note Importer Userscripts.*
