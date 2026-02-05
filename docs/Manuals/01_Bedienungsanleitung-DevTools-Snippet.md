# Bedienungsanleitung: Evento Note Import (DevTools Snippet)

**Version:** 1.0  
**Datum:** 5. Februar 2026

---

## Inhaltsverzeichnis

1. [Übersicht](#1-übersicht)
2. [Voraussetzungen](#2-voraussetzungen)
3. [Excel-Datei vorbereiten](#3-excel-datei-vorbereiten)
4. [Schritt-für-Schritt Anleitung](#4-schritt-für-schritt-anleitung)
5. [Ergebnisse verstehen](#5-ergebnisse-verstehen)
6. [Fehlerbehebung](#6-fehlerbehebung)
7. [FAQ](#7-faq)

---

## 1. Übersicht

Das **Evento Note Import Snippet** ermöglicht das automatische Übertragen von Noten aus einer Excel-Datei in die EventoWeb-Qualifikationsseite. Das Tool:

- Liest eine Excel-Datei mit Studierenden-Namen und Noten
- Gleicht die Namen mit der EventoWeb-Seite ab
- Setzt die Noten automatisch in die entsprechenden Dropdown-Felder
- Erstellt eine Ergebnis-Datei mit Matching-Status

**Zeitersparnis:** Anstatt jede Note manuell einzutragen, werden alle Noten in wenigen Sekunden übertragen.

---

## 2. Voraussetzungen

### Browser
- **Google Chrome** (empfohlen) oder Microsoft Edge
- Andere Browser werden nicht offiziell unterstützt

### Zugang
- Gültiger Login für EventoWeb ZHAW
- Zugriff auf die Qualifikationsseite eines Moduls

### Excel-Datei
- Format: `.xlsx` (Excel 2007 oder neuer)
- Muss ein Tabellenblatt namens **"evento"** enthalten
- Erforderliche Spalten: **"Kombi"** und **"Note"**

---

## 3. Excel-Datei vorbereiten

### 3.1 Tabellenblatt erstellen

1. Öffnen Sie Ihre Excel-Datei
2. Benennen Sie ein Tabellenblatt in **"evento"** um (exakte Schreibweise beachten!)
3. Erstellen Sie folgende Spaltenstruktur:

| Kombi | Note |
|-------|------|
| Müller Hans | 5.0 |
| Meier Anna | 4.5 |
| Schmidt Peter | 3.5 |

### 3.2 Wichtige Hinweise zur Spalte "Kombi"

- Die Namen müssen **exakt** mit den Namen in EventoWeb übereinstimmen
- Achten Sie auf **Gross-/Kleinschreibung**
- Kopieren Sie die Namen am besten direkt aus EventoWeb
- Format: `Nachname Vorname` (wie in EventoWeb angezeigt)

### 3.3 Gültige Noten

Folgende Notenwerte werden akzeptiert:

| Eingabe | Wird zu |
|---------|---------|
| 1, 1.0, 1,0 | 1.0 |
| 1.5, 1,5 | 1.5 |
| 2, 2.0, 2,0 | 2.0 |
| 2.5, 2,5 | 2.5 |
| 3, 3.0, 3,0 | 3.0 |
| 3.5, 3,5 | 3.5 |
| 4, 4.0, 4,0 | 4.0 |
| 4.5, 4,5 | 4.5 |
| 5, 5.0, 5,0 | 5.0 |
| 5.5, 5,5 | 5.5 |
| 6, 6.0, 6,0 | 6.0 |

**Wichtig:** Andere Werte (z.B. 2.3, 4.7) werden **nicht akzeptiert** und nicht gerundet! Diese Einträge erscheinen in der Ergebnisdatei unter "Invalid Grades" und müssen manuell korrigiert werden.

## 4. Schritt-für-Schritt Anleitung

### Schritt 1: EventoWeb öffnen

1. Öffnen Sie **Google Chrome**
2. Navigieren Sie zu EventoWeb: `https://eventoweb.zhaw.ch`
3. Melden Sie sich an
4. Öffnen Sie die **Qualifikationsseite** des gewünschten Moduls

### Schritt 2: DevTools öffnen

1. Drücken Sie **F12** oder **Ctrl+Shift+I** (Windows) / **Cmd+Option+I** (Mac)
2. Die Chrome DevTools öffnen sich am rechten Bildschirmrand oder unten
3. Klicken Sie auf den Tab **"Console"**

### Schritt 3: Snippet einfügen

1. Öffnen Sie die Datei `devtools-snippet/evento-note-import.js`
2. Kopieren Sie den **gesamten Inhalt** (Ctrl+A, dann Ctrl+C)
3. Klicken Sie in die Console-Zeile in den DevTools
4. Fügen Sie den Code ein (Ctrl+V)
5. Eventuell kommt eine Warnmeldung, dann zunächst von Hand "allow pasting" eingeben und mit Enter bestätigen
5. Drücken Sie **Enter**

### Schritt 4: Excel-Datei auswählen

1. Nach dem Ausführen öffnet sich automatisch ein **Datei-Dialog**
2. Navigieren Sie zu Ihrer Excel-Datei
3. Wählen Sie die Datei aus und klicken Sie **"Öffnen"**

### Schritt 5: Ergebnisse prüfen

1. In der Console erscheint eine **Zusammenfassung**:
   ```
   ═══════════════════════════════════════
     EVENTO NOTE IMPORT - ERGEBNIS
   ═══════════════════════════════════════
   ✓ Erfolgreich zugeordnet: 25
   ✓ Noten gesetzt: 25
   ✗ In Excel aber nicht in Evento: 2
   ✗ In Evento aber nicht in Excel: 1
   ⚠ Ungültige Noten: 1
   
   --- In Excel aber nicht in Evento gefunden:
     - Müller Hans (Note: 5.0, Zeile: 5)
   
   --- In Evento aber nicht in Excel:
     - Weber Lisa
   
   --- Ungültige Noten (nicht gesetzt):
     - Schmidt Karl: "2.3" (Zeile: 8)
    Hinweis: Nur halbe Noten erlaubt (1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6)
   ═══════════════════════════════════════
   ```

2. Die **Dropdown-Felder** mit gesetzten Noten werden **grün** markiert
3. Eine Datei **"evento_result.xlsx"** wird automatisch heruntergeladen

### Schritt 6: Änderungen speichern

**Wichtig:** Die Noten sind noch **nicht** in EventoWeb gespeichert!

1. Scrollen Sie zum Ende der Seite
2. Klicken Sie auf **"Speichern"** oder **"Übernehmen"**
3. Warten Sie auf die Bestätigung

---

## 5. Ergebnisse verstehen

### Console-Ausgabe

| Symbol | Bedeutung |
|--------|-----------|
| ✓ Erfolgreich zugeordnet | Anzahl Namen, die in EventoWeb gefunden und gematcht wurden |
| ✓ Noten gesetzt | Anzahl Noten, die erfolgreich in Evento gesetzt wurden |
| ✗ In Excel aber nicht in Evento | Namen aus Excel, die nicht in EventoWeb gefunden wurden |
| ✗ In Evento aber nicht in Excel | Namen aus EventoWeb, die nicht in der Excel-Datei vorhanden sind |
| ⚠ Ungültige Noten | Einträge mit ungültigen Noten (z.B. 2.3) - werden nicht gesetzt |

### Ergebnis-Excel (evento_result.xlsx)

Die heruntergeladene Datei enthält mehrere Tabellenblätter:

#### Tabellenblatt "evento" (Original mit Resultat-Spalte)

Eine neue Spalte **"Resultat"** wird hinzugefügt:

| Kombi | Note | Resultat |
|-------|------|----------|
| Müller Hans | 5.0 | Matched |
| Meier Anna | 4.5 | Matched |
| Unbekannt Peter | 3.5 | Not in Evento |
| Schmidt Karl | 2.3 | Invalid Grade |

- **Matched**: Name wurde gefunden und Note wurde gesetzt
- **Not in Evento**: Name aus Excel wurde in EventoWeb nicht gefunden
- **Invalid Grade**: Ungültige Note (z.B. 2.3) - nur halbe Noten erlaubt

#### Tabellenblatt "Not in Excel" (falls vorhanden)

Zeigt alle Studierenden, die in EventoWeb vorhanden sind, aber nicht in der Excel-Datei:

| Name | Bemerkung |
|------|-----------|
| Weber Lisa | In Evento vorhanden, aber nicht in Excel |

#### Tabellenblatt "Invalid Grades" (falls vorhanden)

Zeigt alle Einträge mit ungültigen Noten:

| Name | Note | Zeile | Bemerkung |
|------|------|-------|-----------|
| Schmidt Karl | 2.3 | 5 | Ungültige Note - nur halbe Noten erlaubt |

**Hinweis:** Gültige Noten sind nur: 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6

### Visuelle Markierung

Gesetzte Dropdown-Felder erhalten einen **hellgrünen Hintergrund** zur besseren Übersicht.

---

## 6. Fehlerbehebung

### Problem: "Sheet 'evento' nicht gefunden"

**Ursache:** Das Tabellenblatt heisst nicht exakt "evento"

**Lösung:**
1. Öffnen Sie die Excel-Datei
2. Rechtsklick auf das Tabellenblatt-Tab unten
3. Wählen Sie "Umbenennen"
4. Tippen Sie exakt: `evento` (Kleinbuchstaben!)

### Problem: "Spalte 'Kombi' nicht gefunden"

**Ursache:** Die Spaltenüberschrift stimmt nicht

**Lösung:**
1. Öffnen Sie die Excel-Datei
2. Prüfen Sie, ob die erste Zeile genau "Kombi" enthält
3. Achten Sie auf Leerzeichen vor/nach dem Text

### Problem: Viele Namen werden nicht gefunden

**Ursache:** Die Namen stimmen nicht exakt überein

**Lösung:**
1. Kopieren Sie die Namen direkt aus EventoWeb
2. Achten Sie auf:
   - Gross-/Kleinschreibung
   - Umlaute (ä, ö, ü)
   - Leerzeichen
   - Sonderzeichen

### Problem: "Keine Studierenden auf der Seite gefunden"

**Ursache:** Sie sind nicht auf der richtigen Seite

**Lösung:**
1. Navigieren Sie zur Qualifikationsseite in EventoWeb
2. Die Seite muss eine Tabelle mit Studierenden-Namen anzeigen
3. Die URL sollte "Brn_Qualifikation" oder ähnlich enthalten

### Problem: Note wird nicht akzeptiert

**Ursache:** Ungültiges Notenformat

**Lösung:**
- Verwenden Sie nur Werte von 1.0 bis 6.0
- Halbe Noten: 1.5, 2.5, 3.5, 4.5, 5.5
- Keine Buchstaben oder Sonderzeichen

---

## 7. FAQ

### Muss ich das Snippet jedes Mal neu einfügen?

Ja, das Snippet muss bei jedem Besuch der Seite neu eingefügt werden. 

**Tipp:** Sie können das Snippet als **Chrome Snippet** speichern:
1. DevTools öffnen (F12)
2. Gehe zu "Sources" > "Snippets"
3. Klicke auf "+ New snippet"
4. Füge den Code ein und speichere

### Kann ich das Tool auch für andere Seiten verwenden?

Nein, das Tool ist speziell für EventoWeb ZHAW entwickelt.

### Was passiert, wenn ich das Snippet zweimal ausführe?

Die Noten werden erneut gesetzt. Bereits gesetzte Noten werden überschrieben (mit demselben Wert).

### Werden meine Daten irgendwo gespeichert?

Nein, alle Daten werden nur lokal in Ihrem Browser verarbeitet. Es werden keine Daten an externe Server gesendet.

### Kann ich das Tool rückgängig machen?

Solange Sie in EventoWeb nicht auf "Speichern" klicken, können Sie die Seite einfach neu laden. Nach dem Speichern müssten Sie die Noten manuell korrigieren.

---

## Support

Bei Problemen oder Fragen sind Sie selbst verantwortlich. Das Tool wird nicht von jemandem unterstützt.

---

*Erstellt für das Excel-to-Evento Note Import Tool Projekt*
