# Product Requirements Document (PRD)
## Excel-to-Evento Note Import Tool

**Version:** 1.0  
**Datum:** 30. Januar 2026  
**Status:** Entwurf

---

## 1. Zusammenfassung

Entwicklung eines Tools, das Noten aus einer Excel-Datei automatisiert in die EventoWeb-Qualifikationsseite der ZHAW überträgt.

---

## 2. Ausgangslage

### 2.1 Webapplikation (EventoWeb)
- **System:** EventoWeb der ZHAW (Zürcher Hochschule für Angewandte Wissenschaften)
- **Seite:** Qualifikation/Bewertung von Studierenden (`brn_qualifikationdurchdozenten.aspx`)
- **Authentifizierung:** Shibboleth-basierte Anmeldung erforderlich
- **Struktur:**
  - HTML-Tabelle mit Studierenden-Einträgen
  - Jede Zeile enthält:
    - Name im Format "NACHNAME VORNAME" (als Link-Text)
    - Adresse und Ort
    - `<select>`-Dropdown für die Note
  - Select-Element-Struktur:
    ```html
    <select name="10730947" id="10730947">
      <option selected="selected" value="0"> </option>
      <option value="23">6.00</option>
      <option value="22">5.50</option>
      <option value="21">5.00</option>
      <option value="20">4.50</option>
      <option value="19">4.00</option>
      <option value="18">3.50</option>
      <option value="17">3.00</option>
      <option value="16">2.50</option>
      <option value="12">2.00</option>
      <option value="11">1.50</option>
      <option value="10">1.00</option>
    </select>
    ```
- **Notenskala:** 1.0 bis 6.0 in 0.5-Schritten (Schweizer Notensystem)

### 2.2 Excel-Datei
- **Tabellenblatt:** "evento"
- **Format:** Semikolon-getrennt (CSV-ähnlich innerhalb von Excel)
- **Struktur:**
  ```
  Kombi;Note
  NACHNAME VORNAME;5.0
  NACHNAME VORNAME;3.5
  NACHNAME VORNAME;6.0
  ...
  ```
- **Hinweise:**
  - Die Anzahl der Einträge variiert
  - Die Namen sollten mit denen in EventoWeb übereinstimmen (mit möglichen Ausnahmen)
  - Noten sind Dezimalzahlen (1.0, 1.5, 2.0, ..., 6.0)

---

## 3. Anforderungen

### 3.1 Funktionale Anforderungen

| ID | Priorität | Anforderung |
|----|-----------|-------------|
| F01 | MUSS | Das Tool muss Excel-Dateien (.xlsx) mit dem Tabellenblatt "evento" einlesen können |
| F02 | MUSS | Das Tool muss Namen aus der Spalte "Kombi" extrahieren (Format: "NACHNAME VORNAME") |
| F03 | MUSS | Das Tool muss Noten aus der Spalte "Note" extrahieren (Format: Dezimalzahl) |
| F04 | MUSS | Das Tool muss die Namen in der EventoWeb-Seite mit den Excel-Namen abgleichen |
| F05 | MUSS | Das Tool muss bei einem Match die entsprechende Note im Select-Dropdown auswählen |
| F06 | MUSS | Das Tool muss im Excel eine neue Spalte "Resultat" erstellen/aktualisieren |
| F07 | MUSS | Bei erfolgreichem Match: "Matched" in Spalte "Resultat" eintragen |
| F08 | MUSS | Bei nicht gefundenem Namen: "Not Found" in Spalte "Resultat" eintragen |
| F09 | MUSS | Bei nicht exakt übereinstimmenden Namen: Manuelle Zuordnung durch Benutzer ermöglichen |
| F11 | KANN | Das Tool könnte eine Vorschau der Matches anzeigen, bevor Änderungen vorgenommen werden |
| F12 | KANN | Das Tool könnte eine Undo-Funktion bereitstellen |

### 3.2 Nicht-funktionale Anforderungen

| ID | Priorität | Anforderung |
|----|-----------|-------------|
| NF01 | MUSS | Das Tool muss im Browser Chrome funktionieren |
| NF02 | MUSS | Das Tool darf keine serverseitigen Änderungen erfordern |
| NF03 | MUSS | Das Tool muss clientseitig (im Browser) laufen |
| NF04 | SOLL | Das Tool soll einfach zu installieren und zu bedienen sein |
| NF05 | SOLL | Das Tool soll keine externen Abhängigkeiten zu Cloud-Diensten haben |
| NF06 | SOLL | Die Excel-Datei soll lokal verbleiben (Datenschutz) |

---

## 4. Anwendungsfall (Use Case)

### 4.1 Hauptszenario: Notenimport

**Akteur:** Dozent/in

**Vorbedingungen:**
1. Dozent/in ist in EventoWeb angemeldet
2. Qualifikationsseite für den entsprechenden Anlass ist geöffnet
3. Excel-Datei mit Noten liegt vor

**Ablauf:**
1. Dozent/in aktiviert das Tool
2. Dozent/in wählt die Excel-Datei aus
3. Tool liest Excel-Datei ein und extrahiert Namen und Noten
4. Tool durchsucht die aktuelle EventoWeb-Seite nach übereinstimmenden Namen
5. Für jeden Match: Tool setzt die Note im entsprechenden Select-Dropdown
6. Tool aktualisiert die Excel-Datei mit der Spalte "Resultat"
7. Dozent/in prüft die Ergebnisse
8. Dozent/in klickt auf "Speichern" in EventoWeb

**Nachbedingungen:**
- Alle gefundenen Studierenden haben ihre Note im Dropdown gesetzt
- Excel-Datei enthält für jeden Eintrag "Matched" oder "Not Found"

### 4.2 Alternativszenario: Name nicht gefunden

**Bedingung:** Name aus Excel existiert nicht in EventoWeb

**Ablauf:**
1. Tool markiert Eintrag in Excel mit "Not Found"
2. Tool zeigt optional eine Warnung/Liste der nicht gefundenen Namen
3. Dozent/in kann diese manuell bearbeiten

---

## 5. Mapping Noten zu Select-Values

| Note | Option Value |
|------|--------------|
| 1.0  | 10 |
| 1.5  | 11 |
| 2.0  | 12 |
| 2.5  | 16 |
| 3.0  | 17 |
| 3.5  | 18 |
| 4.0  | 19 |
| 4.5  | 20 |
| 5.0  | 21 |
| 5.5  | 22 |
| 6.0  | 23 |

**Hinweis:** Die Values sind nicht linear (Sprung von 12 zu 16 zwischen 2.0 und 2.5).

---

## 6. Technische Randbedingungen

- Die Webseite verwendet ASP.NET WebForms
- JavaScript ist auf der Seite aktiviert
- Die Seite enthält jQuery
- Die Select-Elemente haben dynamische numerische IDs
- Namen sind innerhalb von `<a>`-Tags mit `href` zu Personendetails

---

## 7. Risiken und Einschränkungen

| Risiko | Beschreibung | Mitigierung |
|--------|--------------|-------------|
| R01 | Unterschiedliche Namensschreibweisen zwischen Excel und EventoWeb | Manuelle Zuordnung durch Benutzer (kein Fuzzy-Matching) |
| R02 | Seiten-Layout kann sich ändern (Updates) | Robuste Selektoren verwenden |

### 7.1 Explizit ausgeschlossene Risiken

| ID | Beschreibung | Begründung |
|----|--------------|------------|
| ~R03~ | Sonderzeichen-Kodierung | Exakter String-Vergleich, keine Normalisierung nötig |
| ~R04~ | Paginierung bei vielen Studierenden | Alle Studierenden sind immer auf einer Seite |
| ~R05~ | Session-Timeout | Kann ignoriert werden, da Prozess schnell genug |

---

## 8. Erfolgs-Kriterien

1. ≥95% der korrekten Matches werden automatisch gefunden
2. Keine falschen Noten-Zuordnungen
3. Klare Kennzeichnung aller nicht gefundenen Einträge
4. Zeitersparnis von mindestens 80% gegenüber manueller Eingabe

---

## 9. Out of Scope

- Automatisches Speichern in EventoWeb (bleibt manuelle Aktion)
- Anlegen neuer Studierender in EventoWeb
- Bearbeitung mehrerer Anlässe gleichzeitig
- Mobile Nutzung
