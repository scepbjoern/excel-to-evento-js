(async function() {
  'use strict';

  // ========== KONFIGURATION ==========
  const CONFIG = {
    SHEETJS_CDN: 'https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js',
    SHEET_NAME: 'evento',
    COLUMN_NAME: 'Kombi',
    COLUMN_GRADE: 'Note'
  };

  const GRADE_MAP = {
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

  // ========== HILFSFUNKTIONEN ==========

  /**
   * Lädt SheetJS dynamisch von CDN
   * @returns {Promise<boolean>}
   */
  async function loadSheetJS() {
    if (window.XLSX) {
      console.log('SheetJS bereits geladen.');
      return true;
    }

    console.log('SheetJS wird geladen...');
    
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = CONFIG.SHEETJS_CDN;
      script.onload = () => {
        console.log('SheetJS geladen!');
        resolve(true);
      };
      script.onerror = () => {
        reject(new Error('Fehler beim Laden von SheetJS'));
      };
      document.head.appendChild(script);
    });
  }

  /**
   * Erstellt einen File-Input und wartet auf Dateiauswahl
   * @returns {Promise<{file: File, workbook: Object}>}
   */
  function createFileInput() {
    return new Promise((resolve, reject) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.xlsx';
      input.style.display = 'none';
      
      input.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (!file) {
          reject(new Error('Keine Datei ausgewählt'));
          return;
        }
        
        try {
          const result = await parseExcel(file);
          resolve(result);
        } catch (error) {
          reject(error);
        } finally {
          document.body.removeChild(input);
        }
      });
      
      input.addEventListener('cancel', () => {
        document.body.removeChild(input);
        reject(new Error('Dateiauswahl abgebrochen'));
      });
      
      document.body.appendChild(input);
      input.click();
    });
  }

  /**
   * Parst eine Excel-Datei und extrahiert die Daten
   * @param {File} file
   * @returns {Promise<{data: Array, workbook: Object}>}
   */
  async function parseExcel(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const arrayBuffer = e.target.result;
          const workbook = XLSX.read(arrayBuffer, { type: 'array' });
          
          // Prüfe ob Sheet "evento" existiert
          if (!workbook.SheetNames.includes(CONFIG.SHEET_NAME)) {
            reject(new Error(`Sheet '${CONFIG.SHEET_NAME}' nicht gefunden. Verfügbare Sheets: ${workbook.SheetNames.join(', ')}`));
            return;
          }
          
          const sheet = workbook.Sheets[CONFIG.SHEET_NAME];
          const jsonData = XLSX.utils.sheet_to_json(sheet);
          
          // Prüfe ob erforderliche Spalten existieren
          if (jsonData.length === 0) {
            reject(new Error('Das Sheet enthält keine Daten'));
            return;
          }
          
          const firstRow = jsonData[0];
          if (!(CONFIG.COLUMN_NAME in firstRow)) {
            reject(new Error(`Spalte '${CONFIG.COLUMN_NAME}' nicht gefunden`));
            return;
          }
          if (!(CONFIG.COLUMN_GRADE in firstRow)) {
            reject(new Error(`Spalte '${CONFIG.COLUMN_GRADE}' nicht gefunden`));
            return;
          }
          
          // Extrahiere und validiere Daten
          const data = [];
          const invalidGrades = [];
          
          jsonData.forEach((row, index) => {
            const name = row[CONFIG.COLUMN_NAME];
            const grade = row[CONFIG.COLUMN_GRADE];
            
            if (!name || grade === undefined || grade === null) {
              console.warn(`Zeile ${index + 2}: Unvollständige Daten (Name: "${name}", Note: "${grade}")`);
              return;
            }
            
            // Normalisiere Note
            const normalizedGrade = normalizeGrade(grade);
            if (!normalizedGrade) {
              invalidGrades.push({
                name: String(name).trim(),
                grade: String(grade).trim(),
                rowIndex: index + 2,
                reason: 'Ungültige Note (nur halbe Noten erlaubt: 1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6)'
              });
              return;
            }
            
            data.push({
              name: String(name).trim(),
              grade: normalizedGrade,
              rowIndex: index + 2 // +2 weil Header in Zeile 1 und 0-indexed
            });
          });
          
          console.log(`Excel-Datei eingelesen: ${data.length} Einträge`);
          if (invalidGrades.length > 0) {
            console.warn(`${invalidGrades.length} Einträge mit ungültigen Noten übersprungen`);
          }
          resolve({ data, workbook, invalidGrades });
          
        } catch (error) {
          reject(new Error(`Fehler beim Parsen der Excel-Datei: ${error.message}`));
        }
      };
      
      reader.onerror = () => {
        reject(new Error('Fehler beim Lesen der Datei'));
      };
      
      reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Normalisiert eine Note in das Standard-Format (z.B. "5" → "5.0")
   * Keine Rundung! Nur exakte halbe Noten sind erlaubt.
   * @param {string|number} grade
   * @returns {string|null}
   */
  function normalizeGrade(grade) {
    // Konvertiere zu String
    let gradeStr = String(grade).trim();
    
    // Ersetze Komma durch Punkt
    gradeStr = gradeStr.replace(',', '.');
    
    // Prüfe ob es eine gültige Note ist (nur exakte halbe Noten)
    const validGrades = ['1', '1.0', '1.5', '2', '2.0', '2.5', '3', '3.0', '3.5', '4', '4.0', '4.5', '5', '5.0', '5.5', '6', '6.0'];
    
    if (validGrades.includes(gradeStr)) {
      // Normalisiere zu X.X Format
      if (gradeStr === '1' || gradeStr === '1.0') return '1.0';
      if (gradeStr === '1.5') return '1.5';
      if (gradeStr === '2' || gradeStr === '2.0') return '2.0';
      if (gradeStr === '2.5') return '2.5';
      if (gradeStr === '3' || gradeStr === '3.0') return '3.0';
      if (gradeStr === '3.5') return '3.5';
      if (gradeStr === '4' || gradeStr === '4.0') return '4.0';
      if (gradeStr === '4.5') return '4.5';
      if (gradeStr === '5' || gradeStr === '5.0') return '5.0';
      if (gradeStr === '5.5') return '5.5';
      if (gradeStr === '6' || gradeStr === '6.0') return '6.0';
    }
    
    // Keine Rundung - nur exakte halbe Noten sind gültig
    return null;
  }

  /**
   * Scannt die EventoWeb-Seite und extrahiert Namen und Select-Elemente
   * @returns {Map<string, HTMLSelectElement>}
   */
  function scanEventoPage() {
    const domMap = new Map();
    
    // Finde alle Links mit Personendetails
    const nameLinks = document.querySelectorAll('a[href*="Brn_PersonDetailMA"]');
    
    nameLinks.forEach(link => {
      const name = link.textContent.trim();
      
      // Finde die zugehörige Tabellenzeile
      const tr = link.closest('tr');
      if (!tr) {
        console.warn(`Keine Tabellenzeile für "${name}" gefunden`);
        return;
      }
      
      // Finde das Select-Element in dieser Zeile
      const select = tr.querySelector('select[name][id]');
      if (!select) {
        console.warn(`Kein Select-Element für "${name}" gefunden`);
        return;
      }
      
      domMap.set(name, select);
    });
    
    console.log(`DOM gescannt: ${domMap.size} Studierende gefunden`);
    return domMap;
  }

  /**
   * Führt das Matching zwischen Excel-Daten und DOM durch
   * @param {Array} excelData
   * @param {Map} domMap
   * @returns {{matches: Array, notFound: Array, notInExcel: Array}}
   */
  function matchNames(excelData, domMap) {
    const matches = [];
    const notFound = [];
    const excelNames = new Set(excelData.map(e => e.name));
    
    // Finde Excel-Einträge die nicht im DOM sind
    excelData.forEach(entry => {
      if (domMap.has(entry.name)) {
        matches.push({
          excelName: entry.name,
          grade: entry.grade,
          selectElement: domMap.get(entry.name),
          rowIndex: entry.rowIndex
        });
      } else {
        notFound.push({
          name: entry.name,
          grade: entry.grade,
          rowIndex: entry.rowIndex
        });
      }
    });
    
    // Finde DOM-Einträge die nicht im Excel sind
    const notInExcel = [];
    domMap.forEach((selectElement, name) => {
      if (!excelNames.has(name)) {
        notInExcel.push({
          name: name,
          selectElement: selectElement
        });
      }
    });
    
    return { matches, notFound, notInExcel };
  }

  /**
   * Setzt die Noten in den Select-Dropdowns
   * @param {Array} matches
   * @returns {number} Anzahl erfolgreich gesetzter Noten
   */
  function setGrades(matches) {
    let successCount = 0;
    
    matches.forEach(match => {
      const value = GRADE_MAP[match.grade];
      
      if (!value) {
        console.warn(`Kein Mapping für Note "${match.grade}" (${match.excelName})`);
        return;
      }
      
      // Prüfe ob der Wert als Option existiert
      const option = match.selectElement.querySelector(`option[value="${value}"]`);
      if (!option) {
        console.warn(`Option mit Wert "${value}" nicht gefunden für "${match.excelName}"`);
        return;
      }
      
      // Setze den Wert
      match.selectElement.value = value;
      
      // Visuelles Feedback
      match.selectElement.style.backgroundColor = '#90EE90';
      
      // Triggere change-Event für ASP.NET
      const event = new Event('change', { bubbles: true });
      match.selectElement.dispatchEvent(event);
      
      successCount++;
    });
    
    return successCount;
  }

  /**
   * Gibt die Ergebnisse in der Konsole aus
   * @param {Array} matches
   * @param {Array} notFound
   * @param {Array} notInExcel
   * @param {Array} invalidGrades
   * @param {number} setCount
   */
  function logResults(matches, notFound, notInExcel, invalidGrades, setCount) {
    console.log('\n═══════════════════════════════════════');
    console.log('  EVENTO NOTE IMPORT - ERGEBNIS');
    console.log('═══════════════════════════════════════');
    console.log(`✓ Erfolgreich zugeordnet: ${matches.length}`);
    console.log(`✓ Noten gesetzt: ${setCount}`);
    console.log(`✗ In Excel aber nicht in Evento: ${notFound.length}`);
    console.log(`✗ In Evento aber nicht in Excel: ${notInExcel.length}`);
    console.log(`⚠ Ungültige Noten: ${invalidGrades.length}`);
    
    if (notFound.length > 0) {
      console.log('\n--- In Excel aber nicht in Evento gefunden:');
      notFound.forEach(entry => {
        console.log(`  - ${entry.name} (Note: ${entry.grade}, Zeile: ${entry.rowIndex})`);
      });
    }
    
    if (notInExcel.length > 0) {
      console.log('\n--- In Evento aber nicht in Excel:');
      notInExcel.forEach(entry => {
        console.log(`  - ${entry.name}`);
      });
    }
    
    if (invalidGrades.length > 0) {
      console.log('\n--- Ungültige Noten (nicht gesetzt):');
      invalidGrades.forEach(entry => {
        console.log(`  - ${entry.name}: "${entry.grade}" (Zeile: ${entry.rowIndex})`);
      });
      console.log('   Hinweis: Nur halbe Noten erlaubt (1, 1.5, 2, 2.5, 3, 3.5, 4, 4.5, 5, 5.5, 6)');
    }
    
    console.log('═══════════════════════════════════════\n');
  }

  /**
   * Exportiert das Ergebnis als neue Excel-Datei
   * @param {Object} workbook
   * @param {Array} matches
   * @param {Array} notFound
   * @param {Array} notInExcel
   * @param {Array} invalidGrades
   */
  function exportResultExcel(workbook, matches, notFound, notInExcel, invalidGrades) {
    try {
      const sheet = workbook.Sheets[CONFIG.SHEET_NAME];
      
      // Finde die letzte Spalte
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const resultCol = range.e.c + 1;
      
      // Füge Header hinzu
      const headerCell = XLSX.utils.encode_cell({ r: 0, c: resultCol });
      sheet[headerCell] = { t: 's', v: 'Resultat' };
      
      // Erstelle Lookup für Ergebnisse
      const matchedRows = new Set(matches.map(m => m.rowIndex));
      const notFoundRows = new Set(notFound.map(n => n.rowIndex));
      const invalidGradeRows = new Set(invalidGrades.map(i => i.rowIndex));
      
      // Füge Ergebnisse hinzu
      for (let row = 1; row <= range.e.r; row++) {
        const excelRow = row + 1; // Excel ist 1-indexed, erste Datenzeile ist Zeile 2
        const cell = XLSX.utils.encode_cell({ r: row, c: resultCol });
        
        if (matchedRows.has(excelRow)) {
          sheet[cell] = { t: 's', v: 'Matched' };
        } else if (notFoundRows.has(excelRow)) {
          sheet[cell] = { t: 's', v: 'Not in Evento' };
        } else if (invalidGradeRows.has(excelRow)) {
          sheet[cell] = { t: 's', v: 'Invalid Grade' };
        }
      }
      
      // Update range
      range.e.c = resultCol;
      sheet['!ref'] = XLSX.utils.encode_range(range);
      
      // Erstelle zweites Sheet für "Not in Excel" Einträge (die nur in Evento sind)
      if (notInExcel.length > 0) {
        const notInExcelData = [
          ['Name', 'Bemerkung'],
          ...notInExcel.map(entry => [entry.name, 'In Evento vorhanden, aber nicht in Excel'])
        ];
        const notInExcelSheet = XLSX.utils.aoa_to_sheet(notInExcelData);
        workbook.SheetNames.push('Not in Excel');
        workbook.Sheets['Not in Excel'] = notInExcelSheet;
      }
      
      // Erstelle drittes Sheet für "Invalid Grades"
      if (invalidGrades.length > 0) {
        const invalidGradesData = [
          ['Name', 'Note', 'Zeile', 'Bemerkung'],
          ...invalidGrades.map(entry => [entry.name, entry.grade, entry.rowIndex, 'Ungültige Note - nur halbe Noten erlaubt'])
        ];
        const invalidGradesSheet = XLSX.utils.aoa_to_sheet(invalidGradesData);
        workbook.SheetNames.push('Invalid Grades');
        workbook.Sheets['Invalid Grades'] = invalidGradesSheet;
      }
      
      // Erstelle Download
      const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], { type: 'application/octet-stream' });
      
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = 'evento_result.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      console.log('Ergebnis-Excel heruntergeladen: evento_result.xlsx');
      
    } catch (error) {
      console.error('Fehler beim Exportieren:', error.message);
    }
  }

  // ========== HAUPTPROGRAMM ==========
  async function main() {
    try {
      console.log('\n🚀 Evento Note Import gestartet...\n');
      
      // 1. SheetJS laden
      await loadSheetJS();
      
      // 2. Datei auswählen und parsen
      console.log('Bitte wählen Sie die Excel-Datei aus...');
      const { data: excelData, workbook, invalidGrades } = await createFileInput();
      
      // 3. DOM scannen
      const domMap = scanEventoPage();
      
      if (domMap.size === 0) {
        throw new Error('Keine Studierenden auf der Seite gefunden. Sind Sie auf der richtigen EventoWeb-Seite?');
      }
      
      // 4. Matching
      const { matches, notFound, notInExcel } = matchNames(excelData, domMap);
      
      // 5. Noten setzen
      const setCount = setGrades(matches);
      
      // 6. Ergebnisse ausgeben
      logResults(matches, notFound, notInExcel, invalidGrades, setCount);
      
      // 7. Excel exportieren
      exportResultExcel(workbook, matches, notFound, notInExcel, invalidGrades);
      
      console.log('✅ Evento Note Import abgeschlossen!\n');
      
    } catch (error) {
      console.error('❌ Fehler:', error.message);
    }
  }

  // Start
  main();
})();
