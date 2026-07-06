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
    COL_GRADE: 'Note',
    MAX_FILE_SIZE: 10 * 1024 * 1024 // 10MB
  };

  // ============ STYLES ============
  function injectStyles() {
    const styles = `
      #evento-importer-toggle {
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #0064A3;
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        cursor: pointer;
        font-family: Arial, sans-serif;
        font-size: 14px;
        font-weight: bold;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3);
        z-index: 10000;
        transition: all 0.3s ease;
      }
      #evento-importer-toggle:hover {
        background: #005288;
        transform: translateY(-2px);
      }
      #evento-importer-panel {
        position: fixed;
        bottom: 20px;
        right: 20px;
        width: 400px;
        max-height: 80vh;
        background: #FFFFFF;
        border-radius: 12px;
        box-shadow: 0 8px 32px rgba(0,0,0,0.3);
        z-index: 10001;
        font-family: Arial, sans-serif;
        font-size: 14px;
        display: none;
        flex-direction: column;
        overflow: hidden;
      }
      #evento-importer-panel.visible {
        display: flex;
      }
      .evento-header {
        background: #0064A3;
        color: white;
        padding: 16px 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
      }
      .evento-header h3 {
        margin: 0;
        font-size: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
      }
      .evento-close-btn {
        background: none;
        border: none;
        color: white;
        font-size: 24px;
        cursor: pointer;
        padding: 0;
        line-height: 1;
        opacity: 0.8;
        transition: opacity 0.2s;
      }
      .evento-close-btn:hover {
        opacity: 1;
      }
      .evento-body {
        padding: 20px;
        overflow-y: auto;
        flex: 1;
      }
      .evento-dropzone {
        border: 2px dashed #DEE2E6;
        border-radius: 8px;
        padding: 30px 20px;
        text-align: center;
        cursor: pointer;
        transition: all 0.3s ease;
        background: #F8F9FA;
        margin-bottom: 16px;
      }
      .evento-dropzone:hover, .evento-dropzone.dragover {
        border-color: #0064A3;
        background: #E8F4FC;
      }
      .evento-dropzone-icon {
        font-size: 32px;
        margin-bottom: 8px;
      }
      .evento-dropzone-text {
        color: #6C757D;
        font-size: 13px;
      }
      .evento-file-input {
        display: none;
      }
      .evento-section-title {
        color: #6C757D;
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 16px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 1px solid #DEE2E6;
      }
      .evento-status {
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 12px;
        font-size: 13px;
      }
      .evento-status.info {
        background: #E8F4FC;
        color: #0064A3;
      }
      .evento-status.success {
        background: #D4EDDA;
        color: #155724;
      }
      .evento-status.error {
        background: #F8D7DA;
        color: #721C24;
      }
      .evento-status.warning {
        background: #FFF3CD;
        color: #856404;
      }
      .evento-result-item {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 6px 0;
        font-size: 13px;
      }
      .evento-result-item.matched { color: #28A745; }
      .evento-result-item.not-found { color: #DC3545; }
      .evento-result-item.not-in-excel { color: #DC3545; }
      .evento-result-item.invalid { color: #FFC107; }
      .evento-list {
        background: #F8F9FA;
        border-radius: 6px;
        padding: 8px 12px;
        margin: 8px 0;
        max-height: 120px;
        overflow-y: auto;
        font-size: 12px;
      }
      .evento-list-item {
        padding: 4px 0;
        border-bottom: 1px solid #DEE2E6;
      }
      .evento-list-item:last-child {
        border-bottom: none;
      }
      .evento-buttons {
        display: flex;
        gap: 12px;
        margin-top: 16px;
        padding-top: 16px;
        border-top: 1px solid #DEE2E6;
      }
      .evento-btn {
        flex: 1;
        padding: 12px 16px;
        border: none;
        border-radius: 6px;
        cursor: pointer;
        font-size: 14px;
        font-weight: bold;
        transition: all 0.2s ease;
      }
      .evento-btn:disabled {
        opacity: 0.5;
        cursor: not-allowed;
      }
      .evento-btn-primary {
        background: #0064A3;
        color: white;
      }
      .evento-btn-primary:hover:not(:disabled) {
        background: #005288;
      }
      .evento-btn-secondary {
        background: #6C757D;
        color: white;
      }
      .evento-btn-secondary:hover:not(:disabled) {
        background: #5A6268;
      }
      .evento-highlight {
        background-color: #D4EDDA !important;
        transition: background-color 0.3s ease;
      }
      .evento-file-info {
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 10px;
        background: #D4EDDA;
        border-radius: 6px;
        margin-bottom: 12px;
        color: #155724;
        font-size: 13px;
      }
    `;
    const styleEl = document.createElement('style');
    styleEl.textContent = styles;
    document.head.appendChild(styleEl);
  }

  // ============ UI MODULE ============
  const UI = {
    panel: null,
    toggleBtn: null,
    statusEl: null,
    resultEl: null,
    dropzone: null,
    fileInput: null,
    applyBtn: null,
    downloadBtn: null,
    fileInfoEl: null,

    init() {
      this.createToggleButton();
      this.createPanel();
    },

    createToggleButton() {
      this.toggleBtn = document.createElement('div');
      this.toggleBtn.id = 'evento-importer-toggle';
      this.toggleBtn.innerHTML = '📋 Noten Import';
      this.toggleBtn.addEventListener('click', () => this.showPanel());
      document.body.appendChild(this.toggleBtn);
    },

    createPanel() {
      this.panel = document.createElement('div');
      this.panel.id = 'evento-importer-panel';
      this.panel.innerHTML = `
        <div class="evento-header">
          <h3>📋 Evento Note Importer</h3>
          <button class="evento-close-btn">&times;</button>
        </div>
        <div class="evento-body">
          <div class="evento-dropzone">
            <div class="evento-dropzone-icon">📁</div>
            <div class="evento-dropzone-text">Excel-Datei hier ablegen<br>oder klicken zum Auswählen</div>
          </div>
          <input type="file" class="evento-file-input" accept=".xlsx">
          <div class="evento-file-info" style="display: none;"></div>
          <div class="evento-section-title">Status</div>
          <div class="evento-status info">Bereit. Bitte Excel-Datei hochladen.</div>
          <div class="evento-section-title">Ergebnis</div>
          <div class="evento-result-area">(noch keine Daten)</div>
          <div class="evento-buttons">
            <button class="evento-btn evento-btn-primary" disabled>Anwenden</button>
            <button class="evento-btn evento-btn-secondary" disabled>Excel herunterladen</button>
          </div>
        </div>
      `;
      document.body.appendChild(this.panel);

      // Cache references
      this.statusEl = this.panel.querySelector('.evento-status');
      this.resultEl = this.panel.querySelector('.evento-result-area');
      this.dropzone = this.panel.querySelector('.evento-dropzone');
      this.fileInput = this.panel.querySelector('.evento-file-input');
      this.applyBtn = this.panel.querySelector('.evento-btn-primary');
      this.downloadBtn = this.panel.querySelector('.evento-btn-secondary');
      this.fileInfoEl = this.panel.querySelector('.evento-file-info');

      this.setupEventHandlers();
    },

    setupEventHandlers() {
      // Close button
      this.panel.querySelector('.evento-close-btn').addEventListener('click', () => this.hidePanel());

      // Dropzone events
      this.dropzone.addEventListener('click', () => this.fileInput.click());
      this.dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        this.dropzone.classList.add('dragover');
      });
      this.dropzone.addEventListener('dragleave', () => {
        this.dropzone.classList.remove('dragover');
      });
      this.dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        this.dropzone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) this.handleFile(file);
      });

      // File input
      this.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) this.handleFile(file);
      });

      // Apply button
      this.applyBtn.addEventListener('click', () => {
        if (State.matchResult) {
          const count = Evento.applyGrades(State.matchResult.matched);
          this.updateStatus(`✅ ${count} Noten wurden gesetzt.`, 'success');
          State.gradesApplied = true;
        }
      });

      // Download button
      this.downloadBtn.addEventListener('click', () => {
        if (State.workbook && State.matchResult) {
          ExcelModule.export(State.workbook, State.matchResult, State.invalidGrades);
        }
      });
    },

    async handleFile(file) {
      // Validate file type
      if (!file.name.endsWith('.xlsx')) {
        this.updateStatus('❌ Nur .xlsx Dateien werden unterstützt.', 'error');
        return;
      }

      // Validate file size
      if (file.size > CONFIG.MAX_FILE_SIZE) {
        this.updateStatus('❌ Datei zu gross (max. 10MB).', 'error');
        return;
      }

      this.updateStatus('⏳ Datei wird verarbeitet...', 'info');
      this.showFileInfo(file.name);

      try {
        const parseResult = await ExcelModule.parse(file);
        if (!parseResult.success) {
          this.updateStatus(`❌ ${parseResult.error}`, 'error');
          return;
        }

        State.workbook = parseResult.workbook;
        State.excelData = parseResult.data;
        State.invalidGrades = parseResult.invalidGrades;

        // Scan DOM
        Evento.scanPage();

        // Match names
        State.matchResult = Matcher.match(State.excelData, Evento.domMap);

        // Show results
        this.showResults(State.matchResult, State.invalidGrades);
        this.updateStatus(`✅ Datei geladen: ${file.name}`, 'success');

        // Enable buttons
        this.applyBtn.disabled = State.matchResult.matched.length === 0;
        this.downloadBtn.disabled = false;

      } catch (error) {
        console.error('Evento Importer Error:', error);
        this.updateStatus(`❌ Fehler: ${error.message}`, 'error');
      }
    },

    showFileInfo(filename) {
      this.fileInfoEl.innerHTML = `✅ Datei geladen: <strong>${filename}</strong>`;
      this.fileInfoEl.style.display = 'flex';
    },

    showPanel() {
      this.panel.classList.add('visible');
      this.toggleBtn.style.display = 'none';
    },

    hidePanel() {
      this.panel.classList.remove('visible');
      this.toggleBtn.style.display = 'block';
    },

    updateStatus(message, type = 'info') {
      this.statusEl.textContent = message;
      this.statusEl.className = `evento-status ${type}`;
    },

    showResults(results, invalidGrades) {
      let html = '';

      // Summary
      html += `<div class="evento-result-item matched">✓ Gefunden: ${results.matched.length}</div>`;
      if (results.notFound.length > 0) {
        html += `<div class="evento-result-item not-found">✗ In Excel, nicht in Evento: ${results.notFound.length}</div>`;
      }
      if (results.notInExcel.length > 0) {
        html += `<div class="evento-result-item not-in-excel">✗ In Evento, nicht in Excel: ${results.notInExcel.length}</div>`;
      }
      if (invalidGrades.length > 0) {
        html += `<div class="evento-result-item invalid">⚠ Ungültige Noten: ${invalidGrades.length}</div>`;
      }

      // Not found list
      if (results.notFound.length > 0) {
        html += `<div style="margin-top: 12px; font-size: 12px; color: #6C757D;">In Excel, nicht in Evento:</div>`;
        html += '<div class="evento-list">';
        results.notFound.forEach(item => {
          html += `<div class="evento-list-item">• ${item.name} (${item.grade})</div>`;
        });
        html += '</div>';
      }

      // Not in Excel list
      if (results.notInExcel.length > 0) {
        html += `<div style="margin-top: 12px; font-size: 12px; color: #6C757D;">In Evento, nicht in Excel:</div>`;
        html += '<div class="evento-list">';
        results.notInExcel.forEach(item => {
          html += `<div class="evento-list-item">• ${item.name}</div>`;
        });
        html += '</div>';
      }

      // Invalid grades list
      if (invalidGrades.length > 0) {
        html += `<div style="margin-top: 12px; font-size: 12px; color: #6C757D;">Ungültige Noten (nicht gesetzt):</div>`;
        html += '<div class="evento-list">';
        invalidGrades.forEach(item => {
          html += `<div class="evento-list-item">• ${item.name}: "${item.grade}" (Zeile ${item.rowIndex})</div>`;
        });
        html += '</div>';
      }

      this.resultEl.innerHTML = html;
    }
  };

  // ============ STATE ============
  const State = {
    workbook: null,
    excelData: [],
    invalidGrades: [],
    matchResult: null,
    gradesApplied: false
  };

  // ============ EXCEL MODULE ============
  const ExcelModule = {
    async parse(file) {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Check if sheet exists
            if (!workbook.SheetNames.includes(CONFIG.SHEET_NAME)) {
              resolve({ success: false, error: `Sheet '${CONFIG.SHEET_NAME}' nicht gefunden.` });
              return;
            }

            const sheet = workbook.Sheets[CONFIG.SHEET_NAME];
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            if (jsonData.length < 2) {
              resolve({ success: false, error: 'Sheet enthält keine Daten.' });
              return;
            }

            // Find column indices
            const headers = jsonData[0].map(h => String(h).trim());
            const nameColIdx = headers.indexOf(CONFIG.COL_NAME);
            const gradeColIdx = headers.indexOf(CONFIG.COL_GRADE);

            if (nameColIdx === -1) {
              resolve({ success: false, error: `Spalte '${CONFIG.COL_NAME}' nicht gefunden.` });
              return;
            }
            if (gradeColIdx === -1) {
              resolve({ success: false, error: `Spalte '${CONFIG.COL_GRADE}' nicht gefunden.` });
              return;
            }

            const validData = [];
            const invalidGrades = [];

            // Process rows (skip header)
            for (let i = 1; i < jsonData.length; i++) {
              const row = jsonData[i];
              const name = row[nameColIdx] ? String(row[nameColIdx]).trim() : '';
              const grade = row[gradeColIdx] !== undefined ? String(row[gradeColIdx]).trim() : '';

              if (!name) continue; // Skip empty names

              if (!grade) continue; // Skip empty grades

              // Validate grade
              const normalizedGrade = this.normalizeGrade(grade);
              if (normalizedGrade !== null) {
                validData.push({ name, grade: normalizedGrade, rowIndex: i + 1 });
              } else {
                invalidGrades.push({ name, grade, rowIndex: i + 1, reason: 'Ungültige Note (nur Viertelnoten erlaubt)' });
              }
            }

            resolve({
              success: true,
              data: validData,
              invalidGrades,
              workbook,
              error: null
            });

          } catch (error) {
            resolve({ success: false, error: error.message });
          }
        };
        reader.onerror = () => {
          resolve({ success: false, error: 'Datei konnte nicht gelesen werden.' });
        };
        reader.readAsArrayBuffer(file);
      });
    },

    normalizeGrade(grade) {
      const num = parseFloat(String(grade).replace(',', '.'));
      if (isNaN(num) || num < 1 || num > 6) return null;

      // Nur exakte Viertelnoten sind gültig (0.25-Schritte), da dies die
      // feinste in EventoWeb vorkommende Auflösung ist (Bachelor: 0.5,
      // Master: teilweise 0.25).
      const quarterSteps = num * 4;
      if (Math.abs(quarterSteps - Math.round(quarterSteps)) > 1e-9) return null;

      // Format wie im EventoWeb-Dropdown (zwei Nachkommastellen, z.B. "5.25")
      return num.toFixed(2);
    },

    export(workbook, results, invalidGrades) {
      try {
        const sheet = workbook.Sheets[CONFIG.SHEET_NAME];

        // Find headers to get column positions
        const range = XLSX.utils.decode_range(sheet['!ref']);
        let resultColLetter = null;

        // Find or create Resultat column
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
          const cell = sheet[cellAddr];
          if (cell && cell.v === 'Resultat') {
            resultColLetter = XLSX.utils.encode_col(c);
            break;
          }
        }

        if (!resultColLetter) {
          // Add new column
          const newColIdx = range.e.c + 1;
          resultColLetter = XLSX.utils.encode_col(newColIdx);
          sheet[`${resultColLetter}1`] = { t: 's', v: 'Resultat' };
          range.e.c = newColIdx;
          sheet['!ref'] = XLSX.utils.encode_range(range);
        }

        // Set result values for matched
        results.matched.forEach(item => {
          const cellAddr = `${resultColLetter}${item.rowIndex}`;
          sheet[cellAddr] = { t: 's', v: 'Matched' };
        });

        // Set result values for not found
        results.notFound.forEach(item => {
          const cellAddr = `${resultColLetter}${item.rowIndex}`;
          sheet[cellAddr] = { t: 's', v: 'Not in Evento' };
        });

        // Set result values for invalid grades
        invalidGrades.forEach(item => {
          const cellAddr = `${resultColLetter}${item.rowIndex}`;
          sheet[cellAddr] = { t: 's', v: 'Invalid Grade' };
        });

        // Create "Not in Excel" sheet if needed
        if (results.notInExcel.length > 0) {
          const notInExcelData = [
            ['Name', 'Bemerkung'],
            ...results.notInExcel.map(e => [e.name, 'In Evento, nicht in Excel'])
          ];
          const notInExcelSheet = XLSX.utils.aoa_to_sheet(notInExcelData);
          
          // Remove existing sheet if present
          const existingIdx = workbook.SheetNames.indexOf('Not in Excel');
          if (existingIdx > -1) {
            workbook.SheetNames.splice(existingIdx, 1);
            delete workbook.Sheets['Not in Excel'];
          }
          
          workbook.SheetNames.push('Not in Excel');
          workbook.Sheets['Not in Excel'] = notInExcelSheet;
        }

        // Create "Invalid Grades" sheet if needed
        if (invalidGrades.length > 0) {
          const invalidData = [
            ['Name', 'Note', 'Zeile', 'Bemerkung'],
            ...invalidGrades.map(e => [e.name, e.grade, e.rowIndex, 'Ungültige Note - nur Viertelnoten erlaubt'])
          ];
          const invalidSheet = XLSX.utils.aoa_to_sheet(invalidData);
          
          // Remove existing sheet if present
          const existingIdx = workbook.SheetNames.indexOf('Invalid Grades');
          if (existingIdx > -1) {
            workbook.SheetNames.splice(existingIdx, 1);
            delete workbook.Sheets['Invalid Grades'];
          }
          
          workbook.SheetNames.push('Invalid Grades');
          workbook.Sheets['Invalid Grades'] = invalidSheet;
        }

        // Generate and download
        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = 'evento_result.xlsx';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

      } catch (error) {
        console.error('Export error:', error);
        UI.updateStatus(`❌ Export-Fehler: ${error.message}`, 'error');
      }
    }
  };

  // ============ EVENTO MODULE ============
  const Evento = {
    domMap: new Map(),

    scanPage() {
      this.domMap.clear();

      // Find all student name links
      const nameLinks = document.querySelectorAll('a[href*="Brn_PersonDetailMA"]');

      nameLinks.forEach(link => {
        const name = link.textContent.trim();
        const row = link.closest('tr');
        if (row) {
          const select = row.querySelector('select');
          if (select) {
            this.domMap.set(name, {
              name,
              selectElement: select,
              row
            });
          }
        }
      });

      console.log(`Evento Importer: ${this.domMap.size} Studierende auf der Seite gefunden.`);
    },

    setGrade(selectElement, grade) {
      if (!selectElement) return false;

      // Statt einer fest codierten Werte-Tabelle wird die passende Option
      // über ihren sichtbaren Text (z.B. "5.25") gesucht. Damit funktioniert
      // das Skript unabhängig davon, ob die Seite nur halbe Noten (Bachelor)
      // oder auch Viertelnoten (Master) anbietet, und unabhängig von den
      // internen (nicht-linearen) Option-Values.
      const commaGrade = grade.replace('.', ',');
      const option = Array.from(selectElement.options)
        .find(opt => {
          const text = opt.textContent.trim();
          return text === grade || text === commaGrade;
        });

      if (option) {
        selectElement.value = option.value;
        selectElement.dispatchEvent(new Event('change', { bubbles: true }));
        this.highlightElement(selectElement);
        return true;
      }
      return false;
    },

    highlightElement(element) {
      element.classList.add('evento-highlight');
      setTimeout(() => {
        element.classList.remove('evento-highlight');
      }, 3000);
    },

    applyGrades(matches) {
      let count = 0;
      matches.forEach(item => {
        if (this.setGrade(item.selectElement, item.grade)) {
          count++;
        }
      });
      return count;
    }
  };

  // ============ MATCHER MODULE ============
  const Matcher = {
    match(excelData, domMap) {
      const matched = [];
      const notFound = [];
      const matchedNames = new Set();

      // Match Excel entries to DOM
      excelData.forEach(item => {
        const domEntry = domMap.get(item.name);
        if (domEntry) {
          matched.push({
            excelName: item.name,
            grade: item.grade,
            selectElement: domEntry.selectElement,
            rowIndex: item.rowIndex
          });
          matchedNames.add(item.name);
        } else {
          notFound.push({
            name: item.name,
            grade: item.grade,
            rowIndex: item.rowIndex
          });
        }
      });

      // Find DOM entries not in Excel
      const notInExcel = [];
      domMap.forEach((entry, name) => {
        if (!matchedNames.has(name)) {
          // Check if any Excel entry matches this name
          const existsInExcel = excelData.some(e => e.name === name);
          if (!existsInExcel) {
            notInExcel.push({
              name: entry.name,
              selectElement: entry.selectElement
            });
          }
        }
      });

      return { matched, notFound, notInExcel };
    }
  };

  // ============ MAIN ============
  function main() {
    console.log('Evento Note Importer v1.0 geladen');
    injectStyles();
    UI.init();
    Evento.scanPage();
  }

  // Start when DOM ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', main);
  } else {
    main();
  }
})();
