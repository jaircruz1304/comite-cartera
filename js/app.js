(function () {
  'use strict';

  const P = window.PortfolioProcessor;

  const state = {
    templateFile: null,
    sourceFiles: [],
    analyses: [],
    outputBlob: null,
    outputName: '',
    cutDateKey: '',
    selectedCompany: 'ALL'
  };

  const els = {};

  document.addEventListener('DOMContentLoaded', init);

  function init() {
    cacheElements();
    bindEvents();
    updateFilesUI();
    renderEmptyDashboard();
    checkRuntime();
  }

  function cacheElements() {
    [
      'templateInput', 'sourceInput', 'templateDrop', 'sourceDrop', 'templateStatus',
      'sourceStatus', 'sourceList', 'processButton', 'downloadButton', 'resetButton',
      'progressPanel', 'progressBar', 'progressText', 'messageArea', 'dashboard',
      'companyFilter', 'summaryTableBody', 'validationList', 'chartArea',
      'cutDateLabel', 'privacyBadge', 'fileCountBadge'
    ].forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    els.templateInput.addEventListener('change', (event) => setTemplateFile(event.target.files[0]));
    els.sourceInput.addEventListener('change', (event) => addSourceFiles([...event.target.files]));
    els.processButton.addEventListener('click', processFiles);
    els.downloadButton.addEventListener('click', downloadOutput);
    els.resetButton.addEventListener('click', resetSession);
    els.companyFilter.addEventListener('change', (event) => {
      state.selectedCompany = event.target.value;
      renderDashboard();
    });

    setupDropZone(els.templateDrop, (files) => setTemplateFile(files[0]));
    setupDropZone(els.sourceDrop, addSourceFiles);
  }

  function checkRuntime() {
    if (!window.XlsxPopulate) {
      showMessage('error', 'No se cargó el motor de Excel (XlsxPopulate). La red puede estar bloqueando los CDN. Recarga con Ctrl+F5; si continúa, revisa la consola del navegador.');
      els.processButton.disabled = true;
      return;
    }
    if (!window.JSZip) {
      showMessage('error', 'No se cargó el componente de compatibilidad Excel (JSZip). Recarga con Ctrl+F5 o revisa si la red bloquea los CDN.');
      els.processButton.disabled = true;
      return;
    }
    if (!P) {
      showMessage('error', 'No se cargó el motor de reglas de cartera. Confirma que js/processor.js exista y respete mayúsculas/minúsculas.');
      els.processButton.disabled = true;
      return;
    }
    updateFilesUI();
  }

  function setupDropZone(element, callback) {
    ['dragenter', 'dragover'].forEach((type) => {
      element.addEventListener(type, (event) => {
        event.preventDefault();
        element.classList.add('dragging');
      });
    });

    ['dragleave', 'drop'].forEach((type) => {
      element.addEventListener(type, (event) => {
        event.preventDefault();
        element.classList.remove('dragging');
      });
    });

    element.addEventListener('drop', (event) => callback([...event.dataTransfer.files]));
  }

  function isXlsx(file) {
    return file && /\.xlsx$/i.test(file.name);
  }

  function setTemplateFile(file) {
    clearMessage();
    if (!file) return;
    if (!isXlsx(file)) {
      showMessage('error', 'La plantilla debe ser un archivo Excel con extensión .xlsx.');
      return;
    }
    state.templateFile = file;
    state.outputBlob = null;
    updateFilesUI();
  }

  function addSourceFiles(files) {
    clearMessage();
    const valid = files.filter(isXlsx);
    if (!valid.length) {
      showMessage('error', 'Selecciona uno o varios archivos SaldosDeCarteraSencillo_Report en formato .xlsx.');
      return;
    }

    const existingKeys = new Set(state.sourceFiles.map((file) => `${file.name}|${file.size}|${file.lastModified}`));
    valid.forEach((file) => {
      const key = `${file.name}|${file.size}|${file.lastModified}`;
      if (!existingKeys.has(key)) {
        state.sourceFiles.push(file);
        existingKeys.add(key);
      }
    });

    state.outputBlob = null;
    updateFilesUI();
  }

  function removeSourceFile(index) {
    state.sourceFiles.splice(index, 1);
    state.outputBlob = null;
    updateFilesUI();
  }

  function updateFilesUI() {
    els.templateStatus.textContent = state.templateFile
      ? `${state.templateFile.name} · ${formatBytes(state.templateFile.size)}`
      : 'No se ha cargado una plantilla.';
    els.templateDrop.classList.toggle('ready', Boolean(state.templateFile));

    els.sourceStatus.textContent = state.sourceFiles.length
      ? `${state.sourceFiles.length} archivo(s) de cartera seleccionados.`
      : 'No se han cargado archivos de cartera.';
    els.sourceDrop.classList.toggle('ready', state.sourceFiles.length > 0);
    els.fileCountBadge.textContent = String(state.sourceFiles.length);

    els.sourceList.innerHTML = '';
    state.sourceFiles.forEach((file, index) => {
      const item = document.createElement('li');
      item.className = 'file-item';
      item.innerHTML = `
        <div>
          <strong>${escapeHtml(file.name)}</strong>
          <span>${formatBytes(file.size)}</span>
        </div>
        <button type="button" class="icon-button" aria-label="Quitar archivo">×</button>
      `;
      item.querySelector('button').addEventListener('click', () => removeSourceFile(index));
      els.sourceList.appendChild(item);
    });

    els.processButton.disabled = !(state.templateFile && state.sourceFiles.length && window.XlsxPopulate && window.JSZip && P);
    els.downloadButton.disabled = !state.outputBlob;
  }

  async function processFiles() {
    if (!state.templateFile || !state.sourceFiles.length) return;

    setBusy(true);
    clearMessage();
    state.outputBlob = null;
    state.analyses = [];
    state.selectedCompany = 'ALL';

    try {
      updateProgress(4, 'Validando archivos seleccionados…');
      enforceFileLimits();

      updateProgress(10, 'Abriendo una copia limpia de la plantilla COMITE…');
      const reportWorkbook = await openWorkbookSafely(state.templateFile, 'la plantilla COMITE');
      validateTemplateWorkbook(reportWorkbook);
      prepareFreshTemplate(reportWorkbook);

      const analyses = [];
      for (let index = 0; index < state.sourceFiles.length; index += 1) {
        const file = state.sourceFiles[index];
        const start = 18;
        const span = 46;
        updateProgress(start + Math.round((index / state.sourceFiles.length) * span), `Analizando ${file.name}…`);

        const sourceWorkbook = await openWorkbookSafely(file, `el archivo ${file.name}`);
        const sourceSheet = sourceWorkbook.sheet(0);
        if (!sourceSheet) throw new Error(`El archivo ${file.name} no contiene hojas.`);
        const usedRange = sourceSheet.usedRange();
        const values = usedRange ? normalizeMatrix(usedRange.value()) : [];
        const analysis = P.analyzeValues(values, file.name);
        analysis.sourceValues = values;
        analyses.push(analysis);
      }

      updateProgress(67, 'Validando empresas, fechas y consistencia…');
      state.cutDateKey = P.validateDateGroup(analyses);

      analyses.sort((a, b) => companyOrder(a.company.code) - companyOrder(b.company.code));

      updateProgress(73, 'Escribiendo resultados sobre la plantilla limpia…');
      analyses.forEach((analysis) => writeCompanyReport(reportWorkbook, analysis));

      updateProgress(83, 'Creando hojas ORIGEN y control de auditoría…');
      analyses.forEach((analysis) => addOriginSheet(reportWorkbook, analysis));
      addControlSheet(reportWorkbook, analyses);

      updateProgress(92, 'Generando el archivo Excel consolidado…');
      state.outputBlob = await reportWorkbook.outputAsync();
      state.outputName = `COMITE_Consolidado_${P.fileDateKey(analyses[0].metrics.cutDate)}_FINAL.xlsx`;
      state.analyses = analyses;

      updateProgress(100, 'Proceso completado correctamente.');
      renderDashboard();
      showMessage('success', `Se procesaron ${analyses.length} empresa(s). El archivo ${state.outputName} está listo para descargar.`);
    } catch (error) {
      console.error(error);
      showMessage('error', error && error.message ? error.message : 'Ocurrió un error inesperado durante el procesamiento.');
      renderEmptyDashboard();
    } finally {
      setBusy(false);
      updateFilesUI();
    }
  }

  async function openWorkbookSafely(file, label) {
    try {
      return await XlsxPopulate.fromDataAsync(file);
    } catch (initialError) {
      const message = String(initialError && initialError.message ? initialError.message : initialError);
      const isXmlCellError = /children|_parseNode|Cannot read propert/i.test(message);
      if (!isXmlCellError || !window.JSZip) {
        throw new Error(`No se pudo abrir ${label}: ${message}`);
      }

      updateProgress(12, `Corrigiendo compatibilidad interna de ${label}…`);
      const sanitized = await sanitizeWorkbookForPopulate(file);
      try {
        return await XlsxPopulate.fromDataAsync(sanitized.blob);
      } catch (retryError) {
        const retryMessage = String(retryError && retryError.message ? retryError.message : retryError);
        throw new Error(`No se pudo abrir ${label}, incluso después de normalizar ${sanitized.changes} celda(s) vacía(s): ${retryMessage}`);
      }
    }
  }

  async function sanitizeWorkbookForPopulate(file) {
    const input = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(input);
    const worksheetNames = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet[^/]*\.xml$/i.test(name));
    let changes = 0;

    for (const name of worksheetNames) {
      const entry = zip.file(name);
      if (!entry) continue;
      const xml = await entry.async('string');
      let normalized = xml;

      // Excel puede guardar celdas vacías como inlineStr sin nodo <is>.
      // xlsx-populate 1.21.0 intenta leer un hijo inexistente y produce
      // “Cannot read properties of undefined (reading children)”.
      normalized = normalized.replace(/<c\b([^>]*?)\s+t="inlineStr"([^>]*)\/>/g, (match, before, after) => {
        changes += 1;
        return `<c${before}${after}/>`;
      });
      normalized = normalized.replace(/<c\b([^>]*?)\s+t="inlineStr"([^>]*)>\s*<\/c>/g, (match, before, after) => {
        changes += 1;
        return `<c${before}${after}/>`;
      });
      normalized = normalized.replace(/<c\b([^>]*?)\s+t="inlineStr"([^>]*)>\s*<is\s*\/>\s*<\/c>/g, (match, before, after) => {
        changes += 1;
        return `<c${before}${after}/>`;
      });
      normalized = normalized.replace(/<c\b([^>]*?)\s+t="inlineStr"([^>]*)>\s*<is>\s*<\/is>\s*<\/c>/g, (match, before, after) => {
        changes += 1;
        return `<c${before}${after}/>`;
      });

      if (normalized !== xml) zip.file(name, normalized);
    }

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    });

    return { blob, changes };
  }

  function enforceFileLimits() {
    const allFiles = [state.templateFile, ...state.sourceFiles];
    const maxPerFile = 25 * 1024 * 1024;
    const maxTotal = 80 * 1024 * 1024;
    const oversized = allFiles.find((file) => file.size > maxPerFile);
    if (oversized) throw new Error(`El archivo ${oversized.name} supera el límite recomendado de 25 MB.`);
    const total = allFiles.reduce((sum, file) => sum + file.size, 0);
    if (total > maxTotal) throw new Error('El conjunto de archivos supera 80 MB. Procesa menos empresas por ejecución.');
  }

  function validateTemplateWorkbook(workbook) {
    const required = P.COMPANY_RULES.map((company) => company.reportSheet);
    const missing = required.filter((name) => !workbook.sheet(name));
    if (missing.length) {
      throw new Error(`La plantilla no contiene las hojas requeridas: ${missing.join(', ')}.`);
    }
  }

  function prepareFreshTemplate(workbook) {
    workbook.sheets().slice().forEach((sheet) => {
      const name = sheet.name();
      if (name.startsWith('ORIGEN_') || name === 'CONTROL_EJECUCION') sheet.delete();
    });

    P.COMPANY_RULES.forEach((company) => {
      const sheet = workbook.sheet(company.reportSheet);
      sheet.range('B2:C22').value(null);
      sheet.cell('A1').value(null);
    });
  }

  function writeCompanyReport(workbook, analysis) {
    const { metrics, company } = analysis;
    const rows = P.TEMPLATE_ROWS;
    const sheet = workbook.sheet(company.reportSheet);

    sheet.range('B2:C22').value(null);
    sheet.cell('A1').value('Valores monetarios expresados en miles');
    sheet.cell('B2').value(P.periodLabel(metrics.cutDate));

    setMetricRow(sheet, rows.TOTAL, metrics.total);
    setMetricRow(sheet, rows.ORIGINAL, metrics.original);
    setMetricRow(sheet, rows.OVERDUE, metrics.overdue);
    setRatioRow(sheet, rows.OVERDUE_RATIO, metrics.ratios.overdueTotal, metrics.ratios.overdueOperations);
    setMetricRow(sheet, rows.ND_TOTAL, metrics.noDevenga);
    setMetricRow(sheet, rows.ND_NORMAL, metrics.noDevengaNormal);
    setMetricRow(sheet, rows.ND_RESTRUCTURED, metrics.noDevengaRestructured);
    setMetricRow(sheet, rows.ND_60_90, metrics.noDevenga6090);
    setMetricRow(sheet, rows.ND_60_90_NORMAL, metrics.noDevenga6090Normal);
    setMetricRow(sheet, rows.ND_60_90_RESTRUCTURED, metrics.noDevenga6090Restructured);
    setRatioRow(sheet, rows.ND_RATIO_TOTAL, metrics.ratios.noDevengaTotal, metrics.ratios.noDevengaOperations);
    setRatioRow(sheet, rows.ND_RATIO_ORIGINAL, metrics.ratios.noDevengaOriginal, metrics.ratios.noDevengaOperations);
    setMetricRow(sheet, rows.ND_OVER_90, metrics.noDevengaOver90);
    setMetricRow(sheet, rows.ND_OVER_90_NORMAL, metrics.noDevengaOver90Normal);
    setMetricRow(sheet, rows.ND_OVER_90_RESTRUCTURED, metrics.noDevengaOver90Restructured);
    setRatioRow(sheet, rows.ND_OVER_90_RATIO, metrics.ratios.noDevengaOver90Total, metrics.ratios.noDevengaOver90Operations);

    sheet.range(`B${rows.BEP}:C${rows.BEP_NET}`).value(null);
    if (metrics.chargedOff.value == null) {
      sheet.range(`B${rows.CHARGED_OFF}:C${rows.CHARGED_OFF}`).value(null);
    } else {
      setMetricRow(sheet, rows.CHARGED_OFF, metrics.chargedOff);
    }

    sheet.range('B3:B5').style('numberFormat', '#,##0.00');
    sheet.range('B7:B12').style('numberFormat', '#,##0.00');
    sheet.range('B15:B17').style('numberFormat', '#,##0.00');
    sheet.range('B19:B22').style('numberFormat', '#,##0.00');
    sheet.range('C3:C5').style('numberFormat', '0');
    sheet.range('C7:C12').style('numberFormat', '0');
    sheet.range('C15:C17').style('numberFormat', '0');
    sheet.range('C19:C22').style('numberFormat', '0');
    [rows.OVERDUE_RATIO, rows.ND_RATIO_TOTAL, rows.ND_RATIO_ORIGINAL, rows.ND_OVER_90_RATIO]
      .forEach((row) => sheet.range(`B${row}:C${row}`).style('numberFormat', '0.00%'));
  }

  function setMetricRow(sheet, row, metric) {
    const value = metric && metric.value != null ? metric.value / 1000 : null;
    const operations = metric && metric.operations != null ? metric.operations : null;
    sheet.range(`B${row}:C${row}`).value([[value, operations]]);
  }

  function setRatioRow(sheet, row, valueRatio, operationsRatio) {
    sheet.range(`B${row}:C${row}`).value([[valueRatio, operationsRatio]]);
  }

  function addOriginSheet(workbook, analysis) {
    const name = `ORIGEN_${analysis.company.code}`;
    const existing = workbook.sheet(name);
    if (existing) existing.delete();

    const sheet = workbook.addSheet(name);
    const values = analysis.sourceValues;
    if (values.length) sheet.cell('A1').value(values);

    const columnCount = values.reduce((max, row) => Math.max(max, row.length), 0);
    for (let column = 1; column <= columnCount; column += 1) {
      const maxLength = values.slice(0, 35).reduce((max, row) => {
        const value = row[column - 1];
        return Math.max(max, String(value == null ? '' : value).length);
      }, 0);
      sheet.column(column).width(Math.min(32, Math.max(10, maxLength + 2)));
    }

    if (values.length >= 6) {
      sheet.range(`A5:${columnLetter(columnCount)}6`).style({
        bold: true,
        fill: 'DCE6F1',
        wrapText: true,
        verticalAlignment: 'center'
      });
      sheet.row(6).height(52);
      sheet.freezePanes(2, 6);
    }
  }

  function addControlSheet(workbook, analyses) {
    const existing = workbook.sheet('CONTROL_EJECUCION');
    if (existing) existing.delete();

    const sheet = workbook.addSheet('CONTROL_EJECUCION', 0);
    const now = new Date();
    const headers = [
      'EMPRESA', 'FECHA CORTE', 'ARCHIVO FUENTE', 'TOTAL CARTERA', 'OPERACIONES',
      'CARTERA VENCIDA', 'OP. VENCIDA', 'NO DEVENGA', 'OP. NO DEVENGA',
      'NO DEVENGA 60-90', 'OP. 60-90', 'NO DEVENGA +90', 'OP. +90',
      'CASTIGADA', 'OP. CASTIGADA', 'DUPLICADOS EXCLUIDOS', 'ESTADO', 'OBSERVACIONES'
    ];

    sheet.cell('A1').value('CONTROL DE EJECUCIÓN – COMITÉ DE CARTERA');
    sheet.range('A1:R1').style({
      bold: true,
      fontSize: 16,
      fontColor: 'FFFFFF',
      fill: '16324F',
      horizontalAlignment: 'center',
      verticalAlignment: 'center'
    });
    sheet.row(1).height(28);
    sheet.cell('A2').value('Fecha y hora de procesamiento');
    sheet.cell('B2').value(now).style('numberFormat', 'dd/mm/yyyy hh:mm:ss');
    sheet.cell('A3').value('Regla de plantilla');
    sheet.cell('B3').value('Copia limpia de la plantilla cargada en esta ejecución; la plantilla original no se sobrescribe.');
    sheet.cell('A4').value('Unidad monetaria');
    sheet.cell('B4').value('Valores presentados en miles; cálculos realizados en valores originales.');
    sheet.range('A2:A4').style({ bold: true, fill: 'EAF0F6' });
    sheet.range('A6:R6').value([headers]).style({
      bold: true,
      fontColor: 'FFFFFF',
      fill: '24557A',
      horizontalAlignment: 'center',
      verticalAlignment: 'center',
      wrapText: true,
      border: true
    });
    sheet.row(6).height(48);

    const rows = analyses.map(({ metrics }) => [
      metrics.company,
      metrics.cutDate,
      metrics.fileName,
      metrics.total.value / 1000,
      metrics.total.operations,
      metrics.overdue.value / 1000,
      metrics.overdue.operations,
      metrics.noDevenga.value / 1000,
      metrics.noDevenga.operations,
      metrics.noDevenga6090.value / 1000,
      metrics.noDevenga6090.operations,
      metrics.noDevengaOver90.value / 1000,
      metrics.noDevengaOver90.operations,
      metrics.chargedOff.value == null ? null : metrics.chargedOff.value / 1000,
      metrics.chargedOff.operations,
      metrics.duplicatesExcluded,
      metrics.status,
      metrics.warnings.join(' | ')
    ]);

    if (rows.length) sheet.cell('A7').value(rows);
    const lastRow = 6 + rows.length;
    if (lastRow >= 7) {
      sheet.range(`A7:R${lastRow}`).style({ border: true, verticalAlignment: 'center' });
      sheet.range(`B7:B${lastRow}`).style('numberFormat', 'dd/mm/yyyy');
      ['D', 'F', 'H', 'J', 'L', 'N'].forEach((column) => sheet.range(`${column}7:${column}${lastRow}`).style('numberFormat', '#,##0.00'));
      ['E', 'G', 'I', 'K', 'M', 'O', 'P'].forEach((column) => sheet.range(`${column}7:${column}${lastRow}`).style('numberFormat', '0'));
    }

    const widths = [12, 14, 38, 16, 12, 16, 12, 16, 14, 18, 12, 16, 12, 16, 12, 18, 12, 52];
    widths.forEach((width, index) => sheet.column(index + 1).width(width));
    sheet.freezePanes(1, 6);
  }

  function renderDashboard() {
    if (!state.analyses.length) {
      renderEmptyDashboard();
      return;
    }

    els.dashboard.classList.remove('empty');
    els.cutDateLabel.textContent = formatDate(state.analyses[0].metrics.cutDate);
    renderCompanyFilter();

    const filtered = state.selectedCompany === 'ALL'
      ? state.analyses
      : state.analyses.filter((item) => item.company.code === state.selectedCompany);

    renderKpis(filtered);
    renderChart(filtered);
    renderSummaryTable(filtered);
    renderValidations(filtered);
  }

  function renderCompanyFilter() {
    const current = state.selectedCompany;
    els.companyFilter.innerHTML = '<option value="ALL">Todas las empresas</option>';
    state.analyses.forEach((item) => {
      const option = document.createElement('option');
      option.value = item.company.code;
      option.textContent = item.company.code;
      els.companyFilter.appendChild(option);
    });
    els.companyFilter.value = current;
  }

  function renderKpis(analyses) {
    const totals = analyses.reduce((acc, item) => {
      const m = item.metrics;
      acc.total += m.total.value;
      acc.operations += m.total.operations;
      acc.overdue += m.overdue.value;
      acc.noDevenga += m.noDevenga.value;
      acc.noDevengaOver90 += m.noDevengaOver90.value;
      acc.chargedOff += m.chargedOff.value || 0;
      acc.chargedOffOperations += m.chargedOff.operations || 0;
      return acc;
    }, { total: 0, operations: 0, overdue: 0, noDevenga: 0, noDevengaOver90: 0, chargedOff: 0, chargedOffOperations: 0 });

    const ratios = {
      overdue: totals.total ? totals.overdue / totals.total : null,
      nd: totals.total ? totals.noDevenga / totals.total : null,
      nd90: totals.total ? totals.noDevengaOver90 / totals.total : null
    };

    const cards = [
      { label: 'Total cartera', value: P.formatMoneyThousands(totals.total), sub: `${totals.operations.toLocaleString('es-EC')} operaciones`, tone: 'blue' },
      { label: 'Cartera vencida', value: P.formatMoneyThousands(totals.overdue), sub: P.formatPercent(ratios.overdue), tone: P.semaphore(ratios.overdue, P.THRESHOLDS.overdueTotal) },
      { label: 'No devenga', value: P.formatMoneyThousands(totals.noDevenga), sub: P.formatPercent(ratios.nd), tone: P.semaphore(ratios.nd, P.THRESHOLDS.noDevengaTotal) },
      { label: 'No devenga +90', value: P.formatMoneyThousands(totals.noDevengaOver90), sub: P.formatPercent(ratios.nd90), tone: P.semaphore(ratios.nd90, P.THRESHOLDS.noDevengaOver90Total) },
      { label: 'Cartera castigada', value: totals.chargedOff ? P.formatMoneyThousands(totals.chargedOff) : '—', sub: totals.chargedOffOperations ? `${totals.chargedOffOperations} operaciones` : 'Sin información', tone: totals.chargedOff ? 'neutral' : 'muted' }
    ];

    const container = document.getElementById('kpiGrid');
    container.innerHTML = cards.map((card) => `
      <article class="kpi-card tone-${card.tone}">
        <span>${card.label}</span>
        <strong>${card.value}</strong>
        <small>${card.sub}</small>
      </article>
    `).join('');
  }

  function renderChart(analyses) {
    const max = Math.max(...analyses.map((item) => item.metrics.total.value), 1);
    els.chartArea.innerHTML = analyses.map(({ metrics }) => {
      const totalWidth = Math.max(4, (metrics.total.value / max) * 100);
      const ndWidth = metrics.total.value ? (metrics.noDevenga.value / metrics.total.value) * totalWidth : 0;
      const overdueWidth = metrics.total.value ? (metrics.overdue.value / metrics.total.value) * totalWidth : 0;
      return `
        <div class="chart-row">
          <div class="chart-label"><strong>${metrics.company}</strong><span>${P.formatMoneyThousands(metrics.total.value)}</span></div>
          <div class="bar-track" title="Total cartera: ${P.formatMoneyThousands(metrics.total.value)}">
            <div class="bar-total" style="width:${totalWidth}%"></div>
            <div class="bar-nd" style="width:${ndWidth}%" title="No Devenga"></div>
            <div class="bar-overdue" style="width:${overdueWidth}%" title="Vencida"></div>
          </div>
        </div>
      `;
    }).join('');
  }

  function renderSummaryTable(analyses) {
    els.summaryTableBody.innerHTML = analyses.map(({ metrics }) => `
      <tr>
        <td><strong>${metrics.company}</strong></td>
        <td class="number">${P.formatMoneyThousands(metrics.total.value)}</td>
        <td class="number">${metrics.total.operations.toLocaleString('es-EC')}</td>
        <td class="number">${P.formatMoneyThousands(metrics.overdue.value)}</td>
        <td class="number">${P.formatMoneyThousands(metrics.noDevenga.value)}</td>
        <td class="number">${P.formatMoneyThousands(metrics.noDevenga6090.value)}</td>
        <td class="number">${P.formatMoneyThousands(metrics.noDevengaOver90.value)}</td>
        <td class="number">${P.formatMoneyThousands(metrics.chargedOff.value)}</td>
        <td><span class="status-pill ${metrics.status === 'OK' ? 'ok' : 'review'}">${metrics.status}</span></td>
      </tr>
    `).join('');
  }

  function renderValidations(analyses) {
    const items = [];
    analyses.forEach(({ metrics }) => {
      items.push({ type: 'ok', text: `${metrics.company}: ${metrics.detailRows} filas de detalle validadas; no se contó la fila total.` });
      items.push({ type: metrics.duplicatesExcluded ? 'warning' : 'ok', text: `${metrics.company}: ${metrics.duplicatesExcluded} duplicados exactos excluidos.` });
      items.push({ type: 'ok', text: `${metrics.company}: No Devenga calculado solo con sus columnas específicas y segmentado por Dias Morosidad.` });
      if (metrics.chargedOff.value == null) {
        items.push({ type: 'info', text: `${metrics.company}: Cartera Castigada permanece en blanco por ausencia de información.` });
      } else {
        items.push({ type: 'ok', text: `${metrics.company}: ${metrics.chargedOff.operations} operaciones castigadas separadas de la cartera activa.` });
      }
      metrics.warnings.forEach((warning) => items.push({ type: 'warning', text: `${metrics.company}: ${warning}` }));
    });

    els.validationList.innerHTML = items.map((item) => `<li class="validation-${item.type}"><span></span>${escapeHtml(item.text)}</li>`).join('');
  }

  function renderEmptyDashboard() {
    els.dashboard.classList.add('empty');
    els.cutDateLabel.textContent = '—';
    document.getElementById('kpiGrid').innerHTML = '';
    els.chartArea.innerHTML = '<p class="placeholder">Los gráficos aparecerán después de procesar los archivos.</p>';
    els.summaryTableBody.innerHTML = '<tr><td colspan="9" class="placeholder-cell">Sin resultados procesados.</td></tr>';
    els.validationList.innerHTML = '<li class="validation-info"><span></span>Carga la plantilla y los archivos de cartera para iniciar.</li>';
  }

  function downloadOutput() {
    if (!state.outputBlob) return;
    const url = URL.createObjectURL(state.outputBlob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = state.outputName;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function resetSession() {
    state.templateFile = null;
    state.sourceFiles = [];
    state.analyses = [];
    state.outputBlob = null;
    state.outputName = '';
    state.cutDateKey = '';
    state.selectedCompany = 'ALL';
    els.templateInput.value = '';
    els.sourceInput.value = '';
    clearMessage();
    updateFilesUI();
    renderEmptyDashboard();
    updateProgress(0, 'Listo para procesar.');
  }

  function setBusy(isBusy) {
    document.body.classList.toggle('is-busy', isBusy);
    els.progressPanel.hidden = !isBusy;
    els.processButton.disabled = isBusy || !(state.templateFile && state.sourceFiles.length);
    els.resetButton.disabled = isBusy;
    els.downloadButton.disabled = isBusy || !state.outputBlob;
  }

  function updateProgress(percent, text) {
    els.progressBar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
    els.progressText.textContent = text;
  }

  function showMessage(type, text) {
    els.messageArea.className = `message message-${type}`;
    els.messageArea.textContent = text;
    els.messageArea.hidden = false;
  }

  function clearMessage() {
    els.messageArea.hidden = true;
    els.messageArea.textContent = '';
  }

  function normalizeMatrix(value) {
    if (!Array.isArray(value)) return [];
    return value.map((row) => Array.isArray(row) ? row : [row]);
  }

  function companyOrder(code) {
    const index = P.COMPANY_RULES.findIndex((company) => company.code === code);
    return index < 0 ? 999 : index;
  }

  function columnLetter(columnNumber) {
    let result = '';
    let value = columnNumber;
    while (value > 0) {
      const remainder = (value - 1) % 26;
      result = String.fromCharCode(65 + remainder) + result;
      value = Math.floor((value - 1) / 26);
    }
    return result || 'A';
  }

  function formatDate(date) {
    return date.toLocaleDateString('es-EC', { day: '2-digit', month: 'long', year: 'numeric' });
  }

  function formatBytes(bytes) {
    if (!Number.isFinite(bytes) || bytes === 0) return '0 B';
    const units = ['B', 'KB', 'MB', 'GB'];
    const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / (1024 ** index);
    return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
  }

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }
})();
