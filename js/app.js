(function () {
  'use strict';

  const P = window.PortfolioProcessor;
  const X = window.OOXMLWorkbook;

  const state = {
    templateBuffer: null,
    templateReady: false,
    sourceFiles: [],
    analyses: [],
    outputBlob: null,
    outputName: '',
    selectedCompany: 'ALL'
  };

  const els = {};
  const steps = [
    { at: 8, label: 'Validar archivos' },
    { at: 28, label: 'Analizar cartera' },
    { at: 60, label: 'Aplicar reglas' },
    { at: 80, label: 'Construir Excel' },
    { at: 100, label: 'Validar salida' }
  ];

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    cacheElements();
    bindEvents();
    renderProcessSteps(0);
    renderEmptyDashboard();
    checkRuntime();
    await loadEmbeddedTemplate();
    updateFilesUI();
  }

  function cacheElements() {
    [
      'sourceInput', 'sourceDrop', 'sourceStatus', 'sourceList', 'processButton',
      'downloadButton', 'resetButton', 'progressPanel', 'progressText', 'processSteps',
      'messageArea', 'dashboard', 'companyFilter', 'summaryTableBody', 'validationList',
      'companyCards', 'cutDateLabel', 'fileCountBadge', 'templateRepoStatus',
      'runtimeStatus', 'lastOutputLabel'
    ].forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    els.sourceInput.addEventListener('change', (event) => addSourceFiles([...event.target.files]));
    els.processButton.addEventListener('click', processFiles);
    els.downloadButton.addEventListener('click', downloadOutput);
    els.resetButton.addEventListener('click', resetSession);
    els.companyFilter.addEventListener('change', (event) => {
      state.selectedCompany = event.target.value;
      renderDashboard();
    });
    setupDropZone(els.sourceDrop, addSourceFiles);
  }

  function checkRuntime() {
    const missing = [];
    if (!window.JSZip) missing.push('JSZip');
    if (!P) missing.push('motor de reglas');
    if (!X) missing.push('motor OOXML');
    if (missing.length) {
      els.runtimeStatus.textContent = `No disponible: ${missing.join(', ')}`;
      showMessage('error', `No se cargaron componentes requeridos: ${missing.join(', ')}. Presiona Ctrl+F5.`);
    } else {
      els.runtimeStatus.textContent = 'Motores locales listos';
    }
  }

  async function loadEmbeddedTemplate() {
    state.templateReady = false;
    els.templateRepoStatus.textContent = 'Verificando plantilla integrada…';
    try {
      const response = await fetch('assets/COMITE_BASE.xlsx', { cache: 'no-store' });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const buffer = await response.arrayBuffer();
      if (buffer.byteLength < 1000) throw new Error('La plantilla está vacía o incompleta.');
      await X.inspectTemplate(buffer, P.COMPANY_RULES.map((item) => item.reportSheet));
      state.templateBuffer = buffer;
      state.templateReady = true;
      els.templateRepoStatus.textContent = 'Plantilla integrada validada';
    } catch (error) {
      const message = error && error.message ? error.message : String(error);
      els.templateRepoStatus.textContent = 'Plantilla no disponible';
      showMessage('error', `No se pudo validar la plantilla integrada: ${message}`);
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

  function addSourceFiles(files) {
    clearMessage();
    const valid = files.filter(isXlsx);
    if (!valid.length) {
      showMessage('error', 'Selecciona archivos SaldosDeCarteraSencillo_Report en formato .xlsx.');
      return;
    }
    const existing = new Set(state.sourceFiles.map(fileKey));
    valid.forEach((file) => {
      if (!existing.has(fileKey(file))) {
        state.sourceFiles.push(file);
        existing.add(fileKey(file));
      }
    });
    state.outputBlob = null;
    state.analyses = [];
    updateFilesUI();
    renderEmptyDashboard();
  }

  function fileKey(file) {
    return `${file.name}|${file.size}|${file.lastModified}`;
  }

  function removeSourceFile(index) {
    state.sourceFiles.splice(index, 1);
    state.outputBlob = null;
    state.analyses = [];
    updateFilesUI();
    renderEmptyDashboard();
  }

  function updateFilesUI() {
    els.sourceStatus.textContent = state.sourceFiles.length
      ? `${state.sourceFiles.length} archivo(s) preparados para procesar.`
      : 'Selecciona uno o varios reportes de saldos.';
    els.sourceDrop.classList.toggle('ready', state.sourceFiles.length > 0);
    els.fileCountBadge.textContent = String(state.sourceFiles.length);
    els.sourceList.innerHTML = '';

    state.sourceFiles.forEach((file, index) => {
      const item = document.createElement('li');
      item.className = 'file-item';
      item.innerHTML = `
        <div class="file-symbol">XLSX</div>
        <div class="file-copy">
          <strong>${escapeHtml(file.name)}</strong>
          <span>${formatBytes(file.size)}</span>
        </div>
        <button type="button" class="icon-button" aria-label="Quitar archivo">×</button>
      `;
      item.querySelector('button').addEventListener('click', () => removeSourceFile(index));
      els.sourceList.appendChild(item);
    });

    const runtimeReady = Boolean(window.JSZip && P && X);
    els.processButton.disabled = !(state.templateReady && state.sourceFiles.length && runtimeReady);
    els.downloadButton.disabled = !state.outputBlob;
    els.lastOutputLabel.textContent = state.outputName || 'Aún no se ha generado un consolidado';
  }

  async function processFiles() {
    if (!state.templateReady || !state.sourceFiles.length) return;
    setBusy(true);
    clearMessage();
    state.outputBlob = null;
    state.analyses = [];
    state.selectedCompany = 'ALL';

    try {
      updateProgress(8, 'Validando archivos seleccionados');
      enforceFileLimits();

      const analyses = [];
      for (let index = 0; index < state.sourceFiles.length; index += 1) {
        const file = state.sourceFiles[index];
        const percent = 20 + Math.round(((index + 1) / state.sourceFiles.length) * 34);
        updateProgress(percent, `Analizando ${file.name}`);
        await yieldToBrowser();
        const values = await readSourceMatrix(file);
        const analysis = P.analyzeValues(values, file.name);
        analysis.sourceValues = values;
        analyses.push(analysis);
      }

      updateProgress(62, 'Aplicando reglas de consistencia y segmentación');
      P.validateDateGroup(analyses);
      analyses.sort((a, b) => companyOrder(a.company.code) - companyOrder(b.company.code));

      updateProgress(80, 'Construyendo una copia nueva de la plantilla');
      const generated = await withTimeout(
        X.build(state.templateBuffer, analyses, {
          templateRows: P.TEMPLATE_ROWS,
          requiredSheets: P.COMPANY_RULES.map((item) => item.reportSheet),
          periodLabel: P.periodLabel
        }),
        90000,
        'La construcción del consolidado superó 90 segundos.'
      );

      updateProgress(96, 'Verificando integridad y hojas del archivo');
      const outputBuffer = toStandaloneArrayBuffer(generated);
      await validateXlsxPackage(outputBuffer, [
        'CONTROL_EJECUCION',
        ...analyses.flatMap((item) => [item.company.reportSheet, `ORIGEN_${item.company.code}`])
      ]);

      state.outputBlob = new Blob([outputBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      state.outputName = `COMITE_Consolidado_${P.fileDateKey(analyses[0].metrics.cutDate)}_FINAL.xlsx`;
      state.analyses = analyses;
      updateProgress(100, 'Archivo consolidado validado y listo');
      renderDashboard();
      showMessage('success', `Proceso completado. ${state.outputName} está validado y listo para descargar.`);
    } catch (error) {
      console.error(error);
      showMessage('error', error && error.message ? error.message : 'Ocurrió un error inesperado.');
      renderEmptyDashboard();
    } finally {
      setBusy(false);
      updateFilesUI();
    }
  }

  async function readSourceMatrix(file) {
    try {
      return await X.readFirstSheet(await file.arrayBuffer());
    } catch (error) {
      throw new Error(`No se pudo leer ${file.name}: ${error.message || error}`);
    }
  }

  function enforceFileLimits() {
    const maxPerFile = 25 * 1024 * 1024;
    const maxTotal = 80 * 1024 * 1024;
    const oversized = state.sourceFiles.find((file) => file.size > maxPerFile);
    if (oversized) throw new Error(`El archivo ${oversized.name} supera 25 MB.`);
    const total = state.sourceFiles.reduce((sum, file) => sum + file.size, 0);
    if (total > maxTotal) throw new Error('El conjunto de archivos supera 80 MB.');
  }

  async function validateXlsxPackage(arrayBuffer, requiredSheets) {
    let zip;
    try {
      zip = await JSZip.loadAsync(arrayBuffer);
    } catch (error) {
      throw new Error('El archivo generado no es un paquete XLSX válido.');
    }
    ['[Content_Types].xml', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels'].forEach((part) => {
      if (!zip.file(part)) throw new Error(`El Excel generado está incompleto: falta ${part}.`);
    });
    try {
      await X.inspectTemplate(arrayBuffer, requiredSheets);
    } catch (error) {
      throw new Error(`El Excel generado no superó la verificación estructural: ${error.message || error}`);
    }
  }

  function renderDashboard() {
    if (!state.analyses.length) return renderEmptyDashboard();
    els.dashboard.classList.remove('empty');
    els.cutDateLabel.textContent = formatDate(state.analyses[0].metrics.cutDate);
    renderCompanyFilter();
    const filtered = state.selectedCompany === 'ALL'
      ? state.analyses
      : state.analyses.filter((item) => item.company.code === state.selectedCompany);
    renderKpis(filtered);
    renderCompanyCards(filtered);
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

  function aggregate(analyses) {
    return analyses.reduce((acc, item) => {
      const m = item.metrics;
      acc.total += m.total.value;
      acc.operations += m.total.operations;
      acc.overdue += m.overdue.value;
      acc.overdueOps += m.overdue.operations;
      acc.noDevenga += m.noDevenga.value;
      acc.noDevengaOps += m.noDevenga.operations;
      acc.noDevengaOver90 += m.noDevengaOver90.value;
      acc.chargedOff += m.chargedOff.value || 0;
      acc.chargedOffOperations += m.chargedOff.operations || 0;
      return acc;
    }, { total: 0, operations: 0, overdue: 0, overdueOps: 0, noDevenga: 0, noDevengaOps: 0, noDevengaOver90: 0, chargedOff: 0, chargedOffOperations: 0 });
  }

  function renderKpis(analyses) {
    const totals = aggregate(analyses);
    const cards = [
      { icon: 'Σ', label: 'Total cartera', value: P.formatMoneyThousands(totals.total), meta: `${formatInteger(totals.operations)} operaciones`, tone: 'navy' },
      { icon: 'V', label: 'Cartera vencida', value: P.formatMoneyThousands(totals.overdue), meta: P.formatPercent(safeDivide(totals.overdue, totals.total)), tone: P.semaphore(safeDivide(totals.overdue, totals.total), P.THRESHOLDS.overdueTotal) },
      { icon: 'ND', label: 'No Devenga', value: P.formatMoneyThousands(totals.noDevenga), meta: P.formatPercent(safeDivide(totals.noDevenga, totals.total)), tone: P.semaphore(safeDivide(totals.noDevenga, totals.total), P.THRESHOLDS.noDevengaTotal) },
      { icon: '+90', label: 'No Devenga +90', value: P.formatMoneyThousands(totals.noDevengaOver90), meta: P.formatPercent(safeDivide(totals.noDevengaOver90, totals.total)), tone: P.semaphore(safeDivide(totals.noDevengaOver90, totals.total), P.THRESHOLDS.noDevengaOver90Total) },
      { icon: 'C', label: 'Cartera castigada', value: totals.chargedOff ? P.formatMoneyThousands(totals.chargedOff) : '—', meta: totals.chargedOffOperations ? `${formatInteger(totals.chargedOffOperations)} operaciones` : 'Sin información', tone: totals.chargedOff ? 'slate' : 'muted' }
    ];
    document.getElementById('kpiGrid').innerHTML = cards.map((card) => `
      <article class="kpi-card tone-${card.tone}">
        <div class="kpi-icon">${card.icon}</div>
        <div class="kpi-copy"><span>${card.label}</span><strong>${card.value}</strong><small>${card.meta}</small></div>
      </article>
    `).join('');
  }

  function renderCompanyCards(analyses) {
    els.companyCards.innerHTML = analyses.map(({ metrics }) => {
      const ndRatio = safeDivide(metrics.noDevenga.value, metrics.total.value) || 0;
      const overdueRatio = safeDivide(metrics.overdue.value, metrics.total.value) || 0;
      const statusTone = metrics.status === 'OK' ? 'ok' : 'review';
      return `
        <article class="company-card">
          <header>
            <div><span class="company-kicker">EMPRESA</span><h4>${metrics.company}</h4></div>
            <span class="status-pill ${statusTone}">${metrics.status}</span>
          </header>
          <div class="company-total">
            <span>Total cartera · miles</span>
            <strong>${P.formatMoneyThousands(metrics.total.value)}</strong>
            <small>${formatInteger(metrics.total.operations)} operaciones activas</small>
          </div>
          <div class="gauge-grid">
            ${gaugeMarkup('No Devenga', ndRatio, P.semaphore(ndRatio, P.THRESHOLDS.noDevengaTotal))}
            ${gaugeMarkup('Vencida', overdueRatio, P.semaphore(overdueRatio, P.THRESHOLDS.overdueTotal))}
          </div>
          <dl class="metric-list">
            <div><dt>ND 60–90</dt><dd>${P.formatMoneyThousands(metrics.noDevenga6090.value)}</dd></div>
            <div><dt>ND +90</dt><dd>${P.formatMoneyThousands(metrics.noDevengaOver90.value)}</dd></div>
            <div><dt>Castigada</dt><dd>${P.formatMoneyThousands(metrics.chargedOff.value)}</dd></div>
          </dl>
        </article>
      `;
    }).join('');
  }

  function gaugeMarkup(label, ratio, tone) {
    const percent = Math.max(0, Math.min(100, Math.round(ratio * 100)));
    return `
      <div class="gauge-item">
        <div class="gauge gauge-${tone}" style="--gauge-value:${percent}">
          <div><strong>${percent}%</strong><span>${label}</span></div>
        </div>
      </div>
    `;
  }

  function renderSummaryTable(analyses) {
    els.summaryTableBody.innerHTML = analyses.map(({ metrics }) => `
      <tr>
        <td><strong>${metrics.company}</strong></td>
        <td class="number">${P.formatMoneyThousands(metrics.total.value)}</td>
        <td class="number">${formatInteger(metrics.total.operations)}</td>
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
      items.push({ type: 'ok', text: `${metrics.company}: ${formatInteger(metrics.detailRows)} filas de detalle validadas; la fila total fue excluida.` });
      items.push({ type: metrics.duplicatesExcluded ? 'warning' : 'ok', text: `${metrics.company}: ${formatInteger(metrics.duplicatesExcluded)} duplicados exactos excluidos.` });
      items.push({ type: 'ok', text: `${metrics.company}: No Devenga se calculó únicamente desde sus columnas específicas.` });
      items.push({ type: metrics.chargedOff.value == null ? 'info' : 'ok', text: metrics.chargedOff.value == null
        ? `${metrics.company}: Cartera Castigada permanece en blanco.`
        : `${metrics.company}: ${formatInteger(metrics.chargedOff.operations)} operaciones castigadas separadas de la cartera activa.` });
      metrics.warnings.forEach((warning) => items.push({ type: 'warning', text: `${metrics.company}: ${warning}` }));
    });
    els.validationList.innerHTML = items.map((item) => `<li class="validation-${item.type}"><span></span>${escapeHtml(item.text)}</li>`).join('');
  }

  function renderEmptyDashboard() {
    els.dashboard.classList.add('empty');
    els.cutDateLabel.textContent = '—';
    document.getElementById('kpiGrid').innerHTML = '';
    els.companyCards.innerHTML = '<div class="empty-state"><strong>Sin resultados</strong><span>Carga los reportes de saldos y ejecuta el procesamiento.</span></div>';
    els.summaryTableBody.innerHTML = '<tr><td colspan="9" class="placeholder-cell">Sin resultados procesados.</td></tr>';
    els.validationList.innerHTML = '<li class="validation-info"><span></span>La validación aparecerá al finalizar el procesamiento.</li>';
  }

  function renderProcessSteps(percent) {
    if (!els.processSteps) return;
    els.processSteps.innerHTML = steps.map((step, index) => {
      const complete = percent >= step.at;
      const active = !complete && percent > (index === 0 ? 0 : steps[index - 1].at);
      return `<li class="${complete ? 'complete' : active ? 'active' : ''}"><span>${complete ? '✓' : index + 1}</span><strong>${step.label}</strong></li>`;
    }).join('');
  }

  function updateProgress(percent, text) {
    els.progressText.textContent = text;
    renderProcessSteps(percent);
  }

  function setBusy(isBusy) {
    document.body.classList.toggle('is-busy', isBusy);
    els.progressPanel.hidden = !isBusy;
    const runtimeReady = Boolean(window.JSZip && P && X);
    els.processButton.disabled = isBusy || !(state.templateReady && state.sourceFiles.length && runtimeReady);
    els.resetButton.disabled = isBusy;
    els.downloadButton.disabled = isBusy || !state.outputBlob;
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
    window.setTimeout(() => URL.revokeObjectURL(url), 3000);
  }

  function resetSession() {
    state.sourceFiles = [];
    state.analyses = [];
    state.outputBlob = null;
    state.outputName = '';
    state.selectedCompany = 'ALL';
    els.sourceInput.value = '';
    clearMessage();
    updateFilesUI();
    renderEmptyDashboard();
    updateProgress(0, 'Listo para procesar');
  }

  function withTimeout(promise, milliseconds, message) {
    return Promise.race([
      promise,
      new Promise((_, reject) => window.setTimeout(() => reject(new Error(message)), milliseconds))
    ]);
  }

  function toStandaloneArrayBuffer(value) {
    if (value instanceof ArrayBuffer) return value;
    if (ArrayBuffer.isView(value)) return value.buffer.slice(value.byteOffset, value.byteOffset + value.byteLength);
    throw new Error('El generador no devolvió un archivo binario válido.');
  }

  function yieldToBrowser() {
    return new Promise((resolve) => window.setTimeout(resolve, 0));
  }

  function safeDivide(numerator, denominator) {
    return denominator ? numerator / denominator : null;
  }

  function companyOrder(code) {
    const index = P.COMPANY_RULES.findIndex((company) => company.code === code);
    return index < 0 ? 999 : index;
  }

  function formatInteger(value) {
    return Number(value || 0).toLocaleString('es-EC', { maximumFractionDigits: 0 });
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

  function showMessage(type, text) {
    els.messageArea.className = `message message-${type}`;
    els.messageArea.textContent = text;
    els.messageArea.hidden = false;
  }

  function clearMessage() {
    els.messageArea.hidden = true;
    els.messageArea.textContent = '';
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
