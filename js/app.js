(function () {
  'use strict';

  const P = window.PortfolioProcessor;
  const R = window.ProjectionEngine;
  const X = window.OOXMLWorkbook;

  const state = {
    templateBuffer: null,
    templateReady: false,
    sourceFiles: [],
    normalAnalyses: [],
    projectedAnalyses: [],
    normalOutputBlob: null,
    projectedOutputBlob: null,
    normalOutputName: '',
    projectedOutputName: '',
    selectedCompany: 'ALL',
    selectedScenario: 'NORMAL',
    baseDate: null,
    targetDate: null,
    projectionParams: null
  };

  const els = {};
  const steps = [
    { at: 8, label: 'Validar archivos' },
    { at: 28, label: 'Calcular Normal' },
    { at: 48, label: 'Calcular Proyectado' },
    { at: 70, label: 'Construir Excel Normal' },
    { at: 88, label: 'Construir Excel Proyectado' },
    { at: 100, label: 'Validar salidas' }
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
    updateProjectionOutputs();
  }

  function cacheElements() {
    [
      'sourceInput', 'sourceDrop', 'sourceStatus', 'sourceList', 'processButton',
      'downloadNormalButton', 'downloadProjectedButton', 'resetButton', 'progressPanel',
      'progressText', 'processSteps', 'messageArea', 'dashboard', 'companyFilter',
      'companyFilterLabel', 'summaryTableBody', 'validationList', 'companyCards',
      'cutDateLabel', 'dateChipLabel', 'fileCountBadge', 'templateRepoStatus',
      'runtimeStatus', 'lastOutputLabel', 'horizonDays', 'thresholdDays',
      'reclassifyToggle', 'baseDateOutput', 'targetDateOutput', 'projectionStatus',
      'dashboardScenarioTitle', 'dashboardScenarioDescription', 'scenarioDashboard',
      'comparisonDashboard', 'comparisonKpis', 'comparisonTableBody', 'projectionAuditList'
    ].forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    els.sourceInput.addEventListener('change', (event) => addSourceFiles([...event.target.files]));
    els.processButton.addEventListener('click', processFiles);
    els.downloadNormalButton.addEventListener('click', () => downloadBlob(state.normalOutputBlob, state.normalOutputName));
    els.downloadProjectedButton.addEventListener('click', () => downloadBlob(state.projectedOutputBlob, state.projectedOutputName));
    els.resetButton.addEventListener('click', resetSession);
    els.companyFilter.addEventListener('change', (event) => {
      state.selectedCompany = event.target.value;
      renderDashboard();
    });
    document.querySelectorAll('.scenario-tab').forEach((button) => {
      button.addEventListener('click', () => {
        state.selectedScenario = button.dataset.scenario;
        document.querySelectorAll('.scenario-tab').forEach((item) => item.classList.toggle('active', item === button));
        renderDashboard();
      });
    });
    [els.horizonDays, els.thresholdDays, els.reclassifyToggle].forEach((element) => {
      element.addEventListener('change', () => {
        invalidateOutputs();
        updateProjectionOutputs();
      });
    });
    setupDropZone(els.sourceDrop, addSourceFiles);
  }

  function checkRuntime() {
    const missing = [];
    if (!window.JSZip) missing.push('JSZip');
    if (!P) missing.push('motor de reglas');
    if (!R) missing.push('motor de proyección');
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
      els.templateRepoStatus.textContent = 'Plantilla no disponible';
      showMessage('error', `No se pudo validar la plantilla integrada: ${error.message || error}`);
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
    invalidateOutputs();
    updateFilesUI();
    renderEmptyDashboard();
  }

  function fileKey(file) {
    return `${file.name}|${file.size}|${file.lastModified}`;
  }

  function removeSourceFile(index) {
    state.sourceFiles.splice(index, 1);
    invalidateOutputs();
    updateFilesUI();
    renderEmptyDashboard();
  }

  function invalidateOutputs() {
    state.normalAnalyses = [];
    state.projectedAnalyses = [];
    state.normalOutputBlob = null;
    state.projectedOutputBlob = null;
    state.normalOutputName = '';
    state.projectedOutputName = '';
    state.baseDate = null;
    state.targetDate = null;
    state.projectionParams = null;
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
        <div class="file-copy"><strong>${escapeHtml(file.name)}</strong><span>${formatBytes(file.size)}</span></div>
        <button type="button" class="icon-button" aria-label="Quitar archivo">×</button>`;
      item.querySelector('button').addEventListener('click', () => removeSourceFile(index));
      els.sourceList.appendChild(item);
    });

    const runtimeReady = Boolean(window.JSZip && P && R && X);
    els.processButton.disabled = !(state.templateReady && state.sourceFiles.length && runtimeReady);
    els.downloadNormalButton.disabled = !state.normalOutputBlob;
    els.downloadProjectedButton.disabled = !state.projectedOutputBlob;
    const outputs = [state.normalOutputName, state.projectedOutputName].filter(Boolean);
    els.lastOutputLabel.textContent = outputs.length ? `${outputs.length} archivos listos` : 'Aún no se han generado salidas';
  }

  function readProjectionParams() {
    const normalized = R.normalizeParameters({
      horizonDays: els.horizonDays.value,
      thresholdDays: els.thresholdDays.value,
      reclassify: els.reclassifyToggle.checked
    });
    els.horizonDays.value = String(normalized.horizonDays);
    els.thresholdDays.value = String(normalized.thresholdDays);
    return normalized;
  }

  function updateProjectionOutputs() {
    const params = readProjectionParams();
    els.baseDateOutput.textContent = state.baseDate ? formatDate(state.baseDate) : 'Se detectará desde los archivos';
    const target = state.baseDate ? R.addDays(state.baseDate, params.horizonDays) : null;
    els.targetDateOutput.textContent = target ? formatDate(target) : 'Se calculará automáticamente';
    els.projectionStatus.textContent = state.projectedAnalyses.length
      ? `${formatInteger(state.projectedAnalyses.reduce((sum, item) => sum + (item.projection?.reclassifiedOperations || 0), 0))} operaciones reclasificadas`
      : `Horizonte ${params.horizonDays} días · Umbral > ${params.thresholdDays}`;
  }

  async function processFiles() {
    if (!state.templateReady || !state.sourceFiles.length) return;
    setBusy(true);
    clearMessage();
    invalidateOutputs();
    state.selectedCompany = 'ALL';

    try {
      updateProgress(8, 'Validando archivos seleccionados');
      enforceFileLimits();

      const normalAnalyses = [];
      for (let index = 0; index < state.sourceFiles.length; index += 1) {
        const file = state.sourceFiles[index];
        updateProgress(16 + Math.round(((index + 1) / state.sourceFiles.length) * 22), `Calculando escenario Normal: ${file.name}`);
        await yieldToBrowser();
        const values = await readSourceMatrix(file);
        const analysis = P.analyzeValues(values, file.name);
        analysis.sourceValues = values;
        analysis.metrics.scenario = 'NORMAL';
        analysis.metrics.baseDate = new Date(analysis.metrics.cutDate.getTime());
        normalAnalyses.push(analysis);
      }

      P.validateDateGroup(normalAnalyses);
      normalAnalyses.sort((a, b) => companyOrder(a.company.code) - companyOrder(b.company.code));
      const baseDate = new Date(normalAnalyses[0].metrics.cutDate.getTime());
      const params = readProjectionParams();
      const targetDate = R.addDays(baseDate, params.horizonDays);

      updateProgress(48, 'Aplicando la lógica parametrizada de proyección');
      const projectedAnalyses = normalAnalyses.map((analysis) => R.projectAnalysis(analysis, params));
      projectedAnalyses.sort((a, b) => companyOrder(a.company.code) - companyOrder(b.company.code));

      updateProgress(70, 'Construyendo Excel del escenario Normal');
      const normalGenerated = await withTimeout(
        X.build(state.templateBuffer, normalAnalyses, {
          templateRows: P.TEMPLATE_ROWS,
          requiredSheets: P.COMPANY_RULES.map((item) => item.reportSheet),
          periodLabel: P.periodLabel,
          control: { scenario: 'NORMAL', baseDate, targetDate: baseDate, projection: { horizonDays: 0, thresholdDays: null } }
        }),
        90000,
        'La construcción del Excel Normal superó 90 segundos.'
      );

      updateProgress(88, 'Construyendo Excel del escenario Proyectado');
      const projectedGenerated = await withTimeout(
        X.build(state.templateBuffer, projectedAnalyses, {
          templateRows: P.TEMPLATE_ROWS,
          requiredSheets: P.COMPANY_RULES.map((item) => item.reportSheet),
          periodLabel: P.periodLabel,
          control: { scenario: 'PROYECTADO', baseDate, targetDate, projection: params }
        }),
        90000,
        'La construcción del Excel Proyectado superó 90 segundos.'
      );

      updateProgress(96, 'Verificando integridad de ambos archivos');
      const normalBuffer = toStandaloneArrayBuffer(normalGenerated);
      const projectedBuffer = toStandaloneArrayBuffer(projectedGenerated);
      const requiredSheets = (analyses) => [
        'CONTROL_EJECUCION',
        ...analyses.flatMap((item) => [item.company.reportSheet, `ORIGEN_${item.company.code}`])
      ];
      await validateXlsxPackage(normalBuffer, requiredSheets(normalAnalyses));
      await validateXlsxPackage(projectedBuffer, requiredSheets(projectedAnalyses));

      state.normalOutputBlob = xlsxBlob(normalBuffer);
      state.projectedOutputBlob = xlsxBlob(projectedBuffer);
      state.normalOutputName = `COMITE_Normal_${P.fileDateKey(baseDate)}_FINAL.xlsx`;
      state.projectedOutputName = `COMITE_Proyectado_${P.fileDateKey(targetDate)}_BASE_${P.fileDateKey(baseDate)}_FINAL.xlsx`;
      state.normalAnalyses = normalAnalyses;
      state.projectedAnalyses = projectedAnalyses;
      state.baseDate = baseDate;
      state.targetDate = targetDate;
      state.projectionParams = params;

      updateProgress(100, 'Escenarios validados y listos para descargar');
      updateProjectionOutputs();
      renderDashboard();
      const reclassified = projectedAnalyses.reduce((sum, item) => sum + item.projection.reclassifiedOperations, 0);
      showMessage('success', `Proceso completado. Se generaron los escenarios Normal y Proyectado; ${formatInteger(reclassified)} operaciones fueron reclasificadas en la proyección.`);
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
    if (!state.normalAnalyses.length || !state.projectedAnalyses.length) return renderEmptyDashboard();
    els.dashboard.classList.remove('empty');
    renderCompanyFilter();

    const isComparison = state.selectedScenario === 'COMPARATIVO';
    els.scenarioDashboard.hidden = isComparison;
    els.comparisonDashboard.hidden = !isComparison;
    els.companyFilterLabel.hidden = false;

    if (isComparison) {
      els.dashboardScenarioTitle.textContent = 'Comparativo Normal vs. Proyectado';
      els.dashboardScenarioDescription.textContent = 'Variaciones monetarias expresadas en miles, redondeadas y sin decimales.';
      els.dateChipLabel.textContent = 'Horizonte';
      els.cutDateLabel.textContent = `${formatShortDate(state.baseDate)} → ${formatShortDate(state.targetDate)}`;
      renderComparison();
      return;
    }

    const projected = state.selectedScenario === 'PROYECTADO';
    const source = projected ? state.projectedAnalyses : state.normalAnalyses;
    const filtered = state.selectedCompany === 'ALL' ? source : source.filter((item) => item.company.code === state.selectedCompany);
    els.dashboardScenarioTitle.textContent = projected ? 'Escenario Proyectado' : 'Escenario Normal';
    els.dashboardScenarioDescription.textContent = projected
      ? `Proyección a ${state.projectionParams.horizonDays} días con umbral > ${state.projectionParams.thresholdDays}. Valores en miles sin decimales.`
      : 'Resultados obtenidos directamente de los archivos originales. Valores en miles sin decimales.';
    els.dateChipLabel.textContent = projected ? 'Fecha objetivo' : 'Fecha de corte';
    els.cutDateLabel.textContent = formatDate(projected ? state.targetDate : state.baseDate);
    renderKpis(filtered, projected);
    renderCompanyCards(filtered, projected);
    renderSummaryTable(filtered);
    renderValidations(filtered, projected);
  }

  function renderCompanyFilter() {
    const codes = state.normalAnalyses.map((item) => item.company.code);
    els.companyFilter.innerHTML = '<option value="ALL">Todas las empresas</option>' + codes.map((code) => `<option value="${code}">${code}</option>`).join('');
    if (!codes.includes(state.selectedCompany)) state.selectedCompany = 'ALL';
    els.companyFilter.value = state.selectedCompany;
  }

  function aggregate(analyses) {
    return analyses.reduce((accumulator, { metrics }) => {
      accumulator.total += metrics.total.value;
      accumulator.operations += metrics.total.operations;
      accumulator.overdue += metrics.overdue.value;
      accumulator.overdueOps += metrics.overdue.operations;
      accumulator.noDevenga += metrics.noDevenga.value;
      accumulator.noDevengaOps += metrics.noDevenga.operations;
      accumulator.noDevenga6090 += metrics.noDevenga6090.value;
      accumulator.noDevengaOver90 += metrics.noDevengaOver90.value;
      accumulator.chargedOff += metrics.chargedOff.value || 0;
      accumulator.chargedOffOperations += metrics.chargedOff.operations || 0;
      accumulator.reclassified += metrics.projection ? metrics.projection.reclassifiedOperations : 0;
      return accumulator;
    }, { total: 0, operations: 0, overdue: 0, overdueOps: 0, noDevenga: 0, noDevengaOps: 0, noDevenga6090: 0, noDevengaOver90: 0, chargedOff: 0, chargedOffOperations: 0, reclassified: 0 });
  }

  function renderKpis(analyses, projected) {
    const totals = aggregate(analyses);
    const cards = [
      { icon: 'Σ', label: 'Total cartera', value: P.formatMoneyThousands(totals.total), meta: `${formatInteger(totals.operations)} operaciones`, tone: 'navy' },
      { icon: 'V', label: 'Cartera vencida', value: P.formatMoneyThousands(totals.overdue), meta: P.formatPercent(safeDivide(totals.overdue, totals.total)), tone: P.semaphore(safeDivide(totals.overdue, totals.total), P.THRESHOLDS.overdueTotal) },
      { icon: 'ND', label: 'No Devenga', value: P.formatMoneyThousands(totals.noDevenga), meta: P.formatPercent(safeDivide(totals.noDevenga, totals.total)), tone: P.semaphore(safeDivide(totals.noDevenga, totals.total), P.THRESHOLDS.noDevengaTotal) },
      { icon: '+90', label: 'No Devenga +90', value: P.formatMoneyThousands(totals.noDevengaOver90), meta: P.formatPercent(safeDivide(totals.noDevengaOver90, totals.total)), tone: P.semaphore(safeDivide(totals.noDevengaOver90, totals.total), P.THRESHOLDS.noDevengaOver90Total) },
      projected
        ? { icon: '↗', label: 'Reclasificadas', value: formatInteger(totals.reclassified), meta: 'operaciones por proyección', tone: 'projected' }
        : { icon: 'C', label: 'Cartera castigada', value: totals.chargedOff ? P.formatMoneyThousands(totals.chargedOff) : '—', meta: totals.chargedOffOperations ? `${formatInteger(totals.chargedOffOperations)} operaciones` : 'Sin información', tone: totals.chargedOff ? 'slate' : 'muted' }
    ];
    document.getElementById('kpiGrid').innerHTML = cards.map((card) => `
      <article class="kpi-card tone-${card.tone}">
        <div class="kpi-icon">${card.icon}</div>
        <div class="kpi-copy"><span>${card.label}</span><strong>${card.value}</strong><small>${card.meta}</small></div>
      </article>`).join('');
  }

  function renderCompanyCards(analyses, projected) {
    els.companyCards.innerHTML = analyses.map(({ metrics, projection }) => {
      const ndRatio = safeDivide(metrics.noDevenga.value, metrics.total.value) || 0;
      const overdueRatio = safeDivide(metrics.overdue.value, metrics.total.value) || 0;
      const statusTone = metrics.status === 'OK' ? 'ok' : 'review';
      const extra = projected ? `<div><dt>Reclasificadas</dt><dd>${formatInteger(projection?.reclassifiedOperations || 0)}</dd></div>` : '';
      return `
        <article class="company-card ${projected ? 'projected-card' : ''}">
          <header><div><span class="company-kicker">EMPRESA</span><h4>${metrics.company}</h4></div><span class="status-pill ${statusTone}">${metrics.status}</span></header>
          <div class="company-total"><span>Total cartera · miles</span><strong>${P.formatMoneyThousands(metrics.total.value)}</strong><small>${formatInteger(metrics.total.operations)} operaciones activas</small></div>
          <div class="gauge-grid">${gaugeMarkup('No Devenga', ndRatio, P.semaphore(ndRatio, P.THRESHOLDS.noDevengaTotal))}${gaugeMarkup('Vencida', overdueRatio, P.semaphore(overdueRatio, P.THRESHOLDS.overdueTotal))}</div>
          <dl class="metric-list"><div><dt>ND 60–90</dt><dd>${P.formatMoneyThousands(metrics.noDevenga6090.value)}</dd></div><div><dt>ND +90</dt><dd>${P.formatMoneyThousands(metrics.noDevengaOver90.value)}</dd></div>${extra}<div><dt>Castigada</dt><dd>${P.formatMoneyThousands(metrics.chargedOff.value)}</dd></div></dl>
        </article>`;
    }).join('');
  }

  function gaugeMarkup(label, ratio, tone) {
    const percent = Math.max(0, Math.min(100, Math.round(ratio * 100)));
    return `<div class="gauge-item"><div class="gauge gauge-${tone}" style="--gauge-value:${percent}"><div><strong>${percent}%</strong><span>${label}</span></div></div></div>`;
  }

  function renderSummaryTable(analyses) {
    els.summaryTableBody.innerHTML = analyses.map(({ metrics }) => `
      <tr><td><strong>${metrics.company}</strong></td><td class="number">${P.formatMoneyThousands(metrics.total.value)}</td><td class="number">${formatInteger(metrics.total.operations)}</td><td class="number">${P.formatMoneyThousands(metrics.overdue.value)}</td><td class="number">${P.formatMoneyThousands(metrics.noDevenga.value)}</td><td class="number">${P.formatMoneyThousands(metrics.noDevenga6090.value)}</td><td class="number">${P.formatMoneyThousands(metrics.noDevengaOver90.value)}</td><td class="number">${P.formatMoneyThousands(metrics.chargedOff.value)}</td><td><span class="status-pill ${metrics.status === 'OK' ? 'ok' : 'review'}">${metrics.status}</span></td></tr>`).join('');
  }

  function renderValidations(analyses, projected) {
    const items = [];
    analyses.forEach(({ metrics, projection }) => {
      items.push({ type: 'ok', text: `${metrics.company}: ${formatInteger(metrics.detailRows)} filas de detalle validadas; totales y subtotales excluidos.` });
      items.push({ type: metrics.duplicatesExcluded ? 'warning' : 'ok', text: `${metrics.company}: ${formatInteger(metrics.duplicatesExcluded)} duplicados exactos excluidos.` });
      items.push({ type: 'ok', text: `${metrics.company}: No Devenga se calculó únicamente desde sus columnas específicas.` });
      if (projected) {
        items.push({ type: 'info', text: `${metrics.company}: ${formatInteger(projection.reclassifiedOperations)} operaciones migraron al superar > ${projection.thresholdDays} días.` });
      }
      items.push({ type: metrics.chargedOff.value == null ? 'info' : 'ok', text: metrics.chargedOff.value == null
        ? `${metrics.company}: Cartera Castigada permanece en blanco.`
        : `${metrics.company}: ${formatInteger(metrics.chargedOff.operations)} operaciones castigadas separadas de la cartera activa.` });
      metrics.warnings.forEach((warning) => items.push({ type: 'warning', text: `${metrics.company}: ${warning}` }));
    });
    els.validationList.innerHTML = items.map((item) => `<li class="validation-${item.type}"><span></span>${escapeHtml(item.text)}</li>`).join('');
  }

  function renderComparison() {
    const pairs = comparisonPairs();
    const filtered = state.selectedCompany === 'ALL' ? pairs : pairs.filter((item) => item.company === state.selectedCompany);
    const normalTotals = aggregate(filtered.map((item) => item.normal));
    const projectedTotals = aggregate(filtered.map((item) => item.projected));
    const deltaOverdue = projectedTotals.overdue - normalTotals.overdue;
    const deltaNd = projectedTotals.noDevenga - normalTotals.noDevenga;
    const delta6090 = projectedTotals.noDevenga6090 - normalTotals.noDevenga6090;
    const deltaOver90 = projectedTotals.noDevengaOver90 - normalTotals.noDevengaOver90;

    els.comparisonKpis.innerHTML = [
      { label: 'Total cartera', value: P.formatMoneyThousands(projectedTotals.total), meta: `Δ ${formatSignedThousands(projectedTotals.total - normalTotals.total)}`, tone: 'navy' },
      { label: 'Aumento Vencida', value: formatSignedThousands(deltaOverdue), meta: 'miles frente al escenario Normal', tone: 'red' },
      { label: 'Aumento No Devenga', value: formatSignedThousands(deltaNd), meta: 'miles frente al escenario Normal', tone: 'yellow' },
      { label: 'Reclasificadas', value: formatInteger(projectedTotals.reclassified), meta: 'operaciones', tone: 'projected' }
    ].map((card) => `<article class="comparison-kpi tone-${card.tone}"><span>${card.label}</span><strong>${card.value}</strong><small>${card.meta}</small></article>`).join('');

    els.comparisonTableBody.innerHTML = filtered.map((item) => {
      const n = item.normal.metrics;
      const p = item.projected.metrics;
      return `<tr><td><strong>${item.company}</strong></td><td class="number">${P.formatMoneyThousands(n.total.value)}</td><td class="number">${P.formatMoneyThousands(p.total.value)}</td><td class="number delta-positive">${formatSignedThousands(p.overdue.value - n.overdue.value)}</td><td class="number delta-positive">${formatSignedThousands(p.noDevenga.value - n.noDevenga.value)}</td><td class="number">${formatSignedThousands(p.noDevenga6090.value - n.noDevenga6090.value)}</td><td class="number">${formatSignedThousands(p.noDevengaOver90.value - n.noDevengaOver90.value)}</td><td class="number">${formatInteger(item.projected.projection.reclassifiedOperations)}</td></tr>`;
    }).join('');

    const params = state.projectionParams;
    const totalReclassified = filtered.reduce((sum, item) => sum + item.projected.projection.reclassifiedOperations, 0);
    els.projectionAuditList.innerHTML = [
      `Fecha base: ${formatDate(state.baseDate)}.`,
      `Fecha objetivo: ${formatDate(state.targetDate)}.`,
      `Horizonte aplicado: ${params.horizonDays} días.`,
      `Umbral: días proyectados mayores a ${params.thresholdDays}.`,
      `Reclasificación: ${params.reclassify ? 'activada' : 'desactivada'}.`,
      `${formatInteger(totalReclassified)} operaciones reclasificadas en la selección actual.`,
      `El Total Cartera permanece sin variación; la proyección modifica la distribución entre Por Vencer, No Devenga y Vencida.`
    ].map((text, index) => `<li class="validation-${index < 5 ? 'info' : 'ok'}"><span></span>${escapeHtml(text)}</li>`).join('');
  }

  function comparisonPairs() {
    const projectedMap = new Map(state.projectedAnalyses.map((item) => [item.company.code, item]));
    return state.normalAnalyses.map((normal) => ({ company: normal.company.code, normal, projected: projectedMap.get(normal.company.code) })).filter((item) => item.projected);
  }

  function renderEmptyDashboard() {
    els.dashboard.classList.add('empty');
    els.cutDateLabel.textContent = '—';
    document.getElementById('kpiGrid').innerHTML = '';
    els.companyCards.innerHTML = '<div class="empty-state"><strong>Sin resultados</strong><span>Carga los reportes originales y ejecuta ambos escenarios.</span></div>';
    els.summaryTableBody.innerHTML = '<tr><td colspan="9" class="placeholder-cell">Sin resultados procesados.</td></tr>';
    els.validationList.innerHTML = '<li class="validation-info"><span></span>La validación aparecerá al finalizar el procesamiento.</li>';
    els.comparisonKpis.innerHTML = '';
    els.comparisonTableBody.innerHTML = '';
    els.projectionAuditList.innerHTML = '';
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
    const runtimeReady = Boolean(window.JSZip && P && R && X);
    els.processButton.disabled = isBusy || !(state.templateReady && state.sourceFiles.length && runtimeReady);
    els.resetButton.disabled = isBusy;
    els.downloadNormalButton.disabled = isBusy || !state.normalOutputBlob;
    els.downloadProjectedButton.disabled = isBusy || !state.projectedOutputBlob;
  }

  function downloadBlob(blob, name) {
    if (!blob || !name) return;
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = name;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 3000);
  }

  function resetSession() {
    state.sourceFiles = [];
    invalidateOutputs();
    state.selectedCompany = 'ALL';
    state.selectedScenario = 'NORMAL';
    els.sourceInput.value = '';
    document.querySelectorAll('.scenario-tab').forEach((item) => item.classList.toggle('active', item.dataset.scenario === 'NORMAL'));
    clearMessage();
    updateFilesUI();
    updateProjectionOutputs();
    renderEmptyDashboard();
    updateProgress(0, 'Listo para procesar');
  }

  function xlsxBlob(buffer) {
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function withTimeout(promise, milliseconds, message) {
    return Promise.race([promise, new Promise((_, reject) => window.setTimeout(() => reject(new Error(message)), milliseconds))]);
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

  function formatShortDate(date) {
    return date.toLocaleDateString('es-EC', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  function formatSignedThousands(value) {
    const rounded = Math.round(Number(value || 0) / 1000);
    const formatted = Math.abs(rounded).toLocaleString('es-EC', { maximumFractionDigits: 0 });
    if (rounded > 0) return `+${formatted}`;
    if (rounded < 0) return `−${formatted}`;
    return '0';
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
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#039;');
  }
})();
