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
    projectionParams: null,
    logoDataUrl: null
  };

  const els = {};
  const steps = [
    { at: 8, label: 'Preparar' },
    { at: 48, label: 'Analizar' },
    { at: 88, label: 'Generar' },
    { at: 100, label: 'Finalizar' }
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
      'downloadNormalButton', 'downloadProjectedButton', 'downloadNormalEnterpriseButton', 'downloadProjectedEnterpriseButton', 'resetButton', 'progressPanel',
      'progressText', 'processSteps', 'messageArea', 'dashboard', 'companyFilter',
      'companyFilterLabel', 'summaryTableBody', 'companyCards',
      'cutDateLabel', 'dateChipLabel', 'fileCountBadge', 'templateRepoStatus',
      'runtimeStatus', 'lastOutputLabel', 'horizonDays', 'thresholdDays',
      'reclassifyToggle', 'baseDateOutput', 'targetDateOutput', 'projectionStatus',
      'dashboardScenarioTitle', 'dashboardScenarioDescription', 'scenarioDashboard',
      'comparisonDashboard', 'comparisonKpis', 'comparisonTableBody'
    ].forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    els.sourceInput.addEventListener('change', (event) => addSourceFiles([...event.target.files]));
    els.processButton.addEventListener('click', processFiles);
    els.downloadNormalButton.addEventListener('click', () => downloadBlob(state.normalOutputBlob, state.normalOutputName));
    els.downloadProjectedButton.addEventListener('click', () => downloadBlob(state.projectedOutputBlob, state.projectedOutputName));
    els.downloadNormalEnterpriseButton.addEventListener('click', () => downloadEnterpriseDashboard('NORMAL'));
    els.downloadProjectedEnterpriseButton.addEventListener('click', () => downloadEnterpriseDashboard('PROYECTADO'));
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
      els.runtimeStatus.textContent = 'Requiere actualización';
      console.error('Componentes no disponibles:', missing);
      showMessage('error', 'La aplicación no pudo iniciar correctamente. Actualiza la página con Ctrl+F5.');
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
      console.error('Error al validar la plantilla integrada:', error);
      showMessage('error', 'No fue posible preparar la plantilla integrada. Actualiza la página e inténtalo nuevamente.');
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
    els.downloadNormalEnterpriseButton.disabled = !state.normalAnalyses.length;
    els.downloadProjectedEnterpriseButton.disabled = !state.projectedAnalyses.length;
    const outputs = [state.normalOutputName, state.projectedOutputName].filter(Boolean);
    els.lastOutputLabel.textContent = outputs.length ? '4 archivos disponibles' : 'Aún no se han generado salidas';
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
      updateProgress(8, 'Preparando los archivos seleccionados');
      enforceFileLimits();

      const normalAnalyses = [];
      for (let index = 0; index < state.sourceFiles.length; index += 1) {
        const file = state.sourceFiles[index];
        updateProgress(16 + Math.round(((index + 1) / state.sourceFiles.length) * 22), `Analizando ${index + 1} de ${state.sourceFiles.length} reportes`);
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

      updateProgress(48, 'Generando los escenarios Normal y Proyectado');
      const projectedAnalyses = normalAnalyses.map((analysis) => R.projectAnalysis(analysis, params));
      projectedAnalyses.sort((a, b) => companyOrder(a.company.code) - companyOrder(b.company.code));

      updateProgress(70, 'Preparando el archivo del escenario Normal');
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

      updateProgress(88, 'Preparando el archivo del escenario Proyectado');
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

      updateProgress(96, 'Finalizando los archivos de descarga');
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

      updateProgress(100, 'Resultados listos para descargar');
      updateProjectionOutputs();
      renderDashboard();
      const reclassified = projectedAnalyses.reduce((sum, item) => sum + item.projection.reclassifiedOperations, 0);
      showMessage('success', `Proceso completado. Se generaron los escenarios Normal y Proyectado; ${formatInteger(reclassified)} operaciones fueron reclasificadas en la proyección.`);
    } catch (error) {
      console.error(error);
      showMessage('error', userFriendlyError(error));
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
      console.error(`No se pudo leer ${file.name}:`, error);
      throw new Error(`No se pudo leer el archivo ${file.name}. Verifica que sea un reporte Excel válido.`);
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
    els.comparisonKpis.innerHTML = '';
    els.comparisonTableBody.innerHTML = '';
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
    els.progressPanel.style.display = isBusy ? 'flex' : 'none';
    els.progressPanel.setAttribute('aria-hidden', String(!isBusy));
    els.processButton.setAttribute('aria-busy', String(isBusy));
    const runtimeReady = Boolean(window.JSZip && P && R && X);
    els.processButton.disabled = isBusy || !(state.templateReady && state.sourceFiles.length && runtimeReady);
    els.resetButton.disabled = isBusy;
    els.downloadNormalButton.disabled = isBusy || !state.normalOutputBlob;
    els.downloadProjectedButton.disabled = isBusy || !state.projectedOutputBlob;
    els.downloadNormalEnterpriseButton.disabled = isBusy || !state.normalAnalyses.length;
    els.downloadProjectedEnterpriseButton.disabled = isBusy || !state.projectedAnalyses.length;
    if (!isBusy) renderProcessSteps(0);
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

  async function downloadEnterpriseDashboard(scenario) {
    const projected = scenario === 'PROYECTADO';
    const analyses = projected ? state.projectedAnalyses : state.normalAnalyses;
    if (!analyses.length) return;

    try {
      const logoDataUrl = await getEmbeddedLogoDataUrl();
      const reportDate = projected ? state.targetDate : state.baseDate;
      const payload = {
        scenario,
        scenarioLabel: projected ? 'Escenario Proyectado' : 'Escenario Normal',
        subtitle: projected
          ? `Proyección a ${state.projectionParams.horizonDays} días · Umbral > ${state.projectionParams.thresholdDays}`
          : 'Resultados obtenidos directamente de los archivos originales',
        dateLabel: projected ? 'Fecha objetivo' : 'Fecha de corte',
        dateValue: formatDate(reportDate),
        baseDate: formatDate(state.baseDate),
        targetDate: formatDate(state.targetDate),
        horizonDays: projected ? state.projectionParams.horizonDays : 0,
        thresholdDays: projected ? state.projectionParams.thresholdDays : null,
        logoDataUrl,
        generatedAt: new Date().toLocaleString('es-EC'),
        rows: analyses.map(({ metrics, projection }) => ({
          company: metrics.company,
          total: metrics.total.value,
          operations: metrics.total.operations,
          overdue: metrics.overdue.value,
          overdueOperations: metrics.overdue.operations,
          noDevenga: metrics.noDevenga.value,
          noDevengaOperations: metrics.noDevenga.operations,
          noDevenga6090: metrics.noDevenga6090.value,
          noDevengaOver90: metrics.noDevengaOver90.value,
          chargedOff: metrics.chargedOff.value || 0,
          chargedOffOperations: metrics.chargedOff.operations || 0,
          reclassified: projection?.reclassifiedOperations || 0,
          status: metrics.status || 'OK'
        }))
      };

      const html = buildEnterpriseDashboardHtml(payload);
      const keyDate = P.fileDateKey(reportDate);
      const name = `Dashboard_Enterprise_${projected ? 'Proyectado' : 'Normal'}_${keyDate}.html`;
      downloadBlob(new Blob(['\ufeff', html], { type: 'text/html;charset=utf-8' }), name);
      showMessage('success', `Dashboard Enterprise ${projected ? 'Proyectado' : 'Normal'} generado correctamente.`);
    } catch (error) {
      console.error('No se pudo generar el Dashboard Enterprise:', error);
      showMessage('error', 'No fue posible generar el Dashboard Enterprise. Vuelve a intentarlo.');
    }
  }

  async function getEmbeddedLogoDataUrl() {
    if (state.logoDataUrl) return state.logoDataUrl;
    try {
      const response = await fetch('assets/logo_CTH.png', { cache: 'force-cache' });
      if (!response.ok) return '';
      const bytes = new Uint8Array(await response.arrayBuffer());
      let binary = '';
      const chunk = 0x8000;
      for (let index = 0; index < bytes.length; index += chunk) {
        binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
      }
      state.logoDataUrl = `data:image/png;base64,${btoa(binary)}`;
      return state.logoDataUrl;
    } catch (error) {
      console.warn('No se pudo incrustar el logo en el dashboard exportado:', error);
      return '';
    }
  }

  function buildEnterpriseDashboardHtml(payload) {
    const serialized = JSON.stringify(payload).replace(/</g, '\\u003c');
    return `<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CTH | ${escapeHtml(payload.scenarioLabel)} | Enterprise Dashboard</title>
<style>
:root{--navy:#0f2f49;--navy2:#184b70;--green:#18835f;--green2:#57ad8c;--red:#c13e49;--amber:#c68a22;--ink:#182b3a;--muted:#667988;--line:#dbe5ec;--bg:#eef3f6;--white:#fff;--shadow:0 18px 45px rgba(15,47,73,.10)}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font-family:Inter,Segoe UI,Arial,sans-serif}.shell{width:min(1440px,calc(100% - 42px));margin:0 auto}.top{background:linear-gradient(120deg,#0b263b,#154b70);color:white}.topin{min-height:116px;display:flex;align-items:center;justify-content:space-between;gap:28px}.brand{display:flex;align-items:center;gap:22px}.logo{width:148px;height:66px;display:grid;place-items:center;padding:8px 12px;border-radius:15px;background:white}.logo img{max-width:100%;max-height:100%;object-fit:contain}.brand h1{margin:4px 0 0;font-size:27px;letter-spacing:-.03em}.eyebrow{font-size:10px;font-weight:800;letter-spacing:.16em;color:#9dc7db}.brand p{margin:7px 0 0;color:#d1e2ec;font-size:13px}.meta{text-align:right}.meta span,.meta strong{display:block}.meta span{font-size:10px;color:#a9c8da;text-transform:uppercase;letter-spacing:.1em}.meta strong{margin-top:5px;font-size:15px}.toolbar{margin-top:-18px;padding:16px 18px;display:flex;align-items:center;justify-content:space-between;gap:18px;border:1px solid rgba(15,47,73,.08);border-radius:16px;background:white;box-shadow:var(--shadow)}.toolbar-left,.toolbar-right{display:flex;align-items:center;gap:12px}.badge{padding:7px 10px;border-radius:999px;background:#e8f4ef;color:var(--green);font-size:10px;font-weight:850;letter-spacing:.08em}.datebox span,.datebox strong{display:block}.datebox span{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em}.datebox strong{margin-top:2px;font-size:13px}.toolbar select,.toolbar button{height:38px;border:1px solid var(--line);border-radius:10px;background:white;color:var(--ink);font:inherit}.toolbar select{padding:0 34px 0 12px;font-size:12px}.toolbar button{padding:0 14px;font-size:11px;font-weight:750;cursor:pointer}.toolbar button:hover{background:#f4f8fa}.content{padding:28px 0 36px}.section-title{margin:0 0 14px}.section-title span{display:block;color:var(--green);font-size:9px;font-weight:850;letter-spacing:.14em}.section-title h2{margin:5px 0 0;font-size:20px}.kpis{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px}.kpi{position:relative;overflow:hidden;min-height:126px;padding:17px;border:1px solid var(--line);border-radius:16px;background:white;box-shadow:0 9px 24px rgba(15,47,73,.06)}.kpi:after{content:"";position:absolute;right:-24px;bottom:-32px;width:90px;height:90px;border-radius:50%;background:rgba(24,131,95,.08)}.kpi span,.kpi strong,.kpi small{display:block}.kpi span{font-size:10px;color:var(--muted);font-weight:750}.kpi strong{margin-top:12px;color:var(--navy);font-size:27px;letter-spacing:-.04em}.kpi small{margin-top:8px;color:var(--muted);font-size:10px}.grid2{display:grid;grid-template-columns:minmax(0,1.35fr) minmax(320px,.65fr);gap:16px;margin-top:18px}.panel{padding:20px;border:1px solid var(--line);border-radius:17px;background:white;box-shadow:0 10px 26px rgba(15,47,73,.055)}.panel-head{display:flex;align-items:center;justify-content:space-between;gap:14px;margin-bottom:16px}.panel-head h3{margin:0;font-size:15px}.panel-head p{margin:4px 0 0;color:var(--muted);font-size:10px}.metric-switch{display:flex;gap:5px;padding:4px;border-radius:10px;background:#eef3f6}.metric-switch button{border:0;border-radius:7px;padding:7px 9px;background:transparent;color:var(--muted);font-size:9px;font-weight:800;cursor:pointer}.metric-switch button.active{background:white;color:var(--navy);box-shadow:0 2px 7px rgba(15,47,73,.1)}.bars{display:grid;gap:14px}.bar-row{display:grid;grid-template-columns:44px 1fr 80px;align-items:center;gap:10px}.bar-row strong{font-size:11px}.track{height:11px;overflow:hidden;border-radius:99px;background:#e9eff3}.fill{height:100%;min-width:2px;border-radius:99px;background:linear-gradient(90deg,var(--navy2),#4f8cac)}.bar-row[data-tone="risk"] .fill{background:linear-gradient(90deg,var(--amber),#e4b24f)}.bar-row[data-tone="critical"] .fill{background:linear-gradient(90deg,var(--red),#df7079)}.bar-value{text-align:right;font-size:11px;font-weight:800;color:var(--navy)}.rings{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.ring-card{text-align:center;padding:13px 8px;border:1px solid #e5edf2;border-radius:14px;background:#f9fbfc}.ring{--p:0;--c:var(--green);width:88px;height:88px;margin:0 auto;display:grid;place-items:center;border-radius:50%;background:conic-gradient(var(--c) calc(var(--p)*1%),#dfe8ee 0);position:relative}.ring:before{content:"";position:absolute;inset:10px;border-radius:50%;background:white}.ring strong{position:relative;font-size:18px;color:var(--navy)}.ring-card span{display:block;margin-top:9px;color:var(--muted);font-size:9px;font-weight:800}.insight{margin-top:14px;padding:14px;border-left:4px solid var(--green);border-radius:10px;background:#f0f8f4}.insight strong{display:block;font-size:11px;color:var(--navy)}.insight p{margin:5px 0 0;color:var(--muted);font-size:10px;line-height:1.55}.table-panel{margin-top:18px}.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:13px}table{width:100%;border-collapse:collapse;background:white;font-size:10px}th{padding:11px 12px;background:#f2f6f8;color:#4e6474;text-align:left;text-transform:uppercase;letter-spacing:.06em;font-size:8px}td{padding:11px 12px;border-top:1px solid #e6edf2}td.num{text-align:right;font-variant-numeric:tabular-nums}.pill{display:inline-flex;padding:4px 7px;border-radius:999px;background:#eaf6f0;color:var(--green);font-size:8px;font-weight:850}.footer{padding:18px 0 24px;border-top:1px solid #d8e2e9;color:var(--muted);font-size:9px}.footerin{display:flex;justify-content:space-between;gap:18px}.note{margin-top:15px;color:var(--muted);font-size:9px}.empty{padding:36px;text-align:center;color:var(--muted)}
@media(max-width:1050px){.kpis{grid-template-columns:repeat(3,1fr)}.grid2{grid-template-columns:1fr}.topin{align-items:flex-start;padding:22px 0}.meta{display:none}}@media(max-width:680px){.shell{width:min(100% - 20px,1440px)}.brand{align-items:flex-start}.logo{width:100px;height:52px}.brand h1{font-size:21px}.brand p{display:none}.toolbar,.toolbar-left,.toolbar-right{align-items:stretch;flex-direction:column}.kpis{grid-template-columns:1fr}.rings{grid-template-columns:1fr}.bar-row{grid-template-columns:38px 1fr 64px}.footerin{flex-direction:column}}
@media print{body{background:white}.toolbar button,.toolbar select,.metric-switch{display:none}.toolbar{margin-top:0;box-shadow:none}.panel,.kpi{box-shadow:none}.content{padding-top:16px}.shell{width:100%}}
</style>
</head>
<body>
<header class="top"><div class="shell topin"><div class="brand"><div class="logo">${payload.logoDataUrl ? `<img src="${payload.logoDataUrl}" alt="CTH">` : '<strong>CTH</strong>'}</div><div><span class="eyebrow">GESTIÓN Y ANÁLISIS FINANCIERO</span><h1>Enterprise Dashboard · ${escapeHtml(payload.scenarioLabel)}</h1><p>${escapeHtml(payload.subtitle)}</p></div></div><div class="meta"><span>Desarrolladora</span><strong>Lizbeth Sanipatín</strong></div></div></header>
<div class="shell toolbar"><div class="toolbar-left"><span class="badge">VALORES EN MILES</span><div class="datebox"><span>${escapeHtml(payload.dateLabel)}</span><strong>${escapeHtml(payload.dateValue)}</strong></div></div><div class="toolbar-right"><label class="datebox"><span>Empresa</span><select id="companyFilter"><option value="ALL">Todas las empresas</option></select></label><button id="printButton">Imprimir / Guardar PDF</button></div></div>
<main class="shell content"><div class="section-title"><span>RESUMEN EJECUTIVO</span><h2 id="scopeTitle">Consolidado de todas las empresas</h2></div><section class="kpis" id="kpis"></section><div class="grid2"><section class="panel"><div class="panel-head"><div><h3>Exposición por empresa</h3><p>Comparación visual en miles, sin decimales.</p></div><div class="metric-switch"><button class="active" data-metric="total">Total</button><button data-metric="noDevenga">No Devenga</button><button data-metric="overdue">Vencida</button></div></div><div class="bars" id="bars"></div></section><section class="panel"><div class="panel-head"><div><h3>Composición de riesgo</h3><p>Participación sobre la cartera total.</p></div></div><div class="rings" id="rings"></div><div class="insight" id="insight"></div></section></div><section class="panel table-panel"><div class="panel-head"><div><h3>Matriz ejecutiva</h3><p>Resultados oficiales utilizados para esta exportación.</p></div></div><div class="table-wrap"><table><thead><tr><th>Empresa</th><th>Total cartera</th><th>Operaciones</th><th>Vencida</th><th>No Devenga</th><th>ND 60–90</th><th>ND +90</th><th>Castigada</th>${payload.scenario === 'PROYECTADO' ? '<th>Reclasificadas</th>' : ''}<th>Estado</th></tr></thead><tbody id="tableBody"></tbody></table></div><p class="note">Esta presentación es exclusivamente visual. Los valores provienen del mismo motor validado que genera el Excel estándar y no modifican las reglas de negocio.</p></section></main>
<footer class="footer"><div class="shell footerin"><span>CTH · Comité de Cartera · ${escapeHtml(payload.scenarioLabel)}</span><span>Generado: ${escapeHtml(payload.generatedAt)} · Privacidad por diseño</span></div></footer>
<script>
const REPORT=${serialized};let metric='total';const money=v=>Math.round((Number(v)||0)/1000).toLocaleString('es-EC',{maximumFractionDigits:0});const integer=v=>(Number(v)||0).toLocaleString('es-EC',{maximumFractionDigits:0});const pct=(a,b)=>b?Math.round(a/b*100):0;const sum=(rows,key)=>rows.reduce((t,r)=>t+(Number(r[key])||0),0);const filter=document.getElementById('companyFilter');REPORT.rows.forEach(r=>{const o=document.createElement('option');o.value=r.company;o.textContent=r.company;filter.appendChild(o)});document.getElementById('printButton').onclick=()=>window.print();document.querySelectorAll('[data-metric]').forEach(b=>b.onclick=()=>{metric=b.dataset.metric;document.querySelectorAll('[data-metric]').forEach(x=>x.classList.toggle('active',x===b));render()});filter.onchange=render;
function current(){return filter.value==='ALL'?REPORT.rows:REPORT.rows.filter(r=>r.company===filter.value)}
function render(){const rows=current();document.getElementById('scopeTitle').textContent=filter.value==='ALL'?'Consolidado de todas las empresas':'Empresa '+filter.value;const total=sum(rows,'total'),ops=sum(rows,'operations'),over=sum(rows,'overdue'),nd=sum(rows,'noDevenga'),over90=sum(rows,'noDevengaOver90'),cast=sum(rows,'chargedOff'),recl=sum(rows,'reclassified');const cards=[['Total cartera',money(total),integer(ops)+' operaciones'],['Cartera vencida',money(over),pct(over,total)+'% del total'],['No Devenga',money(nd),pct(nd,total)+'% del total'],['No Devenga +90',money(over90),pct(over90,total)+'% del total'],REPORT.scenario==='PROYECTADO'?['Reclasificadas',integer(recl),'operaciones proyectadas']:['Cartera castigada',cast?money(cast):'—',cast?integer(sum(rows,'chargedOffOperations'))+' operaciones':'Sin información']];document.getElementById('kpis').innerHTML=cards.map(c=>'<article class="kpi"><span>'+c[0]+'</span><strong>'+c[1]+'</strong><small>'+c[2]+'</small></article>').join('');renderBars(rows);document.getElementById('rings').innerHTML=[['No Devenga',pct(nd,total),'#c68a22'],['Vencida',pct(over,total),'#c13e49'],['ND +90',pct(over90,total),'#184b70']].map(r=>'<article class="ring-card"><div class="ring" style="--p:'+r[1]+';--c:'+r[2]+'"><strong>'+r[1]+'%</strong></div><span>'+r[0]+'</span></article>').join('');const highest=[...rows].sort((a,b)=>(b.noDevenga/b.total)-(a.noDevenga/a.total))[0];document.getElementById('insight').innerHTML=highest?'<strong>Lectura ejecutiva</strong><p>'+highest.company+' presenta la mayor participación de No Devenga ('+pct(highest.noDevenga,highest.total)+'%). '+(REPORT.scenario==='PROYECTADO'?integer(recl)+' operaciones fueron reclasificadas en el escenario seleccionado.':'Los valores corresponden al corte real procesado.')+'</p>':'<strong>Sin información</strong>';document.getElementById('tableBody').innerHTML=rows.map(r=>'<tr><td><strong>'+r.company+'</strong></td><td class="num">'+money(r.total)+'</td><td class="num">'+integer(r.operations)+'</td><td class="num">'+money(r.overdue)+'</td><td class="num">'+money(r.noDevenga)+'</td><td class="num">'+money(r.noDevenga6090)+'</td><td class="num">'+money(r.noDevengaOver90)+'</td><td class="num">'+(r.chargedOff?money(r.chargedOff):'—')+'</td>'+(REPORT.scenario==='PROYECTADO'?'<td class="num">'+integer(r.reclassified)+'</td>':'')+'<td><span class="pill">'+r.status+'</span></td></tr>').join('')}
function renderBars(rows){const max=Math.max(1,...rows.map(r=>Number(r[metric])||0));const tone=metric==='overdue'?'critical':metric==='noDevenga'?'risk':'';document.getElementById('bars').innerHTML=rows.map(r=>'<div class="bar-row" data-tone="'+tone+'"><strong>'+r.company+'</strong><div class="track"><div class="fill" style="width:'+Math.max(1,Math.round((r[metric]||0)/max*100))+'%"></div></div><span class="bar-value">'+money(r[metric])+'</span></div>').join('')}
render();
<\/script>
</body>
</html>`;
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


  function userFriendlyError(error) {
    const message = String(error && error.message ? error.message : '');
    if (/misma fecha|fecha de corte/i.test(message)) return 'Los reportes deben corresponder a una misma fecha de corte.';
    if (/duplicad|empresa/i.test(message)) return message;
    if (/límite|tamaño|archivo/i.test(message)) return message;
    if (/90 segundos|superó/i.test(message)) return 'El procesamiento tardó más de lo esperado. Reduce el número de archivos o vuelve a intentarlo.';
    if (/xlsx|excel|plantilla|paquete/i.test(message)) return 'No fue posible generar el archivo Excel. Verifica los reportes cargados y vuelve a intentarlo.';
    return 'No se pudo completar el procesamiento. Revisa los archivos y vuelve a intentarlo.';
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
