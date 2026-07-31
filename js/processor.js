(function (global) {
  'use strict';

  const COMPANY_RULES = Object.freeze([
    { code: 'CTH', reportSheet: 'CTH', patterns: ['cth', 'corporacion de desarrollo'] },
    { code: 'F12', reportSheet: 'F12', patterns: ['f12'] },
    { code: 'F8', reportSheet: 'F8', patterns: ['f8'] },
    { code: 'F11', reportSheet: 'F11', patterns: ['f11'] }
  ]);

  const TEMPLATE_ROWS = Object.freeze({
    TOTAL: 3,
    ORIGINAL: 4,
    OVERDUE: 5,
    OVERDUE_RATIO: 6,
    ND_TOTAL: 7,
    ND_NORMAL: 8,
    ND_RESTRUCTURED: 9,
    ND_60_90: 10,
    ND_60_90_NORMAL: 11,
    ND_60_90_RESTRUCTURED: 12,
    ND_RATIO_TOTAL: 13,
    ND_RATIO_ORIGINAL: 14,
    ND_OVER_90: 15,
    ND_OVER_90_NORMAL: 16,
    ND_OVER_90_RESTRUCTURED: 17,
    ND_OVER_90_RATIO: 18,
    BEP: 19,
    BEP_PROVISION: 20,
    BEP_NET: 21,
    CHARGED_OFF: 22
  });

  const THRESHOLDS = Object.freeze({
    overdueTotal: { green: 0.02, yellow: 0.05 },
    noDevengaTotal: { green: 0.10, yellow: 0.20 },
    noDevengaOver90Total: { green: 0.05, yellow: 0.10 }
  });

  function normalizeText(value) {
    return String(value == null ? '' : value)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/\s+/g, ' ')
      .trim()
      .toUpperCase();
  }

  function toNumber(value) {
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    if (value instanceof Date || value == null || value === '') return 0;

    let text = String(value).trim().replace(/[$€£\s]/g, '');
    if (!text) return 0;

    const lastComma = text.lastIndexOf(',');
    const lastDot = text.lastIndexOf('.');

    if (lastComma > -1 && lastDot > -1) {
      if (lastComma > lastDot) {
        text = text.replace(/\./g, '').replace(',', '.');
      } else {
        text = text.replace(/,/g, '');
      }
    } else if (lastComma > -1) {
      const decimals = text.length - lastComma - 1;
      text = decimals === 2 ? text.replace(',', '.') : text.replace(/,/g, '');
    }

    text = text.replace(/[^0-9.\-]/g, '');
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
  }

  function isDetailIdentifier(value) {
    if (typeof value === 'number') return Number.isFinite(value);
    return /^\d+$/.test(String(value == null ? '' : value).trim());
  }

  function stableValue(value) {
    if (value instanceof Date) return value.toISOString();
    if (typeof value === 'number') return Math.round(value * 1000000) / 1000000;
    return normalizeText(value);
  }

  function rowSignature(row) {
    return JSON.stringify((row || []).slice(1).map(stableValue));
  }

  function safeDivide(numerator, denominator) {
    return denominator ? numerator / denominator : null;
  }

  function sumColumns(row, indexes) {
    return indexes.reduce((sum, index) => sum + toNumber(row[index]), 0);
  }

  function findHeaderRow(values) {
    const maxRows = Math.min(values.length, 30);
    for (let r = 0; r < maxRows; r += 1) {
      const row = (values[r] || []).map(normalizeText);
      const hasId = row.some((value) => value === '#');
      const hasClass = row.some((value) => value.includes('CALIF') && value.includes('CONT'));
      const hasDays = row.some((value) => value.includes('DIAS MOROSIDAD'));
      if (hasId && hasClass && hasDays) return r;
    }
    return -1;
  }

  function propagateGroups(groupRow) {
    let current = '';
    return (groupRow || []).map((value) => {
      const normalized = normalizeText(value);
      if (normalized) current = normalized;
      return current;
    });
  }

  function findHeaderIndex(headers, predicate) {
    for (let i = 0; i < headers.length; i += 1) {
      if (predicate(headers[i])) return i;
    }
    return -1;
  }

  function buildColumnMap(headers, groups) {
    const result = {
      id: findHeaderIndex(headers, (h) => h === '#'),
      accountingClass: findHeaderIndex(headers, (h) => h.includes('CALIF') && h.includes('CONT')),
      daysPastDue: findHeaderIndex(headers, (h) => h.includes('DIAS MOROSIDAD')),
      chargedOff: findHeaderIndex(headers, (h) => h.includes('CARTERA CASTIGADA')),
      demand: findHeaderIndex(headers, (h) => h.includes('CARTERA EN DEMANDA')),
      sold: findHeaderIndex(headers, (h) => h.includes('CARTERA ENAJENACION')),
      premium: findHeaderIndex(headers, (h) => h.includes('VALOR PREMIO')),
      active: [],
      noDevenga: [],
      overdue: []
    };

    for (let c = 0; c < headers.length; c += 1) {
      const group = groups[c] || '';
      const header = headers[c] || '';
      const isMaturingBalance = group.includes('CARTERA POR VENCER') && header.includes('SALDO POR VENCER');
      const isNoAccrualBalance = group.includes('CARTERA QUE NO DEVENGA INTERESES') && header.includes('SALDO POR VENCER');
      const isOverdueBalance = group.includes('CARTERA VENCIDA') && header.includes('SALDO VENCIDO');

      if (isMaturingBalance || isNoAccrualBalance || isOverdueBalance) result.active.push(c);
      if (isNoAccrualBalance) result.noDevenga.push(c);
      if (isOverdueBalance) result.overdue.push(c);
    }

    return result;
  }

  function validateRequiredColumns(columns) {
    const missing = [];
    if (columns.id < 0) missing.push('#');
    if (columns.accountingClass < 0) missing.push('CaliF.Cont.');
    if (columns.daysPastDue < 0) missing.push('Dias Morosidad');
    if (columns.chargedOff < 0) missing.push('Cartera Castigada');
    if (!columns.active.length) missing.push('grupos de cartera activa');
    if (!columns.noDevenga.length) missing.push('Cartera que no devenga Intereses');
    if (!columns.overdue.length) missing.push('Cartera Vencida');
    if (missing.length) throw new Error(`Columnas o grupos requeridos no encontrados: ${missing.join(', ')}`);
  }

  function classifyAccounting(value) {
    const text = normalizeText(value);
    if (text.includes('RE')) return 'REESTRUCTURADA';
    if (text.includes('NOR') || text.includes('GRA')) return 'NORMAL';
    return 'OTRA';
  }

  function createMetric() {
    return { value: 0, operations: 0 };
  }

  function createEmptyMetrics(company, cutDate, fileName) {
    return {
      company,
      cutDate,
      fileName,
      total: createMetric(),
      original: createMetric(),
      overdue: createMetric(),
      noDevenga: createMetric(),
      noDevengaNormal: createMetric(),
      noDevengaRestructured: createMetric(),
      noDevenga6090: createMetric(),
      noDevenga6090Normal: createMetric(),
      noDevenga6090Restructured: createMetric(),
      noDevengaOver90: createMetric(),
      noDevengaOver90Normal: createMetric(),
      noDevengaOver90Restructured: createMetric(),
      chargedOff: createMetric(),
      unclassifiedNoDevenga: createMetric(),
      informative: { demand: 0, sold: 0, premium: 0 },
      ratios: {},
      detailRows: 0,
      duplicatesExcluded: 0,
      nonDetailRows: 0,
      warnings: [],
      status: 'OK'
    };
  }

  function addByAccountingClass(normalBucket, restructuredBucket, accountingClass, value) {
    if (accountingClass === 'NORMAL') {
      normalBucket.value += value;
      normalBucket.operations += 1;
    } else if (accountingClass === 'REESTRUCTURADA') {
      restructuredBucket.value += value;
      restructuredBucket.operations += 1;
    }
  }

  function finalizeMetrics(metrics) {
    metrics.ratios.overdueTotal = safeDivide(metrics.overdue.value, metrics.total.value);
    metrics.ratios.overdueOperations = safeDivide(metrics.overdue.operations, metrics.total.operations);
    metrics.ratios.noDevengaTotal = safeDivide(metrics.noDevenga.value, metrics.total.value);
    metrics.ratios.noDevengaOperations = safeDivide(metrics.noDevenga.operations, metrics.total.operations);
    metrics.ratios.noDevengaOriginal = safeDivide(metrics.noDevenga.value, metrics.original.value);
    metrics.ratios.noDevengaOver90Total = safeDivide(metrics.noDevengaOver90.value, metrics.total.value);
    metrics.ratios.noDevengaOver90Operations = safeDivide(metrics.noDevengaOver90.operations, metrics.total.operations);
    metrics.portfolioToMature = Math.max(0, metrics.total.value - metrics.noDevenga.value - metrics.overdue.value);

    if (metrics.chargedOff.value <= 0) metrics.chargedOff = { value: null, operations: null };
  }

  function validateMetrics(metrics) {
    const tolerance = 0.01;
    const ndClassified = metrics.noDevengaNormal.value + metrics.noDevengaRestructured.value + metrics.unclassifiedNoDevenga.value;

    if (Math.abs(metrics.noDevenga.value - ndClassified) > tolerance) {
      metrics.warnings.push('El total No Devenga no coincide con la suma de clasificaciones.');
    }
    if (metrics.unclassifiedNoDevenga.operations > 0) {
      metrics.warnings.push(`${metrics.unclassifiedNoDevenga.operations} operaciones No Devenga no fueron clasificadas como NORMAL o REESTRUCTURADA.`);
    }
    if (metrics.noDevenga6090.value - metrics.noDevenga.value > tolerance) {
      metrics.warnings.push('No Devenga 60-90 supera al total No Devenga.');
    }
    if (metrics.noDevengaOver90.value - metrics.noDevenga.value > tolerance) {
      metrics.warnings.push('No Devenga +90 supera al total No Devenga.');
    }
    if (metrics.duplicatesExcluded > 0) {
      metrics.warnings.push(`${metrics.duplicatesExcluded} registros duplicados exactos fueron excluidos.`);
    }
    if (!metrics.total.operations) metrics.warnings.push('No se encontraron operaciones activas.');

    metrics.status = metrics.warnings.length ? 'REVISAR' : 'OK';
  }

  function parseCutDate(values, fileName, fallbackDate) {
    const maxRows = Math.min(values.length, 15);
    for (let r = 0; r < maxRows; r += 1) {
      const row = values[r] || [];
      for (let c = 0; c < Math.min(row.length, 8); c += 1) {
        const text = normalizeText(row[c]);
        if (!text.includes('FECHA CORTE')) continue;
        const raw = String(row[c] || '');
        const match = raw.match(/(\d{1,2})\s*[-/]\s*(\d{1,2})\s*[-/]\s*(\d{4})/);
        if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]), 12, 0, 0);
      }
    }

    const normalizedName = normalizeText(fileName);
    const numeric = normalizedName.match(/(\d{1,2})[-_ ](\d{1,2})[-_ ](20\d{2})/);
    if (numeric) return new Date(Number(numeric[3]), Number(numeric[2]) - 1, Number(numeric[1]), 12, 0, 0);

    const months = { ENE: 0, FEB: 1, MAR: 2, ABR: 3, MAY: 4, JUN: 5, JUL: 6, AGO: 7, SEP: 8, OCT: 9, NOV: 10, DIC: 11 };
    const named = normalizedName.match(/(\d{1,2})\s*(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)(?:\s*(20\d{2}))?/);
    if (named) {
      const base = fallbackDate || new Date();
      const year = named[3] ? Number(named[3]) : base.getFullYear();
      return new Date(year, months[named[2]], Number(named[1]), 12, 0, 0);
    }

    const fallback = fallbackDate || new Date();
    return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate(), 12, 0, 0);
  }

  function detectCompany(fileName, values) {
    const sampleRows = (values || []).slice(0, 8);
    const sampleText = sampleRows.flat().map(normalizeText).join(' ');
    const haystack = `${normalizeText(fileName)} ${sampleText}`;

    const matches = COMPANY_RULES.filter((company) => company.patterns.some((pattern) => haystack.includes(normalizeText(pattern))));
    if (matches.length === 1) return matches[0];

    if (matches.length > 1) {
      const exactFileMatch = matches.find((company) => normalizeText(fileName).includes(company.code));
      if (exactFileMatch) return exactFileMatch;
    }

    throw new Error(`No se pudo identificar de forma segura la empresa del archivo “${fileName}”.`);
  }

  function analyzeValues(values, fileName, forcedCompany) {
    if (!Array.isArray(values) || !values.length) throw new Error('El archivo no contiene información.');

    const company = forcedCompany || detectCompany(fileName, values);
    const headerRowIndex = findHeaderRow(values);
    if (headerRowIndex < 0) throw new Error('No se encontró la fila de encabezados con #, CaliF.Cont. y Dias Morosidad.');

    const groupRowIndex = Math.max(0, headerRowIndex - 1);
    const headers = (values[headerRowIndex] || []).map(normalizeText);
    const groups = propagateGroups(values[groupRowIndex] || []);
    const columns = buildColumnMap(headers, groups);
    validateRequiredColumns(columns);

    const cutDate = parseCutDate(values, fileName, new Date());
    const metrics = createEmptyMetrics(company.code, cutDate, fileName);
    const seen = new Set();

    for (let r = headerRowIndex + 1; r < values.length; r += 1) {
      const row = values[r] || [];
      const classMarker = normalizeText(row[columns.accountingClass]);
      const isSummaryMarker = /^[-+]?\d+(?:[.,]\d+)?$/.test(classMarker)
        || classMarker.includes('TOTAL')
        || classMarker.includes('SUBTOTAL');

      if (!isDetailIdentifier(row[columns.id]) || isSummaryMarker) {
        metrics.nonDetailRows += 1;
        continue;
      }

      const signature = rowSignature(row);
      if (seen.has(signature)) {
        metrics.duplicatesExcluded += 1;
        continue;
      }
      seen.add(signature);
      metrics.detailRows += 1;

      const activeValue = sumColumns(row, columns.active);
      const noDevengaValue = sumColumns(row, columns.noDevenga);
      const overdueValue = sumColumns(row, columns.overdue);
      const chargedOffValue = toNumber(row[columns.chargedOff]);
      const demandValue = columns.demand >= 0 ? toNumber(row[columns.demand]) : 0;
      const soldValue = columns.sold >= 0 ? toNumber(row[columns.sold]) : 0;
      const premiumValue = columns.premium >= 0 ? toNumber(row[columns.premium]) : 0;
      const daysPastDue = toNumber(row[columns.daysPastDue]);
      const accountingClass = classifyAccounting(row[columns.accountingClass]);
      const isChargedOff = chargedOffValue > 0;

      metrics.informative.demand += demandValue;
      metrics.informative.sold += soldValue;
      metrics.informative.premium += premiumValue;

      if (isChargedOff) {
        metrics.chargedOff.value += chargedOffValue;
        metrics.chargedOff.operations += 1;
        continue;
      }

      if (activeValue > 0) {
        metrics.total.value += activeValue;
        metrics.total.operations += 1;
      }

      if (overdueValue > 0) {
        metrics.overdue.value += overdueValue;
        metrics.overdue.operations += 1;
      }

      if (noDevengaValue <= 0) continue;

      metrics.noDevenga.value += noDevengaValue;
      metrics.noDevenga.operations += 1;

      if (accountingClass === 'NORMAL') {
        metrics.noDevengaNormal.value += noDevengaValue;
        metrics.noDevengaNormal.operations += 1;
      } else if (accountingClass === 'REESTRUCTURADA') {
        metrics.noDevengaRestructured.value += noDevengaValue;
        metrics.noDevengaRestructured.operations += 1;
      } else {
        metrics.unclassifiedNoDevenga.value += noDevengaValue;
        metrics.unclassifiedNoDevenga.operations += 1;
      }

      if (daysPastDue >= 60 && daysPastDue <= 90) {
        metrics.noDevenga6090.value += noDevengaValue;
        metrics.noDevenga6090.operations += 1;
        addByAccountingClass(metrics.noDevenga6090Normal, metrics.noDevenga6090Restructured, accountingClass, noDevengaValue);
      }

      if (daysPastDue > 90) {
        metrics.noDevengaOver90.value += noDevengaValue;
        metrics.noDevengaOver90.operations += 1;
        addByAccountingClass(metrics.noDevengaOver90Normal, metrics.noDevengaOver90Restructured, accountingClass, noDevengaValue);
      }
    }

    metrics.original.value = metrics.total.value;
    metrics.original.operations = metrics.total.operations;
    finalizeMetrics(metrics);
    validateMetrics(metrics);

    return { company, metrics, headerRowIndex, columns };
  }

  function dateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function fileDateKey(date) {
    return dateKey(date).replace(/-/g, '');
  }

  function periodLabel(date) {
    const months = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'];
    return `${months[date.getMonth()]}\n${date.getFullYear()}`;
  }

  function formatMoneyThousands(value) {
    if (value == null) return '—';
    return (Number(value) / 1000).toLocaleString('es-EC', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatPercent(value) {
    if (value == null) return '—';
    return Number(value).toLocaleString('es-EC', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function semaphore(value, threshold) {
    if (value == null) return 'neutral';
    if (value <= threshold.green) return 'green';
    if (value <= threshold.yellow) return 'yellow';
    return 'red';
  }

  function validateDateGroup(analyses) {
    const dates = [...new Set(analyses.map((item) => dateKey(item.metrics.cutDate)))];
    if (dates.length !== 1) {
      throw new Error(`Los archivos no corresponden a una sola fecha de corte: ${dates.join(', ')}.`);
    }

    const companies = analyses.map((item) => item.company.code);
    const duplicates = companies.filter((code, index) => companies.indexOf(code) !== index);
    if (duplicates.length) {
      throw new Error(`Se cargó más de un archivo para la empresa ${[...new Set(duplicates)].join(', ')}.`);
    }

    return dates[0];
  }

  global.PortfolioProcessor = Object.freeze({
    COMPANY_RULES,
    TEMPLATE_ROWS,
    THRESHOLDS,
    normalizeText,
    toNumber,
    analyzeValues,
    detectCompany,
    parseCutDate,
    dateKey,
    fileDateKey,
    periodLabel,
    formatMoneyThousands,
    formatPercent,
    semaphore,
    validateDateGroup
  });

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = global.PortfolioProcessor;
  }
})(typeof window !== 'undefined' ? window : globalThis);
