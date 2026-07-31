(function (global) {
  'use strict';

  const P = global.PortfolioProcessor;

  function assertRuntime() {
    if (!P) throw new Error('El motor de cartera no está disponible.');
  }

  function cloneMatrix(values) {
    return (values || []).map((row) => Array.isArray(row) ? row.slice() : []);
  }

  function isDetailIdentifier(value) {
    if (typeof value === 'number') return Number.isFinite(value);
    return /^\d+$/.test(String(value == null ? '' : value).trim());
  }

  function addDays(date, days) {
    const result = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
    result.setDate(result.getDate() + Number(days || 0));
    return result;
  }

  function normalizeParameters(parameters) {
    const source = parameters || {};
    const horizonDays = Math.max(0, Math.round(Number(source.horizonDays == null ? 8 : source.horizonDays)));
    const thresholdDays = Math.max(0, Math.round(Number(source.thresholdDays == null ? 60 : source.thresholdDays)));
    const reclassify = source.reclassify !== false;
    return { horizonDays, thresholdDays, reclassify };
  }

  function sumAt(row, indexes) {
    return indexes.reduce((sum, index) => sum + P.toNumber(row[index]), 0);
  }

  function validateStructure(columns) {
    if (!columns || columns.daysPastDue < 0) throw new Error('No se encontró la columna Dias Morosidad.');
    if (!Array.isArray(columns.current) || columns.current.length < 5) {
      throw new Error('La proyección requiere cinco columnas de Cartera Por Vencer.');
    }
    if (!Array.isArray(columns.noDevenga) || columns.noDevenga.length < 5) {
      throw new Error('La proyección requiere cinco columnas de Cartera que no devenga Intereses.');
    }
    if (!Array.isArray(columns.overdue) || columns.overdue.length < 1) {
      throw new Error('La proyección requiere columnas de Cartera Vencida.');
    }
  }

  function projectAnalysis(normalAnalysis, parameters) {
    assertRuntime();
    if (!normalAnalysis || !normalAnalysis.sourceValues) throw new Error('No se recibió el escenario normal para proyectar.');

    const params = normalizeParameters(parameters);
    const columns = normalAnalysis.columns;
    validateStructure(columns);

    const projectedValues = cloneMatrix(normalAnalysis.sourceValues);
    const stats = {
      company: normalAnalysis.company.code,
      rowsAdvanced: 0,
      reclassifiedOperations: 0,
      movedToNoDevenga: 0,
      movedToOverdue: 0,
      totalReclassified: 0
    };

    for (let r = normalAnalysis.headerRowIndex + 1; r < projectedValues.length; r += 1) {
      const row = projectedValues[r] || [];
      if (!isDetailIdentifier(row[columns.id])) continue;

      const originalDays = P.toNumber(row[columns.daysPastDue]);
      const projectedDays = originalDays + params.horizonDays;
      row[columns.daysPastDue] = projectedDays;
      stats.rowsAdvanced += 1;

      if (!params.reclassify || projectedDays <= params.thresholdDays) continue;

      const chargedOff = columns.chargedOff >= 0 ? P.toNumber(row[columns.chargedOff]) : 0;
      const currentValue = sumAt(row, columns.current);
      const noDevengaValue = sumAt(row, columns.noDevenga);
      const overdueValue = sumAt(row, columns.overdue);

      // Regla validada con los archivos de referencia 23→31 de julio de 2026:
      // solo migra cartera que estaba íntegramente Por Vencer y que supera el umbral.
      if (chargedOff > 0 || currentValue <= 0 || noDevengaValue > 0 || overdueValue > 0) continue;

      const current = columns.current;
      const noDevenga = columns.noDevenga;
      const overdue = columns.overdue;

      const firstInstallment = P.toNumber(row[current[0]]);
      const futureInstallments = [
        P.toNumber(row[current[1]]),
        P.toNumber(row[current[2]]),
        P.toNumber(row[current[3]]),
        P.toNumber(row[current[4]])
      ];

      current.forEach((index) => { row[index] = 0; });
      noDevenga.forEach((index) => { row[index] = 0; });

      row[noDevenga[0]] = futureInstallments[0];
      row[noDevenga[1]] = futureInstallments[1];
      row[noDevenga[2]] = futureInstallments[2];
      row[noDevenga[3]] = futureInstallments[3];
      row[noDevenga[4]] = 0;
      row[overdue[0]] = firstInstallment;

      const movedToNoDevenga = futureInstallments.reduce((sum, value) => sum + value, 0);
      stats.reclassifiedOperations += 1;
      stats.movedToNoDevenga += movedToNoDevenga;
      stats.movedToOverdue += firstInstallment;
      stats.totalReclassified += currentValue;
    }

    const targetDate = addDays(normalAnalysis.metrics.cutDate, params.horizonDays);
    const projected = P.analyzeValues(projectedValues, normalAnalysis.metrics.fileName, normalAnalysis.company);
    projected.sourceValues = projectedValues;
    projected.metrics.baseDate = new Date(normalAnalysis.metrics.cutDate.getTime());
    projected.metrics.cutDate = targetDate;
    projected.metrics.scenario = 'PROYECTADO';
    projected.metrics.projection = Object.assign({}, params, stats, { targetDate });
    projected.projection = projected.metrics.projection;

    return projected;
  }

  function compareReference(normalAnalysis, projectedAnalysis) {
    const normal = normalAnalysis.metrics;
    const projected = projectedAnalysis.metrics;
    return {
      company: normal.company,
      totalDelta: projected.total.value - normal.total.value,
      overdueDelta: projected.overdue.value - normal.overdue.value,
      overdueOperationsDelta: projected.overdue.operations - normal.overdue.operations,
      noDevengaDelta: projected.noDevenga.value - normal.noDevenga.value,
      noDevengaOperationsDelta: projected.noDevenga.operations - normal.noDevenga.operations,
      noDevenga6090Delta: projected.noDevenga6090.value - normal.noDevenga6090.value,
      noDevengaOver90Delta: projected.noDevengaOver90.value - normal.noDevengaOver90.value,
      chargedOffDelta: (projected.chargedOff.value || 0) - (normal.chargedOff.value || 0),
      reclassifiedOperations: projected.projection ? projected.projection.reclassifiedOperations : 0
    };
  }

  global.ProjectionEngine = Object.freeze({
    addDays,
    normalizeParameters,
    projectAnalysis,
    compareReference
  });

  if (typeof module !== 'undefined' && module.exports) module.exports = global.ProjectionEngine;
})(typeof window !== 'undefined' ? window : globalThis);
