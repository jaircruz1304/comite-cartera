'use strict';

const assert = require('node:assert/strict');
const P = require('../js/processor.js');
const R = require('../js/projection.js');

const values = [
  ['CORPORACION DE DESARROLLO DE MERCADO SECUNDARIO DE HIPOTECAS CTH S.A.'],
  ['SALDOS DE CARTERA'],
  [' Fecha Corte: 23-07-2026'],
  [],
  ['', '', 'Cartera Por Vencer', '', '', '', '', 'Cartera que no devenga Intereses', '', '', '', '', 'Cartera Vencida', '', '', '', '', '', '', '', '', '', '', ''],
  ['#', 'CaliF.Cont.', 'Saldo por vencer 1 a 30 días', 'Saldo por vencer 31 a 90 días', 'Saldo por vencer 91 a 180 días', 'Saldo por vencer 181 a 360 días', 'Saldo por vencer Mas 360 días', 'Saldo por vencer 1 a 30 días', 'Saldo por vencer 31 a 90 días', 'Saldo por vencer 91 a 180 días', 'Saldo por vencer 181 a 360 días', 'Saldo por vencer Mas 360 días', 'Saldo Vencido 1 a 30 días', 'Saldo Vencido 31 a 90 días', 'Saldo Vencido 91 a 270 días', 'Saldo Vencido 271 a 360 días', 'Saldo Vencido 361 a 720 días', 'Saldo Vencido Mas 720 días', 'Cartera en Demanda', 'Cartera Castigada', 'Cartera Enajenación', 'Valor Premio', '', 'Dias Morosidad'],
  [1, 'NOR', 100, 200, 300, 400, 500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 56],
  [2, 'NORMAL', 10, 20, 30, 40, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 52],
  [3, 'RE', 0, 0, 0, 0, 0, 25, 35, 45, 55, 65, 15, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 70]
];

const normal = P.analyzeValues(values, 'SaldosDeCarteraSencillo_Report cth 23jul2026.xlsx');
normal.sourceValues = values;
const projected = R.projectAnalysis(normal, { horizonDays: 8, thresholdDays: 60, reclassify: true });

assert.equal(P.dateKey(projected.metrics.baseDate), '2026-07-23');
assert.equal(P.dateKey(projected.metrics.cutDate), '2026-07-31');
assert.equal(projected.projection.reclassifiedOperations, 1);
assert.equal(projected.sourceValues[6][23], 64);
assert.deepEqual(projected.sourceValues[6].slice(2, 7), [0, 0, 0, 0, 0]);
assert.deepEqual(projected.sourceValues[6].slice(7, 12), [200, 300, 400, 500, 0]);
assert.equal(projected.sourceValues[6][12], 100);
assert.equal(projected.sourceValues[7][23], 60);
assert.deepEqual(projected.sourceValues[7].slice(2, 7), [10, 20, 30, 40, 50]);
assert.equal(projected.sourceValues[8][23], 78);
assert.deepEqual(projected.sourceValues[8].slice(7, 13), [25, 35, 45, 55, 65, 15]);
assert.equal(projected.metrics.total.value, normal.metrics.total.value);

console.log('projection-smoke: OK');
