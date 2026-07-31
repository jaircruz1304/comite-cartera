'use strict';

const assert = require('node:assert/strict');
const P = require('../js/processor.js');

const values = [
  ['CORPORACION DE DESARROLLO DE MERCADO SECUNDARIO DE HIPOTECAS CTH S.A.'],
  ['SALDOS DE CARTERA'],
  [' Fecha Corte: 29-07-2026'],
  [],
  ['', '', 'Cartera Por Vencer', '', '', '', '', 'Cartera que no devenga Intereses', '', '', '', '', 'Cartera Vencida', '', '', '', '', '', '', '', '', '', '', ''],
  ['#', 'CaliF.Cont.', 'Saldo por vencer 1 a 30 días', 'Saldo por vencer 31 a 90 días', 'Saldo por vencer 91 a 180 días', 'Saldo por vencer 181 a 360 días', 'Saldo por vencer Mas 360 días', 'Saldo por vencer 1 a 30 días', 'Saldo por vencer 31 a 90 días', 'Saldo por vencer 91 a 180 días', 'Saldo por vencer 181 a 360 días', 'Saldo por vencer Mas 360 días', 'Saldo Vencido 1 a 30 días', 'Saldo Vencido 31 a 90 días', 'Saldo Vencido 91 a 270 días', 'Saldo Vencido 271 a 360 días', 'Saldo Vencido 361 a 720 días', 'Saldo Vencido Mas 720 días', 'Cartera en Demanda', 'Cartera Castigada', 'Cartera Enajenación', 'Valor Premio', '', 'Dias Morosidad'],
  [1, 'NOR', 1000, 0, 0, 0, 0, 200, 0, 0, 0, 0, 30, 0, 0, 0, 0, 0, 0, 0, 0, 0, '', 65],
  [2, 'RE', 500, 0, 0, 0, 0, 0, 300, 0, 0, 0, 0, 80, 0, 0, 0, 0, 0, 0, 0, 0, '', 120],
  [3, 'GRA', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 400, 0, 0, '', 200],
  ['TOTAL', 1500, 0, 0, 0, 0, 200, 300, 0, 0, 0, 30, 80, 0, 0, 0, 0, 0, 400]
];

const analysis = P.analyzeValues(values, 'SaldosDeCarteraSencillo_Report cth 29jul.xlsx');
const m = analysis.metrics;

assert.equal(m.company, 'CTH');
assert.equal(P.dateKey(m.cutDate), '2026-07-29');
assert.equal(m.total.value, 2110);
assert.equal(m.total.operations, 2);
assert.equal(m.overdue.value, 110);
assert.equal(m.overdue.operations, 2);
assert.equal(m.noDevenga.value, 500);
assert.equal(m.noDevenga.operations, 2);
assert.equal(m.noDevengaNormal.value, 200);
assert.equal(m.noDevengaRestructured.value, 300);
assert.equal(m.noDevenga6090.value, 200);
assert.equal(m.noDevengaOver90.value, 300);
assert.equal(m.chargedOff.value, 400);
assert.equal(m.chargedOff.operations, 1);
assert.equal(m.status, 'OK');

console.log('processor-smoke: OK');
