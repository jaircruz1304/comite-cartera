'use strict';

const assert = require('node:assert/strict');
require('../js/config.js');
require('../js/drive.js');

const D = globalThis.DriveConnector;
assert.ok(D, 'No se exportó DriveConnector');
assert.equal(D.normalizeFolderId('https://drive.google.com/drive/folders/ABC_123-xyz?usp=sharing'), 'ABC_123-xyz');
assert.equal(D.normalizeFolderId('ABC_123-xyz'), 'ABC_123-xyz');

const files = [
  { id: '1', name: 'SaldosDeCarteraSencillo_Report CTH.xlsx', modifiedTime: '2026-07-30T10:00:00Z' },
  { id: '2', name: 'SaldosDeCarteraSencillo_Report CTH nuevo.xlsx', modifiedTime: '2026-07-31T10:00:00Z' },
  { id: '3', name: 'SaldosDeCarteraSencillo_Report F8.xlsx', modifiedTime: '2026-07-31T09:00:00Z' }
];
const detect = (name) => ({ code: /F8/i.test(name) ? 'F8' : 'CTH' });
const selected = D.chooseLatestPerCompany(files, detect);
assert.deepEqual(selected.map((file) => file.id).sort(), ['2', '3']);
console.log('drive-smoke: OK');
