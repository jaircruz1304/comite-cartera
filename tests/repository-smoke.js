'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const template = path.join(root, 'assets', 'COMITE_BASE.xlsx');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(root, 'js', 'app.js'), 'utf8');

assert.ok(fs.existsSync(template), 'Falta assets/COMITE_BASE.xlsx');
assert.ok(fs.statSync(template).size > 1000, 'La plantilla integrada está vacía');
assert.match(index, /id="sourceInput"/);
assert.doesNotMatch(index, /id="templateInput"/);
assert.match(index, /Plantilla COMITE integrada/);
assert.match(app, /assets\/COMITE_BASE\.xlsx/);
assert.match(app, /outputAsync\(\{ type: 'arraybuffer' \}\)/);
assert.match(app, /validateXlsxPackage\(outputBuffer/);

console.log('repository-smoke: OK');
