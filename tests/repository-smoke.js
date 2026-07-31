'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const template = path.join(root, 'assets', 'COMITE_BASE.xlsx');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(root, 'js', 'app.js'), 'utf8');
const engine = fs.readFileSync(path.join(root, 'js', 'ooxml.js'), 'utf8');
const projection = fs.readFileSync(path.join(root, 'js', 'projection.js'), 'utf8');

assert.ok(fs.existsSync(template), 'Falta assets/COMITE_BASE.xlsx');
assert.ok(fs.statSync(template).size > 1000, 'La plantilla integrada está vacía');
assert.ok(fs.existsSync(path.join(root, 'vendor', 'jszip.min.js')), 'Falta JSZip local');
assert.match(index, /id="sourceInput"/);
assert.match(index, /id="horizonDays"/);
assert.match(index, /id="downloadNormalButton"/);
assert.match(index, /id="downloadProjectedButton"/);
assert.match(index, /js\/projection\.js/);
assert.doesNotMatch(index, /id="templateInput"/);
assert.doesNotMatch(index, /xlsx-populate/i);
assert.match(app, /assets\/COMITE_BASE\.xlsx/);
assert.match(app, /R\.projectAnalysis/);
assert.match(app, /X\.build\(/);
assert.match(engine, /OOXMLWorkbook|function build/);
assert.match(projection, /reclassifiedOperations/);

console.log('repository-smoke: OK');
