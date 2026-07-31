'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const template = path.join(root, 'assets', 'COMITE_BASE.xlsx');
const index = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const app = fs.readFileSync(path.join(root, 'js', 'app.js'), 'utf8');
const engine = fs.readFileSync(path.join(root, 'js', 'ooxml.js'), 'utf8');

assert.ok(fs.existsSync(template), 'Falta assets/COMITE_BASE.xlsx');
assert.ok(fs.statSync(template).size > 1000, 'La plantilla integrada está vacía');
assert.ok(fs.existsSync(path.join(root, 'vendor', 'jszip.min.js')), 'Falta JSZip local');
assert.match(index, /id="sourceInput"/);
assert.doesNotMatch(index, /id="templateInput"/);
assert.doesNotMatch(index, /xlsx-populate/i);
assert.match(index, /js\/ooxml\.js/);
assert.match(app, /assets\/COMITE_BASE\.xlsx/);
assert.match(app, /X\.build\(/);
assert.match(app, /X\.readFirstSheet/);
assert.match(engine, /Motor OOXML|OOXMLWorkbook|function build/);

console.log('repository-smoke: OK');
