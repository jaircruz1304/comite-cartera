'use strict';

const assert = require('node:assert/strict');
require('../js/pdf.js');
assert.ok(globalThis.PDFReport, 'No se exportó PDFReport');
assert.equal(typeof globalThis.PDFReport.generate, 'function');
assert.equal(globalThis.PDFReport.available(), false, 'En Node no debe asumirse jsPDF disponible');
console.log('pdf-smoke: OK');
