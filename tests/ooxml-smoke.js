'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

global.JSZip = require('../vendor/jszip.min.js');
const P = require('../js/processor.js');
const X = require('../js/ooxml.js');

function metric(value, operations) { return { value, operations }; }

(async () => {
  const templatePath = path.resolve(__dirname, '..', 'assets', 'COMITE_BASE.xlsx');
  const template = fs.readFileSync(templatePath);
  const templateBuffer = template.buffer.slice(template.byteOffset, template.byteOffset + template.byteLength);
  await X.inspectTemplate(templateBuffer, ['CTH', 'F12', 'F8', 'F11']);

  const metrics = {
    company: 'CTH',
    cutDate: new Date(2026, 6, 29, 12),
    fileName: 'SaldosDeCarteraSencillo_Report cth.xlsx',
    total: metric(11207598.62, 619),
    original: metric(11207598.62, 619),
    overdue: metric(187709.46, 93),
    noDevenga: metric(1906326.83, 89),
    noDevengaNormal: metric(909071.94, 55),
    noDevengaRestructured: metric(997254.89, 34),
    noDevenga6090: metric(705255.65, 30),
    noDevenga6090Normal: metric(326929.10, 18),
    noDevenga6090Restructured: metric(378326.55, 12),
    noDevengaOver90: metric(1201071.18, 59),
    noDevengaOver90Normal: metric(582142.84, 37),
    noDevengaOver90Restructured: metric(618928.34, 22),
    chargedOff: metric(1258491.55, 54),
    ratios: {
      overdueTotal: 0.0167,
      overdueOperations: 0.1502,
      noDevengaTotal: 0.1701,
      noDevengaOperations: 0.1438,
      noDevengaOriginal: 0.1701,
      noDevengaOver90Total: 0.1072,
      noDevengaOver90Operations: 0.0953
    },
    duplicatesExcluded: 0,
    status: 'OK',
    warnings: [],
    detailRows: 673
  };
  const analysis = {
    company: { code: 'CTH', reportSheet: 'CTH' },
    metrics,
    sourceValues: [['Fecha Corte: 29-07-2026'], [], [], [], ['Grupo'], ['#', 'Valor'], [1, 100]]
  };

  const output = await X.build(templateBuffer, [analysis], {
    templateRows: P.TEMPLATE_ROWS,
    requiredSheets: P.COMPANY_RULES.map((item) => item.reportSheet),
    periodLabel: P.periodLabel
  });
  assert.ok(output.byteLength > 1000);
  const inspection = await X.inspectTemplate(output, ['CTH', 'F12', 'F8', 'F11', 'CONTROL_EJECUCION', 'ORIGEN_CTH']);
  assert.ok(inspection.names.includes('CONTROL_EJECUCION'));
  assert.ok(inspection.names.includes('ORIGEN_CTH'));
  console.log('ooxml-smoke: OK');
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
