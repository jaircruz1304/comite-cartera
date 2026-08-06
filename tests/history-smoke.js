'use strict';
const assert=require('node:assert/strict');const fs=require('node:fs');const path=require('node:path');
const root=path.resolve(__dirname,'..');const catalog=JSON.parse(fs.readFileSync(path.join(root,'data/historico/catalogo.json'),'utf8'));const series=JSON.parse(fs.readFileSync(path.join(root,'data/historico/series.json'),'utf8'));
assert.equal(catalog.periodStart,'2013-07');assert.equal(catalog.periodEnd,'2026-07');assert.equal(catalog.periodCount,157);assert.equal(catalog.companyPeriodCount,424);assert.equal(series.records.length,424);
for(const c of ['CTH','F8','F11','F12'])assert.ok(series.records.some(r=>r.company===c));
assert.ok(fs.existsSync(path.join(root,'historico.html')));assert.ok(fs.existsSync(path.join(root,'js/history.js')));console.log('history-smoke: OK');
