'use strict';
const assert=require('node:assert/strict');const fs=require('node:fs');const path=require('node:path');const root=path.resolve(__dirname,'..');
const index=fs.readFileSync(path.join(root,'index.html'),'utf8');const hist=fs.readFileSync(path.join(root,'historico.html'),'utf8');const app=fs.readFileSync(path.join(root,'js/app.js'),'utf8');
['assets/COMITE_BASE.xlsx','assets/logo_CTH.png','vendor/jszip.min.js','js/config.js','js/github.js','js/history.js','data/historico/catalogo.json','data/historico/series.json'].forEach(f=>assert.ok(fs.existsSync(path.join(root,f)),`Falta ${f}`));
assert.match(index,/historico\.html/);assert.match(index,/js\/github\.js/);assert.doesNotMatch(index,/accounts\.google\.com\/gsi\/client/);assert.match(hist,/Análisis histórico/);assert.match(hist,/Exportar Excel/);assert.match(hist,/Exportar PDF/);assert.match(hist,/Lizbeth Sanipatín/);assert.match(app,/assets\/COMITE_BASE\.xlsx/);console.log('repository-smoke: OK');
