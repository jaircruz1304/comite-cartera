'use strict';
const assert=require('node:assert/strict');
globalThis.APP_CONFIG={githubOwner:'demo',githubRepo:'repo',githubSourcePath:'sources/current'};
require('../js/github.js');
const G=globalThis.DriveConnector;
assert.ok(G);assert.equal(G.isConfigured(),true);assert.equal(G.normalizeFolderId('/sources/current/'),'sources/current');
const files=[{id:'1',name:'CTH.xlsx',modifiedTime:'2026-01-01'},{id:'2',name:'CTH nuevo.xlsx',modifiedTime:'2026-02-01'},{id:'3',name:'F8.xlsx',modifiedTime:'2026-01-15'}];
const chosen=G.chooseLatestPerCompany(files,n=>({code:/F8/i.test(n)?'F8':'CTH'}));
assert.deepEqual(chosen.map(x=>x.id).sort(),['2','3']);console.log('github-smoke: OK');
