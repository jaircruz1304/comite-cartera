(function (global) {
  'use strict';
  const config = global.APP_CONFIG || {};
  let connected = false;
  function configuredValue(v){ return String(v||'').trim() && !/PEGUE_AQUI|TU_USUARIO|TU_REPOSITORIO/i.test(String(v)); }
  function isConfigured(){ return configuredValue(config.githubOwner) && configuredValue(config.githubRepo); }
  function isConnected(){ return connected && isConfigured(); }
  async function connect(){ if(!isConfigured()) throw new Error('Configure githubOwner y githubRepo en js/config.js.'); connected=true; return {connected:true}; }
  async function disconnect(){ connected=false; }
  function normalizeFolderId(value){ return String(value||config.githubSourcePath||'sources/current').trim().replace(/^\/+|\/+$/g,''); }
  async function api(path){
    const url=`https://api.github.com${path}`;
    const r=await fetch(url,{headers:{Accept:'application/vnd.github+json'}});
    if(!r.ok) throw new Error(`GitHub respondió ${r.status}. Verifique que el repositorio y la ruta sean públicos.`);
    return r.json();
  }
  async function lastCommitDate(path){
    try{
      const q=new URLSearchParams({path,sha:config.githubBranch||'main',per_page:'1'});
      const rows=await api(`/repos/${encodeURIComponent(config.githubOwner)}/${encodeURIComponent(config.githubRepo)}/commits?${q}`);
      return rows[0] && rows[0].commit && rows[0].commit.committer ? rows[0].commit.committer.date : null;
    }catch(e){ return null; }
  }
  async function listFolderFiles(folderValue){
    if(!isConfigured()) throw new Error('GitHub no está configurado en js/config.js.');
    const folderId=normalizeFolderId(folderValue);
    const ref=encodeURIComponent(config.githubBranch||'main');
    const rows=await api(`/repos/${encodeURIComponent(config.githubOwner)}/${encodeURIComponent(config.githubRepo)}/contents/${folderId}?ref=${ref}`);
    if(!Array.isArray(rows)) throw new Error('La ruta configurada no corresponde a una carpeta.');
    const matches=rows.filter(x=>x.type==='file' && /\.xlsx$/i.test(x.name) && (!config.reportNameContains || x.name.toLowerCase().includes(String(config.reportNameContains).toLowerCase())));
    const files=await Promise.all(matches.map(async x=>({id:x.sha,name:x.name,size:x.size,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',modifiedTime:await lastCommitDate(x.path),download_url:x.download_url,path:x.path,html_url:x.html_url})));
    files.sort((a,b)=>new Date(b.modifiedTime||0)-new Date(a.modifiedTime||0));
    return {folderId,files};
  }
  async function downloadFile(metadata){
    const r=await fetch(metadata.download_url,{cache:'no-store'}); if(!r.ok) throw new Error(`No se pudo descargar ${metadata.name}.`);
    const buffer=await r.arrayBuffer(); if(buffer.byteLength<1000) throw new Error(`${metadata.name} está vacío o incompleto.`);
    return new File([buffer],metadata.name,{type:metadata.mimeType,lastModified:metadata.modifiedTime?new Date(metadata.modifiedTime).getTime():Date.now()});
  }
  async function downloadFiles(files,onProgress){ const out=[]; for(let i=0;i<files.length;i++){ if(onProgress)onProgress(i,files.length,files[i]); out.push(await downloadFile(files[i])); } if(onProgress)onProgress(files.length,files.length,null); return out; }
  function chooseLatestPerCompany(files,detectCompany){ const m=new Map(); (files||[]).forEach(f=>{try{const c=detectCompany(f.name,null); const e=m.get(c.code); if(!e||new Date(f.modifiedTime||0)>new Date(e.modifiedTime||0))m.set(c.code,f);}catch(_){}}); return [...m.values()]; }
  global.DriveConnector=Object.freeze({isConfigured,isConnected,connect,disconnect,normalizeFolderId,listFolderFiles,downloadFile,downloadFiles,chooseLatestPerCompany,getDefaultFolderId:()=>String(config.githubSourcePath||'sources/current')});
})(typeof window!=='undefined'?window:globalThis);
