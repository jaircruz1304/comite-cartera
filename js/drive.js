(function (global) {
  'use strict';

  const config = global.APP_CONFIG || {};
  let accessToken = '';
  let tokenClient = null;
  let expiresAt = 0;

  function isConfigured() {
    const value = String(config.googleOAuthClientId || '').trim();
    return Boolean(value && !/PEGUE_AQUI|YOUR_CLIENT|CLIENT_ID/i.test(value));
  }

  function isConnected() {
    return Boolean(accessToken && Date.now() < expiresAt - 30000);
  }

  async function waitForGoogleIdentity(timeoutMs = 12000) {
    const started = Date.now();
    while (!(global.google && global.google.accounts && global.google.accounts.oauth2)) {
      if (Date.now() - started > timeoutMs) {
        throw new Error('No se pudo cargar Google Identity Services. Verifica la conexión a Internet y las políticas de red.');
      }
      await new Promise((resolve) => global.setTimeout(resolve, 120));
    }
    return global.google.accounts.oauth2;
  }

  async function ensureClient() {
    if (!isConfigured()) {
      throw new Error('La conexión con Google Drive aún no está configurada. Completa googleOAuthClientId en js/config.js.');
    }
    if (tokenClient) return tokenClient;
    const oauth2 = await waitForGoogleIdentity();
    tokenClient = oauth2.initTokenClient({
      client_id: config.googleOAuthClientId,
      scope: config.googleDriveScope || 'https://www.googleapis.com/auth/drive.readonly',
      callback: function () {}
    });
    return tokenClient;
  }

  async function connect(options = {}) {
    if (isConnected()) return { connected: true, expiresAt };
    const client = await ensureClient();

    return new Promise((resolve, reject) => {
      client.callback = (response) => {
        if (!response || response.error) {
          reject(new Error(response && response.error_description ? response.error_description : 'Google no autorizó el acceso a Drive.'));
          return;
        }
        accessToken = response.access_token;
        const seconds = Number(response.expires_in || 3600);
        expiresAt = Date.now() + seconds * 1000;
        resolve({ connected: true, expiresAt });
      };
      client.error_callback = (error) => reject(new Error(error && error.message ? error.message : 'Se cerró o bloqueó la autorización de Google.'));
      client.requestAccessToken({ prompt: options.silent ? '' : 'consent' });
    });
  }

  async function disconnect() {
    const token = accessToken;
    accessToken = '';
    expiresAt = 0;
    if (token && global.google && global.google.accounts && global.google.accounts.oauth2) {
      await new Promise((resolve) => {
        try {
          global.google.accounts.oauth2.revoke(token, resolve);
        } catch (error) {
          resolve();
        }
      });
    }
  }

  function normalizeFolderId(value) {
    const text = String(value || '').trim();
    if (!text) return '';
    const folderMatch = text.match(/\/folders\/([a-zA-Z0-9_-]+)/);
    if (folderMatch) return folderMatch[1];
    const idMatch = text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (idMatch) return idMatch[1];
    return text.replace(/[^a-zA-Z0-9_-]/g, '');
  }

  function escapeDriveQuery(value) {
    return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  }

  async function authorizedFetch(url, options = {}) {
    if (!isConnected()) await connect();
    const headers = new Headers(options.headers || {});
    headers.set('Authorization', `Bearer ${accessToken}`);
    const response = await fetch(url, { ...options, headers });
    if (response.status === 401) {
      accessToken = '';
      expiresAt = 0;
      throw new Error('La sesión de Google Drive venció. Conéctate nuevamente.');
    }
    if (!response.ok) {
      let details = '';
      try {
        const body = await response.json();
        details = body && body.error && body.error.message ? body.error.message : '';
      } catch (error) {
        details = await response.text().catch(() => '');
      }
      throw new Error(`Google Drive respondió ${response.status}${details ? `: ${details}` : ''}`);
    }
    return response;
  }

  async function listFolderFiles(folderValue) {
    const folderId = normalizeFolderId(folderValue || config.googleDriveFolderId);
    if (!folderId) throw new Error('Ingresa el enlace o ID de la carpeta de Google Drive.');

    const query = [
      `'${escapeDriveQuery(folderId)}' in parents`,
      'trashed = false',
      "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    ];
    const expected = String(config.reportNameContains || '').trim();
    if (expected) query.push(`name contains '${escapeDriveQuery(expected)}'`);

    const files = [];
    let pageToken = '';
    const pageSize = Math.min(1000, Math.max(1, Number(config.maxDriveFiles || 500)));

    do {
      const params = new URLSearchParams({
        q: query.join(' and '),
        pageSize: String(pageSize),
        orderBy: 'modifiedTime desc',
        spaces: 'drive',
        includeItemsFromAllDrives: 'true',
        fields: 'nextPageToken,files(id,name,size,modifiedTime,mimeType,parents,driveId,webViewLink)'
      });
      if (pageToken) params.set('pageToken', pageToken);
      const response = await authorizedFetch(`https://www.googleapis.com/drive/v3/files?${params.toString()}`);
      const body = await response.json();
      (body.files || []).forEach((file) => files.push(file));
      pageToken = body.nextPageToken || '';
    } while (pageToken && files.length < Number(config.maxDriveFiles || 500));

    return { folderId, files };
  }

  async function downloadFile(metadata) {
    if (!metadata || !metadata.id) throw new Error('El archivo de Drive no tiene identificador.');
    const params = new URLSearchParams({ alt: 'media', supportsAllDrives: 'true' });
    const response = await authorizedFetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(metadata.id)}?${params.toString()}`);
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength < 1000) throw new Error(`El archivo ${metadata.name} está vacío o incompleto.`);
    return new File([buffer], metadata.name, {
      type: metadata.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      lastModified: metadata.modifiedTime ? new Date(metadata.modifiedTime).getTime() : Date.now()
    });
  }

  async function downloadFiles(files, onProgress) {
    const output = [];
    for (let index = 0; index < files.length; index += 1) {
      if (typeof onProgress === 'function') onProgress(index, files.length, files[index]);
      output.push(await downloadFile(files[index]));
    }
    if (typeof onProgress === 'function') onProgress(files.length, files.length, null);
    return output;
  }

  function chooseLatestPerCompany(files, detectCompany) {
    const selected = new Map();
    (files || []).forEach((file) => {
      let company;
      try {
        company = detectCompany(file.name, null);
      } catch (error) {
        return;
      }
      const existing = selected.get(company.code);
      const modified = new Date(file.modifiedTime || 0).getTime();
      const existingModified = existing ? new Date(existing.modifiedTime || 0).getTime() : -1;
      if (!existing || modified > existingModified) selected.set(company.code, file);
    });
    return [...selected.values()];
  }

  global.DriveConnector = Object.freeze({
    isConfigured,
    isConnected,
    connect,
    disconnect,
    normalizeFolderId,
    listFolderFiles,
    downloadFile,
    downloadFiles,
    chooseLatestPerCompany,
    getDefaultFolderId: () => String(config.googleDriveFolderId || '')
  });
})(typeof window !== 'undefined' ? window : globalThis);
