(function (global) {
  'use strict';

  global.APP_CONFIG = Object.freeze({
    version: '1.9.1',

    // Reemplaza este valor con el ID de cliente OAuth 2.0 tipo "Aplicación web"
    // creado en Google Cloud Console. El ID de cliente no es una contraseña.
    googleOAuthClientId: 'PEGUE_AQUI_SU_CLIENT_ID.apps.googleusercontent.com',

    // Carpeta predeterminada de reportes CTH. El usuario puede cambiarla desde la interfaz.
    googleDriveFolderId: '1mJsSSYjGYcYu26YGX0dCGXS7ubLFsqpQ',

    // drive.readonly permite localizar y descargar automáticamente reportes dentro
    // de la carpeta indicada. No se solicitan permisos de escritura.
    googleDriveScope: 'https://www.googleapis.com/auth/drive.readonly',

    reportNameContains: 'SaldosDeCarteraSencillo_Report',
    maxDriveFiles: 500
  });
})(typeof window !== 'undefined' ? window : globalThis);
