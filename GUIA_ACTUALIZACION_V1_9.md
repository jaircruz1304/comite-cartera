# Guía de actualización · Versión 1.9

## 1. Reemplazar la versión anterior

1. Descomprime `COMITE_CARTERA_GITHUB_PAGES_V1_9_DRIVE_PDF.zip`.
2. Copia **todo el contenido interno** de la carpeta descomprimida dentro del repositorio local `comite-cartera`.
3. Acepta reemplazar los archivos existentes.
4. Confirma que existan:

```text
js/config.js
js/drive.js
js/pdf.js
assets/COMITE_BASE.xlsx
```

## 2. Configurar Google Cloud

### 2.1 Crear el proyecto

1. Ingresa en Google Cloud Console.
2. Crea o selecciona un proyecto institucional.
3. En **APIs y servicios → Biblioteca**, habilita **Google Drive API**.

### 2.2 Configurar OAuth

1. En **Google Auth Platform**, configura la pantalla de consentimiento.
2. Para una organización Google Workspace, selecciona audiencia **Interna** cuando esté disponible.
3. Crea un cliente OAuth de tipo **Aplicación web**.
4. En **Orígenes JavaScript autorizados**, registra únicamente el origen, sin la ruta del repositorio:

```text
https://TU-USUARIO.github.io
```

Con dominio personalizado registra, por ejemplo:

```text
https://reportes.cth.fin.ec
```

### 2.3 Colocar el ID de cliente

Edita `js/config.js`:

```javascript
googleOAuthClientId: '123456789-xxxxx.apps.googleusercontent.com'
```

Opcionalmente deja configurada la carpeta:

```javascript
googleDriveFolderId: '1AbCdEf...'
```

No se utiliza ni debe publicarse un secreto de cliente.

## 3. Publicar

En GitHub Desktop registra:

```text
Agregar procesamiento Drive y PDF Enterprise nativo
```

Luego:

1. **Commit to main**.
2. **Push origin**.
3. Espera que GitHub Actions termine correctamente.
4. Abre la aplicación y presiona `Ctrl + F5`.

En el pie debe aparecer:

```text
Versión 1.9.1 · Local + Drive · PDF Enterprise nativo
```

## 4. Probar procesamiento local

1. Selecciona **Procesamiento local**.
2. Carga los cuatro reportes.
3. Procesa.
4. Verifica las cuatro opciones:
   - Excel estándar Normal.
   - PDF Enterprise Normal.
   - Excel estándar Proyectado.
   - PDF Enterprise Proyectado.

Los PDF se generan directamente. No se descarga ni crea un archivo HTML intermedio.

## 5. Probar Google Drive

1. Selecciona **Procesamiento desde Drive**.
2. Presiona **Conectar Drive**.
3. Autoriza el permiso de solo lectura.
4. Pega el enlace o ID de la carpeta.
5. Presiona **Buscar reportes**.
6. Verifica los archivos seleccionados.
7. Presiona **Usar archivos seleccionados**.
8. Procesa normalmente.

## 6. Errores frecuentes

### Drive no configurado

Revisa `googleOAuthClientId` en `js/config.js`.

### Error `origin_mismatch`

El origen de GitHub Pages o dominio personalizado no está registrado exactamente en el cliente OAuth.

### Aplicación no verificada

Mientras el proyecto esté en prueba, agrega los usuarios como testers. Para uso institucional, configura la audiencia interna y coordina los permisos con el administrador de Google Workspace.

### No aparecen archivos

- Confirma el ID de carpeta.
- Verifica que el usuario conectado tenga acceso.
- Los archivos deben ser `.xlsx` y sus nombres deben contener `SaldosDeCarteraSencillo_Report`.

### PDF no disponible

La generación PDF utiliza jsPDF y AutoTable cargados desde CDN. Revisa que la red institucional permita `cdnjs.cloudflare.com` y actualiza con `Ctrl + F5`.

## 7. Alcance de la automatización

La opción Drive automatiza la localización y descarga durante una sesión autorizada. Debido a que GitHub Pages no ejecuta procesos de servidor, no puede iniciar tareas programadas o desatendidas. Esa fase requeriría Apps Script o un servicio backend.
