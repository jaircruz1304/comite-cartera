# CTH · Comité de Cartera · Versión 1.9

Aplicación estática para GitHub Pages con dos modalidades de origen y dos escenarios de cálculo:

- **Procesamiento local:** carga manual de archivos `.xlsx` desde el navegador.
- **Procesamiento desde Google Drive:** conexión OAuth de solo lectura, búsqueda en una carpeta y descarga directa a la memoria del navegador.
- **Escenario Normal:** reglas vigentes del Comité de Cartera.
- **Escenario Proyectado:** horizonte, umbral y reclasificación parametrizables.
- **Excel estándar:** conserva la plantilla COMITE, precisión contable y hojas ORIGEN.
- **PDF Enterprise nativo:** se genera directamente desde JavaScript con tarjetas, gráficos vectoriales y matriz ejecutiva. No crea un HTML intermedio.

## Reglas de negocio

La versión 1.9 no modifica el motor validado de cartera ni la lógica de proyección. La modalidad Drive solamente cambia el origen de los archivos. El PDF Enterprise únicamente presenta los resultados calculados.

## Configuración de Google Drive

1. Crea o selecciona un proyecto en Google Cloud.
2. Habilita **Google Drive API**.
3. Configura la pantalla de consentimiento OAuth.
4. Crea credenciales OAuth 2.0 de tipo **Aplicación web**.
5. Agrega como origen JavaScript autorizado el dominio de GitHub Pages, por ejemplo:
   - `https://TU-USUARIO.github.io`
   - o el dominio personalizado utilizado por la aplicación.
6. Abre `js/config.js` y reemplaza:

```javascript
googleOAuthClientId: 'PEGUE_AQUI_SU_CLIENT_ID.apps.googleusercontent.com'
```

7. Opcionalmente configura una carpeta predeterminada:

```javascript
googleDriveFolderId: 'ID_DE_LA_CARPETA'
```

No coloques un secreto de cliente en GitHub. Una aplicación JavaScript usa únicamente el ID de cliente público.

## Alcance solicitado

La conexión utiliza:

```text
https://www.googleapis.com/auth/drive.readonly
```

El permiso es de solo lectura y permite localizar y descargar automáticamente archivos de la carpeta indicada. El token se mantiene solamente durante la sesión del navegador y no se guarda en el repositorio.

Para uso institucional, se recomienda configurar la aplicación OAuth como **interna** dentro de la organización Google Workspace y aplicar las políticas del administrador.

## Uso

### Modalidad local

1. Selecciona **Procesamiento local**.
2. Configura la proyección.
3. Carga los reportes `.xlsx`.
4. Procesa los escenarios.
5. Descarga Excel Normal, PDF Normal, Excel Proyectado y PDF Proyectado.

### Modalidad Drive

1. Selecciona **Procesamiento desde Drive**.
2. Presiona **Conectar Drive** y autoriza acceso de solo lectura.
3. Pega el enlace o ID de la carpeta.
4. Presiona **Buscar reportes**.
5. Revisa la selección o usa **Seleccionar últimos por empresa**.
6. Presiona **Usar archivos seleccionados**.
7. Procesa y descarga los resultados.

## Limitación de GitHub Pages

GitHub Pages es una plataforma estática. La conexión implementada permite extracción automática **durante una sesión interactiva**, después de que el usuario autoriza Google Drive. No ejecuta procesos desatendidos ni programados en segundo plano. Para automatización sin presencia del usuario se requiere un componente servidor, por ejemplo Apps Script, Cloud Run o una función administrada.

## Privacidad

- En modo local, los archivos permanecen en la memoria del navegador.
- En modo Drive, los archivos viajan directamente de Google Drive al navegador autenticado.
- La aplicación no tiene backend propio y no almacena los reportes en GitHub.
- Nunca subas archivos reales de cartera al repositorio.

## Pruebas

```bash
npm test
```

Se validan el motor Normal, la proyección, el generador OOXML, utilidades de Drive, módulo PDF y estructura del repositorio.


## Carpeta de Drive preconfigurada

La aplicación incluye como carpeta predeterminada el ID `1mJsSSYjGYcYu26YGX0dCGXS7ubLFsqpQ`. El usuario aún debe autorizar el acceso mediante OAuth 2.0 y puede sustituir la carpeta desde la interfaz cuando sea necesario.
