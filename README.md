# Comité de Cartera – GitHub Pages V1.5

Aplicación web local para procesar reportes `SaldosDeCarteraSencillo_Report`, aplicar las reglas del Comité de Cartera y generar un consolidado Excel a partir de una plantilla integrada en el repositorio.

## Cambios principales de la V1.5

- Se eliminó completamente `XlsxPopulate`, causante del error `Cannot read properties of undefined (reading 'attributes')`.
- Se incorporó un motor OOXML propio basado en JSZip que modifica directamente una copia de la plantilla sin reconstruirla desde cero.
- La descarga se genera y valida antes de habilitar el botón.
- La plantilla está en `assets/COMITE_BASE.xlsx`; el usuario solo carga los reportes de saldos.
- JSZip se incluye localmente en `vendor/jszip.min.js`.
- El lector de reportes también usa OOXML y ya no depende de servicios externos.
- El dashboard fue rediseñado con una interfaz ejecutiva, indicadores circulares, matriz de control y panel de auditoría.
- El HTML presenta cifras monetarias en miles y sin decimales.
- El Excel conserva dos decimales y el formato visual de la plantilla.

## Flujo de uso

1. Abrir la aplicación publicada en GitHub Pages.
2. Cargar uno o varios reportes de saldos correspondientes a una sola fecha.
3. Presionar **Procesar cartera**.
4. Revisar los KPIs, indicadores por empresa y validaciones.
5. Presionar **Descargar consolidado**.

## Archivo generado

El archivo se denomina:

`COMITE_Consolidado_AAAAMMDD_FINAL.xlsx`

Contiene:

- `CTH`
- `F12`
- `F8`
- `F11`
- `CONTROL_EJECUCION`
- Una hoja `ORIGEN_*` por cada empresa cargada

## Publicación

1. Copiar todo el contenido del paquete en la raíz del repositorio.
2. Aceptar el reemplazo de los archivos anteriores.
3. Confirmar que existan:
   - `assets/COMITE_BASE.xlsx`
   - `vendor/jszip.min.js`
   - `js/ooxml.js`
4. Hacer `Commit to main`.
5. Hacer `Push origin`.
6. Esperar que GitHub Actions termine correctamente.
7. Abrir la página y presionar `Ctrl + F5`.

## Seguridad

Los reportes seleccionados se procesan en la memoria del navegador. La plantilla integrada forma parte del repositorio, pero los archivos de cartera no se cargan automáticamente a GitHub.
