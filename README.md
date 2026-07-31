# Comité de Cartera — GitHub Pages v1.4

Aplicación web local-first para procesar archivos `SaldosDeCarteraSencillo_Report.xlsx` y generar un consolidado Excel conforme a las reglas del proyecto **INFORMES CTH 2026**.

## Cambio principal de esta versión

La plantilla ya no se carga manualmente. Se encuentra integrada en el repositorio en:

```text
assets/COMITE_BASE.xlsx
```

El usuario únicamente selecciona los reportes de saldos de cartera. En cada ejecución, la aplicación abre una copia nueva de la plantilla integrada, completa los resultados actuales y genera un archivo distinto. La plantilla del repositorio nunca se sobrescribe.

## Flujo de uso

1. Abrir la aplicación publicada en GitHub Pages.
2. Seleccionar los archivos de CTH, F12, F8 y F11 correspondientes a una misma fecha.
3. Presionar **Procesar cartera**.
4. Revisar KPIs, validaciones y resumen por empresa.
5. Presionar **Descargar consolidado**.

## Reglas funcionales implementadas

- Procesamiento independiente por empresa.
- Exclusión de filas de total, subtotales y duplicados exactos.
- Cartera castigada separada de la cartera activa.
- Cartera vencida calculada desde sus columnas específicas.
- No Devenga calculado exclusivamente desde la agrupación correspondiente.
- `NOR` y `GRA` clasificados como NORMAL.
- `RE` clasificado como REESTRUCTURADA.
- Segmentos 60–90 y +90 calculados con `Dias Morosidad`.
- Bienes en Pago, Provisiones BEP y Neto BEP en blanco cuando no existe información.
- Hojas `ORIGEN_*` con los datos actuales de cada archivo.
- Hoja `CONTROL_EJECUCION` para auditoría.

## Presentación de cifras

En el panel HTML:

- Todos los importes monetarios se dividen para 1.000.
- Se muestran redondeados a números enteros, sin decimales.
- Las operaciones se muestran como números enteros.
- Los porcentajes se muestran sin decimales.

En el archivo Excel:

- Los importes siguen expresados en miles.
- Se conservan dos decimales para el detalle contable.
- Las operaciones se mantienen como enteros.

## Validación del archivo descargado

Antes de habilitar la descarga, la aplicación:

1. Genera el libro como `ArrayBuffer`.
2. Verifica que sea un paquete XLSX válido.
3. Confirma la presencia de `workbook.xml` y las relaciones internas.
4. Intenta abrirlo nuevamente con el lector Excel del navegador.
5. Comprueba que existan las hojas COMITE, ORIGEN y CONTROL requeridas.
6. Solo después crea el archivo descargable.

Esto evita entregar un archivo incompleto o dañado.

## Estructura del repositorio

```text
comite-cartera
├── .github/
│   └── workflows/
│       └── deploy.yml
├── assets/
│   ├── COMITE_BASE.xlsx
│   └── styles.css
├── js/
│   ├── app.js
│   └── processor.js
├── tests/
│   └── processor-smoke.js
├── .nojekyll
├── index.html
├── README.md
├── PRIVACIDAD.md
├── LICENSE
├── VERSION.txt
└── package.json
```

## Actualización en GitHub Desktop

1. Copiar todo el contenido de esta versión sobre la carpeta local del repositorio.
2. Aceptar el reemplazo de archivos.
3. Verificar que `assets/COMITE_BASE.xlsx` esté incluido.
4. Registrar el commit:

```text
Integrar plantilla y validar descarga Excel
```

5. Presionar **Commit to main**.
6. Presionar **Push origin**.
7. Esperar que GitHub Actions termine correctamente.
8. Recargar la página con `Ctrl + F5`.

## Privacidad

Los reportes seleccionados se procesan en la memoria del navegador. No se cargan al repositorio ni se envían a un servidor por esta aplicación. La plantilla integrada no contiene reportes reales ni información histórica de cartera.
