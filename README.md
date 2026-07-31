# Comité de Cartera – GitHub Pages

Aplicación web estática para procesar archivos `SaldosDeCarteraSencillo_Report.xlsx` utilizando una copia limpia de la plantilla `COMITE.xlsx`.

## Qué hace

- Lee la plantilla y los archivos de cartera directamente en el navegador.
- No sobrescribe la plantilla original.
- Procesa cada empresa por separado: CTH, F12, F8 y F11.
- Excluye filas de total, subtotales, encabezados, vacíos y duplicados exactos.
- Separa cartera castigada de la cartera activa.
- Calcula No Devenga exclusivamente con las columnas de esa agrupación.
- Clasifica `NOR` y `GRA` como NORMAL, y `RE` como REESTRUCTURADA.
- Usa `Dias Morosidad` para los segmentos 60–90 y +90.
- Deja en blanco Cartera Castigada y Bienes en Pago cuando no hay información.
- Genera hojas `ORIGEN_*` y una hoja `CONTROL_EJECUCION`.
- Descarga un archivo nuevo: `COMITE_Consolidado_AAAAMMDD_FINAL.xlsx`.

## Privacidad

La aplicación es **local-first**. Los archivos seleccionados se procesan en la memoria del navegador. Este proyecto no incluye código para cargar los Excel a un servidor, a GitHub ni a una base de datos.

No subas al repositorio:

- La plantilla real.
- Archivos de cartera.
- Consolidaciones generadas.
- Datos personales, financieros o confidenciales.

## Estructura

```text
comite-cartera-github-pages/
├── .github/
│   └── workflows/
│       └── deploy.yml
├── assets/
│   └── styles.css
├── js/
│   ├── app.js
│   └── processor.js
├── .nojekyll
├── index.html
└── README.md
```

## Publicación en GitHub Pages

### 1. Crear el repositorio

1. Inicia sesión en GitHub.
2. Crea un repositorio nuevo.
3. Nombre sugerido: `comite-cartera`.
4. Para una cuenta GitHub Free, utiliza un repositorio público.
5. No agregues archivos Excel reales.

### 2. Subir el contenido

1. Descomprime el paquete entregado.
2. En el repositorio, selecciona **Add file → Upload files**.
3. Sube todos los archivos y carpetas conservando la estructura.
4. Confirma el commit en la rama `main`.

### 3. Activar GitHub Pages

1. Abre **Settings** del repositorio.
2. Entra en **Pages**.
3. En **Build and deployment → Source**, selecciona **GitHub Actions**.
4. Abre la pestaña **Actions** y espera que finalice `Deploy to GitHub Pages`.
5. La dirección tendrá el formato:

```text
https://TU-USUARIO.github.io/comite-cartera/
```

Documentación oficial:

- https://docs.github.com/en/pages/getting-started-with-github-pages
- https://docs.github.com/en/get-started/start-your-journey/deploying-your-website-automatically

## Uso diario

1. Abre la dirección publicada.
2. Carga la plantilla `COMITE.xlsx` original.
3. Carga los archivos `SaldosDeCarteraSencillo_Report` del mismo corte.
4. Pulsa **Procesar cartera**.
5. Revisa el dashboard y las validaciones.
6. Pulsa **Descargar consolidado**.
7. Abre el Excel y verifica la hoja `CONTROL_EJECUCION`.

## Requisitos de los archivos

### Plantilla

Debe contener estas hojas:

- `CTH`
- `F12`
- `F8`
- `F11`

### Archivos fuente

Deben contener:

- `#`
- `CaliF.Cont.`
- `Dias Morosidad`
- `Cartera Castigada`
- Agrupación `Cartera Por Vencer`
- Agrupación `Cartera que no devenga Intereses`
- Agrupación `Cartera Vencida`

Todos los archivos cargados en una ejecución deben corresponder a la misma fecha de corte.

## Dependencia externa

La aplicación carga `xlsx-populate` 1.21.0 desde UNPKG mediante una referencia con integridad SRI. Esta biblioteca se usa porque trabaja sobre el contenido XML del libro y está orientada a conservar las características y estilos existentes de una plantilla.

Proyecto y documentación:

- https://github.com/dtjohnson/xlsx-populate
- https://www.npmjs.com/package/xlsx-populate

## Validación realizada

El motor de cálculo fue contrastado con los cuatro archivos del corte del 29 de julio de 2026 y reproduce los resultados del consolidado previamente validado.

## Límites de esta versión

- Requiere conexión a Internet para cargar la biblioteca `xlsx-populate` desde UNPKG.
- No almacena históricos entre sesiones.
- No tiene autenticación propia.
- GitHub Pages es un alojamiento estático; la aplicación no ejecuta un servidor.
- Las hojas ORIGEN conservan los valores y la estructura tabular, pero no necesariamente todos los elementos visuales del archivo fuente.

## Próximas fases sugeridas

1. Comparación automática entre dos fechas de corte.
2. Histórico descargable en JSON o Excel.
3. Configuración editable de umbrales de semáforos.
4. Modo sin conexión con la biblioteca incluida dentro del repositorio.
5. Pruebas automáticas de regresión con archivos anonimizados.

## Corrección 1.2.0 — compatibilidad con celdas vacías

Esta versión incorpora una normalización automática de ciertas celdas vacías que Excel puede guardar como `inlineStr` sin contenido. Ese patrón es válido para Excel, pero provoca en `xlsx-populate 1.21.0` el error `Cannot read properties of undefined (reading 'children')`. La aplicación corrige únicamente ese detalle interno en la copia temporal cargada en memoria, sin sobrescribir ni modificar la plantilla original.
