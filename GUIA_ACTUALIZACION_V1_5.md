# Actualización a la versión 1.5

## 1. Descargar y descomprimir

Descomprime `COMITE_CARTERA_GITHUB_PAGES_V1_5_ENTERPRISE.zip`.

## 2. Reemplazar el repositorio local

Copia todo el contenido interno de la carpeta descomprimida dentro de tu repositorio local `comite-cartera` y acepta reemplazar los archivos anteriores.

No copies la carpeta principal como una subcarpeta. `index.html` debe quedar directamente en la raíz del repositorio.

## 3. Verificar archivos esenciales

La estructura debe incluir:

```text
comite-cartera
├── .github/workflows/deploy.yml
├── assets
│   ├── COMITE_BASE.xlsx
│   └── styles.css
├── js
│   ├── app.js
│   ├── ooxml.js
│   └── processor.js
├── vendor
│   └── jszip.min.js
├── tests
├── index.html
├── README.md
└── package.json
```

## 4. Subir a GitHub

En GitHub Desktop:

1. Summary: `Corregir motor Excel y renovar dashboard enterprise`.
2. Presiona **Commit to main**.
3. Presiona **Push origin**.
4. En GitHub, abre **Actions** y espera la marca verde.

## 5. Limpiar caché

Abre la página publicada y presiona:

`Ctrl + F5`

En el pie de página debe aparecer:

`Versión 1.5.0 · Motor OOXML estable`

## 6. Prueba

1. Carga los cuatro archivos de saldos.
2. Presiona **Procesar cartera**.
3. Verifica que aparezcan cinco KPIs y cuatro tarjetas de empresa.
4. Presiona **Descargar consolidado**.
5. Abre el archivo en Excel.

La descarga fue probada con los reportes del 29 de julio de 2026 y contiene las nueve hojas esperadas.
