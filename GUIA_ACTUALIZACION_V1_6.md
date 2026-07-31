# Actualización a V1.6 — Escenarios Normal y Proyectado

## 1. Reemplazar la versión anterior

1. Descomprime `COMITE_CARTERA_GITHUB_PAGES_V1_6_ESCENARIOS.zip`.
2. Copia **todo el contenido interno** de la carpeta descomprimida dentro del repositorio local `comite-cartera`.
3. Acepta reemplazar los archivos existentes.
4. Verifica que existan:
   - `index.html`
   - `js/processor.js`
   - `js/projection.js`
   - `js/ooxml.js`
   - `js/app.js`
   - `assets/COMITE_BASE.xlsx`

## 2. Publicar

En GitHub Desktop escribe como resumen:

`Agregar escenarios Normal y Proyectado parametrizables`

Después:

1. **Commit to main**.
2. **Push origin**.
3. Espera la marca verde en **Actions**.
4. Abre la aplicación y presiona `Ctrl + F5`.

## 3. Uso

1. Carga únicamente los archivos originales de CTH, F12, F8 y F11.
2. Configura:
   - Horizonte en días.
   - Umbral de migración.
   - Activación o desactivación de la reclasificación.
3. Presiona **Procesar escenarios**.
4. Revisa las pestañas:
   - Normal.
   - Proyectado.
   - Comparativo.
5. Descarga cada Excel con su botón independiente.

## 4. Configuración validada para el ejemplo

- Horizonte: `8`.
- Umbral: `60`.
- Reclasificación: activada.
- Fecha base detectada: `23-07-2026`.
- Fecha objetivo calculada: `31-07-2026`.

Con esos parámetros, el motor reproduce exactamente los cuatro archivos proyectados entregados como referencia.
