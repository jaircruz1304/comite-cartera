# Actualización GitHub Pages — Comité de Cartera v1.4

## Cambios incorporados

1. Ya no se carga manualmente la plantilla COMITE.
2. La plantilla se encuentra dentro del repositorio en `assets/COMITE_BASE.xlsx`.
3. El usuario carga únicamente los archivos `SaldosDeCarteraSencillo_Report.xlsx`.
4. El Excel se genera como binario `ArrayBuffer`, se valida internamente y luego se convierte a descarga `.xlsx`.
5. La descarga solo se habilita cuando el archivo supera la verificación de apertura y contiene todas las hojas requeridas.
6. El panel HTML muestra valores monetarios en miles, redondeados y sin decimales.
7. Los porcentajes del HTML también se muestran sin decimales.
8. El Excel conserva los importes en miles con dos decimales para el detalle contable.

## Cómo actualizar el repositorio

1. Descomprime `COMITE_CARTERA_GITHUB_PAGES_V1_4.zip`.
2. Copia todo el contenido descomprimido dentro de la carpeta local de tu repositorio `comite-cartera`.
3. Acepta reemplazar los archivos existentes.
4. Confirma que exista este archivo:

```text
assets/COMITE_BASE.xlsx
```

5. En GitHub Desktop escribe como resumen:

```text
Integrar plantilla y validar descarga Excel
```

6. Presiona **Commit to main**.
7. Presiona **Push origin**.
8. Espera que la acción `Deploy to GitHub Pages` termine con marca verde.
9. Abre la aplicación publicada y presiona `Ctrl + F5`.

## Prueba

1. Carga los reportes de CTH, F12, F8 y F11 de la misma fecha.
2. Presiona **Procesar cartera**.
3. Espera el mensaje:

```text
Proceso completado y archivo Excel validado.
```

4. Presiona **Descargar consolidado**.
5. Abre el archivo en Microsoft Excel.

No se debe subir al repositorio ningún archivo de cartera real ni consolidado generado.
