# Comité de Cartera — Escenarios Normal y Proyectado

Aplicación estática para GitHub Pages que procesa localmente archivos `SaldosDeCarteraSencillo_Report.xlsx` y genera dos escenarios independientes a partir de una sola carga:

- **Normal:** aplica las reglas vigentes de Comité de Cartera.
- **Proyectado:** incrementa `Dias Morosidad` y aplica una reclasificación parametrizable.
- **Comparativo:** muestra variaciones por empresa y consolidado.

Los archivos de cartera se procesan en la memoria del navegador. No se envían a GitHub ni a un servidor.

## Parámetros por defecto validados

- Fecha base de referencia: **23-07-2026**.
- Horizonte: **8 días**.
- Fecha objetivo: **31-07-2026**.
- Umbral: **días proyectados > 60**.
- Reclasificación activada.

## Regla de proyección

Para cada fila real de detalle:

1. `Dias Morosidad proyectado = Dias Morosidad original + horizonte`.
2. Si la operación está únicamente en **Cartera Por Vencer**, no tiene saldo previo en No Devenga ni Vencida, no está castigada y los días proyectados son mayores al umbral:
   - Por Vencer 1–30 → Vencida 1–30.
   - Por Vencer 31–90 → No Devenga 1–30.
   - Por Vencer 91–180 → No Devenga 31–90.
   - Por Vencer 181–360 → No Devenga 91–180.
   - Por Vencer +360 → No Devenga 181–360.
3. Operaciones que ya están en No Devenga o Vencida conservan sus saldos y solo actualizan días.
4. El Total Cartera y la Cartera Castigada no se alteran.

La regla fue contrastada contra los cuatro archivos proyectados de referencia: **1.276 operaciones**, **37 reclasificaciones** y **0 diferencias** en días, saldos e indicadores.

## Salidas

- `COMITE_Normal_YYYYMMDD_FINAL.xlsx`
- `COMITE_Proyectado_YYYYMMDD_BASE_YYYYMMDD_FINAL.xlsx`

Cada archivo parte de `assets/COMITE_BASE.xlsx`, mantiene hojas independientes por empresa, agrega `CONTROL_EJECUCION` y genera hojas `ORIGEN_*`.

## Publicación

1. Copiar todo el contenido de esta carpeta a la raíz del repositorio.
2. Confirmar que `index.html` y `assets/COMITE_BASE.xlsx` estén en la raíz y subcarpeta indicadas.
3. Hacer commit y push a `main`.
4. Configurar **Settings → Pages → Source: GitHub Actions**.
5. Esperar que la acción `Deploy to GitHub Pages` termine correctamente.
6. Abrir la página y presionar `Ctrl + F5` después de actualizar.

## Pruebas

```bash
npm test
```

Las pruebas verifican reglas normales, proyección, generación OOXML, plantilla integrada y estructura del repositorio.
