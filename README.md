# Comité de Cartera CTH — Versión 2.0

Solución GitHub Pages para procesamiento de cartera, escenarios Normal y Proyectado, Excel estándar, PDF Enterprise y análisis histórico.

## Módulos

- `index.html`: procesamiento operativo local o desde GitHub.
- `historico.html`: filtros, tendencias, comparaciones y exportaciones históricas.
- `data/historico/`: catálogo normalizado de julio de 2013 a julio de 2026.
- `data/operational/`: indicadores generados automáticamente por GitHub Actions desde `sources/current/`.

## Principios

- La plantilla COMITE integrada nunca se sobrescribe.
- Las reglas de negocio actuales permanecen intactas.
- GitHub tiene prioridad para datos operativos.
- El histórico consolidado complementa períodos sin detalle.
- No se inventa información faltante.
- Valores HTML/PDF: miles, sin decimales.
- Excel: precisión contable.

Consulte `GUIA_IMPLEMENTACION_V2.md`, `SEGURIDAD_DATOS.md` e `INFORME_NORMALIZACION_HISTORICA.md`.
