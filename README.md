# CTH · Comité de Cartera — Versión 1.7.1

Aplicación web local para procesar reportes `SaldosDeCarteraSencillo_Report.xlsx`, generar los escenarios **Normal** y **Proyectado**, comparar resultados y descargar dos consolidaciones Excel independientes.

## Corrección V1.7.1

- El indicador de procesamiento permanece oculto antes de iniciar.
- El spinner se detiene y desaparece al finalizar, tanto en éxito como en error.
- Se reforzó el control con `hidden`, `display` y atributos de accesibilidad para evitar que estilos del navegador mantengan visible el indicador.

## Flujo operativo

1. Configurar horizonte, umbral y reclasificación.
2. Cargar los reportes originales correspondientes a una misma fecha.
3. Procesar ambos escenarios.
4. Descargar los archivos Normal y Proyectado.

## Cambios de interfaz V1.7

- Parametrización ubicada antes de la carga y el procesamiento.
- Logo institucional CTH integrado localmente en `assets/logo_CTH.png`.
- Identificación de la desarrolladora: **Lizbeth Sanipatín**.
- Indicadores del sistema compactos y secundarios.
- Eliminación de la sección visible de validaciones y alertas.
- Mensajes de procesamiento simplificados para el usuario final.
- Privacidad presentada en una sola línea discreta.
- Panel ejecutivo reorganizado para aprovechar mejor el espacio.

## Reglas funcionales conservadas

- Procesamiento independiente por empresa.
- Plantilla COMITE integrada y utilizada como base limpia.
- No se sobrescribe la plantilla original.
- Valores monetarios del HTML expresados en miles y sin decimales.
- Excel con precisión contable.
- No Devenga calculado únicamente con sus columnas específicas.
- Castigados excluidos de la cartera activa.
- Escenario Proyectado parametrizable sin alterar el escenario Normal.

## Privacidad

Los archivos seleccionados se procesan localmente en la memoria del navegador. No se transmiten a GitHub ni a servidores externos.

## Desarrolladora

**Lizbeth Sanipatín**
