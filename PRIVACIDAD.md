# Privacidad y tratamiento de archivos

Esta aplicación funciona bajo un modelo **local-first**.

- Los archivos `SaldosDeCarteraSencillo_Report.xlsx` seleccionados por el usuario se leen y procesan en la memoria del navegador.
- Los reportes no se cargan al repositorio de GitHub.
- Los reportes no se envían a un servidor por el código de esta aplicación.
- La plantilla `assets/COMITE_BASE.xlsx` forma parte del repositorio, pero no contiene información histórica ni datos reales de cartera.
- El archivo consolidado se genera localmente y se descarga directamente al equipo del usuario.
- Al recargar o cerrar la página, los archivos cargados dejan de estar disponibles para la aplicación.

No deben subirse al repositorio archivos fuente, consolidaciones generadas ni información financiera institucional.
