# Privacidad y tratamiento de archivos

## Procesamiento local

Los archivos seleccionados se leen en la memoria del navegador. No se cargan a GitHub ni a un servidor de la aplicación.

## Procesamiento desde Google Drive

El usuario autoriza temporalmente acceso de solo lectura mediante Google Identity Services. Los archivos seleccionados se descargan directamente desde Google Drive a la memoria del navegador y se procesan con el mismo motor local.

La aplicación:

- No almacena tokens OAuth en el repositorio.
- No utiliza secretos de cliente.
- No solicita permisos de escritura en Drive.
- No guarda reportes en servidores propios.
- No incorpora los archivos de cartera al repositorio de GitHub.

El token de acceso es temporal y se mantiene solo durante la sesión activa de la página.

## PDF y Excel

Los Excel estándar y PDF Enterprise se generan en el navegador. La versión PDF es nativa y no se crea un archivo HTML intermedio.

## Responsabilidad institucional

El administrador debe configurar la aplicación OAuth, restringir su audiencia, revisar los permisos solicitados y asegurar que únicamente usuarios autorizados accedan a la carpeta de reportes.
