# Guía rápida de implementación

## Paso 1. Crear el repositorio

1. Ingresa a GitHub.
2. Selecciona **New repository**.
3. Nombre sugerido: `comite-cartera`.
4. Selecciona **Public** si utilizas GitHub Free.
5. Crea el repositorio sin subir archivos Excel reales.

## Paso 2. Subir la aplicación

1. Descarga y descomprime `COMITE_CARTERA_GITHUB_PAGES_V1.zip`.
2. En el repositorio, selecciona **Add file → Upload files**.
3. Arrastra todo el contenido de la carpeta descomprimida, no la carpeta contenedora.
4. Verifica que `index.html` quede en la raíz.
5. Confirma el commit en `main`.

## Paso 3. Activar GitHub Pages

1. Entra en **Settings → Pages**.
2. En **Build and deployment**, selecciona **GitHub Actions**.
3. Abre la pestaña **Actions**.
4. Espera que `Deploy to GitHub Pages` finalice en verde.
5. Abre la URL publicada, normalmente:

```text
https://TU-USUARIO.github.io/comite-cartera/
```

## Paso 4. Primera prueba

1. Carga la plantilla `COMITE.xlsx`.
2. Carga los cuatro archivos de cartera de una misma fecha.
3. Pulsa **Procesar cartera**.
4. Revisa el panel y las validaciones.
5. Descarga el consolidado.
6. Comprueba las hojas `CONTROL_EJECUCION`, `CTH`, `F12`, `F8`, `F11` y `ORIGEN_*`.

## Regla de seguridad

No subas al repositorio la plantilla, los archivos fuente ni los resultados. El repositorio debe contener solamente el código incluido en el paquete.
