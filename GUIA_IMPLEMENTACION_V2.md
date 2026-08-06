# Implementación V2.0 — GitHub operativo e histórico

## 1. Configurar el repositorio

Edite `js/config.js`:

```javascript
githubOwner: 'SU_USUARIO',
githubRepo: 'comite-cartera',
githubBranch: 'main',
githubSourcePath: 'sources/current'
```

La modalidad GitHub funciona sin Google Cloud cuando el repositorio y los archivos son públicos.

## 2. Cargar reportes operativos

Coloque los archivos autorizados en:

```text
sources/current/
```

Al realizar `push`, GitHub Actions:

1. Lee los Excel de la carpeta.
2. Aplica el motor de reglas vigente.
3. Genera `data/operational/catalog.json`.
4. Ejecuta las pruebas.
5. Publica GitHub Pages.

La aplicación también puede descargar los Excel directamente desde GitHub y procesarlos en el navegador.

## 3. Consultar el histórico

Abra:

```text
https://SU_USUARIO.github.io/comite-cartera/historico.html
```

El módulo permite seleccionar empresa, período base, comparación de 3, 6 o 12 meses, comparación interanual y rango personalizado. También exporta Excel y PDF directamente.

## 4. Fuentes

- GitHub: fuente operativa prioritaria.
- JSON histórico publicado: datos mensuales normalizados.
- Google Drive: repositorio documental histórico y descarga de los libros originales.

## 5. Seguridad

No coloque reportes reservados en un repositorio público. Para esos archivos utilice la carga local. GitHub Pages no debe contener tokens de acceso a repositorios privados.
