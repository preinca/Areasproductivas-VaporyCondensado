# Dashboard de Checklist de Vapor y Condensado — v1.1

Aplicación independiente del dashboard de Servicios de Planta, creada con HTML, CSS, JavaScript, Google Apps Script, Google Sheets y GitHub Pages.

## Base de datos ya creada

- **Archivo:** [Base_Maestra_Checklist_Vapor_Condensado](https://docs.google.com/spreadsheets/d/1wk8lgVDfJdKgTKo0ZGznBBhVw_pmPmu30r41G5aA9FY/edit)
- **Pestaña:** `LISTA_VERIFICADOS` (16 columnas).
- **Estado al 02-11-2026:** solo encabezados; sin registros históricos en la Base Maestra original consultada. Tampoco hay registros en el `verificados_fallback.js` del repositorio anterior. No se añadieron datos ficticios.
- La Base Maestra original `preinca/dashboard` y sus archivos **no se han modificado**.

## Para ponerlo en marcha (pasos necesarios en tus cuentas)

### 1. Crear y autorizar el Apps Script

1. Abrir [Google Apps Script](https://script.google.com/) **con la cuenta que puede editar la nueva Google Sheet**.
2. Elegir **Nuevo proyecto** y ponerle el nombre `Backend Checklist Vapor y Condensado`.
3. Abrir `apps_script_backend/Code.gs` de este paquete, copiar **todo** y sustituir el contenido inicial de `Código.gs` del proyecto.
4. Guardar el proyecto.
5. Entrar en **Configuración del proyecto → Propiedades de secuencia de comandos** y crear las dos propiedades siguientes:
   - `READ_KEY` = contraseña larga de consulta, **mínimo 16 caracteres**.
   - `WRITE_KEY` = contraseña diferente y larga para cargas, **mínimo 16 caracteres**.
6. En la lista de funciones del editor, seleccionar `validarConfiguracion` y pulsar **Ejecutar**. Autorizar el acceso a la hoja cuando Google lo solicite. La función comprueba la pestaña, los encabezados y la presencia de ambas propiedades sin escribir datos.
7. Pulsar **Implementar → Nueva implementación → Tipo: Aplicación web**. Seleccionar *Ejecutar como: Yo*. Elegir el nivel de acceso conforme a la política interna y copiar la URL final que termina en `/exec`.

**Privacidad:** el backend impide consultar el historial desde GET y requiere `READ_KEY` para leer por POST y `WRITE_KEY` para modificar. Estas contraseñas **no** se almacenan en GitHub ni en el almacenamiento local del navegador. Sin embargo, una contraseña compartida no equivale a autenticación individual, auditoría de usuarios o control corporativo de identidades. Si necesitas acceso estrictamente privado por usuario, no uses una implementación pública con claves compartidas como solución definitiva: conviene implementar autenticación con cuentas Google Workspace antes de cargar registros sensibles.

### 2. Publicar el frontend en un repositorio distinto

1. En GitHub, crear un repositorio **nuevo**, por ejemplo, `checklist-vapor-condensado`. No copiar archivos al repositorio anterior `preinca/dashboard`.
2. Extraer este ZIP y cargar **los archivos del interior** a la **raíz** de la rama `main`. Deben quedar `index.html`, `styles.css`, `app.js` y `config.js` directamente en la raíz. No subir el ZIP como un archivo sin extraer.
3. En el nuevo repositorio, abrir **Settings → Pages → Build and deployment → Deploy from a branch**, elegir `main` y `/(root)`, y guardar.
4. Abrir la URL publicada por GitHub Pages y pulsar **⚙ Configuración**. Pegar la URL `/exec` del Apps Script. La URL se guarda solo en ese navegador; alternativamente puedes fijarla en `config.js` para que todos los dispositivos la conozcan. **Nunca agregar allí las contraseñas**.
5. Pulsar **🔒 Conectar**, introducir `READ_KEY` y comprobar que aparezca `0 registros cargados` o el número real de registros. Después, pulsar **↻ Actualizar** para confirmar.
6. Descargar el formato original con el botón **Formato Excel**. Completar un checklist real autorizado o un archivo de prueba controlada, usar **Cargar checklist**, introducir `WRITE_KEY` y comprobar que los registros aparecen en la hoja y en el histórico. Si se usan datos de prueba, eliminarlos de la nueva hoja antes de comenzar el uso operativo.

### 3. Fotografías de referencia

La interfaz tiene definidas las 23 fotografías utilizadas por la pestaña original. Este paquete incluye la carpeta y un script opcional `herramientas/descargar_referencias.py` para descargarlas del **repositorio público anterior** e incluirlas físicamente en el nuevo proyecto. Requiere conexión a Internet y se ejecuta con `python herramientas/descargar_referencias.py`. Mientras no se descarguen, la interfaz intenta leerlas de `preinca/dashboard` como respaldo temporal.

### 4. Comprobaciones recomendadas

- Filtro de área y fecha, tabla paginada, detalle por equipo y referencias fotográficas.
- Carga de Excel con supervisor, operador, fecha y turno en cada hoja diligenciada.
- Confirmación de escritura, persistencia al actualizar y actualización por mismo ID sin duplicación.
- Prueba en navegador de PC y móvil y verificación de que la hoja original sigue sin cambios.
- Para el uso real, verificar que se cumplen las reglas de acceso y seguridad de la organización.

## Archivos del proyecto

- `index.html`: interfaz web.
- `styles.css`: estilos y adaptación móvil.
- `app.js`: lógica de visualización, filtros, referencias, Excel y conexión con Google.
- `config.js`: URL del Apps Script (vacía por defecto; se puede configurar desde la interfaz).
- `apps_script_backend/Code.gs`: backend dedicado a la nueva hoja.
- `referencias_condensado/`: directorio de imágenes, inicialmente vacío.
- `herramientas/descargar_referencias.py`: descarga opcional de las imágenes originales.
- `.nojekyll`: evita que GitHub Pages procese los archivos mediante Jekyll.

## Limitaciones actuales

**No está desplegada todavía** ni la aplicación web de Apps Script ni una página GitHub Pages nueva. Estas operaciones necesitan realizarse en sus respectivas interfaces con tu cuenta; el archivo ZIP no las publica por sí solo. El resultado de las pruebas con datos simulados no sustituye a una carga y lectura reales.

**Histórico:** la pestaña `LISTA_VERIFICADOS` del archivo original que se pudo consultar estaba vacía; si hay historial en otra hoja, Excel o respaldo, habrá que identificarlo y migrarlo sin suponer registros inexistentes.
