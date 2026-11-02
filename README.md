# Dashboard de Checklist de Vapor y Condensado — v1.2

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
5. Entrar en **Configuración del proyecto → Propiedades de secuencia de comandos** y crear `WRITE_KEY` = contraseña larga para cargas, **mínimo 16 caracteres**. Ya no se usa `READ_KEY` para consultar. Si la propiedad antigua existe, puede permanecer guardada.
6. En la lista de funciones del editor, seleccionar `validarConfiguracion` y pulsar **Ejecutar**. Autorizar el acceso a la hoja cuando Google lo solicite. La función comprueba la pestaña, los encabezados y la presencia de ambas propiedades sin escribir datos.
7. Si todavía no existe la implementación, pulsar **Implementar → Nueva implementación → Tipo: Aplicación web**. Seleccionar *Ejecutar como: Yo*. Para permitir la consulta sin iniciar sesión, configurar el acceso como *Cualquier usuario* (si la política de la organización lo permite), y copiar la URL que termina en `/exec`. Si ya existe, usar **Implementar → Administrar implementaciones → Editar (lápiz) → Versión: Nueva versión → Implementar**, sin crear una URL nueva.

**Privacidad:** las consultas del historial se permiten por POST sin contraseña, mientras que cualquier carga o modificación requiere `WRITE_KEY`. Esta última nunca se guarda en GitHub ni en el navegador. **Importante:** si se habilita acceso público al Web App, la información de verificaciones (operadores, observaciones, desviaciones, etc.) también puede ser consultada por personas externas que conozcan la dirección. Utilizar esta modalidad solo para información autorizada como de consulta general; si se necesita confidencialidad, aplicar autenticación empresarial.

### 2. Publicar el frontend en un repositorio distinto

1. En GitHub, crear un repositorio **nuevo**, por ejemplo, `checklist-vapor-condensado`. No copiar archivos al repositorio anterior `preinca/dashboard`.
2. Extraer este ZIP y cargar **los archivos del interior** a la **raíz** de la rama `main`. Deben quedar `index.html`, `styles.css`, `app.js` y `config.js` directamente en la raíz. No subir el ZIP como un archivo sin extraer.
3. En el nuevo repositorio, abrir **Settings → Pages → Build and deployment → Deploy from a branch**, elegir `main` y `/(root)`, y guardar.
4. Configurar la URL `/exec` en `config.js` (ya está configurada en el repositorio `preinca/ChecklistSyU`). No agregar allí contraseñas.
5. Abrir GitHub Pages: los registros se consultan automáticamente sin contraseña. Pulsar **↻ Actualizar** para volver a consultar Google Sheets en cualquier momento.
6. Descargar el formato original con **Formato Excel**. Completar un checklist real autorizado o de prueba, usar **Cargar checklist**, introducir `WRITE_KEY` y comprobar que los registros aparecen en la hoja y en el histórico. Eliminar datos de prueba de la nueva hoja antes de comenzar el uso operativo.

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
- `config.js`: URL `/exec` del Apps Script, configurada en el repositorio. No contiene contraseñas.
- `apps_script_backend/Code.gs`: backend dedicado a la nueva hoja.
- `referencias_condensado/`: directorio de imágenes, inicialmente vacío.
- `herramientas/descargar_referencias.py`: descarga opcional de las imágenes originales.
- `.nojekyll`: evita que GitHub Pages procese los archivos mediante Jekyll.

## Limitaciones actuales

**Estado actual:** el frontend de `preinca/ChecklistSyU` está publicado en GitHub Pages. El backend actualizado `apps_script_backend/Code.gs` debe copiarse al proyecto de Apps Script y publicarse como **nueva versión de la implementación existente**. Mientras siga desplegado el backend anterior, la consulta sin contraseña fallará. Comprobar la lectura real y la carga con `WRITE_KEY` después del despliegue.

**Histórico:** la pestaña `LISTA_VERIFICADOS` del archivo original que se pudo consultar estaba vacía; si hay historial en otra hoja, Excel o respaldo, habrá que identificarlo y migrarlo sin suponer registros inexistentes.
