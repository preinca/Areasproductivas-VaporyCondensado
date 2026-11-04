# Activación del checklist de refinerías

La interfaz de DE SMET y TIUNA y el formato Excel `FORMATO_REFINERIAS_CHECKLIST.xlsx` ya se encuentran en GitHub.

## Estado de la base
- Google Sheet: `Base_Maestra_Checklist_Vapor_Condensado`
- Pestaña original de áreas productivas: `LISTA_VERIFICADOS` (sin cambios).
- Nueva pestaña de refinerías: `LISTA_REFINERIAS` (con los mismos 16 encabezados, independiente de producción).

## Paso pendiente: actualizar despliegue de Apps Script

1. Abrir el proyecto Google Apps Script del dashboard que utiliza el enlace `/exec` de `config.js`.
2. Sustituir el contenido de `Code.gs` por la versión de `apps_script_backend/Code.gs` de este repositorio.
3. Conservar la propiedad del proyecto `WRITE_KEY` existente; **no publicarla ni ponerla en GitHub**.
4. En Apps Script elegir **Implementar → Gestionar implementaciones**.
5. Editar la implementación del tipo aplicación web utilizada por el dashboard y seleccionar **Nueva versión**.
6. Guardar/implementar y conservar el mismo URL `/exec`. No modificar `config.js` si se actualiza la implementación existente.

## Pruebas

- Abrir **DE SMET** y **TIUNA**; pulsar **Actualizar**: ambos historiales deben consultar solo la pestaña `LISTA_REFINERIAS`.
- Descargar `FORMATO_REFINERIAS_CHECKLIST.xlsx`, completar supervisor, operador, día, turno y al menos una condición con OK / NO OK / N/A.
- Cargar el archivo desde una de las refinerías: solicitará `WRITE_KEY` y una confirmación con el desglose de ambas hojas.
- Verificar que los registros aparecen en la refinería correspondiente, incluido el enlace al diagrama asociado.
- Verificar que **Panel General** y las cinco **Áreas productivas** mantienen su historial en `LISTA_VERIFICADOS`.

El módulo de interfaz `refinerias.js` tiene su propio lector Excel y no cambia el código existente de `app.js`. Las funciones nuevas del backend son `read_refineries` y `upsert_refineries`. Las funciones antiguas `read_verified` y `upsert_verified` se conservan.
