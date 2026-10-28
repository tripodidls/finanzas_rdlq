# Finanzas del Proyecto 🍓🎀

Página de finanzas de **La Reina de la Queja** (MJ 2026) para anotar y actualizar las finanzas del proyecto: rifa del peso, venta de remeras y gastos (maquillaje, luces, escenografía, utilería, vestuario y audio).

- `index.html`: la página web. Arriba tiene botones para ir directo a Rifas, Remeras y Gastos.
- `img/`: imágenes del proyecto (afiche, frutilla, remeras, barco de papel). Los datos se guardan en el navegador de quien la usa.
- `datos.json`: datos que ve quien abre la página por primera vez. Para publicar tus números, tocá "💾 Guardar copia de datos" en la página y subí acá el `datos.json` descargado, reemplazando este.
- `Finanzas_Proyecto.xlsx`: versión planilla para importar en Google Sheets.

## Datos compartidos con Google Sheets

Para que todos los que entran al link vean y editen los mismos datos:

1. Creá una planilla nueva y vacía en Google Sheets.
2. En la planilla: **Extensiones → Apps Script**. Borrá lo que aparece y pegá el contenido de `apps-script.gs`. Guardá.
3. Arriba elegí la función `prepararPlanilla` y tocá **Ejecutar**. Aceptá los permisos (tu cuenta de Google).
4. **Implementar → Nueva implementación → tipo: Aplicación web**. Ejecutar como: **Yo**. Quién tiene acceso: **Cualquier persona**. Tocá **Implementar** y copiá la URL que termina en `/exec`.
5. En `index.html`, pegá esa URL en `const URL_SHEETS="";`.
