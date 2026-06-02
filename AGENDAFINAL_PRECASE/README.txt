ARCHIVOS CORREGIDOS - Agendas PCR-CS

1. En Apps Script, crea/actualiza el archivo HTML llamado exactamente: Index
   - Pega el contenido de Index.html

2. En Apps Script, crea/actualiza el archivo script llamado: Code.gs
   - Pega el contenido de Code.gs

3. IMPORTANTE:
   En Code.gs, pega el ID de tu Google Sheet aquí:

   const SPREADSHEET_ID = '';

   Ejemplo:
   const SPREADSHEET_ID = '1ABCDEF123456789';

4. Vuelve a desplegar la Web App en Apps Script:
   Implementar > Nueva implementación > Aplicación web

5. Abre la URL de la Web App, no el HTML local.

Qué corrige:
- Guardado de citas en la hoja AGENDA PCR-CS.
- Visualización inmediata en agenda después de guardar.
- Click sobre la cita para ver detalle.
- Exportación a la hoja PLATAFORMA.
- Error claro si falta configurar la Sheet.
