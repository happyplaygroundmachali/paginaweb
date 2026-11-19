/**
 * Happy Playground - backend Google Apps Script
 * 1) Crea una hoja de cálculo y copia su ID.
 * 2) En Extensiones > Apps Script, pega este archivo.
 * 3) Configura SHEET_ID y, si corresponde, CALENDAR_ID.
 * 4) Implementar > Nueva implementación > Aplicación web.
 *    Ejecutar como: Yo. Acceso: Cualquier persona.
 * 5) Copia la URL /exec en CONFIG.appsScriptUrl dentro de index.html.
 *
 * Las solicitudes se registran como PENDIENTES. No se confirma disponibilidad
 * ni se bloquea automáticamente una fecha: el equipo debe revisar y confirmar.
 */
const SHEET_ID = "PEGA_AQUI_EL_ID_DE_TU_GOOGLE_SHEET";
const SHEET_NAME = "Solicitudes";
const CALENDAR_ID = ""; // Opcional: ID de calendario. Vacío = no crea eventos.
const CREATE_PENDING_CALENDAR_EVENT = true;

function doPost(e) {
  try {
    const p = (e && e.parameter) ? e.parameter : {};
    if (!p.nombre || !p.telefono || !p.fecha || !p.consentimiento) {
      throw new Error("Faltan campos obligatorios.");
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(["Fecha de solicitud","Nombre","Teléfono","Correo","Fecha cumpleaños","Edad","Invitados","Mensaje","Estado","Evento Calendar ID"]);
      sh.setFrozenRows(1);
    }
    const now = new Date();
    let calendarEventId = "";
    if (CREATE_PENDING_CALENDAR_EVENT && CALENDAR_ID) {
      const cal = CalendarApp.getCalendarById(CALENDAR_ID);
      if (cal) {
        const parts = String(p.fecha).split("-").map(Number);
        const start = new Date(parts[0], parts[1]-1, parts[2], 12, 0, 0);
        const end = new Date(parts[0], parts[1]-1, parts[2], 13, 0, 0);
        const title = "SOLICITUD PENDIENTE - " + p.nombre;
        const description = [
          "Solicitud recibida desde la web Happy Playground.",
          "Estado: PENDIENTE DE CONFIRMACIÓN",
          "Teléfono: " + p.telefono,
          "Correo: " + (p.email || ""),
          "Edad: " + (p.edad || ""),
          "Invitados: " + (p.invitados || ""),
          "Mensaje: " + (p.mensaje || "")
        ].join("\n");
        const event = cal.createEvent(title, start, end, {description: description});
        event.setColor(CalendarApp.EventColor.YELLOW);
        calendarEventId = event.getId();
      }
    }
    sh.appendRow([now,p.nombre,p.telefono,p.email||"",p.fecha,p.edad||"",p.invitados||"",p.mensaje||"","PENDIENTE",calendarEventId]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
