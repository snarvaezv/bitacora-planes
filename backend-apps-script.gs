/**
 * BITÁCORA DE PLANES — Backend en Google Apps Script
 * ----------------------------------------------------
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Crea una hoja de cálculo nueva en Google Sheets (puede llamarse "Planes de Pareja").
 * 2. Ve a Extensiones > Apps Script.
 * 3. Borra el contenido de Code.gs que aparece por defecto y pega TODO este archivo.
 * 4. Guarda (ícono de disco) y dale un nombre al proyecto, ej. "API Planes".
 * 5. Click en "Implementar" (Deploy) > "Nueva implementación".
 *    - Tipo: "Aplicación web"
 *    - Ejecutar como: "Yo" (tu cuenta)
 *    - Quién tiene acceso: "Cualquier usuario"
 * 6. Copia la URL que te da ("URL de la aplicación web") — la necesitas para el HTML.
 * 7. La primera vez tendrás que autorizar permisos (pantalla de Google, click en
 *    "Avanzado" > "Ir a [nombre del proyecto] (no seguro)" — es tu propio script, es seguro).
 * 8. Cada vez que edites este código, debes crear una NUEVA implementación (o editar la
 *    existente) para que los cambios se reflejen en la URL pública.
 */

const SHEET_NAME = "Planes";
const MARCADOR_SHEET_NAME = "Marcador";
const PHOTOS_FOLDER_NAME = "Bitácora de Planes - Fotos";
const HEADERS = [
  "id", "titulo", "descripcion", "categoria", "prioridad",
  "costoEstimado", "fechaTentativa", "fechaReal", "estado",
  "calificacion", "nota", "propuestoPor", "fechaCreacion", "fotoUrl"
];
const MARCADOR_HEADERS = ["id", "juego", "ganador", "fecha", "lugar", "nota", "fechaCreacion"];
const MENSAJES_SHEET_NAME = "Mensajes";
const MENSAJES_HEADERS = ["id", "fotoUrl", "frase", "fechaCreacion"];

function doGet(e) {
  return respond({ planes: getAllPlanes(), marcador: getAllMarcador(), mensajes: getAllMensajes() });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    if (action === "add") result = addPlan(body.plan);
    else if (action === "update") result = updatePlan(body.plan);
    else if (action === "delete") result = deletePlan(body.id);
    else if (action === "uploadPhoto") result = uploadPhoto(body.fileData, body.fileName, body.mimeType);
    else if (action === "addMarcador") result = addMarcador(body.entry);
    else if (action === "deleteMarcador") result = deleteMarcador(body.id);
    else if (action === "addMensaje") result = addMensaje(body.mensaje);
    else if (action === "deleteMensaje") result = deleteMensaje(body.id);
    else result = { error: "Acción desconocida: " + action };
    return respond(result);
  } catch (err) {
    return respond({ error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllPlanes() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexado para getRange
  }
  return -1;
}

function addPlan(plan) {
  const sheet = getSheet();
  plan.id = "p_" + new Date().getTime();
  plan.fechaCreacion = new Date().toISOString();
  const row = HEADERS.map(h => (plan[h] !== undefined && plan[h] !== null) ? plan[h] : "");
  sheet.appendRow(row);
  return plan;
}

function updatePlan(plan) {
  const sheet = getSheet();
  const rowIndex = findRowIndexById(sheet, plan.id);
  if (rowIndex === -1) return { error: "Plan no encontrado: " + plan.id };
  const row = HEADERS.map(h => (plan[h] !== undefined && plan[h] !== null) ? plan[h] : "");
  sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([row]);
  return plan;
}

function deletePlan(id) {
  const sheet = getSheet();
  const rowIndex = findRowIndexById(sheet, id);
  if (rowIndex === -1) return { error: "Plan no encontrado: " + id };
  sheet.deleteRow(rowIndex);
  return { success: true, id: id };
}

function getMarcadorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MARCADOR_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MARCADOR_SHEET_NAME);
    sheet.appendRow(MARCADOR_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllMarcador() {
  const sheet = getMarcadorSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function addMarcador(entry) {
  const sheet = getMarcadorSheet();
  entry.id = "m_" + new Date().getTime();
  entry.fechaCreacion = new Date().toISOString();
  const row = MARCADOR_HEADERS.map(h => (entry[h] !== undefined && entry[h] !== null) ? entry[h] : "");
  sheet.appendRow(row);
  return entry;
}

function deleteMarcador(id) {
  const sheet = getMarcadorSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: "Registro no encontrado: " + id };
}

/* ---------- MENSAJES (ventana de bienvenida) ----------
 * Cada fila es una foto + frase. La app muestra una distinta cada día.
 * Puedes editar la frase directamente en la hoja si quieres corregirla.
 */
function getMensajesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MENSAJES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MENSAJES_SHEET_NAME);
    sheet.appendRow(MENSAJES_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllMensajes() {
  const data = getMensajesSheet().getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0] !== "" && row[0] !== null)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
}

function addMensaje(mensaje) {
  if (!mensaje || !mensaje.fotoUrl) return { error: "Falta la foto del mensaje." };
  const sheet = getMensajesSheet();
  mensaje.id = "w_" + new Date().getTime();
  mensaje.fechaCreacion = new Date().toISOString();
  sheet.appendRow(MENSAJES_HEADERS.map(h => (mensaje[h] !== undefined && mensaje[h] !== null) ? mensaje[h] : ""));
  return mensaje;
}

function deleteMensaje(id) {
  const sheet = getMensajesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { success: true, id: id };
    }
  }
  return { error: "Mensaje no encontrado: " + id };
}

function getPhotosFolder() {
  const folders = DriveApp.getFoldersByName(PHOTOS_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTOS_FOLDER_NAME);
}

/**
 * fileData: string base64 SIN el prefijo "data:image/...;base64,"
 * fileName: nombre sugerido, ej. "plan_p_12345.jpg"
 * mimeType: ej. "image/jpeg"
 */
function uploadPhoto(fileData, fileName, mimeType) {
  if (!fileData) return { error: "No llegó ninguna imagen." };
  const folder = getPhotosFolder();
  const bytes = Utilities.base64Decode(fileData);
  const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", fileName || ("foto_" + Date.now() + ".jpg"));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // Formato directo, más confiable para <img src> que el link de "compartir" estándar.
  const directUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
  return { url: directUrl, fileId: file.getId() };
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
