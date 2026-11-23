// Pegar en Extensiones > Apps Script de una Google Sheet nueva.
// Desplegar > Nueva implementación > Aplicación web
//   Ejecutar como: Yo · Quién tiene acceso: Cualquier persona
const SHEET_NAME = 'Participantes';
const RAFFLE_KEY = 'CAMBIA-ESTA-CLAVE'; // clave para abrir la tómbola; no la compartas

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const data = JSON.parse(e.postData.contents);
    const name = clean_(data.name, 100);
    const email = clean_(data.email, 200).toLowerCase();
    const antenna = clean_(data.antenna, 50);
    if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json_({ ok: false });

    const sheet = sheet_();
    const lastRow = sheet.getLastRow();
    const emails = lastRow > 1 ? sheet.getRange(2, 3, lastRow - 1, 1).getValues().flat() : [];
    if (emails.includes(email)) return json_({ ok: true, duplicate: true });

    sheet.appendRow([new Date(), name, email, antenna]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if ((e.parameter.key || '') !== RAFFLE_KEY) return json_({ ok: false, error: 'Clave incorrecta' });
  const sheet = sheet_();
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 2, lastRow - 1, 3).getValues() : [];
  const participants = rows.map(([name, email, antenna]) => ({ name, email, antenna }));
  return json_({ ok: true, participants });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) sheet.appendRow(['Fecha', 'Nombre', 'Correo', 'Antena']);
  return sheet;
}

// Evita que un texto como "=FORMULA()" se ejecute dentro de la hoja.
function clean_(value, maxLength) {
  const text = String(value || '').trim().slice(0, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
