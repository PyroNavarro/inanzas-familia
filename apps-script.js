// ============================================================
//  FINANZAS FAMILIA — Google Apps Script
//  Pega este código en: Google Sheets → Extensiones → Apps Script
//  Luego: Implementar → Nueva implementación → Web App
//  Acceso: Cualquiera (para que Make y el dashboard puedan llamarlo)
// ============================================================

const SHEET_GASTOS   = 'Gastos';
const SHEET_INGRESOS = 'Ingresos';
const SHEET_LOG      = 'WhatsApp Log';

// ── GET: el dashboard llama esto para obtener todos los datos ──
function doGet(e) {
  const action = e.parameter.action || 'getAll';

  if (action === 'getAll') {
    const gastos   = getHoja(SHEET_GASTOS);
    const ingresos = getHoja(SHEET_INGRESOS);
    return jsonResponse({ gastos, ingresos });
  }

  return jsonResponse({ error: 'Acción no reconocida' });
}

// ── POST: Make llama esto cuando llega un mensaje de WhatsApp ──
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action || 'add';

    if (action === 'add') {
      if (body.tipo === 'gasto') {
        agregarGasto(body);
      } else if (body.tipo === 'ingreso') {
        agregarIngreso(body);
      }
      // Guardar log del mensaje original de WhatsApp
      if (body.mensajeOriginal) {
        guardarLog(body);
      }
      return jsonResponse({ ok: true });
    }

  } catch(err) {
    return jsonResponse({ error: err.message });
  }
}

// ── AGREGAR GASTO ──────────────────────────────────────────────
function agregarGasto(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_GASTOS);

  // Crear hoja si no existe, con encabezados
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_GASTOS);
    sheet.appendRow(['ID','Fecha','Día','Mes','Año','Descripción','Monto','Tipo','Categoría','Origen']);
    sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#f0f0f0');
  }

  const ahora = new Date();
  const mes   = data.mes !== undefined ? parseInt(data.mes) : ahora.getMonth();
  const dia   = data.dia || ahora.getDate();

  sheet.appendRow([
    Date.now(),                      // ID único
    ahora.toISOString(),             // Timestamp completo
    dia,                             // Día del mes
    mes,                             // Mes (0-11)
    ahora.getFullYear(),             // Año
    data.desc   || '',               // Descripción
    parseFloat(data.monto) || 0,     // Monto
    data.tipogasto || 'indispensable',// Tipo
    data.cat    || 'General',        // Categoría
    data.origen || 'whatsapp'        // Origen (manual / whatsapp)
  ]);
}

// ── AGREGAR INGRESO ────────────────────────────────────────────
function agregarIngreso(data) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let sheet  = ss.getSheetByName(SHEET_INGRESOS);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_INGRESOS);
    sheet.appendRow(['ID','Fecha','Día','Mes','Año','Descripción','Quién','Monto','Origen']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#f0f0f0');
  }

  const ahora = new Date();
  const mes   = data.mes !== undefined ? parseInt(data.mes) : ahora.getMonth();

  sheet.appendRow([
    Date.now(),
    ahora.toISOString(),
    data.dia  || ahora.getDate(),
    mes,
    ahora.getFullYear(),
    data.desc  || '',
    data.quien || '',
    parseFloat(data.monto) || 0,
    data.origen || 'manual'
  ]);
}

// ── LOG WhatsApp ───────────────────────────────────────────────
function guardarLog(data) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  let sheet  = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(['Timestamp','Número','Mensaje Original','Parseado','Estado']);
    sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f0f0f0');
  }
  sheet.appendRow([
    new Date().toISOString(),
    data.numero || '',
    data.mensajeOriginal || '',
    JSON.stringify(data),
    'OK'
  ]);
}

// ── LEER HOJA y convertir a JSON ───────────────────────────────
function getHoja(nombreHoja) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) return [];

  const datos = sheet.getDataRange().getValues();
  if (datos.length < 2) return [];

  const headers = datos[0].map(h => String(h).toLowerCase().replace(/ /g,'_'));
  return datos.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    // Asegurar que ID sea numérico para compatibilidad con el frontend
    obj.id = parseInt(obj.id) || Date.now();
    return obj;
  });
}

// ── HELPER: respuesta JSON con CORS ───────────────────────────
function jsonResponse(data) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
