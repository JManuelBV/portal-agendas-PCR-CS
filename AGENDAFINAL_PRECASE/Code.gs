// =============================================================
// CONFIGURACIÓN
// 1) Pega aquí el ID de tu Google Sheet entre comillas.
//    Ejemplo: const SPREADSHEET_ID = '1ABCDEF...';
// 2) Si este Apps Script está ligado directamente a la Sheet,
//    puedes dejarlo vacío: ''.
// =============================================================
const SPREADSHEET_ID = '';

const SHEET_NAME = 'AGENDA PCR-CS';
const EXPORT_SHEET_NAME = 'PLATAFORMA';

const HEADERS = [
  'FECHA APERTURA',
  'ID',
  'NOMBRE DEL CLIENTE',
  'TIPO DE VISA',
  'AGENTE PCR',
  'FECHA LLAMADA',
  'HORARIO DE LLAMADA',
  'LINK VIDEOLLAMADA',
  'NOTAS'
];

function doGet(e) {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Agendas PCR-CS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiGetWeb(params) {
  try {
    params = params || {};
    const action = params.action || '';

    if (action === 'getAgenda') {
      const data = getAgenda_(params.from, params.to, params.person || 'ALL');
      return { success: true, data };
    }

    return { success: true, message: 'Apps Script activo', data: [] };
  } catch (err) {
    return { success: false, message: err.message || String(err), data: [] };
  }
}

function apiPostWeb(body) {
  try {
    body = body || {};
    const action = body.action || '';

    if (action === 'saveAppointment') {
      saveAppointment_(body);
      return { success: true, message: 'Cita guardada correctamente.' };
    }

    if (action === 'deleteAppointment') {
      const deleted = deleteAppointment_(body);
      return {
        success: true,
        deleted,
        message: deleted ? 'Cita eliminada.' : 'No se encontró la cita para eliminar.'
      };
    }

    if (action === 'exportAgendaToSheet') {
      const filters = body.filters || {};
      const exported = exportAgendaToSheet_(filters.from, filters.to, filters.person || 'ALL');
      return {
        success: true,
        message: `Agenda exportada a PLATAFORMA. Registros agregados: ${exported}.`,
        exported
      };
    }

    return { success: false, message: 'Acción no reconocida.' };
  } catch (err) {
    return { success: false, message: err.message || String(err) };
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = apiPostWeb(body);
    return json_(result);
  } catch (err) {
    return json_({ success: false, message: err.message || String(err) });
  }
}

function saveAppointment_(body) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  ensureHeaders_(sheet);

  const row = [
    body.openingDate || '',
    body.id || '',
    body.clientName || '',
    body.visaType || '',
    body.personAssigned || '',
    body.callDate || '',
    body.callTimeLabel || '',
    body.meetingLink || '',
    body.notes || ''
  ];

  sheet.appendRow(row);
}

function getAgenda_(from, to, person) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  ensureHeaders_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const fromDate = normalizeYMD_(from);
  const toDate = normalizeYMD_(to);
  const personFilter = String(person || 'ALL').trim();

  return values
    .map(rowToObject_)
    .filter(item => {
      const callDate = normalizeYMD_(item.callDate);

      if (fromDate && callDate < fromDate) return false;
      if (toDate && callDate > toDate) return false;
      if (personFilter !== 'ALL' && item.personAssigned !== personFilter) return false;

      return true;
    });
}

function deleteAppointment_(body) {
  const sheet = getOrCreateSheet_(SHEET_NAME);
  ensureHeaders_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  const targetId = String(body.id || '').trim();
  const targetDate = normalizeYMD_(body.date || body.callDate || '');
  const targetPerson = String(body.personAssigned || body.person || '').trim();
  const targetTime = String(body.time || body.callTimeLabel || '').trim().toLowerCase();

  for (let i = values.length - 1; i >= 0; i--) {
    const item = rowToObject_(values[i]);

    const sameId = String(item.id || '').trim() === targetId;
    const sameDate = normalizeYMD_(item.callDate) === targetDate;
    const samePerson = String(item.personAssigned || '').trim() === targetPerson;
    const sameTime = !targetTime || String(item.callTimeLabel || '').trim().toLowerCase() === targetTime;

    if (sameId && sameDate && samePerson && sameTime) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }

  return false;
}

function exportAgendaToSheet_(from, to, person) {
  const data = getAgenda_(from, to, person || 'ALL');
  const exportSheet = getOrCreateSheet_(EXPORT_SHEET_NAME);

  // Exporta a PLATAFORMA.
  // No borra contenido previo.
  // Solo escribe columnas A:I.
  // No toca columna J en adelante.
  // Cada exportación nueva se agrega debajo.
  ensureExportHeaders_(exportSheet);

  if (!data.length) return 0;

  const rows = data.map(item => [
    item.openingDate || '',
    item.id || '',
    item.clientName || '',
    item.visaType || '',
    item.personAssigned || '',
    item.callDate || '',
    item.callTimeLabel || '',
    item.meetingLink || '',
    item.notes || ''
  ]);

  const nextRow = getNextAppendRowInAtoI_(exportSheet);
  exportSheet.getRange(nextRow, 1, rows.length, HEADERS.length).setValues(rows);

  exportSheet.setFrozenRows(1);
  exportSheet.autoResizeColumns(1, HEADERS.length);

  return rows.length;
}

function ensureExportHeaders_(sheet) {
  const firstRowValues = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const firstRowHasDataInAtoI = firstRowValues.some(value => String(value || '').trim() !== '');

  if (!firstRowHasDataInAtoI) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const currentFirstHeader = String(firstRowValues[0] || '').trim();

  if (currentFirstHeader === 'FECHA APERTURA') {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function getNextAppendRowInAtoI_(sheet) {
  const maxRows = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, maxRows, HEADERS.length).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const hasData = values[i].some(value => String(value || '').trim() !== '');

    if (hasData) {
      return i + 2;
    }
  }

  return 1;
}

function rowToObject_(row) {
  return {
    openingDate: normalizeYMD_(row[0]),
    fechaApertura: normalizeYMD_(row[0]),

    id: String(row[1] || ''),

    clientName: String(row[2] || ''),
    nombreCliente: String(row[2] || ''),

    visaType: String(row[3] || ''),
    tipoVisa: String(row[3] || ''),

    personAssigned: String(row[4] || ''),
    agentPcr: String(row[4] || ''),

    callDate: normalizeYMD_(row[5]),
    fechaLlamada: normalizeYMD_(row[5]),

    callTimeLabel: String(row[6] || ''),
    horarioLlamada: String(row[6] || ''),

    meetingLink: String(row[7] || ''),
    linkVideollamada: String(row[7] || ''),

    notes: String(row[8] || ''),
    notas: String(row[8] || '')
  };
}

function getSpreadsheet_() {
  const configuredId = String(SPREADSHEET_ID || '').trim();

  if (configuredId) {
    return SpreadsheetApp.openById(configuredId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();

  if (active) {
    return active;
  }

  throw new Error('No hay Google Sheet activa. Pega el ID de tu Sheet en la constante SPREADSHEET_ID dentro de Code.gs.');
}

function getOrCreateSheet_(name) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  return sheet;
}

function ensureHeaders_(sheet) {
  const existing = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const needsHeaders = existing.join('').trim() === '' || existing[0] !== HEADERS[0];

  if (needsHeaders) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function normalizeYMD_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const s = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  const slashDate = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);

  if (slashDate) {
    let first = Number(slashDate[1]);
    let second = Number(slashDate[2]);
    const year = slashDate[3].length === 2 ? `20${slashDate[3]}` : slashDate[3];

    // Soporta MM/DD/YYYY y también DD/MM/YYYY cuando el primer número es mayor a 12.
    let month = first;
    let day = second;
    if (first > 12 && second <= 12) {
      day = first;
      month = second;
    }

    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const parsed = new Date(s);

  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return s;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
