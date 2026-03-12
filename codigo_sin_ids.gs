const SPREADSHEET_ID = 'PEGAR_AQUI_ID_DE_TU_SPREADSHEET';
const RESPUESTAS_SHEET_NAME = 'Respuestas';
const HEADERS = [
  'FechaHoraRegistro',
  'TipoInstitucionCodigo',
  'TipoInstitucion',
  'EntidadNombre',
  'PaisNombre',
  'RUC',
  'TipoDocumentoCodigo',
  'TipoDocumento',
  'NumeroDocumento',
  'NombreCompleto',
  'Celular',
  'CorreoElectronico',
  'ClaveInstitucion'
];

function doGet() {
  return ContentService
    .createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = String(params.action || '').trim();
  const requestId = String(params.requestId || '');

  try {
    if (!action) {
      return buildPostMessageResponse_({
        source: 'onpe-registro-webapp',
        requestId,
        status: 'error',
        message: 'No se recibió una acción válida.'
      });
    }

    if (action === 'checkInstitution') {
      return handleCheckInstitution_(params, requestId);
    }

    if (action === 'submitRegistration') {
      return handleSubmitRegistration_(params, requestId);
    }

    return buildPostMessageResponse_({
      source: 'onpe-registro-webapp',
      requestId,
      status: 'error',
      message: 'Acción no reconocida.'
    });
  } catch (error) {
    return buildPostMessageResponse_({
      source: 'onpe-registro-webapp',
      requestId,
      status: 'error',
      message: error && error.message ? error.message : 'Error inesperado.'
    });
  }
}

function handleCheckInstitution_(params, requestId) {
  const payload = sanitizePayload_(params);
  const sheet = getOrCreateSheet_(SpreadsheetApp.openById(SPREADSHEET_ID), RESPUESTAS_SHEET_NAME);
  ensureHeaders_(sheet);

  const claveInstitucion = buildInstitutionKey_(payload);
  const existe = institutionExists_(sheet, claveInstitucion);

  return buildPostMessageResponse_({
    source: 'onpe-registro-webapp',
    requestId,
    status: existe ? 'duplicate' : 'available',
    message: existe ? 'La institución o entidad ya cuenta con un registro.' : 'Disponible.'
  });
}

function handleSubmitRegistration_(params, requestId) {
  const payload = sanitizePayload_(params);
  validatePayload_(payload);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = getOrCreateSheet_(ss, RESPUESTAS_SHEET_NAME);
    ensureHeaders_(sheet);

    const claveInstitucion = buildInstitutionKey_(payload);
    if (institutionExists_(sheet, claveInstitucion)) {
      return buildPostMessageResponse_({
        source: 'onpe-registro-webapp',
        requestId,
        status: 'duplicate',
        message: 'La institución o entidad ya cuenta con un registro.'
      });
    }

    sheet.appendRow([
      new Date(),
      payload.tipoInstitucion,
      payload.tipoInstitucionLabel,
      payload.entidadNombre,
      payload.paisNombre,
      payload.ruc,
      payload.tipoDocumento,
      payload.tipoDocumentoLabel,
      payload.numeroDocumento,
      payload.nombreCompleto,
      payload.celular,
      payload.correoElectronico,
      claveInstitucion
    ]);

    return buildPostMessageResponse_({
      source: 'onpe-registro-webapp',
      requestId,
      status: 'success',
      message: 'Registro completado correctamente.'
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function sanitizePayload_(params) {
  return {
    tipoInstitucion: cleanText_(params.tipoInstitucion),
    tipoInstitucionLabel: cleanText_(params.tipoInstitucionLabel),
    entidadNombre: cleanText_(params.entidadNombre),
    paisNombre: cleanText_(params.paisNombre),
    ruc: onlyDigits_(params.ruc),
    tipoDocumento: cleanText_(params.tipoDocumento),
    tipoDocumentoLabel: cleanText_(params.tipoDocumentoLabel),
    numeroDocumento: onlyDigits_(params.numeroDocumento),
    nombreCompleto: cleanText_(params.nombreCompleto),
    celular: onlyDigits_(params.celular),
    correoElectronico: cleanText_(params.correoElectronico).toLowerCase()
  };
}

function validatePayload_(payload) {
  const tiposValidos = {
    institucion_publica: true,
    organizacion_politica: true,
    mision_observacion: true,
    encuestadora_vigente: true
  };

  if (!tiposValidos[payload.tipoInstitucion]) {
    throw new Error('Tipo de institución no válido.');
  }

  if (!payload.entidadNombre) {
    throw new Error('Debe seleccionar una institución o entidad válida.');
  }

  if ((payload.tipoInstitucion === 'institucion_publica' || payload.tipoInstitucion === 'encuestadora_vigente') && !/^\d{11}$/.test(payload.ruc)) {
    throw new Error('El RUC debe tener exactamente 11 dígitos.');
  }

  if (payload.tipoInstitucion === 'mision_observacion' && !payload.paisNombre) {
    throw new Error('Debe seleccionar el país de la misión de observación.');
  }

  if (!payload.tipoDocumento || !payload.tipoDocumentoLabel) {
    throw new Error('Debe seleccionar el tipo de documento.');
  }

  if (!/^\d{6,15}$/.test(payload.numeroDocumento)) {
    throw new Error('El número de documento debe contener entre 6 y 15 dígitos.');
  }

  if (!payload.nombreCompleto || payload.nombreCompleto.length < 5) {
    throw new Error('Debe ingresar nombres y apellidos válidos.');
  }

  if (!/^\d{9}$/.test(payload.celular)) {
    throw new Error('El número de celular debe contener exactamente 9 dígitos.');
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(payload.correoElectronico)) {
    throw new Error('Debe ingresar un correo electrónico válido.');
  }
}

function buildInstitutionKey_(payload) {
  if (payload.tipoInstitucion === 'mision_observacion') {
    return [payload.tipoInstitucion, payload.entidadNombre, payload.paisNombre]
      .map(normalizeKeyPart_)
      .join('|');
  }

  return [payload.tipoInstitucion, payload.entidadNombre]
    .map(normalizeKeyPart_)
    .join('|');
}

function institutionExists_(sheet, key) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const values = sheet.getRange(2, HEADERS.length, lastRow - 1, 1).getDisplayValues().flat();
  return values.indexOf(key) !== -1;
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getDisplayValues()[0];
  const mismatch = HEADERS.some(function(header, index) {
    return currentHeaders[index] !== header;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function buildPostMessageResponse_(payload) {
  const json = JSON.stringify(payload);
  const html = '<!DOCTYPE html><html><body><script>' +
    'window.top.postMessage(' + json + ', "*");' +
    '</script></body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cleanText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function onlyDigits_(value) {
  return String(value || '').replace(/\D+/g, '');
}

function normalizeKeyPart_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}
