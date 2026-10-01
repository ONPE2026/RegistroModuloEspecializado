/**
 * ERM2026 - Prerregistro Módulo Especializado
 * Backend en Google Apps Script para el formulario index.html.
 *
 * CÓMO USARLO:
 * 1. Crea una Google Sheet nueva (o usa una que ya tengas para ERM2026).
 * 2. En esa Sheet: Extensiones > Apps Script.
 * 3. Borra el contenido de Code.gs que aparece por defecto y pega este archivo completo.
 * 4. Ajusta, si quieres, el nombre de la hoja en SHEET_NAME y el ID de la carpeta
 *    de Drive en DRIVE_FOLDER_ID (abajo).
 * 5. Implementar > Nueva implementación > tipo "Aplicación web":
 *      - Ejecutar como: Yo (tu cuenta)
 *      - Quién tiene acceso: Cualquier usuario
 *    Copia la URL que termina en /exec.
 * 6. Pega esa URL en la constante API_URL de index.html.
 * 7. Prueba: llena el formulario completo (incluyendo adjuntar un PDF) y revisa
 *    que aparezca una fila nueva en la hoja "Registros" y un archivo nuevo en
 *    la carpeta de Drive indicada.
 *
 * IMPORTANTE sobre la carpeta de Drive: NO hace falta que la carpeta esté
 * compartida públicamente ni "con cualquiera que tenga el enlace". El script
 * se ejecuta con tu propia cuenta de Google (por eso "Ejecutar como: Yo"),
 * así que basta con que la carpeta sea tuya (o esté compartida contigo con
 * permiso de edición). Puedes dejarla completamente privada/restringida.
 */

const SHEET_NAME = 'Registros';

// ID de la carpeta de Drive donde se guardan los PDF firmados.
// Es la parte final de la URL de la carpeta:
// https://drive.google.com/drive/folders/ESTE_ES_EL_ID
const DRIVE_FOLDER_ID = '1JmDr6zYHMZYyF8sF5kyCbgSxDCrZ2lo2';

// Columnas cuyo valor puede empezar con '+' (como el celular con código de país)
// y por eso deben forzarse a texto, o Sheets las toma como el inicio de una fórmula.
const COLUMNAS_TEXTO_FORZADO = ['celular'];

const COLUMNAS = [
  'marcaTemporal',
  'submissionId',
  'sentAt',
  'tipoInstitucion',
  'tipoInstitucionLabel',
  'entidadNombre',
  'paisNombre',
  'ruc',
  'tipoDocumento',
  'tipoDocumentoLabel',
  'numeroDocumento',
  'nombres',
  'apellidos',
  'celular',
  'correoElectronico',
  'documentoUrl'
];

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'checkInstitution') {
      return responderJson(manejarCheckInstitution(body));
    }

    if (action === 'submitRegistration') {
      return responderJson(manejarSubmitRegistration(body));
    }

    return responderJson({ status: 'error', message: 'Acción no reconocida.' });
  } catch (error) {
    return responderJson({ status: 'error', message: String(error) });
  }
}

// Algunos navegadores (sendBeacon con no-cors, o llamadas de prueba) hacen GET;
// respondemos algo simple para que no truene si alguien abre la URL directo.
function doGet(e) {
  return ContentService.createTextOutput('ERM2026 - Prerregistro: endpoint activo.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function manejarCheckInstitution(body) {
  // Solo lectura: no se usa LockService (si otro envío con PDF tenía el candado,
  // la validación se quedaba esperando). Además se leen únicamente las columnas
  // necesarias (tipoInstitucion .. paisNombre) en vez de toda la hoja.
  const hoja = obtenerHoja();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return { status: 'available' };

  const idxTipo = COLUMNAS.indexOf('tipoInstitucion');
  const idxEntidad = COLUMNAS.indexOf('entidadNombre');
  const idxPais = COLUMNAS.indexOf('paisNombre');
  const primeraCol = idxTipo + 1;
  const numCols = idxPais - idxTipo + 1;

  const filas = hoja.getRange(2, primeraCol, ultimaFila - 1, numCols).getValues();

  const tipo = normalizar(body.tipoInstitucion);
  const entidadNombre = normalizar(body.entidadNombre);
  const paisNombre = normalizar(body.paisNombre);

  for (let i = 0; i < filas.length; i++) {
    const fila = filas[i];
    if (normalizar(fila[0]) !== tipo) continue;
    if (normalizar(fila[idxEntidad - idxTipo]) !== entidadNombre) continue;

    if (tipo === 'mision_observacion') {
      if (normalizar(fila[idxPais - idxTipo]) === paisNombre) return { status: 'duplicate' };
    } else {
      return { status: 'duplicate' };
    }
  }

  return { status: 'available' };
}

function manejarSubmitRegistration(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const hoja = obtenerHoja();

    // Deduplicar reintentos: si el navegador reenvía el mismo submissionId
    // (por ejemplo, si el usuario hace doble clic), no lo volvemos a insertar
    // ni volvemos a subir el archivo.
    if (body.submissionId && yaExisteSubmission(hoja, body.submissionId)) {
      return { status: 'ok', deduped: true };
    }

    // Número de respuesta = cuántas filas de datos hay antes de esta
    // (la fila 1 es el encabezado, así que getLastRow() antes de insertar
    // ya coincide con el número de esta nueva respuesta).
    const numeroRespuesta = hoja.getLastRow();

    const nombreArchivoBase = construirNombreArchivo(body, numeroRespuesta);
    const documentoUrl = guardarArchivoEnDrive(body, nombreArchivoBase);

    const fila = COLUMNAS.map((col) => {
      if (col === 'marcaTemporal') return new Date();
      if (col === 'documentoUrl') return documentoUrl;

      const valor = body[col] !== undefined ? body[col] : '';
      if (COLUMNAS_TEXTO_FORZADO.indexOf(col) !== -1 && valor !== '') {
        return "'" + valor; // fuerza texto: evita que "+51 9..." se lea como fórmula
      }
      return valor;
    });

    hoja.appendRow(fila);
    return { status: 'ok' };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Decide el nombre de archivo según el tipo de institución:
 * - Organización política: el nombre de la organización.
 * - Encuestadora vigente: el nombre (razón social) de la encuestadora.
 * - Institución pública / misión de observación: un ID según el número de
 *   respuesta (ID_1, ID_2, ...), porque no hay un nombre corto disponible.
 */
function construirNombreArchivo(body, numeroRespuesta) {
  const tipo = body.tipoInstitucion;
  let base;

  if (tipo === 'partido_nacional' || tipo === 'movimiento_regional' || tipo === 'encuestadora_vigente') {
    base = body.entidadNombre;
  }

  if (!base) {
    base = 'ID_' + numeroRespuesta;
  }

  return sanitizarNombreArchivo(base);
}

function sanitizarNombreArchivo(nombre) {
  const limpio = String(nombre)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // quitar tildes
    .replace(/[^A-Za-z0-9 _-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .substring(0, 100);

  return limpio || 'REGISTRO';
}

/**
 * Guarda el PDF adjunto (viene en base64 desde el formulario) en la carpeta
 * de Drive configurada arriba. Devuelve la URL del archivo, o '' si no venía
 * archivo, o 'ERROR: ...' si algo falló (para poder revisarlo en la hoja).
 */
function guardarArchivoEnDrive(body, nombreBase) {
  if (!body.archivoBase64) return '';

  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const contentType = body.archivoTipo || 'application/pdf';
    const bytes = Utilities.base64Decode(body.archivoBase64);
    const blob = Utilities.newBlob(bytes, contentType, nombreBase + '.pdf');
    const archivo = folder.createFile(blob);
    return archivo.getUrl();
  } catch (error) {
    return 'ERROR: ' + error;
  }
}

function yaExisteSubmission(hoja, submissionId) {
  const filas = hoja.getDataRange().getValues();
  const encabezados = filas[0];
  const idx = encabezados.indexOf('submissionId');
  if (idx === -1) return false;

  for (let i = 1; i < filas.length; i++) {
    if (filas[i][idx] === submissionId) return true;
  }
  return false;
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(SHEET_NAME);

  if (!hoja) {
    hoja = libro.insertSheet(SHEET_NAME);
    hoja.appendRow(COLUMNAS);
    hoja.setFrozenRows(1);

    COLUMNAS_TEXTO_FORZADO.forEach((col) => {
      const idx = COLUMNAS.indexOf(col);
      if (idx === -1) return;
      // Columna completa (menos el encabezado) como texto plano ('@').
      hoja.getRange(2, idx + 1, hoja.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  }

  return hoja;
}

function normalizar(valor) {
  return String(valor || '').trim().toUpperCase();
}

function responderJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
