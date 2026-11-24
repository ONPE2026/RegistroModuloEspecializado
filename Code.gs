/**
 * Segunda Vuelta Regional - Prerregistro Módulo Especializado
 * Backend en Google Apps Script para el formulario index.html.
 *
 * IMPORTANTE: este es un proceso de registro NUEVO e independiente del de
 * ERM2026 (primera vuelta). Usa una Google Sheet nueva y una carpeta de
 * Drive nueva (no reutilices las de la primera vuelta), para no mezclar los
 * datos de ambos procesos.
 *
 * CÓMO USARLO:
 * 1. Crea una Google Sheet nueva para la Segunda Vuelta Regional.
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
const DRIVE_FOLDER_ID = 'PEGAR_AQUI_ID_CARPETA_DRIVE';

// Token compartido: debe coincidir exactamente con APP_TOKEN en index.html.
// Es una barrera simple contra envios automatizados directos a esta URL
// (bots, scripts) que no pasan por el formulario real. No es una autenticacion
// fuerte (cualquiera que vea el codigo fuente de la pagina lo puede ver), pero
// filtra el abuso casual de quien solo encuentra la URL.
const APP_TOKEN = 'SVR-2026-prereg';

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

    if (body.appToken !== APP_TOKEN) {
      return responderJson({ status: 'error', message: 'Solicitud no autorizada.' });
    }

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
  return ContentService.createTextOutput('Segunda Vuelta Regional - Prerregistro: endpoint activo.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// A partir de la Segunda Vuelta Regional se permite más de una persona
// registrada por la misma organización/institución, así que esta función
// ya no bloquea por duplicado: siempre deja continuar al paso 2. Se deja la
// accion 'checkInstitution' (y esta funcion) por compatibilidad con el
// frontend, que todavia la llama antes de mostrar el resto del formulario.
function manejarCheckInstitution(body) {
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
 * Decide el nombre de archivo: usa el nombre de la institucion/entidad
 * (de la lista o escrito a mano en "Otro"), y si por algun motivo viene
 * vacio, cae a un ID segun el numero de respuesta.
 */
function construirNombreArchivo(body, numeroRespuesta) {
  const base = body.entidadNombre || ('ID_' + numeroRespuesta);
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
