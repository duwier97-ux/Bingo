// ===== CONTROL DE INGRESO CON QR (API JSON para GitHub Pages) =====
// El PIN NO va en el código: Configuración del proyecto > Propiedades del script > PIN

const ID_HOJA = '11dB-G51rdEzHs31NqfWQmUVXgw9Re8yGWwl0PzRee48';
const NOMBRE_HOJA = 'Base de datos';
const ZONA = 'America/Bogota';

// Columnas de la hoja "Base de datos"
const COL_CODIGO = 1;  // A: Código Único
const COL_NOMBRE = 2;  // B: Nombre Completo
const COL_ASISTIO = 8; // H: Asistió (casilla)
const COL_HORA = 9;    // I: Hora de ingreso

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function pinValido_(pin) {
  const correcto = String(PropertiesService.getScriptProperties().getProperty('PIN') || '');
  return correcto !== '' && String(pin || '') === correcto;
}

// Comprobación de estado: abrir la URL /exec en el navegador debe mostrar este JSON.
function doGet() {
  return json_({ estado: 'activo', hora: formatearHora_(new Date()) });
}

// Recibe texto plano con JSON: { accion: 'validarPin' | 'registrar', pin: '...', codigo: '...' }
function doPost(e) {
  try {
    const datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const accion = String(datos.accion || 'registrar');

    if (!pinValido_(datos.pin)) {
      return json_({ estado: 'pin' });
    }
    if (accion === 'validarPin') {
      return json_({ estado: 'pin_ok' });
    }
    if (accion === 'registrar') {
      return json_(registrarIngreso_(datos.codigo));
    }
    return json_({ estado: 'error', mensaje: 'Acción no válida' });
  } catch (err) {
    return json_({ estado: 'error', mensaje: String((err && err.message) || err) });
  }
}

function formatearHora_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, ZONA, 'dd/MM/yyyy HH:mm');
  }
  return '';
}

function registrarIngreso_(codigo) {
  const codigoLimpio = String(codigo || '').trim().toUpperCase();
  if (!codigoLimpio) {
    return { estado: 'no_existe', codigo: '' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // evita choques si dos celulares escanean al mismo tiempo (si se agota, doPost devuelve estado 'error')

  try {
    const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(NOMBRE_HOJA);
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) {
      return { estado: 'no_existe', codigo: codigoLimpio };
    }

    const codigos = hoja.getRange(2, COL_CODIGO, ultimaFila - 1, 1).getValues();
    const filas = [];
    for (let i = 0; i < codigos.length; i++) {
      if (String(codigos[i][0]).trim().toUpperCase() === codigoLimpio) {
        filas.push(i + 2);
      }
    }

    if (filas.length === 0) {
      return { estado: 'no_existe', codigo: codigoLimpio };
    }
    if (filas.length > 1) {
      return { estado: 'duplicado', codigo: codigoLimpio };
    }

    const fila = filas[0];
    const nombre = String(hoja.getRange(fila, COL_NOMBRE).getValue()).trim();
    const yaAsistio = hoja.getRange(fila, COL_ASISTIO).getValue() === true;

    if (yaAsistio) {
      const horaPrevia = hoja.getRange(fila, COL_HORA).getValue();
      return { estado: 'repetido', nombre: nombre, codigo: codigoLimpio, hora: formatearHora_(horaPrevia) };
    }

    const ahora = new Date();
    hoja.getRange(fila, COL_ASISTIO).setValue(true);
    hoja.getRange(fila, COL_HORA).setValue(ahora);
    SpreadsheetApp.flush();

    return { estado: 'ok', nombre: nombre, codigo: codigoLimpio, hora: formatearHora_(ahora) };
  } finally {
    lock.releaseLock();
  }
}
