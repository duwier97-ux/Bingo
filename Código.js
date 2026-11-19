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
// La columna de la tabla se busca por su encabezado en la fila 1 (sin importar mayúsculas ni tildes).
const ENCABEZADO_TABLA = 'Tabla N°';

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

// Recibe texto plano con JSON: { accion, pin, codigo?, tabla? }
//   accion: 'validarPin' | 'consultar' | 'registrar' | 'resumen'. Todas exigen PIN.
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
    if (accion === 'consultar') {
      return json_(consultarCodigo_(datos.codigo));
    }
    if (accion === 'registrar') {
      return json_(registrarIngreso_(datos.codigo, datos.tabla));
    }
    if (accion === 'resumen') {
      return json_(resumen_());
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

function limpiarCodigo_(codigo) {
  return String(codigo || '').trim().toUpperCase();
}

// "Tabla N°", "TABLA Nº", "tabla  n°" -> "tabla n°"
function normalizarEncabezado_(texto) {
  return String(texto || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[º˚]/g, '°')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

// Un solo número entero positivo (acepta 7 o "7"; rechaza "", 0, "1,2", "1.5", "-3").
function limpiarTabla_(tabla) {
  const texto = String(tabla === undefined || tabla === null ? '' : tabla).trim();
  if (!/^\d+$/.test(texto)) return null;
  const n = Number(texto);
  return (Number.isSafeInteger(n) && n > 0) ? n : null;
}

function abrirHoja_() {
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(NOMBRE_HOJA);
  if (!hoja) throw new Error('No se encontró la hoja "' + NOMBRE_HOJA + '"');
  return hoja;
}

// Lee la hoja de una sola vez: { filas (desde la fila 2), colTabla (1 = A; 0 si no existe) }
function leerHoja_(hoja) {
  const ultimaFila = hoja.getLastRow();
  const ultimaCol = Math.max(hoja.getLastColumn(), COL_HORA);
  const buscado = normalizarEncabezado_(ENCABEZADO_TABLA);
  const encabezados = ultimaFila >= 1 ? hoja.getRange(1, 1, 1, ultimaCol).getValues()[0] : [];
  let colTabla = 0;
  for (let c = 0; c < encabezados.length; c++) {
    if (normalizarEncabezado_(encabezados[c]) === buscado) { colTabla = c + 1; break; }
  }
  const filas = ultimaFila >= 2 ? hoja.getRange(2, 1, ultimaFila - 1, ultimaCol).getValues() : [];
  return { filas: filas, colTabla: colTabla };
}

// Posiciones (índice dentro de filas) cuyo código coincide.
function buscarCodigo_(filas, codigoLimpio) {
  const encontradas = [];
  for (let i = 0; i < filas.length; i++) {
    if (limpiarCodigo_(filas[i][COL_CODIGO - 1]) === codigoLimpio) encontradas.push(i);
  }
  return encontradas;
}

// Busca el código y devuelve su estado, sin escribir nada.
// Estados: no_existe | duplicado | repetido (ya ingresó) | pendiente (puede ingresar)
function evaluarCodigo_(datos, codigoLimpio) {
  if (!codigoLimpio) return { estado: 'no_existe', codigo: '' };
  const pos = buscarCodigo_(datos.filas, codigoLimpio);
  if (pos.length === 0) return { estado: 'no_existe', codigo: codigoLimpio };
  if (pos.length > 1) return { estado: 'duplicado', codigo: codigoLimpio };

  const i = pos[0];
  const fila = datos.filas[i];
  const nombre = String(fila[COL_NOMBRE - 1]).trim();
  if (fila[COL_ASISTIO - 1] === true) {
    const tabla = datos.colTabla ? fila[datos.colTabla - 1] : '';
    return {
      estado: 'repetido', nombre: nombre, codigo: codigoLimpio,
      hora: formatearHora_(fila[COL_HORA - 1]),
      tabla: (tabla === '' || tabla === null || tabla === undefined) ? '' : String(tabla)
    };
  }
  return { estado: 'pendiente', nombre: nombre, codigo: codigoLimpio, indice: i };
}

function consultarCodigo_(codigo) {
  const r = evaluarCodigo_(leerHoja_(abrirHoja_()), limpiarCodigo_(codigo));
  delete r.indice;
  return r;
}

// Total = filas con código en la columna A; asistentes = de esas, las que tienen H = TRUE.
function resumen_() {
  const datos = leerHoja_(abrirHoja_());
  let total = 0;
  let asistentes = 0;
  for (let i = 0; i < datos.filas.length; i++) {
    if (String(datos.filas[i][COL_CODIGO - 1]).trim() === '') continue;
    total++;
    if (datos.filas[i][COL_ASISTIO - 1] === true) asistentes++;
  }
  return { estado: 'resumen', total: total, asistentes: asistentes, pendientes: total - asistentes };
}

function registrarIngreso_(codigo, tabla) {
  const codigoLimpio = limpiarCodigo_(codigo);
  if (!codigoLimpio) {
    return { estado: 'no_existe', codigo: '' };
  }
  const numTabla = limpiarTabla_(tabla);
  if (numTabla === null) {
    return { estado: 'tabla_invalida', codigo: codigoLimpio, mensaje: 'Escribe un solo número de tabla, entero y mayor que 0.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // evita choques si dos celulares escanean al mismo tiempo (si se agota, doPost devuelve estado 'error')

  try {
    const hoja = abrirHoja_();
    // Se lee DENTRO del lock: refleja lo que otros celulares ya escribieron.
    const datos = leerHoja_(hoja);
    if (!datos.colTabla) {
      return { estado: 'error', mensaje: 'No existe la columna "' + ENCABEZADO_TABLA + '" en la fila 1 de la hoja "' + NOMBRE_HOJA + '". Agrégala y vuelve a intentar.' };
    }

    // Se vuelve a comprobar no_existe / duplicado / repetido.
    const r = evaluarCodigo_(datos, codigoLimpio);
    if (r.estado !== 'pendiente') return r;

    // Un número de tabla no puede estar asignado a otra persona.
    for (let i = 0; i < datos.filas.length; i++) {
      if (i === r.indice) continue;
      if (limpiarTabla_(datos.filas[i][datos.colTabla - 1]) === numTabla) {
        return {
          estado: 'tabla_repetida', tabla: numTabla, codigoSolicitado: codigoLimpio,
          nombre: String(datos.filas[i][COL_NOMBRE - 1]).trim(),
          codigo: limpiarCodigo_(datos.filas[i][COL_CODIGO - 1])
        };
      }
    }

    const fila = r.indice + 2;
    const ahora = new Date();
    hoja.getRange(fila, datos.colTabla).setValue(numTabla);
    hoja.getRange(fila, COL_HORA).setValue(ahora);
    hoja.getRange(fila, COL_ASISTIO).setValue(true);
    SpreadsheetApp.flush();

    return { estado: 'ok', nombre: r.nombre, codigo: codigoLimpio, hora: formatearHora_(ahora), tabla: numTabla };
  } finally {
    lock.releaseLock();
  }
}
