// ===== CONTROL DE INGRESO CON QR + JUEGO DE BINGO (API JSON para GitHub Pages) =====
// Las claves NO van en el código: Configuración del proyecto > Propiedades del script
//   CLAVE_INGRESO  página de ingreso (index.html). Si no existe, se usa la propiedad antigua PIN.
//   CLAVE_JUEGO    control del juego (bingo.html) y carga de tablas.
//   CLAVE_TV       pantalla (pantalla.html): solo lectura del estado.

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

function propiedad_(nombre) {
  return String(PropertiesService.getScriptProperties().getProperty(nombre) || '');
}

function claveCoincide_(dada, correcta) {
  return correcta !== '' && String(dada === undefined || dada === null ? '' : dada) === correcta;
}

function pinValido_(pin) {
  return claveCoincide_(pin, propiedad_('CLAVE_INGRESO') || propiedad_('PIN'));
}

function esClaveJuego_(clave) { return claveCoincide_(clave, propiedad_('CLAVE_JUEGO')); }
function esClaveTv_(clave) { return claveCoincide_(clave, propiedad_('CLAVE_TV')); }

// Comprobación de estado: abrir la URL /exec en el navegador debe mostrar este JSON.
function doGet() {
  return json_({ estado: 'activo', hora: formatearHora_(new Date()) });
}

// Recibe texto plano con JSON.
//   Ingreso: { accion: 'validarPin' | 'consultar' | 'registrar' | 'resumen', pin, codigo?, tabla? } (CLAVE_INGRESO)
//   Juego:   { accion: <ver ACCIONES_JUEGO>, clave, ... } (CLAVE_JUEGO; 'estado' y 'validarClave' también con CLAVE_TV)
function doPost(e) {
  try {
    const datos = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const accion = String(datos.accion || 'registrar');

    if (Object.prototype.hasOwnProperty.call(ACCIONES_JUEGO, accion)) {
      return json_(atenderJuego_(accion, datos));
    }

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

// Página de resumen (resumen.html), solo lectura con CLAVE_TV o CLAVE_JUEGO. Mismos conteos que resumen_():
// tablasVendidas = filas con código en la columna A (vacías ignoradas); asistentes = de esas, H = TRUE.
function resumenVentas_() {
  const r = resumen_();
  return { estado: 'ventas', tablasVendidas: r.total, asistentes: r.asistentes, hora: formatearHora_(new Date()) };
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

  // Tablas existentes (hoja "Tablas"). Se lee antes del lock: el ingreso nunca cambia esa hoja.
  const existentes = numerosDeTablas_();

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

    // La tabla debe existir en la hoja "Tablas" (activa o no). Sin hoja o vacía, no se bloquea.
    if (existentes && !existentes[numTabla]) {
      return { estado: 'tabla_inexistente', tabla: numTabla, codigo: codigoLimpio, mensaje: 'Esa tabla no existe. Revisa el número' };
    }

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

// Números de tabla de la hoja "Tablas" ({ 762: true, ... }), activas o no.
// null si la hoja no existe, está vacía o no tiene la columna "Tabla": en ese caso el ingreso no se bloquea.
function numerosDeTablas_() {
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_TABLAS);
  if (!hoja || hoja.getLastRow() < 2 || hoja.getLastColumn() < 1) return null;
  const enc = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(normalizarTexto_);
  const c = enc.indexOf('tabla');
  if (c < 0) return null;
  const mapa = {};
  let hay = false;
  hoja.getRange(2, c + 1, hoja.getLastRow() - 1, 1).getValues().forEach(function (f) {
    const t = limpiarTabla_(f[0]);
    if (t !== null) { mapa[t] = true; hay = true; }
  });
  return hay ? mapa : null;
}

// =====================================================================
// ===== JUEGO DE BINGO =====
// Estado en Propiedades del script (ESTADO_JUEGO) + copia en caché para el polling de la pantalla.
// Tablas en la hoja privada "Tablas" del mismo Sheet. Ganadores confirmados en la hoja "Ganadores".
// =====================================================================

const HOJA_TABLAS = 'Tablas';
const HOJA_GANADORES = 'Ganadores';
const PROP_ESTADO = 'ESTADO_JUEGO';
const CACHE_ESTADO = 'estado_juego';
const CACHE_TABLAS = 'tablas_juego';
const CENTRO = 12; // celda N3: número de tabla, cuenta siempre como marcada
const CELDAS = ['B1', 'I1', 'N1', 'G1', 'O1', 'B2', 'I2', 'N2', 'G2', 'O2', 'B3', 'I3', 'N3', 'G3', 'O3',
  'B4', 'I4', 'N4', 'G4', 'O4', 'B5', 'I5', 'N5', 'G5', 'O5']; // fila por fila (posición = fila*5 + columna)
const RANGOS = { B: [1, 15], I: [16, 30], N: [31, 45], G: [46, 60], O: [61, 75] };
const MAX_GANADORES_EN_ESTADO = 20;

// 'lectura' = también con CLAVE_TV. 'juego' = solo con CLAVE_JUEGO.
const ACCIONES_JUEGO = {
  validarClave: 'lectura', estado: 'lectura', resumenVentas: 'lectura',
  marcarBalota: 'juego', quitarBalota: 'juego', deshacerBalota: 'juego',
  fijarFigura: 'juego', fijarPremio: 'juego',
  ganadores: 'juego', confirmarGanador: 'juego', ocultarAnuncio: 'juego', nuevaRonda: 'juego', reiniciarJuego: 'juego',
  listarTablas: 'juego', activarTabla: 'juego', cargarTablas: 'juego',
  infoSorteo: 'juego', sortear: 'juego', descartarSorteo: 'juego', ocultarSorteo: 'juego', corregirGenero: 'juego'
};

function atenderJuego_(accion, datos) {
  const juego = esClaveJuego_(datos.clave);
  const tv = !juego && esClaveTv_(datos.clave);
  if (!juego && !(tv && ACCIONES_JUEGO[accion] === 'lectura')) return { estado: 'clave' };
  const rol = juego ? 'juego' : 'tv';

  switch (accion) {
    case 'validarClave': return { estado: 'clave_ok', rol: rol };
    case 'estado': return estadoPublico_(leerEstado_(), rol);
    case 'resumenVentas': return resumenVentas_();
    case 'marcarBalota': return marcarBalota_(datos.balota);
    case 'quitarBalota': return quitarBalota_(datos.balota);
    case 'deshacerBalota': return deshacerBalota_();
    case 'fijarFigura': return fijarFigura_(datos);
    case 'fijarPremio': return fijarPremio_(datos.premio);
    case 'ganadores': return calcularGanadores_(leerEstado_());
    case 'confirmarGanador': return confirmarGanador_(datos.tabla);
    case 'ocultarAnuncio': return mutarEstado_(function (e) { e.anuncio = null; });
    case 'nuevaRonda': return mutarEstado_(function (e) { e.ronda++; e.balotas = []; e.anuncio = null; });
    case 'reiniciarJuego': return reiniciarJuego_(datos.confirmacion);
    case 'listarTablas': return listarTablas_();
    case 'activarTabla': return activarTabla_(datos.tabla, datos.activa);
    case 'cargarTablas': return cargarTablas_(datos.tablas);
    case 'infoSorteo': return infoSorteo_(datos.filtro);
    case 'sortear': return sortear_(datos.premio, datos.filtro);
    case 'corregirGenero': return corregirGenero_(datos.codigo, datos.genero);
    case 'descartarSorteo': return descartarSorteo_(datos.volverASortear !== false);
    case 'ocultarSorteo': return ocultarSorteo_();
  }
  return { estado: 'error', mensaje: 'Acción no válida' };
}

// ---------- Figuras ----------
function figuras_() {
  const filas = [], cols = [];
  for (let i = 0; i < 5; i++) {
    const f = [], c = [];
    for (let j = 0; j < 5; j++) { f.push(i * 5 + j); c.push(j * 5 + i); }
    filas.push(f); cols.push(c);
  }
  const d1 = [0, 6, 12, 18, 24], d2 = [4, 8, 12, 16, 20];
  const todas = []; for (let p = 0; p < 25; p++) todas.push(p);
  const une = function () { return [].concat.apply([], arguments); };
  return {
    lleno: { nombre: 'Cartón lleno', celdas: [todas] },
    linea: { nombre: 'Línea (horizontal, vertical o diagonal)', celdas: filas.concat(cols, [d1, d2]) },
    horizontal: { nombre: 'Línea horizontal', celdas: filas },
    vertical: { nombre: 'Línea vertical', celdas: cols },
    diagonal: { nombre: 'Diagonal', celdas: [d1, d2] },
    esquinas: { nombre: 'Cuatro esquinas', celdas: [[0, 4, 20, 24]] },
    x: { nombre: 'Letra X', celdas: [une(d1, d2)] },
    marco: { nombre: 'Marco', celdas: [une(filas[0], filas[4], cols[0], cols[4])] },
    l: { nombre: 'Letra L', celdas: [une(cols[0], filas[4])] },
    t: { nombre: 'Letra T', celdas: [une(filas[0], cols[2])] },
    u: { nombre: 'Letra U', celdas: [une(cols[0], cols[4], filas[4])] },
    cruz: { nombre: 'Cruz', celdas: [une(filas[2], cols[2])] }
  };
}

function aMascara_(posiciones) {
  const m = []; for (let p = 0; p < 25; p++) m.push('0');
  posiciones.forEach(function (p) { m[p] = '1'; });
  m[CENTRO] = '1';
  return m.join('');
}

function figuraPorId_(id) {
  const lista = figuras_();
  if (!Object.prototype.hasOwnProperty.call(lista, id)) return null;
  return { id: id, nombre: lista[id].nombre, mascaras: lista[id].celdas.map(aMascara_) };
}

// ---------- Estado ----------
function estadoInicial_() {
  return { v: 0, ronda: 1, balotas: [], figura: figuraPorId_('lleno'), premio: 0, ganadores: [], anuncio: null, actualizado: '' };
}

function normalizarEstado_(e) {
  const base = estadoInicial_();
  e = e || {};
  Object.keys(base).forEach(function (k) { if (e[k] === undefined || (e[k] === null && k !== 'anuncio')) e[k] = base[k]; });
  return e;
}

// Lectura rápida (polling): caché y, si no está, propiedades.
function leerEstado_() {
  const c = CacheService.getScriptCache().get(CACHE_ESTADO);
  if (c) return normalizarEstado_(JSON.parse(c));
  const e = leerEstadoDurable_();
  CacheService.getScriptCache().put(CACHE_ESTADO, JSON.stringify(e), 21600);
  return e;
}

function leerEstadoDurable_() {
  const p = PropertiesService.getScriptProperties().getProperty(PROP_ESTADO);
  return normalizarEstado_(p ? JSON.parse(p) : null);
}

function guardarEstado_(e) {
  e.v = (Number(e.v) || 0) + 1;
  e.actualizado = formatearHora_(new Date());
  if (e.ganadores.length > MAX_GANADORES_EN_ESTADO) e.ganadores = e.ganadores.slice(-MAX_GANADORES_EN_ESTADO);
  const texto = JSON.stringify(e);
  PropertiesService.getScriptProperties().setProperty(PROP_ESTADO, texto);
  CacheService.getScriptCache().put(CACHE_ESTADO, texto, 21600);
}

function estadoPublico_(e, rol) {
  const r = {
    estado: 'juego', v: e.v, ronda: e.ronda, balotas: e.balotas,
    actual: e.balotas.length ? e.balotas[e.balotas.length - 1] : null,
    figura: e.figura, premio: e.premio, anuncio: e.anuncio, actualizado: e.actualizado
  };
  if (rol === 'juego') r.ganadores = e.ganadores;
  // Sorteo adicional: estado aparte (no cambia la versión v del bingo). La TV lo pinta en su propia capa.
  const so = leerSorteo_();
  r.sorteo = { vs: so.vs, anuncio: so.anuncio };
  return r;
}

// Cambia el estado bajo lock. fn puede devolver { error: {...} } para no guardar, o { extra: {...} }.
function mutarEstado_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const e = leerEstadoDurable_();
    const r = fn(e) || {};
    if (r.error) return Object.assign(estadoPublico_(e, 'juego'), r.error);
    guardarEstado_(e);
    return Object.assign(estadoPublico_(e, 'juego'), r.extra || {});
  } finally {
    lock.releaseLock();
  }
}

function limpiarBalota_(valor) {
  const t = String(valor === undefined || valor === null ? '' : valor).trim();
  if (!/^\d{1,2}$/.test(t)) return null;
  const n = Number(t);
  return (n >= 1 && n <= 75) ? n : null;
}

function marcarBalota_(valor) {
  const n = limpiarBalota_(valor);
  if (n === null) return { estado: 'error', mensaje: 'La balota debe ser un número de 1 a 75.' };
  return mutarEstado_(function (e) {
    if (e.balotas.indexOf(n) >= 0) return { error: { estado: 'balota_repetida', balota: n } };
    e.balotas.push(n);
    return { extra: { marcada: n } };
  });
}

function quitarBalota_(valor) {
  const n = limpiarBalota_(valor);
  if (n === null) return { estado: 'error', mensaje: 'La balota debe ser un número de 1 a 75.' };
  return mutarEstado_(function (e) {
    const i = e.balotas.indexOf(n);
    if (i < 0) return { error: { estado: 'balota_no_marcada', balota: n } };
    e.balotas.splice(i, 1);
    return { extra: { quitada: n } };
  });
}

function deshacerBalota_() {
  return mutarEstado_(function (e) {
    if (!e.balotas.length) return { error: { estado: 'sin_balotas' } };
    return { extra: { quitada: e.balotas.pop() } };
  });
}

function fijarFigura_(datos) {
  const id = String(datos.id || '');
  let figura;
  if (id === 'personalizada') {
    const m = String(datos.mascara || '');
    if (!/^[01]{25}$/.test(m)) return { estado: 'error', mensaje: 'La figura personalizada debe tener 25 casillas.' };
    const conCelda = m.split('').some(function (c, p) { return c === '1' && p !== CENTRO; });
    if (!conCelda) return { estado: 'error', mensaje: 'Marca al menos una casilla además del centro.' };
    const nombre = String(datos.nombre || '').replace(/\s+/g, ' ').trim().slice(0, 40) || 'Figura personalizada';
    figura = { id: id, nombre: nombre, mascaras: [m.slice(0, CENTRO) + '1' + m.slice(CENTRO + 1)] };
  } else {
    figura = figuraPorId_(id);
    if (!figura) return { estado: 'error', mensaje: 'Figura no válida' };
  }
  return mutarEstado_(function (e) { e.figura = figura; });
}

// Acepta 1500000, "1500000", "$ 1.500.000". Pesos colombianos, enteros.
function fijarPremio_(valor) {
  const t = String(valor === undefined || valor === null ? '' : valor).replace(/[\s$.]/g, '').replace(/,00$/, '');
  if (!/^\d{1,12}$/.test(t)) return { estado: 'error', mensaje: 'Escribe el premio en pesos, solo números.' };
  const premio = Number(t);
  return mutarEstado_(function (e) { e.premio = premio; });
}

// ---------- Tablas ----------
function normalizarTexto_(t) {
  return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').trim().toLowerCase();
}

function esVerdadero_(v) {
  return v === true || /^(true|verdadero|si|sí|1|x)$/i.test(String(v === undefined || v === null ? '' : v).trim());
}

// Devuelve [{ t: tabla, s: serie, a: activa, o: nota, n: [25 números por posición; centro = 0] }]
function tablasJuego_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get(CACHE_TABLAS);
  if (c) return JSON.parse(c);
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_TABLAS);
  if (!hoja || hoja.getLastRow() < 2) throw new Error('Aún no se han cargado las tablas (hoja "' + HOJA_TABLAS + '").');
  const valores = hoja.getRange(1, 1, hoja.getLastRow(), hoja.getLastColumn()).getValues();
  const col = {};
  valores[0].forEach(function (h, i) { col[normalizarTexto_(h)] = i; });
  const falta = ['tabla', 'activa'].concat(CELDAS.filter(function (k) { return k !== 'N3'; }).map(normalizarTexto_))
    .filter(function (k) { return col[k] === undefined; });
  if (falta.length) throw new Error('A la hoja "' + HOJA_TABLAS + '" le faltan columnas: ' + falta.join(', '));
  const lista = [];
  for (let i = 1; i < valores.length; i++) {
    const f = valores[i];
    const t = limpiarTabla_(f[col.tabla]);
    if (t === null) continue;
    lista.push({
      t: t,
      s: col.serie === undefined ? '' : String(f[col.serie]),
      a: esVerdadero_(f[col.activa]),
      o: col.nota === undefined ? '' : String(f[col.nota]),
      n: CELDAS.map(function (k, p) { return p === CENTRO ? 0 : Number(f[col[normalizarTexto_(k)]]) || 0; })
    });
  }
  const texto = JSON.stringify(lista);
  if (texto.length < 95000) cache.put(CACHE_TABLAS, texto, 21600);
  return lista;
}

// Tablas registradas en la página de ingreso: { numero: { nombre, codigo } }
function tablasConIngreso_() {
  const datos = leerHoja_(abrirHoja_());
  const mapa = {};
  if (!datos.colTabla) return mapa;
  datos.filas.forEach(function (fila) {
    if (fila[COL_ASISTIO - 1] !== true) return;
    const t = limpiarTabla_(fila[datos.colTabla - 1]);
    if (t !== null && !mapa[t]) mapa[t] = { nombre: String(fila[COL_NOMBRE - 1]).trim(), codigo: limpiarCodigo_(fila[COL_CODIGO - 1]) };
  });
  return mapa;
}

// Ganadores = tablas activas, con ingreso, que completan alguna variante de la figura con las balotas cantadas.
function calcularGanadores_(e) {
  const orden = {};
  e.balotas.forEach(function (b, i) { orden[b] = i; });
  const tablas = tablasJuego_();
  const ingreso = tablasConIngreso_();
  const conDatos = {};
  const ganadores = [], inactivasConIngreso = [];
  let enJuego = 0, casi = 0;

  tablas.forEach(function (tb) {
    conDatos[tb.t] = true;
    const persona = ingreso[tb.t];
    if (!persona) return;
    if (!tb.a) { inactivasConIngreso.push(tb.t); return; }
    enJuego++;
    const marcadas = tb.n.map(function (num, p) { return p === CENTRO || orden[num] !== undefined; });
    let mejor = null, minFaltan = 99;
    e.figura.mascaras.forEach(function (m) {
      let faltan = 0, ultima = -1;
      for (let p = 0; p < 25; p++) {
        if (m.charAt(p) !== '1' || p === CENTRO) continue;
        if (!marcadas[p]) faltan++; else if (orden[tb.n[p]] > ultima) ultima = orden[tb.n[p]];
      }
      if (faltan < minFaltan) minFaltan = faltan;
      if (faltan === 0 && (mejor === null || ultima < mejor.ultima)) mejor = { ultima: ultima, mascara: m };
    });
    if (mejor) {
      ganadores.push({
        tabla: tb.t, nombre: persona.nombre, codigo: persona.codigo, serie: tb.s,
        orden: mejor.ultima + 1, conBalota: mejor.ultima >= 0 ? e.balotas[mejor.ultima] : null,
        grilla: tb.n, marcadas: marcadas, mascara: mejor.mascara,
        confirmado: e.ganadores.some(function (g) { return g.ronda === e.ronda && g.tabla === tb.t; })
      });
    } else if (minFaltan === 1) {
      casi++;
    }
  });
  ganadores.sort(function (a, b) { return a.orden - b.orden || a.tabla - b.tabla; });
  const sinDatos = Object.keys(ingreso).map(Number).filter(function (t) { return !conDatos[t]; }).sort(function (a, b) { return a - b; });
  return {
    estado: 'ganadores', ronda: e.ronda, balotas: e.balotas.length, figura: e.figura.nombre,
    ganadores: ganadores, enJuego: enJuego, casi: casi, sinDatos: sinDatos, inactivasConIngreso: inactivasConIngreso
  };
}

function confirmarGanador_(tabla) {
  const t = limpiarTabla_(tabla);
  if (t === null) return { estado: 'error', mensaje: 'Número de tabla no válido' };
  let registro = null;
  const r = mutarEstado_(function (e) {
    const calc = calcularGanadores_(e);
    const g = calc.ganadores.filter(function (x) { return x.tabla === t; })[0];
    if (!g) return { error: { estado: 'no_gana', tabla: t } };
    if (g.confirmado) return { error: { estado: 'ya_confirmado', tabla: t } };
    registro = {
      ronda: e.ronda, tabla: t, nombre: g.nombre, codigo: g.codigo, figura: e.figura.nombre, premio: e.premio,
      balotas: e.balotas.length, conBalota: g.conBalota, hora: formatearHora_(new Date())
    };
    e.ganadores.push(registro);
    // El anuncio lleva además el cartón (grilla y figura) para que la TV lo dibuje; el historial no.
    e.anuncio = Object.assign({}, registro, { grilla: g.grilla, mascara: g.mascara });
    return { extra: { ganador: registro } };
  });
  if (registro) {
    try { anotarGanadorEnHoja_(registro); } catch (err) { r.avisoHoja = 'No se pudo anotar en la hoja "' + HOJA_GANADORES + '": ' + String((err && err.message) || err); }
  }
  return r;
}

function anotarGanadorEnHoja_(g) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let hoja = ss.getSheetByName(HOJA_GANADORES);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_GANADORES);
    hoja.appendRow(['Fecha y hora', 'Ronda', 'Tabla N°', 'Nombre', 'Código', 'Figura', 'Premio (COP)', 'Balotas cantadas', 'Completó con']);
    hoja.setFrozenRows(1);
  }
  hoja.appendRow([g.hora, g.ronda, g.tabla, g.nombre, g.codigo, g.figura, g.premio, g.balotas, g.conBalota]);
}

// Deja el juego como nuevo antes del evento: ronda 1, sin balotas, sin anuncio, sin historial de ganadores,
// y la hoja "Ganadores" vacía (se conserva la fila 1 de encabezados). La figura y el premio se conservan.
// También deja los sorteos adicionales como nuevos: sin historial ni anuncio, y la hoja "Sorteos" vacía (con encabezados).
// NO toca las hojas "Tablas" ni "Base de datos" ni las claves.
function reiniciarJuego_(confirmacion) {
  if (String(confirmacion || '').trim() !== 'REINICIAR') return { estado: 'error', mensaje: 'Para reiniciar escribe la palabra REINICIAR.' };
  let filasBorradas = 0, avisoHoja = '', filasSorteos = 0;
  const r = mutarEstado_(function (e) {
    try {
      const hs = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_SORTEOS);
      if (hs && hs.getLastRow() >= 2) {
        filasSorteos = hs.getLastRow() - 1;
        hs.getRange(2, 1, filasSorteos, Math.max(1, hs.getLastColumn())).clearContent();
        SpreadsheetApp.flush();
      }
    } catch (err) {
      avisoHoja = 'No se pudo vaciar la hoja "' + HOJA_SORTEOS + '": ' + String((err && err.message) || err);
    }
    const so = leerSorteoDurable_();
    guardarSorteo_({ vs: so.vs, n: 0, ganadores: [], historial: [], anuncio: null, descartados: null });
    try {
      const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_GANADORES);
      if (hoja && hoja.getLastRow() >= 2) {
        filasBorradas = hoja.getLastRow() - 1;
        hoja.getRange(2, 1, filasBorradas, Math.max(1, hoja.getLastColumn())).clearContent();
        SpreadsheetApp.flush();
      }
    } catch (err) {
      avisoHoja = (avisoHoja ? avisoHoja + ' ' : '') + 'No se pudo vaciar la hoja "' + HOJA_GANADORES + '": ' + String((err && err.message) || err);
    }
    e.ronda = 1; e.balotas = []; e.anuncio = null; e.ganadores = [];
  });
  r.reiniciado = true;
  r.filasBorradas = filasBorradas;
  r.filasSorteos = filasSorteos;
  if (avisoHoja) r.avisoHoja = avisoHoja;
  return r;
}

function listarTablas_() {
  const tablas = tablasJuego_();
  const inactivas = tablas.filter(function (t) { return !t.a; }).map(function (t) { return { tabla: t.t, serie: t.s, nota: t.o }; });
  // Para el control: las inactivas y las que tienen nota (p. ej. 0762 "por confirmar"), aunque ya estén activas.
  const revisar = tablas.filter(function (t) { return !t.a || String(t.o).trim() !== ''; })
    .map(function (t) { return { tabla: t.t, serie: t.s, nota: t.o, activa: t.a }; });
  return { estado: 'tablas', total: tablas.length, activas: tablas.length - inactivas.length, inactivas: inactivas, revisar: revisar };
}

// Activa o desactiva una tabla (p. ej. 0762 "por confirmar"). Solo cambia la columna Activa.
function activarTabla_(tabla, activa) {
  const t = limpiarTabla_(tabla);
  if (t === null) return { estado: 'error', mensaje: 'Número de tabla no válido' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_TABLAS);
    if (!hoja || hoja.getLastRow() < 2) return { estado: 'error', mensaje: 'Aún no se han cargado las tablas.' };
    const valores = hoja.getRange(1, 1, hoja.getLastRow(), hoja.getLastColumn()).getValues();
    const enc = valores[0].map(normalizarTexto_);
    const cT = enc.indexOf('tabla'), cA = enc.indexOf('activa');
    if (cT < 0 || cA < 0) return { estado: 'error', mensaje: 'Faltan las columnas Tabla o Activa en la hoja "' + HOJA_TABLAS + '".' };
    let fila = -1;
    for (let i = 1; i < valores.length; i++) if (limpiarTabla_(valores[i][cT]) === t) { fila = i + 1; break; }
    if (fila < 0) return { estado: 'tabla_no_existe', tabla: t };
    hoja.getRange(fila, cA + 1).setValue(activa === true);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(CACHE_TABLAS);
  } finally {
    lock.releaseLock();
  }
  return Object.assign(listarTablas_(), { cambiada: t, activa: activa === true });
}

// Reemplaza el contenido de la hoja "Tablas" con las filas recibidas, después de validarlas todas.
function cargarTablas_(filas) {
  if (!Array.isArray(filas) || filas.length === 0 || filas.length > 5000) return { estado: 'error', mensaje: 'Se esperaba una lista de tablas.' };
  const errores = [], vistas = {}, salida = [];
  const celdas = CELDAS.filter(function (k) { return k !== 'N3'; });
  filas.forEach(function (f, i) {
    const t = limpiarTabla_(f && f.tabla);
    if (t === null) { errores.push('fila ' + (i + 1) + ': tabla no válida'); return; }
    if (vistas[t]) { errores.push('tabla ' + t + ' repetida'); return; }
    vistas[t] = true;
    const nums = celdas.map(function (k) { return f[k]; });
    const malos = celdas.filter(function (k, j) {
      const n = nums[j], r = RANGOS[k.charAt(0)];
      return !(typeof n === 'number' && Number.isInteger(n) && n >= r[0] && n <= r[1]);
    });
    if (malos.length) { errores.push('tabla ' + t + ': fuera de rango ' + malos.join(',')); return; }
    if (new Set(nums).size !== 24) { errores.push('tabla ' + t + ': números repetidos'); return; }
    salida.push([t, String(f.serie || '').slice(0, 100), String(f.pdf || '').slice(0, 100), Number(f.pagina) || '']
      .concat(nums, [f.activa !== false, String(f.nota || '').slice(0, 300)]));
  });
  if (errores.length) return { estado: 'error', mensaje: errores.length + ' filas con errores; no se cargó nada.', errores: errores.slice(0, 20) };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const hoja = ss.getSheetByName(HOJA_TABLAS) || ss.insertSheet(HOJA_TABLAS);
    hoja.clear();
    const encabezados = ['Tabla', 'Serie', 'PDF', 'Página'].concat(celdas, ['Activa', 'Nota']);
    hoja.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
    hoja.getRange(2, 1, salida.length, encabezados.length).setValues(salida);
    hoja.getRange(2, encabezados.length - 1, salida.length, 1).insertCheckboxes();
    hoja.setFrozenRows(1);
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove(CACHE_TABLAS);
  } finally {
    lock.releaseLock();
  }
  const activas = salida.filter(function (r) { return r[r.length - 2] === true; }).length;
  return { estado: 'tablas_cargadas', total: salida.length, activas: activas, inactivas: salida.length - activas };
}

// =====================================================================
// ===== SORTEOS ADICIONALES (aparte del bingo) =====
// Participan los códigos de la pestaña "Códigos" (se busca sin importar mayúsculas ni tildes) que además
// tienen ingreso registrado en "Base de datos" (H = TRUE). Un código de la pestaña que no está en
// "Base de datos" no participa. No se repiten ganadores. Cada resultado se anota en la hoja "Sorteos".
// Estado propio en Propiedades (ESTADO_SORTEO) + caché: no toca el estado ni la versión del bingo.
// =====================================================================

const HOJA_SORTEOS = 'Sorteos';
const NOMBRES_PESTANA_CODIGOS = ['codigos', 'codigos adicionales'];
const PROP_SORTEO = 'ESTADO_SORTEO';
const CACHE_SORTEO = 'estado_sorteo';
const MAX_HISTORIAL_SORTEO = 25;
const TAM_RULETA = 30;
const ENCABEZADOS_SORTEOS = ['Fecha', 'Hora', 'N° sorteo', 'Premio', 'Código', 'Nombre', 'Estado', 'Género'];
// Forma de un código: PMB-121-T, ABC-12, X1-2-3 … (letras/números separados por guiones)
const FORMA_CODIGO = /^[A-Z0-9]{1,8}(-[A-Z0-9]{1,8}){1,4}$/;

function sorteoInicial_() {
  return { vs: 0, n: 0, ganadores: [], historial: [], anuncio: null, descartados: null };
}

function normalizarSorteo_(s) {
  const base = sorteoInicial_();
  s = s || {};
  Object.keys(base).forEach(function (k) { if (s[k] === undefined) s[k] = base[k]; });
  return s;
}

function leerSorteo_() {
  const c = CacheService.getScriptCache().get(CACHE_SORTEO);
  if (c) return normalizarSorteo_(JSON.parse(c));
  const s = leerSorteoDurable_();
  CacheService.getScriptCache().put(CACHE_SORTEO, JSON.stringify(s), 21600);
  return s;
}

function leerSorteoDurable_() {
  const p = PropertiesService.getScriptProperties().getProperty(PROP_SORTEO);
  return normalizarSorteo_(p ? JSON.parse(p) : null);
}

// Guarda el estado del sorteo (llamar con el lock tomado). Sube vs para que la TV y el control se enteren.
function guardarSorteo_(s) {
  s.vs = (Number(s.vs) || 0) + 1;
  if (s.historial.length > MAX_HISTORIAL_SORTEO) s.historial = s.historial.slice(-MAX_HISTORIAL_SORTEO);
  let texto = JSON.stringify(s);
  while (texto.length > 8500 && s.historial.length > 5) { s.historial.shift(); texto = JSON.stringify(s); }
  PropertiesService.getScriptProperties().setProperty(PROP_SORTEO, texto);
  CacheService.getScriptCache().put(CACHE_SORTEO, texto, 21600);
}

// Busca la pestaña de códigos sin importar mayúsculas, tildes ni espacios: "Códigos", "codigos", "CODIGOS"…
// No toma copias como "Códigos (2)".
function buscarPestanaCodigos_(ss) {
  const hojas = ss.getSheets ? ss.getSheets() : [];
  for (let k = 0; k < NOMBRES_PESTANA_CODIGOS.length; k++) {
    for (let i = 0; i < hojas.length; i++) {
      if (normalizarTexto_(hojas[i].getName()).replace(/\s+/g, ' ') === NOMBRES_PESTANA_CODIGOS[k]) return hojas[i];
    }
  }
  return null;
}

function pareceCodigo_(v) {
  return typeof v === 'string' && FORMA_CODIGO.test(limpiarCodigo_(v));
}

// Lee la pestaña de códigos y se adapta a su forma:
//  - Con encabezados (una celda de la fila 1 dice "Código…"): esa columna; nombre = la que diga "Nombre…".
//  - Sin encabezados: la columna con más valores con forma de código; nombre = la columna siguiente si tiene texto.
// Ignora casillas (TRUE/FALSE), errores (#N/A), vacías y repetidos.
function leerPestanaCodigos_(ss) {
  const hoja = buscarPestanaCodigos_(ss);
  if (!hoja) return { error: 'No se encontró la pestaña de códigos ("Códigos") en el Sheet.' };
  const ultF = hoja.getLastRow(), ultC = hoja.getLastColumn();
  const valores = (ultF >= 1 && ultC >= 1) ? hoja.getRange(1, 1, ultF, ultC).getValues() : [];
  let cCod = -1, cNom = -1, desde = 0;
  const enc = (valores[0] || []).map(function (v) { return typeof v === 'string' ? normalizarTexto_(v) : ''; });
  const conEncabezado = enc.some(function (t) { return /^codigo/.test(t); }) && !(valores[0] || []).some(pareceCodigo_);
  if (conEncabezado) {
    desde = 1;
    cCod = enc.findIndex(function (t) { return /^codigo/.test(t); });
    cNom = enc.findIndex(function (t) { return /^nombre/.test(t); });
  } else {
    let mejor = 0;
    for (let c = 0; c < ultC; c++) {
      let n = 0;
      for (let f = 0; f < valores.length; f++) if (pareceCodigo_(valores[f][c])) n++;
      if (n > mejor) { mejor = n; cCod = c; }
    }
    if (cCod >= 0 && cCod + 1 < ultC) {
      let texto = 0, conCodigo = 0;
      for (let f = 0; f < valores.length; f++) {
        if (!pareceCodigo_(valores[f][cCod])) continue;
        conCodigo++;
        const v = valores[f][cCod + 1];
        if (typeof v === 'string' && v.trim() && !pareceCodigo_(v)) texto++;
      }
      if (conCodigo && texto >= conCodigo / 2) cNom = cCod + 1;
    }
  }
  const lista = [], vistos = {};
  if (cCod >= 0) {
    for (let f = desde; f < valores.length; f++) {
      const v = valores[f][cCod];
      if (!pareceCodigo_(v)) continue;
      const cod = limpiarCodigo_(v);
      if (vistos[cod]) continue;
      vistos[cod] = true;
      lista.push({ codigo: cod, nombre: cNom >= 0 ? String(valores[f][cNom] || '').replace(/\s+/g, ' ').trim() : '' });
    }
  }
  const letra = function (c) { return c >= 0 ? String.fromCharCode(65 + c) : ''; };
  return { hoja: hoja.getName(), encabezado: conEncabezado, colCodigo: letra(cCod), colNombre: letra(cNom), codigos: lista };
}

// Participantes = códigos de la pestaña con fila en "Base de datos" y H = TRUE.
function participantesSorteo_() {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const pc = leerPestanaCodigos_(ss);
  if (pc.error) return pc;
  const base = leerHoja_(abrirHoja_());
  const enBase = {};
  base.filas.forEach(function (f) {
    const cod = limpiarCodigo_(f[COL_CODIGO - 1]);
    if (!cod) return;
    const x = enBase[cod] || (enBase[cod] = { asistio: false, nombre: '' });
    if (f[COL_ASISTIO - 1] === true) x.asistio = true;
    if (!x.nombre) x.nombre = String(f[COL_NOMBRE - 1] || '').replace(/\s+/g, ' ').trim();
  });
  const participantes = [];
  let noEnBase = 0, sinIngreso = 0;
  pc.codigos.forEach(function (c) {
    const b = enBase[c.codigo];
    if (!b) { noEnBase++; return; }
    if (!b.asistio) { sinIngreso++; return; }
    participantes.push({ codigo: c.codigo, nombre: c.nombre || b.nombre, nombreBase: b.nombre });
  });
  pc.participantes = participantes;
  pc.total = pc.codigos.length;
  pc.noEnBase = noEnBase;
  pc.sinIngreso = sinIngreso;
  clasificarParticipantes_(pc, ss, base.filas.map(function (f) { return String(f[COL_NOMBRE - 1] || ''); }));
  return pc;
}

function normPremio_(t) { return normalizarTexto_(t).replace(/\s+/g, ' '); }

// Quiénes pueden salir ahora: participantes que no han ganado; si se está repitiendo el sorteo de un premio
// después de descartar, tampoco los descartados para ese mismo premio.
function elegibles_(pc, s, premio, filtro) {
  const ganaron = {};
  s.ganadores.forEach(function (c) { ganaron[c] = true; });
  const fuera = {};
  if (s.descartados && s.descartados.premio === normPremio_(premio)) s.descartados.codigos.forEach(function (c) { fuera[c] = true; });
  return pc.participantes.filter(function (p) { return !ganaron[p.codigo] && !fuera[p.codigo] && pasaFiltro_(p, filtro); });
}

function resumenSorteo_(pc, s, premio, filtro) {
  filtro = limpiarFiltro_(filtro);
  const r = {
    estado: 'sorteo', vs: s.vs, hoja: pc.hoja, colCodigo: pc.colCodigo, colNombre: pc.colNombre, encabezado: pc.encabezado,
    total: pc.total, participantes: pc.participantes.length, noEnBase: pc.noEnBase, sinIngreso: pc.sinIngreso,
    yaGanaron: s.ganadores.length, historial: s.historial.slice().reverse(), anuncio: s.anuncio
  };
  const pr = premio === undefined ? '' : premio;
  r.disponibles = elegibles_(pc, s, pr, filtro).length;
  // Género (solo para el sorteo): cuentas entre los que ingresaron y disponibles por filtro.
  r.filtro = filtro;
  r.genero = pc.genero;
  r.disponiblesPorFiltro = { todos: elegibles_(pc, s, pr, 'todos').length, mujeres: elegibles_(pc, s, pr, 'mujeres').length, hombres: elegibles_(pc, s, pr, 'hombres').length };
  r.revisar = listaRevisarGenero_(pc);
  if (s.descartados) r.repitiendo = { premio: s.descartados.premio, descartados: s.descartados.codigos.length };
  return r;
}

function infoSorteo_(filtro) {
  const pc = participantesSorteo_();
  if (pc.error) return { estado: 'error', mensaje: pc.error };
  return resumenSorteo_(pc, leerSorteo_(), undefined, filtro);
}

function limpiarPremioSorteo_(premio) {
  return String(premio === undefined || premio === null ? '' : premio).replace(/\s+/g, ' ').trim().slice(0, 60);
}

// Elige al azar (con el lock tomado). Devuelve el registro o null si no hay a quién sortear.
function sortearInterno_(pc, s, premio, avisos, filtro) {
  filtro = limpiarFiltro_(filtro);
  const lista = elegibles_(pc, s, premio, filtro);
  if (!lista.length) return null;
  const g = lista[Math.floor(Math.random() * lista.length)];
  if (s.descartados && s.descartados.premio !== normPremio_(premio)) s.descartados = null;
  s.n = (Number(s.n) || 0) + 1;
  const ahora = new Date();
  const reg = {
    n: s.n, premio: premio, codigo: g.codigo, nombre: g.nombre, estado: 'ganador', filtro: filtro,
    fecha: Utilities.formatDate(ahora, ZONA, 'dd/MM/yyyy'), hora: Utilities.formatDate(ahora, ZONA, 'HH:mm:ss')
  };
  s.ganadores.push(g.codigo);
  s.historial.push(reg);
  // Códigos para la animación de la TV: una muestra de participantes (el ganador va al final).
  const ruleta = [];
  const otros = lista.filter(function (p) { return p.codigo !== g.codigo; }).map(function (p) { return p.codigo; });
  for (let i = otros.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = otros[i]; otros[i] = otros[j]; otros[j] = t; }
  for (let i = 0; i < Math.min(TAM_RULETA, otros.length); i++) ruleta.push(otros[i]);
  // uid: único aunque la numeración vuelva a 1 tras reiniciar (la TV lo usa para no repetir la animación).
  s.anuncio = {
    id: s.n, uid: s.n + '-' + Date.now().toString(36) + '-' + Math.floor(Math.random() * 1e9).toString(36), premio: premio, codigo: g.codigo, nombre: g.nombre, hora: reg.hora,
    participantes: lista.length, filtro: filtro, ruleta: ruleta
  };
  try {
    anotarSorteoEnHoja_(reg);
  } catch (err) {
    avisos.push('No se pudo anotar en la hoja "' + HOJA_SORTEOS + '": ' + String((err && err.message) || err));
  }
  return reg;
}

function anotarSorteoEnHoja_(reg) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let hoja = ss.getSheetByName(HOJA_SORTEOS);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_SORTEOS);
    hoja.appendRow(ENCABEZADOS_SORTEOS);
    hoja.setFrozenRows(1);
  } else if (String(hoja.getRange(1, 8, 1, 1).getValues()[0][0] || '') === '') {
    hoja.getRange(1, 8).setValue('Género');   // hoja creada antes de existir el filtro
  }
  hoja.appendRow([reg.fecha, reg.hora, reg.n, reg.premio, reg.codigo, reg.nombre, 'Ganador', nombreFiltro_(reg.filtro)]);
}

// Marca como "Descartado" la fila del sorteo n en la hoja (busca por N° sorteo y código).
function marcarDescartadoEnHoja_(reg) {
  const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_SORTEOS);
  if (!hoja || hoja.getLastRow() < 2) return false;
  const filas = hoja.getRange(2, 1, hoja.getLastRow() - 1, ENCABEZADOS_SORTEOS.length).getValues();
  for (let i = filas.length - 1; i >= 0; i--) {
    if (Number(filas[i][2]) === reg.n && limpiarCodigo_(filas[i][4]) === reg.codigo) {
      hoja.getRange(i + 2, 7).setValue('Descartado');
      return true;
    }
  }
  return false;
}

function conLockSorteo_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sortear_(premioDado, filtroDado) {
  const premio = limpiarPremioSorteo_(premioDado);
  const filtro = limpiarFiltro_(filtroDado);
  if (!premio) return { estado: 'error', mensaje: 'Escribe el nombre del premio antes de sortear.' };
  const pc = participantesSorteo_();
  if (pc.error) return { estado: 'error', mensaje: pc.error };
  return conLockSorteo_(function () {
    const s = leerSorteoDurable_();
    const avisos = [];
    const reg = sortearInterno_(pc, s, premio, avisos, filtro);
    if (!reg) {
      const r = resumenSorteo_(pc, s, premio, filtro);
      r.estado = 'sin_participantes';
      const delFiltro = pc.participantes.filter(function (p) { return pasaFiltro_(p, filtro); }).length;
      r.mensaje = filtro !== 'todos' && pc.participantes.length && !delFiltro
        ? 'No hay participantes para "' + nombreFiltro_(filtro).toLowerCase() + '": ninguno de los que ingresaron está clasificado como ' + (filtro === 'mujeres' ? 'mujer' : 'hombre') + '.' +
          (pc.genero.sinClasificar ? ' Hay ' + pc.genero.sinClasificar + ' sin clasificar: revísalos en "Revisar género".' : '')
        : pc.participantes.length
        ? 'No quedan participantes: todos los que ingresaron ya ganaron' + (s.descartados && s.descartados.premio === normPremio_(premio) ? ' o fueron descartados para este premio' : '') + '.'
        : 'No hay participantes: ningún código de la pestaña "' + pc.hoja + '" tiene ingreso registrado.';
      return r;
    }
    guardarSorteo_(s);
    const r = resumenSorteo_(pc, s, premio, filtro);
    r.ganador = reg;
    if (avisos.length) r.avisoHoja = avisos.join(' ');
    return r;
  });
}

// Descarta el último resultado (no cuenta como ganador) y, si se pide, vuelve a sortear el mismo premio
// sin ese código.
function descartarSorteo_(volverASortear) {
  const pc = participantesSorteo_();
  if (pc.error) return { estado: 'error', mensaje: pc.error };
  return conLockSorteo_(function () {
    const s = leerSorteoDurable_();
    const ult = s.historial[s.historial.length - 1];
    if (!ult || ult.estado !== 'ganador') {
      const r = resumenSorteo_(pc, s);
      r.estado = 'sin_resultado';
      r.mensaje = 'No hay un resultado para descartar.';
      return r;
    }
    const avisos = [];
    ult.estado = 'descartado';
    const i = s.ganadores.lastIndexOf(ult.codigo);
    if (i >= 0) s.ganadores.splice(i, 1);
    const np = normPremio_(ult.premio);
    if (!s.descartados || s.descartados.premio !== np) s.descartados = { premio: np, codigos: [] };
    if (s.descartados.codigos.indexOf(ult.codigo) < 0) s.descartados.codigos.push(ult.codigo);
    if (s.anuncio && s.anuncio.id === ult.n) s.anuncio = null;
    try {
      if (!marcarDescartadoEnHoja_(ult)) avisos.push('No se encontró la fila del sorteo ' + ult.n + ' en la hoja "' + HOJA_SORTEOS + '".');
    } catch (err) {
      avisos.push('No se pudo marcar en la hoja "' + HOJA_SORTEOS + '": ' + String((err && err.message) || err));
    }
    let reg = null;
    if (volverASortear) reg = sortearInterno_(pc, s, ult.premio, avisos, ult.filtro);
    guardarSorteo_(s);
    const r = resumenSorteo_(pc, s, ult.premio, ult.filtro);
    r.descartado = ult;
    if (reg) r.ganador = reg;
    if (volverASortear && !reg) {
      r.estado = 'sin_participantes';
      r.mensaje = 'Se descartó ' + ult.codigo + ', pero no quedan participantes' + (ult.filtro && ult.filtro !== 'todos' ? ' (' + nombreFiltro_(ult.filtro).toLowerCase() + ')' : '') + ' para volver a sortear "' + ult.premio + '".';
    }
    if (avisos.length) r.avisoHoja = avisos.join(' ');
    return r;
  });
}

function ocultarSorteo_() {
  return conLockSorteo_(function () {
    const s = leerSorteoDurable_();
    s.anuncio = null;
    guardarSorteo_(s);
    return { estado: 'sorteo_oculto', vs: s.vs, anuncio: null };
  });
}

// =====================================================================
// ===== GÉNERO PARA SORTEOS FILTRADOS (solo mujeres / solo hombres) =====
// No usa ninguna columna de género en "Base de datos": se deduce del primer nombre de la columna
// "Nombre Completo" (en la hoja viene primero el nombre y luego los apellidos). Los nombres compuestos se
// deciden por el primero ("María José" -> mujer, "José María" -> hombre, "Juan David" -> hombre).
// Confianza: 'alta' (nombre de la lista), 'baja' (por terminación o por el segundo nombre; participa, pero
// se lista para revisar) y 'sin' (sin clasificar: unisex, extranjero o desconocido; no participa en un
// sorteo filtrado hasta que se revise). Las correcciones de la hoja "Género" (Código, Género H/M) mandan.
// Solo lo usa el sorteo: no toca el bingo ni el ingreso.
// =====================================================================

const HOJA_GENERO = 'Género';
const ENCABEZADOS_GENERO = ['Código', 'Género'];
const FILTROS_SORTEO = { todos: '', mujeres: 'M', hombres: 'H' };

const NOMBRES_MUJER_ = (
  'maria ana luz carmen rosa martha marta gloria blanca sandra claudia patricia adriana diana paola andrea carolina catalina ' +
  'natalia nathalia nataly nathaly daniela valentina valeria sofia isabella isabel camila mariana gabriela laura lorena liliana lilian liliam ' +
  'viviana veronica monica angela angelica alejandra fernanda juliana luisa paula juana johana johanna yohana yuliana yulieth yuli yuly ' +
  'julieth juliet jessica jesica yessica yesica jenny yenny jennifer jenifer yenifer yennifer kelly karen karina katherine katerine kateryn catherine ' +
  'leidy leydi lady deisy deysi daisy dayana eliana elena helena estefania stefania stephanie tatiana tania sara sarah salome antonella ' +
  'milena yamile yamileth yamilet yaneth janeth janet yanet nancy nelly nidia nubia nohora norma olga piedad pilar socorro amparo consuelo ' +
  'mercedes dolores esperanza fanny francy beatriz teresa tereza cecilia celia lucia lucila lucy luzmila ludivia ligia lida lina marleny ' +
  'marlene marlen marly maritza mayra maira melisa melissa mery merly mary miriam myriam olivia omaira oneida rubiela ruby ruth shirley ' +
  'sonia stella estela susana xiomara yolanda yesenia yudy judy yurany yuranis zulma edelmira emilse elsa elvira erika erica esther eugenia ' +
  'emilia emma fabiola flor floralba gladys graciela gisela hilda ines ingrid irene iris jacqueline jaqueline jakeline jazmin yasmin juanita ' +
  'julia leonor libia lilia magdalena margarita marcela mariela marisol marina martina matilde maryuri milady mildred nelsy noemi noralba ' +
  'odilia ofelia otilia rocio rosalba rosmira silvia soraya tulia vanessa vanesa victoria wendy ximena jimena yadira yeimi yeimy yina ' +
  'yolima yorladis yuleidy zoraida adiela aleyda aracelly aracely amanda amalia carlina cielo cielito cindy damaris dary elizabeth ' +
  'eliceth emelida geraldine imelda lucedia luzdary samara selena serena alba alicia aura berta bertha betty brigitte celeste clara cristina ' +
  'dora doris edith elisa eloisa emperatriz evelyn fatima francia hortensia ivon ivonne josefa josefina leticia lizeth lisseth lizet ' +
  'luciana mabel mireya nora pamela regina rebeca sabrina samantha sharon tiffany valery zoila faisully libelly miguelina maribel ' +
  'alexandra angie anyi arelis belly carmenza deyanira dilia dina edna eunice gina greys heidy jhoana jhojana karla keila leidi ' +
  'lorely luzmary mariluz marilyn mirian nayeli nelcy nury rosaura sulay yajaira yurley yuliza zuleima belkis dorys layla leila marily'
).split(' ');

const NOMBRES_HOMBRE_ = (
  'jose juan carlos luis jorge manuel miguel pedro pablo andres felipe david daniel diego alejandro sebastian santiago mateo matias ' +
  'nicolas samuel gabriel rafael fernando francisco javier jaime jairo jhon john jhoan johan yohan jhonatan jonathan jonatan johnny jhonny ' +
  'alexander alberto alfonso alfredo alvaro antonio armando arnulfo arturo augusto bernardo brayan brian bryan camilo cesar cristian ' +
  'christian cristobal dario edgar edison edwin efrain eduardo elkin emilio enrique ernesto esteban fabian fabio faber federico freddy ' +
  'fredy german gerardo gilberto gonzalo guillermo gustavo hector hernan hernando holman horacio hugo humberto ignacio isaac ismael ivan ' +
  'jaider jader jefferson yeferson jeferson jesus joaquin julian julio kevin leonardo leandro lisandro lucas luciano marco marcos mario ' +
  'martin mauricio maximiliano milton moises nelson norberto octavio omar orlando oscar pascual raul ramiro ramon reinaldo ricardo roberto ' +
  'rodrigo rodolfo ruben salvador sergio silvio teodoro tomas ulises uriel victor wilson william willian wilmer yesid yeison jeison yair ' +
  'harold harrison heider henry herney hoover americo didier duvan duban dairo darwin deiby deivis edier eider elmer emerson faiver ' +
  'ferney gildardo hamilton hermes james jhojan kaleth libardo marlon nestor noe ovidio pastor rigoberto robinson rogelio rosendo saul ' +
  'segundo silvestre stiven steven estiven teofilo valentin wilmar wilfredo yamid yerson albeiro alirio aldemar abel adolfo agustin alonso ' +
  'anderson angel aurelio baltazar benjamin bladimir vladimir carmelo celso ciro clemente dagoberto damian danilo edilberto eliecer elias ' +
  'emiro eugenio ezequiel fidel florentino gregorio heriberto jacinto jeronimo joel jonny josue leon leonel lorenzo marcelino marcelo ' +
  'mariano maximo napoleon nicanor oliverio osvaldo placido porfirio reynaldo samir simon tobias wilber yhon yeiner yorman jhan jhordan ' +
  'jordan brandon kener luber ever elvis dennis adrian aldair cristhian dilan dylan emanuel emmanuel erick eric esneider ' +
  'gerson harvey jair jhonier yonier jorman keiner maicol michael maycol miller neider nixon oswaldo sneider wilder yilmer yordi yordy ' +
  'edgardo alcides arley duberney anibal benito bayron byron cristo deiver eduar elver evaristo eusebio feliciano ' +
  'gentil hugo isidro jaminton jhonson leider luciano maikel mairon nolberto onofre patricio reinel sigifredo ' +
  'virgilio wbeimar yeisson yhonatan eider giovanny geovany giovanni yovany jhovany'
).split(' ');

// Unisex o ambiguos: nunca se clasifican solos (aunque la terminación sugiera algo).
const NOMBRES_AMBIGUOS_ = ('alex cruz yuri guadalupe trinidad reyes yury sasha francis andy dany dani karol rene ariel cris kris yan jean ' +
  'alexis noel nikita darcy dayan elian ofir norby amalfi').split(' ');

// Palabras que no son nombres: tratamientos y partículas ("María del Rosario", "Sra. Luz").
const PALABRAS_SALTAR_ = 'sr sra srta don dona dr dra lic prof de del la las los y'.split(' ');

let mapaNombres_ = null;
function mapaNombres_get_() {
  if (mapaNombres_) return mapaNombres_;
  const m = {};
  NOMBRES_MUJER_.forEach(function (n) { if (n) m[n] = 'M'; });
  NOMBRES_HOMBRE_.forEach(function (n) { if (n) m[n] = m[n] && m[n] !== 'H' ? '?' : 'H'; });
  NOMBRES_AMBIGUOS_.forEach(function (n) { m[n] = '?'; });
  mapaNombres_ = m;
  return m;
}

function palabrasNombre_(nombre) {
  return normalizarTexto_(nombre).replace(/[^a-zñ\s]/g, ' ').split(/\s+/)
    .filter(function (p) { return p && PALABRAS_SALTAR_.indexOf(p) < 0; })
    .map(function (p) { return p === 'ma' ? 'maria' : p; });
}

// Deducción por terminación (solo para nombres que no están en la lista).
function generoPorTerminacion_(p) {
  if (p.length < 3) return '';
  if (/(a|eth|ith|yth|ely|elly|ully|illy|ally|ys|dis)$/.test(p)) return 'M';
  if (/(o|son|ton|ier|man|ar)$/.test(p)) return 'H';
  return '';
}

// Orden de la columna de nombres: 'nombre' (nombre primero) o 'apellido' (apellido primero), por votación.
function ordenNombres_(nombres) {
  const m = mapaNombres_get_();
  let primero = 0, despues = 0;
  nombres.forEach(function (n) {
    const p = palabrasNombre_(n);
    if (p.length < 2) return;
    if (m[p[0]] === 'H' || m[p[0]] === 'M') primero++;
    else if (p.slice(1).some(function (x) { return m[x] === 'H' || m[x] === 'M'; })) despues++;
  });
  return despues > primero ? 'apellido' : 'nombre';
}

// { g: 'H' | 'M' | '', conf: 'alta' | 'baja' | 'sin', motivo }
function deducirGenero_(nombre, orden) {
  const m = mapaNombres_get_();
  let p = palabrasNombre_(nombre);
  if (!p.length) return { g: '', conf: 'sin', motivo: 'sin nombre' };
  // Apellido primero: "Mosquera Rivas Luz Marina" -> los nombres empiezan después de los apellidos.
  if (orden === 'apellido' && p.length >= 2) {
    const desde = p.length >= 4 ? 2 : 1;
    let k = -1;
    for (let i = 1; i < p.length; i++) if (m[p[i]] === 'H' || m[p[i]] === 'M') { k = i; break; }
    p = p.slice(k >= 0 ? k : desde);
  }
  const uno = p[0], dos = p.length >= 3 || orden === 'apellido' ? p[1] : undefined;
  const g1 = m[uno], g2 = dos ? m[dos] : undefined;
  if (g1 === 'H' || g1 === 'M') {
    // Compuesto con géneros opuestos ("Luz Carlos"): raro; se pide revisar. María/José segundo es normal.
    if ((g2 === 'H' || g2 === 'M') && g2 !== g1 && ['maria', 'jose', 'rosario'].indexOf(dos) < 0 && uno !== 'maria' && uno !== 'jose') {
      return { g: g1, conf: 'baja', motivo: 'nombres de distinto género' };
    }
    return { g: g1, conf: 'alta', motivo: 'nombre conocido' };
  }
  if (g1 !== '?') {
    const t = generoPorTerminacion_(uno);
    if (t) return { g: t, conf: 'baja', motivo: 'por la terminación' };
  }
  // Primer nombre unisex o desconocido: si el segundo nombre es claro, se usa con baja confianza.
  if (g2 === 'H' || g2 === 'M') return { g: g2, conf: 'baja', motivo: 'por el segundo nombre' };
  return { g: '', conf: 'sin', motivo: g1 === '?' ? 'nombre unisex' : 'nombre desconocido' };
}

// "H", "hombre", "masculino" -> H; "M", "mujer", "F", "femenino" -> M.
function limpiarGenero_(v) {
  const t = normalizarTexto_(v);
  if (/^(h|hombre|masculino|masc)$/.test(t)) return 'H';
  if (/^(m|mujer|f|femenino|fem)$/.test(t)) return 'M';
  return '';
}

function buscarHojaGenero_(ss) {
  const hojas = ss.getSheets ? ss.getSheets() : [];
  for (let i = 0; i < hojas.length; i++) if (normalizarTexto_(hojas[i].getName()) === 'genero') return hojas[i];
  return ss.getSheetByName(HOJA_GENERO);
}

// Correcciones manuales: { CODIGO: 'H' | 'M' }
function leerCorreccionesGenero_(ss) {
  const hoja = buscarHojaGenero_(ss);
  const r = {};
  if (!hoja || hoja.getLastRow() < 2) return r;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(function (f) {
    const c = limpiarCodigo_(f[0]), g = limpiarGenero_(f[1]);
    if (c && g) r[c] = g;
  });
  return r;
}

// Añade a cada participante su género (p.g, p.conf, p.motivo, p.revisado) y deja la cuenta en pc.genero.
function clasificarParticipantes_(pc, ss, nombresBase) {
  const orden = ordenNombres_(nombresBase);
  const corr = leerCorreccionesGenero_(ss);
  let mujeres = 0, hombres = 0, sin = 0, baja = 0;
  pc.participantes.forEach(function (p) {
    const d = deducirGenero_(p.nombreBase || p.nombre, orden);
    p.deducido = d.g; p.conf = d.conf; p.motivo = d.motivo;
    p.revisado = corr[p.codigo] || '';
    p.g = p.revisado || d.g;
    if (p.g === 'M') mujeres++; else if (p.g === 'H') hombres++; else sin++;
    if (!p.revisado && d.conf === 'baja') baja++;
  });
  pc.genero = { orden: orden, mujeres: mujeres, hombres: hombres, sinClasificar: sin, bajaConfianza: baja, correcciones: Object.keys(corr).length };
  return pc;
}

function limpiarFiltro_(f) {
  const t = normalizarTexto_(f);
  return Object.prototype.hasOwnProperty.call(FILTROS_SORTEO, t) ? t : 'todos';
}

function pasaFiltro_(p, filtro) {
  const g = FILTROS_SORTEO[filtro || 'todos'];
  return !g || p.g === g;   // en un sorteo filtrado no participan los sin clasificar
}

function nombreFiltro_(filtro) {
  return filtro === 'mujeres' ? 'Solo mujeres' : filtro === 'hombres' ? 'Solo hombres' : 'Todos';
}

// Participantes (con ingreso) sin clasificar o de baja confianza, y los ya revisados a mano entre ellos.
function listaRevisarGenero_(pc) {
  return pc.participantes
    .filter(function (p) { return p.conf !== 'alta' || p.revisado; })
    .map(function (p) { return { codigo: p.codigo, nombre: p.nombreBase || p.nombre, deducido: p.deducido, conf: p.conf, motivo: p.motivo, revisado: p.revisado }; })
    .sort(function (a, b) { return (a.revisado ? 1 : 0) - (b.revisado ? 1 : 0) || (a.conf === 'sin' ? 0 : 1) - (b.conf === 'sin' ? 0 : 1) || a.nombre.localeCompare(b.nombre); });
}

// Guarda (o cambia) la corrección de un código en la hoja "Género". Solo códigos que están en "Base de datos".
function corregirGenero_(codigoDado, generoDado) {
  const codigo = limpiarCodigo_(codigoDado), g = limpiarGenero_(generoDado);
  if (!codigo) return { estado: 'error', mensaje: 'Falta el código.' };
  if (!g) return { estado: 'error', mensaje: 'El género debe ser H (hombre) o M (mujer).' };
  const base = leerHoja_(abrirHoja_());
  if (!base.filas.some(function (f) { return limpiarCodigo_(f[COL_CODIGO - 1]) === codigo; })) {
    return { estado: 'no_existe', mensaje: 'El código ' + codigo + ' no está en "Base de datos".' };
  }
  conLockSorteo_(function () {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    let hoja = buscarHojaGenero_(ss);
    if (!hoja) {
      hoja = ss.insertSheet(HOJA_GENERO);
      hoja.appendRow(ENCABEZADOS_GENERO);
      hoja.setFrozenRows(1);
    }
    let fila = -1;
    if (hoja.getLastRow() >= 2) {
      const cods = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < cods.length; i++) if (limpiarCodigo_(cods[i][0]) === codigo) { fila = i + 2; break; }
    }
    if (fila > 0) hoja.getRange(fila, 2).setValue(g);
    else hoja.appendRow([codigo, g]);
    SpreadsheetApp.flush();
  });
  const r = infoSorteo_();
  if (r.estado === 'sorteo') r.corregido = { codigo: codigo, genero: g };
  return r;
}
