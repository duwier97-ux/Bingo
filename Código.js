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
  validarClave: 'lectura', estado: 'lectura',
  marcarBalota: 'juego', quitarBalota: 'juego', deshacerBalota: 'juego',
  fijarFigura: 'juego', fijarPremio: 'juego',
  ganadores: 'juego', confirmarGanador: 'juego', ocultarAnuncio: 'juego', nuevaRonda: 'juego', reiniciarJuego: 'juego',
  listarTablas: 'juego', activarTabla: 'juego', cargarTablas: 'juego'
};

function atenderJuego_(accion, datos) {
  const juego = esClaveJuego_(datos.clave);
  const tv = !juego && esClaveTv_(datos.clave);
  if (!juego && !(tv && ACCIONES_JUEGO[accion] === 'lectura')) return { estado: 'clave' };
  const rol = juego ? 'juego' : 'tv';

  switch (accion) {
    case 'validarClave': return { estado: 'clave_ok', rol: rol };
    case 'estado': return estadoPublico_(leerEstado_(), rol);
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
// NO toca las hojas "Tablas" ni "Base de datos" ni las claves.
function reiniciarJuego_(confirmacion) {
  if (String(confirmacion || '').trim() !== 'REINICIAR') return { estado: 'error', mensaje: 'Para reiniciar escribe la palabra REINICIAR.' };
  let filasBorradas = 0, avisoHoja = '';
  const r = mutarEstado_(function (e) {
    try {
      const hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(HOJA_GANADORES);
      if (hoja && hoja.getLastRow() >= 2) {
        filasBorradas = hoja.getLastRow() - 1;
        hoja.getRange(2, 1, filasBorradas, Math.max(1, hoja.getLastColumn())).clearContent();
        SpreadsheetApp.flush();
      }
    } catch (err) {
      avisoHoja = 'No se pudo vaciar la hoja "' + HOJA_GANADORES + '": ' + String((err && err.message) || err);
    }
    e.ronda = 1; e.balotas = []; e.anuncio = null; e.ganadores = [];
  });
  r.reiniciado = true;
  r.filasBorradas = filasBorradas;
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
