/**
 * ═══════════════════════════════════════════════════════════════════════════
 * PUENTES DEL MTC — Consolidación de archivos de los órganos (OGAME)
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * CÓMO FUNCIONA
 *   1. Cada órgano envía su Excel. Se deja, tal como llegó, en su subcarpeta del Drive:
 *        Puentes – Fuentes / PVN    → inventario de puentes de la Red Vial Nacional
 *        Puentes – Fuentes / DGPPT  → inventario general de puentes de las concesiones
 *        Puentes – Fuentes / PVD    → base del estado situacional de puentes de PVD
 *   2. En esta hoja: menú Puentes → Actualizar desde Drive.
 *   3. El script toma el archivo más reciente de cada subcarpeta, lo lee y reescribe
 *      la pestaña del órgano. Las fuentes NO se mezclan: una pestaña por órgano.
 *   4. La pestaña "Control" registra fecha, archivo, puentes leídos y advertencias.
 *      Si un archivo no se puede leer, su pestaña NO se modifica (queda la versión anterior).
 *
 * REQUISITO: servicio avanzado "Drive API" activado (Servicios → + → Drive API),
 * necesario para leer archivos Excel.
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────────────────
// ID de la carpeta "Puentes – Fuentes" (la parte de la URL después de /folders/)
var CARPETA_FUENTES_ID = 'PEGAR_AQUI_EL_ID_DE_LA_CARPETA';

var FUENTES = [
  { clave: 'PVN',   subcarpeta: 'PVN',   pestana: 'PVN',   leer: function (v) { return leerPVN(v); } },
  { clave: 'DGPPT', subcarpeta: 'DGPPT', pestana: 'DGPPT', leer: function (v) { return leerDGPPT(v); } },
  { clave: 'PVD',   subcarpeta: 'PVD',   pestana: 'PVD',   leer: function (v) { return leerPVD(v); } }
];
var PESTANA_CONTROL = 'Control';
var TZ = 'America/Lima';

// ── MENÚ ───────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Puentes')
    .addItem('Actualizar desde Drive (los tres)', 'actualizarTodo')
    .addSeparator()
    .addItem('Actualizar solo PVN', 'actualizarPVN')
    .addItem('Actualizar solo DGPPT', 'actualizarDGPPT')
    .addItem('Actualizar solo PVD', 'actualizarPVD')
    .addToUi();
}
function actualizarTodo()  { procesar_(FUENTES); }
function actualizarPVN()   { procesar_([FUENTES[0]]); }
function actualizarDGPPT() { procesar_([FUENTES[1]]); }
function actualizarPVD()   { procesar_([FUENTES[2]]); }

// ── PROCESO ────────────────────────────────────────────────────────────────
function procesar_(lista) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var raiz = DriveApp.getFolderById(CARPETA_FUENTES_ID);
  var resumen = [];
  lista.forEach(function (fu) {
    var reg = { fuente: fu.clave, archivo: '', modificado: '', hoja: '', puentes: 0, avisos: '', resultado: '' };
    try {
      var archivo = archivoMasReciente_(raiz, fu.subcarpeta);
      reg.archivo = archivo.getName();
      reg.modificado = Utilities.formatDate(archivo.getLastUpdated(), TZ, 'dd/MM/yyyy HH:mm');
      var hojas = leerHojas_(archivo);
      // se usa la hoja del archivo que produce más puentes válidos
      var mejor = null;
      Object.keys(hojas).forEach(function (nombre) {
        var r = fu.leer(hojas[nombre]);
        if (r.ok && (!mejor || r.filas.length > mejor.r.filas.length)) mejor = { nombre: nombre, r: r };
        if (!mejor && !r.ok) reg.avisos = r.error;
      });
      if (!mejor) throw new Error(reg.avisos || 'No se encontró una hoja con el formato esperado');
      escribirPestana_(ss, fu.pestana, mejor.r.cabecera, mejor.r.filas);
      reg.hoja = mejor.nombre;
      reg.puentes = mejor.r.filas.length;
      reg.avisos = mejor.r.avisos.join(' · ');
      reg.resultado = 'Actualizado';
    } catch (e) {
      reg.resultado = 'NO actualizado: ' + e.message;
    }
    registrarControl_(ss, reg);
    resumen.push(fu.clave + ': ' + (reg.resultado === 'Actualizado' ? reg.puentes + ' puentes' : reg.resultado));
  });
  SpreadsheetApp.getUi().alert('Resultado de la actualización', resumen.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function archivoMasReciente_(raiz, nombreSub) {
  var subs = raiz.getFoldersByName(nombreSub);
  if (!subs.hasNext()) throw new Error('No existe la subcarpeta "' + nombreSub + '"');
  var it = subs.next().getFiles(), mejor = null;
  while (it.hasNext()) {
    var f = it.next(), tipo = f.getMimeType(), nombre = f.getName().toLowerCase();
    var esHoja = tipo === MimeType.GOOGLE_SHEETS || /\.(xlsx|xlsm|xls)$/.test(nombre);
    if (esHoja && (!mejor || f.getLastUpdated() > mejor.getLastUpdated())) mejor = f;
  }
  if (!mejor) throw new Error('No hay archivos Excel en la subcarpeta "' + nombreSub + '"');
  return mejor;
}

// Devuelve { nombreHoja: [[textos]] }. Los Excel se convierten a una copia temporal de Google Sheets.
function leerHojas_(archivo) {
  var temporal = null, id = archivo.getId();
  if (archivo.getMimeType() !== MimeType.GOOGLE_SHEETS) {
    temporal = Drive.Files.create({ name: 'tmp_puentes_' + archivo.getName(), mimeType: MimeType.GOOGLE_SHEETS }, archivo.getBlob());
    id = temporal.id;
  }
  try {
    var salida = {};
    SpreadsheetApp.openById(id).getSheets().forEach(function (h) {
      if (h.getLastRow() > 1) salida[h.getName()] = h.getDataRange().getDisplayValues();
    });
    return salida;
  } finally {
    if (temporal) DriveApp.getFileById(temporal.id).setTrashed(true);
  }
}

function escribirPestana_(ss, nombre, cabecera, filas) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.clearContents();
  var datos = [cabecera].concat(filas.map(function (f) { return f.map(function (x) { return x == null ? '' : x; }); }));
  var rango = h.getRange(1, 1, datos.length, cabecera.length);
  rango.setNumberFormat('@');          // todo como texto: evita que Sheets cambie códigos o progresivas
  rango.setValues(datos);
  h.getRange(1, 1, 1, cabecera.length).setFontWeight('bold').setBackground('#E8F0FA');
  h.setFrozenRows(1);
}

function registrarControl_(ss, reg) {
  var h = ss.getSheetByName(PESTANA_CONTROL) || ss.insertSheet(PESTANA_CONTROL);
  var cab = ['Fuente', 'Archivo', 'Modificado', 'Cargado', 'Hoja leída', 'Puentes', 'Advertencias', 'Resultado'];
  var vals = h.getLastRow() > 0 ? h.getDataRange().getValues() : [];
  if (!vals.length) { vals = [cab]; }
  var fila = [reg.fuente, reg.archivo, reg.modificado, Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'),
              reg.hoja, reg.puentes, reg.avisos, reg.resultado];
  // una fila por fuente con su última carga (la página del Hub lee esta pestaña)
  var i = vals.findIndex(function (v, k) { return k > 0 && v[0] === reg.fuente; });
  if (i > 0) vals[i] = fila; else vals.push(fila);
  h.clearContents();
  h.getRange(1, 1, vals.length, cab.length).setValues(vals);
  h.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#E8F0FA');
  h.setFrozenRows(1);
}

// ═══════════════════════════════════════════════════════════════════════════
// REGLAS DE LECTURA (funciones puras: reciben una tabla de textos y devuelven filas)
// Flexibles: buscan la fila de encabezados, reconocen columnas por su nombre
// (sin importar mayúsculas, tildes ni espacios) y no fallan ante columnas nuevas.
// ═══════════════════════════════════════════════════════════════════════════

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[\u2212\u2013\u2014]/g, '-').replace(/\s+/g, ' ').trim();
}
function limpio_(s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); }

var DEPTOS_ = ['Amazonas', 'Áncash', 'Apurímac', 'Arequipa', 'Ayacucho', 'Cajamarca', 'Callao', 'Cusco',
  'Huancavelica', 'Huánuco', 'Ica', 'Junín', 'La Libertad', 'Lambayeque', 'Lima', 'Loreto', 'Madre de Dios',
  'Moquegua', 'Pasco', 'Piura', 'Puno', 'San Martín', 'Tacna', 'Tumbes', 'Ucayali'];
var DEPTO_IDX_ = {};
DEPTOS_.forEach(function (d) { DEPTO_IDX_[norm_(d)] = d; });
DEPTO_IDX_['CUZCO'] = 'Cusco';
function depto_(s) {
  var k = norm_(s);
  if (!k) return '';
  if (DEPTO_IDX_[k]) return DEPTO_IDX_[k];
  return limpio_(s).toLowerCase().replace(/(^|\s)\S/g, function (m) { return m.toUpperCase(); })
    .replace(/ (De|Del|La|Y) /g, function (m) { return m.toLowerCase(); });
}

// Busca la fila (o grupo de hasta 3 filas) de encabezados que mejor calza con los campos.
// defs: [{ key, pats: [RegExp], req: true|false }] en orden de prioridad.
function buscarEncabezado_(valores, defs, maxFilas) {
  var mejor = null;
  var lim = Math.min(maxFilas || 25, valores.length);
  for (var r = 0; r < lim; r++) {
    for (var span = 1; span <= 3 && r + span <= valores.length; span++) {
      var ncol = 0;
      for (var k = 0; k < span; k++) ncol = Math.max(ncol, valores[r + k].length);
      var cab = [];
      for (var c = 0; c < ncol; c++) {
        var partes = [];
        for (var k2 = 0; k2 < span; k2++) { var t = norm_(valores[r + k2][c]); if (t) partes.push(t); }
        cab.push(partes.join(' '));
      }
      var mapa = {}, usadas = {}, req = 0, total = 0;
      defs.forEach(function (d) {
        for (var c2 = 0; c2 < cab.length; c2++) {
          if (usadas[c2] || !cab[c2]) continue;
          if (d.pats.some(function (p) { return p.test(cab[c2]); })) {
            mapa[d.key] = c2; usadas[c2] = true; total++; if (d.req) req++; break;
          }
        }
      });
      var puntaje = req * 100 + total * 10 - span;
      if (!mejor || puntaje > mejor.puntaje) mejor = { fila: r, span: span, mapa: mapa, puntaje: puntaje, req: req };
    }
  }
  return mejor;
}
function faltantes_(enc, defs) {
  return defs.filter(function (d) { return d.req && enc.mapa[d.key] === undefined; })
    .map(function (d) { return d.key; });
}
function celda_(fila, enc, key) {
  var c = enc.mapa[key];
  return c === undefined ? '' : limpio_(fila[c]);
}
function esRepeticionEncabezado_(fila, enc, defs) {
  var coincide = 0;
  defs.forEach(function (d) {
    var v = norm_(celda_(fila, enc, d.key));
    if (v && d.pats.some(function (p) { return p.test(v); })) coincide++;
  });
  return coincide >= 2;
}
function contarValores_(lista) {
  var m = {};
  lista.forEach(function (v) { m[v] = (m[v] || 0) + 1; });
  return Object.keys(m).map(function (k) { return k + ' (' + m[k] + ')'; }).join(', ');
}

// ── PVN: inventario de la Red Vial Nacional ────────────────────────────────
var DEF_PVN_ = [
  { key: 'depto', pats: [/DEPARTAMENTO/, /REGION/], req: true },
  { key: 'nombre', pats: [/NOMBRE/, /^PUENTE$/], req: true },
  { key: 'ruta', pats: [/^RUTA/, /CODIGO DE RUTA/], req: false },
  { key: 'prog', pats: [/PROGRESIVA/, /^KM$/], req: false },
  { key: 'longm', pats: [/LONGITUD ?\(M\)/, /LONG\.? ?\(M\)/, /LONGITUD TOTAL/], req: false },
  { key: 'tipo', pats: [/^TIPO/], req: false },
  { key: 'cond', pats: [/CONDICION/, /^ESTADO/], req: true },
  { key: 'lat', pats: [/LATITUD/], req: false },
  { key: 'lon', pats: [/LONGITUD/], req: false }
];
var CONDICION_ = { 'BUENO': 'Bueno', 'BUENA': 'Bueno', 'BUEN ESTADO': 'Bueno', 'REGULAR': 'Regular', 'MALO': 'Malo', 'MALA': 'Malo', 'CRITICO': 'Crítico' };

function leerPVN(valores) {
  var enc = buscarEncabezado_(valores, DEF_PVN_);
  var falta = faltantes_(enc, DEF_PVN_);
  if (falta.length) return { ok: false, error: 'No se encontraron las columnas: ' + falta.join(', ') };
  var filas = [], avisos = [], desconocidos = [], sinDepto = 0;
  for (var r = enc.fila + enc.span; r < valores.length; r++) {
    var f = valores[r];
    var nombre = celda_(f, enc, 'nombre');
    if (!nombre || esRepeticionEncabezado_(f, enc, DEF_PVN_)) continue;
    var condRaw = celda_(f, enc, 'cond');
    var cond = CONDICION_[norm_(condRaw)] || (condRaw ? limpio_(condRaw) : 'Sin dato');
    if (condRaw && !CONDICION_[norm_(condRaw)]) desconocidos.push(condRaw);
    var dep = depto_(celda_(f, enc, 'depto'));
    if (!dep) sinDepto++;
    filas.push([dep || 'Sin departamento', nombre, celda_(f, enc, 'ruta'), celda_(f, enc, 'prog'),
      celda_(f, enc, 'longm'), celda_(f, enc, 'tipo') || 'Sin dato', cond, celda_(f, enc, 'lat'), celda_(f, enc, 'lon')]);
  }
  if (desconocidos.length) avisos.push('Condiciones no reconocidas, se conservan tal cual: ' + contarValores_(desconocidos));
  if (sinDepto) avisos.push(sinDepto + ' puentes sin departamento');
  return { ok: filas.length > 0, error: filas.length ? '' : 'No se encontraron filas de datos', filas: filas, avisos: avisos,
    cabecera: ['Departamento', 'Puente', 'Ruta', 'Progresiva', 'Longitud (m)', 'Tipo', 'Condición', 'Latitud', 'Longitud'] };
}

// ── DGPPT: inventario de puentes de las concesiones ────────────────────────
var DEF_DGPPT_ = [
  { key: 'num', pats: [/^N\s*[°ºO]?\.?$/, /^NRO/, /^ITEM/, /^N°/], req: false },
  { key: 'nombre', pats: [/NOMBRE/], req: true },
  { key: 'ruta', pats: [/^RUTA/, /UBICACION RUTA/], req: false },
  { key: 'prog', pats: [/PROGRESIVA/], req: false },
  { key: 'distrito', pats: [/DISTRITO/], req: false },
  { key: 'provincia', pats: [/PROVINCIA/], req: false },
  { key: 'depto', pats: [/DEPARTAMENTO/, /REGION/], req: true },
  { key: 'longm', pats: [/LONGITUD ?\(M\)/, /LONG\.? ?\(M\)/], req: false },
  { key: 'estado', pats: [/\bESTADO\b/, /CONDICION/], req: true },
  { key: 'calif', pats: [/CALIFICACION/], req: false },
  { key: 'om', pats: [/O&M/, /OPERACION/], req: false }
];
var CONCESION_CORTO_ = [
  ['AUTOPISTA DEL SOL', 'Autopista del Sol'], ['SIERRA TRAMO 2', 'Longitudinal de la Sierra T2'],
  ['SIERRA TRAMO 4', 'Longitudinal de la Sierra T4'], ['AMAZONAS NORTE', 'IIRSA Norte'],
  ['RED VIAL N° 6', 'Red Vial 6'], ['RED VIAL N 6', 'Red Vial 6'], ['RED VIAL N° 5', 'Red Vial 5'], ['RED VIAL N 5', 'Red Vial 5'],
  ['IIRSA SUR TRAMO 1', 'IIRSA Sur T1'], ['IIRSA SUR TRAMO 2', 'IIRSA Sur T2'], ['IIRSA SUR TRAMO 3', 'IIRSA Sur T3'],
  ['IIRSA SUR TRAMO 4', 'IIRSA Sur T4'], ['IIRSA SUR TRAMO 5', 'IIRSA Sur T5'], ['IIRSA CENTRO', 'IIRSA Centro T2'],
  ['AUTOPISTA DEL NORTE', 'Autopista del Norte'], ['QUILCA', 'Dv. Quilca – La Concordia'],
  ['CANCHAQUE', 'Buenos Aires – Canchaque'], ['HUARAL', 'Chancay – Huaral – Acos'], ['MOCUPE', 'Nuevo Mocupe – Oyotún']
];
var ESTADO_DG_ = { 'BUENO': 'Bueno', 'BUEN ESTADO': 'Bueno', 'BUEN ESTADO (NUEVO)': 'Bueno', 'NUEVO': 'Bueno',
  'REGULAR': 'Regular', 'MALO': 'Malo', 'CRITICO': 'Crítico', 'MUY MALO': 'Crítico' };
var CALIF_ = { 'DEFICIENTE': 'Deficiente', 'DEFICIENTRE': 'Deficiente', 'ALARMANTE': 'Alarmante',
  'RIESGO INACEPTABLE': 'Riesgo inaceptable', 'SERIA': 'Seria', 'SERIO': 'Seria' };

function nombreConcesion_(titulo) {
  var t = norm_(titulo);
  for (var i = 0; i < CONCESION_CORTO_.length; i++) if (t.indexOf(norm_(CONCESION_CORTO_[i][0])) >= 0) return CONCESION_CORTO_[i][1];
  // Concesión nueva: se usa el título, sin "Puentes de la/del" y hasta los dos puntos
  return limpio_(titulo).replace(/^PUENTES?\s+(DE\s+LA|DEL|DE)\s+/i, '').split(':')[0].trim();
}
function esTituloSeccion_(fila) {
  var vals = fila.map(limpio_).filter(function (v) { return v; });
  if (!vals.length || vals.length > 2) return '';
  var t = vals.join(' ');
  if (t.length < 12 || !/[A-Za-zÁÉÍÓÚÑ]{4}/.test(t) || /^\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}$/.test(t)) return '';
  return t;
}

function leerDGPPT(valores) {
  var enc = buscarEncabezado_(valores, DEF_DGPPT_);
  var falta = faltantes_(enc, DEF_DGPPT_);
  if (falta.length) return { ok: false, error: 'No se encontraron las columnas: ' + falta.join(', ') };
  var filas = [], avisos = [], desconocidos = [], calDesc = [], sinDepto = 0, sinConc = 0, omitidas = 0;
  // la concesión del primer bloque está en las filas de título sobre el encabezado
  var conc = '';
  for (var r0 = 0; r0 < enc.fila; r0++) { var t0 = esTituloSeccion_(valores[r0]); if (t0 && /PUENTE|TRAMO|AUTOPISTA|RED VIAL|IIRSA|CARRETERA/i.test(t0)) conc = nombreConcesion_(t0); }
  for (var r = enc.fila + enc.span; r < valores.length; r++) {
    var f = valores[r];
    var titulo = esTituloSeccion_(f);
    var nombre = celda_(f, enc, 'nombre');
    if (titulo && (!nombre || titulo.indexOf(nombre) >= 0) && /PUENTE|TRAMO|AUTOPISTA|RED VIAL|IIRSA|CARRETERA|CONCESI/i.test(titulo)) { conc = nombreConcesion_(titulo); continue; }
    if (!nombre || esRepeticionEncabezado_(f, enc, DEF_DGPPT_)) continue;
    if (titulo && !celda_(f, enc, 'ruta') && !celda_(f, enc, 'estado')) continue;
    // Si el archivo numera sus puentes, las filas sin número son subtítulos de tramo o encabezados repetidos
    if (enc.mapa.num !== undefined && !/^\d+$/.test(celda_(f, enc, 'num'))) { omitidas++; continue; }
    var eRaw = celda_(f, enc, 'estado');
    var est = ESTADO_DG_[norm_(eRaw)] || (eRaw ? limpio_(eRaw) : 'Sin dato');
    if (eRaw && !ESTADO_DG_[norm_(eRaw)]) desconocidos.push(eRaw);
    var cRaw = norm_(celda_(f, enc, 'calif')).replace(/[^A-Z ]/g, '').trim();
    var cal = CALIF_[cRaw] || (cRaw ? limpio_(celda_(f, enc, 'calif')) : '');
    if (cRaw && !CALIF_[cRaw]) calDesc.push(cRaw);
    var dep = depto_(celda_(f, enc, 'depto'));
    if (!dep) sinDepto++;
    if (!conc) sinConc++;
    filas.push([conc || 'Sin concesión', dep || 'Sin departamento', celda_(f, enc, 'distrito'), nombre,
      celda_(f, enc, 'ruta'), celda_(f, enc, 'prog'), celda_(f, enc, 'longm'), est, cal, celda_(f, enc, 'om')]);
  }
  if (desconocidos.length) avisos.push('Estados no reconocidos, se conservan tal cual: ' + contarValores_(desconocidos));
  if (calDesc.length) avisos.push('Calificaciones no reconocidas: ' + contarValores_(calDesc));
  if (sinDepto) avisos.push(sinDepto + ' puentes sin departamento');
  if (sinConc) avisos.push(sinConc + ' puentes sin concesión identificada');
  if (omitidas) avisos.push(omitidas + ' filas sin número omitidas (subtítulos de tramo o encabezados repetidos)');
  return { ok: filas.length > 0, error: filas.length ? '' : 'No se encontraron filas de datos', filas: filas, avisos: avisos,
    cabecera: ['Concesión', 'Departamento', 'Distrito', 'Puente', 'Ruta', 'Progresiva', 'Longitud (m)', 'Estado', 'Calificación', 'Responsable O&M'] };
}

// ── PVD: cartera de puentes de Provías Descentralizado ─────────────────────
var DEF_PVD_ = [
  { key: 'cui', pats: [/^CUI$/, /CODIGO UNICO/], req: false },
  { key: 'region', pats: [/^REGION/, /DEPARTAMENTO/], req: true },
  { key: 'provincia', pats: [/PROVINCIA/], req: false },
  { key: 'distrito', pats: [/DISTRITO/], req: false },
  { key: 'nombre', pats: [/DENOMINACION/, /NOMBRE/], req: true },
  { key: 'estado', pats: [/^ESTADO/, /SITUACION/], req: true },
  { key: 'tinf', pats: [/TIPO DE INFRAESTRUCTURA/, /TIPO DE PUENTE/], req: false },
  { key: 'tvia', pats: [/TIPO DE VIA/, /^RED/], req: false },
  { key: 'tinv', pats: [/TIPO DE INVERSION/], req: false },
  { key: 'grupo', pats: [/^GRUPO/], req: false }
];
function situacionPVD_(lineaSit) {
  var s = norm_(lineaSit);
  if (/CULMINAD/.test(s)) return 'Culminado';
  if (/SIN INICIAR/.test(s)) return 'Sin iniciar';
  if (/PARALIZAD/.test(s)) return 'Paralizado';
  if (/SUSPENDID/.test(s)) return 'Suspendido';
  if (/EJECUCION FISICA/.test(s)) return 'En ejecución física';
  if (/ELABORACION/.test(s)) return 'En estudios';
  if (/ACTOS|CONVOCAD|CONTRATO|PREPARATORI|BUENA PRO/.test(s)) return 'En contratación';
  return '';
}
function etapaPVD_(lineaEtapa) {
  var s = norm_(lineaEtapa);
  if (/^C\.|EJECUCION/.test(s)) return 'Ejecución';
  if (/^B\.|ESTUDIO/.test(s)) return 'Estudio definitivo';
  if (/^A\.|PREINVERSION/.test(s)) return 'Preinversión';
  return '';
}
function leerPVD(valores) {
  var enc = buscarEncabezado_(valores, DEF_PVD_);
  var falta = faltantes_(enc, DEF_PVD_);
  if (falta.length) return { ok: false, error: 'No se encontraron las columnas: ' + falta.join(', ') };
  var filas = [], avisos = [], desconocidos = [], sinRegion = 0;
  for (var r = enc.fila + enc.span; r < valores.length; r++) {
    var f = valores[r];
    var nombre = celda_(f, enc, 'nombre');
    if (!nombre || esRepeticionEncabezado_(f, enc, DEF_PVD_)) continue;
    var lineas = String(f[enc.mapa.estado] || '').split(/\r?\n|\|/).map(limpio_).filter(function (x) { return x; });
    var sitLinea = lineas.length ? lineas[lineas.length - 1] : '';
    var sit = situacionPVD_(sitLinea);
    if (!sit) { sit = sitLinea || 'Sin dato'; if (sitLinea) desconocidos.push(sitLinea); }
    var etapa = lineas.length > 1 ? etapaPVD_(lineas[0]) : '';
    var tRaw = norm_(celda_(f, enc, 'tinf'));
    var tipo = /MODULAR/.test(tRaw) ? 'Modular' : (/DEFINITIV/.test(tRaw) ? 'Definitivo' : (tRaw ? limpio_(celda_(f, enc, 'tinf')) : 'Sin dato'));
    var reg = depto_(celda_(f, enc, 'region'));
    if (!reg) sinRegion++;
    var quitarNum = function (s) { return limpio_(s).replace(/^\d+\.\s*/, ''); };
    filas.push([reg || 'Sin región', celda_(f, enc, 'provincia'), celda_(f, enc, 'distrito'), nombre, tipo,
      quitarNum(celda_(f, enc, 'tvia')), quitarNum(celda_(f, enc, 'tinv')), celda_(f, enc, 'grupo'),
      celda_(f, enc, 'cui'), etapa || '—', sit]);
  }
  if (desconocidos.length) avisos.push('Situaciones no reconocidas, se conservan tal cual: ' + contarValores_(desconocidos));
  if (sinRegion) avisos.push(sinRegion + ' puentes sin región');
  return { ok: filas.length > 0, error: filas.length ? '' : 'No se encontraron filas de datos', filas: filas, avisos: avisos,
    cabecera: ['Región', 'Provincia', 'Distrito', 'Puente', 'Tipo', 'Tipo de vía', 'Tipo de inversión', 'Grupo', 'CUI', 'Etapa', 'Situación'] };
}

