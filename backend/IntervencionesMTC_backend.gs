/**
 * ═══════════════════════════════════════════════════════════════════════════
 * INTERVENCIONES MTC — REGISTRO DEL CORTE (módulo beta, paralelo al Hub)
 * Backend Apps Script del Sheet «Base de Intervenciones MTC»
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * Un solo universo de intervenciones (pestaña Universo) como fuente de datos.
 * Las unidades reportantes registran el estado situacional del corte desde el
 * Hub (página beta/registro-intervenciones.html); el especialista OMEI valida,
 * o corrige y valida; el administrador carga la información financiera y
 * publica el corte, que genera la pestaña Actual con las mismas columnas que la
 * base de datos de la Carpeta Departamental (criterio único del monto).
 *
 * FLUJO (directiva v2, en el Hub):
 *   1. Administrador / Dirección abre el corte (fecha de corte y plazo).
 *   2. Punto focal: actualiza el reporte de cada intervención de su unidad,
 *      propone altas, bajas y correcciones, y envía.
 *   3. Especialista: valida, o corrige y valida con sustento (sin devolver);
 *      devolver a la unidad es la excepción. Aprueba altas (asigna COD INTERNO).
 *   4. Administrador: carga la consulta masiva del MEF (pestaña «base»).
 *   5. Administrador: publica (100 % validado, sin solicitudes pendientes):
 *      respalda Actual como «dd-mm», escribe Actual, CUI vigentes y KPIs.
 *
 * SEGURIDAD: la pestaña Accesos NUNCA se publica en la web. Las contraseñas se
 * verifican aquí; la página solo recibe un token de sesión (6 horas).
 *
 * PUBLICAR: Implementar → Nueva implementación → Aplicación web
 *   (Ejecutar como: yo · Quién tiene acceso: cualquier usuario).
 *   Cambios posteriores: Gestionar implementaciones → lápiz → Nueva versión.
 *
 * Este archivo también corre en el navegador en el modo demo de la página
 * (beta/demo/), sobre un simulador de Sheets. No usar APIs fuera de:
 * SpreadsheetApp, CacheService, LockService, Utilities, ContentService.
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────────────────
var CONFIG_INT = {
  SPREADSHEET_ID: '',          // vacío = el Sheet donde está este script
  ZONA: 'America/Lima',
  HORAS_SESION: 6,
  INTENTOS_LOGIN: 5            // intentos fallidos antes de bloquear 10 minutos
};

var H = {
  UNIVERSO: 'Universo', REPORTES: 'Reportes', ULTIMO: 'Último reporte', HIST: 'Histórico de reportes',
  SOLICITUDES: 'Solicitudes', CORTES: 'Cortes', FINANCIERA: 'Financiera', ASIGNACIONES: 'Asignaciones',
  ACCESOS: 'Accesos', BITACORA: 'Bitácora', ACTUAL: 'Actual', CUIS: 'CUI vigentes', KPIS: 'KPIs',
  CAT_UNIDADES: 'Cat_Unidades', CAT_TIPOS: 'Cat_Tipos', CAT_MF: 'Cat_ModalidadFase', CAT_LISTAS: 'Cat_Listas',
  CARGA_UNIVERSO: 'Carga_Universo', CARGA_ACTUAL: 'Carga_Actual', CARGA_MAESTRO: 'Carga_Maestro',
  CARGA_FINANCIERA: 'Carga_Financiera', CARGA_CREDITO: 'Carga_Credito'
};

var COLS_REPORTE = ['ID', 'Corte', 'COD INTERNO', 'ID solicitud', 'Unidad reportante', 'Monto reportado S/',
  'Beneficiarios', 'Modalidad de ejecución', 'Fase', 'Financiamiento (SI/NO)', 'Estado situacional',
  'Cambios de identificación', 'Motivo de los cambios', 'Estado', 'Observación', 'Registrado por',
  'Fecha de registro', 'Revisado por', 'Fecha de revisión', 'Correcciones OMEI', 'Sustento OMEI'];

var ENC = {};
ENC[H.UNIVERSO] = ['COD INTERNO', 'Tipo', 'Vigente', 'Viceministerio', 'Unidad reportante', 'Nombre de la intervención',
  'Área de influencia', 'Tipo de Código', 'Código', 'Tipo de Infraestructura', 'Tipo de Órgano ejecutor',
  'Crédito suplementario', 'Monto del crédito S/', 'Hoja de Ruta', 'Cartera Priorizada', 'Relevancia sectorial',
  'Incluida en', 'Fecha de alta', 'Registrado por', 'Última modificación', 'Modificado por'];
ENC[H.REPORTES] = COLS_REPORTE;
ENC[H.ULTIMO] = COLS_REPORTE;
ENC[H.HIST] = COLS_REPORTE;
ENC[H.SOLICITUDES] = ['ID', 'Tipo', 'Corte', 'COD INTERNO', 'Unidad reportante', 'Datos', 'Motivo', 'Estado',
  'Solicitado por', 'Fecha de solicitud', 'Resuelto por', 'Fecha de resolución', 'COD asignado', 'Sustento OMEI'];
ENC[H.CORTES] = ['Corte', 'Plazo', 'Estado', 'Abierto por', 'Fecha de apertura', 'Financiera cargada', 'Fecha financiera',
  'Publicado por', 'Fecha de publicación', 'Intervenciones', 'Monto total S/'];
ENC[H.FINANCIERA] = ['CUI', 'Costo de Inversión', 'PIA', 'PIM', 'Devengado Anual', 'Fuente del costo', 'Fecha de consulta'];
ENC[H.ASIGNACIONES] = ['Ámbito', 'Clave', 'Especialista', 'Desde', 'Hasta', 'Asignado por'];
ENC[H.ACCESOS] = ['Usuario', 'Contraseña', 'Rol', 'Unidad reportante', 'Nombre', 'Activo'];
ENC[H.BITACORA] = ['Fecha', 'Usuario', 'Acción', 'Detalle'];
ENC[H.CUIS] = ['COD INTERNO', 'CUI'];
ENC[H.KPIS] = ['Indicador', 'Valor', 'Unidad', 'Corte', 'Calculado el'];
ENC[H.CAT_UNIDADES] = ['Sigla', 'Nombre', 'Viceministerio'];
ENC[H.CAT_TIPOS] = ['Tipo', 'Tipos de código', 'Datos de intervención', 'Módulo Intervenciones MTC'];
ENC[H.CAT_MF] = ['Modalidad', 'N°', 'Fase', 'Etapa'];
ENC[H.CAT_LISTAS] = ['Lista', 'Valor'];

var ROLES = { PF: 'Punto focal', ESP: 'Especialista', ADMIN: 'Administrador', DIR: 'Dirección' };
var AMBITO_CARPETA = 'Carpeta Departamental', AMBITO_HR = 'Hoja de Ruta';
var TIPOS_CODIGO = ['CUI', 'IDEA', 'PENDIENTE', 'No aplica'];
// Ningún dato de identificación (código, CUI, crédito suplementario…) se cambia desde el reporte del corte:
// todos se cambian por «Solicitar corrección», que aprueba OMEI. Así no hay dos vías que se crucen.
var CAMPOS_CAMBIO_DIRECTO = [];
// Datos de identificación que se cambian por solicitud de corrección
var CAMPOS_IDENTIDAD = ['Tipo', 'Nombre de la intervención', 'Área de influencia', 'Tipo de Código', 'Código',
  'Tipo de Infraestructura', 'Tipo de Órgano ejecutor', 'Crédito suplementario', 'Monto del crédito S/',
  'Hoja de Ruta', 'Cartera Priorizada', 'Relevancia sectorial', 'Incluida en'];
var CAMPOS_REPORTE = ['Monto reportado S/', 'Beneficiarios', 'Modalidad de ejecución', 'Fase',
  'Financiamiento (SI/NO)', 'Estado situacional'];
var DEPARTAMENTOS = ['AMAZONAS', 'ÁNCASH', 'APURÍMAC', 'AREQUIPA', 'AYACUCHO', 'CAJAMARCA', 'CALLAO', 'CUSCO',
  'HUANCAVELICA', 'HUÁNUCO', 'ICA', 'JUNÍN', 'LA LIBERTAD', 'LAMBAYEQUE', 'LIMA', 'LORETO', 'MADRE DE DIOS',
  'MOQUEGUA', 'PASCO', 'PIURA', 'PUNO', 'SAN MARTÍN', 'TACNA', 'TUMBES', 'UCAYALI'];
var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'setiembre', 'octubre', 'noviembre', 'diciembre'];


// ── ENTRADA WEB ────────────────────────────────────────────────────────────
function doPost(e) {
  var datos;
  try { datos = JSON.parse(e.postData.contents); }
  catch (err) { return responder({ ok: false, mensaje: 'Solicitud no válida' }); }
  return responder(procesar(datos));
}

function doGet() {
  return responder({ ok: true, mensaje: 'Web App Intervenciones MTC (beta) activo' });
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Punto único de entrada (también lo usa el modo demo)
function procesar(d) {
  try {
    if (!d || !d.accion) return { ok: false, mensaje: 'Falta la acción' };
    if (d.accion === 'login') return login(d.usuario, d.clave);
    var u = sesion(d.token);
    if (!u) return { ok: false, sesion: false, mensaje: 'La sesión expiró. Vuelve a ingresar.' };
    var acciones = {
      datos: accDatos, guardarReportes: accGuardarReportes, registrarAlta: accRegistrarAlta,
      solicitar: accSolicitar, revisarReportes: accRevisarReportes, resolverSolicitud: accResolverSolicitud,
      abrirCorte: accAbrirCorte, cargarFinanciera: accCargarFinanciera, publicarCorte: accPublicarCorte,
      guardarAsignaciones: accGuardarAsignaciones, editarUniverso: accEditarUniverso, salir: accSalir
    };
    var fn = acciones[d.accion];
    if (!fn) return { ok: false, mensaje: 'Acción no reconocida: ' + d.accion };
    if (d.accion === 'datos' || d.accion === 'salir') return fn(u, d);
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { return fn(u, d); } finally { lock.releaseLock(); }
  } catch (err) {
    console.error(err && err.stack || err);
    return { ok: false, mensaje: 'Error: ' + (err && err.message || err) };
  }
}


// ── UTILIDADES ─────────────────────────────────────────────────────────────
function libro() {
  return CONFIG_INT.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG_INT.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function hoja(nombre, crear) {
  var ss = libro(); var h = ss.getSheetByName(nombre);
  if (!h && crear) h = ss.insertSheet(nombre);
  if (!h) throw new Error('No existe la pestaña «' + nombre + '». Ejecuta «Preparar base».');
  return h;
}
function txt(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function nombreNormal(v) { return txt(v).replace(/\s+/g, ' ').toUpperCase(); }
function clave(v) { return txt(v).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' ').toUpperCase(); }
function esSi(v) { return clave(v) === 'SI'; }
function ahora() { return Utilities.formatDate(new Date(), CONFIG_INT.ZONA, 'dd/MM/yyyy HH:mm'); }
function hoyTxt() { return Utilities.formatDate(new Date(), CONFIG_INT.ZONA, 'dd/MM/yyyy'); }
function uuid() { return Utilities.getUuid().slice(0, 8).toUpperCase(); }

// "1,250,000.50" | "1 250 000" | 1250000 → número; '' → NaN
function num(v) {
  if (typeof v === 'number') return v;
  var t = txt(v).replace(/[S\/\s]/g, '').replace(/,/g, '');
  if (t === '' || t === '-') return NaN;
  var n = Number(t); return isNaN(n) ? NaN : n;
}
// "2167707 , 2167708; 2167715" → ['2167707','2167708','2167715']
function partirCodigos(v) {
  return txt(v).split(/[,;\/\n]+/).map(function (p) { return p.trim().replace(/\.0+$/, ''); })
    .filter(function (p) { return p && clave(p) !== 'NO APLICA' && clave(p) !== 'PENDIENTE'; });
}
// 1234567.8 → "1,234,568" (o con decimales)
function miles(n, dec) {
  var p = Number(n || 0).toFixed(dec || 0).split('.');
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return p.join('.');
}
function unirCodigos(lista) { return lista.join(', '); }
function normTipoCodigo(v) {
  var k = clave(v);
  if (k === 'CUI') return 'CUI'; if (k === 'IDEA') return 'IDEA'; if (k === 'PENDIENTE') return 'PENDIENTE';
  if (k === 'NO APLICA' || k === '') return 'No aplica';
  return txt(v);
}
// dd/mm/yyyy → número ordenable
function ordenFecha(s) {
  var m = txt(s).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? Number(m[3]) * 10000 + Number(m[2]) * 100 + Number(m[1]) : 0;
}
function textoCorte(corte) {
  var m = txt(corte).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? 'Corte del ' + Number(m[1]) + ' de ' + MESES[Number(m[2]) - 1] + ' de ' + m[3] : 'Corte del ' + corte;
}
function etiquetaCorte(corte) { var m = txt(corte).match(/^(\d{1,2})\/(\d{1,2})/); return m ? ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2) : txt(corte); }
function anioCorte(corte) { var m = txt(corte).match(/(\d{4})$/); return m ? m[1] : Utilities.formatDate(new Date(), CONFIG_INT.ZONA, 'yyyy'); }
function jsonSeguro(s, def) { try { return s ? JSON.parse(s) : def; } catch (e) { return def; } }

// Tabla con encabezado en la fila 1. Devuelve objetos {encabezado: valor}.
function leer(nombre) {
  var h = hoja(nombre); var v = h.getDataRange().getValues();
  if (!v.length) return [];
  var enc = v[0].map(txt);
  var filas = [];
  for (var i = 1; i < v.length; i++) {
    if (!v[i].some(function (c) { return txt(c) !== ''; })) continue;
    var o = {}; for (var j = 0; j < enc.length; j++) if (enc[j]) o[enc[j]] = txt(v[i][j]);
    filas.push(o);
  }
  return filas;
}
function aMatriz(nombre, filas) {
  var enc = ENC[nombre];
  return filas.map(function (o) { return enc.map(function (k) { var x = o[k]; return x === undefined || x === null ? '' : String(x); }); });
}
// Reescribe la tabla completa (solo pestañas chicas)
function escribir(nombre, filas) {
  var h = hoja(nombre, true); var enc = ENC[nombre];
  h.clearContents();
  h.getRange(1, 1, 1, enc.length).setValues([enc]);
  if (filas.length) {
    var r = h.getRange(2, 1, filas.length, enc.length);
    r.setNumberFormat('@');
    r.setValues(aMatriz(nombre, filas));
  }
}
function agregar(nombre, filas) {
  if (!filas.length) return;
  var h = hoja(nombre, true); var enc = ENC[nombre];
  if (h.getLastRow() === 0) h.getRange(1, 1, 1, enc.length).setValues([enc]);
  var r = h.getRange(h.getLastRow() + 1, 1, filas.length, enc.length);
  r.setNumberFormat('@');
  r.setValues(aMatriz(nombre, filas));
}
function anotar(usuario, accion, detalle) {
  agregar(H.BITACORA, [{ 'Fecha': ahora(), 'Usuario': usuario, 'Acción': accion, 'Detalle': detalle }]);
}


// ── CATÁLOGOS Y REGLAS ─────────────────────────────────────────────────────
function leerCatalogos() {
  var cat = { unidades: {}, tipos: {}, modalidades: {}, ordenModalidades: [], listas: {} };
  leer(H.CAT_UNIDADES).forEach(function (r) {
    if (r['Sigla']) cat.unidades[r['Sigla']] = { sigla: r['Sigla'], nombre: r['Nombre'] || r['Sigla'], vice: r['Viceministerio'] };
  });
  leer(H.CAT_TIPOS).forEach(function (r) {
    if (!r['Tipo']) return;
    cat.tipos[r['Tipo']] = {
      tipo: r['Tipo'],
      codigos: txt(r['Tipos de código']).split(';').map(normTipoCodigo).filter(Boolean),
      datos: esSi(r['Datos de intervención']),
      modulo: esSi(r['Módulo Intervenciones MTC'])
    };
  });
  leer(H.CAT_MF).sort(function (a, b) { return Number(a['N°']) - Number(b['N°']); }).forEach(function (r) {
    var m = r['Modalidad']; if (!m) return;
    if (!cat.modalidades[m]) { cat.modalidades[m] = []; cat.ordenModalidades.push(m); }
    cat.modalidades[m].push({ fase: r['Fase'], etapa: r['Etapa'] });
  });
  leer(H.CAT_LISTAS).forEach(function (r) {
    if (!r['Lista'] || !r['Valor']) return;
    (cat.listas[r['Lista']] = cat.listas[r['Lista']] || []).push(r['Valor']);
  });
  if (!cat.listas['Departamento']) cat.listas['Departamento'] = DEPARTAMENTOS.slice();
  return cat;
}
function enLista(cat, lista, valor) {
  var k = clave(valor);
  return (cat.listas[lista] || []).some(function (v) { return clave(v) === k; });
}
function valorCatalogo(cat, lista, valor) {
  var k = clave(valor); var l = cat.listas[lista] || [];
  for (var i = 0; i < l.length; i++) if (clave(l[i]) === k) return l[i];
  return txt(valor);
}
function tipoDe(cat, tipo) { return cat.tipos[tipo] || null; }
function enModulo(cat, fila) {
  var t = tipoDe(cat, fila['Tipo']);
  return !!(t && t.modulo && esSi(fila['Vigente']));
}

// Reglas de la identificación (universo). Devuelve una lista de errores.
function validarIdentidad(d, cat) {
  var e = [];
  var t = tipoDe(cat, d['Tipo']);
  if (!t) { e.push('Tipo no válido'); return e; }
  if (!cat.unidades[d['Unidad reportante']]) e.push('Unidad reportante no válida');
  if (txt(d['Nombre de la intervención']).length < 10) e.push('Nombre de la intervención: mínimo 10 caracteres');
  var areas = txt(d['Área de influencia']).split(',').map(txt).filter(Boolean);
  if (!areas.length) e.push('Área de influencia obligatoria');
  else if (areas.some(function (a) { return clave(a) === 'NACIONAL'; })) { if (areas.length > 1) e.push('Área: «NACIONAL» va sola'); }
  else areas.forEach(function (a) { if (!enLista(cat, 'Departamento', a)) e.push('Departamento no válido: ' + a); });
  var tc = normTipoCodigo(d['Tipo de Código']);
  if (t.codigos.indexOf(tc) < 0) e.push('Tipo de Código «' + tc + '» no corresponde al tipo ' + t.tipo);
  var cods = partirCodigos(d['Código']);
  if (tc === 'CUI') {
    if (!cods.length) e.push('Ingresa al menos un CUI');
    cods.forEach(function (c) { if (!/^\d{7}$/.test(c)) e.push('CUI no válido (7 dígitos): ' + c); });
    if (cods.length !== new Set(cods).size) e.push('Hay CUI repetidos');
  } else if (tc === 'IDEA') {
    if (!cods.length) e.push('Ingresa al menos un código IDEA');
    cods.forEach(function (c) { if (!/^\d{5,7}$/.test(c)) e.push('Código IDEA no válido: ' + c); });
  }
  if (t.datos) {
    if (!enLista(cat, 'Tipo de Infraestructura', d['Tipo de Infraestructura'])) e.push('Tipo de Infraestructura no válido');
    if (!enLista(cat, 'Tipo de Órgano ejecutor', d['Tipo de Órgano ejecutor'])) e.push('Tipo de Órgano ejecutor no válido');
  }
  if (t.datos) {   // código, monto y crédito solo para los tipos con datos de intervención (Proyecto, Servicio)
    if (['SI', 'NO'].indexOf(clave(d['Crédito suplementario'])) < 0) e.push('Indica si tiene crédito suplementario');
    if (esSi(d['Crédito suplementario']) && !(num(d['Monto del crédito S/']) > 0)) e.push('Monto del crédito obligatorio (mayor que 0)');
  }
  if (['SI', 'NO'].indexOf(clave(d['Hoja de Ruta'])) < 0) e.push('Indica si forma parte de la Hoja de Ruta');
  if (esSi(d['Hoja de Ruta'])) {
    if (['SI', 'NO'].indexOf(clave(d['Cartera Priorizada'])) < 0) e.push('Indica si es Cartera Priorizada');
    if (txt(d['Relevancia sectorial']).length < 20) e.push('Relevancia sectorial: describe por qué es relevante (mínimo 20 caracteres)');
    txt(d['Incluida en']).split(';').map(txt).filter(Boolean).forEach(function (v) {
      if (!enLista(cat, 'Incluida en', v)) e.push('«Incluida en» no válido: ' + v);
    });
  }
  return e;
}
// Limpia los campos que no aplican y normaliza valores al catálogo
function normalizarIdentidad(d, cat) {
  var o = {}; CAMPOS_IDENTIDAD.forEach(function (k) { o[k] = txt(d[k]); });
  o['Unidad reportante'] = txt(d['Unidad reportante']);
  o['Nombre de la intervención'] = nombreNormal(o['Nombre de la intervención']);   // siempre en mayúsculas y sin saltos ni espacios dobles
  var t = tipoDe(cat, o['Tipo']);
  o['Tipo de Código'] = normTipoCodigo(o['Tipo de Código']);
  o['Código'] = (o['Tipo de Código'] === 'CUI' || o['Tipo de Código'] === 'IDEA') ? unirCodigos(partirCodigos(o['Código'])) : '';
  var areas = o['Área de influencia'].split(',').map(txt).filter(Boolean);
  o['Área de influencia'] = areas.some(function (a) { return clave(a) === 'NACIONAL'; }) ? 'NACIONAL'
    : areas.map(function (a) { return valorCatalogo(cat, 'Departamento', a); }).join(', ');
  if (t && !t.datos) { o['Tipo de Infraestructura'] = ''; o['Tipo de Órgano ejecutor'] = ''; }
  else {
    o['Tipo de Infraestructura'] = valorCatalogo(cat, 'Tipo de Infraestructura', o['Tipo de Infraestructura']);
    o['Tipo de Órgano ejecutor'] = valorCatalogo(cat, 'Tipo de Órgano ejecutor', o['Tipo de Órgano ejecutor']);
  }
  if (t && !t.datos) { o['Crédito suplementario'] = ''; o['Monto del crédito S/'] = ''; }
  else {
    o['Crédito suplementario'] = esSi(o['Crédito suplementario']) ? 'SI' : 'NO';
    o['Monto del crédito S/'] = o['Crédito suplementario'] === 'SI' ? String(num(o['Monto del crédito S/'])) : '';
  }
  o['Hoja de Ruta'] = esSi(o['Hoja de Ruta']) ? 'SI' : 'NO';
  if (o['Hoja de Ruta'] === 'SI') {
    o['Cartera Priorizada'] = esSi(o['Cartera Priorizada']) ? 'SI' : 'NO';
    o['Relevancia sectorial'] = txt(o['Relevancia sectorial']);
    o['Incluida en'] = o['Incluida en'].split(';').map(txt).filter(Boolean)
      .map(function (v) { return valorCatalogo(cat, 'Incluida en', v); }).join('; ');
  } else { o['Cartera Priorizada'] = ''; o['Relevancia sectorial'] = ''; o['Incluida en'] = ''; }
  return o;
}

// Reglas del reporte del corte. tipoCodigo = el que quedará vigente (actual o propuesto).
function validarReporte(r, tipo, tipoCodigo, cat) {
  var e = []; var t = tipoDe(cat, tipo);
  if (txt(r['Estado situacional']).length < 15) e.push('Estado situacional: mínimo 15 caracteres');
  if (!t || !t.datos) return e;
  var fases = cat.modalidades[r['Modalidad de ejecución']];
  if (!fases) e.push('Modalidad de ejecución no válida');
  else if (!fases.some(function (f) { return f.fase === r['Fase']; })) e.push('La fase no corresponde a la modalidad ' + r['Modalidad de ejecución']);
  if (['SI', 'NO'].indexOf(clave(r['Financiamiento (SI/NO)'])) < 0) e.push('Financiamiento: SI o NO');
  var b = num(r['Beneficiarios']); if (isNaN(b) || b < 0) e.push('Beneficiarios: número obligatorio');
  var m = num(r['Monto reportado S/']);
  if (normTipoCodigo(tipoCodigo) !== 'CUI') { if (isNaN(m) || m < 0) e.push('Monto obligatorio (sin CUI es el monto oficial)'); }
  else if (txt(r['Monto reportado S/']) !== '' && (isNaN(m) || m < 0)) e.push('Monto no válido');
  return e;
}
function normalizarReporte(r, tipo, cat) {
  var o = {}; CAMPOS_REPORTE.forEach(function (k) { o[k] = txt(r[k]); });
  var t = tipoDe(cat, tipo);
  if (!t || !t.datos) { ['Monto reportado S/', 'Beneficiarios', 'Modalidad de ejecución', 'Fase', 'Financiamiento (SI/NO)'].forEach(function (k) { o[k] = ''; }); return o; }
  var m = num(o['Monto reportado S/']); o['Monto reportado S/'] = isNaN(m) ? '' : String(m);
  var b = num(o['Beneficiarios']); o['Beneficiarios'] = isNaN(b) ? '' : String(b);
  var fz = clave(o['Financiamiento (SI/NO)']); o['Financiamiento (SI/NO)'] = fz === 'SI' || fz === 'NO' ? fz : '';
  return o;
}

// Cambios directos de identificación propuestos en el reporte {campo: valor}
function aplicarCambios(fila, cambios) {
  var o = {}; for (var k in fila) o[k] = fila[k];
  for (var c in cambios) o[c] = cambios[c];
  return o;
}


// ── SESIÓN ─────────────────────────────────────────────────────────────────
function login(usuario, claveTxt) {
  usuario = txt(usuario);
  var cache = CacheService.getScriptCache();
  var kInt = 'int_' + clave(usuario);
  var intentos = Number(cache.get(kInt) || 0);
  if (intentos >= CONFIG_INT.INTENTOS_LOGIN) return { ok: false, mensaje: 'Demasiados intentos. Espera 10 minutos.' };
  var acc = leer(H.ACCESOS).filter(function (a) { return clave(a['Usuario']) === clave(usuario) && !(clave(a['Activo']) === 'NO'); })[0];
  if (!acc || txt(acc['Contraseña']) !== txt(claveTxt)) {
    cache.put(kInt, String(intentos + 1), 600);
    return { ok: false, mensaje: 'Usuario o contraseña incorrectos' };
  }
  cache.remove(kInt);
  var u = { usuario: acc['Usuario'], rol: acc['Rol'], unidad: acc['Unidad reportante'], nombre: acc['Nombre'] || acc['Usuario'] };
  if ([ROLES.PF, ROLES.ESP, ROLES.ADMIN, ROLES.DIR].indexOf(u.rol) < 0) return { ok: false, mensaje: 'Rol no válido en Accesos: ' + u.rol };
  if (u.rol === ROLES.PF && !u.unidad) return { ok: false, mensaje: 'El usuario no tiene unidad reportante en Accesos' };
  var token = Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify(u), CONFIG_INT.HORAS_SESION * 3600);
  return { ok: true, token: token, usuario: u };
}
function sesion(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var s = cache.get('tok_' + token); if (!s) return null;
  cache.put('tok_' + token, s, CONFIG_INT.HORAS_SESION * 3600);
  return JSON.parse(s);
}
function accSalir(u, d) { CacheService.getScriptCache().remove('tok_' + d.token); return { ok: true }; }
function esGestor(u) { return u.rol === ROLES.ADMIN || u.rol === ROLES.DIR; }
function puedeRevisar(u) { return u.rol === ROLES.ESP || esGestor(u); }

// Unidades que ve el usuario (Carpeta Departamental): el punto focal la suya; el especialista
// las asignadas a él y las que no tienen especialista; administrador y Dirección, todas.
function unidadesVisibles(u, cat, asign) {
  var todas = Object.keys(cat.unidades);
  if (u.rol === ROLES.PF) return [u.unidad];
  if (esGestor(u)) return todas;
  var vig = asign.filter(function (a) { return a['Ámbito'] === AMBITO_CARPETA && !a['Hasta']; });
  return todas.filter(function (s) {
    var a = vig.filter(function (x) { return x['Clave'] === s; });
    return !a.length || a.some(function (x) { return clave(x['Especialista']) === clave(u.usuario); });
  });
}
function corteAbierto() {
  return leer(H.CORTES).filter(function (c) { return c['Estado'] === 'Abierto'; })[0] || null;
}


// ── ACCIÓN: DATOS (carga inicial de la página) ─────────────────────────────
function accDatos(u) {
  var cat = leerCatalogos();
  var asign = leer(H.ASIGNACIONES);
  var visibles = unidadesVisibles(u, cat, asign);
  var vis = {}; visibles.forEach(function (s) { vis[s] = true; });
  var universo = leer(H.UNIVERSO);
  var cortes = leer(H.CORTES);
  var abierto = cortes.filter(function (c) { return c['Estado'] === 'Abierto'; })[0] || null;
  var publicados = cortes.filter(function (c) { return c['Estado'] === 'Publicado'; })
    .sort(function (a, b) { return ordenFecha(b['Corte']) - ordenFecha(a['Corte']); });
  var res = {
    ok: true, usuario: u, catalogos: cat, corte: abierto, ultimoPublicado: publicados[0] || null,
    universo: universo.filter(function (r) { return vis[r['Unidad reportante']]; }),
    reportes: abierto ? leer(H.REPORTES).filter(function (r) { return r['Corte'] === abierto['Corte'] && vis[r['Unidad reportante']]; }) : [],
    ultimos: leer(H.ULTIMO).filter(function (r) { return vis[r['Unidad reportante']]; }),
    solicitudes: leer(H.SOLICITUDES).filter(function (s) { return vis[s['Unidad reportante']]; }),
    // Para avisos de duplicados en todo el universo vigente (nombre y CUI), sin datos sensibles
    indice: universo.filter(function (r) { return esSi(r['Vigente']); }).map(function (r) {
      return [r['COD INTERNO'], r['Unidad reportante'], r['Nombre de la intervención'], r['Tipo de Código'] === 'CUI' ? r['Código'] : ''];
    })
  };
  if (puedeRevisar(u)) {
    res.asignaciones = asign.filter(function (a) { return !a['Hasta']; });
    res.especialistas = leer(H.ACCESOS).filter(function (a) { return a['Rol'] === ROLES.ESP && clave(a['Activo']) !== 'NO'; })
      .map(function (a) { return { usuario: a['Usuario'], nombre: a['Nombre'] || a['Usuario'] }; });
  }
  if (esGestor(u)) {
    var fin = leer(H.FINANCIERA);
    res.financiera = { n: fin.length, fecha: fin.length ? fin[0]['Fecha de consulta'] : '',
      cuis: fin.map(function (f) { return f['CUI']; }) };
    res.cortes = cortes;
    res.hojaDeRuta = universo.filter(function (r) { return esSi(r['Vigente']) && esSi(r['Hoja de Ruta']); })
      .map(function (r) { return { cod: r['COD INTERNO'], nombre: r['Nombre de la intervención'], unidad: r['Unidad reportante'], cartera: r['Cartera Priorizada'] }; });
  }
  return res;
}


// ── ACCIÓN: GUARDAR REPORTES (punto focal; también carga de la matriz) ──────
// d.items = [{cod, campos:{...CAMPOS_REPORTE}, cambios:{campo:valor}, motivo, enviar:bool}]
function accGuardarReportes(u, d) {
  var corte = corteAbierto(); if (!corte) return { ok: false, mensaje: 'No hay un corte abierto' };
  var cat = leerCatalogos();
  var universo = leer(H.UNIVERSO); var porCod = {}; universo.forEach(function (r) { porCod[r['COD INTERNO']] = r; });
  var reportes = leer(H.REPORTES); var idx = {};
  reportes.forEach(function (r, i) { if (r['Corte'] === corte['Corte'] && r['COD INTERNO']) idx[r['COD INTERNO']] = i; });
  var errores = [], guardados = 0;
  (d.items || []).forEach(function (it) {
    var fila = porCod[it.cod];
    if (!fila) { errores.push(it.cod + ': no existe en el universo'); return; }
    if (u.rol === ROLES.PF && fila['Unidad reportante'] !== u.unidad) { errores.push(it.cod + ': no es de tu unidad'); return; }
    if (u.rol !== ROLES.PF) { errores.push('Solo el punto focal registra el reporte; OMEI corrige al validar'); return; }
    if (!enModulo(cat, fila)) { errores.push(it.cod + ': no forma parte del módulo'); return; }
    var previo = idx[it.cod] !== undefined ? reportes[idx[it.cod]] : null;
    if (previo && previo['Estado'] === 'Validado') { errores.push(it.cod + ': ya fue validado'); return; }
    if (Object.keys(it.cambios || {}).some(function (k) { return txt(it.cambios[k]) !== txt(fila[k]); })) {
      errores.push(it.cod + ': el código y el crédito se cambian con «Solicitar corrección»'); return;
    }
    var cambios = {};
    Object.keys(it.cambios || {}).forEach(function (k) {
      if (CAMPOS_CAMBIO_DIRECTO.indexOf(k) >= 0 && txt(it.cambios[k]) !== txt(fila[k])) cambios[k] = txt(it.cambios[k]);
    });
    if (Object.keys(cambios).length && txt(it.motivo).length < 5) { errores.push(it.cod + ': indica el motivo del cambio de código o crédito'); return; }
    var propuesta = normalizarIdentidad(aplicarCambios(fila, cambios), cat);
    var rep = normalizarReporte(it.campos || {}, fila['Tipo'], cat);
    var err = Object.keys(cambios).length ? validarIdentidad(aplicarCambios(propuesta, { 'Unidad reportante': fila['Unidad reportante'] }), cat) : [];
    if (it.enviar) err = err.concat(validarReporte(rep, fila['Tipo'], propuesta['Tipo de Código'], cat));
    if (err.length) { errores.push(it.cod + ': ' + err.join('; ')); return; }
    var cambiosNorm = {}; Object.keys(cambios).forEach(function (k) { cambiosNorm[k] = propuesta[k]; });
    var o = previo || { 'ID': uuid(), 'Corte': corte['Corte'], 'COD INTERNO': it.cod, 'Unidad reportante': fila['Unidad reportante'] };
    CAMPOS_REPORTE.forEach(function (k) { o[k] = rep[k]; });
    o['Cambios de identificación'] = Object.keys(cambiosNorm).length ? JSON.stringify(cambiosNorm) : '';
    o['Motivo de los cambios'] = Object.keys(cambiosNorm).length ? txt(it.motivo) : '';
    o['Estado'] = it.enviar ? 'Enviado' : 'Borrador';
    o['Registrado por'] = u.usuario; o['Fecha de registro'] = ahora();
    if (previo) reportes[idx[it.cod]] = o; else { idx[it.cod] = reportes.length; reportes.push(o); }
    guardados++;
  });
  if (guardados) { escribir(H.REPORTES, reportes); anotar(u.usuario, 'Guardar reportes', guardados + ' reportes del corte ' + corte['Corte']); }
  return { ok: errores.length === 0, guardados: guardados, errores: errores,
    mensaje: errores.length ? 'Se guardaron ' + guardados + '. Revisa: ' + errores.slice(0, 5).join(' | ') + (errores.length > 5 ? ' …' : '') : 'Se guardaron ' + guardados };
}


// ── ACCIÓN: ALTA DE INTERVENCIÓN (punto focal) ─────────────────────────────
// d.identidad = {CAMPOS_IDENTIDAD}, d.reporte = {CAMPOS_REPORTE}
function accRegistrarAlta(u, d) {
  if (u.rol !== ROLES.PF) return { ok: false, mensaje: 'La unidad reportante registra las altas; OMEI las valida' };
  var cat = leerCatalogos(); var corte = corteAbierto();
  var ident = normalizarIdentidad(aplicarCambios(d.identidad || {}, { 'Unidad reportante': u.unidad }), cat);
  ident['Unidad reportante'] = u.unidad;
  var rep = normalizarReporte(d.reporte || {}, ident['Tipo'], cat);
  var err = validarIdentidad(ident, cat).concat(validarReporte(rep, ident['Tipo'], ident['Tipo de Código'], cat));
  if (err.length) return { ok: false, mensaje: err.join('; '), errores: err };
  var datos = { identidad: ident, reporte: rep };
  var s = { 'ID': 'S-' + uuid(), 'Tipo': 'Alta', 'Corte': corte ? corte['Corte'] : '', 'COD INTERNO': '',
    'Unidad reportante': u.unidad, 'Datos': JSON.stringify(datos), 'Motivo': txt(d.motivo), 'Estado': 'Pendiente',
    'Solicitado por': u.usuario, 'Fecha de solicitud': ahora() };
  agregar(H.SOLICITUDES, [s]);
  anotar(u.usuario, 'Alta solicitada', s['ID'] + ' · ' + ident['Nombre de la intervención']);
  return { ok: true, id: s['ID'], mensaje: 'Alta registrada y enviada a OMEI' };
}


// ── ACCIÓN: SOLICITAR BAJA O CORRECCIÓN (punto focal) ──────────────────────
function accSolicitar(u, d) {
  if (u.rol !== ROLES.PF) return { ok: false, mensaje: 'Solo el punto focal registra solicitudes' };
  if (['Baja', 'Corrección'].indexOf(d.tipo) < 0) return { ok: false, mensaje: 'Tipo de solicitud no válido' };
  var fila = leer(H.UNIVERSO).filter(function (r) { return r['COD INTERNO'] === d.cod; })[0];
  if (!fila || fila['Unidad reportante'] !== u.unidad) return { ok: false, mensaje: 'La intervención no es de tu unidad' };
  if (txt(d.motivo).length < 5) return { ok: false, mensaje: 'Indica el motivo' };
  var datos = {};
  if (d.tipo === 'Corrección') {
    var cat = leerCatalogos();
    var cambios = {};
    Object.keys(d.cambios || {}).forEach(function (k) { if (CAMPOS_IDENTIDAD.indexOf(k) >= 0 && txt(d.cambios[k]) !== txt(fila[k])) cambios[k] = d.cambios[k]; });
    if (!Object.keys(cambios).length) return { ok: false, mensaje: 'No hay cambios' };
    var prop = normalizarIdentidad(aplicarCambios(fila, cambios), cat); prop['Unidad reportante'] = fila['Unidad reportante'];
    var err = validarIdentidad(prop, cat); if (err.length) return { ok: false, mensaje: err.join('; ') };
    Object.keys(cambios).forEach(function (k) { datos[k] = prop[k]; });
    ['Cartera Priorizada', 'Relevancia sectorial', 'Incluida en', 'Monto del crédito S/', 'Código'].forEach(function (k) { if (prop[k] !== fila[k]) datos[k] = prop[k]; });
  }
  var pend = leer(H.SOLICITUDES).some(function (s) { return s['COD INTERNO'] === d.cod && s['Estado'] === 'Pendiente' && s['Tipo'] === d.tipo; });
  if (pend) return { ok: false, mensaje: 'Ya hay una solicitud de ' + d.tipo.toLowerCase() + ' pendiente para ' + d.cod };
  var corte = corteAbierto();
  var s = { 'ID': 'S-' + uuid(), 'Tipo': d.tipo, 'Corte': corte ? corte['Corte'] : '', 'COD INTERNO': d.cod,
    'Unidad reportante': u.unidad, 'Datos': JSON.stringify(datos), 'Motivo': txt(d.motivo), 'Estado': 'Pendiente',
    'Solicitado por': u.usuario, 'Fecha de solicitud': ahora() };
  agregar(H.SOLICITUDES, [s]);
  anotar(u.usuario, d.tipo + ' solicitada', s['ID'] + ' · ' + d.cod);
  return { ok: true, id: s['ID'], mensaje: 'Solicitud enviada a OMEI' };
}


// ── ACCIÓN: REVISAR REPORTES (especialista) ────────────────────────────────
// d.items = [{cod, accion:'validar'|'devolver', campos:{...} (finales), cambios:{...} (finales), sustento, observacion}]
// Si el especialista modifica algo frente a lo que envió la unidad, se guarda como corrección y el sustento es obligatorio.
function accRevisarReportes(u, d) {
  if (!puedeRevisar(u)) return { ok: false, mensaje: 'Sin permiso para revisar' };
  var corte = corteAbierto(); if (!corte) return { ok: false, mensaje: 'No hay un corte abierto' };
  var cat = leerCatalogos(); var asign = leer(H.ASIGNACIONES);
  var vis = {}; unidadesVisibles(u, cat, asign).forEach(function (s) { vis[s] = true; });
  var universo = leer(H.UNIVERSO); var porCod = {}; universo.forEach(function (r, i) { porCod[r['COD INTERNO']] = i; });
  var reportes = leer(H.REPORTES); var idx = {};
  reportes.forEach(function (r, i) { if (r['Corte'] === corte['Corte'] && r['COD INTERNO']) idx[r['COD INTERNO']] = i; });
  var errores = [], hechos = 0, univCambio = false;
  (d.items || []).forEach(function (it) {
    var iu = porCod[it.cod]; var fila = iu !== undefined ? universo[iu] : null;
    if (!fila) { errores.push(it.cod + ': no existe'); return; }
    if (!vis[fila['Unidad reportante']]) { errores.push(it.cod + ': no está a tu cargo'); return; }
    var ir = idx[it.cod];
    if (ir === undefined) { errores.push(it.cod + ': la unidad aún no registra su reporte'); return; }
    var rep = reportes[ir];
    if (rep['Estado'] === 'Validado') { errores.push(it.cod + ': ya validado'); return; }
    if (it.accion === 'devolver') {
      if (txt(it.observacion).length < 5) { errores.push(it.cod + ': escribe la observación para la unidad'); return; }
      rep['Estado'] = 'Devuelto'; rep['Observación'] = txt(it.observacion);
      rep['Revisado por'] = u.usuario; rep['Fecha de revisión'] = ahora(); hechos++; return;
    }
    if (it.accion !== 'validar') { errores.push(it.cod + ': acción no válida'); return; }
    if (rep['Estado'] !== 'Enviado') { errores.push(it.cod + ': solo se valida un reporte enviado por la unidad'); return; }
    // Valores de la unidad
    var cambiosUnidad = jsonSeguro(rep['Cambios de identificación'], {});
    var repUnidad = {}; CAMPOS_REPORTE.forEach(function (k) { repUnidad[k] = rep[k]; });
    // Valores finales (OMEI puede corregir)
    var camposFin = it.campos ? normalizarReporte(it.campos, fila['Tipo'], cat) : repUnidad;
    var cambiosFin = it.cambios ? it.cambios : cambiosUnidad;
    var limpios = {}; Object.keys(cambiosFin).forEach(function (k) {
      if (CAMPOS_CAMBIO_DIRECTO.indexOf(k) >= 0 && txt(cambiosFin[k]) !== txt(fila[k])) limpios[k] = cambiosFin[k];
    });
    var ident = normalizarIdentidad(aplicarCambios(fila, limpios), cat); ident['Unidad reportante'] = fila['Unidad reportante'];
    var err = (Object.keys(limpios).length ? validarIdentidad(ident, cat) : [])
      .concat(validarReporte(camposFin, fila['Tipo'], ident['Tipo de Código'], cat));
    if (err.length) { errores.push(it.cod + ': ' + err.join('; ')); return; }
    // Correcciones OMEI = diferencias entre lo enviado por la unidad y lo validado
    var corr = [];
    CAMPOS_REPORTE.forEach(function (k) { if (txt(repUnidad[k]) !== txt(camposFin[k])) corr.push({ campo: k, unidad: repUnidad[k], omei: camposFin[k] }); });
    CAMPOS_CAMBIO_DIRECTO.forEach(function (k) {
      var vu = k in cambiosUnidad ? txt(cambiosUnidad[k]) : txt(fila[k]); var vf = txt(ident[k]);
      if (vu !== vf) corr.push({ campo: k, unidad: vu, omei: vf });
    });
    if (corr.length && txt(it.sustento).length < 5) { errores.push(it.cod + ': indica el sustento de la corrección'); return; }
    CAMPOS_REPORTE.forEach(function (k) { rep[k] = camposFin[k]; });
    rep['Correcciones OMEI'] = corr.length ? JSON.stringify(corr) : '';
    rep['Sustento OMEI'] = corr.length ? txt(it.sustento) : '';
    rep['Estado'] = 'Validado'; rep['Observación'] = '';
    rep['Revisado por'] = u.usuario; rep['Fecha de revisión'] = ahora();
    // Aplica al universo los cambios de identificación validados
    var aplicados = {};
    CAMPOS_CAMBIO_DIRECTO.forEach(function (k) { if (txt(ident[k]) !== txt(fila[k])) aplicados[k] = ident[k]; });
    if (Object.keys(aplicados).length) {
      Object.keys(aplicados).forEach(function (k) { fila[k] = aplicados[k]; });
      fila['Última modificación'] = ahora(); fila['Modificado por'] = u.usuario; univCambio = true;
      anotar(u.usuario, 'Cambio de identificación', it.cod + ' · ' + JSON.stringify(aplicados));
    }
    hechos++;
  });
  if (hechos) escribir(H.REPORTES, reportes);
  if (univCambio) escribir(H.UNIVERSO, universo);
  if (hechos) anotar(u.usuario, 'Revisión', hechos + ' reportes del corte ' + corte['Corte']);
  return { ok: errores.length === 0, hechos: hechos, errores: errores,
    mensaje: errores.length ? 'Listo: ' + hechos + '. Revisa: ' + errores.slice(0, 5).join(' | ') : 'Listo: ' + hechos };
}


// ── ACCIÓN: RESOLVER SOLICITUD (especialista) ──────────────────────────────
// d = {id, decision:'aprobar'|'rechazar', identidad?, reporte?, cambios?, codAsignado?, sustento}
function accResolverSolicitud(u, d) {
  if (!puedeRevisar(u)) return { ok: false, mensaje: 'Sin permiso' };
  var cat = leerCatalogos(); var asign = leer(H.ASIGNACIONES);
  var vis = {}; unidadesVisibles(u, cat, asign).forEach(function (s) { vis[s] = true; });
  var sols = leer(H.SOLICITUDES); var s = null, is = -1;
  sols.forEach(function (x, i) { if (x['ID'] === d.id) { s = x; is = i; } });
  if (!s) return { ok: false, mensaje: 'Solicitud no encontrada' };
  if (s['Estado'] !== 'Pendiente') return { ok: false, mensaje: 'La solicitud ya fue resuelta' };
  if (!vis[s['Unidad reportante']]) return { ok: false, mensaje: 'La unidad no está a tu cargo' };
  if (d.decision === 'rechazar') {
    if (txt(d.sustento).length < 5) return { ok: false, mensaje: 'Indica el motivo del rechazo' };
    s['Estado'] = 'Rechazada'; s['Resuelto por'] = u.usuario; s['Fecha de resolución'] = ahora(); s['Sustento OMEI'] = txt(d.sustento);
    escribir(H.SOLICITUDES, sols); anotar(u.usuario, 'Solicitud rechazada', s['ID']);
    return { ok: true, mensaje: 'Solicitud rechazada' };
  }
  if (d.decision !== 'aprobar') return { ok: false, mensaje: 'Acción no válida' };
  var universo = leer(H.UNIVERSO);
  var datos = jsonSeguro(s['Datos'], {});
  var corte = corteAbierto();
  if (s['Tipo'] === 'Alta') {
    var identU = datos.identidad || {}, repU = datos.reporte || {};
    var ident = normalizarIdentidad(aplicarCambios(d.identidad || identU, { 'Unidad reportante': s['Unidad reportante'] }), cat);
    ident['Unidad reportante'] = s['Unidad reportante'];
    var rep = normalizarReporte(d.reporte || repU, ident['Tipo'], cat);
    var err = validarIdentidad(ident, cat).concat(validarReporte(rep, ident['Tipo'], ident['Tipo de Código'], cat));
    var cod = txt(d.codAsignado) || sugerirCodigo(s['Unidad reportante'], universo, corte ? anioCorte(corte['Corte']) : null);
    if (!/^[A-Z]+-\d+(-\d+)?$/.test(cod)) err.push('COD INTERNO no válido: ' + cod);
    if (universo.some(function (r) { return clave(r['COD INTERNO']) === clave(cod); })) err.push('El COD INTERNO ' + cod + ' ya existe');
    if (err.length) return { ok: false, mensaje: err.join('; ') };
    var corr = diferencias(identU, ident, CAMPOS_IDENTIDAD).concat(diferencias(repU, rep, CAMPOS_REPORTE));
    if (corr.length && txt(d.sustento).length < 5) return { ok: false, mensaje: 'Indica el sustento de la corrección' };
    var nueva = { 'COD INTERNO': cod, 'Vigente': 'SI', 'Viceministerio': (cat.unidades[ident['Unidad reportante']] || {}).vice || '',
      'Fecha de alta': ahora(), 'Registrado por': s['Solicitado por'] + ' / ' + u.usuario };
    CAMPOS_IDENTIDAD.forEach(function (k) { nueva[k] = ident[k]; });
    nueva['Unidad reportante'] = ident['Unidad reportante'];
    universo.push(nueva); escribir(H.UNIVERSO, universo);
    if (corte) {
      var reps = leer(H.REPORTES);
      var o = { 'ID': uuid(), 'Corte': corte['Corte'], 'COD INTERNO': cod, 'ID solicitud': s['ID'], 'Unidad reportante': nueva['Unidad reportante'],
        'Estado': 'Validado', 'Registrado por': s['Solicitado por'], 'Fecha de registro': s['Fecha de solicitud'],
        'Revisado por': u.usuario, 'Fecha de revisión': ahora(),
        'Correcciones OMEI': corr.length ? JSON.stringify(corr) : '', 'Sustento OMEI': corr.length ? txt(d.sustento) : '' };
      CAMPOS_REPORTE.forEach(function (k) { o[k] = rep[k]; });
      reps.push(o); escribir(H.REPORTES, reps);
    }
    s['COD INTERNO'] = cod; s['COD asignado'] = cod;
    s['Sustento OMEI'] = corr.length ? txt(d.sustento) : '';
  } else {
    var iu = -1; universo.forEach(function (r, i) { if (r['COD INTERNO'] === s['COD INTERNO']) iu = i; });
    if (iu < 0) return { ok: false, mensaje: 'La intervención ya no existe' };
    var fila = universo[iu];
    if (s['Tipo'] === 'Baja') {
      fila['Vigente'] = 'NO';
      if (corte) {   // deja de exigirse en el corte en curso
        var rs = leer(H.REPORTES).filter(function (r) { return !(r['Corte'] === corte['Corte'] && r['COD INTERNO'] === fila['COD INTERNO']); });
        escribir(H.REPORTES, rs);
      }
    } else {
      var cambios = d.cambios || datos;
      var prop = normalizarIdentidad(aplicarCambios(fila, cambios), cat); prop['Unidad reportante'] = fila['Unidad reportante'];
      var e2 = validarIdentidad(prop, cat); if (e2.length) return { ok: false, mensaje: e2.join('; ') };
      var c2 = diferencias(datos, prop, Object.keys(datos));
      if (c2.length && txt(d.sustento).length < 5) return { ok: false, mensaje: 'Indica el sustento de la corrección' };
      CAMPOS_IDENTIDAD.forEach(function (k) { fila[k] = prop[k]; });
      s['Sustento OMEI'] = txt(d.sustento);
    }
    fila['Última modificación'] = ahora(); fila['Modificado por'] = u.usuario;
    escribir(H.UNIVERSO, universo);
  }
  s['Estado'] = 'Aprobada'; s['Resuelto por'] = u.usuario; s['Fecha de resolución'] = ahora();
  escribir(H.SOLICITUDES, sols);
  anotar(u.usuario, 'Solicitud aprobada', s['ID'] + ' · ' + s['Tipo'] + ' · ' + s['COD INTERNO']);
  return { ok: true, cod: s['COD INTERNO'], mensaje: s['Tipo'] === 'Alta' ? 'Alta aprobada: ' + s['COD INTERNO'] : 'Solicitud aprobada' };
}
function diferencias(a, b, campos) {
  var out = [];
  campos.forEach(function (k) { if (txt(a[k]) !== txt(b[k])) out.push({ campo: k, unidad: txt(a[k]), omei: txt(b[k]) }); });
  return out;
}
// Siguiente correlativo SIGLA-AAAANNN (p. ej. PVN-2026254)
function sugerirCodigo(sigla, universo, anio) {
  anio = anio || Utilities.formatDate(new Date(), CONFIG_INT.ZONA, 'yyyy');
  var pref = sigla + '-' + anio; var max = 0;
  universo.forEach(function (r) {
    var c = txt(r['COD INTERNO']);
    if (c.indexOf(pref) === 0) { var n = Number(c.slice(pref.length).split('-')[0]); if (!isNaN(n) && n > max) max = n; }
  });
  var sig = String(max + 1); while (sig.length < 3) sig = '0' + sig;
  return pref + sig;
}


// ── ACCIÓN: ABRIR CORTE (Dirección o administrador) ────────────────────────
function accAbrirCorte(u, d) {
  if (!esGestor(u)) return { ok: false, mensaje: 'Sin permiso' };
  if (corteAbierto()) return { ok: false, mensaje: 'Ya hay un corte abierto' };
  var corte = txt(d.corte), plazo = txt(d.plazo);
  if (!ordenFecha(corte) || !ordenFecha(plazo)) return { ok: false, mensaje: 'Fechas en formato dd/mm/aaaa' };
  if (ordenFecha(plazo) < ordenFecha(corte)) return { ok: false, mensaje: 'El plazo no puede ser anterior a la fecha de corte' };
  var cortes = leer(H.CORTES);
  if (cortes.some(function (c) { return c['Corte'] === corte; })) return { ok: false, mensaje: 'Ese corte ya existe' };
  var ult = cortes.filter(function (c) { return c['Estado'] === 'Publicado'; }).sort(function (a, b) { return ordenFecha(b['Corte']) - ordenFecha(a['Corte']); })[0];
  if (ult && ordenFecha(corte) <= ordenFecha(ult['Corte'])) return { ok: false, mensaje: 'La fecha de corte debe ser posterior al último publicado (' + ult['Corte'] + ')' };
  agregar(H.CORTES, [{ 'Corte': corte, 'Plazo': plazo, 'Estado': 'Abierto', 'Abierto por': u.usuario, 'Fecha de apertura': ahora(), 'Financiera cargada': 'NO' }]);
  escribir(H.REPORTES, []);
  anotar(u.usuario, 'Abrir corte', corte + ' · plazo ' + plazo);
  return { ok: true, mensaje: 'Corte del ' + corte + ' abierto' };
}


// ── ACCIÓN: CARGAR FINANCIERA (administrador) ──────────────────────────────
// d.filas = [{CUI, costo, pia, pim, dev, fuente}], d.fechaConsulta
function accCargarFinanciera(u, d) {
  if (u.rol !== ROLES.ADMIN) return { ok: false, mensaje: 'Solo el administrador de la plataforma carga la financiera' };
  var corte = corteAbierto(); if (!corte) return { ok: false, mensaje: 'Abre el corte antes de cargar la financiera' };
  var vistos = {}, filas = [], errores = [];
  (d.filas || []).forEach(function (f, i) {
    var cui = txt(f.CUI).replace(/\.0+$/, '');
    if (!/^\d{7}$/.test(cui)) { errores.push('Fila ' + (i + 2) + ': CUI no válido «' + cui + '»'); return; }
    if (vistos[cui]) return; vistos[cui] = true;
    var n = function (v) { var x = num(v); return isNaN(x) ? 0 : x; };
    filas.push({ 'CUI': cui, 'Costo de Inversión': String(n(f.costo)), 'PIA': String(n(f.pia)), 'PIM': String(n(f.pim)),
      'Devengado Anual': String(n(f.dev)), 'Fuente del costo': txt(f.fuente), 'Fecha de consulta': txt(d.fechaConsulta) || hoyTxt() });
  });
  if (!filas.length) return { ok: false, mensaje: 'El archivo no tiene CUI válidos. ' + errores.slice(0, 3).join(' | ') };
  escribir(H.FINANCIERA, filas);
  var cortes = leer(H.CORTES);
  cortes.forEach(function (c) { if (c['Corte'] === corte['Corte']) { c['Financiera cargada'] = 'SI'; c['Fecha financiera'] = ahora(); } });
  escribir(H.CORTES, cortes);
  anotar(u.usuario, 'Cargar financiera', filas.length + ' CUI · consulta ' + (txt(d.fechaConsulta) || hoyTxt()));
  return { ok: true, n: filas.length, errores: errores, mensaje: 'Financiera cargada: ' + filas.length + ' CUI' + (errores.length ? ' (' + errores.length + ' filas descartadas)' : '') };
}


// ── PUBLICACIÓN ────────────────────────────────────────────────────────────
// Verifica si el corte se puede publicar. Devuelve {listo, faltan:[...], ...}
function estadoPublicacion(corte, cat, universo, reportes, sols) {
  var enCorte = {}; reportes.forEach(function (r) { if (r['Corte'] === corte['Corte']) enCorte[r['COD INTERNO']] = r; });
  var alcance = universo.filter(function (r) { return enModulo(cat, r); });
  var sinValidar = alcance.filter(function (r) { var x = enCorte[r['COD INTERNO']]; return !x || x['Estado'] !== 'Validado'; });
  var pendientes = sols.filter(function (s) { return s['Estado'] === 'Pendiente'; });
  var faltan = [];
  if (sinValidar.length) faltan.push(sinValidar.length + ' intervenciones sin reporte validado');
  if (pendientes.length) faltan.push(pendientes.length + ' solicitudes pendientes');
  if (!esSi(corte['Financiera cargada'])) faltan.push('Falta cargar la información financiera del corte');
  return { listo: !faltan.length, faltan: faltan, alcance: alcance, enCorte: enCorte };
}

// Calcula Actual con el criterio único del monto (directiva v2, 7.6.4)
function calcularActual(corte, alcance, enCorte, financiera, ultimos) {
  var fin = {}; financiera.forEach(function (f) { fin[f['CUI']] = f; });
  var ult = {}; (ultimos || []).forEach(function (r) { ult[r['COD INTERNO']] = r; });
  var usoCui = {};
  var base = alcance.map(function (r) {
    var cuis = r['Tipo de Código'] === 'CUI' ? partirCodigos(r['Código']) : [];
    var valido = cuis.length > 0 && cuis.every(function (c) { return /^\d+$/.test(c); });
    if (valido) cuis.forEach(function (c) { (usoCui[c] = usoCui[c] || {})[r['COD INTERNO']] = true; });
    return { r: r, cuis: cuis, valido: valido, rep: enCorte[r['COD INTERNO']] || ult[r['COD INTERNO']] || {} };
  });
  var anio = anioCorte(corte);
  var filas = [], vistos = {}, totalMef = 0, totalUnidad = 0, cuiVig = [];
  base.forEach(function (b) {
    var r = b.r, rep = b.rep;
    var reportado = num(rep['Monto reportado S/']);
    var monto, fuente, compartido = 'No', detalle = '';
    if (b.valido) {
      monto = 0; var maxUso = 1;
      b.cuis.forEach(function (c) {
        var costo = fin[c] ? (num(fin[c]['Costo de Inversión']) || 0) : 0;
        monto += costo;
        if (!vistos[c]) { vistos[c] = true; totalMef += costo; }
        var n = Object.keys(usoCui[c]).length; if (n > maxUso) maxUso = n;
        cuiVig.push({ 'COD INTERNO': r['COD INTERNO'], 'CUI': c });
      });
      fuente = 'MEF';
      if (maxUso > 1) compartido = 'Sí (' + maxUso + ' intervenciones)';
      detalle = b.cuis.map(function (c) {
        var costo = fin[c] ? (num(fin[c]['Costo de Inversión']) || 0) : 0;
        return 'CUI ' + c + ' = ' + miles(costo);
      }).join(' | ');
    } else {
      monto = isNaN(reportado) ? 0 : reportado; fuente = 'Reportado por la unidad'; totalUnidad += monto;
    }
    var fila = {
      'Corte': corte, 'Viceministerio': r['Viceministerio'], 'COD INTERNO': r['COD INTERNO'],
      'Entidad Reportante': r['Unidad reportante'], 'Nombre de la intervención': r['Nombre de la intervención'],
      'Área de influencia': r['Área de influencia'], 'Tipo de Código': r['Tipo de Código'],
      'Código': r['Código'], 'Proyecto / Servicio': clave(r['Tipo']),
      'Tipo de Infraestructura': r['Tipo de Infraestructura'], 'Tipo de Órgano ejecutor': r['Tipo de Órgano ejecutor'],
      'Monto reportado por la unidad S/': isNaN(reportado) ? '' : reportado,
      'Monto de la intervención S/': monto, 'Fuente del monto': fuente, 'CUI compartido': compartido,
      'Detalle del monto por CUI': detalle,
      'Beneficiarios': isNaN(num(rep['Beneficiarios'])) ? '' : num(rep['Beneficiarios']),
      'Modalidad de ejecución': rep['Modalidad de ejecución'] || '', 'Fase (Depende de la Modalidad)': rep['Fase'] || '',
      'Crédito suplementario': r['Crédito suplementario'], 'Monto del crédito S/': r['Monto del crédito S/'] === '' ? '' : num(r['Monto del crédito S/']),
      'Hoja de Ruta': r['Hoja de Ruta'], 'Cartera Priorizada': r['Cartera Priorizada'],
      'Estado Situacional': rep['Estado situacional'] || ''
    };
    fila['Financiamiento ' + anio + ' (SI/NO)'] = rep['Financiamiento (SI/NO)'] || '';
    filas.push(fila);
  });
  var enc = ['Corte', 'Viceministerio', 'COD INTERNO', 'Entidad Reportante', 'Nombre de la intervención', 'Área de influencia',
    'Tipo de Código', 'Código', 'Proyecto / Servicio', 'Tipo de Infraestructura', 'Tipo de Órgano ejecutor',
    'Monto reportado por la unidad S/', 'Monto de la intervención S/', 'Fuente del monto', 'CUI compartido',
    'Detalle del monto por CUI', 'Beneficiarios', 'Modalidad de ejecución', 'Fase (Depende de la Modalidad)',
    'Financiamiento ' + anio + ' (SI/NO)', 'Estado Situacional', 'Crédito suplementario', 'Monto del crédito S/',
    'Hoja de Ruta', 'Cartera Priorizada'];
  var kpis = {
    intervenciones: filas.length,
    proyectos: filas.filter(function (f) { return f['Proyecto / Servicio'] === 'PROYECTO'; }).length,
    servicios: filas.filter(function (f) { return f['Proyecto / Servicio'] === 'SERVICIO'; }).length,
    totalMef: totalMef, totalUnidad: totalUnidad, total: totalMef + totalUnidad,
    cuis: Object.keys(vistos).length,
    cuisSinCosto: Object.keys(vistos).filter(function (c) { return !fin[c]; })
  };
  return { enc: enc, filas: filas, kpis: kpis, cuis: cuiVig };
}

// Escribe Actual: fila 1 «Corte del …», fila 2 encabezados, luego datos
function escribirActual(nombre, enc, filas, corte) {
  var h = hoja(nombre, true); h.clearContents();
  h.getRange(1, 1, 1, 1).setValues([[textoCorte(corte)]]);
  h.getRange(2, 1, 1, enc.length).setValues([enc]);
  if (filas.length) {
    var m = filas.map(function (f) { return enc.map(function (k) { var v = f[k]; return v === undefined || v === null ? '' : v; }); });
    var r = h.getRange(3, 1, m.length, enc.length);
    r.setNumberFormat('@');
    ['Monto reportado por la unidad S/', 'Monto de la intervención S/', 'Beneficiarios', 'Monto del crédito S/'].forEach(function (k) {
      var j = enc.indexOf(k); if (j >= 0) h.getRange(3, j + 1, m.length, 1).setNumberFormat('#,##0');
    });
    r.setValues(m);
  }
}

function accPublicarCorte(u, d) {
  if (!esGestor(u)) return { ok: false, mensaje: 'Sin permiso' };
  var corte = corteAbierto(); if (!corte) return { ok: false, mensaje: 'No hay un corte abierto' };
  var cat = leerCatalogos(); var universo = leer(H.UNIVERSO); var reportes = leer(H.REPORTES);
  var est = estadoPublicacion(corte, cat, universo, reportes, leer(H.SOLICITUDES));
  if (d.soloVerificar) return { ok: true, listo: est.listo, faltan: est.faltan };
  if (u.rol !== ROLES.ADMIN) return { ok: false, mensaje: 'Solo el administrador de la plataforma publica el corte' };
  if (!est.listo) return { ok: false, mensaje: 'Aún no se puede publicar: ' + est.faltan.join('; '), faltan: est.faltan };
  var calc = calcularActual(corte['Corte'], est.alcance, est.enCorte, leer(H.FINANCIERA), []);
  // 1) Respaldo de Actual anterior con la etiqueta de su corte
  var cortes = leer(H.CORTES);
  var previo = cortes.filter(function (c) { return c['Estado'] === 'Publicado'; }).sort(function (a, b) { return ordenFecha(b['Corte']) - ordenFecha(a['Corte']); })[0];
  var ss = libro(); var hAct = ss.getSheetByName(H.ACTUAL);
  if (hAct && hAct.getLastRow() > 1) {
    var nombre = previo ? etiquetaCorte(previo['Corte']) : 'Actual anterior'; var n = nombre, k = 2;
    while (ss.getSheetByName(n)) n = nombre + ' (' + (k++) + ')';
    var vals = hAct.getDataRange().getValues();
    var hc = ss.insertSheet(n); var rr = hc.getRange(1, 1, vals.length, vals[0].length); rr.setNumberFormat('@'); rr.setValues(vals);
  }
  // 2) Actual, CUI vigentes y KPIs
  escribirActual(H.ACTUAL, calc.enc, calc.filas, corte['Corte']);
  escribir(H.CUIS, calc.cuis);
  var t = ahora();
  escribir(H.KPIS, [
    { 'Indicador': 'Intervenciones', 'Valor': calc.kpis.intervenciones, 'Unidad': 'Número', 'Corte': corte['Corte'], 'Calculado el': t },
    { 'Indicador': 'Costo de inversión', 'Valor': (calc.kpis.total / 1e6).toFixed(1), 'Unidad': 'Millones de soles', 'Corte': corte['Corte'], 'Calculado el': t },
    { 'Indicador': 'Proyectos', 'Valor': calc.kpis.proyectos, 'Unidad': 'Número', 'Corte': corte['Corte'], 'Calculado el': t },
    { 'Indicador': 'Servicios', 'Valor': calc.kpis.servicios, 'Unidad': 'Número', 'Corte': corte['Corte'], 'Calculado el': t }
  ]);
  // 3) Reportes del corte → histórico y último reporte por intervención
  var delCorte = reportes.filter(function (r) { return r['Corte'] === corte['Corte']; });
  agregar(H.HIST, delCorte);
  var ult = {}; leer(H.ULTIMO).forEach(function (r) { ult[r['COD INTERNO']] = r; });
  delCorte.forEach(function (r) { ult[r['COD INTERNO']] = r; });
  escribir(H.ULTIMO, Object.keys(ult).map(function (k) { return ult[k]; }));
  escribir(H.REPORTES, []);
  // 4) Corte publicado
  cortes.forEach(function (c) {
    if (c['Corte'] === corte['Corte']) {
      c['Estado'] = 'Publicado'; c['Publicado por'] = u.usuario; c['Fecha de publicación'] = t;
      c['Intervenciones'] = String(calc.kpis.intervenciones); c['Monto total S/'] = String(Math.round(calc.kpis.total));
    }
  });
  escribir(H.CORTES, cortes);
  anotar(u.usuario, 'Publicar corte', corte['Corte'] + ' · ' + calc.kpis.intervenciones + ' intervenciones · S/ ' + (calc.kpis.total / 1e6).toFixed(1) + ' MM');
  return { ok: true, kpis: calc.kpis, mensaje: 'Corte del ' + corte['Corte'] + ' publicado: ' + calc.kpis.intervenciones + ' intervenciones, S/ ' + miles(calc.kpis.total / 1e6, 1) + ' MM' };
}


// ── ACCIÓN: ASIGNACIONES (administrador) ───────────────────────────────────
// d.cambios = [{ambito, clave, especialista}] ; especialista '' = quitar asignación
function accGuardarAsignaciones(u, d) {
  if (!esGestor(u)) return { ok: false, mensaje: 'Sin permiso' };
  var asign = leer(H.ASIGNACIONES); var t = ahora(); var n = 0;
  var esp = {}; leer(H.ACCESOS).forEach(function (a) { if (a['Rol'] === ROLES.ESP) esp[clave(a['Usuario'])] = a['Usuario']; });
  var errores = [];
  (d.cambios || []).forEach(function (c) {
    if ([AMBITO_CARPETA, AMBITO_HR].indexOf(c.ambito) < 0) { errores.push('Ámbito no válido'); return; }
    var nuevo = txt(c.especialista);
    if (nuevo && !esp[clave(nuevo)]) { errores.push(nuevo + ' no es especialista'); return; }
    var vig = asign.filter(function (a) { return a['Ámbito'] === c.ambito && a['Clave'] === c.clave && !a['Hasta']; });
    if (vig.length === 1 && clave(vig[0]['Especialista']) === clave(nuevo)) return;
    vig.forEach(function (a) { a['Hasta'] = t; });
    if (nuevo) asign.push({ 'Ámbito': c.ambito, 'Clave': c.clave, 'Especialista': esp[clave(nuevo)], 'Desde': t, 'Hasta': '', 'Asignado por': u.usuario });
    n++;
  });
  if (n) { escribir(H.ASIGNACIONES, asign); anotar(u.usuario, 'Asignaciones', n + ' cambios'); }
  return { ok: !errores.length, n: n, mensaje: n + ' asignaciones actualizadas' + (errores.length ? '. ' + errores.join('; ') : '') };
}


// ── ACCIÓN: EDITAR UNIVERSO (administrador) ────────────────────────────────
function accEditarUniverso(u, d) {
  if (u.rol !== ROLES.ADMIN) return { ok: false, mensaje: 'Solo el administrador edita el universo directamente' };
  if (txt(d.sustento).length < 5) return { ok: false, mensaje: 'Indica el sustento del cambio' };
  var cat = leerCatalogos(); var universo = leer(H.UNIVERSO);
  var fila = universo.filter(function (r) { return r['COD INTERNO'] === d.cod; })[0];
  if (!fila) return { ok: false, mensaje: 'No existe ' + d.cod };
  var prop = normalizarIdentidad(aplicarCambios(fila, d.cambios || {}), cat); prop['Unidad reportante'] = fila['Unidad reportante'];
  var err = validarIdentidad(prop, cat); if (err.length) return { ok: false, mensaje: err.join('; ') };
  var dif = diferencias(fila, prop, CAMPOS_IDENTIDAD);
  if (d.cambios && 'Vigente' in d.cambios) { var v = esSi(d.cambios['Vigente']) ? 'SI' : 'NO'; if (v !== fila['Vigente']) { dif.push({ campo: 'Vigente', unidad: fila['Vigente'], omei: v }); fila['Vigente'] = v; } }
  if (!dif.length) return { ok: true, mensaje: 'Sin cambios' };
  CAMPOS_IDENTIDAD.forEach(function (k) { fila[k] = prop[k]; });
  fila['Última modificación'] = ahora(); fila['Modificado por'] = u.usuario;
  escribir(H.UNIVERSO, universo);
  anotar(u.usuario, 'Editar universo', d.cod + ' · ' + JSON.stringify(dif) + ' · ' + txt(d.sustento));
  return { ok: true, mensaje: 'Universo actualizado' };
}


// ═══════════════════════════════════════════════════════════════════════════
// PREPARACIÓN DE LA BASE Y CARGA INICIAL (se ejecutan desde el menú del Sheet)
// ═══════════════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Intervenciones')
    .addItem('1. Preparar base (crear pestañas y catálogos)', 'prepararBase')
    .addItem('2. Importar carga inicial (pestañas Carga_*)', 'importarCargaInicial')
    .addToUi();
}

function prepararBase() {
  var ss = libro();
  Object.keys(ENC).forEach(function (n) {
    if (!ss.getSheetByName(n)) { var h = ss.insertSheet(n); h.getRange(1, 1, 1, ENC[n].length).setValues([ENC[n]]); h.setFrozenRows(1); }
  });
  if (!ss.getSheetByName(H.ACTUAL)) ss.insertSheet(H.ACTUAL);
  if (!leer(H.CAT_TIPOS).length) escribir(H.CAT_TIPOS, [
    { 'Tipo': 'Proyecto', 'Tipos de código': 'CUI; IDEA; PENDIENTE; No aplica', 'Datos de intervención': 'SI', 'Módulo Intervenciones MTC': 'SI' },
    { 'Tipo': 'Servicio', 'Tipos de código': 'CUI; IDEA; PENDIENTE; No aplica', 'Datos de intervención': 'SI', 'Módulo Intervenciones MTC': 'SI' },
    { 'Tipo': 'Política / Normativa', 'Tipos de código': 'No aplica', 'Datos de intervención': 'NO', 'Módulo Intervenciones MTC': 'NO' },
    { 'Tipo': 'Supervisión', 'Tipos de código': 'No aplica', 'Datos de intervención': 'NO', 'Módulo Intervenciones MTC': 'NO' },
    { 'Tipo': 'Acción', 'Tipos de código': 'No aplica', 'Datos de intervención': 'NO', 'Módulo Intervenciones MTC': 'NO' }
  ]);
  if (!leer(H.CAT_MF).length) escribir(H.CAT_MF, catalogoModalidadFase());
  var listas = leer(H.CAT_LISTAS);
  if (!listas.some(function (l) { return l['Lista'] === 'Departamento'; })) {
    agregar(H.CAT_LISTAS, DEPARTAMENTOS.map(function (d) { return { 'Lista': 'Departamento', 'Valor': d }; }));
  }
  if (!listas.some(function (l) { return l['Lista'] === 'Incluida en'; })) {
    agregar(H.CAT_LISTAS, ['Discurso 28', 'Discurso Premier', 'Plan de Gobierno'].map(function (d) { return { 'Lista': 'Incluida en', 'Valor': d }; }));
  }
  [H.CARGA_UNIVERSO, H.CARGA_ACTUAL, H.CARGA_MAESTRO, H.CARGA_FINANCIERA, H.CARGA_CREDITO].forEach(function (n) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });
  return 'Base preparada';
}

// Catálogo Modalidades-Fase (Anexo N° 04 de la directiva)
function catalogoModalidadFase() {
  var def = {
    'OBRA PÚBLICA': [['Perfil', 'Preinversión'], ['Factibilidad', 'Preinversión'], ['Adjudicación de Expediente Técnico', 'Expediente técnico'],
      ['Con contrato de ET', 'Expediente técnico'], ['Elaboración del ET', 'Expediente técnico'], ['ET aprobado', 'Expediente técnico'],
      ['Adjudicación de Obra', 'Actos previos'], ['Con contrato de Obra', 'Actos previos'], ['Ejecución Contractual de Obra', 'En ejecución'],
      ['Ejecución Suspendida', 'Ejecución suspendida'], ['Gestión Predial', 'Actos previos'], ['Entrega de Obra', 'Concluida u operando'],
      ['Funcionamiento', 'Concluida u operando'], ['Liquidación', 'Concluida u operando']],
    'APP': [['Código Idea', 'Preinversión'], ['Formulación (Preinversión)', 'Preinversión'], ['Estructuración', 'Preinversión'],
      ['Planeamiento y Programación', 'Preinversión'], ['Proceso de Selección (concurso de licitación)', 'Actos previos'],
      ['Ejecución Contractual', 'En ejecución'], ['Elaboración del ET', 'Expediente técnico'], ['Gestión Predial', 'Actos previos'],
      ['Ejecución Suspendida', 'Ejecución suspendida'], ['Operación', 'Concluida u operando'], ['Concluido', 'Concluida u operando']],
    'APP PA': [['Formulación (Preinversión)', 'Preinversión'], ['Estructuración', 'Preinversión'], ['Planeamiento y Programación', 'Preinversión'],
      ['Ejecución Contractual', 'En ejecución'], ['Ejecución Suspendida', 'Ejecución suspendida'], ['Operación', 'Concluida u operando'], ['Concluido', 'Concluida u operando']],
    'ACUERDO GOBIERNO A GOBIERNO (GTOG)': [['Formulación (Preinversión)', 'Preinversión'], ['Evaluación de la modalidad G2G', 'Preinversión'],
      ['Selección de Gob y suscripción del acuerdo', 'Actos previos'], ['Diseño de la contratación', 'Expediente técnico'],
      ['Procura internacional', 'Actos previos'], ['Ejecución del contrato', 'En ejecución'], ['ET', 'Expediente técnico'],
      ['ET y construcción', 'En ejecución'], ['Construcción', 'En ejecución'], ['Finalización', 'Concluida u operando']],
    'OBRAS POR IMPUESTOS': [['Preinversión', 'Preinversión'], ['Priorización', 'Actos previos'], ['Informe previo de la CGR', 'Actos previos'],
      ['Aprobación de las bases', 'Actos previos'], ['Selección', 'Actos previos'], ['Firma de Convenio', 'Actos previos'], ['Concluido', 'Concluida u operando']],
    'ADM. DIRECTA - POR CONTRATA': [['Actos previos', 'Actos previos'], ['Fase de consentimiento', 'Actos previos'], ['Programado', 'Actos previos'],
      ['Ejecución Contractual de Obra', 'En ejecución'], ['Funcionamiento', 'Concluida u operando'], ['Desactivado', 'Desactivado']],
    'CANON/TASA POR COBERTURA': [['Implementación', 'En ejecución'], ['Concluido', 'Concluida u operando']],
    'INVERSIÓN PRIVADA': [['Implementación', 'En ejecución'], ['Concluido', 'Concluida u operando']],
    'POR DEFINIR': [['No iniciado', 'No iniciado'], ['IDEA', 'Preinversión'], ['Formulación (Preinversión)', 'Preinversión']]
  };
  var out = [];
  Object.keys(def).forEach(function (m) { def[m].forEach(function (f, i) { out.push({ 'Modalidad': m, 'N°': String(i + 1), 'Fase': f[0], 'Etapa': f[1] }); }); });
  return out;
}

// Lee una pestaña Carga_* buscando la fila de encabezados (la que contiene «COD INTERNO» o «CUI»)
function leerCarga(nombre, claveEnc) {
  var h = libro().getSheetByName(nombre); if (!h || h.getLastRow() < 2) return [];
  var v = h.getDataRange().getValues(); var iH = -1;
  for (var i = 0; i < Math.min(8, v.length); i++) if (v[i].some(function (c) { return clave(c) === clave(claveEnc); })) { iH = i; break; }
  if (iH < 0) return [];
  var enc = v[iH].map(function (c) { return clave(c); });
  var out = [];
  for (var r = iH + 1; r < v.length; r++) {
    if (!v[r].some(function (c) { return txt(c) !== ''; })) continue;
    var o = {}; enc.forEach(function (k, j) { if (k) o[k] = txt(v[r][j]); }); out.push(o);
  }
  return out;
}
function campo(o, nombres) { for (var i = 0; i < nombres.length; i++) { var k = clave(nombres[i]); if (o[k] !== undefined && o[k] !== '') return o[k]; } return ''; }

/**
 * Carga inicial desde las pestañas (pegar valores, con su fila de encabezados):
 *   Carga_Universo   ← pestaña IntervencionesMTC (universo) de la Hoja de Ruta
 *   Carga_Actual     ← pestaña Actual del Sheet Carpeta_Departamental (último corte publicado)
 *   Carga_Maestro    ← pestaña Maestro_interv de la Hoja de Ruta
 *   Carga_Financiera ← tabla Financiera del consolidador (CUI, Costo de Inversión, PIA, PIM, Devengado Anual)
 *   Carga_Credito    ← hoja del crédito suplementario (COD INTERNO y/o CUI, Monto asignado) — opcional
 * Reescribe Universo, Último reporte, Financiera, catálogos derivados y el corte inicial publicado.
 */
function importarCargaInicial() {
  prepararBase();
  var cu = leerCarga(H.CARGA_UNIVERSO, 'COD INTERNO');
  if (!cu.length) throw new Error('Carga_Universo está vacía');
  var ca = leerCarga(H.CARGA_ACTUAL, 'COD INTERNO');
  var cm = leerCarga(H.CARGA_MAESTRO, 'COD INTERNO');
  var cf = leerCarga(H.CARGA_FINANCIERA, 'CUI');
  var cc = leerCarga(H.CARGA_CREDITO, 'COD INTERNO'); if (!cc.length) cc = leerCarga(H.CARGA_CREDITO, 'CUI');
  var resumen = [];
  var act = {}; ca.forEach(function (r) { act[campo(r, ['COD INTERNO'])] = r; });
  var mae = {}; cm.forEach(function (r) { var k = campo(r, ['COD INTERNO']); if (k && !mae[k]) mae[k] = r; });
  var cred = { cod: {}, cui: {} };
  cc.forEach(function (r) {
    var m = campo(r, ['Monto asignado', 'Monto']); var k = campo(r, ['COD INTERNO']); var c = campo(r, ['CUI', 'Código']);
    if (k) cred.cod[k] = m || '0'; if (c) partirCodigos(c).forEach(function (x) { cred.cui[x] = m || '0'; });
  });
  var mapTipo = { 'PROYECTO': 'Proyecto', 'SERVICIO': 'Servicio', 'SUPERVISION': 'Supervisión', 'POLITICAS': 'Política / Normativa', 'POLITICA': 'Política / Normativa', 'ACCIONES': 'Acción', 'ACCION': 'Acción' };
  var unidades = {}, listas = { 'Tipo de Infraestructura': {}, 'Tipo de Órgano ejecutor': {}, 'Incluida en': {} };
  var vistos = {}, universo = [], ultimos = [], corteIni = '';
  var t = ahora();
  cu.forEach(function (r) {
    var cod = campo(r, ['COD INTERNO']); if (!cod || vistos[cod]) { if (cod) resumen.push('Duplicado omitido: ' + cod); return; }
    vistos[cod] = true;
    var a = act[cod] || {}; var m = mae[cod] || {};
    var sigla = campo(r, ['Entidad Reportante', 'Unidad reportante']);
    var vice = campo(r, ['Viceministerio']);
    if (sigla) unidades[sigla] = unidades[sigla] || vice;
    var tipo = mapTipo[clave(campo(r, ['Proyecto / Servicio', 'Tipo']))] || 'Proyecto';
    var tc = normTipoCodigo(campo(r, ['Tipo de Código']));
    var cods = partirCodigos(campo(r, ['Código']));
    var infra = campo(r, ['Tipo de Infraestructura']) || campo(a, ['Tipo de Infraestructura']);
    var organo = campo(a, ['Tipo de Órgano ejecutor']);
    if (infra) listas['Tipo de Infraestructura'][infra] = 1;
    if (organo) listas['Tipo de Órgano ejecutor'][organo] = 1;
    var hr = esSi(campo(r, ['HOJA DE RUTA'])) ? 'SI' : 'NO';
    var cartera = '', relev = '', incl = '';
    if (hr === 'SI') {
      var perMae = clave(campo(m, ['Periodicidad'])), vigMae = clave(campo(m, ['Vigente'])), tipoMae = clave(campo(m, ['Tipo']));
      cartera = (perMae === 'SEMANAL' && vigMae !== 'NO' && tipoMae === 'INTERVENCIONES') ? 'SI' : 'NO';
      relev = campo(m, ['Relevancia sectorial']);
      incl = campo(m, ['Incluida en', 'Priorizadas']).split(';').map(txt).filter(function (x) { return x && clave(x) !== 'NO PRIORIZADA'; }).join('; ');
      incl.split(';').map(txt).filter(Boolean).forEach(function (x) { listas['Incluida en'][x] = 1; });
    }
    var mc = cred.cod[cod]; if (mc === undefined) cods.forEach(function (c) { if (cred.cui[c] !== undefined) mc = cred.cui[c]; });
    var area = campo(r, ['Área de influencia']);
    universo.push({
      'COD INTERNO': cod, 'Tipo': tipo, 'Vigente': esSi(campo(r, ['Vigente'])) ? 'SI' : 'NO', 'Viceministerio': vice,
      'Unidad reportante': sigla, 'Nombre de la intervención': nombreNormal(campo(r, ['Nombre de la intervención'])),
      'Área de influencia': clave(area) === 'NACIONAL' ? 'NACIONAL' : area.split(',').map(function (x) { return txt(x).toUpperCase(); }).filter(Boolean).join(', '),
      'Tipo de Código': tc, 'Código': (tc === 'CUI' || tc === 'IDEA') ? unirCodigos(cods) : '',
      'Tipo de Infraestructura': infra, 'Tipo de Órgano ejecutor': organo,
      'Crédito suplementario': mc !== undefined ? 'SI' : 'NO', 'Monto del crédito S/': mc !== undefined ? String(num(mc) || 0) : '',
      'Hoja de Ruta': hr, 'Cartera Priorizada': cartera, 'Relevancia sectorial': relev, 'Incluida en': incl,
      'Fecha de alta': t, 'Registrado por': 'Carga inicial', 'Última modificación': '', 'Modificado por': ''
    });
    if (a && Object.keys(a).length) {
      var corte = campo(a, ['Corte']);
      if (corte && ordenFecha(normalizarFecha(corte)) > ordenFecha(corteIni)) corteIni = normalizarFecha(corte);
      ultimos.push({
        'ID': uuid(), 'Corte': normalizarFecha(corte), 'COD INTERNO': cod, 'Unidad reportante': sigla,
        'Monto reportado S/': String(isNaN(num(campo(a, ['Monto reportado por la unidad S/', 'Monto de la intervención S/', 'Monto']))) ? '' : num(campo(a, ['Monto reportado por la unidad S/', 'Monto de la intervención S/', 'Monto']))),
        'Beneficiarios': String(isNaN(num(campo(a, ['Beneficiarios']))) ? '' : num(campo(a, ['Beneficiarios']))),
        'Modalidad de ejecución': normModalidad(campo(a, ['Modalidad de ejecución']) || campo(r, ['Modalidad de ejecución'])),
        'Fase': campo(a, ['Fase (Depende de la Modalidad)', 'Fase']),
        'Financiamiento (SI/NO)': esSi(campo(a, ['Financiamiento 2026 (SI/NO)', 'Financiamiento 2025 (SI/NO)', 'Financiamiento (SI/NO)'])) ? 'SI' : 'NO',
        'Estado situacional': campo(a, ['Estado Situacional']), 'Estado': 'Validado',
        'Registrado por': 'Carga inicial', 'Fecha de registro': t
      });
    }
  });
  // Catálogos derivados de la data vigente
  escribir(H.CAT_UNIDADES, Object.keys(unidades).sort().map(function (s) { return { 'Sigla': s, 'Nombre': s, 'Viceministerio': unidades[s] }; }));
  var otras = leer(H.CAT_LISTAS).filter(function (l) { return !listas[l['Lista']] || !Object.keys(listas[l['Lista']]).length; });
  var nuevas = [];
  Object.keys(listas).forEach(function (n) { Object.keys(listas[n]).sort().forEach(function (v) { nuevas.push({ 'Lista': n, 'Valor': v }); }); });
  escribir(H.CAT_LISTAS, otras.concat(nuevas));
  // Fases fuera del catálogo: se informa para corregir en el primer corte
  var mf = {}; catalogoModalidadFase().forEach(function (f) { mf[f['Modalidad'] + '|' + clave(f['Fase'])] = f['Fase']; });
  var fueraCat = 0;
  ultimos.forEach(function (u) {
    var k = u['Modalidad de ejecución'] + '|' + clave(u['Fase']);
    if (mf[k]) u['Fase'] = mf[k]; else if (u['Modalidad de ejecución'] || u['Fase']) fueraCat++;
  });
  escribir(H.UNIVERSO, universo);
  escribir(H.ULTIMO, ultimos);
  if (cf.length) {
    escribir(H.FINANCIERA, cf.map(function (f) {
      return { 'CUI': campo(f, ['CUI']).replace(/\.0+$/, ''), 'Costo de Inversión': String(num(campo(f, ['Costo de Inversión', 'costo_inversion'])) || 0),
        'PIA': String(num(campo(f, ['PIA'])) || 0), 'PIM': String(num(campo(f, ['PIM'])) || 0),
        'Devengado Anual': String(num(campo(f, ['Devengado Anual', 'Devengado'])) || 0), 'Fuente del costo': campo(f, ['fuente_costo', 'Fuente del costo']),
        'Fecha de consulta': 'Carga inicial' };
    }));
  }
  if (corteIni && !leer(H.CORTES).some(function (c) { return c['Corte'] === corteIni; })) {
    agregar(H.CORTES, [{ 'Corte': corteIni, 'Plazo': corteIni, 'Estado': 'Publicado', 'Abierto por': 'Carga inicial', 'Fecha de apertura': t,
      'Financiera cargada': cf.length ? 'SI' : 'NO', 'Publicado por': 'Carga inicial', 'Fecha de publicación': t }]);
  }
  // Responsable OGAME del maestro → asignaciones de la Hoja de Ruta (por intervención)
  var asign = leer(H.ASIGNACIONES); var nA = 0;
  universo.forEach(function (u) {
    if (u['Hoja de Ruta'] !== 'SI') return;
    var resp = campo(mae[u['COD INTERNO']] || {}, ['Responsable OGAME']);
    if (!resp || clave(resp) === 'NO APLICA') return;
    if (asign.some(function (a) { return a['Ámbito'] === AMBITO_HR && a['Clave'] === u['COD INTERNO'] && !a['Hasta']; })) return;
    asign.push({ 'Ámbito': AMBITO_HR, 'Clave': u['COD INTERNO'], 'Especialista': resp, 'Desde': t, 'Hasta': '', 'Asignado por': 'Carga inicial' }); nA++;
  });
  escribir(H.ASIGNACIONES, asign);
  var sinMaestro = universo.filter(function (u) { return u['Hoja de Ruta'] === 'SI' && !mae[u['COD INTERNO']]; }).length;
  var maestroFuera = cm.filter(function (m) { var k = campo(m, ['COD INTERNO']); return k && !vistos[k] && clave(campo(m, ['Vigente'])) !== 'NO'; }).length;
  var msg = 'Universo: ' + universo.length + ' (' + universo.filter(function (u) { return u['Vigente'] === 'SI'; }).length + ' vigentes) · Último reporte: ' + ultimos.length +
    ' (corte ' + (corteIni || '—') + ') · Financiera: ' + cf.length + ' CUI · Asignaciones HR: ' + nA +
    ' · Hoja de Ruta sin dato en el maestro: ' + sinMaestro + ' · Maestro vigente fuera del universo (pendiente de alinear): ' + maestroFuera +
    ' · Reportes con fase fuera del catálogo: ' + fueraCat + (resumen.length ? ' · ' + resumen.slice(0, 5).join('; ') : '');
  anotar('Carga inicial', 'Importar carga inicial', msg);
  return msg;
}
function normalizarFecha(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, CONFIG_INT.ZONA, 'dd/MM/yyyy');
  var s = txt(v); var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + y; }
  var d = new Date(s); return isNaN(d) ? s : Utilities.formatDate(d, CONFIG_INT.ZONA, 'dd/MM/yyyy');
}
function normModalidad(v) {
  var k = clave(v);
  var mapa = { 'OBRA PUBLICA': 'OBRA PÚBLICA', 'APP': 'APP', 'APP PA': 'APP PA', 'ACUERDO GOBIERNO A GOBIERNO (GTOG)': 'ACUERDO GOBIERNO A GOBIERNO (GTOG)',
    'OBRAS POR IMPUESTOS': 'OBRAS POR IMPUESTOS', 'ADM. DIRECTA - POR CONTRATA': 'ADM. DIRECTA - POR CONTRATA',
    'CANON/TASA POR COBERTURA': 'CANON/TASA POR COBERTURA', 'INVERSION PRIVADA': 'INVERSIÓN PRIVADA', 'POR DEFINIR': 'POR DEFINIR' };
  return mapa[k] || txt(v);
}
