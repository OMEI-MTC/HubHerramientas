/**
 * ═══════════════════════════════════════════════════════════════════════════
 * HOJA DE RUTA MTC — OGAME
 * Backend Apps Script — flujo de validación con paso OGAME condicional
 * ═══════════════════════════════════════════════════════════════════════════
 *
 * FLUJO DE DATOS:
 *   Registrador registra → Histórico + Prev (Validado = No, Validado OGAME = No)
 *   Validador (Vice/SG/Despacho) valida → Prev cols F/G/H
 *     · Mensual: su validación publica directo en Estado Actual
 *     · Semanal SIN responsable OGAME: su validación también publica directo
 *     · Semanal CON responsable OGAME: queda pendiente de OGAME
 *   OGAME valida → Prev col I = Sí → copia Prev → Estado Actual
 *
 * REGLA DEL PASO OGAME (cambio de noviembre 2026):
 *   Una intervención requiere el visto bueno de OGAME solo si en Maestro_interv,
 *   columna "Responsable OGAME", figura un usuario que en la pestaña Accesos tiene
 *   Tipo de Usuario = OGAME (por ejemplo, OGAME-2). Si la celda está vacía o dice
 *   cualquier otra cosa ("NO APLICA", un nombre, etc.), el paso se salta.
 *   Las mensuales nunca requieren OGAME (la vista OGAME solo muestra semanales).
 *
 * COLUMNAS Histórico de Intervenciones:
 *   A: Fecha de registro   B: COD INTERNO   C: Intervención Estratégica
 *   D: Estado situacional  E: Avance        F: Riesgos de continuidad
 *   G: Registrado por      H: Fecha y hora de registro
 *
 * COLUMNAS Prev - Interv:
 *   A: Fecha de registro (corte del registrador — no cambia cuando actúa el validador/OGAME)
 *   B: COD INTERNO         C: Intervención Estratégica
 *   D: Estado situacional  E: Avance
 *   F: Validado (Sí/No)    G: Estado de Validación (Pendiente/Observado/Validado)
 *   H: Observación         I: Validado OGAME (Sí/No)
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────────────────
var SPREADSHEET_ID = '1Bjp130MWeO7sQx-EzrXh_hOs8GlGc2qg9Kq17WNLdcY';

var HOJA_HIST_INTERV   = 'Histórico de Intervenciones';
var HOJA_HIST_HITOS    = 'Histórico de Hitos';
var HOJA_ESTADO_INTERV = 'Estado Actual - Interv';
var HOJA_ESTADO_HITOS  = 'Estado Actual - Hitos';
var HOJA_PREV_INTERV   = 'Prev - Interv';
var HOJA_PREV_HITOS    = 'Prev - Hitos';
var HOJA_ACCESOS       = 'Accesos';

// Maestro de intervenciones: se ubica por su gid (el mismo que usa el TSV publicado),
// así no depende del nombre de la pestaña. De aquí se leen Periodicidad y Responsable OGAME.
var GID_MAESTRO_INTERV = 352174490;

var TZ = 'America/Lima';


// ── PUNTO DE ENTRADA ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    var datos = JSON.parse(e.postData.contents);
    switch (datos.accion) {
      case 'registrarHojaRuta':        return responder(registrarHojaRuta(datos));
      case 'validarRegistro':          return responder(validarRegistro(datos));
      case 'validarLote':              return responder(validarLote(datos));
      case 'validarOgame':             return responder(validarOgame(datos));
      case 'actualizarPriorizacion':   return responder(actualizarPriorizacion(datos));
      case 'actualizarHitosHabilitados': return responder(actualizarHitosHabilitados(datos));
      default: return responder({ ok: false, mensaje: 'Acción no reconocida: ' + datos.accion });
    }
  } catch (error) {
    return responder({ ok: false, mensaje: 'Error: ' + error.message });
  }
}

function doGet() {
  return responder({ ok: true, mensaje: 'Web App Hoja de Ruta MTC activo' });
}

function responder(objeto) {
  return ContentService
    .createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}


// ── CRONÓMETRO ─────────────────────────────────────────────────────────────
// Escribe en el registro de la ejecución cuántos milisegundos pasaron desde el
// inicio hasta cada paso. Se ve en Apps Script → Ejecuciones → clic en la fila.
function crearCronometro(nombreAccion) {
  var t0 = Date.now();
  return function (paso) {
    console.log('[' + nombreAccion + '] ' + paso + ': ' + (Date.now() - t0) + ' ms');
  };
}


// ── ACCIÓN 1: REGISTRO (Registrador o Validador editando contenido) ────────
function registrarHojaRuta(datos) {
  var marca = crearCronometro('registrar ' + datos.codigo);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  marca('lock obtenido');

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ahora         = new Date();
    var fechaCorte    = Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy');
    var fechaHora     = Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm');
    var registradoPor = datos.registradoPor || datos.organo || '';

    // ── 1. Histórico de Intervenciones ──
    var hojaHistInterv = ss.getSheetByName(HOJA_HIST_INTERV);
    escribirFila(hojaHistInterv, [
      fechaCorte,      // A
      datos.codigo,    // B: COD INTERNO
      datos.nombre,    // C
      datos.estado,    // D
      datos.avance,    // E
      '',              // F: Riesgos de continuidad (se llena directo en Sheets)
      registradoPor,   // G
      fechaHora        // H
    ]);
    marca('histórico interv escrito');

    // ── 2. Histórico de Hitos ──
    if (datos.hitos && datos.hitos.length > 0) {
      var hojaHistHitos = ss.getSheetByName(HOJA_HIST_HITOS);
      var filasHist = datos.hitos.map(function(h) {
        return [
          fechaCorte, datos.codigo, datos.nombre,
          h.codhito, h.nombre || '',
          h.fechaCumplimiento || '', h.estadoCumplimiento || ''
        ];
      });
      escribirFilas(hojaHistHitos, filasHist);
      marca('histórico hitos escrito');
    }

    // ── 3. Prev - Interv ──
    // Al registrar: resetea Validado Vice = No, Estado = Pendiente,
    // Observación vacía, Validado OGAME = No
    var hojaPrevInterv = ss.getSheetByName(HOJA_PREV_INTERV);
    var filaPrev = buscarFilaPorValor(hojaPrevInterv, 2, datos.codigo);
    var valoresPrev = [
      fechaCorte,                        // A: Fecha de registro (corte)
      datos.codigo,                      // B: COD INTERNO
      datos.nombre,                      // C
      datos.estado,                      // D
      datos.avance,                      // E
      'No',                              // F: Validado Vice
      'Pendiente',                       // G: Estado de Validación
      '',                                // H: Observación
      'No'                               // I: Validado OGAME
    ];
    if (filaPrev > 0) {
      hojaPrevInterv.getRange(filaPrev, 1, 1, valoresPrev.length).setValues([valoresPrev]);
    } else {
      escribirFila(hojaPrevInterv, valoresPrev);
    }
    marca('prev interv escrito');

    // ── 4. Prev - Hitos ──
    if (datos.hitos && datos.hitos.length > 0) {
      var hojaPrevHitos = ss.getSheetByName(HOJA_PREV_HITOS);
      datos.hitos.forEach(function(h) {
        var filaH = buscarFilaPorValor(hojaPrevHitos, 4, h.codhito);
        var vals = [
          fechaCorte, datos.codigo, datos.nombre,
          h.codhito, h.nombre || '',
          h.fechaCumplimiento || '', h.estadoCumplimiento || ''
        ];
        if (filaH > 0) {
          hojaPrevHitos.getRange(filaH, 1, 1, vals.length).setValues([vals]);
        } else {
          escribirFila(hojaPrevHitos, vals);
        }
      });
      marca('prev hitos escrito');
    }

    return { ok: true, fechaRegistro: fechaCorte };

  } finally {
    lock.releaseLock();
    marca('fin');
  }
}


// ── ACCIÓN 2: VALIDAR (Validador) ──────────────────────────────────────────
// Dos entradas, un mismo motor (procesarValidaciones):
//   · validarRegistro → una intervención (lápiz). Puede traer estado situacional e hitos.
//   · validarLote     → varias intervenciones ("Guardar cambios"). Solo estado de
//                       validación y observación.
// NO actualiza el corte (col A). Sí escribe en Histórico para trazabilidad.
// Publica en Estado Actual si está Validado y NO requiere OGAME (ver requiereOgame).

function validarRegistro(datos) {
  var r = ejecutarValidaciones([datos], datos.registradoPor || '', 'validar ' + datos.codigo);
  var res = r.resultados[0];
  return res.ok ? { ok: true, publicado: res.publicado } : { ok: false, mensaje: res.mensaje };
}

function validarLote(datos) {
  var cambios = datos.cambios || [];
  if (!cambios.length) return { ok: true, resultados: [] };
  var r = ejecutarValidaciones(cambios, datos.registradoPor || '', 'lote de ' + cambios.length);
  var fallidos = r.resultados.filter(function(x) { return !x.ok; }).length;
  return { ok: fallidos === 0, resultados: r.resultados };
}

function ejecutarValidaciones(lista, registradoPor, etiqueta) {
  var marca = crearCronometro(etiqueta);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  marca('lock obtenido');
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var r = procesarValidaciones(ss, lista, registradoPor, marca);
    marca('escritura terminada');
    return r;
  } finally {
    lock.releaseLock();
    marca('fin');
  }
}

// Motor: FASE 1 lee todo lo necesario una sola vez; FASE 2 solo escribe.
// Así Sheets no tiene que recalcular entre una lectura y otra.
function procesarValidaciones(ss, lista, registradoPor, marca) {
  var fechaHora = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm');

  // ── FASE 1: lectura ──
  var hojaPrev  = ss.getSheetByName(HOJA_PREV_INTERV);
  var hojaPrevH = ss.getSheetByName(HOJA_PREV_HITOS);
  var hojaEst   = ss.getSheetByName(HOJA_ESTADO_INTERV);
  var hojaEstH  = ss.getSheetByName(HOJA_ESTADO_HITOS);
  var hojaHist  = ss.getSheetByName(HOJA_HIST_INTERV);

  var prev  = leerTabla(hojaPrev, 9);
  var prevH = leerTabla(hojaPrevH, 7);
  var est   = leerTabla(hojaEst, 2);
  var estH  = leerTabla(hojaEstH, 4);
  var info  = leerInfoMaestro(ss);          // { per: {...}, resp: {...} }
  var usuariosOgame = leerUsuariosOgame(ss); // { 'OGAME-2': true, ... }

  var filaPrev = indexarFilas(prev, 1);    // COD INTERNO → fila
  var filaPrevH = indexarFilas(prevH, 3);  // CODHITO → fila
  var filaEst  = indexarFilas(est, 1);
  var filaEstH = indexarFilas(estH, 3);
  var hitosPorCod = {};
  prevH.forEach(function(f) {
    var c = String(f[1]).trim();
    if (c) (hitosPorCod[c] = hitosPorCod[c] || []).push({
      codhito: f[3], nombre: f[4], fechaCumplimiento: f[5], estadoCumplimiento: f[6]
    });
  });

  // Próximas filas libres (se calculan ahora, antes de escribir nada)
  var sigHist  = siguienteFilaLibre(hojaHist);
  var sigPrevH = siguienteFilaLibre(hojaPrevH);
  var sigEst   = siguienteFilaLibre(hojaEst);
  var sigEstH  = siguienteFilaLibre(hojaEstH);
  marca('lectura terminada');

  // ── FASE 2: escritura ──
  var filasHistorico = [];
  var resultados = [];

  lista.forEach(function(item) {
    var codigo = String(item.codigo || '').trim();
    var fila = filaPrev[codigo];
    if (!fila) {
      resultados.push({ codigo: codigo, ok: false, mensaje: 'No existe registro previo' });
      return;
    }
    var f = prev[fila - 2];
    var corte  = f[0];
    var nombre = item.nombre || f[2];
    var estado = item.estado !== undefined ? item.estado : f[3];
    var avance = f[4];
    var validadoTxt = (item.validado || '').toString().toLowerCase();
    var esValidado  = validadoTxt === 'sí' || validadoTxt === 'si';
    var estadoVal   = item.estadoVal || (esValidado ? 'Validado' : 'Pendiente');
    var observacion = item.observacion || '';

    // Prev - Interv: C:D solo si el validador editó contenido (lápiz)
    if (item.estado !== undefined) {
      hojaPrev.getRange(fila, 3, 1, 2).setValues([[nombre, estado]]);
    }
    // Prev - Interv: F, G, H — NO se toca A (corte) ni I (Validado OGAME)
    hojaPrev.getRange(fila, 6, 1, 3).setValues([[esValidado ? 'Sí' : 'No', estadoVal, observacion]]);

    // Prev - Hitos (solo si vienen editados desde el lápiz)
    var hitos = hitosPorCod[codigo] || [];
    if (item.hitos && item.hitos.length > 0) {
      item.hitos.forEach(function(h) {
        var vals = [corte, codigo, nombre, h.codhito, h.nombre || '',
                    h.fechaCumplimiento || '', h.estadoCumplimiento || ''];
        var fh = filaPrevH[String(h.codhito).trim()];
        if (!fh) { fh = sigPrevH++; filaPrevH[String(h.codhito).trim()] = fh; }
        hojaPrevH.getRange(fh, 1, 1, 7).setValues([vals]);
      });
      hitos = item.hitos;
    }

    // Histórico (se escribe todo junto al final)
    filasHistorico.push([corte, codigo, nombre, estado, '', '', registradoPor, fechaHora]);

    // Publicación: validado y sin paso OGAME pendiente
    var publicado = false;
    if (esValidado && !requiereOgame(codigo, info, usuariosOgame, item.periodicidad)) {
      if (filaEst[codigo]) {
        hojaEst.getRange(filaEst[codigo], 1, 1, 5).setValues([[corte, codigo, nombre, estado, avance]]);
      } else {
        filaEst[codigo] = sigEst++;
        hojaEst.getRange(filaEst[codigo], 1, 1, 7)
               .setValues([[corte, codigo, nombre, estado, avance, 'No', 'Sí']]);
      }
      hitos.forEach(function(h) {
        var ch = String(h.codhito).trim();
        if (!ch) return;
        if (!filaEstH[ch]) filaEstH[ch] = sigEstH++;
        hojaEstH.getRange(filaEstH[ch], 1, 1, 7).setValues([[
          corte, codigo, nombre, ch, h.nombre || '',
          h.fechaCumplimiento || '', h.estadoCumplimiento || ''
        ]]);
      });
      publicado = true;
    }

    resultados.push({ codigo: codigo, ok: true, publicado: publicado });
  });

  if (filasHistorico.length) {
    hojaHist.getRange(sigHist, 1, filasHistorico.length, 8).setValues(filasHistorico);
  }

  return { resultados: resultados };
}

// Crea un índice { valor de la columna → número de fila en la hoja }
function indexarFilas(tabla, columna) {
  var idx = {};
  tabla.forEach(function(f, i) {
    var v = String(f[columna]).trim();
    if (v) idx[v] = i + 2;
  });
  return idx;
}


// ── ACCIÓN 3: VALIDAR OGAME ────────────────────────────────────────────────
// Puede editar contenido Y/O marcar como validado.
// · Editar: actualiza estado situacional, avance e hitos en Prev. No cambia el corte
//   ni la validación del órgano (cols F/G/H). Deja Validado OGAME = No si no valida.
// · Validar (Validado OGAME = Sí): además copia Prev → Estado Actual.
function validarOgame(datos) {
  var marca = crearCronometro('ogame ' + datos.codigo);
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  marca('lock obtenido');

  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var ahora         = new Date();
    var fechaHora     = Utilities.formatDate(ahora, TZ, 'dd/MM/yyyy HH:mm');
    var registradoPor = datos.registradoPor || '';
    var esValidado    = (datos.validadoOgame || '').toString().toLowerCase() === 'sí' ||
                        (datos.validadoOgame || '').toString().toLowerCase() === 'si';

    var hojaPrev = ss.getSheetByName(HOJA_PREV_INTERV);
    var filaPrev = buscarFilaPorValor(hojaPrev, 2, datos.codigo);

    if (filaPrev < 0) {
      return { ok: false, mensaje: 'No existe registro previo para: ' + datos.codigo };
    }

    var filaData = hojaPrev.getRange(filaPrev, 1, 1, 9).getValues()[0];
    var fechaCorteExistente = filaData[0];

    // Si OGAME editó contenido, actualizar D y E además de I
    var nombre = datos.nombre || filaData[2];
    var estado = datos.estado !== undefined ? datos.estado : filaData[3];
    var avance = datos.avance !== undefined ? datos.avance : filaData[4];

    if (datos.estado !== undefined || datos.avance !== undefined) {
      // OGAME editó contenido: actualizar C, D, E
      hojaPrev.getRange(filaPrev, 3, 1, 3).setValues([[nombre, estado, avance]]);
    }

    // Actualizar col I: Validado OGAME
    hojaPrev.getRange(filaPrev, 9).setValue(esValidado ? 'Sí' : 'No');
    marca('prev escrito');

    // Trazabilidad en Histórico
    var hojaHistInterv = ss.getSheetByName(HOJA_HIST_INTERV);
    escribirFila(hojaHistInterv, [
      fechaCorteExistente, datos.codigo, nombre,
      estado, avance, '',
      registradoPor, fechaHora
    ]);
    marca('histórico escrito');

    // Hitos editados por OGAME: se guardan en Prev SIEMPRE (valide o no) y quedan en el
    // Histórico de Hitos. Antes solo se guardaban si OGAME validaba en el mismo paso,
    // y una edición sin visto bueno se perdía al recargar.
    var hojaPrevH = ss.getSheetByName(HOJA_PREV_HITOS);
    var hitosData = null;
    if (datos.hitos && datos.hitos.length > 0) {
      datos.hitos.forEach(function(h) {
        var filaH = buscarFilaPorValor(hojaPrevH, 4, h.codhito);
        var vals = [
          fechaCorteExistente, datos.codigo, nombre,
          h.codhito, h.nombre || '',
          h.fechaCumplimiento || '', h.estadoCumplimiento || ''
        ];
        if (filaH > 0) {
          hojaPrevH.getRange(filaH, 1, 1, vals.length).setValues([vals]);
        } else {
          escribirFila(hojaPrevH, vals);
        }
      });
      escribirFilas(ss.getSheetByName(HOJA_HIST_HITOS), datos.hitos.map(function(h) {
        return [fechaCorteExistente, datos.codigo, nombre, h.codhito, h.nombre || '',
                h.fechaCumplimiento || '', h.estadoCumplimiento || ''];
      }));
      hitosData = datos.hitos;
      marca('hitos guardados');
    }

    // Si OGAME validó: copiar Prev → Estado Actual
    if (esValidado) {
      if (!hitosData) hitosData = leerHitosPorCodigo(hojaPrevH, datos.codigo);
      copiarPrevAEstadoActual(ss, datos.codigo, nombre, hitosData, fechaCorteExistente, estado, avance);
      marca('publicado en Estado Actual');
    }

    return { ok: true };

  } finally {
    lock.releaseLock();
    marca('fin');
  }
}


// ── HELPER: copiar Prev → Estado Actual ───────────────────────────────────
function copiarPrevAEstadoActual(ss, codigo, nombre, hitos, fechaReg, estado, avance) {
  var hojaEstInterv = ss.getSheetByName(HOJA_ESTADO_INTERV);
  var filaEst = buscarFilaPorValor(hojaEstInterv, 2, codigo);

  if (filaEst > 0) {
    hojaEstInterv.getRange(filaEst, 1, 1, 5).setValues([[
      fechaReg, codigo, nombre,
      estado !== undefined ? estado : '',
      avance  !== undefined ? avance  : ''
    ]]);
  } else {
    escribirFila(hojaEstInterv, [
      fechaReg, codigo, nombre,
      estado !== undefined ? estado : '',
      avance  !== undefined ? avance  : '',
      'No', 'Sí'
    ]);
  }

  if (hitos && hitos.length > 0) {
    var hojaEstHitos = ss.getSheetByName(HOJA_ESTADO_HITOS);
    hitos.forEach(function(h) {
      var codhito = h.codhito  || h[3] || '';
      var nomHito = h.nombre   || h[4] || '';
      var fecha   = h.fechaCumplimiento  || h[5] || '';
      var estH    = h.estadoCumplimiento || h[6] || '';
      var filaH   = buscarFilaPorValor(hojaEstHitos, 4, codhito);
      var vals    = [fechaReg, codigo, nombre, codhito, nomHito, fecha, estH];
      if (filaH > 0) {
        hojaEstHitos.getRange(filaH, 1, 1, 7).setValues([vals]);
      } else {
        escribirFila(hojaEstHitos, vals);
      }
    });
  }
}


// ── HELPER: leer hitos de Prev - Hitos por COD INTERNO ────────────────────
function leerHitosPorCodigo(hoja, codigo) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  var datos = hoja.getRange(2, 1, ultima - 1, 7).getValues();
  var resultado = [];
  datos.forEach(function(fila) {
    if (String(fila[1]).trim() === String(codigo).trim()) {
      resultado.push({
        codhito:            fila[3],
        nombre:             fila[4],
        fechaCumplimiento:  fila[5],
        estadoCumplimiento: fila[6]
      });
    }
  });
  return resultado;
}


// ── HELPER: datos del maestro (periodicidad y responsable OGAME) ──────────
function obtenerHojaMaestro(ss) {
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (hojas[i].getSheetId() === GID_MAESTRO_INTERV) return hojas[i];
  }
  return null;
}

// Devuelve { per: { 'COD-1': 'MENSUAL', ... }, resp: { 'COD-1': 'OGAME-2', ... } }
function leerInfoMaestro(ss) {
  var info = { per: {}, resp: {} };
  var hoja = obtenerHojaMaestro(ss);
  if (!hoja) {
    console.log('No se encontró el maestro con gid ' + GID_MAESTRO_INTERV);
    return info;
  }
  var datos = hoja.getDataRange().getValues();
  // El encabezado se busca en las primeras 5 filas por si hay títulos arriba
  for (var h = 0; h < Math.min(5, datos.length); h++) {
    var enc  = datos[h].map(function(x) { return String(x).trim(); });
    var cCod = enc.indexOf('COD INTERNO');
    var cPer = enc.indexOf('Periodicidad');
    var cRes = enc.indexOf('Responsable OGAME');
    if (cCod >= 0 && cPer >= 0) {
      for (var r = h + 1; r < datos.length; r++) {
        var cod = String(datos[r][cCod]).trim();
        if (!cod) continue;
        info.per[cod] = String(datos[r][cPer]).trim().toUpperCase();
        info.resp[cod] = cRes >= 0 ? String(datos[r][cRes]).trim() : '';
      }
      if (cRes < 0) console.log('El maestro no tiene la columna "Responsable OGAME": no se exigirá OGAME');
      return info;
    }
  }
  console.log('El maestro no tiene las columnas "COD INTERNO" y "Periodicidad"');
  return info;
}

// Compatibilidad con funciones anteriores
function leerMapaPeriodicidad(ss) {
  return leerInfoMaestro(ss).per;
}

function obtenerPeriodicidad(ss, codigo) {
  return leerMapaPeriodicidad(ss)[String(codigo).trim()] || '';
}

// Usuarios con Tipo de Usuario = OGAME en la pestaña Accesos → { 'OGAME-2': true, ... }
// (el usuario se guarda tal cual está escrito, para compararlo igual que la página)
function leerUsuariosOgame(ss) {
  var usuarios = {};
  var hoja = ss.getSheetByName(HOJA_ACCESOS);
  if (!hoja) { console.log('No se encontró la pestaña ' + HOJA_ACCESOS); return usuarios; }
  var datos = hoja.getDataRange().getValues();
  for (var h = 0; h < Math.min(5, datos.length); h++) {
    var enc  = datos[h].map(function(x) { return String(x).trim(); });
    var cUsr = enc.indexOf('Usuario');
    var cTip = enc.indexOf('Tipo de Usuario');
    if (cUsr >= 0 && cTip >= 0) {
      for (var r = h + 1; r < datos.length; r++) {
        var usr = String(datos[r][cUsr]).trim();
        var tip = String(datos[r][cTip]).trim().toUpperCase();
        if (usr && tip === 'OGAME') usuarios[usr] = true;
      }
      return usuarios;
    }
  }
  console.log('Accesos no tiene las columnas "Usuario" y "Tipo de Usuario"');
  return usuarios;
}

// ¿Esta intervención necesita el visto bueno de OGAME para publicarse?
//   · Mensual → nunca.
//   · Otra periodicidad → solo si "Responsable OGAME" es un usuario OGAME de Accesos.
function requiereOgame(codigo, info, usuariosOgame, periodicidadRespaldo) {
  var cod = String(codigo).trim();
  var per = info.per[cod] || String(periodicidadRespaldo || '').toUpperCase();
  if (per === 'MENSUAL') return false;
  // Comparación exacta (mayúsculas incluidas), igual que la vista OGAME de la página:
  // si no coincide exactamente, ningún especialista la vería y quedaría trabada.
  var resp = String(info.resp[cod] || '').trim();
  return resp !== '' && usuariosOgame[resp] === true;
}


// ── ACCIÓN 4: PRIORIZACIÓN ─────────────────────────────────────────────────
function actualizarPriorizacion(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hoja = ss.getSheetByName(HOJA_ESTADO_INTERV);
    var fila = buscarFilaPorValor(hoja, 2, datos.codigo);
    if (fila > 0) {
      hoja.getRange(fila, 6).setValue(datos.priorizado);
    } else {
      var f = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
      escribirFila(hoja, [f, datos.codigo, '', '', '', datos.priorizado, 'Sí']);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}


// ── ACCIÓN 5: HABILITACIÓN DE HITOS ────────────────────────────────────────
function actualizarHitosHabilitados(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    var hoja = ss.getSheetByName(HOJA_ESTADO_INTERV);
    var fila = buscarFilaPorValor(hoja, 2, datos.codigo);
    if (fila > 0) {
      hoja.getRange(fila, 7).setValue(datos.habilitados);
    } else {
      var f = Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy');
      escribirFila(hoja, [f, datos.codigo, '', '', '', 'No', datos.habilitados]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}


// ── UTILIDADES ─────────────────────────────────────────────────────────────
function buscarFilaPorValor(hoja, columna, valor) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return -1;
  var datos = hoja.getRange(2, columna, ultimaFila - 1, 1).getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === String(valor).trim()) return i + 2;
  }
  return -1;
}

function escribirFila(hoja, valores) {
  var fila = siguienteFilaLibre(hoja);
  hoja.getRange(fila, 1, 1, valores.length).setValues([valores]);
}

function escribirFilas(hoja, arrayDeFilas) {
  if (arrayDeFilas.length === 0) return;
  var fila  = siguienteFilaLibre(hoja);
  var ancho = arrayDeFilas[0].length;
  hoja.getRange(fila, 1, arrayDeFilas.length, ancho).setValues(arrayDeFilas);
}

// Busca la primera fila libre mirando SOLO la columna B (COD INTERNO),
// que siempre está llena en todas las hojas donde se escribe.
function siguienteFilaLibre(hoja) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 2;
  var colB = hoja.getRange(1, 2, ultima, 1).getValues();
  for (var i = colB.length - 1; i >= 0; i--) {
    if (colB[i][0] !== '' && colB[i][0] !== null) return i + 2;
  }
  return 2;
}


// ── REPARACIÓN (ejecutar a mano) ─────────────────────────────────────────
// Publica en Estado Actual todas las intervenciones que ya están validadas por
// su órgano en Prev - Interv (col F = Sí) y que NO requieren OGAME:
// todas las mensuales y las semanales sin responsable OGAME válido.
// Sirve para publicar de una vez las que quedaron esperando un visto bueno de
// OGAME que nadie podía dar. No toca el corte (usa la fecha de Prev) ni cambia Prev.
// Se puede ejecutar más de una vez sin duplicar nada.
function republicarValidadosSinOgame() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // ── Fase 1: lectura ──
  var info     = leerInfoMaestro(ss);
  var usrOgame = leerUsuariosOgame(ss);
  var hojaEst  = ss.getSheetByName(HOJA_ESTADO_INTERV);
  var hojaEstH = ss.getSheetByName(HOJA_ESTADO_HITOS);

  var prev  = leerTabla(ss.getSheetByName(HOJA_PREV_INTERV), 9);
  var prevH = leerTabla(ss.getSheetByName(HOJA_PREV_HITOS), 7);
  var est   = leerTabla(hojaEst, 2);   // solo A:B, para ubicar cada código
  var estH  = leerTabla(hojaEstH, 4);  // solo A:D, para ubicar cada hito

  var filaEst = {};
  est.forEach(function(f, i) {
    var c = String(f[1]).trim();
    if (c) filaEst[c] = i + 2;
  });
  var filaEstH = {};
  estH.forEach(function(f, i) {
    var c = String(f[3]).trim();
    if (c) filaEstH[c] = i + 2;
  });
  var hitosPorCod = {};
  prevH.forEach(function(f) {
    var c = String(f[1]).trim();
    if (!c) return;
    (hitosPorCod[c] = hitosPorCod[c] || []).push(f);
  });
  console.log('Lectura terminada: ' + (Date.now() - t0) + ' ms');

  // ── Fase 2: armar los cambios en memoria ──
  var cambiosInterv = [];
  var cambiosHitos  = [];
  var nuevasInterv  = [];
  var nuevosHitos   = [];
  var publicadas    = [];

  prev.forEach(function(f) {
    var codigo   = String(f[1]).trim();
    var validado = String(f[5]).trim().toLowerCase();
    if (!codigo) return;
    if (validado !== 'sí' && validado !== 'si') return;
    if (!info.per[codigo]) return;                       // no está en el maestro
    if (requiereOgame(codigo, info, usrOgame)) return;   // espera a OGAME

    var valores = [f[0], codigo, f[2], f[3], f[4]];
    if (filaEst[codigo]) cambiosInterv.push({ fila: filaEst[codigo], valores: valores });
    else                 nuevasInterv.push(valores.concat(['No', 'Sí']));

    (hitosPorCod[codigo] || []).forEach(function(h) {
      var codhito = String(h[3]).trim();
      if (!codhito) return;
      var vh = [f[0], codigo, f[2], codhito, h[4], h[5], h[6]];
      if (filaEstH[codhito]) cambiosHitos.push({ fila: filaEstH[codhito], valores: vh });
      else                   nuevosHitos.push(vh);
    });

    publicadas.push(codigo);
  });

  // ── Fase 3: escritura (sin lecturas intermedias) ──
  cambiosInterv.forEach(function(c) {
    hojaEst.getRange(c.fila, 1, 1, 5).setValues([c.valores]);
  });
  cambiosHitos.forEach(function(c) {
    hojaEstH.getRange(c.fila, 1, 1, 7).setValues([c.valores]);
  });
  if (nuevasInterv.length) escribirFilas(hojaEst, nuevasInterv);
  if (nuevosHitos.length)  escribirFilas(hojaEstH, nuevosHitos);

  Logger.log('Publicadas ' + publicadas.length + ' intervenciones validadas sin paso OGAME (' +
             nuevasInterv.length + ' nuevas en Estado Actual) en ' +
             Math.round((Date.now() - t0) / 1000) + ' s: ' + publicadas.join(', '));
}

// Nombre anterior: se mantiene para no romper nada que lo invoque
function republicarMensualesValidados() {
  republicarValidadosSinOgame();
}

// Lee desde la fila 2 hasta la última, las primeras "ancho" columnas.
function leerTabla(hoja, ancho) {
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  return hoja.getRange(2, 1, ultima - 1, ancho).getValues();
}


// ── PRUEBAS (solo lectura, no escriben en la hoja) ────────────────────────
function verificarDestino() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('ESCRIBIENDO EN: ' + ss.getName() + ' — ' + ss.getUrl());
  ss.getSheets().forEach(function(h) {
    Logger.log('  · "' + h.getName() + '"  (última fila: ' + h.getLastRow() + ')');
  });
}

function probarConexion() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hojas = [HOJA_HIST_INTERV, HOJA_HIST_HITOS,
               HOJA_ESTADO_INTERV, HOJA_ESTADO_HITOS,
               HOJA_PREV_INTERV, HOJA_PREV_HITOS, HOJA_ACCESOS];
  hojas.forEach(function(nombre) {
    var h = ss.getSheetByName(nombre);
    Logger.log(h ? '✓ "' + nombre + '" — fila ' + h.getLastRow()
                 : '✗ NO ENCONTRADA: "' + nombre + '"');
  });
  var m = obtenerHojaMaestro(ss);
  Logger.log(m ? '✓ Maestro encontrado: "' + m.getName() + '"'
               : '✗ NO ENCONTRADO el maestro (gid ' + GID_MAESTRO_INTERV + ')');
}

function probarPeriodicidad() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var mapa = leerMapaPeriodicidad(ss);
  var codigos = Object.keys(mapa);
  Logger.log('Intervenciones en el maestro: ' + codigos.length);
  codigos.slice(0, 10).forEach(function(c) { Logger.log('  ' + c + ' → ' + mapa[c]); });
}

// Resume qué intervenciones exigirán OGAME con la regla nueva. No escribe nada.
function probarReglaOgame() {
  var ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  var info = leerInfoMaestro(ss);
  var usr  = leerUsuariosOgame(ss);
  Logger.log('Usuarios OGAME en Accesos: ' + Object.keys(usr).join(', '));
  var conOgame = [], sinOgame = 0, mensuales = 0;
  Object.keys(info.per).forEach(function(c) {
    if (info.per[c] === 'MENSUAL') { mensuales++; return; }
    if (requiereOgame(c, info, usr)) conOgame.push(c + ' → ' + info.resp[c]);
    else sinOgame++;
  });
  Logger.log('Mensuales (publican al validar el órgano): ' + mensuales);
  Logger.log('No mensuales SIN responsable OGAME (publican al validar el órgano): ' + sinOgame);
  Logger.log('No mensuales CON responsable OGAME (esperan a OGAME): ' + conOgame.length);
  conOgame.slice(0, 30).forEach(function(x) { Logger.log('  ' + x); });
}

// ⚠ Las funciones siguientes ESCRIBEN una intervención ficticia "PRUEBA-1" en la
// hoja real. No ejecutarlas salvo en una copia del archivo.
function probarRegistro() {
  var r = registrarHojaRuta({
    accion: 'registrarHojaRuta',
    codigo: 'PRUEBA-1', nombre: 'Intervención de prueba',
    estado: 'Estado de prueba', avance: '50',
    validado: 'No', registradoPor: 'DGPRTM-1',
    hitos: [
      { codhito: 'PRUEBA-1-1', nombre: 'Hito 1',
        fechaCumplimiento: 'ago-26', estadoCumplimiento: 'En proceso' }
    ]
  });
  Logger.log(JSON.stringify(r));
}

function probarValidacion() {
  var r = validarRegistro({
    accion: 'validarRegistro', codigo: 'PRUEBA-1',
    estadoVal: 'Observado', validado: 'No',
    observacion: 'Falta precisar el avance físico.',
    registradoPor: 'VMT-1'
  });
  Logger.log(JSON.stringify(r));
}

function probarLote() {
  var r = validarLote({
    accion: 'validarLote', registradoPor: 'VMT-1',
    cambios: [
      { codigo: 'PRUEBA-1', estadoVal: 'Validado', validado: 'Sí', observacion: '' }
    ]
  });
  Logger.log(JSON.stringify(r));
}

function probarOgame() {
  var r = validarOgame({
    accion: 'validarOgame', codigo: 'PRUEBA-1',
    validadoOgame: 'Sí', registradoPor: 'OGAME-2'
  });
  Logger.log(JSON.stringify(r));
}
