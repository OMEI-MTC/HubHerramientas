/**
 * Simulador mínimo de Google Sheets / Apps Script para el MODO DEMO de
 * registro-intervenciones.html (?demo=1) y para las pruebas automáticas.
 * Implementa solo lo que usa backend/IntervencionesMTC_backend.gs.
 * Los datos se guardan en el navegador (localStorage); nada sale del equipo.
 */
(function (global) {
  'use strict';
  const CLAVE_LS = 'intmtc_demo_libro_v1';

  function crearHoja(nombre, datos) {
    const h = { _nombre: nombre, _d: datos || [] };
    const ancho = () => h._d.reduce((m, f) => Math.max(m, f.length), 0);
    h.getName = () => h._nombre;
    h.getLastRow = () => { for (let i = h._d.length - 1; i >= 0; i--) if (h._d[i].some(c => c !== '' && c != null)) return i + 1; return 0; };
    h.getLastColumn = () => ancho();
    h.clearContents = () => { h._d = []; };
    h.setFrozenRows = () => h;
    h.getDataRange = () => rango(h, 1, 1, Math.max(h.getLastRow(), 1), Math.max(ancho(), 1));
    h.getRange = (r, c, nr, nc) => rango(h, r, c, nr || 1, nc || 1);
    return h;
  }
  function rango(h, r, c, nr, nc) {
    return {
      getValues() {
        const out = [];
        for (let i = 0; i < nr; i++) { const f = h._d[r - 1 + i] || []; const row = []; for (let j = 0; j < nc; j++) { const v = f[c - 1 + j]; row.push(v === undefined || v === null ? '' : v); } out.push(row); }
        return out;
      },
      setValues(vals) {
        if (vals.length !== nr || vals.some(v => v.length !== nc)) throw new Error('Las dimensiones no coinciden con el rango');
        for (let i = 0; i < nr; i++) { const k = r - 1 + i; while (h._d.length <= k) h._d.push([]); for (let j = 0; j < nc; j++) h._d[k][c - 1 + j] = vals[i][j]; }
        return this;
      },
      setNumberFormat() { return this; }
    };
  }

  function crearLibro(estado) {
    const hojas = [];
    const ss = {
      getSheetByName: n => hojas.find(h => h._nombre === n) || null,
      getSheets: () => hojas.slice(),
      insertSheet: n => { if (ss.getSheetByName(n)) throw new Error('Ya existe la hoja ' + n); const h = crearHoja(n); hojas.push(h); return h; },
      _exportar: () => hojas.map(h => [h._nombre, h._d]),
    };
    (estado || []).forEach(([n, d]) => hojas.push(crearHoja(n, d)));
    return ss;
  }

  let LIBRO = null;
  const almacen = global.localStorage || null;
  function cargarLibro() {
    if (LIBRO) return LIBRO;
    let estado = null;
    try { estado = almacen ? JSON.parse(almacen.getItem(CLAVE_LS) || 'null') : null; } catch (e) { estado = null; }
    LIBRO = crearLibro(estado);
    return LIBRO;
  }
  function guardarLibro() { if (almacen && LIBRO) try { almacen.setItem(CLAVE_LS, JSON.stringify(LIBRO._exportar())); } catch (e) { console.warn('No se pudo guardar el demo', e); } }
  function reiniciar() { if (almacen) almacen.removeItem(CLAVE_LS); LIBRO = crearLibro(null); }

  // Caché de sesiones: en el navegador se guarda en sessionStorage para sobrevivir a una recarga
  const ses = global.sessionStorage || null;
  const cache = {
    get: k => { if (ses) { const v = ses.getItem('intmtc_cache_' + k); return v === null ? undefined : { v }; } return mapa.get(k); },
    set: (k, e) => { if (ses) ses.setItem('intmtc_cache_' + k, e.v); else mapa.set(k, e); },
    delete: k => { if (ses) ses.removeItem('intmtc_cache_' + k); else mapa.delete(k); }
  };
  const mapa = new Map();
  const pad = n => String(n).padStart(2, '0');

  global.SpreadsheetApp = {
    getActiveSpreadsheet: cargarLibro,
    openById: cargarLibro,
    getUi: () => ({ createMenu: () => ({ addItem() { return this; }, addToUi() {} }) })
  };
  global.CacheService = { getScriptCache: () => ({
    get: k => { const e = cache.get(k); return e && e.v !== undefined ? e.v : null; },
    put: (k, v) => cache.set(k, { v: String(v) }),
    remove: k => cache.delete(k)
  }) };
  global.LockService = { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) };
  global.Utilities = {
    getUuid: () => (global.crypto && global.crypto.randomUUID) ? global.crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx'.replace(/x/g, () => (Math.random() * 16 | 0).toString(16)) + '-' + Date.now().toString(16),
    formatDate: (d, tz, fmt) => fmt.replace('yyyy', d.getFullYear()).replace('MM', pad(d.getMonth() + 1)).replace('dd', pad(d.getDate()))
      .replace('HH', pad(d.getHours())).replace('mm', pad(d.getMinutes()))
  };
  global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: t => ({ setMimeType: () => ({ contenido: t }) }) };
  global.console = global.console || { log() {}, error() {} };

  global.SimuladorSheets = { guardar: guardarLibro, reiniciar, libro: cargarLibro, crearLibro, _usar: l => { LIBRO = l; } };
})(typeof window !== 'undefined' ? window : globalThis);
