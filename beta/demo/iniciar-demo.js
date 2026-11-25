/**
 * Arranque del MODO DEMO: arma una base de ejemplo en el simulador de Sheets con el
 * mismo backend real (backend/IntervencionesMTC_backend.gs) y expone DEMO_API(payload).
 * Requiere cargar antes: simulador-sheets.js y datos-demo.js, y haber definido window.BACKEND_DEMO
 * con las funciones del backend (ver cargarBackendDemo en la página), aisladas del resto de la página.
 */
(function (global) {
  'use strict';
  const CLAVE_DEMO = 'demo';
  const USUARIOS_DEMO = [
    ['admin', 'Administrador', '', 'Administrador de la plataforma (demo)'],
    ['direccion', 'Dirección', '', 'Dirección OMEI (demo)'],
    ['esp.omei', 'Especialista', '', 'Especialista OMEI (demo)'],
    ['esp.ogame', 'Especialista', '', 'Especialista OGAME (demo)']
  ];

  const B = () => global.BACKEND_DEMO;
  function escribirMatriz(nombre, filas) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    const ancho = filas.reduce((m, f) => Math.max(m, f.length), 0);
    h.clearContents();
    h.getRange(1, 1, filas.length, ancho).setValues(filas.map(f => f.concat(Array(ancho - f.length).fill(''))));
  }

  function construirBase() {
    SimuladorSheets.reiniciar();
    Object.entries(global.DEMO_CARGA).forEach(([n, filas]) => escribirMatriz(n, filas));
    B().prepararBase();
    const msg = B().importarCargaInicial();
    const unidades = B().leer(B().H.CAT_UNIDADES).map(u => u['Sigla']);
    const accesos = USUARIOS_DEMO.map(([u, r, un, n]) => ({ 'Usuario': u, 'Contraseña': CLAVE_DEMO, 'Rol': r, 'Unidad reportante': un, 'Nombre': n, 'Activo': 'SI' }))
      .concat(unidades.map(s => ({ 'Usuario': 'pf.' + s.toLowerCase(), 'Contraseña': CLAVE_DEMO, 'Rol': 'Punto focal', 'Unidad reportante': s, 'Nombre': 'Punto focal ' + s + ' (demo)', 'Activo': 'SI' })));
    B().escribir(B().H.ACCESOS, accesos);
    const t = B().ahora();
    const asign = B().leer(B().H.ASIGNACIONES).concat(
      ['PVN', 'DGPPT', 'PVD'].filter(s => unidades.includes(s)).map(s => ({ 'Ámbito': 'Carpeta Departamental', 'Clave': s, 'Especialista': 'esp.omei', 'Desde': t, 'Hasta': '', 'Asignado por': 'Demo' })));
    B().escribir(B().H.ASIGNACIONES, asign);
    B().agregar(B().H.CORTES, [{ 'Corte': '26/11/2026', 'Plazo': '01/12/2026', 'Estado': 'Abierto', 'Abierto por': 'direccion', 'Fecha de apertura': t, 'Financiera cargada': 'NO' }]);
    SimuladorSheets.guardar();
    return msg;
  }

  function asegurarBase() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName('Universo') || !ss.getSheetByName('Accesos')) return construirBase();
    return null;
  }

  // Ejecuta el código del backend en un ámbito propio (sus nombres no chocan con los de la página)
  global.cargarBackendDemo = src => {
    global.BACKEND_DEMO = new Function(src + '\n;return { procesar, prepararBase, importarCargaInicial, leer, escribir, agregar, ahora, H };')();
  };
  global.DEMO = {
    usuarios: () => B().leer(B().H.ACCESOS).map(a => ({ usuario: a['Usuario'], rol: a['Rol'], unidad: a['Unidad reportante'] })),
    reiniciar: construirBase,
    asegurar: asegurarBase,
    // Simula la red: copia profunda de ida y vuelta, como JSON por HTTP
    api: payload => {
      const r = B().procesar(JSON.parse(JSON.stringify(payload)));
      SimuladorSheets.guardar();
      return JSON.parse(JSON.stringify(r));
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
