/**
 * Asistencia en puesto – Frigorífico Bermejo S.A.
 * Aplicación web de Google Apps Script vinculada a una Hoja de cálculo.
 *
 * Hojas que usa:
 *   Padrón        Legajo | Apellido y nombre | Área | Sub área | Tarea | Turno (TM/TT, opcional) | Reporta a
 *                 (se lee por el título de cada columna, así que el orden puede cambiarse)
 *   Supervisores  Nombre | PIN | Activo (SI/NO) | También ve a (opcional, nombres separados por coma)
 *   Asistencia    la completa la aplicación (no editar a mano mientras se usa)
 *   Resumen       totales del día por área y sub área (elegir fecha en B1 y turno en E1)
 *                 K1: correos para el reporte (separados por coma); N1: hora del envío automático o No
 *
 * Primera vez: ejecutar la función configurar() desde el editor.
 */

const TZ = 'America/Argentina/Buenos_Aires';
const HHEE_MAX = 12;
const H_PADRON = 'Padrón';
const H_SUP = 'Supervisores';
const H_ASIS = 'Asistencia';
const COLS_ASIS = ['ID', 'Fecha', 'Turno', 'Legajo', 'Apellido y nombre', 'Área', 'Sub área',
  'Estado', 'HHEE (hs)', 'Observación', 'Supervisor', 'Hora registro', 'Última modificación', 'Tarea'];

/* ---------------- Página ---------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Asistencia en puesto')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------------- Preparar la hoja ---------------- */

function configurar() {
  const ss = SpreadsheetApp.getActive();
  const s = hoja_(ss, H_SUP, ['Nombre', 'PIN', 'Activo', 'También ve a']);
  s.getRange('B:B').setNumberFormat('@');
  if (s.getLastRow() < 2) s.getRange(2, 1, 1, 3).setValues([['Supervisor de ejemplo', '1234', 'SI']]);
  if (!s.getRange('D1').getValue()) s.getRange('D1').setValue('También ve a').setFontWeight('bold');

  const p = hoja_(ss, H_PADRON, ['Legajo', 'Apellido y nombre', 'Área', 'Sub área', 'Tarea', 'Turno', 'Reporta a']);
  let c = columnasPadron_(p);
  [['tar', 'Tarea'], ['tur', 'Turno'], ['rep', 'Reporta a']].forEach(([k, titulo]) => {
    if (c[k] < 0) {
      p.getRange(1, p.getLastColumn() + 1).setValue(titulo).setFontWeight('bold');
      c = columnasPadron_(p);
    }
  });
  p.getRange(2, c.rep + 1, Math.max(p.getMaxRows() - 1, 1), 1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(s.getRange('A2:A'), true).setAllowInvalid(true)
    .setHelpText('Elegí el supervisor tal como figura en la pestaña Supervisores.').build());

  const a = hoja_(ss, H_ASIS, COLS_ASIS);
  if (!a.getRange('N1').getValue()) a.getRange('N1').setValue('Tarea').setFontWeight('bold');
  a.getRange('A:A').setNumberFormat('@');
  a.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  a.getRange('I:I').setNumberFormat('0.0');
  a.getRange('L:L').setNumberFormat('@');
  a.getRange('M:M').setNumberFormat('dd/mm/yyyy hh:mm');
  a.hideColumns(1);

  hojaResumen_(ss);
  const activadores = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (activadores.indexOf('actualizarResumenAutomatico') < 0) {
    ScriptApp.newTrigger('actualizarResumenAutomatico').timeBased().everyMinutes(10).create();
  }
  if (activadores.indexOf('envioAutomatico') < 0) {
    ScriptApp.newTrigger('envioAutomatico').timeBased().everyHours(1).create();
  }
  prepararEnvio_(hojaResumen_(ss));
  actualizarResumen();
  ss.toast('Hojas listas. El resumen se actualiza solo cada 10 minutos.', 'Asistencia', 6);
}

function hoja_(ss, nombre, cab) {
  let h = ss.getSheetByName(nombre);
  if (!h) h = ss.insertSheet(nombre);
  if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

/* ---------------- Supervisores y PIN ---------------- */

function supervisores_() {
  const h = SpreadsheetApp.getActive().getSheetByName(H_SUP);
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 4).getDisplayValues()
    .filter(r => r[0].trim() && !/^no$/i.test(r[2].trim()))
    .map(r => ({ nombre: r[0].trim(), pin: r[1].trim(), tambien: r[3].trim() }));
}

function listarSupervisores() {
  return supervisores_().map(s => s.nombre).sort((a, b) => a.localeCompare(b, 'es'));
}

const DIAS_SESION = 60;

// Acepta la sesión (token) o, para celulares con la versión anterior, nombre y PIN.
// Devuelve el supervisor: { nombre, pin, tambien }.
function validar_(cred) {
  if (cred && cred.token) {
    const raw = PropertiesService.getScriptProperties().getProperty('ses_' + cred.token);
    if (!raw) throw new Error('SESION: Tu sesión venció. Ingresá de nuevo con tu PIN.');
    const ses = JSON.parse(raw);
    const s = supervisores_().find(x => x.nombre === ses.n);
    if (!s) throw new Error('SESION: Ese supervisor no está habilitado.');
    if (hash_([s.nombre, s.pin]) !== ses.h) throw new Error('SESION: Tu PIN cambió. Ingresá de nuevo.');
    if (Date.now() - ses.c > DIAS_SESION * 864e5) throw new Error('SESION: Tu sesión venció. Ingresá de nuevo con tu PIN.');
    return s;
  }
  return validarPin_(cred);
}

function validarPin_(cred) {
  if (!cred || !cred.nombre) throw new Error('SESION: Ingresá con tu nombre y PIN.');
  const cache = CacheService.getScriptCache();
  const k = 'fallos_' + Utilities.base64EncodeWebSafe(String(cred.nombre)).slice(0, 200);
  const fallos = Number(cache.get(k) || 0);
  if (fallos >= 5) throw new Error('SESION: Demasiados intentos con PIN incorrecto. Esperá 10 minutos.');
  const s = supervisores_().find(x => x.nombre === cred.nombre);
  if (!s || !s.pin || s.pin !== String(cred.pin || '').trim()) {
    cache.put(k, String(fallos + 1), 600);
    throw new Error(s ? 'SESION: PIN incorrecto.' : 'SESION: Ese supervisor no está habilitado.');
  }
  cache.remove(k);
  return s;
}

// Con el PIN correcto entrega un token de sesión: el celular guarda el token, nunca el PIN.
// Si RR. HH. cambia el PIN o desactiva al supervisor, la sesión deja de valer.
function ingresar(cred) {
  const s = validarPin_(cred);
  const props = PropertiesService.getScriptProperties();
  limpiarSesiones_(props);
  const token = Utilities.getUuid();
  props.setProperty('ses_' + token, JSON.stringify({ n: s.nombre, h: hash_([s.nombre, s.pin]), c: Date.now() }));
  return { nombre: s.nombre, token: token };
}

function salir(cred) {
  if (cred && cred.token) PropertiesService.getScriptProperties().deleteProperty('ses_' + cred.token);
  return true;
}

function limpiarSesiones_(props) {
  const todas = props.getProperties();
  Object.keys(todas).forEach(k => {
    if (k.indexOf('ses_') !== 0) return;
    try { if (Date.now() - JSON.parse(todas[k]).c > DIAS_SESION * 864e5) props.deleteProperty(k); }
    catch (e) { props.deleteProperty(k); }
  });
}

/* ---------------- Padrón ---------------- */

// Devuelve [legajo, nombre, área, sub área, turno, reporta a, tarea] leyendo cada dato por el título de su columna
function padron_() {
  const h = SpreadsheetApp.getActive().getSheetByName(H_PADRON);
  if (!h || h.getLastRow() < 2) return [];
  const c = columnasPadron_(h);
  if (c.leg < 0) return [];
  const g = (r, i) => (i >= 0 ? String(r[i] || '').trim() : '');
  return h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).getDisplayValues()
    .map(r => [g(r, c.leg).replace(/[.\s]/g, ''), g(r, c.nom), g(r, c.area), g(r, c.sub),
      turno_(g(r, c.tur)), g(r, c.rep), g(r, c.tar)])
    .filter(r => r[0]);
}

function columnasPadron_(h) {
  const n = Math.max(h.getLastColumn(), 1);
  const cab = h.getRange(1, 1, 1, n).getDisplayValues()[0].map(normTxt_);
  const f = re => cab.findIndex(x => re.test(x));
  return {
    leg: f(/^(legajo|leg\b|nro)/), nom: f(/nombre|apellido/), area: f(/^area/), sub: f(/^sub ?area/),
    tar: f(/^(tarea|puesto)/), tur: f(/^turno/), rep: f(/^reporta/)
  };
}

function normTxt_(x) {
  return String(x || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function turno_(t) {
  t = String(t).trim().toUpperCase();
  if (/^(TM|M|MAÑANA|MANANA)$/.test(t)) return 'TM';
  if (/^(TT|T|TARDE)$/.test(t)) return 'TT';
  return '';
}

function hash_(o) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(o), Utilities.Charset.UTF_8));
}

/* ---------------- Lectura del día ---------------- */

function datosDelDia(cred, fecha, hashPadron) {
  const sup = validar_(cred);
  const yo = sup.nombre;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error('Fecha no válida.');
  const pad = padron_();
  const hp = hash_(pad);
  const tambien = sup.tambien ? sup.tambien.split(/[,;]/).map(x => x.trim()).filter(Boolean) : [];
  return {
    yo: yo,
    cubre: [yo].concat(tambien),
    hoy: Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'),
    hashPadron: hp,
    padron: hp === hashPadron ? null : pad,
    registros: registrosDe_(fecha)
  };
}

function registrosDe_(fecha) {
  const h = SpreadsheetApp.getActive().getSheetByName(H_ASIS);
  if (!h || h.getLastRow() < 2) return [];
  const hallados = h.getRange(2, 1, h.getLastRow() - 1, 1)
    .createTextFinder(fecha + '|').matchCase(true).findAll();
  if (!hallados.length) return [];
  // Agrupa las filas en tramos cercanos, para no leer miles de filas de otros días
  const filas = hallados.map(r => r.getRow()).sort((a, b) => a - b);
  const tramos = [];
  filas.forEach(f => {
    const u = tramos[tramos.length - 1];
    if (u && f - u[1] <= 20) u[1] = f; else tramos.push([f, f]);
  });
  let vals = [];
  tramos.forEach(t => { vals = vals.concat(h.getRange(t[0], 1, t[1] - t[0] + 1, COLS_ASIS.length).getDisplayValues()); });
  return vals.filter(r => r[0].indexOf(fecha + '|') === 0).map(r => {
    const id = r[0].split('|');
    return {
      leg: id[2], turno: id[1], sup: r[10],
      e: r[7] === 'Ausente' ? 'A' : 'P',
      h: Number(String(r[8]).replace(',', '.')) || 0,
      o: r[9], t: r[11], n: r[4] === '(fuera del padrón)' ? '' : r[4],
      nom: r[4], area: r[5], sub: r[6]
    };
  });
}

/* ---------------- Guardado ---------------- */

function guardarLote(cred, ops) {
  const yo = validar_(cred).nombre;
  if (!Array.isArray(ops) || !ops.length) return { ok: [] };

  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const h = SpreadsheetApp.getActive().getSheetByName(H_ASIS);
    if (!h) throw new Error('Falta la hoja Asistencia. Ejecutá configurar().');
    const pmap = new Map(padron_().map(p => [p[0], p]));
    const supId = yo.replace(/\|/g, '/');
    const ok = [];
    const nuevas = [];
    const ahora = new Date();

    ops.slice(0, 50).forEach(op => {
      const valida = op && /^\d{4}-\d{2}-\d{2}$/.test(op.fecha) && /^(TM|TT)$/.test(op.turno) &&
        /^[A-Za-z0-9-]{1,20}$/.test(String(op.leg));
      if (!valida) { if (op && op.qid) ok.push(op.qid); return; }

      const id = [op.fecha, op.turno, op.leg, supId].join('|');
      const enLote = nuevas.findIndex(x => x[0] === id);
      const fila = enLote >= 0 ? 0 : buscarFila_(h, id);

      if (op.borrar) {
        if (enLote >= 0) nuevas.splice(enLote, 1);
        else if (fila) h.deleteRow(fila);
        ok.push(op.qid);
        return;
      }

      const p = pmap.get(String(op.leg));
      const f = op.fecha.split('-').map(Number);
      const hhee = op.e === 'P' ? Math.max(0, Math.min(HHEE_MAX, Number(op.h) || 0)) : 0;
      const valores = [
        id,
        new Date(Date.UTC(f[0], f[1] - 1, f[2], 15)),
        op.turno,
        /^\d+$/.test(op.leg) ? Number(op.leg) : op.leg,
        p ? p[1] : (textoSeguro_(op.n, 80) || '(fuera del padrón)'),
        p ? p[2] : '',
        p ? p[3] : '',
        op.e === 'A' ? 'Ausente' : 'Presente',
        hhee || '',
        textoSeguro_(op.o, 300),
        yo,
        op.t ? Utilities.formatDate(new Date(op.t), TZ, 'HH:mm') : '',
        ahora,
        p ? p[6] : ''
      ];

      if (enLote >= 0) nuevas[enLote] = valores;
      else if (fila) h.getRange(fila, 1, 1, valores.length).setValues([valores]);
      else nuevas.push(valores);
      ok.push(op.qid);
    });

    if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, COLS_ASIS.length).setValues(nuevas);
    return { ok: ok };
  } finally {
    lock.releaseLock();
  }
}

// Un texto que empieza con = + - @ se guarda como texto, no como fórmula
function textoSeguro_(s, max) {
  const txt = String(s || '').trim().slice(0, max);
  return /^[=+\-@]/.test(txt) ? "'" + txt : txt;
}

function buscarFila_(h, id) {
  if (h.getLastRow() < 2) return 0;
  const c = h.getRange(2, 1, h.getLastRow() - 1, 1)
    .createTextFinder(id).matchEntireCell(true).matchCase(true).findNext();
  return c ? c.getRow() : 0;
}

/* ---------------- Resumen diario ---------------- */

const H_RES = 'Resumen';
const FUERA = '(fuera del padrón)';
const SIN_SUP = 'Sin supervisor asignado';

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Asistencia')
    .addItem('Ver resumen de hoy', 'resumenDeHoy')
    .addItem('Actualizar resumen', 'actualizarResumen')
    .addSeparator()
    .addItem('Enviar reporte por mail', 'enviarReportePorMail')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() !== H_RES) return;
  const a1 = e.range.getA1Notation();
  if (a1 === 'B1' || a1 === 'E1') actualizarResumen();
}

function actualizarResumenAutomatico() {
  const hora = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  if (hora < 4) return; // de madrugada no hace falta
  actualizarResumen();
}

function resumenDeHoy() {
  const ss = SpreadsheetApp.getActive();
  const h = hojaResumen_(ss);
  h.getRange('B1').setValue(fechaComoCelda_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')));
  actualizarResumen();
  ss.setActiveSheet(h);
}

function fechaComoCelda_(iso) {
  const f = iso.split('-').map(Number);
  return new Date(Date.UTC(f[0], f[1] - 1, f[2], 15));
}

function hojaResumen_(ss) {
  let h = ss.getSheetByName(H_RES);
  if (h) return h;
  h = ss.insertSheet(H_RES, 0);
  h.getRange('A1').setValue('Fecha');
  h.getRange('B1').setValue(fechaComoCelda_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd')))
    .setNumberFormat('dd/mm/yyyy')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  h.getRange('D1').setValue('Turno');
  h.getRange('E1').setValue('Ambos')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Ambos', 'TM', 'TT'], true).setAllowInvalid(false).build());
  h.getRange('G1').setValue('Actualizado');
  h.getRange('A1:H1').setFontWeight('bold');
  h.getRange('B1').setBackground('#FCEDD2');
  h.getRange('E1').setBackground('#FCEDD2');
  h.setFrozenRows(1);
  [110, 230, 150, 150, 90, 230, 160, 120].forEach((w, i) => h.setColumnWidth(i + 1, w));
  return h;
}

// Evita que el activador, un cambio de fecha y el envío por mail escriban el Resumen al mismo tiempo
function actualizarResumen() {
  let lock = null;
  try { lock = LockService.getDocumentLock(); } catch (e) { lock = null; }
  if (lock && !lock.tryLock(20000)) return null;
  try { return actualizarResumen_(); } finally { if (lock) lock.releaseLock(); }
}

function actualizarResumen_() {
  const ss = SpreadsheetApp.getActive();
  const h = hojaResumen_(ss);

  let fv = h.getRange('B1').getValue();
  if (!(fv instanceof Date)) {
    fv = fechaComoCelda_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
    h.getRange('B1').setValue(fv);
  }
  const fecha = Utilities.formatDate(fv, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  const tSel = String(h.getRange('E1').getValue());
  const turno = /^(TM|TT)$/.test(tSel) ? tSel : 'Ambos';

  const pad = padron_();
  const pmap = new Map(pad.map(p => [p[0], p]));
  const regs = registrosDe_(fecha).filter(r => turno === 'Ambos' || r.turno === turno);

  // Un colaborador cuenta una sola vez aunque lo hayan registrado dos supervisores
  const porLeg = new Map();
  regs.forEach(r => { if (!porLeg.has(r.leg)) porLeg.set(r.leg, []); porLeg.get(r.leg).push(r); });

  const unidades = [], ausentes = [], hhee = [], dobles = [];
  porLeg.forEach((lista, leg) => {
    const p = pmap.get(leg);
    const nom = p ? p[1] : (lista[0].nom || FUERA);
    const area = p ? (p[2] || 'Sin área') : FUERA;
    const sub = p ? p[3] : '';
    const tarea = p ? p[6] : '';
    const presente = lista.some(r => r.e === 'P');
    const porTurno = {};
    lista.forEach(r => { porTurno[r.turno] = Math.max(porTurno[r.turno] || 0, r.e === 'P' ? (r.h || 0) : 0); });
    const turnos = Object.keys(porTurno).sort();
    const hs = turnos.reduce((s, k) => s + porTurno[k], 0);
    const rep = p ? (p[5] || SIN_SUP) : FUERA;
    unidades.push({ area: area, sub: sub, rep: rep, e: presente ? 'P' : 'A', h: presente ? hs : 0 });
    if (!presente) ausentes.push([num_(leg), nom, area, sub, turnos.join(' y '), unir_(lista.map(r => r.o)), unir_(lista.map(r => r.sup)), tarea]);
    lista.filter(r => r.e === 'P' && r.h > 0).forEach(r => hhee.push([num_(leg), nom, area, sub, r.turno, r.o, r.sup, r.h]));
    turnos.forEach(t => {
      const del = lista.filter(r => r.turno === t);
      if (del.length > 1) {
        dobles.push([num_(leg), nom, area, sub, t,
          del.map(r => r.e === 'P' ? 'Presente' : 'Ausente').join(' / '), del.map(r => r.sup).join(' / ')]);
      }
    });
    // Presente en un turno y ausente en el otro: cuenta como presente, pero se avisa
    const estados = new Set(lista.map(r => r.e));
    if (turnos.length > 1 && estados.size > 1) {
      dobles.push([num_(leg), nom, area, sub, turnos.join(' y '),
        lista.map(r => r.turno + ': ' + (r.e === 'P' ? 'Presente' : 'Ausente')).join(' / '), unir_(lista.map(r => r.sup))]);
    }
  });

  const sinReg = pad.filter(p => !porLeg.has(p[0]) && (turno === 'Ambos' || !p[4] || p[4] === turno));
  sinReg.forEach(p => unidades.push({ area: p[2] || 'Sin área', sub: p[3], rep: p[5] || SIN_SUP, e: 'S', h: 0 }));

  // Por supervisor (columna Reporta a del padrón)
  const porSup = new Map();
  unidades.forEach(u => { if (!porSup.has(u.rep)) porSup.set(u.rep, []); porSup.get(u.rep).push(u); });
  const filasSup = ordenar_(Array.from(porSup.keys())).map(k => [k].concat(filaConteo_(cuenta_(porSup.get(k)))));

  // Por área y sub área
  const porArea = new Map();
  unidades.forEach(u => {
    if (!porArea.has(u.area)) porArea.set(u.area, new Map());
    const m = porArea.get(u.area);
    if (!m.has(u.sub || '')) m.set(u.sub || '', []);
    m.get(u.sub || '').push(u);
  });
  const filasArea = [], negritas = [];
  ordenar_(Array.from(porArea.keys())).forEach(a => {
    const subs = porArea.get(a);
    const todas = [].concat.apply([], Array.from(subs.values()));
    negritas.push(filasArea.length);
    filasArea.push([a, 'Total área'].concat(filaConteo_(cuenta_(todas))));
    const claves = ordenar_(Array.from(subs.keys()));
    if (!(claves.length === 1 && claves[0] === '')) {
      claves.forEach(s => filasArea.push(['', s || 'Sin sub área'].concat(filaConteo_(cuenta_(subs.get(s))))));
    }
  });

  const porAreaNombre = (x, y) => cmpArea_(x[2], y[2]) || String(x[1]).localeCompare(String(y[1]), 'es');
  ausentes.sort(porAreaNombre); hhee.sort(porAreaNombre); dobles.sort(porAreaNombre);
  const filasSin = sinReg.map(p => [num_(p[0]), p[1], p[2] || 'Sin área', p[3], p[4], p[5], p[6]]).sort(porAreaNombre);

  const t = cuenta_(unidades);
  const totHs = hhee.length ? unidades.reduce((s, u) => s + u.h, 0) : 0;
  const titFecha = Utilities.formatDate(fv, ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy') +
    (turno === 'Ambos' ? ', ambos turnos' : ', turno ' + turno);

  // Escribir
  h.getRange(2, 1, Math.max(h.getMaxRows() - 1, 1), h.getMaxColumns()).clear();
  let r = 3;
  r = bloque_(h, r, 'Totales del ' + titFecha,
    ['Dotación', 'Presentes', 'Ausentes', 'Sin registrar', '% presentismo', '% ausentismo', 'HHEE (hs)', 'Dobles registros'],
    [[t.d, t.p, t.a, t.s, t.d ? t.p / t.d : '', t.d ? t.a / t.d : '', totHs, dobles.length]],
    { pct: [5, 6], dec: [7] });
  r = bloque_(h, r, 'Por área y sub área',
    ['Área', 'Sub área', 'Dotación', 'Presentes', 'Ausentes', 'Sin registrar', '% presentismo', 'HHEE (hs)'],
    filasArea, { pct: [7], dec: [8], negritas: negritas });
  r = bloque_(h, r, 'Por supervisor',
    ['Supervisor', 'Dotación', 'Presentes', 'Ausentes', 'Sin registrar', '% presentismo', 'HHEE (hs)'],
    filasSup, { pct: [6], dec: [7] });
  r = bloque_(h, r, 'Ausentes (' + ausentes.length + ')',
    ['Legajo', 'Apellido y nombre', 'Área', 'Sub área', 'Turno', 'Observación', 'Supervisor', 'Tarea'], ausentes, {});
  r = bloque_(h, r, 'Horas extra (' + String(Math.round(totHs * 10) / 10).replace('.', ',') + ' hs)',
    ['Legajo', 'Apellido y nombre', 'Área', 'Sub área', 'Turno', 'Observación', 'Supervisor', 'HHEE (hs)'], hhee, { dec: [8] });
  r = bloque_(h, r, 'Sin registrar (' + filasSin.length + ')',
    ['Legajo', 'Apellido y nombre', 'Área', 'Sub área', 'Turno en padrón', 'Reporta a', 'Tarea'], filasSin, {});
  if (dobles.length) {
    r = bloque_(h, r, 'Dobles registros o contradictorios (' + dobles.length + ')',
      ['Legajo', 'Apellido y nombre', 'Área', 'Sub área', 'Turno', 'Estados', 'Supervisores'], dobles, {});
  }
  h.getRange('H1').setValue(new Date()).setNumberFormat('dd/mm/yyyy hh:mm');
  return { titFecha: titFecha, t: t, totHs: totHs, dobles: dobles, filasArea: filasArea, negritas: negritas,
    filasSup: filasSup, ausentes: ausentes, hhee: hhee, filasSin: filasSin };
}

function cuenta_(us) {
  const c = { d: us.length, p: 0, a: 0, s: 0, h: 0 };
  us.forEach(u => { if (u.e === 'P') { c.p++; c.h += u.h; } else if (u.e === 'A') c.a++; else c.s++; });
  return c;
}
function filaConteo_(c) { return [c.d, c.p, c.a, c.s, c.d ? c.p / c.d : '', c.h]; }
function num_(leg) { return /^\d+$/.test(leg) ? Number(leg) : leg; }
function unir_(xs) { return Array.from(new Set(xs.map(x => String(x || '').trim()).filter(Boolean))).join(' / '); }
function cmpArea_(a, b) {
  const peso = x => x === FUERA ? 2 : (x === 'Sin área' || x === SIN_SUP ? 1 : 0);
  return peso(a) - peso(b) || String(a).localeCompare(String(b), 'es');
}
function ordenar_(xs) { return xs.slice().sort(cmpArea_); }

function bloque_(h, r, titulo, cab, filas, o) {
  h.getRange(r, 1).setValue(titulo).setFontWeight('bold').setFontSize(12);
  r++;
  h.getRange(r, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#2F4858').setFontColor('#FFFFFF');
  r++;
  if (!filas.length) {
    h.getRange(r, 1).setValue('Sin casos.').setFontColor('#5A6873');
    return r + 2;
  }
  const datos = filas.map(f => { const x = f.slice(0, cab.length); while (x.length < cab.length) x.push(''); return x; });
  h.getRange(r, 1, datos.length, cab.length).setValues(datos);
  (o.pct || []).forEach(c => h.getRange(r, c, datos.length, 1).setNumberFormat('0.0%'));
  (o.dec || []).forEach(c => h.getRange(r, c, datos.length, 1).setNumberFormat('0.0'));
  (o.negritas || []).forEach(i => h.getRange(r + i, 1, 1, cab.length).setFontWeight('bold').setBackground('#EDF1F2'));
  return r + datos.length + 2;
}

/* ---------------- Reporte por mail ---------------- */

const HORAS_ENVIO = ['No'].concat(Array.from({ length: 19 }, (_, i) => String(i + 5).padStart(2, '0') + ' hs'));

// Celdas de configuración en la fila 1 de Resumen (actualizarResumen nunca las borra)
function prepararEnvio_(h) {
  if (!h.getRange('J1').getValue()) h.getRange('J1').setValue('Enviar reporte a').setFontWeight('bold');
  if (!h.getRange('M1').getValue()) h.getRange('M1').setValue('Envío automático').setFontWeight('bold');
  h.getRange('K1').setBackground('#FCEDD2').setNote('Uno o más correos, separados por coma.');
  const n1 = h.getRange('N1');
  if (!n1.getValue()) n1.setValue('No');
  n1.setBackground('#FCEDD2').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(HORAS_ENVIO, true).setAllowInvalid(false).build())
    .setNote('Hora a la que sale el reporte del día, todos los días. Llega dentro de esa hora (por ejemplo, entre 15:00 y 15:59).');
  h.setColumnWidth(10, 130); h.setColumnWidth(11, 280); h.setColumnWidth(13, 130); h.setColumnWidth(14, 90);
}

function destinatarios_(h) {
  return String(h.getRange('K1').getValue()).split(/[,;\s]+/).map(x => x.trim())
    .filter(x => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(x));
}

// Desde el menú
function enviarReportePorMail() {
  const ui = SpreadsheetApp.getUi();
  const h = hojaResumen_(SpreadsheetApp.getActive());
  prepararEnvio_(h);
  if (!destinatarios_(h).length) {
    ui.alert('Escribí al menos un correo en la celda K1 de la hoja Resumen.');
    return;
  }
  try {
    const n = enviarReporte_();
    SpreadsheetApp.getActive().toast('Reporte enviado a ' + n + ' destinatario(s).', 'Asistencia', 6);
  } catch (err) {
    ui.alert('No se pudo enviar el reporte: ' + err.message);
  }
}

// Activador de cada hora: envía una sola vez por día, en la hora elegida en N1
function envioAutomatico() {
  const h = SpreadsheetApp.getActive().getSheetByName(H_RES);
  if (!h) return;
  const m = String(h.getRange('N1').getValue()).match(/^(\d{1,2})/);
  if (!m) return;
  const ahora = new Date();
  if (Number(Utilities.formatDate(ahora, TZ, 'H')) !== Number(m[1])) return;
  const hoy = Utilities.formatDate(ahora, TZ, 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ultimoEnvio') === hoy) return;
  if (!destinatarios_(h).length) return;
  h.getRange('B1').setValue(fechaComoCelda_(hoy));
  enviarReporte_();
  props.setProperty('ultimoEnvio', hoy);
}

function enviarReporte_() {
  const ss = SpreadsheetApp.getActive();
  const h = hojaResumen_(ss);
  const dest = destinatarios_(h);
  const d = actualizarResumen();
  if (!d) throw new Error('El resumen se está actualizando. Probá de nuevo en un minuto.');
  const t = d.t;
  const pct = x => (x === '' || x == null) ? '' : String(Math.round(x * 1000) / 10).replace('.', ',') + ' %';
  const hs = x => String(Math.round((Number(x) || 0) * 10) / 10).replace('.', ',');

  const caja = (n, txt, color) => '<td style="padding:10px 14px;border:1px solid #D2DADE;text-align:left">' +
    '<div style="font-size:24px;font-weight:bold;color:' + color + '">' + n + '</div>' +
    '<div style="font-size:12px;color:#5A6873">' + txt + '</div></td>';

  let html = '<div style="font-family:Arial,Helvetica,sans-serif;color:#14202A;max-width:760px">' +
    '<h2 style="margin:0 0 4px;color:#2F4858">Asistencia en puesto</h2>' +
    '<p style="margin:0 0 16px;color:#5A6873">' + escHtml_(d.titFecha) + '</p>' +
    '<table style="border-collapse:collapse;margin-bottom:8px"><tr>' +
    caja(t.p, 'Presentes', '#14724A') + caja(t.a, 'Ausentes', '#B42318') +
    caja(t.s, 'Sin registrar', '#14202A') + caja(hs(d.totHs), 'HHEE (hs)', '#8F5400') +
    '</tr></table>' +
    '<p style="margin:0 0 20px;font-size:13px;color:#5A6873">Dotación ' + t.d + '. Presentismo ' + pct(t.d ? t.p / t.d : '') +
    ', ausentismo ' + pct(t.d ? t.a / t.d : '') + (d.dobles.length ? '. Dobles registros: ' + d.dobles.length : '') + '.</p>';

  html += tablaHtml_('Por área y sub área',
    ['Área', 'Sub área', 'Dotación', 'Presentes', 'Ausentes', 'Sin registrar', '% presentismo', 'HHEE'],
    d.filasArea.map(f => [f[0], f[1], f[2], f[3], f[4], f[5], pct(f[6]), hs(f[7])]), d.negritas);
  html += tablaHtml_('Por supervisor',
    ['Supervisor', 'Dotación', 'Presentes', 'Ausentes', 'Sin registrar', '% presentismo', 'HHEE'],
    d.filasSup.map(f => [f[0], f[1], f[2], f[3], f[4], pct(f[5]), hs(f[6])]));
  html += tablaHtml_('Ausentes (' + d.ausentes.length + ')',
    ['Legajo', 'Apellido y nombre', 'Área', 'Tarea', 'Observación', 'Registró'],
    d.ausentes.map(f => [f[0], f[1], [f[2], f[3]].filter(Boolean).join(' / '), f[7], f[5], f[6]]));
  html += tablaHtml_('Horas extra (' + hs(d.totHs) + ' hs)',
    ['Legajo', 'Apellido y nombre', 'Área', 'Turno', 'HHEE', 'Observación', 'Registró'],
    d.hhee.map(f => [f[0], f[1], [f[2], f[3]].filter(Boolean).join(' / '), f[4], hs(f[7]), f[5], f[6]]));
  html += tablaHtml_('Sin registrar (' + d.filasSin.length + ')',
    ['Legajo', 'Apellido y nombre', 'Área', 'Tarea', 'Reporta a'],
    d.filasSin.map(f => [f[0], f[1], [f[2], f[3]].filter(Boolean).join(' / '), f[6], f[5]]));

  html += '<p style="margin:24px 0"><a href="' + ss.getUrl() + '" style="background:#2F4858;color:#ffffff;padding:11px 18px;' +
    'text-decoration:none;border-radius:6px;font-weight:bold">Abrir la planilla</a></p>' +
    '<p style="font-size:12px;color:#5A6873">Mensaje automático del sistema de Asistencia en puesto de Frigorífico Bermejo. ' +
    'Generado a las ' + Utilities.formatDate(new Date(), TZ, 'HH:mm') + ' hs.</p></div>';

  MailApp.sendEmail({ to: dest.join(','), subject: 'Asistencia FBSA: ' + d.titFecha, htmlBody: html, name: 'Asistencia FBSA' });
  return dest.length;
}

function tablaHtml_(titulo, cab, filas, negritas) {
  const MAX = 100;
  const neg = new Set(negritas || []);
  const td = 'padding:6px 8px;border-bottom:1px solid #D2DADE;font-size:13px;vertical-align:top';
  let out = '<h3 style="margin:22px 0 8px;font-size:16px;color:#2F4858">' + escHtml_(titulo) + '</h3>';
  if (!filas.length) return out + '<p style="font-size:13px;color:#5A6873;margin:0">Sin casos.</p>';
  out += '<table style="border-collapse:collapse;width:100%"><tr>' +
    cab.map(c => '<th style="' + td + ';background:#2F4858;color:#ffffff;text-align:left">' + escHtml_(c) + '</th>').join('') + '</tr>';
  filas.slice(0, MAX).forEach((f, i) => {
    const extra = neg.has(i) ? ';font-weight:bold;background:#EDF1F2' : '';
    out += '<tr>' + f.map(v => '<td style="' + td + extra + '">' + escHtml_(v) + '</td>').join('') + '</tr>';
  });
  out += '</table>';
  if (filas.length > MAX) out += '<p style="font-size:12px;color:#5A6873">Y ' + (filas.length - MAX) + ' más en la planilla.</p>';
  return out;
}

function escHtml_(v) {
  return String(v == null ? '' : v).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}
