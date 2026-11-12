/**
 * Mis Finanzas — API sobre Google Sheets
 *
 * Esta hoja de cálculo es la base de datos del formulario.
 * Pega este código en Extensiones → Apps Script de tu hoja,
 * cambia la CLAVE y publícalo como "Aplicación web".
 *
 * Versión 2: envío en lote, edición en el mismo lugar y orden sin reescribir filas.
 * Versión 3: cuentas (cada movimiento pertenece a una cuenta; saldo por cuenta).
 */

const VERSION = 3;

// Clave secreta: el formulario la envía en cada pedido. Cámbiala por una tuya.
const CLAVE = 'CAMBIA_ESTA_CLAVE';

const HOJA_MOV = 'Movimientos';
const HOJA_PRES = 'Presupuesto';
const HOJA_RES = 'Resumen mensual';
const HOJA_CAT = 'Gastos por categoría';
const HOJA_CFG = 'Config';
const HOJA_PRE = 'Préstamos';
const HOJA_CTA = 'Cuentas';

// Categorías especiales: no son ingreso ni gasto real
const CAT_TRANSF = 'Transferencia entre mis cuentas';
const CAT_AJUSTE = 'Ajuste de saldo';

const CAT_GASTO = ['Alimentación', 'Transporte', 'Vivienda / Alquiler', 'Servicios (luz, agua, internet)', 'Salud', 'Educación', 'Entretenimiento', 'Ropa y cuidado personal', 'Deudas / Tarjetas', 'Ahorro e inversión', 'Otros gastos'];

// Columnas de Movimientos (1 = A)
const COL = { id: 1, fecha: 2, mes: 3, tipo: 4, categoria: 5, nota: 6, metodo: 7, moneda: 8, monto: 9, tc: 10, montoPEN: 11, creado: 12, prestamo: 13, cuenta: 14, destino: 15 };
const ENCABEZADOS = ['ID', 'Fecha', 'Mes', 'Tipo', 'Categoría', 'Descripción', 'Medio de pago', 'Moneda', 'Monto', 'Tipo de cambio', 'Monto S/', 'Registrado', 'Préstamo', 'Cuenta', 'Cuenta destino'];

// Columnas de Cuentas (1 = A)
const CCOL = { id: 1, nombre: 2, tipo: 3, archivada: 4, creado: 5, saldo: 6 };
const C_ENCABEZADOS = ['ID', 'Nombre', 'Tipo', 'Archivada', 'Registrado', 'Saldo actual S/'];
const TIPOS_CUENTA = ['Billetera digital', 'Banco', 'Efectivo', 'Tarjeta de crédito', 'App de préstamos', 'Otro'];

// Columnas de Préstamos (1 = A)
const PCOL = { id: 1, entidad: 2, nota: 3, fecha: 4, monto: 5, cuotas: 6, cuota: 7, frecuencia: 8, primer: 9, creado: 16 };
const P_ENCABEZADOS = ['ID', 'Entidad', 'Descripción', 'Fecha de desembolso', 'Monto recibido S/', 'N° de cuotas', 'Monto de cuota S/', 'Frecuencia', 'Primer vencimiento', 'Total a pagar S/', 'Interés y costos S/', 'Pagado S/', 'Saldo pendiente S/', 'Cuotas pagadas', 'Estado', 'Registrado'];

/* ---------- Puntos de entrada ---------- */

function doGet(e) {
  return responder_(function () {
    verificarClave_(e.parameter.clave);
    return leerTodo_();
  });
}

function doPost(e) {
  return responder_(function () {
    const body = JSON.parse(e.postData.contents || '{}');
    verificarClave_(body.clave);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      // "lote": varios pasos en una sola llamada, en orden; se responde una sola vez.
      const pasos = body.accion === 'lote' ? (body.pasos || []) : [body];
      pasos.forEach(ejecutar_);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return leerTodo_();
  });
}

function ejecutar_(b) {
  switch (b.accion) {
    case 'agregar': return agregar_(b.mov);
    case 'editar': return editar_(b.mov);
    case 'borrar': return borrar_(b.id);
    case 'reordenar': return reordenar_(b.items);
    case 'config': return guardarConfig_(b.config);
    case 'prestamo_agregar': return agregarPrestamo_(b.prestamo);
    case 'prestamo_borrar': return borrarPrestamo_(b.id);
    case 'cuenta_agregar': return agregarCuenta_(b.cuenta);
    case 'cuenta_editar': return editarCuenta_(b.cuenta);
    case 'cuenta_borrar': return borrarCuenta_(b.id);
    default: throw new Error('Acción desconocida');
  }
}

/** Ejecútalo una vez desde el editor para crear las pestañas y dar permisos. */
function configurar() {
  configurar_();
  SpreadsheetApp.getActive().toast('Hoja lista. Ahora implementa la aplicación web.', 'Mis Finanzas', 8);
}

/* ---------- Lógica ---------- */

function responder_(fn) {
  let out;
  try {
    configurarSiHaceFalta_();
    try {
      out = fn();
    } catch (err) {
      // Si alguien borró o renombró una pestaña, se vuelve a crear y se reintenta una vez.
      if (!/null|undefined/.test(String(err && err.message))) throw err;
      PropertiesService.getScriptProperties().deleteProperty('hoja_lista');
      configurarSiHaceFalta_();
      out = fn();
    }
    out = Object.assign({ ok: true, version: VERSION }, out);
  } catch (err) {
    out = { ok: false, version: VERSION, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Revisa y crea las pestañas solo la primera vez (no en cada llamada). */
function configurarSiHaceFalta_() {
  const props = PropertiesService.getScriptProperties();
  const marca = ss_().getId() + ':' + VERSION;
  if (props.getProperty('hoja_lista') === marca) return;
  configurar_();
  props.setProperty('hoja_lista', marca);
}

function verificarClave_(c) {
  if (!c || c !== CLAVE) throw new Error('Clave incorrecta');
}

function ss_() { return SpreadsheetApp.getActive(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }

function configurar_() {
  const ss = ss_();
  let mov = ss.getSheetByName(HOJA_MOV);
  if (!mov) {
    const vacia = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0 ? ss.getSheets()[0] : null;
    mov = vacia ? vacia.setName(HOJA_MOV) : ss.insertSheet(HOJA_MOV, 0);
    mov.getRange(1, 1, 1, ENCABEZADOS.length).setValues([ENCABEZADOS]).setFontWeight('bold').setBackground('#e8edf7');
    // Columnas calculadas para TODAS las filas, también las que escribas a mano
    mov.getRange('C1').setFormula('={"Mes";ARRAYFORMULA(IF(B2:B="","",TEXT(B2:B,"yyyy-mm")))}');
    mov.getRange('K1').setFormula('={"Monto S/";ARRAYFORMULA(IF(I2:I="","",IF(H2:H="USD",I2:I*J2:J,I2:I)))}');
    mov.getRange('B2:B').setNumberFormat('dd/mm/yyyy');
    mov.getRange('I2:I').setNumberFormat('#,##0.00');
    mov.getRange('J2:J').setNumberFormat('0.000');
    mov.getRange('K2:K').setNumberFormat('#,##0.00');
    mov.setFrozenRows(1);
    mov.setColumnWidths(1, 12, 110);
    mov.setColumnWidth(5, 220);
    mov.setColumnWidth(6, 240);
    const reglas = [
      SpreadsheetApp.newDataValidation().requireValueInList(['Gasto', 'Ingreso'], true).build(),
      SpreadsheetApp.newDataValidation().requireValueInList(['PEN', 'USD'], true).build()
    ];
    mov.getRange('D2:D').setDataValidation(reglas[0]);
    mov.getRange('H2:H').setDataValidation(reglas[1]);
  }
  // Versiones anteriores de la hoja no tenían las columnas Préstamo, Cuenta y Cuenta destino
  [COL.prestamo, COL.cuenta, COL.destino].forEach(function (c) {
    if (mov.getRange(1, c).getValue() === '') mov.getRange(1, c).setValue(ENCABEZADOS[c - 1]).setFontWeight('bold').setBackground('#e8edf7');
  });
  if (!ss.getSheetByName(HOJA_CTA)) {
    const ct = ss.insertSheet(HOJA_CTA, 1);
    ct.getRange(1, 1, 1, C_ENCABEZADOS.length).setValues([C_ENCABEZADOS]).setFontWeight('bold').setBackground('#e8edf7');
    ct.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(TIPOS_CUENTA, true).build());
    ct.getRange('F2:F').setNumberFormat('#,##0.00');
    ct.setFrozenRows(1);
    ct.setColumnWidths(1, 6, 140);
  }
  if (!ss.getSheetByName(HOJA_PRE)) {
    const pr = ss.insertSheet(HOJA_PRE);
    pr.getRange(1, 1, 1, P_ENCABEZADOS.length).setValues([P_ENCABEZADOS]).setFontWeight('bold').setBackground('#e8edf7');
    pr.getRange('D2:D').setNumberFormat('dd/mm/yyyy');
    pr.getRange('I2:I').setNumberFormat('dd/mm/yyyy');
    pr.getRange('E2:E').setNumberFormat('#,##0.00');
    pr.getRange('G2:G').setNumberFormat('#,##0.00');
    pr.getRange('J2:M').setNumberFormat('#,##0.00');
    pr.getRange('H2:H').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Mensual', 'Quincenal', 'Semanal'], true).build());
    pr.setFrozenRows(1);
    pr.setColumnWidths(1, 16, 120);
  }
  if (!ss.getSheetByName(HOJA_CFG)) {
    const cfg = ss.insertSheet(HOJA_CFG);
    cfg.getRange('A1:B1').setValues([['Tipo de cambio (S/ por US$)', 3.75]]);
    cfg.getRange('A1').setFontWeight('bold');
    cfg.setColumnWidth(1, 230);
  }
  if (!ss.getSheetByName(HOJA_PRES)) {
    const p = ss.insertSheet(HOJA_PRES);
    p.getRange('A1').setValue('Mes (AAAA-MM):').setFontWeight('bold');
    p.getRange('B1').setFormula('=TEXT(TODAY(),"yyyy-mm")');
    p.getRange('C1').setValue('Escribe otro mes en B1 para revisarlo');
    p.getRange('A3:F3').setValues([['Categoría', 'Presupuesto S/', 'Gastado S/', 'Restante S/', '% usado', 'Estado']]).setFontWeight('bold').setBackground('#e8edf7');
    const filas = CAT_GASTO.map(function (c, i) {
      const r = i + 4;
      return [c, 0,
        '=SUMIFS(Movimientos!K:K,Movimientos!C:C,$B$1,Movimientos!D:D,"Gasto",Movimientos!E:E,A' + r + ')',
        '=B' + r + '-C' + r,
        '=IF(B' + r + '=0,"",C' + r + '/B' + r + ')',
        '=IF(B' + r + '=0,IF(C' + r + '>0,"Sin límite",""),IF(C' + r + '/B' + r + '>1,"Excedido",IF(C' + r + '/B' + r + '>=0.8,"Cerca del límite","OK")))'];
    });
    p.getRange(4, 1, filas.length, 6).setValues(filas);
    const t = CAT_GASTO.length + 4;
    p.getRange(t, 1, 1, 5).setValues([['Total', '=SUM(B4:B' + (t - 1) + ')', '=SUM(C4:C' + (t - 1) + ')', '=B' + t + '-C' + t, '=IF(B' + t + '=0,"",C' + t + '/B' + t + ')']]);
    p.getRange(t, 1, 1, 6).setFontWeight('bold');
    p.getRange('B4:D' + t).setNumberFormat('#,##0.00');
    p.getRange('E4:E' + t).setNumberFormat('0%');
    p.setColumnWidth(1, 240);
  }
  if (!ss.getSheetByName(HOJA_RES)) {
    const r = ss.insertSheet(HOJA_RES);
    r.getRange('A1:E1').setValues([['Mes', 'Ingresos S/', 'Gastos S/', 'Balance S/', 'Tasa de ahorro']]).setFontWeight('bold').setBackground('#e8edf7');
    r.getRange('A2').setFormula('=IFERROR(SORT(UNIQUE(FILTER(Movimientos!C2:C,Movimientos!C2:C<>""))),"")');
    r.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Movimientos!K:K,Movimientos!C:C,A2:A,Movimientos!D:D,"Ingreso")))');
    r.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Movimientos!K:K,Movimientos!C:C,A2:A,Movimientos!D:D,"Gasto")))');
    r.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))');
    r.getRange('E2').setFormula('=ARRAYFORMULA(IF(A2:A="","",IF(B2:B=0,"",D2:D/B2:B)))');
    r.getRange('B2:D').setNumberFormat('#,##0.00');
    r.getRange('E2:E').setNumberFormat('0%');
    r.setColumnWidths(1, 5, 130);
  }
  if (!ss.getSheetByName(HOJA_CAT)) {
    ss.insertSheet(HOJA_CAT).setColumnWidth(1, 240);
  }
  // Resúmenes sin transferencias entre cuentas ni ajustes de saldo (se reescriben al actualizar la versión)
  const excl = ',Movimientos!E:E,"<>' + CAT_TRANSF + '",Movimientos!E:E,"<>' + CAT_AJUSTE + '"';
  const res = ss.getSheetByName(HOJA_RES);
  res.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Movimientos!K:K,Movimientos!C:C,A2:A,Movimientos!D:D,"Ingreso"' + excl + ')))');
  res.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIFS(Movimientos!K:K,Movimientos!C:C,A2:A,Movimientos!D:D,"Gasto"' + excl + ')))');
  ss.getSheetByName(HOJA_CAT).getRange('A1').setFormula('=IFERROR(QUERY(Movimientos!C:K,"select E, sum(K) where D = \'Gasto\' and E <> \'' + CAT_AJUSTE + '\' group by E pivot C label E \'Categoría\'",1),"Aún no hay gastos")');
}

function aFechaTexto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const t = String(v || '').trim();
  let m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

function leerTodo_() {
  const ss = ss_();
  const sh = ss.getSheetByName(HOJA_MOV);
  const n = ultimaFila_(sh, [COL.fecha, COL.monto]) - 1;
  const movimientos = [];
  if (n > 0) {
    const vals = sh.getRange(2, 1, n, ENCABEZADOS.length).getValues();
    const idsNuevos = [];
    vals.forEach(function (r, i) {
      const fecha = aFechaTexto_(r[COL.fecha - 1]);
      const monto = Number(r[COL.monto - 1]);
      if (!fecha || !(monto > 0)) return;
      let id = String(r[COL.id - 1] || '');
      if (!id) { id = Utilities.getUuid().slice(0, 8); idsNuevos.push([i + 2, id]); }
      const moneda = String(r[COL.moneda - 1]).toUpperCase() === 'USD' ? 'USD' : 'PEN';
      const tc = moneda === 'USD' ? (Number(r[COL.tc - 1]) || leerTc_()) : 1;
      const tipo = String(r[COL.tipo - 1]).toLowerCase().indexOf('ing') === 0 ? 'Ingreso' : 'Gasto';
      movimientos.push({
        id: id, fecha: fecha, mes: fecha.slice(0, 7), tipo: tipo,
        categoria: String(r[COL.categoria - 1] || (tipo === 'Gasto' ? 'Otros gastos' : 'Otros ingresos')),
        nota: String(r[COL.nota - 1] || ''), metodo: String(r[COL.metodo - 1] || ''),
        moneda: moneda, monto: monto, tc: tc, montoPEN: Math.round(monto * tc * 100) / 100,
        creado: r[COL.creado - 1] instanceof Date ? r[COL.creado - 1].toISOString() : String(r[COL.creado - 1] || ''),
        prestamoId: String(r[COL.prestamo - 1] || ''),
        cuenta: String(r[COL.cuenta - 1] || ''), destino: String(r[COL.destino - 1] || '')
      });
    });
    // Filas escritas a mano reciben un ID para poder borrarlas desde el formulario
    idsNuevos.forEach(function (x) { sh.getRange(x[0], COL.id).setValue(x[1]); });
  }
  return { movimientos: movimientos, cuentas: leerCuentas_(), prestamos: leerPrestamos_(), config: leerConfig_(), hoja: { url: ss.getUrl(), id: ss.getId(), nombre: ss.getName() } };
}

/** Última fila con datos reales (getLastRow no sirve: las ARRAYFORMULA llenan toda la columna). */
function ultimaFila_(sh, cols) {
  const max = sh.getMaxRows();
  let ultima = 1;
  cols.forEach(function (c) {
    const abajo = sh.getRange(max, c);
    const r = abajo.getValue() !== '' ? max : abajo.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    if (r > ultima) ultima = r;
  });
  return ultima;
}

/** Fila (1-based) de cada ID en la columna A, en una sola lectura. */
function filasPorId_(sh, cols) {
  const n = ultimaFila_(sh, cols) - 1;
  const mapa = {};
  if (n < 1) return mapa;
  sh.getRange(2, 1, n, 1).getValues().forEach(function (r, i) { if (r[0] !== '') mapa[String(r[0])] = i + 2; });
  return mapa;
}

function leerTc_() {
  return Number(ss_().getSheetByName(HOJA_CFG).getRange('B1').getValue()) || 3.75;
}

function leerConfig_() {
  const p = ss_().getSheetByName(HOJA_PRES);
  const filas = p.getRange(4, 1, CAT_GASTO.length, 2).getValues();
  const presupuestos = {};
  filas.forEach(function (r) { if (r[0] && Number(r[1]) > 0) presupuestos[r[0]] = Number(r[1]); });
  return { tc: leerTc_(), presupuestos: presupuestos };
}

function agregar_(m) {
  const sh = ss_().getSheetByName(HOJA_MOV);
  escribirMov_(sh, ultimaFila_(sh, [COL.fecha, COL.monto]) + 1, m, true);
}

/** Edita un movimiento en su misma fila (conserva su hora de registro y su lugar). */
function editar_(m) {
  if (!m || !m.id) throw new Error('Falta el movimiento');
  const sh = ss_().getSheetByName(HOJA_MOV);
  const fila = filasPorId_(sh, [COL.fecha, COL.monto])[String(m.id)];
  if (!fila) return agregar_(m);   // si ya no existe, se agrega
  escribirMov_(sh, fila, m, false);
}

/** Cambia solo la hora de registro (columna Registrado) para fijar el orden. */
function reordenar_(items) {
  if (!items || !items.length) return;
  const sh = ss_().getSheetByName(HOJA_MOV);
  const filas = filasPorId_(sh, [COL.fecha, COL.monto]);
  items.forEach(function (it) {
    const f = filas[String(it.id)];
    const d = new Date(it.creado);
    if (f && !isNaN(d.getTime())) sh.getRange(f, COL.creado).setValue(d);
  });
}

function escribirMov_(sh, r, m, nuevo) {
  if (!m) throw new Error('Falta el movimiento');
  const fecha = aFechaTexto_(m.fecha);
  const monto = Number(m.monto);
  if (!fecha) throw new Error('Fecha no válida');
  if (!(monto > 0)) throw new Error('Monto no válido');
  const p = fecha.split('-').map(Number);
  const moneda = m.moneda === 'USD' ? 'USD' : 'PEN';
  const tc = moneda === 'USD' ? Number(m.tc) : 1;
  if (!(tc > 0)) throw new Error('Tipo de cambio no válido');
  const id = String(m.id || Utilities.getUuid().slice(0, 8));
  // C (Mes) y K (Monto S/) las calculan las fórmulas de la fila 1
  sh.getRange(r, COL.id, 1, 2).setValues([[id, new Date(p[0], p[1] - 1, p[2])]]);
  sh.getRange(r, COL.tipo, 1, 7).setValues([[m.tipo === 'Ingreso' ? 'Ingreso' : 'Gasto', String(m.categoria || ''), String(m.nota || ''), String(m.metodo || ''), moneda, monto, tc]]);
  const cuenta = String(m.cuenta || ''), destino = m.categoria === CAT_TRANSF ? String(m.destino || '') : '';
  if (nuevo) sh.getRange(r, COL.creado, 1, 4).setValues([[new Date(), String(m.prestamoId || ''), cuenta, destino]]);
  else sh.getRange(r, COL.prestamo, 1, 3).setValues([[String(m.prestamoId || ''), cuenta, destino]]);
}

function borrar_(id) {
  if (!id) throw new Error('Falta el ID');
  const sh = ss_().getSheetByName(HOJA_MOV);
  const fila = filasPorId_(sh, [COL.fecha, COL.monto])[String(id)];
  if (fila) sh.deleteRow(fila);
}

function guardarConfig_(c) {
  if (!c) return;
  const ss = ss_();
  if (Number(c.tc) > 0) ss.getSheetByName(HOJA_CFG).getRange('B1').setValue(Number(c.tc));
  if (c.presupuestos) {
    const p = ss.getSheetByName(HOJA_PRES);
    const cats = p.getRange(4, 1, CAT_GASTO.length, 1).getValues();
    const vals = cats.map(function (r) { const v = Number(c.presupuestos[r[0]]); return [v > 0 ? v : 0]; });
    p.getRange(4, 2, vals.length, 1).setValues(vals);
  }
}

/* ---------- Préstamos ---------- */

function aFecha_(t) {
  const p = aFechaTexto_(t).split('-').map(Number);
  if (p.length !== 3 || !p[0]) throw new Error('Fecha no válida');
  return new Date(p[0], p[1] - 1, p[2]);
}

function leerPrestamos_() {
  const sh = ss_().getSheetByName(HOJA_PRE);
  const n = ultimaFila_(sh, [PCOL.id, PCOL.entidad]) - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, P_ENCABEZADOS.length).getValues();
  const out = [];
  vals.forEach(function (r, i) {
    const cuotas = Number(r[PCOL.cuotas - 1]), cuota = Number(r[PCOL.cuota - 1]);
    if (!r[PCOL.entidad - 1] || !(cuotas > 0) || !(cuota > 0)) return;
    let id = String(r[PCOL.id - 1] || '');
    if (!id) { id = 'p' + Utilities.getUuid().slice(0, 8); sh.getRange(i + 2, PCOL.id).setValue(id); }
    out.push({
      id: id, entidad: String(r[PCOL.entidad - 1]), nota: String(r[PCOL.nota - 1] || ''),
      fecha: aFechaTexto_(r[PCOL.fecha - 1]), monto: Number(r[PCOL.monto - 1]) || 0,
      cuotas: cuotas, cuota: cuota, frecuencia: String(r[PCOL.frecuencia - 1] || 'Mensual'),
      primer: aFechaTexto_(r[PCOL.primer - 1])
    });
  });
  return out;
}

function agregarPrestamo_(p) {
  if (!p || !String(p.entidad || '').trim()) throw new Error('Falta la entidad');
  const cuotas = Math.round(Number(p.cuotas)), cuota = Number(p.cuota), monto = Number(p.monto);
  if (!(monto > 0)) throw new Error('Monto recibido no válido');
  if (!(cuotas > 0)) throw new Error('Número de cuotas no válido');
  if (!(cuota > 0)) throw new Error('Monto de cuota no válido');
  const frec = ['Mensual', 'Quincenal', 'Semanal'].indexOf(p.frecuencia) >= 0 ? p.frecuencia : 'Mensual';
  const sh = ss_().getSheetByName(HOJA_PRE);
  const r = ultimaFila_(sh, [PCOL.id, PCOL.entidad]) + 1;
  const id = String(p.id || ('p' + Utilities.getUuid().slice(0, 8)));
  sh.getRange(r, 1, 1, 9).setValues([[id, String(p.entidad).trim(), String(p.nota || ''), aFecha_(p.fecha), monto, cuotas, cuota, frec, aFecha_(p.primer)]]);
  sh.getRange(r, 10, 1, 7).setValues([[
    '=F' + r + '*G' + r,
    '=J' + r + '-E' + r,
    '=SUMIFS(Movimientos!K:K,Movimientos!M:M,A' + r + ')',
    '=MAX(0,J' + r + '-L' + r + ')',
    '=COUNTIFS(Movimientos!M:M,A' + r + ')',
    '=IF(M' + r + '<=0.009,"Pagado","Pendiente")',
    new Date()
  ]]);
}

function borrarPrestamo_(id) {
  if (!id) throw new Error('Falta el ID');
  const sh = ss_().getSheetByName(HOJA_PRE);
  const n = ultimaFila_(sh, [PCOL.id, PCOL.entidad]) - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, PCOL.id, n, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); return; }
  }
}

/* ---------- Cuentas ---------- */

function leerCuentas_() {
  const sh = ss_().getSheetByName(HOJA_CTA);
  const n = ultimaFila_(sh, [CCOL.id, CCOL.nombre]) - 1;
  if (n < 1) return [];
  const out = [];
  sh.getRange(2, 1, n, C_ENCABEZADOS.length).getValues().forEach(function (r, i) {
    if (!r[CCOL.nombre - 1]) return;
    let id = String(r[CCOL.id - 1] || '');
    if (!id) { id = 'c' + Utilities.getUuid().slice(0, 8); sh.getRange(i + 2, CCOL.id).setValue(id); }
    out.push({
      id: id, nombre: String(r[CCOL.nombre - 1]), tipo: String(r[CCOL.tipo - 1] || 'Otro'),
      archivada: r[CCOL.archivada - 1] === true || String(r[CCOL.archivada - 1]).toLowerCase() === 'sí',
      creado: r[CCOL.creado - 1] instanceof Date ? r[CCOL.creado - 1].toISOString() : String(r[CCOL.creado - 1] || '')
    });
  });
  return out;
}

function formulaSaldo_(r) {
  // ingresos - gastos de la cuenta, - lo que salió hacia otra cuenta, + lo que llegó desde otra cuenta
  const k = 'Movimientos!K:K', n = 'Movimientos!N:N', a = '$A' + r;
  return '=SUMIFS(' + k + ',' + n + ',' + a + ',Movimientos!D:D,"Ingreso",Movimientos!E:E,"<>' + CAT_TRANSF + '")' +
    '-SUMIFS(' + k + ',' + n + ',' + a + ',Movimientos!D:D,"Gasto")' +
    '-SUMIFS(' + k + ',' + n + ',' + a + ',Movimientos!E:E,"' + CAT_TRANSF + '")' +
    '+SUMIFS(' + k + ',Movimientos!O:O,' + a + ')';
}

function agregarCuenta_(c) {
  if (!c || !String(c.nombre || '').trim()) throw new Error('Falta el nombre de la cuenta');
  const sh = ss_().getSheetByName(HOJA_CTA);
  const r = ultimaFila_(sh, [CCOL.id, CCOL.nombre]) + 1;
  const tipo = TIPOS_CUENTA.indexOf(c.tipo) >= 0 ? c.tipo : 'Otro';
  sh.getRange(r, 1, 1, 6).setValues([[String(c.id || ('c' + Utilities.getUuid().slice(0, 8))), String(c.nombre).trim(), tipo, c.archivada ? 'Sí' : 'No', new Date(), formulaSaldo_(r)]]);
}

function editarCuenta_(c) {
  if (!c || !c.id) throw new Error('Falta la cuenta');
  const sh = ss_().getSheetByName(HOJA_CTA);
  const f = filasPorId_(sh, [CCOL.id, CCOL.nombre])[String(c.id)];
  if (!f) return agregarCuenta_(c);
  const tipo = TIPOS_CUENTA.indexOf(c.tipo) >= 0 ? c.tipo : 'Otro';
  sh.getRange(f, CCOL.nombre, 1, 3).setValues([[String(c.nombre || '').trim(), tipo, c.archivada ? 'Sí' : 'No']]);
}

function borrarCuenta_(id) {
  if (!id) throw new Error('Falta el ID');
  const sh = ss_().getSheetByName(HOJA_CTA);
  const f = filasPorId_(sh, [CCOL.id, CCOL.nombre])[String(id)];
  if (f) sh.deleteRow(f);
}
