/**
 * Plan Ninja Empanada · Backend del CRM en Google Sheets
 * Pegar en Extensiones > Apps Script de tu hoja de cálculo.
 */
const PIN = '1111';        // ← CÁMBIALO antes de publicar
const NOTIFY_EMAIL = 'joseph.edu.ceo@gmail.com';   // recibe un aviso por cada pedido web nuevo

const CAB_C = ['id', 'creado', 'nombre', 'tel', 'zona', 'origen', 'etapa', 'notas'];
const CAB_P = ['id', 'cid', 'fecha', 'estado', 'carne3', 'pollo3', 'mixta3', 'carne5', 'pollo5', 'mixta5', 'total', 'nota', 'creado'];
const TXT = ['id', 'cid', 'creado', 'nombre', 'tel', 'zona', 'origen', 'etapa', 'notas', 'fecha', 'estado', 'nota'];

function hoja_(n, cab) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.appendRow(cab);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, cab.length).setFontWeight('bold').setBackground('#ede9fe');
    cab.forEach((k, i) => { if (TXT.indexOf(k) >= 0) s.getRange(2, i + 1, s.getMaxRows() - 1, 1).setNumberFormat('@'); });
  }
  return s;
}
function leer_(n, cab) {
  const tz = Session.getScriptTimeZone();
  const v = hoja_(n, cab).getDataRange().getValues(); v.shift();
  return v.filter(r => r[0] !== '').map(r => {
    const o = {};
    cab.forEach((k, i) => { o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], tz, 'yyyy-MM-dd') : r[i]; });
    return o;
  });
}
const cliRow_ = c => CAB_C.map(k => c[k] == null ? '' : String(c[k]));
const pedRow_ = p => [p.id, p.cid, p.fecha, p.estado].concat(p.q.map(Number), [Number(p.total), p.nota || '', p.creado || '']);
const rowPed_ = o => ({ id: String(o.id), cid: String(o.cid), fecha: o.fecha, estado: o.estado, nota: o.nota,
  q: [o.carne3, o.pollo3, o.mixta3, o.carne5, o.pollo5, o.mixta5].map(Number), total: Number(o.total) });

function upsert_(n, cab, row) {
  const s = hoja_(n, cab);
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => String(r[0]));
  const i = ids.indexOf(String(row[0]));
  if (i > 0) s.getRange(i + 1, 1, 1, row.length).setValues([row]); else s.appendRow(row);
}
function borrar_(n, cab, col, val) {
  const s = hoja_(n, cab), v = s.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][col]) === String(val)) s.deleteRow(i + 1);
}
function vaciar_() {
  [['Clientes', CAB_C], ['Pedidos', CAB_P]].forEach(([n, c]) => { const s = hoja_(n, c); if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, c.length).clearContent(); });
}
function cargar_(n, cab, rows) {
  if (!rows.length) return;
  const s = hoja_(n, cab), falta = rows.length + 1 - s.getMaxRows();
  if (falta > 0) s.insertRowsAfter(s.getMaxRows(), falta);
  s.getRange(2, 1, rows.length, cab.length).setValues(rows);
}
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const lim_ = (s, n) => String(s == null ? '' : s).slice(0, n);
const id_ = () => Utilities.getUuid().slice(0, 8);

/** Pedido que llega desde la web pública (sin PIN). */
function lead_(d) {
  if (d.web) return { ok: true };                       // trampa anti-spam
  const nombre = lim_(d.nombre, 60).trim(), tel = lim_(d.tel, 20).trim(), dig = tel.replace(/\D/g, '');
  if (!nombre || dig.length < 7) return { ok: false, error: 'Datos incompletos' };
  const q = (d.q || []).slice(0, 6).map(x => Math.max(0, Math.min(99, parseInt(x) || 0)));
  while (q.length < 6) q.push(0);
  const total = (q[0] + q[1] + q[2]) * 3 + (q[3] + q[4] + q[5]) * 5;
  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  let c = leer_('Clientes', CAB_C).filter(x => String(x.tel).replace(/\D/g, '').slice(-9) === dig.slice(-9))[0];
  if (!c) c = { id: id_(), creado: hoy, nombre: nombre, tel: tel, zona: lim_(d.zona, 80), origen: 'Web', etapa: 'Nuevo', notas: '' };
  else if (d.zona && !c.zona) c.zona = lim_(d.zona, 80);
  upsert_('Clientes', CAB_C, cliRow_(c));
  if (total > 0) {
    const f = /^\d{4}-\d{2}-\d{2}$/.test(d.fecha || '') ? d.fecha : hoy;
    upsert_('Pedidos', CAB_P, pedRow_({ id: id_(), cid: c.id, fecha: f, estado: 'Pendiente', q: q, total: total, nota: lim_(d.nota, 120), creado: hoy }));
    if (NOTIFY_EMAIL) {
      try { MailApp.sendEmail(NOTIFY_EMAIL, '🥟 Nuevo pedido web: ' + nombre + ' (S/ ' + total + ')', nombre + ' · ' + tel + '\nTotal: S/ ' + total + '\nEntrega: ' + f + '\n' + lim_(d.nota, 120)); } catch (e) {}
    }
  }
  return { ok: true };
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents), d = req.data || {};
    if (req.action === 'lead') return out_(lead_(d));
    if (String(req.pin) !== PIN) return out_({ ok: false, error: 'PIN incorrecto' });
    switch (req.action) {
      case 'all': return out_({ ok: true, clientes: leer_('Clientes', CAB_C), pedidos: leer_('Pedidos', CAB_P).map(rowPed_) });
      case 'upsertC': upsert_('Clientes', CAB_C, cliRow_(d)); break;
      case 'upsertP':
        upsert_('Pedidos', CAB_P, pedRow_(d.pedido));
        if (d.cliente) upsert_('Clientes', CAB_C, cliRow_(d.cliente));
        break;
      case 'delC': borrar_('Clientes', CAB_C, 0, d.id); borrar_('Pedidos', CAB_P, 1, d.id); break;
      case 'delP': borrar_('Pedidos', CAB_P, 0, d.id); break;
      case 'wipe': vaciar_(); break;
      case 'replace': vaciar_(); cargar_('Clientes', CAB_C, (d.clientes || []).map(cliRow_)); cargar_('Pedidos', CAB_P, (d.pedidos || []).map(pedRow_)); break;
      default: return out_({ ok: false, error: 'Acción desconocida' });
    }
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
function doGet() { return ContentService.createTextOutput('CRM Ninja activo ✅'); }
