// Apps Script ligado a la hoja "Libro electronico farmacia Pasquiali".
// Pestaña: datos. Columnas de la hoja (igual que el listado de psicotrópicos):
//   A Nº | B Mat. Médico | C Ape. Médico | D Fecha | E Cant. | F Proveedor | G Producto
// El formulario de index.html no cambia: sigue enviando [id, descripcion, cantidad, fecha, matricula, apellido]
// y este script acomoda cada dato en su columna (toSheet_ / fromSheet_).
// La primera vez que se usa, si la hoja tiene el orden viejo (id, descripcion, cantidad, fecha, matricula, apellido),
// la reordena sola y deja una copia de respaldo en otra pestaña.
var TAB = 'datos';
var TOKEN = ''; // opcional: si lo completás, poné el mismo valor en TOKEN de index.html
var HEAD = ['Nº', 'Mat. Médico', 'Ape. Médico', 'Fecha', 'Cant.', 'Proveedor', 'Producto'];
var NCOLS = HEAD.length;

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() { return SpreadsheetApp.getActive().getSheetByName(TAB); }

// formulario -> hoja:  [id, descripcion, cantidad, fecha, matricula, apellido]  ->  [Nº, Mat, Ape, Fecha, Cant, Proveedor, Producto]
function toSheet_(r) { return [r[0], r[4], r[5], r[3], r[2], '', r[1]]; }
// hoja -> formulario
function fromSheet_(s) { return [s[0], s[6], s[4], s[3], s[1], s[2]]; }

function norm_(t) {
  return String(t == null ? '' : t).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]/g, '');
}

// Deja la hoja con las columnas nuevas. Si tenía el orden viejo, reordena los datos (con respaldo).
function ensureLayout_() {
  var sh = sheet_();
  var h = sh.getRange(1, 1, 1, NCOLS).getValues()[0];
  if (norm_(h[1]) === 'matmedico' && norm_(h[6]) === 'producto') return sh;   // ya está en el formato nuevo
  var viejo = norm_(h[1]) === 'descripcion' || (norm_(h[0]) === 'id' && norm_(h[5]) === 'apellido');
  var vacia = sh.getLastRow() === 0;
  if (!viejo && !vacia) throw new Error('La pestaña "' + TAB + '" tiene un formato que no reconozco; no se modificó nada.');

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var last = sh.getLastRow();
    var filas = [];
    if (viejo && last > 1) {
      var ss = SpreadsheetApp.getActive();
      var copia = sh.copyTo(ss);   // respaldo antes de tocar nada
      copia.setName('datos_respaldo_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm'));
      filas = sh.getRange(2, 1, last - 1, 6).getValues().map(function (r) {
        return toSheet_(r.map(function (c) { return c instanceof Date ? fmt_(c) : String(c == null ? '' : c); }));
      });
    }
    sh.clearContents();
    var hr = sh.getRange(1, 1, 1, NCOLS);
    hr.setNumberFormat('@');
    hr.setValues([HEAD]);
    hr.setFontWeight('bold');
    sh.setFrozenRows(1);
    if (filas.length) {
      var rg = sh.getRange(2, 1, filas.length, NCOLS);
      rg.setNumberFormat('@');
      rg.setValues(filas);
    }
  } finally {
    lock.releaseLock();
  }
  return sh;
}

// Para acomodar la hoja ya mismo: en el editor de Apps Script elegí "acomodarHoja" y tocá ▶ Ejecutar.
function acomodarHoja() { ensureLayout_(); }

// Lee una fecha de la hoja: texto DD/MM/AAAA o fecha real de Sheets. Devuelve Date o null.
function parseDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v).trim());
  if (!m) return null;
  var d = new Date(+m[3], +m[2] - 1, +m[1]);
  return d.getMonth() === +m[2] - 1 ? d : null;
}
function fmt_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function isoToDate_(s) { var p = String(s).split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN && p.action) return out_({ ok: false, error: 'token inválido' });
  var sh;
  try { sh = ensureLayout_(); } catch (err) { return out_({ ok: false, error: String(err) }); }
  var last = sh.getLastRow();

  if (p.action === 'list') {
    if (!p.from || !p.to) return out_({ ok: false, error: 'faltan fechas' });
    var from = isoToDate_(p.from), to = isoToDate_(p.to);
    var rows = [], skipped = 0;
    if (last > 1) {
      var data = sh.getRange(2, 1, last - 1, NCOLS).getValues();
      data.forEach(function (r, i) {
        if (r.every(function (c) { return c === '' })) return;
        var d = parseDate_(r[3]);
        if (!d) { skipped++; return; }
        if (d >= from && d <= to) {
          var v = fromSheet_(r).map(function (c) { return c instanceof Date ? fmt_(c) : String(c); });
          v[3] = fmt_(d);
          rows.push({ row: i + 2, values: { id: v[0], descripcion: v[1], cantidad: v[2], fecha: v[3], matricula: v[4], apellido: v[5] } });
        }
      });
    }
    return out_({ ok: true, rows: rows, skipped: skipped });
  }

  if (p.action === 'matriculas') {
    // pares [matrícula, apellido] de la hoja, para autocompletar el apellido
    var pairs = [];
    if (last > 1) {
      sh.getRange(2, 2, last - 1, 2).getValues().forEach(function (r) {
        var m = String(r[0]).trim(), a = String(r[1]).trim();
        if (m && a) pairs.push([m, a]);
      });
    }
    return out_({ ok: true, items: pairs });
  }

  var lastId = last > 1 ? sh.getRange(last, 1).getValue() : null;
  // cols = campos que maneja el formulario (6); el index.html lo usa para saber que el script está al día
  return out_({ ok: true, lastId: lastId, rows: Math.max(0, last - 1), cols: 6 });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out_({ ok: false, error: 'token inválido' });
    var sh = ensureLayout_();

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'delete') {
        var drow = Number(body.row);
        if (!drow || drow < 2 || drow > sh.getLastRow()) return out_({ ok: false, error: 'fila inválida' });
        if (String(sh.getRange(drow, 1).getValue()) !== String(body.expectId)) return out_({ ok: false, error: 'la fila cambió en la hoja, volvé a buscar' });
        var base = Number(sh.getRange(drow, 1).getValue()); // Nº de la fila que se borra: lo toma la fila siguiente
        sh.deleteRow(drow);
        var lastRow = sh.getLastRow(), renumbered = 0;
        if (lastRow >= drow && !isNaN(base)) {
          var n = lastRow - drow + 1, ids = [];
          for (var k = 0; k < n; k++) ids.push([String(base + k)]);
          var idRange = sh.getRange(drow, 1, n, 1);
          idRange.setNumberFormat('@');
          idRange.setValues(ids);
          renumbered = n;
        }
        return out_({ ok: true, renumbered: renumbered });
      }

      if (body.action === 'update') {
        var row = Number(body.row), v = body.values;
        if (!row || row < 2 || row > sh.getLastRow() || !v || v.length !== 6) return out_({ ok: false, error: 'fila inválida' });
        // seguridad: la fila debe seguir siendo la misma que se mostró
        if (String(sh.getRange(row, 1).getValue()) !== String(body.expectId)) return out_({ ok: false, error: 'la fila cambió en la hoja, volvé a buscar' });
        var nueva = toSheet_(v.map(function (x) { return String(x == null ? '' : x); }));
        nueva[5] = String(sh.getRange(row, 6).getValue());   // el Proveedor no se edita desde el formulario: se conserva
        var rg = sh.getRange(row, 1, 1, NCOLS);
        rg.setNumberFormat('@');
        rg.setValues([nueva]);
        return out_({ ok: true, row: row });
      }

      var rows = body.rows;
      if (!rows || !rows.length) return out_({ ok: false, error: 'sin filas' });
      var start = sh.getLastRow() + 1;
      var range = sh.getRange(start, 1, rows.length, NCOLS);
      range.setNumberFormat('@'); // texto: evita que fechas o matrículas se conviertan
      range.setValues(rows.map(function (r) {
        return toSheet_([r[0], r[1], r[2], r[3], r[4], r[5]].map(function (v) { return String(v == null ? '' : v); }));
      }));
      return out_({ ok: true, added: rows.length, firstRow: start, lastRow: start + rows.length - 1 });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
