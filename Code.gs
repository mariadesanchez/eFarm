// Apps Script ligado a la hoja "Libro electronico farmacia Pasquiali".
// Pestaña: datos. Columnas: id, descripcion, cantida, fecha, matricula, nombre_apellido.
var TAB = 'datos';
var TOKEN = ''; // opcional: si lo completás, poné el mismo valor en TOKEN de index.html

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() { return SpreadsheetApp.getActive().getSheetByName(TAB); }

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
  var sh = sheet_();
  var last = sh.getLastRow();

  if (p.action === 'list') {
    if (!p.from || !p.to) return out_({ ok: false, error: 'faltan fechas' });
    var from = isoToDate_(p.from), to = isoToDate_(p.to);
    var rows = [], skipped = 0;
    if (last > 1) {
      var data = sh.getRange(2, 1, last - 1, 6).getValues();
      data.forEach(function (r, i) {
        if (r.every(function (c) { return c === '' })) return;
        var d = parseDate_(r[3]);
        if (!d) { skipped++; return; }
        if (d >= from && d <= to) {
          var v = r.map(function (c) { return c instanceof Date ? fmt_(c) : String(c); });
          v[3] = fmt_(d);
          rows.push({ row: i + 2, values: { id: v[0], descripcion: v[1], cantidad: v[2], fecha: v[3], matricula: v[4], nombre_apellido: v[5] } });
        }
      });
    }
    return out_({ ok: true, rows: rows, skipped: skipped });
  }

  var lastId = last > 1 ? sh.getRange(last, 1).getValue() : null;
  return out_({ ok: true, lastId: lastId, rows: Math.max(0, last - 1) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out_({ ok: false, error: 'token inválido' });
    var sh = sheet_();

    if (body.action === 'update') {
      var row = Number(body.row), v = body.values;
      if (!row || row < 2 || row > sh.getLastRow() || !v || v.length !== 6) return out_({ ok: false, error: 'fila inválida' });
      // seguridad: la fila debe seguir siendo la misma que se mostró
      if (String(sh.getRange(row, 1).getValue()) !== String(body.expectId)) return out_({ ok: false, error: 'la fila cambió en la hoja, volvé a buscar' });
      var rg = sh.getRange(row, 1, 1, 6);
      rg.setNumberFormat('@');
      rg.setValues([v.map(function (x) { return String(x == null ? '' : x); })]);
      return out_({ ok: true, row: row });
    }

    var rows = body.rows;
    if (!rows || !rows.length) return out_({ ok: false, error: 'sin filas' });
    var start = sh.getLastRow() + 1;
    var range = sh.getRange(start, 1, rows.length, 6);
    range.setNumberFormat('@'); // texto: evita que fechas o matrículas se conviertan
    range.setValues(rows.map(function (r) {
      return [r[0], r[1], r[2], r[3], r[4], r[5]].map(function (v) { return String(v == null ? '' : v); });
    }));
    return out_({ ok: true, added: rows.length, firstRow: start, lastRow: start + rows.length - 1 });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
