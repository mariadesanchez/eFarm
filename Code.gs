// Apps Script ligado a la hoja "Libro electronico farmacia Pasquiali".
// Pestaña: datos. Columnas: id, descripcion, cantida, fecha, matricula, nombre_apellido.
var TAB = 'datos';
var TOKEN = ''; // opcional: si lo completás, poné el mismo valor en TOKEN de index.html

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
  var last = sh.getLastRow();
  var lastId = last > 1 ? sh.getRange(last, 1).getValue() : null;
  return out_({ ok: true, lastId: lastId, rows: Math.max(0, last - 1) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out_({ ok: false, error: 'token inválido' });
    var rows = body.rows;
    if (!rows || !rows.length) return out_({ ok: false, error: 'sin filas' });
    var sh = SpreadsheetApp.getActive().getSheetByName(TAB);
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
