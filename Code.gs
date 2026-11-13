// Apps Script ligado a la hoja "Libro electronico farmacia Pasquiali".
// Pestañas: Enero … Diciembre (una por mes de 2026) y "datos" (los registros anteriores).
// Columnas de cada pestaña (igual que el listado de psicotrópicos):
//   A Nº | B Mat. Médico | C Ape. Médico | D Fecha | E Cant. | F Proveedor | G Producto
// El formulario de index.html envía [id, descripcion, cantidad, fecha, matricula, apellido] y el nombre de la pestaña (tab);
// este script acomoda cada dato en su columna (toSheet_ / fromSheet_) y lo guarda en la pestaña del mes elegido.
var MESES = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
var TABS = MESES.concat(['datos']);
var TOKEN = ''; // opcional: si lo completás, poné el mismo valor en TOKEN de index.html
var HEAD = ['Nº', 'Mat. Médico', 'Ape. Médico', 'Fecha', 'Cant.', 'Proveedor', 'Producto'];
var NCOLS = HEAD.length;

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function norm_(t) {
  return String(t == null ? '' : t).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, '');
}

// nombre de pestaña permitido (mes o "datos"); si no viene o no existe en la lista, "datos"
function tabName_(t) {
  var n = norm_(t);
  for (var i = 0; i < TABS.length; i++) if (norm_(TABS[i]) === n) return TABS[i];
  return 'datos';
}

function sheetByName_(name) {
  var n = norm_(name), all = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < all.length; i++) if (norm_(all[i].getName()) === n) return all[i];
  return null;
}

// formulario -> hoja:  [id, descripcion, cantidad, fecha, matricula, apellido]  ->  [Nº, Mat, Ape, Fecha, Cant, Proveedor, Producto]
function toSheet_(r) { return [r[0], r[4], r[5], r[3], r[2], '', r[1]]; }
// hoja -> formulario
function fromSheet_(s) { return [s[0], s[6], s[4], s[3], s[1], s[2]]; }

// Devuelve la pestaña lista para usar. "datos" con el orden viejo se reordena sola (con respaldo);
// una pestaña de mes vacía recibe los títulos.
function ensureLayout_(tab) {
  var name = tabName_(tab);
  var sh = sheetByName_(name);
  if (!sh) throw new Error('No existe la pestaña "' + name + '" en la hoja.');
  var vacia = sh.getLastRow() === 0;
  var h = sh.getRange(1, 1, 1, NCOLS).getValues()[0];
  if (norm_(h[1]) === 'matmedico' && norm_(h[6]) === 'producto') {   // ya está en el formato nuevo
    if (name === 'datos') repararFilas_(sh);
    return sh;
  }
  var viejo = name === 'datos' && (norm_(h[1]) === 'descripcion' || (norm_(h[0]) === 'id' && norm_(h[5]) === 'apellido'));
  if (!viejo && !vacia) throw new Error('La pestaña "' + name + '" tiene un formato que no reconozco; no se modificó nada.');

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

// Una sola vez: filas de "datos" que quedaron con el orden viejo debajo de los títulos nuevos
// (Mat. Médico con el medicamento, Cant. con la matrícula, Proveedor con el apellido, Producto vacío).
// Se detectan porque "Mat. Médico" tiene letras y "Producto" está vacío; se acomodan y queda un respaldo.
function repararFilas_(sh) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('filasReparadas') === '1') return;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (props.getProperty('filasReparadas') === '1') return;
    var last = sh.getLastRow();
    if (last > 1) {
      var rg = sh.getRange(2, 1, last - 1, NCOLS);
      var data = rg.getValues();
      var malas = 0;
      var nuevas = data.map(function (r) {
        var mal = /[A-Za-z]/.test(String(r[1])) && String(r[6]).trim() === '';
        if (!mal) return r;
        malas++;
        return toSheet_(r.slice(0, 6).map(function (c) { return c instanceof Date ? fmt_(c) : String(c == null ? '' : c); }));
      });
      if (malas) {
        var ss = SpreadsheetApp.getActive();
        sh.copyTo(ss).setName('datos_respaldo_filas_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm'));
        rg.setNumberFormat('@');
        rg.setValues(nuevas);
      }
    }
    props.setProperty('filasReparadas', '1');
  } finally {
    lock.releaseLock();
  }
}

// Para acomodar la hoja "datos" ya mismo: en el editor de Apps Script elegí "acomodarHoja" y tocá ▶ Ejecutar.
function acomodarHoja() { ensureLayout_('datos'); }

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

// pestañas a recorrer en las consultas: la pedida, o todas (las que existan y tengan el formato nuevo)
function sheetsFor_(tab) {
  var names = tab ? [tabName_(tab)] : TABS;
  var res = [];
  names.forEach(function (n) {
    var sh = sheetByName_(n);
    if (!sh) return;
    if (!tab) {   // recorrido general: se salta lo que no tiene el formato nuevo en vez de fallar
      var h = sh.getLastRow() ? sh.getRange(1, 1, 1, NCOLS).getValues()[0] : [];
      if (!(norm_(h[1]) === 'matmedico' && norm_(h[6]) === 'producto')) return;
    } else {
      ensureLayout_(n);
    }
    res.push({ name: n, sh: sh });
  });
  return res;
}

// último Nº usado: en la pestaña pedida; si está vacía, en los meses anteriores y por último en "datos"
function lastIdFor_(tab) {
  var name = tabName_(tab), order = [];
  var mi = MESES.indexOf(name);
  if (mi >= 0) for (var i = mi; i >= 0; i--) order.push(MESES[i]);
  order.push('datos');
  for (var k = 0; k < order.length; k++) {
    var sh = sheetByName_(order[k]);
    if (!sh || sh.getLastRow() < 2) continue;
    var v = sh.getRange(sh.getLastRow(), 1).getValue();
    if (v !== '' && v != null && !isNaN(Number(v))) return v;
  }
  return null;
}

// confirma que la fila sigue siendo la que se mostró: por Nº, o por Producto si la fila no tiene Nº
function mismaFila_(sh, row, body) {
  var id = String(sh.getRange(row, 1).getValue());
  if (id !== String(body.expectId == null ? '' : body.expectId)) return false;
  if (id === '' && String(sh.getRange(row, 7).getValue()) !== String(body.expectDesc == null ? '' : body.expectDesc)) return false;
  return true;
}

// Consulta pública del Buscador Nacional de Profesionales de la Salud (REFEPS): devuelve los MÉDICOS de la provincia
// de Buenos Aires que tienen esa matrícula. Se pide la página, se toma el formulario y se envía igual que lo haría el navegador.
var REFEPS_URL = 'https://www.argentina.gob.ar/salud/buscador-nacional-de-profesionales-de-la-salud';
function buscarProfesional_(matricula) {
  var mat = String(matricula || '').replace(/\D/g, '');
  if (mat.length < 4 || mat.length > 9) throw new Error('matrícula inválida');
  var cache = CacheService.getScriptCache(), key = 'refeps_' + mat;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var page = UrlFetchApp.fetch(REFEPS_URL, { muteHttpExceptions: true }).getContentText();
  var at = page.indexOf('consulta-profesionales-form');
  var bid = at >= 0 ? /name="form_build_id" value="([^"]+)"/.exec(page.slice(at)) : null;
  if (!bid) throw new Error('el buscador nacional cambió su formulario');
  var resp = UrlFetchApp.fetch(REFEPS_URL, {
    method: 'post', muteHttpExceptions: true, followRedirects: true,
    payload: { searchBy: 'matricula', dni: '', matricula: mat, apellidonombre: '', form_build_id: bid[1],
               form_id: 'argobar_consulta_refeps_profesionales', tarro_de_miel: '', op: 'Consultar' }
  });
  var html = resp.getContentText();
  var m = /allItems\s*=\s*(\[[\s\S]*?\]);\s*\n/.exec(html), all = [];
  if (m) all = JSON.parse(m[1]);
  else if (html.indexOf('resultados encontrados') < 0 && html.indexOf('refepsProfesionales') < 0 && html.indexOf('No se encontr') < 0) {
    throw new Error('el buscador nacional no respondió como se esperaba');
  }
  var items = [];
  all.forEach(function (x) {
    var ok = (x.profesiones || []).some(function (p) {
      return norm_(p.profesionReferencia) === 'medico' && (p.matriculas || []).some(function (mm) {
        return String(mm.matricula).replace(/\D/g, '') === mat && norm_(mm.provinciaMatricula) === 'buenosaires';
      });
    });
    if (!ok) return;
    var mm0 = null, esp = '';
    (x.profesiones || []).forEach(function (p) {
      if (norm_(p.profesionReferencia) !== 'medico') return;
      (p.matriculas || []).forEach(function (mm) {
        if (!mm0 && String(mm.matricula).replace(/\D/g, '') === mat && norm_(mm.provinciaMatricula) === 'buenosaires') { mm0 = mm; esp = p.refepsEspecialidad || ''; }
      });
    });
    items.push({ apellido: x.apellido, nombre: x.nombre, situacion: mm0 ? mm0.situacionMatricula : '', especialidad: esp });
  });
  var out = { items: items, total: all.length };
  cache.put(key, JSON.stringify(out), 21600);
  return out;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (TOKEN && p.token !== TOKEN && p.action) return out_({ ok: false, error: 'token inválido' });

  if (p.action === 'list') {
    if (!p.from || !p.to) return out_({ ok: false, error: 'faltan fechas' });
    var from = isoToDate_(p.from), to = isoToDate_(p.to);
    var rows = [], skipped = 0, list;
    try { list = sheetsFor_(p.tab); } catch (err) { return out_({ ok: false, error: String(err) }); }
    list.forEach(function (t) {
      var last = t.sh.getLastRow();
      if (last < 2) return;
      t.sh.getRange(2, 1, last - 1, NCOLS).getValues().forEach(function (r, i) {
        if (r.every(function (c) { return c === '' })) return;
        var d = parseDate_(r[3]);
        if (!d) { skipped++; return; }
        if (d >= from && d <= to) {
          var v = fromSheet_(r).map(function (c) { return c instanceof Date ? fmt_(c) : String(c); });
          v[3] = fmt_(d);
          rows.push({ tab: t.name, row: i + 2, values: { id: v[0], descripcion: v[1], cantidad: v[2], fecha: v[3], matricula: v[4], apellido: v[5] } });
        }
      });
    });
    return out_({ ok: true, rows: rows, skipped: skipped });
  }

  if (p.action === 'profesional') {
    try { var r = buscarProfesional_(p.matricula); return out_({ ok: true, items: r.items, total: r.total }); }
    catch (err) { return out_({ ok: false, error: String(err && err.message || err) }); }
  }

  if (p.action === 'matriculas') {
    // pares [matrícula, apellido] de todas las pestañas, para autocompletar el apellido
    var pairs = [];
    try {
      sheetsFor_(p.tab).forEach(function (t) {
        var last = t.sh.getLastRow();
        if (last < 2) return;
        t.sh.getRange(2, 2, last - 1, 2).getValues().forEach(function (r) {
          var m = String(r[0]).trim(), a = String(r[1]).trim();
          if (m && a) pairs.push([m, a]);
        });
      });
    } catch (err) { return out_({ ok: false, error: String(err) }); }
    return out_({ ok: true, items: pairs });
  }

  // estado: último Nº para la pestaña pedida (y datos para que index.html sepa que el script está al día)
  var tab = p.tab ? tabName_(p.tab) : 'datos';
  try { ensureLayout_(tab); } catch (err) { return out_({ ok: false, error: String(err) }); }
  // cols = campos que maneja el formulario (6); layout 4 = guarda según la pestaña (mes) y admite filas sin Nº
  return out_({ ok: true, lastId: lastIdFor_(tab), tab: tab, cols: 6, layout: 4 });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (TOKEN && body.token !== TOKEN) return out_({ ok: false, error: 'token inválido' });
    var sh = ensureLayout_(body.tab);

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      if (body.action === 'delete') {
        var drow = Number(body.row);
        if (!drow || drow < 2 || drow > sh.getLastRow()) return out_({ ok: false, error: 'fila inválida' });
        if (!mismaFila_(sh, drow, body)) return out_({ ok: false, error: 'la fila cambió en la hoja, volvé a buscar' });
        var rawId = sh.getRange(drow, 1).getValue();
        var base = rawId === '' ? NaN : Number(rawId); // Nº de la fila que se borra: lo toma la fila siguiente (si no tiene Nº, no se renumera)
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
        if (!mismaFila_(sh, row, body)) return out_({ ok: false, error: 'la fila cambió en la hoja, volvé a buscar' });
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
      return out_({ ok: true, added: rows.length, tab: sh.getName(), firstRow: start, lastRow: start + rows.length - 1 });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
