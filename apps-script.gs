/**
 * Conecta la página de finanzas con esta planilla de Google Sheets.
 * Pegá este código en Extensiones → Apps Script y publicalo como "Aplicación web"
 * (Ejecutar como: Yo · Quién tiene acceso: Cualquier persona).
 */
var HOJAS = {
  rifa:    { nombre: 'Rifa del Peso', columnas: ['fecha', 'detalle', 'monto'],
             encabezados: ['Fecha', 'Detalle', 'Monto recaudado'] },
  remeras: { nombre: 'Remeras', columnas: ['fecha', 'detalle', 'cantidad', 'recaudo'],
             encabezados: ['Fecha', 'Detalle', 'Cantidad (unidades)', 'Recaudo'] },
  gastos:  { nombre: 'Gastos', columnas: ['fecha', 'categoria', 'detalle', 'monto'],
             encabezados: ['Fecha', 'Categoría', 'Detalle', 'Monto'] }
};

function doGet() {
  return respuesta_(leerTodo_());
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var datos = JSON.parse(e.postData.contents);
    Object.keys(HOJAS).forEach(function (k) { escribir_(k, datos[k] || []); });
    asegurarResumen_();
    return respuesta_({ ok: true, guardado: new Date().toISOString() });
  } finally {
    lock.releaseLock();
  }
}

function leerTodo_() {
  var out = {};
  Object.keys(HOJAS).forEach(function (k) {
    var h = hoja_(k), cols = HOJAS[k].columnas, n = h.getLastRow() - 1;
    out[k] = n < 1 ? [] : h.getRange(2, 1, n, cols.length).getValues()
      .filter(function (fila) { return fila.some(function (v) { return v !== ''; }); })
      .map(function (fila) {
        var o = {};
        cols.forEach(function (c, i) {
          var v = fila[i];
          o[c] = v instanceof Date
            ? Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd')
            : String(v);
        });
        return o;
      });
  });
  return out;
}

function escribir_(k, filas) {
  var h = hoja_(k), cols = HOJAS[k].columnas;
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, cols.length).clearContent();
  if (!filas.length) return;
  var valores = filas.map(function (f) {
    return cols.map(function (c) {
      var v = f[c] == null ? '' : f[c];
      return (c === 'monto' || c === 'recaudo' || c === 'cantidad') && v !== '' ? Number(v) : v;
    });
  });
  h.getRange(2, 1, valores.length, cols.length).setValues(valores);
}

function hoja_(k) {
  var ss = SpreadsheetApp.getActive(), cfg = HOJAS[k], h = ss.getSheetByName(cfg.nombre);
  if (!h) {
    h = ss.insertSheet(cfg.nombre);
    h.getRange(1, 1, 1, cfg.encabezados.length).setValues([cfg.encabezados])
      .setFontWeight('bold').setBackground('#fd66c3');
    h.setFrozenRows(1);
  }
  return h;
}

function asegurarResumen_() {
  var ss = SpreadsheetApp.getActive();
  if (ss.getSheetByName('Resumen')) return;
  var r = ss.insertSheet('Resumen', 0);
  var cats = ['Maquillaje', 'Luces', 'Escenografía', 'Utilería', 'Vestuario', 'Audio'];
  var filas = [
    ['INGRESOS', ''],
    ['Rifa del peso', "=SUM('Rifa del Peso'!C2:C)"],
    ['Remeras (recaudo)', '=SUM(Remeras!D2:D)'],
    ['Remeras vendidas (unidades)', '=SUM(Remeras!C2:C)'],
    ['TOTAL INGRESOS', '=B2+B3'],
    ['', ''],
    ['GASTOS', '']
  ];
  cats.forEach(function (c) { filas.push([c, '=SUMIF(Gastos!B2:B,"' + c + '",Gastos!D2:D)']); });
  filas.push(['TOTAL GASTOS', '=SUM(B8:B13)'], ['', ''], ['SALDO', '=B5-B14']);
  r.getRange(1, 1, filas.length, 2).setValues(filas);
  r.getRange('A1:A16').setFontWeight('bold');
  r.setColumnWidth(1, 220);
}

/** Corré esta función una vez desde el editor para crear las pestañas y dar permisos. */
function prepararPlanilla() {
  Object.keys(HOJAS).forEach(hoja_);
  asegurarResumen_();
}

function respuesta_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
