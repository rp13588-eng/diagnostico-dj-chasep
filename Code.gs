/**
 * BACKEND – Diagnóstico Distribuidores Junior CHASEP
 * Guarda las respuestas de encuesta.html en Google Sheets
 * y se las entrega a resultados.html (protegido con clave).
 *
 * 1) script.google.com > Nuevo proyecto > pega este código > Guardar
 * 2) Cambia CLAVE_RESULTADOS abajo (mínimo 8 caracteres)
 * 3) Ejecuta la función "setup" una vez (acepta permisos) → crea la hoja
 * 4) Implementar > Nueva implementación > Tipo: Aplicación web
 *    Ejecutar como: Yo · Quién tiene acceso: Cualquier persona > Implementar
 * 5) Copia la URL que termina en /exec y pégala en encuesta.html y resultados.html
 *
 * Si después cambias este código: Implementar > Administrar implementaciones >
 * editar (lápiz) > Versión: Nueva versión > Implementar (la URL no cambia).
 */

var CLAVE_RESULTADOS = 'CAMBIA_ESTA_CLAVE'; // <-- cámbiala (mín. 8 caracteres): se pide para ver resultados

var CAMPOS = [
  ['fecha', 'Fecha'],
  ['nombre', 'Nombre'],
  ['sucursal', 'Sucursal'],
  ['antiguedad', 'Antigüedad'],
  ['seguridad', 'Seguridad demo (1-10)'],
  ['partes', 'Partes difíciles de la demo'],
  ['objecion', 'Objeción más difícil'],
  ['capacitaciones', 'Capacitaciones último mes'],
  ['capUtil', 'Capacitación más útil'],
  ['fuentes', 'Fuentes de prospectos'],
  ['mejorFuente', 'Mejor fuente'],
  ['registro', 'Dónde lleva seguimiento'],
  ['activos', 'Vendedores activos'],
  ['arranque', 'En capacitación / arranque'],
  ['juntas', 'Frecuencia de juntas'],
  ['citas', 'Citas por semana'],
  ['demos', 'Demos por semana'],
  ['ventas', 'Ventas por semana'],
  ['quitaTiempo', 'Qué le quita tiempo'],
  ['pedido', 'Qué pide a la oficina'],
  ['metaDic', 'Meta ventas al 31 dic'],
  ['falta', 'Qué le falta']
];

function setup() {
  var props = PropertiesService.getScriptProperties();
  var existente = props.getProperty('SHEET_ID');
  if (existente) {
    Logger.log('Ya existe la hoja: ' + SpreadsheetApp.openById(existente).getUrl() + ' (no se creó otra)');
    return;
  }
  var ss = SpreadsheetApp.create('Respuestas – Diagnóstico Distribuidores Junior');
  var sh = ss.getSheets()[0];
  sh.setName('Respuestas');
  sh.appendRow(CAMPOS.map(function (c) { return c[1]; }));
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, CAMPOS.length).setFontWeight('bold').setBackground('#14213D').setFontColor('#F6F3EC');
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('Hoja creada: ' + ss.getUrl());
}

function hoja_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Ejecuta primero la función setup');
  return SpreadsheetApp.openById(id).getSheetByName('Respuestas');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Limpia un valor antes de guardarlo: recorta, limita largo y evita que Sheets lo interprete como fórmula
function limpio_(v) {
  if (v === undefined || v === null) return '';
  if (Array.isArray(v)) v = v.map(String).join(', ');
  if (typeof v === 'number') return isFinite(v) ? v : '';
  v = String(v).trim().slice(0, 2000);
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}

function doPost(e) {
  var d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Datos inválidos' });
  }
  if (!d || !String(d.nombre || '').trim()) return json_({ ok: false, error: 'Falta nombre' });
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var fila = CAMPOS.map(function (c) {
      return c[0] === 'fecha' ? new Date() : limpio_(d[c[0]]);
    });
    hoja_().appendRow(fila);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function doGet(e) {
  if (CLAVE_RESULTADOS === 'CAMBIA_ESTA_CLAVE' || CLAVE_RESULTADOS.length < 8) {
    return json_({ ok: false, error: 'Configura CLAVE_RESULTADOS en el Apps Script' });
  }
  var clave = (e && e.parameter && e.parameter.key) || '';
  if (clave !== CLAVE_RESULTADOS) return json_({ ok: false, error: 'Clave incorrecta' });
  var valores = hoja_().getDataRange().getValues();
  var rows = valores.slice(1).map(function (r) {
    var o = {};
    CAMPOS.forEach(function (c, i) {
      o[c[0]] = r[i] instanceof Date ? r[i].toISOString() : r[i];
    });
    return o;
  });
  return json_({ ok: true, rows: rows });
}
