/**
 * Apps Script para licencias de PROASE.
 * Hoja: "Datos" (la tabla dentro se llama "Licencias")
 * Columnas: Usuario | Contraseña | Fecha | Estado | Computadora
 * Desplegar como: Aplicación web → Ejecutar como: Yo → Acceso: Cualquiera
 */

// La hoja (pestaña) donde vive la tabla "Licencias"
var SHEET_NAME = "Datos";
var DEFAULT_EXPIRY = "2027-02-15";

function doGet(e) {
  var id = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim() : "";
  if (!id) {
    return _json({ error: "Falta el parámetro id" });
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    return _json({ error: "No existe la hoja " + SHEET_NAME });
  }

  var data = sheet.getDataRange().getValues();
  // data[0] = encabezados
  var foundRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === id) {
      foundRow = i + 1; // fila en la hoja (base 1)
      break;
    }
  }

  var usuario, password, expiry, estado;

  if (foundRow === -1) {
    // Crear nuevo registro
    usuario = "WS-" + id.replace(/-/g, "").slice(0, 6).toUpperCase();
    password = _genPassword();
    expiry = DEFAULT_EXPIRY;
    estado = _computeEstado(expiry);
    sheet.appendRow([usuario, password, expiry, estado, id]);
    return _json({
      active: estado === "Activo",
      usuario: usuario,
      password: password,
      expiry_date: expiry,
      estado: estado,
      mensaje: "Nuevo registro creado"
    });
  }

  // Actualizar estado si cambió
  usuario = String(data[foundRow - 1][0]);
  password = String(data[foundRow - 1][1]);
  expiry = _normDate(data[foundRow - 1][2]);
  var estadoActual = String(data[foundRow - 1][3]).trim();
  estado = _computeEstado(expiry);
  if (estado !== estadoActual) {
    sheet.getRange(foundRow, 4).setValue(estado);
  }

  return _json({
    active: estado === "Activo",
    usuario: usuario,
    password: password,
    expiry_date: expiry,
    estado: estado,
    mensaje: "Registro encontrado"
  });
}

function _computeEstado(expiryStr) {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var expiry = new Date(expiryStr + "T00:00:00");
  return expiry >= today ? "Activo" : "Inactivo";
}

function _normDate(v) {
  if (v instanceof Date) {
    var y = v.getFullYear();
    var m = ('0' + (v.getMonth() + 1)).slice(-2);
    var d = ('0' + v.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
  }
  return String(v).trim();
}

function _genPassword() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789";
  var out = "";
  for (var i = 0; i < 10; i++) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
