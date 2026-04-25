/**
 * KINETICA — Servidor de Licencias y Actualizaciones (Google Apps Script)
 *
 * doPost — verificación de licencias        (hoja: "licencias")
 * doGet  — consulta de versión disponible   (hoja: "config")
 *
 * Ver docs/licencias.md y docs/actualizaciones.md para el workflow completo.
 */

// ── Constantes de columnas (base 1) ──────────────────────────────────────────
var COL_KEY        = 1;
var COL_CLIENTE    = 2;
var COL_EXPIRACION = 3;
var COL_MACHINE_ID = 4;
var COL_ACTIVADO   = 5;
var COL_NOTAS      = 6;
var SHEET_NAME     = "licencias";

// ── Punto de entrada HTTP GET — consulta de versión ──────────────────────────

function doGet(e) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("config");
    if (!sheet) {
      return _respond({ error: "config_not_found" });
    }

    var values = sheet.getDataRange().getValues();
    var config = {};
    for (var i = 0; i < values.length; i++) {
      var key = String(values[i][0]).trim();
      var val = String(values[i][1]).trim();
      if (key) config[key] = val;
    }

    var version     = config["version"]      || "0.0.0";
    var downloadUrl = config["download_url"] || "";

    if (!version || !downloadUrl) {
      return _respond({ error: "config_incomplete" });
    }

    return _respond({ version: version, download_url: downloadUrl });

  } catch (err) {
    return _respond({ error: err.toString() });
  }
}

// ── Punto de entrada HTTP POST — verificación de licencia ─────────────────────

function doPost(e) {
  try {
    var body       = JSON.parse(e.postData.contents);
    var key        = String(body.key        || "").trim().toUpperCase();
    var machineId  = String(body.machine_id || "").trim();

    if (!key || !machineId) {
      return _respond({ valid: false, reason: "bad_request" });
    }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      var nombres = ss.getSheets().map(function(s) { return s.getName(); });
      return _respond({ valid: false, reason: "server_error", detail: "hoja_no_encontrada", hojas: nombres });
    }

    var values = sheet.getDataRange().getValues();

    for (var i = 1; i < values.length; i++) {
      var row    = values[i];
      var rowKey = String(row[COL_KEY - 1]).trim().toUpperCase();

      if (rowKey !== key) continue;

      // ── Clave encontrada ──────────────────────────────────────────────────

      // La celda puede ser texto "YYYY-MM-DD" o un Date nativo de Sheets —
      // Google convierte automáticamente las celdas con formato de fecha.
      var rawExp  = row[COL_EXPIRACION - 1];
      var expDate = (rawExp instanceof Date) ? rawExp : new Date(String(rawExp).trim() + "T00:00:00");
      var today   = new Date();
      today.setHours(0, 0, 0, 0);

      if (isNaN(expDate.getTime())) {
        return _respond({ valid: false, reason: "server_error" });
      }

      // Siempre devolvemos la fecha como string ISO para consistencia
      var tz     = Session.getScriptTimeZone();
      var expStr = Utilities.formatDate(expDate, tz, "yyyy-MM-dd");

      if (expDate < today) {
        return _respond({ valid: false, reason: "expired", expired_on: expStr });
      }

      var daysLeft        = Math.ceil((expDate - today) / 86400000);
      var storedMachineId = String(row[COL_MACHINE_ID - 1]).trim();

      if (!storedMachineId) {
        // Primera activación: registrar machine_id y fecha
        sheet.getRange(i + 1, COL_MACHINE_ID).setValue(machineId);
        sheet.getRange(i + 1, COL_ACTIVADO).setValue(
          Utilities.formatDate(new Date(), tz, "dd/MM/yyyy")
        );
        return _respond({
          valid:            true,
          first_activation: true,
          days_left:        daysLeft,
          expires:          expStr
        });
      }

      if (storedMachineId === machineId) {
        return _respond({ valid: true, days_left: daysLeft, expires: expStr });
      }

      return _respond({ valid: false, reason: "machine_mismatch" });
    }

    return _respond({ valid: false, reason: "key_not_found" });

  } catch (err) {
    return _respond({ valid: false, reason: "server_error", detail: err.toString() });
  }
}

// ── Helper ────────────────────────────────────────────────────────────────────

function _respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
