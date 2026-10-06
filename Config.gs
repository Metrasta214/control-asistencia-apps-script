const SPREADSHEET_ID = "1GBuWaHHUpZ4xOI9gIfxR-RRJlCF9fhIn4SjzVZwlqdM";

const SHEET_NAMES = Object.freeze({
  PARTICIPANTES: "Participantes",
  SESIONES: "Sesiones",
  ASISTENCIAS: "Asistencias",
});

const HEADERS = Object.freeze({
  PARTICIPANTES: ["id_participante", "nombre", "correo", "activo"],
  SESIONES: ["id_sesion", "numero", "fecha", "tema", "estado"],
  ASISTENCIAS: ["id_sesion", "id_participante", "estado", "hora_registro"],
});

const ESTADOS = Object.freeze({
  PRESENTE: "PRESENTE",
  AUSENTE: "AUSENTE",
  JUSTIFICADA: "JUSTIFICADA",
  SUSPENDIDA: "SUSPENDIDA",
  RETARDO: "RETARDO",
  PROGRAMADA: "PROGRAMADA",
  REGISTRADA: "REGISTRADA",
});

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(
      `No existe la hoja ${sheetName}. Ejecuta prepararBaseDatos().`,
    );
  }

  return sheet;
}

function validateHeaders_(sheet, expectedHeaders) {
  const range = sheet.getRange(1, 1, 1, expectedHeaders.length);

  if (sheet.getLastRow() === 0) {
    range.setValues([expectedHeaders]);
  } else {
    const currentHeaders = range.getValues()[0];
    const invalid = expectedHeaders.some(
      (header, index) => String(currentHeaders[index]).trim() !== header,
    );

    if (invalid) {
      throw new Error(
        `Revisa los encabezados de ${sheet.getName()}. ` +
          `Deben ser: ${expectedHeaders.join(", ")}.`,
      );
    }
  }

  sheet.setFrozenRows(1);
}

function getDataRows_(sheet, width) {
  const lastRow = sheet.getLastRow();

  return lastRow < 2
    ? []
    : sheet.getRange(2, 1, lastRow - 1, width).getValues();
}

function normalizeBoolean_(value) {
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;

  return ["true", "verdadero", "1", "sí", "si"].includes(
    String(value || "")
      .trim()
      .toLowerCase(),
  );
}

function appendRows_(sheet, rows) {
  if (!rows.length) return;

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}

function formatDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd",
    );
  }

  return String(value || "");
}

function formatTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "HH:mm");
  }

  return "08:00";
}
