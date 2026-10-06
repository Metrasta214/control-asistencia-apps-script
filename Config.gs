const SPREADSHEET_ID = "1GBuWaHHUpZ4xOI9gIfxR-RRJlCF9fhIn4SjzVZwlqdM";

const SHEET_NAMES = {
  PARTICIPANTES: "Participantes",
  SESIONES: "Sesiones",
  ASISTENCIAS: "Asistencias",
};

const HEADERS = {
  PARTICIPANTES: ["id_participante", "nombre", "correo", "activo"],
  SESIONES: ["id_sesion", "numero", "fecha", "tema", "estado"],
  ASISTENCIAS: ["id_sesion", "id_participante", "estado", "hora_registro"],
};

const ESTADOS = {
  PRESENTE: "PRESENTE",
  AUSENTE: "AUSENTE",
  JUSTIFICADA: "JUSTIFICADA",
  SUSPENDIDA: "SUSPENDIDA",
  RETARDO: "RETARDO",
  PROGRAMADA: "PROGRAMADA",
  REGISTRADA: "REGISTRADA",
};

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
  const currentHeaders = range.getValues()[0];
  const emptyHeaders = currentHeaders.every((header) => header === "");

  if (sheet.getLastRow() === 0 || emptyHeaders) {
    range.setValues([expectedHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const hasInvalidHeaders = expectedHeaders.some(
    (header, index) => currentHeaders[index] !== header,
  );

  if (hasInvalidHeaders) {
    throw new Error(
      `Encabezados invalidos en ${sheet.getName()}. Se esperaba: ${expectedHeaders.join(", ")}.`,
    );
  }

  sheet.setFrozenRows(1);
}

function getDataRows_(sheet, width) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, width).getValues();
}

function normalizeBoolean_(value) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    return value.toLowerCase() === "true";
  }

  return Boolean(value);
}
