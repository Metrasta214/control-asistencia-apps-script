function prepararBaseDatos() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getSpreadsheet_();

    const participantesSheet = getOrCreateSheet_(
      spreadsheet,
      SHEET_NAMES.PARTICIPANTES,
      HEADERS.PARTICIPANTES,
    );

    const sesionesSheet = getOrCreateSheet_(
      spreadsheet,
      SHEET_NAMES.SESIONES,
      HEADERS.SESIONES,
    );

    const asistenciasSheet = getOrCreateSheet_(
      spreadsheet,
      SHEET_NAMES.ASISTENCIAS,
      HEADERS.ASISTENCIAS,
    );

    const participantesAgregados = seedParticipantes_(participantesSheet);
    const sesionesAgregadas = seedSesiones_(sesionesSheet);

    [participantesSheet, sesionesSheet, asistenciasSheet].forEach((sheet) => {
      sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .setFontWeight("bold")
        .setBackground("#e8eefc");

      sheet.autoResizeColumns(1, sheet.getLastColumn());
    });

    if (sesionesSheet.getLastRow() > 1) {
      sesionesSheet
        .getRange(2, 3, sesionesSheet.getLastRow() - 1, 1)
        .setNumberFormat("yyyy-mm-dd hh:mm");
    }

    SpreadsheetApp.flush();

    const resultado = {
      ok: true,
      participantesAgregados,
      sesionesAgregadas,
      url: spreadsheet.getUrl(),
    };

    console.log(JSON.stringify(resultado));
    return resultado;
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSheet_(spreadsheet, sheetName, headers) {
  const sheet =
    spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  validateHeaders_(sheet, headers);
  return sheet;
}

function seedParticipantes_(sheet) {
  const nombres = [
    ["Mariana", "López Hernández"],
    ["Santiago", "Ramírez Torres"],
    ["Valentina", "García Mendoza"],
    ["Mateo", "Hernández Cruz"],
    ["Regina", "Martínez Flores"],
    ["Sebastián", "González Vargas"],
    ["Camila", "Pérez Navarro"],
    ["Emiliano", "Sánchez Reyes"],
    ["Ximena", "Morales Castillo"],
    ["Leonardo", "Jiménez Ortega"],
    ["Sofía", "Ruiz Domínguez"],
    ["Diego", "Torres Salazar"],
    ["Renata", "Vázquez Medina"],
    ["Daniel", "Mendoza Silva"],
    ["Natalia", "Castro Rojas"],
    ["Alejandro", "Ortiz Guerrero"],
    ["Fernanda", "Ramos Cabrera"],
    ["Nicolás", "Márquez León"],
    ["Lucía", "Vega Fuentes"],
    ["Andrés", "Pineda Campos"],
    ["Isabella", "Núñez Estrada"],
    ["Joaquín", "Mora Ibarra"],
    ["Ana", "Cervantes Soto"],
    ["Pablo", "Valdez Luna"],
    ["Emilia", "Esparza Molina"],
    ["Rodrigo", "Acosta Beltrán"],
    ["Paula", "Miranda Solís"],
    ["Gael", "Franco Bautista"],
    ["Daniela", "Salinas Ponce"],
    ["Marco", "Treviño Arias"],
  ];

  const rows = getDataRows_(sheet, HEADERS.PARTICIPANTES.length);
  const existing = new Map();

  rows.forEach((row, index) => {
    const id = String(row[0]).trim();
    if (!id) return;

    if (existing.has(id)) {
      throw new Error(`ID duplicado en Participantes: ${id}`);
    }

    existing.set(id, { row, numeroFila: index + 2 });
  });

  const nuevas = [];

  nombres.forEach((persona, index) => {
    const id = `P${String(index + 1).padStart(2, "0")}`;
    const nombre = persona.join(" ");
    const correo = `${id.toLowerCase()}@example.com`;
    const actual = existing.get(id);

    if (!actual) {
      nuevas.push([id, nombre, correo, true]);
    } else if (/^participante\s+\d+$/i.test(String(actual.row[1]).trim())) {
      sheet.getRange(actual.numeroFila, 2, 1, 2).setValues([[nombre, correo]]);
    }
  });

  appendRows_(sheet, nuevas);
  return nuevas.length;
}

function seedSesiones_(sheet) {
  const temas = [
    "Introducción a la IA",
    "Aprendizaje automático",
    "Datos y modelos",
    "Redes neuronales",
    "Ética tecnológica",
    "Taller de prompts",
    "Proyecto integrador",
    "Presentación de avances",
    "Evaluación práctica",
    "Cierre del curso",
  ];

  const rows = getDataRows_(sheet, HEADERS.SESIONES.length);
  const existingIds = new Set();

  rows.forEach((row) => {
    const id = String(row[0]).trim();
    if (!id) return;

    if (existingIds.has(id)) {
      throw new Error(`ID duplicado en Sesiones: ${id}`);
    }

    existingIds.add(id);
  });

  // Si ya hay sesiones, usa una de ellas como referencia semanal.
  const referencia = rows.find(
    (row) =>
      row[2] instanceof Date && !isNaN(row[2].getTime()) && Number(row[1]) >= 1,
  );

  const inicio = referencia ? new Date(referencia[2]) : new Date();

  if (referencia) {
    inicio.setDate(inicio.getDate() - (Number(referencia[1]) - 1) * 7);
  }

  inicio.setHours(8, 0, 0, 0);

  const nuevas = [];

  for (let index = 0; index < 10; index += 1) {
    const id = `S${String(index + 1).padStart(2, "0")}`;

    if (!existingIds.has(id)) {
      const fecha = new Date(inicio);
      fecha.setDate(inicio.getDate() + index * 7);

      nuevas.push([id, index + 1, fecha, temas[index], ESTADOS.PROGRAMADA]);
    }
  }

  appendRows_(sheet, nuevas);
  return nuevas.length;
}
