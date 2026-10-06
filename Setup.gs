function prepararBaseDatos() {
  const spreadsheet = getSpreadsheet_();
  const participantesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.PARTICIPANTES, HEADERS.PARTICIPANTES);
  const sesionesSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.SESIONES, HEADERS.SESIONES);
  const asistenciasSheet = getOrCreateSheet_(spreadsheet, SHEET_NAMES.ASISTENCIAS, HEADERS.ASISTENCIAS);

  validateHeaders_(participantesSheet, HEADERS.PARTICIPANTES);
  validateHeaders_(sesionesSheet, HEADERS.SESIONES);
  validateHeaders_(asistenciasSheet, HEADERS.ASISTENCIAS);

  const participantesAgregados = seedParticipantes_(participantesSheet);
  const sesionesAgregadas = seedSesiones_(sesionesSheet);

  return {
    ok: true,
    participantesAgregados,
    sesionesAgregadas,
  };
}

function getOrCreateSheet_(spreadsheet, sheetName, headers) {
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  validateHeaders_(sheet, headers);
  return sheet;
}

function seedParticipantes_(sheet) {
  const rows = getDataRows_(sheet, HEADERS.PARTICIPANTES.length);
  const nombres = [
    ['Mariana', 'López Hernández'], ['Santiago', 'Ramírez Torres'], ['Valentina', 'García Mendoza'],
    ['Mateo', 'Hernández Cruz'], ['Regina', 'Martínez Flores'], ['Sebastián', 'González Vargas'],
    ['Camila', 'Pérez Navarro'], ['Emiliano', 'Sánchez Reyes'], ['Ximena', 'Morales Castillo'],
    ['Leonardo', 'Jiménez Ortega'], ['Sofía', 'Ruiz Domínguez'], ['Diego', 'Torres Salazar'],
    ['Renata', 'Vázquez Medina'], ['Daniel', 'Mendoza Silva'], ['Natalia', 'Castro Rojas'],
    ['Alejandro', 'Ortiz Guerrero'], ['Fernanda', 'Ramos Cabrera'], ['Nicolás', 'Márquez León'],
    ['Lucía', 'Vega Fuentes'], ['Andrés', 'Pineda Campos'], ['Isabella', 'Núñez Estrada'],
    ['Joaquín', 'Mora Ibarra'], ['Ana', 'Cervantes Soto'], ['Pablo', 'Valdez Luna'],
    ['Emilia', 'Esparza Molina'], ['Rodrigo', 'Acosta Beltrán'], ['Paula', 'Miranda Solís'],
    ['Gael', 'Franco Bautista'], ['Daniela', 'Salinas Ponce'], ['Marco', 'Treviño Arias'],
  ];
  const existingIds = new Set(rows.map((row) => String(row[0]).trim()).filter(Boolean));
  const newRows = [];

  rows.forEach((row, index) => {
    if (String(row[1]).toLowerCase().startsWith('participante ') && nombres[index]) {
      const fullName = nombres[index].join(' ');
      const email = `${nombres[index][0].toLowerCase()}.${nombres[index][1].split(' ')[0].toLowerCase()}@escuela.edu.mx`;
      sheet.getRange(index + 2, 2, 1, 2).setValues([[fullName, email]]);
    }
  });

  for (let index = 1; index <= 30; index += 1) {
    const id = `P${String(index).padStart(2, '0')}`;

    if (!existingIds.has(id)) {
      newRows.push([
        id,
        nombres[index - 1].join(' '),
        `${nombres[index - 1][0].toLowerCase()}.${nombres[index - 1][1].split(' ')[0].toLowerCase()}@escuela.edu.mx`,
        true,
      ]);
    }
  }

  appendRows_(sheet, newRows);
  return newRows.length;
}

function seedSesiones_(sheet) {
  const rows = getDataRows_(sheet, HEADERS.SESIONES.length);
  const existingIds = new Set(rows.map((row) => String(row[0]).trim()).filter(Boolean));
  const newRows = [];
  const startDate = new Date();

  startDate.setHours(0, 0, 0, 0);

  for (let index = 1; index <= 10; index += 1) {
    const id = `S${String(index).padStart(2, '0')}`;

    if (!existingIds.has(id)) {
      const sessionDate = new Date(startDate);
      sessionDate.setDate(startDate.getDate() + (index - 1) * 7);
      sessionDate.setHours(8 + ((index - 1) % 4), 0, 0, 0);

      newRows.push([
        id,
        index,
        sessionDate,
        ['Introducción a la IA', 'Aprendizaje automático', 'Datos y modelos', 'Redes neuronales', 'Ética tecnológica', 'Taller de prompts', 'Proyecto integrador', 'Presentación de avances', 'Evaluación práctica', 'Cierre del curso'][index - 1],
        ESTADOS.PROGRAMADA,
      ]);
    }
  }

  appendRows_(sheet, newRows);
  return newRows.length;
}

function appendRows_(sheet, rows) {
  if (rows.length === 0) {
    return;
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
