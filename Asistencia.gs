function obtenerSesiones() {
  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  validateHeaders_(sheet, HEADERS.SESIONES);

  return getDataRows_(sheet, HEADERS.SESIONES.length)
    .filter((row) => String(row[0]).trim())
    .map((row) => ({
      idSesion: String(row[0]).trim(),
      numero: Number(row[1]),
      fecha: formatDate_(row[2]),
      hora: formatTime_(row[2]),
      tema: String(row[3] || ""),
      estado: String(row[4] || ESTADOS.PROGRAMADA),
    }))
    .sort((a, b) => a.numero - b.numero);
}

function obtenerParticipantes(idSesion) {
  const sesion = validarSesion_(idSesion);
  const participantes = obtenerParticipantesActivos_();
  const asistencias = obtenerAsistenciasPorSesion_(sesion.idSesion);

  return participantes.map((participante) => {
    const estado =
      asistencias.get(participante.idParticipante) || ESTADOS.AUSENTE;

    return {
      idParticipante: participante.idParticipante,
      nombre: participante.nombre,
      correo: participante.correo,
      estado,
      presente: estado === ESTADOS.PRESENTE,
    };
  });
}

function guardarAsistencia(idSesion, participantes) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sesion = validarSesion_(idSesion);
    const registros = validarParticipantes_(participantes);
    const activos = obtenerParticipantesActivos_();
    const idsActivos = new Set(
      activos.map((participante) => participante.idParticipante),
    );

    registros.forEach((registro) => {
      if (!idsActivos.has(registro.idParticipante)) {
        throw new Error(
          `Participante no válido o inactivo: ${registro.idParticipante}`,
        );
      }
    });

    if (registros.length !== activos.length) {
      throw new Error(
        "La lista de participantes cambió. Recarga la sesión antes de guardar.",
      );
    }

    const sheet = getSheet_(SHEET_NAMES.ASISTENCIAS);
    validateHeaders_(sheet, HEADERS.ASISTENCIAS);

    const rows = getDataRows_(sheet, HEADERS.ASISTENCIAS.length);
    const indice = new Map();

    rows.forEach((row, index) => {
      const sesionId = String(row[0]).trim();
      const participanteId = String(row[1]).trim();

      if (!sesionId || !participanteId) return;

      const clave = buildAttendanceKey_(sesionId, participanteId);

      if (indice.has(clave)) {
        throw new Error(
          `Hay una asistencia duplicada en la hoja: ` +
            `${sesionId}, ${participanteId}. Revisa esos registros.`,
        );
      }

      indice.set(clave, index + 2);
    });

    const ahora = new Date();
    const nuevas = [];
    let actualizados = 0;

    registros.forEach((registro) => {
      const clave = buildAttendanceKey_(
        sesion.idSesion,
        registro.idParticipante,
      );

      const datos = [
        sesion.idSesion,
        registro.idParticipante,
        registro.estado,
        ahora,
      ];

      if (indice.has(clave)) {
        sheet.getRange(indice.get(clave), 1, 1, 4).setValues([datos]);
        actualizados += 1;
      } else {
        nuevas.push(datos);
      }
    });

    appendRows_(sheet, nuevas);

    if (sheet.getLastRow() > 1) {
      sheet
        .getRange(2, 4, sheet.getLastRow() - 1, 1)
        .setNumberFormat("yyyy-mm-dd hh:mm:ss");
    }

    getSheet_(SHEET_NAMES.SESIONES)
      .getRange(sesion.numeroFila, 5)
      .setValue(ESTADOS.REGISTRADA);

    SpreadsheetApp.flush();

    return {
      ok: true,
      insertados: nuevas.length,
      actualizados,
      total: registros.length,
    };
  } finally {
    lock.releaseLock();
  }
}

function validarSesion_(idSesion) {
  if (typeof idSesion !== "string" || !idSesion.trim()) {
    throw new Error("Selecciona una sesión válida.");
  }

  const id = idSesion.trim();
  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  validateHeaders_(sheet, HEADERS.SESIONES);

  const rows = getDataRows_(sheet, HEADERS.SESIONES.length);
  const coincidencias = [];

  rows.forEach((row, index) => {
    if (String(row[0]).trim() === id) {
      coincidencias.push({
        idSesion: id,
        numeroFila: index + 2,
        estado: String(row[4] || ESTADOS.PROGRAMADA),
      });
    }
  });

  if (!coincidencias.length) {
    throw new Error(`No existe la sesión ${id}.`);
  }

  if (coincidencias.length > 1) {
    throw new Error(`La sesión ${id} está duplicada en Google Sheets.`);
  }

  return coincidencias[0];
}

function validarParticipantes_(participantes) {
  if (!Array.isArray(participantes) || !participantes.length) {
    throw new Error("No hay participantes para guardar.");
  }

  const ids = new Set();
  const permitidos = new Set([
    ESTADOS.PRESENTE,
    ESTADOS.AUSENTE,
    ESTADOS.JUSTIFICADA,
    ESTADOS.SUSPENDIDA,
    ESTADOS.RETARDO,
  ]);

  return participantes.map((participante) => {
    if (!participante || typeof participante !== "object") {
      throw new Error("Cada registro debe ser un objeto.");
    }

    const idParticipante = String(participante.idParticipante || "").trim();

    if (!idParticipante) {
      throw new Error("Falta el ID de un participante.");
    }

    if (ids.has(idParticipante)) {
      throw new Error(`Participante repetido: ${idParticipante}`);
    }

    let estado;

    if (participante.estado !== undefined && participante.estado !== null) {
      estado = String(participante.estado).trim().toUpperCase();
    } else if (typeof participante.presente === "boolean") {
      estado = participante.presente ? ESTADOS.PRESENTE : ESTADOS.AUSENTE;
    } else {
      throw new Error(`Falta el estado de ${idParticipante}.`);
    }

    if (!permitidos.has(estado)) {
      throw new Error(`Estado no válido para ${idParticipante}.`);
    }

    ids.add(idParticipante);
    return { idParticipante, estado };
  });
}

function obtenerParticipantesActivos_() {
  const sheet = getSheet_(SHEET_NAMES.PARTICIPANTES);
  validateHeaders_(sheet, HEADERS.PARTICIPANTES);

  const ids = new Set();
  const participantes = [];

  getDataRows_(sheet, HEADERS.PARTICIPANTES.length).forEach((row) => {
    const idParticipante = String(row[0]).trim();
    if (!idParticipante) return;

    if (ids.has(idParticipante)) {
      throw new Error(`ID duplicado en Participantes: ${idParticipante}`);
    }

    ids.add(idParticipante);

    if (normalizeBoolean_(row[3])) {
      participantes.push({
        idParticipante,
        nombre: String(row[1] || ""),
        correo: String(row[2] || ""),
      });
    }
  });

  return participantes;
}

function obtenerAsistenciasPorSesion_(idSesion) {
  const sheet = getSheet_(SHEET_NAMES.ASISTENCIAS);
  validateHeaders_(sheet, HEADERS.ASISTENCIAS);

  const resultado = new Map();

  getDataRows_(sheet, HEADERS.ASISTENCIAS.length).forEach((row) => {
    if (String(row[0]).trim() !== idSesion || !row[1]) return;

    const id = String(row[1]).trim();

    if (resultado.has(id)) {
      throw new Error(`Asistencia duplicada para ${id} en ${idSesion}.`);
    }

    resultado.set(
      id,
      String(row[2] || ESTADOS.AUSENTE)
        .trim()
        .toUpperCase(),
    );
  });

  return resultado;
}

function buildAttendanceKey_(idSesion, idParticipante) {
  return `${idSesion}::${idParticipante}`;
}

function programarSesion(idSesion, fecha, hora) {
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(String(fecha)) ||
    !/^\d{2}:\d{2}$/.test(String(hora))
  ) {
    throw new Error("Selecciona una fecha y una hora válidas.");
  }

  const zona = Session.getScriptTimeZone();
  const texto = `${fecha} ${hora}`;
  let programada;

  try {
    programada = Utilities.parseDate(texto, zona, "yyyy-MM-dd HH:mm");
  } catch (error) {
    throw new Error("La fecha o la hora no es válida.");
  }

  if (Utilities.formatDate(programada, zona, "yyyy-MM-dd HH:mm") !== texto) {
    throw new Error("La fecha o la hora está fuera de rango.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sesion = validarSesion_(idSesion);
    const sheet = getSheet_(SHEET_NAMES.SESIONES);

    sheet
      .getRange(sesion.numeroFila, 3)
      .setValue(programada)
      .setNumberFormat("yyyy-mm-dd hh:mm");

    // Cambiar el horario conserva la asistencia ya registrada.
    SpreadsheetApp.flush();

    return {
      ok: true,
      fecha: formatDate_(programada),
      hora: formatTime_(programada),
      estado: sesion.estado,
    };
  } finally {
    lock.releaseLock();
  }
}
