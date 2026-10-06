function obtenerSesiones() {
  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  validateHeaders_(sheet, HEADERS.SESIONES);

  return getDataRows_(sheet, HEADERS.SESIONES.length)
    .filter((row) => row[0])
    .map((row) => ({
      idSesion: String(row[0]),
      numero: Number(row[1]),
      fecha: formatDate_(row[2]),
      hora: formatTime_(row[2]),
      tema: String(row[3] || ''),
      estado: String(row[4] || ESTADOS.PROGRAMADA),
    }))
    .sort((a, b) => a.numero - b.numero);
}

function obtenerParticipantes(idSesion) {
  validarSesion_(idSesion);

  const participantes = obtenerParticipantesActivos_();
  const asistencias = obtenerAsistenciasPorSesion_(idSesion);

  return participantes.map((participante) => {
    const estado = asistencias.get(participante.idParticipante) || ESTADOS.AUSENTE;

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
  let lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;
    validarSesion_(idSesion);
    const registros = validarParticipantes_(participantes);
    const participantesValidos = obtenerParticipantesActivos_();
    const participantesMap = new Map(participantesValidos.map((participante) => [participante.idParticipante, participante]));

    registros.forEach((registro) => {
      if (!participantesMap.has(registro.idParticipante)) {
        throw new Error(`Participante no valido o inactivo: ${registro.idParticipante}`);
      }
    });

    const sheet = getSheet_(SHEET_NAMES.ASISTENCIAS);
    validateHeaders_(sheet, HEADERS.ASISTENCIAS);

    const existingRows = getDataRows_(sheet, HEADERS.ASISTENCIAS.length);
    const existingIndexByKey = new Map();

    existingRows.forEach((row, index) => {
      const rowSessionId = String(row[0]).trim();
      const rowParticipantId = String(row[1]).trim();

      if (rowSessionId && rowParticipantId) {
        existingIndexByKey.set(buildAttendanceKey_(rowSessionId, rowParticipantId), index + 2);
      }
    });

    const now = new Date();
    const rowsToInsert = [];
    let updated = 0;

    registros.forEach((registro) => {
      const estado = registro.estado || (registro.presente ? ESTADOS.PRESENTE : ESTADOS.AUSENTE);
      const row = [idSesion, registro.idParticipante, estado, now];
      const existingRow = existingIndexByKey.get(buildAttendanceKey_(idSesion, registro.idParticipante));

      if (existingRow) {
        sheet.getRange(existingRow, 1, 1, HEADERS.ASISTENCIAS.length).setValues([row]);
        updated += 1;
      } else {
        rowsToInsert.push(row);
      }
    });

    appendRows_(sheet, rowsToInsert);
    marcarSesionRegistrada_(idSesion);

    return {
      ok: true,
      insertados: rowsToInsert.length,
      actualizados: updated,
      total: registros.length,
    };
  } catch (error) {
    throw new Error(error.message || 'No se pudo guardar la asistencia.');
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function validarSesion_(idSesion) {
  if (!idSesion || typeof idSesion !== 'string') {
    throw new Error('Selecciona una sesion valida.');
  }

  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  validateHeaders_(sheet, HEADERS.SESIONES);
  const exists = getDataRows_(sheet, HEADERS.SESIONES.length).some((row) => String(row[0]) === idSesion);

  if (!exists) {
    throw new Error(`No existe la sesion ${idSesion}.`);
  }
}

function validarParticipantes_(participantes) {
  if (!Array.isArray(participantes)) {
    throw new Error('La asistencia debe enviarse como una lista de participantes.');
  }

  if (participantes.length === 0) {
    throw new Error('No hay participantes para guardar.');
  }

  const ids = new Set();

  return participantes.map((participante) => {
    if (!participante || typeof participante !== 'object') {
      throw new Error('Cada registro de asistencia debe ser un objeto.');
    }

    const idParticipante = String(participante.idParticipante || '').trim();

    if (!idParticipante) {
      throw new Error('Cada registro requiere idParticipante.');
    }

    if (ids.has(idParticipante)) {
      throw new Error(`Participante repetido en el envio: ${idParticipante}`);
    }

    const estado = String(participante.estado || '').trim().toUpperCase();
    const estadoValido = [ESTADOS.PRESENTE, ESTADOS.AUSENTE, ESTADOS.JUSTIFICADA, ESTADOS.SUSPENDIDA, ESTADOS.RETARDO].includes(estado);

    if (!estadoValido && typeof participante.presente !== 'boolean') {
      throw new Error(`El estado de asistencia no es valido para ${idParticipante}.`);
    }

    ids.add(idParticipante);

    return {
      idParticipante,
      estado: estadoValido ? estado : (participante.presente ? ESTADOS.PRESENTE : ESTADOS.AUSENTE),
    };
  });
}

function obtenerParticipantesActivos_() {
  const sheet = getSheet_(SHEET_NAMES.PARTICIPANTES);
  validateHeaders_(sheet, HEADERS.PARTICIPANTES);

  return getDataRows_(sheet, HEADERS.PARTICIPANTES.length)
    .filter((row) => row[0] && normalizeBoolean_(row[3]))
    .map((row) => ({
      idParticipante: String(row[0]),
      nombre: String(row[1] || ''),
      correo: String(row[2] || ''),
      activo: normalizeBoolean_(row[3]),
    }));
}

function obtenerAsistenciasPorSesion_(idSesion) {
  const sheet = getSheet_(SHEET_NAMES.ASISTENCIAS);
  validateHeaders_(sheet, HEADERS.ASISTENCIAS);
  const asistencias = new Map();

  getDataRows_(sheet, HEADERS.ASISTENCIAS.length).forEach((row) => {
    if (String(row[0]) === idSesion && row[1]) {
      asistencias.set(String(row[1]), String(row[2] || ESTADOS.AUSENTE));
    }
  });

  return asistencias;
}

function marcarSesionRegistrada_(idSesion) {
  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  const rows = getDataRows_(sheet, HEADERS.SESIONES.length);

  for (let index = 0; index < rows.length; index += 1) {
    if (String(rows[index][0]) === idSesion) {
      sheet.getRange(index + 2, 5).setValue(ESTADOS.REGISTRADA);
      return;
    }
  }
}

function buildAttendanceKey_(idSesion, idParticipante) {
  return `${idSesion}::${idParticipante}`;
}

function formatDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '');
}

function formatTime_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }

  return '08:00';
}

function programarSesion(idSesion, fecha, hora) {
  validarSesion_(idSesion);

  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha)) || !/^\d{2}:\d{2}$/.test(String(hora))) {
    throw new Error('Selecciona una fecha y una hora validas.');
  }

  const parts = String(fecha).split('-').map(Number);
  const timeParts = String(hora).split(':').map(Number);
  const scheduled = new Date(parts[0], parts[1] - 1, parts[2], timeParts[0], timeParts[1], 0, 0);
  const sheet = getSheet_(SHEET_NAMES.SESIONES);
  const rows = getDataRows_(sheet, HEADERS.SESIONES.length);

  for (let index = 0; index < rows.length; index += 1) {
    if (String(rows[index][0]) === idSesion) {
      sheet.getRange(index + 2, 3).setValue(scheduled);
      sheet.getRange(index + 2, 5).setValue(ESTADOS.PROGRAMADA);
      return { ok: true, fecha: formatDate_(scheduled), hora: formatTime_(scheduled) };
    }
  }

  throw new Error(`No existe la sesion ${idSesion}.`);
}
