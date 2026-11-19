/**
 * ============================================================================
 * BACKEND GOOGLE APPS SCRIPT - AGRÍCOLA HUARMEY (MÓDULO CAPORALES)
 * Archivo: Codigo.gs
 * Descripción: API RESTful para consulta de catálogos maestros (con caché) y
 * recepción por lotes de registros de caporales con LockService e Idempotencia.
 * ============================================================================
 */

// CONFIGURACIÓN PRINCIPAL (Reemplazar con el ID real del Google Sheets)
const CONFIG = {
  SPREADSHEET_ID: 'https://script.google.com/macros/s/AKfycbyROhN-dK-8Xg4K2q5VBOS90YqLUecINScOs5Hsaiqs6cmVUndtVtLi-Tw4MBzrRcHO/exec',
  API_SECRET_TOKEN: '', // Opcional: Si se define, el frontend debe enviarlo en el header/body
  SHEETS: {
    MODULOS: 'maestro-modulos',
    ACOPIADOR: 'maestro-acopiador',
    CAPORAL: 'maestro-caporal',
    REGISTRO: 'registro-caporal',
    AUDITORIA_UUID: '_sys_uuids_caporal' // Hoja técnica oculta para garantizar idempotencia
  },
  HEADERS: {
    MODULOS: 'MODULO - TURNO',
    ACOPIADOR: 'DATOS ACOPIADOR',
    CAPORAL: 'DATOS CAPORAL'
  },
  SUPERVISORES_DEFAULT: ['Frank Farroñan', 'Roni Cubas'],
  MODALIDADES_DEFAULT: ['CAMPAMENTO', 'IDA Y VUELTA'],
  CACHE_KEY_MAESTROS: 'HUARMEY_MAESTROS_CAPORAL_V1',
  CACHE_TTL_SECONDS: 21600, // 6 horas (máximo permitido por CacheService)
  MAX_BATCH_SIZE: 50
};

/**
 * Endpoint GET: Consulta de datos maestros con caché compartida para 200 usuarios
 */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getMaestros';

    if (action === 'ping') {
      return jsonResponse({ status: 'ok', message: 'API Caporales Agrícola Huarmey operativa', timestamp: new Date().toISOString() });
    }

    if (action === 'getMaestros') {
      const cache = CacheService.getScriptCache();
      const cachedJson = cache.get(CONFIG.CACHE_KEY_MAESTROS);

      // Responder en ~40ms desde memoria RAM sin abrir Google Sheets
      if (cachedJson) {
        return ContentService
          .createTextOutput(cachedJson)
          .setMimeType(ContentService.MimeType.JSON);
      }

      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const data = {
        supervisores: CONFIG.SUPERVISORES_DEFAULT,
        modalidades: CONFIG.MODALIDADES_DEFAULT,
        modulos: extractColumnValues(ss, CONFIG.SHEETS.MODULOS, CONFIG.HEADERS.MODULOS),
        acopiadores: extractColumnValues(ss, CONFIG.SHEETS.ACOPIADOR, CONFIG.HEADERS.ACOPIADOR),
        caporales: extractColumnValues(ss, CONFIG.SHEETS.CAPORAL, CONFIG.HEADERS.CAPORAL),
        serverTimestamp: new Date().toISOString()
      };

      const payloadString = JSON.stringify({ status: 'ok', data: data });
      cache.put(CONFIG.CACHE_KEY_MAESTROS, payloadString, CONFIG.CACHE_TTL_SECONDS);

      return ContentService
        .createTextOutput(payloadString)
        .setMimeType(ContentService.MimeType.JSON);
    }

    return jsonResponse({ status: 'error', message: 'Acción GET no soportada.' });
  } catch (err) {
    console.error('Error en doGet:', err);
    return jsonResponse({ status: 'error', message: err.message || 'Error interno al consultar los datos maestros.' });
  }
}

/**
 * Endpoint POST: Recepción por lotes (Batch) con LockService e Idempotencia
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'error', message: 'Cuerpo de solicitud vacío.' });
    }

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse({ status: 'error', message: 'JSON mal formado.' });
    }

    // Validación opcional de token de seguridad
    if (CONFIG.API_SECRET_TOKEN && payload.token !== CONFIG.API_SECRET_TOKEN) {
      return jsonResponse({ status: 'error', message: 'Token de autorización inválido.' });
    }

    if (payload.action !== 'syncBatch' || !Array.isArray(payload.records)) {
      return jsonResponse({ status: 'error', message: 'Formato de lote inválido.' });
    }

    if (payload.records.length === 0) {
      return jsonResponse({ status: 'ok', results: [] });
    }

    if (payload.records.length > CONFIG.MAX_BATCH_SIZE) {
      return jsonResponse({ status: 'error', message: `El lote excede el máximo de ${CONFIG.MAX_BATCH_SIZE} registros.` });
    }

    const results = processBatchWithLock(payload.records);
    return jsonResponse({ status: 'ok', results: results });

  } catch (err) {
    console.error('Error crítico en doPost:', err);
    return jsonResponse({ status: 'error', message: err.message || 'Error al procesar registros en el servidor.' });
  }
}

/**
 * Extrae una columna específica por encabezado (tolerante a tildes, guiones, espacios y filas superiores)
 */
function extractColumnValues(ss, sheetName, headerName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error(`No existe la hoja: ${sheetName}`);
  
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const normalize = (str) => String(str || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[–—]/g, '-')
    .replace(/\s*-\s*/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();

  const targetHeader = normalize(headerName);
  let headerRowIndex = -1;
  let colIndex = -1;

  const maxSearchRows = Math.min(values.length, 10);
  for (let r = 0; r < maxSearchRows; r++) {
    const normalizedRow = values[r].map(normalize);
    const idx = normalizedRow.indexOf(targetHeader);
    if (idx !== -1) {
      headerRowIndex = r;
      colIndex = idx;
      break;
    }
  }

  if (colIndex === -1) {
    const foundRow1 = values[0].filter(c => String(c).trim() !== '').join(', ') || '(Fila 1 vacía)';
    throw new Error(`No se encontró "${headerName}" en ${sheetName}. Encabezados detectados en fila 1: [${foundRow1}]`);
  }

  const uniqueSet = new Set();
  for (let i = headerRowIndex + 1; i < values.length; i++) {
    const cell = String(values[i][colIndex] || '').trim();
    if (cell !== '') {
      uniqueSet.add(cell);
    }
  }
  return Array.from(uniqueSet);
}

/**
 * Procesa el lote bajo exclusión mutua (LockService) garantizando idempotencia y orden exacto de 7 columnas
 */
function processBatchWithLock(records) {
  const lock = LockService.getScriptLock();
  const acquired = lock.tryLock(25000);
  
  if (!acquired) {
    throw new Error('Servidor ocupado por alta concurrencia. Reintentando automáticamente.');
  }

  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const targetSheet = ss.getSheetByName(CONFIG.SHEETS.REGISTRO);
    if (!targetSheet) throw new Error(`Hoja destino no encontrada: ${CONFIG.SHEETS.REGISTRO}`);

    let uuidSheet = ss.getSheetByName(CONFIG.SHEETS.AUDITORIA_UUID);
    if (!uuidSheet) {
      uuidSheet = ss.insertSheet(CONFIG.SHEETS.AUDITORIA_UUID);
      uuidSheet.appendRow(['UUID', 'TIMESTAMP_SERVIDOR']);
      uuidSheet.hideSheet();
    }

    const existingData = uuidSheet.getDataRange().getValues();
    const processedUuids = new Set();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][0]) processedUuids.add(String(existingData[i][0]).trim());
    }

    const rowsToInsert = [];
    const uuidsToLog = [];
    const results = [];
    const nowIso = new Date().toISOString();

    for (const rec of records) {
      const uuid = String(rec.uuid || '').trim();
      
      if (!uuid || uuid.length < 10) {
        results.push({ uuid: uuid, status: 'INVALID', message: 'UUID inválido o ausente.' });
        continue;
      }
      if (processedUuids.has(uuid)) {
        results.push({ uuid: uuid, status: 'DUPLICATE_IGNORED', message: 'Registro ya procesado previamente.' });
        continue;
      }

      const validation = validateRecord(rec);
      if (!validation.valid) {
        results.push({ uuid: uuid, status: 'INVALID', message: validation.message });
        continue;
      }

      // Orden EXACTO de 7 columnas (A hasta G) en registro-caporal:
      // Col A: FECHA REG. | Col B: SUPERVISOR | Col C: MODULO -TURNO | Col D: DATOS ACOPIADOR | Col E: DATOS CAPORAL | Col F: MODALIDAD | Col G: TOTAL
      rowsToInsert.push([
        sanitizeText(rec.fechaRegistro),
        sanitizeText(rec.supervisor),
        sanitizeText(rec.moduloTurno),
        sanitizeText(rec.datosAcopiador),
        sanitizeText(rec.datosCaporal),
        sanitizeText(rec.modalidad),
        Number(rec.total)
      ]);

      uuidsToLog.push([uuid, nowIso]);
      processedUuids.add(uuid);
      results.push({ uuid: uuid, status: 'SYNCED' });
    }

    // Escritura Atómica en Bloque (setValues) en las 7 columnas (A:G)
    if (rowsToInsert.length > 0) {
      const startRow = targetSheet.getLastRow() + 1;
      targetSheet.getRange(startRow, 1, rowsToInsert.length, 7).setValues(rowsToInsert);

      const uuidStartRow = uuidSheet.getLastRow() + 1;
      uuidSheet.getRange(uuidStartRow, 1, uuidsToLog.length, 2).setValues(uuidsToLog);
      
      SpreadsheetApp.flush();
    }

    return results;

  } finally {
    lock.releaseLock();
  }
}

/**
 * Valida reglas de negocio en servidor
 */
function validateRecord(rec) {
  if (!rec.fechaRegistro || !/^\d{4}-\d{2}-\d{2}$/.test(String(rec.fechaRegistro).trim())) {
    return { valid: false, message: 'Formato de FECHA REG. inválido (esperado YYYY-MM-DD).' };
  }
  if (!rec.supervisor || String(rec.supervisor).trim().length > 100) {
    return { valid: false, message: 'SUPERVISOR es obligatorio o demasiado largo.' };
  }
  if (!rec.moduloTurno || String(rec.moduloTurno).trim().length > 120) {
    return { valid: false, message: 'MODULO - TURNO inválido o demasiado largo.' };
  }
  if (!rec.datosAcopiador || String(rec.datosAcopiador).trim().length > 150) {
    return { valid: false, message: 'DATOS ACOPIADOR inválido o demasiado largo.' };
  }
  if (!rec.datosCaporal || String(rec.datosCaporal).trim().length > 150) {
    return { valid: false, message: 'DATOS CAPORAL inválido o demasiado largo.' };
  }
  if (!rec.modalidad || !['CAMPAMENTO', 'IDA Y VUELTA'].includes(String(rec.modalidad).trim().toUpperCase())) {
    return { valid: false, message: 'MODALIDAD debe ser CAMPAMENTO o IDA Y VUELTA.' };
  }
  
  const total = Number(rec.total);
  if (!Number.isInteger(total) || total <= 0 || total > 100000) {
    return { valid: false, message: 'TOTAL debe ser un entero positivo válido.' };
  }

  return { valid: true };
}

/**
 * Previene inyección de fórmulas en Google Sheets (=, +, -, @)
 */
function sanitizeText(val) {
  let str = String(val || '').trim();
  if (/^[=+\-@]/.test(str)) {
    str = "'" + str;
  }
  return str;
}

/**
 * Helper para respuestas JSON estándar
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Función auxiliar para limpiar caché manualmente si actualizas los maestros en Sheets
 */
function limpiarCacheMaestros() {
  CacheService.getScriptCache().remove(CONFIG.CACHE_KEY_MAESTROS);
  Logger.log('Caché de maestros limpiada exitosamente.');
}
