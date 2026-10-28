/**
 * Página de mantenciones - Apps Script
 * Paso 1: planilla SAMS
 *
 * Para agregar otra planilla más adelante, solo se suma un bloque en HOJAS.
 */

const HOJAS = {
  sams: {
    nombre: 'SAMS',
    archivoId: '1s25_K3_onbvd27gRGesywnS5V7Ty72nI4Kgqw9MxnKI',
    pestana: 'SAMS',
    filaEncabezado: 1,
    columnaClave: 'Centro',  // se usa para confirmar que la fila no cambió de lugar
    historial: 'historial'   // al desactivar un centro, su periodo pasa a esta hoja
  },
  // Periodos terminados: un mismo centro puede aparecer varias veces, por eso la clave es el Código
  historial: {
    nombre: 'Historial',
    archivoId: '1s25_K3_onbvd27gRGesywnS5V7Ty72nI4Kgqw9MxnKI',
    pestana: 'HISTORIAL',
    filaEncabezado: 1,
    columnaClave: 'Código'
  }
};

// ---------- Puntos de entrada ----------

function doGet() {
  return responder({ ok: false, error: 'No autorizado' });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const secreto = PropertiesService.getScriptProperties().getProperty('SECRETO');

    if (!secreto || body.secreto !== secreto) {
      return responder({ ok: false, error: 'No autorizado' });
    }

    const config = HOJAS[body.hoja];
    if (!config) return responder({ ok: false, error: 'Hoja no permitida' });

    switch (body.accion) {
      case 'leer':    return responder(leer(config));
      case 'editar':  return responder(editar(config, body));
      case 'agregar':  return responder(agregar(config, body));
      case 'eliminar': return responder(eliminar(config, body));
      case 'desactivar': return responder(desactivar(config, body));
      default:         return responder({ ok: false, error: 'Acción no válida' });
    }
  } catch (err) {
    return responder({ ok: false, error: err.message });
  }
}

// ---------- Acciones ----------

function leer(config) {
  const sheet = abrirPestana(config);
  const rango = sheet.getDataRange();
  const valores = rango.getDisplayValues();
  const formulas = rango.getFormulas();
  const inicioDatos = config.filaEncabezado;
  const encabezados = valores[config.filaEncabezado - 1].map(h => String(h).trim());

  const columnas = [];
  encabezados.forEach((nombre, i) => {
    if (!nombre) return;
    const tieneFormula = formulas.slice(inicioDatos).some(f => f[i] !== '');
    columnas.push({ nombre: nombre, indice: i, editable: !tieneFormula });
  });

  const filas = [];
  for (let r = inicioDatos; r < valores.length; r++) {
    const fila = valores[r];
    if (fila.every(v => v === '')) continue;
    const celdas = {};
    columnas.forEach(c => { celdas[c.nombre] = fila[c.indice]; });
    filas.push({ fila: r + 1, celdas: celdas });
  }

  return {
    ok: true,
    hoja: config.nombre,
    columnaClave: config.columnaClave,
    columnas: columnas.map(c => ({ nombre: c.nombre, editable: c.editable })),
    filas: filas
  };
}

function editar(config, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = abrirPestana(config);
    const ubicacion = ubicarCelda(sheet, config, body.fila, body.columna, body.clave);
    if (!ubicacion.ok) return ubicacion;

    const celda = ubicacion.celda;
    if (celda.getFormula()) {
      return { ok: false, error: 'Esa celda es una fórmula y no se puede editar' };
    }

    const actual = celda.getDisplayValue();
    const anterior = body.anterior == null ? '' : String(body.anterior);

    if (actual !== anterior) {
      return {
        ok: false,
        conflicto: true,
        error: 'La celda cambió desde que cargaste la página',
        actual: actual
      };
    }

    celda.setValue(body.nuevo == null ? '' : body.nuevo);
    SpreadsheetApp.flush();
    const guardado = celda.getDisplayValue();

    registrar([new Date(), body.usuario || '', config.nombre, body.fila, body.clave, body.columna, actual, guardado]);

    return { ok: true, valor: guardado };
  } finally {
    lock.releaseLock();
  }
}

// Agrega una fila al final. body.celdas: { nombreColumna: valor }
// Las columnas con fórmula (por ejemplo Días restantes) se copian de la última fila,
// junto con su formato y validaciones (listas desplegables).
function agregar(config, body) {
  const celdas = body.celdas;
  if (!celdas || typeof celdas !== 'object') return { ok: false, error: 'Datos inválidos' };

  const nombre = String(celdas[config.columnaClave] || '').trim();
  if (!nombre) return { ok: false, error: 'Falta el nombre en ' + config.columnaClave };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = abrirPestana(config);
    const ultimaFila = sheet.getLastRow();
    const ultimaColumna = sheet.getLastColumn();
    const encabezados = sheet.getRange(config.filaEncabezado, 1, 1, ultimaColumna)
      .getDisplayValues()[0].map(h => String(h).trim());

    const colClave = encabezados.indexOf(config.columnaClave);
    if (colClave === -1) return { ok: false, error: 'No se encontró la columna ' + config.columnaClave };

    // No permitir dos centros con el mismo nombre
    if (ultimaFila > config.filaEncabezado) {
      const existentes = sheet.getRange(config.filaEncabezado + 1, colClave + 1, ultimaFila - config.filaEncabezado, 1)
        .getDisplayValues().map(f => normalizar(f[0]));
      if (existentes.includes(normalizar(nombre))) {
        return { ok: false, error: 'Ya existe un centro llamado "' + nombre + '"' };
      }
    }

    const nuevaFila = ultimaFila + 1;
    // Si la pestaña no tiene filas vacías al final, se crea una
    if (sheet.getMaxRows() < nuevaFila) sheet.insertRowAfter(ultimaFila);
    const destino = sheet.getRange(nuevaFila, 1, 1, ultimaColumna);

    // Copiar la última fila de datos trae fórmulas, formato y validaciones
    let formulas = new Array(ultimaColumna).fill('');
    if (ultimaFila > config.filaEncabezado) {
      sheet.getRange(ultimaFila, 1, 1, ultimaColumna).copyTo(destino);
      formulas = destino.getFormulas()[0];
    }

    const usadas = [];
    const fila = encabezados.map((encabezado, i) => {
      if (formulas[i]) return formulas[i];  // se mantiene la fórmula, ya ajustada a la fila nueva
      if (!encabezado || !(encabezado in celdas)) return '';
      usadas.push(encabezado);
      return comoTexto(celdas[encabezado]);
    });
    destino.setValues([fila]);
    SpreadsheetApp.flush();

    registrar([new Date(), body.usuario || '', config.nombre, nuevaFila, nombre, '(centro nuevo)', '', usadas.join(', ')]);

    return { ok: true, fila: nuevaFila };
  } finally {
    lock.releaseLock();
  }
}

// Borra la fila completa de un centro. El servidor ya verificó que el usuario es administrador.
// Antes de borrar se confirma que en esa fila sigue el mismo centro, y se guarda una copia
// de todos sus datos en el log para poder recuperarlos.
function eliminar(config, body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = abrirPestana(config);
    const fila = Number(body.fila);
    const ultimaFila = sheet.getLastRow();
    const ultimaColumna = sheet.getLastColumn();

    if (!Number.isInteger(fila) || fila <= config.filaEncabezado || fila > ultimaFila) {
      return { ok: false, error: 'Fila no válida' };
    }

    const encabezados = sheet.getRange(config.filaEncabezado, 1, 1, ultimaColumna)
      .getDisplayValues()[0].map(h => String(h).trim());
    const colClave = encabezados.indexOf(config.columnaClave);
    if (colClave === -1) return { ok: false, error: 'No se encontró la columna ' + config.columnaClave };

    const valores = sheet.getRange(fila, 1, 1, ultimaColumna).getDisplayValues()[0];
    if (valores[colClave] !== String(body.clave)) {
      return { ok: false, conflicto: true, error: 'La fila cambió de lugar en la planilla. Recarga la página.' };
    }

    const copia = {};
    encabezados.forEach((h, i) => { if (h) copia[h] = valores[i]; });

    sheet.deleteRow(fila);
    SpreadsheetApp.flush();

    registrar([new Date(), body.usuario || '', config.nombre, fila, body.clave, '(centro eliminado)', JSON.stringify(copia), '']);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Desactiva un centro: copia su fila al Historial (Estado Desactivado, FIN y un Código nuevo)
// y la saca de la hoja de vigentes. Todo en un solo paso: si no se puede sacar de la hoja
// de vigentes, se borra la copia del Historial para que nunca quede en las dos.
// El servidor ya verificó que el usuario es administrador.
function desactivar(config, body) {
  const destinoConfig = HOJAS[config.historial];
  if (!destinoConfig) return { ok: false, error: 'Esta planilla no tiene historial' };
  const fin = String(body.fin || '').trim();
  if (!/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(fin)) return { ok: false, error: 'Fecha de FIN no válida' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // --- Origen: la fila del centro vigente ---
    const origen = abrirPestana(config);
    const fila = Number(body.fila);
    const ultimaFilaOrigen = origen.getLastRow();
    const columnasOrigen = origen.getLastColumn();
    if (!Number.isInteger(fila) || fila <= config.filaEncabezado || fila > ultimaFilaOrigen) {
      return { ok: false, error: 'Fila no válida' };
    }
    const encOrigen = origen.getRange(config.filaEncabezado, 1, 1, columnasOrigen)
      .getDisplayValues()[0].map(h => String(h).trim());
    const rangoOrigen = origen.getRange(fila, 1, 1, columnasOrigen);
    const mostrados = rangoOrigen.getDisplayValues()[0];
    const valores = rangoOrigen.getValues()[0];      // fechas como fecha, números como número
    const colClave = encOrigen.indexOf(config.columnaClave);
    if (colClave === -1 || mostrados[colClave] !== String(body.clave)) {
      return { ok: false, conflicto: true, error: 'La fila cambió de lugar en la planilla. Recarga la página.' };
    }
    const deOrigen = {};
    encOrigen.forEach((h, i) => { if (h) deOrigen[normalizar(h)] = { valor: valores[i], texto: mostrados[i] }; });

    // --- Destino: una fila nueva al final del Historial ---
    const destino = abrirPestana(destinoConfig);
    const ultimaFilaDestino = destino.getLastRow();
    const columnasDestino = destino.getLastColumn();
    const encDestino = destino.getRange(destinoConfig.filaEncabezado, 1, 1, columnasDestino)
      .getDisplayValues()[0].map(h => String(h).trim());

    const centro = String(body.clave).trim();
    const inicio = deOrigen['inicio'] ? deOrigen['inicio'].texto : '';
    const codigo = codigoUnico(destino, destinoConfig, encDestino, centro, inicio);

    const nuevaFila = encDestino.map(h => {
      const n = normalizar(h);
      if (!h) return '';
      if (h === destinoConfig.columnaClave) return codigo;
      if (n === 'estado') return 'Desactivado';
      if (n === 'fin') return fin;
      const dato = deOrigen[n];
      if (!dato) return '';
      const v = dato.valor;
      if (typeof v === 'string') {
        if (/^0\d+$/.test(v)) return "'" + v;           // conserva ceros adelante (SIEP 011854)
        return comoTexto(v);
      }
      return v;
    });

    const filaDestino = ultimaFilaDestino + 1;
    if (destino.getMaxRows() < filaDestino) destino.insertRowAfter(ultimaFilaDestino);
    const rangoDestino = destino.getRange(filaDestino, 1, 1, columnasDestino);
    if (ultimaFilaDestino > destinoConfig.filaEncabezado) {
      destino.getRange(ultimaFilaDestino, 1, 1, columnasDestino)
        .copyTo(rangoDestino, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }
    rangoDestino.setValues([nuevaFila]);
    SpreadsheetApp.flush();

    // --- Sacar de vigentes; si falla, se deshace la copia ---
    try {
      origen.deleteRow(fila);
      SpreadsheetApp.flush();
    } catch (err) {
      destino.deleteRow(filaDestino);
      SpreadsheetApp.flush();
      return { ok: false, error: 'No se pudo sacar el centro de ' + config.nombre + '. No se movió nada.' };
    }

    registrar([new Date(), body.usuario || '', config.nombre, fila, centro, '(desactivado → ' + destinoConfig.nombre + ')', codigo, fin]);
    return { ok: true, codigo: codigo };
  } finally {
    lock.releaseLock();
  }
}

// Código de periodo: Centro-AAAA-MM según INICIO (d/m/aaaa). Si ya existe se agrega -2, -3…
function codigoUnico(sheet, config, encabezados, centro, inicio) {
  const m = String(inicio).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  const base = centro + '-' + (m ? m[3] + '-' + ('0' + m[2]).slice(-2) : 'sin-fecha');
  const col = encabezados.indexOf(config.columnaClave);
  const ultima = sheet.getLastRow();
  const existentes = ultima > config.filaEncabezado
    ? sheet.getRange(config.filaEncabezado + 1, col + 1, ultima - config.filaEncabezado, 1).getDisplayValues().map(f => normalizar(f[0]))
    : [];
  let codigo = base, n = 2;
  while (existentes.includes(normalizar(codigo))) codigo = base + '-' + n++;
  return codigo;
}

// ---------- Ayudantes ----------

function abrirPestana(config) {
  const sheet = SpreadsheetApp.openById(config.archivoId).getSheetByName(config.pestana);
  if (!sheet) throw new Error('No se encontró la pestaña ' + config.pestana);
  return sheet;
}

function ubicarCelda(sheet, config, fila, columna, clave) {
  fila = Number(fila);
  const ultimaFila = sheet.getLastRow();
  const ultimaColumna = sheet.getLastColumn();

  if (!Number.isInteger(fila) || fila <= config.filaEncabezado || fila > ultimaFila) {
    return { ok: false, error: 'Fila no válida' };
  }
  if (!columna) return { ok: false, error: 'Columna no válida' };

  const encabezados = sheet.getRange(config.filaEncabezado, 1, 1, ultimaColumna)
    .getDisplayValues()[0].map(h => String(h).trim());

  const col = encabezados.indexOf(columna);
  if (col === -1) return { ok: false, error: 'Columna no encontrada' };

  const colClave = encabezados.indexOf(config.columnaClave);
  const claveActual = sheet.getRange(fila, colClave + 1).getDisplayValue();
  if (claveActual !== String(clave)) {
    return { ok: false, conflicto: true, error: 'La fila cambió de lugar en la planilla. Recarga la página.' };
  }

  return { ok: true, celda: sheet.getRange(fila, col + 1) };
}

// Compara nombres sin importar mayúsculas, tildes ni espacios extra
function normalizar(texto) {
  return String(texto).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

// Un valor escrito en la página nunca debe convertirse en fórmula
function comoTexto(valor) {
  const v = valor == null ? '' : String(valor).trim();
  return /^[=+]/.test(v) ? "'" + v : v;
}

function registrar(fila) {
  const id = PropertiesService.getScriptProperties().getProperty('LOG_ID');
  if (!id) return;
  SpreadsheetApp.openById(id).getSheetByName('Ediciones').appendRow(fila);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Para correr a mano desde el editor ----------

// Correr UNA vez: genera la clave secreta y crea la planilla de log.
function configurar() {
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty('SECRETO')) {
    props.setProperty('SECRETO', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  }

  if (!props.getProperty('LOG_ID')) {
    const log = SpreadsheetApp.create('Log página mantenciones');
    const ediciones = log.getSheets()[0];
    ediciones.setName('Ediciones');
    ediciones.appendRow(['Fecha', 'Usuario', 'Planilla', 'Fila', 'Clave', 'Columna', 'Valor anterior', 'Valor nuevo']);
    props.setProperty('LOG_ID', log.getId());
  }

  Logger.log('SECRETO: ' + props.getProperty('SECRETO'));
  Logger.log('Log: ' + SpreadsheetApp.openById(props.getProperty('LOG_ID')).getUrl());
}

// Prueba de lectura sin pasar por la web app.
function probar() {
  const r = leer(HOJAS.sams);
  Logger.log('Columnas: ' + JSON.stringify(r.columnas));
  Logger.log('Filas: ' + r.filas.length);
  Logger.log('Ejemplo: ' + JSON.stringify(r.filas[0]));
}
