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

    // Tickets de la planilla Mantenciones (pestaña solicitudes)
    switch (body.accion) {
      case 'solicitudesLeer':    return responder(solicitudesLeer());
      case 'solicitudesAceptar': return responder(solicitudesAceptar(body));
      case 'solicitudesCerrar':  return responder(solicitudesCerrar(body));
      case 'solicitudesGestion': return responder(solicitudesGestion(body));
      case 'solicitudesNuevo':   return responder(solicitudesNuevo(body));
      case 'solicitudesCentro':  return responder(solicitudesCentro(body));
      case 'solicitudesEliminar': return responder(solicitudesEliminar(body));
      // Compensaciones de los técnicos (planilla TURNOS TECNICOS - COMPENSACIONES)
      case 'compLeer':           return responder(compLeer(body));
      case 'compSalida':         return responder(compGuardar('salidas', body));
      case 'compLibre':          return responder(compGuardar('libres', body));
      case 'compEliminar':       return responder(compEliminar(body));
      case 'compZona':           return responder(compZona(body));
      case 'compDescartar':      return responder(compDescartar(body));
      // Posición de los wellboats (pestaña wellboat de la planilla SAMS, solo lectura)
      case 'wellboatLeer':       return responder(wellboatLeer());
      // Programa de biólogos (planilla Programa biólogos)
      case 'progLeer':           return responder(progLeer());
      case 'progAsignar':        return responder(progAsignar(body));
      case 'progCambioCentro':   return responder(progCambioCentro(body));
      case 'progAgregarCentro':  return responder(progAgregarCentro(body));
      case 'progQuitarCentro':   return responder(progQuitarCentro(body));
      case 'progSegunda':        return responder(progSegunda(body));
      case 'progEliminarSegunda': return responder(progEliminarSegunda(body));
      case 'progReemplazo':      return responder(progReemplazo(body));
      case 'progEditarReemplazo': return responder(progEditarReemplazo(body));
      case 'progEliminarReemplazo': return responder(progEliminarReemplazo(body));
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

// ---------- Solicitudes (tickets de mantención) ----------
// Hacen exactamente lo mismo que los botones de la pestaña "ingreso" de la planilla Mantenciones:
// mismas columnas, mismos formatos de fecha y mismos avisos a n8n. Así el bot, las fórmulas
// de "datos" y los botones de la planilla siguen funcionando igual.

const SOLICITUDES = {
  archivoId: '1QCxGh77VEYD0nYUFlvH6qnJIw5PmqohScf1H5tqF1wU',
  pestana: 'solicitudes',
  listas: 'datos',  // E: técnicos, I: tipos de problema/servicio, J: centros
  webhookMensaje: 'https://n8n-n8n.ax8fpu.easypanel.host/webhook/mandar_mensaje',
  webhookNotion: 'https://n8n-n8n.ax8fpu.easypanel.host/webhook/notion'
};
// Columnas de la pestaña solicitudes (1 = A)
const COL = { id: 1, fecha: 2, nombre: 3, telefono: 4, centro: 5, problema: 6, estado: 7, tecnico: 8,
  tipo: 9, observaciones: 10, fechaTecnico: 11, aceptacion: 12, gestionCentro: 13, gestionTerreno: 14 };

function solicitudesLeer() {
  const ss = SpreadsheetApp.openById(SOLICITUDES.archivoId);
  const sheet = ss.getSheetByName(SOLICITUDES.pestana);
  const valores = sheet.getDataRange().getDisplayValues();
  const tickets = [];
  for (let r = 1; r < valores.length; r++) {
    const f = valores[r];
    if (!/^\d+$/.test(String(f[0]).trim())) continue;  // la fila 2 es un aviso, no un ticket
    tickets.push({
      fila: r + 1, id: f[0].trim(), fecha: f[1].trim(), nombre: f[2], telefono: f[3], centro: f[4], problema: f[5],
      estado: f[6], tecnico: f[7], tipo: f[8], observaciones: f[9], fechaTecnico: f[10].trim(),
      aceptacion: f[11].trim(), gestionCentro: f[12] || '', gestionTerreno: f[13] || ''
    });
  }
  const datos = ss.getSheetByName(SOLICITUDES.listas).getDataRange().getDisplayValues();
  const columna = i => [...new Set(datos.map(f => String(f[i] || '').trim()).filter(Boolean))];
  return {
    ok: true,
    tickets: tickets,
    tecnicos: columna(4),
    tipos: columna(8),
    centros: columna(9),
    zona: ss.getSpreadsheetTimeZone()
  };
}

// Aceptar = "abrir ticket": guarda el técnico (H) y la hora de inicio (L)
function solicitudesAceptar(body) {
  if (!body.tecnico) return { ok: false, error: 'Falta el técnico' };
  return conTicket(body, (ss, sheet, fila) => {
    if (sheet.getRange(fila, COL.aceptacion).getDisplayValue().trim()) {
      return { ok: false, conflicto: true, error: 'Ese ticket ya fue aceptado. Recarga la página.' };
    }
    sheet.getRange(fila, COL.tecnico).setValue(body.tecnico);
    sheet.getRange(fila, COL.aceptacion).setValue(ahora(ss, 'dd-MM-yyyy - HH:mm'));
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Solicitudes', fila, body.id, '(ticket aceptado)', '', body.tecnico]);
    return { ok: true };
  });
}

// Cerrar ticket (registrarSolicitud): Finalizado y mensaje a la bióloga.
// Si el tipo de servicio es una gestión, pasa a Gestión centro / terreno como en la planilla.
function solicitudesCerrar(body) {
  const tipo = unirOpciones(body.tipos);
  if (!body.tecnico || !tipo) return { ok: false, error: 'Completa el técnico y el tipo de servicio' };
  return conTicket(body, (ss, sheet, fila) => {
    if (!sheet.getRange(fila, COL.aceptacion).getDisplayValue().trim()) {
      return { ok: false, error: 'Primero hay que aceptar el ticket (no tiene fecha de aceptación).' };
    }
    const fechaHora = ahora(ss, 'yyyy-MM-dd - HH:mm');
    let estado = 'Finalizado', webhook = SOLICITUDES.webhookMensaje;
    if (tipo.toLowerCase() === 'gestión ayuda centro') {
      estado = 'Gestión centro'; webhook = SOLICITUDES.webhookNotion;
      sheet.getRange(fila, COL.gestionCentro).setValue('Centro // ' + fechaHora);
    } else if (tipo.toLowerCase() === 'gestión ingreso terreno') {
      estado = 'Gestión terreno'; webhook = SOLICITUDES.webhookNotion;
      sheet.getRange(fila, COL.gestionTerreno).setValue('Terreno // ' + fechaHora);
    }
    return guardarCierre(ss, sheet, fila, body, estado, tipo, fechaHora, webhook);
  });
}

// Botones "Gestión ayuda centro" / "Gestión ingreso terreno" (registrarSolicitudTipo)
function solicitudesGestion(body) {
  if (!body.tecnico) return { ok: false, error: 'Falta el técnico' };
  if (body.gestion !== 'centro' && body.gestion !== 'terreno') return { ok: false, error: 'Gestión no válida' };
  return conTicket(body, (ss, sheet, fila) => {
    const fechaHora = ahora(ss, 'yyyy-MM-dd - HH:mm');
    let estado, tipo;
    if (body.gestion === 'centro') {
      estado = 'Gestión centro'; tipo = 'Gestión ayuda centro';
      sheet.getRange(fila, COL.gestionCentro).setValue('Centro // ' + fechaHora);
    } else {
      estado = 'Gestión terreno'; tipo = 'Gestión ingreso terreno';
      sheet.getRange(fila, COL.gestionTerreno).setValue('Terreno // ' + fechaHora);
    }
    return guardarCierre(ss, sheet, fila, body, estado, tipo, fechaHora, SOLICITUDES.webhookNotion);
  });
}

// Ticket ingresado por un técnico (crearSolicitudTecnico): queda aceptado por él mismo
function solicitudesNuevo(body) {
  const centro = String(body.centro || '').trim();
  const tecnico = String(body.tecnico || '').trim();
  const problema = unirOpciones(body.problemas);
  if (!centro || !tecnico || !problema) return { ok: false, error: 'Completa centro, técnico y tipo de problema' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SOLICITUDES.archivoId);
    const sheet = ss.getSheetByName(SOLICITUDES.pestana);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
      .map(f => Number(f[0])).filter(n => !isNaN(n) && n > 0);
    const nuevoId = Math.max(0, ...ids) + 1;
    const fecha = ahora(ss, 'dd-MM-yyyy - HH:mm');
    sheet.appendRow([nuevoId, fecha, tecnico, 'Interno', centro, problema, 'Falla', tecnico, '',
      comoTexto(body.observaciones), '', fecha, '', '']);
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Solicitudes', sheet.getLastRow(), nuevoId, '(ticket nuevo)', centro, problema]);
    return { ok: true, id: nuevoId };
  } finally {
    lock.releaseLock();
  }
}

// Corrige el centro (E) de un ticket abierto que llegó con el centro equivocado. No avisa a nadie.
function solicitudesCentro(body) {
  const centro = String(body.centro || '').trim();
  if (!centro) return { ok: false, error: 'Elige el centro' };
  return conTicket(body, (ss, sheet, fila) => {
    const centros = ss.getSheetByName(SOLICITUDES.listas).getRange('J:J').getDisplayValues().flat().map(v => v.trim());
    if (!centros.includes(centro)) return { ok: false, error: 'Ese centro no está en la lista de la planilla' };
    if (/^finaliz/i.test(sheet.getRange(fila, COL.estado).getDisplayValue().trim())) {
      return { ok: false, conflicto: true, error: 'El ticket ya está finalizado. Recarga la página.' };
    }
    const anterior = sheet.getRange(fila, COL.centro).getDisplayValue();
    if (anterior.trim() === centro) return { ok: false, error: 'El ticket ya tiene ese centro' };
    sheet.getRange(fila, COL.centro).setValue(centro);
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Solicitudes', fila, body.id, '(centro corregido)', anterior, centro]);
    return { ok: true };
  });
}

// Borra la fila de un ticket abierto. Antes copia la fila completa a "Solicitudes eliminadas"
// en la planilla de log, con quién y por qué; si no se puede respaldar, no se borra nada.
function solicitudesEliminar(body) {
  const motivo = String(body.motivo || '').trim();
  if (motivo.length < 5) return { ok: false, error: 'Escribe el motivo' };
  if (String(body.confirmacion || '').trim() !== String(body.id)) {
    return { ok: false, error: 'El número escrito no coincide con el ticket' };
  }
  const logId = PropertiesService.getScriptProperties().getProperty('LOG_ID');
  if (!logId) return { ok: false, error: 'No hay planilla de log para guardar el respaldo. No se borró nada.' };
  return conTicket(body, (ss, sheet, fila) => {
    if (/^finaliz/i.test(sheet.getRange(fila, COL.estado).getDisplayValue().trim())) {
      return { ok: false, error: 'Un ticket finalizado no se puede eliminar' };
    }
    const columnas = Math.max(sheet.getLastColumn(), COL.gestionTerreno);
    const valores = sheet.getRange(fila, 1, 1, columnas).getDisplayValues()[0];

    const log = SpreadsheetApp.openById(logId);
    let respaldo = log.getSheetByName('Solicitudes eliminadas');
    if (!respaldo) {
      respaldo = log.insertSheet('Solicitudes eliminadas');
      const enc = sheet.getRange(1, 1, 1, columnas).getDisplayValues()[0];
      respaldo.appendRow(['Eliminado el', 'Usuario', 'Motivo', 'Fila'].concat(enc));
      respaldo.setFrozenRows(1);
    }
    respaldo.appendRow([new Date(), body.usuario || '', motivo, fila].concat(valores.map(comoTexto)));
    SpreadsheetApp.flush();
    // Se comprueba que el respaldo quedó escrito antes de borrar
    const ultima = respaldo.getRange(respaldo.getLastRow(), 5).getDisplayValue().trim();
    if (ultima !== String(body.id)) return { ok: false, error: 'No se pudo guardar el respaldo. No se borró nada.' };

    sheet.deleteRow(fila);
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Solicitudes', fila, body.id, '(ticket eliminado)', valores[COL.centro - 1], motivo]);
    return { ok: true };
  });
}

// ---------- Compensaciones de los técnicos ----------
// Pestañas nuevas en la planilla de compensaciones; las pestañas por técnico quedan como historial.
// Las salidas a terreno suman días; los días libres dados los restan.

const COMPENSACIONES = {
  archivoId: '1xux68mBpvF0flEUX2_j1JiVb81kEVwN_cILzCczQpiM',
  tecnicos: ['Adrián Calcumil', 'Sergio Martin', 'Wilmer Ruiz', 'Matias Pineda'],
  hojas: {
    salidas: { pestana: 'Salidas', columnas: ['ID', 'Técnico', 'Centro', 'Desde', 'Hasta', 'Días', 'Tramos de vuelo', 'Domingos',
      'Feriados', 'Retorno domingo', 'Total', 'Observación', 'Registrado por', 'Registrado el'] },
    libres: { pestana: 'Días libres', columnas: ['ID', 'Técnico', 'Desde', 'Hasta', 'Días', 'Observación', 'Registrado por', 'Registrado el'] },
    zonas: { pestana: 'Zonas', columnas: ['Centro', 'Compensa', 'Actualizado por', 'Actualizado el'] },
    descartados: { pestana: 'Certificados descartados', columnas: ['Clave', 'Técnico', 'Centro', 'Fecha', 'Descartado por', 'Descartado el'] }
  },
  // Formulario "Certificado de Servicios": de aquí sale a dónde fue cada técnico
  certificados: { archivoId: '1ILyE1aSgBv5JAPOAYINz75aYIl6K7pdPyajVSmWQYBc', pestana: 'Respuestas de formulario 1' }
};

// Días compensados de una salida: ⅓ por día trabajado, 0,25 por tramo en avión,
// 1,5 por domingo o feriado trabajado y 0,5 por retorno en domingo (sin redondear: ⅓ se guarda exacto)
function totalSalida(s) {
  return Math.round((s['Días'] / 3 + s['Tramos de vuelo'] * 0.25 + s['Domingos'] * 1.5 + s['Feriados'] * 1.5 +
    s['Retorno domingo'] * 0.5) * 1e6) / 1e6;
}

function hojaComp(ss, tipo) {
  const config = COMPENSACIONES.hojas[tipo];
  let sheet = ss.getSheetByName(config.pestana);
  if (!sheet) {
    sheet = ss.insertSheet(config.pestana);
    sheet.appendRow(config.columnas);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, config.columnas.length).setFontWeight('bold');
  }
  return sheet;
}

// Filas como objetos por encabezado; las fechas salen como yyyy-MM-dd
function filasComp(ss, tipo) {
  const sheet = hojaComp(ss, tipo);
  const valores = sheet.getDataRange().getValues();
  const encabezados = valores[0].map(h => String(h).trim());
  const zona = ss.getSpreadsheetTimeZone();
  return valores.slice(1).map((f, i) => {
    const o = { fila: i + 2 };
    encabezados.forEach((h, j) => {
      const v = f[j];
      o[h] = v instanceof Date ? Utilities.formatDate(v, zona, 'yyyy-MM-dd') : v;
    });
    return o;
  }).filter(o => o[encabezados[0]] !== '' && o[encabezados[0]] != null);
}

function compLeer(body) {
  const ss = SpreadsheetApp.openById(COMPENSACIONES.archivoId);
  const respuesta = {
    ok: true,
    tecnicos: COMPENSACIONES.tecnicos,
    salidas: filasComp(ss, 'salidas'),
    libres: filasComp(ss, 'libres')
  };
  // Lo que solo usa quien administra: zonas, certificados y descartados
  if (body.conCertificados) {
    respuesta.zonas = filasComp(ss, 'zonas');
    respuesta.descartados = filasComp(ss, 'descartados').map(d => String(d['Clave']));
    respuesta.certificados = leerCertificados(Number(body.dias) || 120);
  }
  return respuesta;
}

// Certificados de terreno de los últimos días; uno por técnico (a veces van dos al mismo centro)
function leerCertificados(dias) {
  const ss = SpreadsheetApp.openById(COMPENSACIONES.certificados.archivoId);
  const sheet = ss.getSheetByName(COMPENSACIONES.certificados.pestana);
  if (!sheet) return [];
  const zona = ss.getSpreadsheetTimeZone();
  const valores = sheet.getDataRange().getValues();
  const enc = valores[0].map(h => normalizar(h));
  const col = nombre => enc.findIndex(h => h.startsWith(nombre));
  const c = { marca: col('marca temporal'), fecha: col('fecha mantencion'), centro: col('nombre centro'), tecnico: col('tecnico responsable'),
    tipo: col('tipo de servicio'), empresa: col('empresa'), modalidad: col('modalidad') };
  const desde = new Date(Date.now() - dias * 86400000);
  const aFecha = v => {
    if (v instanceof Date) return v;
    const m = String(v || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
  };
  const lista = [];
  valores.slice(1).forEach(f => {
    const fecha = aFecha(f[c.fecha]) || aFecha(f[c.marca]);
    if (!fecha || fecha < desde) return;
    if (c.modalidad >= 0 && /remot/i.test(String(f[c.modalidad]))) return;
    const centro = String(f[c.centro] || '').trim();
    if (!centro) return;
    const marca = f[c.marca] instanceof Date ? Utilities.formatDate(f[c.marca], zona, "yyyy-MM-dd'T'HH:mm:ss") : String(f[c.marca]);
    String(f[c.tecnico] || '').split(/,| y /).map(t => t.trim()).filter(Boolean).forEach(tecnico => {
      lista.push({
        clave: marca + '|' + normalizar(centro) + '|' + normalizar(tecnico),
        fecha: Utilities.formatDate(fecha, zona, 'yyyy-MM-dd'),
        centro: centro, tecnico: tecnico,
        tipo: String(f[c.tipo] || ''), empresa: String(f[c.empresa] || '').trim()
      });
    });
  });
  return lista;
}

const aFechaHoja = texto => {
  const m = String(texto || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
};
const numeroComp = (v, max) => {
  const n = Number(v);
  return isFinite(n) && n >= 0 && n <= max ? n : null;
};

// Crea o actualiza (si trae id) una salida o un día libre
function compGuardar(tipo, body) {
  const config = COMPENSACIONES.hojas[tipo];
  const tecnico = COMPENSACIONES.tecnicos.find(t => normalizar(t) === normalizar(body.tecnico));
  if (!tecnico) return { ok: false, error: 'Elige un técnico de la lista' };
  const desde = aFechaHoja(body.desde), hasta = aFechaHoja(body.hasta || body.desde);
  if (!desde || !hasta) return { ok: false, error: 'Revisa las fechas' };
  if (hasta < desde) return { ok: false, error: 'La fecha "hasta" es anterior a "desde"' };

  const d = { 'Técnico': tecnico, 'Desde': desde, 'Hasta': hasta, 'Observación': comoTexto(body.observacion) };
  if (tipo === 'salidas') {
    d['Centro'] = comoTexto(body.centro);
    if (!d['Centro']) return { ok: false, error: 'Escribe el centro' };
    d['Días'] = numeroComp(body.dias, 60);
    d['Tramos de vuelo'] = numeroComp(body.tramos, 10);
    d['Domingos'] = numeroComp(body.domingos, 20);
    d['Feriados'] = numeroComp(body.feriados, 20);
    d['Retorno domingo'] = numeroComp(body.retornoDomingo, 5);
    if ([d['Días'], d['Tramos de vuelo'], d['Domingos'], d['Feriados'], d['Retorno domingo']].some(v => v === null)) {
      return { ok: false, error: 'Revisa los números' };
    }
    d['Total'] = totalSalida(d);
  } else {
    d['Días'] = numeroComp(body.dias, 60);
    if (!d['Días']) return { ok: false, error: 'Escribe cuántos días libres' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(COMPENSACIONES.archivoId);
    const sheet = hojaComp(ss, tipo);
    const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(f => Number(f[0])) : [];
    let id = Number(body.id), fila;
    if (body.id != null && body.id !== '') {
      const i = ids.indexOf(id);
      if (i === -1) return { ok: false, conflicto: true, error: 'Ese registro ya no existe. Recarga la página.' };
      fila = i + 2;
    } else {
      id = Math.max(0, ...ids.filter(n => !isNaN(n))) + 1;
      fila = sheet.getLastRow() + 1;
      d['Registrado por'] = body.usuario || '';
      d['Registrado el'] = new Date();
    }
    d['ID'] = id;
    const actual = fila <= sheet.getLastRow() ? sheet.getRange(fila, 1, 1, encabezados.length).getValues()[0] : encabezados.map(() => '');
    const nueva = encabezados.map((h, j) => h in d ? d[h] : actual[j]);
    sheet.getRange(fila, 1, 1, encabezados.length).setValues([nueva]);
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Compensaciones', fila, id, body.id ? '(' + config.pestana + ' editada)' : '(' + config.pestana + ' nueva)',
      tecnico, tipo === 'salidas' ? d['Centro'] + ' · ' + d['Total'] : d['Días'] + ' días']);
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function compEliminar(body) {
  const tipo = body.tipo === 'libres' ? 'libres' : body.tipo === 'salidas' ? 'salidas' : '';
  if (!tipo) return { ok: false, error: 'Tipo no válido' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(COMPENSACIONES.archivoId);
    const sheet = hojaComp(ss, tipo);
    if (sheet.getLastRow() < 2) return { ok: false, conflicto: true, error: 'Ese registro ya no existe. Recarga la página.' };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(f => Number(f[0]));
    const i = ids.indexOf(Number(body.id));
    if (i === -1) return { ok: false, conflicto: true, error: 'Ese registro ya no existe. Recarga la página.' };
    const valores = sheet.getRange(i + 2, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    sheet.deleteRow(i + 2);
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Compensaciones', i + 2, body.id, '(' + COMPENSACIONES.hojas[tipo].pestana + ' eliminada)', valores.join(' | '), '']);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Si un centro compensa o no. compensa: true, false o null (vuelve a la zona automática)
function compZona(body) {
  const centro = String(body.centro || '').trim();
  if (!centro) return { ok: false, error: 'Falta el centro' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(COMPENSACIONES.archivoId);
    const sheet = hojaComp(ss, 'zonas');
    const centros = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(f => normalizar(f[0])) : [];
    const i = centros.indexOf(normalizar(centro));
    if (body.compensa === null) {
      if (i !== -1) sheet.deleteRow(i + 2);
    } else {
      const fila = [comoTexto(centro), body.compensa ? 'Sí' : 'No', body.usuario || '', new Date()];
      if (i === -1) sheet.appendRow(fila);
      else sheet.getRange(i + 2, 1, 1, 4).setValues([fila]);
    }
    SpreadsheetApp.flush();
    registrar([new Date(), body.usuario || '', 'Compensaciones', '', '', '(zona)', centro, body.compensa === null ? 'automática' : body.compensa ? 'compensa' : 'no compensa']);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Certificado que no corresponde compensar (o que ya está cubierto de otra forma)
function compDescartar(body) {
  const clave = String(body.clave || '').trim();
  if (!clave) return { ok: false, error: 'Falta el certificado' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(COMPENSACIONES.archivoId);
    const sheet = hojaComp(ss, 'descartados');
    const claves = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(f => String(f[0])) : [];
    const i = claves.indexOf(clave);
    if (body.deshacer) { if (i !== -1) sheet.deleteRow(i + 2); }
    else if (i === -1) sheet.appendRow([clave, comoTexto(body.tecnico), comoTexto(body.centro), comoTexto(body.fecha), body.usuario || '', new Date()]);
    SpreadsheetApp.flush();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Wellboats ----------
// La pestaña wellboat de SAMS se actualiza sola con la posición de cada buque; aquí solo se lee.
function wellboatLeer() {
  const ss = SpreadsheetApp.openById(HOJAS.sams.archivoId);
  const sheet = ss.getSheetByName('wellboat');
  if (!sheet) return { ok: false, error: 'No se encontró la pestaña wellboat en la planilla SAMS' };
  const v = sheet.getDataRange().getDisplayValues();
  const enc = (v[0] || []).map(h => normalizar(h));
  const col = nombres => enc.findIndex(h => nombres.some(n => h.startsWith(n)));
  const c = { nombre: col(['wellboat', 'buque', 'nombre']), datos: col(['datos', 'estado']), lat: col(['lat']), lon: col(['lon', 'lng']), rango: col(['rango']) };
  const numero = x => parseFloat(String(x || '').trim().replace(',', '.'));
  const buques = v.slice(1).map(f => ({
    nombre: String(f[c.nombre] || '').trim(), datos: c.datos >= 0 ? String(f[c.datos] || '').trim() : '',
    lat: numero(f[c.lat]), lon: numero(f[c.lon]), rango: c.rango >= 0 ? String(f[c.rango] || '').trim() : ''
  })).filter(b => b.nombre && isFinite(b.lat) && isFinite(b.lon));
  return { ok: true, buques: buques, leido: ahora(ss, 'HH:mm') };
}

// ---------- Programa de biólogos ----------
// La página usa pestañas propias en la planilla "Programa biólogos": "Programa (página)" con un centro por fila
// y "Biólogos (página)". La pestaña Programa original no se modifica; solo se lee una vez para la copia inicial.

const PROGRAMA = {
  archivoId: '1kNLxUmWqlEuKNDdycweBmG_bCJSLXWZPt7RT8rP9Tnc',
  original: 'Programa',
  pestana: 'Programa (página)',
  biologos: 'Biólogos (página)',
  segunda: 'Segunda Muestra',
  historial: 'Historial del programa',
  reemplazos: 'Reemplazos',
  plancton: ['PA', 'PAC']   // Plancton Andino (Noruega / Chile): sin biólogo
};
const PROG_COLUMNAS = ['ID', 'Empresa', 'Centro', 'Centro nuevo', 'Fecha cambio', 'Lunes a viernes', 'Sábado y domingo', 'Ubicación', 'Barrio', 'Actualizado por', 'Actualizado el'];
// Columnas de "Programa (página)" (1 = A)
const PCOL = { id: 1, empresa: 2, centro: 3, nuevo: 4, fecha: 5, semana: 6, finde: 7, ubicacion: 8, barrio: 9, por: 10, el: 11 };
const SEGUNDA_COLUMNAS = ['Centro', 'Biologa', 'Tiempo', 'Inicio', 'Fin', 'Horario', 'Asignado por', 'Asignado el', 'Observación'];
// Reemplazos temporales: el titular sigue en el programa; entre Desde y Hasta el centro lo cubre otro biólogo
const REEMP_COLUMNAS = ['ID', 'Centro ID', 'Centro', 'Tramo', 'Titular', 'Reemplazo', 'Desde', 'Hasta', 'Motivo', 'Observación', 'Registrado por', 'Registrado el'];

// "Rada Achao ---> Lincay (15/nov)" → { actual, nuevo, fecha }
function partirCentro(texto) {
  const m = String(texto || '').match(/^(.*?)\s*-+>\s*(.*?)\s*(?:\(([^)]*)\))?\s*$/);
  return m ? { actual: m[1].trim(), nuevo: m[2].trim(), fecha: (m[3] || '').trim() } : { actual: String(texto || '').trim(), nuevo: '', fecha: '' };
}

// Crea las pestañas de la página la primera vez, copiando el programa y los biólogos de la pestaña Programa
function prepararPrograma(ss) {
  let sheet = ss.getSheetByName(PROGRAMA.pestana), bios = ss.getSheetByName(PROGRAMA.biologos);
  if (sheet && bios) return { sheet, bios };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet = ss.getSheetByName(PROGRAMA.pestana);
    bios = ss.getSheetByName(PROGRAMA.biologos);
    const original = ss.getSheetByName(PROGRAMA.original);
    const v = original ? original.getRange(1, 1, original.getLastRow(), 12).getDisplayValues() : [[]];
    let i = 1;
    const centros = [];
    let empresa = '';
    for (; i < v.length; i++) {
      const f = v[i];
      if (/^total/i.test(String(f[1]).trim())) break;
      if (String(f[1]).trim()) empresa = String(f[1]).trim();
      const nombre = String(f[2]).trim();
      if (!nombre) continue;
      const c = partirCentro(nombre);
      centros.push([centros.length + 1, empresa, c.actual, c.nuevo, c.fecha, String(f[3]).trim(), String(f[8]).trim(), String(f[10]).trim(), String(f[11]).trim(), 'Copia de Programa', new Date()]);
    }
    const biologos = [];
    for (; i < v.length; i++) {
      const f = v[i];
      const nombre = String(f[1]).trim(), sigla = String(f[2]).trim();
      if (!nombre || !sigla || /^total|^centros/i.test(nombre)) continue;
      biologos.push([nombre, sigla, f.slice(3, 8).some(x => String(x).trim() !== '') ? 'Semana' : 'Fin de semana']);
    }
    if (!sheet) {
      sheet = ss.insertSheet(PROGRAMA.pestana);
      sheet.getRange(1, 1, 1, PROG_COLUMNAS.length).setValues([PROG_COLUMNAS]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (centros.length) sheet.getRange(2, 1, centros.length, PROG_COLUMNAS.length).setValues(centros);
    }
    if (!bios) {
      bios = ss.insertSheet(PROGRAMA.biologos);
      bios.getRange(1, 1, 1, 3).setValues([['Nombre', 'Sigla', 'Equipo']]).setFontWeight('bold');
      bios.setFrozenRows(1);
      if (biologos.length) bios.getRange(2, 1, biologos.length, 3).setValues(biologos);
    }
    SpreadsheetApp.flush();
    return { sheet, bios };
  } finally {
    lock.releaseLock();
  }
}

function progLeer() {
  const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
  const { sheet, bios } = prepararPrograma(ss);
  const v = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, PROG_COLUMNAS.length).getDisplayValues() : [];
  const centros = v.map((f, k) => ({
    fila: k + 2, id: String(f[0]).trim(), empresa: String(f[1]).trim(), centro: String(f[2]).trim(), nuevo: String(f[3]).trim(), fecha: String(f[4]).trim(),
    semana: String(f[5]).trim(), finde: String(f[6]).trim(), ubicacion: String(f[7]).trim(), barrio: String(f[8]).trim()
  })).filter(c => c.id && c.centro);
  const b = bios.getLastRow() > 1 ? bios.getRange(2, 1, bios.getLastRow() - 1, 3).getDisplayValues() : [];
  const biologos = b.map(f => ({ nombre: String(f[0]).trim(), sigla: String(f[1]).trim(), equipo: /fin/i.test(f[2]) ? 'finde' : 'semana' }))
    .filter(x => x.nombre && x.sigla);
  return { ok: true, centros: centros, biologos: biologos, plancton: PROGRAMA.plancton, segundas: leerSegundas(ss), reemplazos: leerReemplazos(ss) };
}

function hojaReemplazos(ss) {
  let sheet = ss.getSheetByName(PROGRAMA.reemplazos);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRAMA.reemplazos);
    sheet.getRange(1, 1, 1, REEMP_COLUMNAS.length).setValues([REEMP_COLUMNAS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function leerReemplazos(ss) {
  const sheet = ss.getSheetByName(PROGRAMA.reemplazos);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const zona = ss.getSpreadsheetTimeZone();
  const iso = x => x instanceof Date ? Utilities.formatDate(x, zona, 'yyyy-MM-dd') : String(x || '').trim();
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, REEMP_COLUMNAS.length).getValues().map(f => ({
    id: String(f[0]).trim(), centroId: String(f[1]).trim(), centro: String(f[2]).trim(), tramo: /s.bado|fin/i.test(f[3]) ? 'finde' : 'semana',
    titular: String(f[4]).trim(), reemplazo: String(f[5]).trim(), desde: iso(f[6]), hasta: iso(f[7]), motivo: String(f[8] || '').trim(),
    observacion: String(f[9] || '').trim(), registradoPor: String(f[10] || '').trim()
  })).filter(r => r.id && r.centro);
}

// Registra un reemplazo para uno o varios centros del mismo titular (una fila por centro)
function progReemplazo(body) {
  const tramo = body.tramo === 'finde' ? 'finde' : 'semana';
  const titular = String(body.titular || '').trim();
  const desde = aFechaHoja(body.desde), hasta = aFechaHoja(body.hasta);
  const centros = Array.isArray(body.centros) ? body.centros : [];
  if (!titular) return { ok: false, error: 'Elige el biólogo que no estará' };
  if (!desde || !hasta) return { ok: false, error: 'Revisa las fechas' };
  if (hasta < desde) return { ok: false, error: 'La fecha "hasta" es anterior a "desde"' };
  if (!centros.length) return { ok: false, error: 'Marca al menos un centro' };
  if (centros.some(c => !String(c.reemplazo || '').trim() || String(c.reemplazo).trim() === titular)) return { ok: false, error: 'Elige quién cubre cada centro' };
  const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = hojaReemplazos(ss);
    const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(f => Number(f[0])).filter(n => !isNaN(n)) : [];
    let id = Math.max(0, ...ids);
    const filas = centros.map(c => [++id, comoTexto(c.id), comoTexto(c.centro), tramo === 'finde' ? 'Sábado y domingo' : 'Lunes a viernes', comoTexto(titular),
      comoTexto(c.reemplazo), desde, hasta, comoTexto(body.motivo), comoTexto(body.observacion), body.usuario || '', new Date()]);
    sheet.getRange(sheet.getLastRow() + 1, 1, filas.length, REEMP_COLUMNAS.length).setValues(filas);
    SpreadsheetApp.flush();
    centros.forEach(c => historialPrograma(ss, body.usuario, c.centro, 'Reemplazo ' + (tramo === 'finde' ? 'sábado y domingo' : 'lunes a viernes'),
      titular, c.reemplazo + ' · ' + body.desde + ' a ' + body.hasta + (body.motivo ? ' · ' + body.motivo : '')));
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function conReemplazo(body, accion) {
  const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = ss.getSheetByName(PROGRAMA.reemplazos);
    if (!sheet || sheet.getLastRow() < 2) return { ok: false, conflicto: true, error: 'Ese reemplazo ya no existe. Recarga la página.' };
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues().map(f => String(f[0]).trim());
    const i = ids.indexOf(String(body.id));
    if (i === -1) return { ok: false, conflicto: true, error: 'Ese reemplazo ya no existe. Recarga la página.' };
    return accion(ss, sheet, i + 2, sheet.getRange(i + 2, 1, 1, REEMP_COLUMNAS.length).getDisplayValues()[0]);
  } finally {
    lock.releaseLock();
  }
}

function progEditarReemplazo(body) {
  const desde = aFechaHoja(body.desde), hasta = aFechaHoja(body.hasta);
  const reemplazo = String(body.reemplazo || '').trim();
  if (!reemplazo) return { ok: false, error: 'Elige quién cubre' };
  if (!desde || !hasta) return { ok: false, error: 'Revisa las fechas' };
  if (hasta < desde) return { ok: false, error: 'La fecha "hasta" es anterior a "desde"' };
  return conReemplazo(body, (ss, sheet, fila, actual) => {
    if (reemplazo === actual[4]) return { ok: false, error: 'Quien cubre no puede ser el mismo titular' };
    sheet.getRange(fila, 6, 1, 5).setValues([[comoTexto(reemplazo), desde, hasta, comoTexto(body.motivo), comoTexto(body.observacion)]]);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, actual[2], 'Reemplazo editado', actual[5] + ' · ' + actual[6] + ' a ' + actual[7], reemplazo + ' · ' + body.desde + ' a ' + body.hasta);
    return { ok: true };
  });
}

function progEliminarReemplazo(body) {
  return conReemplazo(body, (ss, sheet, fila, actual) => {
    sheet.deleteRow(fila);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, actual[2], 'Reemplazo eliminado', actual[4] + ' → ' + actual[5] + ' · ' + actual[6] + ' a ' + actual[7], '');
    return { ok: true };
  });
}

function hojaSegunda(ss) {
  const sheet = ss.getSheetByName(PROGRAMA.segunda) || ss.insertSheet(PROGRAMA.segunda);
  // Columnas nuevas al final (Horario, Asignado por, Asignado el, Observación) si aún no están
  const actual = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), SEGUNDA_COLUMNAS.length)).getDisplayValues()[0] : [];
  SEGUNDA_COLUMNAS.forEach((h, j) => { if (!String(actual[j] || '').trim()) sheet.getRange(1, j + 1).setValue(h).setFontWeight('bold'); });
  return sheet;
}

function leerSegundas(ss) {
  const sheet = ss.getSheetByName(PROGRAMA.segunda);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const zona = ss.getSpreadsheetTimeZone();
  const v = sheet.getRange(2, 1, sheet.getLastRow() - 1, SEGUNDA_COLUMNAS.length).getValues();
  const iso = x => x instanceof Date ? Utilities.formatDate(x, zona, 'yyyy-MM-dd') : String(x || '').trim();
  return v.map((f, k) => ({ fila: k + 2, centro: String(f[0]).trim(), biologa: String(f[1]).trim(), tiempo: String(f[2]).trim(),
    inicio: iso(f[3]), fin: iso(f[4]), horario: String(f[5] || '').trim(), asignadoPor: String(f[6] || '').trim(),
    asignadoEl: iso(f[7]), observacion: String(f[8] || '').trim() })).filter(s => s.centro);
}

function historialPrograma(ss, usuario, centro, cambio, antes, despues) {
  let sheet = ss.getSheetByName(PROGRAMA.historial);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRAMA.historial);
    sheet.appendRow(['Fecha', 'Usuario', 'Centro', 'Cambio', 'Antes', 'Después']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  sheet.appendRow([new Date(), usuario || '', comoTexto(centro), cambio, comoTexto(antes), comoTexto(despues)]);
}

// Ubica el centro por su ID y confirma que sigue siendo el mismo antes de escribir
function conCentro(body, accion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
    const sheet = ss.getSheetByName(PROGRAMA.pestana);
    if (!sheet || sheet.getLastRow() < 2) return { ok: false, conflicto: true, error: 'El programa cambió. Recarga la página.' };
    const ids = sheet.getRange(2, PCOL.id, sheet.getLastRow() - 1, 1).getDisplayValues().map(f => String(f[0]).trim());
    const i = ids.indexOf(String(body.id));
    if (i === -1) return { ok: false, conflicto: true, error: 'Ese centro ya no está en el programa. Recarga la página.' };
    const fila = i + 2;
    const centro = sheet.getRange(fila, PCOL.centro).getDisplayValue().trim();
    if (centro !== String(body.centro || '').trim()) return { ok: false, conflicto: true, error: 'El programa cambió. Recarga la página.' };
    return accion(ss, sheet, fila, centro);
  } finally {
    lock.releaseLock();
  }
}

const marcarActualizado = (sheet, fila, usuario) => sheet.getRange(fila, PCOL.por, 1, 2).setValues([[usuario || '', new Date()]]);

// Biólogo de lunes a viernes o de sábado y domingo
function progAsignar(body) {
  const tramo = body.tramo === 'finde' ? 'finde' : body.tramo === 'semana' ? 'semana' : '';
  const sigla = String(body.sigla || '').trim();
  if (!tramo || !sigla) return { ok: false, error: 'Faltan datos' };
  return conCentro(body, (ss, sheet, fila, centro) => {
    const celda = sheet.getRange(fila, PCOL[tramo]);
    const antes = celda.getDisplayValue().trim();
    if (PROGRAMA.plancton.includes(antes)) return { ok: false, error: 'Los centros de Plancton Andino no llevan biólogo' };
    celda.setValue(comoTexto(sigla));
    marcarActualizado(sheet, fila, body.usuario);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, centro, tramo === 'semana' ? 'Biólogo lunes a viernes' : 'Biólogo sábado y domingo', antes, sigla);
    return { ok: true };
  });
}

// Cambio de centro: programar (centro nuevo + fecha estimada), completar (queda el nuevo) o quitar (se cancela)
function progCambioCentro(body) {
  return conCentro(body, (ss, sheet, fila, centro) => {
    const [nuevoActual, fechaActual] = sheet.getRange(fila, PCOL.nuevo, 1, 2).getDisplayValues()[0].map(x => String(x).trim());
    const antes = centro + (nuevoActual ? ' → ' + nuevoActual + (fechaActual ? ' (' + fechaActual + ')' : '') : '');
    let despues, cambio;
    if (body.modo === 'completar' || body.modo === 'quitar') {
      if (!nuevoActual) return { ok: false, error: 'Ese centro no tiene un cambio pendiente' };
      if (body.modo === 'completar') sheet.getRange(fila, PCOL.centro).setValue(comoTexto(nuevoActual));
      sheet.getRange(fila, PCOL.nuevo, 1, 2).setValues([['', '']]);
      despues = body.modo === 'completar' ? nuevoActual : centro;
      cambio = body.modo === 'completar' ? 'Cambio de centro completado' : 'Cambio de centro cancelado';
    } else {
      const nuevo = String(body.nuevo || '').trim(), fecha = String(body.fecha || '').trim();
      if (!nuevo) return { ok: false, error: 'Escribe el centro nuevo' };
      sheet.getRange(fila, PCOL.nuevo, 1, 2).setValues([[comoTexto(nuevo), comoTexto(fecha)]]);
      despues = centro + ' → ' + nuevo + (fecha ? ' (' + fecha + ')' : '');
      cambio = nuevoActual ? 'Cambio de centro editado' : 'Cambio de centro programado';
    }
    marcarActualizado(sheet, fila, body.usuario);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, centro, cambio, antes, despues);
    return { ok: true };
  });
}

// Centro nuevo: queda al final de su empresa (o al final si la empresa es nueva)
function progAgregarCentro(body) {
  const centro = String(body.centro || '').trim(), empresa = String(body.empresa || '').trim();
  if (!centro || !empresa) return { ok: false, error: 'Escribe la empresa y el centro' };
  const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
  const { sheet } = prepararPrograma(ss);   // fuera del bloqueo: prepararPrograma usa su propio bloqueo
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const v = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, PROG_COLUMNAS.length).getDisplayValues() : [];
    if (v.some(f => normalizar(f[PCOL.centro - 1]) === normalizar(centro))) return { ok: false, error: 'Ese centro ya está en el programa' };
    const id = Math.max(0, ...v.map(f => Number(f[0])).filter(n => !isNaN(n))) + 1;
    let despuesDe = -1, empresaFinal = empresa;
    v.forEach((f, k) => { if (normalizar(f[PCOL.empresa - 1]) === normalizar(empresa)) { despuesDe = k; empresaFinal = String(f[PCOL.empresa - 1]).trim(); } });
    const fila = despuesDe === -1 ? v.length + 2 : despuesDe + 3;
    if (fila <= sheet.getLastRow()) sheet.insertRowBefore(fila);
    sheet.getRange(fila, 1, 1, PROG_COLUMNAS.length).setValues([[id, comoTexto(empresaFinal), comoTexto(centro), '', '', comoTexto(body.semana), comoTexto(body.finde),
      comoTexto(body.ubicacion), comoTexto(body.barrio), body.usuario || '', new Date()]]);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, centro, 'Centro agregado', '', [empresaFinal, body.semana, body.finde].filter(Boolean).join(' · '));
    return { ok: true, id: id };
  } finally {
    lock.releaseLock();
  }
}

function progQuitarCentro(body) {
  return conCentro(body, (ss, sheet, fila, centro) => {
    const antes = sheet.getRange(fila, 1, 1, PROG_COLUMNAS.length - 2).getDisplayValues()[0].join(' · ');
    sheet.deleteRow(fila);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, centro, 'Centro quitado del programa', antes, '');
    return { ok: true };
  });
}

function textoTiempo(inicio, fin) {
  const dias = Math.round((fin - inicio) / 86400000) + 1;
  return dias % 7 === 0 ? (dias / 7) + (dias === 7 ? ' semana' : ' semanas') : dias + (dias === 1 ? ' día' : ' días');
}

// Crea (sin fila) o edita (con fila) una segunda muestra
function progSegunda(body) {
  const centro = String(body.centro || '').trim(), biologa = String(body.biologa || '').trim();
  const inicio = aFechaHoja(body.inicio), fin = aFechaHoja(body.fin);
  if (!centro || !biologa) return { ok: false, error: 'Elige el centro y la bióloga' };
  if (!inicio || !fin) return { ok: false, error: 'Revisa las fechas' };
  if (fin < inicio) return { ok: false, error: 'La fecha de fin es anterior al inicio' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
    const sheet = hojaSegunda(ss);
    const datos = [comoTexto(centro), comoTexto(biologa), textoTiempo(inicio, fin), inicio, fin, comoTexto(body.horario)];
    let fila = Number(body.fila), antes = '';
    if (body.fila) {
      if (!Number.isInteger(fila) || fila < 2 || fila > sheet.getLastRow()) return { ok: false, conflicto: true, error: 'Ese registro ya no existe. Recarga la página.' };
      const actual = sheet.getRange(fila, 1, 1, 2).getDisplayValues()[0];
      if (actual[0].trim() !== String(body.centroAntes || '').trim()) return { ok: false, conflicto: true, error: 'La planilla cambió. Recarga la página.' };
      antes = actual.join(' · ');
      sheet.getRange(fila, 1, 1, 6).setValues([datos]);
      sheet.getRange(fila, 9).setValue(comoTexto(body.observacion));
    } else {
      fila = sheet.getLastRow() + 1;
      sheet.getRange(fila, 1, 1, 9).setValues([datos.concat([body.usuario || '', new Date(), comoTexto(body.observacion)])]);
    }
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, centro, body.fila ? 'Segunda muestra editada' : 'Segunda muestra asignada', antes,
      biologa + ' · ' + body.inicio + ' a ' + body.fin + (body.horario ? ' · ' + body.horario : ''));
    return { ok: true, fila: fila };
  } finally {
    lock.releaseLock();
  }
}

function progEliminarSegunda(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(PROGRAMA.archivoId);
    const sheet = ss.getSheetByName(PROGRAMA.segunda);
    const fila = Number(body.fila);
    if (!sheet || !Number.isInteger(fila) || fila < 2 || fila > sheet.getLastRow()) return { ok: false, conflicto: true, error: 'Ese registro ya no existe. Recarga la página.' };
    const actual = sheet.getRange(fila, 1, 1, 5).getDisplayValues()[0];
    if (actual[0].trim() !== String(body.centroAntes || '').trim()) return { ok: false, conflicto: true, error: 'La planilla cambió. Recarga la página.' };
    sheet.deleteRow(fila);
    SpreadsheetApp.flush();
    historialPrograma(ss, body.usuario, actual[0], 'Segunda muestra eliminada', actual.join(' · '), '');
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// Escribe G a K, avisa a n8n y deja registro. Común a cerrar y gestión.
function guardarCierre(ss, sheet, fila, body, estado, tipo, fechaHora, webhook) {
  const observaciones = comoTexto(body.observaciones);
  sheet.getRange(fila, COL.estado, 1, 5).setValues([[estado, body.tecnico, tipo, observaciones, fechaHora]]);
  SpreadsheetApp.flush();
  registrar([new Date(), body.usuario || '', 'Solicitudes', fila, body.id, '(ticket → ' + estado + ')', tipo, observaciones]);

  const f = sheet.getRange(fila, 1, 1, COL.problema).getDisplayValues()[0];
  // El "id" que manda la planilla es el texto de la lista: "189 / Ensenada / 28-10 / Bomba sin cebar"
  const payload = {
    id: [f[0], f[COL.centro - 1], String(f[COL.fecha - 1]).slice(0, 5), f[COL.problema - 1]].join(' / '),
    telefono: f[COL.telefono - 1], centro: f[COL.centro - 1], problema: f[COL.problema - 1],
    tecnico: body.tecnico, tipoServicio: tipo, observaciones: observaciones, fechaHora: fechaHora, estado: estado
  };
  let aviso = '';
  if (PropertiesService.getScriptProperties().getProperty('WEBHOOKS') === 'no') {
    aviso = 'Avisos a n8n apagados (propiedad WEBHOOKS = no).';
  } else {
    try {
      enviarConReintento(webhook, {
        method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
      });
    } catch (err) {
      aviso = 'Se guardó, pero no se pudo avisar a n8n: ' + err.message;
    }
  }
  return { ok: true, estado: estado, aviso: aviso };
}

// Ubica el ticket por fila y confirma que ahí sigue el mismo ID antes de escribir
function conTicket(body, accion) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SOLICITUDES.archivoId);
    const sheet = ss.getSheetByName(SOLICITUDES.pestana);
    const fila = Number(body.fila);
    if (!Number.isInteger(fila) || fila < 2 || fila > sheet.getLastRow()) return { ok: false, error: 'Fila no válida' };
    if (sheet.getRange(fila, COL.id).getDisplayValue().trim() !== String(body.id)) {
      return { ok: false, conflicto: true, error: 'El ticket cambió de lugar en la planilla. Recarga la página.' };
    }
    return accion(ss, sheet, fila);
  } finally {
    lock.releaseLock();
  }
}

// Varias opciones en una celda de lista múltiple: "A, B"; las que llevan coma van entre comillas
function unirOpciones(lista) {
  if (!Array.isArray(lista)) lista = [lista];
  return lista.map(v => String(v || '').trim()).filter(Boolean)
    .map(v => v.includes(',') && !/^".*"$/.test(v) ? '"' + v + '"' : v).join(', ');
}

function ahora(ss, formato) {
  return Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), formato);
}

// El motivo del último intento fallido se devuelve a la página y queda en "Ejecuciones" del editor
function enviarConReintento(url, opciones) {
  let motivo = '';
  for (let intento = 1; intento <= 3; intento++) {
    try {
      const respuesta = UrlFetchApp.fetch(url, opciones);
      const code = respuesta.getResponseCode();
      if (code >= 200 && code < 300) return;
      motivo = 'n8n respondió ' + code + ': ' + respuesta.getContentText().slice(0, 200);
    } catch (err) {
      motivo = err.message;
      // Sin el permiso para conectarse a servicios externos no sirve reintentar
      if (/permis|authoriz|autoriz/i.test(motivo)) break;
    }
    if (intento < 3) Utilities.sleep(1500 * intento);
  }
  console.error('Aviso a n8n falló: ' + url + ' → ' + motivo);
  throw new Error(motivo);
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

// Correr a mano después de pegar una versión nueva: pide el permiso para conectarse a
// servicios externos (lo necesitan los avisos a n8n) y comprueba que n8n responde.
// Solo abre la página principal de n8n: no dispara ningún flujo ni manda mensajes.
function probarAvisos() {
  const props = PropertiesService.getScriptProperties();
  Logger.log('Propiedad WEBHOOKS: ' + (props.getProperty('WEBHOOKS') || '(no está: avisos encendidos)'));
  const r = UrlFetchApp.fetch('https://n8n-n8n.ax8fpu.easypanel.host/', { muteHttpExceptions: true });
  Logger.log('n8n respondió ' + r.getResponseCode() + ' (200 = conexión OK)');
}

// Prueba de lectura sin pasar por la web app.
function probar() {
  const r = leer(HOJAS.sams);
  Logger.log('Columnas: ' + JSON.stringify(r.columnas));
  Logger.log('Filas: ' + r.filas.length);
  Logger.log('Ejemplo: ' + JSON.stringify(r.filas[0]));
}
