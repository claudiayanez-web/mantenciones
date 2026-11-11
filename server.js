const express = require('express');
const crypto = require('crypto');
const path = require('path');

const app = express();
app.set('trust proxy', 1);
app.use(express.json({ limit: '100kb' }));
app.use(express.static(path.join(__dirname, 'public')));

const {
  PORT = 3000,
  MANTENCIONES_SCRIPT_URL,
  MANTENCIONES_SECRETO,
  USUARIOS = '',
  ADMINS = '',
  TECNICOS = '',
  SESION_SECRETO
} = process.env;

if (!MANTENCIONES_SCRIPT_URL || !MANTENCIONES_SECRETO || !SESION_SECRETO || !USUARIOS) {
  console.error('Faltan variables de entorno: MANTENCIONES_SCRIPT_URL, MANTENCIONES_SECRETO, USUARIOS o SESION_SECRETO');
  process.exit(1);
}

// ---------- Planillas habilitadas ----------
// bloqueadas: columnas que no se pueden editar desde la página aunque en el Sheet sean valores
const HOJAS = {
  sams: {
    titulo: 'SAMS',
    bloqueadas: ['Días restantes', 'Fecha mantención'],
    desactivable: true   // un administrador puede pasar un centro al Historial
  },
  // Periodos terminados. Cada fila tiene un Código único (Centro-AAAA-MM) que genera la página
  historial: {
    titulo: 'Historial',
    bloqueadas: []
  }
};

// ---------- Usuarios y sesión ----------

const usuarios = Object.fromEntries(
  USUARIOS.split(',')
    .map(par => par.trim())
    .filter(Boolean)
    .map(par => {
      const i = par.indexOf(':');
      return [par.slice(0, i).trim().toLowerCase(), par.slice(i + 1).trim()];
    })
);

// Administradores: pueden eliminar centros. ADMINS=claudia (separados por coma)
const admins = ADMINS.split(',').map(u => u.trim().toLowerCase()).filter(Boolean);
function esAdmin(usuario) {
  return admins.includes(usuario);
}

// Nombre de técnico de cada usuario, como aparece en la lista de técnicos de la planilla Mantenciones.
// TECNICOS=sebastian:Sebastián Arzola,wilmer:Wilmer Ruiz (separados por coma)
const tecnicos = Object.fromEntries(
  TECNICOS.split(',')
    .map(par => par.trim())
    .filter(par => par.includes(':'))
    .map(par => {
      const i = par.indexOf(':');
      return [par.slice(0, i).trim().toLowerCase(), par.slice(i + 1).trim()];
    })
);

const DURACION_SESION_MS = 12 * 60 * 60 * 1000; // 12 horas

function firmar(texto) {
  return crypto.createHmac('sha256', SESION_SECRETO).update(texto).digest('base64url');
}

function crearSesion(usuario) {
  const datos = Buffer.from(JSON.stringify({ u: usuario, exp: Date.now() + DURACION_SESION_MS })).toString('base64url');
  return datos + '.' + firmar(datos);
}

function leerSesion(req) {
  const cookies = Object.fromEntries(
    (req.headers.cookie || '').split(';').map(c => {
      const i = c.indexOf('=');
      return [c.slice(0, i).trim(), decodeURIComponent(c.slice(i + 1).trim())];
    })
  );
  const token = cookies.sesion;
  if (!token || !token.includes('.')) return null;

  const [datos, firma] = token.split('.');
  const esperada = firmar(datos);
  if (firma.length !== esperada.length || !crypto.timingSafeEqual(Buffer.from(firma), Buffer.from(esperada))) {
    return null;
  }
  try {
    const sesion = JSON.parse(Buffer.from(datos, 'base64url').toString());
    if (sesion.exp < Date.now() || !usuarios[sesion.u]) return null;
    return sesion.u;
  } catch {
    return null;
  }
}

function claveCorrecta(usuario, clave) {
  const real = usuarios[usuario];
  if (!real || typeof clave !== 'string') return false;
  const a = Buffer.from(firmar(real));
  const b = Buffer.from(firmar(clave));
  return crypto.timingSafeEqual(a, b);
}

function requiereSesion(req, res, next) {
  const usuario = leerSesion(req);
  if (!usuario) return res.status(401).json({ ok: false, error: 'Sesión vencida. Vuelve a ingresar.' });
  req.usuario = usuario;
  next();
}

// Guías de uso: las capturas muestran centros y nombres, así que solo con sesión.
// Sin sesión se vuelve a la página principal, que muestra el login.
app.use('/ayuda', (req, res, next) => leerSesion(req) ? next() : res.redirect('/'),
  express.static(path.join(__dirname, 'ayuda')));

// Límite de intentos de login por IP: 5 cada 15 minutos
const intentos = new Map();
function puedeIntentar(ip) {
  const ahora = Date.now();
  const lista = (intentos.get(ip) || []).filter(t => ahora - t < 15 * 60 * 1000);
  intentos.set(ip, lista);
  return lista.length < 5;
}
function registrarIntento(ip) {
  intentos.get(ip).push(Date.now());
}

// ---------- Llamada al Apps Script ----------

async function llamarScript(cuerpo) {
  const respuesta = await fetch(MANTENCIONES_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ ...cuerpo, secreto: MANTENCIONES_SECRETO }),
    redirect: 'follow'
  });
  const texto = await respuesta.text();
  try {
    return JSON.parse(texto);
  } catch {
    throw new Error('El Apps Script no respondió JSON. Revisa que esté publicado como aplicación web.');
  }
}

// ---------- Rutas ----------

app.post('/api/login', (req, res) => {
  const ip = req.ip;
  if (!puedeIntentar(ip)) {
    return res.status(429).json({ ok: false, error: 'Demasiados intentos. Espera 15 minutos.' });
  }
  const usuario = String(req.body.usuario || '').trim().toLowerCase();
  const clave = req.body.clave;

  if (!claveCorrecta(usuario, clave)) {
    registrarIntento(ip);
    return res.status(401).json({ ok: false, error: 'Usuario o clave incorrectos' });
  }

  res.cookie('sesion', crearSesion(usuario), {
    httpOnly: true,
    secure: req.secure,
    sameSite: 'strict',
    maxAge: DURACION_SESION_MS
  });
  res.json({ ok: true, usuario });
});

app.post('/api/logout', (req, res) => {
  res.clearCookie('sesion');
  res.json({ ok: true });
});

app.get('/api/yo', requiereSesion, (req, res) => {
  res.json({
    ok: true,
    usuario: req.usuario,
    admin: esAdmin(req.usuario),
    tecnico: tecnicos[req.usuario] || '',
    hojas: Object.entries(HOJAS).map(([id, h]) => ({ id, titulo: h.titulo, desactivable: !!h.desactivable }))
  });
});

app.get('/api/hojas/:hoja', requiereSesion, async (req, res) => {
  const config = HOJAS[req.params.hoja];
  if (!config) return res.status(404).json({ ok: false, error: 'Planilla no encontrada' });

  try {
    const datos = await llamarScript({ accion: 'leer', hoja: req.params.hoja });
    if (!datos.ok) return res.status(502).json(datos);

    datos.columnas = datos.columnas.map(c => ({
      ...c,
      editable: c.editable && !config.bloqueadas.includes(c.nombre)
    }));
    res.json(datos);
  } catch (err) {
    console.error('Error al leer', req.params.hoja, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo leer la planilla. Intenta de nuevo.' });
  }
});

app.post('/api/hojas/:hoja/editar', requiereSesion, async (req, res) => {
  const config = HOJAS[req.params.hoja];
  if (!config) return res.status(404).json({ ok: false, error: 'Planilla no encontrada' });

  const { fila, columna, clave, anterior, nuevo } = req.body;
  if (config.bloqueadas.includes(columna)) {
    return res.status(403).json({ ok: false, error: 'Esta columna no se edita desde la página' });
  }

  try {
    const resultado = await llamarScript({
      accion: 'editar',
      hoja: req.params.hoja,
      usuario: req.usuario,
      fila, columna, clave, anterior, nuevo
    });
    res.status(resultado.ok ? 200 : 409).json(resultado);
  } catch (err) {
    console.error('Error al editar', req.params.hoja, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  }
});

// Agrega una fila nueva al final de la hoja. celdas: { nombreColumna: valor }
app.post('/api/hojas/:hoja/agregar', requiereSesion, async (req, res) => {
  const config = HOJAS[req.params.hoja];
  if (!config) return res.status(404).json({ ok: false, error: 'Planilla no encontrada' });

  const celdas = req.body.celdas;
  if (!celdas || typeof celdas !== 'object' || Array.isArray(celdas)) {
    return res.status(400).json({ ok: false, error: 'Datos inválidos' });
  }
  // Solo textos, y nunca las columnas bloqueadas (esas se calculan en el Sheet)
  const limpias = {};
  for (const [columna, valor] of Object.entries(celdas)) {
    if (config.bloqueadas.includes(columna) || typeof valor !== 'string') continue;
    if (valor.trim()) limpias[columna] = valor.trim();
  }
  if (!Object.keys(limpias).length) {
    return res.status(400).json({ ok: false, error: 'Completa al menos el nombre del centro' });
  }

  try {
    const resultado = await llamarScript({
      accion: 'agregar',
      hoja: req.params.hoja,
      usuario: req.usuario,
      celdas: limpias
    });
    res.status(resultado.ok ? 200 : 409).json(resultado);
  } catch (err) {
    console.error('Error al agregar', req.params.hoja, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo agregar. Intenta de nuevo.' });
  }
});

// Elimina la fila de un centro. Solo administradores.
app.post('/api/hojas/:hoja/eliminar', requiereSesion, async (req, res) => {
  const config = HOJAS[req.params.hoja];
  if (!config) return res.status(404).json({ ok: false, error: 'Planilla no encontrada' });
  if (!esAdmin(req.usuario)) {
    return res.status(403).json({ ok: false, error: 'Solo un administrador puede eliminar centros' });
  }

  const { fila, clave } = req.body;
  if (!Number.isInteger(fila) || typeof clave !== 'string' || !clave.trim()) {
    return res.status(400).json({ ok: false, error: 'Datos inválidos' });
  }

  try {
    const resultado = await llamarScript({
      accion: 'eliminar',
      hoja: req.params.hoja,
      usuario: req.usuario,
      fila, clave
    });
    res.status(resultado.ok ? 200 : 409).json(resultado);
  } catch (err) {
    console.error('Error al eliminar', req.params.hoja, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo eliminar. Intenta de nuevo.' });
  }
});

// Pasa un centro vigente al Historial (lo copia allá y lo saca de la hoja). Solo administradores.
app.post('/api/hojas/:hoja/desactivar', requiereSesion, async (req, res) => {
  const config = HOJAS[req.params.hoja];
  if (!config || !config.desactivable) return res.status(404).json({ ok: false, error: 'Esta planilla no se puede desactivar' });
  if (!esAdmin(req.usuario)) {
    return res.status(403).json({ ok: false, error: 'Solo un administrador puede desactivar centros' });
  }

  const { fila, clave, fin } = req.body;
  if (!Number.isInteger(fila) || typeof clave !== 'string' || !clave.trim()
    || typeof fin !== 'string' || !/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(fin)) {
    return res.status(400).json({ ok: false, error: 'Datos inválidos' });
  }

  try {
    const resultado = await llamarScript({
      accion: 'desactivar',
      hoja: req.params.hoja,
      usuario: req.usuario,
      fila, clave, fin
    });
    res.status(resultado.ok ? 200 : 409).json(resultado);
  } catch (err) {
    console.error('Error al desactivar', req.params.hoja, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo desactivar. Intenta de nuevo.' });
  }
});

// ---------- Solicitudes (tickets de la planilla Mantenciones) ----------
// El Apps Script escribe lo mismo que los botones de la pestaña "ingreso" y avisa a n8n.

const texto = (v, max = 2000) => typeof v === 'string' && v.trim().length <= max ? v.trim() : null;
const listaTextos = v => Array.isArray(v) && v.length && v.length <= 20 && v.every(x => texto(x, 200)) ? v.map(x => x.trim()) : null;

app.get('/api/solicitudes', requiereSesion, async (req, res) => {
  try {
    const datos = await llamarScript({ accion: 'solicitudesLeer' });
    res.status(datos.ok ? 200 : 502).json(datos);
  } catch (err) {
    console.error('Error al leer solicitudes', err.message);
    res.status(502).json({ ok: false, error: 'No se pudieron leer las solicitudes. Intenta de nuevo.' });
  }
});

// Cada acción valida sus datos y arma lo que se manda al Apps Script
const ACCIONES_SOLICITUD = {
  aceptar: b => texto(b.tecnico, 100) && { accion: 'solicitudesAceptar', tecnico: b.tecnico.trim() },
  cerrar: b => texto(b.tecnico, 100) && listaTextos(b.tipos) && texto(b.observaciones ?? '') !== null &&
    { accion: 'solicitudesCerrar', tecnico: b.tecnico.trim(), tipos: listaTextos(b.tipos), observaciones: (b.observaciones || '').trim() },
  gestion: b => texto(b.tecnico, 100) && ['centro', 'terreno'].includes(b.gestion) && texto(b.observaciones ?? '') !== null &&
    { accion: 'solicitudesGestion', tecnico: b.tecnico.trim(), gestion: b.gestion, observaciones: (b.observaciones || '').trim() },
  centro: b => texto(b.centro, 200) && { accion: 'solicitudesCentro', centro: b.centro.trim() },
  // Hay que escribir el número del ticket y un motivo; el script lo vuelve a comprobar
  eliminar: b => texto(b.motivo, 500) && b.motivo.trim().length >= 5 && String(b.confirmacion ?? '').trim() === String(b.id) &&
    { accion: 'solicitudesEliminar', motivo: b.motivo.trim(), confirmacion: String(b.confirmacion).trim() }
};

app.post('/api/solicitudes/:accion', requiereSesion, async (req, res) => {
  const b = req.body || {};
  let cuerpo;
  if (req.params.accion === 'nuevo') {
    cuerpo = texto(b.centro, 200) && texto(b.tecnico, 100) && listaTextos(b.problemas) && texto(b.observaciones ?? '') !== null &&
      { accion: 'solicitudesNuevo', centro: b.centro.trim(), tecnico: b.tecnico.trim(), problemas: listaTextos(b.problemas), observaciones: (b.observaciones || '').trim() };
  } else {
    const armar = ACCIONES_SOLICITUD[req.params.accion];
    if (!armar) return res.status(404).json({ ok: false, error: 'Acción no válida' });
    // El ticket se identifica por su fila y su ID (se comprueba que coincidan antes de escribir)
    if (!Number.isInteger(b.fila) || !/^\d+$/.test(String(b.id || ''))) {
      return res.status(400).json({ ok: false, error: 'Datos inválidos' });
    }
    cuerpo = armar(b);
    if (cuerpo) Object.assign(cuerpo, { fila: b.fila, id: String(b.id) });
  }
  if (!cuerpo) return res.status(400).json({ ok: false, error: 'Faltan datos' });

  // Cada técnico actúa a su nombre; solo un administrador puede elegir a otro técnico
  if (!esAdmin(req.usuario)) {
    if (!tecnicos[req.usuario]) {
      return res.status(403).json({ ok: false, error: 'Tu usuario no tiene un técnico asignado. Pídeselo a un administrador.' });
    }
    cuerpo.tecnico = tecnicos[req.usuario];
  }

  try {
    const resultado = await llamarScript({ ...cuerpo, usuario: req.usuario });
    res.status(resultado.ok ? 200 : 409).json(resultado);
  } catch (err) {
    console.error('Error en solicitud', req.params.accion, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  }
});

// ---------- Compensaciones ----------
// Solo un administrador registra y ve todo; cada técnico ve únicamente lo suyo.
const mismoNombre = (a, b) => String(a || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim() ===
  String(b || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

app.get('/api/compensaciones', requiereSesion, async (req, res) => {
  const admin = esAdmin(req.usuario);
  const tecnico = tecnicos[req.usuario];
  if (!admin && !tecnico) {
    return res.status(403).json({ ok: false, error: 'Tu usuario no tiene un técnico asignado. Pídeselo a un administrador.' });
  }
  try {
    const r = await llamarScript({ accion: 'compLeer', conCertificados: admin });
    if (r.ok && !admin) {
      r.tecnicos = (r.tecnicos || []).filter(t => mismoNombre(t, tecnico));
      r.salidas = (r.salidas || []).filter(s => mismoNombre(s['Técnico'], tecnico));
      r.libres = (r.libres || []).filter(s => mismoNombre(s['Técnico'], tecnico));
      delete r.zonas; delete r.certificados; delete r.descartados;
      // Técnicos remotos u otros que no están en el sistema de compensaciones: sin acceso
      if (!r.tecnicos.length) return res.status(403).json({ ok: false, sinAcceso: true, error: 'No tienes acceso a compensaciones.' });
    }
    res.status(r.ok ? 200 : 502).json(r);
  } catch (err) {
    console.error('Error al leer compensaciones', err.message);
    res.status(502).json({ ok: false, error: 'No se pudieron leer las compensaciones.' });
  }
});

const fechaValida = v => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v);
const ACCIONES_COMP = {
  salida: b => texto(b.tecnico, 100) && texto(b.centro, 200) && fechaValida(b.desde) && fechaValida(b.hasta) &&
    { accion: 'compSalida', id: b.id ?? '', tecnico: b.tecnico, centro: b.centro, desde: b.desde, hasta: b.hasta, dias: b.dias, tramos: b.tramos,
      domingos: b.domingos, feriados: b.feriados, retornoDomingo: b.retornoDomingo, observacion: texto(b.observacion ?? '') ?? '' },
  libre: b => texto(b.tecnico, 100) && fechaValida(b.desde) && fechaValida(b.hasta) &&
    { accion: 'compLibre', id: b.id ?? '', tecnico: b.tecnico, desde: b.desde, hasta: b.hasta, dias: b.dias, observacion: texto(b.observacion ?? '') ?? '' },
  eliminar: b => ['salidas', 'libres'].includes(b.tipo) && Number.isInteger(b.id) && { accion: 'compEliminar', tipo: b.tipo, id: b.id },
  zona: b => texto(b.centro, 200) && [true, false, null].includes(b.compensa) && { accion: 'compZona', centro: b.centro, compensa: b.compensa },
  descartar: b => texto(b.clave, 300) && { accion: 'compDescartar', clave: b.clave, tecnico: texto(b.tecnico ?? '', 100) ?? '',
    centro: texto(b.centro ?? '', 200) ?? '', fecha: texto(b.fecha ?? '', 20) ?? '', deshacer: !!b.deshacer }
};

app.post('/api/compensaciones/:accion', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede registrar compensaciones' });
  const armar = ACCIONES_COMP[req.params.accion];
  if (!armar) return res.status(404).json({ ok: false, error: 'Acción no válida' });
  const cuerpo = armar(req.body || {});
  if (!cuerpo) return res.status(400).json({ ok: false, error: 'Faltan datos' });
  try {
    const r = await llamarScript({ ...cuerpo, usuario: req.usuario });
    res.status(r.ok ? 200 : 409).json(r);
  } catch (err) {
    console.error('Error en compensaciones', req.params.accion, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  }
});

// ---------- Acceso a la plataforma de clientes ----------
// Checklist de centros (SAMS, Chile y México) a los que ya se les dio acceso. Solo administradores.
app.get('/api/plataforma', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede ver este checklist' });
  try {
    const r = await llamarScript({ accion: 'platLeer' });
    res.status(r.ok ? 200 : 502).json(r);
  } catch (err) {
    console.error('Error al leer acceso plataforma', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo leer el checklist.' });
  }
});

app.post('/api/plataforma', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede marcar centros' });
  const b = req.body || {};
  const centro = texto(b.centro ?? '', 200);
  if (!centro || typeof b.listo !== 'boolean') return res.status(400).json({ ok: false, error: 'Faltan datos' });
  const cuerpo = { accion: 'platMarcar', centro, empresa: texto(b.empresa ?? '', 200) ?? '', listo: b.listo, usuario: req.usuario };
  if (b.nota !== undefined) cuerpo.nota = texto(b.nota ?? '', 300) ?? '';
  try {
    const r = await llamarScript(cuerpo);
    res.status(r.ok ? 200 : 409).json(r);
  } catch (err) {
    console.error('Error al marcar acceso plataforma', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  }
});

// ---------- Notion: Operaciones y checklists ----------
// Se lee y se marca directo en Notion (integración "Página operaciones"); Notion sigue siendo el original.
// Operaciones la ven todos; los checklists de instalación y desinstalación, solo administradores.

const NOTION_API = process.env.NOTION_API || 'https://api.notion.com';
const NOTION = {
  operaciones: process.env.NOTION_OPERACIONES || 'afe5b2f5-6dff-83b5-9d1b-8728bc5baf8f',
  instalacion: process.env.NOTION_INSTALACION || '2b05b2f5-6dff-81d9-b59d-000b70ede3f5',
  desinstalacion: process.env.NOTION_DESINSTALACION || '35f5b2f5-6dff-8176-af0c-000bd4ac4fd9'
};
// Pasos en el mismo orden que se ven en la tarjeta de Notion; si aparece una casilla nueva, se agrega al final
const PASOS = {
  instalacion: { titulo: 'ID Mantención', estado: '¿Instalado?', pasos: ['Contacto Centro', 'Enviar info a Centro', 'Disponibilidad',
    'coordinar hora, fecha y lugar', 'Acreditación/permiso  técnico', 'Asignar biólogo', 'Equipo operativo', 'Materiales preparados (checklist)',
    'Instalación OK', 'Agregar certificado de servicio', 'Certificado de servicio enviado', 'Agregar SAMS', 'Avisar nuevo centro grupo PASPA',
    'Enviar programa actualizado PASPA', 'QUICKSIGHT: agregar centro panel global y empresa', 'GTR: activación GTR',
    'QUICK O GTR: Enviar correo utilización', 'Agregar RindeGastos', 'Agregar Agente', 'Dispositivo funcionando correctamente',
    'Muestras subiendo a Quicksight'] },
  desinstalacion: { titulo: 'ID Mantención', estado: '¿Desinstalado?', pasos: ['Quitar asignación biólogo APP', 'Coordinar retiro con centro',
    'Avisar en PASPA termino ciclo', 'Enviar correo con fecha desinstalación', 'Confirmar acreditación técnico', 'Quitar de agente',
    'Quitar de certificado servicio', 'Parar servicio', 'Quitar de Quicksight', 'Quitar de SAMS', 'Desactivar dispositivo y centro de la APP'] }
};

async function notion(ruta, opciones = {}) {
  if (!process.env.NOTION_TOKEN) throw Object.assign(new Error('Falta NOTION_TOKEN'), { sinToken: true });
  const r = await fetch(NOTION_API + ruta, {
    ...opciones,
    headers: { Authorization: 'Bearer ' + process.env.NOTION_TOKEN, 'Notion-Version': '2025-09-03', 'Content-Type': 'application/json' }
  });
  const datos = await r.json().catch(() => ({}));
  if (!r.ok) throw Object.assign(new Error(datos.message || 'Notion respondió ' + r.status), { estado: r.status, codigo: datos.code });
  return datos;
}

async function leerFuente(id) {
  const paginas = [];
  let cursor;
  do {
    const r = await notion('/v1/data_sources/' + id + '/query', { method: 'POST', body: JSON.stringify({ page_size: 100, start_cursor: cursor }) });
    paginas.push(...r.results.filter(p => !p.archived && !p.in_trash));
    cursor = r.has_more ? r.next_cursor : undefined;
  } while (cursor);
  return paginas;
}

// Valor simple de una propiedad de Notion
function valorNotion(p) {
  if (!p) return null;
  switch (p.type) {
    case 'title': case 'rich_text': return p[p.type].map(t => t.plain_text).join('');
    case 'select': case 'status': return p[p.type]?.name || '';
    case 'multi_select': return p.multi_select.map(o => o.name);
    case 'date': return p.date ? { desde: p.date.start, hasta: p.date.end } : null;
    case 'checkbox': return p.checkbox;
    case 'url': return p.url || '';
    case 'last_edited_time': return p.last_edited_time;
    default: return null;
  }
}

function errorNotion(res, err, que) {
  if (err.sinToken) return res.status(503).json({ ok: false, error: 'Falta conectar Notion: agrega NOTION_TOKEN en Easypanel.' });
  console.error('Error con Notion', que, err.estado || '', err.message);
  const sinPermiso = err.estado === 404 || err.estado === 403 || err.codigo === 'object_not_found';
  res.status(502).json({ ok: false, error: sinPermiso
    ? 'Notion no deja ver ' + que + '. Revisa que la integración "Página operaciones" esté agregada en Conexiones.'
    : 'No se pudo leer ' + que + ' desde Notion. Intenta de nuevo.' });
}

// Se guarda un minuto en memoria para no consultar Notion en cada vuelta
const cacheNotion = {};
async function leerConCache(clave, leer, recargar) {
  const c = cacheNotion[clave];
  if (!recargar && c && Date.now() - c.en < 60000) return c.datos;
  const datos = await leer();
  cacheNotion[clave] = { en: Date.now(), datos };
  return datos;
}

app.get('/api/operaciones', requiereSesion, async (req, res) => {
  try {
    const centros = await leerConCache('operaciones', async () => (await leerFuente(NOTION.operaciones)).map(p => {
      const v = n => valorNotion(p.properties[n]);
      return { id: p.id, url: p.url, nombre: v('Nombre') || '', empresa: v('Empresa') || '', estado: v('Estado') || '', zonas: v('Etiqueta') || [],
        ultimaVisita: v('Última visita')?.desde || '', proxima: v('Próxima mantención')?.desde || '', rango: v('Rango actual'),
        observaciones: v('Observaciones') || '', editado: p.last_edited_time };
    }).filter(c => c.nombre), req.query.recargar === '1');
    res.json({ ok: true, centros });
  } catch (err) { errorNotion(res, err, 'Operaciones'); }
});

// ---------- Fotos de Notion achicadas ----------
// Las fotos que suben desde el celular pesan varios MB. El servidor las baja una vez, las achica
// (miniatura para la ficha, mediana para el visor) y las guarda en memoria; la próxima vez salen al instante.
// Solo se aceptan fotos guardadas en Notion, para que esto no sirva para bajar cualquier cosa de internet.
const sharp = require('sharp');
const HOST_FOTOS = /^(prod-files-secure\.s3\.[a-z0-9-]+\.amazonaws\.com|s3\.[a-z0-9-]+\.amazonaws\.com|file\.notion\.so|img\.notionusercontent\.com)$/i;
const TAMANOS_FOTO = { chica: { lado: 320, calidad: 70 }, grande: { lado: 1600, calidad: 82 } };
const cacheFotos = new Map();   // clave -> Buffer (el más antiguo primero)
let bytesCacheFotos = 0;
const fotosEnCamino = new Map(); // misma foto pedida varias veces a la vez: se procesa una sola
const LIMITE_CACHE_FOTOS = 80 * 1024 * 1024;

function esFotoNotion(url) {
  try { const u = new URL(url); return u.protocol === 'https:' && HOST_FOTOS.test(u.hostname); } catch { return false; }
}
// Enlaces que usa la página: miniatura, mediana y la original (por si se quiere descargar)
const hostsAvisados = new Set();
function enlacesFoto(url) {
  if (!esFotoNotion(url)) {
    let host = ''; try { host = new URL(url).hostname; } catch {}
    if (host && !hostsAvisados.has(host)) { hostsAvisados.add(host); console.log('Foto de un sitio no reconocido, se muestra sin achicar:', host); }
    return { chica: url, grande: url, original: url };
  }
  const p = t => '/api/foto?t=' + t + '&u=' + encodeURIComponent(url);
  return { chica: p('chica'), grande: p('grande'), original: url };
}
function guardarFotoEnCache(clave, b) {
  if (cacheFotos.has(clave)) return;
  cacheFotos.set(clave, b);
  bytesCacheFotos += b.length;
  for (const [k, v] of cacheFotos) {
    if (bytesCacheFotos <= LIMITE_CACHE_FOTOS) break;
    cacheFotos.delete(k); bytesCacheFotos -= v.length;
  }
}
// Las fotos de iPhone vienen en HEIC: ni los navegadores ni sharp las leen, así que se decodifican aparte
const decodificarHeic = require('heic-decode');
const esHeic = b => b.length > 12 && b.toString('latin1', 4, 8) === 'ftyp' && /^(heic|heix|hevc|hevx|heim|heis|mif1|msf1)$/.test(b.toString('latin1', 8, 12));
// Como mucho dos fotos procesándose a la vez: una foto de celular descomprimida ocupa unos 50 MB
let fotosProcesando = 0;
const esperandoTurno = [];
async function conTurno(trabajo) {
  if (fotosProcesando >= 2) await new Promise(r => esperandoTurno.push(r));
  fotosProcesando++;
  try { return await trabajo(); } finally {
    fotosProcesando--;
    const siguiente = esperandoTurno.shift();
    if (siguiente) siguiente();
  }
}
async function prepararFoto(original) {
  if (!esHeic(original)) return { entrada: original, opciones: { failOn: 'none' }, girar: true };
  const { width, height, data } = await decodificarHeic({ buffer: original });
  // libheif ya entrega la foto derecha, no hay que girarla
  return { entrada: Buffer.from(data.buffer, data.byteOffset, data.byteLength), opciones: { raw: { width, height, channels: 4 } }, girar: false };
}
const achicar = (foto, tamano) => {
  const { lado, calidad } = TAMANOS_FOTO[tamano];
  const s = sharp(foto.entrada, foto.opciones);
  return (foto.girar ? s.rotate() : s)
    .resize({ width: lado, height: lado, fit: 'inside', withoutEnlargement: true }).jpeg({ quality: calidad, mozjpeg: true }).toBuffer();
};
// Se baja la foto de Notion una sola vez: se prepara el tamaño pedido y, sin esperar, también el otro,
// para que al abrir el visor la grande ya esté lista
async function fotoAchicada(url, tamano) {
  const u = new URL(url);
  const base = u.hostname + u.pathname;  // sin la firma, que cambia cada vez que se lee Notion
  const clave = tamano + '|' + base;
  if (cacheFotos.has(clave)) {
    const b = cacheFotos.get(clave);
    cacheFotos.delete(clave); cacheFotos.set(clave, b);  // queda como la más reciente
    return b;
  }
  if (fotosEnCamino.has(clave)) return fotosEnCamino.get(clave);
  const claveOtro = (tamano === 'chica' ? 'grande' : 'chica') + '|' + base;
  const hacerOtro = !cacheFotos.has(claveOtro) && !fotosEnCamino.has(claveOtro);
  let entregar, fallar, entregarOtro, fallarOtro;
  const trabajo = new Promise((ok, mal) => { entregar = ok; fallar = mal; });
  fotosEnCamino.set(clave, trabajo);
  if (hacerOtro) {
    const otro = new Promise((ok, mal) => { entregarOtro = ok; fallarOtro = mal; });
    otro.catch(() => {});
    fotosEnCamino.set(claveOtro, otro);
  }
  // Todo dentro de un mismo turno: bajar, leer, achicar al tamaño pedido (se entrega apenas está) y luego al otro
  conTurno(async () => {
    try {
      const r = await fetch(url, { redirect: 'error' });
      if (!r.ok) throw Object.assign(new Error('Notion respondió ' + r.status), { estado: r.status });
      if (Number(r.headers.get('content-length') || 0) > 40 * 1024 * 1024) throw new Error('Foto demasiado grande');
      const foto = await prepararFoto(Buffer.from(await r.arrayBuffer()));
      const b = await achicar(foto, tamano);
      guardarFotoEnCache(clave, b);
      entregar(b);
      if (hacerOtro) {
        try { const b2 = await achicar(foto, claveOtro.split('|')[0]); guardarFotoEnCache(claveOtro, b2); entregarOtro(b2); }
        catch (e) { fallarOtro(e); }
      }
    } catch (e) {
      fallar(e);
      if (hacerOtro) fallarOtro(e);
    } finally {
      fotosEnCamino.delete(clave);
      if (hacerOtro) fotosEnCamino.delete(claveOtro);
    }
  });
  return trabajo;
}
app.get('/api/foto', requiereSesion, async (req, res) => {
  const url = String(req.query.u || '');
  const tamano = TAMANOS_FOTO[req.query.t] ? req.query.t : 'chica';
  if (!esFotoNotion(url)) return res.status(400).json({ ok: false, error: 'Foto no válida' });
  try {
    const b = await fotoAchicada(url, tamano);
    res.set('Content-Type', 'image/jpeg').set('Cache-Control', 'private, max-age=3000').send(b);
  } catch (err) {
    console.error('Error al achicar foto', err.estado || '', err.message);
    // Si el enlace venció (duran una hora) o no se pudo achicar, la página reintenta con la original
    res.status(err.estado === 403 ? 410 : 502).json({ ok: false, error: 'No se pudo cargar la foto' });
  }
});

// Contenido de la tarjeta (bitácora técnica): párrafos, títulos, listas y fotos, un nivel de sub-listas
function bloqueSimple(b) {
  const d = b[b.type] || {};
  const texto = (d.rich_text || []).map(t => ({ t: t.plain_text, b: !!t.annotations?.bold, i: !!t.annotations?.italic,
    u: !!t.annotations?.underline, link: t.href && /^https?:/.test(t.href) ? t.href : '' }));
  const base = { tipo: b.type, texto, hijos: b.has_children };
  if (b.type === 'to_do') base.hecho = !!d.checked;
  if (b.type === 'image') {
    base.imagen = d.type === 'external' ? d.external?.url : d.file?.url;
    if (base.imagen) base.foto = enlacesFoto(base.imagen);
  }
  if (b.type === 'image') base.texto = (d.caption || []).map(t => ({ t: t.plain_text }));
  return base;
}
async function leerBloques(id) {
  const r = await notion('/v1/blocks/' + id + '/children?page_size=100');
  return r.results.map(bloqueSimple);
}
app.get('/api/operaciones/:id/bitacora', requiereSesion, async (req, res) => {
  if (!UUID.test(req.params.id)) return res.status(400).json({ ok: false, error: 'Centro no válido' });
  try {
    const r = await notion('/v1/blocks/' + req.params.id + '/children?page_size=100');
    const bloques = [];
    for (const b of r.results) {
      const s = bloqueSimple(b);
      // Sub-listas (por ejemplo "Automático: N° 9" dentro de "Alimentación eléctrica")
      if (b.has_children && /list_item|to_do|toggle/.test(b.type)) s.hijos = await leerBloques(b.id);
      else s.hijos = [];
      bloques.push(s);
    }
    res.json({ ok: true, bloques });
  } catch (err) { errorNotion(res, err, 'la bitácora'); }
});

// Comentarios de la tarjeta (los de la página, no los que van pegados a una línea de la bitácora).
// Desde la página se comenta con la integración, así que el nombre de quien escribe va al inicio del comentario.
const nombresNotion = new Map();
async function autorComentario(c) {
  if (c.display_name?.resolved_name) return c.display_name.resolved_name;
  const id = c.created_by?.id;
  if (!id) return '';
  if (!nombresNotion.has(id)) {
    try { nombresNotion.set(id, (await notion('/v1/users/' + id)).name || ''); } catch { nombresNotion.set(id, ''); }
  }
  return nombresNotion.get(id);
}
const nombreVisible = u => tecnicos[u] || u.split('.').map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');

async function paginaDeOperaciones(id) {
  const pagina = await notion('/v1/pages/' + id);
  return String(pagina.parent?.data_source_id || '').replace(/-/g, '') === NOTION.operaciones.replace(/-/g, '');
}

app.get('/api/operaciones/:id/comentarios', requiereSesion, async (req, res) => {
  if (!UUID.test(req.params.id)) return res.status(400).json({ ok: false, error: 'Centro no válido' });
  try {
    const lista = [];
    let cursor;
    do {
      const r = await notion('/v1/comments?block_id=' + req.params.id + '&page_size=100' + (cursor ? '&start_cursor=' + encodeURIComponent(cursor) : ''));
      lista.push(...r.results);
      cursor = r.has_more ? r.next_cursor : undefined;
    } while (cursor);
    const comentarios = [];
    for (const c of lista) {
      // Fotos adjuntas al comentario (enlaces de Notion que vencen en una hora; se piden de nuevo al abrir la ficha)
      const fotos = (c.attachments || []).filter(a => a.category === 'image' || /^image\//.test(a.content_type || ''))
        .map(a => a.file?.url || a.url).filter(u => typeof u === 'string' && /^https:/.test(u)).map(enlacesFoto);
      comentarios.push({ id: c.id, fecha: c.created_time, autor: await autorComentario(c), esPagina: c.created_by?.type === 'bot', fotos,
        texto: (c.rich_text || []).map(t => ({ t: t.plain_text, b: !!t.annotations?.bold, i: !!t.annotations?.italic, u: !!t.annotations?.underline,
          link: t.href && /^https?:/.test(t.href) ? t.href : '' })) });
    }
    res.json({ ok: true, comentarios });
  } catch (err) {
    if (err.estado === 403) return res.status(502).json({ ok: false, error: 'Falta activar "Leer comentarios" en la integración de Notion.' });
    errorNotion(res, err, 'los comentarios');
  }
});

app.post('/api/operaciones/:id/comentarios', requiereSesion, async (req, res) => {
  const comentario = texto(req.body?.texto ?? '', 1800);
  if (!UUID.test(req.params.id) || !comentario) return res.status(400).json({ ok: false, error: 'Escribe el comentario' });
  try {
    if (!await paginaDeOperaciones(req.params.id)) return res.status(400).json({ ok: false, error: 'La tarjeta no es de Operaciones' });
    await notion('/v1/comments', { method: 'POST', body: JSON.stringify({ parent: { page_id: req.params.id },
      rich_text: [{ type: 'text', text: { content: nombreVisible(req.usuario) + ': ' }, annotations: { bold: true } }, { type: 'text', text: { content: comentario } }] }) });
    console.log('Comentario en Operaciones', req.params.id, 'por', req.usuario);
    res.json({ ok: true });
  } catch (err) {
    if (err.estado === 403) return res.status(502).json({ ok: false, error: 'Falta activar "Insertar comentarios" en la integración de Notion.' });
    errorNotion(res, err, 'los comentarios');
  }
});

app.get('/api/checklists', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede ver los checklists' });
  try {
    const leerTipo = async tipo => (await leerFuente(NOTION[tipo])).map(p => {
      const conf = PASOS[tipo];
      const casillas = Object.entries(p.properties).filter(([, x]) => x.type === 'checkbox').map(([n]) => n);
      const orden = conf.pasos.filter(n => casillas.includes(n)).concat(casillas.filter(n => !conf.pasos.includes(n)));
      const selects = Object.entries(p.properties).filter(([, x]) => x.type === 'select').map(([n, x]) => ({ nombre: n, valor: x.select?.name || '' }));
      return { id: p.id, url: p.url, tipo, centro: valorNotion(p.properties[conf.titulo]) || '(sin nombre)', estado: valorNotion(p.properties[conf.estado]) || 'Sin empezar',
        fecha: valorNotion(p.properties.Fecha)?.desde || '', pasos: orden.map(n => ({ nombre: n, hecho: p.properties[n].checkbox })), selects,
        creado: p.created_time, editado: p.last_edited_time };
    });
    const datos = await leerConCache('checklists', async () => ({ instalacion: await leerTipo('instalacion'), desinstalacion: await leerTipo('desinstalacion') }), req.query.recargar === '1');
    res.json({ ok: true, ...datos });
  } catch (err) { errorNotion(res, err, 'los checklists'); }
});

// Marcar un paso, el estado general o una selección (por ejemplo "GTR O QUICKSIGHT") de un checklist
app.post('/api/checklists/:id', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede marcar los checklists' });
  const b = req.body || {};
  const tipo = PASOS[b.tipo] ? b.tipo : null;
  const nombre = texto(b.propiedad ?? '', 120);
  if (!UUID.test(req.params.id) || !tipo || !nombre) return res.status(400).json({ ok: false, error: 'Faltan datos' });
  let propiedad;
  if (b.clase === 'paso' && typeof b.valor === 'boolean') propiedad = { checkbox: b.valor };
  else if (b.clase === 'estado' && nombre === PASOS[tipo].estado && ['Sin empezar', 'En progreso', 'Listo'].includes(b.valor)) propiedad = { status: { name: b.valor } };
  else if (b.clase === 'select' && texto(b.valor ?? '', 100) !== null) propiedad = { select: b.valor ? { name: b.valor } : null };
  else return res.status(400).json({ ok: false, error: 'Faltan datos' });
  try {
    // Se confirma que la tarjeta es de ese checklist antes de cambiarla
    const pagina = await notion('/v1/pages/' + req.params.id);
    if (String(pagina.parent?.data_source_id || '').replace(/-/g, '') !== NOTION[tipo].replace(/-/g, '')) return res.status(400).json({ ok: false, error: 'La tarjeta no es de este checklist' });
    if (!pagina.properties[nombre]) return res.status(400).json({ ok: false, error: 'Ese paso ya no existe en Notion. Recarga.' });
    await notion('/v1/pages/' + req.params.id, { method: 'PATCH', body: JSON.stringify({ properties: { [nombre]: propiedad } }) });
    delete cacheNotion.checklists;
    console.log('Checklist', tipo, req.params.id, nombre, JSON.stringify(b.valor), 'por', req.usuario);
    res.json({ ok: true });
  } catch (err) { errorNotion(res, err, 'el checklist'); }
});

// ---------- Wellboats ----------
// Posición de los buques (pestaña wellboat de SAMS). Todos la ven; nadie la edita desde aquí.
app.get('/api/wellboats', requiereSesion, async (req, res) => {
  try {
    const r = await llamarScript({ accion: 'wellboatLeer' });
    res.status(r.ok ? 200 : 502).json(r);
  } catch (err) {
    console.error('Error al leer wellboats', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo leer la posición de los wellboats.' });
  }
});

// ---------- Programa de biólogos ----------
// Todos lo ven; solo un administrador asigna, cambia centros y registra segundas muestras.
app.get('/api/programa', requiereSesion, async (req, res) => {
  try {
    const r = await llamarScript({ accion: 'progLeer' });
    res.status(r.ok ? 200 : 502).json(r);
  } catch (err) {
    console.error('Error al leer el programa', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo leer el programa de biólogos.' });
  }
});

const idYCentro = b => /^[0-9]+$/.test(String(b.id ?? '')) && texto(b.centro, 150);
const ACCIONES_PROGRAMA = {
  asignar: b => idYCentro(b) && ['semana', 'finde'].includes(b.tramo) && texto(b.sigla, 10) &&
    { accion: 'progAsignar', id: String(b.id), centro: b.centro, tramo: b.tramo, sigla: b.sigla.trim() },
  cambio: b => idYCentro(b) && ['programar', 'completar', 'quitar'].includes(b.modo) &&
    { accion: 'progCambioCentro', id: String(b.id), centro: b.centro, modo: b.modo, nuevo: texto(b.nuevo ?? '', 150) ?? '', fecha: texto(b.fecha ?? '', 40) ?? '' },
  agregar: b => texto(b.empresa, 100) && texto(b.centro, 150) &&
    { accion: 'progAgregarCentro', empresa: b.empresa.trim(), centro: b.centro.trim(), semana: texto(b.semana ?? '', 10) ?? '', finde: texto(b.finde ?? '', 10) ?? '',
      ubicacion: texto(b.ubicacion ?? '', 100) ?? '', barrio: texto(b.barrio ?? '', 20) ?? '' },
  quitar: b => idYCentro(b) && { accion: 'progQuitarCentro', id: String(b.id), centro: b.centro },
  segunda: b => texto(b.centro, 150) && texto(b.biologa, 100) && fechaValida(b.inicio) && fechaValida(b.fin) &&
    (b.fila == null || Number.isInteger(b.fila)) &&
    { accion: 'progSegunda', fila: b.fila ?? '', centroAntes: texto(b.centroAntes ?? '', 150) ?? '', centro: b.centro.trim(), biologa: b.biologa.trim(),
      inicio: b.inicio, fin: b.fin, horario: texto(b.horario ?? '', 100) ?? '', observacion: texto(b.observacion ?? '', 500) ?? '' },
  eliminarSegunda: b => Number.isInteger(b.fila) && texto(b.centroAntes, 150) &&
    { accion: 'progEliminarSegunda', fila: b.fila, centroAntes: b.centroAntes },
  // Reemplazo temporal de uno o varios centros del mismo titular
  reemplazo: b => texto(b.titular, 10) && ['semana', 'finde'].includes(b.tramo) && fechaValida(b.desde) && fechaValida(b.hasta) &&
    Array.isArray(b.centros) && b.centros.length > 0 && b.centros.length <= 80 &&
    b.centros.every(c => /^[0-9]+$/.test(String(c.id ?? '')) && texto(c.centro, 150) && texto(c.reemplazo, 10)) &&
    { accion: 'progReemplazo', titular: b.titular.trim(), tramo: b.tramo, desde: b.desde, hasta: b.hasta,
      motivo: texto(b.motivo ?? '', 60) ?? '', observacion: texto(b.observacion ?? '', 500) ?? '',
      centros: b.centros.map(c => ({ id: String(c.id), centro: c.centro.trim(), reemplazo: c.reemplazo.trim() })) },
  editarReemplazo: b => /^[0-9]+$/.test(String(b.id ?? '')) && texto(b.reemplazo, 10) && fechaValida(b.desde) && fechaValida(b.hasta) &&
    { accion: 'progEditarReemplazo', id: String(b.id), reemplazo: b.reemplazo.trim(), desde: b.desde, hasta: b.hasta,
      motivo: texto(b.motivo ?? '', 60) ?? '', observacion: texto(b.observacion ?? '', 500) ?? '' },
  eliminarReemplazo: b => /^[0-9]+$/.test(String(b.id ?? '')) && { accion: 'progEliminarReemplazo', id: String(b.id) },
  agregarBiologo: b => texto(b.nombre, 60) && /^[A-Za-zÁÉÍÓÚÑáéíóúñ]{1,4}$/.test(String(b.sigla ?? '').trim()) && ['semana', 'finde', 'reemplazo'].includes(b.equipo) &&
    { accion: 'progAgregarBiologo', nombre: b.nombre.trim(), sigla: b.sigla.trim(), equipo: b.equipo }
};

app.post('/api/programa/:accion', requiereSesion, async (req, res) => {
  if (!esAdmin(req.usuario)) return res.status(403).json({ ok: false, error: 'Solo un administrador puede editar el programa' });
  const armar = ACCIONES_PROGRAMA[req.params.accion];
  if (!armar) return res.status(404).json({ ok: false, error: 'Acción no válida' });
  const cuerpo = armar(req.body || {});
  if (!cuerpo) return res.status(400).json({ ok: false, error: 'Faltan datos' });
  try {
    const r = await llamarScript({ ...cuerpo, usuario: req.usuario });
    res.status(r.ok ? 200 : 409).json(r);
  } catch (err) {
    console.error('Error en programa', req.params.accion, err.message);
    res.status(502).json({ ok: false, error: 'No se pudo guardar. Intenta de nuevo.' });
  }
});

// ---------- Muestras de los equipos (base de Retina) ----------
// Solo lectura: la conexión se abre con todas las transacciones en modo lectura.
// Cuenta todas las muestras tomadas, estén analizadas o no.
// RETINA_DB_URL=postgres://usuario:clave@host:5432/base (la clave puede ir sin codificar)

function datosConexionRetina(url) {
  const m = String(url).trim().replace(/^["']|["']$/g, '')
    .match(/^postgres(?:ql)?:\/\/([^:]+):(.*)@([^@:\/]+)(?::(\d+))?\/([^\/?]+)$/);
  if (!m) return null;
  let password = m[2];
  try { password = decodeURIComponent(m[2]); } catch {}
  return { user: m[1], password, host: m[3], port: Number(m[4] || 5432), database: m[5] };
}

let poolRetina = null;
function conexionRetina() {
  if (poolRetina) return poolRetina;
  const datos = process.env.RETINA_DB_URL && datosConexionRetina(process.env.RETINA_DB_URL);
  if (!datos) return null;
  const { Pool } = require('pg');
  poolRetina = new Pool({ ...datos, max: 2, connectionTimeoutMillis: 15000, idleTimeoutMillis: 60000,
    options: '-c default_transaction_read_only=on' });
  poolRetina.on('error', err => console.error('Base de Retina', err.message));
  return poolRetina;
}

// "Hoy" es el día en Chile, aunque el equipo esté en Noruega
const SQL_MUESTRAS = `
  SELECT d.id AS equipo_id, d.code AS equipo, d.is_active AS activo, ce.id AS centro_id, ce.name AS centro, co.name AS empresa,
         count(s.id)::int AS total,
         count(s.id) FILTER (WHERE s.collected_at >= date_trunc('day', now() AT TIME ZONE 'America/Santiago') AT TIME ZONE 'America/Santiago')::int AS hoy,
         count(s.id) FILTER (WHERE s.collected_at >= now() - interval '7 days')::int AS semana,
         max(s.collected_at) AS ultima
  FROM devices d
  LEFT JOIN centers ce ON ce.id = d.center_id
  LEFT JOIN companies co ON co.id = ce.company_id
  LEFT JOIN samples s ON s.device_id = d.id AND s.deleted_at IS NULL
  WHERE d.deleted_at IS NULL
  GROUP BY d.id, ce.id, co.name`;

// Se guarda un minuto para no consultar la base cada vez que alguien abre la pestaña
let muestrasCache = null;
app.get('/api/muestras', requiereSesion, async (req, res) => {
  const pool = conexionRetina();
  if (!pool) return res.status(503).json({ ok: false, error: 'Falta configurar RETINA_DB_URL en el servidor.' });
  if (muestrasCache && Date.now() - muestrasCache.leido < 60000 && req.query.recargar !== '1') {
    return res.json(muestrasCache.respuesta);
  }
  try {
    const r = await pool.query(SQL_MUESTRAS);
    const respuesta = { ok: true, leido: new Date().toISOString(), equipos: r.rows };
    muestrasCache = { leido: Date.now(), respuesta };
    res.json(respuesta);
  } catch (err) {
    console.error('Error al leer muestras', err.code || '', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo conectar a la base de Retina.' });
  }
});

// Fecha y hora (de Chile) de cada muestra que tomó un equipo, entre dos días. Máximo 62 días por consulta.
const SQL_MUESTRAS_EQUIPO = `
  SELECT to_char(s.collected_at AT TIME ZONE 'America/Santiago', 'YYYY-MM-DD') AS fecha,
         to_char(s.collected_at AT TIME ZONE 'America/Santiago', 'HH24:MI:SS') AS hora,
         dp.name AS profundidad, dp.value_meters AS metros
  FROM samples s
  LEFT JOIN depths dp ON dp.id = s.depth_id
  WHERE s.device_id = $1 AND s.deleted_at IS NULL
    AND s.collected_at >= ($2::date)::timestamp AT TIME ZONE 'America/Santiago'
    AND s.collected_at < ($3::date + 1)::timestamp AT TIME ZONE 'America/Santiago'
  ORDER BY s.collected_at`;

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
app.get('/api/muestras/equipo/:id', requiereSesion, async (req, res) => {
  const { desde, hasta } = req.query;
  if (!UUID.test(req.params.id) || !fechaValida(desde) || !fechaValida(hasta) || desde > hasta) {
    return res.status(400).json({ ok: false, error: 'Faltan datos' });
  }
  if ((new Date(hasta) - new Date(desde)) / 86400000 > 62) {
    return res.status(400).json({ ok: false, error: 'Elige un periodo de máximo 2 meses.' });
  }
  const pool = conexionRetina();
  if (!pool) return res.status(503).json({ ok: false, error: 'Falta configurar RETINA_DB_URL en el servidor.' });
  try {
    const r = await pool.query(SQL_MUESTRAS_EQUIPO, [req.params.id, desde, hasta]);
    res.json({ ok: true, muestras: r.rows });
  } catch (err) {
    console.error('Error al leer muestras del equipo', err.code || '', err.message);
    res.status(502).json({ ok: false, error: 'No se pudo conectar a la base de Retina.' });
  }
});

app.listen(PORT, () => console.log('Mantenciones escuchando en puerto ' + PORT));
