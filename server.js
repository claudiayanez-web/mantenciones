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

app.listen(PORT, () => console.log('Mantenciones escuchando en puerto ' + PORT));
