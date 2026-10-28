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
    bloqueadas: ['Días restantes', 'Fecha mantención']
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
    hojas: Object.entries(HOJAS).map(([id, h]) => ({ id, titulo: h.titulo }))
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

app.listen(PORT, () => console.log('Mantenciones escuchando en puerto ' + PORT));
