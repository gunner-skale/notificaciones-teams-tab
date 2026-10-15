// Teams tab for Consegsa task notifications: serves public/ and proxies /api/* to the backend, so the browser talks
// to a single origin (no CORS) and the reverse proxy only needs to point here. No npm dependencies (Node 18+).
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');

const PORT = Number(process.env.PORT || 6090);
const BACKEND_URL = (process.env.BACKEND_URL || 'http://localhost:6080').replace(/\/$/, '');
const PUBLIC_DIR = path.join(__dirname, 'public');

// Teams shows the tab inside an iframe: allow Teams (web, desktop, Outlook/M365 hosts) as ancestors and no X-Frame-Options
const SECURITY_HEADERS = {
  'Content-Security-Policy': [
    "default-src 'self'",
    "script-src 'self' https://res.cdn.office.net",
    "style-src 'self'",
    "connect-src 'self'",
    "img-src 'self' data:",
    'frame-ancestors https://teams.microsoft.com https://*.teams.microsoft.com https://*.office.com https://*.cloud.microsoft https://*.microsoft365.com',
  ].join('; '),
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
};

const TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.svg': 'image/svg+xml' };

function send(res, status, body, type) {
  res.writeHead(status, { ...SECURITY_HEADERS, 'Content-Type': type, 'Cache-Control': 'no-store' });
  res.end(body);
}

function sendJson(res, status, obj) {
  send(res, status, JSON.stringify(obj), 'application/json; charset=utf-8');
}

// GET /api/tarea?t=<signed link> -> backend GET /integracion/teams/tab/tarea?t=
async function proxyTarea(url, res) {
  const t = url.searchParams.get('t') || '';
  try {
    const r = await fetch(`${BACKEND_URL}/integracion/teams/tab/tarea?t=${encodeURIComponent(t)}`, { signal: AbortSignal.timeout(10000) });
    sendJson(res, r.status, await r.json());
  } catch (e) {
    console.error(`[teams-tab] backend ${BACKEND_URL} no responde: ${e.message}`);
    sendJson(res, 502, { code: 502, server_message: 'No se pudo contactar al servidor.', data: [] });
  }
}

function serveStatic(url, res) {
  const relativo = url.pathname === '/' ? 'index.html' : decodeURIComponent(url.pathname).replace(/^\/+/, '');
  const archivo = path.normalize(path.join(PUBLIC_DIR, relativo));
  if (!archivo.startsWith(PUBLIC_DIR + path.sep)) {
    return send(res, 404, 'Not found', 'text/plain');
  }
  fs.readFile(archivo, (err, contenido) => {
    if (err) return send(res, 404, 'Not found', 'text/plain');
    send(res, 200, contenido, TYPES[path.extname(archivo)] || 'application/octet-stream');
  });
}

http.createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  // Access log without the signed link (t) or other query values
  res.on('finish', () => {
    console.log(`[teams-tab] ${new Date().toISOString()} ${req.method} ${url.pathname} ${res.statusCode} ${req.headers['sec-fetch-dest'] || '-'} ${(req.headers['user-agent'] || '-').slice(0, 80)}`);
  });
  if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'Method not allowed', 'text/plain');
  if (url.pathname === '/salud') return sendJson(res, 200, { ok: true });
  if (url.pathname === '/api/tarea') return proxyTarea(url, res);
  serveStatic(url, res);
}).listen(PORT, () => {
  console.log(`[teams-tab] http://localhost:${PORT} -> backend ${BACKEND_URL}`);
});
