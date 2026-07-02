#!/usr/bin/env node

const http = require('http');
const fs = require('fs');
const path = require('path');
const { consultCedula } = require('./lib/cedula-service');

const PORT = Number(process.env.PORT) || 3456;
const ROOT = __dirname;

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.ico': 'image/x-icon',
};

function sendJson(res, statusCode, payload) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(JSON.stringify(payload));
}

function serveStatic(req, res) {
  const urlPath = req.url === '/' ? '/index.html' : req.url.split('?')[0];
  const safePath = path.normalize(urlPath).replace(/^(\.\.[/\\])+/, '');
  const filePath = path.join(ROOT, 'public', safePath);

  if (!filePath.startsWith(path.join(ROOT, 'public'))) {
    sendJson(res, 403, { error: 'Forbidden' });
    return;
  }

  if (urlPath.startsWith('/lib/')) {
    const libPath = path.join(ROOT, urlPath.slice(1));
    if (libPath.startsWith(path.join(ROOT, 'lib')) && fs.existsSync(libPath)) {
      const ext = path.extname(libPath);
      res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] ?? 'text/plain; charset=utf-8' });
      fs.createReadStream(libPath).pipe(res);
      return;
    }
  }

  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    sendJson(res, 404, { error: 'Not found' });
    return;
  }

  const ext = path.extname(filePath);
  res.writeHead(200, { 'Content-Type': MIME_TYPES[ext] ?? 'application/octet-stream' });
  fs.createReadStream(filePath).pipe(res);
}

async function handleConsult(req, res) {
  let body = '';
  req.on('data', (chunk) => {
    body += chunk;
    if (body.length > 10_000) {
      req.destroy();
    }
  });

  req.on('end', async () => {
    try {
      const payload = JSON.parse(body || '{}');
      const cedula = String(payload.cedula ?? '').replace(/\D/g, '');

      if (!cedula) {
        sendJson(res, 400, { error: 'Cédula inválida' });
        return;
      }

      const result = await consultCedula(cedula);
      sendJson(res, 200, result);
    } catch (error) {
      sendJson(res, 500, { error: error.message ?? 'Error interno' });
    }
  });
}

const server = http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/api/consultar') {
    handleConsult(req, res);
    return;
  }

  if (req.method === 'GET') {
    serveStatic(req, res);
    return;
  }

  sendJson(res, 405, { error: 'Method not allowed' });
});

server.listen(PORT, () => {
  console.log(`UI disponible en http://localhost:${PORT}`);
  console.log('El Excel se procesa en el navegador; el servidor solo consulta cédulas.');
});
