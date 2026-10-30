const http = require('http');
const url = require('url');
const fs = require('fs');
const path = require('path');
const { buscarMatriculaSisa } = require('./buscar_sisa');

const PORT = process.env.PORT || 3050;
const cache = new Map();

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

const server = http.createServer(async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const parsedUrl = url.parse(req.url, true);

  // API SISA Endpoint
  if (parsedUrl.pathname === '/api/sisa') {
    const matricula = parsedUrl.query.matricula || '';
    if (!matricula) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: 'Parámetro matricula faltante' }));
      return;
    }

    const cleanMat = String(matricula).replace(/\D/g, '');
    if (cache.has(cleanMat)) {
      console.log(`[SISA CACHE HIT] Matrícula: ${cleanMat}`);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(cache.get(cleanMat)));
      return;
    }

    console.log(`[SISA BUSCANDO] Consultando matrícula: ${cleanMat}...`);
    try {
      const data = await buscarMatriculaSisa(cleanMat);
      if (data && data.ok) {
        cache.set(cleanMat, data);
      }
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(data));
    } catch (err) {
      console.error('[SISA ERROR]', err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ ok: false, error: String(err.message || err) }));
    }
    return;
  }

  // Static files server for eFarm
  let filePath = path.join(__dirname, parsedUrl.pathname === '/' ? 'index.html' : parsedUrl.pathname);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME_TYPES[ext] || 'application/octet-stream';

  fs.readFile(filePath, (err, content) => {
    if (err) {
      if (err.code === 'ENOENT') {
        res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
        res.end('404 Not Found');
      } else {
        res.writeHead(500);
        res.end(`Server Error: ${err.code}`);
      }
    } else {
      res.writeHead(200, { 'Content-Type': contentType });
      res.end(content);
    }
  });
});

server.listen(PORT, () => {
  console.log(`===============================================`);
  console.log(` Servidor eFarm + SISA activo en:`);
  console.log(` http://localhost:${PORT}`);
  console.log(` Endpoint API: http://localhost:${PORT}/api/sisa?matricula=123456`);
  console.log(`===============================================`);
});
