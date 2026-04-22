// 간단한 정적 파일 서버 (http://localhost 에서 fetch 허용)
// 사용: node server.js
const http = require('http');
const fs   = require('fs');
const path = require('path');
const { generateQuote } = require('./generate_quote');

const ROOT = __dirname;
const PORT = process.env.PORT ? Number(process.env.PORT) : 5173;

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.svg': 'image/svg+xml',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.map': 'application/json; charset=utf-8',
};

function safeJoin(base, target) {
  const targetPath = path.normalize(target).replace(/^([/\\])+/, '');
  const finalPath = path.join(base, targetPath);
  if (!finalPath.startsWith(base)) throw new Error('Invalid path');
  return finalPath;
}

const server = http.createServer((req, res) => {
  try {
    const url      = (req.url || '/').split('?')[0];
    const pathname = decodeURIComponent(url);

    /* ── GET /api/templates : template/ 폴더의 xlsx 목록 반환 ── */
    if (req.method === 'GET' && pathname === '/api/templates') {
      const templateDir = path.join(ROOT, 'template');
      const files = fs.readdirSync(templateDir)
        .filter(f => f.endsWith('.xlsx'))
        .map(f => ({ filename: f, label: f.replace(/\.xlsx$/i, '') }));
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8', 'Access-Control-Allow-Origin': '*' });
      res.end(JSON.stringify(files));
      return;
    }

    /* ── POST /api/generate-excel : 엑셀 견적서 생성 ── */
    if (req.method === 'POST' && pathname === '/api/generate-excel') {
      let body = '';
      req.on('data', chunk => { body += chunk; });
      req.on('end', async () => {
        try {
          const data   = JSON.parse(body);
          const buffer = await generateQuote(data);
          const fname  = encodeURIComponent(
            `${data.client || 'OPENWORKS'}_${data.date || ''}.xlsx`
          );
          res.writeHead(200, {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': `attachment; filename*=UTF-8''${fname}`,
            'Content-Length': buffer.length,
            'Access-Control-Allow-Origin': '*',
          });
          res.end(buffer);
        } catch (e) {
          console.error('[generate-excel ERROR]', e.message, e.stack);
          res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({ error: e.message }));
        }
      });
      return;
    }

    /* ── OPTIONS preflight ── */
    if (req.method === 'OPTIONS') {
      res.writeHead(204, { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Content-Type' });
      res.end();
      return;
    }

    const rel      = pathname === '/' ? '/index.html' : pathname;
    const filePath = safeJoin(ROOT, rel);

    if (!fs.existsSync(filePath)) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME[ext] || 'application/octet-stream';
    res.writeHead(200, { 'Content-Type': contentType });
    fs.createReadStream(filePath).pipe(res);
  } catch (e) {
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(String(e && e.message ? e.message : e));
  }
});

server.listen(PORT, () => {
  console.log(`Server running: http://localhost:${PORT}`);
});

