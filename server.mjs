import https from 'node:https';
import { readFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

const base = fileURLToPath(new URL('./public/', import.meta.url));
const certDirectory = join(homedir(), '.office-addin-dev-certs');

const mimeTypes = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.png': 'image/png',
  '.json': 'application/json'
};

async function main() {
  let key, cert;
  try {
    [key, cert] = await Promise.all([
      readFile(join(certDirectory, 'localhost.key')),
      readFile(join(certDirectory, 'localhost.crt'))
    ]);
  } catch {
    throw new Error('Sertifikat localhost belum tersedia.');
  }

  const server = https.createServer({ key, cert }, async (request, response) => {
    response.setHeader('Access-Control-Allow-Origin', '*');
    response.setHeader('Access-Control-Allow-Methods', 'GET, HEAD, OPTIONS');
    response.setHeader('Access-Control-Allow-Private-Network', 'true');
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Cache-Control', 'no-store');

    if (request.method === 'OPTIONS') { response.writeHead(204); response.end(); return; }
    if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405); response.end(); return; }

    let pathname = new URL(request.url, 'https://localhost:3000').pathname;
    if (pathname === '/') pathname = '/taskpane.html';

    const ext = pathname.substring(pathname.lastIndexOf('.'));
    const contentType = mimeTypes[ext] || 'application/octet-stream';

    try {
      const filePath = join(base, pathname.replace(/^\//, ''));
      const body = await readFile(filePath);
      response.writeHead(200, { 'Content-Type': contentType });
      response.end(request.method === 'HEAD' ? undefined : body);
    } catch {
      response.writeHead(404);
      response.end('File tidak ditemukan');
    }
  });

  server.on('error', error => {
    console.error(error.code === 'EADDRINUSE' ? 'Port 3000 sedang dipakai aplikasi lain. Hentikan aplikasi tersebut lalu jalankan lagi.' : error.message);
    process.exitCode = 1;
  });

  server.listen(3000, '127.0.0.1', () => {
    console.log('✅ Server lokal aktif di: https://localhost:3000/taskpane.html');
    console.log('Biarkan terminal ini tetap terbuka selama menguji di Excel.');
    console.log('Tekan Ctrl+C untuk menghentikan server.');
  });
}

main().catch(err => { console.error(err.message); process.exitCode = 1; });
