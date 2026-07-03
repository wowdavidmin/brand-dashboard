import { createServer } from 'node:http';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname } from 'node:path';

const PORT = 3000;
const FILE = join(import.meta.dirname, 'index.html');

const mimeTypes = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const html = readFileSync(FILE, 'utf-8');

createServer((req, res) => {
  let url = new URL(req.url, 'http://localhost');
  let path = url.pathname;
  if (path === '/') path = '/index.html';
  const filePath = join(import.meta.dirname, path);
  
  if (path === '/index.html' || path === '/') {
    res.writeHead(200, {
      'Content-Type': 'text/html; charset=UTF-8',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-cache',
    });
    res.end(html);
  } else if (existsSync(filePath)) {
    const ext = extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': mimeTypes[ext] || 'application/octet-stream' });
    res.end(readFileSync(filePath));
  } else {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=UTF-8' });
    res.end(html);
  }
}).listen(PORT, () => {
  console.log(`Server running on http://127.0.0.1:${PORT}`);
});
