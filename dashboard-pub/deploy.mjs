import { createServer } from 'node:http';
import { readFileSync, existsSync } from 'node:fs';
import { join, extname } from 'node:path';

const PORT = 3000;
const FILE = join(import.meta.dirname, 'index.html');
const html = readFileSync(FILE, 'utf-8');

const mime = {
  '.html': 'text/html; charset=UTF-8',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.json': 'application/json',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const server = createServer((req, res) => {
  const url = new URL(req.url, 'http://localhost');
  let path = url.pathname;
  if (path === '/') path = '/index.html';
  const filePath = join(import.meta.dirname, path);
  if (path === '/index.html' || !existsSync(filePath)) {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=UTF-8', 'Access-Control-Allow-Origin': '*', 'Cache-Control': 'no-cache' });
    res.end(html);
  } else {
    const ext = extname(filePath).toLowerCase();
    res.writeHead(200, { 'Content-Type': mime[ext] || 'application/octet-stream' });
    res.end(readFileSync(filePath));
  }
});

server.listen(PORT, '127.0.0.1', () => {
  console.log(`✅ 로컬 서버: http://127.0.0.1:${PORT}`);
  console.log(`\n⏳ 터널 열기: 대시보드 파일 열기...`);
  console.log(`\n📋 팀원 공유용 임시 링크 생성 중...`);
  console.log(`(아래 링크를 열려면 브라우저에서 접속하세요)`);
  console.log(`\n  => netlify.com 에서 드래그&드롭 배포 권장 <=`);
});

process.on('SIGINT', () => { server.close(); process.exit(); });
process.on('SIGTERM', () => { server.close(); process.exit(); });
