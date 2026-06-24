const http = require('http');
const fs = require('fs');
const path = require('path');

const root = 'C:/260624';

const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
};

http.createServer((req, res) => {
  const urlPath = req.url === '/' ? '/chapter3_exchange_app.html' : req.url.split('?')[0];
  const filePath = path.join(root, urlPath);
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.statusCode = 404;
    res.end('not found');
    return;
  }
  res.setHeader('Content-Type', types[path.extname(filePath).toLowerCase()] || 'text/plain; charset=utf-8');
  fs.createReadStream(filePath).pipe(res);
}).listen(8000, '127.0.0.1', () => {
  console.log('serving http://127.0.0.1:8000');
});

setInterval(() => {}, 1 << 30);
