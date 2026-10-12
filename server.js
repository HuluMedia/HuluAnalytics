/**
 * 运营数据分析 - 本地静态服务
 * 零依赖(Node >= 18)
 *
 * 数据完全在浏览器本地解析(上传的 xlsx 文件不经过服务端),
 * 本服务只负责托管 public/ 下的前端页面。
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.csv': 'text/csv; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

const PUBLIC_DIR = path.join(__dirname, 'public');

function serveStatic(req, res) {
  const urlPath = decodeURIComponent(req.url.split('?')[0]);

  // 开发者 CSV 从项目根目录提供
  if (urlPath.startsWith('/dev-')) {
    const devFile = path.join(__dirname, '项目详情-' + urlPath.slice(5) + '.csv');
    fs.readFile(devFile, (err, buf) => {
      if (err) { res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' }); res.end('Not Found'); return; }
      res.writeHead(200, { 'Content-Type': MIME['.csv'] });
      res.end(buf);
    });
    return;
  }

  const filePath = path.join(PUBLIC_DIR, urlPath === '/' ? 'index.html' : urlPath);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }
  fs.readFile(filePath, (err, buf) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Not Found');
      return;
    }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(filePath)] || 'application/octet-stream' });
    res.end(buf);
  });
}

const server = http.createServer(serveStatic);
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`服务已启动: http://localhost:${PORT}`);
});
