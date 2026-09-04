import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import ExcelJS from 'exceljs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PORT = 3001;
const TARGET_FILE = path.resolve(__dirname, 'local.xlsx');

async function ensureTargetFile() {
  if (!fs.existsSync(TARGET_FILE)) {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');
    sheet.getCell('A1').value = '欢迎使用在线实时 Excel 编辑器！';
    sheet.getCell('A2').value = '在这里编辑的任何单元格都会在 1.5 秒后自动静默写回本地磁盘。';
    sheet.getColumn(1).width = 40;
    sheet.getColumn(2).width = 60;
    await workbook.xlsx.writeFile(TARGET_FILE);
    console.log(`[OK] 成功创建了示例数据文件: ${TARGET_FILE}`);
  }
}

const server = http.createServer(async (req, res) => {
  // 设置 CORS 跨域响应头，方便前端安全对接
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(200);
    res.end();
    return;
  }

  const url = new URL(req.url || '', `http://${req.headers.host}`);

  // 1. 状态接口
  if (url.pathname === '/api/status') {
    await ensureTargetFile();
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({ status: 'running', fileName: 'local.xlsx', path: TARGET_FILE }));
    return;
  }

  // 2. 加载数据接口 (将 Excel 文件流返回给前端)
  if (url.pathname === '/api/load') {
    await ensureTargetFile();
    const fileStream = fs.createReadStream(TARGET_FILE);
    res.writeHead(200, { 'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    fileStream.pipe(res);
    return;
  }

  // 3. 实时防抖同步保存接口 (直接将请求流 Pipe 写入本地文件)
  if (url.pathname === '/api/save' && req.method === 'POST') {
    const writeStream = fs.createWriteStream(TARGET_FILE);
    req.pipe(writeStream);
    req.on('end', () => {
      console.log(`[Sync] ${new Date().toLocaleTimeString()} 自动防抖保存成功！物理路径: ${TARGET_FILE}`);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true }));
    });
    req.on('error', (err) => {
      console.error('[Error] 保存失败:', err);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: false, error: err.message }));
    });
    return;
  }

  // 4. 前端报错捕获同步接口
  if (url.pathname === '/api/log-error' && req.method === 'POST') {
    let body = '';
    req.on('data', chunk => { body += chunk; });
    req.on('end', () => {
      try {
        const errorData = JSON.parse(body);
        const logMsg = `\n=======================================================\n` +
                       `[BROWSER CRASH DETECTED] ${new Date().toLocaleString()}\n` +
                       `Message: ${errorData.message}\n` +
                       `Stack: ${errorData.stack || 'No stack'}\n` +
                       `ActiveTab: ${errorData.activeTab || 'none'}\n` +
                       `=======================================================\n`;
        console.error(logMsg);
        fs.appendFileSync(path.resolve(__dirname, 'browser-errors.log'), logMsg);
        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: true }));
      } catch (e) {
        res.writeHead(400, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ success: false, error: 'Invalid JSON' }));
      }
    });
    return;
  }

  res.writeHead(404);
  res.end('Not Found');
});

ensureTargetFile().then(() => {
  server.listen(PORT, () => {
    console.log(`=======================================================`);
    console.log(`🚀 Node.js 物理文件伴侣服务器已启动！`);
    console.log(`🔗 前端连接地址: http://localhost:${PORT}`);
    console.log(`📁 物理绑定文件: ${TARGET_FILE}`);
    console.log(`=======================================================`);
  });
});
