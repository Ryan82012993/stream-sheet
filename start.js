import { spawn, execSync } from 'child_process';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

console.log('\x1b[36m%s\x1b[0m', '=== StreamSheet 一键启动工具 ===');

// 1. 检查 node_modules 目录
const nodeModulesPath = path.resolve(__dirname, 'node_modules');
if (!fs.existsSync(nodeModulesPath)) {
  console.log('\x1b[33m%s\x1b[0m', '[提示] 未检测到依赖包，正在自动为您安装，请稍候...');
  try {
    execSync('npm install', { stdio: 'inherit' });
    console.log('\x1b[32m%s\x1b[0m', '[OK] 依赖包安装成功！');
  } catch (error) {
    console.error('\x1b[31m%s\x1b[0m', '[错误] 依赖安装失败，请尝试在终端手动运行 "npm install"');
    process.exit(1);
  }
} else {
  console.log('\x1b[32m%s\x1b[0m', '[OK] 检测到已安装依赖包');
}

// 2. 启动服务
console.log('\x1b[36m%s\x1b[0m', '\n[1/2] 正在启动 Companion 同步服务端 (3001 端口)...');

const isWindows = process.platform === 'win32';
const npmCmd = isWindows ? 'npm.cmd' : 'npm';

const serverProcess = spawn(npmCmd, ['run', 'server'], {
  stdio: ['inherit', 'pipe', 'pipe']
});

serverProcess.stdout.on('data', (data) => {
  const output = data.toString();
  output.split('\n').forEach(line => {
    if (line.trim()) {
      console.log(`\x1b[35m[Backend]\x1b[0m ${line}`);
    }
  });
});

serverProcess.stderr.on('data', (data) => {
  const errOutput = data.toString().trim();
  if (errOutput) {
    console.error(`\x1b[31m[Backend Error]\x1b[0m ${errOutput}`);
  }
});

// 延迟 1.5 秒后启动前端
setTimeout(() => {
  console.log('\x1b[36m%s\x1b[0m', '\n[2/2] 正在启动 Web 网页客户端 (3000 端口)...');
  
  const devProcess = spawn(npmCmd, ['run', 'dev'], {
    stdio: ['inherit', 'pipe', 'pipe']
  });

  devProcess.stdout.on('data', (data) => {
    const output = data.toString();
    output.split('\n').forEach(line => {
      if (line.trim()) {
        console.log(`\x1b[32m[Frontend]\x1b[0m ${line}`);
      }
    });
  });

  devProcess.stderr.on('data', (data) => {
    const errOutput = data.toString().trim();
    if (errOutput) {
      console.error(`\x1b[31m[Frontend Error]\x1b[0m ${errOutput}`);
    }
  });

  // 处理进程关闭
  const cleanup = () => {
    console.log('\x1b[33m%s\x1b[0m', '\n[提示] 正在关闭 StreamSheet 所有服务...');
    try {
      serverProcess.kill('SIGTERM');
    } catch (e) {}
    try {
      devProcess.kill('SIGTERM');
    } catch (e) {}
    console.log('\x1b[32m%s\x1b[0m', '[OK] 所有服务已成功停止！');
    process.exit(0);
  };

  process.on('SIGINT', cleanup);
  process.on('SIGTERM', cleanup);

}, 1500);
