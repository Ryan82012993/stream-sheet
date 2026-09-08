@echo off
chcp 65001 > nul
title StreamSheet 一键启动工具

echo =========================================
echo        StreamSheet 一键启动工具
echo =========================================

:: 1. 检测 Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Node.js，请先安装 Node.js ^(推荐 v16.x 或更高版本^)。
    pause
    exit /b 1
)
for /f "tokens=*" %%i in ('node -v') do set NODE_VER=%%i
echo [OK] 检测到 Node.js: %NODE_VER%

:: 2. 检测依赖
if not exist node_modules (
    echo [提示] 未检测到依赖包，正在自动运行 npm install...
    call npm install
    if %errorlevel% neq 0 (
        echo [错误] 依赖安装失败，请手动运行 npm install 后再试。
        pause
        exit /b 1
    )
    echo [OK] 依赖安装完成！
) else (
    echo [OK] 检测到已安装依赖包。
)

:: 3. 启动服务
echo [1/2] 正在启动 Companion 同步服务端 (3001 端口)...
start "StreamSheet Backend" /b npm run server

:: 等待两秒以确保后端初始化
timeout /t 2 /nobreak > nul

echo [2/2] 正在启动 Web 网页客户端 (3000 端口)...
echo =========================================
echo ✨ StreamSheet 服务已全部启动！
echo - 前端客户端: http://localhost:3000
echo - 后端同步端: http://localhost:3001
echo 提示: 请保持此窗口打开以保持服务运行。直接关闭此窗口或按 Ctrl+C 可以停止。
echo =========================================

:: 启动前端作为前台进程，这样 Ctrl+C 会同时中断该窗口下所有的子进程
call npm run dev
