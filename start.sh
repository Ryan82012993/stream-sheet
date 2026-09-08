#!/bin/bash

# 定义颜色
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m' # 无颜色

echo -e "${BLUE}=========================================${NC}"
echo -e "${BLUE}       StreamSheet 一键启动工具         ${NC}"
echo -e "${BLUE}=========================================${NC}"

# 1. 检测 Node.js 环境
if ! command -v node &> /dev/null; then
    echo -e "${RED}[错误] 未检测到 Node.js，请先安装 Node.js (推荐 v16.x 或更高版本)。${NC}"
    exit 1
fi
echo -e "${GREEN}[OK] 检测到 Node.js: $(node -v)${NC}"

# 2. 检查依赖包是否安装
if [ ! -d "node_modules" ]; then
    echo -e "${YELLOW}[提示] 未检测到依赖包，正在自动运行 npm install...${NC}"
    npm install
    if [ $? -ne 0 ]; then
        echo -e "${RED}[错误] 依赖安装失败，请手动运行 npm install 后再试。${NC}"
        exit 1
    fi
    echo -e "${GREEN}[OK] 依赖安装完成！${NC}"
else
    echo -e "${GREEN}[OK] 检测到已安装依赖包。${NC}"
fi

# 3. 进程清理逻辑（优雅退出，防止端口占用）
pids=()
cleanup() {
    echo -e "\n${YELLOW}[提示] 正在关闭 StreamSheet 所有服务...${NC}"
    for pid in "${pids[@]}"; do
        if kill -0 "$pid" 2>/dev/null; then
            kill "$pid" 2>/dev/null
        fi
    done
    echo -e "${GREEN}[OK] 所有服务已成功关闭！${NC}"
    exit 0
}

# 捕获 Ctrl+C 和终止信号
trap cleanup INT TERM

# 4. 启动后端 Companion 服务
echo -e "${BLUE}[1/2] 正在启动 Companion 同步服务端 (3001 端口)...${NC}"
npm run server &
pids+=($!)

# 稍微等待 1.5 秒，让后端完全初始化 local.xlsx
sleep 1.5

# 5. 启动前端 Vite 服务
echo -e "${BLUE}[2/2] 正在启动 Web 网页客户端 (3000 端口)...${NC}"
npm run dev &
pids+=($!)

echo -e "${GREEN}=========================================${NC}"
echo -e "${GREEN}✨ StreamSheet 服务已全部启动成功！${NC}"
echo -e "- 前端客户端: ${BLUE}http://localhost:3000${NC}"
echo -e "- 后端同步端: ${BLUE}http://localhost:3001${NC}"
echo -e "${YELLOW}提示: 保持此窗口开启。按 Ctrl+C 可以一键安全停止所有服务。${NC}"
echo -e "${GREEN}=========================================${NC}"

# 挂起脚本以保持等待
wait
