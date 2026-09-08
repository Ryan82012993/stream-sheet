# StreamSheet ⚡ 

> **Local-First, Real-Time Web Excel Editor with Dual-Mode Stream Sync.**
>
> 这是一个基于 **React + TypeScript + FortuneSheet** 构建的**本地优先（Local-First）** 实时 Excel 在线编辑器。支持现代浏览器物理句柄直接读写与轻量级 Node.js 伴侣流同步（Companion Stream Server）的双模全自动保存机制，兼具极高的本地安全隐私性与开箱即用的同步体验。

---

## ✨ 核心特色 (Key Features)

### 1. 💻 本地优先双模同步 (Dual-Mode Sync Architecture)
*   **现代物理直读写 (Modern Local-First)**: 默认采用现代浏览器的 **File System Access API** (`showOpenFilePicker` 和 `createWritable`)。一次授权，即可在 Web 端静默物理写盘。
*   **伴侣流同步服务 (Companion Stream Server)**: 针对其他浏览器或本地调试，StreamSheet 自动侦测并建立与本地轻量级 Node 服务的同步。支持 **防抖自动保存 (Debounced Auto-Save)**，在最后一次修改后延时 1.5 秒即启动流式物理写盘。

### 2. 💾 生产级物理原子落盘 (Atomic File Write & Temp Swap)
*   为了防止保存中途因网络波动、前端崩溃或 Node.js 意外挂掉导致原 Excel 文件被截断损坏，伴侣服务器采用**临时写入-完全成功后原子替换**（Write-then-Swap）策略。
*   数据保存时，流管道先写入临时文件 `local.xlsx.tmp`，仅在数据完全接收成功后调用 `fs.rename` 原子覆盖。**即便遇到灾难性断电，您的原 `local.xlsx` 数据库也绝对 100% 完好无损！**

### 3. 🗂️ 独创的离线/在线多标签工作簿空间 (Multi-Tab Workspace)
*   支持同时打开/创建多个本地物理文件句柄、远程同步工作簿和离线临时沙盒，提供标签页无缝切换。
*   独创 **“离场数据净化（Structured Clone Isolation）”** 技术与高性能数据清洗过滤层（`sheetSanitizer.ts`）。在切换、新建、删除标签页时，使用原生 `structuredClone` 深度复制并剥离 FortuneSheet 挂载的 DOM 脏引用，裁剪公式计算链，彻底规避只读冲突与渲染器并发激活崩溃。
*   新工作簿一键秒级创建，默认载入最符合 Microsoft Excel 初始观感的 **84行 × 30列 标准网格**，支持在边界动态一键无限追加行列。

### 4. 📡 零依赖二进制流写盘与防御型日志收集
*   **物理流同步通道**: 伴侣服务器 `server.js` 基于 Node.js 原生的 `http` 与 `fs` 管道构建，**无 Express/Multer/CORS 等任何三方依赖**，开箱即用。
*   **Payload 限制防御**: 服务端内置数据流安全限载器。限制单次日志上报最大为 `1MB` 并自动销毁超标流（`req.destroy()`），强力防范内存溢出漏洞（OutOfMemory）与拒绝服务攻击（DoS）。

### 5. 🩺 浏览器运行时异常监控与暗黑诊断台 (Diagnostics Overlay)
*   前端深度集成组件崩溃与未捕获异常监控。当底层公式计算越界、网络中断或发生运行时 Crash 时，前端将自动上报并将日志同步安全持久化至服务端的 `browser-errors.log` 中。
*   右下角内置精美的 **Catppuccin 暗黑风格错误诊断面板**，提供高保真 StackTrace 显示、一键复制错误报告和防抖静默上报，方便开发者 and 用户进行故障回溯。

### 6. 📊 高保真 Excel 导入导出器 (High-Fidelity Exporter)
*   基于优秀的 `exceljs` 与 `luckyexcel`。
*   **深度翻译**: 自主实现公式翻译（带 `=` 字符兼容）、单元格背景色映射、精细字体字号比对、列宽/行高转换以及合并单元格（`merge`）信息高保真重构，确保多次修改、多次导入导出依然具有专业级排版一致性。

---

## 🏗️ 架构图解 (System Architecture)

```text
┌────────────────────────────────────────────────────────┐
│                      StreamSheet                       │
│              (React + FortuneSheet + TS)               │
└──────────────────────────┬─────────────────────────────┘
                           │
             ┌─────────────┴─────────────┐
             ▼                           ▼
┌─────────────────────────┐ ┌────────────────────────────┐
│ File System Access API  │ │  Companion Stream Server   │
│ (showOpenFilePicker)    │ │   (http + fs.writeStream)  │
├─────────────────────────┤ ├────────────────────────────┤
│   Direct Local Disk     │ │   Binary Byte-Stream POST  │
│   (Serverless, Chrome)  │ │   (Port 3001, All Browsers)│
└─────────────────────────┘ └────────────────────────────┘
```

---

## 🚀 快速启动 (Quick Start)

### 1. 安装项目依赖
确保您的 Node 版本 ≥ 16.x，然后在项目根目录下运行：
```bash
npm install
```

### 2. 启动系统

您可以选用以下 **最简单的一键启动方式**（脚本会自动检测并自动安装缺少的依赖）：

*   **🍏 macOS / Linux 用户**：
    在终端中直接运行以下命令：
    ```bash
    ./start.sh
    ```
*   **🪟 Windows 用户**：
    在项目根目录下双击运行 `start.bat` 文件。
*   **📦 跨平台通用命令**：
    在任何终端运行：
    ```bash
    npm start
    ```

---

#### 💡 (备选) 手动分步启动

如果您希望手动启动前端和后端，可以在两个不同的终端分别运行：

#### 终端 A: 启动 Companion 同步服务端 (3001 端口)
```bash
npm run server
```
*伴侣服务端会全自动在项目根目录下初始化一个精美的示例 `local.xlsx` 数据库。同时也会在这里生成 `browser-errors.log`。*

#### 终端 B: 启动 Web 网页客户端 (3000 端口)
```bash
npm run dev
```

打开浏览器，访问 `http://localhost:3000`：
*   **绿色指示灯 `⚡ Companion Stream Server` 亮起**，即表明已成功握手！
*   您在网页端做出的任何修改，都会在停止输入 1.5 秒后**默默地通过临时原子文件覆盖机制瞬间安全保存**至本地目录下的 `local.xlsx`。
*   您可以随时双击该物理 `.xlsx`，用 Microsoft Excel 或 WPS 进行联合协同查看。

---

## 🛠️ 项目结构 (Project Structure)

```text
├── package.json               # 项目配置与启动脚本 (更名为 stream-sheet)
├── server.js                  # 零依赖原生 Node.js 原子流同步服务器（安全落盘、错误收集、安全限载）
├── browser-errors.log         # 前端产生的浏览器运行时崩溃/异常历史日志
├── vite.config.ts             # Vite 构建配置文件 (监听 3000 端口)
├── index.html                 # 单页应用入口
├── src/
│   ├── main.tsx               # 前端主入口 (Strict Mode 深度兼容)
│   ├── App.tsx                # 多标签工作簿管理与双模同步主界面
│   ├── index.css              # 高颜值 UI 样式表（含 Catppuccin 暗黑终端样式、穿透 hover 桥）
│   ├── luckyexcel.d.ts        # TS 模块与 LuckyExcel 补丁声明
│   ├── vite-env.d.ts          # 全局 TypeScript 声明扩展（含文件系统 API Window 类型扩展）
│   └── utils/
│       ├── excelExporter.ts   # 自主研发的高保真 Excel 导出还原工具
│       └── sheetSanitizer.ts  # 高性能表格数据安全清洗与公式链裁剪层（消除只读属性冲突，消灭 any）
```

---

## 🔒 隐私与安全性声明 (Privacy & Security)

StreamSheet 始终将用户的数据安全和隐私放在首位：
1.  **绝不上传云端**: 所有数据均在本地浏览器的 JS 运行时或您的局域网/本机伴侣服务器处理，绝对不会上传到任何第三方云服务器。
2.  **句柄授权严格**: 浏览器 File System 权限随用随关，页面刷新后需要重新手动开启，防止静默恶意篡改。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 许可发布。
