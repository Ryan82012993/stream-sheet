# StreamSheet ⚡ 

> **Local-First, Real-Time Web Excel Editor with Dual-Mode Stream Sync.**
>
> 这是一个基于 **React + TypeScript + FortuneSheet** 构建的**本地优先（Local-First）** 实时 Excel 在线编辑器。支持现代浏览器物理句柄直接读写与轻量级 Node.js 伴侣流同步（Companion Stream Server）的双模全自动保存机制，兼具极高的本地安全隐私性与开箱即用的同步体验。

---

## ✨ 核心特色 (Key Features)

### 1. 💻 本地优先双模同步 (Dual-Mode Sync Architecture)
*   **现代物理直读写 (Modern Local-First)**: 默认采用现代浏览器的 **File System Access API** (`showOpenFilePicker` 和 `createWritable`)。一次授权，即可在 Web 端静默物理写盘。
*   **伴侣流同步服务 (Companion Stream Server)**: 针对其他浏览器，StreamSheet 自动侦测并建立与本地轻量级 Node 服务的同步。支持 **防抖自动保存 (Debounced Auto-Save)**，在最后一次修改后延时 1.5 秒即启动流式物理回写。

### 2. 📡 零依赖二进制流写盘 (Zero-Dependency Binary Stream Saving)
*   **物理流同步通道**: 伴侣服务器 `server.js` 基于 Node.js 原生的 `http` 与 `fs` 管道构建，**无 Express/Multer/CORS 等任何三方依赖**。
*   **高性能**: 直接通过 `application/octet-stream` 格式的 Raw Buffer 字节流进行传输和写入，绕过 Multipart 表单解析的巨大性能开销，实现瞬间写盘。

### 3. 📊 高保真 Excel 导入导出器 (High-Fidelity Exporter)
*   基于优秀的 `exceljs` 与 `luckyexcel`。
*   **深度翻译**: 自主实现公式翻译（带 `=` 字符兼容）、单元格背景色映射、精细字体字号比对、列宽/行高转换以及合并单元格（`merge`）信息高保真重构，确保多次修改、多次导入导出依然具有专业级排版一致性。

### 4. 🧩 极佳的用户体验
*   支持拖拽、本地选档、一键导出另存为。
*   顶部拥有状态指示灯，完美显示：`💻 物理直写` | `⚡ 伴侣流同步` | `☁️ 离线临时编辑` 以及保存状态。

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
确保您的 Node 版本 $\ge$ 16.x，然后在项目根目录下运行：
```bash
npm install
```

### 2. 启动系统

为了享受完整的伴侣流同步体验，建议在两个终端并发启动前端和后端：

#### 终端 A: 启动 Companion 同步服务端 (3001 端口)
```bash
npm run server
```
*伴侣服务端会全自动在项目根目录下初始化一个精美的示例 `local.xlsx` 数据库。*

#### 终端 B: 启动 Web 网页客户端 (3000 端口)
```bash
npm run dev
```

打开浏览器，访问 `http://localhost:3000`：
*   **绿色指示灯 `⚡ Companion Stream Server` 亮起**，即表明已成功握手！
*   您在网页端做出的任何修改，都会在停止输入 1.5 秒后**默默地瞬间保存**至本地目录下的 `local.xlsx`。
*   您可以随时双击该物理 `.xlsx`，用 Microsoft Excel 或 WPS 进行联合协同查看。

---

## 🛠️ 项目结构 (Project Structure)

```text
├── package.json               # 项目配置与启动脚本 (更名为 stream-sheet)
├── server.js                  # 零依赖原生 Node.js 伴侣流同步服务器
├── vite.config.ts             # Vite 构建配置文件 (监听 3000 端口)
├── index.html                 # 单页应用入口
├── src/
│   ├── main.tsx               # 前端主入口
│   ├── App.tsx                # 双模同步主界面逻辑
│   ├── index.css              # 高颜值 UI 样式表
│   ├── luckyexcel.d.ts        # TS 模块与 LuckyExcel 补丁声明
│   └── utils/
│       └── excelExporter.ts   # 自主研发的高保真 Excel 导出还原工具
```

---

## 🔒 隐私与安全性声明 (Privacy & Security)

StreamSheet 始终将用户的数据安全和隐私放在首位：
1.  **绝不上传云端**: 所有数据均在本地浏览器的 JS 运行时或您的局域网/本机伴侣服务器处理，绝对不会上传到任何第三方云服务器。
2.  **句柄授权严格**: 浏览器 File System 权限随用随关，页面刷新后需要重新手动开启，防止静默恶意篡改。

---

## 📄 开源许可证 (License)

本项目采用 [MIT License](LICENSE) 许可发布。
