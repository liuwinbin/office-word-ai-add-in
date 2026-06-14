# OfficeAI智能助手

基于大模型（DeepSeek API / 本地 Ollama）的 Word Web 加载项，提供自然语言驱动的智能排版与文本处理能力。

## 项目结构

```
office-word-ai-add-in/
├── manifest.xml              # Office Add-in 清单文件
├── taskpane.html             # 侧边栏 UI 界面
├── taskpane.js               # 核心业务逻辑
├── server.ts                 # 开发服务器（HTTPS + API 代理）
├── scripts/
│   └── generate-certs.mjs    # 自签名证书生成脚本
├── assets/
│   ├── icon-16.png           # 16x16 图标
│   ├── icon-32.png           # 32x32 图标
│   └── icon-80.png           # 80x80 图标
└── README.md                 # 本文件
```

## 环境要求

- Microsoft Word 2016 或更高版本（Windows / Mac / Word Online）
- Node.js v18+
- 模型服务二选一：
  - **本地 Ollama**：已安装并运行 Ollama，已拉取所需模型
  - **DeepSeek API**：有效的 API Key

## 快速开始

### 1. 安装依赖

```bash
npm install
```

### 2. 生成 HTTPS 证书（首次使用）

```bash
npm run certs
```

此命令会在 `.certs/` 目录下生成自签名证书 `localhost.pfx`，用于本地 HTTPS 服务。

### 3. 启动开发服务器

```bash
npm run dev
```

服务器会自动：
- 在 `https://localhost:3000` 启动 HTTPS 服务
- 提供 Word 加载项页面（`/taskpane.html`）
- 提供 React 应用页面（`/`）
- 代理 AI API 请求

> 浏览器会提示证书不受信任（自签名），点击"继续访问"即可。

### 4. 在 Word 中旁加载加载项

1. 打开 Microsoft Word
2. 选择 **插入** → **我的加载项**（或 **获取加载项**）
3. 在弹出窗口中，选择 **上传我的加载项**（Upload My Add-in）
4. 浏览选择本项目的 `manifest.xml` 文件
5. 加载完成后，Word **主页**选项卡中会出现 **AI助手** 分组，内含 **大模型** 按钮
6. 点击按钮即可打开侧边栏

## 使用指南

### 模型配置

1. 在侧边栏顶部选择模型服务商：**本地 Ollama** 或 **DeepSeek API**
2. 填写对应的配置信息：
   - Ollama：本地服务地址（默认 `http://localhost:11434`）和模型名称（默认 `deepseek-r1:latest`）
   - DeepSeek API：API Base URL（默认 `https://api.deepseek.com`）和 API Key
3. 点击 **测试连接** 验证配置是否正确

### 文本处理流程

1. 在 Word 文档中选中需要处理的文本
2. 点击侧边栏的 **提取选区文本** 按钮，将选中文本填入输入框
3. 在 **AI指令** 输入框中用自然语言描述你想要的排版或处理操作
4. 点击 **执行AI指令**，AI 处理后的结果将自动替换原选区文本

### 指令示例

| 指令 | 说明 |
|------|------|
| 将标题加粗并设为三号字 | 格式调整 |
| 把这段文字翻译成英文 | 翻译 |
| 将文本排版为正式公文格式 | 排版 |
| 将列表中每项用编号重新排列 | 结构化 |
| 修正文中的错别字和语法错误 | 校对 |

## 跨域说明

- 服务器已在代码中启用 CORS
- Ollama 默认允许本地跨域访问（`localhost:11434`）
- DeepSeek API 已配置 CORS，可直接从浏览器调用
- 如果 Ollama 遇到跨域问题，请设置环境变量 `OLLAMA_ORIGINS=*`

## 技术栈

- 原生 HTML / CSS / JavaScript（Word 加载项侧边栏）
- React + Vite + Express（主应用 + 开发服务器）
- Office JavaScript API (office.js)
- 大模型 API：Ollama `/api/chat` / DeepSeek `/v1/chat/completions`

## License

MIT
