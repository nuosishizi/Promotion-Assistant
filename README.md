# 媒体字幕处理工具 v2.2

> **Powered by** Groq Whisper · Google Gemini · Google Lens OCR · yt-dlp

---

## 📁 文件夹结构（发给别人时保持这个结构）

```
mp4_transcriber/
│
├── transcriber.py          ← 主程序（运行这个）
├── backend.py              ← 后端逻辑（勿删）
├── requirements.txt        ← Python 依赖
│
├── yt-dlp.exe              ← ★ 必须：视频下载工具
├── cookies.txt             ← 可选：浏览器 Cookie（下载私密视频用）
├── credentials.json        ← 仅 Google Sheets 功能需要
├── config.json             ← 自动生成，保存你的配置
│
└── chrome-lens-ocr-main/   ← ★ OCR 功能必须（见下方说明）
    ├── sharex.js
    ├── cli.js
    └── node_modules/
```

---

## ⚡ 快速安装（第一次使用）

### 1. 安装 Python 依赖
```bat
pip install groq google-generativeai gspread google-auth webvtt-py tkinterdnd2
```

### 2. 下载 yt-dlp.exe
前往 https://github.com/yt-dlp/yt-dlp/releases，下载 `yt-dlp.exe`，放到脚本文件夹。

### 3. 安装 FFmpeg
下载 https://github.com/BtbN/FFmpeg-Builds/releases，解压后确认路径：
```
C:\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe
```
（路径写死在 backend.py，可修改 `FFMPEG_EXE` 变量）

### 4. 配置 OCR（图片文字识别）
> 需要安装 Node.js（https://nodejs.org，建议 LTS 版）

**方式 A — 放到脚本文件夹（推荐、可发给别人）:**
```
mp4_transcriber/
└── chrome-lens-ocr-main/   ← 把整个文件夹复制进来
    ├── sharex.js
    └── ...
```

**方式 B — 安装到 C 盘根目录（已有的安装）:**
```
C:\chrome-lens-ocr-main\sharex.js   ← 程序会自动找到这里
```

**验证 OCR 是否正常（在命令行测试）:**
```bat
node C:\chrome-lens-ocr-main\sharex.js "C:\test_image.png"
```
有文字输出则正常。

---

## 🔑 API Keys 获取

| Key | 获取地址 |
|-----|---------|
| Groq API Key | https://console.groq.com → API Keys |
| Gemini API Key | https://aistudio.google.com/app/apikey |

填写后点击「**💾 保存**」按钮，自动存入 `config.json`。

---

## 📊 Google Sheets 配置

1. 前往 https://console.cloud.google.com/
2. 新建项目 → 启用 **Google Sheets API**
3. 创建「服务账户」→ 下载 JSON → 重命名为 **`credentials.json`** 放到脚本文件夹
4. 在 Google Sheet 里点「共享」，把服务账户邮箱加为编辑者

---

## 🖼 图片 OCR 功能说明

标签页「图片 OCR」支持三种模式：

| 模式 | 说明 |
|------|------|
| 📁 图片文件夹 | 自动扫描所有图片（PNG/JPG/WEBP 等），批量识别 |
| 🔗 图片链接 | 每行一个链接（可以是网络 URL 或本地路径） |
| 📊 Sheet 图片列 | 从 Google Sheet 指定列读取图片链接，识别后写回 |

结果会自动出现在下方的 **📊 结果表格** 中，双击可复制，「导出 CSV」可保存。

---

## ▶ 启动

```bat
python transcriber.py
```

---

## ⚠ Facebook 链接说明

检测到 `facebook.com / fb.com / fb.watch` 链接时，**自动跳过字幕下载，直接下载视频**。
勾选「识别画面文字」则额外截帧 OCR。

---

## 🔧 常见问题

| 问题 | 解决 |
|------|------|
| OCR 无输出 | 先命令行测试 `node sharex.js "图片路径"` 是否正常 |
| yt-dlp 下载失败 | 在 `config.json` 旁放 `cookies.txt`（用 EditThisCookie 导出） |
| Google Sheets 连不上 | 确认 credentials.json 格式，并把服务账户邮件加为 Sheet 编辑者 |
| FFmpeg 找不到 | 修改 `backend.py` 第 30 行 `FFMPEG_BIN` 路径 |
