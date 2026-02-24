# Promotion Assistant — 媒体字幕处理工具 v2.2

> **Powered by** Groq Whisper · Google Gemini · Google Lens OCR · yt-dlp

自动转录 MP4 / 视频链接、翻译字幕、批量图片 OCR、写入 Google Sheets，一体化推广助手。

---

## 📁 推荐文件夹结构

```
mp4_transcriber/
│
├── transcriber.py          ← 主程序（运行这个）
├── backend.py              ← 后端逻辑（勿删）
├── requirements.txt        ← Python 依赖列表
├── build.bat               ← 打包脚本（可选）
│
├── yt-dlp.exe              ← ★ 必须手动下载，放到此处
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
![demo](<2026-02-24 142814204.png>) 
## ⚡ 安装步骤（第一次使用）

### 第 1 步：安装 Python

前往 https://www.python.org/downloads/ 下载并安装 Python **3.10+**。

安装时勾选 ☑ **"Add Python to PATH"**（非常重要）。

验证安装：
```bat
python --version
```

---

### 第 2 步：安装 Python 依赖

在项目文件夹下打开命令行，运行：

```bat
pip install -r requirements.txt
```

或者手动安装：

```bat
pip install groq google-genai gspread google-auth webvtt-py tkinterdnd2
```

| 包名 | 用途 |
|------|------|
| `groq` | 调用 Groq Whisper 转录语音 |
| `google-genai` | 调用 Gemini API 翻译/总结 |
| `gspread` | 读写 Google Sheets |
| `google-auth` | Google 服务账号认证 |
| `webvtt-py` | 解析 VTT 字幕文件 |
| `tkinterdnd2` | 拖拽文件到窗口 |

---

### 第 3 步：下载 yt-dlp.exe

1. 前往 https://github.com/yt-dlp/yt-dlp/releases
2. 找到最新版本，下载 **`yt-dlp.exe`**
3. 放到脚本文件夹（和 `transcriber.py` 同级）

---

### 第 4 步：安装 FFmpeg

1. 前往 https://github.com/BtbN/FFmpeg-Builds/releases
2. 下载 **`ffmpeg-master-latest-win64-gpl-shared.zip`**
3. 解压，确保路径为：
   ```
   C:\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe
   ```

> 如果你解压到其他位置，请在 `backend.py` 中修改 `FFMPEG_BIN` 变量。

---

### 第 5 步：配置 OCR（图片文字识别）

OCR 功能依赖 **Node.js** 和 **chrome-lens-ocr** 库。

#### 5-A：安装 Node.js

前往 https://nodejs.org/，下载 **LTS 版本**安装。

验证安装：
```bat
node --version
```

#### 5-B：下载 chrome-lens-ocr

```bat
# 下载源码（方式一：git clone）
git clone https://github.com/dimdenGD/chrome-lens-ocr.git chrome-lens-ocr-main
cd chrome-lens-ocr-main
npm install

# 下载源码（方式二：直接下载 ZIP）
# 前往 https://github.com/dimdenGD/chrome-lens-ocr，点 Code → Download ZIP，解压
```

#### 5-C：放到正确位置

**推荐（可发给别人）：** 把整个 `chrome-lens-ocr-main/` 文件夹放到脚本同级目录：
```
mp4_transcriber/
└── chrome-lens-ocr-main/
    ├── sharex.js
    └── node_modules/
```

**备选：** 放到 `C:\chrome-lens-ocr-main\`，程序会自动找到。

#### 5-D：验证 OCR 是否正常

```bat
node chrome-lens-ocr-main\sharex.js "C:\test_image.png"
```

有文字输出则正常。

---

## 🔑 API Keys 获取

在程序界面的「**⚙ 设置**」里填写以下 Key，填完点「**💾 保存**」：

| Key | 获取地址 |
|-----|---------|
| Groq API Key | https://console.groq.com → API Keys |
| Gemini API Key | https://aistudio.google.com/app/apikey |

---

## 📊 Google Sheets 配置（可选）

1. 前往 https://console.cloud.google.com/
2. 新建项目 → 启用 **Google Sheets API** 和 **Google Drive API**
3. 创建「服务账户」→ 下载 JSON → 重命名为 **`credentials.json`** 放到脚本文件夹
4. 在 Google Sheet 里点「共享」，把服务账户邮箱加为**编辑者**

---

## ▶ 启动

```bat
python transcriber.py
```

双击 `transcriber.py` 也可以（前提是 Python 已关联 `.py` 文件）。

---

## 🖼 图片 OCR 功能

标签页「图片 OCR」支持三种模式：

| 模式 | 说明 |
|------|------|
| 📁 图片文件夹 | 自动扫描所有图片（PNG/JPG/WEBP 等），批量识别 |
| 🔗 图片链接 | 每行一个链接（网络 URL 或本地路径），批量识别 |
| 📊 Sheet 图片列 | 从 Google Sheet 指定列读取图片链接，识别后写回 |

---

## ⚠ Facebook 链接说明

检测到 `facebook.com / fb.com / fb.watch` 链接时，**自动跳过字幕下载，直接下载视频**。  
勾选「识别画面文字」则额外截帧 OCR。

---

## 🔧 常见问题

| 问题 | 解决 |
|------|------|
| `ModuleNotFoundError` | 运行 `pip install -r requirements.txt` |
| OCR 无输出 | 命令行测 `node sharex.js "图片路径"` 是否正常；确认 Node.js 已安装 |
| yt-dlp 下载失败 | 在脚本文件夹放 `cookies.txt`（用 EditThisCookie 扩展导出） |
| Google Sheets 连不上 | 确认 `credentials.json` 格式，服务账户邮件已加为 Sheet 编辑者 |
| FFmpeg 找不到 | 修改 `backend.py` 中的 `FFMPEG_BIN` 变量 |
| `tkinterdnd2` 报错 | `pip install tkinterdnd2` 或到 https://pypi.org/project/tkinterdnd2/ 手动下载 |
