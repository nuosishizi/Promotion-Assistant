@echo off
chcp 65001 >nul
setlocal

echo ============================================================
echo   MediaTool 打包脚本
echo ============================================================
echo.

:: ── 0. 检查工具 ─────────────────────────────────────────────
where node >nul 2>&1 || (echo [错误] 未找到 Node.js，请先安装 & pause & exit /b 1)
where python >nul 2>&1 || (echo [错误] 未找到 Python & pause & exit /b 1)

:: ── 1. 安装 Python 依赖 ──────────────────────────────────────
echo [1/5] 安装 Python 依赖...
pip install -q pyinstaller groq google-genai gspread google-auth webvtt-py customtkinter tkinterdnd2 pillow
if errorlevel 1 (echo [错误] pip install 失败 & pause & exit /b 1)

:: ── 2. 确认 yt-dlp.exe 存在 ─────────────────────────────────
echo.
echo [2/4] 检查 yt-dlp.exe...
if not exist "%~dp0yt-dlp.exe" (
    echo [下载] 正在下载 yt-dlp.exe...
    curl -L "https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe" -o "%~dp0yt-dlp.exe"
    if errorlevel 1 echo [警告] yt-dlp 下载失败，请手动复制 yt-dlp.exe 到项目目录
) else (
    echo [OK] yt-dlp.exe 已存在
)

:: ── 3. 运行 PyInstaller ──────────────────────────────────────
echo.
echo [3/4] 运行 PyInstaller 打包...
pyinstaller MediaTool.spec --noconfirm --clean
if errorlevel 1 (echo [错误] PyInstaller 失败 & pause & exit /b 1)
echo [OK] PyInstaller 完成

:: ── 4. 把 ffmpeg 复制到 dist\MediaTool\ ─────────────────────
echo.
echo [4/4] 复制 ffmpeg...
set "FFMPEG_SRC=C:\ffmpeg-master-latest-win64-gpl-shared\bin"
set "DIST=dist\MediaTool"

if exist "%FFMPEG_SRC%\ffmpeg.exe" (
    mkdir "%DIST%\ffmpeg" 2>nul
    copy /y "%FFMPEG_SRC%\ffmpeg.exe"  "%DIST%\ffmpeg\" >nul
    copy /y "%FFMPEG_SRC%\ffprobe.exe" "%DIST%\ffmpeg\" >nul
    echo [OK] ffmpeg 复制到 %DIST%\ffmpeg\
) else (
    echo [警告] 未找到 ffmpeg，请手动复制
)

:: ── 完成 ─────────────────────────────────────────────────────
echo.
echo ============================================================
echo   打包完成！输出目录: dist\MediaTool\
echo ============================================================
echo.
echo 分发给用户时，把整个 dist\MediaTool\ 文件夹打 zip 即可。
echo 用户解压后直接双击 MediaTool.exe 运行，无需安装任何软件。
echo.
pause
