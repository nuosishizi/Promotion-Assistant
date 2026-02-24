# -*- coding: utf-8 -*-
"""backend.py — 所有后端逻辑（下载 / OCR / 转录 / 翻译）"""

import os, sys, re, json, shutil, tempfile, threading, subprocess, time
from pathlib import Path

# ── 第三方库 ──────────────────────────────────────────────────────────
try:
    import groq as _groq_lib; GROQ_OK = True
except ImportError:
    GROQ_OK = False

try:
    from google import genai as _genai_sdk
    GENAI_OK = True
except ImportError:
    GENAI_OK = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

try:
    import webvtt; WEBVTT_OK = True
except ImportError:
    WEBVTT_OK = False

# ── 路径配置 ─────────────────────────────────────────────────────────
def _script_dir():
    """运行目录：PyInstaller 打包后返回 exe 所在目录，否则返回脚本目录。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

SCRIPT_DIR = _script_dir()

# ffmpeg: 优先使用同目录下的 ffmpeg\ 子文件夹（打包时），其次系统安装
_FFMPEG_CANDIDATES = [
    os.path.join(SCRIPT_DIR, "ffmpeg", "ffmpeg.exe"),       # 打包版
    r"C:\ffmpeg-master-latest-win64-gpl-shared\bin\ffmpeg.exe",  # 开发版
    "ffmpeg",                                                 # 系统 PATH
]
FFMPEG_EXE = next((p for p in _FFMPEG_CANDIDATES
                   if os.path.isfile(p) or p == "ffmpeg"), "ffmpeg")

ENV_FILE     = r"C:\Users\newnew\.openclaw.env"
YT_DLP_EXE   = os.path.join(SCRIPT_DIR, "yt-dlp.exe")
COOKIES_FILE = os.path.join(SCRIPT_DIR, "cookies.txt")
CREDS_FILE   = os.path.join(SCRIPT_DIR, "credentials.json")
CONFIG_FILE  = os.path.join(SCRIPT_DIR, "config.json")

# OCR: 优先用预编译的 sharex.exe（无需 Node.js），其次用 node + js 脚本
_SHAREX_EXE = os.path.join(SCRIPT_DIR, "sharex.exe")
if os.path.isfile(_SHAREX_EXE):
    # 打包发布模式：直接运行 sharex.exe
    NODE_EXE = ""
    LENS_JS  = _SHAREX_EXE   # 复用此变量，lens_ocr() 会检测是否为 .exe
else:
    NODE_EXE = "node"
    _LENS_CANDIDATES = [
        os.path.join(SCRIPT_DIR, "chrome-lens-ocr-main", "sharex.js"),
        r"C:\chrome-lens-ocr-main\sharex.js",
    ]
    LENS_JS = next((p for p in _LENS_CANDIDATES if os.path.isfile(p)),
                   _LENS_CANDIDATES[-1])

GEMINI_MODELS = ["gemini-2.0-flash", "gemini-2.5-flash", "gemini-1.5-pro-latest", "gemini-2.5-pro"]



_NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

# ══════════════════════════════════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════════════════════════════════

def load_env(path: str) -> dict:
    env = {}
    if not os.path.isfile(path):
        return env
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"): continue
            if "=" in line:
                k, _, v = line.partition("=")
                env[k.strip()] = v.strip().strip('"').strip("'")
    return env


def collect_media(paths: list, exts=(".mp4", ".mov", ".mkv", ".avi", ".webm")) -> list:
    result = []
    for p in paths:
        p = p.strip().strip("{}")
        if os.path.isfile(p) and p.lower().endswith(exts):
            result.append(p)
        elif os.path.isdir(p):
            for root, _, files in os.walk(p):
                for f in files:
                    if f.lower().endswith(exts):
                        result.append(os.path.join(root, f))
    return result


def collect_images(paths: list) -> list:
    IMG_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif", ".tiff")
    result = []
    for p in paths:
        p = p.strip().strip("{}")
        if os.path.isfile(p) and p.lower().endswith(IMG_EXTS):
            result.append(p)
        elif os.path.isdir(p):
            for r, _, files in os.walk(p):
                for f in sorted(files):
                    if f.lower().endswith(IMG_EXTS):
                        result.append(os.path.join(r, f))
    return result


def is_facebook_url(url: str) -> bool:
    return bool(re.search(r"(facebook\.com|fb\.com|fb\.watch)", url, re.I))


# ── FFmpeg ──────────────────────────────────────────────────────────
def ffmpeg_extract_wav(mp4: str, out_dir: str) -> str:
    wav = os.path.join(out_dir, Path(mp4).stem + ".wav")
    cmd = [FFMPEG_EXE, "-y", "-i", mp4,
           "-vn", "-ar", "16000", "-ac", "1", "-c:a", "pcm_s16le", wav]
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", creationflags=_NO_WINDOW)
    if r.returncode != 0:
        raise RuntimeError(f"FFmpeg 错误:\n{r.stderr[-400:]}")
    return wav


def ffmpeg_extract_frame(video: str, out_dir: str, ts: str = "00:00:05") -> str:
    """从视频截取一帧（默认第5秒），返回 PNG 路径"""
    png = os.path.join(out_dir, Path(video).stem + "_frame.png")
    cmd = [FFMPEG_EXE, "-y", "-ss", ts, "-i", video,
           "-frames:v", "1", "-q:v", "2", png]
    r = subprocess.run(cmd, capture_output=True, text=True,
                       encoding="utf-8", errors="replace", creationflags=_NO_WINDOW)
    if r.returncode != 0 or not os.path.isfile(png):
        raise RuntimeError(f"截帧失败: {r.stderr[-300:]}")
    return png


# ── OCR (Google Lens) ────────────────────────────────────────────────
def lens_ocr(image_path_or_url: str, script_path: str = "") -> str:
    """调用 sharex.exe 或 node sharex.js 识别图片，返回识别文字"""
    # 优先使用运行时传入的路径，如果为空，则回退寻找到预设的路径
    actual_lens_js = script_path if script_path else LENS_JS

    if not actual_lens_js or not os.path.isfile(actual_lens_js):
        raise RuntimeError("未找到 OCR 脚本或可执行程序，请前往上方设置有效路径。")

    if actual_lens_js.endswith(".exe"):
        # 打包模式：直接运行 sharex.exe（内置 Node.js）
        env = os.environ.copy()
        # 告诉 pkg 的 sharp 模块去哪里找 .node 文件
        sharp_dir = os.path.join(SCRIPT_DIR, "sharp-win32-x64")
        if os.path.isdir(sharp_dir):
            env["SHARP_MODULE_DIR"] = sharp_dir
        cmd = [actual_lens_js, image_path_or_url]
    else:
        # 开发模式：node sharex.js
        env = None
        cmd = [NODE_EXE, actual_lens_js, image_path_or_url]

    r = subprocess.run(cmd, capture_output=True, text=True, env=env,
                       encoding="utf-8", errors="replace",
                       timeout=60, creationflags=_NO_WINDOW)
    if r.returncode != 0:
        raise RuntimeError(f"OCR 失败: {r.stderr[:300]}")
    return r.stdout.strip()



# ── Groq Whisper ────────────────────────────────────────────────────
def groq_transcribe(client, wav: str) -> str:
    with open(wav, "rb") as f:
        resp = client.audio.transcriptions.create(
            model="whisper-large-v3-turbo", file=f, response_format="text")
    return resp if isinstance(resp, str) else resp.text


# ── Gemini 翻译 ──────────────────────────────────────────────────────
def gemini_translate(text: str, api_key: str, model_name: str, log_cb=None) -> str:
    if not text.strip(): return ""
    client = _genai_sdk.Client(api_key=api_key)
    prompt = (
        "将以下内容完整翻译为简体中文，保留原有段落格式，只输出译文，不加任何说明。\n\n"
        f"---\n{text}\n---"
    )
    last_err = None
    for attempt in range(2):
        try:
            resp = client.models.generate_content(model=model_name, contents=prompt)
            return resp.text.strip()
        except Exception as e:
            last_err = e
            if log_cb: log_cb(f"  ⚠ 翻译失败（第{attempt+1}次）: {e}")
            if attempt < 1: time.sleep(10)
    return f"[翻译失败: {str(last_err)[:120]}]"


# ── yt-dlp 封装 ──────────────────────────────────────────────────────
def run_yt_dlp(args: list, log_cb=None) -> subprocess.CompletedProcess:
    cmd = [YT_DLP_EXE] + args
    if os.path.isfile(COOKIES_FILE):
        cmd += ["--cookies", COOKIES_FILE]
        if log_cb: log_cb("  ✅ 使用 cookies.txt")
    return subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        creationflags=_NO_WINDOW, cwd=SCRIPT_DIR)


def download_subtitle(url: str, out_dir: str, log_cb=None) -> str | None:
    uid = f"sub_{int(time.time()*1000)}"
    tmpl = os.path.join(out_dir, uid + ".%(ext)s")
    r = run_yt_dlp([
        url, "--write-sub", "--write-auto-sub",
        "--sub-langs", "en.*,zh.*", "--convert-subs", "vtt",
        "--skip-download", "--no-warnings", "-o", tmpl, "-q",
    ], log_cb)
    for fn in os.listdir(out_dir):
        if fn.startswith(uid) and fn.endswith(".vtt"):
            return os.path.join(out_dir, fn)
    return None


def download_audio_wav(url: str, out_dir: str, log_cb=None) -> str | None:
    uid = f"audio_{int(time.time()*1000)}"
    tmpl = os.path.join(out_dir, uid + ".%(ext)s")
    run_yt_dlp([
        url, "-x", "--audio-format", "wav", "--audio-quality", "0",
        "--no-warnings", "-o", tmpl, "-q",
        "--postprocessor-args", "ffmpeg:-ar 16000 -ac 1",
    ], log_cb)
    for fn in os.listdir(out_dir):
        if fn.startswith(uid) and fn.endswith(".wav"):
            return os.path.join(out_dir, fn)
    return None


def download_video_file(url: str, out_dir: str, log_cb=None) -> str | None:
    uid = f"video_{int(time.time()*1000)}"
    tmpl = os.path.join(out_dir, uid + ".%(ext)s")
    r = run_yt_dlp([url, "--no-warnings", "-o", tmpl, "-q"], log_cb)
    if r.returncode != 0:
        if log_cb: log_cb(f"  ⚠ 视频下载失败: {r.stderr[:180]}")
        return None
    for fn in os.listdir(out_dir):
        if fn.startswith(uid):
            return os.path.join(out_dir, fn)
    return None


# ── VTT 解析 ────────────────────────────────────────────────────────
def parse_vtt(vtt_path: str) -> str:
    if WEBVTT_OK:
        try:
            seen, lines = set(), []
            for cap in webvtt.read(vtt_path):
                t = cap.text.strip().replace("\n", " ")
                if t and t not in seen:
                    seen.add(t); lines.append(t)
            return " ".join(lines)
        except Exception:
            pass
    with open(vtt_path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    return " ".join(re.findall(r"(?m)^(?!\d+:\d).+$", raw) or [])


# ══════════════════════════════════════════════════════════════════════
# 核心处理函数
# ══════════════════════════════════════════════════════════════════════

def process_url(url: str, out_dir: str, groq_client, gemini_key: str,
                gemini_model: str, download_video: bool, ocr_frame: bool,
                ocr_script_path: str,
                log_cb, pause_evt: threading.Event, stop_flag: list,
                audio_ok: bool = True, translate_ok: bool = True) -> dict:
    """处理单个 URL，返回 {source, original, chinese, status, note}"""
    res = dict(source=url, original="", chinese="", status="处理中", note="")
    tmp = tempfile.mkdtemp(prefix="mst_")
    is_fb = is_facebook_url(url)
    try:
        if is_fb:
            # ── Facebook: 直接下载视频，然后截帧 OCR ──────────────────
            log_cb("  📘 检测到 Facebook 链接，直接下载视频…")
            pause_evt.wait()
            if stop_flag[0]: return res
            vpath = download_video_file(url, out_dir, log_cb)
            if not vpath:
                raise RuntimeError("Facebook 视频下载失败")
            log_cb(f"  ✅ 视频已保存: {os.path.basename(vpath)}")
            # Ⅰ 语音转录
            if audio_ok and groq_client:
                try:
                    log_cb("  🎤 提取音频并 Whisper 转录…")
                    wav = ffmpeg_extract_wav(vpath, tmp)
                    if wav:
                        res["original"] = groq_transcribe(groq_client, wav)
                        log_cb(f"  语音: {res['original'][:80]}…")
                    else:
                        log_cb("  ⚠ 音频提取失败，跳过转录")
                except Exception as e:
                    log_cb(f"  ⚠ Whisper 转录失败: {e}")
            elif audio_ok and not groq_client:
                log_cb("  ⚠ 已开启语音识别但未配置 Groq API Key")
            # Ⅱ 画面 OCR
            if ocr_frame:
                log_cb("  🖼 截取视频封面帧并 OCR…")
                frame = ffmpeg_extract_frame(vpath, tmp)
                res["original"] = lens_ocr(frame, script_path=ocr_script_path)
                log_cb(f"  OCR: {res['original'][:80]}…")
            res["status"] = "成功（已下载视频）"
        if not is_fb:
            # ── 普通链接 ──────────────────────────────────────────────
            log_cb("  [1/3] 尝试下载字幕…")
            pause_evt.wait()
            if stop_flag[0]: return res
            vtt = download_subtitle(url, tmp, log_cb)
            if vtt:
                log_cb("  ✅ 字幕获取成功，解析中…")
                res["original"] = parse_vtt(vtt)
            elif audio_ok:
                log_cb("  ⚠ 无字幕，尝试 Whisper 转录…")
                if not groq_client:
                    raise RuntimeError("无字幕且未配置 Groq API Key")
                pause_evt.wait()
                if stop_flag[0]: return res
                wav = download_audio_wav(url, tmp, log_cb)
                if not wav:
                    raise RuntimeError("音频下载失败")
                log_cb("  🎤 Groq Whisper 转录…")
                res["original"] = groq_transcribe(groq_client, wav)
            else:
                log_cb("  ⚠ 无字幕，已关闭语音识别，跳过 Whisper")

            # 可选：截帧 OCR
            if ocr_frame:
                pause_evt.wait()
                if stop_flag[0]: return res
                log_cb("  [2/3] 截取视频帧 OCR…")
                try:
                    vid = download_video_file(url, tmp, log_cb)
                    if vid:
                        frame = ffmpeg_extract_frame(vid, tmp)
                        ocr_txt = lens_ocr(frame, script_path=ocr_script_path)
                        if ocr_txt:
                            res["original"] += f"\n\n[画面文字]\n{ocr_txt}"
                except Exception as e:
                    log_cb(f"  ⚠ 截帧 OCR 失败: {e}")


            # 可选：下载完整视频
            if download_video and not is_fb:
                pause_evt.wait()
                if stop_flag[0]: return res
                log_cb("  [3/3] 下载视频…")
                vpath2 = download_video_file(url, out_dir, log_cb)
                if vpath2: log_cb(f"  ✅ 视频: {os.path.basename(vpath2)}")

        # Gemini 翻译
        if translate_ok and res["original"] and GENAI_OK and gemini_key:
            pause_evt.wait()
            if stop_flag[0]: return res
            log_cb("  🌐 Gemini 翻译…")
            res["chinese"] = gemini_translate(res["original"], gemini_key, gemini_model, log_cb)

        if not res["status"].startswith("成功"):
            res["status"] = "成功"
    except Exception as e:
        res["status"] = "失败"; res["note"] = str(e)
        log_cb(f"  ❌ {e}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return res


def process_image_ocr(source: str, gemini_key: str, gemini_model: str,
                       do_translate: bool, ocr_script_path: str, log_cb) -> dict:
    """对单张图片（本地路径或 URL）进行 OCR，可选 Gemini 翻译"""
    res = dict(source=source, original="", chinese="", status="处理中", note="")
    try:
        log_cb(f"  🖼 OCR: {source[:70]}…")
        res["original"] = lens_ocr(source, script_path=ocr_script_path)
        log_cb(f"  识别: {res['original'][:80]}…")
        if do_translate and GENAI_OK and gemini_key and res["original"]:
            log_cb("  🌐 Gemini 翻译…")
            res["chinese"] = gemini_translate(res["original"], gemini_key, gemini_model, log_cb)
        res["status"] = "成功"
    except Exception as e:
        res["status"] = "失败"; res["note"] = str(e)
        log_cb(f"  ❌ {e}")
    return res
