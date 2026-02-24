# -*- coding: utf-8 -*-
"""媒体字幕处理工具 v2.2  |  启动: python transcriber.py"""

import os, csv, json, re, threading, shutil, tempfile, subprocess
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor
import customtkinter as ctk
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

from backend import (
    SCRIPT_DIR, ENV_FILE, FFMPEG_EXE, YT_DLP_EXE, CREDS_FILE,
    CONFIG_FILE, LENS_JS, GEMINI_MODELS, _NO_WINDOW,
    GROQ_OK, GENAI_OK, GSPREAD_OK, WEBVTT_OK,
    load_env, collect_media, collect_images,
    ffmpeg_extract_wav, groq_transcribe, gemini_translate,
    process_url, process_image_ocr, _groq_lib,
)
try:
    if GSPREAD_OK:
        import gspread
        from google.oauth2.service_account import Credentials
except Exception:
    pass

# DnD: CTk 和 tkinterDnD 兼容性有限，暂时禁用拖放
DND_OK = False
_Base = ctk.CTk

# ══════════════════════════════════════════════════════════════════════
# 🎨  VS Code / Notion Dark 配色
# ══════════════════════════════════════════════════════════════════════
C = dict(
    bg      = "#f8f9fa",   # 主背景极浅灰
    card    = "#ffffff",   # 卡片白
    surface = "#f1f3f5",   # 输入框 / 悬浮
    accent  = "#2563eb",   # 清新蓝 accent
    accent2 = "#3b82f6",   # hover 蓝
    red     = "#ef4444",   # 错误红
    green   = "#22c55e",   # 成功绿
    orange  = "#d97706",   # 警告橙
    text    = "#1e293b",   # 主文字深灰
    sub     = "#64748b",   # 副文字中灰
    dim     = "#94a3b8",   # 更淡
    border  = "#e2e8f0",   # 细边框
    tab_sel = "#ffffff",
    log_bg  = "#f8fafc",   # 日志区极淡
    log_fg  = "#475569",
)

# ══════════════════════════════════════════════════════════════════════
# 结果表格
# ══════════════════════════════════════════════════════════════════════
class ResultTable(tk.Frame):
    COLS   = ("来源", "原文识别", "中文翻译", "状态")
    WIDTHS = (200, 320, 320, 80)

    def __init__(self, master, **kw):
        kw.setdefault("bg", C["bg"])
        super().__init__(master, **kw)
        self._rows: list[tuple] = []
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["bg"])
        hdr.pack(fill="x", pady=(6, 2))
        tk.Label(hdr, text="📊  结果表格",
                 font=("Segoe UI", 10, "bold"),
                 bg=C["bg"], fg=C["accent"]).pack(side="left", padx=4)

        def _ibtn(text, cmd, bg):
            return tk.Button(hdr, text=text, command=cmd,
                             bg=bg, fg=C["text"], activebackground=C["accent2"],
                             activeforeground=C["text"],
                             font=("Segoe UI", 9, "bold"), relief="flat",
                             cursor="hand2", padx=12, pady=4, bd=0)

        _ibtn("📋 复制全部", self._copy_all, C["accent"] ).pack(side="right", padx=3)
        _ibtn("💾 导出 CSV",  self._export_csv, C["surface"]).pack(side="right", padx=3)
        _ibtn("🗑 清空",      self.clear, C["dim"]   ).pack(side="right", padx=3)

        wrap = tk.Frame(self, bg=C["border"], bd=1)
        wrap.pack(fill="both", expand=True)
        vsb = ttk.Scrollbar(wrap, orient="vertical")
        hsb = ttk.Scrollbar(wrap, orient="horizontal")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree = ttk.Treeview(
            wrap, columns=self.COLS, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set, height=7)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        for col, w in zip(self.COLS, self.WIDTHS):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, minwidth=50)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self._on_dbl)
        self.tree.tag_configure("ok",  background="#d1fae5", foreground="#065f46")
        self.tree.tag_configure("err", background="#fee2e2", foreground="#991b1b")

        self._tip = tk.Label(self, text="💡 双击行 → 复制（Tab 分隔，可直接粘贴到 Google Sheet）",
                              font=("Segoe UI", 8), bg=C["bg"], fg=C["dim"])
        self._tip.pack(anchor="w", padx=4, pady=(2,0))

    def add_row(self, source, original, chinese, status):
        full   = (source, original, chinese, status)
        self._rows.append(full)
        tag    = "ok" if "成功" in status else "err"
        disp   = (source[:55],
                  original[:72].replace("\n"," "),
                  chinese[:72].replace("\n"," "),
                  status)
        self.tree.insert("", "end", values=disp, tags=(tag,))
        self.tree.yview_moveto(1.0)

    def clear(self):
        self._rows.clear()
        for i in self.tree.get_children(): self.tree.delete(i)

    def _tsv(self, row): return "\t".join(str(c).replace("\n"," ") for c in row)

    def _on_dbl(self, _):
        sel = self.tree.selection()
        if not sel: return
        idx = self.tree.index(sel[0])
        if idx < len(self._rows):
            self.clipboard_clear()
            self.clipboard_append(self._tsv(self._rows[idx]))
            self._flash("✅ 已复制该行 — 直接 Ctrl+V 粘贴到 Google Sheet")

    def _copy_all(self):
        if not self._rows: return
        lines = ["\t".join(self.COLS)] + [self._tsv(r) for r in self._rows]
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))
        self._flash(f"✅ 已复制 {len(self._rows)} 行（含表头）")

    def _flash(self, msg):
        self._tip.config(text=msg, fg=C["accent"])
        self.after(3500, lambda: self._tip.config(
            text="💡 双击行 → 复制（Tab 分隔，可直接粘贴到 Google Sheet）", fg=C["dim"]))

    def _export_csv(self):
        if not self._rows: messagebox.showinfo("提示", "结果为空"); return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("所有", "*.*")],
            initialfile=f"result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        if not path: return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            csv.writer(f).writerows([self.COLS] + list(self._rows))
        messagebox.showinfo("完成", f"已保存：\n{path}")


# ══════════════════════════════════════════════════════════════════════
# 主窗口
# ══════════════════════════════════════════════════════════════════════
class App(_Base):
    def __init__(self):
        super().__init__()
        self.title("🎬 媒体字幕处理工具 v2.2")
        self.geometry("1040x860")
        self.configure(fg_color=C["bg"])
        self.resizable(True, True)

        self._env       = load_env(ENV_FILE)
        self._running   = False
        self._pause_evt = threading.Event(); self._pause_evt.set()
        self._stop_flag = [False]
        self._mp4_queue = []
        self._all_pause_btns = []
        self._all_stop_btns  = []

        self._setup_styles()
        self._build_header()
        self._build_keys_bar()
        self._build_notebook()
        self._build_progress()
        self._build_table()
        self._build_log()
        self._load_config()
        self._check_deps()
        self.protocol("WM_DELETE_WINDOW", self._on_close)   # 关闭时自动保存

    # ── Styles ──────────────────────────────────────────────────────
    def _setup_styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        # Treeview (no CTk equivalent, keep ttk styled for light mode)
        st.configure("Treeview",
                     background=C["card"], fieldbackground=C["card"],
                     foreground=C["text"], rowheight=28,
                     font=("Segoe UI", 9), borderwidth=0,
                     relief="flat")
        st.configure("Treeview.Heading",
                     background=C["surface"], foreground=C["accent"],
                     font=("Segoe UI", 9, "bold"), relief="flat")
        st.map("Treeview",
               background=[("selected", C["accent"])],
               foreground=[("selected", "#ffffff")])
        # Scrollbar
        st.configure("TScrollbar", background=C["surface"], troughcolor=C["bg"],
                     relief="flat", borderwidth=0)
        # Spinbox
        st.configure("TSpinbox",
                     fieldbackground=C["surface"], background=C["surface"],
                     foreground=C["text"], insertwidth=1)
        # Progressbar (for a2v tab which still uses ttk)
        st.configure("TSep.Horizontal.TProgressbar",
                     troughcolor=C["surface"], background=C["accent"],
                     thickness=6, borderwidth=0)

    # ── Header strip ────────────────────────────────────────────────
    def _build_header(self):
        h = tk.Frame(self, bg=C["card"], height=50)
        h.pack(fill="x"); h.pack_propagate(False)
        tk.Label(h, text="🎬  媒体字幕处理工具  v2.2",
                 font=("Segoe UI", 14, "bold"),
                 bg=C["card"], fg=C["text"]).pack(side="left", padx=20, pady=10)
        tk.Label(h, text="Powered by Groq Whisper + Gemini + Google Lens OCR",
                 font=("Segoe UI", 8), bg=C["card"], fg=C["dim"]
                 ).pack(side="left", padx=0, pady=10)

    # ── API Keys bar ────────────────────────────────────────────────
    def _build_keys_bar(self):
        bar = ctk.CTkFrame(self, fg_color=C["surface"], corner_radius=0,
                           border_color=C["border"], border_width=1)
        bar.pack(fill="x")
        inner = ctk.CTkFrame(bar, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=8)

        self.groq_key_var   = tk.StringVar(value=self._env.get("GROQ_API_KEY",""))
        self.gemini_key_var = tk.StringVar(value=self._env.get("GEMINI_API_KEY",""))
        self.gemini_model_var = tk.StringVar(value=GEMINI_MODELS[0])

        def _lbl(text, col):
            ctk.CTkLabel(inner, text=text, font=("Segoe UI",9,"bold"),
                         text_color=C["sub"]).grid(row=0, column=col, sticky="w", padx=(0,4))

        def _key(var, show="*"):
            return ctk.CTkEntry(inner, textvariable=var, show=show, width=200,
                                font=("Consolas",10), fg_color=C["card"],
                                text_color=C["text"], border_color=C["border"],
                                corner_radius=6)

        _lbl("Groq Key:", 0)
        _key(self.groq_key_var).grid(row=0, column=1, padx=(0,14), pady=(0,4))
        _lbl("Gemini Key:", 2)
        _key(self.gemini_key_var).grid(row=0, column=3, padx=(0,14), pady=(0,4))
        ctk.CTkLabel(inner, text="模型:", font=("Segoe UI",9,"bold"),
                     text_color=C["sub"]).grid(row=0, column=4, sticky="w", padx=(0,4), pady=(0,4))
        ctk.CTkComboBox(inner, values=GEMINI_MODELS, variable=self.gemini_model_var,
                        width=200, font=("Segoe UI",10),
                        fg_color=C["card"], text_color=C["text"],
                        border_color=C["border"], button_color=C["accent"],
                        button_hover_color=C["accent2"],
                        dropdown_fg_color=C["card"],
                        dropdown_text_color=C["text"]
                        ).grid(row=0, column=5, padx=(0,12), pady=(0,4))
        
        # Row 1: OCR Script Path
        self.ocr_script_path_var = tk.StringVar(value=self._env.get("OCR_SCRIPT_PATH", ""))
        _lbl("OCR 脚本:", 0)
        _key(self.ocr_script_path_var, show="").grid(row=1, column=1, columnspan=3, sticky="ew", padx=(0,14), pady=(0,0))
        
        def _pick_script():
            p = filedialog.askopenfilename(title="选择 OCR 脚本", filetypes=[("JS/EXE 脚本", "*.js;*.exe"), ("所有文件", "*.*")])
            if p: self.ocr_script_path_var.set(p)
            
        ctk.CTkButton(inner, text="📁", width=28, height=28,
                      command=_pick_script, fg_color=C["surface"], 
                      text_color=C["text"], hover_color=C["border"]).grid(row=1, column=4, sticky="w", padx=(0,4))

        ctk.CTkButton(inner, text="💾 保存", command=self._save_config,
                      fg_color=C["accent"], hover_color=C["accent2"],
                      text_color="#ffffff", font=("Segoe UI",10,"bold"),
                      corner_radius=8, width=80, height=32
                      ).grid(row=0, rowspan=2, column=6)

    # ── CTkTabview Notebook ─────────────────────────────────────────
    def _build_notebook(self):
        self.nb = ctk.CTkTabview(
            self,
            fg_color=C["bg"],
            segmented_button_fg_color=C["surface"],
            segmented_button_selected_color=C["accent"],
            segmented_button_selected_hover_color=C["accent2"],
            segmented_button_unselected_color=C["surface"],
            segmented_button_unselected_hover_color=C["border"],
            text_color=C["text"],
            text_color_disabled=C["sub"],
            corner_radius=8)
        self.nb.pack(padx=10, pady=(8,0), fill="x")
        for t in ("🎬 MP4转录", "🔗 链接处理",
                  "📊 Google Sheets", "🖼 图片OCR",
                  "🗂 文件命名", "🎵 音频转视频"):
            self.nb.add(t)
        self._build_tab_mp4()
        self._build_tab_url()
        self._build_tab_sheets()
        self._build_tab_ocr()
        self._build_tab_rename()
        self._build_tab_audio2video()

    # ── helpers ─────────────────────────────────────────────────────
    def _frame(self, parent, **kw):
        kw.setdefault("fg_color", C["bg"])
        return ctk.CTkFrame(parent, **kw)

    def _card_frame(self, parent):
        return ctk.CTkFrame(parent, fg_color=C["card"],
                            border_color=C["border"], border_width=1,
                            corner_radius=8)

    def _label(self, p, text, bold=False, fg=None):
        return ctk.CTkLabel(p, text=text,
                            font=("Segoe UI", 10, "bold" if bold else "normal"),
                            text_color=fg or C["sub"])

    def _entry(self, p, var, w=50, show=""):
        return ctk.CTkEntry(p, textvariable=var, show=show, width=max(60, w*7),
                            font=("Consolas", 10), fg_color=C["card"],
                            text_color=C["text"], border_color=C["border"],
                            corner_radius=6)

    _LIGHT_BKGS = None  # resolved lazily
    def _btn(self, p, text, cmd, bg=None, fg=None):
        _bg = bg or C["accent"]
        # Light backgrounds need dark text; dark backgrounds need white text
        _light_bgs = {C["surface"], C["card"], C["dim"], C["border"], C["bg"]}
        if fg is None:
            _fg = C["text"] if _bg in _light_bgs else "#ffffff"
        else:
            _fg = fg
        # Hover: darken light buttons slightly, brighten accent buttons
        if _bg == C["accent"]:
            _hv = C["accent2"]
        elif _bg in _light_bgs:
            _hv = C["border"]  # slightly darker shade on hover
        else:
            _hv = _bg
        return ctk.CTkButton(p, text=text, command=cmd,
                             fg_color=_bg, hover_color=_hv,
                             text_color=_fg,
                             font=("Segoe UI", 10, "bold"),
                             corner_radius=8, height=32)

    def _check(self, p, text, var, command=None):
        kw = dict(command=command) if command else {}
        return ctk.CTkCheckBox(p, text=text, variable=var,
                               font=("Segoe UI", 10),
                               text_color=C["text"],
                               fg_color=C["accent"],
                               hover_color=C["accent2"],
                               checkmark_color="#ffffff",
                               corner_radius=4, **kw)

    def _radio(self, p, text, var, val, command=None):
        kw = dict(command=command) if command else {}
        return ctk.CTkRadioButton(p, text=text, variable=var, value=val,
                                  font=("Segoe UI", 10),
                                  text_color=C["text"],
                                  fg_color=C["accent"],
                                  hover_color=C["accent2"], **kw)

    def _ctrl_bar(self, parent, start_cmd):
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x", padx=10, pady=(4,8))
        self._btn(bar, "▶  开始", start_cmd, C["accent"]).pack(side="right", padx=4)
        sb = self._btn(bar, "⏹ 停止", self._stop, C["red"])
        sb.pack(side="right", padx=4)
        pb = self._btn(bar, "⏸ 暂停", self._toggle_pause, C["surface"])
        pb.pack(side="right", padx=4)
        sb.configure(state="disabled"); pb.configure(state="disabled")
        self._all_pause_btns.append(pb); self._all_stop_btns.append(sb)
        return pb, sb

    def _drop_zone(self, parent, default_text, on_drop=None, on_click=None):
        f = ctk.CTkFrame(parent, fg_color=C["surface"], height=72, cursor="hand2",
                         border_color=C["accent"], border_width=2, corner_radius=8)
        f.pack(padx=10, pady=(10,4), fill="x")
        f.pack_propagate(False)
        lbl = ctk.CTkLabel(f, text=default_text, font=("Segoe UI", 11),
                           text_color=C["accent"], cursor="hand2")
        lbl.pack(expand=True)
        if on_click:
            for w in (f, lbl._label):
                try: w.bind("<Button-1>", lambda _: on_click())
                except Exception: pass
        return f, lbl

    # ── Tab 1: MP4 转录 ─────────────────────────────────────────────
    def _build_tab_mp4(self):
        outer = self.nb.tab("🎬 MP4转录")

        _, self.mp4_lbl = self._drop_zone(
            outer,
            "⬇  点击选择视频文件",
            on_click=self._mp4_pick)

        opt = ctk.CTkFrame(outer, fg_color="transparent")
        opt.pack(padx=10, pady=(4,0), fill="x")
        self._btn(opt, "📂 选择文件", self._mp4_pick,  C["surface"]).pack(side="left", padx=4)
        self._btn(opt, "🗑 清空",     self._mp4_clear, C["surface"]).pack(side="left", padx=4)
        self.mp4_audio_var     = tk.BooleanVar(value=True)
        self.mp4_ocr_var       = tk.BooleanVar(value=False)
        self.mp4_translate_var = tk.BooleanVar(value=True)
        self._check(opt, "🎤 识别语音（Groq Whisper）",
                    self.mp4_audio_var).pack(side="left", padx=10)
        self._check(opt, "🖼 识别画面文字（OCR）",
                    self.mp4_ocr_var).pack(side="left", padx=4)
        self._check(opt, "🌐 翻译为中文",
                    self.mp4_translate_var).pack(side="left", padx=10)
        self._ctrl_bar(outer, self._mp4_start)

    # ── Tab 2: 链接处理 ─────────────────────────────────────────────
    def _build_tab_url(self):
        outer = self.nb.tab("🔗 链接处理")

        self._label(outer, "每行一个链接（YouTube / TikTok / Facebook / Instagram …）:"
                    ).pack(padx=10, pady=(8,2), anchor="w")
        self.url_text = ctk.CTkTextbox(
            outer, font=("Consolas", 10), fg_color=C["card"],
            text_color=C["text"], border_color=C["border"],
            border_width=1, corner_radius=6, height=80)
        self.url_text.pack(padx=10, fill="x")

        drow = ctk.CTkFrame(outer, fg_color="transparent")
        drow.pack(padx=10, pady=(6,2), fill="x")
        self._label(drow, "输出目录:").pack(side="left")
        import pathlib
        self.url_outdir_var = tk.StringVar(value=str(pathlib.Path.home()/"Downloads"))
        self._entry(drow, self.url_outdir_var, w=50).pack(side="left", padx=6)
        self._btn(drow, "📁", self._url_pick_dir, C["surface"]).pack(side="left")

        opt = ctk.CTkFrame(outer, fg_color="transparent")
        opt.pack(padx=10, pady=(4,0), fill="x")
        self.url_audio_var     = tk.BooleanVar(value=True)
        self.dl_video_var      = tk.BooleanVar(value=False)
        self.url_ocr_var       = tk.BooleanVar(value=False)
        self.url_translate_var = tk.BooleanVar(value=True)
        self._check(opt, "🎤 识别语音（Groq Whisper）",
                    self.url_audio_var).pack(side="left", padx=4)
        self._check(opt, "⬇ 同时下载视频", self.dl_video_var).pack(side="left", padx=4)
        self._check(opt, "🖼 识别画面文字", self.url_ocr_var).pack(side="left", padx=4)
        self._check(opt, "🌐 翻译为中文",
                    self.url_translate_var).pack(side="left", padx=10)
        ctk.CTkLabel(opt, text="⚠ Facebook 链接将直接下载视频",
                     font=("Segoe UI", 9), text_color=C["dim"]
                     ).pack(side="right", padx=8)
        self._ctrl_bar(outer, self._url_start)

    # ── Tab 3: Google Sheets ────────────────────────────────────────
    def _build_tab_sheets(self):
        outer = self.nb.tab("📊 Google Sheets")

        card = self._card_frame(outer)
        card.pack(padx=10, pady=(10,4), fill="x")
        g = tk.Frame(card, bg=C["card"])
        g.pack(fill="x", padx=10, pady=8)
        g.columnconfigure(1, weight=1)

        ekw = dict(font=("Consolas",9), bg=C["surface"], fg=C["text"],
                   insertbackground=C["accent"], relief="flat",
                   highlightbackground=C["border"], highlightthickness=1)
        lkw = dict(font=("Segoe UI",9,"bold"), bg=C["card"], fg=C["sub"])

        self.sh_url_var  = tk.StringVar()
        self.sh_name_var = tk.StringVar()
        for r, (lbl, var, w, cols) in enumerate([
            ("Sheet URL:",  self.sh_url_var,  70, 5),
            ("Sheet 名称:", self.sh_name_var, 28, 1),
        ]):
            tk.Label(g, text=lbl, **lkw).grid(row=r, column=0, sticky="w", padx=(0,8), pady=3)
            tk.Entry(g, textvariable=var, width=w, **ekw).grid(
                row=r, column=1, columnspan=cols, sticky="ew"if w>30 else "w", pady=3)

        for i, (lbl, attr, default) in enumerate([
            ("链接列:", "sh_link_col", "A"), ("原文列:", "sh_orig_col", "B"),
            ("中文列:", "sh_zh_col",   "C"), ("状态列:", "sh_stat_col", "D"),
        ]):
            c = 2 + i*2
            tk.Label(g, text=lbl, **lkw).grid(row=1, column=c, sticky="w", padx=(10,2))
            var = tk.StringVar(value=default); setattr(self, attr+"_var", var)
            tk.Entry(g, textvariable=var, width=4, **ekw).grid(row=1, column=c+1, padx=(0,4))

        opt = ctk.CTkFrame(outer, fg_color="transparent")
        opt.pack(padx=10, pady=(2,0), fill="x")
        self.skip_filled_var   = tk.BooleanVar(value=True)
        self.sh_audio_var      = tk.BooleanVar(value=True)
        self.sh_ocr_var        = tk.BooleanVar(value=False)
        self.sh_translate_var  = tk.BooleanVar(value=True)
        self._check(opt, "跳过原文列已有内容的行",
                    self.skip_filled_var).pack(side="left", padx=4)
        self._check(opt, "🎤 识别语音（Groq Whisper）",
                    self.sh_audio_var).pack(side="left", padx=10)
        self._check(opt, "🖼 识别画面文字",
                    self.sh_ocr_var).pack(side="left", padx=4)
        self._check(opt, "🌐 翻译为中文",
                    self.sh_translate_var).pack(side="left", padx=10)
        ctk.CTkLabel(opt, text="线程数:", font=("Segoe UI",9),
                     text_color=C["sub"]).pack(side="left", padx=(14,4))
        self.sh_threads_var = tk.IntVar(value=3)
        ttk.Spinbox(opt, from_=1, to=10, textvariable=self.sh_threads_var, width=5
                    ).pack(side="left")
        ctk.CTkLabel(opt, text="  批量写入(行):", font=("Segoe UI",9),
                     text_color=C["sub"]).pack(side="left", padx=(14,4))
        self.sh_batch_var = tk.IntVar(value=10)
        ttk.Spinbox(opt, from_=1, to=50, textvariable=self.sh_batch_var, width=5
                    ).pack(side="left")
        # 失败写入重试按鈕
        self._sh_failed_writes: list = []
        self.sh_retry_btn = self._btn(opt, "🔁 重试失败写入", self._sheets_retry_write, C["red"])
        self.sh_retry_btn.pack(side="right", padx=8)
        self.sh_retry_btn.configure(state="disabled")
        self._ctrl_bar(outer, self._sheets_start)

    # ── Tab 4: 图片 OCR ─────────────────────────────────────────────
    def _build_tab_ocr(self):
        outer = self.nb.tab("🖼 图片OCR")

        # 模式切换
        mf = self._card_frame(outer)
        mf.pack(padx=10, pady=(10,4), fill="x")
        mc = tk.Frame(mf, bg=C["card"]); mc.pack(fill="x", padx=10, pady=8)
        self.ocr_mode_var = tk.StringVar(value="folder")
        for text, val in [("📁 图片文件夹","folder"),
                           ("🔗 图片链接",  "urls"),
                           ("📊 Sheet 图片列","sheets")]:
            self._radio(mc, text, self.ocr_mode_var, val,
                        command=self._ocr_mode_switch
                        ).pack(side="left", padx=6, pady=4)

        # 面板容器
        self._ocr_panels: dict[str, tk.Frame] = {}

        # Folder 面板
        fp = tk.Frame(outer, bg=C["bg"]); self._ocr_panels["folder"] = fp
        fr = tk.Frame(fp, bg=C["bg"]); fr.pack(padx=10, pady=(4,2), fill="x")
        self._label(fr, "图片文件夹:").pack(side="left")
        self.ocr_dir_var = tk.StringVar()
        self._entry(fr, self.ocr_dir_var, w=50).pack(side="left", padx=6)
        self._btn(fr, "📁", self._ocr_pick_dir, C["surface"]).pack(side="left")

        # URLs 面板
        up = tk.Frame(outer, bg=C["bg"]); self._ocr_panels["urls"] = up
        self._label(up, "图片链接（每行一个，支持本地路径或 URL）:"
                    ).pack(padx=10, pady=(4,2), anchor="w")
        self.ocr_url_text = scrolledtext.ScrolledText(
            up, font=("Consolas",9), bg=C["card"], fg=C["text"],
            insertbackground=C["accent"], relief="flat",
            highlightbackground=C["border"], highlightthickness=1, height=4)
        self.ocr_url_text.pack(padx=10, fill="x")

        # Sheets 面板
        sp = tk.Frame(outer, bg=C["bg"]); self._ocr_panels["sheets"] = sp
        sg = tk.Frame(sp, bg=C["bg"]); sg.pack(padx=10, pady=(4,2), fill="x")
        ekw2 = dict(font=("Consolas",9), bg=C["card"], fg=C["text"],
                    insertbackground=C["accent"], relief="flat",
                    highlightbackground=C["border"], highlightthickness=1)
        for row, (lbl, attr, w) in enumerate([
            ("Sheet URL:",  "ocr_sh_url",  55),
            ("Sheet 名称:", "ocr_sh_name", 25),
        ]):
            tk.Label(sg, text=lbl, font=("Segoe UI",9,"bold"),
                     bg=C["bg"], fg=C["sub"]).grid(row=row, column=0, sticky="w", padx=(0,8), pady=3)
            var = tk.StringVar(); setattr(self, attr+"_var", var)
            tk.Entry(sg, textvariable=var, width=w, **ekw2).grid(row=row, column=1, sticky="w", pady=3)
        r2 = tk.Frame(sp, bg=C["bg"]); r2.pack(padx=10, pady=(2,4), fill="x")
        for lbl2, attr2, def2 in [
            ("🔗 图片列:",      "ocr_sh_img_col",  "A"),
            ("📝 原文列:",      "ocr_sh_orig_col", "B"),
            ("🌐 翻译列:",      "ocr_sh_zh_col",   "C"),
            ("🟥 状态列:",      "ocr_sh_stat_col", "D"),
        ]:
            tk.Label(r2, text=lbl2, font=("Segoe UI",9,"bold"),
                     bg=C["bg"], fg=C["sub"]).pack(side="left", padx=(0,2))
            v = tk.StringVar(value=def2); setattr(self, attr2+"_var", v)
            tk.Entry(r2, textvariable=v, width=4, **ekw2).pack(side="left", padx=(0,10))
        self.ocr_sh_skip_var = tk.BooleanVar(value=True)
        self._check(r2, "跳过原文列已有内容的行",
                    self.ocr_sh_skip_var).pack(side="left", padx=(10,0))

        # 选项
        topt = tk.Frame(outer, bg=C["bg"]); topt.pack(padx=10, pady=(4,0), fill="x")
        self.ocr_translate_var = tk.BooleanVar(value=True)
        self._check(topt, "翻译识别结果为中文（Gemini）",
                    self.ocr_translate_var).pack(side="left", padx=4)

        self._ctrl_bar(outer, self._ocr_start)
        self._ocr_mode_switch()

    def _ocr_mode_switch(self):
        mode = self.ocr_mode_var.get()
        for m, p in self._ocr_panels.items():
            if m == mode: p.pack(fill="x")
            else:         p.pack_forget()

    def _ocr_pick_dir(self):
        d = filedialog.askdirectory(title="选择图片文件夹")
        if d: self.ocr_dir_var.set(d)

    # ── 进度 / 表格 / 日志 ──────────────────────────────────────────
    def _build_progress(self):
        pf = tk.Frame(self, bg=C["bg"]); pf.pack(padx=12, pady=(8,0), fill="x")
        self.progress = ttk.Progressbar(
            pf, orient="horizontal", mode="determinate",
            style="TSep.Horizontal.TProgressbar")
        self.progress.pack(fill="x")
        self.prog_lbl = tk.Label(pf, text="就绪",
                                  font=("Segoe UI",9), bg=C["bg"], fg=C["dim"])
        self.prog_lbl.pack(anchor="w", pady=(2,0))

    def _build_table(self):
        self.result_table = ResultTable(self, bg=C["bg"])
        self.result_table.pack(padx=12, fill="both", expand=False)

    def _build_log(self):
        lf = ctk.CTkFrame(self, fg_color="transparent")
        lf.pack(padx=12, pady=(4,12), fill="both", expand=True)
        ctk.CTkLabel(lf, text="日志", font=("Segoe UI",9,"bold"),
                     text_color=C["accent"]).pack(anchor="w")
        self.log_box = ctk.CTkTextbox(
            lf, font=("Consolas",10), fg_color=C["log_bg"],
            text_color=C["log_fg"], border_color=C["border"],
            border_width=1, corner_radius=6, height=120,
            state="disabled")
        self.log_box.pack(fill="both", expand=True)

    # ── Log / Prog helpers ──────────────────────────────────────────
    def _log(self, msg):
        def _do():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg+"\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _do)

    def _upd_prog(self, done, total, name=""):
        pct = int(done/total*100) if total else 0
        self.after(0, lambda: (
            self.progress.config(value=pct),
            self.prog_lbl.configure(text=f"{pct}%  [{done}/{total}]  {name}")))

    def _add_row(self, source, original, chinese, status):
        self.after(0, lambda: self.result_table.add_row(source, original, chinese, status))

    # ── Deps check ──────────────────────────────────────────────────
    def _check_deps(self):
        checks = [
            (GROQ_OK,  "groq",              "pip install groq"),
            (GENAI_OK, "google-generativeai","pip install google-generativeai"),
            (GSPREAD_OK,"gspread",           "pip install gspread google-auth"),
            (WEBVTT_OK, "webvtt-py",         "pip install webvtt-py"),
            (os.path.isfile(YT_DLP_EXE), "yt-dlp.exe", f"{YT_DLP_EXE}"),
            (os.path.isfile(FFMPEG_EXE), "FFmpeg",      f"{FFMPEG_EXE}"),
            (os.path.isfile(LENS_JS),    "chrome-lens-ocr", f"{LENS_JS}"),
        ]
        issues = [f"  ⚠ {n}: {h}" for ok,n,h in checks if not ok]
        if issues:
            self._log("启动检查（缺失不影响其他功能）:")
            for i in issues: self._log(i)
        else:
            self._log("✅ 所有依赖就绪")

    # ── Config ──────────────────────────────────────────────────────
    def _load_config(self):
        if not os.path.isfile(CONFIG_FILE): return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
            self.groq_key_var.set(cfg.get("groq_key", self.groq_key_var.get()))
            self.gemini_key_var.set(cfg.get("gemini_key", self.gemini_key_var.get()))
            if cfg.get("gemini_model") in GEMINI_MODELS:
                self.gemini_model_var.set(cfg["gemini_model"])
            self.ocr_script_path_var.set(cfg.get("ocr_script_path", self.ocr_script_path_var.get()))
            self.url_outdir_var.set(cfg.get("out_dir", self.url_outdir_var.get()))
            self.sh_url_var.set(cfg.get("sheet_url",""))
            self.sh_name_var.set(cfg.get("sheet_name",""))
            for attr, key, default in [
                ("sh_link_col","link_col","A"), ("sh_orig_col","orig_col","B"),
                ("sh_zh_col","zh_col","C"),     ("sh_stat_col","stat_col","D"),
            ]:
                getattr(self, attr+"_var").set(cfg.get(key, default))
            self.sh_threads_var.set(cfg.get("threads",3))
            self.skip_filled_var.set(cfg.get("skip_filled",True))
            # 勾选状态
            self.mp4_ocr_var.set(    cfg.get("mp4_ocr",    False))
            self.mp4_audio_var.set(  cfg.get("mp4_audio",  True))
            self.mp4_translate_var.set(cfg.get("mp4_translate", True))
            self.dl_video_var.set(   cfg.get("dl_video",   False))
            self.url_ocr_var.set(    cfg.get("url_ocr",    False))
            self.url_audio_var.set(  cfg.get("url_audio",  True))
            self.url_translate_var.set(cfg.get("url_translate", True))
            self.sh_ocr_var.set(     cfg.get("sh_ocr",     False))
            self.sh_audio_var.set(   cfg.get("sh_audio",   True))
            self.sh_translate_var.set(cfg.get("sh_translate", True))
            self.ocr_translate_var.set(cfg.get("ocr_translate", True))
            # OCR Sheet 设置
            if cfg.get("ocr_sh_url"):      self.ocr_sh_url_var.set(cfg["ocr_sh_url"])
            if cfg.get("ocr_sh_name"):     self.ocr_sh_name_var.set(cfg["ocr_sh_name"])
            if cfg.get("ocr_sh_img_col"):  self.ocr_sh_img_col_var.set(cfg["ocr_sh_img_col"])
            if cfg.get("ocr_sh_orig_col"): self.ocr_sh_orig_col_var.set(cfg["ocr_sh_orig_col"])
            if cfg.get("ocr_sh_zh_col"):   self.ocr_sh_zh_col_var.set(cfg["ocr_sh_zh_col"])
            if cfg.get("ocr_sh_stat_col"): self.ocr_sh_stat_col_var.set(cfg["ocr_sh_stat_col"])
            self.ocr_sh_skip_var.set(cfg.get("ocr_sh_skip", True))
            # OCR 模式
            if cfg.get("ocr_mode") in ("folder","urls","sheets"):
                self.ocr_mode_var.set(cfg["ocr_mode"])
                self._ocr_mode_switch()
            # 命名笪设置
            if cfg.get("ren_sort") in ("mtime","ctime","name"):
                self.ren_sort_var.set(cfg["ren_sort"])
            if cfg.get("ren_sep"): self.ren_sep_var.set(cfg["ren_sep"])
            if cfg.get("ren_pad"): self.ren_pad_var.set(cfg["ren_pad"])
            # 音频转视频
            if cfg.get("a2v_src"): self.a2v_src_var.set(cfg["a2v_src"])
            if cfg.get("a2v_out"): self.a2v_out_var.set(cfg["a2v_out"])
            if cfg.get("a2v_ratio") in ("9:16","16:9","1:1"):
                self.a2v_ratio_var.set(cfg["a2v_ratio"])
            if cfg.get("a2v_bg"):       self.a2v_bg_var.set(cfg["a2v_bg"])
            if cfg.get("a2v_fontsize"): self.a2v_fontsize_var.set(cfg["a2v_fontsize"])
            if cfg.get("a2v_chars"):    self.a2v_chars_var.set(cfg["a2v_chars"])
            if cfg.get("a2v_font") and cfg["a2v_font"] in self._a2v_font_map:
                self.a2v_font_var.set(cfg["a2v_font"])
            if cfg.get("a2v_fg"):       self.a2v_fg_var.set(cfg["a2v_fg"])
            self.a2v_multiline_var.set(cfg.get("a2v_multiline", True))
            if cfg.get("a2v_audio_ext"): self.a2v_audio_ext_var.set(cfg["a2v_audio_ext"])
        except Exception as e:
            self._log(f"⚠ 读取 config.json: {e}")

    def _save_config(self):
        cfg = {
            "groq_key":      self.groq_key_var.get().strip(),
            "gemini_key":    self.gemini_key_var.get().strip(),
            "gemini_model":  self.gemini_model_var.get(),
            "ocr_script_path": self.ocr_script_path_var.get().strip(),
            "out_dir":       self.url_outdir_var.get().strip(),
            "sheet_url":     self.sh_url_var.get().strip(),
            "sheet_name":    self.sh_name_var.get().strip(),
            "link_col":      self.sh_link_col_var.get().strip().upper(),
            "orig_col":      self.sh_orig_col_var.get().strip().upper(),
            "zh_col":        self.sh_zh_col_var.get().strip().upper(),
            "stat_col":      self.sh_stat_col_var.get().strip().upper(),
            "threads":       self.sh_threads_var.get(),
            "skip_filled":   self.skip_filled_var.get(),
            # 勾选状态
            "mp4_ocr":       self.mp4_ocr_var.get(),
            "mp4_audio":     self.mp4_audio_var.get(),
            "mp4_translate": self.mp4_translate_var.get(),
            "dl_video":      self.dl_video_var.get(),
            "url_ocr":       self.url_ocr_var.get(),
            "url_audio":     self.url_audio_var.get(),
            "url_translate": self.url_translate_var.get(),
            "sh_ocr":        self.sh_ocr_var.get(),
            "sh_audio":      self.sh_audio_var.get(),
            "sh_translate":  self.sh_translate_var.get(),
            "ocr_translate": self.ocr_translate_var.get(),
            # OCR Sheet
            "ocr_sh_url":      self.ocr_sh_url_var.get().strip(),
            "ocr_sh_name":     self.ocr_sh_name_var.get().strip(),
            "ocr_sh_img_col":  self.ocr_sh_img_col_var.get().strip().upper(),
            "ocr_sh_orig_col": self.ocr_sh_orig_col_var.get().strip().upper(),
            "ocr_sh_zh_col":   self.ocr_sh_zh_col_var.get().strip().upper(),
            "ocr_sh_stat_col": self.ocr_sh_stat_col_var.get().strip().upper(),
            "ocr_sh_skip":     self.ocr_sh_skip_var.get(),
            "ocr_mode":        self.ocr_mode_var.get(),
            "sh_batch":      getattr(self, 'sh_batch_var', tk.IntVar(value=10)).get(),
            # 命名笪
            "ren_sort":      getattr(self, 'ren_sort_var', tk.StringVar(value='mtime')).get(),
            "ren_sep":       getattr(self, 'ren_sep_var',  tk.StringVar(value=' ')).get(),
            "ren_pad":       getattr(self, 'ren_pad_var',  tk.StringVar(value='1')).get(),
            # 音频转视频
            "a2v_src":       self.a2v_src_var.get().strip(),
            "a2v_out":       self.a2v_out_var.get().strip(),
            "a2v_ratio":     self.a2v_ratio_var.get(),
            "a2v_bg":        self.a2v_bg_var.get().strip(),
            "a2v_fontsize":  self.a2v_fontsize_var.get(),
            "a2v_chars":     self.a2v_chars_var.get(),
            "a2v_font":      self.a2v_font_var.get(),
            "a2v_fg":        self.a2v_fg_var.get().strip(),
            "a2v_multiline": self.a2v_multiline_var.get(),
            "a2v_audio_ext": self.a2v_audio_ext_var.get().strip(),
        }
        with open(CONFIG_FILE,"w",encoding="utf-8") as f:
            json.dump(cfg, f, indent=4, ensure_ascii=False)
        self._log("💾 配置已保存")

    def _on_close(self):
        """Override window close: auto-save all settings before quitting."""
        try: self._save_config()
        except Exception: pass
        self.destroy()

    # ── Pause / Stop ────────────────────────────────────────────────
    def _toggle_pause(self):
        if self._pause_evt.is_set():
            self._pause_evt.clear()
            for b in self._all_pause_btns:
                b.configure(text="▶  继续", fg_color=C["accent"])
            self._log("⏸ 已暂停")
        else:
            self._pause_evt.set()
            for b in self._all_pause_btns:
                b.configure(text="⏸ 暂停", fg_color=C["surface"])
            self._log("▶  继续")

    def _stop(self):
        self._stop_flag[0] = True; self._pause_evt.set()
        self._log("⏹ 停止中…")

    def _set_running(self, state: bool):
        self._running = state
        if not state:
            self._stop_flag[0] = False; self._pause_evt.set()
        s = "normal" if state else "disabled"
        for b in self._all_pause_btns + self._all_stop_btns:
            b.configure(state=s)
        if not state:
            for b in self._all_pause_btns:
                b.configure(text="⏸ 暂停", fg_color=C["surface"])

    def _groq_client(self):
        k = self.groq_key_var.get().strip()
        return _groq_lib.Groq(api_key=k) if GROQ_OK and k else None

    # ── MP4 ─────────────────────────────────────────────────────────
    def _mp4_on_drop(self, event):
        paths = [a or b for a,b in re.findall(r'\{([^}]+)\}|(\S+)', event.data)]
        self._mp4_enqueue(paths)

    def _mp4_pick(self):
        paths = filedialog.askopenfilenames(
            title="选择视频文件",
            filetypes=[("视频","*.mp4 *.mov *.mkv *.avi *.webm"),("所有","*.*")])
        if paths: self._mp4_enqueue(list(paths))

    def _mp4_enqueue(self, paths):
        files = collect_media(paths)
        added = [f for f in files if f not in self._mp4_queue]
        self._mp4_queue.extend(added); cnt = len(self._mp4_queue)
        self._log(f"➕ 新增 {len(added)} 个，队列共 {cnt} 个")
        self.mp4_lbl.configure(
            text=f"✅ 队列：{cnt} 个视频" if cnt
                 else "⬇  将视频文件 / 文件夹拖到这里")

    def _mp4_clear(self):
        self._mp4_queue.clear()
        self.mp4_lbl.configure(text="⬇  将视频文件 / 文件夹拖到这里")
        self._log("🗑 队列已清空")

    def _mp4_start(self):
        if not self._mp4_queue:
            messagebox.showwarning("提示","请先添加视频"); return
        if not self.mp4_audio_var.get() and not self.mp4_ocr_var.get():
            messagebox.showwarning("提示","请至少勾选「识别语音」或「识别画面文字」"); return
        if self.mp4_audio_var.get() and not self.groq_key_var.get().strip():
            messagebox.showerror("错误","识别语音需要填写 Groq API Key"); return
        self._set_running(True)
        threading.Thread(target=self._mp4_worker, daemon=True).start()

    def _mp4_worker(self):
        gc = self._groq_client(); gkey = self.gemini_key_var.get().strip()
        gmod = self.gemini_model_var.get()
        do_audio     = self.mp4_audio_var.get()
        do_ocr       = self.mp4_ocr_var.get()
        do_translate = self.mp4_translate_var.get()
        total = len(self._mp4_queue); tmp = tempfile.mkdtemp(prefix="mp4t_")
        try:
            for i, mp4 in enumerate(self._mp4_queue):
                self._pause_evt.wait()
                if self._stop_flag[0]: break
                name = os.path.basename(mp4)
                self._upd_prog(i, total, name)
                self._log(f"\n[{i+1}/{total}] {name}")
                orig = zh = status = ""
                try:
                    audio_text = ""
                    ocr_text   = ""

                    # Ⅰ 语音识别
                    if do_audio:
                        self._log("  ⚙ 提取音频…")
                        wav = ffmpeg_extract_wav(mp4, tmp)
                        self._pause_evt.wait()
                        if self._stop_flag[0]: break
                        self._log("  🎤 Whisper 转录…")
                        audio_text = groq_transcribe(gc, wav)
                        self._log(f"  语音: {audio_text[:80]}…")

                    # Ⅱ 画面 OCR
                    if do_ocr and os.path.isfile(LENS_JS):
                        try:
                            from backend import ffmpeg_extract_frame, lens_ocr as _ocr
                            self._log("  🖼 截帧 OCR…")
                            frm = ffmpeg_extract_frame(mp4, tmp)
                            ocr_text = _ocr(frm)
                            self._log(f"  画面: {ocr_text[:80]}…")
                        except Exception as e:
                            self._log(f"  ⚠ OCR: {e}")

                    # 拼合原文
                    if audio_text and ocr_text:
                        orig = audio_text + f"\n\n[画面文字]\n{ocr_text}"
                    else:
                        orig = audio_text or ocr_text

                    # Ⅲ 翻译（无论成败都写入）
                    if do_translate and GENAI_OK and gkey and orig:
                        self._pause_evt.wait()
                        if self._stop_flag[0]: break
                        self._log("  🌐 Gemini 翻译…")
                        zh = gemini_translate(orig, gkey, gmod, self._log)

                    status = "成功"
                except Exception as e:
                    status = "失败"; self._log(f"  ❌ {e}")
                self._add_row(name, orig, zh, status)
            self._upd_prog(total, total, "完成")
            self._log("\n✅ MP4 转录完毕")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            self.after(0, lambda: self._set_running(False))

    # ── URL ─────────────────────────────────────────────────────────
    def _url_pick_dir(self):
        d = filedialog.askdirectory(); 
        if d: self.url_outdir_var.set(d)

    def _url_start(self):
        urls = [u.strip() for u in self.url_text.get("1.0","end").splitlines() if u.strip()]
        if not urls: messagebox.showwarning("提示","请输入链接"); return
        if not os.path.isfile(YT_DLP_EXE):
            messagebox.showerror("错误",f"yt-dlp.exe 不存在:\n{YT_DLP_EXE}"); return
        out_dir = self.url_outdir_var.get().strip()
        os.makedirs(out_dir, exist_ok=True)
        self._set_running(True)
        threading.Thread(target=self._url_worker, args=(urls, out_dir), daemon=True).start()

    def _url_worker(self, urls, out_dir):
        gc = self._groq_client(); gkey = self.gemini_key_var.get().strip()
        gmod = self.gemini_model_var.get()
        dl    = self.dl_video_var.get()
        ocr   = self.url_ocr_var.get()
        audio = self.url_audio_var.get()
        do_translate = self.url_translate_var.get()
        ocr_script = self.ocr_script_path_var.get().strip()
        total = len(urls)
        try:
            for i, url in enumerate(urls):
                self._pause_evt.wait()
                if self._stop_flag[0]: break
                self._upd_prog(i, total, url[:50])
                self._log(f"\n[{i+1}/{total}]")
                res = process_url(url, out_dir, gc, gkey, gmod,
                                  dl, ocr, ocr_script, self._log,
                                  self._pause_evt, self._stop_flag,
                                  audio_ok=audio, translate_ok=do_translate)
                self._add_row(res["source"], res["original"], res["chinese"], res["status"])
            self._upd_prog(total, total, "完成")
            self._log("\n✅ 链接处理完毕")
        finally:
            self.after(0, lambda: self._set_running(False))

    # ── Sheets ──────────────────────────────────────────────────────
    def _sheets_start(self):
        if not GSPREAD_OK:
            messagebox.showerror("错误","pip install gspread google-auth"); return
        if not os.path.isfile(CREDS_FILE):
            messagebox.showerror("错误",f"找不到 credentials.json:\n{CREDS_FILE}"); return
        if not self.sh_url_var.get().strip():
            messagebox.showerror("错误","请填写 Sheet URL"); return
        self._save_config(); self._set_running(True)
        threading.Thread(target=self._sheets_worker, daemon=True).start()

    def _sheets_worker(self):
        import time as _time
        gc = self._groq_client(); gkey = self.gemini_key_var.get().strip()
        gmod = self.gemini_model_var.get()
        sh_url = self.sh_url_var.get().strip(); sh_name = self.sh_name_var.get().strip()
        lc = self.sh_link_col_var.get().strip().upper()
        oc = self.sh_orig_col_var.get().strip().upper()
        zc = self.sh_zh_col_var.get().strip().upper()
        sc = self.sh_stat_col_var.get().strip().upper()
        skip = self.skip_filled_var.get(); nthrd = self.sh_threads_var.get()
        ocr = self.sh_ocr_var.get(); audio = self.sh_audio_var.get()
        do_translate = self.sh_translate_var.get()
        batch_sz = self.sh_batch_var.get()
        tmp = tempfile.mkdtemp(prefix="shts_")

        def ci(c): return ord(c)-ord('A')
        def sg(row, i): return row[i].strip() if i<len(row) else ""

        def _ws_factory():
            creds2 = Credentials.from_service_account_file(
                CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
            return gspread.authorize(creds2).open_by_url(sh_url).worksheet(sh_name)

        def _batch_write(updates: list, tag=""):
            """updates = [(cell_a1, value), ...]. Retry 3x with 30s wait."""
            for attempt in range(3):
                try:
                    ws_w = _ws_factory()
                    data = [{"range": addr, "values": [[val]]}
                            for addr, val in updates if val is not None]
                    if data: ws_w.batch_update(data)
                    return True
                except Exception as e:
                    if attempt < 2:
                        self._log(f"  ⚠ [{tag}] 写入失败（第{attempt+1}次），30秒后重试…: {e}")
                        _time.sleep(30)
                    else:
                        self._log(f"  ❌ [{tag}] 写入失败，已加入重试队列: {e}")
                        self._sh_failed_writes.extend(updates)
                        self.after(0, lambda: self.sh_retry_btn.configure(state="normal"))
                        return False

        try:
            self._log("🔗 连接 Google Sheets…")
            creds = Credentials.from_service_account_file(
                CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
            ws = gspread.authorize(creds).open_by_url(sh_url).worksheet(sh_name)
            all_vals = ws.get_all_values()
            if not all_vals: self._log("⚠ 表格为空"); return
            self._log("✅ 连接成功")

            tasks = [(ri, sg(row, ci(lc)))
                     for ri, row in enumerate(all_vals[1:], start=2)
                     if sg(row, ci(lc)) and not (skip and sg(row, ci(oc)))]
            total = len(tasks)
            self._log(f"共 {total} 行待处理，跳过 {len(all_vals)-1-total} 行")
            if not total: return

            done_c = [0]; lock = threading.Lock()
            pending_writes: list = []   # buffered (cell, value) pairs
            pend_lock = threading.Lock()

            def flush_pending():
                with pend_lock:
                    if not pending_writes: return
                    chunk = pending_writes[:]
                    pending_writes.clear()
                _batch_write(chunk, "flush")

            def proc(task):
                ri, url = task
                self._pause_evt.wait()
                if self._stop_flag[0]: return
                self._log(f"\n[行 {ri}] {url[:60]}…")
                # 副状态卷 — 单独写回（不入批量队列）
                try:
                    _ws_factory().update_acell(f"{sc}{ri}", "处理中…")
                except Exception:
                    pass

                res = process_url(url, tmp, gc, gkey, gmod,
                                  False, ocr, self._log,
                                  self._pause_evt, self._stop_flag,
                                  audio_ok=audio, translate_ok=do_translate)
                self._add_row(f"行{ri}: {res['source'][:35]}",
                              res["original"], res["chinese"], res["status"])

                # 加入批量队列
                stat_val = "已完成" if res["status"].startswith("成功") else f"失败:{res['note'][:60]}"
                with pend_lock:
                    pending_writes.extend([
                        (f"{oc}{ri}", res["original"]),
                        (f"{zc}{ri}", res["chinese"]),
                        (f"{sc}{ri}", stat_val),
                    ])
                    if len(pending_writes) >= batch_sz * 3:
                        chunk = pending_writes[:]
                        pending_writes.clear()
                        _batch_write(chunk, f"batch@{ri}")

                with lock:
                    done_c[0] += 1
                    self._upd_prog(done_c[0], total, url[:35])

            with ThreadPoolExecutor(max_workers=nthrd) as ex:
                ex.map(proc, tasks)

            flush_pending()   # 写入剩余
            failed = len(self._sh_failed_writes)
            self._log(f"\n✅ Google Sheets 处理完毕（失败队列: {failed} 条）")
        except Exception as e:
            self._log(f"❌ {e}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            self.after(0, lambda: self._set_running(False))

    def _sheets_retry_write(self):
        """One-click retry of all buffered failed writes."""
        import time as _time
        pending = self._sh_failed_writes[:]
        self._sh_failed_writes.clear()
        self.sh_retry_btn.configure(state="disabled")
        if not pending:
            self._log("⚠ 重试队列为空"); return
        self._log(f"🔁 开始重试 {len(pending)} 条写入请求…")
        sh_url  = self.sh_url_var.get().strip()
        sh_name = self.sh_name_var.get().strip()

        def _do():
            for attempt in range(3):
                try:
                    creds = Credentials.from_service_account_file(
                        CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
                    ws = gspread.authorize(creds).open_by_url(sh_url).worksheet(sh_name)
                    data = [{"range": a, "values": [[v]]} for a,v in pending if v is not None]
                    ws.batch_update(data)
                    self._log(f"✅ 重试完成，{len(pending)//3} 行已写入")
                    return
                except Exception as e:
                    if attempt < 2:
                        self._log(f"  ⚠ 重试失败 ({attempt+1}/3), 30s…: {e}")
                        _time.sleep(30)
                    else:
                        self._log(f"  ❌ 重试最终失败: {e}")
                        self._sh_failed_writes.extend(pending)
                        self.after(0, lambda: self.sh_retry_btn.configure(state="normal"))
        threading.Thread(target=_do, daemon=True).start()

    # ── OCR Tab ─────────────────────────────────────────────────────
    def _ocr_start(self):
        if not os.path.isfile(LENS_JS):
            messagebox.showerror("错误",
                f"找不到 sharex.js:\n{LENS_JS}\n\n"
                "请将 chrome-lens-ocr-main 文件夹放到脚本目录，\n"
                "或确认 C:\\chrome-lens-ocr-main\\sharex.js 存在"); return
        mode = self.ocr_mode_var.get()
        gkey = self.gemini_key_var.get().strip()
        gmod = self.gemini_model_var.get()
        do_t = self.ocr_translate_var.get()

        if mode == "folder":
            d = self.ocr_dir_var.get().strip()
            if not d or not os.path.isdir(d):
                messagebox.showerror("错误","请选择有效的图片文件夹"); return
            sources = collect_images([d])
        elif mode == "urls":
            sources = [u.strip() for u in
                       self.ocr_url_text.get("1.0","end").splitlines() if u.strip()]
        else:  # sheets
            if not GSPREAD_OK:
                messagebox.showerror("错误","pip install gspread google-auth"); return
            sources = None   # loaded in worker

        if sources is not None and not sources:
            messagebox.showwarning("提示","没有找到图片"); return
        self._set_running(True)
        threading.Thread(target=self._ocr_worker,
                         args=(mode, sources, gkey, gmod, do_t), daemon=True).start()

    def _ocr_worker(self, mode, sources, gkey, gmod, do_t):
        try:
            sh_url = sh_name = zc = sc = ws_ref = None
            if mode == "sheets":
                sh_url  = self.ocr_sh_url_var.get().strip()
                sh_name = self.ocr_sh_name_var.get().strip()
                img_col = self.ocr_sh_img_col_var.get().strip().upper()
                zc      = self.ocr_sh_zh_col_var.get().strip().upper()
                sc      = self.ocr_sh_stat_col_var.get().strip().upper()
                creds = Credentials.from_service_account_file(
                    CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
                ws_ref = gspread.authorize(creds).open_by_url(sh_url).worksheet(sh_name)
                all_vals = ws_ref.get_all_values()
                idx     = ord(img_col) - ord('A')
                oc      = self.ocr_sh_orig_col_var.get().strip().upper()
                skip_fl = self.ocr_sh_skip_var.get()
                oc_idx  = ord(oc) - ord('A') if oc else -1
                # list of (row_index, url), skip if orig already filled
                sources = [
                    (ri, r[idx].strip())
                    for ri, r in enumerate(all_vals[1:], start=2)
                    if idx < len(r) and r[idx].strip()
                    and not (skip_fl and oc_idx >= 0
                             and oc_idx < len(r) and r[oc_idx].strip())
                ]
                total_rows = len(all_vals) - 1
                skipped = total_rows - len(sources)
                if not sources:
                    self._log(f"⚠ Sheet 中无待处理图片链接（共{total_rows}行，跳过{skipped}行）"); return
                self._log(f"共 {total_rows} 行，跳过 {skipped} 行，待处理 {len(sources)} 行")

            total = len(sources)
            self._log(f"🖼 共 {total} 张图片待 OCR")

            for i, item in enumerate(sources):
                self._pause_evt.wait()
                if self._stop_flag[0]: break
                # item is either a plain URL/path or (row_index, url) for sheets mode
                if isinstance(item, tuple):
                    ri, src = item
                else:
                    ri, src = None, item

                self._upd_prog(i, total, src[:50])
                self._log(f"\n[{i+1}/{total}] {src[:70]}")
                ocr_script = self.ocr_script_path_var.get().strip()
                res = process_image_ocr(src, gkey, gmod, do_t, ocr_script, self._log)

                # ✅ 写入表格（关键修复）
                display_src = f"行{ri}: {src[:40]}" if ri else src[:55]
                self._add_row(display_src, res["original"], res["chinese"], res["status"])

                # 写回 Google Sheet（sheets 模式）
                if ri and ws_ref:
                    try:
                        creds2 = Credentials.from_service_account_file(
                            CREDS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
                        ws2 = gspread.authorize(creds2).open_by_url(sh_url).worksheet(sh_name)
                        oc = self.ocr_sh_orig_col_var.get().strip().upper()
                        if oc: ws2.update_acell(f"{oc}{ri}", res["original"])
                        if zc: ws2.update_acell(f"{zc}{ri}", res["chinese"])
                        if sc: ws2.update_acell(f"{sc}{ri}",
                                                "已完成" if res["status"]=="成功" else "失败")
                    except Exception as e:
                        self._log(f"  ⚠ 写回: {e}")

            self._upd_prog(total, total, "完成")
            self._log(f"\n✅ OCR 完毕，共 {total} 张")
        except Exception as e:
            self._log(f"❌ {e}")
        finally:
            self.after(0, lambda: self._set_running(False))


    # ── Tab 5: 文件命名 ───────────────────────────────────────────────
    def _build_tab_rename(self):
        import os as _os
        outer = self.nb.tab("🗂 文件命名")

        # ── 文件夹 + 扩展名 ──
        row1 = ctk.CTkFrame(outer, fg_color="transparent")
        row1.pack(padx=10, pady=(10,2), fill="x")
        self._label(row1, "目标文件夹:").pack(side="left")
        self.ren_dir_var = tk.StringVar()
        self._entry(row1, self.ren_dir_var, w=46).pack(side="left", padx=6)
        self._btn(row1, "📁",
                  lambda: (d:=filedialog.askdirectory()) and self.ren_dir_var.set(d),
                  C["surface"]).pack(side="left")

        # ── 文件过滤 ──
        row2 = ctk.CTkFrame(outer, fg_color="transparent")
        row2.pack(padx=10, pady=(4,0), fill="x")
        self._label(row2, "文件类型:").pack(side="left")
        self.ren_ext_var = tk.StringVar(value="*")
        presets = ["* (所有文件)", ".zip", ".mp4", ".jpg", ".png", ".pdf", ".txt"]
        ctk.CTkComboBox(row2, values=presets, variable=self.ren_ext_var,
                        width=130, font=("Segoe UI",10),
                        fg_color=C["card"], text_color=C["text"],
                        border_color=C["border"], button_color=C["accent"],
                        button_hover_color=C["accent2"],
                        dropdown_fg_color=C["card"], dropdown_text_color=C["text"]
                        ).pack(side="left", padx=6)
        self._label(row2, "或手动输入后缀（如 .rar）:").pack(side="left", padx=(10,4))
        self.ren_custom_ext = tk.StringVar()
        ctk.CTkEntry(row2, textvariable=self.ren_custom_ext, width=70,
                     font=("Consolas",10), fg_color=C["card"], text_color=C["text"],
                     border_color=C["border"], corner_radius=6).pack(side="left")

        # ── 排序 / 格式 ──
        row3 = tk.Frame(outer, bg=C["bg"])
        row3.pack(padx=10, pady=(6,0), fill="x")

        self._label(row3, "排序依据:").pack(side="left")
        self.ren_sort_var = tk.StringVar(value="mtime")
        for text, val in [("修改时间","mtime"),("创建时间","ctime"),("文件名","name")]:
            self._radio(row3, text, self.ren_sort_var, val
                        ).pack(side="left", padx=3)

        self._label(row3, "  序号格式:").pack(side="left", padx=(14,4))
        self.ren_pad_var = tk.StringVar(value="1")
        ttk.Combobox(row3, textvariable=self.ren_pad_var,
                     values=["1 (1,2,3…)","2 (01,02…)","3 (001,002…)","4 (0001…)"],
                     width=14, state="readonly", font=("Segoe UI",9)
                     ).pack(side="left", padx=4)

        self._label(row3, "  分隔符:").pack(side="left", padx=(10,4))
        self.ren_sep_var = tk.StringVar(value=" ")
        for txt, val in [("空格"," "),("下划线 _","_"),("短线 -","-"),("点 .",".")]:
            self._radio(row3, txt, self.ren_sep_var, val
                        ).pack(side="left", padx=3)

        # ── 按钮 ──
        brow = tk.Frame(outer, bg=C["bg"])
        brow.pack(padx=10, pady=(8,4), fill="x")
        self._btn(brow, "🔍 预览（不执行）", self._rename_preview, C["surface"]
                  ).pack(side="left", padx=4)
        self._btn(brow, "✅ 执行重命名", self._rename_execute, C["accent"]
                  ).pack(side="left", padx=4)
        self.ren_undo_btn = self._btn(brow, "↩ 撤销上次", self._rename_undo, C["red"])
        self.ren_undo_btn.pack(side="left", padx=4)
        self.ren_undo_btn.configure(state="disabled")
        self.ren_status_lbl = tk.Label(brow, text="",
                                       font=("Segoe UI",9), bg=C["bg"], fg=C["accent"])
        self.ren_status_lbl.pack(side="left", padx=12)

        # ── 预览表格 ──
        wrap = tk.Frame(outer, bg=C["border"], bd=1)
        wrap.pack(padx=10, pady=(2,8), fill="both", expand=True)
        vsb = ttk.Scrollbar(wrap, orient="vertical")
        vsb.pack(side="right", fill="y")
        cols = ("#", "原文件名", "新文件名", "大小", "日期")
        self.ren_tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                      yscrollcommand=vsb.set, height=10)
        vsb.config(command=self.ren_tree.yview)
        for col, w in zip(cols, (40, 340, 340, 80, 140)):
            self.ren_tree.heading(col, text=col)
            self.ren_tree.column(col, width=w, minwidth=30)
        self.ren_tree.pack(fill="both", expand=True)

        self._ren_log: list[tuple] = []   # [(new_path, old_path), ...]

    # ── Rename helpers ───────────────────────────────────────────────
    def _ren_collect(self) -> list[str] | None:
        """返回按所选顺序排序的文件列表，失败返回 None"""
        folder = self.ren_dir_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("错误", "请选择有效的文件夹"); return None

        ext = self.ren_custom_ext.get().strip() or self.ren_ext_var.get()
        ext = ext.split("(")[0].strip()   # 去掉 Combobox 描述部分
        if ext == "*": ext = ""
        if ext and not ext.startswith("."): ext = "." + ext

        files = [os.path.join(folder, f) for f in os.listdir(folder)
                 if os.path.isfile(os.path.join(folder, f))
                 and (not ext or f.lower().endswith(ext.lower()))]

        sort_key = self.ren_sort_var.get()
        if sort_key == "mtime":
            files.sort(key=lambda p: os.path.getmtime(p))
        elif sort_key == "ctime":
            files.sort(key=lambda p: os.path.getctime(p))
        else:
            files.sort(key=lambda p: os.path.basename(p).lower())
        return files

    def _ren_make_name(self, idx: int, total: int, original_name: str) -> str:
        pad_raw = self.ren_pad_var.get()
        pad = int(pad_raw[0]) if pad_raw[0].isdigit() else 1
        sep = self.ren_sep_var.get()
        num = str(idx).zfill(pad)
        return f"{num}{sep}{original_name}"

    def _rename_preview(self):
        files = self._ren_collect()
        if files is None: return
        self.ren_tree.delete(*self.ren_tree.get_children())
        total = len(files)
        if total == 0:
            self.ren_status_lbl.config(text="⚠ 没有匹配的文件", fg=C["red"]); return

        from datetime import datetime as _dt
        for i, path in enumerate(files, start=1):
            old_name = os.path.basename(path)
            new_name = self._ren_make_name(i, total, old_name)
            size     = os.path.getsize(path)
            mtime    = _dt.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
            size_str = (f"{size//1024//1024} MB" if size > 1048576
                        else f"{size//1024} KB" if size > 1024
                        else f"{size} B")
            self.ren_tree.insert("", "end", values=(i, old_name, new_name, size_str, mtime))
        self.ren_status_lbl.config(
            text=f"预览 {total} 个文件，确认后点「执行重命名」", fg=C["accent"])

    def _rename_execute(self):
        files = self._ren_collect()
        if files is None: return
        if not files:
            messagebox.showinfo("提示", "没有匹配的文件"); return

        confirm = messagebox.askyesno(
            "确认重命名",
            f"将重命名 {len(files)} 个文件。\n操作可撤销。确定继续？")
        if not confirm: return

        total = len(files); log = []; errors = []
        for i, old_path in enumerate(files, start=1):
            old_name = os.path.basename(old_path)
            new_name = self._ren_make_name(i, total, old_name)
            new_path = os.path.join(os.path.dirname(old_path), new_name)
            if old_path == new_path: continue
            if os.path.exists(new_path):
                errors.append(f"已存在: {new_name}")
                continue
            try:
                os.rename(old_path, new_path)
                log.append((new_path, old_path))   # (current, original)
            except Exception as e:
                errors.append(f"{old_name}: {e}")

        # 保存撤销日志
        self._ren_log = log
        log_path = os.path.join(self.ren_dir_var.get().strip(), "rename_log.json")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                json.dump(log, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

        self.ren_undo_btn.configure(state="normal" if log else "disabled")
        self._rename_preview()   # 刷新显示
        msg = f"✅ 完成 {len(log)} 个"
        if errors: msg += f"，{len(errors)} 个失败"
        self.ren_status_lbl.config(text=msg, fg=C["accent"] if not errors else C["red"])
        if errors:
            messagebox.showwarning("部分失败", "\n".join(errors[:10]))

    def _rename_undo(self):
        log = self._ren_log
        # 尝试从文件加载（如果内存为空）
        if not log:
            log_path = os.path.join(self.ren_dir_var.get().strip(), "rename_log.json")
            if os.path.isfile(log_path):
                try:
                    with open(log_path, "r", encoding="utf-8") as f:
                        log = json.load(f)
                except Exception:
                    pass
        if not log:
            messagebox.showinfo("提示", "没有可撤销的记录"); return

        confirm = messagebox.askyesno("确认撤销", f"将还原 {len(log)} 个文件名。确定？")
        if not confirm: return

        ok = err = 0
        for new_path, old_path in reversed(log):
            if not os.path.isfile(new_path):
                err += 1; continue
            try:
                os.rename(new_path, old_path); ok += 1
            except Exception:
                err += 1

        self._ren_log.clear()
        self.ren_undo_btn.configure(state="disabled")
        self._rename_preview()
        self.ren_status_lbl.config(
            text=f"↩ 撤销完成：{ok} 个成功{'，'+str(err)+'个失败' if err else ''}",
            fg=C["accent"])

    # ── Tab 6: 音频转视频 ─────────────────────────────────────────────
    def _build_tab_audio2video(self):
        outer = self.nb.tab("🎵 音频转视频")

        card = self._card_frame(outer)
        card.pack(padx=10, pady=(10,4), fill="x")
        g = tk.Frame(card, bg=C["card"]); g.pack(fill="x", padx=10, pady=8)
        ekw = dict(font=("Consolas",9), bg=C["surface"], fg=C["text"],
                   insertbackground=C["accent"], relief="flat",
                   highlightbackground=C["border"], highlightthickness=1)
        lkw = dict(font=("Segoe UI",9,"bold"), bg=C["card"], fg=C["sub"])

        # ── 音频文件夹 + 输出目录 ──
        tk.Label(g, text="音频文件夹:", **lkw).grid(row=0, column=0, sticky="w", pady=3, padx=(0,8))
        self.a2v_src_var = tk.StringVar()
        tk.Entry(g, textvariable=self.a2v_src_var, width=52, **ekw).grid(row=0, column=1, columnspan=3, sticky="ew", pady=3)
        self._btn(g, "📁",
                  lambda: [self.a2v_src_var.set(d := filedialog.askdirectory() or self.a2v_src_var.get()),
                            self.a2v_out_var.get() or self.a2v_out_var.set(d)],
                  C["surface"]).grid(row=0, column=4, padx=(4,0), pady=3)

        tk.Label(g, text="输出目录:", **lkw).grid(row=1, column=0, sticky="w", pady=3, padx=(0,8))
        self.a2v_out_var = tk.StringVar()
        tk.Entry(g, textvariable=self.a2v_out_var, width=52, **ekw).grid(row=1, column=1, columnspan=3, sticky="ew", pady=3)
        self._btn(g, "📁",
                  lambda: (d:=filedialog.askdirectory()) and self.a2v_out_var.set(d),
                  C["surface"]).grid(row=1, column=4, padx=(4,0), pady=3)
        g.columnconfigure(1, weight=1)

        # ── 选项卡 ──
        opt = self._card_frame(outer)
        opt.pack(padx=10, pady=(2,4), fill="x")
        of = tk.Frame(opt, bg=C["card"]); of.pack(fill="x", padx=10, pady=8)

        # 宽高比
        tk.Label(of, text="📱 画面比例:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=0, column=0, sticky="w", padx=(0,10))
        self.a2v_ratio_var = tk.StringVar(value="9:16")
        for i, (txt, val) in enumerate([("9:16 竖屏","9:16"),("16:9 横屏","16:9"),("1:1 方形","1:1")]):
            self._radio(of, txt, self.a2v_ratio_var, val
                        ).grid(row=0, column=i+1, padx=4)

        # 背景颜色
        tk.Label(of, text="🎨 背景:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=0, column=5, sticky="w", padx=(16,6))
        self.a2v_bg_var = tk.StringVar(value="#1a1a2e")
        COLORS = [("深藋蓝","#1a1a2e"),("纯黑","#000000"),
                  ("Spotify绿","#0d1b12"),("S深红","#1a0a0a"),("费列灵蓝","#0a0a1a")]
        ttk.Combobox(of, textvariable=self.a2v_bg_var,
                     values=[c for _,c in COLORS], width=10
                     ).grid(row=0, column=6, padx=4)
        # 颜色预览块
        self._bg_preview = tk.Label(of, width=3, relief="flat",
                                    bg=self.a2v_bg_var.get())
        self._bg_preview.grid(row=0, column=7, padx=(0,4))
        def _update_preview(*_):
            try: self._bg_preview.config(bg=self.a2v_bg_var.get())
            except Exception: pass
        self.a2v_bg_var.trace_add("write", _update_preview)

        tk.Label(of, text="  字体大小:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=1, column=0, sticky="w", pady=(6,0))
        self.a2v_fontsize_var = tk.IntVar(value=100)
        ttk.Spinbox(of, from_=20, to=300, textvariable=self.a2v_fontsize_var,
                    width=6).grid(row=1, column=1, sticky="w", pady=(6,0))

        tk.Label(of, text="  显示字数:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=1, column=2, sticky="w", padx=(10,4), pady=(6,0))
        self.a2v_chars_var = tk.IntVar(value=10)
        ttk.Spinbox(of, from_=1, to=50, textvariable=self.a2v_chars_var,
                    width=5).grid(row=1, column=3, sticky="w", pady=(6,0))

        self.a2v_multiline_var = tk.BooleanVar(value=True)
        self._check(of, "自动换行",
                    self.a2v_multiline_var).grid(row=1, column=4, columnspan=1,
                                                sticky="w", padx=(10,0), pady=(6,0))

        tk.Label(of, text="  音频格式:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=1, column=5, sticky="w", padx=(10,4), pady=(6,0))
        self.a2v_audio_ext_var = tk.StringVar(value="mp3,wav,m4a,flac,aac,ogg")
        tk.Entry(of, textvariable=self.a2v_audio_ext_var, width=22,
                 font=("Consolas",9), bg=C["surface"], fg=C["text"],
                 insertbackground=C["accent"], relief="flat",
                 highlightbackground=C["border"], highlightthickness=1
                 ).grid(row=1, column=6, columnspan=2, sticky="w", pady=(6,0))

        # 字体选择 ─────────────────────────────────────────────
        import glob as _glob
        _font_files = sorted(
            _glob.glob(r"C:\Windows\Fonts\*.ttc") +
            _glob.glob(r"C:\Windows\Fonts\*.ttf"))
        self._a2v_font_map = {os.path.basename(f): f for f in _font_files}
        _preferred = ["msyh.ttc","simhei.ttf","simfang.ttf","arial.ttf","msyhbd.ttc"]
        _default_font = next((k for k in _preferred if k in self._a2v_font_map),
                              next(iter(self._a2v_font_map), ""))

        tk.Label(of, text="🔤 字体:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=2, column=0, sticky="w", pady=(6,0))
        self.a2v_font_var = tk.StringVar(value=_default_font)
        ttk.Combobox(of, textvariable=self.a2v_font_var,
                     values=list(self._a2v_font_map.keys()), width=22
                     ).grid(row=2, column=1, columnspan=3, sticky="w", pady=(6,0), padx=(0,4))

        tk.Label(of, text="  文字颜色:", font=("Segoe UI",9,"bold"),
                 bg=C["card"], fg=C["sub"]).grid(row=2, column=4, sticky="w", padx=(10,4), pady=(6,0))
        self.a2v_fg_var = tk.StringVar(value="#ffffff")
        tk.Entry(of, textvariable=self.a2v_fg_var, width=10,
                 font=("Consolas",9), bg=C["surface"], fg=C["text"],
                 insertbackground=C["accent"], relief="flat",
                 highlightbackground=C["border"], highlightthickness=1
                 ).grid(row=2, column=5, sticky="w", pady=(6,0))

        tk.Label(of, text="  (没有 Pillow 则安装: pip install Pillow)",
                 font=("Segoe UI",8), bg=C["card"], fg=C["dim"]
                 ).grid(row=2, column=6, columnspan=2, sticky="w", pady=(6,0))

        # ── 进度条 ──
        pf = tk.Frame(outer, bg=C["bg"])
        pf.pack(padx=10, pady=(4,0), fill="x")
        self.a2v_prog = ttk.Progressbar(
            pf, orient="horizontal", mode="determinate",
            style="TSep.Horizontal.TProgressbar")
        self.a2v_prog.pack(fill="x")
        self.a2v_prog_lbl = tk.Label(
            pf, text="就绪", font=("Segoe UI",9), bg=C["bg"], fg=C["dim"])
        self.a2v_prog_lbl.pack(anchor="w", pady=(2,0))

        # ── 按钮行 ──
        brow = tk.Frame(outer, bg=C["bg"])
        brow.pack(padx=10, pady=(6,8), fill="x")
        self._btn(brow, "▶  开始转换", self._a2v_start, C["accent"]
                  ).pack(side="right", padx=4)
        self.a2v_stop_var = [False]
        self._btn(brow, "⏹ 停止",
                  lambda: self.a2v_stop_var.__setitem__(0, True),
                  C["red"]).pack(side="right", padx=4)
        self.a2v_status_lbl = tk.Label(
            brow, text="", font=("Segoe UI",9),
            bg=C["bg"], fg=C["accent"])
        self.a2v_status_lbl.pack(side="left", padx=4)

        # ── 文件列表 ──
        wrap = tk.Frame(outer, bg=C["border"], bd=1)
        wrap.pack(padx=10, pady=(2,10), fill="both", expand=True)
        vsb = ttk.Scrollbar(wrap, orient="vertical")
        vsb.pack(side="right", fill="y")
        cols = ("#", "音频文件", "标题文字", "输出文件", "状态")
        self.a2v_tree = ttk.Treeview(
            wrap, columns=cols, show="headings",
            yscrollcommand=vsb.set, height=8)
        vsb.config(command=self.a2v_tree.yview)
        for col, w in zip(cols, (40, 260, 140, 260, 80)):
            self.a2v_tree.heading(col, text=col)
            self.a2v_tree.column(col, width=w, minwidth=30)
        self.a2v_tree.pack(fill="both", expand=True)
        self.a2v_tree.tag_configure("ok",  background="#0d2818", foreground=C["green"])
        self.a2v_tree.tag_configure("err", background="#2d0f0f", foreground=C["red"])
        self.a2v_tree.tag_configure("run", background="#0d1e35", foreground=C["accent"])

    # ── Audio2Video helpers ───────────────────────────────────────────────
    def _a2v_start(self):
        src = self.a2v_src_var.get().strip()
        if not src or not os.path.isdir(src):
            messagebox.showerror("错误", "请选择有效的音频文件夹"); return
        if not os.path.isfile(FFMPEG_EXE):
            messagebox.showerror("错误", f"找不到 FFmpeg:\n{FFMPEG_EXE}"); return

        exts = tuple("." + e.strip().lstrip(".") for e in
                     self.a2v_audio_ext_var.get().split(","))
        files = sorted(
            [os.path.join(src, f) for f in os.listdir(src)
             if os.path.isfile(os.path.join(src, f)) and f.lower().endswith(exts)],
            key=lambda p: os.path.basename(p))
        if not files:
            messagebox.showwarning("提示", "没有找到音频文件"); return

        out_dir = self.a2v_out_var.get().strip() or src
        os.makedirs(out_dir, exist_ok=True)

        # 初始化列表
        for row in self.a2v_tree.get_children(): self.a2v_tree.delete(row)
        for i, f in enumerate(files, 1):
            name = os.path.splitext(os.path.basename(f))[0]
            title = name[:self.a2v_chars_var.get()]
            out_name = name + ".mp4"
            self.a2v_tree.insert("", "end",
                                  iid=str(i),
                                  values=(i, os.path.basename(f), title, out_name, "等待"))

        self.a2v_stop_var[0] = False
        font_path = self._a2v_font_map.get(self.a2v_font_var.get(), "")
        threading.Thread(
            target=self._a2v_worker,
            args=(files, out_dir,
                  self.a2v_ratio_var.get(),
                  self.a2v_bg_var.get().strip(),
                  self.a2v_fontsize_var.get(),
                  self.a2v_chars_var.get(),
                  self.a2v_multiline_var.get(),
                  font_path,
                  self.a2v_fg_var.get().strip()),
            daemon=True).start()

    def _a2v_worker(self, files, out_dir, ratio, bg_color, fontsize, nchars,
                    multiline, font_path, fg_color="#ffffff"):
        import tempfile as _tf
        tmp = _tf.mkdtemp(prefix="a2v_")
        DIMS = {"9:16": (1080,1920), "16:9": (1920,1080), "1:1": (1080,1080)}
        W, H = DIMS.get(ratio, (1080,1920))

        # Pillow 检查
        try:
            from PIL import Image, ImageDraw, ImageFont
            PIL_OK = True
        except ImportError:
            PIL_OK = False

        def _hex_to_rgb(h):
            h = h.lstrip("#")
            if len(h) == 3: h = "".join(c*2 for c in h)
            return tuple(int(h[i:i+2],16) for i in (0,2,4))

        def _make_card(title: str, png_path: str):
            """Pillow 用 PNG 渲染标题图，支持中文、emoji、自动换行"""
            if not PIL_OK:
                raise RuntimeError("请先 pip install Pillow")
            bg_rgb = _hex_to_rgb(bg_color) if bg_color.startswith("#") else (26,26,46)
            fg_rgb = _hex_to_rgb(fg_color) if fg_color.startswith("#") else (255,255,255)

            img  = Image.new("RGB", (W, H), color=bg_rgb)
            draw = ImageDraw.Draw(img)

            # 加载字体
            try:
                font = ImageFont.truetype(font_path, fontsize) if font_path else ImageFont.load_default()
            except Exception:
                font = ImageFont.load_default()

            # 自动换行：按像素宽度进行装行
            max_px = int(W * 0.85)
            lines = []
            current = ""
            current_w = 0
            for ch in title:
                try:
                    bb = font.getbbox(ch)
                    cw = bb[2] - bb[0]
                except Exception:
                    cw = fontsize
                if multiline and (current_w + cw) > max_px and current:
                    lines.append(current)
                    current = ch; current_w = cw
                else:
                    current += ch; current_w += cw
            if current: lines.append(current)
            text_block = "\n".join(lines)

            # 居中绘制
            line_h = int(fontsize * 1.35)
            total_h = line_h * len(lines)
            y0 = (H - total_h) // 2
            for ln in lines:
                try:
                    bb = draw.textbbox((0,0), ln, font=font)
                    tw = bb[2] - bb[0]
                except Exception:
                    tw = fontsize * len(ln)
                x0 = (W - tw) // 2
                draw.text((x0, y0), ln, font=font, fill=fg_rgb)
                y0 += line_h

            img.save(png_path, "PNG")

        total = len(files); ok_count = err_count = 0

        def _upd_tree(iid, status, tag):
            self.after(0, lambda: self.a2v_tree.set(iid, "状态", status))
            self.after(0, lambda: self.a2v_tree.item(iid, tags=(tag,)))
            self.after(0, lambda: self.a2v_tree.see(iid))

        def _upd_prog(done):
            pct = int(done/total*100) if total else 0
            self.after(0, lambda: self.a2v_prog.config(value=pct))
            self.after(0, lambda: self.a2v_prog_lbl.config(
                text=f"{pct}%  [{done}/{total}]"))

        for i, audio in enumerate(files, 1):
            if self.a2v_stop_var[0]: break
            iid = str(i)
            _upd_tree(iid, "转换中…", "run")
            self.after(0, lambda lbl=self.a2v_status_lbl, n=os.path.basename(audio):
                lbl.config(text=f"正在处理: {n}"))

            try:
                stem = os.path.splitext(os.path.basename(audio))[0]
                title = stem[:nchars]

                # 1️⃣ 用 Pillow 生成标题图
                png = os.path.join(tmp, f"card_{i}.png")
                _make_card(title, png)

                # 2️⃣ FFmpeg: 静态图片 + 音频 → MP4
                out_mp4 = os.path.join(out_dir, stem + ".mp4")
                cmd = [
                    FFMPEG_EXE, "-y",
                    "-loop", "1", "-framerate", "25", "-i", png,   # 静态封面
                    "-i", audio,                                     # 音频
                    "-vf", f"scale={W}:{H}",
                    "-c:v", "libx264", "-preset", "fast", "-crf", "23",
                    "-pix_fmt", "yuv420p",
                    "-c:a", "aac", "-b:a", "192k",
                    "-shortest", "-movflags", "+faststart",
                    out_mp4
                ]
                r = subprocess.run(cmd, capture_output=True, text=True,
                                   encoding="utf-8", errors="replace",
                                   creationflags=_NO_WINDOW)
                if r.returncode != 0:
                    raise RuntimeError(r.stderr[-400:].strip())

                _upd_tree(iid, "✅ 完成", "ok")
                ok_count += 1
                self._log(f"  ✅ [{i}/{total}] {stem}.mp4")
            except Exception as e:
                _upd_tree(iid, f"❌ {str(e)[:40]}", "err")
                err_count += 1
                self._log(f"  ❌ [{i}/{total}] {os.path.basename(audio)}: {e}")

            _upd_prog(i)

        shutil.rmtree(tmp, ignore_errors=True)
        summary = f"完成 {ok_count} 个"
        if err_count: summary += f"，{err_count} 个失败"
        if self.a2v_stop_var[0]: summary = "已停止 — " + summary
        self.after(0, lambda: self.a2v_status_lbl.config(text=summary))
        self._log(f"\n🎵 音频转视频完毕: {summary}")


# ══════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
