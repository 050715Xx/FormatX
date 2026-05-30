#!/usr/bin/env python3
"""
FormatX — 自动化排版与可编辑公式解析器
依赖：pip install customtkinter python-docx tkinterdnd2
"""

import os
import sys
import threading
import traceback
import tempfile
import base64
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog


# ── 嵌入图标（1.ico 的 base64）─────────────────────────────────
_ICON_B64 = None  # 打包时由 build 脚本注入，或运行时从文件读取

# ── 从 format_conversion 导入核心函数 ──────────────────────────
try:
    from format_conversion import (convert_markdown_to_docx, reformat_docx,
                                     convert_text_to_docx, _detect_suspicious_vars)
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from format_conversion import (convert_markdown_to_docx, reformat_docx,
                                     convert_text_to_docx, _detect_suspicious_vars)

# ── 模板克隆适配器 ───────────────────────────────────────
try:
    from pipeline_adapter import reformat_docx_clone
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from pipeline_adapter import reformat_docx_clone


# ── FormatX 官方原木风色板 ──────────────────────────────────────
COLOR_BG_MAIN     = "#F9F7F2"  # 全局主背景 (米白)
COLOR_BG_PANEL    = "#F5F3ED"  # 侧边栏/日志框背景 (暖灰白)
COLOR_PRIMARY     = "#E5E3DC"  # 激活按钮/主控元素 (浅灰)
COLOR_HOVER       = "#EBE9E4"  # 悬停反馈色 (柔和灰)
COLOR_TEXT        = "#37352F"  # 全局正文/标题 (深炭灰)
COLOR_ACCENT      = "#A3B5A7"  # 进度条填充 (柔和鼠尾草)
COLOR_TRACK       = "#E8EBE7"  # 进度条底槽 (浅灰绿)

THEMES = {
    "moss": {
        "bg_main": "#2B312C", "bg_panel": "#232824", "primary": "#353C36",
        "hover": "#454E46", "text": "#E3DFD2", "text_muted": "#92978D",
        "accent": "#DF593A", "track": "#2B312C"
    },
    "warm": {
        "bg_main": "#F9F7F2", "bg_panel": "#F5F3ED", "primary": "#E5E3DC",
        "hover": "#EBE9E4", "text": "#37352F", "text_muted": "#8B7D72",
        "accent": "#A3B5A7", "track": "#E8EBE7"
    }
}
# ── 窗口配置 ────────────────────────────────────────────────────
APP_TITLE = "FormatX"
APP_SIZE = "1050x800"

# 加载 Montserrat Bold 字体（打包时从 sys._MEIPASS 读取）
def _get_font_path(filename):
    if getattr(sys, '_MEIPASS', ''):
        p = os.path.join(sys._MEIPASS, filename)
        if os.path.exists(p):
            return p
    local = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    if os.path.exists(local):
        return local
    return None

_MONTSERRAT_PATH = _get_font_path('Montserrat-Bold.ttf')

ctk.set_default_color_theme("blue")


# ── 日志面板 ──────────────────────────────────────────────────
class ConsoleLog(ctk.CTkTextbox):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("font", ("Consolas", 12))
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("wrap", "word")
        kwargs.setdefault("fg_color", COLOR_BG_PANEL)
        kwargs.setdefault("text_color", COLOR_TEXT)
        super().__init__(master, **kwargs)
        self.configure(state="disabled")

    def write(self, text):
        self.configure(state="normal")
        self.insert("end", text)
        self.see("end")
        self.configure(state="disabled")

    def clear(self):
        self.configure(state="normal")
        self.delete("1.0", "end")
        self.configure(state="disabled")


# ── 主窗口 ────────────────────────────────────────────────────
class App(ctk.CTk):
    def __init__(self):
        # 强制声明 Windows AppID
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                'yimin.formatx.v2.0')
        except Exception:
            pass
        super().__init__()
        self.withdraw()  # 先隐藏主窗口

        from splash import SplashScreen
        self._splash = SplashScreen()

        self.theme_key = "moss"
        try:
            import json
            cfg_path = os.path.join(os.path.expanduser("~"), ".formatx_ai_config.json")
            if os.path.exists(cfg_path):
                with open(cfg_path, "r", encoding="utf-8") as f:
                    self.theme_key = json.load(f).get("ui_theme", "moss")
        except Exception:
            pass

        self.title(APP_TITLE)
        self.geometry(APP_SIZE)
        self.minsize(950, 700)
        self.configure(fg_color=COLOR_BG_MAIN)

        # 字体
        self.font_normal = ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=14, weight="normal")
        self.font_bold   = ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=14, weight="bold")
        self.font_title  = ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=18, weight="normal")

        self._set_titlebar_color()
        self._set_icon()

        # ── 顶级 Grid 布局 (2行 x 2列) ──
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=0)
        self.grid_columnconfigure(1, weight=1)

        self._build_sidebar()
        self._build_main_workspace()
        self._build_bottom_console()
        self._setup_dnd()

        self._current_page = "page_basic"
        self.select_page("page_basic")

        self.apply_ui_theme(self.theme_key)

        # 启动页淡出 + 显示主窗口
        self.after(800, lambda: (self._splash.fade_out(), self.deiconify()))

    # ── 主题变色引擎 ─────────────────────────────────────────────
    def apply_ui_theme(self, theme_key):
        self.theme_key = theme_key
        t = THEMES[theme_key]

        if theme_key == "moss":
            ctk.set_appearance_mode("dark")
        else:
            ctk.set_appearance_mode("light")
        self.update_idletasks()

        self.configure(fg_color=t["bg_main"])

        def _update_titlebar_dwm():
            try:
                import ctypes
                bgr_hex = 0x00F7FBFD if theme_key == "warm" else 0x002C312B
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 35, ctypes.byref(ctypes.c_int(bgr_hex)),
                    ctypes.sizeof(ctypes.c_int))
            except Exception:
                pass
        self.after(200, _update_titlebar_dwm)

        def _update_widget(parent):
            for child in parent.winfo_children():
                if isinstance(child, ctk.CTkToplevel):
                    continue
                try:
                    w_type = child.__class__.__name__
                    try:
                        raw_fg = child.cget("fg_color")
                        fg_str = str(raw_fg).lower()
                    except Exception:
                        fg_str = ""

                    if child == getattr(self, "bottom_frame", None):
                        child.configure(fg_color=t["bg_main"])
                    elif child == getattr(self, "log", None):
                        child.configure(fg_color=t["bg_panel"], text_color=t["text"],
                                        border_color=t["primary"])
                    elif "CTkLabel" in w_type:
                        if child.cget("text") in ["满怀希望，\n就会所向披靡 ✨",
                                                  "自动化排版与\n可编辑公式解析器"]:
                            child.configure(text_color=t["text_muted"])
                        else:
                            child.configure(text_color=t["text"])
                    elif child in getattr(self, "nav_buttons", {}).values():
                        if "transparent" not in fg_str:
                            child.configure(fg_color=t["primary"])
                        child.configure(text_color=t["text"], hover_color=t["hover"])
                    elif "CTkButton" in w_type:
                        if "transparent" not in fg_str:
                            if child.cget("text") == "一键开始排版":
                                child.configure(fg_color=t["primary"], text_color=t["text"],
                                                hover_color=t["hover"])
                            elif child.cget("fg_color") != "transparent":
                                child.configure(fg_color=t["primary"])
                        child.configure(text_color=t["text"], hover_color=t["hover"])
                    elif "CTkScrollableFrame" in w_type:
                        child.configure(fg_color=t["bg_main"],
                                        scrollbar_button_color=t["primary"])
                    elif "CTkFrame" in w_type:
                        if child == getattr(self, "sidebar_frame", None):
                            child.configure(fg_color=t["bg_panel"])
                        elif "transparent" not in fg_str:
                            child.configure(fg_color=t["bg_main"])
                    elif "CTkSegmentedButton" in w_type:
                        child.configure(fg_color=t["bg_main"],
                                        selected_color=t["primary"],
                                        selected_hover_color=t["hover"],
                                        unselected_color=t["bg_main"],
                                        unselected_hover_color=t["primary"],
                                        text_color=t["text"])
                    elif "CTkEntry" in w_type:
                        child.configure(fg_color=t["bg_panel"], text_color=t["text"],
                                        border_color=t["primary"])
                    elif "CTkOptionMenu" in w_type:
                        child.configure(fg_color=t["primary"], text_color=t["text"],
                                        button_color=t["primary"], button_hover_color=t["hover"],
                                        dropdown_fg_color=t["bg_panel"])
                    elif "CTkTextbox" in w_type:
                        child.configure(fg_color=t["bg_panel"], text_color=t["text"],
                                        border_color=t["primary"])
                    elif "CTkProgressBar" in w_type:
                        child.configure(progress_color=t["accent"], fg_color=t["track"])
                    elif "CTkCheckBox" in w_type or "CTkRadioButton" in w_type:
                        child.configure(fg_color=t["accent"], hover_color=t["hover"],
                                        text_color=t["text"])
                except Exception as e:
                    print(f"[Theme Warn] {w_type} failed: {e}")
                if child.winfo_children():
                    _update_widget(child)

        _update_widget(self)
        if hasattr(self, "btn_inspire") and self.btn_inspire.winfo_exists():
            self.btn_inspire.configure(text_color="#8B7D72",
                                       border_color=t["primary"])
        if hasattr(self, "_custom_widgets"):
            pw = self._custom_widgets["page"]
            pw.configure(fg_color=t["bg_main"])
            try:
                pw._draw()
            except Exception:
                pass
            self._custom_widgets["switch_card"].configure(fg_color=t["bg_panel"],
                border_color=t["primary"])
            self._custom_widgets["indent_card"].configure(fg_color=t["bg_panel"],
                border_color=t["primary"])
            self._custom_widgets["line_card"].configure(fg_color=t["bg_panel"],
                border_color=t["primary"])
            self._custom_widgets["seg"].configure(
                fg_color=t["bg_main"], selected_color=t["primary"],
                selected_hover_color=t["hover"], unselected_color=t["bg_main"],
                unselected_hover_color=t["primary"], text_color=t["text"])
        if hasattr(self, "_ai_chat_apply_theme"):
            chat_t = {"text": t["text"], "primary": t["primary"],
                      "bg_panel": t["bg_panel"], "hover": t["hover"],
                      "bg_main": t["bg_main"], "text_muted": t["text_muted"]}
            self._ai_chat_apply_theme(chat_t)
        if hasattr(self, "_sync_custom_ui_hook"):
            self._sync_custom_ui_hook()
        self.update_idletasks()

    # ── 标题栏颜色 ─────────────────────────────────────────────
    def _set_titlebar_color(self):
        try:
            import ctypes
            self.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 35,
                ctypes.byref(ctypes.c_int(0x00F2F7F9)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

    # ── 窗口图标 ───────────────────────────────────────────────
    def _set_icon(self):
        if getattr(sys, '_MEIPASS', ''):
            p = os.path.join(sys._MEIPASS, '3.ico')
            if os.path.exists(p):
                try:
                    self.iconbitmap(p)
                    return
                except Exception:
                    pass
        try:
            self.iconbitmap(sys.executable)
            return
        except Exception:
            pass
        try:
            tf = tempfile.NamedTemporaryFile(suffix='.ico', delete=False)
            tf.write(_make_ico_data())
            tf.close()
            self.iconbitmap(tf.name)
        except Exception:
            pass

    # ── 左侧导航栏 ──────────────────────────────────────────────
    def _build_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(
            self, width=200, corner_radius=0, fg_color=COLOR_BG_PANEL)
        self.sidebar_frame.grid(row=0, column=0, rowspan=1, sticky="nsew")
        self.sidebar_frame.grid_propagate(False)

        # Montserrat 标题
        title_family = "Montserrat"
        try:
            import ctypes
            m_path = _get_font_path('Montserrat-Bold.ttf')
            if m_path:
                ctypes.windll.gdi32.AddFontResourceExW(m_path, 0x10, 0)
        except Exception:
            pass
        try:
            import tkinter.font as tkf
            tkf.Font(family="Montserrat", size=1)
        except Exception:
            title_family = "Segoe UI"

        ctk.CTkLabel(
            self.sidebar_frame, text="FormatX",
            font=ctk.CTkFont(family=title_family, size=24, weight="normal"),
            text_color=COLOR_TEXT
        ).pack(pady=(30, 20))

        ctk.CTkLabel(
            self.sidebar_frame, text="自动化排版与\n可编辑公式解析器",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=16, slant="italic"),
            text_color="#9B8D82"
        ).pack(pady=(0, 20))

        # 新建文档按钮
        self.btn_new_doc = ctk.CTkButton(
            self.sidebar_frame, text="✨ 新建文档",
            command=self._new_document,
            font=self.font_normal, fg_color=COLOR_PRIMARY,
            text_color=COLOR_TEXT, hover_color=COLOR_HOVER,
            anchor="w", height=40
        )
        self.btn_new_doc.pack(pady=(14, 10), padx=15, fill="x")

        ctk.CTkFrame(self.sidebar_frame, height=2, fg_color=COLOR_HOVER).pack(fill="x", padx=15, pady=(0, 15))

        self.nav_buttons = {}

        self.nav_buttons["page_basic"] = ctk.CTkButton(
            self.sidebar_frame, text="🏠 基础排版",
            fg_color="transparent", text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, anchor="w", font=self.font_normal,
            command=lambda: self.select_page("page_basic")
        )
        self.nav_buttons["page_basic"].pack(pady=5, padx=10, fill="x")

        self.nav_buttons["page_text"] = ctk.CTkButton(
            self.sidebar_frame, text="✏️ AI 创作舱",
            fg_color="transparent", text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, anchor="w", font=self.font_normal,
            command=lambda: self.select_page("page_text")
        )
        self.nav_buttons["page_text"].pack(pady=5, padx=10, fill="x")

        self.nav_buttons["page_template"] = ctk.CTkButton(
            self.sidebar_frame, text="🎨 学校模板克隆",
            fg_color="transparent", text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, anchor="w", font=self.font_normal,
            command=lambda: self.select_page("page_template")
        )
        self.nav_buttons["page_template"].pack(pady=5, padx=10, fill="x")

        self.nav_buttons["page_custom"] = ctk.CTkButton(
            self.sidebar_frame, text="🎛️ 自定义模式",
            fg_color="transparent", text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, anchor="w", font=self.font_normal,
            command=lambda: self.select_page("page_custom")
        )
        self.nav_buttons["page_custom"].pack(pady=5, padx=10, fill="x")

        # ⚙️ 设置入口（侧边栏右上角，仅图标）
        ctk.CTkButton(
            self.sidebar_frame, text="⚙️",
            command=self._show_settings,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            fg_color="transparent", text_color="#8B7D72",
            hover_color=COLOR_HOVER, width=28, height=22
        ).place(relx=0.97, rely=0.02, anchor="ne")

        # 暖茶灰温馨寄语
        ctk.CTkLabel(
            self.sidebar_frame,
            text="满怀希望，\n就会所向披靡 ✨",
            font=ctk.CTkFont(family="Microsoft YaHei", size=18, slant="italic"),
            text_color="#B0A8A0", justify="left"
        ).pack(side="bottom", pady=20)

    # ── 右侧纸牌屋 ──────────────────────────────────────────────
    def _build_main_workspace(self):
        self.workspace_container = ctk.CTkFrame(
            self, fg_color=COLOR_BG_MAIN, corner_radius=0)
        self.workspace_container.grid(
            row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.workspace_container.grid_rowconfigure(0, weight=1)
        self.workspace_container.grid_columnconfigure(0, weight=1)

        self.pages = {}

        # === Page 1: 基础排版 ===
        page_basic = ctk.CTkFrame(self.workspace_container, fg_color="transparent")
        page_basic.grid(row=0, column=0, sticky="nsew")
        page_basic.grid_rowconfigure(0, weight=1)
        page_basic.grid_columnconfigure(0, weight=1)

        basic_content = ctk.CTkFrame(page_basic, fg_color="transparent")
        basic_content.pack(fill="both", expand=True)

        # 文件选择
        ctk.CTkLabel(basic_content, text="📂 选择待排版的文档 (.md / .txt / .docx)",
                     font=self.font_bold, text_color=COLOR_TEXT).pack(anchor="w", pady=(0, 8))

        file_row = ctk.CTkFrame(basic_content, fg_color="transparent")
        file_row.pack(fill="x", pady=(0, 10))

        self.file_entry = ctk.CTkEntry(
            file_row, placeholder_text="请选择文件 或 拖拽文件到此处...",
            font=("Microsoft YaHei", 11),
            fg_color="transparent", text_color=COLOR_TEXT, border_color=COLOR_PRIMARY)
        self.file_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.file_entry.configure(state="readonly")
        self.file_entry.bind("<Double-Button-1>", lambda e: self._browse_input())

        ctk.CTkButton(
            file_row, text="浏览", width=70,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, font=("Microsoft YaHei", 13),
            command=self._browse_input
        ).pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            file_row, text="新建空白.md", width=100,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, font=("Microsoft YaHei", 13),
            command=self._create_blank_md
        ).pack(side="left")

        # 输出路径
        ctk.CTkLabel(basic_content, text="📁 输出路径",
                     font=self.font_bold, text_color=COLOR_TEXT).pack(anchor="w", pady=(15, 8))

        out_label_row = ctk.CTkFrame(basic_content, fg_color="transparent")
        out_label_row.pack(fill="x", pady=(0, 4))

        self.same_dir_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            out_label_row, text="同目录", variable=self.same_dir_var,
            command=self._toggle_output,
            font=("Microsoft YaHei", 11),
            fg_color=COLOR_ACCENT, hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT, checkbox_width=20, checkbox_height=20
        ).pack(side="right")

        out_row = ctk.CTkFrame(basic_content, fg_color="transparent")
        out_row.pack(fill="x")

        self.out_entry = ctk.CTkEntry(
            out_row, placeholder_text="自动生成（同目录）...",
            font=("Microsoft YaHei", 11), state="disabled",
            fg_color="transparent", text_color=COLOR_TEXT, border_color=COLOR_PRIMARY)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.out_btn = ctk.CTkButton(
            out_row, text="浏览", width=70,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER, state="disabled", command=self._browse_output
        )
        self.out_btn.pack(side="right")

        # 输出模式
        self.output_mode = ctk.StringVar(value="export")
        radio_row = ctk.CTkFrame(basic_content, fg_color="transparent")
        radio_row.pack(fill="x", pady=(20, 0))
        ctk.CTkLabel(radio_row, text="输出：",
                     font=("Microsoft YaHei", 11),
                     text_color=COLOR_TEXT).pack(side="left", padx=(0, 8))
        ctk.CTkRadioButton(radio_row, text="导出 .docx", variable=self.output_mode,
                           value="export", font=("Microsoft YaHei", 11),
                           fg_color=COLOR_ACCENT, hover_color=COLOR_HOVER,
                           text_color=COLOR_TEXT
                           ).pack(side="left", padx=(0, 12))
        ctk.CTkRadioButton(radio_row, text="直接显示结果", variable=self.output_mode,
                           value="preview", font=("Microsoft YaHei", 11),
                           fg_color=COLOR_ACCENT, hover_color=COLOR_HOVER,
                           text_color=COLOR_TEXT
                           ).pack(side="left")

        self.pages["page_basic"] = page_basic

        self.btn_inspire = ctk.CTkButton(
            page_basic, text="✨ 寄语", width=70, height=28, corner_radius=14,
            fg_color="transparent", border_width=1, border_color=COLOR_PRIMARY,
            text_color="#8B7D72", hover_color=COLOR_HOVER,
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=12),
            command=self._show_inspiration_modal)
        self.btn_inspire.pack(side="bottom", anchor="e", padx=30, pady=(10, 10))

        # === Page 2: AI 创作舱 (聊天流) ===
        from ai_assist.ai_chat_panel import build_ai_chat_panel
        t = THEMES.get(self.theme_key, THEMES["warm"])
        chat_colors = {
            "text": COLOR_TEXT, "primary": COLOR_PRIMARY,
            "bg_panel": COLOR_BG_PANEL, "hover": COLOR_HOVER,
            "bg_main": COLOR_BG_MAIN,
            "text_muted": t["text_muted"],
        }
        page_ai, self._ai_chat_apply_theme = build_ai_chat_panel(
            self.workspace_container, chat_colors, self._log)
        page_ai.grid(row=0, column=0, sticky="nsew")
        page_ai.grid_rowconfigure(1, weight=1)
        page_ai.grid_columnconfigure(0, weight=1)

        self.pages["page_text"] = page_ai

        # === Page 3: 学校模板克隆 ===
        page_template = ctk.CTkFrame(self.workspace_container, fg_color="transparent")
        page_template.grid(row=0, column=0, sticky="nsew")
        page_template.grid_rowconfigure(0, weight=1)
        page_template.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(page_template, text="🎨 学校模板克隆引擎 (可选)",
                     font=self.font_title, text_color=COLOR_TEXT).pack(anchor="w", pady=(0, 20))

        # Layer 1: 封面与目录物理拼接
        cover_card = ctk.CTkFrame(page_template, fg_color=COLOR_BG_PANEL,
                                  border_width=1, border_color=COLOR_PRIMARY,
                                  corner_radius=10)
        cover_card.pack(fill="x", padx=25, pady=(10, 15), ipady=8)
        ctk.CTkLabel(
            cover_card, text="提取封面与目录 (物理拼接)",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            text_color=COLOR_TEXT
        ).pack(side="left", padx=20, pady=12)
        self.check_cover = ctk.CTkCheckBox(
            cover_card, text="",
            fg_color=COLOR_ACCENT, hover_color=COLOR_HOVER,
            text_color=COLOR_TEXT, font=self.font_bold
        )
        self.check_cover.pack(side="right", padx=20, pady=12)

        cover_frame = ctk.CTkFrame(page_template, fg_color="transparent")
        cover_frame.pack(fill="x", pady=(0, 20), padx=25)
        self.entry_cover = ctk.CTkEntry(
            cover_frame, placeholder_text="选择封面模板.docx...",
            fg_color="transparent", text_color=COLOR_TEXT, border_color=COLOR_PRIMARY)
        self.entry_cover.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._cover_fullpath = ""  # 存储封面完整路径
        ctk.CTkButton(
            cover_frame, text="浏览", width=80,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT, hover_color=COLOR_HOVER,
            command=lambda: self._browse_cover_file()
        ).pack(side="right")

        # Layer 2: 智能克隆排版 DNA
        from core.config_manager import load_full_config, update_config
        from core.state.constants import TypographyMode
        from core.state.mode_manager import mode_manager

        clone_switch_card = ctk.CTkFrame(page_template, fg_color=COLOR_BG_PANEL,
                                         border_width=1, border_color=COLOR_PRIMARY,
                                         corner_radius=10)
        clone_switch_card.pack(fill="x", padx=25, pady=(10, 15), ipady=8)
        ctk.CTkLabel(
            clone_switch_card, text="启用正文模板克隆 (强行接管排版)",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            text_color=COLOR_TEXT
        ).pack(side="left", padx=20, pady=12)

        self.switch_clone = ctk.CTkSwitch(
            clone_switch_card, text="", progress_color="#2D5A27",
            command=lambda: _on_user_click_clone_switch())
        self.switch_clone.pack(side="right", padx=20, pady=12)

        style_frame = ctk.CTkFrame(page_template, fg_color="transparent")
        style_frame.pack(fill="x", pady=0, padx=25)
        self.entry_style = ctk.CTkEntry(
            style_frame, placeholder_text="选择标准正文模板.docx...",
            fg_color="transparent", text_color=COLOR_TEXT, border_color=COLOR_PRIMARY)
        self.entry_style.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self._clone_fullpath = ""  # 存储克隆模板完整路径

        saved_path = load_full_config().get("clone_target_path", "")
        if saved_path:
            self._clone_fullpath = saved_path
            self.entry_style.insert(0, os.path.basename(saved_path))

        # 状态反馈标签
        self._clone_status = ctk.CTkLabel(
            page_template, text="",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color="#8B7D72"
        )
        self._clone_status.pack(anchor="w", padx=25, pady=(0, 10))

        def _browse_and_save_clone():
            self._browse_file(self.entry_style)
            path = self.entry_style.get().strip()
            if not path:
                return
            # 存储完整路径 + 显示文件名
            self._clone_fullpath = path
            self.entry_style.delete(0, "end")
            self.entry_style.insert(0, os.path.basename(path))
            update_config("clone_target_path", path)
            if self.switch_clone.get():
                mode_manager.set_mode(TypographyMode.CLONE)
            # 检测模板有效性
            try:
                from docx import Document
                Document(path)
                self._clone_status.configure(
                    text="✅ 模板文件验证成功，排版时将自动提取样式",
                    text_color="#2D5A27")
            except Exception:
                self._clone_status.configure(
                    text="❌ 模板文件无效，请确保是完整的 .docx 文件",
                    text_color="#C0392B")

        ctk.CTkButton(
            style_frame, text="浏览", width=80,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT, hover_color=COLOR_HOVER,
            command=_browse_and_save_clone
        ).pack(side="right")

        def _on_global_mode_changed_clone(current_mode):
            if current_mode == TypographyMode.CLONE:
                if not self.switch_clone.get():
                    self.switch_clone.select()
                self._clone_mutex_hint.configure(
                    text="⚡ 当前已激活模板克隆引擎，排版格式由模板 DOCX 全权接管。\n"
                         "【自定义模式】中的缩进/行距设置已被休眠。",
                    text_color="#2D5A27")
            else:
                if self.switch_clone.get():
                    self.switch_clone.deselect()
                self._clone_mutex_hint.configure(
                    text="提示：开启克隆将提取文档底层 OOXML 基因（字号、字符缩进、多倍行距等）。\n"
                         "此操作与【自定义模式】互斥。",
                    text_color="#8B7D72")

        def _on_user_click_clone_switch():
            if self.switch_clone.get():
                mode_manager.set_mode(TypographyMode.CLONE)
                if self._clone_fullpath:
                    update_config("clone_target_path", self._clone_fullpath)
            else:
                mode_manager.set_mode(TypographyMode.PRESET)

        def _cleanup_clone_sub(event):
            if event.widget == page_template:
                mode_manager.unsubscribe(_on_global_mode_changed_clone)
        page_template.bind("<Destroy>", _cleanup_clone_sub, add="+")

        # 互斥提示（动态）
        self._clone_mutex_hint = ctk.CTkLabel(
            page_template,
            text="提示：开启克隆将提取文档底层 OOXML 基因（字号、字符缩进、多倍行距等）。\n"
                 "此操作与【自定义模式】互斥。",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color="#8B7D72", justify="left"
        )
        self._clone_mutex_hint.pack(anchor="w", pady=20, padx=25)

        mode_manager.subscribe(_on_global_mode_changed_clone)
        _on_global_mode_changed_clone(mode_manager.get_mode())

        self.pages["page_template"] = page_template

        # === Page 4: 自定义排版微调 ===
        page_custom = ctk.CTkFrame(self.workspace_container, fg_color="transparent")
        page_custom.grid(row=0, column=0, sticky="nsew")
        page_custom.grid_rowconfigure(0, weight=1)
        page_custom.grid_columnconfigure(0, weight=1)

        t = THEMES.get(self.theme_key, THEMES["warm"])

        custom_bg = ctk.CTkFrame(page_custom, fg_color=t["bg_main"], corner_radius=0)
        custom_bg.place(relx=0, rely=0, relwidth=1, relheight=1)

        custom_scroll = ctk.CTkScrollableFrame(
            custom_bg, fg_color="transparent",
            scrollbar_button_color=t["primary"],
            scrollbar_button_hover_color=t["hover"])
        custom_scroll.pack(fill="both", expand=True, padx=60, pady=30)

        ctk.CTkLabel(
            custom_scroll, text="自定义排版微调",
            font=ctk.CTkFont(family="Microsoft YaHei", size=22, weight="bold"),
            text_color=t["text"]
        ).pack(anchor="w", pady=(0, 10))

        ctk.CTkLabel(
            custom_scroll,
            text="开启后，缩进与行距将覆盖预设方案中的对应数值。",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            text_color=t["text_muted"], justify="left"
        ).pack(anchor="w", pady=(0, 20))

        # 总控开关
        switch_card = ctk.CTkFrame(custom_scroll, fg_color=t["bg_panel"],
                                    border_width=1, border_color=t["primary"],
                                    corner_radius=10)
        switch_card.pack(fill="x", pady=(0, 15), ipady=8)
        ctk.CTkLabel(
            switch_card, text="启用自定义排版参数",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, weight="bold"),
            text_color=t["text"]
        ).pack(side="left", padx=20, pady=12)
        master_switch = ctk.CTkSwitch(
            switch_card, text="", progress_color="#2D5A27",
            command=lambda: _on_custom_toggle())
        master_switch.pack(side="right", padx=20, pady=12)

        # 卡片容器 (被蒙版包裹)
        cards_wrapper = ctk.CTkFrame(custom_scroll, fg_color="transparent")
        cards_wrapper.pack(fill="x")

        # 缩进卡片 — 单位用 placeholder
        indent_card = ctk.CTkFrame(cards_wrapper, fg_color=t["bg_panel"],
                                    border_width=1, border_color=t["primary"],
                                    corner_radius=10)
        indent_card.pack(fill="x", pady=(0, 12), ipady=5)
        ctk.CTkLabel(
            indent_card, text="首行缩进",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            text_color=t["text"]
        ).pack(anchor="w", padx=20, pady=(12, 6))
        indent_row = ctk.CTkFrame(indent_card, fg_color="transparent")
        indent_row.pack(fill="x", padx=20, pady=(0, 12))
        ctk.CTkLabel(
            indent_row, text="缩进值", text_color=t["text"],
            font=("Microsoft YaHei", 12)
        ).pack(side="left")
        indent_entry = ctk.CTkEntry(
            indent_row, width=100, placeholder_text="0.85",
            fg_color=t["bg_main"], text_color=t["text"],
            border_color=t["primary"])
        indent_entry.pack(side="right")
        ctk.CTkLabel(
            indent_row, text="cm", text_color=t["text_muted"],
            font=("Microsoft YaHei", 11)
        ).pack(side="right", padx=(0, 8))

        # 行距卡片 — 动态标签
        line_card = ctk.CTkFrame(cards_wrapper, fg_color=t["bg_panel"],
                                 border_width=1, border_color=t["primary"],
                                 corner_radius=10)
        line_card.pack(fill="x", pady=(0, 12), ipady=5)
        ctk.CTkLabel(
            line_card, text="行距设置",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            text_color=t["text"]
        ).pack(anchor="w", padx=20, pady=(12, 6))
        line_mode_row = ctk.CTkFrame(line_card, fg_color="transparent")
        line_mode_row.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(
            line_mode_row, text="行距模式", text_color=t["text"],
            font=("Microsoft YaHei", 12)
        ).pack(side="left")
        line_mode_var = ctk.StringVar(value="exact")
        line_mode_seg = ctk.CTkSegmentedButton(
            line_mode_row,
            values=["固定值", "多倍", "单倍"],
            variable=line_mode_var,
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            fg_color=t["bg_main"],
            selected_color=t["primary"],
            selected_hover_color=t["hover"],
            unselected_color=t["bg_main"],
            unselected_hover_color=t["primary"],
            text_color=t["text"],
            corner_radius=6,
            command=lambda v: _on_mode_changed(
                {"固定值": "exact", "多倍": "multiple", "单倍": "single"}.get(v, "exact")))
        line_mode_seg.pack(side="left", fill="x", expand=True, padx=(10, 0))
        line_val_row = ctk.CTkFrame(line_card, fg_color="transparent")
        line_val_row.pack(fill="x", padx=20, pady=(4, 12))
        line_val_lbl = ctk.CTkLabel(
            line_val_row, text="行距数值 (pt)", text_color=t["text"],
            font=("Microsoft YaHei", 12))
        line_val_lbl.pack(side="left")
        line_val_entry = ctk.CTkEntry(
            line_val_row, width=100, placeholder_text="20.0",
            fg_color=t["bg_main"], text_color=t["text"],
            border_color=t["primary"])
        line_val_entry.pack(side="right")

        def _on_mode_changed(mode):
            t = THEMES.get(self.theme_key, THEMES["warm"])
            if mode == "single":
                line_val_lbl.configure(text="行距数值", text_color=t["text_muted"])
                line_val_entry.configure(state="normal")
                line_val_entry.delete(0, "end")
                line_val_entry.insert(0, "N/A")
                line_val_entry.configure(state="disabled")
            elif mode == "multiple":
                line_val_lbl.configure(text="行距倍数", text_color=t["text"])
                cur = line_val_entry.get()
                try:
                    v = float(cur) if cur else 0
                except ValueError:
                    v = 0
                if cur == "N/A" or v <= 0 or v > 10:
                    line_val_entry.configure(state="normal")
                    line_val_entry.delete(0, "end")
                    line_val_entry.insert(0, "1.5")
                else:
                    line_val_entry.configure(state="normal")
            else:
                line_val_lbl.configure(text="行距数值 (pt)", text_color=t["text"])
                cur = line_val_entry.get()
                try:
                    v = float(cur) if cur else 0
                except ValueError:
                    v = 0
                if cur == "N/A" or v <= 0 or v < 10:
                    line_val_entry.configure(state="normal")
                    line_val_entry.delete(0, "end")
                    line_val_entry.insert(0, "20.0")
                else:
                    line_val_entry.configure(state="normal")

        # 状态标签
        status_lbl = ctk.CTkLabel(
            custom_scroll, text="",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color=t["accent"])
        status_lbl.pack(anchor="w", pady=(8, 5))

        # 底部按钮行 (全宽保存 + 重置)
        btn_row = ctk.CTkFrame(custom_scroll, fg_color="transparent")
        btn_row.pack(fill="x", pady=(5, 0))
        save_btn = ctk.CTkButton(
            btn_row, text="保存自定义参数",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            height=36, fg_color="#6B8E76", hover_color="#5A7A63",
            text_color="#FDFBF7",
            command=lambda: _save_custom())
        save_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))
        reset_btn = ctk.CTkButton(
            btn_row, text="恢复默认值",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            height=36, fg_color=t["primary"], text_color=t["text"],
            hover_color=t["hover"],
            command=lambda: _reset_defaults())
        reset_btn.pack(side="left", padx=(5, 0))

        # 禁用蒙版标签
        lock_overlay = ctk.CTkLabel(
            cards_wrapper, text="",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"),
            text_color=t["text_muted"])

        # ── 同步 UI 状态 ──
        def _sync_custom_ui():
            t = THEMES.get(self.theme_key, THEMES["warm"])
            from core.config_manager import load_full_config
            cfg = load_full_config()
            is_custom = cfg.get("typography_mode") == "custom"
            if is_custom and not master_switch.get():
                master_switch.select()
            elif not is_custom and master_switch.get():
                master_switch.deselect()
            st = "normal" if is_custom else "disabled"
            tc = t["text"] if is_custom else t["text_muted"]
            indent_entry.configure(state=st, text_color=tc)
            line_mode_seg.configure(state=st)
            line_val_entry.configure(state=st, text_color=tc)
            save_btn.configure(state=st)
            reset_btn.configure(state=st)
            if cfg.get("typography_mode") == "clone":
                lock_overlay.configure(
                    text="当前已开启【模板克隆】，自定义参数已休眠。\n如需微调，请前往【学校模板克隆】关闭接管。",
                    text_color="#2D5A27")
                lock_overlay.place(relx=0.5, rely=0.5, anchor="center")
            else:
                lock_overlay.configure(text="")
                lock_overlay.place_forget()
                indent_card.configure(fg_color=t["bg_panel"])
                line_card.configure(fg_color=t["bg_panel"])
            if is_custom:
                custom = cfg.get("custom_settings", {})
                if indent_entry.get() == "":
                    indent_entry.insert(0, str(custom.get("indent_cm", "0.85")))
                _MODE_REV = {"exact": "固定值", "multiple": "多倍", "single": "单倍"}
                saved_mode = custom.get("line_mode", "exact")
                line_mode_var.set(_MODE_REV.get(saved_mode, "固定值"))
                _on_mode_changed(saved_mode)
                if line_val_entry.get() == "":
                    line_val_entry.insert(0, str(custom.get("line_pt", "20.0")))

        from core.state.mode_manager import mode_manager
        from core.state.constants import TypographyMode

        def _on_mode_changed_global(current_mode):
            if current_mode == TypographyMode.CUSTOM:
                if not master_switch.get():
                    master_switch.select()
                _sync_custom_ui()
            else:
                if master_switch.get():
                    master_switch.deselect()
                indent_entry.configure(state="disabled")
                line_mode_seg.configure(state="disabled")
                line_val_entry.configure(state="disabled")
                save_btn.configure(state="disabled")
                reset_btn.configure(state="disabled")
                _sync_custom_ui()

        def _on_custom_toggle():
            if master_switch.get():
                mode_manager.set_mode(TypographyMode.CUSTOM)
            else:
                mode_manager.set_mode(TypographyMode.PRESET)

        mode_manager.subscribe(_on_mode_changed_global)

        def _cleanup_ui(event):
            if event.widget == page_custom:
                mode_manager.unsubscribe(_on_mode_changed_global)
        page_custom.bind("<Destroy>", _cleanup_ui, add="+")

        def _save_custom():
            if not master_switch.get():
                return
            from core.config_manager import update_config
            update_config("custom_settings", {
                "indent_cm": float(indent_entry.get() or "0.85"),
                "line_mode": {"固定值": "exact", "多倍": "multiple", "单倍": "single"}.get(line_mode_var.get(), "exact"),
                "line_pt": 1.0 if line_val_entry.get() == "N/A" else float(line_val_entry.get() or "20.0"),
            })
            status_lbl.configure(text="自定义参数已保存")

        def _reset_defaults():
            indent_entry.delete(0, "end")
            indent_entry.insert(0, "0.85")
            line_mode_var.set("固定值")
            _on_mode_changed("exact")
            line_val_entry.delete(0, "end")
            line_val_entry.insert(0, "20.0")
            from core.config_manager import update_config
            update_config("custom_settings", {
                "indent_cm": 0.85, "line_mode": "exact", "line_pt": 20.0,
            })
            status_lbl.configure(text="已恢复默认参数")

        _sync_custom_ui()

        self.pages["page_custom"] = page_custom
        self._custom_widgets = {
            "page": custom_bg, "switch_card": switch_card,
            "indent_card": indent_card, "line_card": line_card,
            "seg": line_mode_seg,
        }
        self._sync_custom_ui_hook = _sync_custom_ui

    # ── 底部全局控制台 ──────────────────────────────────────────
    def _build_bottom_console(self):
        self.bottom_frame = ctk.CTkFrame(
            self, corner_radius=0, fg_color=COLOR_BG_MAIN)
        self.bottom_frame.grid(
            row=1, column=0, columnspan=2, sticky="nsew", padx=20, pady=(0, 20))

        btn_row = ctk.CTkFrame(self.bottom_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 10))

        self.go_btn = ctk.CTkButton(
            btn_row, text="一键开始排版",
            font=("Microsoft YaHei", 13, "bold"), height=34,
            fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER,
            command=self._start_conversion
        )
        self.go_btn.pack(side="left", fill="x", expand=True)

        self.report_btn = ctk.CTkButton(
            btn_row, text="📄 诊断报告",
            font=("Microsoft YaHei", 12, "bold"),
            height=34, width=110, state="disabled",
            fg_color=COLOR_BG_PANEL, text_color=COLOR_TEXT,
            hover_color=COLOR_HOVER,
            command=self._show_report_dialog
        )
        self.report_btn.pack(side="right", padx=(10, 0))

        self.progress = ctk.CTkProgressBar(
            self.bottom_frame, height=6,
            fg_color=COLOR_TRACK, progress_color=COLOR_ACCENT
        )
        self.progress.pack(fill="x", pady=(0, 8))
        self.progress.set(0)

        self.log = ConsoleLog(self.bottom_frame, height=80)
        self.log.pack(fill="x")

        self._log('FormatX 已就绪')
        self._log('支持 .md / .txt / .docx')

    # ── 页面切换 ────────────────────────────────────────────────
    def select_page(self, page_name):
        self._current_page = page_name
        primary = THEMES.get(self.theme_key, THEMES["warm"])["primary"]
        for name, btn in self.nav_buttons.items():
            if name == page_name:
                btn.configure(fg_color=primary, font=self.font_bold)
            else:
                btn.configure(fg_color="transparent", font=self.font_normal)
        self.pages[page_name].tkraise()
        self._update_go_btn()
        page = self.pages.get(page_name)
        if page and hasattr(page, '_sync_format_state'):
            page._sync_format_state()

    def _new_document(self):
        """打开 Microsoft Word 新建空白文档（带加载提示）"""
        from splash import SplashScreen
        splash = SplashScreen()
        self.after(600, lambda: self._launch_word(splash))

    def _launch_word(self, splash):
        try:
            os.system("start winword")
            self._log("已打开 Microsoft Word 新建文档")
        except Exception as e:
            self._log(f"打开 Word 失败: {str(e)}")
        splash.fade_out()

    # ── 拖拽支持 ────────────────────────────────────────────────
    def _setup_dnd(self):
        try:
            from tkinterdnd2 import DND_FILES
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop)
            self.dnd_bind('<<DropPosition>>', lambda e: None)
            self._has_dnd = True
        except Exception:
            self._has_dnd = False

    def _on_drop(self, event):
        path = event.data.strip().strip('{}')
        if os.path.isfile(path):
            self._set_input_path(path)

    # ── 文件浏览 ────────────────────────────────────────────────
    def _browse_file(self, entry_widget):
        path = filedialog.askopenfilename(
            parent=self,
            title="选择 Word 模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if path:
            entry_widget.configure(state="normal")
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)

    def _browse_cover_file(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="选择封面模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if path:
            self._cover_fullpath = path
            self.entry_cover.configure(state="normal")
            self.entry_cover.delete(0, "end")
            self.entry_cover.insert(0, os.path.basename(path))

    def _set_input_path(self, path):
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, path)
        self.file_entry.configure(state="readonly")
        self._log(f'已选择: {os.path.basename(path)}')
        self._update_go_btn()

    def _browse_input(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="选择文档",
            filetypes=[("文档", "*.md *.txt *.docx"), ("所有文件", "*.*")]
        )
        if path:
            self._set_input_path(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            title="指定输出路径",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")]
        )
        if path:
            self.out_entry.configure(state="normal")
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)
            self._log(f'输出指定: {os.path.basename(path)}')

    def _create_blank_md(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            title="新建 AI 输出中转文件",
            initialfile="AI输出中转站.md",
            defaultextension=".md",
            filetypes=[("Markdown 文档", "*.md")]
        )
        if path:
            try:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write("\n\n")
                self._set_input_path(path)
                os.startfile(path)
                self._log(f'📝 已打开中转站！粘贴并保存后，即可点击排版。')
            except Exception as e:
                self._log(f'❌ 中转站创建失败: {str(e)}')

    # ── 输出路径切换 ────────────────────────────────────────────
    def _toggle_output(self):
        if self.same_dir_var.get():
            self.out_entry.configure(state="normal")
            self.out_entry.delete(0, "end")
            self.out_entry.configure(state="disabled", placeholder_text="自动生成（同目录）...")
            self.out_btn.configure(state="disabled")
        else:
            self.out_entry.configure(state="normal", placeholder_text="点击浏览或输入路径...")
            self.out_btn.configure(state="normal")

    # ── 按钮状态 ────────────────────────────────────────────────
    def _update_go_btn(self):
        primary = THEMES.get(self.theme_key, THEMES["warm"])["primary"]
        if self._current_page == "page_text":
            self.go_btn.configure(state="disabled", text="请使用下方 AI 生成按钮",
                                  fg_color="gray")
        elif self.file_entry.get().strip():
            self.go_btn.configure(state="normal", text="一键开始排版",
                                  fg_color=primary)
        else:
            self.go_btn.configure(state="disabled", text="请先选择文件",
                                  fg_color="gray")

    # ── 日志 ────────────────────────────────────────────────────
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.write(f'[{ts}] {msg}\n')

    # ── 转换 ────────────────────────────────────────────────────
    def _start_conversion(self):
        # ── 收集模板配置（跨页面读取 Page 3 状态）──
        cover_path = None
        template_path = None

        if self.check_cover.get():
            cover_path = self._cover_fullpath or self.entry_cover.get().strip() or None

        if self.switch_clone.get():
            template_path = None  # 克隆模式走 get_active_scene_config() → tf 引擎

        from template_config import TemplateConfig
        config = TemplateConfig(cover_path=cover_path, template_path=template_path)

        if config.layer_mode == 2:
            self._log(f'检测到模板注入：Layer 2 样式克隆模式')

        # ── 文件模式 ──
        if self._current_page != "page_text":
            input_path = self.file_entry.get().strip()
            if not input_path:
                self._log('❌ 请先选择文件')
                return
            if not os.path.exists(input_path):
                self._log('❌ 文件不存在，请重新选择')
                return
            ext = os.path.splitext(input_path)[1].lower()
            if self.same_dir_var.get():
                output_path = os.path.splitext(input_path)[0] + '_formatted.docx'
            else:
                output_path = self.out_entry.get().strip()
                if not output_path.endswith('.docx'):
                    output_path += '.docx'
        else:
            # AI 创作页：使用页内按钮操作，不走此流程
            self._log('💡 AI 创作请使用下方"A"AI 生成并一键排版"按钮')
            return

        is_preview = (self.output_mode.get() == "preview")

        self.go_btn.configure(state="disabled", text="转换中...", fg_color="gray")
        self.progress.set(0.2)

        thread = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, ext, is_preview, None, config),
            daemon=True)
        thread.start()

    def _run_conversion(self, input_path, output_path, ext, is_preview, text, config):
        try:
            if is_preview:
                if not text:
                    self._log('  ❌ 预览仅支持文本输入模式')
                    return
                self._log('  预览模式：生成中...')
                tracker, formula_stats = convert_text_to_docx(text, output_path)
                self.last_tracker = tracker
                self.global_formula_stats = formula_stats
                self.after(0, lambda: os.startfile(output_path))
                self._log(f'  ✅ 已打开预览')
            elif text is not None:
                self._log('  文本模式：转换中...')
                tracker, formula_stats = convert_text_to_docx(text, output_path)
                self.last_tracker = tracker
                self.global_formula_stats = formula_stats
                self._log(f'  ✅ 已导出 {os.path.abspath(output_path)}')
                for li, var, msg in _detect_suspicious_vars(text):
                    self._log(f'  ⚠️ 行{li}: [{var}] — {msg}')
            else:
                self._log(f'  开始处理: {os.path.basename(input_path)}')
                if self.switch_clone.get() and ext == '.docx':
                    self._log('  ⚡ 启用深度模板克隆引擎...')
                    tracker, formula_stats = reformat_docx_clone(input_path, output_path, gui_config=config)
                    self.last_tracker = tracker
                    self.global_formula_stats = formula_stats
                elif ext == '.docx':
                    tracker, formula_stats = reformat_docx(input_path, output_path, config=config)
                    self.last_tracker = tracker
                    self.global_formula_stats = formula_stats
                else:
                    tracker, formula_stats = convert_markdown_to_docx(input_path, output_path)
                    self.last_tracker = tracker
                    self.global_formula_stats = formula_stats
                self._log(f'  ✅ 转换完成，已保存至 {os.path.abspath(output_path)}')
        except Exception:
            self._log(f'  ❌ 转换失败：')
            for line in traceback.format_exc().splitlines()[-4:]:
                self._log(f'     {line}')
        finally:
            self.after(0, self._conversion_done)

    def _conversion_done(self):
        self.progress.set(1.0)
        self._update_go_btn()
        has_report = (
            (hasattr(self, 'last_tracker') and self.last_tracker)
            or (hasattr(self, 'global_formula_stats') and self.global_formula_stats
                and self.global_formula_stats.matched > 0)
        )
        if has_report:
            self.report_btn.configure(state="normal", fg_color=COLOR_PRIMARY, text_color=COLOR_TEXT)
        self.after(3000, lambda: self.progress.set(0))

    def _ocr_to_docx(self, text):
        """OCR 提取后生成唯一排版后的 Word 文档并打开"""
        if not text:
            raise ValueError("API 未返回任何内容，请检查您的网络或 API Key 设置。")
        if text.startswith("⚠️") or text.startswith("❌"):
            raise ValueError(text)

        import os, tempfile, time, threading
        from docx import Document
        from docx_renderer import _add_body
        timestamp = int(time.time())
        filename = f"AI_提取草稿_{timestamp}.docx"
        path = os.path.join(tempfile.gettempdir(), filename)
        raw_path = path + ".raw.docx"

        doc = Document()

        def _wrap_formula(line):
            line = line.strip()
            if "\\" in line and not (line.startswith("$") or line.startswith("\\begin")):
                return f"${line}$"
            return line

        for line in text.split('\n'):
            processed = _wrap_formula(line)
            if processed:
                _add_body(doc, processed, config=None)

        try:
            doc.save(raw_path)
        except Exception as e:
            self._log(f"❌ 写入文件失败 (请关闭已打开的 Word 文档): {str(e)}")
            raise Exception(f"排版底层错误: {str(e)}")

        def _worker():
            try:
                self._log("正在套用当前排版预设...")
                reformat_docx(raw_path, path)
                if os.path.exists(path):
                    os.startfile(path)
                self._log(f"✅ 智能提取完成，已生成文档: {filename}")
            except Exception as e:
                self._log(f"❌ 排版注入失败: {str(e)}")
                if os.path.exists(raw_path):
                    os.startfile(raw_path)
            finally:
                if os.path.exists(raw_path):
                    try:
                        os.remove(raw_path)
                    except Exception:
                        pass

        threading.Thread(target=_worker, daemon=True).start()

    def _show_inspiration_modal(self):
        t = THEMES.get(self.theme_key, THEMES["warm"])
        modal = ctk.CTkToplevel(self)
        modal.overrideredirect(True)
        modal.attributes("-topmost", True)
        w, h = 480, 280
        x = self.winfo_rootx() + (self.winfo_width() - w) // 2
        y = self.winfo_rooty() + (self.winfo_height() - h) // 2
        modal.geometry(f"{w}x{h}+{x}+{y}")
        modal.configure(fg_color=t["bg_main"])

        inner = ctk.CTkFrame(modal, fg_color=t["bg_panel"],
                             border_width=1, border_color=t["primary"],
                             corner_radius=8)
        inner.pack(fill="both", expand=True, padx=2, pady=2)

        ctk.CTkLabel(
            inner, text="一年后，\n\n你会站在曾经仰望的高度，\n\n活成自己喜欢的模样。",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=22),
            text_color=t["text"], justify="center"
        ).place(relx=0.5, rely=0.45, anchor="center")

        ctk.CTkLabel(
            inner, text="— 轻触任意处闭合 —",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=12),
            text_color=t["text_muted"]
        ).pack(side="bottom", pady=25)

        def _close(event=None):
            modal.destroy()
        for wgt in (modal, inner):
            wgt.bind("<Button-1>", _close)
        modal.focus_force()
        modal.bind("<Escape>", _close)

    def _show_settings(self):
        if hasattr(self, "_settings_win") and self._settings_win.winfo_exists():
            self._settings_win.focus()
            return
        from ai_assist.settings_dialog import SettingsDialog
        self._settings_win = SettingsDialog(self,
                            log_callback=self._log,
                            fill_callback=self._ocr_to_docx,
                            theme_callback=self.apply_ui_theme)

    def _show_report_dialog(self):
        has_tracker = hasattr(self, 'last_tracker') and self.last_tracker
        has_formula = (hasattr(self, 'global_formula_stats')
                       and self.global_formula_stats
                       and self.global_formula_stats.matched > 0)
        if not has_tracker and not has_formula:
            return

        # ── 双色系主题嗅探（默认墨绿，暖白留给未来配置开关）──
        current_theme = "moss"
        if hasattr(self, "config_manager"):
            current_theme = self.config_manager.get("ui_theme", "moss")
        elif hasattr(self, "theme_key"):
            current_theme = self.theme_key

        if current_theme == "moss":
            COLOR_BG_DARK       = "#2B312C"
            COLOR_BG_INNER      = "#232824"
            COLOR_TEXT_MAIN     = "#E3DFD2"
            COLOR_TEXT_MUTED    = "#92978D"
            COLOR_ACCENT_ORANGE = "#DF593A"
            COLOR_CARD_BORDER   = "#353C36"
            COLOR_SUCCESS_TAG   = "#A3B5A7"
            COLOR_SAFE_CARD_BG  = "#313933"
            COLOR_HOVER_BTN     = "#454E46"
            WINDOW_TITLE_BAR    = 0x002C312B
            RISK_CARD_BG        = "#3A2926"
            CONFIRM_FG          = COLOR_CARD_BORDER
            CONFIRM_TEXT        = COLOR_TEXT_MAIN
        else:
            COLOR_BG_DARK       = "#FDFBF7"
            COLOR_BG_INNER      = "#F5F3ED"
            COLOR_TEXT_MAIN     = "#37352F"
            COLOR_TEXT_MUTED    = "#8B7D72"
            COLOR_ACCENT_ORANGE = "#DF593A"
            COLOR_CARD_BORDER   = "#E5E3DC"
            COLOR_SUCCESS_TAG   = "#6B8E76"
            COLOR_SAFE_CARD_BG  = "#EBE9E2"
            COLOR_HOVER_BTN     = "#EBE9E4"
            WINDOW_TITLE_BAR    = 0x00F7FBFD
            RISK_CARD_BG        = "#FCEFEA"
            CONFIRM_FG          = COLOR_TEXT_MUTED
            CONFIRM_TEXT        = "#FDFBF7"

        top = ctk.CTkToplevel(self)
        top.title("FormatX  岁时流转纪事")
        top.geometry("740x720")

        top.configure(fg_color=COLOR_BG_DARK)
        top.transient(self)
        top.grab_set()

        poetic_font_title = ctk.CTkFont(
            family="LXGW WenKai, 楷体, STKaiti, Microsoft YaHei",
            size=20, weight="normal")
        poetic_font_card_num = ctk.CTkFont(
            family="JetBrains Mono, Consolas, LXGW WenKai",
            size=24, weight="bold")
        poetic_font_card_lbl = ctk.CTkFont(
            family="LXGW WenKai, 楷体, STKaiti, Microsoft YaHei",
            size=12)
        poetic_font_small = ctk.CTkFont(
            family="LXGW WenKai, 楷体, STKaiti, Microsoft YaHei",
            size=13)
        poetic_font_mono = ctk.CTkFont(
            family="JetBrains Mono, Consolas, monospace", size=12)

        try:
            import ctypes
            top.update_idletasks()
            hwnd = ctypes.windll.user32.GetParent(top.winfo_id())
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, 35, ctypes.byref(ctypes.c_int(WINDOW_TITLE_BAR)),
                ctypes.sizeof(ctypes.c_int))
        except Exception:
            pass

        formula_matched = self.global_formula_stats.matched if has_formula else 0
        formula_converted = self.global_formula_stats.converted if has_formula else 0
        summary = self.last_tracker.summary() if has_tracker else {"total": 0, "failures": 0}
        risk_count = summary["failures"]
        format_count = summary["total"]

        # ── 顶部标题 ──
        ctk.CTkLabel(top, text="拂去冗杂，文档已恢复它最初的秩序 ✨",
                     font=poetic_font_title,
                     text_color=COLOR_TEXT_MAIN).pack(fill="x", padx=30, pady=(25, 15))

        # ── 数据卡片栅格 ──
        cards_frame = ctk.CTkFrame(top, fg_color="transparent")
        cards_frame.pack(fill="x", padx=30, pady=(0, 15))
        cards_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="equal")

        def create_stat_card(parent, col, icon, label_text, num_text,
                             num_color, bg_color=COLOR_BG_INNER):
            card = ctk.CTkFrame(parent, fg_color=bg_color, corner_radius=8,
                                border_width=1, border_color=COLOR_CARD_BORDER)
            card.grid(row=0, column=col, padx=6, sticky="nsew")
            ctk.CTkLabel(card, text=icon, font=poetic_font_title,
                         text_color=num_color).pack(pady=(12, 2))
            ctk.CTkLabel(card, text=label_text, font=poetic_font_card_lbl,
                         text_color=COLOR_TEXT_MUTED).pack(pady=0)
            ctk.CTkLabel(card, text=num_text, font=poetic_font_card_num,
                         text_color=num_color).pack(pady=(2, 12))

        create_stat_card(cards_frame, 0, "🔮", "公式破镜",
                         f"{formula_converted}/{formula_matched}", COLOR_TEXT_MAIN)
        create_stat_card(cards_frame, 1, "🌿", "格式调理",
                         f"{format_count} 处", COLOR_TEXT_MAIN)
        if risk_count == 0:
            create_stat_card(cards_frame, 2, "🛡️", "潜藏隐患", "0 处",
                             COLOR_TEXT_MUTED, bg_color=COLOR_SAFE_CARD_BG)
        else:
            create_stat_card(cards_frame, 2, "⚠️", "潜藏隐患",
                             f"{risk_count} 处", COLOR_ACCENT_ORANGE,
                             bg_color=RISK_CARD_BG)

        # ── 富文本控制台 ──
        text_container = ctk.CTkFrame(top, fg_color="transparent")
        text_container.pack(fill="both", expand=True, padx=30, pady=5)

        text_area = ctk.CTkTextbox(
            text_container, font=poetic_font_mono, wrap="word",
            fg_color=COLOR_BG_INNER, text_color=COLOR_TEXT_MAIN,
            border_width=0, corner_radius=6)
        text_area.pack(fill="both", expand=True)

        raw_text_widget = text_area._textbox
        raw_text_widget.tag_config("poetic_success", foreground=COLOR_SUCCESS_TAG)
        raw_text_widget.tag_config("poetic_alert", foreground=COLOR_ACCENT_ORANGE)
        raw_text_widget.tag_config("poetic_divider", foreground=COLOR_CARD_BORDER)

        divider = " · " * 22 + "\n"

        # ── 纸上风骨 ──
        text_area.insert("end", "\n【 纸上风骨 】\n")
        text_area.insert("end", divider, "poetic_divider")
        if formula_matched == 0 or not has_formula:
            text_area.insert("end", " 全文未见数理公式，卷面清秀流利。\n")
        else:
            fs = self.global_formula_stats
            critical_occs = [o for o in fs.occurrences if o.confidence < 0.85]
            if not critical_occs:
                text_area.insert("end", " 全文公式脉络清晰、骨架端正，并无隐疾。它们将以最优雅的姿态，跃然纸上。\n")
            else:
                source_cn = {
                    "ole_equation": "岁月遗痕 (MathType)",
                    "ocr_fragment": "光影拓印 (视觉提取)",
                    "plain_text": "网页复制",
                }
                for occ in critical_occs[:10]:
                    text_area.insert("end",
                        f"  ✦ 坐标: 第 {occ.paragraph_index} 段 | "
                        f"来源: {source_cn.get(occ.source_type, occ.source_type)}\n")
                    if occ.is_fixed:
                        text_area.insert("end",
                            "    ✓ [拯救] 察觉残缺，已于 ICU 抢救仓中为其筑回完好真身。\n",
                            "poetic_success")
                        text_area.insert("end",
                            f"    修复后代码: {occ.repaired_text}\n")
                    else:
                        text_area.insert("end",
                            f"    ⚠️ [隐患] 诊断结论: {', '.join(occ.warnings)}\n",
                            "poetic_alert")
                if len(critical_occs) > 10:
                    text_area.insert("end",
                        f"    ... (其余 {len(critical_occs) - 10} 处残缺，皆已无感修复并编纂在册)\n")

        # ── 岁时无惊 ──
        text_area.insert("end", "\n【 岁时无惊 】\n")
        text_area.insert("end", divider, "poetic_divider")
        if risk_count == 0:
            text_area.insert("end", " 极好。行文流利，不见波折，没有一行文字流离失所。\n")
        else:
            failures = self.last_tracker.get_failures()
            for f in failures:
                text_area.insert("end",
                    f"  ✦ 遭遇阻碍: 在处理段落 #{f.paragraph_index} [{f.target}] 时\n")
                text_area.insert("end",
                    f"    ❌ [偏离] 缘由: {f.failure_reason}\n", "poetic_alert")

        # ── 细微之处 ──
        if has_tracker:
            text_area.insert("end", "\n【 细微之处 】\n")
            text_area.insert("end", divider, "poetic_divider")

            def _poetic_translator(record):
                act = record.after
                rule = record.rule_name
                if "1 级标题" in act:
                    return "[秩序] 重新雕琢 1 级标题的节律，使其端正归位"
                if "2 级标题" in act:
                    return "[秩序] 梳理 2 级标题的脉络，令其长幼有序"
                if "3 级标题" in act:
                    return "[秩序] 抚平 3 级标题的呼吸，让它轻盈落下"
                if "MathType" in rule:
                    count = "".join(filter(str.isdigit, act))
                    return f"[洗礼] 剥去了 {count} 处岁月的铜锈，重现公式的清透"
                if "Reference" in rule:
                    return "[留白] 为参考文献添了一抹悬挂缩进，以安顿群星"
                if "Heading Formatting" in rule:
                    return f"[塑骨] 为段落重新注入了 {act.replace('注入 ', '')} 的骨架"
                return f"[流转] {act}"

            success_records = [r for r in self.last_tracker.records if r.success][:8]
            for r in success_records:
                text_area.insert("end", "  ✓ ", "poetic_success")
                text_area.insert("end", f"{_poetic_translator(r)}\n")

            if summary['total'] > 8:
                text_area.insert("end",
                    f"\n  ... (余下 {summary['total'] - 8 - risk_count} 处维度的文字精雕，已悄然隐入后台)\n",
                    "poetic_success")

        text_area.configure(state="disabled")

        # ── 底部 Footer 操作栏 ──
        footer_frame = ctk.CTkFrame(top, fg_color="transparent")
        footer_frame.pack(fill="x", padx=30, pady=(15, 25))

        def _action_confirm():
            top.grab_release()
            top.destroy()

        def _action_export_log():
            path = filedialog.asksaveasfilename(
                parent=top,
                title="将纪事铭刻于磐石",
                initialfile=f"FormatX_排版审计报告_{datetime.now().strftime('%m%d')}.txt",
                defaultextension=".txt",
                filetypes=[("纯文本文件", "*.txt")]
            )
            if path:
                try:
                    content_to_save = raw_text_widget.get("1.0", "end")
                    with open(path, 'w', encoding='utf-8') as f:
                        f.write(content_to_save)
                    self._log(f"报告已保存: {os.path.basename(path)}")
                except Exception as ex:
                    self._log(f"导出报告失败: {str(ex)}")

        btn_export = ctk.CTkButton(
            footer_frame, text="导出纪事", width=110, height=36,
            font=poetic_font_small, fg_color="transparent",
            text_color=COLOR_TEXT_MAIN, border_width=1,
            border_color=COLOR_TEXT_MUTED, hover_color=COLOR_HOVER_BTN,
            command=_action_export_log
        )
        btn_export.pack(side="left", padx=(0, 8))

        if has_tracker and has_formula:
            def _action_export_report():
                path = filedialog.asksaveasfilename(
                    parent=top,
                    title="导出排版审计报告",
                    initialfile=f"FormatX_审计报告_{datetime.now().strftime('%m%d')}.md",
                    defaultextension=".md",
                    filetypes=[("Markdown 文件", "*.md")]
                )
                if path:
                    try:
                        from core.report.collector import collect_report
                        from core.report.markdown_report import generate_markdown_report
                        rd = collect_report(self.last_tracker, self.global_formula_stats)
                        generate_markdown_report(rd, path)
                        self._log(f"报告已保存: {os.path.basename(path)}")
                    except Exception as ex:
                        self._log(f"导出失败: {ex}")

            btn_report = ctk.CTkButton(
                footer_frame, text="导出报告", width=110, height=36,
                font=poetic_font_small, fg_color=COLOR_HOVER_BTN,
                text_color=COLOR_TEXT_MAIN, border_width=0,
                hover_color=COLOR_CARD_BORDER,
                command=_action_export_report
            )
            btn_report.pack(side="left")

        btn_confirm = ctk.CTkButton(
            footer_frame, text="容我知悉", width=130, height=36,
            font=ctk.CTkFont(family="LXGW WenKai, Microsoft YaHei", size=13, weight="bold"),
            fg_color=CONFIRM_FG, text_color=CONFIRM_TEXT,
            hover_color=COLOR_HOVER_BTN, command=_action_confirm
        )
        btn_confirm.pack(side="right")


# ── 图标生成 ──────────────────────────────────────────────────
def _make_ico_data():
    import struct
    bmp_size = 32 * 32 * 4 + 40
    ico = bytearray()
    ico += struct.pack('<HHH', 0, 1, 1)
    ico += struct.pack('<BBBBHHII', 32, 32, 0, 0, 1, 32, bmp_size, 22)
    bmp = bytearray()
    bmp += struct.pack('<IiiHHIIiiII', 40, 32, 64, 1, 32, 0, 0, 0, 0, 0, 0)
    for y in range(31, -1, -1):
        for x in range(32):
            bmp += struct.pack('BBBB', 26, 115, 232, 255)
    ico += bmp
    return bytes(ico)


# ── 入口 ──────────────────────────────────────────────────────
if __name__ == '__main__':
    app = App()
    app.mainloop()
