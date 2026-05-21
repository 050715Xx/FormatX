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
# 运行时写入临时文件供 iconbitmap 使用
_ICON_B64 = None  # 打包时由 build 脚本注入，或运行时从文件读取

# ── 从 format_conversion 导入核心函数 ──────────────────────────
try:
    from format_conversion import (convert_markdown_to_docx, reformat_docx,
                                     convert_text_to_docx, _detect_suspicious_vars)
except ImportError:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from format_conversion import (convert_markdown_to_docx, reformat_docx,
                                     convert_text_to_docx, _detect_suspicious_vars)


# ── 配置 ──────────────────────────────────────────────────────
APP_TITLE = "FormatX"
APP_SIZE = "640x720"
THEME_COLOR = "#1a73e8"
CREAM_BG = "#FAF5F0"
CREAM_FRAME = "#FFFDF9"
BTN_HOVER = "#1557b0"

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

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")


# ── 日志面板 ──────────────────────────────────────────────────
class ConsoleLog(ctk.CTkTextbox):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("font", ("Consolas", 12))
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("height", 150)
        kwargs.setdefault("wrap", "word")
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
        # 强制声明 Windows AppID，解除任务栏图标与 Python 默认图标的绑定
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                'yimin.formatx.v1.0')
        except Exception:
            pass
        super().__init__()
        self.title(APP_TITLE)
        self.geometry(APP_SIZE)
        self.minsize(480, 500)
        self.configure(fg_color=CREAM_BG)
        self._set_titlebar_color()
        self._set_icon()
        self._build_ui()
        self._setup_dnd()

    # ── 标题栏颜色 ─────────────────────────────────────────────
    def _set_titlebar_color(self):
        """Windows 11/10: 将标题栏颜色设为与背景一致的奶油白。"""
        try:
            import ctypes
            self.update()
            hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
            # DWMWA_CAPTION_COLOR = 35, 需要 Windows 11 build 22000+
            try:
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 35,
                    ctypes.byref(ctypes.c_int(0x00F0F5FA)),  # #FAF5F0 颠倒为 BGR
                    ctypes.sizeof(ctypes.c_int))
            except Exception:
                pass  # Win10 不支持，静默回退
        except Exception:
            pass

    # ── 窗口图标 ───────────────────────────────────────────────
    def _set_icon(self):
        # 优先用 sys._MEIPASS 的 3.ico
        if getattr(sys, '_MEIPASS', ''):
            p = os.path.join(sys._MEIPASS, '3.ico')
            if os.path.exists(p):
                try:
                    self.iconbitmap(p)
                    return
                except Exception:
                    pass
        # PyInstaller --icon 嵌入在 exe 里，直接读自身上
        try:
            self.iconbitmap(sys.executable)
            return
        except Exception:
            pass
        # 兜底：内嵌 ico 数据写临时文件
        try:
            tf = tempfile.NamedTemporaryFile(suffix='.ico', delete=False)
            tf.write(_make_ico_data())
            tf.close()
            self.iconbitmap(tf.name)
        except Exception:
            pass

    # ── UI 构建 ────────────────────────────────────────────────
    def _build_ui(self):
        lf = ctk.CTkFont(family="Microsoft YaHei", size=12)

        # 主标题
        # 尝试加载 Montserrat，失败则回退系统字体
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
            title_family = "Segoe UI"  # Windows 现代无衬线

        ctk.CTkLabel(self, text="FormatX",
                     font=ctk.CTkFont(family=title_family, size=34, weight="bold"),
                     text_color="#000000"
                     ).pack(pady=(20, 0))
        ctk.CTkLabel(self, text="自动化排版与可编辑公式解析器",
                     font=ctk.CTkFont(family="Microsoft YaHei", size=11),
                     text_color="gray"
                     ).pack(pady=(0, 10))

        # 右上角暖茶灰温馨寄语
        ctk.CTkLabel(
            self,
            text="满怀希望，\n就会所向披靡 ✨",
            font=ctk.CTkFont(family="Microsoft YaHei", size=14, slant="italic"),
            text_color="#B0A8A0",
            justify="right"
        ).place(relx=0.96, rely=0.03, anchor="ne")

        # ── 选项卡 ──
        tab_row = ctk.CTkFrame(self, fg_color="transparent")
        tab_row.pack(fill="x", padx=16, pady=(0, 4))
        self._tab = "file"
        self.tab_file_btn = ctk.CTkButton(
            tab_row, text="📂 文件上传", width=140, corner_radius=6,
            font=("Microsoft YaHei", 13), fg_color=THEME_COLOR,
            hover_color=BTN_HOVER, command=lambda: self._switch_tab("file")
        )
        self.tab_file_btn.pack(side="left", padx=(0, 8))
        self.tab_text_btn = ctk.CTkButton(
            tab_row, text="✏️ 文本输入", width=140, corner_radius=6,
            font=("Microsoft YaHei", 13), fg_color="gray",
            hover_color=BTN_HOVER, command=lambda: self._switch_tab("text")
        )
        self.tab_text_btn.pack(side="left")

        # ── 文件选择区 ──
        self.file_frame = ctk.CTkFrame(self, border_width=1, fg_color=CREAM_FRAME)
        self.file_frame.pack(fill="x", padx=16, pady=(8, 6), ipady=8)

        # ── 文本输入区 ──
        self.text_frame = ctk.CTkFrame(self, border_width=1, fg_color=CREAM_FRAME)
        self.text_area = ctk.CTkTextbox(
            self.text_frame,
            font=("Consolas", 13), height=200, wrap="word",
            border_width=1
        )
        self.text_area.pack(fill="both", expand=True, padx=14, pady=(10, 4))
        self.text_area.insert("1.0",
            "# 在此粘贴 Markdown 文本...\n\n"
            "正文内容，支持 $\\frac{a}{b}$ 公式。\n\n"
            "```python\nprint('hello')\n```")

        # ── 输出模式 ──
        self.output_mode = ctk.StringVar(value="export")
        radio_row = ctk.CTkFrame(self, fg_color="transparent")
        radio_row.pack(fill="x", padx=20, pady=(4, 0))
        ctk.CTkLabel(radio_row, text="输出：",
                     font=("Microsoft YaHei", 11)).pack(side="left", padx=(0, 8))
        ctk.CTkRadioButton(radio_row, text="导出 .docx", variable=self.output_mode,
                           value="export", font=("Microsoft YaHei", 11)
                           ).pack(side="left", padx=(0, 12))
        ctk.CTkRadioButton(radio_row, text="直接显示结果", variable=self.output_mode,
                           value="preview", font=("Microsoft YaHei", 11)
                           ).pack(side="left")

        # ── 文件选择区内容 ──
        file_frame = self.file_frame
        file_frame.pack(fill="x", padx=16, pady=(8, 6), ipady=8)

        ctk.CTkLabel(file_frame, text="📂 选择待排版的文档 (.md / .txt / .docx)",
                     font=lf, anchor="w"
                     ).pack(fill="x", padx=14, pady=(10, 4))

        entry_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        entry_row.pack(fill="x", padx=14, pady=(0, 4))
        entry_row.columnconfigure(0, weight=1)

        self.file_entry = ctk.CTkEntry(entry_row, placeholder_text="请选择文件 或 拖拽文件到此处...",
                                       font=("Microsoft YaHei", 11))
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.file_entry.configure(state="readonly")

        self.file_entry.bind("<Double-Button-1>", lambda e: self._browse_input())

        browse_btn = ctk.CTkButton(
            entry_row, text="浏览", width=70,
            command=self._browse_input,
            hover_color=BTN_HOVER
        )
        browse_btn.grid(row=0, column=1)

        new_file_btn = ctk.CTkButton(
            entry_row, text="新建空白.md", width=90,
            command=self._create_blank_md,
            fg_color="#28a745", hover_color="#218838"
        )
        new_file_btn.grid(row=0, column=2, padx=(5, 0))

        # 输出路径
        out_label_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        out_label_row.pack(fill="x", padx=14, pady=(10, 4))
        ctk.CTkLabel(out_label_row, text="📁 输出路径",
                     font=lf, anchor="w"
                     ).pack(side="left")

        self.same_dir_var = ctk.BooleanVar(value=True)
        self.same_dir_cb = ctk.CTkCheckBox(
            out_label_row, text="同目录", variable=self.same_dir_var,
            command=self._toggle_output,
            font=("Microsoft YaHei", 11), checkbox_width=20, checkbox_height=20
        )
        self.same_dir_cb.pack(side="right", padx=(8, 0))

        out_row = ctk.CTkFrame(file_frame, fg_color="transparent")
        out_row.pack(fill="x", padx=14, pady=(0, 10))
        out_row.columnconfigure(0, weight=1)

        self.out_entry = ctk.CTkEntry(out_row, placeholder_text="自动生成（同目录）...",
                                      font=("Microsoft YaHei", 11), state="disabled")
        self.out_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.out_btn = ctk.CTkButton(
            out_row, text="浏览", width=70,
            command=self._browse_output,
            hover_color=BTN_HOVER, state="disabled"
        )
        self.out_btn.grid(row=0, column=1)

        # ── 操作与日志区 ──
        action_frame = ctk.CTkFrame(self, border_width=1, fg_color=CREAM_FRAME)
        action_frame.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.go_btn = ctk.CTkButton(
            action_frame, text="请先选择文件", height=45,
            corner_radius=8, font=("Microsoft YaHei", 16, "bold"),
            command=self._start_conversion,
            state="disabled", fg_color="gray"
        )
        self.go_btn.pack(fill="x", padx=14, pady=(14, 8))

        self.progress = ctk.CTkProgressBar(action_frame)
        self.progress.pack(fill="x", padx=14, pady=(0, 6))
        self.progress.set(0)

        self.log = ConsoleLog(action_frame)
        self.log.pack(fill="both", expand=True, padx=14, pady=(0, 14))

        self._log('  FormatX 已就绪')
        self._log('  支持 .md / .txt / .docx')
        self._log('')

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

    # ── 选项卡切换 ────────────────────────────────────────────
    def _switch_tab(self, tab):
        self._tab = tab
        if tab == "file":
            self.tab_file_btn.configure(fg_color=THEME_COLOR)
            self.tab_text_btn.configure(fg_color="gray")
            self.file_frame.pack(fill="x", padx=16, pady=(8, 6), ipady=8)
            self.text_frame.pack_forget()
            self._update_go_btn()
        else:
            self.tab_text_btn.configure(fg_color=THEME_COLOR)
            self.tab_file_btn.configure(fg_color="gray")
            self.file_frame.pack_forget()
            self.text_frame.pack(fill="both", expand=True, padx=16, pady=(8, 6))
            self.go_btn.configure(state="normal", text="一键开始排版",
                                  fg_color=THEME_COLOR)

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

    # ── 文件浏览 ────────────────────────────────────────────────
    def _set_input_path(self, path):
        self.file_entry.configure(state="normal")
        self.file_entry.delete(0, "end")
        self.file_entry.insert(0, path)
        self.file_entry.configure(state="readonly")
        self._log(f'  已选择: {os.path.basename(path)}')
        self._update_go_btn()

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择文档",
            filetypes=[("文档", "*.md *.txt *.docx"), ("所有文件", "*.*")]
        )
        if path:
            self._set_input_path(path)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="指定输出路径",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")]
        )
        if path:
            self.out_entry.configure(state="normal")
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, path)
            self._log(f'  输出指定: {os.path.basename(path)}')

    def _create_blank_md(self):
        """一键新建空白中转站，并自动填入路径"""
        path = filedialog.asksaveasfilename(
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
                self._log(f'  📝 已打开中转站！粘贴并保存后，即可点击排版。')
            except Exception as e:
                self._log(f'  ❌ 中转站创建失败: {str(e)}')

    # ── 按钮状态 ────────────────────────────────────────────────
    def _update_go_btn(self):
        if self.file_entry.get().strip():
            self.go_btn.configure(state="normal", text="一键开始排版",
                                  fg_color=THEME_COLOR)
        else:
            self.go_btn.configure(state="disabled", text="请先选择文件",
                                  fg_color="gray")

    # ── 日志 ────────────────────────────────────────────────────
    def _log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log.write(f'[{ts}] {msg}\n')

    # ── 转换 ────────────────────────────────────────────────────
    def _start_conversion(self):
        is_preview = (self.output_mode.get() == "preview")

        if self._tab == "file":
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
            # 文本模式
            text = self.text_area.get("1.0", "end").strip()
            if not text:
                self._log('❌ 请输入文本内容')
                return
            ext = '.md'
            input_path = None
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(tempfile.gettempdir(), f'FormatX_{ts}.docx')

        self.go_btn.configure(state="disabled", text="转换中...", fg_color="gray")
        self.progress.start()

        thread = threading.Thread(
            target=self._run_conversion,
            args=(input_path, output_path, ext, is_preview, text if self._tab == "text" else None),
            daemon=True)
        thread.start()

    def _run_conversion(self, input_path, output_path, ext, is_preview, text):
        try:
            if is_preview:
                if not text:
                    self._log('  ❌ 预览仅支持文本输入模式')
                    return
                self._log('  预览模式：生成中...')
                convert_text_to_docx(text, output_path)
                self.after(0, lambda: os.startfile(output_path))
                self._log(f'  ✅ 已打开预览')
            elif text is not None:
                self._log('  文本模式：转换中...')
                convert_text_to_docx(text, output_path)
                self._log(f'  ✅ 已导出 {os.path.abspath(output_path)}')
                # 变量名异常检测
                for li, var, msg in _detect_suspicious_vars(text):
                    self._log(f'  ⚠️ 行{li}: [{var}] — {msg}')
            else:
                self._log(f'  开始处理: {os.path.basename(input_path)}')
                if ext == '.docx':
                    reformat_docx(input_path, output_path)
                else:
                    convert_markdown_to_docx(input_path, output_path)
                self._log(f'  ✅ 转换完成，已保存至 {os.path.abspath(output_path)}')
        except Exception:
            self._log(f'  ❌ 转换失败：')
            for line in traceback.format_exc().splitlines()[-4:]:
                self._log(f'     {line}')
        finally:
            self.after(0, self._conversion_done)

    def _conversion_done(self):
        self.progress.stop()
        self.progress.set(1.0)
        self._update_go_btn()
        self.after(3000, lambda: self.progress.set(0))


# ── 图标生成 ──────────────────────────────────────────────────
def _make_ico_data():
    """生成一个最小蓝色方块 .ico 兜底。"""
    import struct
    # 32x32 蓝色方块 BMP 数据，嵌入 .ico
    bmp_size = 32 * 32 * 4 + 40
    ico = bytearray()
    ico += struct.pack('<HHH', 0, 1, 1)  # ICO header: reserved, type=icon, count=1
    ico += struct.pack('<BBBBHHII', 32, 32, 0, 0, 1, 32, bmp_size, 22)
    # BMP header
    bmp = bytearray()
    bmp += struct.pack('<IiiHHIIiiII', 40, 32, 64, 1, 32, 0, 0, 0, 0, 0, 0)
    for y in range(31, -1, -1):
        for x in range(32):
            bmp += struct.pack('BBBB', 26, 115, 232, 255)  # #1a73e8
    ico += bmp
    return bytes(ico)


# ── 入口 ──────────────────────────────────────────────────────
if __name__ == '__main__':
    app = App()
    app.mainloop()
