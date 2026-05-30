"""统一设置弹窗 — 左侧导航 + 右侧内容"""
import os
import threading
import customtkinter as ctk
from core.config_manager import load_full_config, save_full_config, update_config
from ai_assist.multi_llm_client import (test_provider_connection, fetch_remote_models,
                                         extract_text_from_image)
from ai_assist.image_capture import get_image_from_clipboard, image_to_base64
from format_conversion import convert_text_to_docx

_BG_MAIN    = "#F9F7F2"
_BG_PANEL   = "#F5F3ED"
_PRIMARY    = "#E5E3DC"
_HOVER      = "#EBE9E4"
_TEXT       = "#37352F"
_TEXT_MUTED = "#8B7D72"


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent, log_callback=None, fill_callback=None, theme_callback=None):
        super().__init__(parent)
        self.title("FormatX 设置中心")
        self.resizable(False, False)
        self.geometry("900x660")

        import sys, os
        ico = os.path.join(getattr(sys, '_MEIPASS', '.'), '3.ico')
        if not os.path.exists(ico):
            ico = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '3.ico')
        if os.path.exists(ico):
            try:
                self.iconbitmap(ico)
            except Exception:
                pass

        def _center():
            try:
                pw = parent.winfo_width()
                ph = parent.winfo_height()
                if pw > 100 and ph > 100:
                    px = parent.winfo_rootx()
                    py = parent.winfo_rooty()
                    self.geometry(f"900x660+{px+(pw-900)//2}+{py+(ph-660)//2}")
            except Exception:
                pass
        self.after(100, _center)

        # 双向前兼容回调
        self.log_callback = log_callback or (lambda m: None)
        self._log = self.log_callback
        self.fill_callback = fill_callback or (lambda t: None)
        self._fill = self.fill_callback
        self.theme_change_callback = theme_callback or (lambda t: None)

        # 持久化主题配置
        self.current_theme = load_full_config().get("ui_theme", "moss")
        self.COLORS = self._get_theme_colors(self.current_theme)

        self.configure(fg_color=self.COLORS["bg_main"])
        self.transient(parent)
        self.attributes("-topmost", True)

        try:
            from tkinterdnd2 import DND_FILES
            self.drop_target_register(DND_FILES)
            self.dnd_bind('<<Drop>>', self._on_drop_image)
        except Exception:
            pass

        # 布局骨架
        self.nav_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=self.COLORS["bg_panel"], width=160)
        self.nav_frame.pack(side="left", fill="y")
        self.nav_frame.pack_propagate(False)

        self.main_container = ctk.CTkFrame(self, corner_radius=0, fg_color=self.COLORS["bg_main"])
        self.main_container.pack(side="right", fill="both", expand=True)

        self.panels = {}
        self.nav_buttons = {}

        self._build_nav()
        self._build_api_panel()
        self.panels["api"] = self.api_panel
        self._build_ocr_panel()
        self.panels["ocr"] = self.ocr_panel
        self._build_theme_tab()
        self._build_preset_tab()
        self._build_text_format_panel()
        self.panels["text_format"] = self.text_format_panel

        self._switch_tab("api")
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._apply_theme_to_self()
        self.after(0, self.grab_set)

    def _set_appearance_mode(self, mode_string):
        pass

    def _on_close(self):
        try:
            if hasattr(self, 'drop_target_unregister'):
                self.drop_target_unregister()
        except Exception:
            pass
        self.grab_release()
        self.destroy()

    def _switch_tab(self, key):
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.configure(fg_color=self.COLORS["primary"])
            else:
                btn.configure(fg_color="transparent")
        for k, panel in self.panels.items():
            if k == key:
                panel.pack(fill="both", expand=True)
            else:
                panel.pack_forget()

    def _get_theme_colors(self, theme_key):
        if theme_key == "moss":
            return {
                "bg_main": "#2B312C", "bg_panel": "#232824", "primary": "#353C36",
                "hover": "#454E46", "text_main": "#E3DFD2", "text_muted": "#92978D",
                "accent": "#DF593A"
            }
        else:
            return {
                "bg_main": "#F9F7F2", "bg_panel": "#F5F3ED", "primary": "#E5E3DC",
                "hover": "#EBE9E4", "text_main": "#37352F", "text_muted": "#8B7D72",
                "accent": "#A3B5A7"
            }

    # ── 导航栏构建 ──
    def _build_nav(self):
        ctk.CTkLabel(
            self.nav_frame, text="配置中心",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=16, pady=(20, 15))

        def _add_nav_btn(key, text, icon):
            btn = ctk.CTkButton(
                self.nav_frame, text=f"{icon}  {text}", anchor="w", height=36,
                corner_radius=6, font=ctk.CTkFont(family="Microsoft YaHei", size=12),
                fg_color="transparent", text_color=self.COLORS["text_muted"],
                hover_color=self.COLORS["primary"],
                command=lambda: self._switch_tab(key)
            )
            btn.pack(fill="x", padx=10, pady=2)
            self.nav_buttons[key] = btn

        _add_nav_btn("api", "API 配置", "🔑")
        _add_nav_btn("ocr", "智能提取", "📸")
        _add_nav_btn("theme", "界面外观", "🎨")
        _add_nav_btn("preset", "排版预设", "📐")
        _add_nav_btn("text_format", "文本直排", "📝")

    # ── 排版预设面板 ──
    def _build_preset_tab(self):
        self.panels["preset"] = ctk.CTkFrame(self.main_container, fg_color="transparent")

        ctk.CTkLabel(
            self.panels["preset"], text="排版预设方案",
            font=ctk.CTkFont(family="Microsoft YaHei", size=18, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=30, pady=(30, 10))

        ctk.CTkLabel(
            self.panels["preset"],
            text="选定预设，便是为文稿裁衣。\n正文因缩进知礼，标题以字号传神，文献随悬挂低眉。\n依矩而变，方见雅致。",
            justify="left",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=13),
            text_color=self.COLORS["text_muted"]
        ).pack(anchor="w", padx=30, pady=(0, 20))

        presets = [
            ("默认通用格式",
             "正文：宋体/TNR 12pt，首行缩进 0.85cm，固定行距 20pt\n"
             "一级标题：黑体 16pt 居中，段前段后 10pt\n"
             "二级标题：黑体 14pt，段前段后 6pt\n"
             "参考文献：10.5pt，悬挂缩进 0.74cm"),
            ("学术毕业论文规范",
             "正文：宋体/TNR 12pt，首行缩进 0.85cm，固定行距 22pt\n"
             "一级标题：黑体 16pt 居中，段前段后 12pt\n"
             "二级标题：黑体 14pt，段前段后 8pt\n"
             "参考文献：10.5pt，悬挂缩进 0.74cm"),
            ("IEEE Conference",
             "正文：TNR 10pt，双栏，首行缩进 0.5cm\n"
             "一级标题：TNR 10pt 居中，罗马数字编号\n"
             "二级标题：TNR 10pt 斜体，字母编号\n"
             "参考文献：8pt，编号引用"),
        ]
        saved = load_full_config().get("format_preset", "默认通用格式")
        self.preset_var = ctk.StringVar(value=saved)

        is_custom_locked = load_full_config().get("typography_mode") == "custom"
        self._preset_radios = []

        if is_custom_locked:
            ctk.CTkLabel(
                self.panels["preset"],
                text="当前已启用【自定义模式】，预设方案已锁定。\n如需切换预设，请先前往主界面关闭自定义模式。",
                justify="left",
                font=ctk.CTkFont(family="Microsoft YaHei", size=11),
                text_color=self.COLORS["accent"]
            ).pack(anchor="w", padx=30, pady=(0, 10))

        scroll = ctk.CTkScrollableFrame(
            self.panels["preset"], height=160,
            fg_color=self.COLORS["bg_panel"], corner_radius=8,
            border_width=1, border_color=self.COLORS["primary"],
            scrollbar_button_color=self.COLORS["primary"])
        scroll.pack(fill="x", padx=30, pady=(0, 10))

        for name, desc in presets:
            card = ctk.CTkFrame(scroll, fg_color="transparent")
            card.pack(fill="x", padx=10, pady=(10, 0))
            rb = ctk.CTkRadioButton(
                card, text="", variable=self.preset_var, value=name,
                fg_color=self.COLORS["accent"],
                hover_color=self.COLORS["primary"])
            if is_custom_locked:
                rb.configure(state="disabled")
            self._preset_radios.append(rb)
            rb.pack(side="left", anchor="n")
            info = ctk.CTkFrame(card, fg_color="transparent")
            info.pack(side="left", fill="x", padx=(6, 0))
            ctk.CTkLabel(info, text=name,
                         font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
                         text_color=self.COLORS["text_main"]).pack(anchor="w")
            ctk.CTkLabel(info, text=desc, justify="left",
                         font=ctk.CTkFont(family="Microsoft YaHei", size=11),
                         text_color=self.COLORS["text_muted"]).pack(anchor="w", pady=(4, 8))

        advanced_scroll = ctk.CTkScrollableFrame(
            self.panels["preset"], height=190,
            fg_color=self.COLORS["bg_panel"], corner_radius=8,
            border_width=1, border_color=self.COLORS["primary"],
            scrollbar_button_color=self.COLORS["primary"])
        advanced_scroll.pack(fill="x", padx=30, pady=(0, 10))

        self._build_advanced_strategy_options(advanced_scroll)

        ctk.CTkButton(
            self.panels["preset"], text="保存当前全部设置",
            height=36, font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            fg_color=self.COLORS["primary"], text_color=self.COLORS["text_main"],
            hover_color=self.COLORS["accent"],
            command=lambda: self._on_preset_changed(self.preset_var.get())
        ).pack(pady=(10, 5))

        self.preset_status = ctk.CTkLabel(
            self.panels["preset"], text="",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color=self.COLORS["accent"])
        self.preset_status.pack(pady=(0, 10))

    # ── 文本直排面板 ──

    def _build_text_format_panel(self):
        self.text_format_panel = ctk.CTkFrame(self.main_container, fg_color="transparent")

        ctk.CTkLabel(
            self.text_format_panel, text="文本直排排版",
            font=ctk.CTkFont(family="Microsoft YaHei", size=18, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=30, pady=(30, 5))

        ctk.CTkLabel(
            self.text_format_panel,
            text="在此粘贴或输入带 Markdown / LaTeX 的文本，一键转化为严格格式的 Word 文档。",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=13),
            text_color=self.COLORS["text_muted"]
        ).pack(anchor="w", padx=30, pady=(0, 15))

        self.txt_input = ctk.CTkTextbox(
            self.text_format_panel,
            font=ctk.CTkFont(family="Consolas, Microsoft YaHei", size=12),
            fg_color=self.COLORS["bg_panel"], text_color=self.COLORS["text_main"],
            border_width=1, border_color=self.COLORS["primary"])
        self.txt_input.pack(fill="both", expand=True, padx=30, pady=(0, 15))

        footer = ctk.CTkFrame(self.text_format_panel, fg_color="transparent")
        footer.pack(fill="x", padx=30, side="bottom")

        self.txt_status_lbl = ctk.CTkLabel(
            footer, text="就绪",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color=self.COLORS["text_muted"])
        self.txt_status_lbl.pack(side="left")

        ctk.CTkButton(
            footer, text="排版导出 Word", height=36,
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            fg_color="#6B8E76", text_color="#FDFBF7", hover_color="#5A7A63",
            command=self._action_direct_format
        ).pack(side="right", pady=(0, 15))

    def _action_direct_format(self):
        text_content = self.txt_input.get("1.0", "end-1c").strip()
        if not text_content:
            self.txt_status_lbl.configure(text="请输入有效的文本内容！", text_color="#DF593A")
            return

        from tkinter import filedialog
        out_file = filedialog.asksaveasfilename(
            title="保存排版文件", filetypes=[("Word 文档", "*.docx")],
            defaultextension=".docx")
        if not out_file:
            return

        self.txt_status_lbl.configure(text="正在组织秩序，请稍候...", text_color=self.COLORS["text_main"])

        def _run_task():
            try:
                from core.formula_stats import FormulaRuleStats
                convert_text_to_docx(text_content, out_file, formula_stats=FormulaRuleStats())
                self.after(0, lambda: self.txt_status_lbl.configure(
                    text="排版成功！文件已导出。", text_color="#6B8E76"))
                self.after(0, lambda: os.startfile(out_file))
            except Exception as e:
                self.after(0, lambda: self.txt_status_lbl.configure(
                    text=f"排版失败：{str(e)}", text_color="#DF593A"))

        threading.Thread(target=_run_task, daemon=True).start()

    # ── 主题选择面板 ──
    def _build_theme_tab(self):
        self.panels["theme"] = ctk.CTkFrame(self.main_container, fg_color="transparent")

        ctk.CTkLabel(
            self.panels["theme"], text="界面外观风格",
            font=ctk.CTkFont(family="Microsoft YaHei", size=18, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=30, pady=(30, 10))

        desc = (
            "FormatX 为你编纂了白昼与黑夜的交替法则。\n"
            "随心唤醒下方的气韵，让文字在其专属的秩序中流淌。"
        )
        ctk.CTkLabel(
            self.panels["theme"], text=desc, justify="left",
            font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=13),
            text_color=self.COLORS["text_muted"]
        ).pack(anchor="w", padx=30, pady=(0, 30))

        cards_row = ctk.CTkFrame(self.panels["theme"], fg_color="transparent")
        cards_row.pack(fill="x", padx=30)
        self.theme_cards = {}

        self._create_theme_card(
            parent=cards_row, theme_key="warm",
            title="暖纸", subtitle="白天",
            bg_color="#FDFBF7", title_color="#37352F", subtitle_color="#4E7082",
            active_border="#4E7082"
        )
        self._create_theme_card(
            parent=cards_row, theme_key="moss",
            title="青夜", subtitle="夜间",
            bg_color="#2B312C", title_color="#E3DFD2", subtitle_color="#92978D",
            active_border="#DF593A"
        )

        self._update_card_borders(self.current_theme)

    def _on_preset_changed(self, choice):
        update_config("format_preset", choice)
        # 同时确保 heading_styles 和 table_border 已持久化
        config = load_full_config()
        if "heading_styles" in config:
            update_config("heading_styles", config["heading_styles"])
        mode = self.table_border_var.get() if hasattr(self, "table_border_var") else "three_line"
        update_config("normal_table_border_mode", mode)
        if hasattr(self, "preset_status"):
            border_names = {"three_line": "学术三线表", "full_grid": "全网格线", "keep": "保持原样"}
            hs = config.get("heading_styles", {})
            hn_names = {"cn_lower_chapter": "中文篇章", "arabic_dotted": "多级嵌套",
                        "arabic": "纯数字", "circled": "圆圈"}
            h1 = hn_names.get(hs.get("1", ""), hs.get("1", "默认"))
            h2 = hn_names.get(hs.get("2", ""), hs.get("2", "默认"))
            self.preset_status.configure(
                text=f"已保存: {choice}  |  表格: {border_names.get(mode, mode)}  |  序号: {h1} / {h2}")

    def _build_advanced_strategy_options(self, parent_frame):
        """高级排版微调策略 — 大纲序号 + 表格边框"""
        self.advanced_card = ctk.CTkFrame(
            parent_frame, fg_color=self.COLORS["bg_panel"], corner_radius=8,
            border_width=1, border_color=self.COLORS["primary"])
        self.advanced_card.pack(fill="x", padx=30, pady=(10, 15))

        ctk.CTkLabel(
            self.advanced_card, text="高级排版微调策略",
            font=ctk.CTkFont(family="Microsoft YaHei", size=13, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=15, pady=(12, 6))

        # ── Part A: 大纲序号引擎 ──
        config = load_full_config()
        heading_styles = config.get("heading_styles", {
            "1": "cn_lower_chapter", "2": "arabic_dotted", "3": "arabic", "4": "arabic", "5": "circled"
        })

        options_map = {
            "第几章 (第一章 绪论)": "cn_lower_chapter",
            "第几节 (第一节 背景)": "cn_lower_section",
            "多级嵌套 (1.1 / 1.1.1)": "arabic_dotted",
            "纯数字 (1, 2, 3)": "arabic",
            "补零数字 (01, 02, 03)": "arabic_pad2",
            "中文小写 (一, 二, 三)": "cn_lower",
            "中文大写 (壹, 贰, 叁)": "cn_upper",
            "罗马大写 (I, II, III)": "roman_upper",
            "圆圈序号 (①, ②, ③)": "circled",
            "括号圆圈 (⑴, ⑵, ⑶)": "circled_paren"
        }

        for level in [1, 2, 3, 4, 5]:
            row = ctk.CTkFrame(self.advanced_card, fg_color="transparent")
            row.pack(fill="x", padx=15, pady=4)

            ctk.CTkLabel(row, text=f"{level} 级标题序号:", width=100, anchor="w",
                         font=ctk.CTkFont(family="Microsoft YaHei", size=12),
                         text_color=self.COLORS["text_main"]).pack(side="left")

            menu = ctk.CTkOptionMenu(
                row, values=list(options_map.keys()), height=28,
                fg_color=self.COLORS["primary"], text_color=self.COLORS["text_main"],
                button_color=self.COLORS["primary"],
                button_hover_color=self.COLORS["hover"],
                dropdown_fg_color=self.COLORS["bg_panel"],
                command=lambda val, l=str(level): self._on_heading_style_changed(
                    l, options_map[val]))
            menu.pack(side="left", fill="x", expand=True, padx=(5, 5))

            current_code = heading_styles.get(str(level))
            for display_name, code in options_map.items():
                if code == current_code:
                    menu.set(display_name)

        # ── 分割线 ──
        ctk.CTkFrame(self.advanced_card, height=1,
                     fg_color=self.COLORS["primary"]).pack(
            fill="x", padx=15, pady=(12, 12))

        # ── Part B: 表格边框策略 ──
        ctk.CTkLabel(
            self.advanced_card, text="表格边框样式:", anchor="w",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"),
            text_color=self.COLORS["text_main"]
        ).pack(anchor="w", padx=15, pady=(0, 4))

        self.table_border_var = ctk.StringVar()
        saved = load_full_config().get("normal_table_border_mode", "three_line")
        self.table_border_var.set(saved)

        for text, value in [
            ("学术三线表（顶底粗线 + 中间细线）", "three_line"),
            ("全网格线（所有单元格均有边框）", "full_grid"),
            ("保持原样（不修改源文档表格格式）", "keep"),
        ]:
            ctk.CTkRadioButton(
                self.advanced_card, text=text, variable=self.table_border_var,
                value=value,
                font=ctk.CTkFont(family="Microsoft YaHei", size=11),
                text_color=self.COLORS["text_main"], fg_color=self.COLORS["accent"],
                hover_color=self.COLORS["primary"],
                command=self._on_table_border_changed
            ).pack(anchor="w", padx=25, pady=2)

    def _on_heading_style_changed(self, str_level: str, style_code: str):
        config = load_full_config()
        if "heading_styles" not in config:
            config["heading_styles"] = {}
        config["heading_styles"][str_level] = style_code
        update_config("heading_styles", config["heading_styles"])

    def _on_table_border_changed(self):
        mode = self.table_border_var.get()
        update_config("normal_table_border_mode", mode)
        names = {"three_line": "学术三线表", "full_grid": "全网格线", "keep": "保持原样"}
        if hasattr(self, "preset_status"):
            self.preset_status.configure(text=f"已更新排版策略: 表格使用 {names.get(mode, mode)}")

    def _create_theme_card(self, parent, theme_key, title, subtitle, bg_color,
                           title_color, subtitle_color, active_border):
        border_frame = ctk.CTkFrame(parent, corner_radius=10, border_width=2,
                                    fg_color=bg_color, border_color=bg_color)
        border_frame.pack(side="left", padx=(0, 20))

        card_content = ctk.CTkFrame(border_frame, fg_color=bg_color,
                                    corner_radius=8, width=150, height=85)
        card_content.pack(padx=2, pady=2)
        card_content.pack_propagate(False)

        lbl_title = ctk.CTkLabel(card_content, text=title, text_color=title_color,
                                 font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"))
        lbl_title.pack(pady=(16, 2))

        lbl_sub = ctk.CTkLabel(card_content, text=subtitle, text_color=subtitle_color,
                               font=ctk.CTkFont(family="Microsoft YaHei", size=12))
        lbl_sub.pack(pady=0)

        self.theme_cards[theme_key] = {
            "border": border_frame, "active": active_border, "inactive": bg_color
        }

        for widget in [border_frame, card_content, lbl_title, lbl_sub]:
            widget.bind("<Button-1>", lambda e, k=theme_key: self._on_theme_card_clicked(k))
            try:
                widget.configure(cursor="hand2")
            except Exception:
                pass

    def _update_card_borders(self, active_key):
        for key, widgets in self.theme_cards.items():
            if key == active_key:
                widgets["border"].configure(border_color=widgets["active"])
            else:
                widgets["border"].configure(border_color=widgets["inactive"])

    def _on_theme_card_clicked(self, theme_key):
        if self.current_theme == theme_key:
            return
        self.current_theme = theme_key
        self._update_card_borders(theme_key)
        update_config("ui_theme", theme_key)

        if self.theme_change_callback:
            self.theme_change_callback(theme_key)
        else:
            if theme_key == "moss":
                ctk.set_appearance_mode("dark")
            else:
                ctk.set_appearance_mode("light")

        self.COLORS = self._get_theme_colors(theme_key)
        self.update_idletasks()
        self._apply_theme_to_self()

    def _apply_theme_to_self(self):
        self.configure(fg_color=self.COLORS["bg_main"])
        self.nav_frame.configure(fg_color=self.COLORS["bg_panel"])
        self.main_container.configure(fg_color=self.COLORS["bg_main"])

        def _update_titlebar_dwm():
            try:
                import ctypes
                bgr_hex = 0x00F7FBFD if self.current_theme == "warm" else 0x002C312B
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    hwnd, 35, ctypes.byref(ctypes.c_int(bgr_hex)),
                    ctypes.sizeof(ctypes.c_int))
            except Exception:
                pass
        self.after(200, _update_titlebar_dwm)

        def _update_widget(parent):
            for child in parent.winfo_children():
                try:
                    w_type = child.__class__.__name__
                    try:
                        raw_fg = child.cget("fg_color")
                        fg_str = str(raw_fg).lower()
                    except Exception:
                        fg_str = ""

                    if "CTkLabel" in w_type:
                        if child.cget("text") not in ["暖纸", "白天", "青夜", "夜间"]:
                            if "行文至此" in child.cget("text") or "选定预设" in child.cget("text"):
                                child.configure(text_color=self.COLORS["text_muted"])
                            else:
                                child.configure(text_color=self.COLORS["text_main"])
                    elif "CTkButton" in w_type:
                        if "transparent" not in fg_str:
                            child.configure(fg_color=self.COLORS["primary"])
                        child.configure(text_color=self.COLORS["text_main"],
                                        hover_color=self.COLORS["primary"])
                    elif "CTkEntry" in w_type:
                        child.configure(fg_color=self.COLORS["bg_panel"],
                                        text_color=self.COLORS["text_main"],
                                        border_color=self.COLORS["primary"])
                    elif "CTkOptionMenu" in w_type:
                        child.configure(fg_color=self.COLORS["primary"],
                                        text_color=self.COLORS["text_main"],
                                        button_color=self.COLORS["primary"],
                                        dropdown_fg_color=self.COLORS["bg_panel"])
                    elif "CTkRadioButton" in w_type:
                        child.configure(fg_color=self.COLORS["accent"],
                                        hover_color=self.COLORS["primary"],
                                        text_color=self.COLORS["text_main"])
                    elif "CTkScrollableFrame" in w_type:
                        child.configure(fg_color="transparent",
                                        scrollbar_button_color=self.COLORS["primary"],
                                        scrollbar_button_hover_color=self.COLORS["accent"])
                    elif "CTkFrame" in w_type:
                        is_theme_card = any(c in fg_str for c in ["#fdfbf7", "#2b312c"])
                        if "transparent" not in fg_str and not is_theme_card:
                            child.configure(fg_color=self.COLORS["bg_main"])
                except Exception as e:
                    print(f"[Theme Warn] {w_type} failed: {e}")
                if "Frame" in w_type and child.winfo_children():
                    _update_widget(child)

        _update_widget(self)
        self.update_idletasks()

        if "theme" in self.panels:
            for widget in self.panels["theme"].winfo_children():
                if ("CTkLabel" in widget.__class__.__name__
                        and "FormatX 为你编纂了" in widget.cget("text")):
                    widget.configure(text_color=self.COLORS["text_muted"])

    # ── API 配置面板（主从联动布局：左栏供应商列表 + 右栏参数详情）──
    def _build_api_panel(self):
        from ai_assist.provider_presets import API_PROVIDER_PRESETS

        self.api_panel = ctk.CTkFrame(self.main_container, fg_color="transparent")

        self.full_config = load_full_config()
        if "providers" not in self.full_config:
            self.full_config["providers"] = {}
        self.selected_provider_id = self.full_config.get(
            "active_provider", API_PROVIDER_PRESETS[0]["id"])

        ctk.CTkLabel(self.api_panel, text="模型供应商",
                     font=ctk.CTkFont(family="Microsoft YaHei", size=18, weight="bold"),
                     text_color=self.COLORS["text_main"]).pack(anchor="w", pady=(0, 10))

        layout_box = ctk.CTkFrame(self.api_panel, fg_color="transparent")
        layout_box.pack(fill="both", expand=True, pady=5)

        left_sidebar = ctk.CTkScrollableFrame(
            layout_box, width=170, fg_color=self.COLORS["bg_panel"], corner_radius=6)
        left_sidebar.pack(side="left", fill="y", padx=(0, 15))

        right_container = ctk.CTkFrame(layout_box, fg_color="transparent")
        right_container.pack(side="left", fill="both", expand=True)

        ctk.CTkLabel(right_container, text="接口节点 (Base URL)",
                     font=ctk.CTkFont(family="Microsoft YaHei", size=12),
                     text_color=self.COLORS["text_main"]).pack(anchor="w", pady=(0, 2))
        self.ui_url_entry = ctk.CTkEntry(
            right_container, height=32,
            fg_color=self.COLORS["bg_panel"], text_color=self.COLORS["text_main"],
            border_color=self.COLORS["primary"])
        self.ui_url_entry.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(right_container, text="身份凭证 (API Key)",
                     font=ctk.CTkFont(family="Microsoft YaHei", size=12),
                     text_color=self.COLORS["text_main"]).pack(anchor="w", pady=(0, 2))
        self.ui_key_entry = ctk.CTkEntry(
            right_container, height=32, show="*",
            fg_color=self.COLORS["bg_panel"], text_color=self.COLORS["text_main"],
            border_color=self.COLORS["primary"])
        self.ui_key_entry.pack(fill="x", pady=(0, 12))

        ctk.CTkLabel(right_container, text="目标模型 (Model)",
                     font=ctk.CTkFont(family="Microsoft YaHei", size=12),
                     text_color=self.COLORS["text_main"]).pack(anchor="w", pady=(0, 2))
        model_row = ctk.CTkFrame(right_container, fg_color="transparent")
        model_row.pack(fill="x", pady=(0, 12))

        self.ui_model_menu = ctk.CTkOptionMenu(
            model_row, height=32, corner_radius=8,
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            fg_color=self.COLORS["bg_panel"], text_color=self.COLORS["text_main"],
            button_color=self.COLORS["primary"], button_hover_color=self.COLORS["hover"],
            dropdown_fg_color=self.COLORS["bg_panel"],
            dropdown_text_color=self.COLORS["text_main"],
            dropdown_hover_color=self.COLORS["primary"])
        self.ui_model_menu.pack(side="left", fill="x", expand=True)

        self.ui_fetch_btn = ctk.CTkButton(
            model_row, text="拉取列表", width=85, height=32,
            fg_color=self.COLORS["primary"], text_color=self.COLORS["text_main"],
            hover_color=self.COLORS["hover"],
            command=lambda: self._exec_fetch_models())
        self.ui_fetch_btn.pack(side="right", padx=(10, 0))

        self.ui_api_status = ctk.CTkLabel(
            right_container, text="就绪",
            font=ctk.CTkFont(family="Microsoft YaHei", size=11),
            text_color=self.COLORS["text_muted"], wraplength=380, height=30)
        self.ui_api_status.pack(fill="x", pady=5)

        bottom_action = ctk.CTkFrame(right_container, fg_color="transparent")
        bottom_action.pack(fill="x", side="bottom", pady=(10, 0))

        ctk.CTkButton(
            bottom_action, text="测试连接", width=90, height=32,
            font=ctk.CTkFont(family="Microsoft YaHei", size=12),
            fg_color=self.COLORS["primary"], text_color=self.COLORS["text_main"],
            hover_color=self.COLORS["hover"],
            command=lambda: self._exec_test_conn()).pack(side="left")

        ctk.CTkButton(
            bottom_action, text="保存当前配置", width=110, height=32,
            font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"),
            fg_color="#6B8E76", text_color="#FDFBF7", hover_color="#5A7A63",
            command=lambda: self._exec_save_settings()).pack(side="right")

        self.provider_buttons = {}
        for p in API_PROVIDER_PRESETS:
            btn = ctk.CTkButton(
                left_sidebar, text=p["label"], anchor="w", height=32,
                fg_color="transparent", text_color=self.COLORS["text_main"],
                hover_color=self.COLORS["primary"],
                command=lambda _id=p["id"]: self._on_provider_selected_switched(_id))
            btn.pack(fill="x", pady=2, padx=4)
            self.provider_buttons[p["id"]] = btn

        self._on_provider_selected_switched(self.selected_provider_id)

    def _on_provider_selected_switched(self, provider_id):
        from ai_assist.provider_presets import API_PROVIDER_PRESETS
        self.selected_provider_id = provider_id

        for p_id, btn in self.provider_buttons.items():
            if p_id == provider_id:
                btn.configure(fg_color=self.COLORS["primary"],
                             font=ctk.CTkFont(family="Microsoft YaHei", size=12, weight="bold"))
            else:
                btn.configure(fg_color="transparent",
                             font=ctk.CTkFont(family="Microsoft YaHei", size=12))

        preset = next((p for p in API_PROVIDER_PRESETS if p["id"] == provider_id),
                      API_PROVIDER_PRESETS[0])
        saved_prov_data = self.full_config.get("providers", {}).get(provider_id, {})

        self.ui_url_entry.configure(state="normal")
        self.ui_url_entry.delete(0, "end")
        self.ui_url_entry.insert(0, saved_prov_data.get("base_url", preset["default_url"]))
        if provider_id != "custom":
            self.ui_url_entry.configure(state="readonly")

        self.ui_key_entry.delete(0, "end")
        self.ui_key_entry.insert(0, saved_prov_data.get("api_key", ""))

        stored_models = saved_prov_data.get("models", [])
        default_models = preset.get("default_models", [])
        merged = list(dict.fromkeys(default_models + stored_models))
        self.ui_model_menu.configure(values=merged if merged else ["无可用模型"])

        active_model_in_cfg = self.full_config.get("active_model", "")
        if active_model_in_cfg in merged:
            self.ui_model_menu.set(active_model_in_cfg)
        elif merged:
            self.ui_model_menu.set(merged[0])

        self.ui_api_status.configure(
            text=f"当前预览供应商: {preset['label']}",
            text_color=self.COLORS["text_muted"])

    def _exec_test_conn(self):
        url = self.ui_url_entry.get().strip()
        key = self.ui_key_entry.get().strip()
        model = self.ui_model_menu.get()

        if not url or not key:
            self.ui_api_status.configure(
                text="接口节点或凭证不能为空", text_color="#DF593A")
            return

        from ai_assist.provider_presets import API_PROVIDER_PRESETS
        preset = next((p for p in API_PROVIDER_PRESETS
                      if p["id"] == self.selected_provider_id), None)
        api_type = preset["api_type"] if preset else "openai-completions"

        self.ui_api_status.configure(
            text="正在向远端节点发起连通性验证...",
            text_color=self.COLORS["text_main"])
        self.update_idletasks()

        def _run():
            ok, msg = test_provider_connection(url, key, model, api_type)
            color = "#6B8E76" if ok else "#DF593A"
            self.after(0, lambda: self.ui_api_status.configure(
                text=f"测试反馈: {msg}", text_color=color))

        threading.Thread(target=_run, daemon=True).start()

    def _exec_fetch_models(self):
        url = self.ui_url_entry.get().strip()
        key = self.ui_key_entry.get().strip()

        if not url or not key:
            self.ui_api_status.configure(
                text="填入基础凭证后方可拉取模型字典", text_color="#DF593A")
            return

        from ai_assist.provider_presets import API_PROVIDER_PRESETS
        preset = next((p for p in API_PROVIDER_PRESETS
                      if p["id"] == self.selected_provider_id), None)
        api_type = preset["api_type"] if preset else "openai-completions"

        self.ui_api_status.configure(
            text="正在全网检索并提取该节点的模型大纲...",
            text_color=self.COLORS["text_main"])
        self.update_idletasks()

        def _run():
            try:
                model_list = fetch_remote_models(url, key, api_type)
                def _success():
                    self.ui_model_menu.configure(values=model_list)
                    self.ui_model_menu.set(model_list[0])
                    self.ui_api_status.configure(
                        text=f"成功加载 {len(model_list)} 个模型实体结构，请保存配置",
                        text_color="#6B8E76")
                self.after(0, _success)
            except Exception as e:
                self.after(0, lambda: self.ui_api_status.configure(
                    text=str(e), text_color="#DF593A"))

        threading.Thread(target=_run, daemon=True).start()

    def _exec_save_settings(self):
        from ai_assist.provider_presets import API_PROVIDER_PRESETS
        p_id = self.selected_provider_id
        preset = next((p for p in API_PROVIDER_PRESETS if p["id"] == p_id),
                      API_PROVIDER_PRESETS[0])

        current_models = list(self.ui_model_menu.cget("values"))
        if "无可用模型" in current_models:
            current_models = []

        self.full_config["providers"][p_id] = {
            "base_url": self.ui_url_entry.get().strip(),
            "api_key": self.ui_key_entry.get().strip(),
            "api_type": preset["api_type"],
            "models": current_models
        }
        self.full_config["active_provider"] = p_id
        self.full_config["active_model"] = self.ui_model_menu.get()
        self.full_config["base_url"] = self.ui_url_entry.get().strip()
        self.full_config["api_key"] = self.ui_key_entry.get().strip()
        self.full_config["model"] = self.ui_model_menu.get()

        save_full_config(self.full_config)
        self.ui_api_status.configure(
            text="全局模型矩阵配置已成功写入磐石！", text_color="#6B8E76")

    # ── 智能提取面板 ──
    def _build_ocr_panel(self):
        """构建诗意化 OCR 灵感破译仓"""
        self.ocr_panel = ctk.CTkFrame(self.main_container, fg_color="transparent")

        font_title = ctk.CTkFont(family="LXGW WenKai, Microsoft YaHei", size=18, weight="bold")
        font_step_icon = ctk.CTkFont(family="Segoe UI Emoji, Microsoft YaHei", size=16)
        font_step_text = ctk.CTkFont(family="LXGW WenKai, Microsoft YaHei", size=13, weight="bold")
        font_poetic = ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=13)
        font_btn_main = ctk.CTkFont(family="LXGW WenKai, Microsoft YaHei", size=15, weight="bold")
        font_hint = ctk.CTkFont(family="Microsoft YaHei", size=11)

        COLOR_BRAND_DARK = "#2B312C"
        COLOR_BRAND_TEXT = "#E3DFD2"
        COLOR_GHOST_BORDER = "#A3B5A7"

        # 页面主标题
        ctk.CTkLabel(
            self.ocr_panel, text="灵感破译仓",
            font=font_title, text_color=_TEXT
        ).pack(anchor="w", pady=(0, 15))

        # ── 1. 图形化引导 (3步流) ──
        steps_frame = ctk.CTkFrame(self.ocr_panel, fg_color="transparent")
        steps_frame.pack(fill="x", pady=(0, 15))

        def _add_step(parent, icon, text, is_last=False):
            ctk.CTkLabel(parent, text=icon, font=font_step_icon, text_color="#A3B5A7").pack(side="left")
            ctk.CTkLabel(parent, text=text, font=font_step_text, text_color=_TEXT).pack(side="left", padx=(4, 0))
            if not is_last:
                ctk.CTkLabel(parent, text=" --> ", font=font_hint, text_color=_TEXT_MUTED).pack(side="left", padx=8)

        _add_step(steps_frame, "[ ]", "截图 / 选取")
        _add_step(steps_frame, "[*]", "智能破译")
        _add_step(steps_frame, "[ ]", "秩序流转", is_last=True)

        # ── 2. 诗意文案 ──
        poetic_text = (
            "行文至此，免受繁琐。\n\n"
            "将那些散落于教科书、PDF 或网页中的复杂公式与混乱图表，\n"
            "随手截一张图暂存于剪贴板，或是保留为一篇孤立的画像。\n"
            "唤醒下方的引擎，FormatX 将在底层为你悉心破译，\n"
            "让它们化作完美的秩序，流淌进主编辑区。"
        )
        ctk.CTkLabel(
            self.ocr_panel, text=poetic_text, justify="left",
            font=font_poetic, text_color="#5D5854"
        ).pack(anchor="w", pady=(0, 25))

        # ── 3. 双子星按钮区 ──
        self.btn_container = ctk.CTkFrame(self.ocr_panel, fg_color="transparent")
        self.btn_container.pack(fill="x", pady=(0, 5))

        self.btn_clip = ctk.CTkButton(
            self.btn_container, text="破译剪贴板中的灵感",
            height=46, font=font_btn_main,
            fg_color=COLOR_BRAND_DARK, text_color=COLOR_BRAND_TEXT, hover_color="#3A423B",
            command=self._action_clipboard_ocr
        )
        self.btn_clip.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_local = ctk.CTkButton(
            self.btn_container, text="唤醒本地珍藏的画像",
            height=46, font=font_btn_main,
            fg_color="transparent", text_color=_TEXT,
            border_width=1, border_color=COLOR_GHOST_BORDER, hover_color=_HOVER,
            command=self._action_local_file_ocr
        )
        self.btn_local.pack(side="left", fill="x", expand=True)

        ctk.CTkLabel(
            self.ocr_panel, text="支持快捷键 Ctrl + V 直接触发破译",
            font=font_hint, text_color="#A9A19A"
        ).pack(pady=(0, 20))

        self.bind("<Control-v>", lambda e: self._action_clipboard_ocr())

        # ── 4. 动态预览区 / 拖拽响应区 ──
        self.drop_zone = ctk.CTkFrame(
            self.ocr_panel, fg_color="#F0EEE9",
            border_width=1, border_color="#D1CDC5", corner_radius=10
        )
        self.drop_zone.pack(fill="both", expand=True)

        self.monitor_lbl = ctk.CTkLabel(
            self.drop_zone,
            text="等待剪贴板中装入新的灵感...\n\n亦可将图片直接投入此间，静候其归位……",
            font=font_poetic, text_color="#9A928B", justify="center"
        )
        self.monitor_lbl.pack(expand=True)

        # ── 5. 加载遮罩与 Spinner ──
        self._overlay = ctk.CTkFrame(self.ocr_panel, fg_color=_BG_MAIN)

        self._ov_phase = 0

        ctk.CTkLabel(
            self._overlay,
            text="正在聆听剪贴板中的动态...",
            font=ctk.CTkFont(family="Microsoft YaHei", size=16, weight="bold"),
            text_color=_TEXT
        ).pack(pady=(60, 10))

        ctk.CTkLabel(
            self._overlay,
            text="FormatX 正在底层为你悉心破译、重新雕琢，请稍候...",
            font=ctk.CTkFont(family="Microsoft YaHei", size=12, slant="italic"),
            text_color=_TEXT_MUTED
        ).pack(pady=0)

        self._ov_spinner = ctk.CTkLabel(
            self._overlay, text=" *",
            font=ctk.CTkFont(family="Consolas", size=16),
            text_color="#A3B5A7"
        )
        self._ov_spinner.pack(pady=15)

        self._monitor_clipboard()

    def _monitor_clipboard(self):
        if not self.winfo_exists():
            return
        if getattr(self, '_pause_monitor', False):
            self.after(1500, self._monitor_clipboard)
            return
        try:
            img_bytes = get_image_from_clipboard()
            if img_bytes:
                self.monitor_lbl.configure(
                    text="察觉到剪贴板中存在图像，已准备就绪。\n\n亦可将图片直接投入此间，静候其归位……",
                    text_color="#6B8E76"
                )
            else:
                self.monitor_lbl.configure(
                    text="等待剪贴板中装入新的灵感...\n\n亦可将图片直接投入此间，静候其归位……",
                    text_color="#9A928B"
                )
        except Exception:
            pass
        self.after(1500, self._monitor_clipboard)

    def _show_overlay(self):
        if hasattr(self, 'monitor_lbl') and self.monitor_lbl.winfo_exists():
            self.monitor_lbl.pack_forget()
        if hasattr(self, 'btn_container') and self.btn_container.winfo_exists():
            self.btn_container.pack_forget()
        if hasattr(self, 'drop_zone') and self.drop_zone.winfo_exists():
            self.drop_zone.pack_forget()
        if hasattr(self, '_overlay') and self._overlay.winfo_exists():
            self._overlay.pack(fill="both", expand=True, pady=20)
            self._pulse_overlay()

    def _hide_overlay(self):
        if hasattr(self, '_overlay') and self._overlay.winfo_exists():
            self._overlay.pack_forget()
        if hasattr(self, 'monitor_lbl') and self.monitor_lbl.winfo_exists():
            self.monitor_lbl.pack(pady=10, fill="x", expand=True)
        if hasattr(self, 'btn_container') and self.btn_container.winfo_exists():
            self.btn_container.pack(fill="x", pady=(0, 5))
        if hasattr(self, 'drop_zone') and self.drop_zone.winfo_exists():
            self.drop_zone.pack(fill="both", expand=True, padx=20, pady=15)

    def _pulse_overlay(self):
        import math
        self._ov_phase += 1
        t = self._ov_phase
        spin_chars = ['*', '**', '***', '**', '*', '.', '..', '...']
        char = spin_chars[t % len(spin_chars)]
        self._ov_spinner.configure(text=f" {char}")
        self._ov_pulse_id = self.after(120, self._pulse_overlay)

    def _action_clipboard_ocr(self):
        self.monitor_lbl.configure(text="正在截获剪贴板中的图像...", text_color=_TEXT)
        img_bytes = get_image_from_clipboard()
        if not img_bytes:
            self.monitor_lbl.configure(text="剪贴板中空空如也，请先截图复制", text_color="#DF593A")
            return
        b64_data = image_to_base64(img_bytes)
        self._process_image_b64(b64_data, "剪贴板图像")

    def _action_local_file_ocr(self, specified_path=None):
        if specified_path:
            path = specified_path
        else:
            from tkinter import filedialog
            self._pause_monitor = True
            try:
                path = filedialog.askopenfilename(
                    parent=self,
                    title="唤醒本地珍藏的图片",
                    filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp"), ("所有文件", "*.*")]
                )
            finally:
                self._pause_monitor = False
        if not path or not os.path.exists(path):
            return
        self.monitor_lbl.configure(text="正在唤醒本地画像...", text_color=_TEXT)
        try:
            import base64 as b64_mod
            with open(path, "rb") as f:
                b64_data = b64_mod.b64encode(f.read()).decode("utf-8")
            self._process_image_b64(b64_data, f"本地画像 ({os.path.basename(path)})")
        except Exception as e:
            self.monitor_lbl.configure(text=f"读取图片失败: {str(e)}", text_color="#DF593A")

    def _on_drop_image(self, event):
        path = event.data.strip().strip('{}')
        if os.path.isfile(path) and path.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
            self._action_local_file_ocr(specified_path=path)
        else:
            self._log("放入的文件格式不正确，请投入图片文件")

    def _process_image_b64(self, b64_data, source_name=""):
        self._show_overlay()
        self.btn_clip.configure(state="disabled")
        self.btn_local.configure(state="disabled")

        def _run():
            try:
                url = self.ui_url_entry.get().strip()
                key = self.ui_key_entry.get().strip()
                mod = self.ui_model_menu.get()

                from ai_assist.provider_presets import API_PROVIDER_PRESETS
                preset = next((p for p in API_PROVIDER_PRESETS
                              if p["id"] == self.selected_provider_id), None)
                api_type = preset["api_type"] if preset else "openai-completions"

                result = extract_text_from_image(
                    b64_data, api_key=key, model=mod, base_url=url,
                    api_type=api_type)

                def _on_ui_thread():
                    self._hide_overlay()
                    self.btn_clip.configure(state="normal")
                    self.btn_local.configure(state="normal")
                    try:
                        self._fill(result)
                        self.monitor_lbl.configure(
                            text=f"{source_name} 破译成功！已化为秩序流淌进 Word。",
                            text_color="#6B8E76"
                        )
                        self.after(1800, self._on_close)
                    except Exception as err:
                        self.monitor_lbl.configure(
                            text=f"注入 Word 时遭遇波折: {str(err)}", text_color="#DF593A")

                self.after(0, _on_ui_thread)
            except Exception as e:
                def _err_thread():
                    self._hide_overlay()
                    self.btn_clip.configure(state="normal")
                    self.btn_local.configure(state="normal")
                    self.monitor_lbl.configure(
                        text=f"破译失败，灵感逃散: {str(e)}", text_color="#DF593A")
                self.after(0, _err_thread)

        import threading
        threading.Thread(target=_run, daemon=True).start()
