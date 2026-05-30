"""AI 学术创作舱 —— 聊天流式面板（富文本渲染）"""
import os
import re
import customtkinter as ctk

from core.config_manager import get_active_model_settings, load_full_config
from ai_assist.stream_buffer import StreamStateBuffer
from ai_assist.stream_worker import StreamWorker
from core.prompt_compiler import PromptCompiler
from core.prompts import SYSTEM_PROMPT_ACADEMIC, SYSTEM_PROMPT_NOVICE
from format_conversion import convert_text_to_docx, reformat_docx


def _configure_markdown_tags(text_widget, colors):
    """为 Tk Text 注册富文本样式与隐身标签"""
    text_widget.tag_config("hidden", elide=True)
    text_widget.tag_config("h1", font=("Microsoft YaHei", 18, "bold"),
                           foreground=colors["text"])
    text_widget.tag_config("h2", font=("Microsoft YaHei", 16, "bold"),
                           foreground=colors["text"])
    text_widget.tag_config("h3", font=("Microsoft YaHei", 14, "bold"),
                           foreground=colors["text"])
    text_widget.tag_config("bold", font=("Microsoft YaHei", 13, "bold"))
    text_widget.tag_config("math", foreground="#4E7082")


def _apply_markdown_tags(text_widget, colors):
    """扫描全文：隐藏 Markdown 符号 (# ## **)，应用富文本样式"""
    for tag in ("h1", "h2", "h3", "bold", "hidden", "math"):
        text_widget.tag_remove(tag, "1.0", "end")

    content = text_widget.get("1.0", "end-1c")

    # ── 标题：隐藏 # 符号，加粗加大 ──
    for m in re.finditer(r'^(#{1,3})\s+(.+)', content, re.MULTILINE):
        hash_end = m.start(1) + len(m.group(1)) + 1  # 跳过 "# "
        text_widget.tag_add("hidden", f"1.0 + {m.start(1)} chars",
                            f"1.0 + {hash_end} chars")
        tag_name = {"#": "h1", "##": "h2", "###": "h3"}.get(m.group(1), "h3")
        text_widget.tag_add(tag_name, f"1.0 + {hash_end} chars",
                            f"1.0 + {m.end()} chars")

    # ── 加粗：隐藏 ** 符号 ──
    for m in re.finditer(r'\*\*(.+?)\*\*', content):
        text_widget.tag_add("hidden", f"1.0 + {m.start()} chars",
                            f"1.0 + {m.start() + 2} chars")
        text_widget.tag_add("bold", f"1.0 + {m.start() + 2} chars",
                            f"1.0 + {m.end() - 2} chars")
        text_widget.tag_add("hidden", f"1.0 + {m.end() - 2} chars",
                            f"1.0 + {m.end()} chars")

    # ── 公式：特殊颜色 ──
    for m in re.finditer(r'\$\$(.+?)\$\$', content, re.DOTALL):
        text_widget.tag_add("math", f"1.0 + {m.start()} chars",
                            f"1.0 + {m.end()} chars")
    for m in re.finditer(r'(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)', content):
        text_widget.tag_add("math", f"1.0 + {m.start()} chars",
                            f"1.0 + {m.end()} chars")


def _update_markdown_textbox(tb, text, colors):
    """刷新 CTkTextbox 内容、应用 Markdown 富文本、自适应高度"""
    tb.configure(state="normal")
    tb.delete("1.0", "end")
    tb.insert("1.0", text)
    raw = getattr(tb, "_textbox", None) or tb
    _configure_markdown_tags(raw, colors)
    _apply_markdown_tags(raw, colors)
    est_lines = text.count('\n') + sum(len(ln) // 35 + 1 for ln in text.split('\n'))
    target_h = min(max(est_lines * 24, 60), 800)
    tb.configure(height=target_h)
    tb.configure(state="disabled")


def build_ai_chat_panel(parent, initial_colors, log_callback):
    COLORS = dict(initial_colors)
    log = log_callback or (lambda m: None)

    page = ctk.CTkFrame(parent, fg_color=COLORS["bg_main"])

    # ── 顶部栏 ──
    top_bar = ctk.CTkFrame(page, fg_color="transparent", height=40)
    top_bar.pack(fill="x", padx=10, pady=(10, 0))
    top_bar.pack_propagate(False)

    title_lbl = ctk.CTkLabel(top_bar, text="✨ AI 创作舱",
                             font=("Microsoft YaHei", 15, "bold"),
                             text_color=COLORS["text"])
    title_lbl.pack(side="left", padx=10)

    mode_var = ctk.StringVar(value="内容质量模式")
    mode_menu = ctk.CTkOptionMenu(
        top_bar, values=["内容质量模式", "查重降重模式"],
        variable=mode_var, width=130, height=28,
        font=("Microsoft YaHei", 11),
        fg_color=COLORS["bg_panel"], text_color=COLORS["text"],
        button_color=COLORS["bg_panel"], button_hover_color=COLORS["hover"])
    mode_menu.pack(side="right", padx=10)

    # ── 排版状态感知栏 ──
    status_bar = ctk.CTkFrame(page, fg_color="transparent")
    status_bar.pack(fill="x", padx=60, pady=(5, 15))

    lbl_status_icon = ctk.CTkLabel(
        status_bar, text="⚡ 搭载约束：",
        font=("Microsoft YaHei", 12, "bold"), text_color="#2D5A27")
    lbl_status_icon.pack(side="left")

    lbl_active_preset = ctk.CTkLabel(
        status_bar, text="", font=("Microsoft YaHei", 12),
        text_color="#2D5A27")
    lbl_active_preset.pack(side="left", padx=5)

    def _sync_format_state():
        config = load_full_config()
        mode = config.get("typography_mode", "preset")
        preset_name = config.get("format_preset", "默认通用格式")

        if mode == "custom":
            lbl_active_preset.configure(
                text=f"[{preset_name}] + [🛠️ 自定义微调接管]",
                text_color="#D97706")
        else:
            lbl_active_preset.configure(
                text=f"[{preset_name}] (引擎自动排版)",
                text_color="#2D5A27")

    _sync_format_state()
    page._sync_format_state = _sync_format_state

    from core.state.mode_manager import mode_manager
    from core.state.constants import TypographyMode

    def _on_mode_changed(current_mode):
        config = load_full_config()
        preset_name = config.get("format_preset", "默认通用格式")
        if current_mode == TypographyMode.CUSTOM:
            lbl_active_preset.configure(
                text=f"[{preset_name}] + [🛠️ 自定义微调接管]",
                text_color="#D97706")
        elif current_mode == TypographyMode.CLONE:
            lbl_active_preset.configure(
                text=f"[🧬 模板克隆解析中]",
                text_color="#2B7A78")
        else:
            lbl_active_preset.configure(
                text=f"[{preset_name}] (引擎自动排版)",
                text_color="#2D5A27")

    mode_manager.subscribe(_on_mode_changed)
    _on_mode_changed(mode_manager.get_mode())

    # ── 聊天滚动区 ──
    chat_scroll = ctk.CTkScrollableFrame(
        page, fg_color=COLORS["bg_main"], corner_radius=0,
        scrollbar_button_color=COLORS["bg_main"],
        scrollbar_button_hover_color=COLORS["primary"])
    chat_scroll.pack(fill="both", expand=True, padx=60, pady=10)

    chat_inner = ctk.CTkFrame(chat_scroll, fg_color="transparent")
    chat_inner.pack(fill="both", expand=True)

    welcome_frame = ctk.CTkFrame(chat_inner, fg_color="transparent")
    welcome_frame.pack(fill="both", expand=True, pady=140)

    greeting_lbl = ctk.CTkLabel(
        welcome_frame,
        text="将凌乱的思绪投入此间\n\n化作排版严谨的篇章",
        font=ctk.CTkFont(family="LXGW WenKai, 楷体, Microsoft YaHei", size=16),
        text_color=COLORS["text_muted"], justify="center")
    greeting_lbl.pack()

    # ── 底部胶囊输入区 ──
    bottom_container = ctk.CTkFrame(page, fg_color="transparent")
    bottom_container.pack(fill="x", side="bottom", padx=60, pady=(0, 20))

    input_pill = ctk.CTkFrame(bottom_container, fg_color=COLORS["bg_panel"],
                               corner_radius=20, border_width=1,
                               border_color=COLORS["primary"])
    input_pill.pack(fill="x", expand=True, ipady=5)

    chat_input = ctk.CTkTextbox(
        input_pill, height=45, fg_color="transparent",
        text_color=COLORS["text"], font=("Microsoft YaHei", 13), border_width=0)
    chat_input.pack(side="left", fill="both", expand=True, padx=(20, 10), pady=10)

    send_btn = ctk.CTkButton(
        input_pill, text="发送 ↵", width=70, height=36, corner_radius=15,
        font=("Microsoft YaHei", 12, "bold"),
        fg_color=COLORS["primary"], text_color=COLORS["text"],
        hover_color=COLORS["hover"],
        command=lambda: _chat_send())
    send_btn.pack(side="right", padx=(0, 15), pady=10)

    # ── 主题回调 ──
    dynamic_ui_elements = []

    def apply_theme_fn(new_colors):
        COLORS.update(new_colors)
        page.configure(fg_color=COLORS["bg_main"])
        chat_scroll.configure(fg_color=COLORS["bg_main"])
        title_lbl.configure(text_color=COLORS["text"])
        mode_menu.configure(fg_color=COLORS["bg_panel"], text_color=COLORS["text"],
                            button_color=COLORS["bg_panel"],
                            button_hover_color=COLORS["hover"])
        greeting_lbl.configure(text_color=COLORS["text_muted"])
        input_pill.configure(fg_color=COLORS["bg_panel"],
                             border_color=COLORS["primary"])
        chat_input.configure(text_color=COLORS["text"])
        send_btn.configure(fg_color=COLORS["primary"], text_color=COLORS["text"],
                           hover_color=COLORS["hover"])
        for el, el_type in dynamic_ui_elements:
            try:
                if el_type == "text":
                    el.configure(text_color=COLORS["text"])
                elif el_type == "text_muted":
                    el.configure(text_color=COLORS["text_muted"])
                elif el_type == "card":
                    el.configure(fg_color=COLORS["bg_panel"])
                elif el_type == "dispatcher":
                    if hasattr(el, "update_theme"):
                        el.update_theme(COLORS)
            except Exception:
                pass

    # ── 状态 ──
    chat_messages = []
    processing = [False]
    chat_input.bind("<Return>", lambda e: _chat_send() if not e.state & 0x1 else None)

    def _scroll_to_bottom():
        chat_scroll._parent_canvas.yview_moveto(1.0)

    def _chat_send():
        prompt = chat_input.get("1.0", "end-1c").strip()
        if not prompt or processing[0]:
            return
        chat_input.delete("1.0", "end")
        processing[0] = True

        if welcome_frame.winfo_exists():
            welcome_frame.pack_forget()

        _add_bubble(chat_inner, prompt, is_user=True, colors=COLORS,
                    element_tracker=dynamic_ui_elements, rich=False)
        page.after(50, _scroll_to_bottom)

        target_cfg = get_active_model_settings()
        if not target_cfg.get("api_key"):
            _add_bubble(chat_inner, "请先在设置中心的 [API 配置] 中填写并保存 API Key。",
                        is_user=False, colors=COLORS,
                        element_tracker=dynamic_ui_elements, rich=False)
            processing[0] = False
            return

        from ai_assist.block_dispatcher import AIMessageBlockDispatcher
        ai_dispatcher = AIMessageBlockDispatcher(
            chat_inner,
            colors=COLORS,
            on_update_markdown=_update_markdown_textbox,
            on_export_docx=_do_export)
        ai_dispatcher.pack(fill="x", pady=(5, 15), anchor="w")
        dynamic_ui_elements.append((ai_dispatcher, "dispatcher"))

        def on_state_flush(state: dict):
            page.after(0, lambda: ai_dispatcher.render_state(state))
            page.after(0, _scroll_to_bottom)

        def on_stream_complete(final_state: dict):
            def _cleanup():
                if final_state["draft_text"]:
                    ai_dispatcher.mount_action_card(final_state["draft_text"], log)
                elif final_state["display_text"] and final_state["display_text"].strip():
                    ai_dispatcher.mount_action_card(final_state["display_text"], log)
                elif not final_state["display_text"] and not final_state["reasoning_text"]:
                    ai_dispatcher.render_state({
                        "display_text": "(AI 返回为空，请重试)",
                        "has_reasoning": False,
                        "thinking_finished": True
                    })

                saved = final_state["display_text"] if final_state["display_text"].strip() else final_state["draft_text"]
                if saved.strip():
                    chat_messages.append({"role": "assistant", "content": saved})

                processing[0] = False
                _scroll_to_bottom()

            page.after(0, _cleanup)

        def on_stream_error(error_msg: str):
            def _err():
                ai_dispatcher.render_state({
                    "display_text": f"生成中断: {error_msg}",
                    "has_reasoning": False,
                    "thinking_finished": True
                })
                processing[0] = False
            page.after(0, _err)

        _sync_format_state()

        system_prompt = (SYSTEM_PROMPT_ACADEMIC if mode_var.get() == "内容质量模式"
                         else SYSTEM_PROMPT_NOVICE)
        system_prompt += "\n" + PromptCompiler.generate_formatting_system_prompt()
        chat_messages.append({"role": "user", "content": prompt})
        msgs = [{"role": "system", "content": system_prompt}] + chat_messages[-10:]

        buffer = StreamStateBuffer(on_flush_callback=on_state_flush, flush_interval=0.2)
        worker = StreamWorker(target_cfg, msgs, buffer, on_complete=on_stream_complete, on_error=on_stream_error)
        worker.start()

    return page, apply_theme_fn


def _add_bubble(parent, text, is_user, colors, element_tracker, rich=True):
    container = ctk.CTkFrame(parent, fg_color="transparent")
    container.pack(fill="x", pady=5, anchor="e" if is_user else "w")

    bg = colors["primary"] if is_user else "transparent"
    bubble = ctk.CTkFrame(container, corner_radius=12, fg_color=bg)
    bubble.pack(side="right" if is_user else "left", padx=10)

    lines = text.count('\n') + 1
    h = min(max(lines * 22, 28), 400)
    font = ("Microsoft YaHei", 13) if is_user else ctk.CTkFont(
        family="LXGW WenKai, 楷体, Microsoft YaHei", size=13)

    tb = ctk.CTkTextbox(bubble, height=h, fg_color="transparent",
                         text_color=colors["text"], font=font,
                         wrap="word", border_width=0, activate_scrollbars=False)
    tb.insert("1.0", text)
    tb.configure(state="disabled")
    tb.pack(padx=15, pady=10)
    element_tracker.append((tb, "text"))


def _do_export(draft_text, log):
    import threading
    from datetime import datetime
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(os.path.expanduser("~"), "Desktop",
                        f"FormatX_AI_Draft_{ts}.docx")
    raw_path = path + ".raw.docx"

    log("AI 创作舱：正在解析 Markdown 结构...")
    try:
        convert_text_to_docx(draft_text, raw_path)
    except Exception as e:
        log(f"❌ 解析失败: {str(e)}")
        return

    def _worker():
        try:
            log("正在套用当前排版预设...")
            reformat_docx(raw_path, path)
            log(f"✅ 生成成功！文件已保存至: {path}")
            if os.name == 'nt':
                os.startfile(path)
        except Exception as e:
            log(f"⚠️ 排版注入失败，已保存基础格式: {str(e)}")
            if os.path.exists(raw_path):
                try:
                    import shutil
                    shutil.move(raw_path, path)
                except Exception:
                    pass
            if os.name == 'nt' and os.path.exists(path):
                os.startfile(path)
        finally:
            if os.path.exists(raw_path):
                try:
                    os.remove(raw_path)
                except Exception:
                    pass

    threading.Thread(target=_worker, daemon=True).start()
