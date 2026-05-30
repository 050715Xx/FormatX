"""
AI 消息积木块分发容器 (对标前端 AssistantMessage)
管理单条回复内部的组件树，按块类型（Thinking/Text/Card）解耦渲染。
"""
import customtkinter as ctk
from ai_assist.typewriter_engine import TkinterTypewriterEffect


class AIMessageBlockDispatcher(ctk.CTkFrame):
    def __init__(self, parent, colors, on_update_markdown, on_export_docx):
        super().__init__(parent, fg_color="transparent")
        self.colors = colors
        self._update_markdown_textbox = on_update_markdown
        self._do_export_docx = on_export_docx

        self.blocks = {}
        self.typewriter = None
        self.ui_flags = {"thinking_collapsed": False}

    def render_state(self, state: dict):
        if state["has_reasoning"]:
            if "thinking" not in self.blocks:
                self._mount_thinking_block()
            if not self.ui_flags["thinking_collapsed"]:
                self.blocks["reasoning_lbl"].configure(text=state["reasoning_text"])
            if state["thinking_finished"] and not self.ui_flags["thinking_collapsed"]:
                self._collapse_thinking()

        if state["display_text"]:
            if "text" not in self.blocks:
                self._mount_text_block()
            self.typewriter.update(state["display_text"])

    def mount_action_card(self, draft_text: str, log_callback):
        if "draft" in self.blocks or not draft_text.strip():
            return

        card = ctk.CTkFrame(self, fg_color=self.colors["bg_panel"], corner_radius=12)
        card.pack(fill="x", pady=(5, 10), padx=(10, 50), anchor="w")

        lbl = ctk.CTkLabel(card, text="📄 排版草稿已就绪",
                           font=("Microsoft YaHei", 12, "bold"),
                           text_color=self.colors["text"])
        lbl.pack(anchor="w", padx=15, pady=(15, 5))

        btn_row = ctk.CTkFrame(card, fg_color="transparent")
        btn_row.pack(fill="x", padx=15, pady=(4, 12))

        ctk.CTkButton(btn_row, text="✨ 确认排版并预览 Word", height=32, corner_radius=8,
                      fg_color="#2D5A27", hover_color="#3A6E33",
                      command=lambda: self._do_export_docx(draft_text, log_callback)).pack(
            side="left", padx=(0, 10))

        self.blocks["draft"] = card
        self.blocks["draft_lbl"] = lbl

    def _mount_thinking_block(self):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", pady=2, anchor="w")

        btn = ctk.CTkButton(frame, text="⌄ 思考过程",
                            font=("Microsoft YaHei", 12, "italic"),
                            text_color=self.colors["text_muted"],
                            fg_color="transparent",
                            hover_color=self.colors["bg_panel"],
                            width=60, height=24, anchor="w",
                            command=self._toggle_reasoning)
        btn.pack(side="top", anchor="w", padx=(5, 0))

        lbl = ctk.CTkLabel(frame, text="", font=("Microsoft YaHei", 11),
                           text_color=self.colors["text_muted"],
                           justify="left", wraplength=450)
        lbl.pack(side="top", padx=(25, 10), pady=(0, 5), anchor="w")

        self.blocks["thinking"] = frame
        self.blocks["toggle_btn"] = btn
        self.blocks["reasoning_lbl"] = lbl

    def _mount_text_block(self):
        textbox = ctk.CTkTextbox(self, font=("Microsoft YaHei", 13),
                                 text_color=self.colors["text"],
                                 fg_color="transparent", wrap="word",
                                 border_width=0, activate_scrollbars=False)
        textbox.pack(side="top", fill="x", padx=(8, 0), anchor="w", expand=True)
        self.blocks["text"] = textbox

        def _render_wrapper(widget, text):
            self._update_markdown_textbox(widget, text, self.colors)

        self.typewriter = TkinterTypewriterEffect(textbox, _render_wrapper)

    def _collapse_thinking(self):
        if "reasoning_lbl" in self.blocks:
            self.blocks["reasoning_lbl"].pack_forget()
            self.blocks["toggle_btn"].configure(text="› 思考完成")
            self.ui_flags["thinking_collapsed"] = True

    def _toggle_reasoning(self):
        if self.ui_flags["thinking_collapsed"]:
            self.blocks["reasoning_lbl"].pack(
                side="top", padx=(25, 10), pady=(0, 5), anchor="w")
            self.blocks["toggle_btn"].configure(text="⌄ 思考过程")
            self.ui_flags["thinking_collapsed"] = False
        else:
            self._collapse_thinking()

    def update_theme(self, new_colors):
        self.colors = new_colors
        if "toggle_btn" in self.blocks:
            self.blocks["toggle_btn"].configure(
                text_color=new_colors["text_muted"],
                hover_color=new_colors["bg_panel"])
        if "reasoning_lbl" in self.blocks:
            self.blocks["reasoning_lbl"].configure(text_color=new_colors["text_muted"])
        if "text" in self.blocks:
            self.blocks["text"].configure(text_color=new_colors["text"])
        if "draft" in self.blocks:
            self.blocks["draft"].configure(fg_color=new_colors["bg_panel"])
            self.blocks["draft_lbl"].configure(text_color=new_colors["text"])
