"""
AI 智能双轨打字机引擎 (对标前端 StreamingMarkdownContent)
根据文本复杂度智能选择「逐字打字机」或「全量秒渲染」，防止 Markdown 语法破损导致界面闪烁。
"""
import re

COMPLEX_MARKDOWN_PATTERNS = [
    r'(^|\n)\s*(```|~~~)',
    r'(^|\n)\s*\$\$',
    r'(^|\n)\s*\\\[',
    r'(^|\n)\s*\|.*\|',
]

def is_typewriter_eligible(source_text: str) -> bool:
    if not source_text.strip() or '`' in source_text:
        return False
    return not any(re.search(pat, source_text) for pat in COMPLEX_MARKDOWN_PATTERNS)


class TkinterTypewriterEffect:
    def __init__(self, widget, on_render_callback, fps=30):
        self.widget = widget
        self.on_render_callback = on_render_callback
        self.delay_ms = int(1000 / fps)
        self._target_text = ""
        self._current_text = ""
        self._timer_id = None

    def update(self, full_text: str):
        self._target_text = full_text
        if not is_typewriter_eligible(full_text):
            self.stop()
            self._current_text = full_text
            self.on_render_callback(self.widget, full_text)
        elif not self._timer_id:
            self._step_forward()

    def _step_forward(self):
        if len(self._current_text) < len(self._target_text):
            delta = len(self._target_text) - len(self._current_text)
            batch = 3 if delta > 15 else 1
            self._current_text += self._target_text[
                len(self._current_text):len(self._current_text) + batch]
            self.on_render_callback(self.widget, self._current_text)
            self._timer_id = self.widget.after(self.delay_ms, self._step_forward)
        else:
            self._timer_id = None

    def stop(self):
        if self._timer_id:
            self.widget.after_cancel(self._timer_id)
            self._timer_id = None
