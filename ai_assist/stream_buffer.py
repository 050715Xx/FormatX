"""
AI 流式状态管理器与节流阀 (对标前端 use-stream-buffer)
负责缓冲高频 Token，解析特殊标签，并控制向 UI 发射数据的频率。
"""
import time
import re
import threading


class StreamStateBuffer:
    def __init__(self, on_flush_callback, flush_interval=0.2):
        self.on_flush_callback = on_flush_callback
        self.flush_interval = flush_interval
        self.last_flush_time = 0.0
        self._flush_timer = None
        self.reset()

    def reset(self):
        self._raw_reasoning = ""
        self._raw_content = ""
        self.state = {
            "reasoning_text": "",
            "display_text": "",
            "draft_text": "",
            "has_reasoning": False,
            "thinking_finished": False
        }
        if self._flush_timer:
            self._flush_timer.cancel()
            self._flush_timer = None

    def push(self, reasoning_chunk: str, content_chunk: str):
        if reasoning_chunk:
            self._raw_reasoning += reasoning_chunk
            self.state["has_reasoning"] = True

        if content_chunk:
            clean_chunk = content_chunk.replace("*[思考中...]*", "")
            self._raw_content += clean_chunk
            if self.state["has_reasoning"] and not self.state["thinking_finished"] and clean_chunk.strip():
                self.state["thinking_finished"] = True

        current_time = time.time()
        elapsed = current_time - self.last_flush_time

        if elapsed >= self.flush_interval:
            self.flush()
        elif not self._flush_timer:
            remaining = self.flush_interval - elapsed
            self._flush_timer = threading.Timer(remaining, self._on_timer_flush)
            self._flush_timer.daemon = True
            self._flush_timer.start()

    def _on_timer_flush(self):
        self._flush_timer = None
        self.flush()

    def flush(self):
        self.last_flush_time = time.time()
        if self._flush_timer:
            self._flush_timer.cancel()
            self._flush_timer = None
        self._parse_raw_content()
        if self.on_flush_callback:
            self.on_flush_callback(self.state)

    def _parse_raw_content(self):
        content = self._raw_content

        think_match = re.search(r'<think>(.*?)(?:</think>|$)', content, flags=re.DOTALL)
        if think_match:
            self.state["has_reasoning"] = True
            self.state["reasoning_text"] = self._raw_reasoning + think_match.group(1)
            if "</think>" in content:
                self.state["thinking_finished"] = True
            content = re.sub(r'<think>.*?(?:</think>|$)', '', content, flags=re.DOTALL)
        else:
            self.state["reasoning_text"] = self._raw_reasoning

        draft_match = re.search(r'<draft>(.*?)(?:</draft>|$)', content, flags=re.DOTALL)
        if draft_match:
            self.state["draft_text"] = draft_match.group(1).strip()
            content = re.sub(r'<draft>.*?(?:</draft>|$)', '', content, flags=re.DOTALL)

        self.state["display_text"] = content
