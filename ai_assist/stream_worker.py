"""
后台异步流式网络任务 (对标前端 ws-message-handler)
运行在子线程，负责网络 I/O，并将数据推入 Buffer。
"""
import threading
from ai_assist.multi_llm_client import stream_chat_completion


class StreamWorker(threading.Thread):
    def __init__(self, target_cfg, msgs, buffer, on_complete, on_error):
        super().__init__(daemon=True)
        self.target_cfg = target_cfg
        self.msgs = msgs
        self.buffer = buffer
        self.on_complete = on_complete
        self.on_error = on_error

    def run(self):
        try:
            generator = stream_chat_completion(
                api_key=self.target_cfg.get("api_key", ""),
                base_url=self.target_cfg.get("base_url", "https://api.deepseek.com"),
                model=self.target_cfg.get("model", "deepseek-chat"),
                messages=self.msgs,
                api_type=self.target_cfg.get("api_type", "openai-completions"),
            )

            for reasoning_chunk, content_chunk in generator:
                self.buffer.push(reasoning_chunk, content_chunk)

            self.buffer.flush()

            if self.on_complete:
                self.on_complete(self.buffer.state)

        except Exception as e:
            if self.on_error:
                self.on_error(str(e))
