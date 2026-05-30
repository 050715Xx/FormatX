"""排版模式的全局互斥状态机（模块级单例）"""
from core.state.constants import TypographyMode
from core.config_manager import load_full_config, update_config


class _ModeManager:
    def __init__(self):
        self._listeners = []
        cfg = load_full_config()
        self._current_mode = cfg.get("typography_mode", TypographyMode.PRESET)

    def subscribe(self, callback):
        if callback not in self._listeners:
            self._listeners.append(callback)

    def unsubscribe(self, callback):
        if callback in self._listeners:
            self._listeners.remove(callback)

    def get_mode(self) -> str:
        return self._current_mode

    def set_mode(self, new_mode: str, auto_save: bool = True):
        if self._current_mode == new_mode:
            return
        self._current_mode = new_mode
        if auto_save:
            update_config("typography_mode", new_mode)
        self._broadcast()

    def _broadcast(self):
        for listener in self._listeners:
            try:
                listener(self._current_mode)
            except Exception as e:
                print(f"[State Manager] Listener failed: {e}")


mode_manager = _ModeManager()
