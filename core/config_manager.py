"""FormatX 统一配置持久化中心"""
import os
import json
from pathlib import Path

_CONFIG_PATH = Path.home() / ".formatx_ai_config.json"


def load_full_config() -> dict:
    if not _CONFIG_PATH.exists():
        return {}
    try:
        with open(_CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        try:
            bak = _CONFIG_PATH.with_suffix(".bak")
            if _CONFIG_PATH.exists():
                os.rename(_CONFIG_PATH, bak)
        except Exception:
            pass
        return {}


def save_full_config(config_dict: dict):
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_dict, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def update_config(key_or_dict, value=None):
    """原子化增量更新：替代原 settings_dialog 的 _write_config_safe"""
    data = load_full_config()
    if isinstance(key_or_dict, dict):
        data.update(key_or_dict)
    else:
        data[key_or_dict] = value
    save_full_config(data)


def get_active_model_settings() -> dict:
    cfg = load_full_config()
    active_prov = cfg.get("active_provider", "deepseek")
    prov_data = cfg.get("providers", {}).get(active_prov, {})

    return {
        "provider_id": active_prov,
        "base_url": prov_data.get("base_url", cfg.get("base_url", "https://api.deepseek.com")),
        "api_key": prov_data.get("api_key", cfg.get("api_key", "")),
        "api_type": prov_data.get("api_type", cfg.get("api_type", "openai-completions")),
        "model": cfg.get("active_model", cfg.get("model", "deepseek-chat"))
    }


_PRESET_FILE_MAP = {
    "默认通用格式": "default_format",
    "学术毕业论文规范": "thesis_format",
    "IEEE Conference": "ieee_conference",
}


def _resolve_preset_path(filename: str):
    import sys
    from pathlib import Path
    if hasattr(sys, '_MEIPASS'):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent.parent
    return base / 'presets' / filename


def get_active_scene_config():
    cfg = load_full_config()

    if cfg.get("typography_mode") == "clone":
        target_path = cfg.get("clone_target_path")
        if target_path:
            import os as _os
            if _os.path.exists(target_path):
                from core.clone_engine.api import create_scene_from_clone
                return create_scene_from_clone(target_path)

    from core.scene.manager import load_scene_config
    preset_name = cfg.get("format_preset", "默认通用格式")
    filename = _PRESET_FILE_MAP.get(preset_name, "default_format")

    preset_path = _resolve_preset_path(f"{filename}.json")
    if preset_path.exists():
        scene_cfg = load_scene_config(preset_path)
    else:
        from core.scene.schema import SceneConfig
        scene_cfg = SceneConfig()

    _TABLE_OVERLAY_FIELDS = [
        "normal_table_border_mode", "table_border_width_pt",
        "three_line_header_width_pt", "three_line_bottom_width_pt",
        "normal_table_layout_mode", "normal_table_smart_levels",
        "normal_table_line_spacing_mode", "normal_table_repeat_header",
    ]
    for field in _TABLE_OVERLAY_FIELDS:
        if field in cfg:
            setattr(scene_cfg, field, cfg[field])

    if cfg.get("typography_mode") == "custom":
        custom = cfg.get("custom_settings", {})
        normal_style = scene_cfg.styles["normal"]
        if "indent_cm" in custom:
            normal_style.first_line_indent_cm = float(custom["indent_cm"])
        if "line_mode" in custom:
            normal_style.line_spacing_mode = custom["line_mode"]
        if "line_pt" in custom:
            normal_style.line_spacing_pt = float(custom["line_pt"])

    return scene_cfg
