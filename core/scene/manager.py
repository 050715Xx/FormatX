"""SceneConfig 的持久化管理器——JSON 文件与内存对象的双向转换"""
import json
from dataclasses import asdict
from pathlib import Path
from core.scene.schema import (SceneConfig, StyleConfig, HeadingPatternConfig,
                                   ColumnConfig, MarginConfig, PageSetupConfig)


def load_scene_config(json_path: Path) -> SceneConfig:
    if not json_path.exists():
        return SceneConfig()
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    config = SceneConfig(name=data.get("name", "默认格式"))
    if "page_setup" in data:
        ps = data["page_setup"]
        config.page_setup = PageSetupConfig(
            paper_size=ps.get("paper_size", "A4"),
            margin=MarginConfig(**ps["margin"]) if "margin" in ps else MarginConfig(),
            columns=ColumnConfig(**ps["columns"]) if "columns" in ps else ColumnConfig(),
        )
    if "styles" in data:
        for key, s_data in data["styles"].items():
            config.styles[key] = StyleConfig(**s_data)
    heading_detect = data.get("heading_detection", {})
    if heading_detect:
        config.front_matter_titles = heading_detect.get("front_matter_titles", [])
        for item in heading_detect.get("numbered_heading_patterns", []):
            config.numbered_heading_patterns.append(
                HeadingPatternConfig(pattern=item["pattern"], level=item["level"])
            )
    _TABLE_FIELDS = [
        "normal_table_border_mode", "table_border_width_pt",
        "three_line_header_width_pt", "three_line_bottom_width_pt",
        "normal_table_layout_mode", "normal_table_smart_levels",
        "normal_table_line_spacing_mode", "normal_table_repeat_header",
    ]
    for field in _TABLE_FIELDS:
        if field in data:
            setattr(config, field, data[field])
    return config


def save_scene_config(config: SceneConfig, json_path: Path):
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(asdict(config), f, ensure_ascii=False, indent=2)
