"""样式翻译官 — Word 原生数据 → FormatX StyleConfig/SceneConfig"""
from core.scene.schema import StyleConfig, SceneConfig
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH


def _translate_alignment(align_enum) -> str:
    if align_enum == WD_ALIGN_PARAGRAPH.LEFT:
        return "left"
    if align_enum == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    if align_enum == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    if align_enum == WD_ALIGN_PARAGRAPH.JUSTIFY:
        return "justify"
    if align_enum == WD_ALIGN_PARAGRAPH.DISTRIBUTE:
        return "justify"
    return "justify"


def translate_to_formatx_schema(raw_data: dict) -> SceneConfig:
    config = SceneConfig(name="智能克隆模板")

    ps = raw_data.get("page_setup", {})
    config.page_setup.columns.count = ps.get("cols", 1)
    config.page_setup.margin.top_cm = round(ps.get("top_cm", 3.8), 2)
    config.page_setup.margin.bottom_cm = round(ps.get("bottom_cm", 3.8), 2)
    config.page_setup.margin.left_cm = round(ps.get("left_cm", 3.2), 2)
    config.page_setup.margin.right_cm = round(ps.get("right_cm", 3.2), 2)

    for key, raw_style in raw_data.get("styles", {}).items():
        target = config.styles.get(key, StyleConfig())
        base_font_pt = raw_style.get("font_size_pt", 12.0)

        target.size_pt = base_font_pt
        if raw_style.get("font_cn"):
            target.font_cn = raw_style["font_cn"]
        if raw_style.get("font_en"):
            target.font_en = raw_style["font_en"]
        if raw_style.get("bold") is not None:
            target.bold = raw_style["bold"]
        if raw_style.get("italic") is not None:
            target.italic = raw_style["italic"]

        if raw_style.get("alignment") is not None:
            target.alignment = _translate_alignment(raw_style["alignment"])

        target.space_before_pt = raw_style.get("space_before_pt", 0.0)
        target.space_after_pt = raw_style.get("space_after_pt", 0.0)

        indent_fields = [
            ("first_line_indent_chars", "first_line_indent_pt",
             "first_line_indent_cm"),
            ("left_indent_chars", "left_indent_pt", "left_indent_cm"),
            ("hanging_indent_chars", "hanging_indent_pt", "hanging_indent_cm"),
        ]
        for char_key, pt_key, cm_key in indent_fields:
            chars = raw_style.get(char_key, 0.0)
            if chars > 0:
                setattr(target, cm_key,
                        round((chars * base_font_pt) * 0.03527, 2))
            elif raw_style.get(pt_key, 0.0) > 0:
                setattr(target, cm_key,
                        round(raw_style[pt_key] * 0.03527, 2))

        spacing_rule = raw_style.get("line_spacing_rule")
        spacing_val = raw_style.get("line_spacing")
        target.line_spacing_mode = "exact"

        if spacing_rule in (WD_LINE_SPACING.EXACT, WD_LINE_SPACING.AT_LEAST):
            if hasattr(spacing_val, 'pt'):
                target.line_spacing_pt = spacing_val.pt
            elif spacing_val:
                target.line_spacing_pt = float(spacing_val) / 12700
            else:
                target.line_spacing_pt = base_font_pt * 1.2
        elif spacing_rule == WD_LINE_SPACING.ONE_POINT_FIVE:
            target.line_spacing_pt = base_font_pt * 1.5
        elif spacing_rule == WD_LINE_SPACING.DOUBLE:
            target.line_spacing_pt = base_font_pt * 2.0
        elif spacing_rule == WD_LINE_SPACING.MULTIPLE:
            mult = float(spacing_val) if spacing_val else 1.25
            target.line_spacing_pt = base_font_pt * mult
        else:
            target.line_spacing_pt = base_font_pt * 1.25

        config.styles[key] = target

    return config


def translate_tf_to_formatx(tf_scene_cfg) -> SceneConfig:
    """将 tf (src) 的 SceneConfig 翻译为 FormatX 的 SceneConfig"""
    fx = SceneConfig(name=getattr(tf_scene_cfg, "name", "智能克隆模板"))

    if hasattr(tf_scene_cfg, "page_setup"):
        ps = tf_scene_cfg.page_setup
        if hasattr(ps, "columns"):
            try:
                fx.page_setup.columns.count = int(getattr(ps.columns, "count", 1))
            except (TypeError, ValueError):
                pass
        if hasattr(ps, "margin"):
            fx.page_setup.margin.top_cm = round(getattr(ps.margin, "top_cm", 3.8), 2)
            fx.page_setup.margin.bottom_cm = round(getattr(ps.margin, "bottom_cm", 3.8), 2)
            fx.page_setup.margin.left_cm = round(getattr(ps.margin, "left_cm", 3.2), 2)
            fx.page_setup.margin.right_cm = round(getattr(ps.margin, "right_cm", 3.2), 2)

    for key, tf_style in getattr(tf_scene_cfg, "styles", {}).items():
        target = fx.styles.get(key, StyleConfig())

        target.font_cn = getattr(tf_style, "font_cn", "宋体")
        target.font_en = getattr(tf_style, "font_en", "Times New Roman")
        target.size_pt = getattr(tf_style, "size_pt", 12.0)
        target.bold = getattr(tf_style, "bold", False)
        target.italic = getattr(tf_style, "italic", False)

        align_val = getattr(tf_style, "alignment", None)
        if align_val == WD_ALIGN_PARAGRAPH.LEFT:
            target.alignment = "left"
        elif align_val == WD_ALIGN_PARAGRAPH.CENTER:
            target.alignment = "center"
        elif align_val == WD_ALIGN_PARAGRAPH.RIGHT:
            target.alignment = "right"
        else:
            target.alignment = "justify"

        chars = getattr(tf_style, "first_line_indent_chars", 0.0)
        if chars > 0:
            target.first_line_indent_cm = round(
                (chars * target.size_pt) * 0.03527, 2)

        left_chars = getattr(tf_style, "left_indent_chars", 0.0)
        if left_chars > 0:
            target.left_indent_cm = round(
                (left_chars * target.size_pt) * 0.03527, 2)

        hang_chars = getattr(tf_style, "hanging_indent_chars", 0.0)
        if hang_chars > 0:
            target.hanging_indent_cm = round(
                (hang_chars * target.size_pt) * 0.03527, 2)

        spacing_type = getattr(tf_style, "line_spacing_type", "multiple")
        spacing_val = getattr(tf_style, "line_spacing_pt", target.size_pt * 1.25)
        target.line_spacing_mode = "exact"
        if spacing_type == "multiple":
            mult = float(spacing_val) if spacing_val else 1.25
            target.line_spacing_pt = target.size_pt * mult
        else:
            target.line_spacing_pt = float(spacing_val) if spacing_val else target.size_pt * 1.25

        target.space_before_pt = getattr(tf_style, "space_before_pt", 0.0)
        target.space_after_pt = getattr(tf_style, "space_after_pt", 0.0)

        fx.styles[key] = target

    return fx
