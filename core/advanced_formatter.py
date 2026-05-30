"""FormatX 高精度排版中台 — 缩进单位感知 + 行距 OOXML 同步"""
import logging
from docx.enum.text import WD_ALIGN_PARAGRAPH
from core.indent import apply_style_config_indents
from core.line_spacing import apply_line_spacing, sync_spacing_ooxml

logger = logging.getLogger("FormatX.Formatter")


class FormatXStyleAdapter:
    def __init__(self, formatx_style):
        self.alignment = getattr(formatx_style, "alignment", "justify")
        self.space_before_pt = getattr(formatx_style, "space_before_pt", 0.0)
        self.space_after_pt = getattr(formatx_style, "space_after_pt", 0.0)
        self.line_spacing_mode = getattr(formatx_style, "line_spacing_mode", "exact")
        self.line_spacing_pt = getattr(formatx_style, "line_spacing_pt", 20.0)
        self.size_pt = getattr(formatx_style, "size_pt", 12.0)

        left_cm = getattr(formatx_style, "left_indent_cm", 0.0)
        right_cm = getattr(formatx_style, "right_indent_cm", 0.0)
        first_line_cm = getattr(formatx_style, "first_line_indent_cm", 0.0)
        hanging_cm = getattr(formatx_style, "hanging_indent_cm", 0.0)

        self.left_indent_chars = left_cm
        self.left_indent_unit = "cm"
        self.right_indent_chars = right_cm
        self.right_indent_unit = "cm"

        if hanging_cm > 0:
            self.special_indent_mode = "hanging"
            self.special_indent_value = hanging_cm
            self.special_indent_unit = "cm"
            self.hanging_indent_chars = hanging_cm
            self.hanging_indent_unit = "cm"
            self.first_line_indent_chars = 0.0
            self.first_line_indent_unit = "cm"
        elif first_line_cm > 0:
            self.special_indent_mode = "first_line"
            self.special_indent_value = first_line_cm
            self.special_indent_unit = "cm"
            self.first_line_indent_chars = first_line_cm
            self.first_line_indent_unit = "cm"
            self.hanging_indent_chars = 0.0
            self.hanging_indent_unit = "cm"
        else:
            self.special_indent_mode = "none"
            self.special_indent_value = 0.0
            self.special_indent_unit = "cm"
            self.first_line_indent_chars = 0.0
            self.first_line_indent_unit = "cm"
            self.hanging_indent_chars = 0.0
            self.hanging_indent_unit = "cm"


class TypographyEngine:
    @staticmethod
    def apply_paragraph_style(paragraph, formatx_style_config):
        if not paragraph or not formatx_style_config:
            return False
        try:
            adapted = FormatXStyleAdapter(formatx_style_config)
            p_format = paragraph.paragraph_format
            p_element = paragraph._element

            align_map = {
                "left": WD_ALIGN_PARAGRAPH.LEFT,
                "right": WD_ALIGN_PARAGRAPH.RIGHT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            }
            p_format.alignment = align_map.get(
                str(adapted.alignment).lower(), WD_ALIGN_PARAGRAPH.JUSTIFY)

            apply_style_config_indents(p_format, p_element, adapted)
            apply_line_spacing(p_format, adapted.line_spacing_mode,
                              adapted.line_spacing_pt)
            sync_spacing_ooxml(
                p_element,
                space_before_pt=adapted.space_before_pt,
                space_after_pt=adapted.space_after_pt,
                line_spacing_type=adapted.line_spacing_mode,
                line_spacing_value=adapted.line_spacing_pt,
            )
            return True
        except Exception as e:
            logger.error(f"TypographyEngine 注入失败: {e}")
            return False
