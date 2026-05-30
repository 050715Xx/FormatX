"""Word 原生样式提取器 — 继承链遍历 + OOXML 底层提取"""
from docx import Document
from docx.oxml.ns import qn


def _get_xml_ind_attr(style, attr_name) -> float:
    current = style
    while current is not None:
        pf = current.paragraph_format
        if pf is not None and pf._element is not None:
            ind = pf._element.find(qn('w:ind'))
            if ind is not None:
                val = ind.get(qn(attr_name))
                if val is not None:
                    return int(val) / 100.0
        current = current.base_style
    return 0.0


def _resolve_attr(style, attr_group, attr_name):
    current = style
    while current is not None:
        group = getattr(current, attr_group, None)
        if group is not None:
            val = getattr(group, attr_name, None)
            if val is not None:
                return val
        current = current.base_style
    return None


def _get_east_asia_font(style) -> str:
    current = style
    while current is not None:
        font = current.font
        if font is not None and font._element is not None:
            rFonts = font._element.find(qn('w:rFonts'))
            if rFonts is not None:
                val = rFonts.get(qn('w:eastAsia'))
                if val:
                    return val
        current = current.base_style
    return ""


def extract_raw_style_data(docx_path: str) -> dict:
    doc = Document(docx_path)
    raw_data = {"styles": {}, "page_setup": {}}

    section = doc.sections[0]
    cols = section._sectPr.xpath('./w:cols/@w:num') if section._sectPr is not None else []
    raw_data["page_setup"] = {
        "cols": int(cols[0]) if cols else 1,
        "top_cm": section.top_margin.cm if section.top_margin else 3.8,
        "bottom_cm": section.bottom_margin.cm if section.bottom_margin else 3.8,
        "left_cm": section.left_margin.cm if section.left_margin else 3.2,
        "right_cm": section.right_margin.cm if section.right_margin else 3.2,
    }

    STYLE_MAPPINGS = {
        "normal": ["Normal", "正文"],
        "heading1": ["Heading 1", "标题 1", "章标题", "一级标题"],
        "heading2": ["Heading 2", "标题 2", "二级标题"],
        "heading3": ["Heading 3", "标题 3", "三级标题"],
        "heading4": ["Heading 4", "标题 4", "四级标题"],
        "heading5": ["Heading 5", "标题 5", "五级标题"],
        "heading6": ["Heading 6", "标题 6", "六级标题"],
        "heading7": ["Heading 7", "标题 7", "七级标题"],
        "heading8": ["Heading 8", "标题 8", "八级标题"],
        "abstract": ["Abstract", "摘要", "摘要正文"],
        "toc_1": ["TOC 1", "目录 1"],
        "toc_2": ["TOC 2", "目录 2"],
        "toc_3": ["TOC 3", "目录 3"],
        "caption": ["Caption", "题注", "图表题注"],
        "header": ["Header", "页眉"],
        "footer": ["Footer", "页脚"],
    }

    for target_key, docx_names in STYLE_MAPPINGS.items():
        style = None
        for name in docx_names:
            try:
                style = doc.styles[name]
            except KeyError:
                continue
            if style:
                break

        if style:
            first_line_pt_obj = _resolve_attr(style, 'paragraph_format', 'first_line_indent')
            left_pt_obj = _resolve_attr(style, 'paragraph_format', 'left_indent')
            hanging_pt_obj = _resolve_attr(style, 'paragraph_format', 'hanging_indent')
            space_before_obj = _resolve_attr(style, 'paragraph_format', 'space_before')
            space_after_obj = _resolve_attr(style, 'paragraph_format', 'space_after')
            font_size_obj = _resolve_attr(style, 'font', 'size')

            raw_data["styles"][target_key] = {
                "first_line_indent_chars": _get_xml_ind_attr(style, 'w:firstLineChars'),
                "left_indent_chars": _get_xml_ind_attr(style, 'w:leftChars'),
                "hanging_indent_chars": _get_xml_ind_attr(style, 'w:hangingChars'),
                "first_line_indent_pt": first_line_pt_obj.pt if first_line_pt_obj else 0.0,
                "left_indent_pt": left_pt_obj.pt if left_pt_obj else 0.0,
                "hanging_indent_pt": hanging_pt_obj.pt if hanging_pt_obj else 0.0,
                "line_spacing_rule": _resolve_attr(style, 'paragraph_format', 'line_spacing_rule'),
                "line_spacing": _resolve_attr(style, 'paragraph_format', 'line_spacing'),
                "space_before_pt": space_before_obj.pt if space_before_obj else 0.0,
                "space_after_pt": space_after_obj.pt if space_after_obj else 0.0,
                "alignment": _resolve_attr(style, 'paragraph_format', 'alignment'),
                "font_size_pt": font_size_obj.pt if font_size_obj else 12.0,
                "font_en": _resolve_attr(style, 'font', 'name') or "",
                "font_cn": _get_east_asia_font(style),
                "bold": _resolve_attr(style, 'font', 'bold') or False,
                "italic": _resolve_attr(style, 'font', 'italic') or False,
            }

    return raw_data
