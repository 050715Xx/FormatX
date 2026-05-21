#!/usr/bin/env python3

"""
文本 / Word → 严格排版 Word 文档转换器

支持两种输入模式：
    1. .md / .txt → 按 Markdown 标记（# / ## / ### / ```）解析并排版
    2. .docx      → 读取已有文档，自动识别段落角色后重新套用严格格式

依赖：
    pip install python-docx

用法：
    from format_conversion import convert_markdown_to_docx, reformat_docx

    convert_markdown_to_docx('input.md', 'output.docx')
    reformat_docx('input.docx', 'output.docx')

    或直接运行（自动按扩展名选模式）：
    python format_conversion.py input.md output.docx
    python format_conversion.py input.docx output.docx
"""

import os

import sys

import re

from docx import Document

from docx.shared import Pt, Cm, RGBColor

from docx.enum.text import WD_BREAK

from docx.oxml.ns import qn

from docx.oxml import OxmlElement

try:

    from latex_to_omml import replace_latex_with_omml

    _HAS_OMML = True

except ImportError:

    _HAS_OMML = False

    def replace_latex_with_omml(para): return False

# ── 辅助：为文档注册简单列表编号 ──────────────────────────────

def _ensure_list_numbering(doc):

    """

    确保文档有一个简单十进制列表编号定义，返回 numId。

    同一个 doc 只创建一次（缓存在 doc 对象上）。

    """

    cache_attr = '_formatx_list_num_id'

    if hasattr(doc, cache_attr):

        return getattr(doc, cache_attr)

    numbering_part = doc.part.numbering_part

    num_root = numbering_part._element

    max_abs, max_num = 0, 0

    for el in num_root.findall(qn('w:abstractNum')):

        v = int(el.get(qn('w:abstractNumId'), '0'))

        if v > max_abs: max_abs = v

    for el in num_root.findall(qn('w:num')):

        v = int(el.get(qn('w:numId'), '0'))

        if v > max_num: max_num = v

    abs_id, num_id = max_abs + 1, max_num + 1

    abs_el = OxmlElement('w:abstractNum')

    abs_el.set(qn('w:abstractNumId'), str(abs_id))

    mlt = OxmlElement('w:multiLevelType')

    mlt.set(qn('w:val'), 'multilevel')

    abs_el.append(mlt)

    for ilvl, fmt in enumerate(('%1.', '%1.%2.', '%1.%2.%3.')):

        lvl = OxmlElement('w:lvl')

        lvl.set(qn('w:ilvl'), str(ilvl))

        s = OxmlElement('w:start'); s.set(qn('w:val'), '1'); lvl.append(s)

        nf = OxmlElement('w:numFmt'); nf.set(qn('w:val'), 'decimal'); lvl.append(nf)

        lt = OxmlElement('w:lvlText'); lt.set(qn('w:val'), fmt); lvl.append(lt)

        lj = OxmlElement('w:lvlJc'); lj.set(qn('w:val'), 'left'); lvl.append(lj)

        abs_el.append(lvl)

    num_root.insert(0, abs_el)

    num_el = OxmlElement('w:num')

    num_el.set(qn('w:numId'), str(num_id))

    ref = OxmlElement('w:abstractNumId'); ref.set(qn('w:val'), str(abs_id))

    num_el.append(ref)

    num_root.insert(0, num_el)

    setattr(doc, cache_attr, num_id)

    return num_id

def _apply_list_numpr(para, num_id, ilvl=0):

    """给段落加上 w:numPr，使其成为自动编号列表项。"""

    pPr = para._element.get_or_add_pPr()

    old = pPr.find(qn('w:numPr'))

    if old is not None: pPr.remove(old)

    numPr = OxmlElement('w:numPr')

    iel = OxmlElement('w:ilvl'); iel.set(qn('w:val'), str(ilvl)); numPr.append(iel)

    nid = OxmlElement('w:numId'); nid.set(qn('w:val'), str(num_id)); numPr.append(nid)

    pPr.append(numPr)

# ── 标题多级编号模板引擎 ─────────────────────────────────────

def _compile_heading_template(template, ilvl):

    """编译标题编号模板：'第{current}章' → '第%1章' / '{level1}.{current}' → '%1.%2'"""

    def _repl(m):

        name = m.group(1)

        if name in ('current', 'n', 'cn'):

            return f'%{ilvl + 1}'

        m2 = re.match(r'level(\d+)', name)

        if m2:

            return f'%{m2.group(1)}'

        return m.group(0)

    return re.sub(r'\{([a-zA-Z][a-zA-Z0-9_]*)\}', _repl, template)

def _register_heading_numbering(doc, templates=None):

    """

    注册标题多级编号定义（abstractNum + num），并链接到 Heading 1/2/3 样式。

    templates: None 用默认模板，或 dict 如 {0: '第{current}章', 1: '{level1}.{current}'}

    返回 numId。

    """

    if templates is None:

        # 默认：一级用"第X章"，二级用"X.X"，三级用"X.X.X"

        templates = {0: '第{current}章', 1: '{level1}.{current}', 2: '{level1}.{level2}.{current}'}

    numbering_part = doc.part.numbering_part

    num_root = numbering_part._element

    max_abs, max_num = 0, 0

    for el in num_root.findall(qn('w:abstractNum')):

        v = int(el.get(qn('w:abstractNumId'), '0'))

        if v > max_abs: max_abs = v

    for el in num_root.findall(qn('w:num')):

        v = int(el.get(qn('w:numId'), '0'))

        if v > max_num: max_num = v

    abs_id, num_id = max_abs + 1, max_num + 1

    abs_el = OxmlElement('w:abstractNum')

    abs_el.set(qn('w:abstractNumId'), str(abs_id))

    mlt = OxmlElement('w:multiLevelType')

    mlt.set(qn('w:val'), 'multilevel')

    abs_el.append(mlt)

    for ilvl in range(max(templates.keys()) + 1):

        tmpl = templates.get(ilvl, '{current}')

        lvl_text = _compile_heading_template(tmpl, ilvl)

        lvl = OxmlElement('w:lvl')

        lvl.set(qn('w:ilvl'), str(ilvl))

        s = OxmlElement('w:start'); s.set(qn('w:val'), '1'); lvl.append(s)

        nf = OxmlElement('w:numFmt')

        nf.set(qn('w:val'), 'chineseCounting' if ilvl == 0 and '第' in tmpl else 'decimal')

        lvl.append(nf)

        lt = OxmlElement('w:lvlText'); lt.set(qn('w:val'), lvl_text); lvl.append(lt)

        lj = OxmlElement('w:lvlJc'); lj.set(qn('w:val'), 'left'); lvl.append(lj)

        abs_el.append(lvl)

    num_root.insert(0, abs_el)

    num_el = OxmlElement('w:num')

    num_el.set(qn('w:numId'), str(num_id))

    ref = OxmlElement('w:abstractNumId'); ref.set(qn('w:val'), str(abs_id))

    num_el.append(ref)

    num_root.insert(0, num_el)

    # 【已禁用】链接 Heading 1/2/3 样式到编号 —— 改为所见即所得，不再强制自动编号
    # for ilvl, style_name in enumerate(('Heading 1', 'Heading 2', 'Heading 3')):
    #     try:
    #         style = doc.styles[style_name]
    #     except KeyError:
    #         continue
    #     style_el = style.element
    #     pPr = style_el.find(qn('w:pPr'))
    #     if pPr is None:
    #         pPr = OxmlElement('w:pPr')
    #         style_el.insert(0, pPr)
    #     old = pPr.find(qn('w:numPr'))
    #     if old is not None:
    #         pPr.remove(old)
    #     numPr = OxmlElement('w:numPr')
    #     iel = OxmlElement('w:ilvl'); iel.set(qn('w:val'), str(ilvl)); numPr.append(iel)
    #     nid = OxmlElement('w:numId'); nid.set(qn('w:val'), str(num_id)); numPr.append(nid)
    #     pPr.append(numPr)

    setattr(doc, '_formatx_heading_num_id', num_id)

    return num_id

def _get_heading_num_id(doc):

    """获取或创建标题编号 numId（惰性，只创建一次）。"""

    if hasattr(doc, '_formatx_heading_num_id'):

        return getattr(doc, '_formatx_heading_num_id')

    return _register_heading_numbering(doc)

# ---------------------------------------------------------------------------

# 底层格式工具

# ---------------------------------------------------------------------------

def _set_run_font(run, cn_font='宋体', en_font='Times New Roman',

                  size_pt=12, bold=False):

    """

    为 run 同时设置中文字体（w:eastAsia）和西文字体（w:ascii / w:hAnsi）。

    这是整个排版的关键——python-docx 的 run.font.name 只能设西文，

    中文必须通过 XML 层的 w:eastAsia 属性单独指定。

    """

    run.font.size = Pt(size_pt)

    run.bold = bold

    run.font.color.rgb = RGBColor(0, 0, 0)  # 强制黑色，屏蔽 Heading 样式默认蓝色

    # 西文字体（影响 ASCII 数字、英文、符号）

    run.font.name = en_font

    # 东亚字体（影响中文、日文、韩文）

    rPr = run._r.get_or_add_rPr()

    rFonts = rPr.find(qn('w:rFonts'))

    if rFonts is None:

        rFonts = OxmlElement('w:rFonts')

        rPr.insert(0, rFonts)

    rFonts.set(qn('w:eastAsia'), cn_font)

def _set_para_format(para, line_spacing_pt=20, first_line_indent=None):

    """

    设置段落行距（固定值）和首行缩进。

    参数

    ----

    line_spacing_pt : float

        固定行距，单位磅。Pt(20) 即为「固定值 20 磅」。

    first_line_indent : Cm | Pt | None

        首行缩进距离。传 Cm(0) 即明确取消缩进（代码块）；

        传 Cm(0.85) 即正文两字符缩进；None 表示不修改。

    """

    pf = para.paragraph_format

    pf.line_spacing = Pt(line_spacing_pt)

    # 去掉段前段后间距，保证严格 20 磅行距栅格

    pf.space_before = Pt(0)

    pf.space_after = Pt(0)

    if first_line_indent is not None:

        pf.first_line_indent = first_line_indent

# ---------------------------------------------------------------------------

# 段落构建辅助

# ---------------------------------------------------------------------------

def _copy_runs(para, source_para, cn_font, en_font, size_pt, default_bold):

    """

    从 source_para 逐 run 复制到 para，保留原有的 bold/italic/underline，

    仅替换字体名和字号。若无 source_para 则退化为单 run 模式。

    返回 para（便于链式调用）。

    """

    if source_para is None:

        return None  # 调用方自行处理

    # 兜底：获取源段落样式本身的加粗倾向
    p_bold = source_para.style.font.bold if (source_para.style and source_para.style.font) else False

    for src_run in source_para.runs:

        # OMML 模式：保留 $...$ 原样，让 replace_latex_with_omml 处理

        if _HAS_OMML and '$' in src_run.text:

            run_text = src_run.text

        else:

            run_text = _latex_to_text(src_run.text)

        run = para.add_run(run_text)

        # 四级回退：run 自身 → 字符样式 → 段落样式 → 默认值
        r_style_bold = src_run.style.font.bold if (src_run.style and src_run.style.font) else None

        if src_run.bold is not None:
            actual_bold = src_run.bold
        elif r_style_bold is not None:
            actual_bold = r_style_bold
        else:
            actual_bold = p_bold if p_bold is not None else default_bold

        run.bold = actual_bold

        if src_run.italic is not None:

            run.italic = src_run.italic

        if src_run.underline is not None:

            run.underline = src_run.underline

        _set_run_font(run, cn_font=cn_font, en_font=en_font,

                      size_pt=size_pt, bold=run.bold)

    return para

def _strip_heading_prefix(text, level):

    """【所见即所得】：绝对不切除任何前缀，原样保留用户复制的内容。"""

    return text.strip()

def _add_heading(doc, text, level, alignment=None, source_para=None,

                 extra_prefix=''):

    """

    添加一级 / 二级 / 三级标题。

    - 一级：宋体 3号（16 pt）加粗，段前段后 0.5 行（10 pt）

    - 二级：宋体 4号（14 pt）加粗，段前段后 0.5 行（10 pt）

    - 三级：宋体 小4号（12 pt）加粗，段前段后为 0

    均不缩进。alignment 为 None 时默认左对齐。

    source_para 不为 None 时，逐 run 复制保留原加粗/斜体。

    extra_prefix 用于自动编号前缀（如 "1. "）。

    """

    size_map = {1: 16, 2: 14, 3: 12}

    size = size_map.get(level, 12)

    para = doc.add_paragraph()

    _set_para_format(para, line_spacing_pt=20, first_line_indent=Cm(0))

    if level in (1, 2):

        para.paragraph_format.space_before = Pt(10)

        para.paragraph_format.space_after = Pt(10)

    from docx.enum.text import WD_ALIGN_PARAGRAPH

    if alignment is not None:

        para.paragraph_format.alignment = alignment

    elif re.match(r'^第(\d+|[一二三四五六七八九十]+)章', text.strip()) or \
         re.match(r'^(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)', text.strip()):

        # "第X章" / 摘要 … → 居中 + 另起一页

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        para.paragraph_format.page_break_before = True

    else:

        # 其余标题强制左对齐（防止两端对齐拉开字间距）

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 标题文本去多余空格

    title_text = re.sub(r'\s+', '', text.strip()) if re.search(r'[一-鿿]', text) else text.strip()

    # 启用原生标题编号（Heading 样式已绑定 abstractNum）

    heading_styles = {1: 'Heading 1', 2: 'Heading 2', 3: 'Heading 3'}

    heading_style = heading_styles.get(level)

    use_native = (heading_style is not None and extra_prefix == '')

    if use_native:

        try:

            para.style = doc.styles[heading_style]

            _get_heading_num_id(doc)

        except KeyError:

            pass

        title_text = _strip_heading_prefix(title_text, level)

    # 自动编号前缀（仅 .docx 重排时作为兜底）

    if extra_prefix:

        pre_run = para.add_run(extra_prefix)

        _set_run_font(pre_run, cn_font='宋体', en_font='Times New Roman',

                      size_pt=size, bold=True)

    # 原生编号洗掉了前缀 → 用干净文本，不抄原文（防双重编号）

    is_cleaned_native = use_native and title_text != text.strip()

    if source_para is not None and not is_cleaned_native:

        _copy_runs(para, source_para, '宋体', 'Times New Roman', size, True)

    else:

        run = para.add_run(title_text)

        _set_run_font(run, cn_font='宋体', en_font='Times New Roman',

                      size_pt=size, bold=True)

    # 将段落中的 $...$ LaTeX 转为 OMML 公式

    if _HAS_OMML:

        replace_latex_with_omml(para)

def _add_body(doc, text, alignment=None, left_indent=None,

              first_line_indent=None, source_para=None, extra_prefix='',

              style_name=None):

    """

    添加正文段落。

    - 中文：宋体 小4号（12 pt）

    - 英文 / 数字：Times New Roman 小4号（12 pt）

    - 首行缩进 2 字符 ≈ 0.85 厘米（可通过参数覆盖）

    - 两端对齐（Justify）

    - 固定行距 20 磅

    - 若原段落有左缩进（如子项编号），则沿用原值

    - style_name 不为空时保留原段落样式（如 List Number）

    """

    from docx.enum.text import WD_ALIGN_PARAGRAPH

    first = first_line_indent if first_line_indent is not None else Cm(0.85)

    para = doc.add_paragraph()

    _set_para_format(para, line_spacing_pt=20, first_line_indent=first)

    if alignment is not None:

        para.paragraph_format.alignment = alignment

    else:

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    if left_indent is not None:

        para.paragraph_format.left_indent = left_indent

    # 保留原列表样式

    if style_name and 'list' in style_name.lower():

        try:

            para.style = doc.styles[style_name]

        except KeyError:

            pass

    # 自动编号前缀（如有）

    if extra_prefix:

        pre_run = para.add_run(extra_prefix)

        _set_run_font(pre_run, cn_font='宋体', en_font='Times New Roman',

                      size_pt=12, bold=True)

    if source_para is not None:

        # 联合判断：run 自身 + 段落样式继承
        has_runs = len(source_para.runs) > 0
        p_style_bold = source_para.style.font.bold if (source_para.style and source_para.style.font) else False

        def _is_run_bold(r):
            r_style_bold = r.style.font.bold if (r.style and r.style.font) else None
            if r.bold is not None: return r.bold
            if r_style_bold is not None: return r_style_bold
            if p_style_bold is not None: return p_style_bold
            return False

        all_runs_bold = has_runs and all(_is_run_bold(r) for r in source_para.runs)

        if all_runs_bold:
            full_text = source_para.text
            # 冒号雷达：整段加粗时，若存在冒号则只保留冒号前的标题加粗
            if '：' in full_text or ':' in full_text:
                delim = '：' if '：' in full_text else ':'
                parts = full_text.split(delim, 1)

                r1 = para.add_run(parts[0] + delim)
                _set_run_font(r1, cn_font='宋体', en_font='Times New Roman', size_pt=12, bold=True)

                r2 = para.add_run(parts[1].strip())
                _set_run_font(r2, cn_font='宋体', en_font='Times New Roman', size_pt=12, bold=False)
            else:
                # 无冒号全段加粗：整体降级，防止伪装成标题
                _copy_runs(para, source_para, '宋体', 'Times New Roman', 12, False)
                for r in para.runs:
                    r.bold = False
        else:
            # 局部加粗：完美继承原文档的局部格式
            _copy_runs(para, source_para, '宋体', 'Times New Roman', 12, False)

    else:

        raw_text = text.strip()
        parts = re.split(r'\*\*(.*?)\*\*', raw_text)

        for i, part in enumerate(parts):
            if not part:
                continue
            run = para.add_run(part)
            is_bold = (i % 2 != 0)
            _set_run_font(run, cn_font='宋体', en_font='Times New Roman',
                          size_pt=12, bold=is_bold)

    # 将段落中的 $...$ LaTeX 转为 OMML 公式

    if _HAS_OMML:

        replace_latex_with_omml(para)

def _to_unicode_scripts(text):

    """将 ^{x} / _{y} / \bar{A} 转为 Unicode 上/下标/重音符。"""

    SUPER = str.maketrans('0123456789+-=()ij', '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁱʲ')

    SUB = str.maketrans('0123456789+-=()in', '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢₙ')

    OVERLINE = '̅'  # 结合上划线

    def _replace_sup(m):   return m.group(1).translate(SUPER)

    def _replace_sub(m):   return m.group(1).translate(SUB)

    def _replace_bar(m):   return m.group(1) + OVERLINE

    def _replace_hat(m):   return m.group(1) + '̂'

    def _replace_tilde(m): return m.group(1) + '̃'

    def _replace_vec(m):   return m.group(1) + '⃗'

    def _replace_dot(m):   return m.group(1) + '̇'

    # \bar{A} / \bar{AB}

    text = re.sub(r'\\bar\{([^}]+)\}', _replace_bar, text)

    text = re.sub(r'\\bar\s+(\w)', _replace_bar, text)

    # \hat{A}

    text = re.sub(r'\\hat\{([^}]+)\}', _replace_hat, text)

    text = re.sub(r'\\hat\s+(\w)', _replace_hat, text)

    # \tilde{A}

    text = re.sub(r'\\tilde\{([^}]+)\}', _replace_tilde, text)

    text = re.sub(r'\\tilde\s+(\w)', _replace_tilde, text)

    # \vec{A}

    text = re.sub(r'\\vec\{([^}]+)\}', _replace_vec, text)

    text = re.sub(r'\\vec\s+(\w)', _replace_vec, text)

    # \dot{A}

    text = re.sub(r'\\dot\{([^}]+)\}', _replace_dot, text)

    text = re.sub(r'\\dot\s+(\w)', _replace_dot, text)

    # _{ABC} → 下标（花括号）

    text = re.sub(r'_\{([^}]+)\}', _replace_sub, text)

    # ^{ABC} → 上标（花括号）

    text = re.sub(r'\^\{([^}]+)\}', _replace_sup, text)

    # _数字 → 下标（裸下划线+数字，如 a_1→a₁，不影响代码变量 hc_half）

    text = re.sub(r'_(\d+)', _replace_sub, text)

    # ^数字 → 上标

    text = re.sub(r'\^(\d+)', _replace_sup, text)

    return text

# 全局 LaTeX 命令映射（按长度降序，防 \in 截断 \int）

_LATEX_GLOBAL_MAP_SORTED = sorted([

    ('\\implies', '⇒'), ('\\Leftrightarrow', '⇔'), ('\\leftrightarrow', '↔'),

    ('\\Rightarrow', '⇒'), ('\\rightarrow', '→'), ('\\longrightarrow', '→'),

    ('\\Leftarrow', '⇐'), ('\\leftarrow', '←'), ('\\longleftarrow', '←'),

    ('\\varepsilon', 'ε'), ('\\vartheta', 'ϑ'), ('\\varphi', 'φ'),

    ('\\subseteq', '⊆'), ('\\supseteq', '⊇'), ('\\notin', '∉'),

    ('\\parallel', '∥'), ('\\emptyset', '∅'), ('\\forall', '∀'),

    ('\\exists', '∃'), ('\\propto', '∝'), ('\\simeq', '≃'),

    ('\\approx', '≈'), ('\\equiv', '≡'),

    ('\\infty', '∞'), ('\\partial', '∂'), ('\\nabla', '∇'),

    ('\\times', '×'), ('\\cdot', '·'), ('\\ldots', '…'), ('\\cdots', '…'),

    ('\\alpha', 'α'), ('\\beta', 'β'), ('\\gamma', 'γ'), ('\\delta', 'δ'),

    ('\\epsilon', 'ε'), ('\\zeta', 'ζ'), ('\\eta', 'η'), ('\\theta', 'θ'),

    ('\\iota', 'ι'), ('\\kappa', 'κ'), ('\\lambda', 'λ'), ('\\mu', 'µ'),

    ('\\nu', 'ν'), ('\\xi', 'ξ'), ('\\rho', 'ρ'), ('\\sigma', 'σ'),

    ('\\tau', 'τ'), ('\\upsilon', 'υ'), ('\\phi', 'φ'), ('\\chi', 'χ'),

    ('\\psi', 'ψ'), ('\\omega', 'ω'),

    ('\\Gamma', 'Γ'), ('\\Delta', 'Δ'), ('\\Theta', 'Θ'),

    ('\\Lambda', 'Λ'), ('\\Xi', 'Ξ'), ('\\Pi', 'Π'), ('\\Sigma', 'Σ'),

    ('\\Upsilon', 'ϒ'), ('\\Phi', 'Φ'), ('\\Psi', 'Ψ'), ('\\Omega', 'Ω'),

    ('\\left', ''), ('\\right', ''),

    ('\\int', '∫'), ('\\sum', 'Σ'), ('\\prod', 'Π'),

    ('\\neq', '≠'), ('\\leq', '≤'), ('\\geq', '≥'),

    ('\\div', '÷'), ('\\pm', '±'), ('\\mp', '∓'),

    ('\\iff', '⇔'), ('\\to', '→'), ('\\pi', 'π'),

    ('\\ll', '≪'), ('\\gg', '≫'), ('\\mid', '|'),

    ('\\oplus', '⊕'), ('\\otimes', '⊗'), ('\\odot', '⊙'),

    ('\\ominus', '⊖'), ('\\oslash', '⊘'), ('\\star', '⋆'),

    ('\\ast', '∗'), ('\\prime', '′'),

    ('\\circ', '∘'), ('\\bullet', '•'), ('\\diamond', '⋄'),

    ('\\bigoplus', '⨁'), ('\\bigotimes', '⨂'), ('\\bigodot', '⨀'),

    ('\\subset', '⊂'), ('\\supset', '⊃'), ('\\land', '∧'), ('\\lor', '∨'),

    ('\\cup', '∪'), ('\\cap', '∩'), ('\\sim', '∼'), ('\\neg', '¬'),

    ('\\in', '∈'), ('\\ni', '∋'), ('\\perp', '⊥'),

], key=lambda x: -len(x[0]))

def _sqrt_replace(m):

    return '√(' + m.group(1) + ')'

def _latex_to_text(text):

    """

    将简单 LaTeX 数学公式转为可读的纯文本。

    $\\text{投后估值} = \\frac{a}{b}$  →  投后估值 = (a)/(b)

    """

    # --- \frac 替换（处理嵌套花括号） ---

    def _replace_frac(s):

        result = []

        i = 0

        while i < len(s):

            if s[i:i + 5] == '\\frac':

                # 找第一个 {

                j = i + 5

                while j < len(s) and s[j] != '{':

                    j += 1

                if j >= len(s):

                    result.append(s[i:])

                    break

                # 匹配花括号提取分子

                depth = 0

                k = j

                while k < len(s):

                    if s[k] == '{':

                        depth += 1

                    elif s[k] == '}':

                        depth -= 1

                        if depth == 0:

                            break

                    k += 1

                numerator = s[j + 1:k]

                # 匹配花括号提取分母

                m = k + 1

                while m < len(s) and s[m] != '{':

                    m += 1

                depth = 0

                n = m

                while n < len(s):

                    if s[n] == '{':

                        depth += 1

                    elif s[n] == '}':

                        depth -= 1

                        if depth == 0:

                            break

                    n += 1

                denominator = s[m + 1:n]

                # 递归处理分子分母内部的 frac

                numerator = _replace_frac(numerator)

                denominator = _replace_frac(denominator)

                result.append(f'({numerator})/({denominator})')

                i = n + 1

            else:

                result.append(s[i])

                i += 1

        return ''.join(result)

    # --- 提取所有 $...$ 公式并转换 ---

    def _convert_dollar(match):

        formula = match.group(1)

        formula = re.sub(r'\\text\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', r'\1', formula)

        formula = re.sub(r'\\sqrt\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', _sqrt_replace, formula)

        formula = _replace_frac(formula)

        for cmd, sym in _LATEX_GLOBAL_MAP_SORTED:

            formula = formula.replace(cmd, sym)

        for fn in ('sin', 'cos', 'tan', 'csc', 'sec', 'cot',

                   'arcsin', 'arccos', 'arctan',

                   'log', 'ln', 'lim', 'exp', 'det', 'gcd',

                   'max', 'min', 'sup', 'inf'):

            formula = formula.replace('\\' + fn, fn)

        formula = formula.replace('\\,', ' ').replace('\\;', ' ').replace('\\ ', ' ')

        return formula

    # 独立命令替换（不碰 $...$）：符号 / 函数名 / 上下标

    text = _apply_latex_commands(text)

    # $$...$$ 和 $...$ 处理（仅在不走 OMML 时使用）

    text = re.sub(r'\$\$([^$]+)\$\$', _convert_dollar, text)

    text = re.sub(r'\$([^$]+)\$', _convert_dollar, text)

    return text

def _apply_latex_commands(text):

    """将 $...$ 外的 LaTeX 命令转为 Unicode（$...$ 内保持原样供 OMML 处理）。"""

    parts = re.split(r'(\$[^$]+\$|\$\$[^$]+\$\$)', text)

    for i, part in enumerate(parts):

        if part.startswith('$'):

            continue  # 跳过公式区域

        # \mathcal{X} / \mathbb{X} → X

        part = re.sub(r'\\mathcal\{([^}]+)\}', r'\1', part)

        part = re.sub(r'\\mathbb\{([^}]+)\}', r'\1', part)

        part = re.sub(r'\\mathrm\{([^}]+)\}', r'\1', part)

        part = re.sub(r'\\mathbf\{([^}]+)\}', r'\1', part)

        part = re.sub(r'\\mathit\{([^}]+)\}', r'\1', part)

        part = re.sub(r'\\sqrt\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}', _sqrt_replace, part)

        for cmd, sym in _LATEX_GLOBAL_MAP_SORTED:

            part = part.replace(cmd, sym)

        for fn in ('sin', 'cos', 'tan', 'csc', 'sec', 'cot',

                   'arcsin', 'arccos', 'arctan',

                   'log', 'ln', 'lim', 'exp', 'det', 'gcd',

                   'max', 'min', 'sup', 'inf'):

            part = part.replace('\\' + fn, fn)

        part = _to_unicode_scripts(part)

        parts[i] = part

    return ''.join(parts)

def _add_cover_page(doc, title=''):

    """

    添加封面预留页。

    - 如果提供了 title，则按封面标题格式渲染：

      黑体，初号（42 pt），加粗，居中

    - 如果 title 为空字符串，则留空白封面

    - 末尾插入分页符，后续内容从第 2 页开始

    """

    from docx.enum.text import WD_ALIGN_PARAGRAPH

    if doc.paragraphs:

        para = doc.paragraphs[0]

    else:

        para = doc.add_paragraph()

    if title:

        # 封面标题：黑体 初号（42 pt）加粗 居中

        run = para.add_run(title.strip())

        _set_run_font(run, cn_font='黑体', en_font='黑体',

                      size_pt=42, bold=True)

        para.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

        _set_para_format(para, line_spacing_pt=20, first_line_indent=Cm(0))

    para.add_run().add_break(WD_BREAK.PAGE)

def _add_code_line(doc, line_text, source_para=None):

    """

    添加代码块中的一行。

    - Times New Roman 5号（11 pt）

    - 不缩进，保留原始前导空格（维持缩进层级）

    - 固定行距 20 磅

    source_para 不为 None 时逐 run 复制保留原加粗/斜体。

    """

    para = doc.add_paragraph()

    _set_para_format(para, line_spacing_pt=20, first_line_indent=Cm(0))

    if source_para is not None:

        _copy_runs(para, source_para, '宋体', 'Times New Roman', 11, False)

    else:

        run = para.add_run(line_text)

        _set_run_font(run, cn_font='宋体', en_font='Times New Roman',

                      size_pt=11, bold=False)

    # 将段落中的 $...$ LaTeX 转为 OMML 公式

    if _HAS_OMML:

        replace_latex_with_omml(para)

# ---------------------------------------------------------------------------

# .docx 输入：段落角色检测

# ---------------------------------------------------------------------------

# 常见等宽字体（用于识别代码块）

_MONOSPACE_FONTS = {

    'consolas', 'courier', 'courier new', 'source code pro',

    'fira code', 'jetbrains mono', 'monaco', 'menlo', 'dejavu sans mono',

    'lucida console', 'inconsolata', 'cascadia code', 'ubuntu mono',

}

# 中文序号

_CN_NUM_PLAIN = r'(?:[一二三四五六七八九十]{1,3})'

# 纯文本标题模式 —— 按「段数」区分层级：

#   1 段（1. / 1、） → Level 2

#   2 段（1.1.）     → Level 2

#   3+段（1.1.1.）   → Level 3

# 必须按段数从多到少排列，防止短模式抢先匹配。

_HEADING_PATTERNS_READ = [

    # 第X章 / 第四章 … → 一级标题（兼容前导空格/全角空格）

    (re.compile(r'^\s*第(\d+|[一二三四五六七八九十]+)章'), 1),

    # 摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录 → 一级标题

    (re.compile(r'^\s*(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)'), 1),

    # 1.1.1. （3+ 段）→ 三级标题

    (re.compile(r'^\s*(\d+(?:\.\d+){2,})[\.\、\)）]?\s*'), 3),

    # 1.1. （2 段）→ 二级标题

    (re.compile(r'^\s*(\d+\.\d+)[\.\、\)）]?\s*'), 2),

    # 一、/ 一．→ 一级标题

    (re.compile(r'^\s*(' + _CN_NUM_PLAIN + r')[、．]\s*'), 1),

]

def _looks_like_code(text):

    """

    启发式判断文本是否像代码行（不含中文 + 含代码特征）。

    用于 .docx 读取时辅助识别无前导空格的代码行（如 def / function / 花括号行）。

    """

    stripped = text.strip()

    if not stripped:

        return False

    if re.search(r'[一-鿿]', stripped):

        return False

    code_starters = (

        'def ', 'class ', 'if ', 'for ', 'while ', 'import ', 'from ',

        'var ', 'let ', 'const ', 'function ', 'return ', 'print(',

        'public ', 'private ', 'protected ', 'static ', 'void ',

        'int ', 'string ', 'bool ', 'float ', 'double ',

        'console.', 'System.', 'using ', 'package ', '#include',

    )

    if any(stripped.startswith(kw) for kw in code_starters):

        return True

    if re.search(r'[{}();\[\]]', stripped):

        return True

    return False

def _para_has_image(para):

    """递归检查段落中是否包含图片、图形、图表等（含 Visio 绘图、mc:AlternateContent）。"""

    # 检查段落 XML 全树（含 run 内外、mc:AlternateContent 等所有嵌套）

    for child in para._element.iter():

        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag in ('drawing', 'pict', 'object', 'inline', 'anchor'):

            return True

    return False

def _format_table(tbl_element):

    """

    对深拷贝后的表格 XML 元素统一应用排版规范：

    - 对齐：单元格上下左右全居中

    - 环绕：无（嵌入式，移除 tblpPr）

    - 字号：10.5 pt（五号）

    - 边距：单元格左右边距 0.19 cm

    - 宽度：指定宽度 0 cm（自动列宽）

    """

    # 表格级别宽度设为 0 = 自动

    tblPr = tbl_element.find(qn('w:tblPr'))

    if tblPr is None:

        tblPr = OxmlElement('w:tblPr')

        tbl_element.insert(0, tblPr)

    tblW = tblPr.find(qn('w:tblW'))

    if tblW is None:

        tblW = OxmlElement('w:tblW')

        tblPr.insert(0, tblW)

    tblW.set(qn('w:w'), '0')

    tblW.set(qn('w:type'), 'auto')

    # 移除文字环绕（确保嵌入式）

    tblpPr = tblPr.find(qn('w:tblpPr'))

    if tblpPr is not None:

        tblPr.remove(tblpPr)

    # 表格整体在页面上居中

    jc = tblPr.find(qn('w:jc'))

    if jc is None:

        jc = OxmlElement('w:jc')

        tblPr.append(jc)

    jc.set(qn('w:val'), 'center')

    margin_twips = str(int(Cm(0.19).emu / 635))

    for tc in tbl_element.iter(qn('w:tc')):

        # ---- 单元格属性 ----

        tcPr = tc.find(qn('w:tcPr'))

        if tcPr is None:

            tcPr = OxmlElement('w:tcPr')

            tc.insert(0, tcPr)

        # 移除单元格固定宽度（配合表格自动列宽）

        tcW = tcPr.find(qn('w:tcW'))

        if tcW is not None:

            tcPr.remove(tcW)

        # 垂直居中

        vAlign = tcPr.find(qn('w:vAlign'))

        if vAlign is None:

            vAlign = OxmlElement('w:vAlign')

            tcPr.append(vAlign)

        vAlign.set(qn('w:val'), 'center')

        # 左右边距 0.19 cm

        tcMar = tcPr.find(qn('w:tcMar'))

        if tcMar is not None:

            tcPr.remove(tcMar)

        tcMar = OxmlElement('w:tcMar')

        for side in ('left', 'right'):

            mar = OxmlElement(f'w:{side}')

            mar.set(qn('w:w'), margin_twips)

            mar.set(qn('w:type'), 'dxa')

            tcMar.append(mar)

        tcPr.append(tcMar)

        # ---- 单元格内段落 ----

        for p in tc.iter(qn('w:p')):

            pPr = p.find(qn('w:pPr'))

            if pPr is None:

                pPr = OxmlElement('w:pPr')

                p.insert(0, pPr)

            # 水平居中

            jc = pPr.find(qn('w:jc'))

            if jc is None:

                jc = OxmlElement('w:jc')

                pPr.append(jc)

            jc.set(qn('w:val'), 'center')

            # 字号 10.5 pt（五号）= 21 半磅

            for r in p.iter(qn('w:r')):

                rPr = r.find(qn('w:rPr'))

                if rPr is None:

                    rPr = OxmlElement('w:rPr')

                    r.insert(0, rPr)

                for tag_name in ('w:sz', 'w:szCs'):

                    sz = rPr.find(qn(tag_name))

                    if sz is None:

                        sz = OxmlElement(tag_name)

                        rPr.append(sz)

                    sz.set(qn('w:val'), '21')

# ---------------------------------------------------------------------------
# 段落角色检测与异常扫描
# ---------------------------------------------------------------------------

def _is_auto_numbered(para):
    """检测段落是否启用了 Word 原生自动编号"""
    if para is None or para._element is None:
        return False
    pPr = para._element.find(qn('w:pPr'))
    if pPr is None:
        return False
    return pPr.find(qn('w:numPr')) is not None


def _detect_para_role(para, text_override=None):
    """检测段落角色：heading(标题), code(代码), list(列表), body(正文)"""
    text = text_override if text_override is not None else (para.text if para else "")

    # 1. 优先级最高：Word 原生自动编号列表
    if _is_auto_numbered(para):
        return ('list',)

    # 2. 手动列表检测 (兼容 Markdown 和纯文本)
    if re.match(r'^\s*(?:\(|（)?\d+[\.\、\)）]\s+', text) or re.match(r'^\s*[-*]\s+', text):
        return ('list',)

    # 3. 代码检测：恢复原有的严谨条件 (小字号 + 无缩进 + 像代码)
    if para is not None and para.runs:
        run = para.runs[0]
        fn = (run.font.name or '').lower()
        size = run.font.size
        indent = para.paragraph_format.first_line_indent

        if fn in _MONOSPACE_FONTS or (size and (indent is None or indent == 0) and _looks_like_code(text)):
            return ('code',)

    # 4. 标题检测
    for pattern, level in _HEADING_PATTERNS_READ:
        if pattern.match(text):
            return ('heading', level)

    # 5. 默认正文
    return ('body',)

def _detect_suspicious_vars(text):
    """异常变量扫描（供 GUI 面板日志使用）"""
    results = []
    lines = text.split('\n')
    for i, line in enumerate(lines, 1):
        if re.search(r'[A-Za-z]_[A-Za-z]', line) and '$' not in line:
            results.append((i, line.strip(), "疑似漏写公式符号 $...$"))
    return results

# ---------------------------------------------------------------------------
# 文本清洗与 AI 归一化管线 (DeepSeek, 豆包, ChatGPT 等)
# ---------------------------------------------------------------------------

def _clean_pasted_text(text):
    """清洗从网页或 AI 粘贴的底层不可见字符与 UI 干扰词"""
    text = text.replace('\x0c', '\\f').replace('\x0b', '\\v')
    text = re.sub(r'[\x00-\x08\x0e-\x1f\x7f]', '', text)
    for ch in ('​', '‌', '‍', '﻿', '\u200E', '\u200F',
               '\u202A', '\u202B', '\u202C', '\u202D', '\u202E'):
        text = text.replace(ch, "")
    text = text.replace('　', ' ')
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # 拦截豆包/Kimi 网页复制带来的冗余按钮文本
    text = re.sub(r'^\s*复制代码\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(r'(?<!^)[ \t]{2,}', ' ', text, flags=re.MULTILINE)

    # 消除 ≥4 个字符的重复短语（网页剪贴板重影），避免误伤短词
    text = re.sub(r'([a-zA-Z]{4,})\1', r'\1', text)

    return text

def _normalize_ai_markdown(text):
    """将各家 AI 特殊格式统一翻译为引擎标准格式"""
    # 剥离 DeepSeek 的思考过程
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)

    # ── 强力反转义补丁：修复 AI 平台过度转义的 Markdown 语法 ──
    text = text.replace(r'\*\*', '**')    # 恢复加粗
    text = text.replace(r'\*', '*')        # 恢复独立星号
    text = text.replace(r'\_', '_')        # 恢复下划线
    text = text.replace(r'\---', '---')    # 恢复分割线

    # 行首标记恢复
    text = re.sub(r'^(\s*)\\#', r'\1#', text, flags=re.MULTILINE)       # 恢复标题 \# -> #
    text = re.sub(r'^(\s*)\\-', r'\1-', text, flags=re.MULTILINE)       # 恢复无序列表 \- -> -
    text = re.sub(r'^(\s*\d+)\\\.', r'\1.', text, flags=re.MULTILINE)   # 恢复数字列表 1\. -> 1.

    # 修复被错误转义的 Markdown 链接（必须在公式处理之前）
    text = re.sub(r'\\\[(.*?)\\\]\(', r'[\1](', text)

    # ── 以下为原有公式修复逻辑 ──
    # 统一块级公式和行内公式
    text = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    text = re.sub(r'\\\((.*?)\\\)', r'$\1$', text)

    # 修复豆包的过度转义 \$\$
    text = re.sub(r'\\\$\\\$(.*?)\\\$\\\$', r'$$\1$$', text, flags=re.DOTALL)
    text = re.sub(r'\\\$([^$]+?)\\\$', r'$\1$', text)

    # 修复公式多余空格及嵌套错误
    text = re.sub(r'\$\s+([^\$]+?)\s+\$', r'$\1$', text)
    text = re.sub(r'^[\*\+]\s+', '- ', text, flags=re.MULTILINE)

    def _strip_bold_in_math(match):
        return f'${match.group(1).replace("**", "")}$'
    text = re.sub(r'\$\*+(.*?)\*+\$', _strip_bold_in_math, text)
    text = re.sub(r'\*\*\$([^$]+?)\$\*\*', r'$\1$', text)
    return text

# ---------------------------------------------------------------------------
# 核心排版主引擎
# ---------------------------------------------------------------------------

def convert_text_to_docx(text, output_file):
    """
    纯文本模式流水线：带表格检测功能的升级版
    """
    text = _clean_pasted_text(text)
    text = _normalize_ai_markdown(text)

    doc = Document()
    _add_cover_page(doc, title="")

    lines = text.split('\n')
    in_code_block = False
    in_table = False
    table_lines = []

    for line in lines:
        # 1. 代码块检测
        if line.strip().startswith('```'):
            in_code_block = not in_code_block
            continue
        if in_code_block:
            _add_code_line(doc, line)
            continue

        # 2. 表格检测
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            if not in_table:
                in_table = True
                table_lines = []
            table_lines.append(line)
            continue
        else:
            if in_table:
                from table_processor import add_markdown_table_to_doc
                add_markdown_table_to_doc(doc, table_lines)
                table_lines = []
                in_table = False

        # 3. 常规内容检测
        if not line.strip():
            continue

        role = _detect_para_role(None, text_override=line)
        if role[0] == 'heading':
            _add_heading(doc, line, role[1])
        elif role[0] == 'list':
            _add_body(doc, line, left_indent=Cm(0.85), first_line_indent=Cm(-0.85))
        elif role[0] == 'code':
            _add_code_line(doc, line)
        else:
            _add_body(doc, line)

    # 安全收尾：如果文档以表格结尾
    if in_table and table_lines:
        from table_processor import add_markdown_table_to_doc
        add_markdown_table_to_doc(doc, table_lines)

    doc.save(output_file)

def convert_markdown_to_docx(input_file, output_file):
    """处理 .md / .txt 文件的入口"""
    with open(input_file, 'r', encoding='utf-8') as f:
        text = f.read()
    convert_text_to_docx(text, output_file)

def reformat_docx(input_file, output_file):
    """
    重排版已有 Word 文档的入口
    【完全贯彻所见即所得，坚决阻断双重编号】
    """
    import copy
    src = Document(input_file)
    dst = Document()
    _add_cover_page(dst, title="")

    for child in src._body._body:
        if child.tag.endswith('p'):
            para = [p for p in src.paragraphs if p._element == child][0]
            text = para.text.strip()

            # 检测图像
            if _para_has_image(para):
                new_p = copy.deepcopy(child)
                # 修复图像被固定行距切掉一半的问题（设为单倍行距）
                pPr = new_p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    new_p.insert(0, pPr)

                # 移除旧的行距设置
                old_spacing = pPr.find(qn('w:spacing'))
                if old_spacing is not None:
                    pPr.remove(old_spacing)

                spacing = OxmlElement('w:spacing')
                spacing.set(qn('w:line'), '240')
                spacing.set(qn('w:lineRule'), 'auto')
                pPr.append(spacing)
                dst._body._body.append(new_p)
                continue

            if not text:
                continue

            # 使用提取的文本做身份判定
            role = _detect_para_role(para, text_override=text)

            if role[0] == 'heading':
                # 传入 source_para=para 保留原文档里的加粗/斜体
                _add_heading(dst, text, role[1], source_para=para)

            elif role[0] == 'code':
                _add_code_line(dst, text, source_para=para)

            elif role[0] == 'list':
                _add_body(dst, text, left_indent=Cm(0.85), first_line_indent=Cm(-0.85), source_para=para)

            else:
                _add_body(dst, text, source_para=para)

        elif child.tag.endswith('tbl'):
            new_tbl = copy.deepcopy(child)
            sectPr = dst._body._element.find(qn('w:sectPr'))
            if sectPr is not None:
                sectPr.addprevious(new_tbl)
            else:
                dst._body._element.append(new_tbl)
            _format_table(new_tbl)

    dst.save(output_file)

# ---------------------------------------------------------------------------
# CLI 命令行运行入口
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    print("=========================================")
    print(" FormatX - 全能 AI 格式排版引擎已启动")
    print("=========================================")
    if len(sys.argv) >= 3:
        input_file, output_file = sys.argv[1], sys.argv[2]
        ext = os.path.splitext(input_file)[1].lower()
        if ext == '.docx':
            reformat_docx(input_file, output_file)
        else:
            convert_markdown_to_docx(input_file, output_file)
        print(f"\n✅ 转换成功！文件已保存至：{output_file}")
    elif len(sys.argv) == 2:
        input_file = sys.argv[1]
        ext = os.path.splitext(input_file)[1].lower()
        output_file = input_file.replace(ext, '_formatted.docx')
        if ext == '.docx':
            reformat_docx(input_file, output_file)
        else:
            convert_markdown_to_docx(input_file, output_file)
        print(f"\n✅ 转换成功！文件已保存至：{output_file}")
    else:
        print("\n【使用方法】")
        print("1. 拖拽文件到 exe 上直接运行。")
        print("2. 命令行：python format_conversion.py <输入文件> <输出文件>")
