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
from docx.shared import Pt, Cm
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

try:
    from latex_to_omml import replace_latex_with_omml
    _HAS_OMML = True
except ImportError:
    _HAS_OMML = False
    def replace_latex_with_omml(para): return False


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

    for src_run in source_para.runs:
        # OMML 模式：保留 $...$ 原样，让 replace_latex_with_omml 处理
        if _HAS_OMML and '$' in src_run.text:
            run_text = src_run.text
        else:
            run_text = _latex_to_text(src_run.text)
        run = para.add_run(run_text)
        run.bold = src_run.bold if src_run.bold is not None else default_bold
        if src_run.italic is not None:
            run.italic = src_run.italic
        if src_run.underline is not None:
            run.underline = src_run.underline
        _set_run_font(run, cn_font=cn_font, en_font=en_font,
                      size_pt=size_pt, bold=run.bold)
    return para


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

    # 标题文本去多余空格（中文间不该有空格）
    title_text = re.sub(r'\s+', '', text.strip()) if \
        re.search(r'[一-鿿]', text) else text.strip()

    # 自动编号前缀
    if extra_prefix:
        pre_run = para.add_run(extra_prefix)
        _set_run_font(pre_run, cn_font='宋体', en_font='Times New Roman',
                      size_pt=size, bold=True)

    if source_para is not None:
        # 需要同时清理 source runs 中的空格
        for src_run in source_para.runs:
            run_text = _latex_to_text(src_run.text)
            if re.search(r'[一-鿿]', run_text):
                run_text = re.sub(r'\s+', '', run_text)
            run = para.add_run(run_text)
            run.bold = src_run.bold if src_run.bold is not None else True
            if src_run.italic is not None:
                run.italic = src_run.italic
            if src_run.underline is not None:
                run.underline = src_run.underline
            _set_run_font(run, cn_font='宋体', en_font='Times New Roman',
                          size_pt=size, bold=run.bold)
    else:
        run = para.add_run(title_text)
        _set_run_font(run, cn_font='宋体', en_font='Times New Roman',
                      size_pt=size, bold=True)
    # 将段落中的 $...$ LaTeX 转为 OMML 公式
    if _HAS_OMML:
        replace_latex_with_omml(para)


def _add_body(doc, text, alignment=None, left_indent=None,
              first_line_indent=None, source_para=None, extra_prefix=''):
    """
    添加正文段落。

    - 中文：宋体 小4号（12 pt）
    - 英文 / 数字：Times New Roman 小4号（12 pt）
    - 首行缩进 2 字符 ≈ 0.85 厘米（可通过参数覆盖）
    - 两端对齐（Justify）
    - 固定行距 20 磅
    - 若原段落有左缩进（如子项编号），则沿用原值
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

    # 自动编号前缀（如有）
    if extra_prefix:
        pre_run = para.add_run(extra_prefix)
        _set_run_font(pre_run, cn_font='宋体', en_font='Times New Roman',
                      size_pt=12, bold=True)

    if source_para is not None:
        # 正文全段加粗 → 去掉加粗（否则跟三级标题看不出区别）
        all_runs_bold = source_para.runs and all(
            r.bold for r in source_para.runs
        )
        force_no_bold = all_runs_bold
        _copy_runs(para, source_para, '宋体', 'Times New Roman', 12, False)
        if force_no_bold:
            for r in para.runs:
                r.bold = False
    else:
        run = para.add_run(text.strip())
        _set_run_font(run, cn_font='宋体', en_font='Times New Roman',
                      size_pt=12, bold=False)
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
    # 第X章 / 第四章 … → 一级标题
    (re.compile(r'^第(\d+|[一二三四五六七八九十]+)章'), 1),
    # 摘要 / Abstract / 目录 / 参考文献 / 致谢 / 附录 → 一级标题
    (re.compile(r'^(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)'), 1),
    # 1.1.1. （3+ 段）→ 三级标题
    (re.compile(r'^(\d+(?:\.\d+){2,})[\.\、\)）]?\s*'), 3),
    # 1.1. （2 段）→ 二级标题
    (re.compile(r'^(\d+\.\d+)[\.\、\)）]?\s*'), 2),
    # 1. / 1、 （1 段）→ 三级标题
    (re.compile(r'^(\d+)[\.\、\)）]\s*'), 3),
    # 一、/ 一．→ 二级标题
    (re.compile(r'^(' + _CN_NUM_PLAIN + r')[、．]\s*'), 2),
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


def _detect_para_role(para, text_override=None):
    """
    检测已有 .docx 中一个段落的角色。

    text_override 用于传入拼接了自动编号后的文本，使模式匹配能看到编号。
    若为 None 则使用 para.text。

    返回
    ----
    ('heading', level)   标题（level=1/2/3）
    ('code',)            代码行
    ('body',)            正文
    """
    text = text_override if text_override is not None else para.text
    if not text.strip():
        return ('body',)

    stripped = text.strip()
    # 硬限制：除去空格后纯文本超 30 字不可能是标题（第X章除外）
    text_no_spaces = re.sub(r'\s+', '', stripped)
    if len(text_no_spaces) > 30 and not re.match(r'^第(\d+|[一二三四五六七八九十]+)章', stripped):
        return ('body',)

    # --- 方式 1：代码识别（最优先，代码就是代码） ---
    if para.runs:
        run = para.runs[0]
        fn = (run.font.name or '').lower()
        size = run.font.size
        indent = para.paragraph_format.first_line_indent

        if fn in _MONOSPACE_FONTS:
            return ('code',)
        if size and size < Pt(12) and (indent is None or indent == 0):
            return ('code',)
        if size and indent is not None and indent == 0 and _looks_like_code(text):
            return ('code',)

    # --- 方式 2：内容模式优先于样式名 ---
    # 原文可能把不同层级标题都设成同一 Word 样式（如全是 Heading 3），
    # 导致样式名不可信。先用编号格式（5.1 → L2, 5.1.1 → L3）判断。
    stripped = text.strip()
    for pattern, level in _HEADING_PATTERNS_READ:
        if pattern.match(stripped):
            if len(stripped) < 80:
                return ('heading', level)
            break

    # --- 方式 3：通过 Word 内置样式名判断 ---
    # 样式名仅供参考——若内容特征明显是正文（以句号结尾、长段落等），降级为正文
    style_name = (para.style.name if para.style else '').lower()
    style_level = None
    if 'heading 1' in style_name or '标题 1' in style_name:
        style_level = 1
    elif 'heading 2' in style_name or '标题 2' in style_name:
        style_level = 2
    elif 'heading 3' in style_name or '标题 3' in style_name:
        style_level = 3

    if style_level is not None:
        # 有样式名：再用内容特征复核——长文/以标点结尾/有缩进 → 正文
        _style_body = (
            (indent is not None and indent > 0) or
            stripped.endswith(('。', '；', '！', '？', '：')) or
            len(stripped) >= 40  # 中文正文通常超过 40 字
        )
        if not _style_body:
            return ('heading', style_level)

    # --- 方式 4：通过字体属性判断 ---
    if para.runs:
        bold = run.bold
        # 多 run 段落：只有全部 run 都加粗才可能是标题
        multi_run_not_all_bold = (
            len(para.runs) > 1 and not all(r.bold for r in para.runs)
        )

        # 正文特征：以句号/分号结尾、或首行有缩进
        looks_like_body = (
            (indent is not None and indent > 0) or
            stripped.endswith(('。', '；', '！', '？', '：'))
        )
        # 标题识别：加粗 + 字号 ≥ 12pt + 内容较短 + 非正文特征 + 非部分加粗
        if (bold and size and size >= Pt(12) and len(stripped) < 80
                and not multi_run_not_all_bold
                and not looks_like_body):
            if indent is None or indent == 0:
                if size >= Pt(16):
                    return ('heading', 1)
                if size >= Pt(14):
                    return ('heading', 2)
                return ('heading', 3)

    return ('body',)


# ---------------------------------------------------------------------------
# .docx 输入：提取自动编号
# ---------------------------------------------------------------------------

def _format_list_number(count, fmt):
    """将计数器值按格式转为字符串（仅处理常见中文文档编号格式）。"""
    if fmt == 'decimal':
        return str(count)
    if fmt == 'lowerLetter':
        return chr(ord('a') + count - 1) if 1 <= count <= 26 else str(count)
    if fmt == 'upperLetter':
        return chr(ord('A') + count - 1) if 1 <= count <= 26 else str(count)
    if fmt in ('lowerRoman', 'upperRoman'):
        # 罗马数字简化映射（1~30 足够文档使用）
        roman_map = [
            (100, 'c'), (90, 'xc'), (50, 'l'), (40, 'xl'),
            (10, 'x'), (9, 'ix'), (5, 'v'), (4, 'iv'), (1, 'i')
        ]
        n = count
        result = ''
        for val, chars in roman_map:
            while n >= val:
                result += chars
                n -= val
        return result.upper() if fmt == 'upperRoman' else result
    if fmt == 'chineseCounting':
        digits = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十']
        if count <= 10:
            return digits[count]
        if count < 20:
            return '十' + digits[count - 10]
        if count < 100:
            tens = count // 10
            ones = count % 10
            return digits[tens] + '十' + digits[ones]
        return str(count)
    # 兜底：直接数字
    return str(count)


def _extract_auto_numbers(doc):
    """
    从 Word 自动编号定义中提取每个段落的编号文本。

    Word 的自动编号（w:numPr）不会出现在 para.text 中，
    必须解析 numbering.xml 定义并追踪各级计数器才能还原。

    返回
    ----
    dict[int, str]   段落索引 → 编号文本（如 '2.1.1 '）
    """
    try:
        numbering_part = doc.part.numbering_part
        num_root = numbering_part._element
    except (NotImplementedError, AttributeError, Exception):
        return {}  # 文档编号定义无法解析，跳过自动编号

    # ----- 解析抽象编号定义：abstractNumId → {ilvl: {numFmt, lvlText}} -----
    abstract_nums = {}
    for abs_elem in num_root.findall(qn('w:abstractNum')):
        abs_id = abs_elem.get(qn('w:abstractNumId'))
        levels = {}
        for lvl_elem in abs_elem.findall(qn('w:lvl')):
            ilvl = int(lvl_elem.get(qn('w:ilvl')))
            numFmt_el = lvl_elem.find(qn('w:numFmt'))
            lvlText_el = lvl_elem.find(qn('w:lvlText'))
            levels[ilvl] = {
                'numFmt': numFmt_el.get(qn('w:val')) if numFmt_el is not None else 'decimal',
                'lvlText': lvlText_el.get(qn('w:val')) if lvlText_el is not None else '%1.',
            }
        abstract_nums[abs_id] = levels

    # ----- 编号实例：numId → abstractNumId -----
    num_to_abs = {}
    for num_elem in num_root.findall(qn('w:num')):
        nid = num_elem.get(qn('w:numId'))
        ref = num_elem.find(qn('w:abstractNumId'))
        if ref is not None:
            num_to_abs[nid] = ref.get(qn('w:val'))

    # ----- 逐段扫描，追踪计数器 -----
    counters = {}       # (numId, ilvl) → 当前计数值
    result = {}         # para_index → 编号文本

    for idx, para in enumerate(doc.paragraphs):
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            continue
        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            continue

        numId_el = numPr.find(qn('w:numId'))
        ilvl_el = numPr.find(qn('w:ilvl'))
        if numId_el is None:
            continue

        numId = numId_el.get(qn('w:val'))
        ilvl = int(ilvl_el.get(qn('w:val'))) if ilvl_el is not None else 0

        # 一个段落有编号时，重置比它更深的所有级别计数器
        for (nid, lv), _ in list(counters.items()):
            if nid == numId and lv > ilvl:
                del counters[(nid, lv)]

        # 递增本级计数器
        key = (numId, ilvl)
        counters[key] = counters.get(key, 0) + 1

        # 按抽象编号的 lvlText 模板替换生成最终编号文本
        abs_id = num_to_abs.get(numId)
        if abs_id and abs_id in abstract_nums:
            lvl_info = abstract_nums[abs_id].get(ilvl, {})
            template = lvl_info.get('lvlText', '%1.')
            number_str = template
            for lv in range(ilvl + 1):
                placeholder = f'%{lv + 1}'
                cnt = counters.get((numId, lv), 1)
                fmt = abstract_nums[abs_id].get(lv, {}).get('numFmt', 'decimal')
                formatted = _format_list_number(cnt, fmt)
                number_str = number_str.replace(placeholder, formatted)
            result[idx] = number_str

    return result


def _normalize_number_prefix(prefix):
    """
    将含点的混合编号归一化（"一.1." → "1.1."），
    但保留纯中文编号不动（"第一章" / "一、" 等不转换）。
    """
    # 纯中文章节编号不转换：第一章 / 一、 / 一.
    if not re.search(r'\d', prefix):
        return prefix
    cn_map = {
        '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
        '六': '6', '七': '7', '八': '8', '九': '9', '十': '10',
        '十一': '11', '十二': '12', '十三': '13', '十四': '14',
        '十五': '15', '十六': '16', '十七': '17', '十八': '18',
        '十九': '19', '二十': '20',
    }
    result = prefix
    for cn, ar in sorted(cn_map.items(), key=lambda x: -len(x[0])):
        result = result.replace(cn, ar)
    return result


# ---------------------------------------------------------------------------
# .docx 输入：页眉页脚拷贝
# ---------------------------------------------------------------------------

def _copy_image_blobs(src, dst):
    """
    将源文档所有图片的二进制数据拷贝到目标文档。
    返回 old_rId → new_rId 映射，因为目标文档的 rId 可能已被占用。
    """
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    rId_map = {}
    for rId, rel in src.part.rels.items():
        if 'image' not in rel.reltype:
            continue
        try:
            src_img = rel.target_part
            base = os.path.basename(src_img.partname)
            existing_names = set()
            for r in dst.part.rels.values():
                tp = getattr(r, 'target_part', None)
                if tp:
                    existing_names.add(tp.partname)
            name, ext = os.path.splitext(base)
            partname = f'/word/media/{base}'
            counter = 1
            while partname in existing_names:
                partname = f'/word/media/{name}_{counter}{ext}'
                counter += 1
            new_img = Part(
                PackURI(partname),
                src_img.content_type,
                src_img.blob,
                dst.part.package,
            )
            new_rId = dst.part.relate_to(new_img, rel.reltype)
            rId_map[rId] = new_rId
        except Exception:
            pass  # 损坏的图片引用跳过
    return rId_map


def _copy_headers_footers(src, dst, img_rId_map):
    """
    将源文档所有节的页眉/页脚深拷贝到目标文档（含 logo、校训、页码等）。
    """
    # 确保目标文档节数与源文档一致
    while len(dst.sections) < len(src.sections):
        dst.add_section()

    for i, src_sec in enumerate(src.sections):
        dst_sec = dst.sections[i]

        # 默认页眉
        if src_sec.header:
            _copy_header_part(src_sec.header, dst_sec.header, img_rId_map)
        # 首页不同页眉
        if src_sec.different_first_page_header_footer:
            dst_sec.different_first_page_header_footer = True
            if src_sec.first_page_header:
                _copy_header_part(src_sec.first_page_header,
                                  dst_sec.first_page_header, img_rId_map)
        # 奇偶页不同
        if (hasattr(src_sec, 'even_page_header') and src_sec.even_page_header):
            dst_sec.different_even_odd_header_footer = True
            _copy_header_part(src_sec.even_page_header,
                              dst_sec.even_page_header, img_rId_map)

        # 默认页脚
        if src_sec.footer:
            _copy_header_part(src_sec.footer, dst_sec.footer, img_rId_map)
        # 首页不同页脚
        if src_sec.different_first_page_header_footer and src_sec.first_page_footer:
            _copy_header_part(src_sec.first_page_footer,
                              dst_sec.first_page_footer, img_rId_map)


def _copy_header_part(src_part, dst_part, img_rId_map):
    """深拷贝页眉/页脚的全部内容元素，并重映射图片 rId。"""
    import copy
    src_element = src_part._element
    dst_element = dst_part._element
    for child in list(dst_element):
        dst_element.remove(child)
    for child in src_element:
        child_copy = copy.deepcopy(child)
        for blip in child_copy.iter(qn('a:blip')):
            old_embed = blip.get(qn('r:embed'))
            if old_embed and old_embed in img_rId_map:
                blip.set(qn('r:embed'), img_rId_map[old_embed])
        dst_element.append(child_copy)


# ---------------------------------------------------------------------------
# .docx 输入：重新格式化
# ---------------------------------------------------------------------------

def reformat_docx(input_path, output_path):
    """
    读取已有 .docx 文件，检测每个段落的角色（标题 / 代码 / 正文），
    然后按照严格排版规范重新输出一个新的 .docx 文件。

    检测策略（优先级从高到低）：
        1. Word 内置样式名（Heading 1 / 标题 1 等）
        2. 字体属性（加粗 + 字号大小 → 标题级别；等宽字体 → 代码）
        3. 内容模式（数字.  /  一、 等纯文本标题）

    参数
    ----
    input_path : str
        输入 .docx 文件路径。
    output_path : str
        输出 .docx 文件路径。
    """
    import copy
    src = Document(input_path)
    dst = Document()

    # 将源文档图片文件拷贝到目标文档（rId 自动重映射）
    img_rId_map = _copy_image_blobs(src, dst)
    # 将源文档的页眉/页脚拷贝到目标文档（含 logo、校训等）
    _copy_headers_footers(src, dst, img_rId_map)

    def _append_preserved(dst_body, element):
        """将保留元素插入 sectPr 之前，确保文档顺序正确。"""
        children = list(dst_body)
        for i, c in enumerate(children):
            tag = c.tag.split('}')[-1] if '}' in c.tag else c.tag
            if tag == 'sectPr':
                dst_body.insert(i, element)
                return
        dst_body.append(element)

    # 提取 Word 自动编号（w:numPr），这些编号不出现在 para.text 中
    auto_nums = _extract_auto_numbers(src)

    # 不另建封面；原文档的封面部分保留不动

    # 按文档 body 原始顺序遍历：段落 + 表格 + 其他元素
    para_counter = 0
    prev_was_heading = False  # 标题下一行强制为正文
    body = src._element.body
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        # 源文档的 sectPr 跳过，目标文档有自己的
        if tag == 'sectPr':
            continue

        if tag == 'tbl':
            tbl_copy = copy.deepcopy(child)
            _format_table(tbl_copy)
            _append_preserved(dst._element.body, tbl_copy)
            continue

        if tag != 'p':
            _append_preserved(dst._element.body, copy.deepcopy(child))
            continue

        # --- 以下处理段落 (w:p) ---
        para = src.paragraphs[para_counter]
        text = _latex_to_text(para.text)

        # 段落含图片/图形 → 原文深拷贝，rId 重映射后不重新格式化
        if _para_has_image(para):
            img_para = copy.deepcopy(child)
            for blip in img_para.iter(qn('a:blip')):
                old_embed = blip.get(qn('r:embed'))
                if old_embed and old_embed in img_rId_map:
                    blip.set(qn('r:embed'), img_rId_map[old_embed])
            _append_preserved(dst._element.body, img_para)
            para_counter += 1
            continue

        # 自动编号提取
        num_prefix = auto_nums.get(para_counter, '')
        has_auto_num = bool(num_prefix)
        if has_auto_num:
            num_prefix = _normalize_number_prefix(num_prefix)

        # 【核心修复】：软回车（\\n）拆分——标题和正文粘在一起时逐行独立判断
        lines = [ln for ln in text.split('\n') if ln.strip()]
        if not lines:
            para_counter += 1
            continue

        orig_align = para.paragraph_format.alignment
        orig_left_indent = para.paragraph_format.left_indent

        for idx, sub_text in enumerate(lines):
            current_has_auto = (has_auto_num if idx == 0 else False)
            if current_has_auto:
                sub_text = num_prefix + ' ' + sub_text

            role = _detect_para_role(para, text_override=sub_text)

            if prev_was_heading and role[0] == 'heading':
                clean_text = sub_text.strip()
                _is_clear_heading = bool(
                    re.match(r'^第(\d+|[一二三四五六七八九十]+)章', clean_text) or
                    re.match(r'^(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)', clean_text) or
                    re.match(r'^(\d+(?:\.\d+)+)', clean_text)
                )
                if not _is_clear_heading:
                    role = ('body',)

            prev_was_heading = (role[0] == 'heading')

            # 软回车拆成多行时，舍弃原始加粗格式避免串味
            use_source_para = para if len(lines) == 1 else None

            if role[0] == 'heading':
                h_align = orig_align if (
                    re.match(r'^第(\d+|[一二三四五六七八九十]+)章', sub_text) or
                    re.match(r'^(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)', sub_text)
                ) else None
                _add_heading(dst, sub_text, role[1], alignment=h_align,
                             source_para=use_source_para)
            elif role[0] == 'code':
                _add_code_line(dst, sub_text, source_para=use_source_para)
            else:
                if current_has_auto:
                    _add_body(dst, sub_text, alignment=orig_align,
                              left_indent=None, first_line_indent=Cm(0),
                              source_para=use_source_para)
                else:
                    _add_body(dst, sub_text, alignment=orig_align,
                              left_indent=orig_left_indent,
                              source_para=use_source_para)

        para_counter += 1

    dst.save(output_path)
    print(f'转换完成：{input_path} → {output_path}')


# ---------------------------------------------------------------------------
# 主转换逻辑（Markdown / 纯文本 → .docx）
# ---------------------------------------------------------------------------

def convert_markdown_to_docx(input_path, output_path):
    """
    将 Markdown 纯文本文件转换为格式严格的 Word 文档。

    支持的 Markdown 语法：
        # 一级标题（宋体 16pt 加粗）
        ## 二级标题（宋体 14pt 加粗）
        ### 三级标题（宋体 12pt 加粗）
        ``` ... ``` 代码块（可带语言标识，如 ```python）
        其余为正文段落

    纯文本自动识别（当文件中不含 # 标题时自动启用）：
        1. xxx  /  1、xxx  /  1) xxx      → 二级标题（14pt 加粗）
        1.1 xxx  /  1.1.1 xxx             → 三级标题（12pt 加粗）
        一、xxx  /  二．xxx                → 二级标题（14pt 加粗）

    参数
    ----
    input_path : str
        输入 .md / .txt 文件的路径（UTF-8 编码）。
    output_path : str
        输出 .docx 文件的路径。
    """
    # 读入全文
    with open(input_path, 'r', encoding='utf-8') as f:
        raw_text = f.read()

    # 统一换行符
    raw_text = raw_text.replace('\r\n', '\n').replace('\r', '\n')
    lines = raw_text.split('\n')

    doc = Document()

    # ------------------------------------------------------------------
    # 封面预留页
    # ------------------------------------------------------------------
    _add_cover_page(doc)

    # ------------------------------------------------------------------
    # 自动检测：文件中是否有 # 式 Markdown 标题
    # ------------------------------------------------------------------
    has_md_heading = any(
        re.match(r'^#{1,3}\s+', ln) for ln in lines
    )

    # ------------------------------------------------------------------
    # 纯文本标题识别（仅在无 # 标题时启用）
    # ------------------------------------------------------------------
    # 中文序号：一、二、三 … 十、十一 … 二十 …
    _cn_num = r'(?:[一二三四五六七八九十]{1,3})'
    PLAIN_HEADING_PATTERNS = [
        # 第X章 / 第四章 xxx → 一级标题
        (re.compile(r'^第(\d+|[一二三四五六七八九十]+)章\s*(.*)$'), 1),
        # 摘要 / Abstract / 目录 / 参考文献 → 一级标题
        (re.compile(r'^(摘要|Abstract|ABSTRACT|目录|参考文献|致谢|附录)(.*)$'), 1),
        # 1.1.1 xxx（3+ 段）→ 三级标题
        (re.compile(r'^(\d+(?:\.\d+){2,})[\.\、\)）]?\s*(.*)$'), 3),
        # 1.1 xxx（2 段）→ 二级标题
        (re.compile(r'^(\d+\.\d+)[\.\、\)）]?\s*(.*)$'), 2),
        # 1. xxx / 1、xxx（1 段）→ 三级标题
        (re.compile(r'^(\d+)[\.\、\)）]\s*(.*)$'), 3),
        # 一、xxx / 一．xxx → 二级标题
        (re.compile(r'^(' + _cn_num + r')[、．]\s*(.*)$'), 2),
    ]

    # ------------------------------------------------------------------
    # 逐行解析
    # ------------------------------------------------------------------
    in_code_block = False
    code_buffer = []
    prev_was_heading_md = False

    MD_HEADING = re.compile(r'^(#{1,3})\s+(.*)$')

    for line in lines:
        stripped = _latex_to_text(line.strip())

        # --- 代码块边界 ---
        if stripped.startswith('```'):
            if not in_code_block:
                in_code_block = True
                code_buffer = []
            else:
                in_code_block = False
                for code_line in code_buffer:
                    _add_code_line(doc, code_line)
                code_buffer = []
            continue

        # --- 代码块内部 ---
        if in_code_block:
            code_buffer.append(line)
            continue

        # --- 空行 ---
        if not stripped:
            continue

        # --- Markdown 标题（始终生效） ---
        #   # → 一级标题   ## → 二级标题   ### → 三级标题
        m = MD_HEADING.match(line)
        if m:
            md_level = len(m.group(1))  # 1~3
            _add_heading(doc, _latex_to_text(m.group(2)), md_level)
            continue

        # --- 连续编号拆分：正文里 "1. xxx 2. yyy" → 拆行但保持正文格式 ---
        _SPLIT_NUMS = re.compile(r'(?<!\d)(\d+)\.\s*')
        # 用原始行文本拆分（保留 $...$ 供 OMML 使用）
        raw_line = line.strip()
        pieces = _SPLIT_NUMS.split(raw_line)
        if len(pieces) >= 5:
            for i in range(1, len(pieces) - 1, 2):
                num = pieces[i]
                content = pieces[i + 1].strip() if i + 1 < len(pieces) else ''
                next_match = _SPLIT_NUMS.search(content)
                if next_match:
                    content = content[:next_match.start()].strip()
                piece_text = _latex_to_text(f'{num}. {content}')
                _add_body(doc, piece_text)
            continue

        # --- 纯文本标题（始终作为 Markdown 标题的补充） ---
        matched = False
        for pattern, level in PLAIN_HEADING_PATTERNS:
            m = pattern.match(line)
            if m:
                if prev_was_heading_md and level >= 2 and not re.match(r'^(\d+(?:\.\d+)+)', stripped):
                    break
                _add_heading(doc, stripped, level)
                prev_was_heading_md = True
                matched = True
                break
        if matched:
            continue

        prev_was_heading_md = False

        # --- 其余：正文段落 ---
        # OMML 模式：用原始文本（含 $...$），让 replace_latex_with_omml 处理公式
        if _HAS_OMML and '$' in line:
            body_text = _apply_latex_commands(line.strip())
        else:
            body_text = stripped
        _add_body(doc, body_text)

    # ------------------------------------------------------------------
    # 处理未闭合的代码块（容错：当作普通代码行输出）
    # ------------------------------------------------------------------
    if in_code_block and code_buffer:
        for code_line in code_buffer:
            _add_code_line(doc, code_line)

    # 保存
    doc.save(output_path)
    print(f'转换完成：{input_path} → {output_path}')


def _clean_pasted_text(text):
    """
    清洗从网页/AI 粘贴的文本：
    - 零宽字符、HTML 实体、智能引号 → 标准字符
    - 全角空格 → 半角
    - 统一换行
    """
    # 零宽字符
    text = text.replace('​', '').replace('‌', '').replace('‍', '')
    text = text.replace('﻿', '').replace('\u200E', '').replace('\u200F', '')
    # HTML 实体
    import html
    text = html.unescape(text)
    # 智能引号 → 直引号
    text = text.replace('“', '"').replace('”', '"')
    text = text.replace('‘', "'").replace('’', "'")
    # 全角空格 → 半角
    text = text.replace('　', ' ')
    # 特殊连字符
    text = text.replace('–', '--').replace('—', '---')
    # 换行统一
    text = text.replace('\r\n', '\n').replace('\r', '\n')
    return text


def convert_text_to_docx(text, output_path):
    """
    接收字符串文本，写入临时文件后调用 convert_markdown_to_docx。
    供 GUI / API 等内存字符串场景使用。
    先清洗粘贴文本中的隐形垃圾字符。
    """
    text = _clean_pasted_text(text)
    import tempfile
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md',
                                     encoding='utf-8', delete=False) as tf:
        tf.write(text)
        tmp_path = tf.name
    try:
        convert_markdown_to_docx(tmp_path, output_path)
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    return output_path


def preview_text(text):
    """
    LaTeX→Unicode 预览，不生成 .docx。
    返回转换后的纯文本，供 GUI 预览使用。
    """
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    result = []
    in_code = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith('```'):
            in_code = not in_code
            result.append('——— 代码块 ———')
            continue
        if in_code:
            result.append(line)
            continue
        if not stripped:
            result.append('')
            continue
        # 用 _latex_to_text 完整转换（含 $...$ 公式）
        result.append(_latex_to_text(stripped))
    return '\n'.join(result)


# ---------------------------------------------------------------------------
# GUI 入口
# ---------------------------------------------------------------------------

def _pick_file_gui():
    """弹出文件选择框，返回 (input_path, output_path)。"""
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        return None, None

    root = tk.Tk()
    root.withdraw()
    input_file = filedialog.askopenfilename(
        title='选择要转换的文件',
        filetypes=[('文档', '*.md *.txt *.docx'), ('所有文件', '*.*')]
    )
    if not input_file:
        return None, None
    ext = os.path.splitext(input_file)[1]
    output_file = input_file.replace(ext, '_formatted.docx')
    return input_file, output_file


def _run_conversion(input_file, output_file):
    """执行转换并弹出结果提示。"""
    if not os.path.exists(input_file):
        print(f'错误：输入文件不存在 —— {input_file}')
        return False
    ext = os.path.splitext(input_file)[1].lower()
    if ext == '.docx':
        reformat_docx(input_file, output_file)
    else:
        convert_markdown_to_docx(input_file, output_file)
    return True


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

DEFAULT_INPUT  = 'test2.docx'
DEFAULT_OUTPUT = 'output.docx'

if __name__ == '__main__':
    if len(sys.argv) >= 3:
        # 命令行：python format_conversion.py input.md output.docx
        input_file, output_file = sys.argv[1], sys.argv[2]
    elif len(sys.argv) == 2:
        # 拖拽文件到 exe：python format_conversion.py input.md
        input_file = sys.argv[1]
        ext = os.path.splitext(input_file)[1]
        output_file = input_file.replace(ext, '_formatted.docx')
    else:
        # 无参数：弹出 GUI 文件选择框
        input_file, output_file = _pick_file_gui()
        if not input_file:
            print('未选择文件，已退出。')
            sys.exit(0)

    ok = _run_conversion(input_file, output_file)
    if ok:
        # GUI 模式弹出完成提示
        try:
            import tkinter.messagebox as mb
            mb.showinfo('完成', f'已保存到：\n{output_file}')
        except Exception:
            pass
