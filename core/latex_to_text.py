"""
LaTeX → Unicode 纯文本转换器
将简单 LaTeX 数学公式转为可读纯文本，供 _copy_runs 使用
"""
import re


def _to_unicode_scripts(text):
    """将 ^{x} / _{y} / \bar{A} 转为 Unicode 上/下标/重音符。"""
    SUPER = str.maketrans('0123456789+-=()ij', '⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁱʲ')
    SUB = str.maketrans('0123456789+-=()in', '₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ᵢₙ')
    OVERLINE = '̅'

    def _replace_sup(m):   return m.group(1).translate(SUPER)
    def _replace_sub(m):   return m.group(1).translate(SUB)
    def _replace_bar(m):   return m.group(1) + OVERLINE
    def _replace_hat(m):   return m.group(1) + '̂'
    def _replace_tilde(m): return m.group(1) + '̃'
    def _replace_vec(m):   return m.group(1) + '⃗'
    def _replace_dot(m):   return m.group(1) + '̇'

    text = re.sub(r'\\bar\{([^}]+)\}', _replace_bar, text)
    text = re.sub(r'\\bar\s+(\w)', _replace_bar, text)
    text = re.sub(r'\\hat\{([^}]+)\}', _replace_hat, text)
    text = re.sub(r'\\hat\s+(\w)', _replace_hat, text)
    text = re.sub(r'\\tilde\{([^}]+)\}', _replace_tilde, text)
    text = re.sub(r'\\tilde\s+(\w)', _replace_tilde, text)
    text = re.sub(r'\\vec\{([^}]+)\}', _replace_vec, text)
    text = re.sub(r'\\vec\s+(\w)', _replace_vec, text)
    text = re.sub(r'\\dot\{([^}]+)\}', _replace_dot, text)
    text = re.sub(r'\\dot\s+(\w)', _replace_dot, text)
    text = re.sub(r'_\{([^}]+)\}', _replace_sub, text)
    text = re.sub(r'\^\{([^}]+)\}', _replace_sup, text)
    text = re.sub(r'_(\d+)', _replace_sub, text)
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
    """将简单 LaTeX 数学公式转为可读的纯文本。"""
    def _replace_frac(s):
        result = []
        i = 0
        while i < len(s):
            if s[i:i + 5] == '\\frac':
                j = i + 5
                while j < len(s) and s[j] != '{':
                    j += 1
                if j >= len(s):
                    result.append(s[i:])
                    break
                depth = 0
                k = j
                while k < len(s):
                    if s[k] == '{': depth += 1
                    elif s[k] == '}':
                        depth -= 1
                        if depth == 0: break
                    k += 1
                numerator = s[j + 1:k]
                m = k + 1
                while m < len(s) and s[m] != '{':
                    m += 1
                depth = 0
                n = m
                while n < len(s):
                    if s[n] == '{': depth += 1
                    elif s[n] == '}':
                        depth -= 1
                        if depth == 0: break
                    n += 1
                denominator = s[m + 1:n]
                numerator = _replace_frac(numerator)
                denominator = _replace_frac(denominator)
                result.append(f'({numerator})/({denominator})')
                i = n + 1
            else:
                result.append(s[i])
                i += 1
        return ''.join(result)

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

    text = _apply_latex_commands(text)
    text = re.sub(r'\$\$([^$]+)\$\$', _convert_dollar, text)
    text = re.sub(r'\$([^$]+)\$', _convert_dollar, text)
    return text


def _apply_latex_commands(text):
    """将 $...$ 外的 LaTeX 命令转为 Unicode（$...$ 内保持原样供 OMML 处理）。"""
    parts = re.split(r'(\$[^$]+\$|\$\$[^$]+\$\$)', text)
    for i, part in enumerate(parts):
        if part.startswith('$'):
            continue
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
