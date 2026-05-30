"""
FormatX 公式预归一化器 — 位于管线最前端，主动预防破损公式导致的解析错误
"""
import re

from core.symbols import UNICODE_MATH_TO_LATEX

_RE_LATEX_COMMAND = re.compile(r"\\[A-Za-z]+\b")
_RE_UNICODE_SUPERSCRIPT = re.compile(
    r"([一-鿿A-Za-z0-9)]+)([²³⁰ⁱ⁴-⁹⁺-⁾])"
)
_RE_CHINESE = re.compile(r"[一-鿿]")
_RE_MATH_OPERATORS = re.compile(
    r"[=+\-*/<>±×÷∈∑∫√]"
)

_UNICODE_SUPER_MAP = {
    "²": "^2", "³": "^3",
    "⁰": "^0", "ⁱ": "^1",
    "⁴": "^4", "⁵": "^5", "⁶": "^6",
    "⁷": "^7", "⁸": "^8", "⁹": "^9",
    "⁺": "^+", "⁻": "^-", "⁼": "^=",
    "⁽": "^(", "⁾": "^)",
}


def looks_like_formula_text(text: str) -> bool:
    raw = text.strip()
    if raw.startswith("$") and raw.endswith("$"):
        return False
    if _RE_LATEX_COMMAND.search(raw):
        return True
    if any(char in raw for char in UNICODE_MATH_TO_LATEX):
        return True
    return False


def fix_ocr_spaces(text: str) -> str:
    res = re.sub(r"(\b[A-Za-z])\s+(\d)\b", r"\1^\2", text)
    return res


def text_formula_to_latex(text: str) -> str:
    if not text:
        return text
    res = text
    for uni_char, latex_cmd in UNICODE_MATH_TO_LATEX.items():
        res = res.replace(uni_char, latex_cmd)

    def _sub_super(match):
        base = match.group(1)
        sup = match.group(2)
        trans_sup = _UNICODE_SUPER_MAP.get(sup, sup)
        return f"{base}{trans_sup}"

    res = _RE_UNICODE_SUPERSCRIPT.sub(_sub_super, res)
    res = fix_ocr_spaces(res)
    return res


def analyze_formula_mixing(text: str) -> tuple:
    raw = text.strip()
    if not raw:
        return False, False
    if (raw.startswith("$") and raw.endswith("$")) or (
        raw.startswith("\\[") and raw.endswith("\\]")
    ):
        return False, False

    has_math = bool(
        _RE_LATEX_COMMAND.search(raw)
        or any(char in raw for char in UNICODE_MATH_TO_LATEX)
        or _RE_MATH_OPERATORS.search(raw)
    )
    if not has_math:
        return False, False

    chinese_count = len(_RE_CHINESE.findall(raw))
    clean_text = _RE_LATEX_COMMAND.sub("", raw)
    for func in ["sin","cos","tan","log","ln","lim","exp","max","min"]:
        clean_text = re.sub(rf"\b{func}\b", "", clean_text, flags=re.IGNORECASE)

    english_word_count = len(re.findall(r"\b[A-Za-z]{4,}\b", clean_text))
    text_score = chinese_count + english_word_count
    is_pure = text_score == 0 or (text_score / len(raw) < 0.05)

    return has_math, is_pure


def auto_wrap_and_normalize_context(text: str) -> str:
    lines = text.split("\n")
    processed = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            processed.append(line)
            continue

        def _repair_existing(match):
            return f"${text_formula_to_latex(match.group(1))}$"

        line_clean = re.sub(r"\$(.*?)\$", _repair_existing, line)

        has_math, is_pure = analyze_formula_mixing(line_clean)

        if is_pure:
            normalized = text_formula_to_latex(stripped)
            line_clean = f"${normalized}$"
        elif has_math:
            inline_pat = (
                r"(?:"
                r"(?:\\[A-Za-z]+|[Ͱ-Ͽἀ-῿]"
                r"|[a-zA-Z0-9_\^\+\-\*/=\<>\(\)\{\}\[\]\.,])+"
                r"[\s\+\-\*/=\<\>]*"
                r")+"
            )

            def _mixed_filter(match):
                chunk = match.group(0)
                if re.match(r'^\*{1,3}$', chunk.strip()):
                    return chunk
                if re.search(
                    r"\\[A-Za-z]+|[=+\-*/<>²³⁴-⁹]",
                    chunk,
                ):
                    if len(re.findall(r"(?<!\\)\b[A-Za-z]{4,}\b", chunk)) < 2:
                        return f"${text_formula_to_latex(chunk.strip())}$"
                return chunk

            line_clean = re.sub(inline_pat, _mixed_filter, line_clean)

        processed.append(line_clean)

    return "\n".join(processed)
