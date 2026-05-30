"""FormatX 动态大纲序号生成器"""
from core.utils.chinese import int_to_chinese

_CN_UPPER_MAP = str.maketrans({
    "一": "壹", "二": "贰", "三": "叁", "四": "肆", "五": "伍",
    "六": "陆", "七": "柒", "八": "捌", "九": "玖", "十": "拾",
})
_CIRCLED_DIGITS = {idx: ch for idx, ch in enumerate(
    "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳", start=1)}
_CIRCLED_PAREN_DIGITS = {idx: ch for idx, ch in enumerate(
    "⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇", start=1)}


def _to_roman(value: int) -> str:
    if value <= 0:
        return ""
    mapping = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    result = ""
    for threshold, symbol in mapping:
        while value >= threshold:
            result += symbol
            value -= threshold
    return result


class HeadingNumberingManager:
    def __init__(self, level_configs: dict):
        self.configs = {}
        for k, v in (level_configs or {}).items():
            try:
                self.configs[int(k)] = v
            except (ValueError, TypeError):
                pass
        self.counters = {i: 0 for i in range(1, 10)}

    def get_next_number(self, level: int) -> str:
        if level not in self.counters:
            return ""

        self.counters[level] += 1
        for l in range(level + 1, 10):
            self.counters[l] = 0

        style = self.configs.get(level, "arabic")
        current_val = self.counters[level]

        if style == "arabic":
            return str(current_val)
        if style == "arabic_pad2":
            return f"{current_val:02d}"

        if style == "arabic_dotted":
            prefix = ".".join(str(self.counters[i]) for i in range(1, level))
            return f"{prefix}.{current_val}" if prefix else str(current_val)

        if style == "cn_lower_chapter":
            return f"第{int_to_chinese(current_val)}章"
        if style == "cn_lower_section":
            return f"第{int_to_chinese(current_val)}节"

        if style == "cn_lower":
            return int_to_chinese(current_val)
        if style == "cn_upper":
            return int_to_chinese(current_val).translate(_CN_UPPER_MAP)
        if style == "roman_upper":
            return _to_roman(current_val)
        if style == "circled":
            return _CIRCLED_DIGITS.get(current_val, str(current_val))
        if style == "circled_paren":
            return _CIRCLED_PAREN_DIGITS.get(current_val, f"({current_val})")

        return str(current_val)
