"""中文数字转换与全角字符处理工具"""

_CN_DIGITS = "零一二三四五六七八九"
_CN_UNITS = ["", "十", "百", "千"]


def int_to_chinese(n: int) -> str:
    """将阿拉伯数字转换为中文数字（支持 0-9999）"""
    if n == 0:
        return "零"
    if n < 0 or n > 9999:
        return str(n)

    result = ""
    s = str(n)
    length = len(s)

    for i, ch in enumerate(s):
        digit = int(ch)
        unit_index = length - 1 - i

        if digit == 0:
            if result and not result.endswith("零"):
                result += "零"
        else:
            if unit_index == 1 and digit == 1 and i == 0:
                result += "十"
            else:
                result += _CN_DIGITS[digit] + _CN_UNITS[unit_index]

    return result.rstrip("零")


def chinese_to_int(text: str) -> int | None:
    """将中文数字转换为阿拉伯数字（支持常见格式）"""
    text = text.strip()
    if not text:
        return None

    direct_map = {
        "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
        "六": 6, "七": 7, "八": 8, "九": 9, "十": 10,
        "十一": 11, "十二": 12, "十三": 13, "十四": 14, "十五": 15,
        "十六": 16, "十七": 17, "十八": 18, "十九": 19, "二十": 20,
    }
    if text in direct_map:
        return direct_map[text]

    result = 0
    current = 0
    for ch in text:
        if ch in _CN_DIGITS:
            current = _CN_DIGITS.index(ch)
        elif ch == "十":
            result += (current if current else 1) * 10
            current = 0
        elif ch == "百":
            result += current * 100
            current = 0
        elif ch == "千":
            result += current * 1000
            current = 0

    result += current
    return result if result > 0 else None
