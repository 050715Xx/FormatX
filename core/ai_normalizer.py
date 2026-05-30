"""
AI 内容清洗管线
处理 DeepSeek、豆包、ChatGPT 等 AI 输出的格式问题
"""
import re


def _clean_pasted_text(text):
    """清洗从网页或 AI 粘贴的底层不可见字符与 UI 干扰词"""
    text = text.replace('\x0c', '\\f').replace('\x0b', '\\v')
    text = re.sub(r'[\x00-\x08\x0e-\x1f\x7f]', '', text)
    for ch in ('​', '‌', '‍', '﻿', '\u200E', '\u200F',
               '\u202A', '\u202B', '\u202C', '\u202D', '\u202E'):
        text = text.replace(ch, "")
    text = text.replace('　', ' ')
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    text = re.sub(r'^\s*复制代码\s*$', '', text, flags=re.MULTILINE)
    text = re.sub(r'(?<!^)[ \t]{2,}', ' ', text, flags=re.MULTILINE)

    # 消除 >=4 个字符的重复短语（网页剪贴板重影）
    text = re.sub(r'([a-zA-Z]{4,})\1', r'\1', text)

    return text


def _normalize_ai_markdown(text):
    """将各家 AI 特殊格式统一翻译为引擎标准格式"""
    text = re.sub(r'<think>.*?</think>', '', text, flags=re.DOTALL)

    # 强力反转义补丁
    text = text.replace(r'\*\*', '**')
    text = text.replace(r'\*', '*')
    text = text.replace(r'\_', '_')
    text = text.replace(r'\---', '---')

    # 行首标记恢复
    text = re.sub(r'^(\s*)\\#', r'\1#', text, flags=re.MULTILINE)
    text = re.sub(r'^(\s*)\\-', r'\1-', text, flags=re.MULTILINE)
    text = re.sub(r'^(\s*\d+)\\\.', r'\1.', text, flags=re.MULTILINE)

    # 链接恢复（必须在公式处理之前）
    text = re.sub(r'\\\[(.*?)\\\]\(', r'[\1](', text)

    # 公式统一
    text = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    text = re.sub(r'\\\((.*?)\\\)', r'$\1$', text, flags=re.DOTALL)

    # 豆包修复
    text = re.sub(r'\\\$\\\$(.*?)\\\$\\\$', r'$$\1$$', text, flags=re.DOTALL)
    text = re.sub(r'\\\$([^$]+?)\\\$', r'$\1$', text)

    # 公式空格修复
    text = re.sub(r'\$\s+([^\$]+?)\s+\$', r'$\1$', text)

    # 圆圈方块统一为 -
    text = re.sub(r'^(\s*)[\*\+•●○■·]\s+', r'\1- ', text, flags=re.MULTILINE)

    def _strip_bold_in_math(match):
        return f'${match.group(1).replace("**", "")}$'
    text = re.sub(r'\$\*+(.*?)\*+\$', _strip_bold_in_math, text)
    text = re.sub(r'\*\*\$([^$]+?)\$\*\*', r'$\1$', text)
    return text


# ── 转义隐身斗篷 ────────────────────────────────────────────────

import uuid

_RE_FORMATX_ESCAPE = re.compile(r'\\([\\`*_{}\[\]()#+\-.!|>~])')


def hide_escapes(text):
    escapes_map = {}

    def replacer(match):
        char = match.group(1)
        placeholder = f"__ESC_{uuid.uuid4().hex[:8]}__"
        escapes_map[placeholder] = char
        return placeholder

    return _RE_FORMATX_ESCAPE.sub(replacer, text), escapes_map


def restore_escapes(text, escapes_map):
    if not escapes_map or not text:
        return text
    for placeholder, char in escapes_map.items():
        text = text.replace(placeholder, char)
    return text
