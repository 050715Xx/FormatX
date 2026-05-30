"""DeepSeek API 排版修复客户端"""
import requests
import json
import os
from pathlib import Path

_CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".formatx_ai_config.json")


def load_config() -> dict:
    if os.path.exists(_CONFIG_PATH):
        try:
            with open(_CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(base_url: str, api_key: str, model: str):
    config_data = {"base_url": base_url, "api_key": api_key, "model": model}
    path = Path.home() / ".formatx_ai_config.json"
    data = {}
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:
            data = {}
    data.update(config_data)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def format_text_with_deepseek(raw_text: str,
                               api_key: str = "",
                               model: str = "deepseek-chat",
                               base_url: str = "https://api.deepseek.com/chat/completions") -> str:
    """利用 DeepSeek 将粗糙 OCR 文本修复为完美排版的 Markdown/LaTeX 格式。"""
    if not api_key:
        cfg = load_config()
        api_key = cfg.get("api_key", "")
        base_url = cfg.get("base_url", base_url)
        model = cfg.get("model", model)

    if not api_key:
        return "⚠️ 未配置 API Key，请在设置中输入。"

    system_prompt = (
        "你是一个顶级的学术论文排版专家和 LaTeX 公式修复大师。"
        "用户会发给你一段通过 OCR 扫描出来的学术段落，里面可能包含错别字、标点错误、以及识别稀烂的数学公式。"
        "请你的任务是：\n"
        "1. 修复所有的错别字和语病，保持原意不变。\n"
        "2. 将所有的数学符号、公式（无论是行内还是独立行）全部转换为极其标准的 LaTeX 语法。\n"
        "3. 行内公式严格使用 $...$ 包裹，独立公式使用 $$...$$ 包裹。\n"
        "4. 直接输出修复后的最终文本，绝对不要输出任何问候语、解释或者 Markdown 代码块标记。"
    )

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": f"请帮我修复并排版以下内容：\n\n{raw_text}"}
        ],
        "stream": False,
        "temperature": 0.1
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    try:
        response = requests.post(base_url, headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        result = response.json()
        return result['choices'][0]['message']['content'].strip()
    except requests.exceptions.Timeout:
        return "⚠️ 请求超时：API 响应时间过长，请检查网络或稍后再试。"
    except Exception as e:
        return f"⚠️ 接口调用失败：{str(e)}"


def extract_text_from_image(image_base64: str,
                           api_key: str = "",
                           model: str = "deepseek-chat",
                           base_url: str = "https://api.deepseek.com/chat/completions") -> str:
    """利用 DeepSeek Vision 多模态模型从图片中提取并排版文字。

    image_base64: base64 编码的图片数据（不含 data:image/... 前缀）
    """
    if not api_key:
        cfg = load_config()
        api_key = cfg.get("api_key", "")
        base_url = cfg.get("base_url", base_url)
        model = cfg.get("model", model)

    if not api_key:
        return "⚠️ 未配置 API Key，请在设置中输入。"

    system_prompt = (
        "你是专业的 OCR 识别与学术排版助手。"
        "请提取图片中的所有文字内容，包括：\n"
        "1. 数学公式转为标准 LaTeX 语法（行内 $...$，独立 $$...$$）\n"
        "2. 保留原文结构（标题、段落、列表、表格）\n"
        "3. 修复明显错别字和标点错误\n"
        "4. 直接输出最终文本，不要加任何问候、解释或代码块标记"
    )

    payload = {
        "model": model,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": system_prompt},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/png;base64,{image_base64}"}}
            ]
        }],
        "stream": False,
        "temperature": 0.1
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    try:
        response = requests.post(base_url, headers=headers, json=payload, timeout=90)

        if response.status_code != 200:
            err_msg = response.text
            try:
                err_json = response.json()
                if "error" in err_json:
                    err_msg = err_json["error"].get("message", err_msg)
            except Exception:
                pass
            return f"❌ API 拒绝请求 (状态码 {response.status_code}): {err_msg}"

        result = response.json()
        return result['choices'][0]['message']['content'].strip()

    except requests.exceptions.Timeout:
        return "⚠️ 请求超时：图片可能较大，或者网络不佳，请稍后重试。"
    except Exception as e:
        return f"⚠️ 接口底层异常: {str(e)}"


def test_connection(base_url: str, api_key: str, model: str) -> str:
    """测试 API 连接是否有效"""
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": [{"type": "text", "text": "请回复：连接成功"}]}],
        "stream": False,
        "temperature": 0.1
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }
    try:
        response = requests.post(base_url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        result = response.json()
        content = result['choices'][0]['message']['content'].strip()
        return f"✅ 连接成功！模型回复：{content}"
    except Exception as e:
        return f"❌ 连接失败：{str(e)}"


def generate_document_via_deepseek(user_prompt: str,
                                   api_key: str = "",
                                   model: str = "deepseek-chat",
                                   base_url: str = "https://api.deepseek.com/chat/completions",
                                   style_mode: str = "academic") -> str:
    """根据用户提示词生成学术文档（Markdown + LaTeX 格式）。"""
    from core.prompts import SYSTEM_PROMPT_ACADEMIC, SYSTEM_PROMPT_NOVICE

    if not api_key:
        cfg = load_config()
        api_key = cfg.get("api_key", "")
        base_url = cfg.get("base_url", base_url)
        model = cfg.get("model", model)

    if not api_key:
        raise ValueError("未配置 API Key，请在设置中输入。")

    system_prompt = SYSTEM_PROMPT_ACADEMIC if style_mode == "academic" else SYSTEM_PROMPT_NOVICE

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "stream": False,
        "temperature": 0.7,
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }

    response = requests.post(base_url, headers=headers, json=payload, timeout=120)
    response.raise_for_status()
    result = response.json()
    return result['choices'][0]['message']['content'].strip()


def stream_chat_completion(api_key: str, base_url: str, model: str, messages: list):
    """流式 API 生成器——纯原生 requests 实现 SSE，零第三方依赖"""
    import json

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }
    payload = {
        "model": model,
        "messages": messages,
        "stream": True,
        "temperature": 0.7,
    }

    try:
        response = requests.post(base_url, headers=headers, json=payload,
                                 stream=True, timeout=120)
        response.raise_for_status()

        for line in response.iter_lines():
            if line:
                decoded = line.decode("utf-8")
                if decoded.startswith("data: "):
                    data_str = decoded[6:]
                    if data_str == "[DONE]":
                        break
                    try:
                        data_json = json.loads(data_str)
                        delta = data_json["choices"][0].get("delta", {})
                        if "content" in delta:
                            yield delta["content"]
                    except json.JSONDecodeError:
                        continue
    except Exception as e:
        yield f"\n[API 通讯中断: {str(e)}]"
