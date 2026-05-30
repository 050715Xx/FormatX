"""FormatX 多协议路由通讯引擎"""
import json
import requests


def _clean_base_url(url: str) -> str:
    """清理 URL 后缀，只保留基地址"""
    url = url.rstrip("/")
    suffixes = ["/chat/completions", "/v1/messages", "/models", "/responses"]
    for s in suffixes:
        if url.endswith(s):
            url = url[:-len(s)]
            break
    return url


def _assemble_request(base_url: str, api_key: str, api_type: str, action: str):
    """
    单一真相源：负责构造所有 API 的 Endpoint 和 Headers
    action: "chat" | "models" | "ping"
    """
    base = _clean_base_url(base_url)

    if api_type == "anthropic-messages":
        endpoints = {"chat": "/v1/messages", "models": "/v1/models", "ping": "/v1/messages"}
        headers = {
            "x-api-key": api_key,
            "anthropic-version": "2023-06-01",
            "Content-Type": "application/json"
        }
    else:
        endpoints = {"chat": "/chat/completions", "models": "/models", "ping": "/chat/completions"}
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }

    return base + endpoints[action], headers


# ── 1. 连接测试 ──
def test_provider_connection(base_url: str, api_key: str, model: str,
                             api_type: str = "openai-completions") -> tuple:
    endpoint, headers = _assemble_request(base_url, api_key, api_type, "ping")

    payload = {
        "model": model,
        "messages": [{"role": "user", "content": "ping"}],
        "max_tokens": 5,
        "stream": False
    }

    try:
        res = requests.post(endpoint, headers=headers, json=payload, timeout=10)
        if res.status_code == 200:
            return True, "连接成功，信道畅通"
        try:
            err_msg = res.json().get("error", {}).get("message", res.text)
        except Exception:
            err_msg = res.text[:200]
        return False, f"HTTP {res.status_code}: {err_msg}"
    except Exception as e:
        return False, str(e)


# ── 2. 远程模型拉取 ──
def fetch_remote_models(base_url: str, api_key: str,
                        api_type: str = "openai-completions") -> list:
    endpoint, headers = _assemble_request(base_url, api_key, api_type, "models")
    try:
        res = requests.get(endpoint, headers=headers, timeout=12)
        res.raise_for_status()
        data = res.json()
        models = [item["id"] for item in data.get("data", []) if "id" in item]
        return sorted(list(set(models)))
    except Exception as e:
        raise RuntimeError(f"模型字典拉取失败: {str(e)}")


# ── 3. OCR 图像提取 ──
def extract_text_from_image(image_base64: str, api_key: str, model: str,
                            base_url: str, api_type: str = "openai-completions") -> str:
    endpoint, headers = _assemble_request(base_url, api_key, api_type, "chat")

    if api_type == "anthropic-messages":
        payload = {
            "model": model,
            "max_tokens": 4096,
            "messages": [{"role": "user", "content": [
                {"type": "image", "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": image_base64
                }},
                {"type": "text", "text": (
                    "你是专业的 OCR 识别与学术排版助手。"
                    "请提取图片中的所有文字内容，包括：\n"
                    "1. 数学公式转为标准 LaTeX 语法（行内 $...$，独立 $$...$$）\n"
                    "2. 保留原文结构（标题、段落、列表、表格）\n"
                    "3. 修复明显错别字和标点错误\n"
                    "4. 直接输出最终文本，不要加任何问候、解释或代码块标记"
                )}
            ]}]
        }
    else:
        payload = {
            "model": model,
            "messages": [{"role": "user", "content": [
                {"type": "text", "text": (
                    "你是专业的 OCR 识别与学术排版助手。"
                    "请提取图片中的所有文字内容，包括：\n"
                    "1. 数学公式转为标准 LaTeX 语法（行内 $...$，独立 $$...$$）\n"
                    "2. 保留原文结构（标题、段落、列表、表格）\n"
                    "3. 修复明显错别字和标点错误\n"
                    "4. 直接输出最终文本，不要加任何问候、解释或代码块标记"
                )},
                {"type": "image_url",
                 "image_url": {"url": f"data:image/png;base64,{image_base64}"}}
            ]}],
            "stream": False
        }

    try:
        res = requests.post(endpoint, headers=headers, json=payload, timeout=90)
        if res.status_code != 200:
            try:
                err_msg = res.json().get("error", {}).get("message", res.text)
            except Exception:
                err_msg = res.text[:300]
            raise RuntimeError(f"API 拒绝请求 (状态码 {res.status_code}): {err_msg}")

        data = res.json()
        if api_type == "anthropic-messages":
            content_blocks = data.get("content", [])
            return "\n".join(
                c["text"] for c in content_blocks if c.get("type") == "text"
            ).strip()
        else:
            return data["choices"][0]["message"]["content"].strip()

    except requests.exceptions.Timeout:
        raise RuntimeError("请求超时：图片可能较大，或者网络不佳，请稍后重试。")


# ── 4. 流式生成 ──
def stream_chat_completion(api_key: str, base_url: str, model: str, messages: list,
                           api_type: str = "openai-completions"):
    endpoint, headers = _assemble_request(base_url, api_key, api_type, "chat")

    if api_type == "anthropic-messages":
        system_msg = "\n".join(
            m["content"] for m in messages if m["role"] == "system"
        )
        user_msgs = [m for m in messages if m["role"] != "system"]
        payload = {
            "model": model,
            "messages": user_msgs,
            "stream": True,
            "max_tokens": 4096
        }
        if system_msg:
            payload["system"] = system_msg
    else:
        payload = {
            "model": model,
            "messages": messages,
            "stream": True,
            "temperature": 0.7
        }
        if "reasoner" in model or "r1" in model.lower():
            payload["temperature"] = 1.0

    try:
        response = requests.post(
            endpoint, headers=headers, json=payload, stream=True, timeout=120
        )
        response.raise_for_status()

        for line in response.iter_lines():
            if line:
                decoded = line.decode("utf-8").strip()
                if decoded.startswith("data: "):
                    data_str = decoded[6:]
                    if data_str == "[DONE]":
                        break
                    try:
                        data_json = json.loads(data_str)
                        if api_type == "anthropic-messages":
                            delta = data_json.get("delta", {})
                            if delta.get("type") == "text_delta":
                                text = delta.get("text", "")
                                if text:
                                    yield ("", text)
                        else:
                            delta = data_json["choices"][0].get("delta", {})
                            reasoning = delta.get("reasoning_content", "")
                            content = delta.get("content", "")
                            if reasoning or content:
                                yield (reasoning, content)
                    except (json.JSONDecodeError, KeyError, IndexError):
                        continue
    except Exception as e:
        yield ("", f"\n[信道突发中断: {str(e)}]")
