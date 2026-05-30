"""FormatX 模型供应商静态预设"""

API_PROVIDER_PRESETS = [
    {
        "id": "deepseek",
        "label": "DeepSeek (官方)",
        "default_url": "https://api.deepseek.com",
        "api_type": "openai-completions",
        "default_models": ["deepseek-v4-pro", "deepseek-v4-flash"]
    },
    {
        "id": "dashscope",
        "label": "阿里云 (DashScope)",
        "default_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "api_type": "openai-completions",
        "default_models": ["qwen-vl-plus"]
    },
    {
        "id": "siliconflow",
        "label": "硅基流动 (SiliconFlow)",
        "default_url": "https://api.siliconflow.cn/v1",
        "api_type": "openai-completions",
        "default_models": ["deepseek-ai/DeepSeek-V3"]
    },
    {
        "id": "github",
        "label": "GitHub Models (免费 GPT-4o)",
        "default_url": "https://models.inference.ai.azure.com",
        "api_type": "openai-completions",
        "default_models": ["gpt-4o"]
    },
    {
        "id": "zhipu",
        "label": "智谱清言 (GLM)",
        "default_url": "https://open.bigmodel.cn/api/paas/v4",
        "api_type": "openai-completions",
        "default_models": ["glm-4v"]
    },
    {
        "id": "custom",
        "label": "自定义服务商 (高级)",
        "default_url": "",
        "api_type": "openai-completions",
        "default_models": []
    }
]
