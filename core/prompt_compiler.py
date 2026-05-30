"""AI 提示词排版约束编译器"""
from core.config_manager import load_full_config


class PromptCompiler:
    @staticmethod
    def generate_formatting_system_prompt() -> str:
        config = load_full_config()
        mode = config.get("typography_mode", "preset")
        preset_name = config.get("format_preset", "默认通用格式")
        heading_styles = config.get("heading_styles")

        lines = [
            "\n【系统级排版与结构铁律】",
            "你的输出将被直接送入高精排版内核进行物理级 Word 渲染。你必须严格遵守以下规范：",
            "- 绝对禁止手动缩进：正文必须顶格写，禁止使用任何空格或制表符来模拟缩进。",
            "- 禁止使用 HTML 换行：必须使用标准 Markdown 双回车分段。",
            "- 结构组织：使用 Markdown 标题（# ## ###）组织层级结构。",
        ]

        if heading_styles:
            lines.append(
                "- 绝对禁止手动为标题编号：排版引擎已接管多级自动编号。"
                "使用纯 Markdown 标题（如 ## 研究背景），绝不允许自带数字序号。"
            )

        if mode == "custom":
            lines.append(
                f"- 当前文章载体：[{preset_name}] 的内容结构文风 + "
                f"[自定义高精物理排版微调]。请确保逻辑分段紧凑（每段 150-300 字），"
                f"以配合高密度的自定义版面。"
            )
        else:
            lines.append(
                f"- 当前文章载体：严格匹配 [{preset_name}] 的标准文体规范。"
            )

        if "学术毕业论文" in preset_name:
            lines.append(
                "- 学术结构规范：语气必须极度严谨，杜绝口语。"
                "最高支持到 4 级标题。"
                "若有引用，在文末生成独立的 ## 参考文献 层级。"
            )
        elif preset_name == "IEEE Conference":
            lines.append(
                "- IEEE 论文规范：采用紧凑的英文或标准学术分栏语境。"
                "多用专业术语，逻辑直接，避免冗长描述。"
            )

        return "\n".join(lines)
