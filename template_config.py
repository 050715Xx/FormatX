"""
三层模板注入架构控制器

Layer 1: 静态结构注入（封面/目录）
Layer 2: 动态样式克隆（正文 DNA）
Layer 3: 底层硬编码兜底（format_conversion.py 默认值）

用法:
    config = TemplateConfig(cover_path="cover.docx", template_path="template.docx")
    reformat_docx("input.docx", "output.docx", config=config)
"""
import os


class TemplateConfig:
    def __init__(self, cover_path=None, template_path=None):
        self.cover_path = cover_path          # Layer 1: 封面/目录路径
        self.template_path = template_path    # Layer 2: 样式模板路径

    @property
    def layer_mode(self):
        """核心判定：有模板走 Layer 2，无模板走 Layer 3"""
        if self.template_path and os.path.exists(self.template_path):
            return 2
        return 3
