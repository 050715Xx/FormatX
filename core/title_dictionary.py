"""标题字典——根据预设配置动态生成标题检测正则，替代静态 _HEADING_PATTERNS_READ"""
import re
from core.scene.schema import SceneConfig


class TitleDictionary:
    """从 SceneConfig 提取标题检测规则，向下游提供动态正则表达式"""

    def __init__(self, scene_config: SceneConfig = None):
        scene_config = scene_config or SceneConfig()

        # 基础兜底——中文 / 通用论文惯例标题
        self._default_front_matter = [
            "摘要", "Abstract", "ABSTRACT", "目录", "参考文献", "致谢", "附录",
        ]
        self._default_custom_titles = [
            "个人简历", "勘误页", "勘误", "声明", "申明", "授权书",
        ]

        # 融合 preset 特有的无编号标题
        preset_front = getattr(scene_config, "front_matter_titles", []) or []
        self.front_matter_titles = list(dict.fromkeys(
            self._default_front_matter + preset_front
        ))
        self._custom_titles = list(dict.fromkeys(self._default_custom_titles))

        # 预设的多层级有编号规则
        self.numbered_patterns = (
            getattr(scene_config, "numbered_heading_patterns", []) or []
        )

    def generate_front_matter_pipe(self) -> str:
        """生成无编号标题的管道正则片段，例如 摘要|Abstract|目录|参考文献"""
        all_titles = set(self.front_matter_titles + self._custom_titles)
        escaped = [re.escape(t) for t in all_titles if t.strip()]
        return "|".join(escaped)
