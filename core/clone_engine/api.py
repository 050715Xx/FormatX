"""智能克隆引擎主入口"""
import os
from src.docx_io.style_clone import clone_scene_style_from_docx
from src.scene.schema import SceneConfig as TfSceneConfig
from .translator import translate_tf_to_formatx
from core.scene.schema import SceneConfig


def create_scene_from_clone(docx_path: str) -> SceneConfig:
    if not os.path.exists(docx_path):
        return SceneConfig(name="智能克隆模板")

    tf_config = TfSceneConfig(name="智能克隆模板")
    clone_scene_style_from_docx(tf_config, docx_path)
    return translate_tf_to_formatx(tf_config)
