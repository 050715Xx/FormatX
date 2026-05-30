"""FormatX 运行时特性中台"""
import os
import sys


def _env_flag(name: str) -> bool:
    value = os.environ.get(name, "")
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


class RuntimeProbe:
    @staticmethod
    def is_frozen_bundle() -> bool:
        return hasattr(sys, '_MEIPASS')

    @staticmethod
    def is_macos_bundle() -> bool:
        return _env_flag("DOCX_MACOS_BUNDLE")

    @staticmethod
    def is_page_probing_disabled() -> bool:
        return _env_flag("DOCX_DISABLE_WORD_PAGE_SCOPE")

    @staticmethod
    def is_omml_fallback_disabled() -> bool:
        return _env_flag("DOCX_DISABLE_MATHTYPE_OFFICE_FALLBACK")
