"""文档结构防护网与角色精密嗅探中台"""
from core.ooxml_paragraph import replace_paragraph_payload_with_omml
from core.toc_entry import looks_like_reference_entry_line, looks_like_toc_entry_line


class SafePayloadRewriter:
    @staticmethod
    def inject_math_omml(paragraph, omml_element) -> bool:
        if paragraph is None or omml_element is None:
            return False
        result = replace_paragraph_payload_with_omml(
            paragraph._element, omml_element)
        return result.applied


class DocumentRoleInspector:
    @staticmethod
    def evaluate_line_role(text: str) -> str | None:
        raw = str(text or "").strip()
        if not raw:
            return None
        if looks_like_reference_entry_line(raw):
            return "reference_item"
        if looks_like_toc_entry_line(raw):
            return "toc_item"
        return None
