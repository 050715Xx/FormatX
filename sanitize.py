"""
文档免疫清洗模块：在 python-docx 解析前，直接对 .docx 压缩包进行
二进制级别的手术，清除会触发 lxml 解析器崩溃的脏数据。
"""
import zipfile
import tempfile
import os
import re
import atexit
from pathlib import Path
from lxml import etree

_MAX_ATTR_VALUE_LEN = 200_000
_tmp_files = []


def _cleanup():
    for p in _tmp_files:
        try:
            os.unlink(p)
        except OSError:
            pass


atexit.register(_cleanup)


def sanitize_docx(src_path: str) -> str:
    """
    文档免疫清洗：在读取前进行二进制手术。
    返回安全的临时文件路径，供 python-docx 调用。
    """
    src = Path(src_path)
    fd, tmp_path = tempfile.mkstemp(suffix=".docx", prefix="sanitized_")
    os.close(fd)
    _tmp_files.append(tmp_path)

    try:
        with zipfile.ZipFile(src, 'r') as zin, \
             zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename.endswith(".rels"):
                    data = _fix_rels(data)

                if item.filename.endswith(".xml"):
                    data = _fix_oversized_xml_attrs(data)

                zout.writestr(item, data)
    except zipfile.BadZipFile:
        return src_path

    return tmp_path


def _fix_rels(data: bytes) -> bytes:
    """清除所有 Target="NULL" 的畸形链接"""
    try:
        root = etree.fromstring(data)
        ns = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
        removed = False
        for rel in root.xpath("//rel:Relationship", namespaces=ns):
            target = rel.get("Target", "")
            if target.upper() == "NULL":
                rel.getparent().remove(rel)
                removed = True
        if removed:
            return etree.tostring(root, xml_declaration=True, encoding="UTF-8")
    except Exception:
        pass
    return data


def _fix_oversized_xml_attrs(data: bytes) -> bytes:
    """使用正则预处理，截断过长属性值，防止解析器崩溃"""

    def _truncate(match):
        attr_name = match.group(1)
        value = match.group(2)
        if len(value) > _MAX_ATTR_VALUE_LEN:
            return f'{attr_name}=""'.encode()
        return match.group(0)

    return re.sub(rb'(\s[\w:.-]+)\s*=\s*"([^"]*)"', _truncate, data)
