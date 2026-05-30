"""
MathType OLE 预处理清洗器
扫描文档中的 MathType OLE 二进制对象，提取 LaTeX 字符串，
原地替换为 $...$ 文本，供后续 OMML 引擎接管。
"""
import io
from lxml import etree
from docx.oxml.ns import qn, nsmap

# 确保 o 命名空间已注册（部分 docx 文件需显式声明）
nsmap.setdefault('o', 'urn:schemas-microsoft-com:office:office')

try:
    import olefile
    _HAS_OLEFILE = True
except ImportError:
    _HAS_OLEFILE = False


def wash_mathtype_oles(doc) -> int:
    """扫描整个文档，将 MathType OLE 对象洗白为 $LaTeX$ 文本。

    Returns: 清洗的公式数量
    """
    if not _HAS_OLEFILE:
        return 0

    from formula_core.mathtype_ole import decode_equation_native_to_latex

    cleaned_count = 0

    for para in doc.paragraphs:
        for run in para.runs:
            objects = run._element.findall('.//' + qn('w:object'))
            for obj in objects:
                ole_node = obj.find('.//' + qn('o:OLEObject'))
                if ole_node is None:
                    continue

                prog_id = ole_node.get('ProgID', '')
                if 'Equation' not in prog_id:
                    continue

                r_id = ole_node.get(qn('r:id'))
                if not r_id:
                    continue

                try:
                    part = doc.part.related_parts[r_id]
                    ole_blob = part.blob

                    if not olefile.isOleFile(io.BytesIO(ole_blob)):
                        continue

                    with olefile.OleFileIO(io.BytesIO(ole_blob)) as ole:
                        if not ole.exists('Equation Native'):
                            continue
                        native_stream = ole.openstream('Equation Native').read()

                        latex, _ = decode_equation_native_to_latex(native_stream)

                        # 删除原 OLE 对象
                        run._element.remove(obj)

                        # 在当前 w:r 节点下直接追加 w:t 文本节点
                        t = etree.SubElement(run._element, qn('w:t'))
                        t.text = f"[OLE_EQUATION]${latex}$"
                        t.set(qn('xml:space'), 'preserve')
                        cleaned_count += 1

                except Exception:
                    pass

    return cleaned_count
