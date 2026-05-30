"""
DOCX 输入分析器
段落角色预判：代码识别、图片检测、标题模式匹配、表格格式化
"""
import re
from docx.shared import Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# 常见等宽字体（用于识别代码块）
_MONOSPACE_FONTS = {
    'consolas', 'courier', 'courier new', 'source code pro',
    'fira code', 'jetbrains mono', 'monaco', 'menlo', 'dejavu sans mono',
    'lucida console', 'inconsolata', 'cascadia code', 'ubuntu mono',
}

# 中文序号
_CN_NUM_PLAIN = r'(?:[一二三四五六七八九十]{1,3})'



def _looks_like_code(text):
    """启发式判断文本是否像代码行。"""
    stripped = text.strip()
    if not stripped:
        return False
    if re.search(r'[一-鿿]', stripped):
        return False
    code_starters = (
        'def ', 'class ', 'if ', 'for ', 'while ', 'import ', 'from ',
        'var ', 'let ', 'const ', 'function ', 'return ', 'print(',
        'public ', 'private ', 'protected ', 'static ', 'void ',
        'int ', 'string ', 'bool ', 'float ', 'double ',
        'console.', 'System.', 'using ', 'package ', '#include',
    )
    if any(stripped.startswith(kw) for kw in code_starters):
        return True
    if re.search(r'[{}();\[\]]', stripped):
        return True
    return False


def _para_has_image(para):
    """递归检查段落中是否包含图片、图形、图表等。"""
    for child in para._element.iter():
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('drawing', 'pict', 'object', 'inline', 'anchor'):
            return True
    return False


def _format_table(tbl_element):
    """对深拷贝后的表格 XML 元素统一应用排版规范。"""
    tblPr = tbl_element.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl_element.insert(0, tblPr)

    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.insert(0, tblW)
    tblW.set(qn('w:w'), '0')
    tblW.set(qn('w:type'), 'auto')

    tblpPr = tblPr.find(qn('w:tblpPr'))
    if tblpPr is not None:
        tblPr.remove(tblpPr)

    jc = tblPr.find(qn('w:jc'))
    if jc is None:
        jc = OxmlElement('w:jc')
        tblPr.append(jc)
    jc.set(qn('w:val'), 'center')

    margin_twips = str(int(Cm(0.19).emu / 635))

    for tc in tbl_element.iter(qn('w:tc')):
        tcPr = tc.find(qn('w:tcPr'))
        if tcPr is None:
            tcPr = OxmlElement('w:tcPr')
            tc.insert(0, tcPr)

        tcW = tcPr.find(qn('w:tcW'))
        if tcW is not None:
            tcPr.remove(tcW)

        vAlign = tcPr.find(qn('w:vAlign'))
        if vAlign is None:
            vAlign = OxmlElement('w:vAlign')
            tcPr.append(vAlign)
        vAlign.set(qn('w:val'), 'center')

        tcMar = tcPr.find(qn('w:tcMar'))
        if tcMar is not None:
            tcPr.remove(tcMar)
        tcMar = OxmlElement('w:tcMar')
        for side in ('left', 'right'):
            mar = OxmlElement(f'w:{side}')
            mar.set(qn('w:w'), margin_twips)
            mar.set(qn('w:type'), 'dxa')
            tcMar.append(mar)
        tcPr.append(tcMar)

        for p in tc.iter(qn('w:p')):
            pPr = p.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                p.insert(0, pPr)
            jc = pPr.find(qn('w:jc'))
            if jc is None:
                jc = OxmlElement('w:jc')
                pPr.append(jc)
            jc.set(qn('w:val'), 'center')

            for r in p.iter(qn('w:r')):
                rPr = r.find(qn('w:rPr'))
                if rPr is None:
                    rPr = OxmlElement('w:rPr')
                    r.insert(0, rPr)
                for tag_name in ('w:sz', 'w:szCs'):
                    sz = rPr.find(qn(tag_name))
                    if sz is None:
                        sz = OxmlElement(tag_name)
                        rPr.append(sz)
                    sz.set(qn('w:val'), '21')


# ── Word 自动编号解析 ────────────────────────────────────────────

def _build_numbering_maps(doc):
    num_maps = {"abstract": {}, "instance": {}}
    try:
        numbering_part = doc.part.numbering_part
        if numbering_part is None:
            return num_maps
        root = numbering_part.element
    except (AttributeError, KeyError, NotImplementedError):
        return num_maps

    for abs_num in root.findall(qn('w:abstractNum')):
        abs_str = abs_num.get(qn('w:abstractNumId'))
        if abs_str is None:
            continue
        abs_id = int(abs_str)
        levels = {}
        for lvl in abs_num.findall(qn('w:lvl')):
            ilvl_str = lvl.get(qn('w:ilvl'))
            if ilvl_str is None:
                continue
            ilvl = int(ilvl_str)
            numFmt_el = lvl.find(qn('w:numFmt'))
            lvlText_el = lvl.find(qn('w:lvlText'))
            levels[ilvl] = {
                "numFmt": numFmt_el.get(qn('w:val')) if numFmt_el is not None else "decimal",
                "lvlText": lvlText_el.get(qn('w:val')) if lvlText_el is not None else "",
            }
        num_maps["abstract"][abs_id] = levels

    for num in root.findall(qn('w:num')):
        num_id_str = num.get(qn('w:numId'))
        if num_id_str is None:
            continue
        num_id = int(num_id_str)
        abs_ref = num.find(qn('w:abstractNumId'))
        if abs_ref is None:
            continue
        abs_id = int(abs_ref.get(qn('w:val')))
        overrides = {}
        for ovr in num.findall(qn('w:lvlOverride')):
            ilvl_str = ovr.get(qn('w:ilvl'))
            if ilvl_str is None:
                continue
            ilvl_ov = int(ilvl_str)
            lvl_el = ovr.find(qn('w:lvl'))
            if lvl_el is not None:
                nf = lvl_el.find(qn('w:numFmt'))
                lt = lvl_el.find(qn('w:lvlText'))
                overrides[ilvl_ov] = {
                    "numFmt": nf.get(qn('w:val')) if nf is not None else None,
                    "lvlText": lt.get(qn('w:val')) if lt is not None else None,
                }
        num_maps["instance"][num_id] = {"abstractNumId": abs_id, "overrides": overrides}

    return num_maps


def _container_numpr(para):
    try:
        pPr = para._element.find(qn('w:pPr'))
        if pPr is None:
            return None
        numPr = pPr.find(qn('w:numPr'))
        if numPr is None:
            return None
        numId_el = numPr.find(qn('w:numId'))
        ilvl_el = numPr.find(qn('w:ilvl'))
        numId = int(numId_el.get(qn('w:val'))) if numId_el is not None else None
        ilvl = int(ilvl_el.get(qn('w:val'))) if ilvl_el is not None else None
        if numId is not None and ilvl is not None:
            return (numId, ilvl)
    except Exception:
        pass
    return None


def _find_numbering_lvl(num_maps, num_id, ilvl):
    if not num_maps or num_id not in num_maps["instance"]:
        return None
    inst = num_maps["instance"][num_id]
    if ilvl in inst["overrides"]:
        ov = inst["overrides"][ilvl]
        if ov.get("numFmt") and ov.get("lvlText"):
            return ov
    abs_id = inst["abstractNumId"]
    if abs_id in num_maps["abstract"] and ilvl in num_maps["abstract"][abs_id]:
        base = dict(num_maps["abstract"][abs_id][ilvl])
        if ilvl in inst["overrides"]:
            ov = inst["overrides"][ilvl]
            if ov.get("numFmt"):
                base["numFmt"] = ov["numFmt"]
            if ov.get("lvlText"):
                base["lvlText"] = ov["lvlText"]
        return base
    return None
