"""
Word XML 编号注册引擎
负责在文档中创建列表编号定义和多级标题编号模板
"""
import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _ensure_list_numbering(doc):
    """确保文档有一个简单十进制列表编号定义，返回 numId。"""
    cache_attr = '_formatx_list_num_id'
    if hasattr(doc, cache_attr):
        return getattr(doc, cache_attr)

    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        return None

    num_root = numbering_part._element
    max_abs, max_num = 0, 0
    for el in num_root.findall(qn('w:abstractNum')):
        v = int(el.get(qn('w:abstractNumId'), '0'))
        if v > max_abs: max_abs = v
    for el in num_root.findall(qn('w:num')):
        v = int(el.get(qn('w:numId'), '0'))
        if v > max_num: max_num = v

    abs_id, num_id = max_abs + 1, max_num + 1
    abs_el = OxmlElement('w:abstractNum')
    abs_el.set(qn('w:abstractNumId'), str(abs_id))
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'multilevel')
    abs_el.append(mlt)

    for ilvl, fmt in enumerate(('%1.', '%1.%2.', '%1.%2.%3.')):
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))
        s = OxmlElement('w:start'); s.set(qn('w:val'), '1'); lvl.append(s)
        nf = OxmlElement('w:numFmt'); nf.set(qn('w:val'), 'decimal'); lvl.append(nf)
        lt = OxmlElement('w:lvlText'); lt.set(qn('w:val'), fmt); lvl.append(lt)
        lj = OxmlElement('w:lvlJc'); lj.set(qn('w:val'), 'left'); lvl.append(lj)
        abs_el.append(lvl)

    num_root.insert(0, abs_el)

    num_el = OxmlElement('w:num')
    num_el.set(qn('w:numId'), str(num_id))
    ref = OxmlElement('w:abstractNumId'); ref.set(qn('w:val'), str(abs_id))
    num_el.append(ref)
    num_root.insert(0, num_el)

    setattr(doc, cache_attr, num_id)
    return num_id


def _apply_list_numpr(para, num_id, ilvl=0):
    """给段落加上 w:numPr，使其成为自动编号列表项。"""
    pPr = para._element.get_or_add_pPr()
    old = pPr.find(qn('w:numPr'))
    if old is not None: pPr.remove(old)
    numPr = OxmlElement('w:numPr')
    iel = OxmlElement('w:ilvl'); iel.set(qn('w:val'), str(ilvl)); numPr.append(iel)
    nid = OxmlElement('w:numId'); nid.set(qn('w:val'), str(num_id)); numPr.append(nid)
    pPr.append(numPr)


def _compile_heading_template(template, ilvl):
    """编译标题编号模板：'第{current}章' → '第%1章'"""
    def _repl(m):
        name = m.group(1)
        if name in ('current', 'n', 'cn'):
            return f'%{ilvl + 1}'
        m2 = re.match(r'level(\d+)', name)
        if m2:
            return f'%{m2.group(1)}'
        return m.group(0)
    return re.sub(r'\{([a-zA-Z][a-zA-Z0-9_]*)\}', _repl, template)


def _register_heading_numbering(doc, templates=None):
    """注册标题多级编号定义（abstractNum + num），返回 numId。"""
    if templates is None:
        templates = {0: '第{current}章', 1: '{level1}.{current}', 2: '{level1}.{level2}.{current}'}

    try:
        numbering_part = doc.part.numbering_part
    except Exception:
        return None

    num_root = numbering_part._element
    max_abs, max_num = 0, 0
    for el in num_root.findall(qn('w:abstractNum')):
        v = int(el.get(qn('w:abstractNumId'), '0'))
        if v > max_abs: max_abs = v
    for el in num_root.findall(qn('w:num')):
        v = int(el.get(qn('w:numId'), '0'))
        if v > max_num: max_num = v

    abs_id, num_id = max_abs + 1, max_num + 1
    abs_el = OxmlElement('w:abstractNum')
    abs_el.set(qn('w:abstractNumId'), str(abs_id))
    mlt = OxmlElement('w:multiLevelType')
    mlt.set(qn('w:val'), 'multilevel')
    abs_el.append(mlt)

    for ilvl in range(max(templates.keys()) + 1):
        tmpl = templates.get(ilvl, '{current}')
        lvl_text = _compile_heading_template(tmpl, ilvl)
        lvl = OxmlElement('w:lvl')
        lvl.set(qn('w:ilvl'), str(ilvl))
        s = OxmlElement('w:start'); s.set(qn('w:val'), '1'); lvl.append(s)
        nf = OxmlElement('w:numFmt')
        nf.set(qn('w:val'), 'chineseCounting' if ilvl == 0 and '第' in tmpl else 'decimal')
        lvl.append(nf)
        lt = OxmlElement('w:lvlText'); lt.set(qn('w:val'), lvl_text); lvl.append(lt)
        lj = OxmlElement('w:lvlJc'); lj.set(qn('w:val'), 'left'); lvl.append(lj)
        abs_el.append(lvl)

    num_root.insert(0, abs_el)

    num_el = OxmlElement('w:num')
    num_el.set(qn('w:numId'), str(num_id))
    ref = OxmlElement('w:abstractNumId'); ref.set(qn('w:val'), str(abs_id))
    num_el.append(ref)
    num_root.insert(0, num_el)

    # 【已禁用】链接 Heading 1/2/3 样式到编号（所见即所得）
    setattr(doc, '_formatx_heading_num_id', num_id)
    return num_id


def _get_heading_num_id(doc):
    """获取或创建标题编号 numId（惰性，只创建一次）。"""
    if hasattr(doc, '_formatx_heading_num_id'):
        return getattr(doc, '_formatx_heading_num_id')
    return _register_heading_numbering(doc)
