# -*- coding: utf-8 -*-
"""宏：**表头格式统一**。

用户 2026-09-21 定下的规矩（前三条是从他文档里量出来的，已用 12 个真实表头验证）：

1. **段首两格 = 首行缩进 2 字符**（``w:ind w:firstLineChars="200"``），不是空格；
2. 文本 = ``表X-Y`` + N 个**半角空格** + 表格名字，其中
   ``N`` 取"让**名字的中点**落在『编号之后 → 版心右边界』这段的中点"（规则 B：
   与他现有 12 个表头的空格数平均差 +0.5 格、10/12 在 ±3 以内；规则 A「对齐版心中点」
   平均差 −5 格，不是他的做法）；
3. 编号写法**统一成 ``表X-Y``**（去掉 `表` 与编号之间的空格）；
4. 表格**一起居中**（原来是左对齐）。

**幂等**：每次都把编号与名字之间那段空格整段重算（不是往后追加），所以跑第二遍不会越加越多。
"""

import re

from ..document import FIRST_LINE_CHARS, TWIPS_PER_POINT, run_size_of
from ..ooxml import qn
from ..text import Paragraph

#: 表头编号：`表4.2-1` / `表 2.3-1` / `续表6-1` 都认
NUMBER_RE = re.compile(u"^\u7eed?\\s*\u8868\\s*(\\d+(?:\\.\\d+)*?-\\d+)")
#: 「表」字（含"续表"）与编号之间允许的空隙
PREFIX_SPACES = re.compile(u"^\u7eed?\u8868\\s*")

DEFAULT_OPTIONS = {
    "normalize_number": True,     # 编号统一成 表X-Y（去内部空格）
    "center_table": True,         # 表格一起居中
    "min_spaces": 1,              # 编号与名字之间至少留一个空格
    "only_before_table": True,    # 只处理"紧跟着表格"的那一段（表题的定义）
}


class HeaderPlan(object):
    """一个表头的改动计划。"""

    def __init__(self, paragraph, number, name, spaces_have, spaces_want, text_have,
                 text_want, table=None, reason=None):
        self.paragraph = paragraph
        self.number = number
        self.name = name
        self.spaces_have = spaces_have
        self.spaces_want = spaces_want
        self.text_have = text_have
        self.text_want = text_want
        self.table = table
        self.reason = reason

    @property
    def changed(self):
        return self.text_have != self.text_want or bool(self.reason)

    def to_dict(self):
        return {"number": self.number, "name": self.name,
                "spaces_before": self.spaces_have, "spaces_after": self.spaces_want,
                "changed": self.changed, "text_before": self.text_have,
                "text_after": self.text_want, "note": self.reason}


def _number_of(text):
    match = NUMBER_RE.match(text)
    if not match:
        return None
    prefix = PREFIX_SPACES.match(text).group(0)
    keep = u"\u7eed" if prefix.startswith(u"\u7eed") else u""
    return u"%s\u8868%s" % (keep, match.group(1))


def _split(text):
    """``(编号文本, 名字)``；不是表头返回 ``(None, None)``。"""
    number = _number_of(text)
    if number is None:
        return None, None
    rest = text[PREFIX_SPACES.match(text).end():]
    rest = rest[len(re.match(u"\\d+(?:\\.\\d+)*?-\\d+", rest).group(0)):]
    return number, rest.strip()


def _current_spaces(text):
    """编号与名字之间现有的半角空格数（用来报告"原来几个"）。"""
    match = re.match(u"^\u7eed?\u8868\\s*\\d+(?:\\.\\d+)*?-\\d+(\\s*)\\S", text)
    return match.group(1).count(u" ") if match else 0


def plan(document, options=None):
    """算出要改哪些表头（**只读**）。"""
    opts = dict(DEFAULT_OPTIONS)
    opts.update(options or {})
    geometry = document.geometry()
    children = document.block_children()
    plans = []
    for index, element in enumerate(children):
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element)
        raw = paragraph.text
        if not raw.strip():
            continue
        number, name = _split(raw.strip())
        if number is None:
            continue
        table = None
        for follow in children[index + 1:index + 2]:
            if follow.tag == qn("w:tbl"):
                table = follow
        if opts["only_before_table"] and table is None:
            continue

        size = run_size_of(paragraph)
        # 首行缩进 2 字符 = 2 × 字号点数 × 20 dxa/磅, 而字号点数 = sz/2
        #   → 2 × (sz/2) × 20 = sz × 20 dxa（sz=24 即小四 12pt → 480 dxa）
        first_line = 2 * (size / 2.0) * TWIPS_PER_POINT
        stripped = raw.strip()
        prefix_match = PREFIX_SPACES.match(stripped)
        prefix = prefix_match.group(0)
        digits = re.match(u"\\d+(?:\\.\\d+)*?-\\d+",
                          stripped[prefix_match.end():]).group(0)
        number_text = number if opts["normalize_number"] else (prefix + digits)
        number_w = geometry.width_of(number_text, size)
        name_w = geometry.width_of(name, size)
        # 规则 B：名字的中点在「编号之后 → 版心右边界」这段的中点
        target = (first_line + number_w + geometry.column_right) / 2.0
        want = int(round((target - first_line - number_w - name_w / 2.0)
                         / geometry.space_width(size)))
        want = max(want, opts["min_spaces"])
        text_want = number_text + u" " * want + name
        plans.append(HeaderPlan(paragraph, number_text, name, _current_spaces(stripped),
                                want, stripped, text_want, table,
                                None if table is not None else u"（后面不是表格）"))
    return plans


def _set_indent(paragraph, size):
    """段首两格 = 首行缩进 2 字符（``firstLineChars=200``），并给出对应的 dxa 值。"""
    from xml.etree import ElementTree as ET
    pr = paragraph.element.find(qn("w:pPr"))
    if pr is None:
        pr = ET.Element(qn("w:pPr"))
        paragraph.element.insert(0, pr)
    ind = pr.find(qn("w:ind"))
    if ind is None:
        ind = ET.SubElement(pr, qn("w:ind"))
    ind.set(qn("w:firstLineChars"), str(FIRST_LINE_CHARS))
    ind.set(qn("w:firstLine"), str(int(2 * (size / 2.0) * TWIPS_PER_POINT)))
    ind.set(qn("w:leftChars"), u"0")
    jc = pr.find(qn("w:jc"))
    if jc is None:
        jc = ET.SubElement(pr, qn("w:jc"))
    jc.set(qn("w:val"), u"left")


def _center_table(table):
    """表格居中：``w:tblPr/w:jc = center``。"""
    from xml.etree import ElementTree as ET
    pr = table.find(qn("w:tblPr"))
    if pr is None:
        pr = ET.Element(qn("w:tblPr"))
        table.insert(0, pr)
    jc = pr.find(qn("w:jc"))
    if jc is None:
        jc = ET.SubElement(pr, qn("w:jc"))
    if jc.get(qn("w:val")) == "center":
        return False
    jc.set(qn("w:val"), u"center")
    return True


def apply(document, options=None, dry_run=False):
    """把表头格式统一到整个文档；返回报告（``dry_run`` 时一个字节都不改）。"""
    opts = dict(DEFAULT_OPTIONS)
    opts.update(options or {})
    plans = plan(document, opts)
    details = []
    changed = 0
    tables_centered = 0
    if not dry_run:
        for item in plans:
            if item.changed:
                item.paragraph.replace(item.text_have, item.text_want, count=1)
                _set_indent(item.paragraph, run_size_of(item.paragraph))
                changed += 1
            if opts["center_table"] and item.table is not None:
                if _center_table(item.table):
                    tables_centered += 1
            details.append(item.to_dict())
        if changed or tables_centered:
            document.mark_dirty()
    else:
        details = [item.to_dict() for item in plans]
        changed = len([item for item in plans if item.changed])
        # dry-run 也要如实说"会居中几张表"，否则预览数字是假的
        tables_centered = 0
        if opts["center_table"]:
            for item in plans:
                if item.table is None:
                    continue
                jc = item.table.find(qn("w:tblPr") + "/" + qn("w:jc"))
                if jc is None or jc.get(qn("w:val")) != "center":
                    tables_centered += 1
    notes = []
    if opts["only_before_table"]:
        notes.append(u"只处理「紧跟着表格」的表题段（默认）；用 only_before_table=False 可放开")
    return {"op": "header", "changed": changed, "planned": len(plans),
            "tables_centered": tables_centered, "details": details, "notes": notes}
