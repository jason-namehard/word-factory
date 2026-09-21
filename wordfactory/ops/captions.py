# -*- coding: utf-8 -*-
"""宏：**题注格式统一**（表题 + 图题，两套规矩）。

规矩全部来自用户（2026-09-21），并且**表的那一套是用他文档里 12 个真实表题量出来/验证过的**：

**表题**（`表X-Y`）
1. 段首 = 首行缩进 **2 字符**（``w:ind w:firstLineChars="200"``），不是空格
   （他文档里 10 个是 2 字符、`表2.0-1`/`表6-1` 是 3 字符 —— 一并归到 2）；
2. 文本 = `表X-Y` + N 个**半角空格** + 名字，N 用**规则 B**：让**名字的中点**落在
   「编号之后 → 版心右边界」这段的中点（与他现有值平均差 +0.67 格、10/12 在 ±3 内；
   另一条候选规则 A「对齐版心中点」平均差 −5 格，不是他的做法）；
3. 编号统一成 `表X-Y`（去掉「表」与编号之间的空格）；
4. 表格**一起居中**。

**图题**（`图X-Y`）—— 规矩不一样（用户原话：「整个居中，从图这个字开始就居中，然后中间空一个格子」）
1. **整段居中**（``w:jc=center``），不靠空格凑；
2. 编号与名字之间**正好一个空格**；
3. 编号同样统一成 `图X-Y`（去掉内部空格）；
4. 段首缩进要**清掉**（悬挂/首行缩进会把居中顶偏）。

**幂等**：两套都是"整段重算"而不是追加，所以跑第二遍不会越加越多。
"""

import re

from ..document import FIRST_LINE_CHARS, TWIPS_PER_POINT, run_size_of
from ..ooxml import qn
from ..text import Paragraph

#: 题注：`表4.2-1` / `表 2.3-1` / `续表6-1` / `图4.2-1` / `图 6-2`
CAPTION_RE = re.compile(u"^(\u7eed?\u8868|\u56fe)\\s*(\\d+(?:\\.\\d+)*?-\\d+)")
DENOMINATOR_RE = re.compile(u"\\d+(?:\\.\\d+)*?-\\d+")

DEFAULT_OPTIONS = {
    "normalize_number": True,     # 编号统一成 表X-Y / 图X-Y（去内部空格）
    "center_table": True,         # 表题：表格一起居中
    "min_spaces": 1,              # 表题：编号与名字之间至少留几个空格
    "only_before_table": True,    # 表题：只处理"紧跟着表格"的那一段
    "figure_space": 1,            # 图题：编号与名字之间正好几个空格
    "center_figure": True,        # 图题：整段居中
    "clear_figure_indent": True,  # 图题：清掉段首缩进（否则居中会被顶偏）
}

FIGURE = u"\u56fe"
TABLE = u"\u8868"


class CaptionPlan(object):
    """一个题注段的改动计划。"""

    def __init__(self, paragraph, kind, number, name, spaces_have, spaces_want,
                 text_have, text_want, table=None, note=None):
        self.paragraph = paragraph
        self.kind = kind
        self.number = number
        self.name = name
        self.spaces_have = spaces_have
        self.spaces_want = spaces_want
        self.text_have = text_have
        self.text_want = text_want
        self.table = table
        self.note = note

    @property
    def changed(self):
        return self.text_have != self.text_want or bool(self.note)

    #: 验证版要标蓝的范围 = 工具动过的那一段（编号 + 中间的空格）
    @property
    def mark_end(self):
        return len(self.number) + self.spaces_want

    def to_dict(self):
        return {"kind": self.kind, "number": self.number, "name": self.name,
                "spaces_before": self.spaces_have, "spaces_after": self.spaces_want,
                "changed": self.changed, "text_before": self.text_have,
                "text_after": self.text_want, "note": self.note}


def _split(text):
    """``(类型, 编号文本, 名字)``；不是题注返回 ``(None, None, None)``。"""
    match = CAPTION_RE.match(text)
    if not match:
        return None, None, None
    kind = FIGURE if match.group(1).endswith(FIGURE) else TABLE
    keep = u"\u7eed" if match.group(1).startswith(u"\u7eed") else u""
    number = u"%s%s%s" % (keep, TABLE if kind == TABLE else FIGURE, match.group(2))
    rest = text[match.end():]
    return kind, number, rest.strip()


def _spaces_between(text):
    match = re.match(u"^(?:\u7eed?\u8868|\u56fe)\\s*\\d+(?:\\.\\d+)*?-\\d+(\\s*)\\S", text)
    return match.group(1).count(u" ") if match else 0


def plan(document, options=None):
    """算出要改哪些题注（**只读**）。"""
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
        kind, number, name = _split(raw.strip())
        if kind is None:
            continue
        table = None
        if children[index + 1:index + 2] and children[index + 1].tag == qn("w:tbl"):
            table = children[index + 1]
        if kind == TABLE and opts["only_before_table"] and table is None:
            continue

        size = run_size_of(paragraph)
        first_line = 2 * (size / 2.0) * TWIPS_PER_POINT      # 2 字符（dz 见 Geometry 注释）
        if kind == FIGURE:
            want = opts["figure_space"]
            text_want = number + u" " * want + name
            plans.append(CaptionPlan(paragraph, kind, number, name,
                                     _spaces_between(raw.strip()), want,
                                     raw.strip(), text_want, None,
                                     _figure_layout_note(paragraph, opts)))
            continue

        number_w = geometry.width_of(number, size)
        name_w = geometry.width_of(name, size)
        target = (first_line + number_w + geometry.column_right) / 2.0
        want = int(round((target - first_line - number_w - name_w / 2.0)
                         / geometry.space_width(size)))
        want = max(want, opts["min_spaces"])
        text_want = number + u" " * want + name
        notes = [text for text in (_table_layout_note(paragraph),
                                   None if table is not None
                                   else u"\uff08\u540e\u9762\u4e0d\u662f\u8868\u683c\uff09") if text]
        plans.append(CaptionPlan(paragraph, kind, number, name,
                                 _spaces_between(raw.strip()), want,
                                 raw.strip(), text_want, table,
                                 u" + ".join(notes) if notes else None))
    return plans


def _needs_indent_clear(pr):
    """这一段有没有需要清掉的段首/悬挂缩进（**只读**）。"""
    ind = pr.find(qn("w:ind")) if pr is not None else None
    if ind is None:
        return False
    for key in ("w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars", "w:left"):
        value = ind.get(qn(key))
        if value is not None and value != u"0":
            return True
    return False


def _figure_layout_note(paragraph, opts):
    """图题还需要动布局吗？返回要说的话，或 ``None``（**只读**）。

    这条判断必须真的查，不能看到 ``center_figure`` 就写死一句"整段居中"——
    实测写死的后果是：图题每次都报"要改"（``changed`` 看的就是 note），
    于是第二遍跑仍然报"5 个要改"，数字是假的，还会把没变的东西再标一次蓝。
    """
    pr = paragraph.element.find(qn("w:pPr"))
    wants = []
    if opts["center_figure"]:
        jc = pr.find(qn("w:jc")) if pr is not None else None
        if jc is None or jc.get(qn("w:val")) != u"center":
            wants.append(u"\u6574\u6bb5\u5c45\u4e2d")
    if opts["clear_figure_indent"] and _needs_indent_clear(pr):
        wants.append(u"\u6e05\u6389\u6bb5\u9996\u7f29\u8fdb")
    return u" + ".join(wants) if wants else None


def _table_layout_note(paragraph):
    """表题还需要动布局吗？（段首缩进 2 字符 + 左对齐）只说真正缺的。"""
    pr = paragraph.element.find(qn("w:pPr"))
    ind = pr.find(qn("w:ind")) if pr is not None else None
    jc = pr.find(qn("w:jc")) if pr is not None else None
    wants = []
    if ind is None or ind.get(qn("w:firstLineChars")) != str(FIRST_LINE_CHARS) \
            or ind.get(qn("w:firstLine")) is None:
        wants.append(u"\u6bb5\u9996\u7f29\u8fdb 2 \u5b57\u7b26")
    if jc is None or jc.get(qn("w:val")) != u"left":
        wants.append(u"\u5de6\u5bf9\u9f50")
    return u" + ".join(wants) if wants else None


def _paragraph_properties(paragraph):
    from xml.etree import ElementTree as ET
    pr = paragraph.element.find(qn("w:pPr"))
    if pr is None:
        pr = ET.Element(qn("w:pPr"))
        paragraph.element.insert(0, pr)
    return pr


def _set_table_indent(paragraph, size):
    """表题：首行缩进 2 字符 + 左对齐。"""
    pr = _paragraph_properties(paragraph)
    ind = pr.find(qn("w:ind"))
    if ind is None:
        from xml.etree import ElementTree as ET
        ind = ET.SubElement(pr, qn("w:ind"))
    ind.set(qn("w:firstLineChars"), str(FIRST_LINE_CHARS))
    ind.set(qn("w:firstLine"), str(int(2 * (size / 2.0) * TWIPS_PER_POINT)))
    ind.set(qn("w:leftChars"), u"0")
    for drop in ("w:hanging", "w:hangingChars", "w:left"):
        if ind.get(qn(drop)) is not None:
            del ind.attrib[qn(drop)]
    _set_jc(pr, u"left")


def _set_figure_layout(paragraph, center=True, clear_indent=True):
    """图题：整段居中 + 清掉段首/悬挂缩进。"""
    pr = _paragraph_properties(paragraph)
    if clear_indent:
        ind = pr.find(qn("w:ind"))
        if ind is not None:
            for drop in ("w:firstLine", "w:firstLineChars", "w:hanging", "w:hangingChars",
                         "w:left", "w:leftChars"):
                if ind.get(qn(drop)) is not None:
                    del ind.attrib[qn(drop)]
            if len(ind.attrib) == 0:
                pr.remove(ind)
    if center:
        _set_jc(pr, u"center")


def _set_jc(pr, value):
    from xml.etree import ElementTree as ET
    jc = pr.find(qn("w:jc"))
    if jc is None:
        jc = ET.SubElement(pr, qn("w:jc"))
    jc.set(qn("w:val"), value)


def _center_table(table):
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
    """把题注格式统一到整个文档；返回报告。

    ``details`` 里每一项都带 ``paragraph`` 与 ``mark_end``，验证版据此把"动过的部分"标蓝。
    """
    opts = dict(DEFAULT_OPTIONS)
    opts.update(options or {})
    plans = plan(document, opts)
    changed = 0
    tables_centered = 0
    details = []
    if not dry_run:
        for item in plans:
            if item.changed:
                if item.text_have != item.text_want:      # 文字没变就别碰它（免得白切一遍 run）
                    item.paragraph.replace(item.text_have, item.text_want, count=1)
                if item.kind == FIGURE:
                    _set_figure_layout(item.paragraph, opts["center_figure"],
                                       opts["clear_figure_indent"])
                else:
                    _set_table_indent(item.paragraph, run_size_of(item.paragraph))
            if item.kind == TABLE and opts["center_table"] and item.table is not None:
                if _center_table(item.table):
                    tables_centered += 1
    changed = len([item for item in plans if item.changed])
    tables = len([item for item in plans if item.kind == TABLE])
    figures = len([item for item in plans if item.kind == FIGURE])
    changed_tables = len([item for item in plans if item.kind == TABLE and item.changed])
    changed_figures = len([item for item in plans if item.kind == FIGURE and item.changed])
    if dry_run and opts["center_table"]:
        for item in plans:
            if item.kind != TABLE or item.table is None:
                continue
            jc = item.table.find(qn("w:tblPr") + "/" + qn("w:jc"))
            if jc is None or jc.get(qn("w:val")) != "center":
                tables_centered += 1
    for item in plans:
        row = item.to_dict()
        row["paragraph"] = item.paragraph
        row["mark_end"] = item.mark_end
        details.append(row)
    return {"op": "captions", "changed": changed, "planned": len(plans),
            "tables": tables, "figures": figures,
            "changed_tables": changed_tables, "changed_figures": changed_figures,
            "tables_centered": tables_centered,
            "details": details, "options": opts}
