# -*- coding: utf-8 -*-
"""表格款式（外置规则文件 `rules/tablestyle.json`）+ 采集 + 预览。

用户 2026-09-22 的要求：

> 「我支持再开一个外置规则文件 rules/tablestyle.json，另外表格的格式设置需要考虑，
>   保存已经调好的设置（一个调得满意的表格时间成本比较高，因此需要保存）和预览两个功能」

所以这里有四件事：

1. **款式的数据模型**（`:class:`TableStyle` / :class:`StyleSet`）：框线、单元格内边距、
   垂直居中、表头处理、列宽策略、表格对齐/布局、字号字体 —— 全是可以外置编辑的 JSON；
2. **套用**（:func:`apply`）：把款式写进 `w:tblPr` / `w:tcPr` / `w:trPr` / `w:pPr`；
3. **采集**（:func:`capture`）：反过来读一个**你已经调好的表格**，把它现在的样子存成一个命名款式
   （调一次、以后所有报告复用）；
4. **预览**（:func:`build_preview`）：造一份小文档，里面用**代表性内容**把候选款式各渲染一张表，
   你用 Word 打开挑选 —— 不靠渲染器算的东西我们不假装能算。

**单位**（OOXML 里的坑，写清楚免得改错）：框线 `w:sz` 是 **1/8 磅**；单元格内边距 `w:w` 是 **dxa（1/20 磅）**；
字号 `w:sz` 是 **半磅**；列宽 `w:w` 也是 dxa。

**幂等**：所有 setter 都"值一样就不动、也不计数"，所以同一个款式跑第二遍报 0 处（有单测钉）。
"""

import collections
import io
import json
import os
from xml.etree import ElementTree as ET

from .ooxml import NAMESPACES, qn

#: `w:tblPr` 的子元素顺序（OOXML 是**序列**，插错位置 Word 会嫌文件坏）
TBLPR_ORDER = ("w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual",
               "w:tblStyleRowBandSize", "w:tblStyleColBandSize", "w:tblW", "w:jc",
               "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
               "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
#: `w:tcPr` 的子元素顺序
TCPR_ORDER = ("w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders",
              "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
              "w:vAlign", "w:hideMark")
#: `w:trPr` 的子元素顺序
TRPR_ORDER = ("w:cnfStyle", "w:divId", "w:gridBefore", "w:gridAfter", "w:wBefore",
              "w:wAfter", "w:cantSplit", "w:trHeight", "w:tblHeader", "w:tblCellSpacing",
              "w:jc", "w:hidden")
#: `w:pPr` 的子元素顺序（只列我们会插的，插到比它靠后的元素之前即可）
PPR_ORDER = ("w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr",
             "w:widowControl", "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd",
             "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap",
             "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN",
             "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
             "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
             "w:textDirection", "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl",
             "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr", "w:pPrChange")
#: `w:rPr` 的子元素顺序
RPR_ORDER = ("w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
             "w:smallCaps", "w:strike", "w:dstrike", "w:outline", "w:shadow", "w:emboss",
             "w:imprint", "w:noProof", "w:snapToGrid", "w:vanish", "w:webHidden",
             "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz", "w:szCs",
             "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText",
             "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
             "w:specVanish", "w:oMath")

#: 款式里允许出现的键（校验用，写错键名会被 `check()` 点出来）
STYLE_KEYS = ("note", "table_align", "table_layout", "width_percent", "borders",
              "cell_margins", "v_align", "header", "body", "column_widths",
              "font_size_half_points", "font_east_asia")
BORDER_EDGES = ("top", "bottom", "left", "right", "insideH", "insideV", "header_bottom")
BORDER_VALUES = ("single", "double", "thick", "dotted", "dashed", "none", "nil")

#: 出厂款式：都是"工程报告里常见的样子"，用户可改可加
DEFAULT_STYLES = {
    "schema": 1,
    "name": u"表格款式（外置规则；可改可加）",
    "styles": {
        u"三线表": {
            "note": u"报告里最常用的学术款：上下粗线 + 表头下细线，无竖线",
            "table_align": "center",
            "table_layout": "fixed",
            "borders": {"top": {"val": "single", "sz": 12, "color": "000000"},
                        "bottom": {"val": "single", "sz": 12, "color": "000000"},
                        "left": {"val": "none", "sz": 0, "color": "auto"},
                        "right": {"val": "none", "sz": 0, "color": "auto"},
                        "insideH": {"val": "none", "sz": 0, "color": "auto"},
                        "insideV": {"val": "none", "sz": 0, "color": "auto"},
                        "header_bottom": {"val": "single", "sz": 6, "color": "000000"}},
            "cell_margins": {"top": 40, "left": 80, "bottom": 40, "right": 80},
            "v_align": "center",
            "header": {"bold": True, "align": "center", "repeat": True},
            "body": {"align": None},
            "column_widths": "keep",
        },
        u"全框线": {
            "note": u"每格都有细框线（数据表常用），表头加粗居中并跨页重复",
            "table_align": "center",
            "table_layout": "fixed",
            "borders": {"top": {"val": "single", "sz": 6, "color": "000000"},
                        "bottom": {"val": "single", "sz": 6, "color": "000000"},
                        "left": {"val": "single", "sz": 6, "color": "000000"},
                        "right": {"val": "single", "sz": 6, "color": "000000"},
                        "insideH": {"val": "single", "sz": 6, "color": "000000"},
                        "insideV": {"val": "single", "sz": 6, "color": "000000"}},
            "cell_margins": {"top": 40, "left": 80, "bottom": 40, "right": 80},
            "v_align": "center",
            "header": {"bold": True, "align": "center", "repeat": True},
            "body": {"align": None},
            "column_widths": "keep",
        },
        u"三线表_等宽": {
            "note": u"三线表 + 各列等宽（列宽不均匀时用它救一下）",
            "table_align": "center",
            "table_layout": "fixed",
            "borders": {"top": {"val": "single", "sz": 12, "color": "000000"},
                        "bottom": {"val": "single", "sz": 12, "color": "000000"},
                        "left": {"val": "none", "sz": 0, "color": "auto"},
                        "right": {"val": "none", "sz": 0, "color": "auto"},
                        "insideH": {"val": "none", "sz": 0, "color": "auto"},
                        "insideV": {"val": "none", "sz": 0, "color": "auto"},
                        "header_bottom": {"val": "single", "sz": 6, "color": "000000"}},
            "cell_margins": {"top": 40, "left": 80, "bottom": 40, "right": 80},
            "v_align": "center",
            "header": {"bold": True, "align": "center", "repeat": True},
            "body": {"align": None},
            "column_widths": "equal",
        },
    },
}


class TableStyleError(Exception):
    """款式文件的问题（人话）。"""


class TableStyle(object):
    """一个命名款式。字段全是可选，缺的就不动。"""

    def __init__(self, name, data=None):
        self.name = name
        data = dict(data or {})
        self.raw = data
        self.note = data.get("note") or u""
        self.table_align = data.get("table_align")
        self.table_layout = data.get("table_layout")
        self.width_percent = data.get("width_percent")
        self.borders = dict(data.get("borders") or {})
        self.cell_margins = dict(data.get("cell_margins") or {})
        self.v_align = data.get("v_align")
        self.header = dict(data.get("header") or {})
        self.body = dict(data.get("body") or {})
        self.column_widths = data.get("column_widths")
        self.font_size_half_points = data.get("font_size_half_points")
        self.font_east_asia = data.get("font_east_asia")

    def check(self):
        problems = []
        for key in self.raw:
            if key not in STYLE_KEYS:
                problems.append(u"款式 %s 里有不认识的键 %r（允许：%s）"
                                % (self.name, key, u"、".join(STYLE_KEYS)))
        for edge in self.borders:
            if edge not in BORDER_EDGES:
                problems.append(u"款式 %s 的 borders 里有不认识的边 %r（允许：%s）"
                                % (self.name, edge, u"、".join(BORDER_EDGES)))
                continue
            value = self.borders[edge] or {}
            val = value.get("val")
            if val and val not in BORDER_VALUES:
                problems.append(u"款式 %s 的 %s.val=%r 不是合法线型（%s）"
                                % (self.name, edge, val, u"、".join(BORDER_VALUES)))
            for key, item in value.items():
                if key not in ("val", "sz", "color", "space"):
                    problems.append(u"款式 %s 的 %s 里有不认识的键 %r" % (self.name, edge, key))
                if key == "sz" and not isinstance(item, int):
                    problems.append(u"款式 %s 的 %s.sz 必须是整数（1/8 磅）" % (self.name, edge))
        if self.table_align not in (None, "left", "center", "right"):
            problems.append(u"款式 %s 的 table_align=%r 只能是 left/center/right" % (self.name, self.table_align))
        if self.table_layout not in (None, "fixed", "autofit"):
            problems.append(u"款式 %s 的 table_layout=%r 只能是 fixed/autofit" % (self.name, self.table_layout))
        if self.column_widths not in (None, "keep", "equal", "content"):
            problems.append(u"款式 %s 的 column_widths=%r 只能是 keep/equal/content"
                            % (self.name, self.column_widths))
        if self.v_align not in (None, "top", "center", "bottom"):
            problems.append(u"款式 %s 的 v_align=%r 只能是 top/center/bottom" % (self.name, self.v_align))
        for key in ("bold", "align", "repeat"):
            if key in self.header and self.header[key] not in (None, True, False, "left", "center", "right"):
                problems.append(u"款式 %s 的 header.%s=%r 不合法" % (self.name, key, self.header[key]))
        if self.header.get("align") not in (None, "left", "center", "right"):
            problems.append(u"款式 %s 的 header.align=%r 只能是 left/center/right" % (self.name, self.header.get("align")))
        if self.body.get("align") not in (None, "left", "center", "right"):
            problems.append(u"款式 %s 的 body.align=%r 只能是 left/center/right" % (self.name, self.body.get("align")))
        for key in self.cell_margins:
            if key not in ("top", "left", "bottom", "right"):
                problems.append(u"款式 %s 的 cell_margins 里有不认识的键 %r" % (self.name, key))
        return problems


class StyleSet(object):
    """款式文件的全部内容。"""

    def __init__(self, data=None, path=None):
        data = dict(data or {})
        self.name = data.get("name") or u"表格款式"
        self.note = data.get("note") or u""
        self.schema = data.get("schema") or 1
        self.styles = collections.OrderedDict()
        for key, value in (data.get("styles") or {}).items():
            self.styles[key] = TableStyle(key, value)
        self.path = path
        self.raw = data

    @classmethod
    def load(cls, path):
        if not path or not os.path.exists(path):
            return cls(DEFAULT_STYLES)
        with io.open(path, "r", encoding="utf-8-sig") as handle:
            return cls(json.load(handle, object_pairs_hook=collections.OrderedDict), path=path)

    def save(self, path):
        parent = os.path.dirname(os.path.abspath(path))
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        data = collections.OrderedDict()
        data["schema"] = self.schema
        data["name"] = self.name
        if self.note:
            data["note"] = self.note
        data["styles"] = collections.OrderedDict(
            (key, style.raw) for key, style in self.styles.items())
        with io.open(path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(data, ensure_ascii=False, indent=2) + u"\n")
        return path

    def get(self, name):
        if name not in self.styles:
            raise TableStyleError(u"款式 %r 不在 %s 里（有的：%s）"
                                  % (name, self.path or u"内置默认",
                                     u"、".join(self.styles) or u"（空）"))
        return self.styles[name]

    def check(self):
        problems = []
        if not self.styles:
            problems.append(u"款式文件里一个款式都没有")
        for style in self.styles.values():
            problems.extend(style.check())
        return problems

    def put(self, name, data):
        self.styles[name] = TableStyle(name, data)
        self.raw["styles"] = collections.OrderedDict(
            (key, style.raw) for key, style in self.styles.items())
        return self.styles[name]


# --------------------------------------------------------------------- 套用
def parse_indexes(selector, total):
    """``"all"`` / ``"3"`` / ``"1,4-6"`` → 要处理的表格序号（1 起）。"""
    if not selector or selector == "all":
        return list(range(1, total + 1))
    picked = []
    for chunk in str(selector).split(u","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if u"-" in chunk:
            start, _, end = chunk.partition(u"-")
            try:
                picked.extend(range(int(start), int(end) + 1))
            except ValueError:
                raise TableStyleError(u"表格序号 %r 看不懂" % chunk)
        else:
            try:
                picked.append(int(chunk))
            except ValueError:
                raise TableStyleError(u"表格序号 %r 看不懂" % chunk)
    return [index for index in picked if 1 <= index <= total]


def tables_of(document):
    """正文里的表格（含嵌套；嵌套的也算一张，但套用时的列宽按各自的 grid 算）。"""
    return [element for element in document.body().iter(qn("w:tbl"))]


def apply(document, style, selector="all", dry_run=False):
    """把款式套到选中的表格上；返回报告（改了几张表、每类改动几处）。"""
    tables = tables_of(document)
    picked = parse_indexes(selector, len(tables))
    changes = collections.Counter()
    touched = []
    for index, table in enumerate(tables, start=1):
        if index not in picked:
            continue
        count = _apply_to_table(table, style, dry_run)
        if count:
            touched.append(index)
            changes.update(count)
    total = sum(changes.values())
    if not dry_run and total:
        document.mark_dirty()
    return {"op": "tablestyle", "style": style.name, "tables": len(tables),
            "selected": len(picked), "changed_tables": len(touched),
            "tables_changed": touched, "changes": dict(changes), "total": total,
            "dry_run": bool(dry_run), "style_raw": style.raw}


def _apply_to_table(table, style, dry_run):
    changes = collections.Counter()
    pr = _ensure(table, "w:tblPr", None, 0)                 # tblPr 必须是 tbl 的第一个孩子
    if style.table_align and _set_attribute(_ensure(pr, "w:jc", TBLPR_ORDER, dry_run=dry_run), qn("w:val"),
                                            style.table_align, dry_run):
        changes["表格对齐"] += 1
    if style.table_layout and _set_attribute(_ensure(pr, "w:tblLayout", TBLPR_ORDER, dry_run=dry_run),
                                             qn("w:type"), style.table_layout, dry_run):
        changes["表格布局"] += 1
    if style.width_percent:
        node = _ensure(pr, "w:tblW", TBLPR_ORDER, dry_run=dry_run)
        if _set_attribute(node, qn("w:type"), "pct", dry_run):
            changes["表格宽度"] += 1
        if _set_attribute(node, qn("w:w"), int(style.width_percent * 50), dry_run):
            changes["表格宽度"] += 1
    if style.borders and _apply_borders(pr, style, dry_run):
        changes["框线"] += 1
    if style.cell_margins and _apply_cell_margins(pr, style.cell_margins, dry_run):
        changes["单元格内边距"] += 1

    rows = [element for element in table if element.tag == qn("w:tr")]
    for row_index, row in enumerate(rows):
        is_header = row_index == 0
        for cell in [element for element in row if element.tag == qn("w:tc")]:
            if style.v_align:
                tcpr = _ensure(cell, "w:tcPr", None, 0, dry_run)     # tcPr 必须是 tc 的第一个孩子
                if _set_attribute(_ensure(tcpr, "w:vAlign", TCPR_ORDER, dry_run=dry_run), qn("w:val"),
                                  style.v_align, dry_run):
                    changes["垂直居中"] += 1
            if is_header and style.borders.get("header_bottom"):
                tcpr = _ensure(cell, "w:tcPr", None, 0, dry_run)
                borders = _ensure(tcpr, "w:tcBorders", TCPR_ORDER, dry_run=dry_run)
                if _set_border(borders, "w:bottom", style.borders["header_bottom"], dry_run):
                    changes["表头下框线"] += 1
        if is_header and style.header.get("repeat"):
            trpr = _ensure(row, "w:trPr", None, 0, dry_run)          # trPr 必须是 tr 的第一个孩子
            if _ensure_flag(trpr, "w:tblHeader", TRPR_ORDER, dry_run):
                changes["表头跨页重复"] += 1
        want_align = (style.header.get("align") if is_header else style.body.get("align"))
        want_bold = style.header.get("bold") if is_header else None
        for paragraph in row.iter(qn("w:p")):
            if want_align:
                ppr = _ensure(paragraph, "w:pPr", None, 0, dry_run)
                if _set_attribute(_ensure(ppr, "w:jc", PPR_ORDER), qn("w:val"),
                                  want_align, dry_run):
                    changes["段落对齐"] += 1
            if want_bold:
                for run in paragraph.iter(qn("w:r")):
                    rpr = _ensure(run, "w:rPr", None, 0, dry_run)
                    if _ensure_flag(rpr, "w:b", RPR_ORDER, dry_run):
                        changes["表头加粗"] += 1
        if is_header and style.body.get("align"):
            pass                                    # 表头不受 body.align 影响
    if style.column_widths in ("equal", "content"):
        if _apply_column_widths(table, style.column_widths, dry_run):
            changes["列宽"] += 1
    if style.font_size_half_points or style.font_east_asia:
        if _apply_table_font(table, style, dry_run):
            changes["字号字体"] += 1
    return changes


def _apply_borders(pr, style, dry_run):
    borders = _ensure(pr, "w:tblBorders", TBLPR_ORDER, dry_run=dry_run)
    changed = False
    for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
        if edge in style.borders:
            if _set_border(borders, "w:" + edge, style.borders[edge], dry_run):
                changed = True
    return changed


def _apply_cell_margins(pr, margins, dry_run):
    node = _ensure(pr, "w:tblCellMar", TBLPR_ORDER, dry_run=dry_run)
    changed = False
    for edge in ("top", "left", "bottom", "right"):
        if edge in margins:
            child = _ensure(node, "w:" + edge, ("w:top", "w:left", "w:bottom", "w:right"),
                            dry_run=dry_run)
            # **两个属性都要设、都要判**：写成 `a or b` 会被短路 —— 第一次只设了 w:w，
            # 第二次才发现 w:type 还缺着 → 报告"又改了一处"（幂等就破了）。踩过。
            if _set_attribute(child, qn("w:w"), int(margins[edge]), dry_run):
                changed = True
            if _set_attribute(child, qn("w:type"), "dxa", dry_run):
                changed = True
    return changed


def _apply_column_widths(table, mode, dry_run):
    """列宽：``equal`` = 各列等宽；``content`` = 按内容长度估（渲染器说了算的事我们只能估）。"""
    grid = table.find(qn("w:tblGrid"))
    columns = [element for element in grid] if grid is not None else []
    if not columns:
        return False
    total = sum(int(column.get(qn("w:w")) or 0) for column in columns) or 0
    if mode == "equal":
        each = int(total / len(columns)) if total else 0
        if not each:
            return False
        widths = [each] * len(columns)
    else:
        weights = _content_weights(table, len(columns))
        if not weights or not total:
            return False
        base = sum(weights)
        widths = [max(1, int(total * weight / base)) for weight in weights]
    changed = False
    for column, width in zip(columns, widths):
        if _set_attribute(column, qn("w:w"), width, dry_run):
            changed = True
    # 同步到单元格（Word 里以 `w:tcW` 为准）
    for row in [element for element in table if element.tag == qn("w:tr")]:
        cells = [element for element in row if element.tag == qn("w:tc")]
        if len(cells) != len(widths):
            continue
        for cell, width in zip(cells, widths):
            tcpr = _ensure(cell, "w:tcPr", None, 0, dry_run)
            tcw = _ensure(tcpr, "w:tcW", TCPR_ORDER, dry_run=dry_run)
            if _set_attribute(tcw, qn("w:type"), "dxa", dry_run):
                changed = True
            if _set_attribute(tcw, qn("w:w"), width, dry_run):
                changed = True
    return changed


def _content_weights(table, column_count):
    """按每列内容的显示宽度估权重（中文按 2 个单位）。"""
    weights = [0] * column_count
    for row in [element for element in table if element.tag == qn("w:tr")]:
        index = 0
        for cell in [element for element in row if element.tag == qn("w:tc")]:
            span = 1
            pr = cell.find(qn("w:tcPr"))
            if pr is not None:
                node = pr.find(qn("w:gridSpan"))
                if node is not None:
                    try:
                        span = max(1, int(node.get(qn("w:val")) or 1))
                    except (TypeError, ValueError):
                        span = 1
            if index >= column_count:
                break
            text = u"".join((node.text or "") for node in cell.iter(qn("w:t")))
            width = sum(2.0 if ord(ch) > 0x2E80 else 1.0 for ch in text) + 2.0
            weights[index] = max(weights[index], width)
            index += span
    return [weight or 1.0 for weight in weights]


def _apply_table_font(table, style, dry_run):
    changed = False
    for run in table.iter(qn("w:r")):
        rpr = _ensure(run, "w:rPr", None, 0, dry_run)                 # rPr 必须是 run 的第一个孩子
        if style.font_size_half_points:
            node = _ensure(rpr, "w:sz", RPR_ORDER, dry_run=dry_run)
            if _set_attribute(node, qn("w:val"), int(style.font_size_half_points), dry_run):
                changed = True
        if style.font_east_asia:
            node = _ensure(rpr, "w:rFonts", RPR_ORDER, dry_run=dry_run)
            if _set_attribute(node, qn("w:eastAsia"), style.font_east_asia, dry_run):
                changed = True
    return changed


# --------------------------------------------------------------------- 采集
def capture(document, table_index=1):
    """读一个**已经调好的表格**，把它现在的样子存成款式数据（用户的核心诉求：调一次、以后复用）。"""
    tables = tables_of(document)
    if not tables:
        raise TableStyleError(u"这份文档里没有表格")
    if not 1 <= table_index <= len(tables):
        raise TableStyleError(u"表格序号 %d 超出范围（共 %d 张）" % (table_index, len(tables)))
    table = tables[table_index - 1]
    data = collections.OrderedDict()
    pr = table.find(qn("w:tblPr"))
    if pr is not None:
        jc = pr.find(qn("w:jc"))
        if jc is not None and jc.get(qn("w:val")):
            data["table_align"] = jc.get(qn("w:val"))
        layout = pr.find(qn("w:tblLayout"))
        if layout is not None and layout.get(qn("w:type")):
            data["table_layout"] = layout.get(qn("w:type"))
        borders = pr.find(qn("w:tblBorders"))
        if borders is not None:
            captured = collections.OrderedDict()
            for edge in ("top", "bottom", "left", "right", "insideH", "insideV"):
                node = borders.find(qn("w:" + edge))
                if node is not None:
                    captured[edge] = _border_dict(node)
            if captured:
                data["borders"] = captured
        margins = pr.find(qn("w:tblCellMar"))
        if margins is not None:
            captured = collections.OrderedDict()
            for edge in ("top", "left", "bottom", "right"):
                node = margins.find(qn("w:" + edge))
                if node is not None:
                    captured[edge] = int(node.get(qn("w:w")) or 0)
            if captured:
                data["cell_margins"] = captured
    rows = [element for element in table if element.tag == qn("w:tr")]
    if rows:
        header_cell = next((element for element in rows[0] if element.tag == qn("w:tc")), None)
        if header_cell is not None:
            tcpr = header_cell.find(qn("w:tcPr"))
            if tcpr is not None:
                valign = tcpr.find(qn("w:vAlign"))
                if valign is not None and valign.get(qn("w:val")):
                    data["v_align"] = valign.get(qn("w:val"))
                tcborders = tcpr.find(qn("w:tcBorders"))
                if tcborders is not None:
                    bottom = tcborders.find(qn("w:bottom"))
                    if bottom is not None:
                        data.setdefault("borders", collections.OrderedDict())[u"header_bottom"] \
                            = _border_dict(bottom)
        trpr = rows[0].find(qn("w:trPr"))
        header = collections.OrderedDict()
        if trpr is not None and trpr.find(qn("w:tblHeader")) is not None:
            header["repeat"] = True
        first_paragraph = next(iter(rows[0].iter(qn("w:p"))), None)
        if first_paragraph is not None:
            ppr = first_paragraph.find(qn("w:pPr"))
            jc = ppr.find(qn("w:jc")) if ppr is not None else None
            if jc is not None and jc.get(qn("w:val")):
                header["align"] = jc.get(qn("w:val"))
        first_run = next(iter(rows[0].iter(qn("w:r"))), None)
        if first_run is not None:
            rpr = first_run.find(qn("w:rPr"))
            if rpr is not None and _flag_on(rpr.find(qn("w:b"))):
                header["bold"] = True
        if header:
            data["header"] = header
        grid = table.find(qn("w:tblGrid"))
        if grid is not None:
            widths = [int(column.get(qn("w:w")) or 0) for column in grid]
            if widths and len(set(widths)) == 1:
                data["column_widths"] = "equal"
    return data


def _border_dict(node):
    out = collections.OrderedDict()
    val = node.get(qn("w:val"))
    if val:
        out["val"] = val
    for key in ("sz", "space"):
        value = node.get(qn("w:" + key))
        if value is not None:
            try:
                out[key] = int(value)
            except ValueError:
                out[key] = value
    color = node.get(qn("w:color"))
    if color:
        out["color"] = color
    return out


def _flag_on(node):
    if node is None:
        return False
    return (node.get(qn("w:val")) or "true").lower() not in ("false", "0", "off")


# --------------------------------------------------------------------- 小工具
def _ensure(parent, tag, order=None, position=None, dry_run=False):
    """取子元素；没有就**按 OOXML 的顺序**插进去（顺序错了 Word 会嫌文件坏）。

    ``order`` 给"有序容器"用（如 `w:tblPr`）：新元素插到**第一个序号比它大的**已有元素之前；
    ``position`` 给"必须排在最前"的用（如 `w:tc` 里的 `w:tcPr`、`w:p` 里的 `w:pPr`）。

    **``dry_run`` 时返回一个游离元素**（不挂进文档树）：上层"取值—比较—设值"的代码不用改写，
    而"缺元素 ⇒ 会新增 ⇒ 算一处改动"的语义天然成立，同时**树上一个字节都不动**
    —— 踩过一次：早先 `dry_run` 也会 `insert`，`--dry-run` 变成了"偷偷改内存"。
    """
    node = parent.find(qn(tag))
    if node is not None:
        return node
    node = ET.Element(qn(tag))
    if dry_run:
        return node
    if position is not None:
        parent.insert(position, node)
        return node
    if order:
        target = order.index(tag) if tag in order else len(order)
        index = len(parent)
        for offset, existing in enumerate(parent):
            name = _name_of(existing.tag)
            if name in order and order.index(name) > target:
                index = offset
                break
        parent.insert(index, node)
        return node
    parent.append(node)
    return node


def _name_of(tag):
    for prefix in ("w", "r", "wp", "a", "mc", "v", "o"):
        namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                     "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                     "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006"}.get(prefix)
        if namespace and tag.startswith("{%s}" % namespace):
            return "%s:%s" % (prefix, tag.split("}", 1)[1])
    return tag


def _set_attribute(node, attribute, value, dry_run):
    if node is None:
        return False
    if node.get(attribute) == u"%s" % value:
        return False
    if not dry_run:
        node.set(attribute, u"%s" % value)
    return True


def _set_border(borders, tag, spec, dry_run):
    node = _ensure(borders, tag, ("w:top", "w:left", "w:bottom", "w:right",
                                  "w:insideH", "w:insideV"), dry_run=dry_run)
    changed = False
    for key, attribute in (("val", "w:val"), ("sz", "w:sz"), ("space", "w:space"),
                           ("color", "w:color")):
        if key in spec and spec[key] is not None:
            if _set_attribute(node, qn(attribute), spec[key], dry_run):
                changed = True
    return changed


def _ensure_flag(parent, tag, order, dry_run):
    """确保某个"开关型"元素在场（如 `w:tblHeader`）；返回是否真的加了。"""
    if parent.find(qn(tag)) is not None:
        return False
    if not dry_run:
        _ensure(parent, tag, order)
    return True


# --------------------------------------------------------------------- 预览
#: 预览用的代表性表格内容（三种形状覆盖"报告里最常见的难点"）
PREVIEW_SHEETS = (
    (u"两列表头（含单位换行）", [[u"项目", u"数值"],
                                 [u"流域面积", u"0.8"],
                                 [u"干流长度", u"1.1"]]),
    (u"多列（列宽不均匀）", [[u"序号", u"项目", u"单位", u"数量", u"备注"],
                            [u"1", u"土方开挖", u"万m3", u"12.8", u"不含超挖"],
                            [u"2", u"混凝土", u"万m3", u"3.5", u""]]),
    (u"长文本单元格", [[u"项目", u"说明"],
                       [u"防洪标准", u"按50年一遇设计、200年一遇校核，枢纽建筑物级别为3级"]],
     ),
)


def preview_body(styles, gap_paragraph=True):
    """把若干候选款式各渲染一张表 → 拼成 `word/document.xml` 的正文片段。"""
    body = []
    for style in styles:
        body.append(_heading_paragraph(u"【%s】%s" % (style.name, _style_note(style))))
        for caption, rows in PREVIEW_SHEETS:
            body.append(_caption_paragraph(caption))
            body.append(_table_xml(rows, style))
        if gap_paragraph:
            body.append(u"<w:p/>")
    return u"".join(body)


def build_preview(path, styles):
    """写一份**预览文档**：每种候选款式各一张表，你用 Word 打开直接看效果。"""
    from .build import write_document
    return write_document(path, preview_body(styles))


def _style_note(style):
    return style.note or u"（无备注）"


def _heading_paragraph(text):
    return (u'<w:p><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>'
            u'<w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr>'
            u'<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % _escape(text)


def _caption_paragraph(text):
    return (u'<w:p><w:r><w:rPr><w:sz w:val="18"/><w:color w:val="595959"/></w:rPr>'
            u'<w:t xml:space="preserve">%s</w:t></w:r></w:p>') % _escape(text)


def _table_xml(rows, style):
    """造一张表：先按款式把 `w:tblPr` 写全，再由 `apply()` 对内存里的 XML 套一遍
    —— **预览与真跑走同一条代码路径**，不是两套逻辑。"""
    from xml.etree import ElementTree as ET
    from .ooxml import qn

    columns = max(len(row) for row in rows)
    grid = u"".join(u'<w:gridCol w:w="%d"/>' % int(9000 / columns) for _ in range(columns))
    xml = [u'<w:tbl xmlns:w="%s"><w:tblPr><w:tblW w:w="0" w:type="auto"/>'
           u'<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>%s</w:tblGrid>'
           % (NAMESPACES["w"], grid)]
    for row_index, row in enumerate(rows):
        xml.append(u"<w:tr>")
        for value in row:
            xml.append(u'<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/></w:tcPr>'
                       u'<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r></w:p></w:tc>'
                       % (int(9000 / columns), _escape(value)))
        xml.append(u"</w:tr>")
    xml.append(u"</w:tbl>")
    table = ET.fromstring(u"".join(xml))
    # 套用款式（内存里改，不写文件）—— 与真跑同一个函数、同一个 dry_run=False 路径
    _apply_to_table(table, style, False)
    # 序列化后去掉根上多余的 xmlns（外层 document 已经声明过；留着虽然合法但很丑）
    return ET.tostring(table, encoding="unicode").replace(
        u' xmlns:w="%s"' % NAMESPACES["w"], u"", 1)


def _escape(text):
    return (u"%s" % text).replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;")
