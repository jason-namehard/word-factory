# -*- coding: utf-8 -*-
"""宏「表格空格回车删除」（去无意义空格）：清掉表格单元格里的空格 / 换行。

参考宏：`表格空格回车删除.bas`（202 行），规格见 `docs/REFERENCE-MACROS.md` §2.11。
三档删除（宏的权威表，`:144-158`）：

| 档 | 删什么 |
|---|---|
| 1 | 半角空格 `U+0020`、不间断空格 `U+00A0` |
| 2 | 段落标记 `Chr(13)`、手动换行 `Chr(11)`、单元格标记 `Chr(7)` |
| 3 | 上面全部（宏的默认档） |

**两处与宏**有意**不同，都要向用户说清楚**：

1. **只删字符，不重建单元格文本**。宏是 `cellRange.text = newText` 整格写回（`：162`）——
   单元格里的字体/加粗/上下标/颜色全被压成一种。XML 层只删字符、**保留 run 属性**，
   结果是"更保真"。要跟宏一样压平就加 `--flat`（会先把整格文字并成一个 run）。
2. **不含全角空格 `U+3000`**（宏也删不掉它，见 §2.11 边界）。中文报告里全角空格常被用来做
   段首缩进，删了可能伤到原意，所以默认**不删**；要删加 `--full-width-space`。

`Chr(13)`（段落标记）在 OOXML 里是 `w:p` 的边界 → "删它" = **把单元格里的多个段落合并成一个**，
保留第一个段落的 `w:pPr`。`Chr(7)`（单元格标记）在 XML 里就是 `w:tc` 本身，**无事可做**。
"""

from xml.etree import ElementTree as ET

from ..ooxml import qn

#: 三档删除的字符集（照宏的表）
SPACE_CHARS = (u"\u0020", u"\u00a0")
FULL_WIDTH_SPACE = u"\u3000"
RETURN_CHARS = (u"\r", u"\v", u"\x07")          # 段落标记 / 手动换行 / 单元格标记

LEVELS = {1: u"仅空格", 2: u"仅回车", 3: u"空格和回车"}

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
#: `w:tab` / `w:br` 这些"不是文字但要算进去"的节点
BREAK_TAGS = ("w:br", "w:cr")


def clean(document, options=None, dry_run=False):
    """清理表格单元格。返回报告（扫了几个单元格、改了几个、删了多少字符）。"""
    opts = dict(options or {})
    level = int(opts.get("level") or 3)
    if level not in LEVELS:
        raise ValueError(u"档位只能是 1 / 2 / 3（现在是 %r）" % level)
    drop_spaces = level in (1, 3)
    drop_returns = level in (2, 3)
    chars = []
    if drop_spaces:
        chars.extend(SPACE_CHARS)
        if opts.get("full_width_space"):
            chars.append(FULL_WIDTH_SPACE)
    tables = _tables(document)
    cells = 0
    changed_cells = 0
    removed = 0
    merged = 0
    for table in tables:
        for cell in table.iter(qn("w:tc")):
            cells += 1
            hit, how_many, merges = _clean_cell(cell, chars, drop_returns,
                                                bool(opts.get("flat")), dry_run)
            removed += how_many
            merged += merges
            if hit:
                changed_cells += 1
    if not dry_run and changed_cells:
        document.mark_dirty()
    return {"op": "table-clean", "level": level, "level_note": LEVELS[level],
            "tables": len(tables), "cells": cells, "changed_cells": changed_cells,
            "chars_removed": removed, "paragraphs_merged": merged,
            "full_width_space": bool(opts.get("full_width_space")),
            "flat": bool(opts.get("flat")), "dry_run": bool(dry_run)}


def _tables(document):
    """正文里所有表格（含嵌套表格：它们也是 `w:tbl`）。"""
    return [element for element in document.body().iter(qn("w:tbl"))]


def _text_nodes(element):
    return list(element.iter(qn("w:t")))


def _clean_cell(cell, chars, drop_returns, flat, dry_run):
    """清理一个单元格。返回 ``(有没有变, 删了几个字符, 合并了几个段落)``。"""
    before = _text_of(cell)
    if not before:
        return False, 0, 0                      # 宏也是"空单元格不动"（`:142`）

    if flat:
        # `--flat`：照宏的做法把整格压平（段合并 → run 合并 → 再删字符），会丢 run 级格式
        removed = sum(before.count(ch) for ch in chars)
        merges = _merge_paragraphs(cell, dry_run) + _merge_runs(cell, dry_run)
        if drop_returns:
            removed += _drop_breaks(cell, dry_run)
        if not dry_run:
            _strip_chars(cell, chars)
            for node in _text_nodes(cell):
                _fix_xml_space(node)
        return (removed > 0 or merges > 0), removed, merges

    removed = _strip_chars(cell, chars, dry_run)
    merges = 0
    if drop_returns:
        removed += _drop_breaks(cell, dry_run)
        merges = _merge_paragraphs(cell, dry_run)          # Chr(13) = 段落边界 → 合并段落
    if not dry_run:
        for node in _text_nodes(cell):
            _fix_xml_space(node)
    return (removed > 0 or merges > 0), removed, merges


def _strip_chars(cell, chars, dry_run=False):
    """删字符（`w:t` 里的空格 / 不间断空格）；返回删掉的个数。"""
    removed = 0
    for node in _text_nodes(cell):
        text = node.text or u""
        new = text
        for ch in chars:
            removed += new.count(ch)
            new = new.replace(ch, u"")
        if new != text and not dry_run:
            node.text = new
    return removed


def _drop_breaks(cell, dry_run=False):
    """删 `w:br` / `w:cr`（= Word 的手动换行 `Chr(11)`）；返回删掉的个数。"""
    tags = tuple(qn(tag) for tag in BREAK_TAGS)
    removed = 0
    for run in list(cell.iter(qn("w:r"))):
        for node in list(run):
            if node.tag in tags:
                removed += 1
                if not dry_run:
                    run.remove(node)
    return removed


def _merge_paragraphs(cell, dry_run=False):
    """把单元格里的多个段落合并成一个（"删段落标记"的 XML 等价动作）。

    保留**第一个**段落的 `w:pPr`（缩进/对齐），后面的段落只贡献它的 run，随后整个段落被移除
    —— 这样"删掉回车"就真的等于把两段文字接在一起（与宏 `cellRange.text = newText` 同效）。
    """
    paragraphs = [element for element in cell if element.tag == qn("w:p")]
    if len(paragraphs) < 2:
        return 0
    first = paragraphs[0]
    merged = 0
    for paragraph in paragraphs[1:]:
        merged += 1
        if dry_run:
            continue
        for child in list(paragraph):
            if child.tag == qn("w:pPr"):
                continue                          # 后段的段落属性丢掉（只留第一段的）
            paragraph.remove(child)
            first.append(child)
        cell.remove(paragraph)
    return merged


def _merge_runs(cell, dry_run=False):
    """把单元格里的 run 合并成一个（`--flat` 用；结果与宏"整格写回"等价）。

    文字并进**第一个** run（保留它的 `w:rPr`），其余 run 连它的文字节点一起移除。
    """
    merged = 0
    for paragraph in [element for element in cell if element.tag == qn("w:p")]:
        runs = [child for child in paragraph if child.tag == qn("w:r")]
        if len(runs) < 2:
            continue
        first = runs[0]
        target = first.find(qn("w:t"))
        if target is None:
            target = ET.SubElement(first, qn("w:t"))
            target.text = u""
        for run in runs[1:]:
            merged += 1
            if dry_run:
                continue
            target.text = (target.text or u"") + _text_of(run)
            paragraph.remove(run)
    return merged


def _fix_xml_space(node):
    """按需维护 `xml:space="preserve"`：首尾有空格就必须保留，否则会被 Word 吃掉。"""
    text = node.text or u""
    if text != text.strip():
        node.set(XML_SPACE, "preserve")
    elif node.get(XML_SPACE) is not None:
        del node.attrib[XML_SPACE]


def _text_of(element):
    return u"".join((node.text or "") for node in element.iter(qn("w:t")))
