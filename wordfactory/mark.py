# -*- coding: utf-8 -*-
"""两版输出：**验证版**把工具动过的地方标成蓝色，让用户用眼睛判对错；**正式版**通体黑、字体合规。

用户 2026-09-21 的要求：

> 「修改完成的部分应该有验证版和正式版两种选择方式，验证版把文字修改后的部分改成 255 蓝色，
>    让用户可以判断改的是否对，正式版则是通体黑色，没有不合格的字体。」

- 蓝色取 **`0000FF`**（Word 的"蓝色"；他的文档里本来就有 1 处 `0000FF`，口径一致）。
- 验证版只标**工具动过的那一段**（例如题注的"编号 + 中间空格"），不把没动过的名字也涂蓝 ——
  否则用户没法判断到底是哪几个字被改了。
- 正式版由 :mod:`wordfactory.fonts` 负责：全部文字设成黑色 + 去高亮 + 按规则换字体。
"""

from .ooxml import qn
from .text import apply_character_property

#: Word 的"蓝色"
VERIFY_BLUE = u"0000FF"


def mark_ranges(document, ranges, color=VERIFY_BLUE):
    """把若干逻辑区间标成给定颜色；返回标过的 run 数。

    ``ranges`` 是 ``[(paragraph, start, end), …]``。
    """
    touched = 0
    for paragraph, start, end in ranges:
        if end <= start:
            continue
        touched += apply_character_property(paragraph, start, end, "color", color)
    if touched:
        document.mark_dirty()
    return touched


def clear_marks(document, color=VERIFY_BLUE):
    """把验证版留下的蓝色撤掉（正式版之前的清理；正式版自己也会把颜色统一成黑）。"""
    removed = 0
    for part_name in sorted(getattr(document, "writable_parts", {"word/document.xml"})):
        for run in document.part(part_name).iter(qn("w:r")):
            pr = run.find(qn("w:rPr"))
            if pr is None:
                continue
            node = pr.find(qn("w:color"))
            if node is not None and node.get(qn("w:val")) == color:
                pr.remove(node)
                removed += 1
    if removed:
        document.mark_dirty()
    return removed


def count_color(document, color=VERIFY_BLUE):
    """数一数文档里**本来就有**多少处这个颜色（不是我标的）。

    实测：这份文档的标题里本来就有 1 处 ``0000FF``。不先说清楚，用户看到蓝色的标题
    会以为是我改的 —— 那就反而干扰了"判断改的对不对"。
    """
    found = 0
    for run in document.part().iter(qn("w:r")):
        pr = run.find(qn("w:rPr"))
        node = pr.find(qn("w:color")) if pr is not None else None
        if node is not None and node.get(qn("w:val")) == color:
            found += 1
    return found


def verify(document, plans):
    """验证版：把计划里"动过的那一段"标蓝。"""
    ranges = []
    for item in plans:
        if not item["changed"]:
            continue
        ranges.append((item["paragraph"], 0, item["mark_end"]))
    return mark_ranges(document, ranges)
