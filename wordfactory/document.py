# -*- coding: utf-8 -*-
"""文档层：把"一个打开的 .docx"包装成算子好用的对象。

算子（宏）不该自己去翻 XML 找页面尺寸、也不该自己关心哪个部件被改过。这里提供三件事：

* ``Document.part(name)`` / ``mark_dirty(name)``：部件读写与"改过"登记；
* ``Document.paragraphs()``：正文的逻辑段落（带 run 映射）；
* ``Document.geometry()``：**页面几何**——版心宽度、页边距、左右边界。
  表头居中这类事必须用它（"居中"是相对版心的，不是相对页面的）。
"""

from .ooxml import DocxPackage, local_name, qn
from .text import Paragraph

#: Word 的 dxa：1/20 磅
TWIPS_PER_POINT = 20
#: 表头那"两格"的写法：Word 里是**首行缩进 2 字符**（不是空格）
FIRST_LINE_CHARS = 200


class Geometry(object):
    """一节的页面几何（单位：dxa）。"""

    def __init__(self, page_width, margin_left, margin_right):
        self.page_width = page_width
        self.margin_left = margin_left
        self.margin_right = margin_right
        self.column_width = page_width - margin_left - margin_right

    @property
    def column_center(self):
        """版心中点，**相对左侧页边距**（正文里的坐标都从这里算起）。"""
        return self.column_width / 2.0

    @property
    def column_right(self):
        return self.column_width

    def width_of(self, text, size_half_points):
        """文本宽度（dxa）：全角 = 字号，半角 = 字号一半。

        注意 ``w:sz`` 的单位是**半磅**：``w:sz="24"`` 是 12 磅，一个全角字就是 12 磅 = 240 dxa。
        （这里最容易错的就是忘了除 2 —— 一错，算出来的空格数会差一倍。）
        """
        points = size_half_points / 2.0            # 半磅 → 磅
        total = 0
        for ch in text:
            total += points * TWIPS_PER_POINT * (1.0 if is_wide(ch) else 0.5)
        return total

    def space_width(self, size_half_points):
        """一个半角空格的宽度（dxa）：字号的一半。"""
        return (size_half_points / 2.0) * TWIPS_PER_POINT * 0.5

    def __repr__(self):
        return "<Geometry column=%d dxa>" % self.column_width


def is_wide(ch):
    """全角判定：CJK、全角标点、以及 Word 里按全角排的那几个符号。"""
    return (u"\u2e80" <= ch <= u"\u9fff" or u"\uff00" <= ch <= u"\uffef"
            or ch in u"\u3000\u2018\u2019\u201c\u201d\u2014\u2026")


class Document(object):
    """一个打开的文档（``.docx`` / ``.docm``）。"""

    def __init__(self, path_or_package, writable_parts=("word/document.xml",)):
        if isinstance(path_or_package, DocxPackage):
            self.package = path_or_package
            self._own = False
        else:
            self.package = DocxPackage(path_or_package)
            self._own = True
        #: 允许改的部件 —— 算子只能动白名单里的部件（防止"顺手改了样式表"）
        self.writable_parts = set(writable_parts)

    # ------------------------------------------------------------------ 部件
    @property
    def path(self):
        return self.package.path

    def part(self, name=None):
        name = name or DocxPackage.MAIN
        return self.package.xml(name)

    def mark_dirty(self, name=None):
        name = name or DocxPackage.MAIN
        if name not in self.writable_parts and name != DocxPackage.MAIN:
            raise ValueError(u"这个算子不允许改部件 %s（白名单：%s）"
                             % (name, sorted(self.writable_parts)))
        self.package.mark_dirty(name)

    @property
    def dirty_parts(self):
        return self.package.dirty_parts

    # ------------------------------------------------------------------ 内容
    def body(self, name=None):
        root = self.part(name)
        body = root.find(qn("w:body"))
        return body if body is not None else root

    def paragraphs(self, name=None):
        """正文里的逻辑段落（含表格单元格里的段落 —— 它们也是 ``w:p``）。"""
        out = []
        for element in self.part(name).iter(qn("w:p")):
            out.append(Paragraph(element))
        return out

    def block_children(self, name=None):
        """正文的直接子块（``w:p`` / ``w:tbl`` / ``w:sectPr``），用来判断"下一块是不是表格"。"""
        return [child for child in self.body(name)]

    # ------------------------------------------------------------------ 几何
    def geometry(self, name=None):
        """取**文档级** sectPr 的页面几何。

        局限（写清楚，免得踩坑）：分节符里的 sectPr（一页不同版心）这里不处理——
        取的是正文最后那个 sectPr，也就是文档默认节。按需再扩。
        """
        body = self.body(name)
        sect = body.find(qn("w:sectPr"))
        if sect is None:
            for child in body:
                pr = child.find(qn("w:pPr"))
                if pr is not None:
                    found = pr.find(qn("w:sectPr"))
                    if found is not None:
                        sect = found
        page = sect.find(qn("w:pgSz")) if sect is not None else None
        margins = sect.find(qn("w:pgMar")) if sect is not None else None
        page_width = int(page.get(qn("w:w")) or 11906) if page is not None else 11906
        left = int(margins.get(qn("w:left")) or 1800) if margins is not None else 1800
        right = int(margins.get(qn("w:right")) or 1800) if margins is not None else 1800
        return Geometry(page_width, left, right)

    # ------------------------------------------------------------------ 生命周期
    def save(self, out_path):
        return self.package.save(out_path)

    def close(self):
        if self._own:
            self.package.close()

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()
        return False


def run_size_of(paragraph, default=21):
    """段落里第一个带 ``w:sz`` 的 run 的字号（``w:sz`` 的值 = 半磅）；没有就用默认。

    ``w:sz=21`` 即五号（10.5pt），``24`` 即小四（12pt）—— 没写 sz 就按默认 10.5pt 估。
    """
    for run in paragraph.runs:
        pr = run.find(qn("w:rPr"))
        if pr is not None:
            node = pr.find(qn("w:sz"))
            if node is not None:
                try:
                    return int(node.get(qn("w:val")))
                except (TypeError, ValueError):
                    continue
    return default
