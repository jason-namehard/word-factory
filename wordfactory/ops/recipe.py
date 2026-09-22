# -*- coding: utf-8 -*-
"""宏「段落配方」这套的两个算子：**重配**（读配方 + xlsx → 重建段落）与**生成**（扫描 → 配方 + xlsx）。

参考宏：`段落重配.bas`（281 行）与 `段落配方生成器.bas`（419 行），契约见
`docs/REFERENCE-MACROS.md` §2.8 / §2.9 / §3。这里复刻的是**行为**，不是代码结构。

没有"选区"怎么办（外置工具没有 Word 的选区）：用户 2026-09-21 已裁决**接受规则化** ——
所以"识别高亮"= 处理**正文里所有高亮片段**，"识别特定字符"= 处理**指定的占位符字符串**。
两条规则都是确定的、可复现的，比"当时选中了什么"更可靠。

**两条已知差异（会写进报告，不藏着）**：

1. 参考宏把结果**追加到当前文档末尾**；本工具同样追加（保持行为一致），但写到**新文件**（绝不覆盖输入）。
2. 参考宏用 Excel 的 `AutoFit` 定列宽；XML 层没有等价写法，我们按内容**估**一个宽度。
"""

from xml.etree import ElementTree as ET

from ..ooxml import qn
from ..recipe import MISSING, Recipe, RecipeError
from ..text import Paragraph
from ..xlsx import XlsxError, read_column_b

#: 参考宏写结果时的固定装饰行（`段落重配.bas:276`）
REBUILD_OPEN = u"=== 重建段落 ==="
REBUILD_CLOSE = u"=== 结束 ==="


def find_recipe_text(document):
    """从文档正文里找最后一段配方（`=== 段落配方` 到 `=== 配方结束`）。

    参考宏的做法是"读选区"；外置工具没有选区，就找**文档里最后一份配方**——
    生成端把配方追加在文档末尾，所以"最后一份"就是最新的那份。
    """
    texts = []
    for element in document.body():
        if element.tag == qn("w:p"):
            texts.append(Paragraph(element).text)
    start = None
    end = None
    for index, text in enumerate(texts):
        if u"=== 段落配方" in text:
            start = index
        elif u"=== 配方结束" in text:
            end = index
    if start is None or end is None or end < start:
        return None
    return u"\r\n".join(texts[start:end + 1])


def read_recipe(path):
    """从 `.docx`（末尾的配方）或 `.txt`（直接是配方文本）里读配方。"""
    import io
    import os

    if not os.path.exists(path):
        raise RecipeError(u"文件不存在：%s" % path)
    if path.lower().endswith((".docx", ".docm")):
        from ..document import Document
        with Document(path) as document:
            text = find_recipe_text(document)
        if text is None:
            raise RecipeError(u"%s 里没找到配方（找 `=== 段落配方` … `=== 配方结束` 之间那一段）" % path)
        return Recipe.parse(text)
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        return Recipe.parse(handle.read())


def resolve_xlsx(recipe, document_path, data_dir=None):
    """按参考宏的查找顺序定位 xlsx（§3.1）：**文档同目录 → 当前目录**，外加一个显式目录。

    参考宏的顺序是 ①`ActiveDocument.Path` ②`CurDir`；外置工具再加一个 `--data-dir`
    （批量处理时数据表常常集中放一处）。
    """
    import os

    name = recipe.excel_file
    candidates = []
    if os.path.isabs(name):
        candidates.append(name)
    if data_dir:
        candidates.append(os.path.join(data_dir, name))
    if document_path:
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(document_path)), name))
    candidates.append(os.path.join(os.getcwd(), name))
    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    raise XlsxError(u"找不到数据表 %r。找过这些位置：\n  - %s\n（可以用 --data-dir 直接指定目录）"
                    % (name, u"\n  - ".join(candidates)))


def values_for(recipe, xlsx_path):
    """按契约取值：`SHEET_NAME:` 指定的表、B 列、第 n 个变量在第 n+1 行。"""
    values = read_column_b(xlsx_path, recipe.sheet_name or None,
                           max_rows=max(recipe.variable_count, 0))
    return values


def rebuild(document, recipe, values, dry_run=False):
    """把重建出来的段落**追加到文档末尾**（复刻 `段落重配.bas:274-277`）。"""
    text = recipe.reconstruct(values)
    lines = [u""] + [u""] + [REBUILD_OPEN] + text.split(u"\n") + [REBUILD_CLOSE]
    report = {"op": "recipe-rebuild", "recipe": recipe.name,
              "excel": recipe.excel_file, "sheet": recipe.sheet_name,
              "variables": recipe.variable_count,
              "values_used": len(values), "missing": text.count(MISSING),
              "paragraphs": len(lines), "dry_run": bool(dry_run),
              "text": text}
    if dry_run:
        return report
    body = document.body()
    anchor = body.find(qn("w:sectPr"))
    for line in lines:
        paragraph = _plain_paragraph(line)
        if anchor is None:
            body.append(paragraph)
        else:
            body.insert(list(body).index(anchor), paragraph)
    document.mark_dirty()
    return report


def _plain_paragraph(text):
    """一个不带任何格式的 `w:p`（参考宏写出来的就是纯文本）。

    `w:t` 上的 `xml:space="preserve"` 不能省：否则首尾空格会被 Word 吃掉。
    """
    paragraph = ET.Element(qn("w:p"))
    if text == u"":
        return paragraph
    run = ET.SubElement(paragraph, qn("w:r"))
    node = ET.SubElement(run, qn("w:t"))
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    node.text = text
    return paragraph
