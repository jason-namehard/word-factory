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


# --------------------------------------------------------------- 生成（段落配方生成器）
#: `cleanText`（`段落配方生成器.bas:406-418`）：先 Trim，再去掉**一个**尾部的 CR / LF。
#: 注意它**不**去内部的换行 —— 这正是"多段落 TEXT 会在配方里折成一行无前缀文本"的来源。
def _clean_text(text):
    text = text.strip()
    if text.endswith(u"\r"):
        text = text[:-1]
    if text.endswith(u"\n"):
        text = text[:-1]
    return text


def _is_highlighted(pr):
    if pr is None:
        return False
    node = pr.find(qn("w:highlight"))
    if node is None:
        return False
    return (node.get(qn("w:val")) or u"").lower() not in (u"", u"none")


def _paragraph_mark_highlighted(paragraph):
    pr = paragraph.element.find(qn("w:pPr"))
    rpr = pr.find(qn("w:rPr")) if pr is not None else None
    return _is_highlighted(rpr)


def scan_highlights(document, trim_last_char=False, skip_recipes=True):
    """识别高亮（模式 1）：**正文里所有高亮片段**都算变量。

    与参考宏的差别（两处，都要说清楚）：

    * 参考宏扫的是**选区**；外置工具没有选区，规则化后扫**整篇正文**（用户 2026-09-21 已裁决接受）。
      已经存在的配方区间会被跳过，免得把上一份配方当正文再生成一遍。
    * 参考宏的"隐形修正"（把高亮片段**去掉最后一个字符**，`段落配方生成器.bas:177-213`）
      默认**不照抄**（`trim_last_char=False`）：那是为了绕开 Word 在边界字符上读
      `HighlightColorIndex` 的偏差；XML 里高亮状态是**精确**的，多砍一个字符反而会把值改错
      （比如高亮的 `3.5` 会变成 `3.`）。要与宏逐字一致就打开 `trim_last_char`。
    """
    chars, states = [], []
    for paragraph in _scope_paragraphs(document, skip_recipes):
        text = paragraph.text
        position = 0
        for run in paragraph.runs:
            run_text = u"".join((node.text or "") for node in run.findall(qn("w:t")))
            if not run_text:
                continue
            highlighted = _is_highlighted(run.find(qn("w:rPr")))
            for offset in range(len(run_text)):
                chars.append(text[position + offset])
                states.append(highlighted)
            position += len(run_text)
        # 段落标记也进流（参考宏的选区文本里就有它），它的高亮状态决定它归哪一段
        chars.append(u"\r")
        states.append(_paragraph_mark_highlighted(paragraph))
    if chars and chars[-1] == u"\r":
        chars.pop()                      # 末尾那个不代表"下一段"
        states.pop()
    return _segments(chars, states, trim_last_char)


def _text_lines(content):
    """把一段 TEXT 内容切成正文行：内容的**内部换行**在文档里是折行的（无前缀行，即 RAW）。

    为什么在这里就切开：参考宏是 `"TEXT:" & content & vbCrLf` 一次写进文档，Word 把里面的 CR
    变成段落分隔 —— 于是**文档里的样子**是 `TEXT:前半` + `后半` 两段。我们让内存里的行与该形态一致，
    这样"生成出来的文档"再解析回来，行结构分毫不差（有单测钉这条）。
    """
    parts = content.split(u"\r")
    lines = [("TEXT", _clean_text(parts[0]))]
    for part in parts[1:]:
        lines.append(("RAW", _clean_text(part)))
    return [item for item in lines if item[1] or item[0] == "TEXT"]


def _segments(chars, states, trim_last_char):
    """按参考宏的扫描方式切段：非高亮→高亮 时吐 TEXT，高亮→非高亮 时吐 VAR。"""
    lines = []
    pending = []
    index = 0
    total = len(chars)
    while index < total:
        if states[index]:
            text = _clean_text(u"".join(pending))
            pending = []
            if text:
                lines.extend(_text_lines(text))
            start = index
            while index < total and states[index]:
                index += 1
            span = u"".join(chars[start:index])
            if trim_last_char and len(span) > 1:
                lines.append(("VAR", span[:-1]))
                pending = [span[-1]]      # 被"修正"掉的那个字符归到后面的文本
            else:
                lines.append(("VAR", span))
        else:
            pending.append(chars[index])
            index += 1
    text = _clean_text(u"".join(pending))
    if text:
        lines.extend(_text_lines(text))
    return lines


def scan_special_chars(document, spec_char, skip_recipes=True):
    """识别特定字符（模式 2）：占位符每次出现算一个变量，**值就是占位符本身**
    （`段落配方生成器.bas:315`），不是该位置原有的文字。"""
    if not spec_char:
        raise RecipeError(u"识别模式 2 必须给一个占位符字符串（--char）")
    lines = []
    paragraphs = _scope_paragraphs(document, skip_recipes)
    for index, paragraph in enumerate(paragraphs):
        text = paragraph.text
        if not text.strip():
            lines.append(("TEXT", u""))            # 空段落写成 TEXT:（`:292-295`）
        else:
            position = 0
            while True:
                found = text.find(spec_char, position)
                if found < 0:
                    if position < len(text):
                        lines.extend(_text_lines(text[position:]))
                    break
                if found > position:
                    lines.extend(_text_lines(text[position:found]))
                lines.append(("VAR", spec_char))
                position = found + len(spec_char)
        if index < len(paragraphs) - 1:
            lines.append(("TEXT", u""))            # 段间补一个 TEXT:（`:332-334`）
    return lines


def _scope_paragraphs(document, skip_recipes):
    """正文段落；`skip_recipes` 时跳过已有的配方区间（从 `=== 段落配方` 到 `=== 配方结束`）。"""
    paragraphs = []
    inside = False
    for element in document.body():
        if element.tag != qn("w:p"):
            continue
        paragraph = Paragraph(element)
        text = paragraph.text
        if u"=== 段落配方" in text:
            inside = True
            continue
        if u"=== 配方结束" in text:
            inside = False
            continue
        if inside and skip_recipes:
            continue
        paragraphs.append(paragraph)
    return paragraphs


def rows_with_prefixes(lines):
    """从正文行算出 xlsx 的行：``[(A 列前缀, B 列值), …]``。

    A 列 = 该变量**前面那一段 TEXT**（`varPrefixes.Add lastTextSegment`），给人看的上下文；
    每遇到一个变量就清空（`lastTextSegment = ""`），所以相邻两个变量之间没文本时 A 列是空的。
    """
    rows = []
    prefix = u""
    for kind, content in lines:
        if kind == "TEXT":
            prefix = content
        elif kind == "VAR":
            rows.append((prefix, content))
            prefix = u""
    return rows


def generate(document, options, dry_run=False):
    """生成配方（模式 1 高亮 / 模式 2 特定字符）：写 xlsx + 把配方文本追加到文档末尾。

    ``options``：``mode``（`highlight`|`chars`）、``char``（模式 2 的占位符）、``name``（配方名）、
    ``excel_file``、``sheet_name``、``out_xlsx``、``append``（是否把配方写进文档，默认 True）、
    ``trim_last_char``。
    """
    mode = options.get("mode") or "highlight"
    if mode == "highlight":
        lines = scan_highlights(document, bool(options.get("trim_last_char")))
    elif mode == "chars":
        lines = scan_special_chars(document, options.get("char"))
    else:
        raise RecipeError(u"未知模式 %r（只有 highlight / chars 两种）" % mode)
    variables = [content for kind, content in lines if kind == "VAR"]
    recipe = Recipe(options.get("excel_file") or u"数据表.xlsx",
                    options.get("sheet_name") or u"Sheet1",
                    len(variables), options.get("name") or u"默认配方", lines)
    rows = rows_with_prefixes(lines)
    report = {"op": "recipe-generate", "mode": mode, "recipe": recipe.name,
              "variables": len(variables), "rows": rows,
              "excel": recipe.excel_file, "sheet": recipe.sheet_name,
              "paragraphs": 0, "dry_run": bool(dry_run), "recipe_text": recipe.to_text()}
    if dry_run:
        return report, recipe, None
    from ..xlsx import write_workbook
    xlsx_path = write_workbook(options["out_xlsx"], recipe.sheet_name, rows)
    if options.get("append", True):
        body = document.body()
        anchor = body.find(qn("w:sectPr"))
        # 参考宏是 `Range.text = ...`，Word 会把 CR 变成段落分隔 → 一行一个 w:p
        # （TEXT 内容里带的 CR 也会各自成为一段）
        for raw_line in recipe.to_text().split(u"\r\n"):
            for part in raw_line.split(u"\r"):
                paragraph = _plain_paragraph(part)
                if anchor is None:
                    body.append(paragraph)
                else:
                    body.insert(list(body).index(anchor), paragraph)
                report["paragraphs"] += 1
        document.mark_dirty()
    return report, recipe, xlsx_path
