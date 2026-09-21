# -*- coding: utf-8 -*-
"""看一眼一个 ``.docx`` 里有什么 —— 只读，不改。

这个命令现在是本工具**唯一能跑的**能力，但它不只是"玩具"：它量出来的正是后面每个宏都要面对的难点——
**一句话被拆成多少个 run**（Word 按格式变化切 run，还有拼写检查/修订/lang 造成的碎片），
以及文档里有多少高亮、彩色字、上下标。这些数字直接决定替换类宏的难度，也是差分测试夹具的依据。
"""

from .ooxml import DocxPackage, local_name, qn

#: 常见"故事"部件（正文之外的文字也在文档里，宏很容易漏掉）
STORY_PARTS = ("word/document.xml", "word/header1.xml", "word/header2.xml",
               "word/header3.xml", "word/footer1.xml", "word/footer2.xml",
               "word/footer3.xml", "word/footnotes.xml", "word/endnotes.xml")


def _paragraph_runs(paragraph):
    """一个段落里直接属于它的 ``w:r`` 数量（不含文本框内的）。"""
    return [child for child in paragraph if child.tag == qn("w:r")]


def _text_of(paragraph):
    parts = []
    for run in _paragraph_runs(paragraph):
        for node in run.iter(qn("w:t")):
            parts.append(node.text or "")
    return "".join(parts)


def _has_highlight(run):
    return run.find(qn("w:rPr") + "/" + qn("w:highlight")) is not None \
        or run.find(qn("w:rPr") + "/" + qn("w:shd")) is not None


def _color_of(run):
    node = run.find(qn("w:rPr") + "/" + qn("w:color"))
    return node.get(qn("w:val")) if node is not None else None


def _vert_align(run):
    node = run.find(qn("w:rPr") + "/" + qn("w:vertAlign"))
    return node.get(qn("w:val")) if node is not None else None


def scan_part(pkg, name):
    """统计一个 XML 部件里的段落/表格/run 情况。"""
    root = pkg.xml(name)
    paragraphs = list(root.iter(qn("w:p")))
    tables = list(root.iter(qn("w:tbl")))
    runs = list(root.iter(qn("w:r")))
    split = []                     # 一段被切成多个 run 的段落
    highlight = 0
    colored = 0
    super_or_sub = 0
    empty_runs = 0
    for paragraph in paragraphs:
        own = _paragraph_runs(paragraph)
        if len(own) > 1:
            split.append((len(own), _text_of(paragraph)[:40]))
        for run in own:
            text = "".join((node.text or "") for node in run.iter(qn("w:t")))
            if not text:
                empty_runs += 1
            if _has_highlight(run):
                highlight += 1
            if _color_of(run) not in (None, "auto"):
                colored += 1
            if _vert_align(run) in ("superscript", "subscript"):
                super_or_sub += 1
    return {
        "part": name,
        "paragraphs": len(paragraphs),
        "tables": len(tables),
        "runs": len(runs),
        "paragraphs_with_multiple_runs": len(split),
        "worst_split": max([n for n, _ in split], default=0),
        "highlighted_runs": highlight,
        "colored_runs": colored,
        "super_or_subscript_runs": super_or_sub,
        "empty_runs": empty_runs,
        "samples": sorted(split, reverse=True)[:5],
    }


def text_report(path, limit=20, grep=None, part=None):
    """逐段看"逻辑文本 ↔ run"的映射 —— 这是判断一个宏能不能做对的第一步。

    ``grep`` 给关键词时只列包含它的段落（对配方/表头这类"按内容挑段落"的功能特别有用）。
    """
    from .text import Paragraph
    from .ooxml import qn as _qn

    with DocxPackage(path) as pkg:
        name = part or DocxPackage.MAIN
        root = pkg.xml(name)
        rows = []
        for index, element in enumerate(root.iter(_qn("w:p")), 1):
            paragraph = Paragraph(element)
            text = paragraph.text
            if not text.strip():
                continue
            if grep and grep not in text:
                continue
            rows.append({"index": index, "text": text, "runs": paragraph.run_map()})
        return {"file": pkg.path, "part": name, "paragraphs": rows,
                "shown": len(rows[:limit])}


def format_text_report(report, limit=20):
    lines = [u"文件：%s ｜ 部件：%s ｜ 命中段落：%d"
             % (report["file"], report["part"], len(report["paragraphs"]))]
    for row in report["paragraphs"][:limit]:
        pieces = [u"[%s]" % text for _run, text in row["runs"]]
        lines.append(u"")
        lines.append(u"#%d  %d 个 run：%s" % (row["index"], len(row["runs"]), row["text"]))
        if len(row["runs"]) > 1:
            lines.append(u"    切成：%s" % u" + ".join(
                p.replace(u"\n", u"\\n").replace(u"\t", u"\\t") for p in pieces))
    if len(report["paragraphs"]) > limit:
        lines.append(u"")
        lines.append(u"…还有 %d 段（用 --limit 调大）"
                     % (len(report["paragraphs"]) - limit))
    return "\n".join(lines)


def inspect(path, show_parts=True):
    """文档结构总览（部件清单 + 每个故事的段落/表格/run 统计）。"""
    with DocxPackage(path) as pkg:
        result = {"file": pkg.path, "size": len(pkg.read_bytes(pkg.MAIN)),
                  "parts": None, "stories": []}
        if show_parts:
            result["parts"] = [(name, len(pkg.read_bytes(name)))
                               for name in pkg.part_names]
        for name in STORY_PARTS:
            if pkg.has(name):
                result["stories"].append(scan_part(pkg, name))
        return result


def format_report(info):
    lines = [u"文件：%s" % info["file"]]
    if info.get("parts") is not None:
        lines.append(u"部件：%d 个" % len(info["parts"]))
        for name, size in info["parts"]:
            marker = u"  ← 正文" if name == DocxPackage.MAIN else u""
            lines.append(u"    %-38s %8d B%s" % (name, size, marker))
    for story in info["stories"]:
        lines.append(u"")
        lines.append(u"【%s】" % story["part"])
        lines.append(u"  段落 %d ｜ 表格 %d ｜ run %d ｜ 空 run %d"
                     % (story["paragraphs"], story["tables"], story["runs"],
                        story["empty_runs"]))
        lines.append(u"  被切成多个 run 的段落：%d（最多一段 %d 个 run）← 这是替换类宏的难点"
                     % (story["paragraphs_with_multiple_runs"], story["worst_split"]))
        lines.append(u"  高亮 run %d ｜ 彩色 run %d ｜ 上下标 run %d"
                     % (story["highlighted_runs"], story["colored_runs"],
                        story["super_or_subscript_runs"]))
        for count, sample in story["samples"]:
            lines.append(u"    %d 个 run：「%s…」" % (count, sample))
    return "\n".join(lines)
