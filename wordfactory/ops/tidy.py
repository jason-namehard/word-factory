# -*- coding: utf-8 -*-
"""一键清理：无意义的空白行 + 段首/段尾的空格（从网上粘来的文字最常见的两种脏）。

用户 2026-09-22 的原话：「我希望补充…然后再补充一键去除空格、或者无意义的空白行功能，
这个功能抄 `Copy++.exe` 就行」。**两点说明**：

1. **不去拆别人的二进制**：`Copy++.exe` 是第三方程序，反编译它的代码拿来用既不合规也不必要 ——
   这里按**行为**实现（这类工具做的就是"去掉多余空行、去掉多余空格"）。它还有哪些具体行为，
   你指出来我照着补。
2. **默认口径**（2026-10-08 用户看图定稿）：正文里的**空白行整段删掉**、**空白页删掉**、
   段尾空格删掉。**不删**的：前置区（封面/扉页/签字页/前言/目录）里的空白、带换页记号的段、
   挂图的空段。两件要显式开的事，以及为什么：

   * `--trim-leading`：删段首空格 —— **危险**，有的文档用两个半角空格当"首行缩进"
     （`docs/REFERENCE-MACROS.md` §一 记过"段首空两格有四种写法"），删了缩进就没了；
   * `--collapse-space-runs`：段内连续空格压成一个 —— **会破坏表题的"空格居中"**，
     所以它**自动跳过题注段落**（`表X-Y` / `图X-Y` 开头那种）。

   想回到"连续空段只压成一个"的旧保守档：``collapse_blank_lines=True``。
"""

import collections
import re

from ..ooxml import is_on, qn
from ..text import Paragraph

#: 题注段落的识别（与 `ops/captions` 同一套口径）—— 这类段落的空格是排版用的，不能压
CAPTION_RE = re.compile(u"^(?:\\u7eed?\\u8868|\\u56fe)\\s*\\d+")

#: 算作"空格"的字符（**不含制表符**：制表符常是排版用的）
SPACE_CHARS = u" \u00a0\u3000"
#: 「去除空格」删的字符：半角空格 + 不间断空格。**全角空格默认不删**（中文里它常被用来做段首缩进），
#: 依据是用户 2026-09-22 给的 Copy++ 前后对照样本（样本里被删的都是半角空格）。
_REMOVE_SPACES_CHARS = u" \u00a0"
_RUN_RE = re.compile(u"[ \u00a0\u3000]{2,}")
_TRAIL_RE = re.compile(u"[ \u00a0\u3000]+$")
_LEAD_RE = re.compile(u"^[ \u00a0\u3000]+")

DEFAULT_TIDY = {
    # **删掉空白行**（用户 2026-10-08 实测定稿：正文里的空段整段删掉）。
    # 旧的"连续空段压成一个"太保守 —— 用户看图指出"还是有空白行"（一段被空行劈成两半）。
    # 想回到旧口径用 ``collapse_blank_lines``（两者互斥，见 tidy()）。
    "blank_lines": True,
    "collapse_blank_lines": False,  # 旧口径：连续空段压成一个 + 首尾空段删除（保守档）
    "trailing_spaces": True,       # 段尾空格/全角空格
    "trim_leading": False,         # 段首空格（危险：可能是缩进）
    "collapse_space_runs": False,  # 段内连续空格（跳过题注段）
    "caption_skip": True,          # 去空格/压空格时跳过题注段落（它们的空格是排版用的）
    "merge_lines": False,          # 别名：照 Copy++ 的「合并换行」= 删掉全部空段落
    "remove_spaces": False,        # 去除空格：**删掉全部半角/不间断空格**（Copy++ 的「去除空格」）
    "scope": "body",               # body = 正文段落；all = 连表格里的段落一起
    "blank_pages": True,           # 空白页删除：整页只有空格/回车/分页符的页（用户
                                   # 2026-09-30 要求；前置区的空白页由前置区保护挡住）
    # **真实页码**（可选）：``pageprobe`` 用 Word/WPS 排版后读出的"第几块在第几页"，
    # 与 ``list(document.body())`` 一一对应。给了它空白页判定才认得出"自然溢出"的页边界
    # （实测：那份报告第 1 页 22 个空段、第 2 页才是封面，两页之间没有分页符；
    #  不按真实页码，第 1 页的空白页永远删不掉）。由 pipeline/GUI/CLI 探好后传进来。
    "block_pages": None,
    # **前置区保护**（用户 2026-09-27）：封面/扉页/签字页里的空白行是排版，不许删。
    # 识别口径见 frontmatter.py；识别不出前置区时这条自然不生效（等于没保护）。
    "protect_frontmatter": True,
}


def tidy(document, options=None, dry_run=False):
    """按上面那套规则清理；返回报告（`--dry-run` 只数不改）。"""
    opts = dict(DEFAULT_TIDY)
    opts.update(options or {})
    report = collections.Counter()
    protected = set()
    if opts.get("protect_frontmatter", True):
        from .. import frontmatter
        pages = int(opts.get("frontmatter_pages") or 0)
        protected = frontmatter.protected_elements(document, pages=pages or None, block_pages=opts.get("block_pages"))
        if protected:
            report["前置区保护（跳过）"] = len(protected)
    for element, paragraph in _scope_paragraphs(document, opts.get("scope")):
        if id(element) in protected:
            continue
        text = paragraph.text
        if not text:
            continue
        if opts.get("trailing_spaces") and _TRAIL_RE.search(text):
            report["段尾空格"] += 1
            if not dry_run:
                paragraph.replace_regex(u"[ \u00a0\u3000]+$", u"", count=1)
                text = paragraph.text
        if opts.get("trim_leading") and _LEAD_RE.search(text):
            report["段首空格"] += 1
            if not dry_run:
                paragraph.replace_regex(u"^[ \u00a0\u3000]+", u"", count=1)
                text = paragraph.text
        if opts.get("remove_spaces"):
            # **题注段落的空格是排版用的**（"空格居中"规则 B 就靠它）—— 默认整段跳过，
            # 否则一跑就把 17 个表题的空格居中压平（实测踩到）。要连它一起删用 include_captions。
            if opts.get("caption_skip", True) and CAPTION_RE.match(text.strip()):
                report["题注段落（跳过）"] += 1
            else:
                hits = sum(text.count(ch) for ch in _REMOVE_SPACES_CHARS)
                if hits:
                    report["去空格"] += hits
                    if not dry_run:
                        for ch in _REMOVE_SPACES_CHARS:
                            paragraph.replace(ch, u"", count=0)
                        text = paragraph.text
        if opts.get("collapse_space_runs"):
            if opts.get("caption_skip", True) and CAPTION_RE.match(text.strip()):
                report["题注段落（跳过）"] += 1
            else:
                hits = len(_RUN_RE.findall(text))
                if hits:
                    report["连续空格压缩"] += hits
                    if not dry_run:
                        paragraph.replace_regex(u"[ \u00a0\u3000]{2,}", u" ", count=0)
    if opts.get("blank_pages"):
        _remove_blank_pages(document, report, dry_run, opts.get("block_pages"))
    # **三种口径互斥**（同一批段落的三套做法，同时开会重复计数、结果也说不清）：
    # * ``collapse_blank_lines`` —— 旧保守档：连续空段压成一个 + 首尾空段删除；
    # * ``merge_lines`` / ``blank_lines``（**默认**）—— 空行整段删掉（用户 2026-10-08 定稿）。
    if opts.get("collapse_blank_lines"):
        _collapse_blank_paragraphs(document, report, dry_run, protected)
    elif opts.get("merge_lines") or opts.get("blank_lines"):
        _remove_blank_paragraphs(document, report, dry_run, protected)
    # ``total`` = **真改了多少处**：跳过类（…跳过）、子计数（其中…）、以及"空白页数"
    # （那是页数不是改动处数）都不计进来 —— 数字要能对上，别虚报。
    total = sum(count for key, count in report.items()
                if u"跳过" not in key and u"其中" not in key and key != u"空白页数")
    if not dry_run and total:
        document.mark_dirty()
    shown = dict((key, value) for key, value in opts.items() if key != "block_pages")
    return {"op": "tidy", "options": shown, "changes": dict(report), "total": total,
            "dry_run": bool(dry_run)}


def _scope_paragraphs(document, scope):
    """要处理的段落：默认只动**正文段落**（表格里的交给 `tableclean`）。"""
    out = []
    seen = set()
    for element in document.body():
        if element.tag == qn("w:p"):
            seen.add(id(element))
            out.append((element, Paragraph(element)))
    if scope == "all":
        for element in document.part().iter(qn("w:p")):
            if id(element) in seen:
                continue
            out.append((element, Paragraph(element)))
    return out


def carries_page_break(element):
    """这个段是不是**用换页记号撑版面**的（空段但删了会塌版）。

    用户 2026-09-30 的原话（这就是本函数的判据）：
    "这个目录前面的分页符也是不计入删除的……你的检索逻辑没能区分分页符和回车
    和真正的空白行。你需要处理的只有真正的回车引起的空白行和空白格，
    像是 tab 的空格不需要你删除，**分页符也不需要**"。

    三种记号都算：
    * ``w:pPr/w:pageBreakBefore`` —— 段前换页（**只有"开"才算**：``w:val`` 是
      0/false/off 时是显式关掉，不是换页）；
    * ``w:r/w:br w:type="page"``  —— 段内插入的分页符；
    * ``w:pPr/w:sectPr``            —— 分节符（下一页），也决定版面。

    ⚠️ 2026-10-08 修：旧代码只看 ``w:pageBreakBefore`` 元素**在不在**，不看 ``w:val``。
    WPS 导出的报告每个段落都带 ``<w:pageBreakBefore w:val="0"/>``（关），
    结果 288 段全被当成换页 → 空白行一个都删不掉、页码估成 219 页。
    """
    if element.tag != qn("w:p"):
        return False
    pr = element.find(qn("w:pPr"))
    if pr is not None:
        if is_on(pr.find(qn("w:pageBreakBefore"))):
            return True
        if pr.find(qn("w:sectPr")) is not None:
            return True
    for br in element.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return True
    return False


#: 页面上算"有内容"的东西：图片/图形/对象（空白页判定时不能把它们当空）
_CONTENT_TAGS = (qn("w:drawing"), qn("w:pict"), qn("w:object"),
                 qn("w:txbxContent"), qn("w:commentReference"))


def block_has_content(element):
    """这个块**看得到东西吗**（空白页判定用）。

    * 任意后代段落里有非空白文字 → 有内容；
    * 表格里的文字也算（``iter(w:p)`` 覆盖单元格）；
    * 图片 / 图形 / 文本框 / 对象 → 有内容（它没文字但绝不是空白页）。

    注意 ``w:sdt``（内容控件，目录就是它）里的段落也算 —— 只看直接 ``w:p``
    会把整页目录当成空白页。
    """
    for paragraph in element.iter(qn("w:p")):
        if Paragraph(paragraph).text.strip():
            return True
    for node in element.iter():
        if node.tag in _CONTENT_TAGS:
            return True
    return False


def _page_groups(document, block_pages=None):
    """把正文块分成"页"：``[(页号, [块, …]), …]``。

    * 给了 ``block_pages``（**真实页码**，由 ``pageprobe`` 用 Word/WPS 排版后读出来，
      与 ``list(document.body())`` 一一对应）→ 按真实页码分组。**这是唯一能识别
      "自然溢出"造成的空白页的办法**（实测：那份报告第 1 页 22 个空段、第 2 页才是封面，
      两页之间**没有分页符**，只靠分页记号根本分不开）；
    * 没给 → 退回"按换页记号切"（估算法，认不出自然溢出的页边界）。
    """
    blocks = list(document.body())
    groups = []
    if block_pages:
        current = None
        for index, element in enumerate(blocks):
            number = (block_pages[index] if index < len(block_pages) else None)
            if number is None:
                number = current if current is not None else 1
            if current is None or number != current:
                groups.append([number, []])
                current = number
            groups[-1][1].append(element)
        return groups
    bucket = []
    for element in blocks:
        bucket.append(element)
        if element.tag == qn("w:p") and carries_page_break(element):
            groups.append([len(groups) + 1, bucket])
            bucket = []
    if bucket:
        groups.append([len(groups) + 1, bucket])
    return groups


def is_blank_paragraph(element):
    """这个段是不是**可以去掉的空白段**（没文字、也没图/对象）。

    ⚠️ 2026-10-08 实测踩到（**图丢了**）：那份报告里"图 6-1 …"这种图题段是空的，
    图本身挂在**紧跟着的空段**里（``段落文字=''`` 但里面有 ``w:drawing``）。
    旧代码只看 ``Paragraph.text``，把挂图的空段当成空白行删了 —— **整张图没了**。
    所以"空段"的判据必须是 :func:`block_has_content`（文字 + 图/对象一起看）。
    """
    return element.tag == qn("w:p") and not block_has_content(element)


def _remove_blank_paragraphs(document, report, dry_run, protected=frozenset()):
    """**删掉全部空段落**（= Copy++ 的「合并换行」，也是 `tidy` 的**默认**口径）。

    用户 2026-10-08 定稿：正文里的空白行整段删掉（旧口径"连续空段压成一个"太保守 ——
    用户看图指出"还是有空白行"，那是一段话被一个空行劈成两半）。

    三类**不删**（逐条都有实测依据）：
    * ``protected``（前置区封面/扉页/签字页/前言/目录）—— 那些空白是版面；
    * 带换页记号的段 —— 删了版面会塌，报告里记进"空行里带分页符（保留）"；
    * **挂图的空段**（文字为空但里面有 ``w:drawing``）—— 删了图就没了（见 :func:`is_blank_paragraph`）。
    """
    body = document.body()
    doomed = []
    for element in body:
        if id(element) in protected or not is_blank_paragraph(element):
            continue
        if carries_page_break(element):
            report["空行里带分页符（保留）"] = report.get("空行里带分页符（保留）", 0) + 1
            continue
        doomed.append(element)
    for element in doomed:
        report["删除空行"] += 1
        if not dry_run:
            body.remove(element)


def _collapse_blank_paragraphs(document, report, dry_run, protected=frozenset()):
    """连续空白段压成一个；文档开头/结尾的空白段删掉。

    先算出"要删哪些"（纯逻辑），再决定动不动手 —— 这样 `--dry-run` 天然不会改树。
    ``protected`` 里的段落（前置区：封面/扉页/签字页）不参与 —— 封面的空白行是排版，
    删了版面就垮（用户 2026-09-27 明确）。
    """
    body = document.body()
    paragraphs = [element for element in body
                  if element.tag == qn("w:p") and id(element) not in protected]
    if not paragraphs:
        return
    # **带换页记号的空段一律不算"空白段"** —— 它是版面撑出来的，删了就塌
    for element in paragraphs:
        if is_blank_paragraph(element) and carries_page_break(element):
            report["空行里带分页符（保留）"] = report.get("空行里带分页符（保留）", 0) + 1
    blank = [is_blank_paragraph(element) and not carries_page_break(element)
             for element in paragraphs]
    doomed = []
    previous_blank = False
    for element, is_blank in zip(paragraphs, blank):
        if is_blank and previous_blank:
            doomed.append((element, u"空白段压缩"))
            continue
        previous_blank = is_blank
    survivors = [element for element in paragraphs if element not in [item[0] for item in doomed]]
    if survivors:
        if blank[paragraphs.index(survivors[0])]:
            doomed.append((survivors[0], u"首尾空白段删除"))
        last = survivors[-1]
        if last is not survivors[0] and blank[paragraphs.index(last)]:
            doomed.append((last, u"首尾空白段删除"))
    for element, key in doomed:
        report[key] += 1
        if not dry_run:
            body.remove(element)


def _remove_blank_pages(document, report, dry_run, block_pages=None):
    """**空白页检索**（用户 2026-09-30 要求加到一键整理里，2026-10-08 重写）。

    一页里**一个看得到的东西都没有**（只有空段、空格、制表符、回车、分页符；
    页眉页脚不算；**图片/表格/内容控件（目录）都算有内容**）→ 空白页。

    处理：**把这一页里的空段落全删掉**（整页没内容，留着就是给文档多出一张白纸）。
    两种要**保留**的情形：
    * 块上带 ``w:sectPr``（分节符属性）—— 那是**结构**，删了页边距/纸张会变，不赌；
    * 页里有内容（有文字/图/表）—— 不动，那空白段可能是版面。

    **前置区不在此列**：封面/扉页/签字页/前言/目录页都有文字，天然不会被判成空白页
    （用户 2026-10-08 实测口径："第 1 页明明不是封面却留下空白页及其空白行"）。
    所以这里**不再套 ``protected``** —— 一个**没有任何文字**的页不可能是那五种页面。

    页的划分：有 ``block_pages``（真实页码）就按它；没有才退回分页记号估算。
    """
    body = document.body()
    groups = _page_groups(document, block_pages)
    removed = 0
    kept_section = 0
    kept_break = 0
    blank_pages = 0
    for number, group in groups:
        if any(block_has_content(el) for el in group):
            continue                              # 有内容的页不动
        paragraphs = [el for el in group if el.tag == qn("w:p")]
        if not paragraphs:
            continue                              # 只有表格/图片的"页"（无段落）不碰
        blank_pages += 1
        for el in paragraphs:
            pr = el.find(qn("w:pPr"))
            if pr is not None and pr.find(qn("w:sectPr")) is not None:
                kept_section += 1                 # 分节符属性：结构，保留
                continue
            if carries_page_break(el):
                kept_break += 1
                if not block_pages:
                    continue                      # **估算口径**：认不准页边界，换页段一律不赌
            if not dry_run:
                body.remove(el)
            removed += 1
    if removed:
        report["空白页删段"] = removed
        report["空白页数"] = blank_pages
        if not dry_run:
            document.mark_dirty()
    if kept_section:
        report["空白页保留分节符段（跳过）"] = kept_section
    if kept_break:
        report["（其中原带分页符）"] = kept_break

