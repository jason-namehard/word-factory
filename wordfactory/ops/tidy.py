# -*- coding: utf-8 -*-
"""一键清理：无意义的空白行 + 段首/段尾的空格（从网上粘来的文字最常见的两种脏）。

用户 2026-09-22 的原话：「我希望补充…然后再补充一键去除空格、或者无意义的空白行功能，
这个功能抄 `Copy++.exe` 就行」。**两点说明**：

1. **不去拆别人的二进制**：`Copy++.exe` 是第三方程序，反编译它的代码拿来用既不合规也不必要 ——
   这里按**行为**实现（这类工具做的就是"去掉多余空行、去掉多余空格"）。它还有哪些具体行为，
   你指出来我照着补。
2. **默认取保守口径**（与你在表格那件事上的取向一致），只删"明显无意义"的：
   连续空白段压成一个、文档首尾的空白段删掉、段尾的空格/全角空格删掉。
   两件**默认不做**、要显式开的事，以及为什么：

   * `--trim-leading`：删段首空格 —— **危险**，有的文档用两个半角空格当"首行缩进"
     （`docs/REFERENCE-MACROS.md` §一 记过"段首空两格有四种写法"），删了缩进就没了；
   * `--collapse-space-runs`：段内连续空格压成一个 —— **会破坏表题的"空格居中"**，
     所以它**自动跳过题注段落**（`表X-Y` / `图X-Y` 开头那种）。
"""

import collections
import re

from ..ooxml import qn
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
    "blank_lines": True,           # 连续空白段压成一个 + 首尾空白段删除
    "trailing_spaces": True,       # 段尾空格/全角空格
    "trim_leading": False,         # 段首空格（危险：可能是缩进）
    "collapse_space_runs": False,  # 段内连续空格（跳过题注段）
    "caption_skip": True,          # 去空格/压空格时跳过题注段落（它们的空格是排版用的）
    "merge_lines": False,          # 合并换行：**删掉全部空段落**（Copy++ 的「合并换行」）
    "remove_spaces": False,        # 去除空格：**删掉全部半角/不间断空格**（Copy++ 的「去除空格」）
    "scope": "body",               # body = 正文段落；all = 连表格里的段落一起
}


def tidy(document, options=None, dry_run=False):
    """按上面那套规则清理；返回报告（`--dry-run` 只数不改）。"""
    opts = dict(DEFAULT_TIDY)
    opts.update(options or {})
    report = collections.Counter()
    for element, paragraph in _scope_paragraphs(document, opts.get("scope")):
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
    # **两种口径互斥**：`merge_lines`（照 Copy++ 删光空行）与 `blank_lines`（保守：连续空段压一个）
    # 是对同一批段落的两种做法；同时开会重复计数、结果也说不清（踩过：43 + 28 两边都在数）。
    if opts.get("merge_lines"):
        _remove_blank_paragraphs(document, report, dry_run)
    elif opts.get("blank_lines"):
        _collapse_blank_paragraphs(document, report, dry_run)
    total = sum(count for key, count in report.items() if u"跳过" not in key)
    if not dry_run and total:
        document.mark_dirty()
    return {"op": "tidy", "options": opts, "changes": dict(report), "total": total,
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


def _remove_blank_paragraphs(document, report, dry_run):
    """**删掉全部空段落**（Copy++ 的「合并换行」：把被空行隔开的行并成连续行）。

    与 :func:`_collapse_blank_paragraphs`（保守：连续空段只压成一个）**不同** —— 那个是默认口径，
    这个是"照 Copy++ 的行为"（用户 2026-09-22 给了前后对照样本，见 `tests/test_tidy.py`）。
    """
    body = document.body()
    doomed = [element for element in body
              if element.tag == qn("w:p") and not Paragraph(element).text.strip()]
    for element in doomed:
        report["删除空行"] += 1
        if not dry_run:
            body.remove(element)


def _collapse_blank_paragraphs(document, report, dry_run):
    """连续空白段压成一个；文档开头/结尾的空白段删掉。

    先算出"要删哪些"（纯逻辑），再决定动不动手 —— 这样 `--dry-run` 天然不会改树。
    """
    body = document.body()
    paragraphs = [element for element in body if element.tag == qn("w:p")]
    if not paragraphs:
        return
    blank = [not Paragraph(element).text.strip() for element in paragraphs]
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
