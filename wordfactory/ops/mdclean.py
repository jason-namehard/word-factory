# -*- coding: utf-8 -*-
"""宏「MarkDown语言清除」：把 Markdown 标记清成普通中文文本。

参考宏 `MarkDown语言清除.bas`（92 行），规格见 `docs/REFERENCE-MACROS.md` §2.1。规则照抄：

| # | 规则 | 出处 |
|---|---|---|
| 1 | 行内代码 `` `X` `` → **中文双引号** `“X”` | `:37-38` |
| 2 | `**粗体**` → `粗体` | `:41-42` |
| 3 | `*斜体*` → `斜体` | `:44-45` |
| 4 | 段首的 `#` **全部**剥掉（`## 标题` → `标题`） | `:55-62` |
| 5 | 再剥掉**一个**段首的 `*` / `-` / `+`（列表标记） | `:64-67` |
| 6 | 段内所有 `*` 删掉 | `:69` |
| 7 | 连续空格折叠成一个、每段 `Trim` | `:83-87` |

**与宏的两处差异（都要说清）**：

1. **实现方式**：宏是 `selectedRange.text = processedText` 整段回写；我们**逐条规则做局部替换**，
   只动匹配到的那一小段，**其余文字的 run 结构（字体/加粗/上下标）原样保留**。
   所以同一个 `**很粗**` 两边都会变成 `很粗`，但我们不会顺手把整段的格式压平。
2. 宏里 `#.` 那段（`:71-81`）是**死代码**（`#` 在前面已被剥光，永远进不去），我们**不实现** ——
   有单测把这个判断留了痕，免得以后被当成漏做。

规则表只写一份（:data:`RULES`），纯文本清洗 `clean_text` 与 XML 清洗 `apply` **共用它**，
两边不会漂移（有单测交叉验证）。
"""

import collections
import re

from ..ooxml import qn
from ..text import Paragraph

#: ``(规则名, 正则, 替换, 次数上限)``；顺序即执行顺序（与宏一致：先行内代码 → 粗体 → 斜体 → 逐段）
#: ``count=0`` 表示全文替换，``1`` 表示只替换第一次出现。
RULES = (
    (u"行内代码", u"`([^`]+)`", u"\u201c\\1\u201d", 0),
    (u"粗体", u"\\*\\*([^*]+)\\*\\*", u"\\1", 0),
    (u"斜体", u"\\*([^*]+)\\*", u"\\1", 0),
    (u"段首井号", u"^[ \\t]*#{1,6}[ \\t]*", u"", 1),
    (u"列表标记", u"^[ \\t]*[*\\-+][ \\t]*", u"", 1),
    (u"段内星号", u"\\*", u"", 0),
    (u"连续空格", u"  +", u" ", 0),
)

#: 值不值得动它（省得对全篇无差别改写）
_MARKERS = re.compile(u"[*`]|^[ \\t]*#|^[ \\t]*[*\\-+][ \\t]")


def looks_like_markdown(text):
    return bool(text) and bool(_MARKERS.search(text))


def clean_text(text):
    """纯文本版清洗（逐段处理，段落用 CR 分隔；与宏的 `Split(text, vbCr)` 对应）。"""
    out = []
    for raw in (text or u"").split(u"\r"):
        para = raw
        for _, pattern, replacement, count in RULES:
            para = re.sub(pattern, replacement, para, count=count)
        out.append(para.strip())
    return u"\r".join(out)


def apply(document, options=None, dry_run=False):
    """按 :data:`RULES` 清洗正文里的 Markdown 标记。返回报告。"""
    opts = dict(options or {})
    paragraphs = _scope_paragraphs(document, opts.get("scope") or "body")
    changes = collections.Counter()
    touched = 0
    samples = []
    for paragraph in paragraphs:
        before = paragraph.text
        if not looks_like_markdown(before):
            continue
        after = clean_text(before)
        if after == before:
            continue
        touched += 1
        for name, pattern, _replacement, _count in RULES:
            hits = len(re.findall(pattern, before))
            if hits:
                changes[name] += hits
        # 行内代码/加粗/斜体这三条会"改字"，逐段规则会"删字符" —— 两类都在逻辑文本上局部替换
        if len(samples) < 5:
            samples.append({"before": before[:60], "after": after[:60]})
        if not dry_run:
            _apply_rules(paragraph)
    total = sum(counts for counts in changes.values() for counts in [counts])
    if not dry_run and touched:
        document.mark_dirty()
    return {"op": "mdclean", "paragraphs": len(paragraphs), "touched": touched,
            "changes": dict(changes), "samples": samples, "total": total,
            "dry_run": bool(dry_run)}


def _apply_rules(paragraph):
    """逐条规则在**逻辑文本**上做局部替换（只动匹配到的那一段，其余 run 不动）。"""
    for _name, pattern, replacement, count in RULES:
        if re.search(pattern, paragraph.text):
            paragraph.replace_regex(pattern, replacement, count=count)


def _scope_paragraphs(document, scope):
    if scope == "all":
        return [Paragraph(element) for element in document.part().iter(qn("w:p"))]
    return [Paragraph(element) for element in document.body() if element.tag == qn("w:p")]
