# -*- coding: utf-8 -*-
"""宏「规划报告一键宏」的文本部分：批量文本替换 + 两端对齐改左对齐。

参考宏：`规划报告一键宏.bas`（105 行），规格见 `docs/REFERENCE-MACROS.md` §2.12。它做四件事：

| # | 动作 | 本工具 |
|---|---|---|
| 1 | `其它` → `其他` | ✅ 本模块（走**逻辑文本层**，所以跨 run 也能替换） |
| 2 | `东流流经` → `向东流经` | ✅ 本模块 |
| 3 | 段落对齐：两端对齐 → 左对齐 | ✅ 本模块（**含样式继承解析**，见下） |
| 4 | 中文字体 `仿宋` → `宋体` | ✅ 已有 `fonts` / `captions --mode formal` |

**规则外置**（`rules/replacements.json`）：替换对与对齐口径都是数据，不是写死的代码 ——
报告里换个错别字不用改程序。

**为什么对齐必须解析样式继承**（本模块的核心）：Word 的"查找并替换格式"匹配的是**显示出来的
有效格式**（含样式继承）。实测用户的报告：样式 `Normal` 里写着 `w:jc="both"`，
于是**49 个自己没写 `w:jc` 的段落其实是两端对齐的**。只看段落自己的属性会漏掉这 49 段。
所以这里沿 `w:pStyle` 链一路解析到 `docDefaults`，再决定要不要改。
"""

import collections
import io
import json
import os

from ..ooxml import qn
from ..text import Paragraph

#: 出厂规则（取自参考宏 `规划报告一键宏.bas` 的行号）
DEFAULT_REPLACEMENTS = {
    "schema": 1,
    "name": u"文本替换与对齐（默认取自参考宏 规划报告一键宏.bas）",
    "replacements": [
        {"from": u"其它", "to": u"其他", "note": u"参考宏 :20-21"},
        {"from": u"东流流经", "to": u"向东流经", "note": u"参考宏 :33-34（原宏注释为「新增功能」）"},
    ],
    "align": {
        "from": ["both", "distribute"],
        "to": "left",
        "fix_styles": False,
        "note": u"两端对齐→左对齐（参考宏 :44-55）。fix_styles=true 时连样式表里的两端对齐也一起改",
    },
}

VALID_ALIGN = ("left", "center", "right", "both", "distribute", "start", "end")


class TextFixError(Exception):
    """规则文件的问题（人话）。"""


class ReplacementRuleSet(object):
    def __init__(self, data=None, path=None):
        data = dict(data or {})
        self.name = data.get("name") or u"文本替换"
        self.replacements = [dict(item) for item in (data.get("replacements") or [])]
        align = dict(data.get("align") or {})
        self.align_from = [value for value in (align.get("from") or [])]
        self.align_to = align.get("to")
        self.fix_styles = bool(align.get("fix_styles", False))
        self.align_note = align.get("note") or u""
        self.raw = data
        self.path = path

    @classmethod
    def load(cls, path):
        if not path or not os.path.exists(path):
            return cls(DEFAULT_REPLACEMENTS)
        with io.open(path, "r", encoding="utf-8-sig") as handle:
            return cls(json.load(handle), path=path)

    def save(self, path):
        parent = os.path.dirname(os.path.abspath(path))
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with io.open(path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(self.raw, ensure_ascii=False, indent=2) + u"\n")
        return path

    def check(self):
        problems = []
        for index, item in enumerate(self.replacements):
            if not item.get("from"):
                problems.append(u"第 %d 条替换缺 'from'" % (index + 1))
            if "to" not in item:
                problems.append(u"第 %d 条替换缺 'to'" % (index + 1))
            if item.get("from") and item.get("from") == item.get("to"):
                problems.append(u"第 %d 条替换的 from 与 to 一样（%r），会白跑"
                                % (index + 1, item.get("from")))
            extra = set(item) - {"from", "to", "note"}
            if extra:
                problems.append(u"第 %d 条替换里有不认识的键：%s"
                                % (index + 1, u"、".join(sorted(extra))))
        for value in self.align_from:
            if value not in VALID_ALIGN:
                problems.append(u"align.from 里的 %r 不是合法对齐值（%s）"
                                % (value, u"、".join(VALID_ALIGN)))
        if self.align_to is not None and self.align_to not in VALID_ALIGN:
            problems.append(u"align.to=%r 不是合法对齐值" % self.align_to)
        return problems


class AlignResolver(object):
    """按 `w:pStyle` 链解析段落的**有效对齐**（段落自己 → 段落样式 → basedOn → docDefaults）。"""

    def __init__(self, styles_root=None):
        self.styles = {}                       # styleId -> {"based": id, "jc": val, "name": str}
        self.default = None
        #: 段落不写 `w:pStyle` 时用的是**默认段落样式**（`w:default="1"` 那个，通常叫 Normal）——
        #: 它自己也可能写着 `jc=both`。漏了这一步，"没写样式的段落"就会被判成"未指定"而放过
        #: （实测：用户的报告里就是这么漏掉 18 段的）。
        self.default_style_id = None
        if styles_root is None:
            return
        for style in styles_root.iter(qn("w:style")):
            sid = style.get(qn("w:styleId"))
            if not sid:
                continue
            if (style.get(qn("w:type")) in (None, "paragraph")
                    and (style.get(qn("w:default")) or "").lower() in ("1", "true")):
                self.default_style_id = sid
            pr = style.find(qn("w:pPr"))
            jc = pr.find(qn("w:jc")) if pr is not None else None
            based = style.find(qn("w:basedOn"))
            name = style.find(qn("w:name"))
            self.styles[sid] = {"based": based.get(qn("w:val")) if based is not None else None,
                                "jc": jc.get(qn("w:val")) if jc is not None else None,
                                "name": name.get(qn("w:val")) if name is not None else None}
        dd = styles_root.find(qn("w:docDefaults"))
        if dd is not None:
            pr = dd.find(qn("w:pPrDefault"))
            ppr = pr.find(qn("w:pPr")) if pr is not None else None
            jc = ppr.find(qn("w:jc")) if ppr is not None else None
            if jc is not None:
                self.default = jc.get(qn("w:val"))

    def style_ids_with(self, values):
        """样式表里**直接写了**这些对齐值的样式 id（`fix_styles` 用）。"""
        return sorted(sid for sid, entry in self.styles.items() if entry["jc"] in values)

    def effective(self, paragraph_element):
        """返回 ``(对齐值, 来源说明)``。"""
        pr = paragraph_element.find(qn("w:pPr"))
        own = pr.find(qn("w:jc")) if pr is not None else None
        if own is not None and own.get(qn("w:val")):
            return own.get(qn("w:val")), u"段落自己"
        pstyle = pr.find(qn("w:pStyle")) if pr is not None else None
        sid = pstyle.get(qn("w:val")) if pstyle is not None else None
        if not sid:
            sid = self.default_style_id            # 没写样式 = 用默认段落样式
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            entry = self.styles[sid]
            if entry["jc"]:
                return entry["jc"], u"样式 %s" % (entry["name"] or sid)
            sid = entry["based"]
        if self.default:
            return self.default, u"docDefaults"
        return None, u"（未指定）"


def apply(document, rule_set=None, dry_run=False, fix_align=True, fix_styles=False):
    """替换文本 + 把两端对齐改成左对齐。返回报告。"""
    rule_set = rule_set or ReplacementRuleSet(DEFAULT_REPLACEMENTS)
    counts = collections.Counter()
    align_fixed = 0
    align_by_source = collections.Counter()
    # **全文 = 正文 + 表格里的段落**：参考宏作用在 `ActiveDocument.Content`，
    # 那是"整个主故事"（含表格单元格、文本框），不只是 body 的直接子段落。
    # 实测踩过：只遍历 body 的直接子元素会漏掉表格里的 49 个段落（它们才是两端对齐的重灾区）。
    paragraphs = [Paragraph(element) for element in document.part().iter(qn("w:p"))]
    replacements = []
    for paragraph in paragraphs:
        for rule in rule_set.replacements:
            source, target = rule.get("from"), rule.get("to")
            if not source or target is None:
                continue
            hit = paragraph.text.count(source)
            if not hit:
                continue
            replacements.append({"from": source, "to": target, "count": hit,
                                 "paragraph": paragraph.text[:40]})
            counts[u"%s → %s" % (source, target)] += hit
            if not dry_run:
                paragraph.replace(source, target)
    if fix_align and rule_set.align_to:
        resolver = _align_resolver(document)
        from_values = set(rule_set.align_from or [])
        for element in document.part().iter(qn("w:p")):
            current, source = resolver.effective(element)
            if current not in from_values:
                continue
            align_fixed += 1
            align_by_source[source] += 1
            if not dry_run:
                from ..ops.captions import _set_jc, _paragraph_properties
                _set_jc(_paragraph_properties(Paragraph(element)), rule_set.align_to)
        if fix_styles or rule_set.fix_styles:
            styles_root = _styles_root(document)
            if styles_root is not None:
                for style in styles_root.iter(qn("w:style")):
                    pr = style.find(qn("w:pPr"))
                    jc = pr.find(qn("w:jc")) if pr is not None else None
                    if jc is not None and jc.get(qn("w:val")) in from_values:
                        counts[u"样式对齐 %s → %s" % (jc.get(qn("w:val")), rule_set.align_to)] += 1
                        if not dry_run:
                            jc.set(qn("w:val"), rule_set.align_to)
    total = sum(counts.values()) + align_fixed
    if not dry_run and total:
        document.mark_dirty()
        if (fix_styles or rule_set.fix_styles) and any(
                key.startswith(u"样式对齐") for key in counts):
            document.mark_dirty(document.package.STYLES)
    return {"op": "textfix", "rules": rule_set.name, "paragraphs": len(paragraphs),
            "replacements": replacements, "replaced": sum(counts[key] for key in counts
                                                          if not key.startswith(u"样式对齐")),
            "align_fixed": align_fixed, "align_by_source": dict(align_by_source),
            "align_to": rule_set.align_to if fix_align else None,
            "changes": dict(counts), "total": total, "dry_run": bool(dry_run)}


def _align_resolver(document):
    styles_root = _styles_root(document)
    return AlignResolver(styles_root)


def _styles_root(document):
    if not document.package.has(document.package.STYLES):
        return None
    return document.part(document.package.STYLES)
