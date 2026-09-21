# -*- coding: utf-8 -*-
"""上下标规则的**外置接口**：规则写在 JSON 文件里，工具按它用正则查找并应用。

用户的要求（2026-09-21）：

> 「上标需要做规则，有两种上下标，一个是智能上下标，另一个是通用上标，m 后加数字无脑上标是通用上标，
>    这是早期宏的局限性导致的，其实接下来可以统一做智能上下标，工具箱有个外置的接口，
>    可以录入上下标规则，然后工具箱可以用正则查找并应用。」

所以本模块提供**两种规则写法**，一个文件里可以混用：

1. **字面量 + 逐字符类型**（复刻参考宏的字典，`智能上下标.bas:74-97`）::

       {"id": "m2", "match": "m2", "kinds": "NS", "note": "m2 —— m 正常、2 上标"}

   ``kinds`` 与 ``match`` **等长**，逐字符：``N`` 正常、``S`` 上标、``B`` 下标。

2. **正则 + 作用目标**（用户自己要的"用正则查找并应用"）::

       {"id": "unit-m", "pattern": "m([2-9]+)", "target": "group:1",
        "kind": "superscript", "note": "m 后跟 2-9 → 上标（通用上标）"}

   也可以 ``"target": "all"`` 把整段匹配设成上标/下标。

共同字段：``id``（稳定标识，报告里引用）、``note``（给人看）、``enabled``（默认 true）、
``not_before`` / ``not_after``（正则，匹配的**前一个/后一个字符**必须**不**符合它，
用来表达参考宏里那张"边界字符表"）。

顺序即优先级：**先匹配先应用**，被前面规则占用的字符不会被后面的规则再动
（对应参考宏的"特殊符号最高优先级"）。
"""

import io
import json
import os
import re

KINDS = {"N": None, "S": "superscript", "B": "subscript"}
MISSING_SENTINEL = u"\u0000"

#: 参考宏 `智能上下标.bas:100-111` 的边界字符表（逐字符实测 20 个，含半角 = 与全角 ＝）
BOUNDARY_CHARS = u"，。！？；：\"=-+*×/÷（）【】《》＝％"


class RuleError(Exception):
    """规则文件本身有问题（用户可见）。"""


class Rule(object):
    def __init__(self, data):
        self.raw = dict(data)
        self.id = data.get("id") or data.get("match") or data.get("pattern")
        self.note = data.get("note") or u""
        self.enabled = bool(data.get("enabled", True))
        self.match = data.get("match")
        self.kinds = data.get("kinds")
        self.pattern = data.get("pattern")
        self.target = data.get("target", "all")
        self.kind = data.get("kind")
        self.not_before = data.get("not_before")
        self.not_after = data.get("not_after")
        self._regex = None
        self._before = None
        self._after = None
        self.problems = []
        self.compile()          # 构造时就编译：不先 validate 也能直接用

    # ------------------------------------------------------------------ 校验与编译
    def compile(self):
        problems = []
        self._regex = self._before = self._after = None
        if not self.id:
            problems.append(u"规则缺少 id")
        if self.match and self.pattern:
            problems.append(u"%s：'match' 与 'pattern' 只能给一个" % self.id)
        if self.match:
            if not self.kinds:
                problems.append(u"%s：用 'match' 就必须给等长的 'kinds'（N/S/B）" % self.id)
            elif len(self.kinds) != len(self.match):
                problems.append(u"%s：'kinds'(%d) 与 'match'(%d) 长度不等"
                                % (self.id, len(self.kinds), len(self.match)))
            elif any(ch not in KINDS for ch in self.kinds):
                problems.append(u"%s：'kinds' 里只能出现 N/S/B" % self.id)
        elif self.pattern:
            try:
                self._regex = re.compile(self.pattern)
            except re.error as exc:
                problems.append(u"%s：正则编译失败 —— %s" % (self.id, exc))
            if self.kind not in (None, "superscript", "subscript"):
                problems.append(u"%s：'kind' 只能是 superscript / subscript" % self.id)
            if isinstance(self.target, str) and self.target.startswith("group:"):
                if self._regex is not None:
                    try:
                        index = int(self.target.split(":", 1)[1])
                    except ValueError:
                        problems.append(u"%s：'target' 的 group 编号不是整数" % self.id)
                    else:
                        if index < 1 or index > self._regex.groups:
                            problems.append(u"%s：正则只有 %d 个捕获组，取不到 group:%d"
                                            % (self.id, self._regex.groups, index))
            elif self.target != "all":
                problems.append(u"%s：'target' 只能是 'all' 或 'group:N'" % self.id)
        else:
            problems.append(u"%s：既没有 'match' 也没有 'pattern'" % self.id)
        for name, value in (("not_before", self.not_before), ("not_after", self.not_after)):
            if value:
                try:
                    compiled = re.compile(value)
                except re.error as exc:
                    problems.append(u"%s：%s 正则编译失败 —— %s" % (self.id, name, exc))
                else:
                    if name == "not_before":
                        self._before = compiled
                    else:
                        self._after = compiled
        self.problems = problems
        return problems

    def __repr__(self):
        return "<Rule %s %s>" % (self.id, "enabled" if self.enabled else "off")


class RuleSet(object):
    def __init__(self, data=None, path=None):
        data = data or {}
        self.schema = data.get("schema", 1)
        self.name = data.get("name") or u"上下标规则"
        self.note = data.get("note") or u""
        self.rules = [Rule(raw) for raw in data.get("rules", [])]
        self.path = path

    # ------------------------------------------------------------------ 文件
    @classmethod
    def load(cls, path):
        if not os.path.exists(path):
            raise RuleError(u"规则文件不存在：%s" % path)
        with io.open(path, "r", encoding="utf-8-sig") as handle:
            try:
                data = json.load(handle)
            except ValueError as exc:
                raise RuleError(u"%s 不是合法 JSON：%s" % (path, exc))
        return cls(data, path=path)

    def to_dict(self):
        return {"schema": self.schema, "name": self.name, "note": self.note,
                "rules": [rule.raw for rule in self.rules]}

    def save(self, path):
        parent = os.path.dirname(os.path.abspath(path))
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with io.open(path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(self.to_dict(), ensure_ascii=False, indent=2))
            handle.write("\n")
        self.path = path
        return path

    # ------------------------------------------------------------------ 校验
    def validate(self):
        """返回问题清单（空 = 没问题）。**应用到文档之前必须先过这一关。**"""
        problems = []
        seen = set()
        for rule in self.rules:
            problems.extend(rule.problems or rule.compile())
            if rule.id in seen:
                problems.append(u"id 重复：%s" % rule.id)
            seen.add(rule.id)
        return problems

    @property
    def active(self):
        return [rule for rule in self.rules if rule.enabled]

    # ------------------------------------------------------------------ 应用
    def apply_paragraph(self, paragraph):
        """按顺序把规则应用到**一个段落**上；返回改动清单。

        已匹配过的字符不会被后面的规则再动 —— 这就是"先匹配先应用"。
        """
        from .text import set_vertical_align

        text = paragraph.text
        if not text:
            return []
        claimed = []
        changes = []

        def overlaps(start, end):
            for c_start, c_end in claimed:
                if start < c_end and c_start < end:
                    return True
            return False

        for rule in self.active:
            for start, end, spans in self._matches(rule, text):
                if overlaps(start, end):
                    continue
                if not self._neighbours_ok(rule, text, start, end):
                    continue
                claimed.append((start, end))
                for span_start, span_end, kind in spans:
                    if kind is None or span_start >= span_end:
                        continue
                    touched = set_vertical_align(paragraph, span_start, span_end, kind)
                    if touched:
                        changes.append({"rule": rule.id, "kind": kind,
                                        "start": span_start, "end": span_end,
                                        "text": text[span_start:span_end]})
        return changes

    def _neighbours_ok(self, rule, text, start, end):
        """``not_before`` / ``not_after`` 是**排除**条件：前/后一个字符**命中了就否决**这条匹配。

        （段首/段尾没有邻居，用 NUL 代替，普通模式匹配不上它 —— 即"没有邻居"不算被排除。）
        """
        if rule._before is not None:
            before = text[start - 1] if start > 0 else MISSING_SENTINEL
            if rule._before.match(before):
                return False
        if rule._after is not None:
            after = text[end] if end < len(text) else MISSING_SENTINEL
            if rule._after.match(after):
                return False
        return True

    def _matches(self, rule, text):
        """产出 ``(整段起, 整段止, [(要设属性的起, 止, kind), …])``。"""
        if rule.match:
            index = text.find(rule.match)
            while index >= 0:
                # 相邻且同类型的字符合成一段：否则一个词会被切成一个字一个 run
                spans = []
                for offset, letter in enumerate(rule.kinds):
                    kind = KINDS[letter]
                    at = index + offset
                    if spans and spans[-1][2] == kind and spans[-1][1] == at:
                        spans[-1] = (spans[-1][0], at + 1, kind)
                    else:
                        spans.append((at, at + 1, kind))
                yield index, index + len(rule.match), spans
                index = text.find(rule.match, index + 1)
            return
        for match in rule._regex.finditer(text):
            if rule.target == "all":
                spans = [(match.start(), match.end(), rule.kind)]
            else:
                index = int(rule.target.split(":", 1)[1])
                try:
                    span = match.span(index)
                except IndexError:
                    span = (-1, -1)
                if span[0] < 0:
                    continue
                spans = [(span[0], span[1], rule.kind)]
            yield match.start(), match.end(), spans


def apply_to_part(root, ruleset, limit=None):
    """对整篇（一个 XML 部件）应用规则；返回 ``{rule_id: 次数}`` 与明细。"""
    from .text import paragraphs

    counts = {}
    details = []
    for paragraph in paragraphs(root):
        for change in ruleset.apply_paragraph(paragraph):
            counts[change["rule"]] = counts.get(change["rule"], 0) + 1
            if limit is None or len(details) < limit:
                details.append(change)
    return {"counts": counts, "details": details,
            "total": sum(counts.values())}


DEFAULT_RULES = {
    "schema": 1,
    "name": u"上下标规则（默认，取自参考宏 智能上下标.bas:74-97）",
    "note": (u"两种写法都能用：1) match+kinds 逐字符（N 正常 / S 上标 / B 下标）；"
             u"2) pattern+target+kind 正则。顺序即优先级，先匹配先应用。"
             u"改完先用 `rules check` 校验，再用 --dry-run 看会改多少处。"),
    "rules": [
        {"id": "Vmax", "match": "Vmax", "kinds": "NBBB", "note": u"Vmax → max 下标"},
        {"id": "Qpl", "match": "Qpl", "kinds": "NBB", "note": u"Qpl → pl 下标"},
        {"id": "m2", "match": "m2", "kinds": "NS", "note": u"m2"},
        {"id": "m3", "match": "m3", "kinds": "NS", "note": u"m3"},
        {"id": "m4", "match": "m4", "kinds": "NS", "note": u"m4"},
        {"id": "m5", "match": "m5", "kinds": "NS", "note": u"m5"},
        {"id": "cm2", "match": "cm2", "kinds": "NNS", "note": u"cm2"},
        {"id": "km2", "match": "km2", "kinds": "NNS", "note": u"km2"},
        {"id": "qm", "match": "qm", "kinds": "NB", "note": u"qm → m 下标"},
        {"id": "Qm", "match": "Qm", "kinds": "NB", "note": u"Qm → m 下标"},
        {"id": "KP", "match": "KP", "kinds": "NB", "note": u"KP → P 下标"},
        {"id": "CV", "match": "CV", "kinds": "NB", "note": u"CV → V 下标"},
        {"id": "H24P", "match": "H24P", "kinds": "NBBB", "note": u"H24P"},
        {"id": "hR", "match": "hR", "kinds": "NB", "note": u"hR → R 下标"},
        {"id": "H24", "match": "H24", "kinds": "NBB", "note": u"H24 → 24 下标"},
        {"id": "celsius", "match": u"°C", "kinds": "SN", "note": u"°C —— ° 上标、C 正常"},
        {"id": "mm", "match": "mm", "kinds": "NN", "note": u"mm —— 两个都正常（参考宏里的空操作）"},
        {"id": "unit-m-generic",
         "pattern": "m([2-9]+)", "target": "group:1", "kind": "superscript",
         "enabled": False,
         "note": (u"通用上标：m 后跟 2-9 就无脑上标（早期宏 单位上标.bas 的做法，"
                  u"m1 不上标）。默认关着 —— 上面的字典规则已经覆盖 m2..m5；"
                  u"要「无脑」覆盖 m6..m9 就把它打开。")},
    ],
}


def default_ruleset():
    return RuleSet(DEFAULT_RULES)


def write_default(path):
    return default_ruleset().save(path)
