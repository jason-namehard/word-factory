# -*- coding: utf-8 -*-
"""正式版的**字体规范化**规则（外置文件，和上下标规则一个路子）。

用户 2026-09-21 的要求：「正式版则是通体黑色，没有不合格的字体」。
"通体黑色"是明确的；"不合格的字体"是**业务规矩**，只能由他定 —— 所以这里把规矩**外置成文件**：

    {
      "schema": 1,
      "name": "字体规范化",
      "keep":    ["宋体", "黑体", "Times New Roman"],   // 这些字体不动
      "replace": {"仿宋": "宋体", "仿宋_GB2312": "宋体"}, // 明确要换掉的
      "default": "宋体",                                 // 其它"不合格"的字体统一成它
      "black_all": true,                                 // 所有文字设成黑色
      "remove_highlight": true                           // 去掉高亮
    }

**默认值取自参考宏 `规划报告一键宏.bas`（它做的就是 仿宋 → 宋体）**，但那是 2026 年的旧宏口径，
必须由用户确认后才算数 —— 所以在报告里会把"改了多少、动了哪些字体"逐项列出来给他看。

**"没有不合格的字体"要落地，必须覆盖三处，否则会留下看不见的漏洞（这条是实测踩出来的）**：

1. **`w:rFonts` 有四个属性**，不是一个。`ascii` 管 ASCII，`hAnsi` 管 Latin-1（`°`、`é` 之类），
   `eastAsia` 管中日韩，`cs` 管复杂文种。只改 `ascii`/`eastAsia`，`hAnsi`/`cs` 里的旧字体还在
   （实测：1611 个 `eastAsia="仿宋"` 改完之后，仍有 738 个 `rFonts` 含着仿宋）。
2. **段落标记也算**（`w:pPr/w:rPr/w:rFonts`）。它不显示成文字，但那是"这一段回车的字体"，
   用户把光标放到段尾时字体框里看到的就是它。
3. **继承来的字体**。run 自己不写字体时，字体来自字符样式 → 段落样式 → `docDefaults`。
   所以样式表（`word/styles.xml`）和编号表（`word/numbering.xml`）里的字体也得一起换
   （实测：这份文档的字符样式 `26` 就是仿宋；只改正文，继承它的文字照样是仿宋）。
4. **default 只管中文字体**（`default_scope="eastAsia"`，可改成 `"all"`）。实测这份文档里
   `w:cs` 上有 Tahoma、样式表里有 Arial、字体表里有 Symbol/Wingdings——那些是西文与符号字体，
   一股脑换成宋体只会把版面搞坏（符号字体换成宋体，✔ ➜ ★ 会掉字形）。
   `replace` 里的映射是**明确指定**的，四个属性都换；`default` 只兜中文字体。
"""

import io
import json
import os
from xml.etree import ElementTree as ET

from .ooxml import qn

#: 字体规则会动的部件：正文 + 样式表 + 编号表（后两个是"继承来源"，见模块说明第 3 条）。
FONT_PARTS = ("word/document.xml", "word/styles.xml", "word/numbering.xml")

#: `w:rFonts` 的四个属性（见模块说明第 1 条）。
FONT_ATTRS = ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs")

#: 符号字体**永远不碰**：它们靠"哪个码位"出字形，换成宋体后 ✔ ★ ➜ 会变成乱码或直接掉字形。
SYMBOL_FONTS = [u"Symbol", u"Wingdings", u"Wingdings 2", u"Wingdings 3", u"Webdings",
                u"Marlett", u"ZapfDingbats"]

DEFAULT_FONTS = {
    "schema": 1,
    "name": u"字体规范化（keep/replace 取自参考宏 规划报告一键宏.bas:71 —— 需用户确认）",
    "keep": [u"宋体", u"黑体", u"Times New Roman"],
    "replace": {u"仿宋": u"宋体", u"仿宋_GB2312": u"宋体"},
    "default": u"宋体",
    #: default 管哪些属性。``"eastAsia"``（默认）= 只管中文字体；
    #: ``"all"`` = 四个属性都换（西文也会被换成 default，慎用）。
    "default_scope": u"eastAsia",
    "symbol_fonts": SYMBOL_FONTS,
    "black_all": True,
    "remove_highlight": True,
    "note": (u"keep=原样保留；replace=明确映射（四个属性都换）；default=其余中文字体统一成它；"
             u"default_scope 决定 default 管不管西文属性；symbol_fonts 永不触碰。"
             u"改完用 wordfactory audit 复核（末行 AUDIT=PASS）。"),
}

BLACK = u"000000"
#: 中文字体所在的属性
EAST_ASIA = "w:eastAsia"


class FontRuleSet(object):
    def __init__(self, data=None, path=None):
        data = data or {}
        self.name = data.get("name") or u"字体规范化"
        self.keep = set(data.get("keep") or [])
        self.replace = dict(data.get("replace") or {})
        self.default = data.get("default")
        self.default_scope = data.get("default_scope") or u"eastAsia"
        self.symbol_fonts = set(data.get("symbol_fonts") or SYMBOL_FONTS)
        self.black_all = bool(data.get("black_all", True))
        self.remove_highlight = bool(data.get("remove_highlight", True))
        self.note = data.get("note") or u""
        self.path = path
        self.raw = dict(data)

    @classmethod
    def load(cls, path):
        if not path or not os.path.exists(path):
            return cls(DEFAULT_FONTS)
        with io.open(path, "r", encoding="utf-8-sig") as handle:
            return cls(json.load(handle), path=path)

    def save(self, path):
        parent = os.path.dirname(os.path.abspath(path))
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)
        with io.open(path, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(self.raw, ensure_ascii=False, indent=2) + "\n")
        return path

    def target_for(self, font_name, attr=None):
        """这个字体该换成什么？``None`` = 不动。

        ``attr`` 是它在 `w:rFonts` 的哪个属性上（``w:eastAsia`` = 中文字体）。
        区分它是因为 **default 默认只管中文字体**：实测这份文档里 `w:cs` 上有 Tahoma、
        样式表里有 Arial，那都是西文/符号字体，一股脑换成宋体只会把版面搞坏
        （符号字体尤其不能换 —— 见 :data:`SYMBOL_FONTS`）。
        """
        if not font_name:
            return None
        if font_name in self.symbol_fonts:
            return None                                  # 符号字体永不触碰
        if font_name in self.keep:
            return None
        if font_name in self.replace:
            return self.replace[font_name] or None
        if self.default_scope != u"all" and attr and attr != EAST_ASIA:
            return None                                  # default 只管中文属性
        return self.default or None

    def check(self):
        problems = []
        if not isinstance(self.default, (str, type(None))):
            problems.append(u"'default' 必须是字符串或 null")
        if self.default_scope not in (u"eastAsia", u"all"):
            problems.append(u"'default_scope' 只能是 \"eastAsia\" 或 \"all\"，现在是 %r"
                            % self.default_scope)
        for key, value in self.replace.items():
            value_ok = value is None or isinstance(value, str)
            if not value_ok:
                problems.append(u"replace[%r] 必须是字符串或 null" % key)
        overlap = self.keep & set(self.replace)
        if overlap:
            problems.append(u"这些字体同时在 keep 和 replace 里：%s" % u"、".join(sorted(overlap)))
        clash = self.symbol_fonts & (set(self.replace) | self.keep)
        if clash:
            problems.append(u"这些字体同时在 symbol_fonts 和 keep/replace 里，规则会打架：%s"
                            % u"、".join(sorted(clash)))
        return problems


def normalize(document, rule_set=None, dry_run=False, parts=None):
    """正式版：把全文文字统一成黑色（可选去高亮）+ 按规则换字体。

    换字体覆盖三处：run 自己的 `w:rFonts`（四个属性）、段落标记的 `w:rFonts`、
    以及样式表/编号表里的 `w:rFonts`（继承来源）。理由见模块说明。

    返回报告：改了哪些字体（每种、每个属性各多少次）、颜色改了多少、高亮去了多少、动了哪些部件。
    """
    rule_set = rule_set or FontRuleSet(DEFAULT_FONTS)
    if parts is None:
        parts = [name for name in FONT_PARTS
                 if name == document.package.MAIN or document.package.has(name)]
    for name in parts:
        # 宁可当场报错，也不"偷偷只改正文" —— 继承来源没改却报"没有不合格字体"是假报告
        if name != document.package.MAIN and name not in document.writable_parts:
            raise ValueError(
                u"这个算子不允许改部件 %s（白名单：%s）。换字体要连样式表/编号表一起改"
                u"（它们是继承来源），请把 %s 加进 writable_parts。"
                % (name, sorted(document.writable_parts), list(FONT_PARTS)))
    font_changes = {}
    color_changes = 0
    highlight_removed = 0
    text_runs = 0
    touched = set()
    for part_name in parts:
        root = document.part(part_name)
        is_main = part_name == document.package.MAIN
        for element in root.iter():
            if element.tag == qn("w:rFonts"):
                # 正文的 run、段落标记（w:pPr/w:rPr）、样式定义都从这里过 —— 一处覆盖三处
                if _normalize_rfonts(element, rule_set, font_changes, dry_run):
                    touched.add(part_name)
                continue
            if not is_main or element.tag != qn("w:r"):
                continue
            text = "".join((node.text or "") for node in element.findall(qn("w:t")))
            if not text:
                continue
            text_runs += 1
            pr = _run_properties(element, dry_run)
            if rule_set.black_all and _set_color(pr, BLACK, dry_run):
                color_changes += 1
                touched.add(part_name)
            if rule_set.remove_highlight and _drop_highlight(pr, dry_run):
                highlight_removed += 1
                touched.add(part_name)
    if not dry_run:
        for name in sorted(touched):
            document.mark_dirty(name)
    return {"op": "fonts", "runs": text_runs, "fonts": font_changes,
            "colors": color_changes, "highlights": highlight_removed,
            "parts": sorted(touched), "rules": rule_set.name}


def _normalize_rfonts(node, rule_set, font_changes, dry_run):
    """按规则换一个 ``w:rFonts`` 的字体；返回是否真的改了。"""
    changed = False
    for attr in FONT_ATTRS:
        current = node.get(qn(attr))
        if not current:
            continue
        target = rule_set.target_for(current, attr)
        if not target or target == current:
            continue
        key = u"%s → %s（%s）" % (current, target, attr.split(":")[1])
        font_changes[key] = font_changes.get(key, 0) + 1
        changed = True
        if not dry_run:
            node.set(qn(attr), target)
    return changed


def _run_properties(run, dry_run=False):
    """取 run 的 ``w:rPr``；没有就在最前面补一个（**dry-run 时不许动文档**）。"""
    pr = run.find(qn("w:rPr"))
    if pr is None and not dry_run:
        pr = ET.Element(qn("w:rPr"))
        run.insert(0, pr)
    return pr


def _set_color(pr, value, dry_run):
    if pr is None:                             # dry-run 且缺 rPr：算作"要补一个黑色"
        return True
    node = pr.find(qn("w:color"))
    if node is None:
        if dry_run:
            return True
        node = ET.SubElement(pr, qn("w:color"))
        node.set(qn("w:val"), value)
        return True
    if node.get(qn("w:val")) == value:
        return False
    if not dry_run:
        node.set(qn("w:val"), value)
    return True


def _drop_highlight(pr, dry_run):
    """去掉真高亮。``val="none"`` 是"本来就无高亮"，**不动它** ——

    少改一处，也少报一处：原来把它算进"去高亮 N 处"，实测这份文档报的 25 处里
    有 24 处是这种空操作，只有 1 处是真高亮。数字必须对得上眼睛看到的东西。
    """
    if pr is None:
        return False
    node = pr.find(qn("w:highlight"))
    if node is None:
        return False
    if (node.get(qn("w:val")) or u"").lower() == u"none":
        return False
    if not dry_run:
        pr.remove(node)
    return True
