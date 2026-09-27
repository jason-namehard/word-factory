# -*- coding: utf-8 -*-
"""替换规则：**文本替换 / 文字格式替换 / 段落格式替换** 三段合一（用户 2026-09-27 定的形态）。

它取代"直接改 JSON 的规则文件"那套（用户原话：「规则文件…修改难度太高了」）——
规则在界面「替换规则」页用弹窗做出来，存成 ``rules/replace-rules/<名字>.json``；
执行方案里挂一步 ``replace``（**同时只挂一套**，GUI 负责保证）。

规则文件长样::

    {"name": "我的替换",
     "text": [{"find": "其它", "replace": "其他"}],
     "font": [{"from": {"eastAsia": "仿宋_GB2312"}, "to": {"eastAsia": "宋体"}},
              {"from": {"size": "24"}, "to": {"size": "28"}}],
     "para": [{"from": {"align": "both"}, "to": {"align": "left"}}]}

三条实现口径：

* **文本替换**走文本层（跨 run、格式跟第一个 run 走），逐条局部替换 —— 不压平格式；
* **文字格式替换**按 run 的**直接**字符属性匹配（``w:rFonts`` 四属性 / ``w:sz`` 半磅），
  命中才写目标；**符号字体永不碰**（换掉 ✔ ★ 会掉字形）。样式表继承来的匹配留给
  「格式规范化」（那个本来就全局统一），这里只管点名替换；
* **段落格式替换**按段落 ``w:pPr`` 的直接属性匹配（对齐 / 行距 / 首行缩进），命中写目标。
"""

import collections
import io
import json
import os

from .ooxml import qn

#: 符号字体：按码位出字形，换掉会掉字形 —— 任何替换都不许碰
SYMBOL_FONTS = (u"Symbol", u"Wingdings", u"Wingdings 2", u"Wingdings 3",
                u"Webdings", u"Marlett", u"ZapfDingbats", u"OpenSymbol")
#: 字号（半磅）→ 中文名，界面下拉框用（Word 的字号清单里常用的那些）
SIZE_NAMES = {u"42": u"小初", u"36": u"初号", u"28": u"一号", u"32": u"小一",
              u"26": u"二号", u"24": u"小二（12pt）", u"22": u"三号", u"20": u"小三",
              u"18": u"四号", u"16": u"小四", u"15": u"五号", u"14": u"小五",
              u"10.5": u"五号（10.5pt）", u"12": u"六号", u"9": u"小六", u"7": u"七号"}
FONT_ATTRS = ("ascii", "hAnsi", "eastAsia", "cs")


class ReplaceRuleError(Exception):
    """替换规则的问题（人话）。"""


def load(path):
    """读一份规则文件；缺段给空列表。"""
    if not os.path.exists(path):
        raise ReplaceRuleError(u"规则文件不存在：%s" % path)
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        data = json.load(handle)
    data.setdefault("text", [])
    data.setdefault("font", [])
    data.setdefault("para", [])
    return data


def rules_dir(base=None):
    base = base or os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "rules")
    return os.path.join(base, "replace-rules")


def list_rules(base=None):
    """已存的规则清单：``[{name, text, font, para, path}]``（一眼看到有什么）。"""
    directory = rules_dir(base)
    out = []
    if not os.path.isdir(directory):
        return out
    for name in sorted(os.listdir(directory)):
        if not name.endswith(u".json"):
            continue
        try:
            data = load(os.path.join(directory, name))
        except (ValueError, ReplaceRuleError, OSError):
            continue
        out.append({"name": data.get("name") or name[:-5],
                    "text": len(data.get("text") or []),
                    "font": len(data.get("font") or []),
                    "para": len(data.get("para") or []),
                    "path": os.path.join(directory, name)})
    return out


def save(base, name, data):
    """存一份规则（名字即文件名，白名单校验防路径穿越）。"""
    import re
    if not re.match(u"^[\w一-龥·（）()\\- ]{1,40}$", name):
        raise ReplaceRuleError(u"规则名不太好：%r。1–40 个字，别带 . / \\ : 这类字符" % name)
    directory = rules_dir(base)
    if not os.path.isdir(directory):
        os.makedirs(directory)
    data = dict(data)
    data["name"] = name
    path = os.path.join(directory, name + u".json")
    with io.open(path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
    return path


def delete(base, name):
    path = os.path.join(rules_dir(base), name + u".json")
    if not os.path.isfile(path):
        raise ReplaceRuleError(u"没有这套规则：%s" % name)
    os.remove(path)


def apply(document, rules, dry_run=False, scope="all"):
    """执行三段规则；返回报告。``scope``：``all`` 全文（含表格）｜``body`` 只正文段。"""
    report = {"op": "replace", "name": rules.get("name") or u"",
              "text": 0, "font": 0, "para": 0, "details": collections.Counter(),
              "dry_run": bool(dry_run)}
    report["text"] = _apply_text(document, rules.get("text") or [], dry_run, scope)
    report["font"] = _apply_font(document, rules.get("font") or [], dry_run)
    report["para"] = _apply_para(document, rules.get("para") or [], dry_run)
    total = report["text"] + report["font"] + report["para"]
    if not dry_run and total:
        document.mark_dirty()
    report["total"] = total
    return report


# ---------------------------------------------------------------------- 文本
def _scope_paragraph_elements(document, scope):
    if scope == "body":
        for element in document.body():
            if element.tag == qn("w:p"):
                yield element
        return
    for element in document.part().iter(qn("w:p")):
        yield element


def _apply_text(document, entries, dry_run, scope):
    """逐条替换；一条规则吃完全文再下一条（先匹配先应用，链式替换要用户自己留意）。"""
    from .text import Paragraph
    hits = 0
    for entry in entries:
        find = entry.get("find")
        if not find:
            continue
        replacement = entry.get("replace")
        if replacement is None:
            replacement = u""
        for element in _scope_paragraph_elements(document, scope):
            paragraph = Paragraph(element)
            if find not in paragraph.text:
                continue
            # replace 返回"真的换了几处"（文本层只数真变化 → 幂等）
            hits += 1 if paragraph.replace(find, replacement, count=0) else 0
    return hits


# ---------------------------------------------------------------------- 字体
def _run_rpr(run):
    return run.find(qn("w:rPr"))


def _run_text(run):
    return u"".join((node.text or u"") for node in run.findall(qn("w:t")))


def _font_matches(rpr, want):
    """run 的直接字符属性是否命中 ``from`` 条件（列出的键都要相等；没列的不管）。

    支持的键（对齐 Word 字体对话框的栏目，用户 2026-09-27 点名要中文字体/西文字体/颜色）：
    ``eastAsia``/``ascii``（中文字体/西文字体，走 ``w:rFonts``）、``size``（字号，
    ``w:sz`` 半磅）、``color``（字体颜色，``w:color``）、``bold``（字形加粗，``w:b``）。
    """
    if not want:
        return False
    rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
    for key, value in want.items():
        if key in ("eastAsia", "ascii"):
            if rfonts is None or (rfonts.get(qn("w:%s" % key)) or u"") != u"%s" % value:
                return False
        elif key == "size":
            node = rpr.find(qn("w:sz")) if rpr is not None else None
            if node is None or (node.get(qn("w:val")) or u"") != u"%s" % value:
                return False
        elif key == "color":
            node = rpr.find(qn("w:color")) if rpr is not None else None
            if node is None or (node.get(qn("w:val")) or u"").lower() != (u"%s" % value).lower():
                return False
        elif key == "bold":
            node = rpr.find(qn("w:b")) if rpr is not None else None
            on = node is not None and (node.get(qn("w:val")) or u"1") not in (u"0", u"false")
            if on != bool(value):
                return False
        else:
            return False                    # 不认识的键宁可不命中，也不误替换
    return True


def _apply_font(document, entries, dry_run):
    """按 from 命中 → 写 to。只动**带文字**的 run；符号字体永不碰。"""
    hits = 0
    for element in document.part().iter(qn("w:r")):
        text = _run_text(element)
        if not text.strip():
            continue                       # 无文字 run 不参与匹配（ fonts.normalize 管它们）
        rpr = _run_rpr(element)
        rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if rfonts is not None and any(
                (rfonts.get(qn("w:%s" % attr)) or u"") in SYMBOL_FONTS
                for attr in FONT_ATTRS):
            continue
        for entry in entries:
            want = entry.get("from") or {}
            if not _font_matches(rpr, want):
                continue
            if _set_font(rpr, entry.get("to") or {}, dry_run):
                hits += 1
            break                            # 一个 run 只被一条字体规则改（先命中先用）
    return hits


def _set_font(rpr, target, dry_run):
    """把 to 里的字体/字号/颜色/字形写到 rPr（只写点名的项）；没变化返回 False（幂等）。"""
    changed = False
    if rpr is None:
        return False
    from .tablestyle import _ensure, RPR_ORDER
    font_target = {key: value for key, value in target.items() if key in FONT_ATTRS}
    if font_target:
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            if dry_run:
                changed = True
            else:
                rfonts = _ensure(rpr, "w:rFonts", RPR_ORDER, 0, dry_run=False)
        for key, value in font_target.items():
            attr = qn("w:%s" % key)
            if rfonts.get(attr) != u"%s" % value:
                if not dry_run:
                    rfonts.set(attr, u"%s" % value)
                changed = True
    if "size" in target:
        value = u"%s" % target["size"]
        for tag in ("w:sz", "w:szCs"):
            node = rpr.find(qn(tag))
            if node is None:
                if dry_run:
                    changed = True
                    continue
                node = _ensure(rpr, tag, RPR_ORDER, 0, dry_run=False)
            if node.get(qn("w:val")) != value:
                if not dry_run:
                    node.set(qn("w:val"), value)
                changed = True
    if "color" in target:
        value = u"%s" % target["color"]
        node = rpr.find(qn("w:color"))
        if node is None:
            if dry_run:
                changed = True
            else:
                node = _ensure(rpr, "w:color", RPR_ORDER, 0, dry_run=False)
        if (node.get(qn("w:val")) or u"").lower() != value.lower():
            if not dry_run:
                node.set(qn("w:val"), value)
            changed = True
    if "bold" in target:
        want_on = bool(target["bold"])
        node = rpr.find(qn("w:b"))
        is_on = node is not None and (node.get(qn("w:val")) or u"1") not in (u"0", u"false")
        if is_on != want_on:
            if not dry_run:
                if node is None and want_on:
                    _ensure(rpr, "w:b", RPR_ORDER, 0, dry_run=False)
                elif node is not None and not want_on:
                    node.set(qn("w:val"), u"0")    # 关字形不是删元素：Word 的口径就是 val=0
            changed = True
    return changed


# ---------------------------------------------------------------------- 段落
def _para_matches(pr, want):
    """段落直接属性是否命中（align/line/first_line_chars；列出的键都要相等）。"""
    if not want:
        return False
    if "align" in want:
        jc = pr.find(qn("w:jc")) if pr is not None else None
        if jc is None or (jc.get(qn("w:val")) or u"") != u"%s" % want["align"]:
            return False
    if "line" in want:
        spacing = pr.find(qn("w:spacing")) if pr is not None else None
        if spacing is None or (spacing.get(qn("w:line")) or u"") != u"%s" % want["line"]:
            return False
    if "first_line_chars" in want:
        ind = pr.find(qn("w:ind")) if pr is not None else None
        if ind is None or (ind.get(qn("w:firstLineChars")) or u"") != u"%s" % want["first_line_chars"]:
            return False
    return True


def _apply_para(document, entries, dry_run):
    hits = 0
    for element in document.part().iter(qn("w:p")):
        pr = element.find(qn("w:pPr"))
        if pr is None:
            continue
        for entry in entries:
            want = entry.get("from") or {}
            if not _para_matches(pr, want):
                continue
            if _set_para(pr, entry.get("to") or {}, dry_run):
                hits += 1
            break
    return hits


def _set_para(pr, target, dry_run):
    from .tablestyle import _ensure, PPR_ORDER
    changed = False
    if "align" in target:
        node = _ensure(pr, "w:jc", PPR_ORDER, 0, dry_run=dry_run)
        if node.get(qn("w:val")) != u"%s" % target["align"]:
            if not dry_run:
                node.set(qn("w:val"), u"%s" % target["align"])
            changed = True
    if "line" in target:
        node = _ensure(pr, "w:spacing", PPR_ORDER, 0, dry_run=dry_run)
        if node.get(qn("w:line")) != u"%s" % target["line"]:
            if not dry_run:
                node.set(qn("w:line"), u"%s" % target["line"])
            changed = True
    if "first_line_chars" in target:
        node = _ensure(pr, "w:ind", PPR_ORDER, 0, dry_run=dry_run)
        if node.get(qn("w:firstLineChars")) != u"%s" % target["first_line_chars"]:
            if not dry_run:
                node.set(qn("w:firstLineChars"), u"%s" % target["first_line_chars"])
            changed = True
    return changed
