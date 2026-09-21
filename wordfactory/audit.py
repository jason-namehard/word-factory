# -*- coding: utf-8 -*-
"""体检（audit）：把「正式版 = 通体黑色、没有不合格的字体」变成**可复算**的判据。

为什么不看工具自己的报告：工具报告的是"它以为改了哪些"，而用户要的是"文件现在到底怎么样"。
所以这里**重新打开写出来的那个文件**，按 OOXML 的继承链算出每个有文字的 run **实际生效**的字体，
再逐项核对，末行给 ``AUDIT=PASS`` / ``AUDIT=FAIL <原因>``。

字体继承链（少算一层就会漏判）：run 自己的 ``w:rFonts`` → 字符样式（``w:rStyle``）
→ 段落样式（``w:pStyle``）→ ``docDefaults`` → 主题字体（``w:eastAsiaTheme`` 之类指向 theme1.xml）。

分两档报，区别在"这段字体到底会不会显示成文字"：

* **FAIL**：有文字的 run 生效字体不合格 / 任何非黑颜色 / 真高亮（``val != none``）。
* **提示**：段落标记、无文字的 run、样式定义、编号表、字体表里还提到不合格字体——
  它们当前不显示，但"留着迟早咬人"（用户下次在那儿打字就是旧字体），所以必须说出来。
"""

import collections
import os

from .ooxml import DocxPackage, local_name, qn
from .fonts import BLACK, DEFAULT_FONTS, FONT_ATTRS, FontRuleSet

THEME_PART = "word/theme/theme1.xml"
FONT_TABLE_PART = "word/fontTable.xml"

#: `w:rFonts` 的"主题字体"属性 → 它在主题里对应哪个槽位
THEME_SLOTS = {"w:asciiTheme": "latin", "w:hAnsiTheme": "latin",
               "w:eastAsiaTheme": "ea", "w:cstheme": "cs"}

#: 明确写了字体的属性 → 主题槽位（用于回落到主题字体时取哪一项）
ATTR_SLOTS = {"w:ascii": "latin", "w:hAnsi": "latin", "w:eastAsia": "ea", "w:cs": "cs"}

UNSET = u"（未指定）"


def _theme_fonts(theme_root):
    """主题里的字体槽位：``{"latin": "Calibri", "ea": "", "cs": ""}``。"""
    slots = {"latin": "", "ea": "", "cs": ""}
    if theme_root is None:
        return slots
    for block in theme_root.iter():
        if local_name(block.tag) not in ("a:majorFont", "a:minorFont"):
            continue
        if local_name(block.tag) == "a:majorFont":
            continue                                  # 正文用 minorFont；major 只给标题
        for child in block:
            slot = local_name(child.tag)
            if slot in slots:
                slots[slot] = child.get("typeface") or ""
    return slots


def _fonts_of(node, theme):
    """读一个 ``w:rFonts``：直接写的字体 + （有主题属性时）解析出的主题字体。

    **主题槽位是空的时候不算"取到了字体"**：实测这份文档的主题里 `a:ea typeface=""`，
    而目录样式（toc 1/toc 2 → Normal）写的是 ``w:eastAsiaTheme="minorEastAsia"``。
    空槽位等于"没指定"，继承链要继续往下走（走到 docDefaults），否则会把
    36 个目录段落误报成"字体算不出来"。
    """
    direct, themed = {}, {}
    if node is None:
        return direct, themed
    for attr in FONT_ATTRS:
        value = node.get(qn(attr))
        if value:
            direct[attr] = value
    for key, slot in THEME_SLOTS.items():
        if node.get(qn(key)) and theme.get(slot):
            themed[slot] = theme[slot]
    return direct, themed


class FontResolver(object):
    """按继承链算"某个 run 实际生效的字体"。"""

    def __init__(self, styles_root=None, theme_root=None):
        self.theme = _theme_fonts(theme_root)
        self.styles = {}                 # styleId -> {"based","direct","themed","name","type"}
        self.defaults = ({}, {})
        if styles_root is None:
            return
        for style in styles_root.iter(qn("w:style")):
            sid = style.get(qn("w:styleId"))
            if not sid:
                continue
            name = style.find(qn("w:name"))
            based = style.find(qn("w:basedOn"))
            pr = style.find(qn("w:rPr"))
            direct, themed = _fonts_of(pr.find(qn("w:rFonts")) if pr is not None else None,
                                       self.theme)
            self.styles[sid] = {"based": based.get(qn("w:val")) if based is not None else None,
                                "direct": direct, "themed": themed,
                                "name": name.get(qn("w:val")) if name is not None else None,
                                "type": style.get(qn("w:type"))}
        dd = styles_root.find(qn("w:docDefaults"))
        pr = dd.find(qn("w:rPrDefault")) if dd is not None else None
        rpr = pr.find(qn("w:rPr")) if pr is not None else None
        self.defaults = _fonts_of(rpr.find(qn("w:rFonts")) if rpr is not None else None,
                                  self.theme)

    def _lookup(self, source, attr):
        """从一处字体定义里取值：先直接写的，再主题的。"""
        direct, themed = source
        if direct.get(attr):
            return direct[attr]
        slot = ATTR_SLOTS.get(attr)
        if slot and themed.get(slot):
            return themed[slot]
        return None

    def _style_chain(self, sid, attr):
        seen = set()
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            entry = self.styles[sid]
            value = self._lookup((entry["direct"], entry["themed"]), attr)
            if value:
                return value, u"样式 %s" % (entry["name"] or sid)
            sid = entry["based"]
        return None, None

    def effective(self, rfonts_node, rstyle, pstyle, attr):
        """返回 ``(字体名, 来源)``；全都没有就是 ``(None, None)``。"""
        value = self._lookup(_fonts_of(rfonts_node, self.theme), attr)
        if value:
            return value, u"run 自带"
        if rstyle:
            value, source = self._style_chain(rstyle, attr)
            if value:
                return value, u"字符%s" % source
        if pstyle:
            value, source = self._style_chain(pstyle, attr)
            if value:
                return value, u"段落%s" % source
        value = self._lookup(self.defaults, attr)
        if value:
            return value, u"docDefaults"
        return None, None


def _text_of(run):
    return u"".join((node.text or "") for node in run.findall(qn("w:t")))


def _paragraph_text(paragraph):
    return u"".join((node.text or "") for node in paragraph.iter(qn("w:t")))


def _bad(rule_set, font_name, attr=None):
    """这个字体按规则算不算"不合格"（会不会被换掉）。``attr`` 决定 default 管不管它。"""
    if not font_name or font_name == UNSET:
        return False
    return bool(rule_set.target_for(font_name, attr))


def _rstyle_of(pr):
    node = pr.find(qn("w:rStyle")) if pr is not None else None
    return node.get(qn("w:val")) if node is not None else None


def _pstyle_of(paragraph):
    pr = paragraph.find(qn("w:pPr"))
    node = pr.find(qn("w:pStyle")) if pr is not None else None
    return node.get(qn("w:val")) if node is not None else None


def audit(source, rule_set=None):
    """体检一个文档（路径或已打开的 :class:`~wordfactory.document.Document`）。"""
    from .document import Document
    from .ooxml import PackageError

    rule_set = rule_set or FontRuleSet(DEFAULT_FONTS)
    own = not isinstance(source, Document)
    if own:
        document = Document(source)
    else:
        document = source
    package = document.package
    try:
        styles_root = package.xml(DocxPackage.STYLES) if package.has(DocxPackage.STYLES) else None
        theme_root = package.xml(THEME_PART) if package.has(THEME_PART) else None
        resolver = FontResolver(styles_root, theme_root)

        effective = collections.Counter()
        bad_effective = collections.defaultdict(list)
        declared = []                      # 提示档：不显示成文字的不合格字体
        for part_name in ["word/document.xml"] + sorted(
                n for n in package.part_names
                if n.endswith(".xml") and n != "word/document.xml"):
            if not package.has(part_name):
                continue
            root = package.xml(part_name)
            for paragraph in root.iter(qn("w:p")):
                pstyle = _pstyle_of(paragraph)
                for run in paragraph.findall(qn("w:r")):
                    pr = run.find(qn("w:rPr"))
                    rfonts = pr.find(qn("w:rFonts")) if pr is not None else None
                    rstyle = _rstyle_of(pr)
                    text = _text_of(run)
                    if not text:
                        _collect_declared(declared, rule_set, rfonts,
                                          u"%s（段内没有文字的 run）" % part_name, text)
                        continue
                    for attr in FONT_ATTRS:
                        name, where = resolver.effective(rfonts, rstyle, pstyle, attr)
                        if name is None:
                            name = UNSET
                        if attr == "w:eastAsia":
                            effective[name] += 1
                        if _bad(rule_set, name, attr):
                            bad_effective[name].append(
                                {"where": where, "attr": attr, "text": text[:40]})
                # 段落标记：w:pPr/w:rPr/w:rFonts
                pr = paragraph.find(qn("w:pPr"))
                rpr = pr.find(qn("w:rPr")) if pr is not None else None
                rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
                _collect_declared(declared, rule_set, rfonts,
                                  u"%s（段落标记）" % part_name, _paragraph_text(paragraph))
            if part_name != "word/document.xml":
                for node in root.iter(qn("w:rFonts")):
                    _collect_declared(declared, rule_set, node,
                                      u"%s（部件里的字体定义）" % part_name, u"")

        bad_theme = [{"slot": slot, "font": font} for slot, font in sorted(resolver.theme.items())
                     if _bad(rule_set, font)]

        bad_colors = []
        real_highlights = []
        for part_name in sorted(n for n in package.part_names if n.endswith(".xml")):
            root = package.xml(part_name)
            for node in root.iter(qn("w:color")):
                value = node.get(qn("w:val"))
                if value and value.upper() not in (BLACK.upper(), "AUTO"):
                    bad_colors.append({"part": part_name, "value": value})
            for node in root.iter(qn("w:highlight")):
                value = node.get(qn("w:val"))
                if value and value != "none":
                    real_highlights.append({"part": part_name, "value": value})

        # 字体表只是"这份文档提过这些字体"的清单，没有文字用它。**不改它**——把 仿宋 改名会跟
        # 已有的 宋体 条目撞成重复项。所以只提示，不算 FAIL。
        if package.has(FONT_TABLE_PART):
            for node in package.xml(FONT_TABLE_PART).iter(qn("w:font")):
                name = node.get(qn("w:name"))
                if _bad(rule_set, name):
                    declared.append({"where": u"%s（字体清单）" % FONT_TABLE_PART,
                                     "font": name, "attr": u"name", "text": u""})

        reasons = []
        if bad_effective:            reasons.append(u"有文字的 run 里还有不合格字体：%s"
                           % u"、".join(u"%s %d 个" % (k, len(v))
                                        for k, v in sorted(bad_effective.items())))
        if bad_colors:
            reasons.append(u"还有 %d 处非黑颜色" % len(bad_colors))
        if real_highlights:
            reasons.append(u"还有 %d 处高亮" % len(real_highlights))
        return {
            "file": package.path,
            "rules": rule_set.name,
            "runs_with_text": sum(effective.values()),
            "effective": dict(effective.most_common()),
            "bad_effective": dict((k, v[:3]) for k, v in bad_effective.items()),
            "declared": declared,
            "theme_fonts": dict(resolver.theme),
            "bad_theme_fonts": bad_theme,
            "bad_colors": bad_colors[:10],
            "highlights": real_highlights[:10],
            "verdict": "PASS" if not reasons else "FAIL",
            "reasons": reasons,
        }
    finally:
        if own:
            document.close()


def _collect_declared(out, rule_set, rfonts, where, text):
    """收集"当前不显示成文字、但留着不合格字体"的位置。"""
    if rfonts is None:
        return
    for attr in FONT_ATTRS:
        value = rfonts.get(qn(attr))
        if value and _bad(rule_set, value, attr):
            out.append({"where": where, "font": value,
                        "attr": attr.split(":")[1], "text": text[:30]})


def format_audit(report):
    """人看的体检报告。末行是 ``AUDIT=PASS`` / ``AUDIT=FAIL``（方便当闸门用）。"""
    lines = [u"文件：%s" % report["file"],
             u"字体规则：%s" % report["rules"],
             u"有文字的 run：%d 个 ｜ 生效中文字体（eastAsia 口径）：%s"
             % (report["runs_with_text"],
                u"、".join(u"%s %d" % (k, v) for k, v in report["effective"].items())
                or u"（无）"),
             u""]
    if report["bad_effective"]:
        lines.append(u"✗ 生效字体里还有不合格的：")
        for font, items in sorted(report["bad_effective"].items()):
            lines.append(u"    %s（例：%s ← %s「%s」）"
                         % (font, items[0]["text"] or u"（空）",
                            items[0]["where"], items[0]["attr"]))
    else:
        lines.append(u"✓ 生效字体里没有不合格的字体")
    if report["bad_colors"]:
        lines.append(u"✗ 还有非黑颜色 %d 处：%s"
                     % (len(report["bad_colors"]),
                        u"、".join(u"%s=%s" % (c["part"], c["value"])
                                   for c in report["bad_colors"][:5])))
    else:
        lines.append(u"✓ 没有非黑颜色（所有 xml 部件都查了）")
    if report["highlights"]:
        lines.append(u"✗ 还有高亮 %d 处：%s"
                     % (len(report["highlights"]),
                        u"、".join(u"%s=%s" % (h["part"], h["value"])
                                   for h in report["highlights"][:5])))
    else:
        lines.append(u"✓ 没有高亮")
    if report["bad_theme_fonts"]:
        lines.append(u"⚠ 主题字体（word/theme/theme1.xml）里有不合格的：%s"
                     % u"、".join(u"%s=%s" % (t["slot"], t["font"])
                                  for t in report["bad_theme_fonts"]))
        lines.append(u"    （主题是继承来源之一，但本工具**不改主题**；它会影响所有靠主题取字体的"
                     u"文字，要改请手工改或把该字体加进 keep）")
    if report["declared"]:
        grouped = collections.Counter((d["where"], d["font"]) for d in report["declared"])
        lines.append(u"")
        lines.append(u"提示（当前不显示成文字，但留着旧字体）：")
        for (where, font), count in sorted(grouped.items()):
            lines.append(u"    %s 里的 %s（%d 处）" % (where, font, count))
    lines.append(u"")
    lines.append(u"AUDIT=%s%s" % (report["verdict"],
                                  u"" if report["verdict"] == "PASS"
                                  else u" " + u"；".join(report["reasons"])))
    return u"\n".join(lines)
