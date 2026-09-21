# -*- coding: utf-8 -*-
"""字体规范化（正式版的后半程）单测。

这些用例都是**实测踩出来的缺口**，不是照着实现抄的：

* `w:rFonts` 有四个属性，只改 ascii/eastAsia 会留下 hAnsi/cs 的旧字体；
* 段落标记（`w:pPr/w:rPr/w:rFonts`）也是字体，也得改；
* 样式表/编号表是"继承来源"，不改它们，那些 run 自己不写字体的文字照样是旧字体；
* 没有文字的 run 也要改（否则残留计数对不上，实测漏了 7 处）；
* **`--dry-run` 一个字节都不许动**（连"补一个空 rPr"都不行）；
* 幂等：跑第二遍报 0 处。
"""

import os
import shutil
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from wordfactory.document import Document
from wordfactory.fonts import DEFAULT_FONTS, FONT_PARTS, FontRuleSet, normalize
from wordfactory.ooxml import qn

from . import fixtures

FANG = u"\u4eff\u5b8b"        # 仿宋
SONG = u"\u5b8b\u4f53"        # 宋体


def styles_xml(doc_defaults=SONG, char_style=FANG):
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="%s"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="26"><w:name w:val="font31"/>
    <w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/><w:color w:val="000000"/></w:rPr></w:style>
</w:styles>""" % (fixtures.W, doc_defaults, char_style, char_style, char_style, char_style)


def numbering_xml(font=FANG):
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="%s">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:rPr><w:rFonts w:eastAsia="%s"/></w:rPr></w:lvl></w:abstractNum>
</w:numbering>""" % (fixtures.W, font)


class FontCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_fonts_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body, extra=None, dirty=True):
        """``dirty=False`` 时装一份"本来就合规"的样式表/编号表，方便断言精确的改动清单。"""
        parts = {"word/styles.xml": styles_xml() if dirty else styles_xml(char_style=SONG),
                 "word/numbering.xml": numbering_xml() if dirty else numbering_xml(SONG)}
        parts.update(extra or {})
        return fixtures.write_fixture(self.path, body=body, extra=parts)

    def parts_of(self, path):
        with zipfile.ZipFile(path) as archive:
            return dict((name, archive.read(name)) for name in archive.namelist())

    def fonts_in(self, element):
        node = element.find(qn("w:rPr"))
        node = node.find(qn("w:rFonts")) if node is not None else None
        if node is None:
            return {}
        return dict((key, node.get(qn(key))) for key in
                    ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs") if node.get(qn(key)))

    def normalize(self, path, out=None, **kwargs):
        out = out or os.path.join(self.dir, u"出.docx")
        package = Document(path, writable_parts=FONT_PARTS)
        report = normalize(package, FontRuleSet(DEFAULT_FONTS), **kwargs)
        if not kwargs.get("dry_run"):
            package.save(out)
        package.close()
        return report, out

    def open(self, path):
        return Document(path, writable_parts=FONT_PARTS)


class TestFourAttributes(FontCase):
    def test_all_four_attributes_are_replaced(self):
        """四个属性都要换 —— 只换 ascii/eastAsia 的话 hAnsi/cs 会留下旧字体。"""
        body = fixtures.paragraph(fixtures.run(
            u"水位～库容", rfonts={"ascii": FANG, "hAnsi": FANG, "eastAsia": FANG, "cs": FANG}))
        report, out = self.normalize(self.build(body, dirty=False))
        self.assertEqual(sorted(report["fonts"].keys()),
                         [u"%s → %s（%s）" % (FANG, SONG, name)
                          for name in (u"ascii", u"cs", u"eastAsia", u"hAnsi")])
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run),
                             {"w:ascii": SONG, "w:hAnsi": SONG, "w:eastAsia": SONG, "w:cs": SONG},
                             u"四个属性必须全是新字体")

    def test_the_paragraph_mark_font_is_replaced_too(self):
        """段落标记的字体也要换（用户把光标放到段尾时看到的就是它）。"""
        body = fixtures.paragraph(fixtures.run(SONG and u"正文"),
                                 mark_fonts={"ascii": FANG, "eastAsia": FANG})
        _, out = self.normalize(self.build(body))
        with self.open(out) as doc:
            para = list(doc.part().iter(qn("w:p")))[0]
            rpr = para.find(qn("w:pPr")).find(qn("w:rPr"))
            self.assertEqual(rpr.find(qn("w:rFonts")).get(qn("w:eastAsia")), SONG)

    def test_a_run_without_text_is_still_fixed(self):
        """没有文字的 run 也要换 —— 实测就是它让"还有 7 处仿宋"露了出来。"""
        body = fixtures.paragraph(
            fixtures.run(u"有字", rfonts={"eastAsia": SONG}),
            u'<w:r><w:rPr><w:rFonts w:eastAsia="%s"/></w:rPr></w:r>' % FANG)
        report, out = self.normalize(self.build(body))
        self.assertEqual(report["runs"], 1, u"报告的'有文字的 run'只算有字的")
        with self.open(out) as doc:
            fonts = [self.fonts_in(run) for run in doc.part().iter(qn("w:r"))]
            self.assertEqual(fonts[1], {"w:eastAsia": SONG}, u"没文字的 run 也要换")


class TestInheritanceSources(FontCase):
    def test_styles_and_numbering_are_rewritten(self):
        """样式表 / 编号表里的字体是继承来源，必须一起换，否则继承它的文字还是旧字体。"""
        body = fixtures.paragraph(fixtures.run(u"正文"))
        report, out = self.normalize(self.build(body))
        self.assertEqual(report["parts"],
                         ["word/document.xml", "word/numbering.xml", "word/styles.xml"])
        with self.open(out) as doc:
            style = [s for s in doc.part("word/styles.xml").iter(qn("w:style"))
                     if s.get(qn("w:styleId")) == "26"][0]
            fonts = style.find(qn("w:rPr")).find(qn("w:rFonts"))
            self.assertEqual(fonts.get(qn("w:eastAsia")), SONG)
            lvl = doc.part("word/numbering.xml").find(qn("w:abstractNum"))
            self.assertEqual(lvl.iter(qn("w:rFonts")).__next__().get(qn("w:eastAsia")), SONG)

    def test_an_unrelated_part_stays_byte_identical(self):
        """不在白名单里的部件必须**原字节**搬过去（这是"只动该动的"的硬证据）。"""
        body = fixtures.paragraph(fixtures.run(u"正文", rfonts={"eastAsia": FANG}))
        _, out = self.normalize(self.build(body))
        before = self.parts_of(self.path)
        after = self.parts_of(out)
        self.assertEqual(before["customXml/item1.xml"], after["customXml/item1.xml"])
        self.assertEqual(before["word/styles.xml"] == after["word/styles.xml"], False)


class TestDryRunAndIdempotence(FontCase):
    def test_dry_run_does_not_touch_the_document_at_all(self):
        """dry-run 连"补一个空 rPr"都不许做 —— 否则同一棵树保存出来就跟 dry-run 说的不一样。"""
        body = fixtures.paragraph(fixtures.run(u"没有 rPr 的文字"))
        path = self.build(body)
        with self.open(path) as doc:
            before = ET.tostring(doc.part(), encoding="utf-8")
            report = normalize(doc, FontRuleSet(DEFAULT_FONTS), dry_run=True)
            after = ET.tostring(doc.part(), encoding="utf-8")
            dirty = doc.package.dirty_parts
        self.assertEqual(before, after, u"dry-run 改了内存里的 XML")
        self.assertEqual(report["colors"], 1, u"dry-run 要能报出'会补多少黑色'")
        self.assertEqual(dirty, [], u"dry-run 一个部件都不该登记为已改")
        self.assertEqual(report["parts"], ["word/document.xml", "word/numbering.xml",
                                           "word/styles.xml"],
                         u"dry-run 的 parts 是「会动哪些部件」的预告")

    def test_second_run_reports_zero(self):
        body = fixtures.paragraph(fixtures.run(u"正文", rfonts={"eastAsia": FANG},
                                              color="FF0000", highlight="yellow"))
        _, out = self.normalize(self.build(body))
        report, _ = self.normalize(out, out=os.path.join(self.dir, u"出2.docx"))
        self.assertEqual(report["fonts"], {}, u"第二遍不该再有字体要换")
        self.assertEqual(report["colors"], 0, u"第二遍不该再有颜色要改")
        self.assertEqual(report["highlights"], 0, u"第二遍不该再有高亮要去")

    def test_a_kept_font_is_left_alone(self):
        body = fixtures.paragraph(fixtures.run(u"标题", rfonts={"eastAsia": u"黑体"}))
        report, out = self.normalize(self.build(body, dirty=False))
        self.assertEqual([key for key in report["fonts"] if u"黑体" in key], [],
                         u"keep 里的字体不该出现在改动清单里")
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run).get("w:eastAsia"), u"黑体")


class TestWhatDefaultMayTouch(FontCase):
    """default 的边界：只管中文属性、符号字体永不碰（这两条都是实测逼出来的）。"""

    def test_a_chinese_font_outside_the_list_becomes_the_default(self):
        body = fixtures.paragraph(fixtures.run(u"楷体字", rfonts={"eastAsia": u"楷体"}))
        report, out = self.normalize(self.build(body, dirty=False))
        self.assertEqual(report["fonts"], {u"楷体 → 宋体（eastAsia）": 1})
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run).get("w:eastAsia"), SONG)

    def test_a_western_font_is_left_alone_by_default(self):
        """实测：这份文档 w:cs 上有 Tahoma、样式表里有 Arial。西文不许被 default 顺手换掉。"""
        body = fixtures.paragraph(fixtures.run(
            u"Latin", rfonts={"ascii": "Arial", "hAnsi": "Arial", "cs": "Tahoma",
                              "eastAsia": SONG}))
        report, out = self.normalize(self.build(body, dirty=False))
        self.assertEqual(report["fonts"], {}, u"西文属性不归 default 管")
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run).get("w:cs"), "Tahoma")

    def test_default_scope_all_does_replace_western_fonts(self):
        """想连西文一起统一，就把 default_scope 改成 all（它是个开关，不是写死的）。"""
        rule_set = FontRuleSet(dict(DEFAULT_FONTS, default_scope=u"all"))
        body = fixtures.paragraph(fixtures.run(u"Latin", rfonts={"cs": "Tahoma"}))
        path = self.build(body, dirty=False)
        with self.open(path) as doc:
            report = normalize(doc, rule_set)
            out = doc.save(os.path.join(self.dir, u"出.docx"))
        self.assertEqual(report["fonts"], {u"Tahoma → 宋体（cs）": 1})
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run).get("w:cs"), SONG)

    def test_symbol_fonts_are_never_touched_even_with_scope_all(self):
        """符号字体换成宋体 = 掉字形（✔ ➜ ★）。就算 default_scope=all 也不许碰。"""
        rule_set = FontRuleSet(dict(DEFAULT_FONTS, default_scope=u"all"))
        body = fixtures.paragraph(fixtures.run(u"\u2714", rfonts={"ascii": "Wingdings",
                                                                  "eastAsia": SONG}))
        path = self.build(body, dirty=False)
        with self.open(path) as doc:
            report = normalize(doc, rule_set)
            out = doc.save(os.path.join(self.dir, u"出.docx"))
        self.assertEqual(report["fonts"], {}, u"符号字体不该出现在改动清单里")
        with self.open(out) as doc:
            run = list(doc.part().iter(qn("w:r")))[0]
            self.assertEqual(self.fonts_in(run).get("w:ascii"), "Wingdings")

    def test_the_rule_file_flags_a_symbol_font_clash(self):
        problems = FontRuleSet(dict(DEFAULT_FONTS, replace={u"Wingdings": SONG})).check()
        self.assertTrue([p for p in problems if u"打架" in p], problems)


class TestHighlightHonesty(FontCase):
    def test_a_no_op_highlight_is_left_alone(self):
        """``val="none"`` 是「本来就无高亮」。不动它，也**不报**它 —— 数字要跟眼睛看到的一致。"""
        body = fixtures.paragraph(fixtures.run(u"普通文字"),
                                  u'<w:p><w:pPr><w:rPr><w:highlight w:val="none"/></w:rPr></w:pPr>')
        body += u'<w:r><w:t>段落标记的 no-op</w:t></w:r></w:p>'
        report, _ = self.normalize(self.build(body))
        self.assertEqual(report["highlights"], 0, u"空操作不该被算成去过高亮")

    def test_a_real_highlight_is_removed_and_counted_once(self):
        body = fixtures.paragraph(fixtures.run(u"真的高亮", highlight="yellow"))
        report, out = self.normalize(self.build(body))
        self.assertEqual(report["highlights"], 1)
        with self.open(out) as doc:
            self.assertEqual(len(list(doc.part().iter(qn("w:highlight")))), 0)


if __name__ == "__main__":
    unittest.main()
