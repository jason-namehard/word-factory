# -*- coding: utf-8 -*-
"""体检（audit）单测。

体检的职责是**证伪**："正式版 = 通体黑 + 没有不合格的字体"这句话由它来判。
所以这里的用例主要盯"容易漏判的继承路径"，而不是照实现抄一遍：

* 文字自己的字体是从**字符样式**继承来的（run 没写 rFonts）→ 必须判 FAIL；
* 从**段落样式**继承来的 → 必须判 FAIL；
* 合格文档 → PASS；不合格的 → 末行 `AUDIT=FAIL` 且原因点名；
* 提示档（段落标记、样式定义、字体清单）不算 FAIL，但必须报出来。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.audit import audit, format_audit
from wordfactory.document import Document
from wordfactory.fonts import DEFAULT_FONTS, FONT_PARTS, FontRuleSet, normalize
from wordfactory.ooxml import qn

from . import fixtures
from .test_fonts import FANG, SONG, numbering_xml, styles_xml


def clean_styles():
    return styles_xml(doc_defaults=SONG, char_style=SONG)


class AuditCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_audit_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body, extra=None):
        parts = {"word/styles.xml": clean_styles()}
        parts.update(extra or {})
        return fixtures.write_fixture(self.path, body=body, extra=parts)

    def normalize(self, path, out):
        with Document(path, writable_parts=FONT_PARTS) as doc:
            normalize(doc, FontRuleSet(DEFAULT_FONTS))
            return doc.save(out)


class TestInheritedFonts(AuditCase):
    def test_a_font_inherited_from_a_character_style_is_caught(self):
        """实测踩到的就是这一条：run 自己没写字体，字体来自字符样式 26（仿宋）。"""
        body = fixtures.paragraph(fixtures.run(u"水位～库容", rstyle="26"))
        path = self.build(body, extra={"word/styles.xml": styles_xml()})
        report = audit(path)
        self.assertEqual(report["verdict"], "FAIL")
        self.assertIn(FANG, report["bad_effective"])
        self.assertEqual(report["bad_effective"][FANG][0]["where"], u"字符样式 font31")
        self.assertIn(u"不合格字体", u"".join(report["reasons"]))

    def test_the_same_document_passes_after_the_formal_mode(self):
        body = fixtures.paragraph(fixtures.run(u"水位～库容", rstyle="26"),
                                  fixtures.run(u"普通文字"))
        path = self.build(body, extra={"word/styles.xml": styles_xml()})
        out = self.normalize(path, os.path.join(self.dir, u"出.docx"))
        report = audit(out)
        self.assertEqual(report["verdict"], "PASS", report["reasons"])
        self.assertEqual(report["bad_effective"], {})

    def test_a_font_inherited_from_a_paragraph_style_is_caught(self):
        styles = clean_styles().replace(
            u'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
            u'<w:style w:type="paragraph" w:styleId="BodyText2"><w:name w:val="Body Text 2"/>'
            u'<w:rPr><w:rFonts w:eastAsia="%s"/></w:rPr></w:style>' % FANG)
        body = fixtures.paragraph(fixtures.run(u"正文段落"), style="BodyText2")
        report = audit(self.build(body, extra={"word/styles.xml": styles}))
        self.assertEqual(report["verdict"], "FAIL")
        self.assertIn(FANG, report["bad_effective"])

    def test_inheritance_chain_is_resolved_through_basedOn(self):
        """样式继承：BodyText2 → base（base 才是仿宋）。链子要跟到底。"""
        styles = clean_styles().replace(
            u'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
            u'<w:style w:type="paragraph" w:styleId="base1"><w:name w:val="base1"/>'
            u'<w:rPr><w:rFonts w:eastAsia="%s"/></w:rPr></w:style>'
            u'<w:style w:type="paragraph" w:styleId="BodyText2"><w:name w:val="Body Text 2"/>'
            u'<w:basedOn w:val="base1"/></w:style>' % FANG)
        body = fixtures.paragraph(fixtures.run(u"正文段落"), style="BodyText2")
        report = audit(self.build(body, extra={"word/styles.xml": styles}))
        self.assertEqual(report["verdict"], "FAIL")
        self.assertEqual(report["bad_effective"][FANG][0]["where"], u"段落样式 base1")


class TestColorAndHighlight(AuditCase):
    def test_a_non_black_color_fails(self):
        body = fixtures.paragraph(fixtures.run(u"红字", color="FF0000"))
        report = audit(self.build(body))
        self.assertEqual(report["verdict"], "FAIL")
        self.assertEqual(report["bad_colors"][0]["value"], "FF0000")

    def test_a_highlight_fails_but_highlight_none_does_not(self):
        body = fixtures.paragraph(fixtures.run(u"高亮", highlight="yellow"))
        # 段落标记上写 highlight val="none" 是"无高亮"的意思，不该被判成高亮
        body += u'<w:p><w:pPr><w:rPr><w:highlight w:val="none"/></w:rPr></w:pPr>'
        body += u'<w:r><w:t>段落标记的 no-op</w:t></w:r></w:p>'
        report = audit(self.build(body))
        self.assertEqual(report["verdict"], "FAIL")
        self.assertEqual(len(report["highlights"]), 1)
        self.assertEqual(report["highlights"][0]["value"], "yellow")

    def test_formal_mode_clears_both(self):
        body = fixtures.paragraph(fixtures.run(u"红字高亮", color="FF0000", highlight="yellow"))
        out = self.normalize(self.build(body), os.path.join(self.dir, u"出.docx"))
        report = audit(out)
        self.assertEqual(report["verdict"], "PASS", report["reasons"])


class TestHintsVersusFailures(AuditCase):
    def test_a_cleaned_document_is_pass_and_says_so(self):
        path = self.build(fixtures.paragraph(fixtures.run(u"干净文字")))
        text = format_audit(audit(path))
        self.assertTrue(text.endswith(u"AUDIT=PASS"), text)
        self.assertIn(u"✓ 没有非黑颜色", text)

    def test_a_leftover_font_in_a_style_is_a_hint_not_a_failure(self):
        """样式表里留着坏字体、但**没有文字用它** → 只提示，不算 FAIL。"""
        body = fixtures.paragraph(fixtures.run(u"正文"))
        path = self.build(body, extra={"word/styles.xml": styles_xml()})
        report = audit(path)
        self.assertEqual(report["verdict"], "PASS", report["reasons"])
        self.assertTrue([d for d in report["declared"] if u"styles.xml" in d["where"]])
        self.assertIn(u"提示", format_audit(report))

    def test_a_paragraph_mark_font_is_disclosed(self):
        body = fixtures.paragraph(fixtures.run(u"正文"),
                                 mark_fonts={"eastAsia": FANG})
        report = audit(self.build(body))
        self.assertEqual(report["verdict"], "PASS", u"段落标记不显示成文字 → 只提示")
        self.assertEqual([d["where"] for d in report["declared"]],
                         [u"word/document.xml（段落标记）"])

    def test_the_font_table_is_disclosed_but_never_a_failure(self):
        extra = {"word/fontTable.xml": u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="%s"><w:font w:name="%s"><w:charset w:val="86"/></w:font>
<w:font w:name="%s"><w:charset w:val="86"/></w:font></w:fonts>"""
                 % (fixtures.W, FANG, SONG)}
        report = audit(self.build(fixtures.paragraph(fixtures.run(u"正文")), extra=extra))
        self.assertEqual(report["verdict"], "PASS")
        self.assertTrue([d for d in report["declared"] if u"fontTable" in d["where"]])


if __name__ == "__main__":
    unittest.main()
