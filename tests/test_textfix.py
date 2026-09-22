# -*- coding: utf-8 -*-
"""宏「规划报告一键宏」文本部分的单测（`ops/textfix.py`）。

判据来自参考宏 `规划报告一键宏.bas`（§2.12）与 OOXML 语义：

* 替换 `其它`→`其他`、`东流流经`→`向东流经`，**作用范围是全文**（含表格单元格、文本框），
  且**必须跨 run**（Word 的 Find 能跨 run，按 run 逐个替换会漏）；
* 段落对齐：两端对齐（`both`/`distribute`）→ 左对齐；**必须解析样式继承** ——
  实测用户的报告里 `Normal` 样式写着 `jc=both`，只看段落自己的属性会漏掉 18 段；
* 幂等（第二遍 0 处）、`dry-run` 零写入、只重写改过的部件。
"""

import os
import shutil
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.ops import textfix
from wordfactory.text import Paragraph

from . import fixtures

STYLES_WITH_JUSTIFY = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
  <w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="LeftStyle">
    <w:name w:val="left style"/><w:basedOn w:val="Normal"/>
    <w:pPr><w:jc w:val="left"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="InheritsJustify">
    <w:name w:val="inherits justify"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>""" % fixtures.W


class TextFixCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_textfix_")
        self.path = os.path.join(self.dir, u"夹具.docx")
        self.rules = textfix.ReplacementRuleSet(textfix.DEFAULT_REPLACEMENTS)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body, styles=None):
        extra = {"word/styles.xml": styles} if styles else None
        fixtures.write_fixture(self.path, body=body, extra=extra)
        return self.path

    def run_op(self, body, styles=None, **kwargs):
        self.build(body, styles)
        with Document(self.path, writable_parts=("word/document.xml", "word/styles.xml")) as doc:
            report = textfix.apply(doc, self.rules, **kwargs)
            out = None
            if report["total"] and not kwargs.get("dry_run"):
                out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, out

    def body_texts(self, path):
        with Document(path) as doc:
            return [Paragraph(element).text for element in doc.part().iter(qn("w:p"))]

    def jc_of(self, path, index=0):
        with Document(path) as doc:
            element = list(doc.part().iter(qn("w:p")))[index]
        pr = element.find(qn("w:pPr"))
        jc = pr.find(qn("w:jc")) if pr is not None else None
        return jc.get(qn("w:val")) if jc is not None else None


class TestReplacement(TextFixCase):
    def test_a_simple_replacement(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"其它情况")))
        self.assertEqual(report["replaced"], 1)
        self.assertEqual(self.body_texts(out), [u"其他情况"])

    def test_a_replacement_across_runs(self):
        """Word 的 Find 能跨 run；按 run 逐个替换会漏 —— 必须走逻辑文本层。"""
        body = fixtures.paragraph(fixtures.run(u"其"), fixtures.run(u"它"), fixtures.run(u"事项"))
        report, out = self.run_op(body)
        self.assertEqual(report["replaced"], 1)
        self.assertEqual(self.body_texts(out), [u"其他事项"])

    def test_replacements_inside_table_cells(self):
        """参考宏作用在 `ActiveDocument.Content` —— 表格里的文字也算。"""
        body = fixtures.table([[u"<w:r><w:t>其它</w:t></w:r>", u"<w:r><w:t>值</w:t></w:r>"]])
        report, out = self.run_op(body)
        self.assertEqual(report["replaced"], 1)
        self.assertIn(u"其他", u"".join(self.body_texts(out)))

    def test_the_replacement_is_idempotent(self):
        self.build(fixtures.paragraph(fixtures.run(u"其它")))
        with Document(self.path, writable_parts=("word/document.xml",)) as doc:
            first = textfix.apply(doc, self.rules)
            second = textfix.apply(doc, self.rules)
        self.assertEqual(first["total"], 1)
        self.assertEqual(second["total"], 0, u"第二遍不该再改（文本层替换也是幂等的）")

    def test_multiple_hits_in_one_paragraph_are_all_replaced(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"其它其它其它")))
        self.assertEqual(report["replaced"], 3)
        self.assertEqual(self.body_texts(out), [u"其他其他其他"])


class TestAlignment(TextFixCase):
    def test_justify_becomes_left(self):
        body = u'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>两端</w:t></w:r></w:p>'
        report, out = self.run_op(body)
        self.assertEqual(report["align_fixed"], 1)
        self.assertEqual(self.jc_of(out), "left")

    def test_distribute_also_becomes_left(self):
        body = u'<w:p><w:pPr><w:jc w:val="distribute"/></w:pPr><w:r><w:t>分散</w:t></w:r></w:p>'
        report, out = self.run_op(body)
        self.assertEqual(report["align_fixed"], 1)
        self.assertEqual(self.jc_of(out), "left")

    def test_the_alignment_inherited_from_a_style_is_caught(self):
        """**这条是重点**：`Normal` 里写着两端对齐，段落自己没写 jc —— 也得改。"""
        body = fixtures.paragraph(fixtures.run(u"继承来的两端对齐"))
        report, out = self.run_op(body, styles=STYLES_WITH_JUSTIFY)
        self.assertEqual(report["align_fixed"], 1)
        self.assertEqual(report["align_by_source"], {u"样式 Normal": 1})
        self.assertEqual(self.jc_of(out), "left")

    def test_the_inheritance_chain_is_followed_through_basedOn(self):
        """样式 `InheritsJustify` → `Normal`（both）→ 也是两端对齐。"""
        body = fixtures.paragraph(fixtures.run(u"更深一层"), style="InheritsJustify")
        report, out = self.run_op(body, styles=STYLES_WITH_JUSTIFY)
        self.assertEqual(report["align_fixed"], 1)
        self.assertEqual(self.jc_of(out), "left")

    def test_a_style_that_overrides_to_left_is_left_alone(self):
        body = fixtures.paragraph(fixtures.run(u"已经是左对齐"), style="LeftStyle")
        report, out = self.run_op(body, styles=STYLES_WITH_JUSTIFY)
        self.assertEqual(report["align_fixed"], 0)
        self.assertEqual(report["total"], 0)

    def test_center_and_left_paragraphs_are_untouched(self):
        body = (u'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>居中</w:t></w:r></w:p>'
                u'<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>左</w:t></w:r></w:p>')
        report, _ = self.run_op(body)
        self.assertEqual(report["total"], 0, u"居中和左对齐都不该动")

    def test_no_align_switch_skips_the_alignment(self):
        body = u'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>两端</w:t></w:r></w:p>'
        report, out = self.run_op(body, fix_align=False)
        self.assertEqual(report["align_fixed"], 0)
        self.assertIsNone(out)

    def test_fix_styles_can_change_the_style_itself(self):
        body = fixtures.paragraph(fixtures.run(u"正文"))
        self.build(body, STYLES_WITH_JUSTIFY)
        with Document(self.path,
                      writable_parts=("word/document.xml", "word/styles.xml")) as doc:
            report = textfix.apply(doc, self.rules, fix_styles=True)
            out = doc.save(os.path.join(self.dir, u"出.docx"))
        self.assertTrue([key for key in report["changes"] if u"样式对齐" in key],
                        report["changes"])
        with Document(out) as doc:
            style = [s for s in doc.part("word/styles.xml").iter(qn("w:style"))
                     if s.get(qn("w:styleId")) == "Normal"][0]
            jc = style.find(qn("w:pPr")).find(qn("w:jc"))
        self.assertEqual(jc.get(qn("w:val")), "left")

    def test_only_the_changed_parts_are_rewritten(self):
        body = u'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>两端</w:t></w:r></w:p>'
        self.build(body, STYLES_WITH_JUSTIFY)
        before = self.parts(self.path)
        with Document(self.path, writable_parts=("word/document.xml",)) as doc:
            textfix.apply(doc, self.rules)
            out = doc.save(os.path.join(self.dir, u"出.docx"))
        after = self.parts(out)
        changed = sorted(name for name in after if after[name] != before[name])
        self.assertEqual(changed, ["word/document.xml"], u"没动 styles 就不该重写它")

    def parts(self, path):
        with zipfile.ZipFile(path) as archive:
            return dict((name, archive.read(name)) for name in archive.namelist())


class TestDryRunAndRules(TextFixCase):
    def test_dry_run_writes_nothing(self):
        body = (fixtures.paragraph(fixtures.run(u"其它"))
                + u'<w:p><w:pPr><w:jc w:val="both"/></w:pPr><w:r><w:t>两端</w:t></w:r></w:p>')
        self.build(body)
        with Document(self.path, writable_parts=("word/document.xml",)) as doc:
            report = textfix.apply(doc, self.rules, dry_run=True)
            self.assertEqual(doc.dirty_parts, [])
            self.assertEqual(self.body_texts(self.path), [u"其它", u"两端"])
        self.assertEqual(report["total"], 2, u"dry-run 要报出「会改多少」")

    def test_rule_problems_are_reported(self):
        rules = textfix.ReplacementRuleSet({"replacements": [
            {"from": u"甲"}, {"from": u"乙", "to": u"乙"}, {"from": u"丙", "to": u"丁", "x": 1}],
            "align": {"from": ["middle"], "to": "left"}})
        problems = rules.check()
        self.assertTrue([p for p in problems if u"缺 'to'" in p], problems)
        self.assertTrue([p for p in problems if u"一样" in p], problems)
        self.assertTrue([p for p in problems if u"不认识的键" in p], problems)
        self.assertTrue([p for p in problems if u"middle" in p], problems)

    def test_the_rules_file_round_trips(self):
        path = os.path.join(self.dir, u"replacements.json")
        self.rules.save(path)
        again = textfix.ReplacementRuleSet.load(path)
        self.assertEqual([(i["from"], i["to"]) for i in again.replacements],
                         [(u"其它", u"其他"), (u"东流流经", u"向东流经")])
        self.assertEqual(again.align_to, "left")

    def test_the_cli_runs(self):
        body = fixtures.paragraph(fixtures.run(u"其它情况"))
        self.build(body)
        from wordfactory.cli import main
        out = os.path.join(self.dir, u"从命令来.docx")
        code = main(["textfix", self.path, "--out", out])
        self.assertEqual(code, 0)
        self.assertEqual(self.body_texts(out), [u"其他情况"])


if __name__ == "__main__":
    unittest.main()
