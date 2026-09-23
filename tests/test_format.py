# -*- coding: utf-8 -*-
"""宏「格式规范化」（`ops/normalize.py` + `format` 命令）的单测。

它就是把已有步骤**串成一条固定顺序的命令**：先上下标规则（会切 run）→ 再字体与颜色。
判据：

* 两步都真的生效（上标落在该上标的字上、字体换成规则里的、颜色统一成黑）；
* `--dry-run` 一个字节都不写、树也不动；
* 该重写的部件都登记（正文 + 样式表 + 编号表），**别的部件不许动**；
* 幂等（第二遍 0 处）。
"""

import os
import shutil
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from wordfactory.document import Document
from wordfactory.fonts import DEFAULT_FONTS, FontRuleSet
from wordfactory.ooxml import qn
from wordfactory.ops import normalize as normalize_op
from wordfactory.rules import RuleSet
from wordfactory.text import Paragraph

from . import fixtures

SUBSCRIPT_RULES = RuleSet({"rules": [
    {"id": "m2", "match": u"m2", "kinds": "NS"},
    {"id": "cm2", "match": u"cm2", "kinds": "NNS"},
]})
FONT_RULES = FontRuleSet(DEFAULT_FONTS)


class FormatCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_format_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def run_op(self, body, dry_run=False, superscripts=True):
        fixtures.write_fixture(self.path, body=body)
        with Document(self.path, writable_parts=("word/document.xml", "word/styles.xml",
                                                "word/numbering.xml")) as doc:
            report = normalize_op.format_normalize(
                doc, FONT_RULES, SUBSCRIPT_RULES if superscripts else None, dry_run=dry_run)
            out = None
            if report["total"] and not dry_run:
                out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, out

    def parts(self, path):
        with zipfile.ZipFile(path) as archive:
            return dict((name, archive.read(name)) for name in archive.namelist())


class TestBothStepsRun(FormatCase):
    def test_superscripts_and_fonts_both_apply(self):
        body = fixtures.paragraph(
            fixtures.run(u"面积 12m2"), fixtures.run(u"水位 3.5m", color="FF0000"))
        report, out = self.run_op(body)
        steps = dict((step["step"], step) for step in report["steps"])
        self.assertEqual(steps[u"上下标规则"]["total"], 1, u"m2 的 2 要上标")
        # 红字那个 run 变黑；另外上标规则把 `12m2` 切成了两个 run，新 run 没有颜色也是补黑
        self.assertGreaterEqual(steps[u"字体与颜色"]["colors"], 1)
        self.assertGreater(steps[u"字体与颜色"]["runs"], 1)
        with Document(out) as doc:
            root = doc.part()
        verticals = [node.get(qn("w:val")) for node in root.iter(qn("w:vertAlign"))]
        self.assertIn("superscript", verticals)
        colors = [node.get(qn("w:val")) for node in root.iter(qn("w:color"))]
        self.assertNotIn("FF0000", colors, u"红字要被改掉")
        self.assertIn("000000", colors)

    def test_superscripts_can_be_skipped(self):
        body = fixtures.paragraph(fixtures.run(u"面积 12m2"))
        report, out = self.run_op(body, superscripts=False)
        steps = dict((step["step"], step) for step in report["steps"])
        self.assertNotIn(u"上下标规则", steps)
        with Document(out) as doc:
            self.assertEqual(len(list(doc.part().iter(qn("w:vertAlign")))), 0)

    def test_the_right_parts_are_rewritten(self):
        body = fixtures.paragraph(fixtures.run(u"正文 仿宋", color="FF0000"))
        report, out = self.run_op(body)
        before, after = self.parts(self.path), self.parts(out)
        changed = sorted(name for name in after if after[name] != before.get(name))
        # 报告里声明的部件必须与实际被改写的部件一致（不许"说改了却没说清"）
        self.assertEqual(changed, sorted(report["parts"]))
        self.assertIn("word/document.xml", changed)
        # 夹具的样式表里没有需要换的字体、也没有 numbering.xml → 它们不该被动
        self.assertNotIn("word/styles.xml", changed)
        self.assertEqual(before["customXml/item1.xml"], after["customXml/item1.xml"],
                         u"无关部件必须逐字节不变")

    def test_dry_run_touches_nothing(self):
        body = fixtures.paragraph(fixtures.run(u"面积 12m2"), fixtures.run(u"红字", color="FF0000"))
        fixtures.write_fixture(self.path, body=body)
        with Document(self.path, writable_parts=("word/document.xml", "word/styles.xml")) as doc:
            before = ET.tostring(doc.part(), encoding="utf-8")
            report = normalize_op.format_normalize(doc, FONT_RULES, SUBSCRIPT_RULES, dry_run=True)
            after = ET.tostring(doc.part(), encoding="utf-8")
            self.assertEqual(doc.dirty_parts, [])
        self.assertEqual(before, after, u"dry-run 连树都不该动")
        self.assertGreater(report["total"], 0, u"但要报出会改多少")

    def test_running_twice_changes_nothing(self):
        body = fixtures.paragraph(fixtures.run(u"面积 12m2"), fixtures.run(u"红字", color="FF0000"))
        fixtures.write_fixture(self.path, body=body)
        with Document(self.path, writable_parts=("word/document.xml", "word/styles.xml",
                                                "word/numbering.xml")) as doc:
            first = normalize_op.format_normalize(doc, FONT_RULES, SUBSCRIPT_RULES)
            second = normalize_op.format_normalize(doc, FONT_RULES, SUBSCRIPT_RULES)
        self.assertGreater(first["total"], 0)
        self.assertEqual(second["total"], 0, u"第二遍必须是 0（两步都是幂等的）")

    def test_the_cli_runs_and_reports(self):
        fixtures.write_fixture(self.path, body=fixtures.paragraph(fixtures.run(u"面积 12m2")))
        from wordfactory.cli import main
        out = os.path.join(self.dir, u"从命令来.docx")
        self.assertEqual(main(["format", self.path, "--out", out]), 0)
        with Document(out) as doc:
            self.assertIn("superscript",
                          [node.get(qn("w:val")) for node in doc.part().iter(qn("w:vertAlign"))])


if __name__ == "__main__":
    unittest.main()
