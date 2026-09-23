# -*- coding: utf-8 -*-
"""一键清理 `ops/tidy.py` 的单测。

口径（用户 2026-09-22 要的"一键去空格 / 去无意义空白行"，默认保守）：
* 默认只动三件明显无意义的事：段尾空格、连续空白段、文档首尾空白段；
* 段首空格与段内连续空格**默认不动**（前者可能是缩进、后者会把表题的空格居中压掉）；
* 段内连续空格一旦开了，**必须跳过题注段落**；
* `--dry-run` 一个字节都不许动（树也不许动）。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.ops import tidy as tidy_op
from wordfactory.text import Paragraph

from . import fixtures


class TidyCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_tidy_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body):
        fixtures.write_fixture(self.path, body=body)
        return self.path

    def run_op(self, body, **options):
        self.build(body)
        with Document(self.path) as doc:
            report = tidy_op.tidy(doc, options, dry_run=options.pop("_dry", False))
            out = None
            if report["total"] and not options.get("_dry"):
                out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, out

    def texts(self, path=None):
        with Document(path or self.path) as doc:
            return [Paragraph(element).text for element in doc.body()
                    if element.tag == qn("w:p")]

    def all_texts(self, path=None):
        """全文（含表格里的段落）—— 判断"表格有没有被碰"要用它。"""
        with Document(path or self.path) as doc:
            return u"".join(Paragraph(element).text
                            for element in doc.part().iter(qn("w:p")))


class TestSpaces(TidyCase):
    def test_trailing_spaces_go_away(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"正文末尾带空格   ")))
        self.assertEqual(report["changes"]["段尾空格"], 1)
        self.assertEqual(self.texts(out), [u"正文末尾带空格"])

    def test_full_width_and_nbsp_trailing_spaces_too(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"末尾\u3000\u00a0")))
        self.assertEqual(report["changes"]["段尾空格"], 1)
        self.assertEqual(self.texts(out), [u"末尾"])

    def test_leading_spaces_are_kept_by_default(self):
        """段首空格可能是"用空格做的缩进"，默认不碰。"""
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"  首行缩进")))
        self.assertEqual(report["total"], 0)
        self.assertIsNone(out)

    def test_leading_spaces_go_away_when_asked(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"  首行缩进")),
                                  trim_leading=True)
        self.assertEqual(report["changes"]["段首空格"], 1)
        self.assertEqual(self.texts(out), [u"首行缩进"])

    def test_space_runs_are_kept_unless_asked(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"A    B")))
        self.assertEqual(report["total"], 0)

    def test_space_runs_collapse_when_asked_but_captions_are_skipped(self):
        body = (fixtures.paragraph(fixtures.run(u"普通段落 A     B"))
                + fixtures.paragraph(fixtures.run(u"表2.3-1        库容特性表")))
        report, out = self.run_op(body, collapse_space_runs=True)
        self.assertEqual(report["changes"]["连续空格压缩"], 1)
        self.assertEqual(report["changes"]["题注段落（跳过）"], 1)
        self.assertEqual(self.texts(out),
                         [u"普通段落 A B", u"表2.3-1        库容特性表"])


class TestBlankLines(TidyCase):
    def test_consecutive_blank_paragraphs_collapse_to_one(self):
        body = (fixtures.paragraph(fixtures.run(u"第一段")) + u"<w:p/>" + u"<w:p/>"
                + u"<w:p/>" + fixtures.paragraph(fixtures.run(u"第二段")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["空白段压缩"], 2, u"三个空段留一个")
        self.assertEqual(self.texts(out), [u"第一段", u"", u"第二段"])

    def test_a_single_blank_line_is_a_real_gap_and_stays(self):
        body = (fixtures.paragraph(fixtures.run(u"上")) + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"下")))
        report, out = self.run_op(body)
        self.assertEqual(report["total"], 0, u"单独一个空行是「想留的」，不动")

    def test_blank_paragraphs_at_the_edges_go_away(self):
        body = (u"<w:p/>" + u"<w:p/>" + fixtures.paragraph(fixtures.run(u"正文"))
                + u"<w:p/>")
        report, out = self.run_op(body)
        self.assertGreaterEqual(report["changes"]["首尾空白段删除"], 2)
        self.assertEqual(self.texts(out), [u"正文"])

    def test_whitespace_only_paragraphs_count_as_blank(self):
        body = (fixtures.paragraph(fixtures.run(u"甲"))
                + fixtures.paragraph(fixtures.run(u"   "))
                + fixtures.paragraph(fixtures.run(u"乙")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["段尾空格"], 1, u"全空白段落会被清成真空段")
        self.assertEqual(self.texts(out), [u"甲", u"", u"乙"],
                         u"清完它就是一个空行 —— 单个空行是「想留的」，保留")


class TestScopeAndSafety(TidyCase):
    def test_table_paragraphs_are_left_alone_by_default(self):
        body = (fixtures.table([[u'<w:r><w:t>表内带空格   </w:t></w:r>',
                                u'<w:r><w:t>值</w:t></w:r>']])
                + fixtures.paragraph(fixtures.run(u"正文带空格   ")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["段尾空格"], 1, u"默认只动正文段落")
        self.assertIn(u"表内带空格   ", self.all_texts(out), u"表格里的空格默认不碰")

    def test_scope_all_reaches_into_tables(self):
        body = fixtures.table([[u'<w:r><w:t>表内带空格   </w:t></w:r>',
                                u'<w:r><w:t>值</w:t></w:r>']])
        report, out = self.run_op(body, scope="all")
        self.assertEqual(report["changes"]["段尾空格"], 1)
        self.assertNotIn(u"表内带空格   ", self.all_texts(out), u"--scope all 才伸进表格")

    def test_dry_run_touches_nothing(self):
        body = (fixtures.paragraph(fixtures.run(u"末尾  ")) + u"<w:p/>" + u"<w:p/>")
        self.build(body)
        with Document(self.path) as doc:
            before = [Paragraph(element).text for element in doc.body()
                      if element.tag == qn("w:p")]
            report = tidy_op.tidy(doc, dry_run=True)
            after = [Paragraph(element).text for element in doc.body()
                     if element.tag == qn("w:p")]
            self.assertEqual(doc.dirty_parts, [])
        self.assertEqual(before, after, u"dry-run 连树都不该动")
        self.assertGreater(report["total"], 0, u"但要把「会改多少」报出来")

    def test_running_twice_changes_nothing(self):
        body = (fixtures.paragraph(fixtures.run(u"末尾  ")) + u"<w:p/>" + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"正文")))
        self.build(body)
        with Document(self.path) as doc:
            first = tidy_op.tidy(doc)
            second = tidy_op.tidy(doc)
        self.assertGreater(first["total"], 0)
        self.assertEqual(second["total"], 0, u"第二遍必须是 0 处（幂等）")

    def test_the_cli_runs(self):
        body = fixtures.paragraph(fixtures.run(u"末尾  "))
        self.build(body)
        from wordfactory.cli import main
        out = os.path.join(self.dir, u"从命令来.docx")
        self.assertEqual(main(["tidy", self.path, "--out", out]), 0)
        self.assertEqual(self.texts(out), [u"末尾"])


if __name__ == "__main__":
    unittest.main()


class TestCopyPlusPlusSample(TidyCase):
    """**金标准**：用户 2026-09-22 给的 `copy++演示.txt`（Copy++ 开「合并换行」+「去除空格」）。

    前后对照在 `tests/copypp_sample.py` 里（从用户文件逐字抽出）：

    * 「合并换行」= **删掉空行**（不是"压成一个"）→ 被空行隔开的行变成连续行；
    * 「去除空格」= **删掉所有半角空格**（`工程经验 + 客户对接` → `工程经验+客户对接`、
      `Word Factory` → `WordFactory`）。
    """

    def run_sample(self, options):
        from tests.copypp_sample import BEFORE
        body = u"".join(fixtures.paragraph(fixtures.run(line)) if line.strip() else u"<w:p/>"
                        for line in BEFORE)
        report, out = self.run_op(body, **options)
        return report, out

    def test_merge_lines_and_remove_spaces_reproduce_the_sample(self):
        from tests.copypp_sample import AFTER
        report, out = self.run_sample({"merge_lines": True, "remove_spaces": True})
        self.assertEqual(self.texts(out), AFTER, u"必须与 Copy++ 的结果逐行一致")
        self.assertEqual(report["changes"]["删除空行"], 13)
        self.assertGreater(report["changes"]["去空格"], 0)

    def test_merge_lines_alone_removes_blank_lines(self):
        from tests.copypp_sample import BEFORE
        report, out = self.run_sample({"merge_lines": True})
        expected = [line for line in BEFORE if line.strip()]
        self.assertEqual(self.texts(out), expected)

    def test_remove_spaces_alone_keeps_blank_lines_but_strips_spaces(self):
        from tests.copypp_sample import BEFORE
        report, out = self.run_sample({"remove_spaces": True, "blank_lines": False})
        expected = []
        for line in BEFORE:
            if line.strip():
                expected.append(line.replace(u" ", u""))
            else:
                expected.append(u"")
        while expected and expected[-1] == u"":
            expected.pop()
        self.assertEqual(self.texts(out), expected)

    def test_the_two_modes_are_off_by_default(self):
        """默认口径不变（保守）—— 这两个是"照 Copy++ 行为"的显式开关。"""
        from wordfactory.ops.tidy import DEFAULT_TIDY
        self.assertFalse(DEFAULT_TIDY["merge_lines"])
        self.assertFalse(DEFAULT_TIDY["remove_spaces"])

    def test_remove_spaces_skips_caption_paragraphs_by_default(self):
        """题注的空格是"空格居中"用的 —— 默认整段跳过（否则一跑就把表题压平）。"""
        body = (fixtures.paragraph(fixtures.run(u"正文 A B"))
                + fixtures.paragraph(fixtures.run(u"表2.3-1        库容特性表")))
        report, out = self.run_op(body, remove_spaces=True)
        self.assertEqual(report["changes"]["去空格"], 2, u"数的是**空格字符**（A 与 B 之间两个空格）")
        self.assertEqual(report["changes"]["题注段落（跳过）"], 1)
        self.assertEqual(self.texts(out), [u"正文AB", u"表2.3-1        库容特性表"])

    def test_include_captions_can_turn_the_skip_off(self):
        body = fixtures.paragraph(fixtures.run(u"表2.3-1    库容表"))
        report, out = self.run_op(body, remove_spaces=True, caption_skip=False)
        self.assertGreater(report["changes"]["去空格"], 0)
        self.assertEqual(self.texts(out), [u"表2.3-1库容表"])

    def test_merge_lines_and_the_conservative_mode_do_not_double_count(self):
        """两种口径互斥：merge_lines 时不再做"压成一个"的保守处理（否则两边都在数）。"""
        body = (fixtures.paragraph(fixtures.run(u"甲")) + u"<w:p/>" + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"乙")))
        report, out = self.run_op(body, merge_lines=True)
        self.assertEqual(report["changes"].get("空白段压缩"), None)
        self.assertEqual(report["changes"]["删除空行"], 2)
        self.assertEqual(self.texts(out), [u"甲", u"乙"])

    def test_the_full_width_space_is_kept_unless_asked(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"甲乙\u3000丙 丁")),
                                  remove_spaces=True)
        self.assertEqual(self.texts(out), [u"甲乙\u3000丙丁"], u"全角空格默认留着（可能是缩进）")
