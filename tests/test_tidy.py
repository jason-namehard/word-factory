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
import zipfile
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
    """空白行口径（**2026-10-08 用户定稿改过一次默认**）。

    旧默认是"连续空段压成一个"（保守）；用户看图指出"还是有空白行"（一段话被一个空行劈成两半）
    → **现在的默认是整段删掉**。"压成一个"退成显式可选档 ``collapse_blank_lines``，
    下面每条都分别钉这两种口径。
    """

    def test_blank_lines_are_deleted_by_default(self):
        body = (fixtures.paragraph(fixtures.run(u"第一段")) + u"<w:p/>" + u"<w:p/>"
                + u"<w:p/>" + fixtures.paragraph(fixtures.run(u"第二段")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["删除空行"], 3, u"默认：空行整段删掉")
        self.assertEqual(self.texts(out), [u"第一段", u"第二段"])

    def test_a_single_blank_line_is_deleted_too(self):
        """**单个空行也删**（用户 2026-10-08 的截图就是这种情况：一个空行把一段话劈成两半）。"""
        body = (fixtures.paragraph(fixtures.run(u"上")) + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"下")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["删除空行"], 1)
        self.assertEqual(self.texts(out), [u"上", u"下"])

    def test_collapse_mode_is_still_available(self):
        body = (fixtures.paragraph(fixtures.run(u"第一段")) + u"<w:p/>" + u"<w:p/>"
                + u"<w:p/>" + fixtures.paragraph(fixtures.run(u"第二段")))
        report, out = self.run_op(body, collapse_blank_lines=True)
        self.assertEqual(report["changes"]["空白段压缩"], 2, u"旧档：三个空段留一个")
        self.assertEqual(self.texts(out), [u"第一段", u"", u"第二段"])
        self.assertIsNone(report["changes"].get(u"删除空行"), u"两种口径互斥，不能都算")

    def test_collapse_mode_keeps_a_lone_blank_line(self):
        body = (fixtures.paragraph(fixtures.run(u"上")) + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"下")))
        report, out = self.run_op(body, collapse_blank_lines=True)
        self.assertEqual(report["total"], 0, u"旧档：单独一个空行是「想留的」，不动")

    def test_blank_paragraphs_at_the_edges_go_away(self):
        body = (u"<w:p/>" + u"<w:p/>" + fixtures.paragraph(fixtures.run(u"正文"))
                + u"<w:p/>")
        report, out = self.run_op(body)
        self.assertGreaterEqual(report["changes"]["删除空行"], 3)
        self.assertEqual(self.texts(out), [u"正文"])

    def test_whitespace_only_paragraphs_count_as_blank(self):
        body = (fixtures.paragraph(fixtures.run(u"甲"))
                + fixtures.paragraph(fixtures.run(u"   "))
                + fixtures.paragraph(fixtures.run(u"乙")))
        report, out = self.run_op(body)
        self.assertEqual(report["changes"]["段尾空格"], 1, u"全空白段落会被清成真空段")
        self.assertEqual(report["changes"]["删除空行"], 1, u"清完就是一个空行 → 默认删掉")
        self.assertEqual(self.texts(out), [u"甲", u"乙"])


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


class TestFrontmatterProtection(unittest.TestCase):
    """前置区保护（用户 2026-09-27）：封面/扉页/签字页里的空白行**不许删**。

    识别口径：从文档开头走到「目录/前言」为止都是前置区；签字页靠 编制/校核/审核… 角色词认。
    """

    def _doc_with_frontmatter(self, path):
        cover = (fixtures.paragraph(fixtures.run(u"某水库工程防洪安全复核报告", sz="72"))
                 + u"<w:p/>"                      # 封面里的排版空行（**必须保住**）
                 + u"<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
                 + fixtures.paragraph(fixtures.run(u"编制：张三"))
                 + fixtures.paragraph(fixtures.run(u"校核：李四"))
                 + fixtures.paragraph(fixtures.run(u"审核：王五"))
                 + u"<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"
                 + fixtures.paragraph(fixtures.run(u"前 言"))
                 + fixtures.paragraph(fixtures.run(u"受 XX 委托，我院开展了复核工作。"))
                 + fixtures.paragraph(fixtures.run(u""))
                 + fixtures.paragraph(fixtures.run(u"  "))
                 + fixtures.paragraph(fixtures.run(u"正文第二段。"))
                 + fixtures.paragraph(fixtures.run(u"")))
        fixtures.write_fixture(path, body=cover)
        return path

    def test_detect_finds_cover_signature_and_marker(self):
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_")
        try:
            path = self._doc_with_frontmatter(os.path.join(work, u"报告.docx"))
            with Document(path) as doc:
                info = frontmatter.detect(doc)
            self.assertTrue(info["present"])
            self.assertEqual(info["marker"], u"前言")
            # 前言也属于前置区（用户 2026-09-30 补充的五种页面），所以是三页
            self.assertEqual(info["pages"], 3)
            self.assertTrue(info["cover"])
            self.assertTrue(info["title_page"])
            self.assertTrue(info["signature_page"])
            self.assertIn(u"编制", info["roles"])
            human = frontmatter.format_report(info)
            self.assertIn(u"封面：有", human)
            self.assertIn(u"签字页：有", human)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_tidy_does_not_touch_frontmatter_blanks(self):
        """前置区（封面/扉页/签字页/前言）里的空白行**不许删**；正文里的该删就删。

        （2026-10-08：默认口径从"空段压成一个"改成"整段删掉"后重写 —— 夹具里给封面
        加了一个排版空行，钉住"它在、正文那些不在"。）
        """
        work = tempfile.mkdtemp(prefix="wf_fm_")
        try:
            path = self._doc_with_frontmatter(os.path.join(work, u"报告.docx"))
            out = os.path.join(work, u"清理后.docx")
            with Document(path) as doc:
                report = tidy_op.tidy(doc)
                doc.save(out)
            self.assertGreaterEqual(report["changes"].get(u"删除空行", 0), 1,
                                    u"正文里的空段要删")
            self.assertGreater(report["changes"].get(u"前置区保护（跳过）", 0), 0)
            with Document(out) as doc:
                elements = [el for el in doc.body() if el.tag == qn("w:p")]
            texts = [Paragraph(el).text for el in elements]
            self.assertIn(u"", texts, u"封面里的排版空行必须原样保留")
            # 换页段本身没文字，但它是版面不是空白行 —— 排除掉再数
            real_blanks = [el for el in elements
                           if not Paragraph(el).text.strip()
                           and not tidy_op.carries_page_break(el)]
            self.assertEqual(len(real_blanks), 1,
                             u"正文里的空段该删的都删了，只剩封面那一个")
            with zipfile.ZipFile(out) as archive:
                text = archive.read("word/document.xml").decode("utf-8")
            self.assertEqual(text.count(u"<w:p><w:r><w:br"), 2,
                             u"封面/签字页的分页符段落必须原样保留")
            self.assertIn(u"编制：张三", text)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_no_marker_means_no_protection(self):
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_")
        try:
            path = os.path.join(work, u"没目录.docx")
            body = (fixtures.paragraph(fixtures.run(u"第一段。"))
                    + fixtures.paragraph(fixtures.run(u""))
                    + fixtures.paragraph(fixtures.run(u"第二段。")))
            fixtures.write_fixture(path, body=body)
            with Document(path) as doc:
                info = frontmatter.detect(doc)
            self.assertFalse(info["present"])
            self.assertIn(u"没遍历到", info["note"])
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_manual_pages_add_to_auto_instead_of_replacing_it(self):
        """手动页数是"至少再加保护到第 N 页"，不能把自动识别改小（用户 2026-09-30）。

        实测真实报告：封面在第 4–5 页，手动填"1 页"若替代自动识别，
        前置区 13 个排版空段会被清掉 10 个 —— 封面直接垮。
        """
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_add_")
        try:
            path = os.path.join(work, u"报告.docx")
            cover = (fixtures.paragraph(fixtures.run(u"某某水库工程", sz="72"))
                     + u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                     + fixtures.paragraph(fixtures.run(u"编制：张三  校核：李四"))
                     + fixtures.paragraph(fixtures.run(u"前 言"))
                     + fixtures.paragraph(fixtures.run(u"正文。")))
            fixtures.write_fixture(path, body=cover)
            with Document(path) as doc:
                auto = frontmatter.protected_elements(doc)
                manual = frontmatter.protected_elements(doc, pages=1)
            self.assertTrue(auto)
            self.assertEqual(manual, auto,
                             u"手动 1 页不能比自动识别更少（自动永远保护）")
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_cover_is_the_first_page_with_text_not_always_page_one(self):
        """真实报告封面在第 4–5 页（前面是排版空段）—— 按"第 1 页"判会误报无。"""
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_cover_")
        try:
            path = os.path.join(work, u"报告.docx")
            body = u"".join(u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>' for _ in range(3))
            body += (fixtures.paragraph(fixtures.run(u"某某水库工程防洪安全复核报告", sz="72"))
                     + fixtures.paragraph(fixtures.run(u"前 言"))
                     + fixtures.paragraph(fixtures.run(u"正文。")))
            fixtures.write_fixture(path, body=body)
            with Document(path) as doc:
                info = frontmatter.detect(doc)
            self.assertEqual(info["page_map"]["cover"], u"第 4 页")
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_real_page_numbers_win_over_the_estimate(self):
        """给了**真实页码**（Word/WPS 排版结果）就按它分页 —— 排版空段/回车不换页。"""
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_real_")
        try:
            path = os.path.join(work, u"报告.docx")
            # 15 个"带分页符的空段"——按符号数会数成 15 页，真实排版全在第 1 页
            body = u"".join(u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>' for _ in range(10))
            body += (fixtures.paragraph(fixtures.run(u"某某水库工程防洪安全复核报告", sz="72"))
                     + fixtures.paragraph(fixtures.run(u"前 言"))
                     + fixtures.paragraph(fixtures.run(u"正文。")))
            fixtures.write_fixture(path, body=body)
            with Document(path) as doc:
                estimate = frontmatter.detect(doc)
                # 真实页码：前 11 块都在第 1 页，"前言"起第 2 页
                block_pages = [1] * 11 + [2, 2]
                real = frontmatter.detect(doc, block_pages=block_pages)
            self.assertGreater(estimate["pages"], real["pages"],
                               u"估算口径会把排版空段数成很多页")
            self.assertEqual(real["page_map"]["cover"], u"第 1 页")
            self.assertTrue(real["page_map"]["accurate"])
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_a_title_containing_review_is_not_a_signature_page(self):
        """"防洪安全**复核**报告"里的复核不是签字角色（2026-09-30 实测误判）。"""
        from wordfactory import frontmatter
        work = tempfile.mkdtemp(prefix="wf_fm_role_")
        try:
            path = os.path.join(work, u"报告.docx")
            body = (fixtures.paragraph(fixtures.run(u"某某水库防洪安全复核报告", sz="72"))
                    + fixtures.paragraph(fixtures.run(u"前 言"))
                    + fixtures.paragraph(fixtures.run(u"编制方案：本报告。")))   # 长行 + 编制
            fixtures.write_fixture(path, body=body)
            with Document(path) as doc:
                info = frontmatter.detect(doc)
            self.assertEqual(info["roles"], [], u"长行里的编制不算签字角色")
            self.assertFalse(info["signature_page"])
        finally:
            shutil.rmtree(work, ignore_errors=True)


class TestPageBreaksAreNeverBlankLines(unittest.TestCase):
    """带换页记号的空段**不是**空白行，删了就塌版（用户 2026-09-30 实测 40 个）。

    原话："这个目录前面的分页符也是不计入删除的……你需要处理的只有真正的回车引起的
    空白行和空白格，像是 tab 的空格不需要你删除，分页符也不需要"。
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_break_")
        self.path = os.path.join(self.dir, u"分页.docx")
        body = (u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'
                + fixtures.paragraph(fixtures.run(u"前言第一段。"))
                + u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'
                + fixtures.paragraph(fixtures.run(u"前言第二段。"))
                + u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                + u'<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
                  u'</w:sectPr></w:pPr></w:p>'
                + fixtures.paragraph(fixtures.run(u"正文一段。"))
                + fixtures.paragraph(fixtures.run(u"  "))     # 相邻的两个真空白段（该删一个）
                + fixtures.paragraph(fixtures.run(u""))
                + fixtures.paragraph(fixtures.run(u"结尾。"))
                )
        fixtures.write_fixture(self.path, body=body)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _breaks(self, document):
        from wordfactory.ops.tidy import carries_page_break
        return [el for el in list(document.body())
                if el.tag == qn("w:p") and not Paragraph(el).text.strip()
                and carries_page_break(el)]

    def test_page_break_paragraphs_survive_tidy(self):
        with Document(self.path) as doc:
            before = len(self._breaks(doc))
            report = tidy_op.tidy(doc)
            after = len(self._breaks(doc))
        self.assertEqual(before, 4, u"夹具里应有 4 个带换页记号的空段")
        self.assertEqual(after, before, u"分页符一个都不能少")
        self.assertGreaterEqual(report["changes"].get(u"空行里带分页符（保留）", 0), 4)

    def test_real_blank_paragraphs_are_still_removed(self):
        with Document(self.path) as doc:
            report = tidy_op.tidy(doc)
        self.assertGreaterEqual(
            report["changes"].get(u"删除空行", 0)
            + report["changes"].get(u"空白段压缩", 0)
            + report["changes"].get(u"首尾空白段删除", 0), 1,
            u"真正的空白行还是要删的 —— 分页符保护不能变成什么都不删")
        self.assertGreaterEqual(report["changes"].get(u"空行里带分页符（保留）", 0), 4,
                                u"留了几条要报出来（用户要能看出「哪些没动、为什么」）")

    def test_dry_run_leaves_everything(self):
        with Document(self.path) as doc:
            tidy_op.tidy(doc, dry_run=True)
            self.assertEqual(len(self._breaks(doc)), 4)


class TestFrontMatterPartsAndToc(unittest.TestCase):
    """五种前置页面（用户 2026-09-30 补充的完整概念）。

    顺序一般是 **封面 → 扉页 → 签字页 → 前言 → 目录**；前言/目录本身也算前置区
    （不删空白）；**目录后面那个分页符之后才是正文**。
    「前 言」「目 录」中间有 1~2 个空格也要认。
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_fm_parts_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _build(self, toc_as_sdt=True, toc_text=u"目 录"):
        path = os.path.join(self.dir, u"报告.docx")
        parts = [fixtures.paragraph(fixtures.run(u"某某水库工程防洪安全复核报告", sz="72")),
                 u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>',
                 fixtures.paragraph(fixtures.run(u"编制：张三")),
                 u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>',
                 fixtures.paragraph(fixtures.run(u"前 言")),
                 fixtures.paragraph(fixtures.run(u"受委托开展复核。")),
                 u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>']
        if toc_as_sdt:
            parts.append(
                u'<w:sdt><w:sdtPr><w:docPartGallery w:val="目录"/></w:sdtPr>'
                u'<w:sdtContent><w:p><w:r><w:t>1 水库概况</w:t></w:r></w:p>'
                u'</w:sdtContent></w:sdt>')
        else:
            parts.append(fixtures.paragraph(fixtures.run(toc_text)))
            parts.append(fixtures.paragraph(fixtures.run(u"1 水库概况…………1")))
        parts.append(u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>')
        parts.append(fixtures.paragraph(fixtures.run(u"水库概况")))
        parts.append(fixtures.paragraph(fixtures.run(u"基本情况")))
        fixtures.write_fixture(path, body=u"".join(parts))
        return path

    def test_toc_as_sdt_is_part_of_the_front_matter(self):
        from wordfactory import frontmatter
        from wordfactory.ooxml import qn as _qn
        from wordfactory.text import Paragraph
        path = self._build(toc_as_sdt=True)
        with Document(path) as doc:
            info = frontmatter.detect(doc)
            blocks = list(doc.body())
            first_body = blocks[info["blocks"]]
        self.assertEqual(Paragraph(first_body).text.strip(), u"水库概况",
                         u"目录后面的分页符之后才是正文（正文第一块 = 水库概况）")
        self.assertNotEqual(info["page_map"]["toc"], u"无", u"目录应该被认出来")
        self.assertIn(u"目录", info["marker"])

    def test_toc_heading_with_spaces_is_recognized(self):
        """「目 录」「前 言」中间有空格也要认（用户明确要求作为辨别依据）。"""
        from wordfactory import frontmatter
        from wordfactory.text import Paragraph
        for heading in (u"目 录", u"目  录", u"目录"):
            path = self._build(toc_as_sdt=False, toc_text=heading)
            with Document(path) as doc:
                info = frontmatter.detect(doc)
                blocks = list(doc.body())
                first_body = blocks[info["blocks"]]
            self.assertEqual(Paragraph(first_body).text.strip(), u"水库概况",
                             u"标题「%s」没被认出来" % heading)

    def test_preface_is_protected_too(self):
        from wordfactory import frontmatter
        path = self._build(toc_as_sdt=True)
        with Document(path) as doc:
            info = frontmatter.detect(doc)
        self.assertTrue(info["preface_covered"] if "preface_covered" in info
                        else info["page_map"]["preface"] != u"无",
                        u"前言属于前置区（不删空白）")
