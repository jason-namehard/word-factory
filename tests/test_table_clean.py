# -*- coding: utf-8 -*-
"""宏「表格空格回车删除」（去无意义空格）的单测。

判据来自参考宏 `表格空格回车删除.bas:144-158` 的三档字符集，以及两条**有意不同**的口径：
* 只删字符、保留 run 级格式（宏是整格写回、会压平）——`--flat` 才照抄宏；
* 默认不删全角空格 U+3000（宏也删不掉）。
另外要钉住的：段落标记（Chr(13)）在 XML 里是 `w:p` 边界 → "删它" = 合并段落；
`Chr(7)`（单元格标记）在 XML 里就是 `w:tc`，无事可做。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.ops import table_clean as op
from wordfactory.text import Paragraph

from . import fixtures


def cell(*runs):
    return u"<w:tc><w:tcPr/>%s</w:tc>" % u"".join(runs)


def row(*cells):
    return u"<w:tr>%s</w:tr>" % u"".join(cells)


def table(*rows):
    return (u"<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + u"".join(rows) + u"</w:tbl>")


def para(runs):
    return u"<w:p>%s</w:p>" % runs


class TableCleanCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_tableclean_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body):
        fixtures.write_fixture(self.path, body=body)
        return self.path

    def run_op(self, body, **options):
        self.build(body)
        with Document(self.path) as doc:
            report = op.clean(doc, options)
            out = None
            if report["changed_cells"]:
                out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, out

    def cells_text(self, path):
        with Document(path) as doc:
            return [[Paragraph(p).text for p in cell.iter(qn("w:p"))]
                    for cell in doc.part().iter(qn("w:tc"))]


class TestLevels(TableCleanCase):
    def test_level_1_removes_half_width_and_nbsp_only(self):
        body = table(row(cell(para(fixtures.run(u"水 库\u00a0水位")))))
        report, out = self.run_op(body, level=1)
        self.assertEqual(report["changed_cells"], 1)
        self.assertEqual(self.cells_text(out), [[u"水库水位"]])
        self.assertEqual(report["chars_removed"], 2)

    def test_level_2_removes_breaks_and_merges_paragraphs(self):
        """`--exact-macro`：照抄宏，Chr(13)（段落标记）在 XML 里是 `w:p` 边界 → 合并两个段落。"""
        body = table(row(cell(para(fixtures.run(u"第一行")) + para(fixtures.run(u"第二行")))))
        report, out = self.run_op(body, level=2, exact_macro=True)
        self.assertEqual(report["paragraphs_merged"], 1)
        self.assertEqual(self.cells_text(out), [[u"第一行第二行"]])

    def test_level_2_removes_manual_line_breaks(self):
        """`--exact-macro` 时连内部换行也删。"""
        body = table(row(cell(para(u'<w:r><w:t>上</w:t><w:br/><w:t>下</w:t></w:r>'))))
        report, out = self.run_op(body, level=2, exact_macro=True)
        self.assertEqual(report["chars_removed"], 1)
        self.assertEqual(self.cells_text(out), [[u"上下"]])

    def test_level_2_keeps_spaces(self):
        body = table(row(cell(para(fixtures.run(u"水 库")))))
        report, _ = self.run_op(body, level=2)
        self.assertEqual(report["changed_cells"], 0)

    def test_level_3_does_both(self):
        body = table(row(cell(para(fixtures.run(u"水 库")) + para(fixtures.run(u"位\u00a0置")))))
        report, out = self.run_op(body, level=3, exact_macro=True)
        self.assertEqual(self.cells_text(out), [[u"水库位置"]])

    def test_the_default_keeps_intentional_wraps(self):
        """**默认口径（用户 2026-09-22 选 1+2）**：非空段落是正经结构，不合并。"""
        body = table(row(cell(para(fixtures.run(u"防洪标")) + para(fixtures.run(u"准")))))
        report, out = self.run_op(body, level=3)
        self.assertEqual(report["changed_cells"], 0, u"两个非空段落 → 不动它")
        self.assertIsNone(out, u"没改就不写文件")
        self.assertEqual(self.cells_text(self.path), [[u"防洪标", u"准"]])

    def test_the_default_keeps_a_single_internal_manual_break(self):
        """`防洪标↵准` 这种"为了窄列故意折的行"必须留着。"""
        body = table(row(cell(para(u'<w:r><w:t>防洪标</w:t><w:br/><w:t>准</w:t></w:r>'))))
        report, out = self.run_op(body, level=3)
        self.assertEqual(report["changed_cells"], 0)
        self.assertEqual(report["breaks_kept"], 1)

    def test_the_default_drops_leading_trailing_and_double_breaks(self):
        body = table(row(cell(para(u'<w:r><w:br/><w:t>值</w:t><w:br/><w:br/></w:r>'))))
        report, out = self.run_op(body, level=3)
        self.assertEqual(report["changed_cells"], 1, u"首尾的与连续重复的换行要删")
        self.assertEqual(self.cells_text(out), [[u"值"]])

    def test_a_numeric_cell_is_flattened_anyway(self):
        """整格是纯数字 → 换行一律删（数字里断行只能是脏数据）。"""
        body = table(row(cell(para(u'<w:r><w:t>37.</w:t><w:br/><w:t>47</w:t></w:r>'))))
        report, out = self.run_op(body, level=3)
        self.assertEqual(self.cells_text(out), [[u"37.47"]])
        self.assertEqual(report["breaks_kept"], 0, u"数字格不留内部换行")

    def test_a_blank_paragraph_is_dropped(self):
        body = table(row(cell(para(fixtures.run(u"值")) + u"<w:p/>")))
        report, out = self.run_op(body, level=3)
        self.assertEqual(report["paragraphs_merged"], 1)
        self.assertEqual(self.cells_text(out), [[u"值"]])

    def test_level_3_is_the_default(self):
        body = table(row(cell(para(fixtures.run(u"水 库")))))
        report, _ = self.run_op(body)
        self.assertEqual(report["level"], 3)
        self.assertEqual(report["changed_cells"], 1)

    def test_a_bad_level_says_so(self):
        self.build(table(row(cell(para(fixtures.run(u"x"))))))
        with Document(self.path) as doc:
            with self.assertRaises(ValueError):
                op.clean(doc, {"level": 9})


class TestWhatIsDeliberatelyDifferent(TableCleanCase):
    def test_run_formatting_is_kept_by_default(self):
        """宏整格写回会把格式压平；我们只删字符 → 加粗那半截还是加粗。"""
        body = table(row(cell(para(
            u'<w:r><w:t>水 库</w:t></w:r>'
            u'<w:r><w:rPr><w:b/></w:rPr><w:t> 水位</w:t></w:r>'))))
        report, out = self.run_op(body, level=1)
        with Document(out) as doc:
            runs = list(doc.part().iter(qn("w:r")))
        self.assertEqual(len(runs), 2, u"run 结构不该被重建")
        self.assertIsNotNone(runs[1].find(qn("w:rPr")).find(qn("w:b")), u"加粗要留着")
        self.assertEqual(self.cells_text(out), [[u"水库水位"]])

    def test_flat_replicates_the_macro(self):
        body = table(row(cell(para(
            u'<w:r><w:t>水 库</w:t></w:r>'
            u'<w:r><w:rPr><w:b/></w:rPr><w:t> 水位</w:t></w:r>'))))
        report, out = self.run_op(body, level=1, flat=True)
        with Document(out) as doc:
            runs = list(doc.part().iter(qn("w:r")))
        self.assertEqual(len(runs), 1, u"--flat 要照宏压成一个 run")
        self.assertEqual(self.cells_text(out), [[u"水库水位"]])

    def test_the_full_width_space_is_kept_unless_asked(self):
        body = table(row(cell(para(fixtures.run(u"水\u3000库")))))
        report, _ = self.run_op(body, level=1)
        self.assertEqual(report["changed_cells"], 0, u"全角空格默认不删（宏也删不掉）")
        report, out = self.run_op(body, level=1, full_width_space=True)
        self.assertEqual(self.cells_text(out), [[u"水库"]])


class TestScopeAndEdges(TableCleanCase):
    def test_only_table_cells_are_touched(self):
        body = (fixtures.paragraph(fixtures.run(u"正文里的 空格\u00a0不动"))
                + table(row(cell(para(fixtures.run(u"格 内"))))))
        report, out = self.run_op(body, level=3)
        texts = self.cells_text(out)
        with Document(out) as doc:
            body_texts = [Paragraph(p).text for p in doc.body()
                          if p.tag == qn("w:p")]
        self.assertEqual(texts, [[u"格内"]])
        self.assertIn(u"正文里的 空格\u00a0不动", body_texts, u"表格外的文字不许动")

    def test_an_empty_cell_is_skipped(self):
        body = table(row(cell(para(u"")), cell(para(fixtures.run(u" 有 空格")))))
        report, out = self.run_op(body, level=1)
        self.assertEqual(report["cells"], 2)
        self.assertEqual(report["changed_cells"], 1)

    def test_xml_space_is_maintained(self):
        """删掉中间空格后，首尾若出现空格必须补 `xml:space="preserve"`，否则会被 Word 吃掉。"""
        body = table(row(cell(para(fixtures.run(u" x ")))))
        report, out = self.run_op(body, level=1)
        with Document(out) as doc:
            node = list(doc.part().iter(qn("w:t")))[0]
        self.assertEqual(node.text, u"x")
        self.assertIsNone(node.get("{http://www.w3.org/XML/1998/namespace}space"))

    def test_leading_space_keeps_preserve(self):
        """删完空格后若**首尾仍有空白**（例如保留的全角空格），`xml:space` 必须补上。"""
        body = table(row(cell(para(u'<w:r><w:t>\u3000a b</w:t></w:r>'))))
        report, out = self.run_op(body, level=1)
        with Document(out) as doc:
            node = list(doc.part().iter(qn("w:t")))[0]
        self.assertEqual(node.text, u"\u3000ab")
        self.assertEqual(node.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve",
                         u"首尾还有空白 → 必须留 preserve，否则 Word 会吃掉")

    def test_dry_run_changes_nothing(self):
        body = table(row(cell(para(fixtures.run(u"水 库")))))
        self.build(body)
        with Document(self.path) as doc:
            report = op.clean(doc, {"level": 1}, dry_run=True)
            self.assertEqual(doc.dirty_parts, [])
            self.assertEqual(self.cells_text(self.path), [[u"水 库"]])
        self.assertEqual(report["changed_cells"], 1)


if __name__ == "__main__":
    unittest.main()
