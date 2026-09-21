# -*- coding: utf-8 -*-
"""表头格式这个宏的单测。

判据尽量写成**不变量**（而不是"把公式再抄一遍"），这样公式改了、测试仍然有意义：

* 编号里不许再有空格；编号与名字之间至少有 ``min_spaces`` 个空格；
* **名字的中点落在「编号之后 → 版心右边界」的中点附近**（±1 个空格内）—— 这就是规则 B 本身；
* 段首是 ``firstLineChars=200``（首行缩进 2 字符，不是空格）；
* 表格 ``w:tblPr/w:jc = center``；
* 不是表头的段落一个字节都不许动；紧跟着表格才算表题（可关）；
* **幂等**：跑第二遍报 0 处、字节不变。
"""

import os
import shutil
import tempfile
import unittest
from xml.etree import ElementTree as ET

from wordfactory.document import Document, run_size_of
from wordfactory.ooxml import DocxPackage, qn
from wordfactory.ops import captions as header
from wordfactory.text import Paragraph

from . import fixtures


def build_docx(path, body):
    return fixtures.write_fixture(path, body=body)


def header_paragraph(number, name, spaces, table_after=True):
    """造一个表头段：``表X-Y`` + N 空格 + 名字（可选后面跟一张表）。"""
    parts = [fixtures.paragraph(fixtures.run(number + u" " * spaces + name))]
    if table_after:
        parts.append(fixtures.table([[u"<w:r><w:t>A</w:t></w:r>"] * 2]))
    return u"".join(parts)


class HeaderCase(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp(prefix="wf_header_")
        self.body = u"".join([
            fixtures.paragraph(fixtures.run(u"前言：这一段不是表头。")),
            header_paragraph(u"表 2.3-1", u"欧峪水库历史水位特性表", 6),
            header_paragraph(u"表4.2-1", u"水库洪水计算成果表", 20),
            u"".join([fixtures.paragraph(fixtures.run(u"表5-1 这个后面没有表格")),
                      fixtures.paragraph(fixtures.run(u"普通段落"))]),
        ])
        self.src = build_docx(os.path.join(self.work, "fixture.docx"), self.body)

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def _headers(self, doc):
        return [Paragraph(element) for element in doc.body()
                if element.tag == qn("w:p")
                and u"\u8868" in Paragraph(element).text]


class TestHeaderOp(HeaderCase):
    def test_plan_finds_only_captions_that_are_followed_by_a_table(self):
        with Document(self.src) as doc:
            plans = header.plan(doc)
        self.assertEqual(len(plans), 2, u"只该认「紧跟着表格」的两个表题")
        self.assertEqual([p.number for p in plans], [u"表2.3-1", u"表4.2-1"])

    def test_the_number_loses_its_internal_space(self):
        with Document(self.src) as doc:
            plans = header.plan(doc)
            self.assertEqual(plans[0].number, u"表2.3-1", u"表 与编号之间的空格要去掉")
            self.assertEqual(plans[0].text_want[:4], u"表2.3")

    def test_the_name_ends_up_centred_in_the_span_after_the_number(self):
        """规则 B 的不变量：名字中点 ≈ (编号之后 … 版心右边界) 的中点。"""
        with Document(self.src) as doc:
            geometry = doc.geometry()
            plans = header.plan(doc)
            for item in plans:
                size = run_size_of(item.paragraph)
                first_line = 2 * (size / 2.0) * 20
                number_w = geometry.width_of(item.number, size)
                name_w = geometry.width_of(item.name, size)
                start = first_line + number_w + item.spaces_want * geometry.space_width(size)
                center = start + name_w / 2.0
                target = (first_line + number_w + geometry.column_right) / 2.0
                self.assertLessEqual(abs(center - target), geometry.space_width(size) + 1,
                                     u"%s 的名字中点离目标太远" % item.number)

    def test_spaces_never_go_below_the_minimum(self):
        with Document(self.src) as doc:
            plans = header.plan(doc, {"min_spaces": 3})
        self.assertTrue(all(item.spaces_want >= 3 for item in plans))

    def test_a_caption_without_a_table_is_skipped_by_default(self):
        with Document(self.src) as doc:
            plans = header.plan(doc, {"only_before_table": False})
        self.assertEqual(len(plans), 3, u"放开限制后应把「表5-1」也算进来")


class TestHeaderApply(HeaderCase):
    def test_apply_sets_text_indent_and_table_alignment(self):
        out = os.path.join(self.work, "out.docx")
        with Document(self.src) as doc:
            report = header.apply(doc)
            self.assertEqual(report["changed"], 2)
            self.assertGreaterEqual(report["tables_centered"], 1)
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            body = pkg.xml(DocxPackage.MAIN).find(qn("w:body"))
            paragraphs = [c for c in body if c.tag == qn("w:p")]
            caption = None
            for element in paragraphs:
                if Paragraph(element).text.startswith(u"表2.3-1"):
                    caption = element
            self.assertIsNotNone(caption, u"找不到表题段")
            text = Paragraph(caption).text
            self.assertTrue(text.startswith(u"表2.3-1 "), text[:20])
            self.assertNotIn(u"表 2.3", text, u"编号里不该再有空格")
            ind = caption.find(qn("w:pPr") + "/" + qn("w:ind"))
            self.assertIsNotNone(ind)
            self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
            tables = [c for c in body if c.tag == qn("w:tbl")]
            jc = tables[0].find(qn("w:tblPr") + "/" + qn("w:jc"))
            self.assertIsNotNone(jc, u"表格应该被设成居中")
            self.assertEqual(jc.get(qn("w:val")), "center")

    def test_other_paragraphs_are_left_alone(self):
        out = os.path.join(self.work, "out2.docx")
        with Document(self.src) as doc:
            header.apply(doc)
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            texts = [Paragraph(e).text for e in pkg.xml(DocxPackage.MAIN).iter(qn("w:p"))]
        self.assertIn(u"前言：这一段不是表头。", texts)
        self.assertIn(u"普通段落", texts)
        self.assertTrue(any(t.startswith(u"表5-1 这个后面没有表格") for t in texts),
                        u"后面没有表格的表题不该被动")

    def test_dry_run_changes_nothing(self):
        with Document(self.src) as doc:
            before = ET.tostring(doc.part(), encoding="unicode")
            report = header.apply(doc, dry_run=True)
            after = ET.tostring(doc.part(), encoding="unicode")
            self.assertEqual(before, after, u"dry-run 不该改树")
            self.assertEqual(doc.dirty_parts, [])
        self.assertEqual(report["changed"], 2)

    def test_running_twice_changes_nothing_the_second_time(self):
        out = os.path.join(self.work, "out3.docx")
        with Document(self.src) as doc:
            first = header.apply(doc)
            second = header.apply(doc)
            self.assertGreater(first["changed"], 0)
            self.assertEqual(second["changed"], 0, u"第二遍不该再改")
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            body = pkg.xml(DocxPackage.MAIN).find(qn("w:body"))
        # 再开一次，用同样的规则算一遍，仍然 0 处
        with Document(out) as doc2:
            third = header.apply(doc2)
            self.assertEqual(third["changed"], 0, u"对输出再跑一次仍是 0 处")


class TestFigureRule(HeaderCase):
    """图题走另一套规矩：**整段居中** + 编号与名字之间**正好一个空格** + 清掉段首缩进。

    （用户 2026-09-21 原话：「图是"图4.2-1 欧峪水库30年一遇调洪演算图"整个居中，
      从 图这个字开始就居中，然后中间空一个格子」。）
    """

    def test_figure_gets_one_space_and_whole_paragraph_centred(self):
        body = u"".join([
            fixtures.paragraph(fixtures.run(u"图 4.2-1" + u" " * 12 + u"某张调洪演算图")),
            fixtures.paragraph(fixtures.run(u"图6-1" + u" " * 17 + u"水位库容关系图")),
        ])
        src = build_docx(os.path.join(self.work, "figs.docx"), body)
        out = os.path.join(self.work, "figs_out.docx")
        with Document(src) as doc:
            report = header.apply(doc)
            self.assertEqual(report["figures"], 2)
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            body_element = pkg.xml(DocxPackage.MAIN).find(qn("w:body"))
            seen = 0
            for element in body_element.iter(qn("w:p")):
                para = Paragraph(element)
                text = para.text
                if not text.startswith(u"图"):
                    continue
                seen += 1
                number = text.split(u" ")[0]
                self.assertEqual(text[len(number)], u" ", u"图题后要有一个空格")
                rest = text[len(number) + 1:]
                self.assertFalse(rest.startswith(u" "),
                                 u"图题中间只留一个空格：%r" % text[:24])
                jc = element.find(qn("w:pPr") + "/" + qn("w:jc"))
                self.assertIsNotNone(jc, u"图题要整段居中")
                self.assertEqual(jc.get(qn("w:val")), "center")
            self.assertEqual(seen, 2)

    def test_figure_loses_its_first_line_indent(self):
        body = (u'<w:p><w:pPr><w:ind w:firstLine="474" w:firstLineChars="200"/></w:pPr>'
                u'<w:r><w:t>图4.2-2 某图</w:t></w:r></w:p>')
        src = build_docx(os.path.join(self.work, "fig2.docx"), body)
        out = os.path.join(self.work, "fig2_out.docx")
        with Document(src) as doc:
            header.apply(doc)
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            element = list(pkg.xml(DocxPackage.MAIN).iter(qn("w:p")))[0]
            ind = element.find(qn("w:pPr") + "/" + qn("w:ind"))
            self.assertTrue(ind is None or ind.get(qn("w:firstLine")) is None,
                            u"图题的段首缩进必须清掉，否则居中会被顶偏")
            jc = element.find(qn("w:pPr") + "/" + qn("w:jc"))
            self.assertEqual(jc.get(qn("w:val")), "center")

    def test_figure_normalizes_its_number_too(self):
        body = fixtures.paragraph(fixtures.run(u"图 6-2 名字"))
        src = build_docx(os.path.join(self.work, "fig3.docx"), body)
        with Document(src) as doc:
            plans = header.plan(doc)
        self.assertEqual([p.number for p in plans], [u"图6-2"])

    def test_table_caption_still_gets_the_padding_rule(self):
        """同一份文档里两套规矩并存：表题仍是"空格居中"，不是整段居中。"""
        body = header_paragraph(u"表 2.3-1", u"库容特性表", 6)
        src = build_docx(os.path.join(self.work, "mix.docx"), body)
        out = os.path.join(self.work, "mix_out.docx")
        with Document(src) as doc:
            header.apply(doc)
            doc.mark_dirty()
            doc.save(out)
        with DocxPackage(out) as pkg:
            element = list(pkg.xml(DocxPackage.MAIN).iter(qn("w:p")))[0]
            jc = element.find(qn("w:pPr") + "/" + qn("w:jc"))
            self.assertEqual(jc.get(qn("w:val")), "left",
                             u"表题必须保持左对齐（靠空格居中），不能整段居中")


class TestAlreadyCorrectCaptionsAreNotTouched(HeaderCase):
    """**报"要改"就必须真改**：实测踩过一次——图题的 note 是写死的"整段居中"，
    于是图题每跑一遍都报"5 个要改"，第二遍的数字纯属假的，还会把没变的段落又标一次蓝。
    """

    def _figure(self, center=True, indent=u""):
        pr = (u"<w:pPr>%s%s</w:pPr>"
              % (u'<w:jc w:val="center"/>' if center else u"", indent))
        return u'<w:p>%s<w:r><w:t>图4.2-1 某调洪演算图</w:t></w:r></w:p>' % pr

    def test_a_figure_that_is_already_right_is_not_changed(self):
        src = build_docx(os.path.join(self.work, "ok.docx"), self._figure())
        with Document(src) as doc:
            report = header.apply(doc)
        self.assertEqual(report["changed"], 0, u"已经对了就不该报要改")
        self.assertEqual(report["figures"], 1, u"认出来了 1 个图题")
        self.assertEqual(report["changed_figures"], 0, u"但它不该被算成'要改'")
        self.assertEqual([item.note for item in header.plan(
            Document(src))], [None], u"已经对了的图题不该有备注")

    def test_a_figure_that_only_lacks_centring_is_still_fixed(self):
        """文字已经对了、但没居中 → 该改，而且第二遍要报 0。"""
        body = self._figure(center=False, indent=u'<w:ind w:firstLine="474"/>')
        src = build_docx(os.path.join(self.work, "nocentre.docx"), body)
        out = os.path.join(self.work, "nocentre_out.docx")
        with Document(src) as doc:
            first = header.apply(doc)
            doc.mark_dirty()
            doc.save(out)
        self.assertEqual(first["changed"], 1, u"缺居中的图题要改")
        with DocxPackage(out) as pkg:
            element = list(pkg.xml(DocxPackage.MAIN).iter(qn("w:p")))[0]
            self.assertEqual(element.find(qn("w:pPr") + "/" + qn("w:jc")).get(qn("w:val")), "center")
            self.assertIsNone(element.find(qn("w:pPr") + "/" + qn("w:ind")),
                              u"段首缩进要清掉")
        with Document(out) as doc:
            second = header.apply(doc)
        self.assertEqual(second["changed"], 0, u"改完第二遍必须是 0")

    def test_a_table_caption_whose_text_is_right_but_indent_is_missing_is_fixed(self):
        """表题也一样：文字对了但段首没缩进 → 该补上（不然那颗"首行空两格"就丢了）。"""
        body = u'<w:p><w:r><w:t>表4.2-1 %s水库洪水计算成果表</w:t></w:r></w:p>' % (u" " * 20)
        body += fixtures.table([[u"<w:r><w:t>A</w:t></w:r>"] * 2])
        src = build_docx(os.path.join(self.work, "tabs.docx"), body)
        out = os.path.join(self.work, "tabs_out.docx")
        with Document(src) as doc:
            first = header.apply(doc)
            doc.mark_dirty()
            doc.save(out)
        self.assertEqual(first["changed"], 1, u"缺段首缩进的表题要改")
        with DocxPackage(out) as pkg:
            element = list(pkg.xml(DocxPackage.MAIN).iter(qn("w:p")))[0]
            ind = element.find(qn("w:pPr") + "/" + qn("w:ind"))
            self.assertEqual(ind.get(qn("w:firstLineChars")), "200")
        with Document(out) as doc:
            second = header.apply(doc)
        self.assertEqual(second["changed"], 0, u"改完第二遍必须是 0")


class TestHeaderInvariants(HeaderCase):
    def test_a_longer_name_gets_fewer_spaces(self):
        """单调性：名字越长，需要的空格越少（这条不依赖任何公式细节）。"""
        body = u"".join([
            header_paragraph(u"表1-1", u"短名字", 10),
            header_paragraph(u"表1-2", u"这是一个明显更长的表格名字用来占位", 10),
        ])
        src = build_docx(os.path.join(self.work, "mono.docx"), body)
        with Document(src) as doc:
            plans = header.plan(doc)
        self.assertEqual(len(plans), 2)
        self.assertGreater(plans[0].spaces_want, plans[1].spaces_want)


class TestHeaderAgainstTheUsersOwnDocument(unittest.TestCase):
    """在你那份真实文档上做回归：规则 B 应该落在他现有值 ±3 格内（≥10/12）。

    （文件不在就跳过 —— 这条测试是"用真实数据验收公式"，不该让别人的环境失败。）
    """

    DOC = os.path.join(os.environ.get("USERPROFILE", ""), "Desktop",
                       u"\u67a3\u5e84\u5e02\u5c71\u4ead\u533a\u6b27\u5cea\u6c34\u5e93"
                       u"\u9632\u6d2a\u5b89\u5168\u590d\u6838\u62a5\u544a.docx")

    def test_rule_b_reproduces_his_current_spacing_within_three_spaces(self):
        if not os.path.exists(self.DOC):
            self.skipTest(u"用户文档不在本机：%s" % self.DOC)
        with Document(self.DOC) as doc:
            plans = header.plan(doc)
        # 只拿**表题**验规则 B —— 图题走另一套规矩（整段居中 + 一个空格），不在这条判据里
        tables = [item for item in plans if item.kind == u"\u8868"]
        self.assertGreaterEqual(len(tables), 10, u"应该能认出十来个表题")
        deltas = [item.spaces_want - item.spaces_have for item in tables]
        within = len([d for d in deltas if abs(d) <= 3])
        self.assertGreaterEqual(within, len(deltas) * 0.8,
                                u"落在 ±3 格内的比例太低：%r（deltas=%r）"
                                % (within / len(deltas), deltas))

    def test_figures_are_planned_with_the_figure_rule(self):
        if not os.path.exists(self.DOC):
            self.skipTest(u"用户文档不在本机：%s" % self.DOC)
        with Document(self.DOC) as doc:
            plans = header.plan(doc)
        figures = [item for item in plans if item.kind == u"\u56fe"]
        self.assertGreaterEqual(len(figures), 3, u"他文档里有几个图题")
        for item in figures:
            self.assertEqual(item.spaces_want, 1, u"图题中间只留一个空格：%s" % item.number)
            self.assertTrue(item.number.startswith(u"图"),
                            u"图题编号不该带空格：%r" % item.number)


if __name__ == "__main__":
    unittest.main()
