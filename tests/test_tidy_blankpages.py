# -*- coding: utf-8 -*-
"""空白页 / 空白行清理的回归测试（2026-10-08 那轮实测补的）。

用户 2026-10-08 的原话（本轮任务的由来）：

> 大部分都改对了，现在空白页检索和空白行删除做的还是不够好……测试报告的第一页
> 明明检查的时候发现了不是封面，却还保留了这个空白页及其空白行，然后是页码的第 4 页，
> 还有大量的空白行没有删除……关于第 4 页的那个空白行，你可以先检索下，
> 确认下它的空白行是什么字符，确定是否是字符原因导致的格式修正失败

**检索结果**（`tmp/probe_*` 探针实测）：

* 那些"空白行"的实际字符是**制表符 ``\\t`` (U+0009)**（`Paragraph.text` 读出来是 ``"\\t"``，
  ``.strip()`` 之后确实是空的）—— **字符本身不是删不掉的原因**；
* 真正的原因是 WPS 导出的报告**每个段落**都带 ``<w:pageBreakBefore w:val="0"/>``，
  旧代码只看元素在不在，把 288 段全当成"带换页记号的空段"保护起来了；
* 第 4 页整页只有一个 ``w:br type="page"`` 的空段（**不是空白文字行**）；
* 第 1 页 22 个空段与第 2 页的封面之间**没有分页符**（自然溢出），
  页边界只有真实页码（`pageprobe` 起 Word/WPS 排版）才认得出来。
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

#: WPS 导出的报告里"空段"长这样：显式关掉的段前分页 + 一个制表符
TAB_BLANK = (u'<w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr>'
             u'<w:r><w:tab/></w:r></w:p>')
#: 一个真分页符段（段内插入分页符，段本身是空的）
PAGE_BREAK = u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


class BreakCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_pbb_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body):
        fixtures.write_fixture(self.path, body=body)
        return self.path

    def paragraph_texts(self, path=None):
        with Document(path or self.path) as doc:
            return [Paragraph(el).text for el in doc.body() if el.tag == qn("w:p")]


class TestPageBreakBeforeValIsHonoured(BreakCase):
    """``w:pageBreakBefore`` 必须看 ``w:val`` —— 这是本轮修掉的根因。"""

    def _carries(self, body):
        from wordfactory.ops.tidy import carries_page_break
        self.build(body)
        with Document(self.path) as doc:
            return [carries_page_break(el) for el in doc.body() if el.tag == qn("w:p")]

    def test_val_zero_is_off(self):
        self.assertEqual(self._carries(
            u'<w:p><w:pPr><w:pageBreakBefore w:val="0"/></w:pPr></w:p>'), [False],
            u'w:val="0" 是显式关掉，不是换页')
        for value in (u"false", u"off", u"none"):
            self.assertEqual(self._carries(
                u'<w:p><w:pPr><w:pageBreakBefore w:val="%s"/></w:pPr></w:p>' % value),
                [False], u"w:val=%s 也算关" % value)

    def test_bare_element_and_val_one_are_on(self):
        self.assertEqual(self._carries(
            u'<w:p><w:pPr><w:pageBreakBefore/></w:pPr></w:p>'
            u'<w:p><w:pPr><w:pageBreakBefore w:val="1"/></w:pPr></w:p>'
            u'<w:p><w:pPr><w:pageBreakBefore w:val="true"/></w:pPr></w:p>'),
            [True, True, True], u"没写 val 或写成真值都是开")

    def test_br_page_and_sectpr_still_count(self):
        self.assertEqual(self._carries(
            u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
            u'<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            u'</w:sectPr></w:pPr></w:p>'), [True, True])

    def test_tab_only_blank_lines_are_now_deleted(self):
        """制表符空段：``strip()`` 认得它 —— 删不掉是"被误判成换页段"，不是字符问题。"""
        body = (fixtures.paragraph(fixtures.run(u"甲")) + TAB_BLANK * 3
                + fixtures.paragraph(fixtures.run(u"乙")))
        self.build(body)
        with Document(self.path) as doc:
            report = tidy_op.tidy(doc)
        self.assertEqual(report["changes"].get(u"删除空行", 0), 3,
                         u"三个带制表符的空段该整段删掉（旧代码一个都不动）")

    def test_frontmatter_page_estimate_not_inflated(self):
        """十个 ``w:val="0"`` 空段全在第 1 页 —— 页数必须还是 1（旧代码算成 11）。"""
        from wordfactory import frontmatter
        body = TAB_BLANK * 10
        body += (fixtures.paragraph(fixtures.run(u"某某水库工程防洪安全复核报告", sz="72"))
                 + fixtures.paragraph(fixtures.run(u"前 言")))
        self.build(body)
        with Document(self.path) as doc:
            info = frontmatter.detect(doc)
        self.assertEqual(info["pages"], 1)


class TestBlankPages(BreakCase):
    """整页没内容 → 那页的空段全删。页边界要真实页码才认得出（自然溢出）。"""

    def test_leading_blank_page_goes_away_with_real_pages(self):
        cover = (fixtures.paragraph(fixtures.run(u"某某水库工程防洪安全复核报告", sz="72"))
                 + PAGE_BREAK
                 + fixtures.paragraph(fixtures.run(u"前 言"))
                 + fixtures.paragraph(fixtures.run(u"正文。")))
        path = self.build(TAB_BLANK * 5 + cover)
        out = path + u".out.docx"
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 1, 1, 1, 1, 2, 2, 2, 2]})
            doc.save(out)
        self.assertEqual(report["changes"].get(u"空白页数"), 1)
        self.assertEqual(report["changes"].get(u"空白页删段"), 5)
        texts = self.paragraph_texts(out)
        self.assertNotIn(u"\t", texts, u"空白首页的空段要删干净")
        self.assertTrue(any(u"前 言" in text for text in texts), u"前言不能被误删")

    def test_stray_page_break_blank_page_goes_away_with_real_pages(self):
        path = self.build(fixtures.paragraph(fixtures.run(u"签字页"))
                          + PAGE_BREAK
                          + fixtures.paragraph(fixtures.run(u"前 言")))
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 2, 3]})
        self.assertEqual(report["changes"].get(u"空白页删段"), 1)
        self.assertEqual(report["changes"].get(u"（其中原带分页符）"), 1,
                         u"整页只有分页符 → 连它一起删，白页才会消失")

    def test_stray_page_break_paragraph_is_kept_without_real_pages(self):
        """没给真实页码时**不敢删换页段**（估算口径认不准页边界，用户 2026-09-30 的规矩）。"""
        path = self.build(fixtures.paragraph(fixtures.run(u"签字页"))
                          + PAGE_BREAK
                          + fixtures.paragraph(fixtures.run(u"前 言")))
        with Document(path) as doc:
            report = tidy_op.tidy(doc)
            breaks = [el for el in doc.body()
                      if el.tag == qn("w:p") and any(
                          br.get(qn("w:type")) == "page" for br in el.iter(qn("w:br")))]
        self.assertEqual(len(breaks), 1, u"估算口径下分页符段必须留着")
        self.assertEqual(report["changes"].get(u"空白页删段", 0), 0)

    def test_page_with_only_a_drawing_is_not_blank(self):
        path = self.build(u'<w:p><w:r><w:drawing/></w:r></w:p>'
                          + fixtures.paragraph(fixtures.run(u"正文")))
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 2]})
        self.assertEqual(report["changes"].get(u"空白页删段", 0), 0, u"有图的页不是空白页")

    def test_paragraph_holding_only_a_drawing_survives_blank_line_removal(self):
        """**挂图的空段不能被当空白行删掉**（2026-10-08 实测：图 6-1 的图真被删了）。

        那份报告里图题段是空的，图挂在紧随其后的空段里（文字为空、里面有 ``w:drawing``）。
        旧代码只看 ``Paragraph.text`` → 把挂图的段当空白行删 → **整张图没了**。
        """
        image = u'<w:p><w:r><w:drawing/></w:r></w:p>'
        body = (fixtures.paragraph(fixtures.run(u"图 6-1  水位～库容关系图"))
                + image + u"<w:p/>" + u"<w:p/>"
                + fixtures.paragraph(fixtures.run(u"正文继续")))
        path = self.build(body)
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 1, 1, 1, 1]})
            doc.save(path + u".out.docx")
        with Document(path + u".out.docx") as doc:
            drawings = len(list(doc.part().iter(qn("w:drawing"))))
        self.assertEqual(drawings, 1, u"挂图的段删了 = 图丢了")
        self.assertEqual(report["changes"].get(u"删除空行", 0), 2,
                         u"挂图的段不算空段；后面那两个真空段才该删")

    def test_page_with_only_a_table_is_not_blank(self):
        path = self.build(fixtures.table([[u'<w:r><w:t>表内文字</w:t></w:r>',
                                           u'<w:r><w:t>值</w:t></w:r>']])
                          + fixtures.paragraph(fixtures.run(u"正文")))
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 2]})
        self.assertEqual(report["changes"].get(u"空白页删段", 0), 0, u"表格页有内容")

    def test_page_with_only_a_toc_sdt_is_not_blank(self):
        """目录整块包在 ``w:sdt`` 里 —— 只看直接 ``w:p`` 会把目录页误判成空白页。"""
        path = self.build(
            u'<w:sdt><w:sdtPr><w:docPartGallery w:val="目录"/></w:sdtPr>'
            u'<w:sdtContent><w:p><w:r><w:t>1 水库概况</w:t></w:r></w:p>'
            u'</w:sdtContent></w:sdt>'
            + u"".join(u'<w:p/>' for _ in range(2))
            + fixtures.paragraph(fixtures.run(u"正文")))
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"block_pages": [1, 1, 1, 2]})
        self.assertEqual(report["changes"].get(u"空白页删段", 0), 0,
                         u"sdt（目录）里的文字算内容，目录页不能当空白页")
