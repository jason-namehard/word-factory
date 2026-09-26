# -*- coding: utf-8 -*-
"""PDF 导出编排（`ops/pdf.py`）的单测。

要钉住的：

* **页数按 PDF 本身数**（不信渲染器自报 —— 实测 Word 的 `ComputeStatistics` 给 1，PDF 里 /Count 是 26）；
* 用绝对路径打开（Word 的 COM 会话有自己的工作目录，相对路径会解析到 system32）；
* 渲染器探测：Word/WPS/LibreOffice 有没有要如实说；**一个都没有时给人话提示**，不假装成功；
* `--dry-run` 只报计划不真跑。

真正导出（会打开 Word）放在 `TestRealExportIfAnyRenderer` 里 —— 没有渲染器就跳过。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.ops import pdf as pdf_op
from wordfactory.ooxml import qn

from . import fixtures


class PdfCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_pdf_")
        self.path = os.path.join(self.dir, u"夹具.docx")
        body = (fixtures.paragraph(fixtures.run(u"第一段。"))
                + fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>",
                                   u"<w:r><w:t>库容</w:t></w:r>"]]))
        fixtures.write_fixture(self.path, body=body)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_plan_names_the_renderer_and_the_command(self):
        plan = pdf_op.plan(self.path, os.path.join(self.dir, u"出.pdf"))
        self.assertTrue(os.path.isabs(plan["out"]))
        self.assertIn(plan["renderer"], ("word", "wps", "libreoffice"))
        self.assertTrue(plan["command"], u"要说明打算怎么导")

    def test_plan_rejects_an_unavailable_renderer_with_words(self):
        renderers = [item["name"] for item in pdf_op.detect_renderers()
                     if item["available"]]
        if not renderers:
            self.skipTest(u"本机没有渲染器")
        missing = [name for name in ("word", "wps", "libreoffice")
                   if name not in renderers]
        if not missing:
            self.skipTest(u"三种渲染器都有，没法测「没有」的情况")
        with self.assertRaises(pdf_op.PdfError) as caught:
            pdf_op.plan(self.path, os.path.join(self.dir, u"x.pdf"), prefer=missing[0])
        self.assertIn(missing[0], u"%s" % caught.exception)

    def test_plan_says_so_when_there_is_no_renderer_at_all(self):
        original = pdf_op.detect_renderers
        pdf_op.detect_renderers = lambda: [{"name": "word", "kind": "com",
                                            "available": False, "detail": "no"}]
        try:
            with self.assertRaises(pdf_op.PdfError) as caught:
                pdf_op.plan(self.path, os.path.join(self.dir, u"x.pdf"))
            self.assertIn(u"编排", u"%s" % caught.exception)
            self.assertIn(u"不自己渲染", u"%s" % caught.exception)
        finally:
            pdf_op.detect_renderers = original

    def test_a_missing_file_is_rejected(self):
        with self.assertRaises(pdf_op.PdfError):
            pdf_op.plan(os.path.join(self.dir, u"没有.docx"), os.path.join(self.dir, u"x.pdf"))

    def test_pages_are_counted_from_the_pdf_itself(self):
        """页数按交付物算：Word 自报 1、PDF /Count 26 的那次实测，就是这条用例的由来。"""
        fake = os.path.join(self.dir, u"假.pdf")
        with open(fake, "wb") as handle:
            handle.write(b"%PDF-1.7\n/Type /Pages /Count 26\n/Type /Page /Type /Page")
        self.assertEqual(pdf_op._pdf_page_count(fake), 26)

    def test_a_pdf_without_a_count_falls_back_to_counting_pages(self):
        fake = os.path.join(self.dir, u"假2.pdf")
        with open(fake, "wb") as handle:
            handle.write(b"%PDF-1.7\n/Type /Page\n/Type /Page\n/Type /Page")
        self.assertEqual(pdf_op._pdf_page_count(fake), 3)

    def test_a_missing_pdf_counts_as_unknown(self):
        self.assertIsNone(pdf_op._pdf_page_count(os.path.join(self.dir, u"没有.pdf")))


class TestRealExportIfAnyRenderer(unittest.TestCase):
    """真跑一次（会打开 Word/WPS，只读打开、导完不保存）。"""

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_pdfreal_")
        self.path = os.path.join(self.dir, u"夹具.docx")
        body = (fixtures.paragraph(fixtures.run(u"第一段。"))
                + fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>",
                                   u"<w:r><w:t>库容</w:t></w:r>"]]))
        fixtures.write_fixture(self.path, body=body)

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_export_writes_a_real_pdf(self):
        renderers = [item for item in pdf_op.detect_renderers() if item["available"]]
        if not renderers:
            self.skipTest(u"本机没有渲染器（Word/WPS/LibreOffice 都没装）")
        out = os.path.join(self.dir, u"出.pdf")
        report = pdf_op.export(self.path, out, prefer=renderers[0]["name"], visible=True)
        self.assertTrue(os.path.exists(out))
        with open(out, "rb") as handle:
            self.assertTrue(handle.read(5).startswith(b"%PDF"), u"要真是个 PDF")
        self.assertGreater(report["bytes"], 1000)
        self.assertIn(report["renderer"], ("word", "wps", "libreoffice"))


if __name__ == "__main__":
    unittest.main()
