# -*- coding: utf-8 -*-
"""配方引擎（`pipeline.py`）的单测。

要钉住的：

* **顺序就是用户排的顺序**（不是代码写死的）—— 同一个配方，顺序不同结果要不同；
* 模式收尾：``verify`` = 改动标蓝；``formal`` = 通体黑 + 字体合规 + 体检；
* **「字体/颜色」不是配方里的一步**（否则验证版会被全部改黑把标蓝盖掉）；
* ``dry_run`` 一个字节都不写（上下标那步是在副本上跑的）；
* 配方里不认识的步骤要报错，且说清有哪些；
* 输出文件**不覆盖**已有文件（加序号）。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory import pipeline
from wordfactory.ooxml import qn

from . import fixtures


class PipelineCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_pipe_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body=None):
        body = body or u"".join([
            fixtures.paragraph(fixtures.run(u"1、本期"), ),
        ])
        # 一个带题注 + 表格的文档
        body = (fixtures.paragraph(fixtures.run(u"表2.3-1     库容特性表"))
                + fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>",
                                   u"<w:r><w:t>库容</w:t></w:r>"]]))
        return fixtures.write_fixture(self.path, body=body)

    def out(self, name=u"出.docx"):
        return os.path.join(self.dir, name)


class TestRecipe(PipelineCase):
    def test_steps_run_in_the_given_order(self):
        recipe = pipeline.normalize_recipe(["tidy", "captions", "sup"])
        self.assertEqual([step["op"] for step in recipe], ["tidy", "captions", "sup"])

    def test_params_merge_over_the_defaults(self):
        recipe = pipeline.normalize_recipe([{"op": "captions",
                                            "params": {"figure_space": 3}}])
        self.assertEqual(recipe[0]["params"]["figure_space"], 3)
        self.assertEqual(recipe[0]["params"]["center_table"], True,
                         u"没给的参数仍取默认值")

    def test_an_unknown_step_is_rejected_with_the_list(self):
        with self.assertRaises(pipeline.PipelineError) as caught:
            pipeline.normalize_recipe(["nope"])
        message = u"%s" % caught.exception
        self.assertIn(u"captions", message, u"报错里要列出已注册的宏")
        self.assertIn(u"nope", message)

    def test_an_empty_recipe_is_allowed(self):
        self.assertEqual(pipeline.normalize_recipe([]), [])


class TestModes(PipelineCase):
    def test_verify_marks_the_change_blue(self):
        self.build()
        report = pipeline.run_pipeline(self.path, ["captions"], mode="verify",
                                       out_path=self.out())
        self.assertEqual(report["finish"], u"改动标蓝（1 处）")
        from wordfactory.document import Document
        with Document(self.out()) as doc:
            colors = [node.get(qn("w:val")) for node in doc.part().iter(qn("w:color"))]
        self.assertIn("0000FF", colors, u"验证版要把改动标蓝")

    def test_formal_makes_it_black_and_audits(self):
        self.build()
        report = pipeline.run_pipeline(self.path, ["captions"], mode="formal",
                                       out_path=self.out())
        self.assertEqual(report["audit"]["verdict"], "PASS")
        from wordfactory.document import Document
        with Document(self.out()) as doc:
            colors = [node.get(qn("w:val")) for node in doc.part().iter(qn("w:color"))]
        self.assertNotIn("0000FF", colors, u"正式版不许留蓝色")
        self.assertIn("000000", colors)

    def test_fonts_is_not_a_recipe_step(self):
        self.assertNotIn(u"fonts", pipeline.STEPS,
                         u"字体/颜色是收尾，不是配方步骤（否则验证版会被改黑）")

    def test_dry_run_writes_nothing(self):
        self.build()
        before = os.path.getsize(self.path)
        report = pipeline.run_pipeline(self.path, ["captions"], mode="formal", dry_run=True)
        self.assertTrue(report["dry_run"])
        self.assertIsNone(report.get("out"))
        self.assertEqual(os.path.getsize(self.path), before, u"输入文件一个字节都不该动")

    def test_nothing_to_do_writes_nothing(self):
        self.build()
        first = pipeline.run_pipeline(self.path, ["captions"], mode="verify",
                                      out_path=self.out(u"一.docx"))
        second = pipeline.run_pipeline(first["out"], ["captions"], mode="verify",
                                       out_path=self.out(u"二.docx"))
        self.assertIn(u"没有需要改的地方", second["finish"])
        self.assertIsNone(second.get("out"))


class TestOutputNaming(PipelineCase):
    def test_the_output_is_named_after_the_mode_and_never_overwrites(self):
        self.build()
        first = pipeline.run_pipeline(self.path, ["captions"], mode="formal")
        self.assertTrue(first["out"].endswith(u"（正式版）.docx"), first["out"])
        again = pipeline.run_pipeline(self.path, ["captions"], mode="formal")
        self.assertNotEqual(again["out"], first["out"], u"不许覆盖上一次的输出")
        self.assertTrue(os.path.exists(again["out"]))


class TestFormatReport(PipelineCase):
    def test_the_report_names_the_steps_and_the_finish(self):
        self.build()
        report = pipeline.run_pipeline(self.path, ["captions", "tidy"], mode="formal",
                                       out_path=self.out())
        text = pipeline.format_report(report)
        self.assertIn(u"captions → tidy", text)
        self.assertIn(u"AUDIT=PASS", text)
        self.assertIn(self.out(), text)


if __name__ == "__main__":
    unittest.main()
