# -*- coding: utf-8 -*-
"""宏「段落配方」写的一半（生成）的单测。

判据来自参考宏 `段落配方生成器.bas`：
* 模式 1（高亮）：非高亮→高亮 吐 `TEXT:`，高亮→非高亮 吐 `VAR:`，值 = 高亮片段原文；
  段落标记也进字符流（所以跨段的 TEXT 会在配方里折成"无前缀行"）；
* 模式 2（特定字符）：**值就是占位符本身**（`:315`），不是该位置的原文；空段落写 `TEXT:`、段间补 `TEXT:`；
* A 列 = 该变量前面那段 TEXT（给人看的），B 列 = 值；
* 生成完把配方文本**追加到文档末尾**。

还有一条端到端不变量：**生成 → 重配**应当还原出模板段落本身（值填回去）。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.document import Document
from wordfactory.ops import recipe as recipe_op
from wordfactory.recipe import Recipe
from wordfactory.text import Paragraph

from . import fixtures


def run(text, highlight=None):
    if highlight:
        return u'<w:r><w:rPr><w:highlight w:val="%s"/></w:rPr><w:t>%s</w:t></w:r>' \
            % (highlight, text)
    return fixtures.run(text)


class GenCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_gen_")
        self.docx = os.path.join(self.dir, u"模板.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body):
        return fixtures.write_fixture(self.docx, body=body)

    def gen(self, **overrides):
        options = {"mode": "highlight", "char": None, "name": u"土方计算",
                   "excel_file": u"数据表.xlsx", "sheet_name": u"Sheet1",
                   "out_xlsx": os.path.join(self.dir, u"数据表.xlsx"),
                   "append": True, "trim_last_char": False}
        options.update(overrides)
        with Document(self.docx) as doc:
            report, recipe, xlsx = recipe_op.generate(doc, options)
            out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, recipe, xlsx, out

    def texts(self, path):
        import zipfile
        from xml.etree import ElementTree as ET
        from wordfactory.ooxml import qn
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        return [Paragraph(p).text for p in root.iter(qn("w:p"))]


class TestHighlightMode(GenCase):
    def test_two_highlighted_spans_become_two_variables(self):
        body = u"".join([
            fixtures.paragraph(run(u"1、本期"), run(u"3.5", "yellow"), run(u"万m")),
            fixtures.paragraph(run(u"2、设计工程量"), run(u"12.8", "yellow"), run(u"万m")),
        ])
        self.build(body)
        report, recipe, xlsx, _ = self.gen()
        self.assertEqual(report["variables"], 2)
        self.assertEqual([value for _, value in report["rows"]], [u"3.5", u"12.8"])
        # 第 2 个变量的前缀跨了段（"万m" + 段落标记 + "2、设计工程量"）→ 折成 TEXT + RAW 两行
        self.assertEqual([kind for kind, _ in recipe.lines],
                         ["TEXT", "VAR", "TEXT", "RAW", "VAR", "TEXT"])

    def test_the_value_is_the_highlighted_text_itself(self):
        """默认不照抄"隐形修正"：高亮 `3.5` 就存 `3.5`（照抄会变成 `3.`）。"""
        body = fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow"), run(u"万m"))
        self.build(body)
        report, _, _, _ = self.gen()
        self.assertEqual(report["rows"], [(u"本期", u"3.5")])

    def test_the_trim_switch_reproduces_the_macro_quirk(self):
        body = fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow"), run(u"万m"))
        self.build(body)
        report, _, _, _ = self.gen(trim_last_char=True)
        self.assertEqual(report["rows"], [(u"本期", u"3.")])
        # 被"修正"掉的那个字符按参考宏的写法归到后面的文本（`rng.Start = correctedEnd`）
        self.assertIn(u"5万m", [content for kind, content in
                               recipe_op.scan_highlights(Document(self.docx), True)
                               if kind == "TEXT"])

    def test_the_prefix_column_holds_the_text_before_the_variable(self):
        body = u"".join([
            fixtures.paragraph(run(u"1、本期"), run(u"3.5", "yellow"), run(u"万m")),
            fixtures.paragraph(run(u"2、设计"), run(u"12.8", "yellow"), run(u"万m")),
        ])
        self.build(body)
        report, _, _, _ = self.gen()
        self.assertEqual(report["rows"][0][0], u"1、本期")
        self.assertTrue(report["rows"][1][0].startswith(u"万m"), report["rows"][1][0])

    def test_a_paragraph_mark_between_texts_keeps_the_paragraph_split(self):
        """两段非高亮文本之间的段落标记要留在 TEXT 里 → 配方里折成一行"无前缀行"。"""
        body = u"".join([
            fixtures.paragraph(run(u"第一段")),
            fixtures.paragraph(run(u"第二段"), run(u"9", "yellow")),
        ])
        self.build(body)
        report, recipe, _, _ = self.gen()
        kinds = [kind for kind, _ in recipe.lines]
        self.assertEqual(kinds, ["TEXT", "RAW", "VAR"], recipe.lines)
        self.assertEqual(recipe.lines[2][1], u"9")
        # 重配出来应当是"第一段 \n 第二段 9"
        self.assertEqual(recipe.reconstruct([u"X"]), u"第一段\n第二段X")

    def test_a_document_without_highlights_yields_no_variables(self):
        self.build(fixtures.paragraph(run(u"没有高亮")))
        report, _, _, _ = self.gen()
        self.assertEqual(report["variables"], 0)

    def test_the_recipe_is_appended_to_the_document(self):
        body = fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow"))
        before = None
        self.build(body)
        report, recipe, _, out = self.gen()
        texts = self.texts(out)
        self.assertIn(u"=== 段落配方 [土方计算] ===", texts)
        self.assertIn(u"EXCEL_FILE:数据表.xlsx", texts)
        self.assertIn(u"VARIABLE_COUNT:1", texts)
        self.assertIn(u"=== 配方结束 ===", texts)
        self.assertGreater(report["paragraphs"], 0)

    def test_no_append_writes_only_the_xlsx(self):
        self.build(fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow")))
        report, _, xlsx, out = self.gen(append=False)
        self.assertEqual(report["paragraphs"], 0)
        self.assertTrue(os.path.exists(xlsx))
        self.assertNotIn(u"=== 段落配方 [土方计算] ===", self.texts(out))

    def test_an_existing_recipe_section_is_not_scanned_again(self):
        """否则第二次生成会把上一份配方也当成正文（它的开头也有高亮？不，但它的文字会污染 TEXT）。"""
        body = u"".join([
            fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow")),
            fixtures.paragraph(run(u"=== 段落配方 [旧] ===")),
            fixtures.paragraph(run(u"TEXT:旧内容")),
            fixtures.paragraph(run(u"=== 配方结束 ===")),
        ])
        self.build(body)
        report, _, _, _ = self.gen()
        self.assertEqual(report["variables"], 1)
        self.assertNotIn(u"旧内容", u"".join(text for _, text in report["rows"]))


class TestSpecialCharMode(GenCase):
    def test_each_placeholder_is_a_variable_whose_value_is_the_placeholder(self):
        body = u"".join([
            fixtures.paragraph(run(u"1、本期xx万m")),
            fixtures.paragraph(run(u"2、设计工程量xx万m")),
        ])
        self.build(body)
        report, _, _, _ = self.gen(mode="chars", char=u"xx")
        self.assertEqual(report["variables"], 2)
        self.assertEqual([value for _, value in report["rows"]], [u"xx", u"xx"])
        self.assertEqual([prefix for prefix, _ in report["rows"]],
                         [u"1、本期", u"2、设计工程量"])

    def test_paragraph_boundaries_get_text_markers(self):
        body = u"".join([fixtures.paragraph(run(u"甲xx")), fixtures.paragraph(run(u"乙"))])
        self.build(body)
        report, recipe, _, _ = self.gen(mode="chars", char=u"xx")
        self.assertEqual([kind for kind, _ in recipe.lines][:2], ["TEXT", "VAR"])
        self.assertIn(("TEXT", u""), recipe.lines, u"段间要补一个空 TEXT:")

    def test_an_empty_paragraph_becomes_an_empty_text_line(self):
        body = u"<w:p/>" + fixtures.paragraph(run(u"甲xx"))
        self.build(body)
        _, recipe, _, _ = self.gen(mode="chars", char=u"xx")
        self.assertEqual(recipe.lines[0], ("TEXT", u""))

    def test_no_placeholder_says_so(self):
        self.build(fixtures.paragraph(run(u"没有占位符")))
        with self.assertRaises(Exception):
            self.gen(mode="chars", char=u"")


class TestRoundTrip(GenCase):
    def test_generate_then_rebuild_reproduces_the_template(self):
        """**端到端不变量**：生成配方（值 = 原文）后重配，应当还原出模板段落本身。"""
        body = u"".join([
            fixtures.paragraph(run(u"1、本期"), run(u"3.5", "yellow"), run(u"万m")),
            fixtures.paragraph(run(u"2、工程量"), run(u"12.8", "yellow"), run(u"万m")),
        ])
        self.build(body)
        _, _, xlsx, out = self.gen()
        texts = self.texts(out)
        template = u"\n".join(texts[:texts.index(u"=== 段落配方 [土方计算] ===")])
        rebuilt = recipe_op.read_recipe(out)
        text = rebuilt.reconstruct(recipe_op.values_for(rebuilt, xlsx))
        self.assertEqual(text.lstrip(u"\n"), template, u"重配结果应当等于模板正文")

    def test_generate_then_rebuild_with_new_numbers(self):
        """换成新数据 → 重配出来的就是新数字，模板骨架不变。"""
        body = fixtures.paragraph(run(u"本期"), run(u"3.5", "yellow"), run(u"万m"))
        self.build(body)
        _, _, _, out = self.gen()
        from wordfactory.xlsx import write_workbook
        write_workbook(os.path.join(self.dir, u"新数据.xlsx"), u"Sheet1", [(u"本期", u"99.9")])
        recipe = recipe_op.read_recipe(out)
        values = recipe_op.values_for(recipe, os.path.join(self.dir, u"新数据.xlsx"))
        self.assertEqual(recipe.reconstruct(values), u"本期99.9万m")


if __name__ == "__main__":
    unittest.main()
