# -*- coding: utf-8 -*-
"""宏「段落配方」两个算子中**读的一半**（重配）的单测。

判据来自参考宏的行为，不是照实现抄：
* 结果**追加到文档末尾**（原有段落一个都不许动）；
* 追加块 = 空段 + 空段 + `=== 重建段落 ===` + 正文行 + `=== 结束 ===`（`段落重配.bas:276`）；
* 变量值按 `EXCEL_FILE:`/`SHEET_NAME:` 定位、取 **B 列第 n+1 行**；
* 值不够 → `#数据缺失#`；只重写 `word/document.xml`。
"""

import os
import shutil
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.ops import recipe as recipe_op
from wordfactory.recipe import Recipe
from wordfactory.text import Paragraph
from wordfactory.xlsx import write_workbook

from . import fixtures

RECIPE_LINES = [
    u"=== 段落配方 [土方计算] ===",
    u"TEXT:1、",
    u"VAR:1|数据表.xlsx!Sheet1!B2",
    u"TEXT:本期",
    u"VAR:2|数据表.xlsx!Sheet1!B3",
    u"TEXT:万m",
    u"VAR:3|数据表.xlsx!Sheet1!B4",
    u"TEXT:",
    u"TEXT:2、",
    u"VAR:4|数据表.xlsx!Sheet1!B5",
    u"TEXT:设计工程量",
    u"EXCEL_FILE:数据表.xlsx",
    u"SHEET_NAME:Sheet1",
    u"VARIABLE_COUNT:4",
    u"=== 配方结束 ===",
]


def body_with_recipe(prefix=u"正文第一段。"):
    parts = [fixtures.paragraph(fixtures.run(prefix))]
    for line in RECIPE_LINES:
        parts.append(fixtures.paragraph(fixtures.run(line)))
    return u"".join(parts)


class RebuildCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_oprecipe_")
        self.docx = os.path.join(self.dir, u"带配方.docx")
        self.xlsx = os.path.join(self.dir, u"数据表.xlsx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body=None, values=None):
        fixtures.write_fixture(self.docx, body=body or body_with_recipe())
        rows = [(u"a", v) for v in (values if values is not None
                                    else [u"1.5", u"2.0", u"3.0", u"4.0"])]
        write_workbook(self.xlsx, u"Sheet1", rows)
        return self.docx

    def texts(self, path):
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        return [Paragraph(p).text for p in root.iter(qn("w:p"))]

    def parts_of(self, path):
        with zipfile.ZipFile(path) as archive:
            return dict((n, archive.read(n)) for n in archive.namelist())


class TestFindAndRebuild(RebuildCase):
    def test_the_recipe_is_found_in_the_document_tail(self):
        self.build()
        with Document(self.docx) as doc:
            text = recipe_op.find_recipe_text(doc)
        recipe = Recipe.parse(text)
        self.assertEqual(recipe.name, u"土方计算")
        self.assertEqual(recipe.variable_count, 4)

    def test_a_document_without_a_recipe_says_so(self):
        self.build(body=fixtures.paragraph(fixtures.run(u"无关正文")))
        with Document(self.docx) as doc:
            self.assertIsNone(recipe_op.find_recipe_text(doc))

    def test_rebuild_appends_and_leaves_everything_else_alone(self):
        self.build()
        before = self.texts(self.docx)
        out = os.path.join(self.dir, u"重建后.docx")
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
            values = recipe_op.values_for(recipe, self.xlsx)
            report = recipe_op.rebuild(doc, recipe, values)
            doc.save(out)
        self.assertEqual(report["values_used"], 4)
        self.assertEqual(report["missing"], 0)
        after = self.texts(out)
        self.assertEqual(after[:len(before)], before, u"原有段落一个都不许动")
        appended = after[len(before):]
        self.assertEqual(appended, [u"", u"", u"=== 重建段落 ===",
                                    u"", u"1、1.5本期2.0万m3.0", u"",
                                    u"2、4.0设计工程量", u"=== 结束 ==="])

    def test_only_document_xml_is_rewritten(self):
        self.build()
        out = os.path.join(self.dir, u"重建后.docx")
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
            recipe_op.rebuild(doc, recipe, recipe_op.values_for(recipe, self.xlsx))
            doc.save(out)
        before, after = self.parts_of(self.docx), self.parts_of(out)
        changed = sorted(n for n in after if after[n] != before.get(n))
        self.assertEqual(changed, ["word/document.xml"])

    def test_missing_values_become_the_placeholder(self):
        self.build(values=[u"1.5"])
        out = os.path.join(self.dir, u"重建后.docx")
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
            report = recipe_op.rebuild(doc, recipe, recipe_op.values_for(recipe, self.xlsx))
            doc.save(out)
        self.assertEqual(report["missing"], 3)
        self.assertIn(u"#数据缺失#", report["text"])

    def test_dry_run_writes_nothing(self):
        self.build()
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
            before = len(self.texts(self.docx))
            report = recipe_op.rebuild(doc, recipe, recipe_op.values_for(recipe, self.xlsx),
                                       dry_run=True)
            self.assertEqual(doc.dirty_parts, [])
            self.assertEqual(len(list(doc.part().iter(qn("w:p")))), before)
        self.assertTrue(report["dry_run"])

    def test_the_appended_paragraphs_are_plain_and_keep_spaces(self):
        self.build()
        out = os.path.join(self.dir, u"重建后.docx")
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
            recipe_op.rebuild(doc, recipe, recipe_op.values_for(recipe, self.xlsx))
            doc.save(out)
        with zipfile.ZipFile(out) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
        paragraphs = list(root.iter(qn("w:p")))
        node = paragraphs[-2].find(qn("w:r") + "/" + qn("w:t"))
        self.assertEqual(node.get("{http://www.w3.org/XML/1998/namespace}space"), "preserve",
                         u"带首尾空格的文本必须保留 xml:space")
        self.assertIsNone(paragraphs[-2].find(qn("w:pPr")), u"重建段落不该带额外格式")


class TestExcelLookup(RebuildCase):
    def test_the_data_table_is_looked_up_next_to_the_document(self):
        self.build()
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
        found = recipe_op.resolve_xlsx(recipe, self.docx)
        self.assertEqual(os.path.abspath(found), os.path.abspath(self.xlsx))

    def test_data_dir_wins_over_the_document_directory(self):
        self.build()
        other = os.path.join(self.dir, u"别处")
        os.makedirs(other)
        target = write_workbook(os.path.join(other, u"数据表.xlsx"), u"Sheet1", [(u"a", u"9")])
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
        self.assertEqual(os.path.abspath(recipe_op.resolve_xlsx(recipe, self.docx, other)),
                         os.path.abspath(target))

    def test_a_missing_data_table_lists_the_places_it_looked(self):
        self.build()
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
        os.remove(self.xlsx)
        with self.assertRaises(Exception) as caught:
            recipe_op.resolve_xlsx(recipe, self.docx)
        self.assertIn(u"找不到数据表", u"%s" % caught.exception)

    def test_the_sheet_name_from_the_recipe_is_used(self):
        """配方里写 Sheet1，数据表里叫别的名字 → 必须报错，不能"随便挑一个表"。"""
        fixtures.write_fixture(self.docx, body=body_with_recipe())
        write_workbook(self.xlsx, u"别的表", [(u"a", u"1")])
        with Document(self.docx) as doc:
            recipe = recipe_op.read_recipe(self.docx)
        with self.assertRaises(Exception):
            recipe_op.values_for(recipe, self.xlsx)


class TestLastRecipeWins(RebuildCase):
    def test_the_last_recipe_in_the_document_is_used(self):
        """生成端总是往末尾追加 —— 所以"最后一份"就是最新的那份。"""
        body = body_with_recipe() + u"".join(
            fixtures.paragraph(fixtures.run(line)) for line in
            [u"=== 段落配方 [新配方] ===", u"TEXT:甲",
             u"VAR:1|新的.xlsx!S!B2", u"EXCEL_FILE:新的.xlsx",
             u"SHEET_NAME:S", u"VARIABLE_COUNT:1", u"=== 配方结束 ==="])
        self.build(body=body)
        with Document(self.docx) as doc:
            recipe = Recipe.parse(recipe_op.find_recipe_text(doc))
        self.assertEqual(recipe.name, u"新配方")
        self.assertEqual(recipe.excel_file, u"新的.xlsx")

    def test_a_recipe_section_without_any_body_line_is_reported(self):
        """区间里光有头部字段、没有一行 `TEXT:`/`VAR:` → 多半是取错区间了，要说出来。"""
        body = u"".join(fixtures.paragraph(fixtures.run(line)) for line in
                        [u"=== 段落配方 [空配方] ===", u"EXCEL_FILE:数据表.xlsx",
                         u"SHEET_NAME:Sheet1", u"VARIABLE_COUNT:1", u"=== 配方结束 ==="])
        self.build(body=body)
        with Document(self.docx) as doc:
            with self.assertRaises(Exception) as caught:
                Recipe.parse(recipe_op.find_recipe_text(doc))
        self.assertIn(u"没有正文行", u"%s" % caught.exception)


if __name__ == "__main__":
    unittest.main()
