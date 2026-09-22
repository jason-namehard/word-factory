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


class TestTwoCompleteRecipes(RebuildCase):
    """一份文档里放**两份完整配方**时会怎样 —— 明确记下来，免得以后当成 bug 猜。

    参考宏的读法是状态机：所有 `开`…`闭` 区间里的正文行**都收**（按出现顺序），
    头部三项在全文里扫、后出现的覆盖先出现的。外置工具没有选区（等价于"整篇当选区"），
    所以行为与宏一致：**两份会被并成一份**（变量翻倍）。
    → 结论：**别把两份配方放同一个文档**。真实文档里那行多余的标题在同一个区间内，
    不受影响（见 `TestTheRealDocumentShape`）。
    """

    def test_two_bodies_are_merged_the_way_the_macro_would(self):
        def section(name, excel, prefix):
            return [u"=== 段落配方 [%s] ===" % name, u"TEXT:%s" % prefix,
                    u"VAR:1|%s!S!B2" % excel, u"EXCEL_FILE:%s" % excel,
                    u"SHEET_NAME:S", u"VARIABLE_COUNT:1", u"=== 配方结束 ==="]

        body = u"".join(fixtures.paragraph(fixtures.run(line))
                        for line in section(u"甲", u"甲.xlsx", u"A")
                        + section(u"乙", u"乙.xlsx", u"B"))
        self.build(body=body)
        with Document(self.docx) as doc:
            recipe = Recipe.parse(recipe_op.find_recipe_text(doc))
        self.assertEqual([kind for kind, _ in recipe.lines], ["TEXT", "VAR", "TEXT", "VAR"])
        self.assertEqual([content for kind, content in recipe.lines if kind == "TEXT"],
                         [u"A", u"B"])
        self.assertEqual(recipe.excel_file, u"乙.xlsx",
                         u"头部三项在全文里扫，后出现的覆盖先出现的（与宏一致）")

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


class TestTheRealDocumentShape(RebuildCase):
    """用户 2026-09-22 给的金标准样本暴露的形态（`段落重配.docx` 的真实样子）：

    配方正文在文档**开头**，中间又夹着一行 `=== 段落配方 [名] ===`（用户粘贴留下的），
    三个头部字段与 `=== 配方结束 ===` 在它后面。之前"取最后一次出现"的启发式会刚好取到
    那个只有头部、没有正文的残块 → 解析直接失败。

    参考宏是**状态机**：一对 `开`…`闭` 之间的行才算正文，期间再出现开标记只是重新置位
    （`段落重配.bas:219-229`）。外置工具没有选区，等价规则 = "整篇当选区"。
    """

    def body(self):
        lines = ([u"=== 段落配方 [段落重配] ===", u"TEXT:一、量算", u"流域面积F=",
                  u"VAR:1|数据表.xlsx!Sheet1!B2", u"TEXT:km2，干流长度L=",
                  u"VAR:2|数据表.xlsx!Sheet1!B3", u"TEXT:m/m。", u"TEXT:h"]
                 + [u"=== 段落配方 [段落重配] ===", u"", u"",
                    u"EXCEL_FILE:数据表.xlsx", u"SHEET_NAME:Sheet1",
                    u"VARIABLE_COUNT:2", u"=== 配方结束 ===", u"",
                    u"=== 重建段落 ==="])
        return u"".join(fixtures.paragraph(fixtures.run(line)) if line else u"<w:p/>"
                        for line in lines)

    def test_the_whole_span_is_taken_so_the_body_is_not_lost(self):
        self.build(body=self.body())
        with Document(self.docx) as doc:
            recipe = Recipe.parse(recipe_op.find_recipe_text(doc))
        self.assertEqual(recipe.variable_count, 2)
        self.assertEqual([kind for kind, _ in recipe.lines],
                         ["TEXT", "RAW", "VAR", "TEXT", "VAR", "TEXT", "TEXT"])

    def test_the_first_title_wins(self):
        self.build(body=self.body())
        with Document(self.docx) as doc:
            recipe = Recipe.parse(recipe_op.find_recipe_text(doc))
        self.assertEqual(recipe.name, u"段落重配")

    def test_the_span_starts_at_the_first_open_marker(self):
        body = fixtures.paragraph(fixtures.run(u"前面还有无关正文")) + self.body()
        self.build(body=body)
        with Document(self.docx) as doc:
            text = recipe_op.find_recipe_text(doc)
        self.assertNotIn(u"无关正文", text)
        self.assertTrue(text.startswith(u"=== 段落配方"))
