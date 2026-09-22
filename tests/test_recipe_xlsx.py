# -*- coding: utf-8 -*-
"""`.xlsx` 读写与「段落配方」文本格式的单测。

这里的判据全部来自 `docs/REFERENCE-MACROS.md` §3 的契约（不是照我的实现抄的）：
* xlsx：一个表、`A1=项目`/`B1=数值`、**第 n 个变量在第 n+1 行 B 列**；
* 配方文本：`TEXT:`/`VAR:` 行、三个头部字段、`?` 切行、空 `TEXT:` = 换行、
  `VAR:` 只数序号、数字开头的 TEXT 要补换行（`段落重配.bas:237-268`）。
"""

import os
import shutil
import tempfile
import unittest
import zipfile

from wordfactory.recipe import MISSING, Recipe, RecipeError
from wordfactory.xlsx import read_column_b, read_sheet, sheet_names, write_workbook

#: 规格书 §3.5 里的那份配方实例（逐字抄下来当夹具）
SPEC_RECIPE = u"""\r
=== 段落配方 [土方计算] ===\r
TEXT:1、\r
VAR:1|数据表.xlsx!Sheet1!B2\r
TEXT:本期\r
VAR:2|数据表.xlsx!Sheet1!B3\r
TEXT:万m\r
VAR:3|数据表.xlsx!Sheet1!B4\r
TEXT:\r
TEXT:2、\r
VAR:4|数据表.xlsx!Sheet1!B5\r
TEXT:设计工程量\r
EXCEL_FILE:数据表.xlsx\r
SHEET_NAME:Sheet1\r
VARIABLE_COUNT:4\r
=== 配方结束 ===\r
"""


class RecipeCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_recipe_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)


class TestTheSpecExample(RecipeCase):
    def test_the_spec_recipe_parses(self):
        recipe = Recipe.parse(SPEC_RECIPE)
        self.assertEqual(recipe.name, u"土方计算")
        self.assertEqual(recipe.excel_file, u"数据表.xlsx")
        self.assertEqual(recipe.sheet_name, u"Sheet1")
        self.assertEqual(recipe.variable_count, 4)
        self.assertEqual([kind for kind, _ in recipe.lines],
                         ["TEXT", "VAR", "TEXT", "VAR", "TEXT", "VAR", "TEXT", "TEXT", "VAR", "TEXT"])

    def test_reconstruct_matches_the_spec_semantics(self):
        """按 `段落重配.bas:237-268` 逐行推：

        * `previousLineWasText` 初值是 False，而第一行 `TEXT:1、` 以数字开头 → **最前面会多一个换行**
          （参考宏是 `vbCrLf & vbCrLf & "=== 重建段落 ===" & ... & outputText`，多出来的是空段落）；
        * 空 `TEXT:` = 一个换行；`VAR:` 取值的顺序 = 它在配方里出现的顺序。
        """
        recipe = Recipe.parse(SPEC_RECIPE)
        text = recipe.reconstruct([u"1.0", u"2.0", u"3.0", u"4.0"])
        self.assertEqual(text, u"\n1、1.0本期2.0万m3.0\n\n2、4.0设计工程量")

    def test_reconstruct_reports_missing_values(self):
        """变量比 Excel 里的值多 → 补 `#数据缺失#`（`段落重配.bas:261`）。"""
        recipe = Recipe.parse(SPEC_RECIPE)
        text = recipe.reconstruct([u"1.0"])
        self.assertEqual(text.count(MISSING), 3)
        self.assertIn(u"1.0", text)

    def test_text_that_starts_with_a_number_gets_a_break_only_after_a_var(self):
        recipe = Recipe.parse(SPEC_RECIPE)
        # 变量之后的 "2、" 要换行；紧跟 TEXT 的 "设计工程量" 不换行
        text = recipe.reconstruct([u"A", u"B", u"C", u"D"])
        self.assertTrue(text.endswith(u"D设计工程量"))
        self.assertIn(u"\n2、", text)

    def test_round_trip_keeps_the_contract(self):
        """写出去再读回来，契约三项（文件/表名/变量数）与正文行序都不变。"""
        recipe = Recipe.parse(SPEC_RECIPE)
        again = Recipe.parse(recipe.to_text())
        self.assertEqual((again.excel_file, again.sheet_name, again.variable_count,
                          [k for k, _ in again.lines]),
                         (recipe.excel_file, recipe.sheet_name, recipe.variable_count,
                          [k for k, _ in recipe.lines]))
        self.assertEqual(again.name, recipe.name)

    def test_the_var_number_in_the_line_is_ignored(self):
        """读端只数序号、不看行里写的编号（`段落重配.bas:255-262`）。"""
        text = SPEC_RECIPE.replace(u"VAR:1|数据表.xlsx!Sheet1!B2", u"VAR:99|别的.xlsx!别的!B7")
        recipe = Recipe.parse(text)
        self.assertEqual(recipe.reconstruct([u"A", u"B", u"C", u"D"]),
                         u"\n1、A本期B万mC\n\n2、D设计工程量")

    def test_any_line_ending_works(self):
        for newline in (u"\r\n", u"\r", u"\n"):
            recipe = Recipe.parse(SPEC_RECIPE.replace(u"\r\n", newline))
            self.assertEqual(recipe.variable_count, 4)


class TestRecipeErrors(RecipeCase):
    def test_missing_header_fields_are_reported_in_words(self):
        with self.assertRaises(RecipeError) as caught:
            Recipe.parse(u"=== 段落配方 [x] ===\r\nTEXT:abc\r\n=== 配方结束 ===")
        message = u"%s" % caught.exception
        self.assertIn(u"EXCEL_FILE", message)
        self.assertIn(u"SHEET_NAME", message)
        self.assertIn(u"VARIABLE_COUNT", message)

    def test_a_count_that_does_not_match_the_body_is_reported(self):
        text = SPEC_RECIPE.replace(u"VARIABLE_COUNT:4", u"VARIABLE_COUNT:2")
        with self.assertRaises(RecipeError) as caught:
            Recipe.parse(text)
        self.assertIn(u"变量数对不上", u"%s" % caught.exception)

    def test_lines_outside_the_section_are_ignored(self):
        text = u"前面还有别的文字\r\nTEXT:这句不属于配方\r\n" + SPEC_RECIPE
        recipe = Recipe.parse(text)
        self.assertEqual(len([1 for k, _ in recipe.lines if k == "TEXT"]), 6, recipe.lines)


class TestXlsx(RecipeCase):
    def write(self, rows, sheet=u"Sheet1", name=None):
        return write_workbook(name or os.path.join(self.dir, u"数据表.xlsx"), sheet, rows)

    def test_round_trip_the_contract(self):
        path = self.write([(u"1、", u"1.5"), (u"本期", u"2.5")])
        self.assertEqual(sheet_names(path), [u"Sheet1"])
        cells = read_sheet(path)
        self.assertEqual(cells[(1, 1)], u"项目")
        self.assertEqual(cells[(1, 2)], u"数值")
        self.assertEqual(cells[(2, 1)], u"1、")
        self.assertEqual(cells[(2, 2)], u"1.5")
        self.assertEqual(read_column_b(path), [u"1.5", u"2.5"])

    def test_the_variable_index_maps_to_row_plus_one(self):
        """契约里最容易错的一条：第 n 个变量在 **B(n+1)**。"""
        path = self.write([(u"a", u"v1"), (u"b", u"v2"), (u"c", u"v3")])
        values = read_column_b(path)
        for index, value in enumerate(values, start=1):
            with zipfile.ZipFile(path) as archive:
                pass
            self.assertEqual(read_sheet(path)[(index + 1, 2)], value)

    def test_empty_cells_become_empty_strings(self):
        path = self.write([(u"a", u""), (u"b", u"v2")])
        self.assertEqual(read_column_b(path), [u"", u"v2"])

    def test_the_sheet_name_is_used(self):
        path = self.write([(u"a", u"1")], sheet=u"土方")
        self.assertEqual(sheet_names(path), [u"土方"])
        self.assertEqual(read_sheet(path, u"土方")[(2, 2)], u"1")
        with self.assertRaises(Exception):
            read_sheet(path, u"不存在的表")

    def test_special_characters_survive(self):
        path = self.write([(u"a&b<c>\"d\"", u"e'f")])
        cells = read_sheet(path)
        self.assertEqual(cells[(2, 1)], u"a&b<c>\"d\"")
        self.assertEqual(cells[(2, 2)], u"e'f")

    def test_a_multiline_value_is_not_silently_mangled(self):
        """`cleanText` 会把配方文本里的换行去掉，所以值里不该有换行；真出现时也不能丢内容。"""
        path = self.write([(u"a", u"第一行第二行")])
        self.assertEqual(read_column_b(path), [u"第一行第二行"])

    def test_the_file_is_a_valid_zip_with_the_expected_parts(self):
        path = self.write([(u"a", u"1")])
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
        self.assertEqual(names, {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                                 "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                                 "xl/sharedStrings.xml", "xl/worksheets/sheet1.xml"})

    def test_two_writes_are_byte_identical(self):
        """和 docx 一样：输出要可复现（zip 时间戳固定）。"""
        import time
        first = self.write([(u"a", u"1")], name=os.path.join(self.dir, u"一.xlsx"))
        time.sleep(1.1)
        second = self.write([(u"a", u"1")], name=os.path.join(self.dir, u"二.xlsx"))
        with open(first, "rb") as h1, open(second, "rb") as h2:
            self.assertEqual(h1.read(), h2.read())

    def test_what_the_reference_macro_would_read(self):
        """把"参考宏会读到什么"按它的取值规则还原一遍：`B(n+1)`、空值给空串。

        值序：`1.5`（B2）、空（B3）、`300`（B4）→ 第 4 个变量超出范围 → `#数据缺失#`。
        """
        path = self.write([(u"1、", u"1.5"), (u"本期", u""), (u"万m", u"300")])
        values = read_column_b(path, max_rows=3)
        self.assertEqual(values, [u"1.5", u"", u"300"])
        recipe = Recipe.parse(SPEC_RECIPE)
        text = recipe.reconstruct(values)
        self.assertEqual(text, u"\n1、1.5本期万m300\n\n2、#数据缺失#设计工程量")


if __name__ == "__main__":
    unittest.main()
