# -*- coding: utf-8 -*-
"""表格款式（`rules/tablestyle.json`）的单测。

用户 2026-09-22 的要求里有两件事必须钉住：
* **采集**（`capture`）：一个**已经调好的表格**能被读成款式数据 —— 这是"调一次、以后复用"的前提；
* **预览**（`build_preview`）：造一份**新建的** docx，里面每种候选款式各渲染一张表，
  且**预览与真跑走同一条套用逻辑**（不是两套代码）。
另外要钉的是工程纪律：款式文件里的错键/错值要被 `check()` 点出来；套用**幂等**（第二遍报 0 处）；
`dry_run` 不写一个字节。
"""

import os
import shutil
import tempfile
import unittest
import zipfile
from xml.etree import ElementTree as ET

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.tablestyle import (DEFAULT_STYLES, StyleSet, TableStyle, TableStyleError,
                                    apply, build_preview, capture, parse_indexes, tables_of)

from . import fixtures


def cell(text, **attrs):
    pr = u"<w:tcPr/>" if not attrs else u"<w:tcPr><w:vAlign w:val=\"%s\"/></w:tcPr>" % attrs["v_align"]
    return u"<w:tc>%s<w:p><w:r><w:t>%s</w:t></w:r></w:p></w:tc>" % (pr, text)


def row(*cells):
    return u"<w:tr>%s</w:tr>" % u"".join(cells)


def table(rows, columns=2, widths=(2000, 4000)):
    grid = u"".join(u'<w:gridCol w:w="%d"/>' % w for w in widths[:columns])
    return (u'<w:tbl><w:tblPr><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
            % (grid, u"".join(rows)))


class StyleCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_tablestyle_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, body):
        fixtures.write_fixture(self.path, body=body)
        return self.path

    def style(self, name=u"三线表"):
        return StyleSet(DEFAULT_STYLES).get(name)

    def first_table(self, path):
        with Document(path) as doc:
            return list(tables_of(doc))[0]


class TestApply(StyleCase):
    def test_a_three_line_table_gets_its_borders_and_header(self):
        body = table([row(cell(u"项目"), cell(u"数值")), row(cell(u"面积"), cell(u"0.8"))])
        self.build(body)
        out = os.path.join(self.dir, u"出.docx")
        with Document(self.path) as doc:
            report = apply(doc, self.style())
            doc.save(out)
        self.assertEqual(report["changed_tables"], 1)
        self.assertIn("框线", report["changes"])
        table_element = self.first_table(out)
        pr = table_element.find(qn("w:tblPr"))
        self.assertEqual(pr.find(qn("w:jc")).get(qn("w:val")), "center")
        borders = pr.find(qn("w:tblBorders"))
        self.assertEqual(borders.find(qn("w:top")).get(qn("w:sz")), "12")
        self.assertEqual(borders.find(qn("w:insideV")).get(qn("w:val")), "none")
        # 表头下框线 + 跨页重复 + 加粗居中
        header_cell = [element for element in table_element.iter(qn("w:tc"))][0]
        tcpr = header_cell.find(qn("w:tcPr"))
        self.assertEqual(tcpr.find(qn("w:vAlign")).get(qn("w:val")), "center")
        self.assertIsNotNone(tcpr.find(qn("w:tcBorders")))
        first_paragraph = next(iter(table_element.iter(qn("w:p"))))
        self.assertEqual(first_paragraph.find(qn("w:pPr")).find(qn("w:jc")).get(qn("w:val")),
                         "center")
        self.assertIsNotNone(next(iter(table_element.iter(qn("w:r")))).find(qn("w:rPr")).find(qn("w:b")))
        self.assertIsNotNone(table_element.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader")))

    def test_the_child_order_stays_legal(self):
        """`w:tblPr` 是**序列**，插错位置 Word 会嫌文件坏 —— 这里按 schema 顺序核对。"""
        body = table([row(cell(u"A"), cell(u"B"))])
        self.build(body)
        with Document(self.path) as doc:
            apply(doc, self.style())
        table_element = self.body_table(document=doc)
        order = [child.tag for child in table_element.find(qn("w:tblPr"))]
        expected = [qn("w:tblW"), qn("w:jc"), qn("w:tblBorders"), qn("w:tblLayout"),
                    qn("w:tblCellMar"), qn("w:tblLook")]
        positions = [expected.index(tag) for tag in order if tag in expected]
        self.assertEqual(positions, sorted(positions),
                         u"tblPr 的子元素顺序不合法：%s" % u", ".join(order))

    def body_table(self, document=None, path=None):
        with Document(path or self.path) as doc:
            return list(tables_of(doc))[0]

    def test_running_twice_changes_nothing(self):
        body = table([row(cell(u"A"), cell(u"B")), row(cell(u"C"), cell(u"D"))])
        self.build(body)
        with Document(self.path) as doc:
            first = apply(doc, self.style())
            second = apply(doc, self.style())
        self.assertGreater(first["total"], 0)
        self.assertEqual(second["total"], 0, u"同一款式跑第二遍必须是 0 处（幂等）")

    def test_dry_run_writes_nothing(self):
        body = table([row(cell(u"A"), cell(u"B"))])
        self.build(body)
        with Document(self.path) as doc:
            before = len(list(doc.part().iter(qn("w:tblBorders"))))
            report = apply(doc, self.style(), dry_run=True)
            self.assertEqual(doc.dirty_parts, [])
            self.assertEqual(len(list(doc.part().iter(qn("w:tblBorders")))), before)
        self.assertGreater(report["total"], 0, u"dry-run 要报出\"会改多少\"")

    def test_the_selector_picks_tables(self):
        body = (table([row(cell(u"1"), cell(u"2"))]) + fixtures.paragraph(fixtures.run(u"间隔"))
                + table([row(cell(u"3"), cell(u"4"))]))
        self.build(body)
        with Document(self.path) as doc:
            report = apply(doc, self.style(), selector=u"2")
        self.assertEqual(report["selected"], 1)
        self.assertEqual(report["tables_changed"], [2])

    def test_equal_widths_rewrites_the_grid(self):
        body = table([row(cell(u"A"), cell(u"B"))], widths=(2000, 4000))
        self.build(body)
        with Document(self.path) as doc:
            report = apply(doc, self.style(u"三线表_等宽"))
        table_element = list(tables_of(doc))[0]
        widths = [int(column.get(qn("w:w"))) for column in table_element.find(qn("w:tblGrid"))]
        self.assertEqual(widths[0], widths[1], u"等宽款要把各列拉平")
        self.assertIn("列宽", report["changes"])

    def test_indexes_parsing(self):
        self.assertEqual(parse_indexes("all", 3), [1, 2, 3])
        self.assertEqual(parse_indexes("2", 3), [2])
        self.assertEqual(parse_indexes("1,3-4", 5), [1, 3, 4])
        self.assertEqual(parse_indexes("9", 3), [])
        with self.assertRaises(TableStyleError):
            parse_indexes("x", 3)


class TestStyleFile(StyleCase):
    def test_bad_keys_are_reported(self):
        style = TableStyle(u"x", {"boders": {}, "table_align": "middle", "column_widths": "wide"})
        problems = style.check()
        self.assertTrue([p for p in problems if u"不认识的键" in p], problems)
        self.assertTrue([p for p in problems if u"table_align" in p], problems)
        self.assertTrue([p for p in problems if u"column_widths" in p], problems)

    def test_a_bad_edge_or_line_type_is_reported(self):
        style = TableStyle(u"x", {"borders": {"diagonal": {"val": "single"},
                                              "top": {"val": "wavy", "sz": "粗"}}})
        problems = style.check()
        self.assertTrue([p for p in problems if u"不认识的边" in p], problems)
        self.assertTrue([p for p in problems if u"线型" in p], problems)
        self.assertTrue([p for p in problems if u"必须是整数" in p], problems)

    def test_round_trip_through_the_file(self):
        path = os.path.join(self.dir, u"tablestyle.json")
        StyleSet(DEFAULT_STYLES).save(path)
        again = StyleSet.load(path)
        self.assertEqual(sorted(again.styles), sorted(DEFAULT_STYLES["styles"]))
        self.assertEqual(again.get(u"三线表").borders["top"]["sz"], 12)

    def test_a_missing_style_says_what_exists(self):
        with self.assertRaises(TableStyleError) as caught:
            StyleSet(DEFAULT_STYLES).get(u"不存在的款式")
        self.assertIn(u"三线表", u"%s" % caught.exception)


class TestCapture(StyleCase):
    """**采集**：把一张"手调好的表"读成款式 —— 用户说调一张满意的时间成本高，所以要能存下来。"""

    def build_tuned(self):
        """手调过的表：三线表样子 + 表头下细线 + 垂直居中 + 等宽列 + 表头重复。"""
        body = (u'<w:tbl><w:tblPr><w:jc w:val="center"/>'
                u'<w:tblBorders>'
                u'<w:top w:val="single" w:sz="12" w:color="000000"/>'
                u'<w:bottom w:val="single" w:sz="12" w:color="000000"/>'
                u'<w:left w:val="none" w:sz="0" w:color="auto"/>'
                u'<w:right w:val="none" w:sz="0" w:color="auto"/>'
                u'<w:insideH w:val="none" w:sz="0" w:color="auto"/>'
                u'<w:insideV w:val="none" w:sz="0" w:color="auto"/>'
                u'</w:tblBorders>'
                u'<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/>'
                u'<w:bottom w:w="40" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
                u'<w:tblLayout w:type="fixed"/></w:tblPr>'
                u'<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>'
                u'<w:tr><w:trPr><w:tblHeader/></w:trPr>'
                u'<w:tc><w:tcPr><w:vAlign w:val="center"/>'
                u'<w:tcBorders><w:bottom w:val="single" w:sz="6" w:color="000000"/></w:tcBorders>'
                u'</w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                u'<w:r><w:rPr><w:b/></w:rPr><w:t>表头</w:t></w:r></w:p></w:tc>'
                u'<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
                u'<w:p><w:r><w:t>表头二</w:t></w:r></w:p></w:tc></w:tr>'
                u'<w:tr><w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
                u'<w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc>'
                u'<w:tc><w:tcPr><w:vAlign w:val="center"/></w:tcPr>'
                u'<w:p><w:r><w:t>值二</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')
        return self.build(body)

    def test_capture_reads_back_what_was_tuned(self):
        self.build_tuned()
        with Document(self.path) as doc:
            data = capture(doc, 1)
        self.assertEqual(data["table_align"], "center")
        self.assertEqual(data["table_layout"], "fixed")
        self.assertEqual(data["borders"]["top"]["sz"], 12)
        self.assertEqual(data["borders"]["header_bottom"]["sz"], 6)
        self.assertEqual(data["cell_margins"], {"top": 40, "left": 80, "bottom": 40, "right": 80})
        self.assertEqual(data["v_align"], "center")
        self.assertEqual(data["header"], {"repeat": True, "align": "center", "bold": True})
        self.assertEqual(data["column_widths"], "equal")

    def test_a_captured_style_reproduces_the_table_on_another_document(self):
        """**核心诉求**：调一次 → 存成款式 → 套到别的文档上，得到同一副样子。"""
        self.build_tuned()
        with Document(self.path) as doc:
            data = capture(doc, 1)
        style = TableStyle(u"我的表", data)

        other = os.path.join(self.dir, u"另一个.docx")
        fixtures.write_fixture(other, body=table([row(cell(u"甲"), cell(u"乙")),
                                                 row(cell(u"1"), cell(u"2"))]))
        out = os.path.join(self.dir, u"套用后.docx")
        with Document(other) as doc:
            report = apply(doc, style)
            doc.save(out)
        self.assertGreater(report["total"], 0)
        with Document(out) as doc:
            pr = list(tables_of(doc))[0].find(qn("w:tblPr"))
        self.assertEqual(pr.find(qn("w:jc")).get(qn("w:val")), "center")
        self.assertEqual(pr.find(qn("w:tblBorders")).find(qn("w:top")).get(qn("w:sz")), "12")
        self.assertEqual(pr.find(qn("w:tblCellMar")).find(qn("w:left")).get(qn("w:w")), "80")

    def test_capture_through_the_cli_saves_into_the_style_file(self):
        self.build_tuned()
        styles_path = os.path.join(self.dir, u"tablestyle.json")
        from wordfactory.cli import main
        code = main(["tablestyle", "capture", self.path, "--out", styles_path,
                     "--name", u"我调的", "--table", "1"])
        self.assertEqual(code, 0)
        saved = StyleSet.load(styles_path)
        self.assertIn(u"我调的", saved.styles)
        self.assertEqual(saved.get(u"我调的").borders["header_bottom"]["sz"], 6)


class TestPreview(StyleCase):
    def test_the_preview_is_a_new_document_with_every_style(self):
        names = [u"三线表", u"全框线"]
        styles = [StyleSet(DEFAULT_STYLES).get(name) for name in names]
        path = build_preview(os.path.join(self.dir, u"preview.docx"), styles)
        with zipfile.ZipFile(path) as archive:
            names_in_zip = set(archive.namelist())
            doc = archive.read("word/document.xml").decode("utf-8")
        self.assertEqual(names_in_zip, {"[Content_Types].xml", "_rels/.rels",
                                        "word/document.xml"})
        self.assertIn(u"【三线表】", doc)
        self.assertIn(u"【全框线】", doc)
        self.assertEqual(doc.count("<w:tbl>"), len(names) * 3, u"每种款式 3 张代表性表格")

    def test_the_preview_goes_through_the_same_apply_path(self):
        """预览里的表必须**真的被套过款式**（不是另写一套渲染代码）。"""
        styles = [StyleSet(DEFAULT_STYLES).get(u"三线表")]
        path = build_preview(os.path.join(self.dir, u"p.docx"), styles)
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("word/document.xml").decode("utf-8"))
        table = list(root.iter(qn("w:tbl")))[0]
        pr = table.find(qn("w:tblPr"))
        self.assertIsNotNone(pr.find(qn("w:tblBorders")), u"预览里的表要带框线")
        self.assertEqual(pr.find(qn("w:jc")).get(qn("w:val")), "center")
        self.assertIsNotNone(table.find(qn("w:tr")).find(qn("w:trPr")).find(qn("w:tblHeader")))

    def test_the_preview_is_reproducible(self):
        import time
        styles = [StyleSet(DEFAULT_STYLES).get(u"三线表")]
        first = build_preview(os.path.join(self.dir, u"a.docx"), styles)
        time.sleep(1.1)
        second = build_preview(os.path.join(self.dir, u"b.docx"), styles)
        with open(first, "rb") as h1, open(second, "rb") as h2:
            self.assertEqual(h1.read(), h2.read(), u"预览文档也要逐字节可复现")

    def test_the_preview_command_writes_a_file(self):
        from wordfactory.cli import main
        out = os.path.join(self.dir, u"从命令来.docx")
        code = main(["tablestyle", "preview", "--out", out, "--style", u"三线表,全框线"])
        self.assertEqual(code, 0)
        self.assertTrue(os.path.exists(out))


if __name__ == "__main__":
    unittest.main()
