# -*- coding: utf-8 -*-
"""2026-09-27 那轮引擎件的测试：表格模板进管线 / 替换规则 / 就地更新 / 前置区手动页数。"""

import os
import shutil
import tempfile
import unittest
import zipfile

from wordfactory.document import Document
from wordfactory import pipeline as pipeline_mod
from wordfactory import replace_rules as replace_rules_mod
from wordfactory import frontmatter
from wordfactory.ooxml import qn
from wordfactory.ops import recipe as recipe_op

from . import fixtures


class PlanCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_plan_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _doc(self, name=u"报告.docx", body=None):
        path = os.path.join(self.dir, name)
        body = body if body is not None else (
            fixtures.paragraph(fixtures.run(u"表2.3-1     库容特性表"))
            + fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>", u"<w:r><w:t>库容</w:t></w:r>"],
                              [u"<w:r><w:t>135.50</w:t></w:r>", u"<w:r><w:t>1286</w:t></w:r>"]])
            + fixtures.paragraph(fixtures.run(u"正文一段。")))
        fixtures.write_fixture(path, body=body)
        return path

    # ---------------------------------------------------------- 表格模板进管线
    def test_tablestyle_step_applies_the_fallback_style(self):
        path = self._doc()
        out = os.path.join(self.dir, u"出.docx")
        report = pipeline_mod.run_pipeline(
            path, [{"op": "tablestyle", "params": {"style": u"三线表·学术款"}}],
            mode="formal", out_path=out)
        step = report["steps"][0]["report"]
        self.assertEqual(step["op"], "tablestyle")
        self.assertEqual(step["style"], u"三线表·学术款")
        self.assertTrue(os.path.exists(out))

    def test_tablestyle_mapping_picks_per_table(self):
        # 两张表：一张"水位"表、一张"流量"表；映射把流量表换成三线表
        body = (fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>"],
                                [u"<w:r><w:t>135.50</w:t></w:r>"]])
                + fixtures.table([[u"<w:r><w:t>流量</w:t></w:r>"],
                                  [u"<w:r><w:t>12.8</w:t></w:r>"]]))
        path = self._doc(body=body)
        with Document(path) as doc:
            report = pipeline_mod._run_step(
                doc, "tablestyle",
                {"style": u"通用款·外粗内细", "uniform": False, "tables": "all",
                 "mapping": [{"match": u"流量", "style": u"三线表·学术款"}]},
                dry_run=True)
        self.assertEqual(report["style_per_table"], {1: u"通用款·外粗内细",
                                                     2: u"三线表·学术款"})

    def test_tablestyle_uniform_ignores_mapping(self):
        body = (fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>"]])
                + fixtures.table([[u"<w:r><w:t>流量</w:t></w:r>"]]))
        path = self._doc(body=body)
        with Document(path) as doc:
            report = pipeline_mod._run_step(
                doc, "tablestyle",
                {"style": u"通用款·外粗内细", "uniform": True,
                 "mapping": [{"match": u"流量", "style": u"三线表·学术款"}]},
                dry_run=True)
        self.assertEqual(set(report["style_per_table"].values()), {u"通用款·外粗内细"})

    def test_default_templates_fill_the_body_width(self):
        from wordfactory.tablestyle import DEFAULT_STYLES, StyleSet
        styles = StyleSet(DEFAULT_STYLES)
        for name, style in styles.styles.items():
            if style.allow_overflow:
                continue
            self.assertEqual(style.width_percent, 100,
                             u"%s 要铺满版心（用户：最少与页边距一样宽）" % name)

    # ---------------------------------------------------------- 替换规则
    def test_replace_rules_text_font_para(self):
        body = (fixtures.paragraph(fixtures.run(u"用其它的方法。"))
                + fixtures.paragraph(fixtures.run(u"仿宋字。", rfonts={"eastAsia": u"仿宋_GB2312"}))
                + fixtures.paragraph(fixtures.run(u"两端对齐的段。"),
                                     style=None))
        path = self._doc(body=body)
        rules = {"name": u"测试规则",
                 "text": [{"find": u"其它", "replace": u"其他"}],
                 "font": [{"from": {"eastAsia": u"仿宋_GB2312"},
                           "to": {"eastAsia": u"宋体"}}],
                 "para": [{"from": {"align": "both"}, "to": {"align": "left"}}]}
        with Document(path) as doc:
            report = replace_rules_mod.apply(doc, rules)
            out = os.path.join(self.dir, u"替换后.docx")
            doc.save(out)
        self.assertGreaterEqual(report["text"], 1)
        self.assertGreaterEqual(report["font"], 1)
        with zipfile.ZipFile(out) as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertIn(u"其他", text)
        self.assertIn(u'w:eastAsia="宋体"', text)
        self.assertNotIn(u"仿宋_GB2312", text)

    def test_replace_rules_are_idempotent_and_never_touch_symbol_fonts(self):
        body = (fixtures.paragraph(fixtures.run(u"√", rfonts={"ascii": u"Symbol"}))
                + fixtures.paragraph(fixtures.run(u"仿宋字。", rfonts={"eastAsia": u"仿宋"})))
        path = self._doc(body=body)
        rules = {"name": u"r", "text": [], "para": [],
                 "font": [{"from": {"eastAsia": u"仿宋"}, "to": {"eastAsia": u"宋体"}}]}
        with Document(path) as doc:
            first = replace_rules_mod.apply(doc, rules)
            second = replace_rules_mod.apply(doc, rules)
        self.assertGreaterEqual(first["font"], 1)
        self.assertEqual(second["font"], 0, u"第二遍必须 0（幂等）")
        with Document(path) as doc:
            text = doc.part().find(qn("w:body")).find(qn("w:p")).find(qn("w:r")) \
                .find(qn("w:rPr")).find(qn("w:rFonts")).get(qn("w:ascii"))
        self.assertEqual(text, u"Symbol", u"符号字体永不碰")

    def test_replace_step_in_pipeline_by_name(self):
        path = self._doc(body=fixtures.paragraph(fixtures.run(u"用其它的方法。")))
        saved = replace_rules_mod.save(
            os.path.join(os.path.dirname(self.dir.rstrip(os.sep)), u"无用"),  # 不落用户目录
            u"测试规则", {"text": [{"find": u"其它", "replace": u"其他"}]}) \
            if False else None
        # 直接用路径参数（GUI 会传名字；名字解析单独测）
        out = os.path.join(self.dir, u"出.docx")
        rules_path = os.path.join(self.dir, u"规则.json")
        import json as json_mod
        import io as io_mod
        io_mod.open(rules_path, "w", encoding="utf-8").write(
            json_mod.dumps({"name": u"规则", "text": [{"find": u"其它", "replace": u"其他"}]},
                           ensure_ascii=False))
        report = pipeline_mod.run_pipeline(
            path, [{"op": "replace", "params": {"rules": rules_path}}],
            mode="formal", out_path=out)
        self.assertGreaterEqual(report["steps"][0]["report"]["text"], 1)

    # ---------------------------------------------------------- 就地更新
    def test_update_values_body_and_tables(self):
        body = (fixtures.paragraph(fixtures.run(u"库容 "), fixtures.run(u"1286", highlight="yellow"),
                                   fixtures.run(u" 万m³。"))
                + fixtures.table([[u"<w:r><w:t>135.50</w:t></w:r>"],
                                  [u"<w:r><w:t>1286</w:t></w:r>"]]))
        # 表格里的数值也要高亮才能被认出来
        body = body.replace(u"<w:r><w:t>1286</w:t></w:r>",
                            u'<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>'
                            u'<w:t>1286</w:t></w:r>')
        path = self._doc(body=body)
        with Document(path) as doc:
            body_report = recipe_op.update_values(doc, [u"2048"], scope="body", dry_run=True)
            self.assertEqual(body_report["replaced"], 1)
            self.assertEqual(body_report["missing"], 0)
            recipe_op.update_values(doc, [u"2048"], scope="body")
            out1 = os.path.join(self.dir, u"正文更新.docx")
            doc.save(out1)
        with Document(out1) as doc:
            texts = [p.text for p in doc.paragraphs()]
        self.assertIn(u"库容 2048 万m³。", texts)

        # 表格单元格里的高亮值走 scope="tables"
        with Document(path) as doc:
            table_report = recipe_op.update_values(doc, [u"2048"], scope="tables")
            out2 = os.path.join(self.dir, u"表格更新.docx")
            doc.save(out2)
        self.assertEqual(table_report["replaced"], 1)
        with zipfile.ZipFile(out2) as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertEqual(text.count(u">2048<"), 1, u"表格里的那个 1286 换成了 2048")
        self.assertIn(u"库容 ", text, u"正文那个不在本次 scope，不该被动")

    def test_update_values_reports_missing_when_values_run_out(self):
        body = (fixtures.paragraph(fixtures.run(u"A"), fixtures.run(u"1", highlight="yellow"))
                + fixtures.paragraph(fixtures.run(u"B"), fixtures.run(u"2", highlight="yellow")))
        path = self._doc(body=body)
        with Document(path) as doc:
            report = recipe_op.update_values(doc, [u"9"], scope="body", dry_run=True)
        self.assertEqual(report["replaced"], 1)
        self.assertEqual(report["missing"], 1)

    # ---------------------------------------------------------- 前置区手动页数
    def test_frontmatter_pages_override(self):
        body = (fixtures.paragraph(fixtures.run(u"封面", sz="72"))
                + u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                + fixtures.paragraph(fixtures.run(u"签字：张三"))
                + u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                + fixtures.paragraph(fixtures.run(u"前 言"))
                + fixtures.paragraph(fixtures.run(u""))
                + fixtures.paragraph(fixtures.run(u"正文。")))
        path = self._doc(body=body)
        with Document(path) as doc:
            auto = frontmatter.protected_elements(doc)
            manual = frontmatter.protected_elements(doc, pages=1)
        self.assertTrue(auto, u"自动识别：封面+签字页都在前言之前")
        self.assertTrue(manual, u"手动第 1 页也要有保护")
        self.assertLess(len(manual), len(auto), u"只保护第 1 页时比自动（前 2 页）少")

    def test_tidy_honors_frontmatter_pages(self):
        from wordfactory.ops import tidy as tidy_op
        body = (fixtures.paragraph(fixtures.run(u"封面"))
                + u'<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                + fixtures.paragraph(fixtures.run(u"前 言"))
                + fixtures.paragraph(fixtures.run(u""))
                + fixtures.paragraph(fixtures.run(u"")))
        path = self._doc(body=body)
        with Document(path) as doc:
            report = tidy_op.tidy(doc, {"frontmatter_pages": 1})
        self.assertGreater(report["changes"].get(u"前置区保护（跳过）", 0), 0)


if __name__ == "__main__":
    unittest.main()


class TestPreviewDistinctness(PlanCase):
    """六款预览必须**画得出来区别**（用户 2026-09-27：根本看不出区别）。

    判据：把样图里点名款式的标题文字去掉后，六张 SVG 两两不同 ——
    列宽分布（keep/equal/content）、长文本列对齐（居中/左）、三线表（无竖线）都得体现在图上。
    """

    def test_six_previews_are_mutually_distinct(self):
        from wordfactory.tablestyle import DEFAULT_STYLES, StyleSet, preview_svg
        styles = StyleSet(DEFAULT_STYLES)
        bodies = []
        for name, style in styles.styles.items():
            svg = preview_svg(style)
            self.assertTrue(svg.startswith(u"<svg"))
            # 去掉标题（含款式名）再比——剩下的是"画了什么"
            body = svg.split(u"</text>", 1)[1]
            bodies.append((name, body))
        for i, (name_a, body_a) in enumerate(bodies):
            for name_b, body_b in bodies[i + 1:]:
                self.assertNotEqual(body_a, body_b,
                                    u"预览画不出区别：%s 和 %s" % (name_a, name_b))

    def test_column_width_strategies_draw_different_columns(self):
        from wordfactory.tablestyle import TableStyle, preview_svg
        keep = TableStyle(u"keep", {"column_widths": "keep",
                                    "borders": {"top": {"val": "single", "sz": 12}}})
        equal = TableStyle(u"equal", {"column_widths": "equal",
                                      "borders": {"top": {"val": "single", "sz": 12}}})
        def xs(svg):
            import re
            return re.findall(r'<line x1="([\d.]+)"', svg)
        self.assertNotEqual(xs(preview_svg(keep)), xs(preview_svg(equal)))


class TestFontColorBoldRule(PlanCase):
    """文字格式替换：颜色与字形（用户 2026-09-27 点名要颜色栏目）。"""

    def test_color_replacement(self):
        body = fixtures.paragraph(fixtures.run(u"红字。", color="FF0000"))
        path = self._doc(body=body)
        rules = {"name": u"颜色", "text": [], "para": [],
                 "font": [{"from": {"color": "FF0000"}, "to": {"color": "000000"}}]}
        with Document(path) as doc:
            report = replace_rules_mod.apply(doc, rules)
            out = os.path.join(self.dir, u"颜色.docx")
            doc.save(out)
        self.assertGreaterEqual(report["font"], 1)
        with zipfile.ZipFile(out) as archive:
            text = archive.read("word/document.xml").decode("utf-8")
        self.assertNotIn(u"FF0000", text)
        self.assertIn(u'w:val="000000"', text)

    def test_second_pass_is_zero(self):
        body = fixtures.paragraph(fixtures.run(u"红字。", color="FF0000"))
        path = self._doc(body=body)
        rules = {"name": u"颜色", "text": [], "para": [],
                 "font": [{"from": {"color": "FF0000"}, "to": {"color": "000000"}}]}
        with Document(path) as doc:
            replace_rules_mod.apply(doc, rules)
            second = replace_rules_mod.apply(doc, rules)
        self.assertEqual(second["font"], 0)
