# -*- coding: utf-8 -*-
"""Markdown 清理（`ops/mdclean.py`）的单测。

规则来自参考宏 `MarkDown语言清除.bas`（§2.1），逐条对照：
`行内代码 → 中文双引号`、`**粗体**`/`*斜体*` 去星号、段首 `#` 全剥、一个列表标记、
段内星号全删、连续空格折叠、每段 Trim。

另外要钉两条**实现口径**：

* 纯文本版 `clean_text` 与 XML 版 `apply` **共用同一张规则表** → 同一段文字两边结果必须一致；
* `apply` 是**局部替换**：没匹配到的文字，run 结构（加粗等）原样保留（宏是整段回写、会压平）。
"""

import os
import shutil
import tempfile
import unittest

from wordfactory.document import Document
from wordfactory.ooxml import qn
from wordfactory.ops import mdclean
from wordfactory.text import Paragraph

from . import fixtures


class CleanTextCase(unittest.TestCase):
    def test_the_rules_match_the_macro(self):
        cases = [
            (u"# 一级标题", u"一级标题"),
            (u"### 三级标题", u"三级标题"),
            (u"- 列表项", u"列表项"),
            (u"* 列表项", u"列表项"),
            (u"+ 列表项", u"列表项"),
            (u"这是**粗体**字", u"这是粗体字"),
            (u"这是*斜体*字", u"这是斜体字"),
            (u"行内代码 `var` 结束", u"行内代码 \u201cvar\u201d 结束"),
            (u"两个  空格", u"两个 空格"),
            (u"  前后有空白  ", u"前后有空白"),
            (u"**整段粗体**", u"整段粗体"),
        ]
        for source, expected in cases:
            self.assertEqual(mdclean.clean_text(source), expected, source)

    def test_the_dead_branch_is_not_implemented(self):
        """宏里 `#.` 那段（`:71-81`）是死代码 —— 我们不实现，这里把判断留痕。"""
        # `#. 标题` 先被"段首井号"剥掉 `#` → 剩 `. 标题`；宏的 `#.*` 分支永远进不去
        self.assertEqual(mdclean.clean_text(u"#. 标题"), u". 标题")

    def test_paragraphs_are_kept_apart(self):
        text = u"# 第一段\n".replace(u"\n", u"") and u"# 第一段\r**第二段**"
        self.assertEqual(mdclean.clean_text(text), u"第一段\r第二段")

    def test_plain_text_is_left_alone(self):
        plain = u"这是一段没有标记的普通文字，含 1.5 与 3.5 这样的数字。"
        self.assertEqual(mdclean.clean_text(plain), plain)
        self.assertFalse(mdclean.looks_like_markdown(plain))


class ApplyCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_mdclean_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def run_op(self, body, scope="body", dry_run=False):
        fixtures.write_fixture(self.path, body=body)
        with Document(self.path) as doc:
            report = mdclean.apply(doc, {"scope": scope}, dry_run=dry_run)
            out = None
            if report["touched"] and not dry_run:
                out = doc.save(os.path.join(self.dir, u"出.docx"))
        return report, out

    def texts(self, path):
        with Document(path) as doc:
            return [Paragraph(element).text for element in doc.part().iter(qn("w:p"))]

    def test_a_heading_paragraph_loses_its_hashes(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"## 二、设计暴雨")))
        self.assertEqual(report["touched"], 1)
        self.assertEqual(self.texts(out), [u"二、设计暴雨"])

    def test_bold_inside_a_paragraph_is_cleaned_but_other_text_keeps_its_runs(self):
        """**关键口径**：局部替换 —— 没匹配到的文字，run（格式）原样保留。"""
        body = fixtures.paragraph(
            fixtures.run(u"普通文字"),
            u'<w:r><w:rPr><w:b/></w:rPr><w:t>**粗体**</w:t></w:r>',
            fixtures.run(u"后面的字"))
        report, out = self.run_op(body)
        self.assertEqual(report["touched"], 1)
        with Document(out) as doc:
            runs = list(doc.part().iter(qn("w:r")))
        texts = [u"".join((node.text or "") for node in run.findall(qn("w:t"))) for run in runs]
        self.assertEqual(u"".join(texts), u"普通文字粗体后面的字")
        self.assertEqual(len(runs), 3, u"run 数量不变（宏的整段回写会压成一个）")
        self.assertIsNotNone(runs[1].find(qn("w:rPr")).find(qn("w:b")), u"加粗还在")

    def test_paragraphs_without_markers_are_untouched(self):
        report, out = self.run_op(fixtures.paragraph(fixtures.run(u"干干净净的一段")))
        self.assertEqual(report["touched"], 0)
        self.assertIsNone(out)

    def test_the_markdown_rules_survive_a_round_trip_through_xml(self):
        """XML 版与纯文本版必须给同一个结果（共用规则表，不能漂移）。"""
        samples = [u"# 标题", u"- 项目", u"**粗**与*斜*", u"代码 `x` 结束", u"两个  空格"]
        for sample in samples:
            body = fixtures.paragraph(fixtures.run(sample))
            report, out = self.run_op(body)
            if report["touched"]:
                self.assertEqual(self.texts(out), [mdclean.clean_text(sample)], sample)

    def test_dry_run_touches_nothing(self):
        fixtures.write_fixture(self.path, body=fixtures.paragraph(fixtures.run(u"# 标题")))
        with Document(self.path) as doc:
            report = mdclean.apply(doc, dry_run=True)
            self.assertEqual(doc.dirty_parts, [])
            self.assertEqual([Paragraph(e).text for e in doc.body() if e.tag == qn("w:p")],
                             [u"# 标题"])
        self.assertEqual(report["touched"], 1, u"dry-run 要报出「会改多少」")

    def test_running_twice_changes_nothing(self):
        fixtures.write_fixture(self.path, body=fixtures.paragraph(fixtures.run(u"## 标题")))
        with Document(self.path) as doc:
            first = mdclean.apply(doc)
            second = mdclean.apply(doc)
        self.assertEqual(first["touched"], 1)
        self.assertEqual(second["touched"], 0, u"清完第二遍没有标记了（幂等）")

    def test_scope_all_reaches_tables(self):
        body = fixtures.table([[u"<w:r><w:t># 表内标题</w:t></w:r>",
                                u"<w:r><w:t>值</w:t></w:r>"]])
        report, _ = self.run_op(body, scope="body")
        self.assertEqual(report["touched"], 0, u"默认不动表格")
        report, out = self.run_op(body, scope="all")
        self.assertEqual(report["touched"], 1)

    def test_the_cli_runs(self):
        fixtures.write_fixture(self.path, body=fixtures.paragraph(fixtures.run(u"# 标题")))
        from wordfactory.cli import main
        out = os.path.join(self.dir, u"从命令来.docx")
        self.assertEqual(main(["mdclean", self.path, "--out", out]), 0)
        self.assertEqual(self.texts(out), [u"标题"])


if __name__ == "__main__":
    unittest.main()
