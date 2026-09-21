# -*- coding: utf-8 -*-
"""文本层的单测：跨 run 替换是本项目的命门，这里把它的行为钉死。

夹具刻意做成 Word 真实的样子：一句话被切成 3 个 run（格式变化处就会切），
中间那个 run 还带颜色——替换后**必须**保留第一个 run 的格式，并且不能吃掉别的 run 的属性。
"""

import unittest
from xml.etree import ElementTree as ET

from wordfactory.ooxml import DocxPackage, local_name, qn
from wordfactory.text import Paragraph, paragraphs, story_text

from . import fixtures


def parse(body_xml):
    return ET.fromstring(fixtures.build_document(body_xml))


class TextLayerCase(unittest.TestCase):
    def paragraph(self, *runs):
        root = parse(fixtures.paragraph(*runs))
        return Paragraph(root.find(qn("w:body")).find(qn("w:p")))


class TestLogicalText(TextLayerCase):
    def test_joins_runs_into_one_logical_sentence(self):
        para = self.paragraph(fixtures.run(u"山"), fixtures.run(u"亭区"),
                              fixtures.run(u"的报告"))
        self.assertEqual(para.text, u"山亭区的报告")
        self.assertEqual(para.run_count, 3)

    def test_run_map_shows_how_the_sentence_was_split(self):
        para = self.paragraph(fixtures.run(u"山"), fixtures.run(u"亭区"),
                              fixtures.run(u"的报告"))
        self.assertEqual([text for _run, text in para.run_map()],
                         [u"山", u"亭区", u"的报告"])

    def test_tab_and_break_count_as_characters(self):
        root = parse(u"<w:p><w:r><w:t>前</w:t><w:tab/><w:t>后</w:t></w:r></w:p>")
        para = Paragraph(root.find(qn("w:body")).find(qn("w:p")))
        self.assertEqual(para.text, u"前\t后")

    def test_runs_inside_a_hyperlink_are_part_of_the_paragraph(self):
        body = (u"<w:p><w:r><w:t>见</w:t></w:r>"
                u'<w:hyperlink r:id="rId1"><w:r><w:t>链接</w:t></w:r></w:hyperlink>'
                u"<w:r><w:t>处</w:t></w:r></w:p>")
        root = parse(body)
        self.assertEqual(Paragraph(root.find(qn("w:body")).find(qn("w:p"))).text,
                         u"见链接处")

    def test_a_textbox_paragraph_is_not_swallowed(self):
        body = (u"<w:p><w:r><w:t>正文</w:t></w:r>"
                u"<w:r><w:pict><w:txbxContent><w:p><w:r><w:t>文本框里的话</w:t></w:r>"
                u"</w:p></w:txbxContent></w:pict></w:r></w:p>")
        root = parse(body)
        outer = Paragraph(root.find(qn("w:body")).find(qn("w:p")))
        self.assertEqual(outer.text, u"正文", "外层段落不该把文本框里的字算进来")


class TestReplacement(TextLayerCase):
    def test_replaces_across_run_boundaries(self):
        para = self.paragraph(fixtures.run(u"山"), fixtures.run(u"亭区"),
                              fixtures.run(u"的报告"))
        self.assertEqual(para.replace(u"山亭区", u"寒亭区"), 1)
        self.assertEqual(para.text, u"寒亭区的报告")

    def test_the_first_run_keeps_its_formatting(self):
        """Word 自己的"全部替换"就是这个行为：格式跟命中的第一个 run 走。"""
        para = self.paragraph(fixtures.run(u"山"),
                              fixtures.run(u"亭区", color="FF0000"),
                              fixtures.run(u"的报告"))
        para.replace(u"山亭区", u"寒亭区")
        first_run = para.runs[0]
        self.assertEqual(first_run.find(qn("w:t")).text, u"寒亭区",
                         "替换文本应该落在第一个命中的 run 上")
        color = para.runs[1].find(qn("w:rPr") + "/" + qn("w:color"))
        self.assertIsNotNone(color, "被清空的 run 不该丢掉它自己的 rPr")

    def test_replace_within_a_single_run(self):
        para = self.paragraph(fixtures.run(u"甲乙丙丁"))
        self.assertEqual(para.replace(u"丙丁", u"丙"), 1)
        self.assertEqual(para.text, u"甲乙丙")

    def test_replace_all_occurrences_by_default(self):
        para = self.paragraph(fixtures.run(u"旧名称 旧名称 保持不动"))
        self.assertEqual(para.replace(u"旧名称", u"新名称"), 2)
        self.assertEqual(para.text, u"新名称 新名称 保持不动")

    def test_count_limits_the_number_of_replacements(self):
        para = self.paragraph(fixtures.run(u"a a a"))
        self.assertEqual(para.replace(u"a", u"b", count=2), 2)
        self.assertEqual(para.text, u"b b a")

    def test_delete_across_runs(self):
        para = self.paragraph(fixtures.run(u"前"), fixtures.run(u"要删的"),
                              fixtures.run(u"后"))
        para.delete(u"要删的")
        self.assertEqual(para.text, u"前后")

    def test_replacing_twice_is_idempotent_when_the_needle_is_gone(self):
        para = self.paragraph(fixtures.run(u"山亭区"))
        self.assertEqual(para.replace(u"山亭区", u"寒亭区"), 1)
        self.assertEqual(para.replace(u"山亭区", u"寒亭区"), 0)
        self.assertEqual(para.text, u"寒亭区")

    def test_regex_replacement_with_backreferences(self):
        para = self.paragraph(fixtures.run(u"面积 12m2 与 3m3"))
        self.assertEqual(para.replace_regex(r"(\d+)m(\d)", r"\1 m\2"), 2)
        self.assertEqual(para.text, u"面积 12 m2 与 3 m3")

    def test_replacement_does_not_touch_other_parts(self):
        para = self.paragraph(fixtures.run(u"山"), fixtures.run(u"亭区"))
        para.replace(u"山亭区", u"寒亭区")
        self.assertEqual(para.text, u"寒亭区")


class TestWhitespaceCare(TextLayerCase):
    def test_leading_or_trailing_space_gets_the_preserve_attribute(self):
        para = self.paragraph(fixtures.run(u"表5-1 名称"))
        para.replace(u"名称", u"  名称  ")          # 两侧带空格
        node = para.runs[0].find(qn("w:t"))
        self.assertEqual(node.text, u"表5-1   名称  ")
        self.assertEqual(node.get("{http://www.w3.org/XML/1998/namespace}space"),
                         "preserve", "首尾有空格必须声明 xml:space=preserve")

    def test_the_attribute_is_removed_again_when_spaces_go_away(self):
        para = self.paragraph(fixtures.run(u"  去空格  "))
        node = para.runs[0].find(qn("w:t"))
        node.text = u"  去空格  "
        para.replace(u"  去空格  ", u"去空格")
        self.assertIsNone(
            para.runs[0].find(qn("w:t")).get(
                "{http://www.w3.org/XML/1998/namespace}space"),
            "不再有首尾空格时不该留着 preserve（留着无害但会越积越多）")

    def test_deleting_spaces_around_chinese_works_across_runs(self):
        para = self.paragraph(fixtures.run(u"防 洪 "), fixtures.run(u"安 全"))
        self.assertEqual(para.replace_regex(u"[ \u3000]+", u""), 3)
        self.assertEqual(para.text, u"防洪安全")


class TestThroughARealPackage(unittest.TestCase):
    """在真正的 .docx 上跑一遍（打开 → 改 → 存 → 重开验证）。"""

    def setUp(self):
        import os
        import tempfile
        import shutil
        self.work = tempfile.mkdtemp(prefix="wf_text_")
        self.src = fixtures.write_fixture(os.path.join(self.work, "fixture.docx"))
        self.shutil = shutil

    def tearDown(self):
        self.shutil.rmtree(self.work, ignore_errors=True)

    def test_edit_a_real_document_and_read_it_back(self):
        import os
        out = os.path.join(self.work, "out.docx")
        with DocxPackage(self.src) as pkg:
            root = pkg.xml(DocxPackage.MAIN)
            body = root.find(qn("w:body"))
            changed = 0
            for paragraph in paragraphs(root):
                changed += paragraph.replace(u"山亭区", u"寒亭区")
            self.assertEqual(changed, 1, "夹具里应该正好有一处")
            self.assertEqual(changed, 1)
            pkg.mark_dirty(DocxPackage.MAIN)
            pkg.save(out)
        with DocxPackage(out) as pkg:
            text = story_text(pkg.xml(DocxPackage.MAIN))
        self.assertIn(u"寒亭区", text)
        self.assertNotIn(u"山亭区", text)
        self.assertIn(u"这是一个普通段落。", text, "别的段落不该被动过")


if __name__ == "__main__":
    unittest.main()
