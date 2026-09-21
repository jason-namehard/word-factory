# -*- coding: utf-8 -*-
"""上下标规则引擎的单测：规则文件怎么写、怎么校验、怎么应用、能不能重跑。

用户要的是"外置接口 + 正则查找并应用"，所以这里把四件事钉死：
字面量字典规则、正则规则、先匹配先应用、以及**幂等**（重跑报 0 处）。
"""

import os
import shutil
import tempfile
import unittest
from xml.etree import ElementTree as ET

from wordfactory.ooxml import DocxPackage, qn
from wordfactory.rules import (BOUNDARY_CHARS, RuleSet, RuleError, apply_to_part,
                               default_ruleset, write_default)
from wordfactory.text import (Paragraph, paragraphs, set_vertical_align,
                              split_run_at, story_text, vertical_align_of)

from . import fixtures


def parse_paragraph(*runs):
    root = ET.fromstring(fixtures.build_document(fixtures.paragraph(*runs)))
    return Paragraph(root.find(qn("w:body")).find(qn("w:p")))


def marks(paragraph):
    """``[(文本, 上标/下标)]`` —— 用来断言"哪几个字被设了什么"。"""
    out = []
    for run in paragraph.runs:
        text = "".join(node.text or "" for node in run.findall(qn("w:t")))
        if text:
            out.append((text, vertical_align_of(run)))
    return out


class TestRuleValidation(unittest.TestCase):
    def test_a_good_default_ruleset_has_no_problems(self):
        self.assertEqual(default_ruleset().validate(), [])

    def test_bad_regex_is_reported(self):
        rule_set = RuleSet({"rules": [{"id": "bad", "pattern": "[unclosed"}]})
        problems = rule_set.validate()
        self.assertTrue(problems)
        self.assertIn("bad", u" ".join(problems))

    def test_kinds_must_match_the_literal_length(self):
        rule_set = RuleSet({"rules": [{"id": "x", "match": "m2", "kinds": "NSB"}]})
        self.assertIn(u"长度不等", u" ".join(rule_set.validate()))

    def test_kinds_letters_are_limited(self):
        rule_set = RuleSet({"rules": [{"id": "x", "match": "m2", "kinds": "NX"}]})
        self.assertIn(u"只能出现 N/S/B", u" ".join(rule_set.validate()))

    def test_group_index_must_exist(self):
        rule_set = RuleSet({"rules": [{"id": "x", "pattern": "m2", "target": "group:1",
                                       "kind": "superscript"}]})
        self.assertIn(u"捕获组", u" ".join(rule_set.validate()))

    def test_duplicate_ids_are_reported(self):
        rule_set = RuleSet({"rules": [{"id": "dup", "match": "a", "kinds": "N"},
                                      {"id": "dup", "match": "b", "kinds": "N"}]})
        self.assertIn(u"id 重复", u" ".join(rule_set.validate()))

    def test_a_rule_needs_match_or_pattern(self):
        rule_set = RuleSet({"rules": [{"id": "empty"}]})
        self.assertIn(u"既没有", u" ".join(rule_set.validate()))


class TestRuleApplication(unittest.TestCase):
    def setUp(self):
        self.rules = default_ruleset()

    def test_literal_dictionary_rule_sets_only_the_letter_it_names(self):
        """m2 → m 正常、2 上标。"""
        para = parse_paragraph(fixtures.run(u"面积 12m2"))
        changes = self.rules.apply_paragraph(para)
        self.assertEqual([c["text"] for c in changes], [u"2"])
        self.assertEqual(marks(para), [(u"面积 12m", None), (u"2", u"superscript")])

    def test_celsius_rule_superscripts_only_the_degree_sign(self):
        para = parse_paragraph(fixtures.run(u"气温 25°C"))
        self.rules.apply_paragraph(para)
        self.assertEqual(marks(para), [(u"气温 25", None), (u"°", u"superscript"),
                                       (u"C", None)])

    def test_multi_character_rule_sets_several_letters(self):
        """Vmax → m/a/x 下标（V 正常）。"""
        para = parse_paragraph(fixtures.run(u"Vmax"))
        self.rules.apply_paragraph(para)
        self.assertEqual(marks(para), [(u"V", None), (u"max", u"subscript")])

    def test_regex_rule_targets_one_capture_group(self):
        rules = RuleSet({"rules": [
            {"id": "unit-m", "pattern": u"m([2-9]+)", "target": "group:1",
             "kind": "superscript"}]})
        para = parse_paragraph(fixtures.run(u"12m3 与 m9 与 m1"))
        changes = rules.apply_paragraph(para)
        self.assertEqual([c["text"] for c in changes], [u"3", u"9"])
        self.assertEqual(marks(para), [(u"12m", None), (u"3", u"superscript"),
                                       (u" 与 m", None), (u"9", u"superscript"),
                                       (u" 与 m1", None)])

    def test_regex_rule_can_target_the_whole_match(self):
        rules = RuleSet({"rules": [{"id": "deg", "pattern": u"°C", "target": "all",
                                    "kind": "superscript"}]})
        para = parse_paragraph(fixtures.run(u"25°C"))
        rules.apply_paragraph(para)
        self.assertEqual(marks(para), [(u"25", None), (u"°C", u"superscript")])

    def test_the_first_matching_rule_wins(self):
        """先匹配先应用：字典规则占了 m2，后面的通用正则就不再动它。"""
        rules = RuleSet({"rules": [
            {"id": "dict-m2", "match": u"m2", "kinds": "NS"},
            {"id": "generic", "pattern": u"m([2-9]+)", "target": "group:1",
             "kind": "superscript"}]})
        para = parse_paragraph(fixtures.run(u"m2 m3"))
        rules.apply_paragraph(para)
        self.assertEqual(marks(para), [(u"m", None), (u"2", u"superscript"),
                                       (u" m", None), (u"3", u"superscript")])

    def test_a_disabled_rule_does_nothing(self):
        rules = RuleSet({"rules": [{"id": "off", "pattern": u"m([0-9]+)",
                                    "target": "group:1", "kind": "superscript",
                                    "enabled": False}]})
        para = parse_paragraph(fixtures.run(u"m3"))
        self.assertEqual(rules.apply_paragraph(para), [])
        self.assertEqual(marks(para), [(u"m3", None)])

    def test_not_before_and_not_after_narrow_a_match(self):
        """只给"前面不是数字"的 2 设上标 —— 用来表达边界字符表那种约束。"""
        rules = RuleSet({"rules": [{"id": "caret", "pattern": u"2",
                                    "target": "all", "kind": "superscript",
                                    "not_before": u"[0-9]"}]})
        para = parse_paragraph(fixtures.run(u"12 与 2"))
        rules.apply_paragraph(para)
        self.assertEqual(marks(para), [(u"12 与 ", None), (u"2", u"superscript")])

    def test_boundary_characters_match_the_source_expansion(self):
        """规格书 §2.5 的展开表里的 22 个字符。

        （注：规格书正文写"20 个字符"，但它自己的逐字符展开列了 22 个；
        以展开表与源码字节为准 —— 这条记在 docs/REFERENCE-MACROS.md 的待核对项里。）
        """
        self.assertEqual(len(BOUNDARY_CHARS), 22)
        for ch in u"，。！？；：\"=-+*×/÷（）【】《》＝％":
            self.assertIn(ch, BOUNDARY_CHARS, ch)
        self.assertNotIn(u" ", BOUNDARY_CHARS, "空格不在边界表里")


class TestIdempotence(unittest.TestCase):
    def test_running_the_rules_twice_changes_nothing_the_second_time(self):
        para = parse_paragraph(fixtures.run(u"面积 12m2，气温 25°C，Vmax"))
        rules = default_ruleset()
        first = rules.apply_paragraph(para)
        once = ET.tostring(para.element, encoding="unicode")
        second = rules.apply_paragraph(para)
        twice = ET.tostring(para.element, encoding="unicode")
        self.assertTrue(first, "第一次就该有改动")
        self.assertEqual(second, [], "第二次不该再报改动")
        self.assertEqual(once, twice, "第二次不该再改任何字节")


class TestSplitRun(unittest.TestCase):
    """切 run 是所有"只改几个字"的功能的地基，单独钉一遍。"""

    def test_splitting_in_the_middle_keeps_both_halves_and_the_formatting(self):
        para = parse_paragraph(fixtures.run(u"甲乙丙丁", color="FF0000"))
        self.assertTrue(split_run_at(para, 2))
        self.assertEqual([text for _r, text in para.run_map()], [u"甲乙", u"丙丁"])
        for run in para.runs:
            self.assertIsNotNone(run.find(qn("w:rPr") + "/" + qn("w:color")),
                                 "切开后两半都要保留原来的 rPr")

    def test_splitting_at_an_existing_boundary_does_nothing(self):
        para = parse_paragraph(fixtures.run(u"甲乙"), fixtures.run(u"丙丁"))
        before = ET.tostring(para.element, encoding="unicode")
        split_run_at(para, 2)
        self.assertEqual(ET.tostring(para.element, encoding="unicode"), before)

    def test_setting_a_property_on_part_of_a_run_splits_it(self):
        para = parse_paragraph(fixtures.run(u"面积12m2"))     # 逻辑长度 6：面 积 1 2 m 2
        touched = set_vertical_align(para, 5, 6, "superscript")
        self.assertEqual(touched, 1)
        self.assertEqual(marks(para), [(u"面积12m", None), (u"2", u"superscript")])

    def test_clearing_the_property_removes_the_element(self):
        para = parse_paragraph(fixtures.run(u"面积12m2"))
        set_vertical_align(para, 5, 6, "superscript")
        set_vertical_align(para, 5, 6, None)
        for run in para.runs:
            self.assertIsNone(vertical_align_of(run))


class TestRulesThroughARealDocument(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp(prefix="wf_rules_")
        body = u"".join([
            fixtures.paragraph(fixtures.run(u"面积 12"), fixtures.run(u"m2")),
            fixtures.paragraph(fixtures.run(u"气温 25°C，Vmax")),
            fixtures.paragraph(fixtures.run(u"不应被改的一行")),
        ])
        self.src = fixtures.write_fixture(os.path.join(self.work, "fixture.docx"),
                                         body=body)
        self.rules_path = os.path.join(self.work, "rules.json")
        write_default(self.rules_path)

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)

    def test_rules_file_round_trips(self):
        loaded = RuleSet.load(self.rules_path)
        self.assertEqual(len(loaded.rules), len(default_ruleset().rules))

    def test_apply_save_and_read_back(self):
        out = os.path.join(self.work, "out.docx")
        rules = RuleSet.load(self.rules_path)
        with DocxPackage(self.src) as pkg:
            root = pkg.xml(DocxPackage.MAIN)
            report = apply_to_part(root, rules)
            self.assertGreater(report["total"], 0)
            pkg.mark_dirty(DocxPackage.MAIN)
            pkg.save(out)
        with DocxPackage(out) as pkg:
            root = pkg.xml(DocxPackage.MAIN)
            all_marks = []
            for para in paragraphs(root):
                all_marks.extend(marks(para))
        superscripted = [text for text, kind in all_marks if kind == "superscript"]
        subscripted = [text for text, kind in all_marks if kind == "subscript"]
        self.assertIn(u"2", superscripted)          # m2 的 2
        self.assertIn(u"°", superscripted)          # 25°C 的 °
        self.assertIn(u"max", subscripted)          # Vmax 的 max（一整段，不是三个 run）
        self.assertNotIn(u"4", superscripted, "未涉及的段落不该被改")

    def test_a_missing_rules_file_says_so(self):
        with self.assertRaises(RuleError):
            RuleSet.load(os.path.join(self.work, "nope.json"))


if __name__ == "__main__":
    unittest.main()
