# -*- coding: utf-8 -*-
"""路径与"首次运行写出默认规则"（打包成 exe 后规则从哪来）。

实测过的坑：``json.dumps`` 不收 ``object_pairs_hook``（那是 ``load`` 的参数），
生成默认规则一失败就留下 **0 字节文件**，下次启动当它存在 —— 打包版首次运行
因此写出三个空规则，功能一进去就报 500。
"""

import json
import os
import shutil
import tempfile
import unittest

from wordfactory import paths


class EnsureDefaultsCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_paths_")
        self.original = paths.data_dir
        paths.data_dir = lambda: self.dir

    def tearDown(self):
        paths.data_dir = self.original
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_all_default_rules_are_written_and_non_empty(self):
        written = paths.ensure_defaults(verbose=False)
        self.assertEqual(len(written), len(paths.DEFAULTS))
        for name, _maker in paths.DEFAULTS:
            path = paths.rules_path(name)
            self.assertTrue(os.path.exists(path), name)
            self.assertGreater(os.path.getsize(path), 0,
                               u"%s 是空文件 —— 打包版首次运行踩过这个" % name)
            with open(path, encoding="utf-8") as handle:
                self.assertIsInstance(json.load(handle), (dict, list))

    def test_user_edits_are_never_overwritten(self):
        paths.ensure_defaults(verbose=False)
        target = paths.rules_path(u"fonts.json")
        with open(target, "w", encoding="utf-8") as handle:
            handle.write(u'{"note": "我自己改的"}')
        self.assertEqual(paths.ensure_defaults(verbose=False), [],
                         u"已有内容不该被重写")

    def test_an_empty_file_is_treated_as_missing(self):
        """0 字节文件等于没有 —— 要重新写（这正是打包版踩的坑）。"""
        os.makedirs(paths.rules_dir(), exist_ok=True)
        with open(paths.rules_path(u"fonts.json"), "w", encoding="utf-8") as handle:
            handle.write(u"")
        written = paths.ensure_defaults(verbose=False)
        self.assertIn(u"fonts.json", written)
        self.assertGreater(os.path.getsize(paths.rules_path(u"fonts.json")), 0)

    def test_a_failing_maker_does_not_leave_a_broken_file(self):
        os.makedirs(paths.rules_dir(), exist_ok=True)
        def boom():
            raise RuntimeError(u"故意失败")

        broken = (u"坏规则.json", boom)
        original = paths.DEFAULTS
        paths.DEFAULTS = list(original) + [broken]
        try:
            paths.ensure_defaults(verbose=False)
        finally:
            paths.DEFAULTS = original
        self.assertFalse(os.path.exists(paths.rules_path(u"坏规则.json")),
                         u"生成失败就不该留下文件")


class FrozenPathsCase(unittest.TestCase):
    def test_frozen_app_uses_the_folder_next_to_the_exe(self):
        original_frozen = paths.is_frozen
        original_exec = None
        import sys
        original_exec = sys.executable
        try:
            paths.is_frozen = lambda: True
            sys.executable = os.path.join(u"C:\Tools", u"word工厂.exe")
            self.assertEqual(paths.app_dir(), u"C:\Tools", )
            self.assertEqual(paths.data_dir(), os.path.join(u"C:\Tools", u"word工厂数据"))
            self.assertEqual(paths.temp_dir(), os.path.join(u"C:\Tools", u"word工厂数据", u"临时文件"))
        finally:
            paths.is_frozen = original_frozen
            sys.executable = original_exec


if __name__ == "__main__":
    unittest.main()
