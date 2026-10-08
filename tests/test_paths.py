# -*- coding: utf-8 -*-
"""路径与"首次运行写出默认规则"（打包成 exe 后规则从哪来）。

实测过的坑：``json.dumps`` 不收 ``object_pairs_hook``（那是 ``load`` 的参数），
生成默认规则一失败就留下 **0 字节文件**，下次启动当它存在 —— 打包版首次运行
因此写出三个空规则，功能一进去就报 500。
"""

import json
import os
import shutil
import sys
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
    """打包版（exe）路径口径。

    ⚠️ 2026-10-08 修：原来这个测试拿 ``C:\\Tools`` 当"exe 所在目录"，
    跑一次就在 **C 盘根上真建** ``C:\\Tools\\word工厂数据``（改工作区外留垃圾）。
    现在改成临时目录 —— 测试不许有这种副作用。
    """

    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_frozen_")
        self.exe_dir = os.path.join(self.dir, u"程序")
        os.makedirs(self.exe_dir)
        self.original_frozen = paths.is_frozen
        self.original_exec = sys.executable
        self.original_chosen = paths._chosen_data_dir
        paths._chosen_data_dir = None
        paths.is_frozen = lambda: True
        sys.executable = os.path.join(self.exe_dir, u"word工厂.exe")

    def tearDown(self):
        paths.is_frozen = self.original_frozen
        sys.executable = self.original_exec
        paths._chosen_data_dir = self.original_chosen
        shutil.rmtree(self.dir, ignore_errors=True)

    def test_frozen_app_uses_the_folder_next_to_the_exe(self):
        self.assertEqual(paths.app_dir(), self.exe_dir)
        self.assertEqual(paths.data_dir(), os.path.join(self.exe_dir, u"word工厂数据"))
        self.assertEqual(paths.temp_dir(),
                         os.path.join(self.exe_dir, u"word工厂数据", u"临时文件"))
        self.assertTrue(paths.data_dir_is_portable())

    def test_unwritable_exe_folder_falls_back_to_localappdata(self):
        """U 盘写保护 / 装在 Program Files → 数据落 %LOCALAPPDATA%，别让程序起不来。"""
        blocker = os.path.join(self.dir, u"只读位置")
        with open(blocker, "w", encoding="utf-8") as handle:      # 拿"文件"当目录用
            handle.write(u"x")
        sys.executable = os.path.join(blocker, u"子目录", u"word工厂.exe")
        fallback = os.path.join(self.dir, u"localappdata")
        original_env = os.environ.get("LOCALAPPDATA")
        os.environ["LOCALAPPDATA"] = fallback
        paths._chosen_data_dir = None
        try:
            self.assertFalse(paths._writable(os.path.join(blocker, u"子目录", u"word工厂数据")))
            data = paths.data_dir()
            self.assertEqual(data, os.path.join(fallback, u"word工厂", u"word工厂数据"))
            self.assertFalse(paths.data_dir_is_portable(), u"已经不在程序旁边了")
        finally:
            if original_env is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = original_env

    def test_writable_check_is_true_for_a_normal_folder(self):
        self.assertTrue(paths._writable(os.path.join(self.dir, u"能写")))
        self.assertFalse(os.path.exists(os.path.join(self.dir, u"能写", u".wf-write-test")),
                         u"探针文件用完要删掉，别留垃圾")


if __name__ == "__main__":
    unittest.main()
