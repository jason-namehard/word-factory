# -*- coding: utf-8 -*-
"""2026-10-09 第二轮加固的四条（用户复核上一轮后指出的"安全过头"副作用口子）。

上一轮把"杀错人"堵死了，但留了三个**静默退化**的口子 —— 不会再杀错，却会悄悄清不掉自家实例，
最后又回到用户最初抱怨的"卡 60 秒 + 页码全错"，比误杀更难发现：

* **必做 1**：闸② 把 prog_id 映射到**单个**映像名 → WPS 当 Word 默认打开程序的机器上，
  `Word.Application` 起出来的是 `wps.exe`，被判"映像名不符"→ 自家实例永远清不掉；
  另一处：新增候选忘了同步映射 → `expected_name` 为 None → 静默全跳过。
  → 改成**允许的映像名集合**，并加一条"每个 ProgID 都必须有映射"的测试当场红。
* **必做 2**：`DispatchEx` 把进程拉起来之后自己抛异常 → 那批 pid 不在名单里，清不掉。
  → 异常分支里再拍一次快照补收。
* **必做 3**：扫尾与 `/T` 都删掉之后，"WPS 慢半拍起的子孙"会变成清不掉的残留。
  → 用**窄窗口 diff**（`Quit` 前后）替代，候选仍过全部闸门；`expected_dir` 复用本次会话的，
  为空就整轮放弃（绝不退化成"无目录校验"）。

测试全程打桩，**绝不真的调 taskkill、绝不真的起 Office**。
"""

import contextlib
import os
import time
import unittest
from unittest.mock import patch

from wordfactory import officecom
from wordfactory import subproc as subproc_module

WINWORD_DIR = r"C:\Program Files\Microsoft Office\Root\Office16"
WINWORD = WINWORD_DIR + r"\WINWORD.EXE"
WPS_DIR = r"D:\wps\WPS Office\12.1.0.28505\office6"
WPS = WPS_DIR + r"\wps.exe"


class ExpectedImageMappingCase(unittest.TestCase):
    """必做 1：闸② 认"允许的映像名集合"，而不是单个名字。"""

    def test_every_prog_id_we_use_has_an_expected_image_mapping(self):
        """**加候选忘了同步映射**要当场红，别等线上静默全跳过。"""
        from wordfactory.ops import pdf as pdf_op
        used = {prog_id for _name, prog_id in pdf_op._word_prog_ids()}
        used |= {u"Word.Application", u"KWPS.Application", u"WPS.Application"}
        missing = sorted(prog_id for prog_id in used
                         if prog_id not in officecom._EXPECTED_IMAGE)
        self.assertEqual(missing, [], u"这些 ProgID 没有配「允许的映像名」：%s" % missing)
        for prog_id, names in officecom._EXPECTED_IMAGE.items():
            self.assertTrue(names, prog_id)
            self.assertTrue(all(name in officecom._OFFICE_EXES for name in names),
                            u"%s 的映像名必须在白名单里：%s" % (prog_id, names))

    def test_word_prog_id_may_start_wps_and_that_is_ours(self):
        """真机场景：起 Word.Application，实际起出来的是 wps.exe（路径也在 WPS 目录）
        → **必须过闸、必须被清掉**。"""
        calls = []
        started = time.time()
        info = {4242: (u"wps.exe", WPS, started + 1)}
        expected_dir = os.path.dirname(os.path.normcase(WPS))
        with patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False), \
             patch.object(officecom, "_process_ids", return_value={4242}), \
             patch.object(subproc_module, "run",
                          side_effect=lambda cmd, **kw: calls.append(cmd)):
            reason = officecom._gate_reason(4242, u"Word.Application", started, expected_dir)
            self.assertIsNone(reason, u"Word.Application 起出 wps.exe 必须过闸：%s" % reason)
            verified, _dir, skipped = officecom._verified_pids(
                {4242}, u"Word.Application", started, expected_dir=expected_dir)
            self.assertEqual(verified, {4242})
            self.assertEqual(skipped, [])
            self.assertTrue(officecom._terminate(4242, u"Word.Application", started,
                                                 expected_dir))
        self.assertEqual(len(calls), 1, u"该清掉的必须清掉")


class StartFailureResidueCase(unittest.TestCase):
    """必做 2：`_create_app()` 把进程拉起来之后自己抛异常 —— 那批也要收干净。"""

    def setUp(self):
        self.calls = []
        self._run = patch.object(subproc_module, "run",
                                 side_effect=lambda cmd, **kw: self.calls.append(cmd))
        self._run.start()
        self.addCleanup(self._run.stop)
        self._grace = patch.object(officecom, "QUIT_GRACE", 0)
        self._grace.start()
        self.addCleanup(self._grace.stop)

    def test_a_process_left_by_a_failed_dispatch_is_cleaned(self):
        state = {"ids": set()}
        info = {}

        def fake_create(prog_id):
            # DispatchEx 已经把进程拉起来、然后自己炸了
            info[5150] = (u"winword.exe", WINWORD, time.time() + 1)
            state["ids"] = {1, 5150}
            raise RuntimeError(u"DispatchEx 炸了")

        with patch.object(officecom, "_create_app", side_effect=fake_create), \
             patch.object(officecom, "_process_ids",
                          side_effect=lambda: set(state["ids"])), \
             patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False):
            with self.assertRaises(RuntimeError):
                officecom.Session(u"Word.Application")
        self.assertEqual(len(self.calls), 1, u"起实例炸了之后留下的进程必须被收掉")
        self.assertIn("5150", self.calls[0])


class NarrowWindowSweepCase(unittest.TestCase):
    """必做 3：窄窗口 diff —— 只抓 `Quit` 前后新冒出来的自家进程。"""

    def setUp(self):
        self.calls = []
        self._run = patch.object(subproc_module, "run",
                                 side_effect=lambda cmd, **kw: self.calls.append(cmd))
        self._run.start()
        self.addCleanup(self._run.stop)
        self._grace = patch.object(officecom, "QUIT_GRACE", 0)
        self._grace.start()
        self.addCleanup(self._grace.stop)

    @contextlib.contextmanager
    def _scenario(self, info, before_quit, after_quit, prog_id=u"KWPS.Application",
                  visible=None):
        """造场景：**起实例之前**机器上没有我们的进程，起完才有主进程（before_quit），
        `Quit` 之后 ids 变成 after_quit。

        用上下文管理器，是因为 `close()` 也必须在这套假进程表里跑 —— 补丁提前退出的话，
        `close()` 就换成真实进程表了（那样测不出窄窗口逻辑）。
        """
        state = {"phase": "pre"}

        class App(object):
            Version = "12.0"

            def Quit(self, *_args):
                state["phase"] = "after_quit"

        def fake_ids():
            if state["phase"] == "pre":
                return {1, 2}                    # 起实例之前：只有无关进程
            return set(before_quit if state["phase"] == "before_quit" else after_quit)

        def fake_create(prog_id):
            state["phase"] = "before_quit"       # 起完实例 → 主进程出现了
            return App()

        with patch.object(officecom, "_create_app", side_effect=fake_create), \
             patch.object(officecom, "_process_ids", side_effect=fake_ids), \
             patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window",
                          side_effect=visible or (lambda pid: False)):
            yield officecom.Session(prog_id)

    def test_a_slow_grandchild_is_caught_and_cleaned(self):
        info = {6001: (u"wps.exe", WPS, time.time() + 1),
                6002: (u"wps.exe", WPS, time.time() + 2)}   # 慢半拍的那个
        with self._scenario(info, before_quit={6001}, after_quit={6002}) as session:
            self.assertEqual(session.expected_dir, os.path.dirname(os.path.normcase(WPS)))
            session.close()
        self.assertEqual(len(self.calls), 1, u"慢半拍的子孙要被抓住并清掉")
        self.assertIn("6002", self.calls[0])

    def test_system_pids_in_the_window_are_skipped(self):
        info = {6001: (u"wps.exe", WPS, time.time() + 1),
                6002: (u"wps.exe", WPS, time.time() + 2),
                6003: (u"", u"", 0.0)}                     # 读不到映像名（系统/受保护进程）
        with self._scenario(info, before_quit={6001}, after_quit={6002, 6003}) as session:
            session.close()
        self.assertEqual(len(self.calls), 1)
        self.assertIn("6002", self.calls[0])
        self.assertNotIn("6003", self.calls[0], u"读不到名字的一律不动")

    def test_a_visible_word_in_the_window_is_skipped(self):
        info = {6005: (u"winword.exe", WINWORD, time.time() + 1),   # 起会话时的主进程
                6004: (u"winword.exe", WINWORD, time.time() + 2)}   # 窗口里冒出来的
        with self._scenario(info, before_quit={6005}, after_quit={6004},
                            prog_id=u"Word.Application",
                            visible=lambda pid: pid == 6004) as session:
            self.assertEqual(session.expected_dir, WINWORD_DIR.lower())
            session.close()
        self.assertEqual(self.calls, [], u"有窗口的（用户自己开着的）Word 不许动")

    def test_no_expected_dir_means_no_narrow_window_cleanup(self):
        """`expected_dir` 为空 → 一个都不许动（不许退化成"无目录校验"）。"""
        info = {6002: (u"wps.exe", WPS, time.time() + 2)}
        with self._scenario(info, before_quit={6009}, after_quit={6002}) as session:
            self.assertEqual(session.pids, set(), u"起会话时没有可验证的进程")
            self.assertEqual(session.expected_dir, u"")
            session.close()
            self.assertTrue(any(u"放弃窄窗口清理" in line
                                for line in session.cleanup_log))
        self.assertEqual(self.calls, [], u"没有期望目录 → 放弃窄窗口清理，一个都不动")


class StatsAndAuditLogCase(unittest.TestCase):
    """顺手项 4：清理详情要落 `cleanup.log`（exe 里 print 没人看得见），并给一行统计。"""

    def test_cleanup_lines_are_written_to_the_log_file(self):
        import tempfile
        import shutil
        folder = tempfile.mkdtemp(prefix="wf_cleanup_")
        try:
            with patch.object(officecom, "cleanup_log_path",
                              return_value=os.path.join(folder, u"cleanup.log")):
                officecom._log_cleanup(u"结束自己起的 Office 进程：pid=1 映像=winword.exe")
            text = open(os.path.join(folder, u"cleanup.log"), encoding="utf-8").read()
            self.assertIn(u"pid=1 映像=winword.exe", text)
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def test_stats_line_says_what_happened(self):
        stats = officecom.new_stats()
        stats["sessions"] = 2
        stats["killed"].append({"pid": 11, "image": "winword.exe", "path": WINWORD,
                                "created": 1.0})
        stats["skipped"].append({"pid": 12, "reason": u"读不到映像名（多半是系统/受保护进程）"})
        line = officecom.format_stats(stats)
        self.assertIn(u"起 2 个", line)
        self.assertIn(u"清掉 1 个", line)
        self.assertIn(u"跳过 1 个", line)
        self.assertIn(u"读不到映像名", line)
        self.assertIn(u"pid 11", line)

    def test_a_kill_is_recorded_in_the_stats(self):
        calls = []
        started = time.time()
        info = {77: (u"winword.exe", WINWORD, started + 1)}
        stats = officecom.new_stats()
        with patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False), \
             patch.object(officecom, "_process_ids", return_value={77}), \
             patch.object(subproc_module, "run",
                          side_effect=lambda cmd, **kw: calls.append(cmd)):
            officecom._terminate(77, u"Word.Application", started,
                                 WINWORD_DIR.lower(), stats)
        self.assertEqual(len(stats["killed"]), 1)
        self.assertEqual(stats["killed"][0]["pid"], 77)
        self.assertEqual(stats["killed"][0]["image"], u"winword.exe")

    def test_a_skip_is_recorded_with_its_reason(self):
        stats = officecom.new_stats()
        with patch.object(officecom, "_process_info", return_value=(u"", u"", 0.0)), \
             patch.object(officecom, "_process_ids", return_value={204}), \
             patch.object(subproc_module, "run") as fake_run:
            officecom._terminate(204, u"Word.Application", time.time(), u"x", stats)
        self.assertEqual(stats["killed"], [])
        self.assertEqual(len(stats["skipped"]), 1)
        self.assertIn(u"读不到映像名", stats["skipped"][0]["reason"])
        self.assertEqual(fake_run.call_count, 0)


if __name__ == "__main__":
    unittest.main()
