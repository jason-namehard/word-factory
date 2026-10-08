# -*- coding: utf-8 -*-
"""Word/WPS COM 会话的纪律 —— 重点是**结束进程这件事绝不能再伤人**。

⚠️ 2026-10-09 事故（本文件的由来）：``_is_office_app()`` 原来写成
``(not name) or (name in _OFFICE_EXES)`` —— 读不到映像名（系统/受保护进程必然读不到）
被当成"这是我们起的 Office"；又叠加"扫全机新进程差集"的兜底 + ``taskkill /T``，
结果把会话宿主（sihost/svchost）连树杀掉 → **整个 Windows 外壳重建**（当晚 5 次）。

现在必须守住的：
1. 读不到映像名 → **不是我们的**，跳过；
2. 白名单外的进程（svchost / explorer / sihost…）→ 跳过；
3. 白名单内**且**路径在本次 Office 目录里**且**创建时间晚于本次会话开始 → 才是我们的，才允许结束；
4. 候选集**只能**是 DispatchEx 前后 diff 出来的 pid（不许扫全机）；
5. ``_process_ids()`` 失败返回 ``None`` 并放弃本轮清理（绝不返回空集）；
6. 真的结束进程时打印 pid / 映像名 / 完整路径 / 创建时间，供人工核对。

测试里**一律打桩**，绝不真的调 taskkill。
"""

import ctypes
import time
import unittest
from unittest.mock import patch

from wordfactory import officecom
from wordfactory import subproc as subproc_module

WINWORD_DIR = r"C:\Program Files\Microsoft Office\Root\Office16"
WINWORD = WINWORD_DIR + r"\WINWORD.EXE"


class ProcessInfoCase(unittest.TestCase):
    """闸门本身：三道闸逐条单独验。"""

    def setUp(self):
        self.started = 1000.0
        self.expected_dir = WINWORD_DIR.lower()
        self.info = {}
        self._patch = patch.object(officecom, "_process_info",
                                   side_effect=lambda pid: self.info.get(pid, (u"", u"", 0.0)))
        self._patch.start()
        self.addCleanup(self._patch.stop)
        self._window = patch.object(officecom, "_has_visible_window", return_value=False)
        self._window.start()
        self.addCleanup(self._window.stop)
        # 闸门口会先看"这个 pid 还在不在"——测试里统一让它们都"在"
        self._alive = patch.object(officecom, "_process_ids",
                                  return_value=set(range(1, 1000)))
        self._alive.start()
        self.addCleanup(self._alive.stop)

    def reason(self, pid, prog_id=u"Word.Application", created=1001.0, name=u"winword.exe",
               path=WINWORD, expected_dir=None):
        self.info[pid] = (name, path, created)
        return officecom._gate_reason(pid, prog_id, self.started,
                                     self.expected_dir if expected_dir is None else expected_dir)

    def test_gate_one_unreadable_image_name_is_not_ours(self):
        """**事故根因**：读不到映像名（Secure System / Registry / csrss 必然读不到）→ 跳过。"""
        self.assertFalse(officecom._is_office_app(204))
        reason = self.reason(204, name=u"", path=u"", created=0.0)
        self.assertIsNotNone(reason, u"读不到名字的 pid 绝不能过闸")
        self.assertIn(u"读不到映像名", reason)

    def test_gate_one_whitelist_only(self):
        for name in (u"svchost.exe", u"explorer.exe", u"sihost.exe", u"csrss.exe",
                     u"python.exe", u"taskkill.exe"):
            self.assertFalse(officecom._is_office_app(4), name)
            reason = self.reason(4, name=name, path=r"C:\Windows\System32\%s" % name)
            self.assertIn(u"白名单", reason, u"%s 不该过闸" % name)

    def test_gate_two_path_must_be_this_office(self):
        """名字对但路径不对（别处的同名程序）→ 跳过。"""
        reason = self.reason(5, path=r"D:\别人的目录\WINWORD.EXE")
        self.assertIn(u"路径", reason)

    def test_gate_two_no_expected_dir_means_no_kill(self):
        reason = self.reason(7, path=WINWORD, expected_dir=u"")
        self.assertIn(u"期望目录", reason)

    def test_gate_three_creation_time_must_be_after_the_session(self):
        reason = self.reason(8, created=999.0)
        self.assertIn(u"创建时间早于", reason)
        reason = self.reason(8, created=0.0)
        self.assertIn(u"创建时间", reason)

    def test_a_pid_with_a_visible_window_is_never_killed(self):
        """用户自己开着的 Office（有窗口）绝不许动。"""
        self.info[9] = (u"winword.exe", WINWORD, 1001.0)
        with patch.object(officecom, "_has_visible_window", return_value=True):
            reason = officecom._gate_reason(9, u"Word.Application", self.started,
                                            self.expected_dir)
        self.assertIn(u"可见窗口", reason)

    def test_the_good_case_passes_all_gates(self):
        self.assertIsNone(self.reason(10))

    def test_wrong_office_binary_for_this_session_is_skipped(self):
        """本次起的是 Word，就不要去动 wps.exe（名字在白名单里也不能杀）。"""
        reason = self.reason(11, prog_id=u"Word.Application", name=u"wps.exe",
                             path=r"D:\wps\WPS Office\12.1\office6\wps.exe")
        self.assertIn(u"不符", reason)


class VerifiedCandidateCase(unittest.TestCase):
    """候选集只会来自 diff；混合一批系统 pid 时，只有我们的那个留下。"""

    def test_only_the_verified_office_pid_survives(self):
        started = time.time()
        info = {
            100: (u"winword.exe", WINWORD, started + 1),
            200: (u"", u"", 0.0),                       # 读不到名字（系统进程）
            300: (u"svchost.exe", r"C:\Windows\System32\svchost.exe", started + 1),
            400: (u"explorer.exe", r"C:\Windows\explorer.exe", started + 1),
            500: (u"WINWORD.EXE".lower(), WINWORD, started - 5),   # 比会话还老
        }
        with patch.object(officecom, "_process_info", side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False), \
             patch.object(officecom, "_process_ids", return_value=set(info)):
            verified, expected_dir, skipped = officecom._verified_pids(
                set(info), u"Word.Application", started)
        self.assertEqual(verified, {100}, u"只有 100 是我们的")
        self.assertEqual(expected_dir, WINWORD_DIR.lower())
        self.assertEqual({pid for pid, _r in skipped}, {200, 300, 400, 500})


class KillCase(unittest.TestCase):
    """真杀之前：再核一遍闸门 + 打印详情；测试里 subproc 全打桩。"""

    def setUp(self):
        self.started = time.time()
        self.calls = []
        self._run = patch.object(subproc_module, "run",
                                 side_effect=lambda cmd, **kw: self.calls.append(cmd))
        self._run.start()
        self.addCleanup(self._run.stop)
        self._alive = patch.object(officecom, "_process_ids",
                                   return_value=set(range(1, 10000)))
        self._alive.start()
        self.addCleanup(self._alive.stop)

    def test_a_verified_pid_is_terminated_without_tree_flag(self):
        info = {66: (u"winword.exe", WINWORD, self.started + 1)}
        with patch.object(officecom, "_process_info", side_effect=lambda pid: info[pid]), \
             patch.object(officecom, "_has_visible_window", return_value=False):
            ok = officecom._terminate(66, u"Word.Application", self.started,
                                      WINWORD_DIR.lower())
        self.assertTrue(ok)
        self.assertEqual(len(self.calls), 1)
        command = self.calls[0]
        self.assertEqual(command[0], "taskkill")
        self.assertIn("66", command)
        self.assertNotIn("/T", command,
                         u"绝不用 /T：事故里正是连「子孙整棵树」一起杀才端掉会话宿主的")

    def test_an_unreadable_pid_is_never_terminated(self):
        with patch.object(officecom, "_process_info", return_value=(u"", u"", 0.0)):
            ok = officecom._terminate(204, u"Word.Application", self.started,
                                      WINWORD_DIR.lower())
        self.assertFalse(ok)
        self.assertEqual(self.calls, [], u"读不到信息的 pid 一律不许动手")

    def test_a_system_process_is_never_terminated(self):
        info = {300: (u"svchost.exe", r"C:\Windows\System32\svchost.exe", self.started + 1)}
        with patch.object(officecom, "_process_info", side_effect=lambda pid: info[pid]):
            ok = officecom._terminate(300, u"Word.Application", self.started,
                                      WINWORD_DIR.lower())
        self.assertFalse(ok)
        self.assertEqual(self.calls, [])

    def test_kill_registered_clears_the_registry(self):
        info = {66: (u"winword.exe", WINWORD, self.started + 1)}
        registry = {66: (u"Word.Application", self.started, WINWORD_DIR.lower()),
                    204: (u"Word.Application", self.started, WINWORD_DIR.lower())}
        with patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False):
            killed = officecom.kill_registered(registry)
        self.assertEqual(killed, [66], u"只有全过闸的那个被杀")
        self.assertEqual(registry, {}, u"登记表用完即清（不跨会话累积）")


class NoMachineWideSweepCase(unittest.TestCase):
    """把"扫全机"那套彻底拿掉：没有兜底扫描、没有模块级登记表。"""

    def test_there_is_no_machine_wide_sweep(self):
        self.assertFalse(hasattr(officecom, "_sweep_late_strays"),
                         u"扫全机新进程差集的兜底必须删掉（它会误伤系统进程）")
        self.assertFalse(hasattr(officecom, "terminate_leftovers"))
        self.assertFalse(hasattr(officecom, "_OUR_PIDS"),
                         u"模块级登记表不许存在（会跨会话累积）")

    def test_process_ids_returns_none_when_the_api_fails(self):
        class BrokenPsapi(object):
            def EnumProcesses(self, *args):
                return 0

        class BrokenWindll(object):
            psapi = BrokenPsapi()

        with patch.object(ctypes, "windll", BrokenWindll()):
            self.assertIsNone(officecom._process_ids(),
                              u"拿不到进程表必须是 None —— 空集会放大成「全机进程」")

    def test_process_ids_returns_none_when_the_call_raises(self):
        class BoomPsapi(object):
            def EnumProcesses(self, *args):
                raise OSError(u"拒绝访问")

        class BoomWindll(object):
            psapi = BoomPsapi()

        with patch.object(ctypes, "windll", BoomWindll()):
            self.assertIsNone(officecom._process_ids())


class SessionCase(unittest.TestCase):
    """会话层：拿不到快照就不清理；只动 diff 出来的、过了闸的 pid。"""

    def setUp(self):
        self.calls = []
        self._run = patch.object(subproc_module, "run",
                                 side_effect=lambda cmd, **kw: self.calls.append(cmd))
        self._run.start()
        self.addCleanup(self._run.stop)
        self._grace = patch.object(officecom, "QUIT_GRACE", 0)
        self._grace.start()
        self.addCleanup(self._grace.stop)

    def test_no_snapshot_means_no_cleanup_at_all(self):
        class App(object):
            Version = "16.0"

            def Quit(self, *_args):
                pass

        with patch.object(officecom, "_create_app", return_value=App()), \
             patch.object(officecom, "_process_ids", return_value=None):
            session = officecom.Session(u"Word.Application")
        self.assertEqual(session.pids, set())
        self.assertTrue(any(u"不做任何清理" in line for line in session.cleanup_log))
        session.close()
        self.assertEqual(self.calls, [], u"拿不到进程表就一个都不许杀")

    def test_only_the_diffed_and_gated_pid_is_killed(self):
        started = {}
        info = {}

        class App(object):
            Version = "16.0"

            def Quit(self, *_args):
                pass

        def fake_create(prog_id):
            # 起实例"顺便"让系统里多出一个系统进程和一个我们的 Office 进程
            info[8001] = (u"winword.exe", WINWORD, time.time() + 1)
            info[8002] = (u"", u"", 0.0)
            return App()

        def fake_ids():
            return {1, 2, 8001, 8002} if info else {1, 2}

        with patch.object(officecom, "_create_app", side_effect=fake_create), \
             patch.object(officecom, "_process_ids", side_effect=fake_ids), \
             patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False):
            session = officecom.Session(u"Word.Application")
            started["pids"] = set(session.pids)
            session.close()
        self.assertEqual(started["pids"], {8001}, u"只有我们的那个进候选")
        self.assertEqual(len(self.calls), 1)
        self.assertIn("8001", self.calls[0])
        self.assertNotIn("8002", self.calls[0], u"系统进程绝不许碰")


class RunCase(unittest.TestCase):
    def test_original_failure_is_returned_instead_of_unknown_reason(self):
        with patch.object(officecom, "Session",
                          side_effect=AttributeError("Open.ExportAsFixedFormat")):
            with self.assertRaises(officecom.OfficeError) as caught:
                officecom.run(lambda session: None, ["Word.Application"], timeout=3)
        self.assertIn("Open.ExportAsFixedFormat", str(caught.exception))

    def test_failure_in_first_renderer_can_fall_back_and_closes_both_sessions(self):
        closed = []

        class Session(object):
            def __init__(self, prog_id, registry=None):
                self.prog_id = prog_id

            def close(self):
                closed.append(self.prog_id)

        def read(session):
            if session.prog_id == "Word.Application":
                raise RuntimeError("first renderer failed")
            return {"pages": 31}

        with patch.object(officecom, "Session", Session):
            result = officecom.run(read, ["Word.Application", "KWPS.Application"], timeout=3)
        self.assertEqual(result["pages"], 31)
        self.assertEqual(closed, ["Word.Application", "KWPS.Application"])

    def test_a_hanging_run_only_kills_registered_and_gated_pids(self):
        """超时兜底只动**本轮登记表**里、**仍过闸**的进程。"""
        calls = []
        state = {"ids": set()}
        info = {9001: (u"winword.exe", WINWORD, time.time() + 1),
                9002: (u"", u"", 0.0)}

        class App(object):
            Version = "16.0"

            def Quit(self, *_args):
                pass

        def fake_ids():
            return set(state["ids"])

        def fake_create(prog_id):
            # 起实例"顺便"让机器上多出两个系统 pid + 我们那个 Office pid
            state["ids"] = {1, 2, 9001, 9002}
            return App()

        with patch.object(officecom, "_create_app", side_effect=fake_create), \
             patch.object(officecom, "_process_ids", side_effect=fake_ids), \
             patch.object(officecom, "_process_info",
                          side_effect=lambda pid: info.get(pid, (u"", u"", 0.0))), \
             patch.object(officecom, "_has_visible_window", return_value=False), \
             patch.object(subproc_module, "run",
                          side_effect=lambda cmd, **kw: calls.append(cmd)):

            def hangs(session):
                time.sleep(5)
                return {"pages": 1}

            with self.assertRaises(officecom.OfficeError) as caught:
                officecom.run(hangs, ["Word.Application"], timeout=0.6)
        self.assertIn(u"没完成", str(caught.exception))
        self.assertEqual(len(calls), 1, u"只杀我们那个，系统进程不动")
        self.assertIn("9001", calls[0])


if __name__ == "__main__":
    unittest.main()
