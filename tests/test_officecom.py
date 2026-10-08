# -*- coding: utf-8 -*-
"""Word/WPS COM 会话的纪律（2026-10-08 用户报"页码全不对、还特别慢"之后重写）。

两条必须钉死的规矩：

1. **每次都自己起私有实例**（``DispatchEx``），**绝不**去借用用户已经开着的 Office ——
   实测：机器上只要有一个自动化残留的 Word，连上去之后 ``Documents.Open`` 不返回，
   整整 60 秒超时 → 读不到真实页码 → 首页/目录旁那两只空白页就删不掉。
2. **自己起的实例必须死透**：``Quit`` 之后还活着就按 pid 结束它自己 ——
   实测攒了 24 个 WINWORD + 32 个 wps 没有窗口的残留进程，越攒越慢。
"""

import unittest
from unittest.mock import patch

from wordfactory import officecom


class FakeApp(object):
    """假装是 Word/WPS 的 Application 对象（只实现我们用到的那点东西）。"""

    def __init__(self):
        self.Version = "16.0"
        self.quitted = False

    def Quit(self, *_args):
        self.quitted = True


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
            def __init__(self, prog_id):
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


class SessionCase(unittest.TestCase):
    def test_we_always_create_our_own_instance(self):
        """**不借别人的**：起手就是自己起一个（DispatchEx），不去连现成的 Office。"""
        made = []

        def fake_create(prog_id):
            made.append(prog_id)
            return FakeApp()

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", return_value={1, 2, 4242}):
            session = officecom.Session("Word.Application")
        self.assertEqual(made, ["Word.Application"])
        self.assertEqual(session.pids, set(), u"快照没变化时不该把别人的进程算成自己的")
        session.close()

    def test_close_quits_our_own_instance(self):
        state = {"ids": set()}
        app = FakeApp()

        def quit_and_die(*_args):
            app.quitted = True
            state["ids"] = set()          # Quit 之后进程真的退了

        app.Quit = quit_and_die

        def fake_create(prog_id):
            state["ids"] = {555}          # "起了一个新进程"
            return app

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", lambda: set(state["ids"])):
            session = officecom.Session("Word.Application")
            self.assertEqual(session.pids, {555}, u"新起来的那个进程才算我们的")
            session.close()
        self.assertTrue(app.quitted, u"自己起的实例必须 Quit")

    def test_a_stubborn_instance_is_killed_by_pid(self):
        """``Quit`` 之后还赖着不走 → 按 **pid** 结束它自己（绝不用 /IM 名字）。"""
        state = {"ids": set()}
        app = FakeApp()                   # 它的 Quit 什么都不做（赖着不走）
        killed = []

        def fake_create(prog_id):
            state["ids"] = {666}
            return app

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", lambda: set(state["ids"])), \
             patch.object(officecom, "_terminate",
                          side_effect=lambda pid: killed.append(pid)), \
             patch.object(officecom, "QUIT_GRACE", 0):
            session = officecom.Session("Word.Application")
            session.close()
        self.assertEqual(killed, [666])

    def test_open_is_read_only_and_not_added_to_recent(self):
        state = {"ids": set()}
        app = FakeApp()
        opened = {}

        class Documents(object):
            Count = 0

            def Open(self, path, confirm, readonly, add_recent):
                opened.update({"confirm": confirm, "readonly": readonly,
                               "add_recent": add_recent})
                return "DOC"

        app.Documents = Documents()

        def fake_create(prog_id):
            state["ids"] = {777}
            return app

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", lambda: set(state["ids"])):
            session = officecom.Session("Word.Application")
            session.open(u"C:\\tmp\\报告.docx")
        self.assertTrue(opened["readonly"], u"一律只读打开")
        self.assertFalse(opened["add_recent"], u"别往用户的「最近使用的文档」里塞")

    def test_a_hanging_instance_is_cleaned_up_by_pid(self):
        """卡住时按 **pid** 清掉我们自己起的实例 —— 不做"每次调用套一层线程"
        （那样 COM 会跑到没 ``CoInitialize`` 的线程上，报"尚未调用 CoInitialize"）。"""
        state = {"ids": set()}
        app = FakeApp()
        killed = []

        def fake_create(prog_id):
            state["ids"] = {888}              # 我们起了一个实例
            return app

        def hanging_callback(session):
            import time
            time.sleep(5)                     # 假装卡在 Documents.Open 上
            return {"pages": 1}

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", lambda: set(state["ids"])), \
             patch.object(officecom, "_terminate",
                          side_effect=lambda pid: killed.append(pid)), \
             patch.object(officecom, "_OUR_PIDS", set()):
            with self.assertRaises(officecom.OfficeError) as caught:
                officecom.run(hanging_callback, ["Word.Application"], timeout=0.6)
        self.assertIn(u"没完成", str(caught.exception))
        self.assertEqual(killed, [888], u"超时必须清掉自己起的实例，不能留残留")

    def test_leftover_registry_only_holds_processes_we_started(self):
        """超时清理只动"我们起过"的 pid —— 用户自己开的 Word 绝不能被牵连。"""
        state = {"ids": set()}
        app = FakeApp()

        def fake_create(prog_id):
            state["ids"] = {1111}
            return app

        with patch.object(officecom, "_create_app", fake_create), \
             patch.object(officecom, "_process_ids", lambda: set(state["ids"])), \
             patch.object(officecom, "_OUR_PIDS", set()) as registry:
            session = officecom.Session("Word.Application")
            self.assertIn(1111, officecom._OUR_PIDS)
            session.close()
            self.assertEqual(officecom._OUR_PIDS, set(),
                             u"正常关掉之后不该再留在待清理名单里")


if __name__ == "__main__":
    unittest.main()


class OfficeProcessFilterCase(unittest.TestCase):
    """收尾清理只动**Office 应用本体**：``DispatchEx`` 有时会连带拉起共享辅助服务
    （云同步之类），那不是我们的东西。"""

    def test_only_office_apps_are_terminated(self):
        with patch.object(officecom, "_image_name", return_value="wpscloudsvr.exe"):
            self.assertFalse(officecom._is_office_app(4321))
        with patch.object(officecom, "_image_name", return_value="winword.exe"):
            self.assertTrue(officecom._is_office_app(4321))
        with patch.object(officecom, "_image_name", return_value="wps.exe"):
            self.assertTrue(officecom._is_office_app(4321))

    def test_a_non_office_process_is_not_terminated(self):
        """不是 Office 本体的进程（辅助服务）一个都不许动。"""
        from wordfactory import subproc as subproc_module
        with patch.object(officecom, "_is_office_app", return_value=False), \
             patch.object(subproc_module, "run") as fake_run:
            officecom._terminate(999)
        self.assertEqual(fake_run.call_count, 0, u"不是 Office 本体的进程不许动")
