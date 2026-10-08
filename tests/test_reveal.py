# -*- coding: utf-8 -*-
""""产物生成之后"的两件小事（用户 2026-10-08 报的致命 bug + 新要求）：

1. **不能弹控制台黑窗**：打包成 exe 之后进程没有控制台，这时用 ``subprocess`` 起
   ``taskkill`` 会被 Windows 新开一个黑窗口闪一下 —— 用户点「导出 PDF」看到"频繁弹终端窗口"，
   而 PDF 其实早导好了。所有外部命令都走 :mod:`wordfactory.subproc`（带 ``CREATE_NO_WINDOW``）。
2. **导完要弹窗告诉用户**，并给"打开文件 / 打开所在位置 / 什么都不做" —— 对应
   ``POST /api/reveal``（**只认本次跑出来的产物**，白名单之外一律拒绝）。
"""

import io
import json
import os
import shutil
import subprocess
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from unittest.mock import patch

from wordfactory import subproc
from wordfactory.gui import server as gui_server

from . import fixtures


class SubprocCase(unittest.TestCase):
    def test_run_hides_the_console_window(self):
        calls = {}

        def fake_run(command, **kwargs):
            calls["command"] = command
            calls["kwargs"] = kwargs
            return "ok"

        with patch.object(subprocess, "run", fake_run):
            subproc.run(["taskkill", "/PID", "1", "/T", "/F"], capture_output=True)
        self.assertEqual(calls["command"][0], "taskkill")
        self.assertTrue(calls["kwargs"].get("creationflags"),
                        u"必须带 CREATE_NO_WINDOW，否则 exe 里会闪黑窗口")
        self.assertEqual(calls["kwargs"]["creationflags"], subprocess.CREATE_NO_WINDOW)

    def test_popen_hides_the_console_window_too(self):
        calls = {}

        def fake_popen(command, **kwargs):
            calls["kwargs"] = kwargs
            return "proc"

        with patch.object(subprocess, "Popen", fake_popen):
            subproc.popen(["explorer", "/select,", "x.docx"])
        self.assertEqual(calls["kwargs"].get("creationflags"), subprocess.CREATE_NO_WINDOW)

    def test_caller_flags_are_not_overwritten(self):
        calls = {}

        def fake_run(command, **kwargs):
            calls["kwargs"] = kwargs
            return "ok"

        with patch.object(subprocess, "run", fake_run):
            subproc.run(["x"], creationflags=123)
        self.assertEqual(calls["kwargs"]["creationflags"], 123,
                         u"调用方自己指定的 flags 不能被覆盖")


class RevealCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp(prefix="wf_reveal_")
        cls.server = gui_server.ThreadingHTTPServer(
            ("127.0.0.1", 0), type("BoundHandler", (gui_server.Handler,),
                                   {"root": os.path.abspath(cls.dir)}))
        cls.port = cls.server.server_address[1]
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:%d" % cls.port
        cls.output = os.path.join(cls.dir, u"报告（正式版）.docx")
        fixtures.write_fixture(cls.output, body=fixtures.paragraph(fixtures.run(u"正文")))
        # 登记进"本次产物"白名单（跟真跑一轮产物时走的是同一个函数）
        cls.name = gui_server.allow_download(cls.output)

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        shutil.rmtree(cls.dir, ignore_errors=True)

    def post(self, path, payload):
        request = urllib.request.Request(
            self.base + path, data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST")
        try:
            with urllib.request.urlopen(request) as response:
                return json.loads(response.read().decode("utf-8")), response.status
        except urllib.error.HTTPError as error:
            return json.loads(error.read().decode("utf-8")), error.code

    def test_open_uses_the_default_program(self):
        opened = []
        with patch("os.startfile", create=True,
                   side_effect=lambda path: opened.append(path)):
            data, status = self.post("/api/reveal", {"name": self.name, "action": "open"})
        self.assertEqual(status, 200)
        self.assertTrue(data["ok"])
        self.assertEqual(opened, [self.output])

    def test_folder_reveals_the_file_in_explorer(self):
        calls = []
        with patch.object(subproc, "popen",
                          side_effect=lambda command, **kwargs: calls.append(command)):
            data, _status = self.post("/api/reveal", {"name": self.name, "action": "folder"})
        self.assertTrue(data["ok"])
        self.assertEqual(calls[0][:2], ["explorer", "/select,"])
        self.assertEqual(calls[0][2], self.output)

    def test_a_file_we_did_not_produce_is_refused(self):
        """白名单之外一律拒绝 —— 否则这个接口就等于"随便开本机任何文件"。"""
        opened = []
        with patch("os.startfile", create=True,
                   side_effect=lambda path: opened.append(path)):
            data, _status = self.post("/api/reveal",
                                      {"name": u"别人的文件.docx", "action": "open"})
        self.assertFalse(data["ok"])
        self.assertIn(u"找不到", data.get("error", u""))
        self.assertEqual(opened, [], u"没登记过的文件一个都不许开")


if __name__ == "__main__":
    unittest.main()
