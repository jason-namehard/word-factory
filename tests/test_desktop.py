# -*- coding: utf-8 -*-
"""桌面版（exe）相关的小单测。

这里**不开真窗口**（窗口是 pywebview + 系统 WebView2，纯人工看）—— 只钉住
"桌面版靠的那两块地基"真的能立起来：

* `gui_server.make_server()`：桌面版要"先把服务建好、拿到真实 URL 再交给窗口"，这是那条路；
* `_free_port()`：桌面版不再死磕 8765（撞端口就起不来），改成取一个空闲端口。
"""

import os
import shutil
import socket
import tempfile
import threading
import unittest
import urllib.request

from wordfactory import desktop
from wordfactory.gui import server as gui_server


class TestDesktopGlue(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp(prefix="wf_desktop_")
        cls.server = None
        cls.thread = None

    @classmethod
    def tearDownClass(cls):
        if cls.server is not None:
            cls.server.shutdown()
            cls.server.server_close()
        shutil.rmtree(cls.dir, ignore_errors=True)

    def test_free_port_is_actually_free(self):
        port = desktop._free_port()
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.bind(("127.0.0.1", port))          # 能绑上 = 真空闲
        finally:
            sock.close()

    def test_make_server_gives_a_url_and_serves_the_page(self):
        server, url = gui_server.make_server(host="127.0.0.1", port=0,
                                             plans=os.path.join(self.dir, "plans"))
        self.assertTrue(url.startswith("http://127.0.0.1:"))
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with urllib.request.urlopen(url, timeout=20) as response:
                page = response.read().decode("utf-8")
            self.assertIn(u"word工厂", page)
            self.assertIn(u"/api/run", page)
        finally:
            server.shutdown()
            server.server_close()

    def test_desktop_main_exists(self):
        self.assertTrue(callable(desktop.main))
