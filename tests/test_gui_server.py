# -*- coding: utf-8 -*-
"""GUI 后端的单测：起真服务、发真请求（不 mock）。

要钉住的：

* 页面能出来、步骤表/模板表/表格清单能读；
* `/api/run` 按配方跑，**返回报告文本 + 可下载的文件名**；
* 下载只给本次跑出来的文件（白名单）—— 不把文件系统暴露出去；
* 目录浏览限定在 root 里（跳出 root 要报错）；
* 只监听 127.0.0.1。
"""

import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.parse
import urllib.request

from wordfactory.gui import server as gui_server
from wordfactory.ooxml import qn

from . import fixtures


class GuiCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp(prefix="wf_gui_")
        cls.path = os.path.join(cls.dir, u"夹具.docx")
        body = (fixtures.paragraph(fixtures.run(u"表2.3-1     库容特性表"))
                + fixtures.table([[u"<w:r><w:t>水位</w:t></w:r>",
                                   u"<w:r><w:t>库容</w:t></w:r>"]]))
        fixtures.write_fixture(cls.path, body=body)
        cls.server = gui_server.ThreadingHTTPServer(
            ("127.0.0.1", 0), type("BoundHandler", (gui_server.Handler,),
                                   {"root": os.path.abspath(cls.dir)}))
        cls.port = cls.server.server_address[1]
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:%d" % cls.port

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        shutil.rmtree(cls.dir, ignore_errors=True)

    def get(self, path):
        with urllib.request.urlopen(self.base + path) as response:
            return response.read(), response.status

    def post(self, path, payload):
        request = urllib.request.Request(
            self.base + path, data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(request) as response:
            return json.loads(response.read().decode("utf-8")), response.status

    # ------------------------------------------------------------------ 页面与元数据
    def test_the_page_comes_up(self):
        body, status = self.get("/")
        self.assertEqual(status, 200)
        text = body.decode("utf-8")
        self.assertIn(u"word 工厂", text)
        self.assertIn(u"/api/run", text, u"页面要能调到 run 接口")

    def test_steps_and_templates_are_listed(self):
        body, _ = self.get("/api/steps")
        names = [step["op"] for step in json.loads(body.decode("utf-8"))["steps"]]
        self.assertIn("captions", names)
        self.assertNotIn("fonts", names, u"字体是收尾，不是配方步骤")

        body, _ = self.get("/api/templates")
        styles = json.loads(body.decode("utf-8"))["styles"]
        self.assertTrue(any(style["name"] == u"通用款·外粗内细" for style in styles),
                        u"模板表要能读出来（默认款要在）")

    def test_tables_are_listed_by_their_header(self):
        body, _ = self.get("/api/tables?path=" + urllib.parse.quote(self.path))
        tables = json.loads(body.decode("utf-8"))["tables"]
        self.assertEqual(len(tables), 1)
        self.assertEqual(tables[0]["header"], [u"水位", u"库容"])

    # ------------------------------------------------------------------ 跑 + 下载
    def test_run_returns_a_report_and_a_download(self):
        payload, status = self.post("/api/run", {"path": self.path,
                                                 "steps": ["captions"],
                                                 "mode": "verify"})
        self.assertEqual(status, 200)
        self.assertIn(u"captions", payload["text"])
        self.assertTrue(payload["download"], u"要给出下载名")
        self.assertTrue(os.path.exists(payload["report"]["out"]))

        body, status = self.get("/api/download?name=" + urllib.parse.quote(payload["download"]))
        self.assertEqual(status, 200)
        self.assertGreater(len(body), 1000, u"下载到的应该是个 docx")

    def test_formal_mode_reports_the_audit(self):
        payload, _ = self.post("/api/run", {"path": self.path, "steps": ["captions"],
                                            "mode": "formal"})
        self.assertEqual(payload["report"]["audit"]["verdict"], "PASS")
        self.assertIn(u"AUDIT=PASS", payload["text"])

    def test_dry_run_does_not_write(self):
        payload, _ = self.post("/api/run", {"path": self.path, "steps": ["captions"],
                                            "mode": "verify", "dry_run": True})
        self.assertTrue(payload["report"]["dry_run"])
        self.assertIsNone(payload["download"])

    def test_an_unknown_step_is_rejected_with_the_step_list(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.post("/api/run", {"path": self.path, "steps": ["nope"], "mode": "verify"})
        body = caught.exception.read().decode("utf-8")
        self.assertIn(u"nope", body)
        self.assertIn(u"captions", body, u"要列出已注册的宏")

    def test_an_empty_recipe_is_allowed_it_only_finishes(self):
        """空配方合法：正式版 = 只做收尾（通体黑 + 字体）。GUI 那边会提示至少勾一个。"""
        payload, _ = self.post("/api/run", {"path": self.path, "steps": [],
                                            "mode": "formal"})
        self.assertEqual(payload["report"]["mode"], "formal")
        self.assertTrue(payload["download"])

    def test_an_unknown_download_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError):
            self.get("/api/download?name=" + urllib.parse.quote(u"../../../etc/passwd"))

    # ------------------------------------------------------------------ 浏览边界
    def test_browsing_outside_the_root_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError):
            self.get("/api/browse?dir=" + urllib.parse.quote(os.path.dirname(self.dir)))

    def test_reading_a_missing_file_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError):
            self.get("/api/read?path=" + urllib.parse.quote(os.path.join(self.dir, u"没有.json")))


if __name__ == "__main__":
    unittest.main()
