# -*- coding: utf-8 -*-
"""GUI 后端的单测：起真服务、发真请求（不 mock）。

要钉住的：

* 页面能出来、步骤表/模板表/表格清单能读；
* `/api/run` 按配方跑，**返回报告文本 + 可下载的文件名**；
* 下载只给本次跑出来的文件（白名单）—— 不把文件系统暴露出去；
* 目录浏览限定在 root 里（跳出 root 要报错）；
* 只监听 127.0.0.1。
"""

import io
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


class TestPlansAndDrives(GuiCase):
    """执行方案管理 + 盘符浏览（用户 2026-09-26 提的五条里的第 1、3 条）。"""

    @classmethod
    def setUpClass(cls):
        super(TestPlansAndDrives, cls).setUpClass()
        cls.plans = os.path.join(cls.dir, u"plans")
        cls.loose = gui_server.ThreadingHTTPServer(
            ("127.0.0.1", 0), type("LooseHandler", (gui_server.Handler,),
                                   {"root": None, "plans_dir": cls.plans}))
        cls.loose_port = cls.loose.server_address[1]
        cls.loose_thread = threading.Thread(target=cls.loose.serve_forever, daemon=True)
        cls.loose_thread.start()
        cls.loose_base = "http://127.0.0.1:%d" % cls.loose_port

    @classmethod
    def tearDownClass(cls):
        cls.loose.shutdown()
        cls.loose.server_close()
        super(TestPlansAndDrives, cls).tearDownClass()

    def loose_get(self, path):
        with urllib.request.urlopen(self.loose_base + path) as response:
            return response.read(), response.status

    def loose_post(self, path, payload):
        request = urllib.request.Request(
            self.loose_base + path, data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"}, method="POST")
        with urllib.request.urlopen(request) as response:
            return json.loads(response.read().decode("utf-8")), response.status

    # ------------------------------------------------------------------ 盘符
    def test_drives_are_listed(self):
        body, status = self.loose_get("/api/drives")
        self.assertEqual(status, 200)
        drives = json.loads(body.decode("utf-8"))["drives"]
        self.assertTrue(all(drive.endswith(":\\") for drive in drives))

    def test_browsing_without_a_dir_shows_the_drives(self):
        body, _ = self.loose_get("/api/browse?dir=")
        data = json.loads(body.decode("utf-8"))
        self.assertEqual(data["dir"], u"")
        self.assertTrue(data["drives"], u"没给目录要给盘符（此电脑）")
        self.assertEqual(data["entries"], [])

    def test_browsing_outside_the_old_root_now_works(self):
        """默认不限目录：要能挑到别的盘/别的目录里的文件（用户的第 1 条意见）。"""
        body, status = self.loose_get("/api/browse?dir="
                                      + urllib.parse.quote(os.path.dirname(self.dir)))
        self.assertEqual(status, 200)
        data = json.loads(body.decode("utf-8"))
        self.assertTrue(data["dir"])
        self.assertTrue(data["drives"], u"任何一层都要带盘符栏")

    def test_a_missing_dir_falls_back_to_the_drives(self):
        body, status = self.loose_get("/api/browse?dir="
                                      + urllib.parse.quote(os.path.join(self.dir, u"没有这个目录")))
        self.assertEqual(status, 200)
        self.assertTrue(json.loads(body.decode("utf-8"))["drives"])

    # ------------------------------------------------------------------ 执行方案
    def test_builtin_plans_are_listed_first(self):
        body, _ = self.get("/api/plans")
        data = json.loads(body.decode("utf-8"))
        self.assertTrue(data["plans"], u"至少要有出厂自带的几套")
        self.assertTrue(data["plans"][0]["builtin"])
        self.assertTrue(all(plan["steps"] for plan in data["plans"]),
                        u"每套方案都要带步骤，不然「一眼看到有什么」是空的")

    def test_save_list_and_delete_a_plan(self):
        payload, status = self.loose_post("/api/plans/save",
                                          {"name": u"我的方案", "steps": ["captions", "tidy"]})
        self.assertEqual(status, 200)
        self.assertTrue(os.path.isfile(os.path.join(self.plans, u"我的方案.json")))

        body, _ = self.loose_get("/api/plans")
        names = [plan["name"] for plan in json.loads(body.decode("utf-8"))["plans"]]
        self.assertIn(u"我的方案", names)

        payload, _ = self.loose_post("/api/plans/delete", {"name": u"我的方案"})
        self.assertTrue(payload["ok"])
        self.assertFalse(os.path.exists(os.path.join(self.plans, u"我的方案.json")))

    def test_a_builtin_plan_cannot_be_deleted(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.loose_post("/api/plans/delete", {"name": gui_server.BUILTIN_PLANS[0]["name"]})
        self.assertIn(u"出厂", caught.exception.read().decode("utf-8"))

    def test_a_nasty_plan_name_is_refused(self):
        """方案名会变成文件名 —— 路径穿越必须挡住。"""
        for nasty in (u"../跑出去了", u"a/b", u"..", u"", u"x" * 41):
            with self.assertRaises(urllib.error.HTTPError):
                self.loose_post("/api/plans/save", {"name": nasty, "steps": ["captions"]})

    def test_saving_a_plan_without_steps_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError):
            self.loose_post("/api/plans/save", {"name": u"空方案", "steps": []})

    def test_saving_a_plan_with_an_unknown_step_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.loose_post("/api/plans/save", {"name": u"瞎写", "steps": ["nope"]})
        self.assertIn(u"captions", caught.exception.read().decode("utf-8"))


class TestExtdataEndpoint(GuiCase):
    """文档数据外置更新（gen / rebuild）走 GUI：与 CLI 同一条路。"""

    def test_gen_then_rebuild_end_to_end(self):
        # 造一份带高亮变量的文档
        path = os.path.join(self.dir, u"带高亮.docx")
        body = (fixtures.paragraph(fixtures.run(u"库容为 "), fixtures.run(u"3.5", highlight="yellow"),
                                   fixtures.run(u" 万m³。"),
                                   fixtures.run(u"水位 "), fixtures.run(u"12.8", highlight="yellow"),
                                   fixtures.run(u" m。"))
                + fixtures.paragraph(fixtures.run(u"表2.3-1     库容特性表")))
        fixtures.write_fixture(path, body=body)

        payload, status = self.post("/api/extdata", {"action": "gen", "path": path,
                                                     "dry_run": True})
        self.assertEqual(status, 200)
        self.assertEqual(payload["report"]["variables"], 2)
        self.assertIsNone(payload["download"], u"dry-run 不该给下载")

        payload, _ = self.post("/api/extdata", {"action": "gen", "path": path})
        self.assertTrue(payload["download"], u"要给出文档下载名")
        self.assertTrue(payload["downloads"], u"数据表也要能下载")
        xlsx = [name for name in payload["downloads"] if name.endswith(".xlsx")]
        self.assertEqual(len(xlsx), 1, u"要写出一份 xlsx 数据表")
        self.assertTrue(xlsx[0].endswith(u"外置数据.xlsx"),
                        u"数据表默认叫「XX外置数据.xlsx」")
        generated = payload["report"]["out"]
        self.assertTrue(os.path.exists(generated))

        body_bytes, status = self.get("/api/download?name="
                                      + urllib.parse.quote(xlsx[0]))
        self.assertEqual(status, 200)
        self.assertGreater(len(body_bytes), 1000)

        # 配方追加在**生成出来的文档**末尾 → 更新要对它跑，不是对原件跑
        payload, _ = self.post("/api/extdata", {"action": "rebuild", "path": generated,
                                                "xlsx": payload["report"]["xlsx"]})
        self.assertIn(u"3.5", payload["text"])
        self.assertTrue(payload["download"])
        self.assertTrue(os.path.exists(payload["report"]["out"]))

    def test_gen_refuses_to_overwrite_an_existing_data_table(self):
        path = os.path.join(self.dir, u"带高亮2.docx")
        body = fixtures.paragraph(fixtures.run(u"值 "), fixtures.run(u"1", highlight="yellow"))
        fixtures.write_fixture(path, body=body)
        taken = os.path.join(self.dir, u"带高亮2外置数据.xlsx")
        with io.open(taken, "w", encoding="utf-8") as handle:
            handle.write("占位")
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.post("/api/extdata", {"action": "gen", "path": path})
        self.assertIn(u"已经存在", caught.exception.read().decode("utf-8"))

    def test_an_unknown_action_is_refused(self):
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.post("/api/extdata", {"action": "nope", "path": self.path})
        self.assertIn(u"gen", caught.exception.read().decode("utf-8"))

    def test_rebuild_without_a_data_table_says_where_it_looked(self):
        path = os.path.join(self.dir, u"没数据表.docx")
        body = fixtures.paragraph(fixtures.run(u"值 "), fixtures.run(u"1", highlight="yellow"))
        fixtures.write_fixture(path, body=body)
        generated = self.post("/api/extdata",
                              {"action": "gen", "path": path})[0]["report"]["out"]
        # 把刚写出来的数据表挪走，模拟"找不到"
        xlsx = os.path.join(self.dir, u"没数据表外置数据.xlsx")
        if os.path.exists(xlsx):
            os.remove(xlsx)
        with self.assertRaises(urllib.error.HTTPError) as caught:
            self.post("/api/extdata", {"action": "rebuild", "path": generated})
        self.assertIn(u"找不到数据表", caught.exception.read().decode("utf-8"))


class TestPdfEndpoint(GuiCase):
    """GUI 的「导出 PDF」按钮真的接上了（不再是指向命令行的假按钮）。"""

    def test_pdf_export_through_the_gui(self):
        renderers = [item["name"] for item in
                     __import__("wordfactory.ops.pdf", fromlist=["pdf"]).detect_renderers()
                     if item["available"]]
        if not renderers:
            self.skipTest(u"本机没有 PDF 渲染器")
        payload, status = self.post("/api/pdf", {"path": self.path, "hidden": True})
        self.assertEqual(status, 200)
        self.assertTrue(payload["download"].endswith(u".pdf"))
        self.assertTrue(os.path.exists(payload["report"]["out"]))
        self.assertGreater(payload["report"]["bytes"], 1000)
        self.assertIn(u"渲染器", payload["text"])
        body, status = self.get("/api/download?name="
                                + urllib.parse.quote(payload["download"]))
        self.assertEqual(status, 200)
        self.assertTrue(body.startswith(b"%PDF-"), u"下载到的要是真 PDF")

    def test_a_broken_endpoint_never_surfaces_as_an_import_error(self):
        """接口里的相对导入写错会变成 500 + ModuleNotFoundError —— 用户看到的是天书。

        判据：出错可以，但不许是 "No module named" 这种实现细节。
        """
        try:
            self.post("/api/pdf", {"path": self.path, "hidden": True})
        except urllib.error.HTTPError as caught:
            message = caught.read().decode("utf-8")
            self.assertNotIn(u"No module named", message)
            self.assertTrue(u"渲染器" in message or u"PDF" in message, message)
        else:
            self.assertTrue(True)


class TestPortBusyCheck(unittest.TestCase):
    """双击启动器连点两下，不许悄悄起来两个服务（Windows 允许重复绑定同一端口）。"""

    def test_a_listening_port_is_reported_as_busy(self):
        import socket
        probe = socket.socket()
        probe.bind(("127.0.0.1", 0))
        probe.listen(1)
        port = probe.getsockname()[1]
        try:
            self.assertTrue(gui_server._port_busy("127.0.0.1", port))
            self.assertFalse(gui_server._port_busy("127.0.0.1", 0) or
                             gui_server._port_busy("127.0.0.1", 1),
                             u"没人听的端口不该报占用")
        finally:
            probe.close()

    def test_serve_refuses_to_start_twice_and_says_what_to_do(self):
        import socket
        probe = socket.socket()
        probe.bind(("127.0.0.1", 0))
        probe.listen(1)
        port = probe.getsockname()[1]
        try:
            with self.assertRaises(gui_server.PackageError) as caught:
                gui_server.serve(port=port, open_browser=False)
            message = u"%s" % caught.exception
            self.assertIn(u"已经起了一个", message)
            self.assertIn(u"--port", message, u"要告诉用户换端口")
        finally:
            probe.close()
