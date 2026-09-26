# -*- coding: utf-8 -*-
"""本地网页 GUI 的后端（**只用标准库**）。

设计要点（用户 2026-09-23 定的那版布局照做）：

* **只监听 127.0.0.1** —— 这是本机工具，不对外；默认端口 8765，可改。
* 前端是一个静态页（``web/index.html``）+ 几个 JSON 接口；**所有活都由
  :mod:`wordfactory.pipeline` 干**（GUI 和 CLI 同一条路，行为必然一致）。
* **下载只许下本次跑出来的文件**（内存里一份白名单）—— 不把整个文件系统暴露出去。
* 目录浏览限定在一个根目录里（默认用户主目录），也只读。

接口：
    GET  /                    页面
    GET  /api/steps           配方步骤表
    GET  /api/templates       表格模板
    GET  /api/tables?path=..  文档里的表格清单（按表头）
    GET  /api/browse?dir=..   列目录（选文件用）
    GET  /api/read?path=..    读一个文本文件（规则文件预览）
    POST /api/run             {path, steps, mode, out} → 报告
    GET  /api/download?name=..下载本次产出
    POST /api/shutdown        关掉服务
"""

import io
import json
import os
import posixpath
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from .. import pipeline as pipeline_mod
from .. import tablestyle as tablestyle_mod
from ..ooxml import PackageError
from ..pipeline import PipelineError

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(HERE, "web", "index.html")

#: 本次进程跑出来的文件（下载白名单；重启即失效）
_DOWNLOADS = {}
_LOCK = threading.Lock()


def allow_download(path):
    with _LOCK:
        _DOWNLOADS[os.path.basename(path)] = os.path.abspath(path)
    return os.path.basename(path)


class Handler(BaseHTTPRequestHandler):
    server_version = "WordFactoryGUI/1.0"
    root = os.path.expanduser(u"~")

    # ------------------------------------------------------------- 基础
    def log_message(self, fmt, *args):        # 控制台别刷屏
        pass

    def _send(self, code, body, content_type="text/html; charset=utf-8"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, payload, code=200):
        self._send(code, json.dumps(payload, ensure_ascii=False), "application/json")

    def _error(self, code, message):
        self._json({"ok": False, "error": message}, code)

    # ------------------------------------------------------------- GET
    def do_GET(self):
        path, _, query = self.path.partition("?")
        try:
            if path in ("/", "/index.html"):
                with io.open(PAGE, "rb") as handle:
                    self._send(200, handle.read())
            elif path == "/api/steps":
                self._json({"ok": True, "steps": [
                    {"op": name, "note": note, "params": params}
                    for name, (note, params) in pipeline_mod.STEPS.items()]})
            elif path == "/api/templates":
                styles = tablestyle_mod.StyleSet.load(
                    _rules_path(u"tablestyle.json"))
                self._json({"ok": True, "styles": [
                    {"name": style.name, "note": style.note} for style in styles.styles.values()]})
            elif path == "/api/tables":
                self._tables(_query(query).get("path", u""))
            elif path == "/api/browse":
                self._browse(_query(query).get("dir", u""))
            elif path == "/api/read":
                self._read(_query(query).get("path", u""))
            elif path == "/api/download":
                self._download(_query(query).get("name", u""))
            else:
                self._error(404, u"没有这个接口：%s" % path)
        except Exception as exc:                        # noqa: BLE001 - 前端要看到人话
            self._error(500, u"%s: %s" % (type(exc).__name__, exc))

    # ------------------------------------------------------------- POST
    def do_POST(self):
        path = self.path.partition("?")[0]
        try:
            if path == "/api/run":
                self._run()
            elif path == "/api/audit":
                self._audit()
            elif path == "/api/shutdown":
                self._json({"ok": True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
            else:
                self._error(404, u"没有这个接口：%s" % path)
        except Exception as exc:                        # noqa: BLE001 - 前端要看到人话
            self._error(500, u"%s: %s" % (type(exc).__name__, exc))

    # ------------------------------------------------------------- 具体接口
    def _body_json(self):
        length = int(self.headers.get("Content-Length") or 0)
        raw = self.rfile.read(length) if length else b"{}"
        try:
            return json.loads(raw.decode("utf-8"))
        except ValueError:
            raise PipelineError(u"请求体不是 JSON")

    def _tables(self, path):
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        from ..document import Document
        with Document(path) as doc:
            tables = tablestyle_mod.table_summaries(doc)
        self._json({"ok": True, "tables": tables})

    def _browse(self, directory):
        directory = directory or self.root
        directory = os.path.abspath(directory)
        if not _inside(directory, self.root):
            raise PipelineError(u"只能浏览 %s 以内的目录" % self.root)
        if not os.path.isdir(directory):
            raise PipelineError(u"不是目录：%s" % directory)
        entries = []
        for name in sorted(os.listdir(directory), key=lambda n: (not os.path.isdir(
                os.path.join(directory, n)), n.lower())):
            full = os.path.join(directory, name)
            entries.append({"name": name, "dir": os.path.isdir(full),
                            "docx": name.lower().endswith((".docx", ".docm")),
                            "size": os.path.getsize(full) if os.path.isfile(full) else 0})
        self._json({"ok": True, "dir": directory, "parent": os.path.dirname(directory),
                    "entries": entries[:500]})

    def _read(self, path):
        if not path or not os.path.isfile(path):
            raise PipelineError(u"文件不存在：%s" % path)
        if os.path.getsize(path) > 512 * 1024:
            raise PipelineError(u"文件太大，不读了")
        with io.open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
            self._send(200, handle.read(), "text/plain; charset=utf-8")

    def _download(self, name):
        with _LOCK:
            path = _DOWNLOADS.get(name)
        if not path or not os.path.exists(path):
            raise PipelineError(u"没有这个文件（可能服务重启过，重新跑一次）")
        with io.open(path, "rb") as handle:
            self._send(200, handle.read(),
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        self.send_header("Content-Disposition",
                         u'attachment; filename="%s"' % _quote(name))

    def _audit(self):
        """体检：重新打开文件按继承链算一遍（末行 AUDIT=PASS/FAIL）。"""
        from ..audit import audit as audit_op
        from ..audit import format_audit
        from ..fonts import FontRuleSet
        data = self._body_json()
        path = data.get("path") or u""
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        rule_set = FontRuleSet.load(_rules_path(u"fonts.json"))
        report = audit_op(path, rule_set)
        self._json({"ok": True, "verdict": report["verdict"],
                    "reasons": report["reasons"], "text": format_audit(report)})

    def _run(self):
        data = self._body_json()
        path = data.get("path") or u""
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        report = pipeline_mod.run_pipeline(
            path, data.get("steps") or [], mode=data.get("mode") or "verify",
            out_path=data.get("out") or None, dry_run=bool(data.get("dry_run")))
        download = None
        if report.get("out"):
            download = allow_download(report["out"])
        self._json({"ok": True, "report": _slim(report),
                    "text": pipeline_mod.format_report(report),
                    "download": download})


def _slim(report):
    """报告里有些字段不必/不能给前端（比如整棵 XML 树、成片明细），这里裁一下。"""
    slim = dict(report)
    slim.pop("details", None)
    if slim.get("audit"):
        audit = dict(slim["audit"])
        audit.pop("bad_effective", None)
        slim["audit"] = audit
    original_steps = report.get("steps") or []
    slim["steps"] = [dict((k, v) for k, v in step.items() if k != "report")
                     for step in original_steps]
    for step, original in zip(slim["steps"], original_steps):
        step["report"] = _slim(original.get("report") or {})
    return slim


def _query(query):
    out = {}
    for chunk in query.split(u"&"):
        if not chunk:
            continue
        key, _, value = chunk.partition(u"=")
        out[_unquote(key)] = _unquote(value)
    return out


def _unquote(text):
    import urllib.parse
    return urllib.parse.unquote_plus(text or u"")


def _quote(text):
    import urllib.parse
    return urllib.parse.quote(text or u"")


def _inside(path, root):
    path = os.path.abspath(path)
    root = os.path.abspath(root)
    return path == root or path.startswith(root + os.sep)


def _rules_path(name):
    base = os.path.dirname(os.path.dirname(HERE))
    return os.path.join(base, "rules", name)


def serve(host="127.0.0.1", port=8765, open_browser=True, root=None):
    """起服务；``open_browser`` 时顺手打开浏览器。"""
    if not os.path.exists(PAGE):
        raise PackageError(u"找不到页面文件：%s" % PAGE)
    handler = type("BoundHandler", (Handler,), {"root": os.path.abspath(root)
                                                if root else os.path.expanduser(u"~")})
    server = ThreadingHTTPServer((host, port), handler)
    url = "http://%s:%d/" % (host, port)
    print(u"word 工厂 GUI 已启动：%s" % url)
    print(u"（只监听本机 %s；关掉这个窗口或点页面上的「退出」即停）" % host)
    if open_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print(u"\n已停止。")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    serve()
