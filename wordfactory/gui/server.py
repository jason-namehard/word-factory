# -*- coding: utf-8 -*-
"""本地网页 GUI 的后端（**只用标准库**）。

设计要点：

* **只监听 127.0.0.1** —— 这是本机工具，不对外；默认端口 8765，可改。
* 前端是一个静态页（``web/index.html``）+ 几个 JSON 接口；**所有活都由
  :mod:`wordfactory.pipeline` 与各算子干**（GUI 和 CLI 同一条路，行为必然一致）。
* **下载只许下本次跑出来的文件**（内存里一份白名单）—— 不把整个文件系统暴露出去。
* 目录浏览**默认不限目录**（本机单用户工具，要能挑到其它盘的文件）；给了 ``--root``
  才关起来（测试与"只想让它在某个目录里挑"的场景）。
* **执行方案**（记录"勾了哪些功能、什么顺序"）存用户目录
  ``~/.wordfactory/plans/*.json``，整个目录拷走就能带到别的机器。

接口：
    GET  /                    页面
    GET  /api/steps           功能（宏）步骤表
    GET  /api/templates       表格模板
    GET  /api/tables?path=..  文档里的表格清单（按表头）
    GET  /api/browse?dir=..   列目录（选文件用；顺带给盘符列表）
    GET  /api/drives          本机盘符
    GET  /api/plans           执行方案清单（出厂的 + 存过的）
    POST /api/plans/save      存一个执行方案
    POST /api/plans/delete    删一个执行方案（出厂的删不掉）
    GET  /api/read?path=..    读一个文本文件（规则文件预览）
    POST /api/run             {path, steps, mode, out} → 报告
    POST /api/pdf             {path, out?, renderer?} → PDF 报告
    POST /api/extdata         {action: gen|rebuild, …} → 文档数据外置
    GET  /api/download?name=..下载本次产出
    POST /api/shutdown        关掉服务
"""

import io
import json
import os
import re
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

from .. import paths as paths_mod
from .. import pipeline as pipeline_mod
from .. import replace_rules as replace_rules_mod
from .. import tablestyle as tablestyle_mod
from ..ooxml import PackageError
from ..pipeline import PipelineError

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(HERE, "web", "index.html")

#: 读"第几块在第几页"最多探多少个块（COM 逐段读，块太多会慢）。
#: 实测那份 26 页报告 288 块，1 秒左右；1500 块够 100 页上下。
PROBE_BLOCKS_MAX = 1500

#: 本次进程跑出来的文件（下载白名单；重启即失效）
_DOWNLOADS = {}
_LOCK = threading.Lock()

#: 执行方案存哪（打包后跟着 exe 走；源码运行时在用户目录）
PLANS_DIR = paths_mod.plans_dir()

#: **临时文件夹**（用户 2026-09-27 拍板：放 wordfactory 项目文件夹内，就叫「临时文件」）。
#: 「运行此方案」跑出的临时版本（标蓝）都落这里，「清理临时文件」一键清空。
TEMP_DIR = paths_mod.temp_dir()

#: 出厂自带的执行方案（只读；用户改完"另存为"就成了自己的）
BUILTIN_PLANS = [
    {u"name": u"报告规范化（轻）", u"steps": [u"captions", u"tidy"],
     u"note": u"题注统一 + 一键整理：最常用、最不容易出错的一套"},
    {u"name": u"报告规范化（全）",
     u"steps": [u"captions", u"sup", u"tableclean", u"mdclean", u"tidy"],
     u"note": u"五个功能全上：上下标规则 → 表格清理 → MD 清理 → 整理"},
    {u"name": u"只清表格", u"steps": [u"tableclean"],
     u"note": u"只清表格单元格里的空格/回车（三档里的默认档）"},
    {u"name": u"只调题注", u"steps": [u"captions"],
     u"note": u"只统一表题/图题：编号、缩进、居中规则"},
]

#: 方案名只允许这些字符（挡住路径穿越：`..`、`/`、`\` 一律不进文件名）
_PLAN_NAME_OK = re.compile(r"^[\w一-龥·（）()\- ]{1,40}$", re.UNICODE)


def allow_download(path):
    with _LOCK:
        _DOWNLOADS[os.path.basename(path)] = os.path.abspath(path)
    return os.path.basename(path)


class Handler(BaseHTTPRequestHandler):
    server_version = "WordFactoryGUI/1.0"
    root = os.path.expanduser(u"~")
    plans_dir = PLANS_DIR
    #: 替换规则存哪（项目 rules/；测试里可以指去别处，免得测试数据落进真规则库）
    rules_base_dir = None

    # ------------------------------------------------------------- 基础
    def log_message(self, fmt, *args):        # 控制台别刷屏
        pass

    def _send(self, code, body, content_type="text/html; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
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
                    page = handle.read()
                stamp = str(int(os.path.getmtime(PAGE)))
                page = page.replace(b"</head>",
                                    (b'<meta name="wf-build" content="' + stamp.encode()
                                     + b'"></head>'))
                self._send(200, page)
            elif path == "/api/steps":
                self._json({"ok": True,
                            "labels": dict(pipeline_mod.STEP_LABELS),
                            "steps": [
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
            elif path == "/api/drives":
                self._json({"ok": True, "drives": _drives()})
            elif path == "/api/plans":
                self._plans()
            elif path == "/api/temp-dir":
                self._temp_dir()
            elif path == "/api/template-previews":
                self._template_previews()
            elif path == "/api/replace-rules":
                self._json({"ok": True,
                            "rules": replace_rules_mod.list_rules(self.rules_base())})
            elif path == "/api/build":
                self._json({"ok": True, "build": str(int(os.path.getmtime(PAGE)))})
            elif path == "/api/subs-rules":
                self._json({"ok": True, "data": _read_json(_rules_path(u"subscripts.json"))})
            elif path == "/api/frontmatter":
                self._frontmatter(_query(query).get("path", u""),
                                  _query(query).get("accurate") == "1")
            elif path == "/api/doc-formats":
                self._doc_formats(_query(query).get("path", u""))
            elif path == "/api/read":
                self._read(_query(query).get("path", u""))
            elif path == "/api/plugin-guide":
                base = os.path.dirname(os.path.dirname(HERE))   # 项目根
                guide = os.path.join(base, "docs", "PLUGIN-GUIDE.md")
                with io.open(guide, "r", encoding="utf-8") as handle:
                    self._send(200, handle.read(), "text/plain; charset=utf-8")
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
            elif path == "/api/pdf":
                self._pdf()
            elif path == "/api/extdata":
                self._extdata()
            elif path == "/api/plans/save":
                self._plan_save()
            elif path == "/api/plans/delete":
                self._plan_delete()
            elif path == "/api/temp/clean":
                self._temp_clean()
            elif path == "/api/replace-rules/save":
                self._rule_save()
            elif path == "/api/replace-rules/delete":
                self._rule_delete()
            elif path == "/api/subs-rules/save":
                self._subs_save()
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
        """列目录（选文件用）。

        ``dir`` 空或不存在时返回**盘符列表**（"此电脑"）—— 用户要能挑到其它盘的文件。
        ``root`` 为 None 表示不限目录；给了才关起来（测试用）。
        """
        directory = os.path.abspath(directory) if directory else u""
        if not directory or not os.path.isdir(directory):
            self._json({"ok": True, "dir": u"", "parent": u"",
                        "drives": _drives(), "entries": []})
            return
        if self.root and not _inside(directory, self.root):
            raise PipelineError(u"只能浏览 %s 以内的目录" % self.root)
        entries = []
        for name in sorted(os.listdir(directory), key=lambda n: (not os.path.isdir(
                os.path.join(directory, n)), n.lower())):
            full = os.path.join(directory, name)
            entries.append({"name": name, "dir": os.path.isdir(full),
                            "docx": name.lower().endswith((".docx", ".docm")),
                            "size": os.path.getsize(full) if os.path.isfile(full) else 0})
        parent = os.path.dirname(directory)
        self._json({"ok": True, "dir": directory,
                    "parent": parent if parent != directory else u"",
                    "drives": _drives(), "entries": entries[:500]})

    def _plans(self):
        """执行方案清单：出厂的在前，用户存过的在后（一眼能看到有什么）。"""
        saved = []
        directory = self.plans_dir
        if os.path.isdir(directory):
            for name in sorted(os.listdir(directory)):
                if not name.endswith(".json"):
                    continue
                try:
                    with io.open(os.path.join(directory, name), "r",
                                 encoding="utf-8") as handle:
                        data = json.load(handle)
                    saved.append({u"name": data.get("name") or name[:-5],
                                  u"steps": data.get("steps") or [],
                                  u"note": data.get("note") or u"",
                                  u"builtin": False})
                except (ValueError, OSError):
                    continue                     # 读不动的文件不往清单里露底
        self._json({"ok": True, "dir": directory,
                    "plans": [dict(plan, builtin=True) for plan in BUILTIN_PLANS] + saved})

    def _plan_save(self):
        data = self._body_json()
        name = (data.get("name") or u"").strip()
        steps = data.get("steps") or []
        if not _PLAN_NAME_OK.match(name):
            raise PipelineError(
                u"方案名不太好：%r。1–40 个字，别带 . / \\ : 这类字符" % name)
        if not steps:
            raise PipelineError(u"一个功能都没勾，存它干嘛")
        known = set(pipeline_mod.STEPS)
        # 步骤允许两种写法：字符串（用默认参数）或 {"op": …, "params": …, "label": …}
        # —— 表格模板/替换规则带参数，必须原样存进方案，载入才能还原（实测踩过：
        #    以前只认字符串，带参数的步骤一存就报"没有这些功能"）。
        normalized = []
        for step in steps:
            op = step.get("op") if isinstance(step, dict) else step
            if op not in known:
                raise PipelineError(u"没有这个功能：%s（可用的：%s）"
                                    % (op, u"、".join(known)))
            normalized.append(step if isinstance(step, dict) else {u"op": step})
        directory = self.plans_dir
        if not os.path.isdir(directory):
            os.makedirs(directory)
        path = os.path.join(directory, name + u".json")
        if not _inside(path, directory):        # 双保险：名字再怎么花样也跑不出去
            raise PipelineError(u"方案名不太好：%s" % name)
        payload = {"name": name, "steps": normalized,
                   "note": (data.get("note") or u"").strip()}
        with io.open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        log_note = u"已存执行方案：%s（%d 步）→ %s" % (name, len(normalized), path)
        self._json({"ok": True, "name": name, "path": path, "text": log_note})

    def _plan_delete(self):
        data = self._body_json()
        name = (data.get("name") or u"").strip()
        if any(plan["name"] == name for plan in BUILTIN_PLANS):
            raise PipelineError(u"「%s」是出厂自带的，删不掉；你可以改完另存为一个" % name)
        path = os.path.join(self.plans_dir, name + u".json")
        if not os.path.isfile(path):
            raise PipelineError(u"没有这个执行方案：%s" % name)
        os.remove(path)
        self._json({"ok": True, "name": name, "text": u"已删除执行方案：%s" % name})

    # ------------------------------------------------------------- 临时文件夹
    def _temp_dir(self):
        if not os.path.isdir(TEMP_DIR):
            os.makedirs(TEMP_DIR)
        files = []
        for name in sorted(os.listdir(TEMP_DIR)):
            full = os.path.join(TEMP_DIR, name)
            if os.path.isfile(full):
                files.append({"name": name, "size": os.path.getsize(full)})
        self._json({"ok": True, "dir": TEMP_DIR, "files": files})

    def _temp_clean(self):
        """清空临时文件夹（只清这个目录里的**文件**，目录本身保留；防误删别处）。"""
        if not os.path.isdir(TEMP_DIR):
            self._json({"ok": True, "removed": 0, "text": u"临时文件夹还没有创建，无需清理"})
            return
        removed = 0
        for name in os.listdir(TEMP_DIR):
            full = os.path.join(TEMP_DIR, name)
            if os.path.isfile(full) and _inside(full, TEMP_DIR):
                os.remove(full)
                removed += 1
        self._json({"ok": True, "removed": removed,
                    "text": u"已清空临时文件夹（%d 个文件）" % removed})

    # ------------------------------------------------------------- 表格模板预览
    def _template_previews(self):
        styles = tablestyle_mod.StyleSet.load(_rules_path(u"tablestyle.json"))
        self._json({"ok": True, "styles": [
            {"name": style.name, "note": style.note,
             "svg": tablestyle_mod.preview_svg(style)}
            for style in styles.styles.values()]})

    # ------------------------------------------------------------- 替换规则
    def _rule_save(self):
        data = self._body_json()
        name = (data.get("name") or u"").strip()
        rules = data.get("rules") or {}
        path = replace_rules_mod.save(self.rules_base(), name, rules)
        self._json({"ok": True, "name": name, "path": path,
                    "text": u"已存替换规则：%s → %s" % (name, path)})

    def _rule_delete(self):
        data = self._body_json()
        name = (data.get("name") or u"").strip()
        replace_rules_mod.delete(self.rules_base(), name)
        self._json({"ok": True, "name": name, "text": u"已删除替换规则：%s" % name})

    def rules_base(self):
        return self.rules_base_dir or paths_mod.rules_dir()

    def _subs_save(self):
        """上下标规则（字典）存盘 —— 界面「上下标规则」页用（用户 2026-09-28 要的管理区）。"""
        data = self._body_json()
        payload = data.get("data")
        if not isinstance(payload, dict) or not isinstance(payload.get("rules"), list):
            raise PipelineError(u"规则内容不对：要是一个带 rules 数组的对象")
        path = _rules_path(u"subscripts.json")
        with io.open(path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        count = len(payload["rules"])
        self._json({"ok": True, "count": count, "path": path,
                    "text": u"已保存上下标规则（%d 条）→ %s" % (count, path)})

    # ------------------------------------------------------------- 前置区
    def _frontmatter(self, path, accurate=False):
        """前置区识别。``accurate=True`` 时用 **Word/WPS 真实页码**（最准，但要起排版引擎）。

        顺带把**整篇的"第几块在第几页"**（``block_pages``）一起回给界面：
        一键整理要拿它认"自然溢出"造成的空白页（实测：那份报告第 1 页 22 个空段、
        第 2 页才是封面，两页之间没有分页符；不按真实页码，第 1 页的空白页永远删不掉）。
        """
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        from .. import frontmatter
        from ..document import Document
        block_pages = None
        probe_info = None
        with Document(path) as doc:
            if accurate:
                from .. import pageprobe
                blocks = list(doc.body())
                limit = min(len(blocks), PROBE_BLOCKS_MAX)
                try:
                    info_probe = pageprobe.probe(
                        path, blocks=limit + 5,
                        block_starts=pageprobe.block_start_paragraphs(doc, limit=limit + 5))
                    block_pages = info_probe["block_pages"]
                    probe_info = {u"renderer": info_probe.get("renderer"),
                                  u"pages": info_probe.get("pages")}
                except pageprobe.PageProbeError as exc:
                    raise PipelineError(u"读真实页码失败：%s（可以先用估算口径）" % exc)
            info = frontmatter.detect(doc, block_pages=block_pages)
        self._json({"ok": True, "info": info,
                    "text": frontmatter.format_report(info),
                    "accurate": bool(block_pages),
                    "block_pages": block_pages,
                    "renderer": probe_info or None})

    # ------------------------------------------------------------- 文档格式清单
    def _doc_formats(self, path):
        """读出报告里**实际用过**的格式：给「替换规则」的下拉框做数据源。"""
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        from ..ooxml import qn
        from ..document import Document
        east_asia, latin, sizes, aligns, colors = set(), set(), set(), set(), set()
        with Document(path) as doc:
            for run in doc.part().iter(qn("w:r")):
                pr = run.find(qn("w:rPr"))
                if pr is None:
                    continue
                rfonts = pr.find(qn("w:rFonts"))
                if rfonts is not None:
                    for attr, bucket in (("eastAsia", east_asia), ("ascii", latin)):
                        value = rfonts.get(qn("w:%s" % attr))
                        if value:
                            bucket.add(value)
                for tag in ("w:sz", "w:szCs"):
                    node = pr.find(qn(tag))
                    if node is not None and node.get(qn("w:val")):
                        sizes.add(node.get(qn("w:val")))
                color = pr.find(qn("w:color"))
                if color is not None and color.get(qn("w:val")):
                    colors.add((color.get(qn("w:val")) or u"").upper())
            for pr in doc.part().iter(qn("w:pPr")):
                jc = pr.find(qn("w:jc"))
                if jc is not None and jc.get(qn("w:val")):
                    aligns.add(jc.get(qn("w:val")))
        self._json({"ok": True,
                    "east_asia": sorted(east_asia), "latin": sorted(latin),
                    "sizes": sorted(sizes, key=lambda v: float(v) if v.replace(".", u"").isdigit() else 0),
                    "aligns": sorted(aligns), "colors": sorted(colors)})

    def _pdf(self):
        """导出 PDF：编排本机装着的渲染器（Word/WPS/LibreOffice）。"""
        from ..ops import pdf as pdf_op
        data = self._body_json()
        path = data.get("path") or u""
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        out = data.get("out") or os.path.splitext(path)[0] + u".pdf"
        if os.path.abspath(out) == os.path.abspath(path):
            raise PipelineError(u"输出不能跟输入同一个文件")
        report = pdf_op.export(path, out, prefer=data.get("renderer") or None,
                               timeout=int(data.get("timeout") or 300),
                               visible=not data.get("hidden"))
        download = allow_download(report["out"])
        text = (u"文件：%s\n渲染器：%s（%s）\n已写出：%s（%d 字节，%s 页，耗时 %.1f 秒）"
                % (path, report.get("kind"), report.get("renderer"), report["out"],
                   report.get("bytes") or 0, report.get("pages") or u"?",
                   report.get("seconds") or 0))
        self._json({"ok": True, "report": report, "text": text, "download": download})

    def _extdata(self):
        """文档数据外置更新：``gen`` 生成外置数据表 + 配方，``rebuild`` 按数据表重建，
        ``update`` 把高亮数值**就地**换成数据表新值（文档数据更新 / 表格数据更新）。"""
        from ..document import Document
        from ..ops import recipe as recipe_op
        data = self._body_json()
        action = data.get("action") or u""
        path = data.get("path") or u""
        if action not in ("gen", "rebuild", "update"):
            raise PipelineError(u"动作只有 gen（生成外置数据）/ rebuild（按数据表重建段落）/"
                                u" update（就地更新数值）")
        if not path or not os.path.exists(path):
            raise PipelineError(u"文件不存在：%s" % path)
        stem = os.path.splitext(os.path.basename(path))[0]

        if action == "update":
            scope = data.get("scope") or u"body"
            if scope not in ("body", "tables"):
                raise PipelineError(u"scope 只能是 body（正文）/ tables（表格单元格）")
            xlsx = data.get("xlsx")
            if not xlsx or not os.path.exists(xlsx):
                raise PipelineError(u"先选数据表（.xlsx）再更新")
            from ..xlsx import read_column_b
            values = read_column_b(xlsx, max_rows=500)
            out = data.get("out") or os.path.join(
                os.path.dirname(os.path.abspath(path)),
                stem + (u"_表格更新.docx" if scope == "tables" else u"_数据更新.docx"))
            if not data.get("dry_run") and os.path.abspath(out) == os.path.abspath(path):
                raise PipelineError(u"输出不能跟输入同一个文件")
            with Document(path) as doc:
                report = recipe_op.update_values(doc, values, scope=scope,
                                                 dry_run=bool(data.get("dry_run")))
                written = None
                if not data.get("dry_run"):
                    written = doc.save(os.path.abspath(out))
            label = u"表格数据更新" if scope == "tables" else u"文档数据更新"
            lines = [u"%s：%s" % (label, path),
                     u"数据表：%s ｜ 取到 %d 个值" % (xlsx, len(values)),
                     u"就地替换 %d 处%s" % (report["replaced"],
                                           u"（还有 %d 处高亮没给到值）" % report["missing"]
                                           if report["missing"] else u""),
                     u"（只换字，不动段落结构；没有高亮的地方不碰）"]
            if data.get("dry_run"):
                lines.insert(0, u"--dry-run：一个字节都没写")
            else:
                lines.insert(0, u"已写出：%s" % written)
            self._json({"ok": True, "report": dict(report, out=written, xlsx=xlsx),
                        "text": u"\n".join(lines),
                        "download": allow_download(written) if written else None})
            return

        if action == "gen":
            mode = data.get("mode") or u"highlight"
            if mode == u"chars" and not data.get("char"):
                raise PipelineError(u"占位符模式要给 --char（占位符字符串，如 xx）")
            excel_file = data.get("excel_file") or (stem + u"外置数据.xlsx")
            out_doc = data.get("out") or os.path.join(
                os.path.dirname(os.path.abspath(path)), stem + u"_外置数据.docx")
            out_xlsx = data.get("xlsx") or os.path.join(
                os.path.dirname(os.path.abspath(out_doc)), excel_file)
            options = {"mode": mode, "char": data.get("char"),
                       "name": data.get("name") or (stem + u"外置数据"),
                       "excel_file": excel_file,
                       "sheet_name": data.get("sheet") or u"Sheet1",
                       "out_xlsx": os.path.abspath(out_xlsx),
                       "append": data.get("append", True),
                       "trim_last_char": bool(data.get("trim_last_char"))}
            if not data.get("dry_run") and os.path.exists(out_xlsx):
                raise PipelineError(
                    u"%s 已经存在（不覆盖已有数据表）；换个位置再存" % out_xlsx)
            with Document(path) as doc:
                report, recipe, xlsx_path = recipe_op.generate(
                    doc, options, dry_run=bool(data.get("dry_run")))
                written = None
                if not data.get("dry_run"):
                    written = doc.save(os.path.abspath(out_doc))
            lines = [u"识别模式：%s" % (u"高亮（所有高亮片段）" if mode == u"highlight"
                                       else u"占位符 %r" % data.get("char")),
                     u"配方：%s ｜ 变量 %d 个" % (recipe.name, report["variables"]),
                     u"数据表：%s ｜ 工作表 %s" % (excel_file, recipe.sheet_name)]
            for index, row in enumerate(report["rows"], start=1):
                prefix, value = row[0], row[1]
                context = row[2] if len(row) > 2 else u""
                lines.append(u"    %2d. A=%s ｜ B=%s ｜ 段前=%s"
                             % (index, prefix[:16] or u"（空）", value[:24],
                                (context[:16] + u"…") if len(context) > 16
                                else (context or u"（无）")))
            if data.get("dry_run"):
                lines.insert(0, u"--dry-run：一个字节都没写")
            else:
                lines.insert(0, u"已写出：%s" % written)
                lines.append(u"已写出数据表：%s" % xlsx_path)
            payload = dict(report, out=written, xlsx=xlsx_path)
            downloads = [allow_download(item) for item in (written, xlsx_path) if item]
            self._json({"ok": True, "report": payload, "text": u"\n".join(lines),
                        "download": downloads[0] if downloads else None,
                        "downloads": downloads})
            return

        # rebuild：按数据表把段落重建出来
        out = data.get("out") or os.path.join(
            os.path.dirname(os.path.abspath(path)), stem + u"_已更新.docx")
        if not data.get("dry_run") and os.path.abspath(out) == os.path.abspath(path):
            raise PipelineError(u"输出不能跟输入同一个文件")
        with Document(path) as doc:
            recipe = (recipe_op.read_recipe(data["recipe_file"]) if data.get("recipe_file")
                      else recipe_op.read_recipe(path))
            xlsx = data.get("xlsx") or recipe_op.resolve_xlsx(
                recipe, path, data.get("data_dir"))
            values = recipe_op.values_for(recipe, xlsx)
            report = recipe_op.rebuild(doc, recipe, values,
                                       dry_run=bool(data.get("dry_run")))
            written = None
            if not data.get("dry_run"):
                written = doc.save(os.path.abspath(out))
        lines = [u"配方：%s（%d 个变量）" % (recipe.name, recipe.variable_count),
                 u"数据表：%s ｜ 工作表 %s ｜ 取到 %d 个值"
                 % (xlsx, recipe.sheet_name, len(values)),
                 u"重建：%d 个段落%s" % (report["paragraphs"],
                                        u"（其中 %d 处 #数据缺失#）" % report["missing"]
                                        if report["missing"] else u""),
                 u"",
                 u"重建出来的段落正文："]
        for text in report["text"].split(u"\n"):
            lines.append(u"    %s" % text)
        if data.get("dry_run"):
            lines.insert(0, u"--dry-run：一个字节都没写")
        else:
            lines.insert(0, u"已写出：%s" % written)
        self._json({"ok": True, "report": dict(report, out=written, xlsx=xlsx),
                    "text": u"\n".join(lines),
                    "download": allow_download(written) if written else None})

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
            body = handle.read()
        lower = name.lower()
        if lower.endswith(u".pdf"):
            content_type = u"application/pdf"
        elif lower.endswith(u".xlsx"):
            content_type = (u"application/vnd.openxmlformats-officedocument"
                            u".spreadsheetml.sheet")
        else:
            content_type = (u"application/vnd.openxmlformats-officedocument"
                            u".wordprocessingml.document")
        # Content-Disposition 必须**在正文之前**发（原来写在 _send 之后，
        # 头已经出去了 → 浏览器拿不到附件名，只能靠 <a download> 兜）
        self._send(200, body,
                   "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                   extra={"Content-Disposition":
                          u"attachment; filename*=UTF-8''%s" % _quote(name)})

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
        out = data.get("out") or None
        # **重名先问一句**（用户 2026-09-30："告知用户是替换还是另建一个"）：
        # 目标文件已存在时，不自作主张 —— 回一个 conflict 让界面弹确认。
        mode = data.get("mode") or "verify"
        if not out and not data.get("dry_run"):
            # 空字符串要当"没给"处理：否则后缀为空 → 路径退回原文件，永远冲突
            candidate = pipeline_mod.default_out(
                path, mode, suffix=data.get("suffix") or None)
            if os.path.exists(candidate):
                if not data.get("allow_overwrite"):
                    self._json({"ok": True, "conflict": {
                        "path": candidate,
                        "name": os.path.basename(candidate),
                        "dir": os.path.dirname(candidate),
                        "mode": mode}})
                    return
            # 算好的名字直接用（别再让引擎"已存在就加序号"那套兜底插手 ——
            # 覆盖就是覆盖、改后缀就是改后缀）
            out = candidate
        report = pipeline_mod.run_pipeline(
            path, data.get("steps") or [], mode=data.get("mode") or "verify",
            out_path=out or None, dry_run=bool(data.get("dry_run")),
            block_pages=_block_pages_for_run(path, data))
        download = None
        if report.get("out"):
            download = allow_download(report["out"])
        self._json({"ok": True, "report": _slim(report),
                    "text": pipeline_mod.format_report(report),
                    "download": download})


def _block_pages_for_run(path, data):
    """跑之前把**真实页码**（第几块在第几页）备好，交给一键整理判空白页。

    优先用界面带来的（选文件校准过的那份，省一次排版）；**没带就服务端自己探** ——
    2026-10-08 实测踩到：前端页缓存/校准失败时运行请求里没有 ``block_pages``，
    结果又退回"按分页符估算"，**纯空白首页一个都删不掉**（用户报的就是这个现象）。
    "一键整理"要判空白页就必须拿到真实页码，这条不能指望界面状态。

    没装 Word/WPS 或探测失败 → 返回 None，退回估算口径（不阻断运行）。
    """
    steps = [step.get("op") if isinstance(step, dict) else step
             for step in (data.get("steps") or [])]
    if u"tidy" not in steps:
        return data.get("block_pages") or None
    from ..document import Document
    from .. import pipeline as _pipeline
    with Document(path) as doc:
        pages = _pipeline._probe_block_pages(doc)
    return pages or data.get("block_pages") or None


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


def _read_json(path, default=None):
    if not os.path.exists(path):
        return default if default is not None else {}
    with io.open(path, "r", encoding="utf-8-sig") as handle:
        return json.load(handle)


def _inside(path, root):
    path = os.path.abspath(path)
    root = os.path.abspath(root)
    return path == root or path.startswith(root + os.sep)


def _drives():
    """本机盘符（选文件时要能跳到别的盘）。非 Windows 返回空列表。"""
    import string
    out = []
    for letter in string.ascii_uppercase:
        root = letter + u":\\"
        if os.path.isdir(root):
            out.append(root)
    return out


def _rules_path(name):
    return paths_mod.rules_path(name)


def _port_busy(host, port):
    """端口上是不是已经有东西在听了（我们自己的另一个实例最可能）。

    为什么必须查：Windows 的 ``SO_REUSEADDR`` 允许**两个** socket 绑同一个端口，
    于是双击启动器连点两下会起来两个服务，请求在两个进程之间乱跳 ——
    表现是"时好时坏"，极难排查。所以绑之前先连一下，有人就明说。
    """
    import socket
    probe = socket.socket()
    probe.settimeout(0.6)
    try:
        return probe.connect_ex((host or "127.0.0.1", port)) == 0
    finally:
        probe.close()


def serve(host="127.0.0.1", port=8765, open_browser=True, root=None, plans=None):
    """起服务；``open_browser`` 时顺手打开浏览器。

    ``root=None`` 表示**不限目录**（要能挑到其它盘的文件）；给了才关起来。
    ``plans=None`` 用默认的用户目录，可指到别处（测试用）。
    """
    if not os.path.exists(PAGE):
        raise PackageError(u"找不到页面文件：%s" % PAGE)
    if _port_busy(host, port):
        raise PackageError(
            u"端口 %d 已经有人在听了 —— 多半是 word 工厂已经起了一个。\n"
            u"  先试试打开 http://%s:%d/ ，能用就直接用；\n"
            u"  打不开就关掉之前那个黑色命令行窗口，或者换个端口：\n"
            u"  python -m wordfactory.cli gui --port 8766" % (port, host, port))
    handler = type("BoundHandler", (Handler,),
                   {"root": os.path.abspath(root) if root else None,
                    "plans_dir": os.path.abspath(plans) if plans else PLANS_DIR})
    try:
        server = ThreadingHTTPServer((host, port), handler)
    except OSError as exc:
        # 万一预检没拦住（端口在两步之间被抢），也别甩 traceback
        raise PackageError(
            u"端口 %d 绑不上（%s）。换个端口试试：python -m wordfactory.cli gui --port 8766"
            % (port, exc))
    url = "http://%s:%d/" % (host, port)
    print(u"word 工厂 GUI 已启动：%s" % url)
    print(u"（只监听本机 %s；关掉这个窗口或点页面上的「退出」即停）" % host)
    print(u"执行方案存在：%s" % handler.plans_dir)
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
