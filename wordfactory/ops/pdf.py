# -*- coding: utf-8 -*-
"""PDF 导出：**编排**外部渲染器，自己不做渲染。

用户的原始需求里就有"还能支持导出为 PDF 格式"，而从项目一开始就定清的边界是：
**`.docx → .pdf` 必须有排版引擎**（Word / WPS / LibreOffice），XML 层算不出分页。
所以本模块只干三件事：**找渲染器 → 拼命令 → 收结果**，并把"用哪个、花了多久、出了多大"
如实报出来。

三条纪律（沿用上一轮跑 Word COM 的教训）：

* **只用我们自己起的那个实例**，绝不 `taskkill /F /IM <名字>`（那会打死用户正在编辑的 Word）；
* 文档**只读打开**，导完原样关掉，**不保存**；
* 有超时；超时就明说"可能有个 Word 窗口还开着"，不假装成功。
"""

import io
import os
import re
import shutil
import subprocess
import threading
import time

#: Word 的导出格式常量
WD_EXPORT_FORMAT_PDF = 17
WD_STATISTIC_PAGES = 2
#: 默认超时（秒）
DEFAULT_TIMEOUT = 300


class PdfError(Exception):
    """导出 PDF 时的用户可见错误。"""


def _word_prog_ids():
    """可能的 Word/WPS COM  ProgID（按优先级）。"""
    return [("word", "Word.Application"), ("wps", "KWPS.Application")]


def detect_renderers():
    """本机有哪些渲染器可用。返回 ``[{name, kind, detail}, …]``。"""
    found = []
    for name, prog_id in _word_prog_ids():
        try:
            import win32com.client
            app = win32com.client.Dispatch(prog_id)
        except Exception as exc:                        # noqa: BLE001 - 探测失败就是没有
            found.append({"name": name, "kind": "com", "available": False,
                          "detail": u"%s 起不来：%s" % (prog_id, exc)})
            continue
        try:
            version = app.Version
        except Exception:
            version = u"?"
        try:
            app.Quit()
        except Exception:
            pass
        found.append({"name": name, "kind": "com", "available": True,
                      "detail": u"%s（版本 %s）" % (prog_id, version)})
    soffice = shutil.which("soffice") or shutil.which("soffice.bin")
    found.append({"name": "libreoffice", "kind": "cli",
                  "available": bool(soffice),
                  "detail": soffice or u"PATH 里没有 soffice"})
    return found


def plan(docx_path, out_pdf, prefer=None):
    """打算怎么导（不真跑）：用哪个渲染器、什么命令。"""
    if not os.path.exists(docx_path):
        raise PdfError(u"文件不存在：%s" % docx_path)
    renderers = [item for item in detect_renderers() if item["available"]]
    if not renderers:
        raise PdfError(u"本机没有可用的 PDF 渲染器（Word / WPS / LibreOffice 都没找到）。"
                       u"装一个再来；本工具只做编排，不自己渲染。")
    chosen = None
    if prefer:
        chosen = next((item for item in renderers if item["name"] == prefer), None)
        if chosen is None:
            raise PdfError(u"指定用 %r，但本机没找到它（可用：%s）"
                           % (prefer, u"、".join(item["name"] for item in renderers)))
    chosen = chosen or renderers[0]
    return {"file": docx_path, "out": os.path.abspath(out_pdf),
            "renderer": chosen["name"], "kind": chosen["kind"], "detail": chosen["detail"],
            "command": _describe(chosen, docx_path, out_pdf)}


def _describe(renderer, docx_path, out_pdf):
    if renderer["kind"] == "com":
        return u"%s 只读打开 → ExportAsFixedFormat(PDF) → 关闭不保存" % renderer["name"]
    return u'soffice --headless --convert-to pdf --outdir "%s" "%s"' % (
        os.path.dirname(os.path.abspath(out_pdf)) or ".", docx_path)


def export(docx_path, out_pdf, prefer=None, timeout=DEFAULT_TIMEOUT, visible=True):
    """导出 PDF。返回报告（渲染器、耗时、字节数、页数）。"""
    plan_info = plan(docx_path, out_pdf, prefer)
    started = time.time()
    if plan_info["kind"] == "com":
        report = _export_with_com(plan_info, timeout=timeout, visible=visible)
    else:
        report = _export_with_soffice(plan_info, timeout=timeout)
    report["file"] = plan_info["file"]
    report["detail"] = plan_info["detail"]
    report["seconds"] = round(time.time() - started, 1)
    if not os.path.exists(plan_info["out"]):
        raise PdfError(u"渲染器说成功，但没有看到文件：%s" % plan_info["out"])
    report["bytes"] = os.path.getsize(plan_info["out"])
    report["out"] = plan_info["out"]
    report["pages"] = _pdf_page_count(plan_info["out"]) or report.get("pages")
    report["pages_reported"] = report.get("pages_pdf")
    return report


def _pdf_page_count(path):
    """数 PDF 的页数：取页树里的 /Count N（没有就数 /Type /Page）。

    为什么不信渲染器自报的页数：实测 Word 的 ComputeStatistics 在这份文档上返回 1，
    而 WPS 返回 26、PDF 里的 /Count 也是 26 —— 交付物是 PDF，页数就该按 PDF 算。
    （轻量解析：不解压对象流，只认页树计数；数不出来返回 None，不猜。）
    """
    try:
        with io.open(path, "rb") as handle:
            raw = handle.read()
    except OSError:
        return None
    counts = [int(value) for value in re.findall(rb"/Count\s+(\d+)", raw)]
    pages = max(counts) if counts else None
    if pages is None:
        # 用否定环视而不是 [^s]：文件末尾的 "/Type /Page" 后面没有字符，[^s] 会漏数。
        pages = len(re.findall(rb"/Type\s*/Page(?![s])", raw)) or None
    return pages


def _export_with_com(plan_info, timeout, visible):
    """用 Word/WPS 的 COM 导出。"""
    import win32com.client

    result = {}

    def worker():
        app = None
        doc = None
        try:
            import pythoncom
            # **COM 必须在自己的线程里初始化单元**（`CoInitialize`）—— 不调就是
            # "尚未调用 CoInitialize"（实测踩过；上一轮 MacroToolbox 也是这个坑）。
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            prog_id = dict((name, prog) for name, prog in _word_prog_ids())[plan_info["renderer"]]
            app = win32com.client.Dispatch(prog_id)
            app.Visible = bool(visible)
            # **必须给绝对路径**：Word 的 COM 会话有自己的工作目录（实测解析成
            # C:\Windows\system32	mp.docx → "找不到您的文件"）。
            doc = app.Documents.Open(os.path.abspath(plan_info["file"]), ReadOnly=True)
            try:
                # **先重排再统计**：实测直接 ComputeStatistics 会返回 1（Word 还没分页），
                # 而同一份文档 WPS 给的是 26 —— 真实页数 26（PDF 里 /Count 也是 26）。
                doc.Repaginate()
                result["pages"] = int(doc.ComputeStatistics(WD_STATISTIC_PAGES))
            except Exception:
                result["pages"] = None
            doc.ExportAsFixedFormat(plan_info["out"], WD_EXPORT_FORMAT_PDF)
        except Exception as exc:                        # noqa: BLE001 - 原样带回
            result["error"] = u"%s: %s" % (type(exc).__name__, exc)
        finally:
            try:
                if doc is not None:
                    doc.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if app is not None:
                    app.Quit()
            except Exception:
                pass
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        raise PdfError(u"导出超过 %d 秒还没结束，已放弃等待（可能有个 %s 窗口还开着，"
                       u"请手工关掉；文档是只读打开的，不会被改）"
                       % (timeout, plan_info["renderer"]))
    if "error" in result:
        raise PdfError(u"导出失败：%s" % result["error"])
    return {"renderer": plan_info["renderer"], "kind": "com",
            "pages": result.get("pages"), "pages_pdf": result.get("pages")}


def _export_with_soffice(plan_info, timeout):
    out_dir = os.path.dirname(plan_info["out"]) or "."
    command = ["soffice", "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", out_dir, plan_info["file"]]
    try:
        finished = subprocess.run(command, capture_output=True, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise PdfError(u"LibreOffice 导出超过 %d 秒，已中止" % timeout)
    if finished.returncode != 0:
        raise PdfError(u"LibreOffice 导出失败：%s"
                       % (finished.stderr.decode("utf-8", "replace")[:300] or u"（无输出）"))
    return {"renderer": "libreoffice", "kind": "cli", "pages": None, "pages_pdf": None}


def format_report(report):
    lines = [u"文件：%s" % report["file"],
             u"渲染器：%s（%s）" % (report["renderer"], report.get("detail") or u""),
             u"已写出：%s（%d 字节，%s 页，耗时 %s 秒）"
             % (report["out"], report.get("bytes") or 0,
                report.get("pages") if report.get("pages") else u"?",
                report.get("seconds"))]
    return u"\n".join(lines)
