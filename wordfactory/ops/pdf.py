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
DEFAULT_TIMEOUT = 90


class PdfError(Exception):
    """导出 PDF 时的用户可见错误。"""


def _word_prog_ids():
    """可能的 Word/WPS COM ProgID（按优先级）。

    * ``Word.Application`` —— 装了 MS Word 是它；**WPS 设成 Word 默认打开程序时也会占这个名字**；
    * ``KWPS.Application`` —— WPS Office 注册的名字（本机实测 12.0 就是它）；
    * ``WPS.Application``  —— 另一些 WPS 版本/个人版注册的名字（多写一个候选，代价为零）。
    """
    return [("word", "Word.Application"), ("wps", "KWPS.Application"),
            ("wps", "WPS.Application")]


def _in_com_thread(func, timeout=60):
    """在一个**自己初始化 COM** 的线程里跑 ``func``，拿回它的返回值。

    为什么需要它：COM 对象必须在初始化过 `CoInitialize` 的线程里创建和使用。
    GUI 的 HTTP 处理线程是 `ThreadingHTTPServer` 现拉的，没初始化过 COM ——
    在那里直接 `Dispatch` 只会拿到"尚未调用 CoInitialize"，渲染器探测于是被误判成
    "本机没有渲染器"（实测：GUI 点「导出 PDF」说没渲染器，命令行同一个文件却能导出来）。
    """
    box = {}

    def worker():
        try:
            import pythoncom
            try:
                pythoncom.CoInitialize()
            except Exception:                     # 已经初始化过就接着用
                pass
        except ImportError:                       # 没装 pywin32：让 func 自己报错
            pass
        try:
            box["value"] = func()
        except Exception as exc:                  # noqa: BLE001 - 原样带出线程
            box["error"] = exc
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        raise PdfError(u"COM 调用超过 %d 秒没回来（渲染器可能卡在弹窗上）" % timeout)
    if "error" in box:
        raise box["error"]
    return box.get("value")


def _probe_com_renderer(prog_id):
    """Check registration without starting and quitting an Office session."""
    import pywintypes
    pywintypes.IID(prog_id)
    return '已注册'


def detect_renderers():
    found = []
    for name,prog_id in _word_prog_ids():
        try:
            detail = _probe_com_renderer(prog_id)
            available = True
        except Exception as error:
            detail,available = str(error),False
        found.append({'name':name,'kind':'com','available':available,
                      'prog_id':prog_id,'detail':detail})
    soffice = shutil.which('soffice') or shutil.which('soffice.bin')
    found.append({'name':'libreoffice','kind':'cli','available':bool(soffice),'detail':soffice or '未安装'})
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
            "renderer": chosen["name"], "kind": chosen["kind"],
            "prog_id": chosen.get("prog_id"), "detail": chosen["detail"],
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
    with open(plan_info["out"],"rb") as handle:
        if not handle.read(5).startswith(b"%PDF-"):
            raise PdfError("导出的文件不是有效 PDF")
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
    from .. import doclinks, officecom
    first = plan_info.get('prog_id') or 'Word.Application'
    candidates = [first] + [prog_id for _name, prog_id in _word_prog_ids() if prog_id != first]
    # 文档里若有指向网络共享的图表外链，Word 打开时会干等 SMB 超时（实测 40 秒+）——
    # 换成"外链改指向本地空文件"的**临时副本**再转（版式一样，图是缓存数据画的），用完即删。
    temp_path = doclinks.neutralized_copy(plan_info['file'])
    target = temp_path or plan_info['file']
    # 一轮的 Office 会话统计：起了几个实例 / 清掉几个 / 跳过几个（什么原因）。
    # 不只为好看 —— "残留"要能在**发生的那一次**就暴露，而不是等下次卡 60 秒才发现。
    stats = officecom.new_stats()

    def convert(session):
        document = session.open(target)
        document.Repaginate()                    # 先重排，页数才准
        pages = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
        document.ExportAsFixedFormat(os.path.abspath(plan_info['out']), WD_EXPORT_FORMAT_PDF)
        actual = 'wps' if session.prog_id != 'Word.Application' else 'word'
        return {'renderer': actual, 'prog_id': session.prog_id, 'kind': 'com',
                'pages': pages, 'pages_pdf': pages,
                'links_neutralized': bool(temp_path)}

    try:
        result = officecom.run(convert, candidates, timeout=timeout, stats=stats)
    except officecom.OfficeError as error:
        # 失败时统计行由 officecom.run 的收尾写进 cleanup.log，这里只把原因抛出去
        raise PdfError(u'PDF 导出失败：%s' % error) from error
    finally:
        doclinks.discard(temp_path)
    result['office_stats'] = officecom.format_stats(stats)
    return result


def _export_with_soffice(plan_info, timeout):
    from .. import subproc
    out_dir = os.path.dirname(plan_info["out"]) or "."
    command = ["soffice", "--headless", "--norestore", "--convert-to", "pdf",
               "--outdir", out_dir, plan_info["file"]]
    try:
        # 走 subproc：exe（无控制台）里直接 subprocess 会闪一个黑窗口
        finished = subproc.run(command, capture_output=True, timeout=timeout)
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
