# -*- coding: utf-8 -*-
"""页码探针：**用 Word/WPS 打开文档，读出每个正文块真正落在第几页**。

用户 2026-09-30 的原话（这也是本模块存在的理由）：

> 你找的是排版空段……我说的页面就是 word 里识别的左下角这个页面 3/26……
> 你找出那么多页面只能说明找的符号不对，**你是否能读取 word 里这种的页面，
> 它预示的是这个报告要打几张纸**，这才是我最想要的

为什么不能用 XML 里的分页符算（那是我们原来��办法）：
* 排版用的**空段 + 回车**（``w:pageBreakBefore``）**不换页** —— 实测真实报告前 15 个
  这样的空段全都落在**第 1 页**，按符号数会数成 16 页；
* ``w:lastRenderedPageBreak`` 只有 Word/WPS 保存时写了才有（这份报告里一个都没有）；
* ``docProps/app.xml`` 的 ``<Pages>`` 只给**总页数**，不给"第几段在第几页"。

所以准确的办法只有一个：**让排版引擎自己算**。本模块用 COM（Word / WPS）打开文档、
``Repaginate()``，再逐段读 ``Range.Information(wdActiveEndPageNumber)``。
实测 40 段约 0.5 秒 —— 只探测**前置区那一段**（几十个块）完全够用，不用全文扫。

边界（写清楚免得踩）：
* 需要本机装 Word 或 WPS（与 PDF 导出同一套依赖），没装就明说"读不了真实页码"；
* **只读不写**：ReadOnly 打开、关掉不保存，用户的文件一个字都不会动；
* COM 调用必须在初始化了 COM 的线程里（本模块每次都自己起线程并 CoInitialize）。
"""

import os
import threading

#: wdActiveEndPageNumber —— Range 在哪一页
WD_ACTIVE_END_PAGE = 3
#: wdStatisticPages —— 总页数
WD_STATISTIC_PAGES = 2


class PageProbeError(Exception):
    """读不到真实页码时的用户可见错误。"""


def block_start_paragraphs(document, limit=None):
    """每个 body 块的**起始段落序号**（1-based，与 Word/WPS 的 Paragraphs(i) 对应）。

    为什么需要它：COM 的 ``doc.Paragraphs(i)`` 按**文档顺序**数所有段落——
    表格里的、目录 sdt 里的段落都算；而 body 的**块索引**只数直接子元素。
    目录一个 sdt 里往往有几十个段落，不按段落序号对齐的话，sdt 之后的
    所有块页码全部错位（2026-09-30 实测：用户报告目录后页码整体偏小）。
    """
    from .ooxml import qn
    starts = []
    counter = 0
    blocks = list(document.body())
    if limit:
        blocks = blocks[:limit]
    for element in blocks:
        first = None
        for node in element.iter(qn("w:p")):
            counter += 1
            if first is None:
                first = counter
        starts.append(first if first is not None else counter + 1)
    return starts


def probe(path, blocks=60, renderer=None, timeout=180, block_starts=None):
    """…（docstring 不动，只加参数说明）…

    ``block_starts``：每个 body 块的**起始段落序号**（1-based，由
    :func:`block_start_paragraphs` 算出）。给了就按段落序号读页码 ——
    **必须给**，否则 sdt（目录）之后的块页码全部错位
    （2026-09-30 实测：sdt 里 20 个段落没算进 Paragraphs 序号，
    目录之后的块页码整体偏小）。"""
    """返回 ``{"pages": 总页数, "block_pages": [每个块的页码, …]}``。

    ``blocks`` 只探测正文开头这么多个块（前置区判断足够；全文扫慢且没必要）。
    ``renderer`` 可指定 ``"word"`` / ``"wps"``；不给就按 Word → WPS 顺序找第一个能用的。
    """
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise PageProbeError(u"文件不存在：%s" % path)
    result = {}

    def worker():
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        app = doc = None
        #: **这个实例是不是我们起的**。用户自己开着 Word 时 ``Dispatch`` 会**连到他那个实例**上
        #  （Word 是单实例 COM 服务器）—— 那种情况下我们**绝不能**改它的 Visible、更不能 Quit，
        #  否则会把用户正在编辑的文档一起关掉。所以：先试 ``GetActiveObject``，连上了就只借不用。
        ours = False
        try:
            prog_ids = ([renderer] if renderer else
                        [u"Word.Application", u"KWPS.Application"])
            app = None
            errors = []
            for prog_id in prog_ids:
                try:
                    app = win32com.client.GetActiveObject(prog_id)   # 已经开着 → 借用
                    break
                except Exception:                                    # noqa: BLE001 - 没开就自己起
                    try:
                        app = win32com.client.Dispatch(prog_id)
                        ours = True
                        break
                    except Exception as exc:                         # noqa: BLE001 - 换下一个
                        errors.append(u"%s: %s" % (prog_id, exc))
            if app is None:
                raise PageProbeError(u"本机没有可用的 Word/WPS（%s）"
                                     % u"；".join(errors))
            if ours:
                app.Visible = False
            doc = app.Documents.Open(path, ReadOnly=True)
            try:
                doc.Repaginate()                       # 先排版，页码才是准的
            except Exception:                          # noqa: BLE001 - 老版本没这方法
                pass
            result["pages"] = doc.ComputeStatistics(WD_STATISTIC_PAGES)
            pages = []
            total = doc.Paragraphs.Count
            if block_starts:
                for start in block_starts:
                    try:
                        pages.append(int(doc.Paragraphs(int(start)).Range.Information(
                            WD_ACTIVE_END_PAGE)))
                    except Exception:                  # noqa: BLE001 - 拿不到沿用上一个
                        pages.append(pages[-1] if pages else 1)
            else:
                for index in range(1, min(int(blocks), total) + 1):
                    try:
                        pages.append(int(doc.Paragraphs(index).Range.Information(
                            WD_ACTIVE_END_PAGE)))
                    except Exception:                  # noqa: BLE001 - 拿不到就沿用上一个
                        pages.append(pages[-1] if pages else 1)
            result["block_pages"] = pages
            result["renderer"] = prog_id
        finally:
            try:
                if doc is not None:
                    doc.Close(SaveChanges=False)
            except Exception:                          # noqa: BLE001
                pass
            try:
                if app is not None and ours:
                    app.Quit()                         # **只关我们自己起的那个实例**
            except Exception:                          # noqa: BLE001
                pass
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        raise PageProbeError(u"读页码超过 %d 秒没回来（Word/WPS 可能卡在弹窗上）" % timeout)
    if "block_pages" not in result:
        raise PageProbeError(u"没读到页码：%s" % result.get("error", u"未知原因"))
    return result


def block_pages_for(document, limit=120, renderer=None):
    """给一份已打开的文档补上"每个正文块在第几页"（供前置区识别用）。

    返回 ``[页码, …]``，与 ``list(document.body())[:limit]`` 一一对应；读不到就返回 None。
    ``limit`` 必须覆盖**前置区全部块**（2026-09-30 实测：默认 60 时目录 sdt 的段落
    序号 66 超界，页码沿用 1 —— 目录被识别成"第 1 页"）。前置区最多 120 块。
    """
    path = getattr(document.package, "path", None)
    if not path or not os.path.exists(path):
        return None
    starts = block_start_paragraphs(document, limit=limit + 5)
    try:
        info = probe(path, blocks=limit + 5, renderer=renderer,
                     block_starts=starts)
    except PageProbeError:
        return None
    return info["block_pages"][:limit]


def describe(info):
    """人话版（报告里显示用）。"""
    return u"用 %s 排版后共 %s 页" % (info.get("renderer", u"?"), info.get("pages", "?"))
