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


_PROBE_CACHE = {}
_CACHE_LOCK = threading.Lock()

#: 上次成功的渲染器（进程内记着）—— 下次先试它，别每次都拿 Word 去撞墙
_PREFERRED = []


def fingerprint(path):
    import hashlib
    with open(path, 'rb') as handle:
        return hashlib.sha256(handle.read()).hexdigest()


def _candidates(renderer, network_links):
    """先试谁：显式指定 > 上次成功的 > （有网络外链就先 WPS）> 默认 Word→WPS。"""
    if renderer:
        return [renderer]
    order = [u"Word.Application", u"KWPS.Application", u"WPS.Application"]
    if network_links:
        order = [u"KWPS.Application", u"WPS.Application", u"Word.Application"]
    for preferred in reversed(_PREFERRED):
        if preferred in order:
            order.remove(preferred)
            order.insert(0, preferred)
    return order


def probe(path, blocks=60, renderer=None, timeout=60, block_starts=None, force=False):
    """用 Word/WPS 排一遍版，读"每个正文块落在第几页"。

    ``block_starts``：每个 body 块的**起始段落序号**（1-based）—— **必须给**，
    否则 sdt（目录）之后的块页码全部错位（2026-09-30 实测）。

    ``timeout``：**整件事**的预算（秒）；单个候选的预算是它的均分（见
    :func:`wordfactory.officecom.run`）。超时/起不来就抛 ``PageProbeError``，
    调用方退回"按分页符估算"。实测正常一次 1~3 秒。

    结果按"文件内容指纹"缓存（``force=True`` 强制重探）：同一份文件在一轮里
    会被问好几次（前置区、运行前、导 PDF），没必要每次都起一遍 Office。
    """
    from . import doclinks, officecom
    import copy
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        raise PageProbeError(u"文件不存在：%s" % path)
    digest = fingerprint(path)
    key = (os.path.normcase(path), digest, int(blocks), tuple(block_starts or ()), renderer)
    with _CACHE_LOCK:
        cached = _PROBE_CACHE.get(key)
    if cached is not None and not force:
        return copy.deepcopy(cached)
    network_links = doclinks.network_ole_links(path)
    candidates = _candidates(renderer, network_links)

    def make_read(target):
        """造一个"用这个排版引擎打开 ``target`` 并读页码"的回调。"""
        def read(session):
            document = session.open(target)
            document.Repaginate()                   # 先排版，页码才是准的
            count = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
            paragraphs = document.Paragraphs
            total = int(paragraphs.Count)
            starts = (list(block_starts) if block_starts
                      else list(range(1, min(int(blocks), total) + 1)))
            pages = []
            for start in starts:
                start = int(start)
                if not 1 <= start <= total:
                    raise PageProbeError(u"正文块的段落序号超出 Word/WPS 文档范围：%s / %s"
                                         % (start, total))
                # wdActiveEndPageNumber —— Word 状态栏那个口径
                page = int(paragraphs(start).Range.Information(WD_ACTIVE_END_PAGE))
                if not 1 <= page <= count:
                    raise PageProbeError(u"Word/WPS 返回了无效物理页码：%s / %s" % (page, count))
                pages.append(page)
            return {"pages": count, "block_pages": pages,
                    "renderer": session.prog_id, "fingerprint": digest,
                    "network_links": network_links}
        return read

    temp_path = doclinks.neutralized_copy(path)
    errors = []
    try:
        # 有网络外链时**用那份改过链接的副本**排版（版式一样，但 Word 不会再去等网络）
        target = temp_path or path
        for candidate in candidates:
            try:
                result = officecom.run(make_read(target), [candidate], timeout=timeout)
            except officecom.OfficeError as error:
                errors.append(u"%s: %s" % (candidate, error))
                continue
            with _CACHE_LOCK:
                if len(_PROBE_CACHE) > 32:
                    _PROBE_CACHE.clear()
                _PROBE_CACHE[key] = copy.deepcopy(result)
            _PREFERRED[:] = [result["renderer"]]
            return result
    finally:
        doclinks.discard(temp_path)
    raise PageProbeError(u"读取真实页码失败：%s" % u"；".join(errors))


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
