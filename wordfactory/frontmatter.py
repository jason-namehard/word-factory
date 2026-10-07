# -*- coding: utf-8 -*-
"""前置区识别：把封面、扉页、签字页从正文里**拆出来保护**。

用户 2026-09-27 的原话（口径全部来自这里）：

> 封面扉页会不可避免的出现空白行，但是它不允许被删除……空白行和空白单元格应该规避
> 封面、扉页以及前面的签字页。他们的特征很明显，**有全文最大字体**，**位于文档开头几页**，
> 这些页码后面一般会跟着"目录"或者"前言"又或者"前 言"。……判断是否没有的依据就是
> **是否遍历到目录或者前言了**，一旦遍历到那就说明前面就都是。有时候封面和扉页是一个，
> 签字页是在扉页基础上多出了人员签字，常见内容为**审核、编制、校核**等。

实现口径（v1，均为启发式，识别结果**允许手动修正**——手动修正窗口属界面层，随布局改版一起做）：

* **边界**：从文档第一个块开始走，遇到"目录 / 目次 / 前言"（去掉全部空格后**整段等于**，
  或该段里有 TOC 域指令 `TOC \o`）就停 —— 前面全部算前置区；没遇到 = 无前置区（不瞎猜）。
* **分页**：靠文档里现成的分页痕迹（显式分页符 / 段前分页 / `w:lastRenderedPageBreak`
  ——Word 上次排版时留下的记号）把前置区切成"页"。第 1 页 = 封面；中间页 = 扉页；
  最后一页若带 编制/校核/审核… 等角色词 = 签字页。
* **保护**：清理类算子（`tidy` 的空白段/空格）默认跳过前置区段落 —— 封面上的空白行
  是排版，删了版面就垮。
"""

import re

from .ooxml import qn
from .text import Paragraph

#: 前置区结束的标志（去空格后整段等于这些词，或该段带 TOC 域）
MARKERS = (u"目录", u"目次", u"前言")
#: 签字页的角色词（用户原话：常见内容为审核、编制、校核等）。
#: **不含"复核"** —— 实测（2026-09-30）报告标题"防洪安全复核报告"里的"复核"会被误判成签字页。
ROLES = (u"编制", u"校核", u"审核", u"审查", u"审定", u"核定", u"批准", u"设代")
#: 角色词必须出现在**短行**里（签字页通常一行"编制：张三"），否则正文里的"编制方案"也会误判
ROLE_LINE_MAX = 20
#: 安全阀：最多往前找多少个块（真报告的封面区不会超过这个；找不到标志就不能无限吞正文）
MAX_BLOCKS = 120


def _normalized(text):
    """去全部空白（含全角空格）——「前 言」「目 录」都要认得。"""
    return re.sub(u"[\\s\u00a0\u3000]+", u"", text or u"")


def _has_toc_field(paragraph_element):
    """这段里有没有 TOC 域指令（目录页常是一个域，文字可能还是"目录"两个字）。"""
    for node in paragraph_element.iter(qn("w:instrText")):
        if "TOC" in (node.text or u""):
            return True
    return False


def _page_breaks(paragraph_element):
    """这一段里带着几处"换页"痕迹。"""
    breaks = 0
    for node in paragraph_element.iter(qn("w:br")):
        if node.get(qn("w:type")) == "page":
            breaks += 1
    if paragraph_element.find(qn("w:pPr")) is not None:
        if paragraph_element.find(qn("w:pPr")).find(qn("w:pageBreakBefore")) is not None:
            breaks += 1
    breaks += len(list(paragraph_element.iter(qn("w:lastRenderedPageBreak"))))
    return breaks


#: 目录项的样子："1·水库概况……………………1" / "1.1 基本情况 ……… 3"
_TOC_ENTRY_RE = re.compile(u"[.．·…]{2,}\s*\d+\s*$")


def _is_toc_marker(element):
    """是不是"目录"那一块。

    三种形态都要认（实测真实报告用的是第 3 种）：
    1. 文字是「目录/目次」——**允许中间有 1~2 个空格**（用户 2026-09-30 补充的辨别依据）；
    2. 段落里有 TOC 域指令；
    3. **整块包在 ``w:sdt``（内容控件）里** —— Word/WPS 生成的目录默认是这种，
       目录标题和条目都在 sdt 里，光扫段落会整个漏掉。
    """
    tag = element.tag
    if tag == qn("w:p"):
        if _normalized(Paragraph(element).text) in (u"目录", u"目次"):
            return True
        return _has_toc_field(element)
    if tag == qn("w:sdt"):
        for node in element.iter():
            if node.tag == qn("w:docPartGallery"):
                if u"目录" in (node.get(qn("w:val")) or u"") or                         "Table of Contents" in (node.get(qn("w:val")) or u""):
                    return True
            if node.tag == qn("w:instrText") and "TOC" in (node.text or u""):
                return True
    return False


def _block_text(element):
    """一块（段落或 sdt）的文字。"""
    if element.tag == qn("w:p"):
        return Paragraph(element).text or u""
    if element.tag == qn("w:sdt"):
        return u"".join(Paragraph(node).text or u""
                        for node in element.iter(qn("w:p")))
    return u""


def _is_preface(element):
    """「前 言」「前言」——同样允许中间有空格（用户 2026-09-30 补充的辨别依据）。"""
    return element.tag == qn("w:p") and _normalized(Paragraph(element).text) == u"前言"


def _front_matter_end(blocks):
    """前置区到哪儿为止（返回正文起始下标；没找到标志返回 None）。

    口径完全按用户 2026-09-30 的描述：
    顺序一般是 **封面 → 扉页 → 签字页 → 前言 → 目录**；
    **前言、目录本身也属于前置区（不删空白）**；
    **目录后面多半有一个分页符，从那个分页符往后就是正文了**。

    找法：定位「目录」（文字或 TOC 域；没有就退用「前言」），然后从它往后连续吃掉
    空段、目录项、带分页/分节记号的段，遇到第一段**有内容且不是目录项**的就是正文。
    """
    limit = min(len(blocks), MAX_BLOCKS)
    marker = None
    for index in range(limit):
        if _is_toc_marker(blocks[index]):
            marker = index
            break
    if marker is None:
        # 没有目录 → 用「前言」当标志（但仍要越过它继续找目录）
        for index in range(limit):
            if _is_preface(blocks[index]):
                marker = index
                break
        if marker is None:
            return None
    position = marker + 1
    while position < limit:
        element = blocks[position]
        if element.tag == qn("w:tbl"):
            break
        if _is_toc_marker(element):        # 目录整块（sdt）
            position += 1
            continue
        raw = _block_text(element)
        text = _normalized(raw)
        if not text:
            position += 1                 # 空段：还在前置区（目录后的空段也算）
            continue
        if _TOC_ENTRY_RE.search(raw.strip()):
            position += 1                 # 目录项 "1·水库概况……1"
            continue
        break                              # 有内容且不是目录项 → 正文开始
    return position


def detect(document, block_pages=None):
    """识别前置区。返回 dict；**只读**，不动文档。
    block_pages = **真实页码**（pageprobe 用 Word/WPS 排版后读出的，
                       第几块落在第几页）。有它按真实页码分页；没有才
                       退回"分页符估算"（排版空段/回车不换页，估算会数多）。

    结果字段：``present``（有没有识别出前置区）、``marker``（靠哪个词定的界）、
    ``paragraphs``（前置区段数）、``pages``（按分页痕迹估的页数）、
    ``cover`` / ``title_page`` / ``signature_page``（有 / 无）、``roles``（命中的角色词）、
    ``note``（人话说明，识别不出来时写清为什么）。
    """
    body = document.body()
    blocks = list(body)[:MAX_BLOCKS]
    end = _front_matter_end(blocks)
    # **没找到标志 = 没有可靠边界 = 前置区为空**（绝不能"整篇都是前置区"——
    # 那样清理功能等于全废，测试钉过这条）
    front = list(blocks[:end]) if end is not None else []
    marker = u"目录" if any(_is_toc_marker(el) for el in front) else (
        u"前言" if any(_is_preface(el) for el in front) else u"开头")
    if not front:
        return {"present": False, "marker": None, "paragraphs": 0, "pages": 0,
                "cover": False, "title_page": False, "signature_page": False,
                "roles": [], "blocks": 0,
                "page_map": {u"cover": u"无", u"title": u"无", u"signature": u"无",
                             u"preface": u"无", u"toc": u"无", u"accurate": False},
                "note": u"没遍历到「目录/前言」，不能确定哪里是正文开头 —— 不保护（可手动指定）"}
    # 分页：优先用**真实页码**（Word/WPS 排版结果），否则退回分页符估算
    body_blocks = list(body)[:MAX_BLOCKS]
    index_of = {}
    for position, element in enumerate(body_blocks):
        index_of[id(element)] = position
    groups = []
    if block_pages:
        for element in front:
            index = index_of.get(id(element), 0)
            # block_pages 只探了开头 70 块；正文在它之外的照常用估算页号
            number = (block_pages[index] if index < len(block_pages) else 1) or 1
            if not groups or groups[-1][0] != number:
                groups.append([number, []])
            groups[-1][1].append((number, element))
    else:
        current_page = 1
        bucket = []
        for element in front:
            bucket.append((current_page, element))
            if element.tag == qn("w:p") and _page_breaks(element):
                groups.append([current_page, bucket])
                bucket = []
                current_page += 1
        if bucket:
            groups.append([current_page, bucket])
    pages = max(number for number, _blocks in groups) if groups else 1

    def group_has_text(group):
        """这一页（group = [页号, [(页号, 块), …]]）里有没有非空文字。"""
        _number, blocks = group
        return any(_normalized(Paragraph(element).text)
                   for _page, element in blocks if element.tag == qn("w:p"))

    texts = [_normalized(Paragraph(element).text) for element in front
             if element.tag == qn("w:p")]
    roles = [role for role in ROLES
             if any(role in text and len(text) <= ROLE_LINE_MAX for text in texts)]
    # **封面 = 前置区里第一个有非空文字的页**（用户 2026-09-30 实测：真实报告的封面
    # 在第 4–5 页——前面 3 页是排版空段；按"第 1 页"判会把封面判成"无"）。
    # 整片前置区都没有文字才报"无"。
    cover_index = None
    for position, group in enumerate(groups, start=1):
        if group_has_text(group):
            cover_index = position
            break
    cover = cover_index is not None
    signature_page = bool(groups) and bool(roles) and group_has_text(groups[-1])
    start = cover_index or 1
    middle = groups[start:-1] if (signature_page and len(groups) > start + 1) else (
        groups[start:] if len(groups) > start else [])
    title_page = bool(middle) and any(group_has_text(g) for g in middle)
    # 给人看的"到第几页"（用户 2026-09-28：光说"有"没法判断对错，要看到页码）
    def page_at(position):
        return groups[position][0] if 0 <= position < len(groups) else pages

    # **页码范围按用户 2026-09-30 的规则**（真实页码 groups = [[页号, 块], …]）：
    # 封面 = 第一个有文字的页；**扉页一定紧挨封面、一般只有一页**（大字体标题，
    # 签字页常与扉页同页）；前言 = 标题页 .. 第一个分页符前（前言一定有分页符收尾）；
    # 目录 = 标题页 .. 前置区末页（目录后的分页符之后就是正文）。
    # 每个标志的范围 = 自己的页 .. 下一个标志页 - 1；最后一个标志到前置区末页。
    cover_page = page_at(cover_index - 1) if cover else None      # 封面标题所在**页码**
    title_page = (cover_page + 1) if cover_page else None         # 扉页紧挨封面
    signature_page = None
    preface_page = None
    toc_page = None
    for position, group in enumerate(groups):
        blocks_here = group[1]
        texts_here = [_normalized(Paragraph(el).text) for _p, el in blocks_here
                      if el.tag == qn("w:p")]
        if preface_page is None and any(_is_preface(el) for _p, el in blocks_here):
            preface_page = group[0]
        if toc_page is None and any(_is_toc_marker(el) for _p, el in blocks_here):
            toc_page = group[0]
        if signature_page is None and any(
                role in t and len(t) <= ROLE_LINE_MAX
                for role in roles for t in texts_here):
            signature_page = group[0]
    front_last = groups[-1][0] if groups else pages

    entries = [(name, page) for name, page in
               ((u"cover", cover_page), (u"title", title_page),
                (u"signature", signature_page), (u"preface", preface_page),
                (u"toc", toc_page)) if page is not None]
    entries.sort(key=lambda item: item[1])
    distinct = sorted(set(page for _name, page in entries))

    def _range(first_page, last_page):
        if last_page <= first_page:
            return u"第 %d 页" % first_page
        return u"第 %d–%d 页" % (first_page, last_page)

    page_map = {}
    for order, (name, page) in enumerate(entries):
        following = [value for value in distinct if value > page]
        nxt = following[0] if following else front_last + 1
        last_page = nxt - 1
        if name == u"title":
            last_page = page            # **扉页固定一页**（用户：扉页一般只有一页）
        page_map[name] = _range(page, last_page)
    # 签字页与扉页同页 → 并入扉页行（显示同一页码，不单独占一行）
    if signature_page is not None and signature_page == title_page:
        page_map[u"signature"] = page_map.get(u"title", u"无")
    page_map["total_pages"] = pages
    page_map["accurate"] = bool(block_pages)
    return {"present": True, "marker": marker, "paragraphs":
            sum(1 for element in front if element.tag == qn("w:p")),
            "pages": pages, "cover": cover, "title_page": title_page,
            "signature_page": signature_page, "roles": roles, "blocks": len(front),
            "page_map": page_map,
            "note": u"「%s」之前共 %d 段、约 %d 页，已整体划为前置区" % (marker, len(front), pages)}


def _page_groups(document):
    """把正文块按分页痕迹分组成"物理页"：``[(页号, [块, …]), …]``，页号从 1 起。

    分页痕迹 = 显式分页符（``w:br type=page``）/ 段前分页 / ``w:lastRenderedPageBreak``
    （Word 上次排版留下的记号）。是**估算**：文档没重新排版过时基本准，
    差一页的情形靠手动修正兜底（用户 2026-09-27 的"手动调整"窗口就是干这个的）。
    """
    groups = []
    current = 1
    bucket = []
    for element in document.body():
        bucket.append(element)
        breaks = 0
        if element.tag == qn("w:p"):
            breaks = _page_breaks(element)
        if breaks:
            groups.append((current, bucket))
            bucket = []
            current += breaks
    if bucket:
        groups.append((current, bucket))
    return groups


def protected_elements(document, pages=None):
    """前置区里的**段落元素集合**（按 id），清理类算子用它跳过。

    * 自动识别：从头走到「目录/目次/前言」，前面全是前置区；**没遍历到标志 =
      没有可靠边界 = 一个都不保护**（绝不能把整篇当前置区）。
    * 手动页数 ``pages=N``：**与自动识别取并集**，不是替代（用户 2026-09-30 实测：
      真实报告封面在第 4–5 页，手动填"1 页"时若替代掉自动识别，封面那些排版空段
      会被清理掉 10 个 —— 页面直接垮掉）。手动数字的语义是"**至少**保护到第 N 页"。

    所以：**自动识别的永远都保护**，手动只能加不能减。
    """
    protected = set()
    if pages:
        for number, blocks in _page_groups(document):
            if number > int(pages):
                break
            for element in blocks:
                if element.tag == qn("w:p"):
                    protected.add(id(element))
    body = document.body()
    blocks = list(body)[:MAX_BLOCKS]
    seen = []
    for element in blocks:
        if element.tag != qn("w:p"):
            seen.append(element)
            continue
        text = _normalized(Paragraph(element).text)
        if text in MARKERS or _has_toc_field(element):
            protected.update(id(item) for item in seen)
            return protected
        seen.append(element)
    return protected


def format_report(info):
    """给人看的拆解结果（CLI / GUI 报告用）。"""
    if not info.get("present"):
        return u"前置区：无（%s）" % info.get("note", u"")
    lines = [u"前置区：%s" % info["note"],
             u"  封面：%s ｜ 扉页：%s ｜ 签字页：%s"
             % (u"有" if info["cover"] else u"无",
                u"有" if info["title_page"] else u"无",
                u"有" if info["signature_page"] else u"无")]
    if info["roles"]:
        lines.append(u"  签字角色词：%s" % u"、".join(info["roles"]))
    return u"\n".join(lines)
