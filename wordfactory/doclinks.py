# -*- coding: utf-8 -*-
"""`.docx` 里的**网络外链**：Word 打开这类文档会干等 SMB/HTTP 超时，得先处理掉。

实测（2026-10-08，用户那份报告）：图表的数据链到别人机器上的共享
``\\Bf-230206\\2025共享文件\\…\\欧峪水库.xlsm``，Word 打开时**40 秒都回不来**，
而 WPS 不受影响（3 秒）。把这类外链在**临时副本**里改指向本地不存在的文件之后，
Word 7.9 秒就开好了，页码/版面与 WPS 完全一致 —— 因为图是拿**缓存数据**画的，版式不变。

两条规矩：

* **只动 ``oleObject`` 外链**（图表数据那一类）；外链的**图片一律不碰** —— 换了图片版面就变；
* 只在**临时副本**上改，**用户的文件一个字都不动**；副本落在数据目录的「临时文件」里
  （系统 Temp 会被 Office 当成"不安全位置"进受保护视图），用完即删。

读页码（`pageprobe`）与导 PDF（`ops/pdf`）都用它。
"""

import os
import re
import shutil
import tempfile
import zipfile

#: 外链指向"网络位置"的样子：UNC（\\机器\共享）或 http(s)
NETWORK_TARGET = re.compile(r'^\s*(?:file:///)?(?:\\\\|//|https?://)', re.I)

#: 副本里把外链改成这个（本地、不存在 → Word 立刻放弃，不去等网络）
DEAD_TARGET = u"file:///C:/__wordfactory_no_such_file__.xlsm"

_REL_RE = re.compile(r'<Relationship[^>]*>')


def network_ole_links(path):
    """文档里指向**网络位置**的 ``oleObject`` 外链有几个。"""
    count = 0
    try:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if not name.endswith(u".rels"):
                    continue
                text = archive.read(name).decode(u"utf-8", u"replace")
                for tag in _REL_RE.findall(text):
                    if u'TargetMode="External"' not in tag or u"oleObject" not in tag:
                        continue
                    target = re.search(r'Target="([^"]*)"', tag)
                    if target and NETWORK_TARGET.match(target.group(1)):
                        count += 1
    except Exception:                      # noqa: BLE001 - 读不动就当没有，别拖累主流程
        return 0
    return count


def neutralized_copy(path, folder=None):
    """生成"网络外链改指向本地空文件"的**临时副本**；没有网络外链就返回 ``None``。

    调用方负责用完删掉（``os.remove``）。
    """
    if not network_ole_links(path):
        return None
    if folder is None:
        from . import paths
        folder = paths.temp_dir()
    try:
        os.makedirs(folder, exist_ok=True)
    except OSError:
        folder = None                      # 落回系统临时目录
    handle, temp_path = tempfile.mkstemp(prefix=u"外链已屏蔽-", suffix=u".docx", dir=folder)
    os.close(handle)
    try:
        with zipfile.ZipFile(path) as source, \
                zipfile.ZipFile(temp_path, u"w", zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                data = source.read(item.filename)
                if item.filename.endswith(u".rels"):
                    text = data.decode(u"utf-8", u"replace")
                    text = re.sub(
                        r'(<Relationship[^>]*TargetMode="External"[^>]*oleObject[^>]*Target=")[^"]*(")',
                        lambda m: m.group(1) + DEAD_TARGET + m.group(2), text)
                    text = re.sub(
                        r'(<Relationship[^>]*oleObject[^>]*Target=")[^"]*("[^>]*TargetMode="External")',
                        lambda m: m.group(1) + DEAD_TARGET + m.group(2), text)
                    data = text.encode(u"utf-8")
                target.writestr(item, data)
    except Exception:                      # noqa: BLE001 - 复制失败就退回用原件
        try:
            os.remove(temp_path)
        except OSError:
            pass
        return None
    if network_ole_links(temp_path):
        try:
            os.remove(temp_path)
        except OSError:
            pass
        return None                        # 没改干净就别用它，免得白等
    return temp_path


def discard(temp_path):
    """删掉 :func:`neutralized_copy` 造出来的临时副本（没有就什么都不做）。"""
    if not temp_path:
        return
    try:
        os.remove(temp_path)
    except OSError:
        pass
