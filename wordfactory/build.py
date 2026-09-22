# -*- coding: utf-8 -*-
"""从零造一个最小可用的 `.docx`（给"预览"用：把候选款式各渲染一张表，你用 Word 打开挑）。

为什么需要它：本项目平时都是"打开→改→另存"，但**预览**要的是一份**新建**的文档
（不能拿用户的报告当草稿）。这里用与 `ooxml.py` 同一套纪律手搓最小包：

* 部件最少但要合法（`[Content_Types].xml` / `_rels/.rels` / `word/document.xml`）；
* **zip 时间戳固定** → 输出逐字节可复现（与项目其余部分一致）；
* 只依赖标准库。
"""

import os
import zipfile

from .ooxml import _XML_HEADER, NAMESPACES

CONTENT_TYPES = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

#: 固定时间戳（1980-01-01）—— 输出可复现
FIXED_DATE = (1980, 1, 1, 0, 0, 0)

#: A4 纵向 + 常用页边距（跟用户报告一个量级，预览才有意义）
SECT_PR = (u'<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
           u'<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
           u'w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>')


def document_xml(body_xml):
    """把正文片段包成一份合法的 `word/document.xml`。"""
    return (u'%s<w:document xmlns:w="%s" xmlns:r="%s"><w:body>%s%s</w:body></w:document>'
            % (_XML_HEADER, NAMESPACES["w"], NAMESPACES["r"], body_xml, SECT_PR))


def write_document(path, body_xml):
    """写一个最小 `.docx`；返回路径。"""
    parts = [("[Content_Types].xml", CONTENT_TYPES),
             ("_rels/.rels", ROOT_RELS),
             ("word/document.xml", document_xml(body_xml))]
    out_path = os.path.abspath(path)
    parent = os.path.dirname(out_path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    temp = out_path + ".writing"
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, text in parts:
                info = zipfile.ZipInfo(name, date_time=FIXED_DATE)
                info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(info, text.encode("utf-8"))
        os.replace(temp, out_path)
    except Exception:
        if os.path.exists(temp):
            try:
                os.remove(temp)
            except OSError:
                pass
        raise
    return out_path
