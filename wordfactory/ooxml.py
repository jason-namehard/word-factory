# -*- coding: utf-8 -*-
"""OOXML 容器层：打开 / 读取 / 保存一个 ``.docx``。

它只是 zip + XML，但有两件事必须做对，否则会悄悄弄坏文档：

**一、保真——只重写改动过的部件。**
一个 ``.docx`` 里除了正文还有样式、编号、图片、页眉页脚、字体表、嵌入对象……
我们只改 ``word/document.xml`` 之类的少数部件，保存时**其余部件必须按原字节搬过去**
（不重新序列化、不重新压缩它们的内容），这样才敢说"只动了该动的"。

**二、命名空间前缀不能变。**
文档里会出现 ``mc:Ignorable="w14 w15 wp14"`` 这种**引用前缀**的属性。
用 ``ElementTree`` 反序列化再序列化，如果不把常用前缀注册回去，它会写成 ``ns0:``，
那些引用就会指向不存在的名字——Word 可能直接报文件损坏。所以本模块在导入时就注册全部前缀。
"""

import io
import os
import shutil
import zipfile
from xml.etree import ElementTree as ET

#: OOXML 里会互相引用的前缀。写全一点，宁多勿漏。
NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v": "urn:schemas-microsoft-com:vml",
    "o": "urn:schemas-microsoft-com:office:office",
    "w10": "urn:schemas-microsoft-com:office:word",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "sl": "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
}

_XML_HEADER = u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'


def register_namespaces():
    """把 OOXML 前缀注册回 ElementTree（幂等）。"""
    for prefix, uri in NAMESPACES.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:                      # 保留前缀（ns\d+）或不合法时忽略
            pass


register_namespaces()


def qn(tag):
    """``"w:p"`` -> ``"{http://…/main}p"``（ElementTree 用的花括号写法）。"""
    prefix, _, local = tag.partition(":")
    return "{%s}%s" % (NAMESPACES[prefix], local)


def local_name(tag):
    """``"{http://…/main}p"`` -> ``"w:p"``（打印用）。"""
    if tag.startswith("{"):
        uri, _, local = tag[1:].partition("}")
        for prefix, known in NAMESPACES.items():
            if known == uri:
                return "%s:%s" % (prefix, local)
        return local
    return tag


class PackageError(Exception):
    """打开/保存 .docx 时的用户可见错误。"""


class DocxPackage(object):
    """一个打开的 ``.docx``（``.docm`` 同样可用，宏部件只是被原样保留）。

    用法::

        with DocxPackage("a.docx") as pkg:
            root = pkg.xml(pkg.MAIN)
            ...                          # 改 root
            pkg.mark_dirty(pkg.MAIN)     # 声明这个部件要重写
            pkg.save("b.docx")

    只有被 ``mark_dirty`` 的部件会重新序列化；其余部件按**原字节**复制。
    """

    MAIN = "word/document.xml"
    STYLES = "word/styles.xml"
    NUMBERING = "word/numbering.xml"
    SETTINGS = "word/settings.xml"
    CONTENT_TYPES = "[Content_Types].xml"

    def __init__(self, path):
        self.path = os.path.abspath(path)
        if not os.path.exists(self.path):
            raise PackageError(u"文件不存在：%s" % self.path)
        if not zipfile.is_zipfile(self.path):
            raise PackageError(
                u"%s 不是 .docx/.docm（不是 zip）。老式 .doc 是二进制格式，"
                u"本工具不处理，请先另存为 .docx。" % self.path)
        self._zip = zipfile.ZipFile(self.path, "r")
        self._names = self._zip.namelist()
        if self.MAIN not in self._names:
            self._zip.close()
            raise PackageError(u"%s 里没有 %s —— 这看起来不是 Word 文档。" % (self.path, self.MAIN))
        self._trees = {}          # part name -> ElementTree
        self._dirty = set()

    # ------------------------------------------------------------------ 读
    @property
    def part_names(self):
        return list(self._names)

    def has(self, name):
        return name in self._names

    def read_bytes(self, name):
        if name not in self._names:
            raise PackageError(u"部件不存在：%s" % name)
        return self._zip.read(name)

    def xml(self, name):
        """解析并缓存一个 XML 部件；返回它的根元素。"""
        if name not in self._trees:
            try:
                self._trees[name] = ET.fromstring(self.read_bytes(name))
            except ET.ParseError as exc:
                raise PackageError(u"%s 解析失败：%s" % (name, exc))
        return self._trees[name]

    def mark_dirty(self, name):
        """声明该部件被改过，保存时要重写它。"""
        if name not in self._trees:
            raise PackageError(u"还没读过 %s，不能标记为已改" % name)
        self._dirty.add(name)

    @property
    def dirty_parts(self):
        return sorted(self._dirty)

    # ------------------------------------------------------------------ 写
    def serialize(self, name):
        """把（已改的）XML 部件序列化成字节：声明头 + 注册好的前缀。"""
        element = self._trees[name]
        body = ET.tostring(element, encoding="utf-8", xml_declaration=False)
        return _XML_HEADER.encode("utf-8") + body

    def save(self, out_path):
        """写到新文件：改过的部件重写，其余按原字节搬。"""
        out_path = os.path.abspath(out_path)
        if os.path.abspath(out_path) == self.path:
            raise PackageError(
                u"拒绝直接覆盖原文件（%s）。本工具默认另存一份；"
                u"要覆盖请先显式确认，或保存到新文件后再自行替换。" % out_path)
        parent = os.path.dirname(out_path)
        if parent and not os.path.isdir(parent):
            os.makedirs(parent)

        temp = out_path + ".writing"
        try:
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as out:
                for name in self._names:
                    if name in self._dirty:
                        out.writestr(name, self.serialize(name))
                    else:
                        info = self._zip.getinfo(name)
                        out.writestr(info, self._zip.read(name))   # 原字节 + 原压缩信息
            os.replace(temp, out_path)
        except Exception:
            if os.path.exists(temp):
                try:
                    os.remove(temp)
                except OSError:
                    pass
            raise
        return out_path

    # ------------------------------------------------------------------ 生命周期
    def close(self):
        if self._zip is not None:
            self._zip.close()
            self._zip = None

    def __enter__(self):
        return self

    def __exit__(self, *exc_info):
        self.close()
        return False


def copy_file(src, dst):
    """保险起见：需要纯字节复制时用它（带父目录创建）。"""
    parent = os.path.dirname(os.path.abspath(dst))
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    shutil.copy2(src, dst)
    return dst


def file_sha256(path):
    import hashlib
    digest = hashlib.sha256()
    with io.open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 16), b""):
            digest.update(chunk)
    return digest.hexdigest()
