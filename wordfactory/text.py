# -*- coding: utf-8 -*-
"""文本层：把"逻辑文本"与 XML 里的 ``w:r`` / ``w:t`` 对应起来。

**这是整个项目的地基，也是唯一真正的难点。**

Word 里肉眼看到的"一句话"，在 XML 里往往是好几块：
格式一变就切一个 ``w:r``，拼写检查、修订、`lang` 标记、甚至保存时的随机因素都会再切碎 ``w:t``。
所以：

- **不能**对 XML 做朴素字符串替换（会把半句话吃掉、或把标记写坏）；
- **必须**先建立"逻辑文本 ↔ 文本节点"的映射，在逻辑文本上做替换，再按映射把结果写回。

本模块只认一件事：**一个段落（``w:p``）内部**。
跨段落的事（例如"整篇替换"）由调用方逐段落调用——这既安全（不会吃段落标记）又够快。

两个必须自己负责的细节：

1. ``w:t`` 的 ``xml:space="preserve"``：文本首尾有空格时**必须**带这个属性，
   否则 Word 会把空格吃掉（我们做"删空格/加空格"的功能，全靠它）。
2. 不越界：不进入嵌套段落（文本框、``w:txbxContent``），不碰修订删除文本（``w:delText`` 不是 ``w:t``）。
"""

import re

from .ooxml import local_name, qn

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

#: 这些元素在逻辑文本里各算一个字符，替换时要当成"原子"对待
_ATOMIC = {
    qn("w:tab"): u"\t",
    qn("w:br"): u"\n",
    qn("w:cr"): u"\n",
    qn("w:noBreakHyphen"): u"-",
    qn("w:softHyphen"): u"",
    qn("w:sym"): u"",
}

#: 遇到这些就不往里走了（嵌套段落 / 文本框 / 修订删除的内容）
_STOP = {qn("w:p"), qn("w:txbxContent"), qn("w:del"), qn("w:ins"),
         qn("w:footnote"), qn("w:endnote")}


class TextNode(object):
    """一个文本载体：要么是 ``w:t``（可改文本），要么是 ``w:tab`` 这类固定字符。"""

    __slots__ = ("element", "start", "end", "run", "editable")

    def __init__(self, element, start, end, run, editable):
        self.element = element
        self.start = start
        self.end = end
        self.run = run
        self.editable = editable

    @property
    def text(self):
        if self.editable:
            return self.element.text or ""
        return _ATOMIC.get(self.element.tag, u"")

    def set_text(self, value):
        if not self.editable:
            return
        self.element.text = value
        _fix_space_attribute(self.element, value)

    def __repr__(self):
        return "<TextNode %s [%d:%d] %r>" % (local_name(self.element.tag),
                                             self.start, self.end, self.text[:12])


def _fix_space_attribute(element, text):
    """首尾有空白就必须声明 ``xml:space="preserve"``，否则 Word 会吃掉它。"""
    if text[:1].isspace() or text[-1:].isspace():
        element.set(XML_SPACE, "preserve")
    elif XML_SPACE in element.attrib:
        del element.attrib[XML_SPACE]


class Paragraph(object):
    """一个 ``w:p`` 的逻辑视图。"""

    def __init__(self, element):
        self.element = element
        self._nodes = None
        self._text = None
        self._runs = None

    # ------------------------------------------------------------------ 构建映射
    def _walk(self, element, run, out):
        for child in element:
            tag = child.tag
            if tag == qn("w:r"):
                self._walk(child, child, out)               # 进入 run，带上它作"归属"
            elif tag == qn("w:t"):
                out.append((child, run, True))
            elif tag in _ATOMIC:
                out.append((child, run, False))
            elif tag in _STOP:
                continue                                     # 嵌套段落/文本框/修订：不走进去
            else:
                self._walk(child, run, out)                  # w:hyperlink、w:smartTag 等包一层
        return out

    def _build(self):
        out = self._walk(self.element, None, [])

        nodes = []
        cursor = 0
        for element, run, editable in out:
            text = (element.text or "") if editable else _ATOMIC.get(element.tag, u"")
            if not text:
                # 空 w:t 也留着（它属于某个 run，删掉可能影响该 run 的其它属性）
                nodes.append(TextNode(element, cursor, cursor, run, editable))
                continue
            nodes.append(TextNode(element, cursor, cursor + len(text), run, editable))
            cursor += len(text)
        self._nodes = nodes
        self._text = u"".join(node.text for node in nodes)
        return nodes

    @property
    def nodes(self):
        return self._nodes if self._nodes is not None else self._build()

    @property
    def text(self):
        if self._text is None:
            self._build()
        return self._text

    @property
    def runs(self):
        """按文档顺序的 ``w:r`` 元素列表（去重）。"""
        seen = []
        for node in self.nodes:
            if node.run is not None and node.run not in seen:
                seen.append(node.run)
        return seen

    @property
    def run_count(self):
        return len(self.runs)

    def run_map(self):
        """``[(run 元素, 逻辑文本片段), …]`` —— 用来给人看"这句话被切成了几块"。"""
        out = []
        for node in self.nodes:
            if not node.text:
                continue
            if out and out[-1][0] is node.run:
                out[-1][1] += node.text
            else:
                out.append([node.run, node.text])
        return [(run, text) for run, text in out]

    # ------------------------------------------------------------------ 替换
    def find(self, needle, start=0):
        """下一个匹配的 (起, 止) 逻辑下标；没有则 None。"""
        index = self.text.find(needle, start)
        return None if index < 0 else (index, index + len(needle))

    def replace(self, needle, replacement, count=0):
        """在**逻辑文本**上替换，跨 run 也能正确落回。

        返回实际替换的处数。``count=0`` 表示全部替换。
        """
        if not needle:
            return 0
        done = 0
        cursor = 0
        while True:
            hit = self.find(needle, cursor)
            if hit is None:
                break
            start, end = hit
            self._replace_span(start, end, replacement)
            done += 1
            cursor = start + len(replacement)
            if count and done >= count:
                break
        return done

    def _replace_span(self, start, end, replacement):
        """把 ``[start, end)`` 这段逻辑文本换成 ``replacement``。

        手法：命中的**第一个**文本节点接收整段替换文本（于是**保留了首个 run 的格式**，
        这也是 Word 自己"查找替换"的行为），其余节点的命中部分删掉。
        """
        self._build()
        first = True
        for node in self._nodes:
            if node.end <= start or node.start >= end or not node.editable:
                continue
            local_start = max(start, node.start) - node.start
            local_end = min(end, node.end) - node.start
            text = node.text
            if first:
                node.set_text(text[:local_start] + replacement + text[local_end:])
                first = False
            else:
                node.set_text(text[:local_start] + text[local_end:])
        self._nodes = None                       # 让下次访问重建映射
        self._text = None

    def delete(self, needle, count=0):
        return self.replace(needle, u"", count=count)

    def replace_regex(self, pattern, replacement, count=0):
        """按正则替换（``replacement`` 里可用 ``\\1`` 反向引用）。"""
        compiled = re.compile(pattern)
        done = 0
        cursor = 0
        while True:
            match = compiled.search(self.text, cursor)
            if match is None:
                break
            self._replace_span(match.start(), match.end(),
                               match.expand(replacement))
            done += 1
            cursor = match.start() + len(match.expand(replacement))
            if count and done >= count:
                break
        return done


def paragraphs(root):
    """一个 XML 部件里**最外层**的段落（不进文本框、不进表格单元格之外的嵌套）。"""
    out = []
    for element in root.iter(qn("w:p")):
        out.append(Paragraph(element))
    return out


def story_text(root):
    """整篇的逻辑文本（段落之间用 ``\\n`` 连接）——给人看/做断言用。"""
    return u"\n".join(paragraph.text for paragraph in paragraphs(root))


def set_run_text(run_element, text):
    """把一个 run 的文本整体换成 ``text``（保留它的 ``w:rPr``）。"""
    texts = [node for node in run_element if node.tag == qn("w:t")]
    if not texts:
        from xml.etree import ElementTree as ET
        node = ET.SubElement(run_element, qn("w:t"))
        node.text = text
        _fix_space_attribute(node, text)
        return
    texts[0].text = text
    _fix_space_attribute(texts[0], text)
    for extra in texts[1:]:
        run_element.remove(extra)
