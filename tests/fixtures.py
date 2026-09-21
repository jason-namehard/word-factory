# -*- coding: utf-8 -*-
"""夹具工厂：造一个**最小但合法**的 .docx，供单测与差分测试用。

为什么要自己造而不是拿真文档当夹具：真文档会随 Word 版本漂移，而这个文件是我方**逐字节控制**的，
所以"改了哪些部件、有没有多改"可以被机械断言。
"""

import io
import zipfile

from wordfactory.ooxml import NAMESPACES

W = NAMESPACES["w"]
R = NAMESPACES["r"]

CONTENT_TYPES = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

ROOT_RELS = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="%s">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
</w:styles>""" % W

DOC_HEAD = u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="%s" xmlns:r="%s">
  <w:body>""" % (W, R)

DOC_TAIL = u"""
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>
  </w:body>
</w:document>"""


#: ``w:rPr`` 子元素的顺序（OOXML 里是固定序列，夹具也照这个顺序写，别造出 Word 不认的 XML）
RPR_ORDER = ("rstyle", "rfonts", "color", "highlight", "sz", "u", "vertAlign")


def run(text, **props):
    """一个 ``w:r``。

    - 普通 ``props`` 写成 ``w:rPr`` 里的元素（如 ``color='FF0000'``）；
    - ``rfonts={'eastAsia': '仿宋', 'ascii': '仿宋'}`` 写成 ``w:rFonts``（夹具的字体测试要用）；
    - ``rstyle='26'`` 写成 ``w:rStyle w:val="26"``（用来测"字体是从样式继承来的"）。
    """
    rfonts = props.pop("rfonts", None)
    rstyle = props.pop("rstyle", None)
    bits = {}
    if rstyle:
        bits["rstyle"] = u'<w:rStyle w:val="%s"/>' % rstyle
    if rfonts:
        bits["rfonts"] = u"<w:rFonts %s/>" % u" ".join(
            u'w:%s="%s"' % (key, value) for key, value in sorted(rfonts.items()))
    for key, value in props.items():
        bits[key] = u'<w:%s w:val="%s"/>' % (key, value)
    rpr = u""
    if bits:
        ordered = [bits[key] for key in RPR_ORDER if key in bits]
        ordered += [bits[key] for key in sorted(bits) if key not in RPR_ORDER]
        rpr = u"<w:rPr>%s</w:rPr>" % u"".join(ordered)
    return u'<w:r>%s<w:t xml:space="preserve">%s</w:t></w:r>' % (rpr, text)


def paragraph(*runs, **props):
    """一个 ``w:p``；``props`` 支持 ``style``（段落样式）与 ``mark_fonts``（段落标记的字体）。"""
    pr = u""
    style = props.get("style")
    mark_fonts = props.get("mark_fonts")
    if style:
        pr += u'<w:pStyle w:val="%s"/>' % style
    if mark_fonts:
        pr += u"<w:rPr><w:rFonts %s/></w:rPr>" % u" ".join(
            u'w:%s="%s"' % (key, value) for key, value in sorted(mark_fonts.items()))
    ppr = u"<w:pPr>%s</w:pPr>" % pr if pr else u""
    return u"<w:p>%s%s</w:p>" % (ppr, u"".join(runs))


def table(rows, cols=2):
    out = [u"<w:tbl><w:tblPr/><w:tblGrid>%s</w:tblGrid>"
           % (u'<w:gridCol w:w="2000"/>' * cols)]
    for row in rows:
        cells = u"".join(u"<w:tc><w:tcPr/><w:p>%s</w:p></w:tc>" % cell for cell in row)
        out.append(u"<w:tr>%s</w:tr>" % cells)
    out.append(u"</w:tbl>")
    return u"".join(out)


def build_document(body_xml):
    return DOC_HEAD + body_xml + DOC_TAIL


#: 一套固定的夹具正文：故意把一句话切成 3 个 run（这是替换类宏的真实难点）
FIXTURE_BODY = u"".join([
    paragraph(run(u"这是一个普通段落。")),
    paragraph(run(u"山"), run(u"亭区", color="FF0000"), run(u"的报告")),
    paragraph(run(u"高亮文字", highlight="yellow"), run(u"普通文字")),
    paragraph(run(u"面积 12", ), run(u"3", vertAlign="superscript"), run(u" 平方米")),
    paragraph(run(u"标题一"), style="Heading1"),
    table([[u"<w:r><w:t>表头一</w:t></w:r>", u"<w:r><w:t>表头二</w:t></w:r>"],
           [u"<w:r><w:t>值1</w:t></w:r>", u"<w:r><w:t>值2</w:t></w:r>"]]),
])

EXTRA_PARTS = {
    "word/styles.xml": STYLES,
    "customXml/item1.xml": u"<root><note>某个与本工具无关的部件，用来验证「只改该改的」</note></root>",
}


def write_fixture(path, body=None, extra=None):
    """写一个最小 .docx；``extra`` 可加部件（用来验证"没被改的部件应原字节不变"）。"""
    body = FIXTURE_BODY if body is None else body
    parts = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/document.xml", build_document(body)),
    ]
    merged = dict(EXTRA_PARTS)
    merged.update(extra or {})
    parts.extend(sorted(merged.items()))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, text in parts:
            archive.writestr(name, text.encode("utf-8"))
    return path
