# -*- coding: utf-8 -*-
"""最小可用的 `.xlsx` 读写（**只用标准库**）。

为什么自己写而不装 openpyxl：本项目的硬约束是"只依赖标准库"（用户要的是能长期跑的外置工具，
少一个依赖少一个坏掉的理由）。而「段落配方」要的 xlsx 契约极窄（见 `docs/REFERENCE-MACROS.md` §3）：

* 只有**一个**工作表，名字由配方里的 `SHEET_NAME:` 指定；
* `A1="项目"`、`B1="数值"`；
* **第 n 个变量固定落在第 n+1 行的 B 列**（写端与读端各自独立算 `"B" & (n+1)`，必须一致）；
* A 列是"给人看的上下文"，读端**不读**它。

`.xlsx` 就是 zip + XML：`xl/workbook.xml`（表名与 r:id）、`xl/_rels/workbook.xml.rels`（r:id → 文件）、
`xl/worksheets/sheet1.xml`（单元格）、`xl/sharedStrings.xml`（字符串表）。

**写出来要能被 Excel / WPS / Word 的 VBA 读**（这是 M3 的验收标志），所以走最常规的写法：
sharedStrings + 最小 styles.xml；不要用 `inlineStr` 这种"能省部件但少见"的写法。
另外**输出必须可复现**（zip 时间戳固定），理由同 `ooxml.py`。
"""

import io
import os
import re
import zipfile
from xml.etree import ElementTree as ET

#: 主命名空间
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

#: 固定时间戳（1980-01-01 是 zip 能表示的最早时间）—— 让输出逐字节可复现
FIXED_DATE = (1980, 1, 1, 0, 0, 0)


def _qn(namespace, tag):
    return "{%s}%s" % (namespace, tag)


class XlsxError(Exception):
    """读/写 .xlsx 时的用户可见错误。"""


# --------------------------------------------------------------------- 读
def sheet_names(path):
    """工作簿里的表名（按出现顺序）。"""
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [node.get("name") for node in root.iter(_qn(NS_MAIN, "sheet"))]


def _sheet_path(archive, sheet_name):
    """表名 → 工作表部件的路径（经 workbook.xml 的 r:id 查 rels —— 别去猜 sheet1.xml）。"""
    root = ET.fromstring(archive.read("xl/workbook.xml"))
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = dict((rel.get("Id"), rel.get("Target") or "") for rel in rels)
    sheets = [(node.get("name"), node.get(_qn(NS_REL, "id")))
              for node in root.iter(_qn(NS_MAIN, "sheet"))]
    if not sheets:
        raise XlsxError(u"工作簿里没有工作表")
    for name, rid in sheets:
        if sheet_name is None or name == sheet_name:
            target = targets.get(rid)
            if not target:
                raise XlsxError(u"找不到表 %r 对应的部件（r:id=%s）" % (name, rid))
            # Target 可能是 "worksheets/sheet1.xml"（相对 xl/）或 "/xl/worksheets/sheet1.xml"（绝对）
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise XlsxError(u"没有名为 %r 的表（有的表：%s）"
                    % (sheet_name, u"、".join(n for n, _ in sheets)))


def _shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    values = []
    for item in root.iter(_qn(NS_MAIN, "si")):
        # 一个 si 里可能有多个 r/t（分段样式），拼起来才是完整字符串
        values.append(u"".join(node.text or "" for node in item.iter(_qn(NS_MAIN, "t"))))
    return values


def _column_index(ref):
    """``"B12"`` → 列号 2（1 起）。只支持 A..Z 够用，但写全也无妨。"""
    letters = re.match(u"^([A-Z]+)", ref or u"")
    if not letters:
        return None
    index = 0
    for ch in letters.group(1):
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index


def _row_index(ref):
    digits = re.search(u"(\\d+)$", ref or u"")
    return int(digits.group(1)) if digits else None


def read_sheet(path, sheet_name=None):
    """读一个工作表 → ``{(行, 列): 文本}``（行列都从 1 开始）。

    只取值、不管格式。数字按 Excel 的原样写法给（`1.5` 就是 `"1.5"`），
    因为参考宏拿到的是 `CStr(cellValue)`。
    """
    if not os.path.exists(path):
        raise XlsxError(u"文件不存在：%s" % path)
    if not zipfile.is_zipfile(path):
        raise XlsxError(u"%s 不是 .xlsx（不是 zip）。老的 .xls 是二进制格式，请先另存为 .xlsx。" % path)
    with zipfile.ZipFile(path) as archive:
        try:
            part = _sheet_path(archive, sheet_name)
            root = ET.fromstring(archive.read(part))
        except KeyError as exc:
            raise XlsxError(u"%s 里读不到 %s：%s" % (path, part, exc))
        shared = _shared_strings(archive)
    cells = {}
    for cell in root.iter(_qn(NS_MAIN, "c")):
        ref = cell.get("r")
        row, col = _row_index(ref), _column_index(ref)
        if not row or not col:
            continue
        kind = cell.get("t")
        if kind == "inlineStr":
            node = cell.find(_qn(NS_MAIN, "is"))
            value = u"".join(t.text or "" for t in node.iter(_qn(NS_MAIN, "t"))) if node is not None else u""
        else:
            node = cell.find(_qn(NS_MAIN, "v"))
            raw = node.text if node is not None and node.text is not None else u""
            if kind == "s":
                try:
                    value = shared[int(raw)]
                except (ValueError, IndexError):
                    value = u""
            elif kind == "b":
                value = u"TRUE" if raw == "1" else u"FALSE"
            else:
                value = raw
        cells[(row, col)] = value
    return cells


def read_column_b(path, sheet_name=None, max_rows=None):
    """按契约取 B 列：返回 ``[B2 的值, B3 的值, …]``（下标 0 = 第 1 个变量）。

    第 1 行是表头（`B1="数值"`），所以变量 n 在 ``row n+1``。**空单元格给空串** ——
    参考宏对 `IsNull/""` 的处理就是给空串（`段落重配.bas:122-126`）。
    """
    cells = read_sheet(path, sheet_name)
    if not cells:
        return []
    last = max(row for row, _ in cells)
    if max_rows is not None:
        last = min(last, max_rows + 1)
    out = []
    for row in range(2, last + 1):
        out.append(cells.get((row, 2), u""))
    return out


# --------------------------------------------------------------------- 写
def _content_types():
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="%s">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""" % NS_CT


def _root_rels():
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""" % (NS_PKG_REL, NS_REL)


def _workbook_rels():
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="%s">
  <Relationship Id="rId1" Type="%s/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="%s/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="%s/styles" Target="styles.xml"/>
</Relationships>""" % (NS_PKG_REL, NS_REL, NS_REL, NS_REL)


def _workbook(sheet_name):
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="%s" xmlns:r="%s"><sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>""" \
        % (NS_MAIN, NS_REL, _escape(sheet_name))


def _styles():
    """最小样式表：一个默认字体 + 一个默认填充/边框 + 一个 cellXf（全部走默认）。"""
    return u"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="%s"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border/></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>""" % NS_MAIN


def _escape(text):
    # **CR 必须写成 `&#13;`**：XML 会把字面 CR 规范成 LF，而 Excel/宏写出来的单元格里
    # 段落标记就是 CR（实测用户的金标准 `数据表.xlsx`：sharedStrings 里 34 个 `&#13;`）。
    # 不这么写，"写进去的"与"读出来的"就不是同一串字符。
    text = (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;")
            .replace(u'"', u"&quot;"))
    return text.replace(u"\r", u"&#13;")


def _column_width(text, minimum=8.0, maximum=60.0):
    """列宽（Excel 的宽度单位）的近似值：中文按 2 个单位算，再加一点边距。

    参考宏用的是 Excel 的 `AutoFit`（Excel 自己按字体量）。XML 层没有"让 Excel 自己算"的写法，
    所以这里**估**一个值：能看出内容、与原宏的差别只是"不一定刚好贴合"。**这是已知差异，会写在报告里。**
    """
    width = sum(2.0 if ord(ch) > 0x2E80 else 1.0 for ch in text) + 2.0
    return max(minimum, min(maximum, width))


def write_workbook(path, sheet_name, rows, header=(u"项目", u"数值")):
    """写一个"配方用"的 xlsx：``rows`` 是 ``[(A 列文本, B 列文本), …]``（不含表头）。

    ``header=None`` 时不写表头。返回写出的路径。
    """
    sheet_name = sheet_name or u"Sheet1"
    strings = []
    index = {}

    def sid(text):
        # 单元格文本**原样保留**（含 CR）—— 与 Excel/参考宏写出来的形态一致（见 `_escape`）。
        text = u"" if text is None else u"%s" % text
        if text not in index:
            index[text] = len(strings)
            strings.append(text)
        return index[text]

    head_row = None
    body = []
    if header:
        head_row = (sid(header[0]), sid(header[1]))
    for left, right in rows:
        body.append((sid(left), sid(right)))

    def cell(ref, value_index):
        return u'<c r="%s" t="s"><v>%d</v></c>' % (ref, value_index)

    parts = []
    if head_row:
        parts.append(u'<row r="1">%s%s</row>' % (cell("A1", head_row[0]), cell("B1", head_row[1])))
    for offset, (lidx, ridx) in enumerate(body):
        row = offset + (2 if head_row else 1)
        parts.append(u'<row r="%d">%s%s</row>'
                     % (row, cell("A%d" % row, lidx), cell("B%d" % row, ridx)))
    # 列宽是按内容**估**的（Excel 的 AutoFit 由 Excel 自己按字体算，XML 里没有等价写法）。
    # 两列各取本列最长内容来估，够看即可；这条差异会写进报告。
    left_texts = [u"项目"] + [u"%s" % left for left, _ in rows]
    right_texts = [u"数值"] + [u"%s" % right for _, right in rows]
    widths = [u'<col min="1" max="1" width="%.2f" customWidth="1"/>'
              % _column_width(max(left_texts, key=lambda t: _column_width(t))),
              u'<col min="2" max="2" width="%.2f" customWidth="1"/>'
              % _column_width(max(right_texts, key=lambda t: _column_width(t)))]
    sheet = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             u'<worksheet xmlns="%s"><cols>%s</cols><sheetData>%s</sheetData></worksheet>'
             % (NS_MAIN, u"".join(widths), u"".join(parts)))
    shared = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              u'<sst xmlns="%s" count="%d" uniqueCount="%d">%s</sst>'
              % (NS_MAIN, len(strings), len(strings),
                 u"".join(u"<si><t xml:space=\"preserve\">%s</t></si>" % _escape(s)
                          for s in strings)))

    out_path = os.path.abspath(path)
    parent = os.path.dirname(out_path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    temp = out_path + ".writing"
    try:
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as archive:
            for name, text in (("[Content_Types].xml", _content_types()),
                               ("_rels/.rels", _root_rels()),
                               ("xl/workbook.xml", _workbook(sheet_name)),
                               ("xl/_rels/workbook.xml.rels", _workbook_rels()),
                               ("xl/styles.xml", _styles()),
                               ("xl/sharedStrings.xml", shared),
                               ("xl/worksheets/sheet1.xml", sheet)):
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
