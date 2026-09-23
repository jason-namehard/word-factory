# -*- coding: utf-8 -*-
"""最小可用的 DXF 读写（**只用标准库**）—— 用来跟 CAD 交换"界面布局"。

用户 2026-09-23：「我擅长使用 CAD…你生成一个 CAD 的调整 UI，我来修改（CAD 涉及比例问题，
所以你要给我一个**必须不改大小的框**，然后我在框内修改图标大小，你可以按比例计算）」。

所以这里干两件事：

* :func:`write_dxf` 写一份 **R12 ASCII DXF**（最老的、兼容性最好的那种：`POLYLINE` + `TEXT`，
  不需要 handle/owner，AutoCAD / 中望 / 浩辰 / LibreCAD 都能开），
* :func:`read_dxf` 读回来 —— 他改完保存的可能是 R2000+ 的 `LWPOLYLINE`/`MTEXT`，所以**两种都认**。

坐标就是**图纸单位**：我们约定 1 单位 = 1 像素（画布框 1440×900），这样他改完我按比例换算就行。
"""

import io
import os

#: 写出来的版本（R12 ASCII：兼容性最好、结构最简）
ACAD_VERSION = "AC1009"
#: 常用图层（颜色号按 AutoCAD 索引色：7=白/黑 1=红 3=绿 5=蓝 2=黄 8=灰）
LAYERS = {
    "FRAME": 1,        # 画布外框（不许改大小）
    "ELEMENTS": 3,     # 框内的界面元件（可拖可改）
    "PALETTE": 5,      # 框外的"零件库"（拖进来用）
    "NOTES": 8,        # 文字说明
    "GRID": 8,         # 参考网格
}


class DxfError(Exception):
    """读/写 DXF 时的用户可见错误。"""


# --------------------------------------------------------------------- 写
def _pair(code, value):
    return u"%s\n%s\n" % (code, value)


#: 中文 CAD 用 GBK（cp936）读 DXF；有少数符号它编不出来（实测：`³` 上标立方）→ 先换成安全写法。
_TEXT_SUBSTITUTES = {u"³": u"3", u"²": u"2", u"¹": u"1", u"≤": u"<=",
                     u"≥": u">=", u"→": u"->", u"·": u".", u"　": u" "}


def safe_text(text):
    """把文字里 GBK 编不出来的字符换掉 —— DXF 是**代码页**文本，写不进去就是坏文件。"""
    text = u"%s" % (text if text is not None else u"")
    for bad, good in _TEXT_SUBSTITUTES.items():
        text = text.replace(bad, good)
    out = []
    for ch in text:
        try:
            ch.encode("cp936")
            out.append(ch)
        except UnicodeEncodeError:
            out.append(u"?")
    return u"".join(out)


def write_dxf(path, shapes, layers=None, extents=None):
    """写一份最小的 R12 ASCII DXF。

    ``shapes`` 是 ``[{"kind": "rect"|"text"|"circle", ...}, …]``：

    * ``rect``：``x, y, w, h``（左下角 + 宽高）＋可选 ``layer``/``note``
    * ``text``：``x, y, text``（左下角）＋可选 ``height``/``layer``
    * ``circle``：``x, y, r``

    ``extents`` 给 ``(xmin, ymin, xmax, ymax)``，写进 HEADER 让 CAD 一打开就看到全图。
    """
    layers = dict(layers or LAYERS)
    out = []
    out.append(_pair(0, "SECTION"))
    out.append(_pair(2, "HEADER"))
    out.append(_pair(9, "$ACADVER"))
    out.append(_pair(1, ACAD_VERSION))
    if extents:
        out.append(_pair(9, "$EXTMIN"))
        out.append(_pair(10, extents[0]))
        out.append(_pair(20, extents[1]))
        out.append(_pair(9, "$EXTMAX"))
        out.append(_pair(10, extents[2]))
        out.append(_pair(20, extents[3]))
    out.append(_pair(0, "ENDSEC"))

    out.append(_pair(0, "SECTION"))
    out.append(_pair(2, "TABLES"))
    out.append(_pair(0, "TABLE"))
    out.append(_pair(2, "LAYER"))
    out.append(_pair(70, len(layers)))
    for name, color in layers.items():
        out.append(_pair(0, "LAYER"))
        out.append(_pair(2, name))
        out.append(_pair(70, 0))
        out.append(_pair(62, color))
        out.append(_pair(6, "CONTINUOUS"))
    out.append(_pair(0, "ENDTAB"))
    out.append(_pair(0, "ENDSEC"))

    out.append(_pair(0, "SECTION"))
    out.append(_pair(2, "ENTITIES"))
    for shape in shapes:
        kind = shape.get("kind")
        layer = shape.get("layer") or "ELEMENTS"
        if kind == "rect":
            x, y = shape["x"], shape["y"]
            w, h = shape["w"], shape["h"]
            points = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
            out.append(_pair(0, "POLYLINE"))
            out.append(_pair(8, layer))
            out.append(_pair(66, 1))
            out.append(_pair(70, 1))            # 1 = 闭合
            for px, py in points:
                out.append(_pair(0, "VERTEX"))
                out.append(_pair(8, layer))
                out.append(_pair(10, _num(px)))
                out.append(_pair(20, _num(py)))
            out.append(_pair(0, "SEQEND"))
            out.append(_pair(8, layer))
        elif kind == "text":
            out.append(_pair(0, "TEXT"))
            out.append(_pair(8, layer))
            out.append(_pair(10, _num(shape["x"])))
            out.append(_pair(20, _num(shape["y"])))
            out.append(_pair(40, _num(shape.get("height") or 18)))
            out.append(_pair(1, safe_text(shape.get("text"))))
        elif kind == "line":
            # 折线按点对拆成 LINE 段（R12 里 LINE 最省事；参考网格用它）
            points = [point for point in (shape.get("points") or []) if point]
            for start, end in zip(points, points[1:]):
                out.append(_pair(0, "LINE"))
                out.append(_pair(8, layer))
                out.append(_pair(10, _num(start[0])))
                out.append(_pair(20, _num(start[1])))
                out.append(_pair(11, _num(end[0])))
                out.append(_pair(21, _num(end[1])))
        elif kind == "circle":
            out.append(_pair(0, "CIRCLE"))
            out.append(_pair(8, layer))
            out.append(_pair(10, _num(shape["x"])))
            out.append(_pair(20, _num(shape["y"])))
            out.append(_pair(40, _num(shape["r"])))
        else:
            raise DxfError(u"不认识的图形：%r" % kind)
    out.append(_pair(0, "ENDSEC"))
    out.append(_pair(0, "EOF"))

    out_path = os.path.abspath(path)
    parent = os.path.dirname(out_path)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    with io.open(out_path, "w", encoding="cp936", newline="\r\n") as handle:
        handle.write(u"".join(out))
    return out_path


def _num(value):
    """DXF 里的数字：整数不要写成 20.0（老 CAD 更喜欢干净的值）。"""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return value
    if abs(number - round(number)) < 1e-9:
        return int(round(number))
    return round(number, 4)


# --------------------------------------------------------------------- 读
def read_dxf(path):
    """读 DXF 的 ENTITIES（新老两种折线都认）。返回 ``[shape, …]``。

    支持的实体：``POLYLINE``(+``VERTEX``/``SEQEND``)、``LWPOLYLINE``、``TEXT``、``MTEXT``、
    ``CIRCLE``、``LINE``、``SOLID``/``3DFACE``（后两个按四边形处理，AutoCAD 手画的框可能是它）。
    """
    if not os.path.exists(path):
        raise DxfError(u"文件不存在：%s" % path)
    pairs = _pairs(path)
    shapes = []
    index = 0
    while index < len(pairs):
        code, value = pairs[index]
        if code == 0 and value in ("POLYLINE", "LWPOLYLINE", "TEXT", "MTEXT", "CIRCLE",
                                   "LINE", "SOLID", "3DFACE"):
            entity, index = _read_entity(pairs, index)
            shapes.append(entity)
            continue
        index += 1
    return shapes


def _pairs(path):
    """把 DXF 读成 ``[(code, value), …]``；编码按 cp936/utf-8 都试（中文注释很常见）。"""
    raw = io.open(path, "rb").read()
    for encoding in ("utf-8-sig", "cp936", "utf-8"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise DxfError(u"%s 的编码读不出来（既不是 UTF-8 也不是 GBK）" % path)
    lines = text.replace(u"\r\n", u"\n").replace(u"\r", u"\n").split(u"\n")
    pairs = []
    for position in range(0, len(lines) - 1, 2):
        code = lines[position].strip()
        value = lines[position + 1].strip()
        if code == u"" and value == u"":
            continue
        try:
            pairs.append((int(code), value))
        except ValueError:
            raise DxfError(u"DXF 结构不对：第 %d 行应该是组码（数字），实际是 %r"
                           % (position + 1, code))
    return pairs


def _read_entity(pairs, start):
    kind = pairs[start][1]
    index = start + 1
    points = []
    texts = []
    layer = None
    flags = 0
    height = None
    radius = None
    while index < len(pairs):
        code, value = pairs[index]
        if code == 0:
            break
        if code == 8:
            layer = value
        elif code == 10:
            points.append([_to_float(value), None])
        elif code == 20 and points and points[-1][1] is None:
            points[-1][1] = _to_float(value)
        elif code == 11:                        # MTEXT 的第二插入点 / 文本对齐点
            points.append([_to_float(value), None])
        elif code == 21 and points and points[-1][1] is None:
            points[-1][1] = _to_float(value)
        elif code == 70:
            try:
                flags = int(value)
            except ValueError:
                flags = 0
        elif code == 40:
            if kind == "CIRCLE":
                radius = _to_float(value)
            elif height is None:
                height = _to_float(value)
        elif code == 1:
            texts.append(value)
        elif code == 3:                          # MTEXT 的长文本分片
            texts.append(value)
        index += 1
    shape = {"kind": _kind_of(kind), "layer": layer, "points": points,
             "text": u"".join(texts), "height": height, "closed": bool(flags & 1)}
    if kind == "CIRCLE":
        shape["radius"] = radius
    if kind in ("SOLID", "3DFACE"):
        shape["closed"] = True
    if kind == "POLYLINE":                       # 折线的顶点在后面的 VERTEX 里
        index = _read_vertices(pairs, index, shape)
    return shape, index


def _read_vertices(pairs, index, shape):
    """接在 POLYLINE 后面的 VERTEX…SEQEND。"""
    while index < len(pairs):
        code, value = pairs[index]
        if code != 0:
            index += 1
            continue
        if value == "VERTEX":
            index += 1
            x = y = None
            while index < len(pairs) and pairs[index][0] != 0:
                c, v = pairs[index]
                if c == 10:
                    x = _to_float(v)
                elif c == 20:
                    y = _to_float(v)
                index += 1
            if x is not None and y is not None:
                shape["points"].append([x, y])
            continue
        if value == "SEQEND":
            index += 1
            while index < len(pairs) and pairs[index][0] != 0:
                index += 1
        return index
    return index


def _kind_of(dxf_kind):
    return {"POLYLINE": "polyline", "LWPOLYLINE": "polyline", "LINE": "line",
            "TEXT": "text", "MTEXT": "text", "CIRCLE": "circle",
            "SOLID": "polygon", "3DFACE": "polygon"}[dxf_kind]


def _to_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def rect_of(shape):
    """把一个折线/四点实体变成 ``(x, y, w, h)``；不是四边形就返回 None。"""
    points = [point for point in shape.get("points", []) if point and point[1] is not None]
    if len(points) < 3:
        return None
    xs = [point[0] for point in points]
    ys = [point[1] for point in points]
    x, y = min(xs), min(ys)
    return x, y, max(xs) - x, max(ys) - y


def bounds(shapes):
    """所有图形的外包围盒；用来验"画布框没被改大小"。"""
    xs, ys = [], []
    for shape in shapes:
        for point in shape.get("points", []):
            if point and point[1] is not None:
                xs.append(point[0])
                ys.append(point[1])
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)
