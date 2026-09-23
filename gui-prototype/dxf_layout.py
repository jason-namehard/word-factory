# -*- coding: utf-8 -*-
"""CAD（DXF）版的界面布局：**生成模板给他改，改完读回来变成界面。**

用户 2026-09-23：「我擅长使用 CAD…你生成一个 CAD 的调整 UI，我来修改（CAD 涉及比例问题，
所以你要给我一个**必须不改大小的框**，然后我在框内修改图标大小，你可以按比例计算，
比如给我一个画布矩形，画布矩形就是最大的软件外围边框）」。

约定（写进 DXF 的说明图层里，他一看就知道）：

* **画布框 = 1440 × 900 图纸单位，1 单位 = 1 像素** —— 就是软件窗口的外框，**不要改它的大小**；
* 框内随便画/拖动矩形当界面元件，**矩形里写文字**就是那个元件的名字（我按文字认元件）；
* 框外右下角有个**零件库**：想用什么就从那儿复制一个矩形拖进框里；
* 网格是参考线（100 单位一格），不用管它。

用法::

    python dxf_layout.py write  输出.dxf            # 生成模板（给他用 CAD 改）
    python dxf_layout.py read   他改过的.dxf 输出.json  # 读回来 → 与网页版同一种布局 JSON
"""

import io
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from wordfactory import dxf  # noqa: E402

#: 画布框（软件窗口外框）：1440×900，1 单位 = 1 像素
CANVAS_W, CANVAS_H = 1440, 900
#: 不算"界面元件"的图层（框、零件库、说明、网格）
NON_ELEMENT_LAYERS = ("FRAME", "PALETTE", "NOTES", "GRID")

#: 零件库里的模板（名字 → 宽高）；文字就是元件名，我按关键字认类型
PALETTE = [
    (u"侧栏", 240, 420), (u"主工作区", 600, 300), (u"运行日志", 600, 120),
    (u"标签页", 120, 34), (u"按钮", 120, 34), (u"复选框", 120, 28),
    (u"下拉选择", 160, 34), (u"清单", 240, 180), (u"模板选择器", 220, 120),
    (u"表格清单", 260, 170), (u"预览窗", 220, 150),
    (u"运行（dry-run）", 130, 34), (u"出验证版", 130, 34), (u"出正式版", 130, 34),
    (u"体检 AUDIT", 130, 34),
]

#: 文字 → 元件类型（我读回时的判定表；认不出来就当普通分区）
TYPE_BY_KEYWORD = [
    (u"侧栏", "zone"), (u"主工作区", "zone"), (u"分区", "zone"), (u"面板", "zone"),
    (u"日志", "log"), (u"表格清单", "table-list"), (u"清单", "list"),
    (u"模板选择", "style-picker"), (u"预览", "preview"),
    (u"标签", "tab"), (u"复选", "checkbox"), (u"下拉", "select"),
    (u"运行", "action"), (u"验证版", "action"), (u"正式版", "action"), (u"体检", "action"),
    (u"按钮", "button"),
]


def type_of(label):
    for keyword, kind in TYPE_BY_KEYWORD:
        if keyword in (label or u""):
            return kind
    return "zone"


# --------------------------------------------------------------------- 生成模板
def build_shapes():
    shapes = []
    shapes.append({"kind": "rect", "layer": "FRAME", "x": 0, "y": 0,
                   "w": CANVAS_W, "h": CANVAS_H})
    shapes.append({"kind": "text", "layer": "FRAME", "x": 0, "y": CANVAS_H + 20,
                   "height": 22,
                   "text": u"① 画布框 %d×%d（1 单位 = 1 像素）—— 这就是软件窗口外框，"
                           u"**这个框不要改大小**；框内随便画" % (CANVAS_W, CANVAS_H)})
    shapes.append({"kind": "text", "layer": "FRAME", "x": 0, "y": CANVAS_H + 50,
                   "height": 20,
                   "text": u"② 框内每个矩形 = 一个界面元件；矩形里写的字就是它的名字"
                           u"（如「按钮」「表格清单」「预览窗」「侧栏」）"})

    # 参考网格（100 单位一格）
    for x in range(0, CANVAS_W + 1, 100):
        shapes.append({"kind": "line", "layer": "GRID", "points": [[x, 0], [x, CANVAS_H]]})
    for y in range(0, CANVAS_H + 1, 100):
        shapes.append({"kind": "line", "layer": "GRID", "points": [[0, y], [CANVAS_W, y]]})

    # 初始布局（和网页版的"左栏 + 主区"预设一致，他接着改就行）
    initial = [
        (u"侧栏：宏库 / 表格 / 规则 / 体检", 20, 60, 200, 430),
        (u"主工作区", 240, 60, 540, 220),
        (u"清单（勾选 + 拖排序）", 260, 240, 240, 180),
        (u"预览窗", 520, 240, 240, 180),
        (u"运行日志", 240, 440, 540, 90),
        (u"运行（dry-run）", 800, 70, 130, 34),
        (u"出验证版", 800, 120, 130, 34),
        (u"出正式版", 800, 170, 130, 34),
        (u"体检 AUDIT", 800, 220, 130, 34),
    ]
    for text, x, y, w, h in initial:
        shapes.append({"kind": "rect", "layer": "ELEMENTS", "x": x, "y": y, "w": w, "h": h})
        shapes.append({"kind": "text", "layer": "ELEMENTS", "x": x + 10, "y": y + h - 26,
                       "height": 16, "text": text})

    # 零件库（放在画布框右边，不参与识别）
    base_x = CANVAS_W + 120
    shapes.append({"kind": "text", "layer": "PALETTE", "x": base_x, "y": CANVAS_H + 20,
                   "height": 22, "text": u"③ 零件库：要用哪个就把它的矩形**复制**到画布框里"})
    y = CANVAS_H - 60
    for name, w, h in PALETTE:
        shapes.append({"kind": "rect", "layer": "PALETTE", "x": base_x, "y": y, "w": w, "h": h})
        shapes.append({"kind": "text", "layer": "PALETTE", "x": base_x + 8, "y": y + h - 24,
                       "height": 16, "text": name})
        y -= h + 26
        if y < -400:
            y = CANVAS_H - 60
            base_x += 320

    # 表格模板名（给"模板选择器"元件当内容参考）
    shapes.append({"kind": "text", "layer": "PALETTE", "x": base_x, "y": -120, "height": 18,
                   "text": u"③.2 表格模板（供「模板选择器」用）：通用款·外粗内细 / 通用款·全居中 / "
                           u"通用款·均布列宽 / 通用款·自适应列宽 / 三线表·学术款 / 三线表·学术款·均布列宽"})
    shapes.append({"kind": "text", "layer": "PALETTE", "x": base_x, "y": -160, "height": 18,
                   "text": u"③.3 表格清单里勾选的行（按表头）：水位（m）/ 库容（万m³） ｜ "
                           u"序号 / 高程 / 面积 ｜ 项目 / 特性指标 / 单位"})

    # 说明
    notes = [
        u"④ 改完保存成 DXF（R12/R2000 都行），把文件给我；我读回来自动生成界面代码。",
        u"⑤ 只要矩形**在画布框内**就算界面元件；框外的零件库和这堆说明我不认。",
        u"⑥ 想要什么新元件，就在框内画个矩形、里面写字（认不出来我当普通分区处理）。",
        u"⑦ 比例：1 单位 = 1 像素。就算你把整张图缩放，我也会按画布框的实际宽度等比换算。",
        u"⑧ 画布框请保持 1440:900 的比例、别改——它是「最大软件外围边框」，我按它对齐。",
    ]
    y = -40
    for note in notes:
        shapes.append({"kind": "text", "layer": "NOTES", "x": 0, "y": y, "height": 18, "text": note})
        y -= 28
    return shapes


def write_template(path):
    shapes = build_shapes()
    return dxf.write_dxf(path, shapes, extents=(-200, -220, CANVAS_W + 1500, CANVAS_H + 120))


# --------------------------------------------------------------------- 读回 → 布局
def layout_from_dxf(path):
    """把他改过的 DXF 读成布局 JSON（与网页版导出的 `wordfactory-gui-layout.json` 同一种）。"""
    shapes = dxf.read_dxf(path)
    frame = None
    for shape in shapes:
        if (shape.get("layer") or "").upper() == "FRAME":
            rect = dxf.rect_of(shape)
            if rect and rect[2] > 200 and rect[3] > 200:
                frame = rect
                break
    if frame is None:                              # 没找到画布框就拿最大的矩形当框
        rects = [dxf.rect_of(s) for s in shapes if s.get("kind") in ("polyline", "polygon")]
        rects = [r for r in rects if r]
        if not rects:
            raise dxf.DxfError(u"这份 DXF 里没找到任何矩形 —— 画布框和界面元件都是矩形")
        frame = max(rects, key=lambda r: r[2] * r[3])
    fx, fy, fw, fh = frame
    scale = CANVAS_W / fw if fw else 1.0           # 他要是缩放过，就等比换算

    texts = []
    for shape in shapes:
        if shape.get("kind") == "text" and shape.get("text"):
            point = shape["points"][0] if shape.get("points") else None
            if point and point[1] is not None:
                texts.append((point[0], point[1], shape["text"]))

    elements = []
    notes = []
    for shape in shapes:
        layer = (shape.get("layer") or "").upper()
        if layer in NON_ELEMENT_LAYERS:
            if layer == "NOTES" and shape.get("kind") == "text" and shape.get("text"):
                notes.append(shape["text"])
            continue
        rect = dxf.rect_of(shape)
        if not rect:
            continue
        x, y, w, h = rect
        cx, cy = x + w / 2.0, y + h / 2.0
        if not (fx <= cx <= fx + fw and fy <= cy <= fy + fh):
            continue                               # 框外的不算（零件库）
        label = u""
        for tx, ty, text in texts:
            if x - 5 <= tx <= x + w and y - 5 <= ty <= y + h:
                if len(text) > len(label):
                    label = text
        elements.append({
            "id": "e%d" % (len(elements) + 1),
            "type": type_of(label),
            "x": round((x - fx) * scale), "y": round((fy + fh - (y + h)) * scale),
            "w": round(w * scale), "h": round(h * scale),
            "text": label or u"（未命名）",
        })
    return {
        "schema": 1,
        "source": os.path.basename(path),
        "frame": {"w": CANVAS_W, "h": CANVAS_H},
        "dxf_frame": {"w": round(fw), "h": round(fh), "scale": round(scale, 4)},
        "elements": elements,
        "notes": notes,
    }


def main(argv):
    if len(argv) < 3:
        print(__doc__)
        return 2
    action, path = argv[1], argv[2]
    if action == "write":
        out = write_template(path)
        print(u"已写出布局模板：%s" % out)
        print(u"  画布框 %d×%d ｜ 图层：FRAME(框) ELEMENTS(元件) PALETTE(零件库) NOTES(说明) GRID(网格)"
              % (CANVAS_W, CANVAS_H))
        return 0
    if action == "read":
        layout = layout_from_dxf(path)
        out = argv[3] if len(argv) > 3 else os.path.splitext(path)[0] + u".json"
        with io.open(out, "w", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(layout, ensure_ascii=False, indent=2) + u"\n")
        print(u"已读回：%s → %s" % (path, out))
        print(u"  画布框 %s ｜ 元件 %d 个 ｜ 说明 %d 条"
              % (layout["dxf_frame"], len(layout["elements"]), len(layout["notes"])))
        for item in layout["elements"][:20]:
            print(u"    %-4s %-14s @(%4d,%4d) %4d×%-4d %s"
                  % (item["id"], item["type"], item["x"], item["y"], item["w"], item["h"],
                     item["text"][:26]))
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
