# -*- coding: utf-8 -*-
"""DXF 读写的单测（给 CAD 版界面布局用，`wordfactory/dxf.py`）。

要钉住的：

* **往返**：写出去再读回来，矩形/文字/图层/坐标都对得上；
* **读得进他 CAD 存出来的格式**：AutoCAD 一般存 R2000+，折线是 `LWPOLYLINE`（不是 R12 的
  `POLYLINE`+`VERTEX`），文字可能是 `MTEXT` —— 两种都得认（用户是在 CAD 里改我的图）；
* **中文**：DXF 是代码页文本，`³` 这种 GBK 编不出来的字符要先换成安全写法（不然写文件直接抛异常）；
* 说明图层、零件库图层不能被当成界面元件（那是给用户看的，不是界面）。
"""

import io
import os
import shutil
import tempfile
import unittest

from wordfactory import dxf


class DxfCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_dxf_")
        self.path = os.path.join(self.dir, u"布局.dxf")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)


class TestRoundTrip(DxfCase):
    def test_rect_text_and_layers_survive(self):
        shapes = [
            {"kind": "rect", "layer": "FRAME", "x": 0, "y": 0, "w": 1440, "h": 900},
            {"kind": "rect", "layer": "ELEMENTS", "x": 20, "y": 60, "w": 200, "h": 430},
            {"kind": "text", "layer": "ELEMENTS", "x": 30, "y": 460, "text": u"侧栏",
             "height": 16},
            {"kind": "line", "layer": "GRID", "points": [[0, 0], [1440, 0]]},
            {"kind": "circle", "layer": "ELEMENTS", "x": 100, "y": 100, "r": 20},
        ]
        dxf.write_dxf(self.path, shapes)
        back = dxf.read_dxf(self.path)
        kinds = [shape["kind"] for shape in back]
        self.assertEqual(kinds, ["polyline", "polyline", "text", "line", "circle"])
        rects = [dxf.rect_of(shape) for shape in back if shape["kind"] == "polyline"]
        self.assertEqual(rects[0], (0.0, 0.0, 1440.0, 900.0))
        self.assertEqual(rects[1], (20.0, 60.0, 200.0, 430.0))
        self.assertEqual(back[2]["text"], u"侧栏")
        self.assertEqual(back[2]["layer"], "ELEMENTS")
        self.assertEqual(back[3]["points"], [[0.0, 0.0], [1440.0, 0.0]])

    def test_the_file_is_a_readable_text_dxf(self):
        dxf.write_dxf(self.path, [{"kind": "rect", "x": 0, "y": 0, "w": 10, "h": 10}])
        raw = io.open(self.path, "rb").read()
        self.assertIn(b"AC1009", raw, u"版本号要在")
        self.assertIn(b"ENTITIES", raw)
        self.assertIn(b"\r\n", raw, u"DXF 用 CRLF 换行")
        self.assertTrue(raw.rstrip().endswith(b"EOF"))

    def test_bounds_reports_the_extents(self):
        shapes = [{"kind": "rect", "x": 0, "y": 0, "w": 100, "h": 50},
                  {"kind": "rect", "x": 200, "y": -30, "w": 10, "h": 10}]
        dxf.write_dxf(self.path, shapes)
        self.assertEqual(dxf.bounds(dxf.read_dxf(self.path)), (0.0, -30.0, 210.0, 50.0))


class TestReadsWhatCadSaves(DxfCase):
    """AutoCAD 存出来的通常不是 R12 —— 这里喂它真实的写法。"""

    LWPOLYLINE_DXF = u"""0
SECTION
2
ENTITIES
0
LWPOLYLINE
8
ELEMENTS
90
4
70
1
10
20.0
20
60.0
10
220.0
20
60.0
10
220.0
20
490.0
10
20.0
20
490.0
0
TEXT
8
ELEMENTS
10
30.0
20
470.0
40
16.0
1
侧栏
0
MTEXT
8
NOTES
10
0.0
20
-40.0
40
18.0
1
第一段说明
3
第二段说明
0
ENDSEC
0
EOF
"""

    def write_raw(self, text):
        with io.open(self.path, "w", encoding="cp936", newline="\r\n") as handle:
            handle.write(text)
        return self.path

    def test_lwpolyline_and_mtext_are_read(self):
        self.write_raw(self.LWPOLYLINE_DXF)
        shapes = dxf.read_dxf(self.path)
        self.assertEqual([shape["kind"] for shape in shapes], ["polyline", "text", "text"])
        self.assertEqual(dxf.rect_of(shapes[0]), (20.0, 60.0, 200.0, 430.0))
        self.assertEqual(shapes[1]["text"], u"侧栏")
        self.assertEqual(shapes[2]["text"], u"第一段说明第二段说明", u"MTEXT 的 3 码分片要拼起来")

    def test_a_utf8_file_is_read_too(self):
        with io.open(self.path, "w", encoding="utf-8", newline="\r\n") as handle:
            handle.write(self.LWPOLYLINE_DXF)
        shapes = dxf.read_dxf(self.path)
        self.assertEqual(shapes[1]["text"], u"侧栏", u"UTF-8 存的也认（他可能用别的工具）")

    def test_a_broken_file_says_so(self):
        with io.open(self.path, "w", encoding="utf-8") as handle:
            handle.write(u"这不是 DXF\n随便写点\n")
        with self.assertRaises(dxf.DxfError):
            dxf.read_dxf(self.path)


class TestChineseSafety(DxfCase):
    def test_gbk_unencodable_characters_are_replaced(self):
        self.assertEqual(dxf.safe_text(u"库容（万m³）"), u"库容（万m3）")
        self.assertEqual(dxf.safe_text(u"a≤b≥c→d"), u"a<=b>=c->d")
        self.assertEqual(dxf.safe_text(u"表情😀掉不了"), u"表情?掉不了")

    def test_writing_a_superscript_does_not_crash(self):
        dxf.write_dxf(self.path, [{"kind": "text", "x": 0, "y": 0, "text": u"万m³"}])
        shapes = dxf.read_dxf(self.path)
        self.assertEqual(shapes[0]["text"], u"万m3", u"写之前就换掉了，文件仍然可读")


if __name__ == "__main__":
    unittest.main()
