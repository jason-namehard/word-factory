# -*- coding: utf-8 -*-
"""容器层（OOXML 读写）的硬性纪律：保真、幂等、拒覆盖。

这三条是"敢拿它批量处理真文档"的前提，所以它们必须是单测里最先被钉住的。
"""

import io
import os
import re
import shutil
import tempfile
import time
import unittest
import zipfile

from wordfactory.inspect import inspect, scan_part
from wordfactory.ooxml import NAMESPACES, DocxPackage, PackageError, file_sha256, qn

from . import fixtures


class PackageCase(unittest.TestCase):
    def setUp(self):
        self.work = tempfile.mkdtemp(prefix="wf_test_")
        self.src = fixtures.write_fixture(os.path.join(self.work, "fixture.docx"))

    def tearDown(self):
        shutil.rmtree(self.work, ignore_errors=True)


class TestPackageDesign(PackageCase):
    def test_package_reads_a_minimal_docx(self):
        with DocxPackage(self.src) as pkg:
            self.assertIn(DocxPackage.MAIN, pkg.part_names)
            root = pkg.xml(DocxPackage.MAIN)
            self.assertTrue(root.tag.endswith("document"))
            self.assertEqual(pkg.dirty_parts, [])

    def test_package_refuses_a_non_zip_file(self):
        bad = os.path.join(self.work, "old.doc")
        with io.open(bad, "wb") as handle:
            handle.write(b"\xd0\xcf\x11\xe0 not a zip")
        with self.assertRaises(PackageError) as caught:
            DocxPackage(bad)
        self.assertIn(".docx", u"%s" % caught.exception)

    def test_package_refuses_to_overwrite_its_own_file(self):
        with DocxPackage(self.src) as pkg:
            with self.assertRaises(PackageError):
                pkg.save(self.src)

    def test_package_copies_untouched_parts_byte_for_byte(self):
        out = os.path.join(self.work, "copy.docx")
        with DocxPackage(self.src) as pkg:
            pkg.xml(DocxPackage.MAIN)             # 先读，才能声明"改过它"
            pkg.mark_dirty(DocxPackage.MAIN)      # 假装改了正文
            pkg.save(out)
        with zipfile.ZipFile(self.src) as before, zipfile.ZipFile(out) as after:
            names_before = set(before.namelist())
            names_after = set(after.namelist())
            self.assertEqual(names_before, names_after, "保存后部件清单变了")
            for name in sorted(names_before):
                if name == DocxPackage.MAIN:
                    continue
                self.assertEqual(before.read(name), after.read(name),
                                 u"未改动部件 %s 的字节变了 —— 保真被破坏" % name)

    def test_package_keeps_the_part_list_and_total_content(self):
        out = os.path.join(self.work, "copy2.docx")
        with DocxPackage(self.src) as pkg:
            pkg.save(out)
        with zipfile.ZipFile(self.src) as before, zipfile.ZipFile(out) as after:
            self.assertEqual(before.namelist(), after.namelist())

    def test_serializing_an_unmodified_part_keeps_it_parse_equal(self):
        """读出来再写回去，内容必须等价（字节可能不同——这正说明"能不重写就不重写"）。"""
        from xml.etree import ElementTree as ET

        out = os.path.join(self.work, "copy3.docx")
        with DocxPackage(self.src) as pkg:
            original = pkg.read_bytes(DocxPackage.MAIN)
            pkg.xml(DocxPackage.MAIN)
            pkg.mark_dirty(DocxPackage.MAIN)
            reserialized = pkg.serialize(DocxPackage.MAIN)
            pkg.save(out)
        first = ET.tostring(ET.fromstring(original.decode("utf-8")), encoding="unicode")
        second = ET.tostring(ET.fromstring(reserialized.decode("utf-8")), encoding="unicode")
        self.assertEqual(first, second, "重序列化改变了文档结构")
        if original != reserialized:
            # 这不是错误，但必须**被记录**：它说明"声明改了就会重写整个部件"
            print("    note: 重序列化后字节不同（原 %d B -> 新 %d B）—— 所以只能改真的动过的部件"
                  % (len(original), len(reserialized)))
        self.assertEqual(reserialized[:5], b"<?xml", "XML 声明头丢了")


class TestInspector(PackageCase):
    def test_inspector_counts_paragraphs_tables_and_runs(self):
        info = inspect(self.src)
        story = info["stories"][0]
        self.assertEqual(story["part"], DocxPackage.MAIN)
        self.assertGreaterEqual(story["paragraphs"], 5)
        self.assertEqual(story["tables"], 1)

    def test_inspector_notices_a_paragraph_split_into_runs(self):
        """夹具里"山/亭区/的报告"是 3 个 run —— 这正是替换类宏必须处理的情形。"""
        info = inspect(self.src)
        story = info["stories"][0]
        self.assertGreaterEqual(story["paragraphs_with_multiple_runs"], 1)
        self.assertGreaterEqual(story["worst_split"], 3)

    def test_inspector_notices_highlight_color_and_superscript(self):
        info = inspect(self.src)
        story = info["stories"][0]
        self.assertEqual(story["highlighted_runs"], 1)
        self.assertEqual(story["colored_runs"], 1)
        self.assertEqual(story["super_or_subscript_runs"], 1)

    def test_inspector_is_read_only(self):
        before = file_sha256(self.src)
        inspect(self.src)
        self.assertEqual(file_sha256(self.src), before, "inspect 改了原文件")


if __name__ == "__main__":
    unittest.main()


class TestOutputIsReproducible(unittest.TestCase):
    """同一输入两次跑，输出**字节相同**。

    实测踩过：改过的部件原来用 `writestr(name, data)` 写，zip 时间戳取"写入这一刻"，
    于是两次跑出来的文件只有时间戳不同、内容完全一样 —— 声称"同一输入同一输出"就不成立了，
    还会把写入时间泄进文件。判据就是逐字节比。
    """

    def test_two_saves_are_byte_identical(self):
        work = tempfile.mkdtemp(prefix="wf_repro_")
        try:
            src = fixtures.write_fixture(os.path.join(work, "in.docx"))
            out1 = os.path.join(work, "out1.docx")
            out2 = os.path.join(work, "out2.docx")
            for out in (out1, out2):
                with DocxPackage(src) as pkg:
                    root = pkg.xml(DocxPackage.MAIN)
                    body = root.find(qn("w:body"))
                    body[0].find(qn("w:r")).find(qn("w:t")).text = u"改过的文字"
                    pkg.mark_dirty(DocxPackage.MAIN)
                    pkg.save(out)
                time.sleep(1.1)                      # 跨过一秒，让"写入时刻"必然不同
            with io.open(out1, "rb") as h1, io.open(out2, "rb") as h2:
                self.assertEqual(h1.read(), h2.read(), u"两次保存的字节必须完全相同")
        finally:
            shutil.rmtree(work, ignore_errors=True)


class TestIgnorablePrefixesSurvive(unittest.TestCase):
    """``mc:Ignorable`` 点名的前缀必须还在，否则 Word 打不开我们产出的文件。

    实测踩过（2026-09-26）：真实文档根上声明了 29 个命名空间，其中 ``wp14`` 全文一次都没用到；
    ElementTree 只给"树里真用到的"发声明，于是 ``mc:Ignorable="w14 w15 wp14"`` 里的 ``wp14``
    变成未声明前缀 —— 工具产出的每个 .docx（连只跑一步 tidy 的也是）Word 都报
    "文件可能已经损坏"，而原样复制的部件一点问题没有。修法：保存时按原件把根上的声明补齐。
    """

    WP14 = NAMESPACES["wp14"]
    MC = NAMESPACES["mc"]

    def _write(self, document_xml):
        work = tempfile.mkdtemp(prefix="wf_ignorable_")
        path = os.path.join(work, "in.docx")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", (
                u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                u'<Default Extension="xml" ContentType="application/xml"/>'
                u'<Override PartName="/word/document.xml" ContentType="application/vnd.'
                u'openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                u'</Types>').encode("utf-8"))
            archive.writestr("_rels/.rels", (
                u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                u'relationships/officeDocument" Target="word/document.xml"/></Relationships>'
            ).encode("utf-8"))
            archive.writestr("word/document.xml", document_xml.encode("utf-8"))
        return work, path

    def test_unused_prefix_declaration_is_kept(self):
        work, src = self._write(
            u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            u'<w:document xmlns:w="%s" xmlns:mc="%s" xmlns:wp14="%s" mc:Ignorable="wp14">'
            u'<w:body><w:p><w:r><w:t>文字</w:t></w:r></w:p></w:body></w:document>'
            % (NAMESPACES["w"], self.MC, self.WP14))
        try:
            out = os.path.join(work, "out.docx")
            with DocxPackage(src) as pkg:
                pkg.xml(DocxPackage.MAIN).find(qn("w:body")).find(qn("w:p")) \
                    .find(qn("w:r")).find(qn("w:t")).text = u"改过的文字"
                pkg.mark_dirty(DocxPackage.MAIN)
                pkg.save(out)
            with zipfile.ZipFile(out) as archive:
                text = archive.read("word/document.xml").decode("utf-8")
            self.assertIn('xmlns:wp14="%s"' % self.WP14, text,
                          u"根上声明过、但树里没用到的命名空间必须原样保留")
            self.assertIn('mc:Ignorable="wp14"', text)
            ignorable = re.search(r'mc:Ignorable="([^"]*)"', text).group(1)
            declared = set(re.findall(r'xmlns:([A-Za-z0-9_]+)=', text))
            missing = [p for p in ignorable.split() if p not in declared]
            self.assertEqual(missing, [], u"Ignorable 点名了未声明的前缀：%s" % missing)
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def test_serializing_without_any_change_keeps_root_declarations(self):
        work, src = self._write(
            u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            u'<w:document xmlns:w="%s" xmlns:mc="%s" xmlns:w14="%s" xmlns:wp14="%s" '
            u'mc:Ignorable="w14 wp14"><w:body><w:p><w:r><w:t>文字</w:t></w:r></w:p></w:body>'
            u'</w:document>' % (NAMESPACES["w"], self.MC, NAMESPACES["w14"], self.WP14))
        try:
            out = os.path.join(work, "out.docx")
            with DocxPackage(src) as pkg:
                pkg.xml(DocxPackage.MAIN)             # 先读一遍（一个字不改）
                pkg.mark_dirty(DocxPackage.MAIN)      # 也要能保存
                pkg.save(out)
            with zipfile.ZipFile(out) as archive:
                text = archive.read("word/document.xml").decode("utf-8")
            for prefix, uri in (("w14", NAMESPACES["w14"]), ("wp14", self.WP14)):
                self.assertIn('xmlns:%s="%s"' % (prefix, uri), text)
        finally:
            shutil.rmtree(work, ignore_errors=True)
