# -*- coding: utf-8 -*-
"""页码探测里"网络外链"那套处理（2026-10-08 用户报"页码全不对、特别慢"的根因）。

那份报告的图表链到别人机器上的共享：
``word/charts/_rels/chart2.xml.rels -> file:///\\\\Bf-230206\\…\\欧峪水库.xlsm``。
Word 打开时会**一直等 SMB 超时**（实测 >40 秒没返回），于是"读不到真实页码 → 退回估算口径
→ 首页/目录旁那两只空白页删不掉"。把这类外链在**临时副本**里改指向本地不存在的文件之后，
Word 7.9 秒就开好了，页码与 WPS 完全一致（31 页）——因为图是拿缓存数据画的，版式没变。
"""

import os
import shutil
import tempfile
import unittest
import zipfile

from wordfactory import doclinks, pageprobe

from . import fixtures

CHART_RELS = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    u'relationships/oleObject" Target="file:///\\\\Bf-230206\\2025共享文件\\欧峪水库.xlsm" '
    u'TargetMode="External"/></Relationships>')

IMAGE_RELS_LOCAL = (
    u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    u'relationships/image" Target="media/image1.png"/></Relationships>')


class NetworkLinkCase(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="wf_probe_")
        self.path = os.path.join(self.dir, u"夹具.docx")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def build(self, extra):
        fixtures.write_fixture(self.path, body=fixtures.paragraph(fixtures.run(u"正文")),
                               extra=extra)
        return self.path

    def test_network_ole_link_is_detected(self):
        self.build({u"word/charts/_rels/chart1.xml.rels": CHART_RELS})
        self.assertEqual(doclinks.network_ole_links(self.path), 1)

    def test_local_rel_is_not_a_network_link(self):
        self.build({u"word/_rels/x.xml.rels": IMAGE_RELS_LOCAL})
        self.assertEqual(doclinks.network_ole_links(self.path), 0)

    def test_plain_document_has_no_network_links(self):
        self.build({})
        self.assertEqual(doclinks.network_ole_links(self.path), 0)

    def test_neutralized_copy_rewrites_only_the_network_ole_link(self):
        self.build({u"word/charts/_rels/chart1.xml.rels": CHART_RELS,
                    u"word/_rels/x.xml.rels": IMAGE_RELS_LOCAL})
        copy_path = doclinks.neutralized_copy(self.path)
        try:
            self.assertIsNotNone(copy_path, u"有网络外链时才该生成副本")
            self.assertEqual(doclinks.network_ole_links(copy_path), 0,
                             u"副本里不能再有网络外链（否则 Word 还是会去等网络）")
            with zipfile.ZipFile(copy_path) as archive:
                text = archive.read("word/_rels/x.xml.rels").decode("utf-8")
                self.assertIn(u"media/image1.png", text, u"本地图片链接一个字都不许动")
                self.assertEqual(sorted(archive.namelist()),
                                 sorted(zipfile.ZipFile(self.path).namelist()),
                                 u"副本除链接目标外不该增删部件")
        finally:
            if copy_path:
                os.remove(copy_path)

    def test_no_copy_is_made_when_there_is_nothing_to_fix(self):
        self.build({})
        self.assertIsNone(doclinks.neutralized_copy(self.path),
                          u"没有网络外链就别多此一举造副本")

    def test_candidate_order_prefers_wps_for_linked_documents(self):
        original = list(pageprobe._PREFERRED)
        pageprobe._PREFERRED[:] = []
        try:
            order = pageprobe._candidates(None, network_links=1)
            self.assertEqual(order[0], u"KWPS.Application",
                             u"有网络外链时先试 WPS（Word 会干等网络）")
            self.assertIn(u"Word.Application", order, u"WPS 不行还得能退回 Word")
        finally:
            pageprobe._PREFERRED[:] = original

    def test_last_good_renderer_goes_first(self):
        original = list(pageprobe._PREFERRED)
        pageprobe._PREFERRED[:] = [u"WPS.Application"]
        try:
            self.assertEqual(pageprobe._candidates(None, network_links=0)[0],
                             u"WPS.Application", u"上次成功的那个先试，别每次撞墙")
        finally:
            pageprobe._PREFERRED[:] = original

    def test_explicit_renderer_wins(self):
        self.assertEqual(pageprobe._candidates(u"Word.Application", network_links=1),
                         [u"Word.Application"])


if __name__ == "__main__":
    unittest.main()
