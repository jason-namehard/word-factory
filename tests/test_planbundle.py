import json
import os
import tempfile
import unittest
from pathlib import Path

from wordfactory import planbundle, pipeline, fonts
from wordfactory.document import Document
from wordfactory.text import Paragraph
from wordfactory.ooxml import qn
from . import fixtures


class PortablePlanCase(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.root=Path(self.tmp.name)
        (self.root/'replace-rules').mkdir()
        self.replacement={'name':'训练好的替换','text':[{'find':'##其它','replace':'替换已执行'}],'font':[],'para':[]}
        self.subs={'schema':1,'rules':[{'id':'m2','match':'m2','kinds':'NS'}]}
        self.styles={'schema':1,'styles':{'我的款式':{'column_widths':'equal','header':{'bold':True}}}}
        for name,data in [('replace-rules/训练好的替换.json',self.replacement),('subscripts.json',self.subs),('tablestyle.json',self.styles),('fonts.json',fonts.DEFAULT_FONTS)]:
            (self.root/name).write_text(json.dumps(data,ensure_ascii=False),encoding='utf-8')
    def tearDown(self): self.tmp.cleanup()

    def test_export_carries_rules_and_preserves_order_parameters(self):
        plan={'name':'我的方案','note':'需要带走','steps':[
            {'op':'replace','label':'训练替换','params':{'rules':'训练好的替换','scope':'body'}},
            {'op':'mdclean','params':{}},{'op':'sup','params':{}},
            {'op':'tablestyle','params':{'style':'我的款式','uniform':True}}]}
        bundle=planbundle.export(plan,str(self.root))
        self.assertEqual([s['op'] for s in bundle['steps']],['replace','mdclean','sup','tablestyle'])
        self.assertEqual(bundle['steps'][0]['params']['rules_data'],self.replacement)
        self.assertEqual(bundle['steps'][2]['params']['rules_data'],self.subs)
        self.assertEqual(bundle['steps'][3]['params']['styles_data'],self.styles)
        self.assertEqual(bundle['steps'][0]['params']['scope'],'body')
        self.assertEqual(bundle['font_rules_data'],fonts.DEFAULT_FONTS)
        self.assertNotIn('rules_data',plan['steps'][0]['params'])

    def test_imported_plan_runs_after_original_rules_are_unavailable(self):
        bundle=planbundle.export({'name':'顺序验证','steps':[{'op':'replace','params':{'rules':'训练好的替换'}},{'op':'mdclean'}]},str(self.root))
        (self.root/'replace-rules').rename(self.root/'原规则已移走')
        imported=planbundle.import_plan(json.loads(json.dumps(bundle)))
        source=self.root/'输入.docx';output=self.root/'输出.docx'
        fixtures.write_fixture(str(source),body=fixtures.paragraph(fixtures.run('##其它')))
        report=pipeline.run_pipeline(str(source),imported['steps'],out_path=str(output),mode='formal')
        self.assertEqual([s['op'] for s in report['steps']],['replace','mdclean'])
        with Document(str(output)) as doc:
            text=''.join(Paragraph(p).text for p in doc.part().iter(qn('w:p')))
        self.assertEqual(text,'替换已执行')

    def test_missing_replacement_content_cannot_be_exported_as_a_broken_reference(self):
        with self.assertRaises(pipeline.PipelineError):
            planbundle.export({'steps':[{'op':'replace','params':{'rules':'不存在'}}]},str(self.root))

    def test_legacy_plan_json_is_still_accepted(self):
        plan=planbundle.import_plan({'name':'旧方案','steps':['tidy','captions']})
        self.assertEqual(plan['steps'],['tidy','captions'])

    def test_inline_sup_dictionary_is_used_instead_of_machine_defaults(self):
        source=self.root/'输入.docx';output=self.root/'输出.docx'
        fixtures.write_fixture(str(source),body=fixtures.paragraph(fixtures.run('Z9')))
        data={'rules':[{'id':'custom','match':'Z9','kinds':'NS'}]}
        report=pipeline.run_pipeline(str(source),[{'op':'sup','params':{'rules_data':data}}],out_path=str(output),mode='formal')
        self.assertEqual(report['steps'][0]['report']['total'],1)


class FormalPropertiesCase(unittest.TestCase):
    def test_red_paragraph_mark_and_empty_run_are_normalized_and_audited(self):
        with tempfile.TemporaryDirectory() as tmp:
            source=os.path.join(tmp,'输入.docx');output=os.path.join(tmp,'正式.docx')
            body='<w:p><w:pPr><w:rPr><w:color w:val="FF0000"/></w:rPr></w:pPr>'+fixtures.run('正文')+'<w:r><w:rPr><w:color w:val="FF0000"/></w:rPr></w:r></w:p>'
            fixtures.write_fixture(source,body=body)
            report=pipeline.run_pipeline(source,[],mode='formal',out_path=output,force_write=True)
            self.assertEqual(report['audit']['verdict'],'PASS')

    def test_pdf_can_force_an_output_even_when_preview_requires_no_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            source=os.path.join(tmp,'输入.docx');output=os.path.join(tmp,'预览.docx')
            fixtures.write_fixture(source,body=fixtures.paragraph(fixtures.run('没有要改的正文')))
            report=pipeline.run_pipeline(source,[],mode='verify',out_path=output,force_write=True)
            self.assertEqual(report['out'],output)
            self.assertTrue(os.path.isfile(output))
