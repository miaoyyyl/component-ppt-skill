"""制作流程的约束回归：结构锁定、来源、缺图与原字号强调规则。"""
import copy
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from lxml import etree as E
from PIL import Image
from deck_assets import collect_assets
from deck_native import fill_text,fill_component,native_targets,form_info,read_package,slide_part,xml,NS,tag
from deck_pipeline import validate_render_plan,reflow_toc,optional_regions
from production_tasks import ROOT,read_json,write_json
from progress_rings import CHART_NS,related_part
from suite_toc import suite_number


class PipelineChecks(unittest.TestCase):
    def setUp(self):
        self.docs=[{'blocks':[{'id':'D1:p1'}],'assets':[{'id':'D1:image1'}]}]
        self.analysis={'sources':[]}
        self.manifest={'pages':[{'pageNumber':1,'componentId':'BW-TRIPLE-COMMON-01','variantId':'all-text',
            'optionalRegions':[{'id':'column-links','shapeIds':[40,41,42,43]}],
            'targets':[{'id':'body:0','kind':'text'},{'id':'figure:0','kind':'image'}]}]}
        self.plan={'pages':[{'pageNumber':1,'componentId':'BW-TRIPLE-COMMON-01','variantId':'all-text','title':'测试',
            'hiddenRegionIds':['column-links'],'texts':[{'targetId':'body:0','paragraphs':[{'text':'测试内容','emphasisPhrases':['内容']}],'sourceRefs':['D1:p1']}],
            'images':[{'targetId':'figure:0','mode':'word','assetId':'D1:image1','sourceRef':'D1:image1','url':'','description':'真实原图'}],'notes':''}],'warnings':[]}

    def validate(self):return validate_render_plan(self.plan,self.manifest,self.analysis,self.docs)

    def test_only_registered_optional_links_can_hide(self):
        self.validate();self.plan['pages'][0]['hiddenRegionIds']=['shared-links']
        with self.assertRaises(ValueError):self.validate()

    def test_parallel_variant_has_no_deleted_optional_links(self):
        catalog=read_json(ROOT/'library/components.json');inventory=read_json(ROOT/'library/inventory.json')
        component=next(c for c in catalog['components'] if c['id']=='BW-TRIPLE-PROGRESS-01')
        base,_=form_info(component,'base',inventory);parallel,_=form_info(component,'parallel',inventory)
        self.assertEqual(len(optional_regions(component,base)),1);self.assertEqual(optional_regions(component,parallel),[])

    def test_variant_page_order_and_targets_are_locked(self):
        for mutate in [lambda p:p.update(variantId='base'),lambda p:p.update(pageNumber=2),lambda p:p['texts'].clear(),lambda p:p['texts'].append(copy.deepcopy(p['texts'][0]))]:
            original=copy.deepcopy(self.plan);mutate(self.plan['pages'][0])
            with self.assertRaises(ValueError):self.validate()
            self.plan=original

    def test_image_and_emphasis_need_existing_evidence(self):
        self.plan['pages'][0]['images'][0]['assetId']='D9:image9'
        with self.assertRaises(ValueError):self.validate()
        self.plan['pages'][0]['images'][0]['assetId']='D1:image1';self.plan['pages'][0]['texts'][0]['paragraphs'][0]['emphasisPhrases']=['不存在的词']
        with self.assertRaises(ValueError):self.validate()

    def test_display_copy_rejects_citations_and_sentence_endings(self):
        paragraph=self.plan['pages'][0]['texts'][0]['paragraphs'][0];paragraph['emphasisPhrases']=[]
        for content in ['土壤Cd为2.58 mg/kg [1–2]','筛选有效材料。']:
            paragraph['text']=content
            with self.assertRaises(ValueError):self.validate()
        paragraph['text']='土壤Cd：2.58 → 1.53 mg/kg'
        self.plan['pages'][0]['notes']='文献[1]及试验条件，完整说明。'
        self.validate()

    def test_page_titles_are_declarative(self):
        self.manifest['pages'][0]['targets'][0]['id']='pageTitle:0'
        text=self.plan['pages'][0]['texts'][0];text['targetId']='pageTitle:0';text['paragraphs']=[{'text':'哪些因素控制镉迁移','emphasisPhrases':[]}]
        with self.assertRaises(ValueError):self.validate()
        text['paragraphs'][0]['text']='控制镉迁移的因素';self.validate()

    def test_missing_asset_stays_missing_and_placeholder_decodes(self):
        qa=ROOT/'_qa';qa.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=qa) as folder:
            image=self.plan['pages'][0]['images'][0];image['mode']='missing';image['assetId']=''
            assets,records,missing=collect_assets(Path(folder),[],self.plan)
            self.assertEqual(len(missing),1);self.assertEqual(records,[])
            with Image.open(next(iter(assets.values()))) as picture:
                picture.load();self.assertEqual(picture.size,(1,1))

    def test_new_emphasis_preserves_size_and_blue_background_contrast(self):
        theme=read_json(ROOT/'library/theme.json')
        for color,expected in [('FFFFFF','FFFF00'),('0046A5','C00000')]:
            node=E.Element(tag('p','sp'));body=E.SubElement(node,tag('p','txBody'));E.SubElement(body,tag('a','bodyPr'));E.SubElement(body,tag('a','p'))
            profile={'paragraphs':[{'sizePt':22,'isBlank':False,'bold':False,'italic':False,'color':color}]}
            fill_text(node,[{'text':'本次重点内容','emphasisPhrases':['重点']}],profile,theme)
            runs=body.findall('a:p/a:r',NS)
            self.assertTrue(all(r.find('a:rPr',NS).get('sz')=='2200' for r in runs))
            highlighted=next(r for r in runs if r.find('a:t',NS).text=='重点')
            self.assertEqual(highlighted.find('a:rPr/a:solidFill/a:srgbClr',NS).get('val'),expected)
            self.assertIsNotNone(body.find('a:bodyPr/a:noAutofit',NS))

    def test_dynamic_toc_keeps_number_color_and_equal_spacing(self):
        catalog=read_json(ROOT/'library/components.json');inventory=read_json(ROOT/'library/inventory.json');theme=read_json(ROOT/'library/theme.json')
        qa=ROOT/'_qa';qa.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=qa) as folder:
            for component in [c for c in catalog['components'] if c.get('pageRole')=='toc']:
                with self.subTest(component=component['id']):
                    detail,slots=form_info(component,'base',inventory)
                    layout=component.get('tocLayout')
                    number_style=layout['numberStyle'] if layout else 'two-digit'
                    template=Path(folder)/'toc.pptx';output=Path(folder)/'filled.pptx'
                    rows=reflow_toc(ROOT/'library'/detail['templatePath'],['第一部分','第二部分','第三部分'],1,theme,template,layout=layout)
                    self.assertAlmostEqual(rows[1]['y']-rows[0]['y'],rows[2]['y']-rows[1]['y'])
                    targets=native_targets(template,slots);profile={'shapes':{}}
                    texts=[]
                    for target in targets:
                        color='0046A5' if 20010<=target['shapeId']<20020 else 'D9D9D9'
                        profile['shapes'][str(target['shapeId'])]={'paragraphs':[{'sizePt':28,'color':color,'bold':True,'italic':False,'isBlank':p['isBlank']} for p in target['paragraphs']]}
                        texts.append({'targetId':target['id'],'paragraphs':[{'text':'错误99','emphasisPhrases':['错误']}],'sourceRefs':[]})
                    fill_component(template,targets,profile,{'title':'目录测试','texts':texts,'images':[]},{},theme,output)
                    data=read_package(output);part,_=slide_part(data,1);root=xml(data[part])
                    for target in targets:
                        if target['slotKey']!='chapterNumbers':continue
                        node=root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr/@id='{target['shapeId']}']",namespaces=NS)[0]
                        text=''.join(node.xpath(".//a:t/text()",namespaces=NS))
                        self.assertEqual(text,suite_number(target['index'],number_style))
                        self.assertNotIn('C00000',[n.get('val') for n in node.findall('.//a:srgbClr',NS)])
                    for ident in component.get('fixedTextShapeIds',[]):
                        text=''.join(root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr/@id='{ident}']//a:t/text()",namespaces=NS))
                        self.assertTrue(text)

    def test_progress_percent_updates_chart_and_workbook(self):
        catalog=read_json(ROOT/'library/components.json')
        component=next(c for c in catalog['components'] if c['id']=='BW-RESEARCH-PROGRESS-01')
        theme=read_json(ROOT/'library/theme.json')
        template=ROOT/'library/placeholders/templates/BW-RESEARCH-PROGRESS-01.pptx'
        targets=native_targets(template,[s for s in component['slots'] if s['role']=='progress-percent'])
        profile={'shapes':{str(t['shapeId']):{'paragraphs':[{'sizePt':16,'color':'000000','bold':True,
                  'italic':False,'isBlank':False}]} for t in targets}}
        values=[('0%',['0','100']),('62.5%',['62.5','37.5']),('100%',['100','0'])]
        page={'title':'进度回归','images':[],'texts':[{'targetId':t['id'],
              'paragraphs':[{'text':percent,'emphasisPhrases':[]}]} for t,(percent,_) in zip(targets,values,strict=True)]}
        qa=ROOT/'_qa';qa.mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=qa) as folder:
            output=Path(folder)/'progress.pptx'
            fill_component(template,targets,profile,page,{},theme,output)
            data=read_package(output);part,_=slide_part(data,1);root=xml(data[part])
            for target,(label,expected) in zip(targets,values,strict=True):
                label_text=root.xpath(f".//p:sp[p:nvSpPr/p:cNvPr/@id='{target['shapeId']}']//a:t/text()",namespaces=NS)
                self.assertEqual(label_text,[label])
                chart_id=target['chartShapeId']
                rid=root.xpath(f".//p:graphicFrame[p:nvGraphicFramePr/p:cNvPr/@id='{chart_id}']//c:chart/@r:id",namespaces=CHART_NS)[0]
                chart_path=related_part(data,part,rid)
                chart=xml(data[chart_path])
                self.assertEqual(chart.xpath(".//c:numCache/c:pt/c:v/text()",namespaces=CHART_NS),expected)
                workbook_id=chart.find("c:externalData",CHART_NS).get(tag('r','id'))
                with ZipFile(BytesIO(data[related_part(data,chart_path,workbook_id)])) as workbook:
                    sheet=xml(workbook.read('xl/worksheets/sheet1.xml'))
                    cells=sheet.xpath(".//*[local-name()='c'][@r='B2' or @r='B3']/*[local-name()='v']/text()")
                    self.assertEqual(cells,expected)
            page['texts'][0]['paragraphs'][0]['text']='100.1%'
            with self.assertRaisesRegex(ValueError,'0% 至 100%'):
                fill_component(template,targets,profile,page,{},theme,Path(folder)/'invalid.pptx')


if __name__=='__main__':
    run=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(PipelineChecks))
    write_json(ROOT/'docs/deck-pipeline-check.json',{'passed':run.wasSuccessful(),'testsRun':run.testsRun,'mode':'constraint-tests-no-model-or-powerpoint','failures':len(run.failures),'errors':len(run.errors)})
    raise SystemExit(0 if run.wasSuccessful() else 1)
