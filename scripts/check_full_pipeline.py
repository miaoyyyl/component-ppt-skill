"""生成明确标注的功能测试Word，并检查真实自动制作结果。"""
from __future__ import annotations
import argparse
import io
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as E

from production_tasks import ROOT,read_json,write_json,digest


def prepare():
    import pythoncom
    from win32com.client import DispatchEx
    from powerpoint_worker import powerpoint_lock
    folder=ROOT/'_qa/full-pipeline';folder.mkdir(parents=True,exist_ok=True)
    with powerpoint_lock():
        pythoncom.CoInitialize();application=None;presentation=None
        try:
            application=DispatchEx('PowerPoint.Application');presentation=application.Presentations.Add(False)
            presentation.PageSetup.SlideWidth=960;presentation.PageSetup.SlideHeight=540
            for index,label in enumerate(['A','B'],1):
                slide=presentation.Slides.Add(index,12);rect=slide.Shapes.AddShape(1,0,0,960,540)
                rect.Fill.ForeColor.RGB=0xA54600 if index==1 else 0xF8F5F3;rect.Line.Visible=0
                box=slide.Shapes.AddTextbox(1,80,150,800,170);box.TextFrame.TextRange.Text='图像 '+label
                box.TextFrame.TextRange.Font.Name='微软雅黑';box.TextFrame.TextRange.Font.Size=80;box.TextFrame.TextRange.Font.Bold=-1
                box.TextFrame.TextRange.Font.Color.RGB=0xFFFFFF if index==1 else 0xA54600;box.TextFrame.TextRange.ParagraphFormat.Alignment=2
                note=slide.Shapes.AddTextbox(1,80,350,800,70);note.TextFrame.TextRange.Text='仅供图片绑定与比例保持测试'
                note.TextFrame.TextRange.Font.Name='微软雅黑';note.TextFrame.TextRange.Font.Size=26;note.TextFrame.TextRange.Font.Color.RGB=0xFFFFFF if index==1 else 0xA54600;note.TextFrame.TextRange.ParagraphFormat.Alignment=2
                slide.Export(str(folder/f'image-{label}.png'),'PNG',1280,720)
            presentation.SaveAs(str(folder/'test-image-sources.pptx'),24)
        finally:
            if presentation is not None:presentation.Close()
            if application is not None and application.Presentations.Count==0:application.Quit()
            pythoncom.CoUninitialize()
    w='http://schemas.openxmlformats.org/wordprocessingml/2006/main';a='http://schemas.openxmlformats.org/drawingml/2006/main';r='http://schemas.openxmlformats.org/officeDocument/2006/relationships';wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';pic='http://schemas.openxmlformats.org/drawingml/2006/picture'
    document=E.Element('{'+w+'}document');body=E.SubElement(document,'{'+w+'}body')
    paragraphs=['通用PPT流程功能验证（非研究材料）','本文件仅用于检查自动制作工具，不是任何科研主题的正式汇报，也没有实验数据。',
        '第一页用三类并列与共同特征归纳：文字绑定、图像绑定、成稿检查。文字绑定要保持字号和位置；图像绑定要保持原比例和图框；成稿检查要实际打开、渲染与保存重开。三类共同要求原生可编辑、来源可追溯。',
        '第二页用总述与双栏图文要点：左栏介绍图像A，右栏介绍图像B。两张图均由程序为本次功能测试生成，只用于确认不同真实输入图像已进入对应图框，不承担科研证据职责。',
        '左栏三项要点：原图可替换、图框位置固定、比例保持。右栏三项要点：原字节保存、对应来源记录、实际渲染检查。',
        '本次只制作两页，不需要封面、结束或文献专页。所有内容已给出，无需联网补充，不引入外部科研结论。']
    for text in paragraphs:E.SubElement(E.SubElement(E.SubElement(body,'{'+w+'}p'),'{'+w+'}r'),'{'+w+'}t').text=text
    for index,label in enumerate(['A','B'],1):
        p=E.SubElement(body,'{'+w+'}p');E.SubElement(E.SubElement(p,'{'+w+'}r'),'{'+w+'}t').text='本次功能测试原图'+label+'：'
        drawing=E.SubElement(E.SubElement(p,'{'+w+'}r'),'{'+w+'}drawing');inline=E.SubElement(drawing,'{'+wp+'}inline')
        E.SubElement(inline,'{'+wp+'}extent',cx='4572000',cy='2571750');E.SubElement(inline,'{'+wp+'}docPr',id=str(index),name='功能测试图'+label)
        graphic=E.SubElement(inline,'{'+a+'}graphic');gd=E.SubElement(graphic,'{'+a+'}graphicData',uri=pic);image=E.SubElement(gd,'{'+pic+'}pic')
        nv=E.SubElement(image,'{'+pic+'}nvPicPr');E.SubElement(nv,'{'+pic+'}cNvPr',id=str(index),name='image-'+label+'.png');E.SubElement(nv,'{'+pic+'}cNvPicPr')
        fill=E.SubElement(image,'{'+pic+'}blipFill');E.SubElement(fill,'{'+a+'}blip',{'{'+r+'}embed':'rIdImage'+str(index)});E.SubElement(E.SubElement(fill,'{'+a+'}stretch'),'{'+a+'}fillRect')
        sp=E.SubElement(image,'{'+pic+'}spPr');x=E.SubElement(sp,'{'+a+'}xfrm');E.SubElement(x,'{'+a+'}off',x='0',y='0');E.SubElement(x,'{'+a+'}ext',cx='4572000',cy='2571750');E.SubElement(E.SubElement(sp,'{'+a+'}prstGeom',prst='rect'),'{'+a+'}avLst')
    output=folder/'full-pipeline-test.docx'
    with zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+r+'/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr('word/document.xml',E.tostring(document,encoding='utf-8',xml_declaration=True))
        z.writestr('word/_rels/document.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rIdImage{i}" Type="{r}/image" Target="media/image-{label}.png"/>' for i,label in enumerate(['A','B'],1))+'</Relationships>')
        for label in ['A','B']:z.writestr('word/media/image-'+label+'.png',(folder/f'image-{label}.png').read_bytes())
    print(output)


def verify(job_id,data_dir):
    from pptx import Presentation
    from deck_native import read_package,slide_part,xml,NS,rels_path,resolve_target,tag
    output=Path(data_dir)/'results'/job_id;report=read_json(output/'presentation-check.json');plan=read_json(output/'render-plan.json');bindings=read_json(output/'image-bindings.json')
    assert report['passed'] and report['opened'] and report['rendered'] and report['savedCopyAndReopened']
    deck=Presentation(str(output/'presentation.pptx'));assert len(deck.slides)==2
    assert len(bindings)==2 and {b['assetId'] for b in bindings}=={'D1:image1','D1:image2'}
    package=read_package(output/'presentation.pptx');checked=0
    manifest=read_json(Path(data_dir)/'runs'/job_id/'render-manifest.json')
    for page in plan['pages']:
        assert page['componentId'] in {'BW-TRIPLE-COMMON-01','BW-TWO-PERSPECTIVE-01'}
        if page['componentId']=='BW-TRIPLE-COMMON-01':assert page['variantId']=='all-text'
        part,_=slide_part(package,page['pageNumber']);root=xml(package[part]);rels={r.get('Id'):r for r in xml(package[rels_path(part)])}
        for binding in [b for b in bindings if b['pageNumber']==page['pageNumber']]:
            target=next(t for p in manifest['pages'] if p['pageNumber']==page['pageNumber'] for t in p['targets'] if t['id']==binding['targetId'])
            picture=next(p for p in root.findall('.//p:pic',NS) if p.find('p:nvPicPr/p:cNvPr',NS).get('name')=='prod-'+str(target['shapeId']))
            rid=picture.find('p:blipFill/a:blip',NS).get(tag('r','embed'));path=resolve_target(part,rels[rid].get('Target'))
            assert digest(package[path])==binding['sha256'];checked+=1
    assert checked==2
    first_plan=next(p for p in plan['pages'] if p['componentId']=='BW-TRIPLE-COMMON-01')
    assert first_plan['hiddenRegionIds']==['column-links']
    part,_=slide_part(package,first_plan['pageNumber']);root=xml(package[part])
    assert all(next(n for n in root.findall('.//p:cNvPr',NS) if n.get('name')=='prod-'+str(ident)).get('hidden')=='1' for ident in [40,41,42,43])
    assert all(s.notes_slide.notes_text_frame and '初版完整内容' in s.notes_slide.notes_text_frame.text for s in deck.slides)
    write_json(ROOT/'docs/full-pipeline-check.json',{'passed':True,'executionMode':'live-model','inputKind':'generated-functional-test-docx-with-two-images','jobId':job_id,'slides':2,
        'geometryPreserved':report['geometryPreserved'],'nativeTextShapes':report['nativeTextShapes'],'wordImagesBytesVerified':checked,'powerPointOpened':True,'powerPointRendered':True,'savedCopyAndReopened':True,
        'modelVisualReviews':len(report['visualReviews']),'modelVisualWarnings':sum(i['severity']=='warning' for r in report['visualReviews'][-1]['reviews'] for i in r['issues']),
        'optionalLinksHidden':True,'presentationSha256':digest((output/'presentation.pptx').read_bytes()),'missingImages':len(report['missingImages']),'powerPointIssues':len(report['powerPointIssues'])})
    print('完整自动制作检查通过：2页可编辑PPTX，2张Word原图字节一致，实际渲染与模型检查通过。')


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--prepare',action='store_true');parser.add_argument('--job');parser.add_argument('--data-dir',default='_qa/full-service');args=parser.parse_args()
    if args.prepare:prepare()
    else:verify(args.job,args.data_dir)
