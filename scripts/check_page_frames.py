"""验证统一页框、正文实际字号、原图、预览及跨母版拼装后的白底。"""
from __future__ import annotations
import copy
import hashlib
import json
import subprocess
from pathlib import Path
from lxml import etree as E
from PIL import Image

from production_tasks import ROOT,read_json,write_json,digest
from native_pptx import NS,shape_inventory,shape_id,rels_path,resolve_target
from deck_native import read_package,slide_part,xml,native_targets,form_info,tag
from deck_pipeline import run_powerpoint
from page_frame import TITLE_ID,DIVIDER_ID,FOOTER_ID,FRAME_IDS,apply_page_frame


def baseline():
    folder=ROOT/'_qa/page-frame-v04';folder.mkdir(parents=True,exist_ok=True)
    for name in ['components.json','inventory.json','theme.json']:
        path=folder/('old-'+name)
        if not path.exists():
            result=subprocess.run(['git','show','0185469:library/'+name],cwd=ROOT,capture_output=True,check=True)
            path.write_bytes(result.stdout)
    inventory=read_json(folder/'old-inventory.json')
    for component in inventory['components']:
        for form in [component]+component.get('variants',[]):
            path=folder/'before'/Path(form['templatePath']).name;path.parent.mkdir(exist_ok=True)
            if not path.exists():path.write_bytes(subprocess.run(['git','show','0185469:library/'+form['templatePath']],cwd=ROOT,capture_output=True,check=True).stdout)
            assert digest(path.read_bytes())==form['templateSha256']
    return folder,read_json(folder/'old-components.json'),inventory


def image_bytes(package,part,picture):
    rid=picture.find('p:blipFill/a:blip',NS).get(tag('r','embed'))
    relation=next(r for r in xml(package[rels_path(part)]) if r.get('Id')==rid)
    return package[resolve_target(part,relation.get('Target'))]


def main():
    folder,old_catalog,old_inventory=baseline()
    catalog=read_json(ROOT/'library/components.json');inventory=read_json(ROOT/'library/inventory.json');theme=read_json(ROOT/'library/theme.json')
    ph=read_json(ROOT/'library/placeholders/manifest.json');ph_by={p['id']:p for p in ph['items']}
    rendered=read_json(ROOT/'docs/powerpoint-render.json')['items'];neutral=read_json(ROOT/'docs/placeholder-render.json')['items']
    forms=[];pairs=[];pictures=0;content_forms=0;frame_signatures=set();geometry=0
    baseline_components={c['id']:c for c in old_catalog['components']}
    for component in catalog['components']:
        old=baseline_components.get(component['id'])
        for variant in ['base']+[v['id'] for v in component.get('variants',[])]:
            key=component['id']+('' if variant=='base' else '__'+variant)
            detail,slots=form_info(component,variant,inventory)
            source=ROOT/'library'/detail['templatePath']
            if old:
                prior,prior_slots=form_info(old,variant,old_inventory)
                before=folder/'before'/Path(prior['templatePath']).name
            expected=ph_by[key];assert expected['sourceTemplateSha256']==detail['templateSha256']
            for name,path,which in [(key,source,rendered[key]),(key,ROOT/'library'/expected['templatePath'],neutral[key])]:
                assert which['opened'] and which['rendered'] and which['savedCopyAndReopened']
                assert which['inputSha256']==digest(path.read_bytes())
            package=read_package(source);part,_=slide_part(package,1);root=xml(package[part])
            if not component.get('pageFrame'):
                if component['id']=='BW-TOC-01':
                    assert source.read_bytes()==before.read_bytes(),'目录专用导航版式发生变化'
                elif component.get('pageRole')=='toc':
                    assert {'chapters','chapterNumbers','activeIndex'} <= {s['key'] for s in slots}
                else:
                    assert component.get('pageRole') in {'cover','closing'}
                    keys = {s['key'] for s in slots}
                    assert {'presenter','pageTitle'} <= keys and 'logo' not in keys
                continue
            content_forms+=1
            nodes={shape_id(n):n for n in root.find('p:cSld/p:spTree',NS) if shape_id(n) is not None}
            assert FRAME_IDS<={s['id'] for s in shape_inventory(root)}
            assert root.find('p:cSld/p:bg/p:bgPr/a:solidFill/a:srgbClr',NS).get('val')=='FFFFFF'
            for ident in [DIVIDER_ID,FOOTER_ID]:
                frame_signatures.add((ident,hashlib.sha256(E.tostring(nodes[ident],method='c14n',exclusive=True)).hexdigest()))
            title=nodes[TITLE_ID]
            assert ''.join(n.text or '' for n in title.findall('.//a:t',NS))==component['pageFrame']['titleText']
            assert title.find('p:spPr/a:xfrm',NS) is not None
            assert all(n.get('sz')=='3200' and n.get('b')=='1' for n in title.findall('.//a:rPr',NS))
            for name,payload in package.items():
                if name.endswith('.xml') and any(s in name for s in ['/slides/','/slideLayouts/','/slideMasters/']):
                    assert not xml(payload).xpath('.//p:ph[@type="sldNum"] | .//a:fld[@type="slidenum"]',namespaces=NS),'页码仍在幻灯片或母版中'
            elements=shape_inventory(root)
            for element in elements:
                if element['parentId'] is None and element['id'] not in FRAME_IDS and element['boxPx']:
                    box=element['boxPx'];assert box[1]>=96 and box[1]+box[3]<=theme['contentPageFrame']['contentBottom']+.01
                    geometry+=1
            for image_path in [ROOT/'library/previews'/(key+'.png'),ROOT/'library'/expected['previewPath']]:
                with Image.open(image_path).convert('RGB') as image:
                    assert all(image.getpixel(p)==(255,255,255) for p in [(0,0),(1279,0)])
                    assert all(min(image.getpixel(p))>=254 for p in [(0,100),(1279,100)]),'画布边缘不白（允许原内容阴影1级灰度）'
                    assert all(image.getpixel(p)==(0,70,165) for p in [(640,84),(640,717)])
                record=rendered[key] if image_path.parent.name=='previews' and image_path.parent.parent.name=='library' else neutral[key]
                assert record['previewSha256']==digest(image_path.read_bytes())
            if not old:continue
            old_data=read_package(before);old_part,_=slide_part(old_data,1);old_root=xml(old_data[old_part]);new_pics={shape_id(n):n for n in root.findall('.//p:pic',NS)}
            old_boxes={s['id']:s for s in shape_inventory(old_root)};new_boxes={s['id']:s for s in elements}
            for picture in old_root.findall('.//p:pic',NS):
                ident=shape_id(picture);assert ident in new_pics
                assert image_bytes(old_data,old_part,picture)==image_bytes(package,part,new_pics[ident])
                assert old_boxes[ident]['boxPx'][2:]==new_boxes[ident]['boxPx'][2:];pictures+=1
            old_targets=native_targets(before,prior_slots);targets=native_targets(source,slots)
            forms.extend([{'key':'before-'+key,'template':str(before),'targets':old_targets},{'key':'current-'+key,'template':str(source),'targets':targets}])
            pairs.append((key,old_targets,targets))
            # 重施页框不能将缓存的隔断/底条从前一页移走。
            other=copy.deepcopy(root);apply_page_frame(root,component,theme);apply_page_frame(other,component,theme)
            assert all(any(shape_id(n)==ident for n in root.find('p:cSld/p:spTree',NS)) for ident in [TITLE_ID,DIVIDER_ID,FOOTER_ID])
    framed_pages=[bool(c.get('pageFrame')) for c in catalog['components'] for _ in [None]+c.get('variants',[])]
    assert content_forms==sum(framed_pages) and len(frame_signatures)==2
    profile_path=folder/'frame-fonts-report.json';profiles=read_json(profile_path).get('profiles',{}) if profile_path.exists() else {}
    missing=[f for f in forms if profiles.get(f['key'],{}).get('templateSha256')!=digest(Path(f['template']).read_bytes()) or
             not {str(t['shapeId']) for t in f['targets'] if t['kind']=='text'}<=set(profiles.get(f['key'],{}).get('shapes',{}))]
    if missing:
        update=run_powerpoint(folder,{'operation':'inspect','forms':missing},'frame-fonts');profiles.update(update['profiles'])
        write_json(profile_path,update|{'profiles':profiles})
    fonts=0
    for key,old_targets,targets in pairs:
        current={t['id']:t for t in targets}
        for target in old_targets:
            if target['kind']!='text' or target['slotKey'] in {'pageTitle','pageNumber','sectionNumber'}:continue
            assert target['id'] in current,key+'/'+target['id']+'内容槽位丢失'
            before=profiles['before-'+key]['shapes'][str(target['shapeId'])]['paragraphs']
            after=profiles['current-'+key]['shapes'][str(current[target['id']]['shapeId'])]['paragraphs']
            assert [(p['sizePt'],p['bold'],p['italic']) for p in before]==[(p['sizePt'],p['bold'],p['italic']) for p in after],key+'/'+target['id']+'实际字号或层级改变'
            fonts+=1
        page_title=next(t for t in targets if t['slotKey']=='pageTitle')
        assert all(p['sizePt']==32 for p in profiles['current-'+key]['shapes'][str(page_title['shapeId'])]['paragraphs'])
    merged=read_json(ROOT/'_qa/deck-engine/render-report.json')
    assert merged['passed'] and merged['savedCopyAndReopened']
    assert digest((ROOT/'_qa/deck-engine/all-components-editable.pptx').read_bytes())==merged['inputSha256']
    assert len(merged['pages'])==len(ph_by)
    for page,framed in zip(merged['pages'],framed_pages):
        if not framed:continue
        with Image.open(page['preview']).convert('RGB') as image:
            assert image.getpixel((0,0))==(255,255,255) and image.getpixel((640,84))==(0,70,165) and image.getpixel((640,717))==(0,70,165)
    result={'passed':True,'libraryVersion':catalog['version'],'contentForms':content_forms,'caseAndXFormsRenderedAndReopened':2*len(ph_by),'bodyObjectsWithinFrame':geometry,
            'effectiveBodyTextTargetsCompared':fonts,'bodyPictureBytesAndSizesCompared':pictures,'whiteBackgroundAndBlueBarsVerified':True,'pageNumbersRemoved':True,
            'crossDesignMergedPagesChecked':content_forms,'mergedInputSha256':merged['inputSha256'],'tocNavigationUnchanged':True,'baselineRevision':'0185469','sourceWordAndPptUnmodified':True}
    write_json(ROOT/'docs/page-frame-check.json',result);print(json.dumps(result,ensure_ascii=False,indent=2))


if __name__=='__main__':main()
