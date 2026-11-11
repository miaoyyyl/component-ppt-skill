"""为组件生成等字数X占位副本；原段落、字号、框体和原图像保持不变。"""
from __future__ import annotations
import argparse
import json
import re
import unicodedata
from pathlib import Path

from native_pptx import NS,is_fixed_text,shape_id,xml,serialize,slide_part,shape_inventory
from deck_native import read_package,save_package
from production_tasks import ROOT,read_json,write_json,digest


def placeholder_text(value,keep_number=False):
    if keep_number:return value
    result=[]
    for char in value:
        if char.isalpha() or char.isdigit():
            result.append('Ｘ' if unicodedata.east_asian_width(char) in {'W','F'} else 'X')
        else:result.append(char)
    return ''.join(result)


def build():
    catalog=read_json(ROOT/'library/components.json');inventory=read_json(ROOT/'library/inventory.json');output=ROOT/'library/placeholders';entries=[]
    for component in catalog['components']:
        detail=next(d for d in inventory['components'] if d['id']==component['id'])
        for form in [detail]+detail.get('variants',[]):
            variant='base' if form is detail else form['id'];key=component['id']+('' if variant=='base' else '__'+variant)
            original=ROOT/'library'/form['templatePath'];data=read_package(original);part,_=slide_part(data,1);root=xml(data[part]);before=shape_inventory(root);capacities=[]
            number_keys={'chapterNumbers','listNumbers','thirdListNumbers','numbers','pageNumber','sectionNumber'}
            numeric_ids={i for b in form['slotBindings'] if b['key'] in number_keys for i in b['textShapeIds']}
            references={i for b in form['slotBindings'] if b['key']=='bibliography' for i in b['textShapeIds']}
            for shape in root.findall('.//p:sp',NS):
                ident=shape_id(shape);body=shape.find('p:txBody',NS)
                if body is None:continue
                keep_text=ident in numeric_ids or is_fixed_text(shape)
                paragraphs=[]
                for paragraph in body.findall('a:p',NS):
                    count=0
                    for text in paragraph.findall('.//a:t',NS):
                        source=text.text or '';masked=placeholder_text(source,keep_text)
                        if ident in references and count==0:
                            prefix=re.match(r'^\[\d+\]',source)
                            if prefix:masked=source[:prefix.end()]+masked[prefix.end():]
                        assert len(source)==len(masked);text.text=masked;count+=sum(not c.isspace() for c in source)
                    paragraphs.append(count)
                capacities.append({'shapeId':ident,'paragraphCharacters':paragraphs,'characters':sum(paragraphs)})
            data[part]=serialize(root)
            for name in [n for n in data if n.startswith('ppt/notesSlides/') and n.endswith('.xml')]:
                note=xml(data[name])
                for t in note.findall('.//a:t',NS):t.text=placeholder_text(t.text or '')
                data[name]=serialize(note)
            after=shape_inventory(root)
            for a,b in zip(before,after):assert a['id']==b['id'] and a['boxPx']==b['boxPx'] and a['fontSizesPt']==b['fontSizesPt'] and a['textFrame']==b['textFrame']
            target=output/'templates'/(key+'.pptx');save_package(data,target)
            entries.append({'id':key,'componentId':component['id'],'variantId':variant,'sourceTemplateSha256':digest(original.read_bytes()),
                            'templatePath':'placeholders/templates/'+key+'.pptx','templateSha256':digest(target.read_bytes()),'previewPath':'placeholders/previews/'+key+'.png',
                            'textCapacities':capacities,'textCharacters':sum(c['characters'] for c in capacities),'geometryPreserved':True,'imagesUnchanged':True})
    write_json(output/'manifest.json',{'schemaVersion':1,'libraryVersion':catalog['version'],'textPolicy':'中文全角Ｘ、英文和数字半角X；保留标点、空白、段落、结构序号及固定标签','items':entries})
    return entries


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--render',action='store_true')
    parser.add_argument('--only',nargs='+',help='仅渲染这些组件，保留其他预览记录')
    args=parser.parse_args();entries=build()
    if args.only:
        unknown=set(args.only)-{p['componentId'] for p in entries}
        if unknown:raise ValueError(f'组件未登记：{sorted(unknown)}')
    if args.render:
        from deck_pipeline import run_powerpoint
        cache=ROOT/'_qa/placeholders';cache.mkdir(parents=True,exist_ok=True)
        items=[]
        for item in entries:
            if args.only and item['componentId'] not in args.only:continue
            items.append({'input':str(ROOT/'library'/item['templatePath']),'output':str(cache/item['id']),
                          'image':str(ROOT/'library'/item['previewPath']),'id':item['id']})
        result=run_powerpoint(cache,{'operation':'placeholder-batch','items':items},'placeholder-render')
        if args.only:
            previous=read_json(ROOT/'docs/placeholder-render.json')
            result['items']={**previous['items'],**result['items']}
        write_json(ROOT/'docs/placeholder-render.json',result)
    print('已构建等字数占位组件：',len(entries),'种形式')


if __name__=='__main__':main()
