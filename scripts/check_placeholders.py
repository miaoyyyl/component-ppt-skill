"""核对X占位的原字数、富文本格式、框体、原图及实际预览。"""
from production_tasks import ROOT,read_json,write_json,digest
from deck_native import read_package,slide_part,xml,NS
from lxml import etree as E


def main():
    manifest=read_json(ROOT/'library/placeholders/manifest.json');inventory=read_json(ROOT/'library/inventory.json');render=read_json(ROOT/'docs/placeholder-render.json')['items']
    sources={c['id']:c for c in inventory['components']}
    for c in inventory['components']:
        for v in c.get('variants',[]):sources[c['id']+'__'+v['id']]=v
    characters=paragraphs=0;images=0
    for item in manifest['items']:
        original=ROOT/'library'/sources[item['id']]['templatePath'];neutral=ROOT/'library'/item['templatePath'];before=read_package(original);after=read_package(neutral)
        part,_=slide_part(before,1);a=xml(before[part]);b=xml(after[part]);left=a.findall('.//p:sp/p:txBody//a:t',NS);right=b.findall('.//p:sp/p:txBody//a:t',NS)
        assert len(left)==len(right)
        for x,y in zip(left,right):
            source=x.text or '';placeholder=y.text or '';assert len(source)==len(placeholder)
            assert [i for i,c in enumerate(source) if c.isspace()]==[i for i,c in enumerate(placeholder) if c.isspace()]
            x.text='';y.text=''
        assert E.tostring(a,method='c14n')==E.tostring(b,method='c14n'),'占位改变了段落、样式或框体'
        for name,payload in before.items():
            if '/media/' in name:assert after[name]==payload;images+=1
        assert digest(original.read_bytes())==item['sourceTemplateSha256'];assert digest(neutral.read_bytes())==item['templateSha256']
        assert render[item['id']]['inputSha256']==item['templateSha256'];assert render[item['id']]['previewSha256']==digest((ROOT/'library'/item['previewPath']).read_bytes())
        characters+=item['textCharacters'];paragraphs+=sum(len(t['paragraphCharacters']) for t in item['textCapacities'])
    write_json(ROOT/'docs/placeholder-check.json',{'passed':True,'forms':len(manifest['items']),'characters':characters,'paragraphs':paragraphs,'mediaPartsVerified':images,'textRunLengthsPreserved':True,'richTextStylesPreserved':True,'geometryPreserved':True,'powerPointPreviewsVerified':True})
    print('X占位检查通过：',len(manifest['items']),'种形式，',characters,'字符；字数、样式、段落、图框和原图保持一致。')


if __name__=='__main__':main()
