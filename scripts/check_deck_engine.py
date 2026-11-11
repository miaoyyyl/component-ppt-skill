"""检查全部组件形式的原生填充、图片比例、跨母版拼装和实际渲染。"""
from __future__ import annotations
import argparse
import json
import subprocess
import sys
from pathlib import Path

from deck_native import fill_component,form_info,merge_presentations,native_targets,read_package,xml,slide_part,NS,rels_path,resolve_target,tag
from production_tasks import ROOT,read_json,write_json,digest


def main():
    parser=argparse.ArgumentParser();parser.add_argument("--prepare-only",action="store_true");parser.add_argument("--build-only",action="store_true");args=parser.parse_args()
    folder=ROOT/"_qa/deck-engine";folder.mkdir(parents=True,exist_ok=True)
    catalog=read_json(ROOT/"library/components.json");inventory=read_json(ROOT/"library/inventory.json");theme=read_json(ROOT/"library/theme.json")
    forms=[]
    for component in catalog["components"]:
        for variant in ["base"]+[v["id"] for v in component.get("variants",[])]:
            detail,slots=form_info(component,variant,inventory);path=ROOT/"library"/detail["templatePath"]
            forms.append({"key":component["id"]+"__"+variant,"template":str(path),"targets":native_targets(path,slots)})
    inspect_request=folder/"inspect-request.json"
    write_json(inspect_request,{"operation":"inspect","forms":forms,"report":str(folder/"styles.json")})
    if args.prepare_only:return
    cached=read_json(folder/'styles.json').get('profiles',{}) if (folder/'styles.json').is_file() else {}
    if any(cached.get(f['key'],{}).get('templateSha256')!=digest(Path(f['template']).read_bytes()) or
           not {str(t['shapeId']) for t in f['targets'] if t['kind']=='text'}<=set(cached.get(f['key'],{}).get('shapes',{})) for f in forms):
        subprocess.run([sys.executable,str(ROOT/"scripts/powerpoint_worker.py"),str(inspect_request)],check=True,timeout=180)
    profiles=read_json(folder/"styles.json")["profiles"];pages=[];checks=[];notes=[]
    asset=ROOT/"library/references/reference-10.png";before=digest(asset.read_bytes())
    for number,form in enumerate(forms,1):
        assert profiles[form["key"]]["templateSha256"]==digest(Path(form["template"]).read_bytes())
        page={"pageNumber":number,"title":"组件填充功能测试","texts":[],"images":[]}
        for target in form["targets"]:
            if target["kind"]=="image":page["images"].append({"targetId":target["id"],"assetId":"test-image"});continue
            word="1" if target["role"] in {"section-number","hierarchy-number","page-number"} or target["slotKey"]=="chapterNumbers" else "测"
            if target["role"]=="progress-percent":word=["100%","62.5%","0%"][target["index"]]
            sizes={p['sizePt'] for p in profiles[form['key']]['shapes'][str(target['shapeId'])]['paragraphs'] if not p['isBlank']}
            count=max(1,sum(not p["isBlank"] for p in target["paragraphs"])) if len(sizes)>1 else 1
            page["texts"].append({"targetId":target["id"],"paragraphs":[{"text":word,"emphasisPhrases":[]} for _ in range(count)]})
        output=folder/f"filled-{number:02d}.pptx"
        detail=fill_component(form["template"],form["targets"],profiles[form["key"]],page,{"test-image":asset},theme,output)
        assert detail["geometryPreserved"]
        for target in detail["targets"]:
            if target["kind"]=="image":assert target["sha256"]==before and all(n<=0 for n in target["containCrop"])
        pages.append(output);checks.append(detail|{"componentForm":form["key"],"pageNumber":number})
        notes.append("功能测试，不是主题成稿。\n组件形式："+form["key"]+"\n图片仅复用用户参考图验证绑定，不能作为新研究证据。")
    merged=folder/"all-components-editable.pptx";merge=merge_presentations(pages,merged,notes,"全部组件原生填充检查")
    assert merge['geometryPreserved']
    from pptx import Presentation
    reopened=Presentation(str(merged))
    assert all(form['key'] in slide.notes_slide.notes_text_frame.text for slide,form in zip(reopened.slides,forms))
    data=read_package(merged)
    for number,check in enumerate(checks,1):
        part,_=slide_part(data,number);root=xml(data[part]);texts="".join(n.text or "" for n in root.findall(".//a:t",NS))
        assert "微塑料" not in texts and "太湖" not in texts and "干旱" not in texts
        rels={r.get('Id'):r for r in xml(data[rels_path(part)])}
        for target in check['targets']:
            if target['kind']!='image':continue
            picture=next(p for p in root.findall('.//p:pic',NS) if p.find('p:nvPicPr/p:cNvPr',NS).get('name')=='prod-'+str(target['shapeId']))
            rid=picture.find('p:blipFill/a:blip',NS).get(tag('r','embed'));media=resolve_target(part,rels[rid].get('Target'))
            assert digest(data[media])==target['sha256'],'拼装改变了本次图片字节'
    write_json(folder/"render-request.json",{"operation":"render","input":str(merged),"output":str(folder/"previews"),"report":str(folder/"render-report.json"),"pages":checks})
    if not args.build_only:subprocess.run([sys.executable,str(ROOT/"scripts/powerpoint_worker.py"),str(folder/"render-request.json")],check=True,timeout=240)
    assert before==digest(asset.read_bytes())
    report={"nativeFillPassed":True,"forms":len(forms),"targets":sum(len(f["targets"]) for f in forms),"geometryPreserved":True,"mergedObjectsChecked":merge['objectsChecked'],"mergedImagesChecked":sum(t['kind']=='image' for f in forms for t in f['targets']),"sourceImageBytesUnchanged":True,"notesPreserved":True,"editableTextShapes":merge["nativeTextShapes"],"mergedSlides":merge["slides"]}
    if (folder/"render-report.json").is_file():
        actual=read_json(folder/"render-report.json");report["powerPoint"]= {k:actual[k] for k in ["opened","rendered","savedCopyAndReopened","passed","issues","inputSha256"]}
    write_json(ROOT/"docs/deck-engine-check.json",report)
    print(json.dumps(report,ensure_ascii=False,indent=2))
    if 'powerPoint' in report:assert report['powerPoint']['passed'],'PowerPoint实际渲染检查未通过'


if __name__=="__main__":main()
