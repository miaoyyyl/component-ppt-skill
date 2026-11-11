"""按组件槽位填充原生对象，并合并其OOXML依赖；不把幻灯片扁平化为图片。"""
from __future__ import annotations
import copy
import io
import json
import posixpath
import re
import subprocess
import sys
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree as E
from PIL import Image

from native_pptx import (NS, SHAPES, is_fixed_text, xml, serialize, shape_id, shape_inventory, slide_part,
                         rels_path, resolve_target, tag, set_fonts, solid_text_color)
from component_variants import find_shape
from production_tasks import digest

CT="http://schemas.openxmlformats.org/package/2006/content-types"
REL="http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
GENERATED_DATE=(2026,11,2,0,0,0)


def read_package(path):
    with zipfile.ZipFile(path) as archive:return {n:archive.read(n) for n in archive.namelist()}


def save_package(data,path):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as archive:
        for name in sorted(data):
            info=zipfile.ZipInfo(name,date_time=GENERATED_DATE);info.compress_type=zipfile.ZIP_DEFLATED
            archive.writestr(info,data[name])


def form_info(component,variant_id,inventory):
    detail=next(v for v in inventory["components"] if v["id"]==component["id"])
    if variant_id=="base":return detail,component["slots"]
    variant=next((v for v in detail.get("variants",[]) if v["id"]==variant_id),None)
    if variant is None:raise ValueError("组件变体未登记")
    return variant,[s for s in component["slots"] if s["key"] not in variant["replacedSlotKeys"]]+variant["addedSlots"]


def native_targets(template,slots):
    data=read_package(template);part,_=slide_part(data,1);root=xml(data[part]);elements={e["id"]:e for e in shape_inventory(root)}
    targets=[];seen=set()
    for slot in slots:
        if slot["kind"]=="state":continue
        nodes=[]
        ids=slot["shapeIds"]
        if slot["key"] in {"chapters","chapterNumbers"} and any(i>=20000 for i in elements):
            suffix="-label" if slot["key"]=="chapters" else "-number"
            ids=[shape_id(n) for n in root.findall("p:cSld/p:spTree/p:sp",NS) if n.find("p:nvSpPr/p:cNvPr",NS).get("name","").endswith(suffix)]
        for ident in ids:
            node=find_shape(root,ident)
            if slot["kind"].startswith("image"):
                nodes.extend([node] if E.QName(node).localname=="pic" else node.findall(".//p:pic",NS))
            elif node.find("p:txBody",NS) is not None:nodes.append(node)
            else:nodes.extend(n for n in node.findall(".//p:sp",NS) if n.find("p:txBody",NS) is not None)
        for index,node in enumerate(nodes):
            ident=shape_id(node)
            if ident in seen:raise ValueError(f"对象{ident}被多个内容槽位重复绑定")
            seen.add(ident);body=node.find("p:txBody",NS);paragraphs=[]
            if body is not None:
                for p in body.findall("a:p",NS):
                    runs=[];offset=0
                    for run_index,child in enumerate(n for n in p if E.QName(n).localname in {"r","fld","br"}):
                        kind=E.QName(child).localname
                        if kind not in {"r","fld","br"}:continue
                        text="\n" if kind=="br" else "".join(n.text or "" for n in child.findall("a:t",NS))
                        pr=child.find("a:rPr",NS);color=pr.find("a:solidFill/a:srgbClr",NS) if pr is not None else None
                        runs.append({"offset":offset,"length":len(text.encode("utf-16-le"))//2,"color":color.get("val") if color is not None else None,"runIndex":run_index})
                        offset+=len(text.encode("utf-16-le"))//2
                    normal=[r for r in runs if r["length"] and r["color"] not in {"C00000","FF0000","FFFF00"}]
                    candidate=max(normal or runs or [{"offset":0,"length":0}],key=lambda r:r["length"])
                    pp=p.find("a:pPr",NS)
                    visible_text=''.join(n.text or '' for n in p.findall('.//a:t',NS)).strip()
                    paragraphs.append({"sampleOffset":candidate["offset"],"runIndex":candidate.get("runIndex",0),"isBlank":not visible_text,"sourceLength":offset,
                                       "bullet":pp is not None and any(E.QName(n).localname in {"buChar","buAutoNum"} for n in pp)})
            targets.append({"id":slot["key"]+":"+str(index),"slotKey":slot["key"],"slotKind":slot["kind"],"role":slot["role"],"index":index,
                            "shapeId":ident,"kind":"image" if slot["kind"].startswith("image") else "text","boxPx":elements[ident]["boxPx"],
                            "paragraphs":paragraphs,"purpose":slot["replace"]})
            if slot.get("chartShapeIds"):
                targets[-1]["chartShapeId"]=slot["chartShapeIds"][index]
    return targets


def drop_hyperlinks(node):
    for item in list(node.xpath('.//*[local-name()="hlinkClick" or local-name()="hlinkMouseOver"]')):item.getparent().remove(item)


def fill_text(node,paragraphs,profile,theme):
    body=node.find("p:txBody",NS);props=body.find("a:bodyPr",NS)
    for n in list(props):
        if E.QName(n).localname in {"normAutofit","spAutoFit","noAutofit"}:props.remove(n)
    props.append(E.Element(tag("a","noAutofit")))
    original=list(body.findall("a:p",NS));styles=profile["paragraphs"]
    nonempty=[i for i,p in enumerate(styles) if not p["isBlank"]] or [0]
    for p in original:body.remove(p)
    expected=[]
    for index,content in enumerate(paragraphs):
        prototype_index=nonempty[min(index,len(nonempty)-1)];prototype=original[prototype_index];style=styles[prototype_index]
        paragraph=E.SubElement(body,tag("a","p"));pp=prototype.find("a:pPr",NS)
        if pp is not None:paragraph.append(copy.deepcopy(pp))
        pp=paragraph.find("a:pPr",NS)
        if pp is None:pp=E.SubElement(paragraph,tag("a","pPr"))
        default=pp.find("a:defRPr",NS)
        if default is None:default=E.SubElement(pp,tag("a","defRPr"))
        default.set("sz",str(round(style["sizePt"]*100)));set_fonts(default,theme["typography"]["family"])
        solid_text_color(default,style["color"])
        phrases=[s for s in content.get("emphasisPhrases",[]) if s and s in content["text"]]
        spans=[]
        for phrase in phrases:
            start=0
            while (start:=content["text"].find(phrase,start))>=0:
                spans.append((start,start+len(phrase)));start+=len(phrase)
        boundaries=sorted({0,len(content["text"]),*(i for pair in spans for i in pair)})
        for start,end in zip(boundaries,boundaries[1:]):
            highlighted=any(start>=a and end<=b for a,b in spans)
            # 行内换行保持同段字号、行距与项目符号。
            for line_index,text in enumerate(content["text"][start:end].split("\n")):
                if line_index:E.SubElement(paragraph,tag("a","br"))
                if not text:continue
                run=E.SubElement(paragraph,tag("a","r"))
                original_runs=[n for n in prototype if E.QName(n).localname in {"r","fld","br"}]
                seed=original_runs[min(style.get("runIndex",0),len(original_runs)-1)].find("a:rPr",NS) if original_runs else None
                pr=copy.deepcopy(seed) if seed is not None else E.Element(tag("a","rPr"));run.append(pr)
                pr.set("sz",str(round(style["sizePt"]*100)));pr.set("b","1" if highlighted or style["bold"] else "0");pr.set("i","1" if style["italic"] else "0")
                set_fonts(pr,theme["typography"]["family"])
                color=(theme['palette']['emphasisYellow'] if style['color'] in {'FFFFFF','FFFF00'} else theme['palette']['emphasisRed']) if highlighted else style['color']
                solid_text_color(pr,color);E.SubElement(run,tag("a","t")).text=text
        end=E.SubElement(paragraph,tag("a","endParaRPr"),sz=str(round(style["sizePt"]*100)));set_fonts(end,theme["typography"]["family"])
        first_color=paragraph.find('a:r/a:rPr/a:solidFill/a:srgbClr',NS)
        expected.append({"text":content["text"],"sizePt":style["sizePt"],"color":first_color.get('val') if first_color is not None else style['color']})
    drop_hyperlinks(node)
    return expected


def image_details(path):
    path=Path(path)
    with Image.open(path) as image:
        width,height=image.size;format=image.format
        try:orientation=image.getexif().get(274,1)
        except (AttributeError,ValueError):orientation=1
        if orientation in {5,6,7,8}:width,height=height,width
    suffix={"PNG":"png","JPEG":"jpg","GIF":"gif","BMP":"bmp","TIFF":"tiff","WMF":"wmf"}.get(format)
    if suffix is None:raise ValueError("图片格式不能直接嵌入PowerPoint："+path.name)
    return width,height,suffix,{"png":"image/png","jpg":"image/jpeg","gif":"image/gif","bmp":"image/bmp","tiff":"image/tiff","wmf":"image/x-wmf"}[suffix]


def replace_picture(data,part,node,asset,box):
    payload=Path(asset).read_bytes();width,height,suffix,content_type=image_details(asset)
    media=f"ppt/media/production-{digest(payload)}.{suffix}";data[media]=payload
    relationships=xml(data[rels_path(part)]);used={r.get("Id") for r in relationships};index=1
    while "rIdProdImage"+str(index) in used:index+=1
    rid="rIdProdImage"+str(index)
    E.SubElement(relationships,"{"+NS["rel"]+"}Relationship",Id=rid,Type=REL+"image",Target=posixpath.relpath(media,posixpath.dirname(part)))
    data[rels_path(part)]=serialize(relationships)
    types=xml(data["[Content_Types].xml"])
    if not any(n.get("PartName")=="/"+media for n in types):E.SubElement(types,"{"+CT+"}Override",PartName="/"+media,ContentType=content_type)
    data["[Content_Types].xml"]=serialize(types)
    fill=node.find("p:blipFill",NS);blip=fill.find("a:blip",NS);blip.set(tag("r","embed"),rid);blip.attrib.pop(tag("r","link"),None)
    for child in list(blip):blip.remove(child)
    for child in list(fill):
        if E.QName(child).localname in {"srcRect","tile","stretch"}:fill.remove(child)
    scale=min(box[2]/width,box[3]/height);draw_width,draw_height=width*scale,height*scale
    crop_x=round((draw_width-box[2])/(2*draw_width)*100000);crop_y=round((draw_height-box[3])/(2*draw_height)*100000)
    rect=E.Element(tag("a","srcRect"),l=str(crop_x),r=str(crop_x),t=str(crop_y),b=str(crop_y));fill.insert(1,rect)
    E.SubElement(E.SubElement(fill,tag("a","stretch")),tag("a","fillRect"))
    return {"sha256":digest(payload),"pixels":[width,height],"framePx":box,"containCrop":[crop_x,crop_y,crop_x,crop_y]}


def remove_example_metadata(root):
    for item in list(root.xpath('.//*[local-name()="custDataLst" or local-name()="timing" or local-name()="transition"]')):item.getparent().remove(item)
    drop_hyperlinks(root)


def fill_component(template,targets,profile,page,assets,theme,output,hidden_shape_ids=()):
    data=read_package(template);part,_=slide_part(data,1);root=xml(data[part]);before={s["id"]:s for s in shape_inventory(root)}
    text_values={t["targetId"]:t for t in page["texts"]};image_values={t["targetId"]:t for t in page["images"]};entries=[]
    text_ids={t["shapeId"] for t in targets if t["kind"]=="text"}
    for target in targets:
        node=find_shape(root,target["shapeId"])
        if target["kind"]=="text":
            item=text_values[target["id"]]
            if target['slotKey'] in {'chapters','chapterNumbers'}:
                item=copy.deepcopy(item)
                for p in item['paragraphs']:p['emphasisPhrases']=[]
                if target['slotKey']=='chapterNumbers':
                    number = "".join(node.xpath(".//a:t/text()", namespaces=NS))
                    item['paragraphs']=[{'text':number,'emphasisPhrases':[]}]
            paragraphs=item["paragraphs"]
            if "chartShapeId" in target:
                from progress_rings import set_ring_percentage
                if len(paragraphs)!=1:
                    raise ValueError("完成进度只能填写一个百分比")
                chart=find_shape(root,target["chartShapeId"])
                value=set_ring_percentage(data,part,chart,paragraphs[0]["text"])
                paragraphs=[{"text":value,"emphasisPhrases":[]}]
            expected=fill_text(node,paragraphs,profile["shapes"][str(target["shapeId"])],theme)
            entries.append({"targetId":target["id"],"shapeId":target["shapeId"],"kind":"text","paragraphs":expected,"boxPx":target["boxPx"]})
        else:
            item=image_values[target["id"]];asset=assets.get(item["assetId"])
            if asset is None:raise ValueError("缺少图片资产："+item["assetId"])
            details=replace_picture(data,part,node,asset,target["boxPx"])
            entries.append({"targetId":target["id"],"shapeId":target["shapeId"],"kind":"image",**details})
    # 未绑定的数字标识和通用流程标签属于框架；其他隐藏案例文字清空。
    for shape in root.findall(".//p:sp",NS):
        ident=shape_id(shape);body=shape.find("p:txBody",NS)
        if body is None or ident in text_ids or is_fixed_text(shape):continue
        text="".join(t.text or "" for t in body.findall(".//a:t",NS)).strip()
        if text in {"1","2","技术流程"}:continue
        for t in body.findall(".//a:t",NS):t.text=""
    remove_example_metadata(root)
    for ident in hidden_shape_ids:
        node=find_shape(root,ident)
        if node is None:raise ValueError('可选连接对象不存在')
        node.find('.//p:cNvPr',NS).set('hidden','1')
    for node in root.iter():
        if E.QName(node).localname in SHAPES:
            props=node.find('.//p:cNvPr',NS)
            if props is not None:props.set('name','prod-'+str(shape_id(node)))
    # 移除旧图像关系，避免源案例图片虽不显示仍留在新PPT中。
    referenced={value for n in root.iter() for key,value in n.attrib.items() if key.startswith("{"+NS["r"]+"}")}
    relations=xml(data[rels_path(part)])
    for relation in list(relations):
        typ=relation.get("Type","")
        if relation.get("TargetMode")=="External" or typ.endswith(("/notesSlide","/comments","/tags","/commentAuthors")) or (typ.endswith("/image") and relation.get("Id") not in referenced):relations.remove(relation)
    data[rels_path(part)]=serialize(relations);root.find("p:cSld",NS).set("name",page["title"][:120]);data[part]=serialize(root)
    if root.get('showMasterSp')=='0':
        for master in [p for p in data if p.endswith('.xml') and ('/slideMasters/' in p or '/slideLayouts/' in p)]:
            node=xml(data[master])
            for shape in list(node.xpath('.//p:spTree/*',namespaces=NS)):
                if shape.findall('.//p:oleObj',NS):shape.getparent().remove(shape)
            data[master]=serialize(node);rp=rels_path(master)
            if rp in data:
                rs=xml(data[rp]);refs={v for n in node.iter() for k,v in n.attrib.items() if k.startswith('{'+NS['r']+'}')}
                for r in list(rs):
                    if r.get('Type','').endswith('/oleObject') and r.get('Id') not in refs:rs.remove(r)
                data[rp]=serialize(rs)
    after={s["id"]:s for s in shape_inventory(root)}
    if set(before)!=set(after):raise ValueError("填充改变了原生对象集合")
    for ident,old in before.items():
        if old["boxPx"]!=after[ident]["boxPx"] or old["rotationDegrees"]!=after[ident]["rotationDegrees"]:raise ValueError("填充改变了组件框体位置或尺寸")
    save_package(data,output)
    return {"templateSha256":digest(Path(template).read_bytes()),"filledSha256":digest(Path(output).read_bytes()),"objectCount":len(before),"geometryPreserved":True,"hiddenShapeIds":list(hidden_shape_ids),"targets":entries}


def merge_presentations(pages,output,notes,title,workdir=None):
    from pptx import Presentation
    output=Path(output);output.parent.mkdir(parents=True,exist_ok=True);dimensions=None
    for path in pages:
        source=read_package(path);size=xml(source['ppt/presentation.xml']).find('p:sldSz',NS);current=(int(size.get('cx')),int(size.get('cy')))
        if dimensions is not None and dimensions!=current:raise ValueError('页面画幅不一致，不能缩放组件拼装')
        dimensions=current
    scratch=Path(workdir or output.parent);scratch.mkdir(parents=True,exist_ok=True)
    request=scratch/'assemble-request.json';report=scratch/'assemble-report.json'
    request.write_text(json.dumps({'operation':'assemble','inputs':[str(Path(p).resolve()) for p in pages],'dimensions':dimensions,'notes':notes,'output':str(output.resolve()),'report':str(report.resolve())},ensure_ascii=False),encoding='utf-8')
    subprocess.run([sys.executable,str(Path(__file__).with_name('powerpoint_worker.py')),str(request.resolve())],check=True,timeout=240)
    data=read_package(output);types=xml(data['[Content_Types].xml'])
    # 原生导入后恢复本次图像原字节，避免Office默认图片压缩改变科学图。
    for number,path in enumerate(pages,1):
        source=read_package(path);sp,_=slide_part(source,1);sr=xml(source[sp]);source_rels={r.get('Id'):r for r in xml(source[rels_path(sp)])}
        dp,_=slide_part(data,number);dr=xml(data[dp]);rels=xml(data[rels_path(dp)]);by_name={n.find('p:nvPicPr/p:cNvPr',NS).get('name'):n for n in dr.findall('.//p:pic',NS)}
        for index,pic in enumerate(sr.findall('.//p:pic',NS),1):
            name=pic.find('p:nvPicPr/p:cNvPr',NS).get('name');dest=by_name.get(name)
            if dest is None:raise ValueError('PowerPoint原生导入丢失图片对象')
            old_fill=pic.find('p:blipFill',NS);rid=old_fill.find('a:blip',NS).get(tag('r','embed'))
            if rid not in source_rels:raise ValueError('本次图片关系缺失')
            media=resolve_target(sp,source_rels[rid].get('Target'));payload=source[media];ext=media.rsplit('.',1)[-1];new_media='ppt/media/raw-'+digest(payload)+'.'+ext
            data[new_media]=payload;new_rid='rIdLossless'+str(index)
            while any(r.get('Id')==new_rid for r in rels):new_rid+='x'
            E.SubElement(rels,'{'+NS['rel']+'}Relationship',Id=new_rid,Type=REL+'image',Target=posixpath.relpath(new_media,posixpath.dirname(dp)))
            new_fill=copy.deepcopy(old_fill);new_fill.find('a:blip',NS).set(tag('r','embed'),new_rid)
            dest.replace(dest.find('p:blipFill',NS),new_fill)
            source_types=xml(source['[Content_Types].xml']);ctype=next((t.get('ContentType') for t in source_types if t.get('PartName')=='/'+media),None) or next(t.get('ContentType') for t in source_types if t.get('Extension')==ext)
            if not any(t.get('PartName')=='/'+new_media for t in types):E.SubElement(types,'{'+CT+'}Override',PartName='/'+new_media,ContentType=ctype)
        used={v for n in dr.iter() for k,v in n.attrib.items() if k.startswith('{'+NS['r']+'}')}
        for rel in list(rels):
            if rel.get('Type','').endswith('/image') and rel.get('Id') not in used:rels.remove(rel)
        data[dp]=serialize(dr);data[rels_path(dp)]=serialize(rels)
    pres=xml(data['ppt/presentation.xml']);pres.set('autoCompressPictures','0');data['ppt/presentation.xml']=serialize(pres);data['[Content_Types].xml']=serialize(types)
    kept={}
    def walk(part):
        if part in kept:return
        if part:kept[part]=data[part]
        rp=rels_path(part)
        if rp not in data:return
        kept[rp]=data[rp]
        for rel in xml(data[rp]):
            if rel.get('TargetMode')!='External':walk(resolve_target(part,rel.get('Target')))
    walk('')
    for t in list(types):
        if t.get('PartName') and t.get('PartName').lstrip('/') not in kept:types.remove(t)
    kept['[Content_Types].xml']=serialize(types);save_package(kept,output)
    reopened=Presentation(str(output))
    if len(reopened.slides)!=len(pages):raise ValueError('拼装页数不一致')
    final=read_package(output);checked=0
    def named(root):
        inventory={e['id']:e for e in shape_inventory(root)};result={}
        for node in root.iter():
            if E.QName(node).localname not in SHAPES:continue
            prop=node.find('.//p:cNvPr',NS)
            if prop is not None:result[prop.get('name')]=inventory[shape_id(node)]
        return result
    for number,path in enumerate(pages,1):
        source=read_package(path);sp,_=slide_part(source,1);dp,_=slide_part(final,number)
        before=named(xml(source[sp]));after=named(xml(final[dp]))
        if set(before)!=set(after):raise ValueError('原生导入增减了组件对象')
        for name,old in before.items():
            new=after[name]
            if old['kind']!=new['kind'] or abs(old['rotationDegrees']-new['rotationDegrees'])>.01:raise ValueError('原生导入改变了组件对象类型或旋转')
            if old['boxPx'] and (not new['boxPx'] or any(abs(a-b)>.03 for a,b in zip(old['boxPx'],new['boxPx']))):raise ValueError('原生导入改变了组件位置或大小：'+name)
            checked+=1
    return {'slides':len(pages),'sha256':digest(output.read_bytes()),'nativeImport':True,'geometryPreserved':True,'objectsChecked':checked,'nativeTextShapes':sum(len(slide._element.xpath('.//p:sp/p:txBody')) for slide in reopened.slides)}
