"""用户授权范围内的栏内变体；保留外部结构，不调用模型或外部服务。"""
from __future__ import annotations

import copy
import hashlib
import posixpath
from pathlib import Path
import zipfile

from lxml import etree as E

from native_pptx import (NS, SHAPES, xml, serialize, slide_part, shape_id,
                         shape_inventory, set_box, set_fonts, solid_text_color, tag, rels_path)


def find_shape(root, ident):
    return next(n for n in root.iter() if E.QName(n).localname in SHAPES and shape_id(n) == ident)


def clone_ids(node, start):
    mapping = {}
    for offset, info in enumerate(node.findall(".//p:cNvPr", NS)):
        old = int(info.get("id"))
        mapping[old] = start + offset
        info.set("id", str(start + offset))
        info.set("name", f"variant-object-{start+offset}")
    for item in list(node.xpath('.//*[local-name()="creationId"]')):
        item.getparent().remove(item)
    for item in node.findall(".//p:custDataLst", NS):
        item.getparent().remove(item)
    for item in node.findall(".//a:stCxn", NS) + node.findall(".//a:endCxn", NS):
        if int(item.get("id")) in mapping:
            item.set("id", str(mapping[int(item.get("id"))]))
    return mapping


def replace_text(node, texts, font_pt, color, theme, *, bold=False, align="ctr", bullet=None):
    body = node.find("p:txBody", NS)
    if body is None:
        raise ValueError(f"对象{shape_id(node)}不是文字槽位")
    props = body.find("a:bodyPr", NS)
    props.set("anchor", "ctr" if align == "ctr" else "t")
    props.set("wrap", "square")
    for name in ("lIns", "rIns"):
        props.set(name, "57150")
    for name in ("tIns", "bIns"):
        props.set(name, "28575")
    for item in list(props):
        if E.QName(item).localname in {"spAutoFit", "normAutofit", "noAutofit"}:
            props.remove(item)
    props.append(E.Element(tag("a", "noAutofit")))
    for paragraph in list(body.findall("a:p", NS)):
        body.remove(paragraph)
    for text in texts:
        paragraph = E.SubElement(body, tag("a", "p"))
        paragraph_props = E.SubElement(paragraph, tag("a", "pPr"), algn=align)
        if bullet:
            paragraph_props.set("marL", "190500")
            paragraph_props.set("indent", "-190500")
            E.SubElement(paragraph_props, tag("a", "buFont"), typeface=theme["typography"]["family"])
            E.SubElement(paragraph_props, tag("a", "buChar"), char=bullet)
        run = E.SubElement(paragraph, tag("a", "r"))
        pr = E.SubElement(run, tag("a", "rPr"), sz=str(round(font_pt*100)), b="1" if bold else "0")
        solid_text_color(pr, color)
        set_fonts(pr, theme["typography"]["family"])
        E.SubElement(run, tag("a", "t")).text = text


def add_source_copies(root, component):
    for rule in component.get("nativeAdditions", []):
        original = find_shape(root, rule["sourceShapeId"])
        node = copy.deepcopy(original)
        box = next(s["boxPx"] for s in shape_inventory(root) if s["id"] == rule["sourceShapeId"])
        clone_ids(node, rule["newShapeId"])
        set_box(node, box[0]+rule.get("translateX", 0), box[1]+rule.get("translateY", 0), box[2], box[3])
        root.find("p:cSld/p:spTree", NS).append(node)


def insert_panel_image(root,data,part,seed,asset_path,box,ident,seed_shape_id=117):
    from PIL import Image
    asset_path=Path(asset_path)
    payload=asset_path.read_bytes()
    with Image.open(asset_path) as image:
        pixels=image.size
        extension={"PNG":"png","JPEG":"jpeg"}.get(image.format)
    if not extension:
        raise ValueError("栏内图片当前支持PNG与JPEG；不改写原图像素")
    sha=hashlib.sha256(payload).hexdigest()
    media=f"ppt/media/variant-{ident}-{sha[:12]}.{extension}"
    data[media]=payload
    relation_path=rels_path(part)
    relations=xml(data[relation_path])
    relation_id=f"rIdVariant{ident}"
    if any(r.get("Id")==relation_id for r in relations):
        raise ValueError("图片关系ID冲突")
    E.SubElement(relations,"{"+NS["rel"]+"}Relationship",Id=relation_id,
        Type=NS["r"]+"/image",Target=posixpath.relpath(media,posixpath.dirname(part)))
    data[relation_path]=serialize(relations)
    types=xml(data["[Content_Types].xml"])
    if not any(n.get("Extension")==extension for n in types):
        node=E.Element("{http://schemas.openxmlformats.org/package/2006/content-types}Default",Extension=extension,
                       ContentType="image/png" if extension=="png" else "image/jpeg")
        types.insert(0,node)
    data["[Content_Types].xml"]=serialize(types)
    picture=copy.deepcopy(find_shape(seed,seed_shape_id))
    clone_ids(picture,ident)
    blip=picture.find("p:blipFill/a:blip",NS)
    blip.set(tag("r","embed"),relation_id)
    blip.attrib.pop(tag("r","link"),None)
    for crop in picture.findall("p:blipFill/a:srcRect",NS):
        crop.getparent().remove(crop)
    x,y,w,h=box
    scale=min(w/pixels[0],h/pixels[1])
    actual=[x+(w-pixels[0]*scale)/2,y+(h-pixels[1]*scale)/2,pixels[0]*scale,pixels[1]*scale]
    set_box(picture,*actual)
    root.find("p:cSld/p:spTree",NS).append(picture)
    return {"shapeId":ident,"sourceAssetName":asset_path.name,"sourceAssetSha256":sha,
            "pixels":list(pixels),"frameBoxPx":box,"actualBoxPx":actual}


def nested_panels(root, component, recipe, theme, library, data, part):
    rules = component["panelChoices"]
    modes = recipe["panelModes"]
    if len(modes) != 3 or any(m not in {"image-text", "nested-text"} for m in modes):
        raise ValueError("三栏模式必须各为image-text或nested-text")
    if recipe.get("connectionMode", "related") not in {"related", "parallel"}:
        raise ValueError("栏间关系只能在已授权的并列或相关中选择")
    seed_path = library / "templates/BW-TRIPLE-PROGRESS-01.pptx"
    with zipfile.ZipFile(seed_path) as z:
        seed_data = {n:z.read(n) for n in z.namelist()}
    seed_part, _ = slide_part(seed_data, 1)
    seed = xml(seed_data[seed_part])
    tree = root.find("p:cSld/p:spTree", NS)
    region_ids, added_slots, replaced_slots, added_images = {}, [], [], []
    for index, (column, mode) in enumerate(zip(rules["columns"], modes)):
        asset=recipe.get("images",{}).get(str(index+1))
        if asset and mode!="image-text":
            raise ValueError(f"第{index+1}栏选了无图，但同时提供了图片")
        if mode == column["defaultMode"] and not asset:
            continue
        if mode == "image-text":
            if not asset:
                raise ValueError(f"第{index+1}栏原本无图；切换图文需先提供本次真实图片，不能借用别栏图片")
            removed=column["imageShapeIds"]+column.get("nestedShapeIds",[])
            for ident in removed:
                node=find_shape(root,ident);node.getparent().remove(node)
            replaced_slots.extend(column["imageSlotKeys"]+column.get("nestedSlotKeys",[]))
            ident=54000+index*100
            added_images.append(insert_panel_image(root,data,part,seed,asset,column["innerBox"],ident))
            region_ids[column["regionId"]]=[i for i in column["retainedShapeIds"] if i not in removed]+[ident]
            added_slots.append({"key":f"column{index+1}Image","shapeIds":[ident],"kind":"image","role":"local-evidence",
                                "regionId":column["regionId"],"replace":"本次内容提供的真实图片，等比适配固定图框"})
            continue
        for ident in column["imageShapeIds"]:
            node = find_shape(root, ident)
            node.getparent().remove(node)
        replaced_slots.extend(column["imageSlotKeys"])
        x, y, width, available_h = column["innerBox"]
        height = min(160, available_h)
        y += (available_h-height)/2
        background = copy.deepcopy(find_shape(seed, 123))
        heading = copy.deepcopy(find_shape(seed, 126))
        content = copy.deepcopy(find_shape(seed, 128))
        first = 50000 + index*100
        for offset, node in enumerate((background, heading, content)):
            clone_ids(node, first+offset)
            tree.append(node)
        set_box(background, x, y, width, height)
        props=background.find("p:spPr",NS)
        for fill in list(props):
            if E.QName(fill).localname in {"noFill","solidFill","gradFill","pattFill","grpFill","blipFill"}:
                props.remove(fill)
        fill=E.Element(tag("a","solidFill"))
        E.SubElement(fill,tag("a","srgbClr"),val=theme["palette"]["surface"])
        position=next((i for i,n in enumerate(props) if E.QName(n).localname in {"ln","effectLst","effectDag","scene3d","sp3d","extLst"}),len(props))
        props.insert(position,fill)
        line=props.find("a:ln",NS)
        if line is not None:
            solid_text_color(line,theme["palette"]["primary"])
        set_box(heading, x+8, y+10, width-16, 38)
        set_box(content, x+10, y+height-65, width-20, 52)
        replace_text(heading, ["条件、要点或特征"], 14, theme["palette"]["text"], theme, bold=True)
        replace_text(content, ["本栏的补充说明", "或局部判断"], 16, theme["palette"]["emphasisRed"], theme, bold=True)
        region_ids[column["regionId"]] = column["retainedShapeIds"] + [first, first+1, first+2]
        for suffix, ident, role, text in [("Title",first+1,"nested-title","本栏条件、要点或特征"),
                                         ("Body",first+2,"nested-body","本栏补充说明或局部判断")]:
            added_slots.append({"key":f"column{index+1}Nested{suffix}","shapeIds":[ident],
                "kind":"text","role":role,"regionId":column["regionId"],"replace":text})
    if recipe.get("connectionMode") == "parallel":
        for ident in rules["connectionShapeIds"]:
            node = find_shape(root, ident)
            node.getparent().remove(node)
        region_ids["column-links"] = []
    return {"regionShapeIds":region_ids,"addedSlots":added_slots,"replacedSlotKeys":replaced_slots,
            "panelModes":modes,"addedImages":added_images,"sampleContent":"新内嵌框内为明确的版式占位文字，正式制作时需绑定本次内容"}


def decompose_list(root, component, recipe, theme):
    bullet = recipe.get("bullet", "◆")
    if bullet not in {"◆", "•", "▪", "–"}:
        raise ValueError("列表标记需使用已登记的普通无序符号")
    if len(recipe["groupTitles"]) != 2 or len(recipe["groupBodies"]) != 2:
        raise ValueError("当前列表原型承载两组主序号，每组包含下级无序条目")
    tree = root.find("p:cSld/p:spTree", NS)
    original_head = copy.deepcopy(find_shape(root, 18))
    original_background = copy.deepcopy(find_shape(root, 12))
    original_body = copy.deepcopy(find_shape(root, 13))
    for ident in (42,41):
        node = find_shape(root, ident)
        node.getparent().remove(node)
    added_ids, title_ids, body_ids, number_ids = [], [], [], []
    for index in range(2):
        head = copy.deepcopy(original_head)
        background = copy.deepcopy(original_background)
        body = copy.deepcopy(original_body)
        first = 52000+index*100
        mapping = clone_ids(head, first)
        clone_ids(background, first+20)
        clone_ids(body, first+21)
        y = 242+index*170
        set_box(head, 933, y, 312.1619, 42.0063)
        set_box(background, 940, y+46, 310, 110)
        set_box(body, 947, y+49, 296, 104)
        replace_text(find_shape(head, mapping[11]), [recipe["groupTitles"][index]], 20,
                     theme["palette"]["primary"], theme, bold=True, align="l")
        find_shape(head, mapping[11]).find("p:txBody/a:bodyPr",NS).set("lIns","571500")
        replace_text(find_shape(head, mapping[16]), [str(index+1)], 16,
                     theme["palette"]["white"], theme, bold=True)
        number_props = find_shape(head,mapping[16]).find("p:txBody/a:bodyPr",NS)
        for attribute in ("lIns","rIns","tIns","bIns"):
            number_props.set(attribute,"0")
        replace_text(body, recipe["groupBodies"][index], 16,
                     theme["palette"]["text"], theme, align="l", bullet=bullet)
        for node in (head,background,body):
            tree.append(node)
        added_ids.extend([first,first+20,first+21])
        title_ids.append(mapping[11]); number_ids.append(mapping[16]); body_ids.append(first+21)
    # 左侧同样允许切换无序标记，序号层级保持不变。
    for ident in (13,20):
        for marker in find_shape(root,ident).findall(".//a:buChar",NS):
            marker.set("char",bullet)
            paragraph_props=marker.getparent()
            font=paragraph_props.find("a:buFont",NS)
            if font is None:
                font=E.Element(tag("a","buFont"))
                paragraph_props.insert(paragraph_props.index(marker),font)
            font.set("typeface",theme["typography"]["family"])
    return {"regionShapeIds":{"column-3":[30]+added_ids},
        "replacedSlotKeys":["optionalProcess","optionalProcessConclusion"],
        "addedSlots":[
            {"key":"thirdListTitles","shapeIds":title_ids,"kind":"text-group","role":"module-title","regionId":"column-3","replace":"两组主层级小标题"},
            {"key":"thirdListNumbers","shapeIds":number_ids,"kind":"generated-text","role":"hierarchy-number","regionId":"column-3","replace":"主层级序号"},
            {"key":"thirdListBodies","shapeIds":body_ids,"kind":"rich-text-group","role":"hierarchy-body","regionId":"column-3","replace":"每个主层级下的无序条目"}],
        "sourceSupportShapeIds":recipe["sourceSupportShapeIds"],"bullet":bullet,
        "sampleContent":"将源流程的四个已有步骤分组为检测与配比、施用与监测，没有补充新的实验结论"}


def build_variant(base_path, component, variant, theme, library, output):
    if component.get("sourceKind") == "reference-reconstruction":
        from reference_layouts import variant_form
        return variant_form(component, variant, theme, output)
    with zipfile.ZipFile(base_path) as z:
        data = {n:z.read(n) for n in z.namelist()}
    part,_ = slide_part(data,1)
    root = xml(data[part])
    if component.get('pageFrame'):
        from page_frame import restore_content_coordinates
        restore_content_coordinates(root,component)
    recipe = variant["recipe"]
    allowed={v["recipe"]["kind"] for v in component.get("variants",[])}|set(component.get("customRecipeKinds",[]))
    if recipe["kind"] not in allowed:
        raise ValueError("所选形式不属于此人工指定组件的授权范围")
    if recipe["kind"] == "panel-formats":
        detail = nested_panels(root,component,recipe,theme,library,data,part)
    elif recipe["kind"] == "decompose-list":
        detail = decompose_list(root,component,recipe,theme)
    elif recipe["kind"] == "column-labels":
        labels = recipe["labels"]
        slot = next(s for s in component["slots"] if s["key"] == "columnLabels")
        if len(labels) != len(slot["shapeIds"]):
            raise ValueError("栏名数量与三栏结构不符")
        for ident,label in zip(slot["shapeIds"],labels):
            replace_text(find_shape(root,ident),[label],20,theme["palette"]["primary"],theme,bold=True)
        detail = {"regionShapeIds":{},"addedSlots":[],"replacedSlotKeys":[],"columnLabels":labels}
    elif recipe["kind"] in {"factor-positive","four-content","four-categories","bilateral-bottom-text"}:
        from batch3_variants import apply_recipe
        detail=apply_recipe(root,data,part,component,recipe,theme)
    elif recipe["kind"] in {"additive-three","guided-labels"}:
        from batch6_components import apply_recipe
        detail=apply_recipe(root,component,recipe,theme)
    else:
        raise ValueError(f"未授权的组件变体：{recipe['kind']}")
    if component.get('pageFrame'):
        from page_frame import apply_page_frame
        detail['pageFrame']=apply_page_frame(root,component,theme)
    data[part] = serialize(root)
    output = Path(output)
    output.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(output,"w",zipfile.ZIP_DEFLATED) as z:
        for name in sorted(data):
            info=zipfile.ZipInfo(name,date_time=(2026,11,1,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED
            z.writestr(info,data[name])
    detail["templateElements"] = shape_inventory(root)
    detail["id"] = variant["id"]
    detail["name"] = variant["name"]
    detail["recipe"] = recipe
    return detail
