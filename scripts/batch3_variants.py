"""双因素作用、四栏内容和双侧汇入中心的授权形式。"""
import copy
from lxml import etree as E

from native_pptx import NS,tag,shape_inventory,set_box
from component_variants import find_shape,replace_text,insert_panel_image


def native_box(ident,box,theme,text=False):
    node=E.Element(tag("p","sp"))
    nv=E.SubElement(node,tag("p","nvSpPr"))
    E.SubElement(nv,tag("p","cNvPr"),id=str(ident),name=f"content-choice-{ident}")
    E.SubElement(nv,tag("p","cNvSpPr"),**({"txBox":"1"} if text else {}))
    E.SubElement(nv,tag("p","nvPr"))
    props=E.SubElement(node,tag("p","spPr"))
    trans=E.SubElement(props,tag("a","xfrm"))
    E.SubElement(trans,tag("a","off"),x="0",y="0")
    E.SubElement(trans,tag("a","ext"),cx="0",cy="0")
    geom=E.SubElement(props,tag("a","prstGeom"),prst="rect")
    E.SubElement(geom,tag("a","avLst"))
    if text:
        E.SubElement(props,tag("a","noFill"))
    else:
        fill=E.SubElement(props,tag("a","solidFill"))
        E.SubElement(fill,tag("a","srgbClr"),val=theme["palette"]["surface"])
    line=E.SubElement(props,tag("a","ln"),w="6350")
    if text:
        E.SubElement(line,tag("a","noFill"))
    else:
        fill=E.SubElement(line,tag("a","solidFill"))
        E.SubElement(fill,tag("a","srgbClr"),val=theme["palette"]["primary"])
    if text:
        body=E.SubElement(node,tag("p","txBody"))
        E.SubElement(body,tag("a","bodyPr"))
        E.SubElement(body,tag("a","lstStyle"))
    set_box(node,*box)
    return node


def bounds(root,ids):
    elements={s["id"]:s for s in shape_inventory(root)}
    boxes=[elements[i]["boxPx"] for i in ids]
    x=min(b[0] for b in boxes);y=min(b[1] for b in boxes)
    return [x,y,max(b[0]+b[2] for b in boxes)-x,max(b[1]+b[3] for b in boxes)-y]


def evidence_text(root,areas,theme,start=58000):
    region_map,slots={ },[]
    tree=root.find("p:cSld/p:spTree",NS)
    for index,area in enumerate(areas):
        x,y,w,h=bounds(root,area["boxFromIds"])
        for ident in area["removeIds"]:
            node=find_shape(root,ident);node.getparent().remove(node)
        ident=start+index*100
        title_h=30 if h<160 else 42
        frame=native_box(ident,[x,y,w,h],theme)
        title=native_box(ident+1,[x+8,y+6,w-16,title_h],theme,True)
        body=native_box(ident+2,[x+8,y+title_h+8,w-16,h-title_h-18],theme,True)
        size=area.get("fontPt",16)
        replace_text(title,[area["title"]],size,theme["palette"]["primary"],theme,bold=True)
        replace_text(body,area["text"],size,theme["palette"]["text"],theme)
        for node in (frame,title,body):tree.append(node)
        region_map[area["regionId"]]=[ident,ident+1,ident+2]
        slots.extend([{"key":f'{area["regionId"]}Title',"shapeIds":[ident+1],"kind":"text","role":"module-title","regionId":area["regionId"],"replace":"此区域的短标题"},
                      {"key":f'{area["regionId"]}Text',"shapeIds":[ident+2],"kind":"rich-text","role":"summary-text","regionId":area["regionId"],"replace":"本次内容的相关说明、总结或明确的换图提示"}])
    return region_map,slots


def apply_recipe(root,data,part,component,recipe,theme):
    result={"regionShapeIds":{},"addedSlots":[],"replacedSlotKeys":[]}
    kind=recipe["kind"]
    if kind=="factor-positive":
        choice=component["factorPositive"]
        replace_text(find_shape(root,64),["两组促进因素通过不同路径形成共同作用，具体方向与强度由本次内容和证据确定"],22,theme["palette"]["text"],theme,bold=True)
        for ident,text in zip(choice["names"],["促进因素","支撑因素"]):
            replace_text(find_shape(root,ident),[text],18,theme["palette"]["white"],theme,bold=True)
        for index,ident in enumerate(choice["upperIds"]):
            replace_text(find_shape(root,ident),[f"条件{index+1}"],14,theme["palette"]["text"],theme,bold=True)
        for index,ident in enumerate(choice["lowerIds"]):
            replace_text(find_shape(root,ident),[f"条件{index+5}"],14,theme["palette"]["text"],theme,bold=True)
        for ident,text in zip(choice["outcomeIds"],["支持","促进","协同"]):
            replace_text(find_shape(root,ident),[text],20,theme["palette"]["primary"],theme,bold=True)
        result["regionShapeIds"],result["addedSlots"]=evidence_text(root,choice["evidenceAreas"],theme)
        result["replacedSlotKeys"]=choice["replacedSlotKeys"]
        result["sampleContent"]="正向作用是版式占位；已移除负面案例图证，需绑定支持本次作用方向的真实内容"
    elif kind=="bilateral-bottom-text":
        area=component["evidenceTextChoices"]
        result["regionShapeIds"],result["addedSlots"]=evidence_text(root,[area],theme)
        result["replacedSlotKeys"]=area["replacedSlotKeys"]
        result["sampleContent"]="下方总结文字是占位；上方左右相向汇入中心的结构保持"
    elif kind=="four-categories":
        replace_text(find_shape(root,52),["四类内容共享同一主题，分别展示特征与对应响应"],16,theme["palette"]["text"],theme,bold=True,align="l")
        for index,column in enumerate(component["fourChoices"]["columns"]):
            replace_text(find_shape(root,column["labelId"]),[f"类别{['一','二','三','四'][index]}"],13,theme["palette"]["primary"],theme,bold=True,align="l")
        result["sampleContent"]="时期标签已换为并行类别占位；示例正文仍用于说明文字层级，正式制作时需逐栏替换"
    elif kind=="four-content":
        columns=component["fourChoices"]["columns"]
        modes=recipe["modes"]
        if len(modes)!=4 or any(m not in {"text","image-text"} for m in modes):raise ValueError("四栏必须各选择text或image-text")
        images=recipe.get("images",{})
        if not set(images)<={"1","2","3","4"}:raise ValueError("四栏图片键必须为1至4")
        seed=copy.deepcopy(root)
        result["addedImages"]=[]
        for index,(column,mode) in enumerate(zip(columns,modes)):
            asset=images.get(str(index+1))
            if asset and mode!="image-text":raise ValueError("文字形式不能同时指定图片")
            if mode==column["defaultMode"] and not asset:continue
            if mode=="text":
                area={"regionId":f"column-{index+1}","removeIds":column["pictureIds"],"boxFromIds":column["boxFromIds"],
                      "title":"补充要点","text":["在此填写本栏的补充解释","或与上文对应的条件、证据"],"fontPt":10}
                mapping,slots=evidence_text(root,[area],theme,57000+index*100)
                # 保持整个栏目可定位，而不是只定位替换后的局部。
                outer=next(r["shapeIds"] for r in component["regions"] if r["id"]==area["regionId"])
                result["regionShapeIds"][area["regionId"]]=[i for i in outer if i not in column["pictureIds"]]+mapping[area["regionId"]]
                result["addedSlots"].extend(slots)
            else:
                if not asset:raise ValueError(f"第{index+1}栏需要本次真实图片")
                box=bounds(root,column["boxFromIds"])
                removed=column["detailIds"]+column["pictureIds"]
                for ident in removed:
                    node=find_shape(root,ident);node.getparent().remove(node)
                ident=55000+index*100
                result["addedImages"].append(insert_panel_image(root,data,part,seed,asset,box,ident,seed_shape_id=102))
                outer=next(r["shapeIds"] for r in component["regions"] if r["id"]==f"column-{index+1}")
                result["regionShapeIds"][f"column-{index+1}"]=[i for i in outer if i not in removed]+[ident]
                result["addedSlots"].append({"key":f"column{index+1}Image","shapeIds":[ident],"kind":"image","role":"local-evidence","regionId":f"column-{index+1}","replace":"本次真实辅助图，等比适配原有局部证据区"})
            result["replacedSlotKeys"].extend(column["detailSlotKeys"])
        result["sampleContent"]="四栏内部保留主标题、小标题及正文层级，局部证据区可选择相关图或补充文字"
    return result
