"""核对组件来源、原生对象、依赖关系和目录等距/激活状态，不代替内容验收。"""
from pathlib import Path
from datetime import date
import json
import zipfile
import hashlib

from PIL import Image
from pptx import Presentation

from native_pptx import (NS, xml, rels_path, resolve_target, slide_part, shape_inventory,
                         build_single_slide, toc_positions)
from component_variants import build_variant, find_shape
from suite_toc import ROW_ROLES, ROW_START, suite_number

ROOT = Path(__file__).resolve().parents[1]
LIB = ROOT / "library"
count = 0


def check(condition, message):
    global count
    if not condition:
        raise AssertionError(message)
    count += 1


def read(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def validate_package(path):
    with zipfile.ZipFile(path) as z:
        check(z.testzip() is None, f"压缩包损坏：{path.name}")
        data = {n: z.read(n) for n in z.namelist()}
    for name, content in data.items():
        if name.endswith(".rels"):
            part = "" if name == "_rels/.rels" else name.replace("/_rels/", "/")[:-5]
            for relation in xml(content):
                if relation.get("TargetMode") != "External":
                    target = resolve_target(part, relation.get("Target", ""))
                    check(target in data, f"部件缺失：{name} -> {target}")
        elif name.endswith(".xml"):
            root = xml(content)
            relations = rels_path(name)
            ids = {r.get("Id") for r in xml(data[relations])} if relations in data else set()
            for node in root.iter():
                for key, value in node.attrib.items():
                    if key.startswith("{" + NS["r"] + "}"):
                        check(value in ids, f"对象关系失效：{name} -> {value}")
    part, _ = slide_part(data, 1)
    root = xml(data[part])
    ids = [n.get("id") for n in root.findall(".//p:cNvPr", NS)]
    check(len(ids) == len(set(ids)), f"原生对象ID重复：{path.name}")
    presentation = Presentation(path)
    check(len(presentation.slides) == 1, f"组件应为单页：{path.name}")
    check((presentation.slide_width, presentation.slide_height) == (12192000, 6858000), f"画幅变化：{path.name}")
    return root


def validate_toc(root, amount, active, theme):
    elements = {s["id"]: s for s in shape_inventory(root)}
    frames = [elements[20000 + i * 10]["boxPx"] for i in range(amount)]
    rule = theme["toc"]
    gaps = [frames[0][1] - rule["panelY"]]
    gaps += [frames[i][1] - frames[i-1][1] - frames[i-1][3] for i in range(1, amount)]
    gaps += [rule["panelY"] + rule["panelHeight"] - frames[-1][1] - frames[-1][3]]
    check(max(gaps) - min(gaps) < .001, f"{amount}项目录上下与行间不等距：{gaps}")
    check(max(b[3] for b in frames) - min(b[3] for b in frames) < .001, "目录行高不一致")
    check(min(gaps) >= rule["minimumGap"] - .001, "目录净间距低于可读下限")
    check(frames[0][3] >= rule["minimumRowHeight"] - .001, "目录行高低于可读下限")
    blue_labels = 0
    blue_numbers = 0
    for i in range(amount):
        for offset, size in [(1, theme["typography"]["tocTitlePt"]), (2, theme["typography"]["tocNumberPt"])]:
            ident = 20000 + i * 10 + offset
            node = next(n for n in root.findall("p:cSld/p:spTree/p:sp", NS) if n.find("p:nvSpPr/p:cNvPr", NS).get("id") == str(ident))
            run = node.find("p:txBody/a:p/a:r/a:rPr", NS)
            actual = run.find("a:solidFill/a:srgbClr", NS).get("val")
            expected = theme["palette"]["primary" if i == active else "inactive"]
            check(actual == expected, f"第{i+1}项激活/浅灰状态不符")
            check(int(run.get("sz")) / 100 == size, "目录在适配数量时改变了字号")
            if actual == theme["palette"]["primary"]:
                if offset == 1: blue_labels += 1
                else: blue_numbers += 1
    check(blue_labels == blue_numbers == 1, "目录不是唯一高亮")


def validate_suite_toc(root, labels, active, layout):
    elements = {s["id"]: s for s in shape_inventory(root)}
    positions = toc_positions(len(labels), {"toc": layout["spacing"]})
    for index, (label, row) in enumerate(zip(labels, positions)):
        ids = [ROW_START + index * 10 + i for i in range(len(ROW_ROLES))]
        boxes = [elements[i]["boxPx"] for i in ids]
        top = min(b[1] for b in boxes)
        bottom = max(b[1] + b[3] for b in boxes)
        check(abs(top - row["y"]) < .001 and abs(bottom - top - row["height"]) < .001,
              "套系目录的行高或净间距偏离登记公式")
        check(elements[ids[1]]["text"] == label, "套系目录章名不符")
        check(elements[ids[2]]["text"] == suite_number(index, layout["numberStyle"]), "套系目录编号样式不符")
        for ident, size in [(ids[1], layout["labelPt"]), (ids[2], layout["numberPt"])]:
            check(elements[ident]["fontSizesPt"] == [size], "套系目录在重排时改变了字号")
        label_run = find_shape(root, ids[1]).find("p:txBody/a:p/a:r/a:rPr/a:solidFill/a:srgbClr", NS)
        check(label_run.get("val") == layout["labelColors"][0 if index == active else 1], "目录高亮状态不符")


def validate_variant(component, detail, base_root, theme):
    target=LIB/detail["templatePath"]
    check(digest(target)==detail["templateSha256"],"变体与登记哈希不符")
    root=validate_package(target)
    ids={s["id"] for s in shape_inventory(root)}
    for binding in detail["slotBindings"]:
        check(set(binding["templateShapeIds"])<=ids,f"变体槽位失效：{component['id']}/{detail['id']}/{binding['key']}")
    for region_ids in detail["regionShapeIds"].values():
        check(set(region_ids)<=ids,"变体功能分区绑定失效")
    recipe=detail["recipe"]
    if recipe["kind"]=="panel-formats":
        for column,mode in zip(component["panelChoices"]["columns"],recipe["panelModes"]):
            if mode=="nested-text" and column["defaultMode"]=="image-text":
                check(not(set(column["imageShapeIds"]) & ids),"无图小矩形变体仍保留原图片对象")
        if recipe.get("connectionMode")=="parallel":
            check(not(set(component["panelChoices"]["connectionShapeIds"]) & ids),"并列变体仍保留栏间箭头")
        protected={"BW-TRIPLE-PROGRESS-01":[19,115,20,134,72,135,86,139,88,140,85,87],
                   "BW-TRIPLE-COMMON-01":[45,49,50,53,63,64,65]}[component["id"]]
        from lxml import etree as E
        for ident in protected:
            check(E.tostring(find_shape(root,ident),method="c14n")==E.tostring(find_shape(base_root,ident),method="c14n"),"栏内变体修改了底部递进/共同归纳结构")
        for s in detail["templateElements"]:
            if s["id"] in (50000,50100,50200):
                frame=find_shape(root,s["id"])
                check(frame.find("p:spPr/a:ln/a:solidFill/a:srgbClr",NS).get("val")==theme["palette"]["primary"],"内嵌矩形边框不是统一蓝色")
    elif recipe["kind"]=="decompose-list":
        check(42 not in ids and 41 not in ids,"列表变体没有移除可选流程")
        check(len(next(b for b in detail["slotBindings"] if b["key"]=="thirdListNumbers")["textShapeIds"])==2,"第三栏主层级序号缺失")
        for ident in (13,20):
            for marker in find_shape(root,ident).findall(".//a:buChar",NS):
                check(marker.get("char")==recipe["bullet"],"无序符号未按所选方案替换")
                check(marker.getparent().find("a:buFont",NS).get("typeface")==theme["typography"]["family"],"更换符号后仍继承旧符号字体")
    elif recipe["kind"]=="column-labels":
        expected=recipe["labels"]
        slot=next(s for s in component["slots"] if s["key"]=="columnLabels")
        actual=["".join(n.text or "" for n in find_shape(root,i).findall(".//a:t",NS)) for i in slot["shapeIds"]]
        check(actual==expected,"路线或方法标签没有对应替换")
    elif recipe["kind"]=="factor-positive":
        text=''.join(n.text or '' for n in root.findall('.//a:t',NS))
        check('促进因素' in text and '支撑因素' in text,"正向作用名称未替换")
        check('干旱' not in text and not root.findall('.//p:pic',NS),"正向占位仍沿用负面标签或原图证")
        from lxml import etree as E
        for ident in range(82,94):
            if ident==92:continue
            check(E.tostring(find_shape(root,ident),method='c14n')==E.tostring(find_shape(base_root,ident),method='c14n'),"正向选择改变了因素到结果的连线")
    elif recipe["kind"]=="four-categories":
        actual=[''.join(n.text or '' for n in find_shape(root,c['labelId']).findall('.//a:t',NS)) for c in component['fourChoices']['columns']]
        check(actual==['类别一','类别二','类别三','类别四'],"四栏仍强制使用日期标签")
    elif recipe["kind"]=="four-content":
        check(len(recipe['modes'])==4,"四栏形式数量不符")
        if recipe['modes']==['text']*4:check(102 not in ids,"四栏文字形式仍保留局部图片")
        from lxml import etree as E
        for c in component['fourChoices']['columns']:
            for ident in [c['labelId'],c['titleId']]+c['responseIds']:
                check(E.tostring(find_shape(root,ident),method='c14n')==E.tostring(find_shape(base_root,ident),method='c14n'),"四栏局部图文选择改变了栏目主层级或响应")
    elif recipe["kind"]=="bilateral-bottom-text":
        check(not({175,177,181,183}&ids),"底部文字形式未替换图组")
        from lxml import etree as E
        for ident in [124,126,128,133,13]:
            check(E.tostring(find_shape(root,ident),method='c14n')==E.tostring(find_shape(base_root,ident),method='c14n'),"底部内容选择改变了上方汇入中心的逻辑")
    preview=LIB/detail["previewPath"]
    with Image.open(preview) as image:
        check(image.size==(1280,720),"变体实际预览画幅不符")
    return {"id":detail["id"],"template":detail["templatePath"],"passed":True}


def check_reference_layout(component, detail, root):
    slots = {binding["key"]: binding for binding in detail["slotBindings"]}
    figure = component["referenceNumber"] - 16
    if figure == 1:
        for key in ("topics","issues","issuePhotos","points","caseNames","casePhotos"):
            check(len(slots[key]["boxesPx"]) == 4, f"四类问题与案例槽位不完整：{key}")
        for upper,lower in zip(slots["issuePhotos"]["boxesPx"],slots["casePhotos"]["boxesPx"]):
            check(upper[1]+upper[3] < lower[1], "案例照片应在对应问题下方")
        markers=root.xpath('.//p:sp[p:nvSpPr/p:cNvPr[@name="typical-case-marker"]]',namespaces=NS)
        check(len(markers)==4,"四栏案例蓝色标识缺失")
        for ident in slots["points"]["templateShapeIds"]:
            props=find_shape(root,ident).findall("p:txBody/a:p/a:pPr",NS)
            check(len(props)==3 and all(p.find("a:buClr/a:srgbClr",NS).get("val")=="0046A5" and p.find("a:buChar",NS) is not None for p in props),"四栏上部项目符号没有保留蓝色")
    elif figure == 2:
        for i in range(3):
            boxes=[slots[key]["boxesPx"][i] for key in ("overviews","figures","captions","analysis","sources")]
            boxes.append(slots["takeaway"]["boxesPx"][0])
            check(all(a[1]+a[3]<=b[1]+.01 for a,b in zip(boxes,boxes[1:])),"三栏图证的阅读顺序或底部总述位置有误")
    elif figure == 3:
        for key in ("routeNames","methodLabels","means","caseLabels","examples","caseNames","listNumbers","points","localResults"):
            boxes=slots[key]["boxesPx"];half=len(boxes)//2
            for left,right in zip(boxes[:half],boxes[half:]):
                check(left[1:]==right[1:] and right[0]-left[0]==634,"案例双栏未按同一几何结构排列")
        for page in [root]+[validate_package(LIB/v["templatePath"]) for v in detail.get("variants",[])]:
            markers=page.xpath('.//p:sp[p:nvSpPr/p:cNvPr[@name="typical-case-marker"]]',namespaces=NS)
            check(len(markers)==2,"双路径案例蓝色标识缺失")
            for ident in slots["routeNames"]["templateShapeIds"]:
                check(find_shape(page,ident).find("p:spPr/a:gradFill",NS) is not None,"双路径标题渐变底板缺失")
    elif figure == 4:
        dividers=root.findall(".//p:cxnSp/p:spPr/a:ln/a:prstDash",NS)
        check(len(dividers)==1 and dividers[0].get("val")=="dash","缺少中央虚线")
        check(len(slots["axes"]["templateShapeIds"])==2,"分析维度应为两项")
    elif figure == 5:
        body=find_shape(root,slots["bibliography"]["templateShapeIds"][0])
        entries=body.findall("p:txBody/a:p",NS)
        check(len(entries)==9 and all("".join(p.itertext()).startswith(f"[{i}]") for i,p in enumerate(entries,1)),"参考文献示例缺少连续方括号编号")
    elif figure == 6:
        alternate={s["key"] for s in detail["variants"][0]["slotBindings"]}
        check("photo" in slots and "photo" not in alternate and "outcome" in alternate,"两行动形式未将右侧现场图切换为效果")
    elif figure == 7:
        order=[slots[k]["boxesPx"][0][1] for k in ("caseName","context","actions","results")]
        check(order==sorted(order) and slots["scenePhoto"]["boxesPx"][0][0]>slots["actions"]["boxesPx"][0][0],"纵向案例链或现场图区域错误")
        for name in ("background-blue-white-frame","technique-blue-white-frame"):
            node=next(n for n in root.findall(".//p:sp",NS) if n.find("p:nvSpPr/p:cNvPr",NS).get("name")==name)
            check(node.find("p:spPr/a:solidFill/a:srgbClr",NS).get("val")=="FFFFFF","案例链外框未保留白填充")
            check(node.find("p:spPr/a:ln/a:solidFill/a:srgbClr",NS).get("val")=="0046A5","案例链缺少统一蓝色外框")
    elif figure in (8,9):
        check(not component.get("pageFrame") and {"pageTitle","presenter"}<=slots.keys(),"首尾专用版式缺少标题或署名")
        check("logo" not in slots and len(root.findall(".//p:pic",NS))==1,"首尾专用版式应仅保留底带纹理图")
        check(("thanks" in slots)==(figure==9),"结束语槽位与页面用途不符")
    elif figure==10:
        check(not root.findall(".//p:pic",NS),"纯文字四象限出现插图")
        alt=validate_package(LIB/detail["variants"][0]["templatePath"])
        for name in ("items1","items2","items3","items4"):
            ident=slots[name]["templateShapeIds"][0]
            check(len(find_shape(root,ident).findall(".//a:buAutoNum",NS))==2,"象限编号未采用原生列表")
            check(len(find_shape(alt,ident).findall(".//a:buChar",NS))==2,"无序变体未保留独立列表")
        for ident in slots["labels"]["templateShapeIds"]:
            for page in (root,alt):
                check(find_shape(page,ident).find("p:spPr/a:custGeom/a:pathLst/a:path/a:cubicBezTo",NS) is not None,"四象限标签的右上圆弧丢失")
    elif figure in (11,12):
        check(len(root.findall(".//p:pic",NS))==1 and len(root.findall(".//a:buChar",NS))>=6,"无框案例缺少证据图或分点分析")
        if figure==12:check(len(slots["steps"]["templateShapeIds"])==4,"实施流程缺少步骤")
        summary=find_shape(root,slots["summary"]["templateShapeIds"][0])
        stops=summary.findall("p:spPr/a:gradFill/a:gsLst/a:gs",NS)
        check([n.get("pos") for n in stops]==["0","50000","100000"],"案例总述的双向渐变底板丢失")


def check_batch6_layout(component,detail,root):
    slots={s["key"]:s for s in detail["slotBindings"]}
    boxes={s["id"]:s["boxPx"] for s in detail["templateElements"]}
    figure=component["referenceNumber"]-28
    if figure==1:
        check(len(slots["addends"]["templateShapeIds"])==2,"两项合成缺少加项")
        for ident,expected in ((96,"+"),(97,"=")):
            check("".join(find_shape(root,ident).xpath(".//a:t/text()",namespaces=NS))==expected,"组合符号有误")
        check(len(slots["technicalImages"]["templateShapeIds"])==4,"合并后的第一栏丢失原图")
        variant=detail["variants"][0]
        check(len(next(s for s in variant["slotBindings"] if s["key"]=="addends")["templateShapeIds"])==3,"三项变体缺少独立加项槽位")
    elif figure==2:
        check(len(slots["overviewItems"]["templateShapeIds"])==4,"左侧总览条目缺失")
        check(len(slots["contextBodies"]["templateShapeIds"])==2 and len(slots["stageBodies"]["templateShapeIds"])==3,"右侧上下分组被合并")
        check(boxes[156][1]<boxes[172][1] and boxes[82][2]<boxes[96][2]+boxes[158][2],"左右分组布局不符")
    elif figure==3:
        check(boxes[26][2]<boxes[32][2],"左右不等宽结构被改为等宽")
        check(len(slots["featureChain"]["templateShapeIds"])==3 and len(slots["resultFigures"]["templateShapeIds"])==2,"宽栏内部图文结构缺失")
    elif figure==4:
        check(boxes[17][1]+boxes[17][3]<boxes[27][1],"上下两层应独立排列")
        check(len(slots["lowerFigures"]["templateShapeIds"])==2 and len(slots["branches"]["templateShapeIds"])==2,"下层图文或分支落点缺失")
    elif figure==5:
        check(len(slots["guidanceLabels"]["templateShapeIds"])==2 and len(slots["supportFigures"]["templateShapeIds"])==3,"导向模块与图文支撑不完整")
        for variant in detail["variants"]:
            page=validate_package(LIB/variant["templatePath"])
            actual=["".join(find_shape(page,i).xpath(".//a:t/text()",namespaces=NS)) for i in (50,107)]
            check(actual==variant["recipe"]["labels"],"导向名称未随所选形式替换")


def main():
    catalog = read(LIB / "components.json")
    theme = read(LIB / "theme.json")
    inventory = read(LIB / "inventory.json")
    component_ids=[c["id"] for c in catalog["components"]]
    check(len(component_ids)==len(set(component_ids)),"组件ID重复")
    groups = catalog["groups"]
    group_ids = [g["id"] for g in groups]
    check(len(group_ids) == len(set(group_ids)), "组别ID重复")
    check({c["groupId"] for c in catalog["components"]} == set(group_ids), "存在空组或未归组组件")
    order = [group_ids.index(c["groupId"]) for c in catalog["components"]]
    check(order == sorted(order), "组件组未连续排列")
    suites = [g["id"] for g in groups if g["kind"] == "navigation"]
    check(group_ids[:len(suites)] == suites, "首尾页套系未置于库首")
    for suite in suites:
        members = [c for c in catalog["components"] if c["groupId"] == suite]
        check([c.get("pageRole") for c in members] == ["cover", "toc", "closing"], "首尾页套系角色不完整或次序有误")
    check({"BW-TOC-01","BW-POLICY-01","BW-METHOD-01","BW-MECHANISM-01","BW-APPLICATION-01",
           "BW-TRIPLE-PROGRESS-01","BW-TRIPLE-COMMON-01","BW-TRIPLE-DECOMPOSE-01","BW-PROBLEM-IDEA-TECH-01"}<=set(component_ids),"已有组件或本次四类组件缺失")
    check({'BW-DUAL-FACTOR-01','BW-TWO-CLASS-01','BW-TWO-HIERARCHY-01','BW-TWO-PERSPECTIVE-01','BW-FOUR-PARALLEL-01','BW-POLICY-TIMELINE-01','BW-BILATERAL-CENTER-01'}<=set(component_ids),"第三批结构缺失")
    check(catalog["selectionPolicy"]["selectedBy"] == "human", "人工选型策略丢失")
    check(catalog["selectionPolicy"]["structureLocked"] is True, "结构锁定策略丢失")
    by_reference = {c["referenceNumber"]: c for c in catalog["components"]}
    for reference in read(LIB / "references/manifest.json"):
        public_path = LIB / "references" / reference["file"]
        path = ROOT / reference["sourcePath"] if reference.get("sourcePath") else public_path
        if path != public_path:
            component = by_reference[reference["referenceNumber"]]
            check(not public_path.exists() and not component.get("referenceFile"), f"归档原图仍在展示：{path.name}")
        check(digest(path) == reference["sha256"], f"用户原图被改写：{path.name}")
    details = {c["id"]: c for c in inventory["components"]}
    result = []
    for c in catalog["components"]:
        d = details[c["id"]]
        check(digest(ROOT / d["sourcePath"]) == d["sourceSha256"], "原始PPT变化，需重新核对来源")
        template = LIB / d["templatePath"]
        check(digest(template) == d["templateSha256"], "母版与登记哈希不符")
        root = validate_package(template)
        if c.get("sourceKind")=="reference-reconstruction":check_reference_layout(c,d,root)
        if 29<=c["referenceNumber"]<=33:check_batch6_layout(c,d,root)
        source_ids = {s["id"] for s in d["sourceElements"]}
        source_ids.update(d.get("generatedShapeIds",[]))
        if c.get('pageFrame'):
            from page_frame import FRAME_IDS
            source_ids|=FRAME_IDS
        for region in c["regions"]:
            check(set(region["shapeIds"]) <= source_ids, f"区域对象无来源：{c['id']}/{region['id']}")
            check(set(region.get("generatedShapeIds",[])) <= {s["id"] for s in d["templateElements"]},"补充的原生对象未绑定到功能分区")
        for slot in c["slots"]:
            check(set(slot["shapeIds"]) <= source_ids, f"槽位对象无来源：{c['id']}/{slot['key']}")
        for run in root.findall(".//a:r", NS):
            pr = run.find("a:rPr", NS)
            for script in ["latin", "ea", "cs"]:
                font = pr.find(f"a:{script}", NS)
                check(font is not None and font.get("typeface") == theme["typography"]["family"], "可编辑文字字体未统一")
        if c["titleShapeId"] is not None:
            title = next(n for n in root.findall(".//p:sp", NS) if n.find(".//p:cNvPr", NS).get("id") == str(c["titleShapeId"]))
            for pr in title.findall(".//a:rPr", NS):
                check(pr.get("sz") == "3200" and pr.get("b") == "1", "页标题不符合统一规范")
        elif c["id"]=="BW-TOC-01":
            validate_toc(root, 6, 0, theme)
        elif c.get("tocLayout"):
            validate_suite_toc(root, d["tocLabels"], d["activeIndex"], c["tocLayout"])
        if c.get("pageRole") != "toc":
            final = {s["id"]: s for s in d["templateElements"]}
            base={s['id']:s for s in d.get('layoutElements',d['sourceElements'])}
            text_widths={ident:rule["centeredWidthScale"] for rule in c.get("textFrameOverrides",[])
                         if "centeredWidthScale" in rule for ident in rule["shapeIds"]}
            for s in d.get("layoutElements",d["sourceElements"]):
                if not c.get('pageFrame'):
                    check(final[s["id"]]["boxPx"] == s["boxPx"], "内容组件的原生对象位置或尺寸发生变化")
                    continue
                frame=c['pageFrame'];ancestor=s['id'];chain=[ancestor]
                while base[ancestor]['parentId'] is not None:ancestor=base[ancestor]['parentId'];chain.append(ancestor)
                removed=bool(set(chain)&set(frame['removeShapeIds']))
                if removed:check(s['id'] not in final,'旧页头/页码仍存在');continue
                check(s['id'] in final,'统一页框丢失了内容对象')
                check(final[s['id']]['fontSizesPt']==s['fontSizesPt'],'统一页框改变了正文显式字号')
                if s['boxPx'] is None:continue
                sizes={r['shapeId']:r for r in frame.get('sizeOverrides',[])}
                if any(i in sizes and i!=s['id'] for i in chain):continue # 装饰渐变的子图形随登记外框变化，无文字或图片。
                expected=s['boxPx'].copy();expected[1]+=round((frame.get('translateY',0)+frame.get('extraTranslateY',{}).get(str(ancestor),0))*9525)/9525
                if s['id'] in sizes:expected[2:]=sizes[s['id']]['size']
                if s["id"] in text_widths:
                    width=expected[2]*text_widths[s["id"]]
                    expected[0]+=(expected[2]-width)/2
                    expected[2]=width
                check(max(abs(a-b) for a,b in zip(final[s['id']]['boxPx'],expected))<.001,'正文没有按登记位移/留白调整，或尺寸意外改变')
            if c.get("sourceSelection"):
                expected={s["id"]:s for s in d["sourceElements"]}
                for s in d["layoutElements"]:
                    if s["boxPx"] is not None:
                        a,b=expected[s["id"]]["boxPx"],s["boxPx"]
                        check(max(abs(a[i]-b[i]) for i in (2,3))<.001,"提取局部模块时缩放了内部对象")
                        check(abs(b[1]-a[1]-c["sourceSelection"].get("translateY",0))<.001,"局部模块不是整体平移")
            for addition in c.get("nativeAdditions",[]):
                original,added=final[addition["sourceShapeId"]]["boxPx"],final[addition["newShapeId"]]["boxPx"]
                check(original[2:]==added[2:],"补齐箭头改变了原型尺寸")
        if c["id"] == "BW-POLICY-01":
            colors = {n.get("val") for n in root.findall(".//a:srgbClr", NS)}
            check({"C00000", "FFFF00"} <= colors, "用户要求的红黄强调未保留")
        if c['id']=='BW-BILATERAL-CENTER-01':
            check(find_shape(root,124).find('p:spPr/a:prstGeom',NS).get('prst')=='rightArrow',"左侧未指向中心")
            check(find_shape(root,126).find('p:spPr/a:xfrm',NS).get('flipH')=='1',"右侧箭头变成从中心向外")
            check(c['flowContract']['edges'][1]['to']=='center',"共同作用被误登记为横向连续流程")
        if c['id']=='BW-TWO-HIERARCHY-01':
            check(next(s for s in c['slots'] if s['key']=='numbers')['shapeIds']==[14,73],'双栏内容编号槽位丢失')
            for ident in [14,73]:
                bp=find_shape(root,ident).find('p:txBody/a:bodyPr',NS)
                check(bp.get('wrap')=='none' and bp.get('lIns')==bp.get('rIns')=='0',"两位编号存在错误换行风险")
            summary=find_shape(root,c['contentSummary']['shapeId'])
            check(summary.find('p:spPr/a:ln/a:solidFill/a:srgbClr',NS).get('val')==theme['palette']['primary'],"底部总述矩形缺少蓝色边框")
            check(next(s for s in c['slots'] if s['key']=='summary')['role']=='page-conclusion',"底部总述未登记为可替换槽位")
        if c['id']=='BW-POLICY-TIMELINE-01':
            check(len(next(s for s in c['slots'] if s['key']=='dates')['shapeIds'])==4 and len(next(s for s in c['slots'] if s['key']=='tasks')['shapeIds'])==8,"时间节点与治理任务混淆")
        preview = LIB / "previews" / f"{c['id']}.png"
        with Image.open(preview) as image:
            check(image.size == (1280, 720), "预览不是统一画幅")
        variants=[validate_variant(c,v,root,theme) for v in d.get("variants",[])]
        result.append({"id":c["id"], "sourceSlide":c["sourceSlide"], "nativeObjects":len(d["templateElements"]),
                       "regions":len(c["regions"]), "slots":len(c["slots"]), "geometryMode":"page-frame-declared-translations" if c.get('pageFrame') else "translation-only" if c.get("sourceSelection") else "toc-reflow" if c.get("pageRole")=="toc" else "source-positions",
                       "variants":variants})
    qa = ROOT / "_qa"
    qa.mkdir(exist_ok=True)
    toc = next(c for c in catalog["components"] if c["id"]=="BW-TOC-01")
    # 检查1至10项；2项和10项另留原生页做实际渲染。
    for amount in range(1, 11):
        path = qa / f"toc-check-{amount:02d}.pptx"
        build_single_slide(ROOT / "materials" / toc["sourceFile"], toc, theme, path,
                           labels=[f"第{i+1}部分" for i in range(amount)], active_index=amount-1)
        validate_toc(validate_package(path), amount, amount-1, theme)
    for active in range(5):
        path = qa / f"toc-active-{active+1:02d}.pptx"
        build_single_slide(ROOT / "materials" / toc["sourceFile"], toc, theme, path,
                           labels=[f"第{i+1}部分" for i in range(6)], active_index=active)
        validate_toc(validate_package(path), 6, active, theme)
    for invalid in [0, -1, True, 20]:
        try:
            toc_positions(invalid, theme)
        except ValueError:
            check(True, "非法或超容量目录被阻止")
        else:
            raise AssertionError(f"非法/超容量目录未被阻止：{invalid}")
    from deck_pipeline import reflow_toc
    suite_counts = {}
    for component in catalog["components"]:
        layout = component.get("tocLayout")
        if not layout:
            continue
        rule = layout["spacing"]
        maximum = int((rule["panelHeight"] - rule["minimumGap"]) //
                      (rule["minimumRowHeight"] + rule["minimumGap"]))
        ident = component["id"]
        template = LIB / details[ident]["templatePath"]
        suite_counts[ident] = list(range(1, maximum + 1))
        for amount in suite_counts[ident]:
            labels = [f"第{i + 1}部分的章节名称" for i in range(amount)]
            path = qa / f"{ident}-{amount:02d}.pptx"
            reflow_toc(template, labels, amount - 1, theme, path, layout=layout)
            validate_suite_toc(validate_package(path), labels, amount - 1, layout)
    # 模型可逐栏选择：一栏无图、一栏原有内框、一栏保留图，且可同时取消栏间连线。
    progress=next(c for c in catalog["components"] if c["id"]=="BW-TRIPLE-PROGRESS-01")
    mixed=build_variant(LIB/"templates"/f'{progress["id"]}.pptx',progress,
        {"id":"mixed-check","name":"部分栏无图检查","recipe":{"kind":"panel-formats","panelModes":["nested-text","nested-text","image-text"],"connectionMode":"parallel"}},
        theme,LIB,qa/"mixed-panels.pptx")
    mixed_root=validate_package(qa/"mixed-panels.pptx")
    mixed_ids={s["id"] for s in shape_inventory(mixed_root)}
    check(117 not in mixed_ids and 150 in mixed_ids,"逐栏图文选择未按输入保留或删除图片")
    check(not({60,69}&mixed_ids),"部分无图的并列选择没有取消栏间连线")
    with zipfile.ZipFile(LIB/"templates"/f'{progress["id"]}.pptx') as z:
        data={n:z.read(n) for n in z.namelist()}
    part,_=slide_part(data,1)
    root=xml(data[part])
    relation_id=find_shape(root,117).find("p:blipFill/a:blip",NS).get("{"+NS["r"]+"}embed")
    relation=next(r for r in xml(data[rels_path(part)]) if r.get("Id")==relation_id)
    sample=qa/"source-image-sample.png"
    sample.write_bytes(data[resolve_target(part,relation.get("Target"))])
    image_variant=build_variant(LIB/"templates"/f'{progress["id"]}.pptx',progress,
        {"id":"image-check","name":"本次真实图片绑定检查","recipe":{"kind":"panel-formats","panelModes":["image-text"]*3,
        "images":{"2":str(sample)},"connectionMode":"related"}},theme,LIB,qa/"middle-image.pptx")
    image_root=validate_package(qa/"middle-image.pptx")
    image_ids={s["id"] for s in shape_inventory(image_root)}
    check(54100 in image_ids and not({123,124,126,128}&image_ids),"原无图栏切换图片时未替换原内嵌结构")
    asset=image_variant["addedImages"][0]
    check(asset["sourceAssetSha256"]==digest(sample),"插入图片时改写了原图像素")
    check(abs(asset["actualBoxPx"][2]/asset["actualBoxPx"][3]-asset["pixels"][0]/asset["pixels"][1])<.0001,"插图被拉伸")
    decomposition=next(c for c in catalog["components"] if c["id"]=="BW-TRIPLE-DECOMPOSE-01")
    import copy
    spec=copy.deepcopy(decomposition["variants"][0]);spec["recipe"]["bullet"]="▪"
    build_variant(LIB/"templates"/f'{decomposition["id"]}.pptx',decomposition,spec,theme,LIB,qa/"list-square-bullet.pptx")
    square=validate_package(qa/"list-square-bullet.pptx")
    for ident in (13,20):
        check(all(n.get("char")=="▪" for n in find_shape(square,ident).findall(".//a:buChar",NS)),"替代无序符号未生效")
    four=next(c for c in catalog['components'] if c['id']=='BW-FOUR-PARALLEL-01')
    with zipfile.ZipFile(LIB/'templates'/f'{four["id"]}.pptx') as z:fourdata={n:z.read(n) for n in z.namelist()}
    fp,_=slide_part(fourdata,1);fr=xml(fourdata[fp]);rid=find_shape(fr,102).find('p:blipFill/a:blip',NS).get('{'+NS['r']+'}embed')
    rel=next(r for r in xml(fourdata[rels_path(fp)]) if r.get('Id')==rid)
    photo=qa/'four-source-photo.png';photo.write_bytes(fourdata[resolve_target(fp,rel.get('Target'))])
    four_detail=build_variant(LIB/'templates'/f'{four["id"]}.pptx',four,
        {'id':'four-mixed-check','name':'四栏图文选择检查','recipe':{'kind':'four-content','modes':['image-text','text','text','text'],'images':{'1':str(photo)}}},theme,LIB,qa/'four-mixed.pptx')
    four_root=validate_package(qa/'four-mixed.pptx');four_ids={s['id'] for s in shape_inventory(four_root)}
    check(55000 in four_ids and 94 not in four_ids and 102 not in four_ids,"四栏各自图片/文字选择未生效")
    asset=four_detail['addedImages'][0]
    check(asset['sourceAssetSha256']==digest(photo),"四栏图片像素被改写")
    check(abs(asset['actualBoxPx'][2]/asset['actualBoxPx'][3]-asset['pixels'][0]/asset['pixels'][1])<.0001,"四栏图片发生拉伸")
    render_report=read(ROOT/"docs/powerpoint-render.json")["items"]
    for d in inventory["components"]:
        for key,path,preview in [(d["id"],d["templatePath"],f'previews/{d["id"]}.png')]+[(d["id"]+"__"+v["id"],v["templatePath"],v["previewPath"]) for v in d.get("variants",[])]:
            record=render_report[key]
            check(record["inputSha256"]==digest(LIB/path),"实际预览对应旧版母版")
            check(record["previewSha256"]==digest(LIB/preview),"实际预览文件被替换或未更新")
    report = {"version":catalog["version"], "date":date.today().isoformat(), "passed":True, "checksPassed":count,
              "components":result, "tocCountsChecked":list(range(1,11)), "sixItemActiveStatesChecked":list(range(1,7)),
              "navigationSuites":len(suites), "suiteTocCountsChecked":suite_counts,
              "scope":"文件结构、来源哈希、槽位绑定、字体、颜色、原生对象几何、目录净间距与状态；不包含新Word内容制作验收"}
    (ROOT / "docs/validation.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"组件校验通过：{len(catalog['components'])}个组件，{sum(len(c.get('variants',[])) for c in catalog['components'])}个变体，{count}项检查")


if __name__ == "__main__":
    main()
