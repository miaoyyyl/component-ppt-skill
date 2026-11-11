"""第六批组件使用仓库原生PPT，登记组合逻辑、局部调整和授权变体。"""
import copy
import hashlib
import json
from pathlib import Path

from PIL import Image

ROOT=Path(__file__).resolve().parents[1]
LIB=ROOT/"library"
TERAHERTZ="重点研发计划（太赫兹运用 蓝色）.pptx"


def region(key,name,ids,purpose):
    return dict(id=key,name=name,shapeIds=ids,function=purpose,replace=purpose)


def slot(key,ids,kind,role,replace):
    return dict(key=key,shapeIds=ids,kind=kind,role=role,replace=replace)


def component(code,name,figure,file,slide,title,task,argument,regions,slots,remove):
    return dict(id="BW-"+code+"-01",name=name,referenceNumber=28+figure,
        referenceLabel=f"第六批图{figure}",referenceFile=f"reference-{28+figure}.png",
        sourceFile=file,sourceSlide=slide,titleShapeId=61000,task=task,argument=argument,
        fits=[task],doesNotFit=["内容缺少该结构所表达的逻辑联系"],
        fixed=["纯白画布、32pt蓝色页标题、双层蓝色隔断及底条，无页码"]+[r["name"] for r in regions],
        adaptationRules=["以用户给定的组件选型为主，替换独立文字和图片槽位，保留边框、渐变、标签与箭头",
            "示例图片和文字保留作结构参照，实际使用时按本次材料替换并核对来源"],
        regions=[region("header","统一页标题",[61000,61001],"简明标明本页主题")]+regions,
        slots=[slot("pageTitle",[61000],"text","page-title","32pt蓝色页标题")]+slots,
        pageFrame=dict(titleText=title,sourceTitleShapeId=None,removeShapeIds=remove,translateY=0))


def resize(c,ident,width,height):
    from native_pptx import shape_inventory,xml
    from zipfile import ZipFile
    with ZipFile(ROOT/"materials"/c["sourceFile"]) as z:
        nodes=shape_inventory(xml(z.read(f'ppt/slides/slide{c["sourceSlide"]}.xml')))
    box=next(s["boxPx"] for s in nodes if s["id"]==ident)
    c["pageFrame"].setdefault("sizeOverrides",[]).append(
        dict(shapeId=ident,originalSize=box[2:],size=[width,height]))


def definitions():
    lead=lambda:region("lead","悬挂标签与总述框",[3,4,5,6,7],"悬挂标签命名主线，带边框的总述介绍组合判断")
    lead_slots=lambda:[slot("leadLabel",[6],"text","module-title","总述的功能标签"),
                       slot("lead",[7],"rich-text","lead","统领后续模块的一句话")]
    a=component("ADDITIVE-RATIONALE","多项合成与分组阐述",1,TERAHERTZ,4,"多项条件共同支撑研究价值",
        "将多项条件通过加号合成总判断，再以三组图文解释其依据",
        "总述引题，必要性与需求组合为价值判断，下方分别给技术、公共需求和应用环境的证据",
        [lead(),region("composition","多项相加与合成结果",[93,94,95,96,97,131,132,133],"各加项列出独立条件，等号右侧给合成判断"),
         region("evidence-one","第一组图文证据",[98,99,38,137,135,139,141],"以四张相关图片支撑第一组主题"),
         region("evidence-two","第二组图文证据",[106,107,108,148,143,33],"主图与图注解释第二组需求"),
         region("evidence-three","第三组图文证据",[120,121,122,151,129],"主图和多行文字展开第三组应用")],
        lead_slots()+[slot("addends",[131,132],"text-group","module-title","两个可替换的加项"),
         slot("result",[133],"text","page-conclusion","多项条件合成的判断"),
         slot("groupTitles",[38,108,122],"text-group","module-title","三组阐述的标题"),
         slot("technicalImages",[137,135,139,141],"image-group","technical-evidence","第一组的四张图"),
         slot("needImages",[143,148],"image-group","need-evidence","第二组的两张图"),
         slot("applicationImage",[151],"image","application-evidence","第三组主图"),
         slot("captions",[33,129],"rich-text-group","local-conclusion","第二、三组的解释")],
        [2,100,101,102])
    resize(a,98,396.1,366.6)
    a["pageFrame"]["extraTranslateY"]={"137":30.7,"135":30.4,"139":-49.9,"141":-49.7}
    a["variants"]=[dict(id="three-addends",name="三项相加",recipe=dict(kind="additive-three"))]
    a["doesNotFit"]=["按时间顺序展开多个独立阶段"]
    a["adaptationRules"].append("选择两项或三项相加，保留加号和等号的组合逻辑；下方三组证据按材料确定主题")

    b=component("OVERVIEW-STACK","左侧总览与右侧分组递进",2,TERAHERTZ,5,"总览与分组行动共同说明发展路径",
        "左侧在一个外框中列出多点目标，右上并列描述，右下以三个阶段和说明框展开",
        "总述限定共同主题，左侧汇总发展条件并向下落到平台，右侧从背景描述转入行动阶段",
        [lead(),region("overview","左侧总览",[82,81,133,153,151,150,149,154,140],"一个外框中列出四项内容，向下箭头引出共同落点"),
         region("context","右上并列卡片",[96,100,109,156,158,159,160,161],"并列描述两类背景、对象或状态"),
         region("stages","右下阶段与说明",[83,84,85,92,93,94,165,167,168,172,173,175],"三个蓝色阶段标签分别对应下方说明框")],
        lead_slots()+[slot("overviewTitle",[133],"text","module-title","左栏总标题"),
         slot("overviewItems",[134,144,146,148],"text-group","hierarchy-body","左栏四个独立条目"),
         slot("overviewResult",[140],"text","local-conclusion","向下箭头所指的共同落点"),
         slot("contextTitles",[100,159],"text-group","module-title","右上两个卡片的标题"),
         slot("contextBodies",[156,161],"rich-text-group","hierarchy-body","右上多点描述"),
         slot("stageTitles",[165,167,168],"text-group","module-title","三个阶段名称"),
         slot("stageBodies",[172,173,175],"rich-text-group","hierarchy-body","每阶段的具体说明")],[2])
    resize(b,175,206.4,145.0)
    b["doesNotFit"]=["各组之间要求一一配对比较的结果"]

    c=component("ASYMMETRIC-EVIDENCE","不等宽并列图文分析",3,TERAHERTZ,11,"不等宽模块分别解释设计与验证",
        "通过不等宽并列区域安排图文，左侧聚焦方案，右侧容纳实验、特征链与结果",
        "目标框和总领句统领，下方窄栏解释设计，宽栏以实验图、三点特征和结果回扣",
        [region("objective","研究目标标签与目标框",[3,4,5,10],"左侧功能标签引出宽目标描述"),
         region("lead","目标拆解总领句",[7,8],"用一句话点明两侧模块共同支撑的能力"),
         region("design","左侧窄幅方案",[26,27,38,41,46,45],"方案主图配标题，图内带说明框强调局部机理"),
         region("validation","右侧宽幅验证",[32,33,16,17,21,22,23,24,49,51,55,57,59,61,62,65,67],"上部图证和三点特征链，下部结果图及总结")],
        [slot("objectiveLabel",[5],"text","module-title","目标标签"),
         slot("objective",[10],"rich-text","lead","总体研究目标"),
         slot("lead",[8],"text","lead","能力或目标的拆解"),
         slot("groupTitles",[41,61],"text-group","module-title","左右模块标题"),
         slot("designFigure",[38],"image","technical-evidence","左侧方案图"),
         slot("designExplanation",[45],"rich-text","local-conclusion","左图的局部说明"),
         slot("experimentFigures",[49,51],"image-group","experimental-evidence","右上两张实验图"),
         slot("featureChain",[55,57,59],"text-group","hierarchy-body","右上三个连续特征"),
         slot("resultFigures",[65,67],"image-group","result-evidence","右下两张结果图"),
         slot("result",[62],"rich-text","local-conclusion","右下结果总结")],[2])
    c["adaptationRules"].append("左右栏宽保持原图比例；右栏内部的图证、特征链和结果分区保留")
    c["doesNotFit"]=["左右内容要求完全对称的量化对照"]

    d=component("TWO-LEVEL-FEATURES","上下两层要点与图文归纳",4,"联合基金(亚热带岩性森林 蓝色）.pptx",12,"上下两层分别阐述要点与特色",
        "上层以多条文字列举一个层面，下层以图片和分支总结列举另一个层面",
        "总述定位整体价值，上层解释创新或要点，下层用两组图文汇入共同主题后分出两个落点",
        [region("lead","顶部总述矩形",[14,15,16],"概括两个层面共同支持的价值判断"),
         region("upper","上层标签与多点说明",[17,18,19,20,22,23,24],"悬挂式标签统领三条说明"),
         region("lower","下层图文组",[27,28,29,30,46,6150,33,34,35,36],"两组图片与说明并列支撑下层主题"),
         region("synthesis","向右汇总与分支",[25,26,2,40,39,37,47,48,49],"图文经箭头汇入核心主题，再向下分出两个落点")],
        [slot("lead",[15],"text","lead","全页总述"),
         slot("levelLabels",[19,30],"text-group","module-title","两个层面的标签"),
         slot("upperPoints",[22,23,24],"rich-text-group","hierarchy-body","上层三项说明"),
         slot("lowerFigures",[46,6150],"image-group","supporting-evidence","下层两组配图"),
         slot("figureLabels",[34,36],"text-group","figure-title","两个图文组的简短概括"),
         slot("synthesis",[47],"text","yellow-label","汇入后的核心判断"),
         slot("branches",[48,49],"text-group","local-conclusion","两个落点的说明")],[4,8,9,11,41])
    d["colorOverrides"]=[dict(shapeId=47,mode="yellow-text")]
    d["doesNotFit"]=["必须连续讲述时间顺序的过程"]

    e=component("GUIDED-EXPLANATION","问题、方法或技术导向展开",5,"3深蓝白色系-国家自然科学项目申请答辩结题汇报PPT.pptx",4,"以问题、方法或技术组织支撑内容",
        "顶部总述引题，左侧两组支撑内容通过箭头对应右侧两个导向模块，底部汇总行动",
        "总述、对应支撑、分组解释和后续行动形成闭合叙述",
        [region("lead","顶部总述",[40,41],"交代本页讨论范围"),
         region("support-one","左上支撑模块",[89,90,91,94,95,96,97],"一个标题加两个并列目标或特征"),
         region("support-two","左下图文模块",[58,72,73,5132,5130,5128],"三图与简短图题提供第二组支撑"),
         region("links","左右对应箭头",[87,88],"将左侧两组内容分别连接到右侧导向模块"),
         region("guidance","右侧导向说明",[44,50,107,108,114,115,116,119,123],"两个导向标题和分点解释，第二组附两个小框"),
         region("takeaway","底部两条归纳",[42,43],"用两条行动或结论收束本页")],
        [slot("lead",[41],"rich-text","lead","本页总述"),
         slot("supportTitle",[95],"text","module-title","左上模块名称"),
         slot("supportGoals",[91,97],"text-group","hierarchy-body","两个目标或特征"),
         slot("supportFigures",[5132,5128,5130],"image-group","supporting-evidence","左下三张相关图片"),
         slot("figureCaption",[73],"text","figure-title","三图的共同说明"),
         slot("guidanceLabels",[50,107],"text-group","red-subtitle","问题一/二，也可切换为方法或技术"),
         slot("guidanceBodies",[108,114],"rich-text-group","hierarchy-body","两个导向模块的分点内容"),
         slot("detailBoxes",[116,123],"text-group","hierarchy-body","第二模块下方的两项补充"),
         slot("takeaways",[43],"rich-text","page-conclusion","底部两条归纳或行动")],[4,5,6,7,10,3076,5122,122])
    resize(e,108,681.0,77.1)
    resize(e,114,700.0,42.2)
    e["variants"]=[dict(id=key,name=name,recipe=dict(kind="guided-labels",labels=labels)) for key,name,labels in
        [("methods","方法导向",["方法一","方法二"]),("technologies","技术导向",["技术一","技术二"])]]
    e["doesNotFit"]=["缺少支撑材料和分解内容的孤立结论"]
    e["adaptationRules"].append("导向名称用于标明模块角色，正文制作时按本次材料一并替换")
    return [a,b,c,d,e]


def apply_recipe(root,component,recipe,theme):
    from native_pptx import NS,set_box
    from component_variants import find_shape,clone_ids
    if recipe["kind"]=="guided-labels":
        labels=recipe.get("labels")
        if not isinstance(labels,list) or len(labels)!=2 or any(not isinstance(s,str) or not s.strip() for s in labels):
            raise ValueError("导向名称须包含两个非空标题")
        for ident,label in zip((50,107),labels):
            for i,node in enumerate(find_shape(root,ident).findall(".//a:t",NS)):
                node.text=label if i==0 else ""
        return dict(regionShapeIds={},addedSlots=[],replacedSlotKeys=[])
    tree=root.find("p:cSld/p:spTree",NS)
    copies=[]
    for source,ident in ((93,63000),(131,63001),(96,63002)):
        node=copy.deepcopy(find_shape(root,source));clone_ids(node,ident);tree.append(node);copies.append(ident)
    for ident,x,w in ((93,60.8,232),(94,356.8,232),(63000,652.8,232),(95,948.8,284.7)):
        set_box(find_shape(root,ident),x,243.1,w,57.7)
    for ident,x,label in ((131,70.8,"研究必要性"),(132,366.8,"重大需求"),(63001,662.8,"技术基础"),(133,958.8,"科学与应用价值")):
        node=find_shape(root,ident);set_box(node,x,251,264.7 if ident==133 else 212,42)
        for i,t in enumerate(node.findall(".//a:t",NS)):t.text=label if i==0 else ""
    for ident,x in ((96,296.8),(63002,592.8),(97,888.8)):
        set_box(find_shape(root,ident),x,238,60,67.9)
    region_ids=next(r["shapeIds"] for r in component["regions"] if r["id"]=="composition")
    return dict(regionShapeIds={"composition":region_ids+copies},replacedSlotKeys=["addends"],
        addedSlots=[slot("addends",[131,132,63001],"text-group","module-title","三个独立加项")])


def register():
    path=LIB/"components.json";data=json.loads(path.read_text(encoding="utf8"))
    items=definitions();ids={c["id"] for c in items}
    data["components"]=[c for c in data["components"] if c["id"] not in ids]+items
    if data["version"]=="0.6.0":
        data["compatibleTaskVersions"].append("0.6.0");data["version"]="0.7.0"
    path.write_text(json.dumps(data,ensure_ascii=False,indent=2)+"\n",encoding="utf8")
    path=LIB/"references/manifest.json";refs=json.loads(path.read_text(encoding="utf8"))
    refs=[r for r in refs if r["referenceNumber"] not in range(29,34)]
    for i in range(29,34):
        picture=LIB/"references"/f"reference-{i}.png"
        with Image.open(picture) as image:size=list(image.size)
        refs.append(dict(referenceNumber=i,referenceBatch="batch-06",batchFigure=i-28,file=picture.name,
            pixels=size,sha256=hashlib.sha256(picture.read_bytes()).hexdigest(),status="用户原图，像素未改写"))
    path.write_text(json.dumps(refs,ensure_ascii=False,indent=2)+"\n",encoding="utf8")


if __name__=="__main__":register()
