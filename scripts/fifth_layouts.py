"""第五批原生组件布局，分别用于导航页面、纯文字比较和无框案例。"""
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Pt

from reference_layouts import BLUE, PALE, WHITE, RED, px
from layout_styles import clear_effects, summary_gradient, corner_tab


def listing(p,key,label,x,y,w,h,text,size,numbered=False,mark="•"):
    p.text(key,label,x,y,w,h,text,size)
    for para in p.slide.shapes[-1].text_frame.paragraphs:
        props=para._p.get_or_add_pPr()
        props.set("marL",str(px(38)))
        props.set("indent",str(px(-34)))
        bullet=OxmlElement("a:buAutoNum" if numbered else "a:buChar")
        bullet.set("type" if numbered else "char","romanUcPeriod" if numbered else mark)
        props.append(bullet)
        para.space_after=Pt(5)


def bookend(p,closing):
    p.regions.clear()
    p.slots.clear()
    p.region("title","汇报主题","封面以大标题引题，结束页以小标题回扣同一主题")
    for x in (180,966):
        p.rect(x,90,134,1,"A6BDE0")
        p.rect(x+130 if x==180 else x,87,7,7,BLUE,shape_type=MSO_SHAPE.DIAMOND)
    p.text("pageTitle","汇报标题",68,65 if closing else 167,1144,52 if closing else 176,
           "PPT标题XXXXXXXXXX",26 if closing else 58,BLUE,closing,True)
    p.slots["pageTitle"]["role"]="page-title"
    if closing:
        p.region("thanks","结束语","中央两行依次放致谢对象与讨论或反馈邀请")
        p.text("thanks","致谢与邀请",65,143,1150,276,"感谢各位同学和老师\n敬请批评指正！",70,BLUE,False,True)
    p.region("signature","蓝色底带与汇报人","底部蓝带确定页面重心，交界处胶囊显示汇报人")
    p.rect(0,476,1280,244,BLUE)
    # 只截取无文字的蓝色纹理，标题、致谢与署名单独绑定。
    crop=(19,572,1440,243) if closing else (26,580,1439,244)
    p.pic("bandTexture","底带背景纹理",0,500,1280,220,crop)
    badge=p.rect(472,450,336,52,BLUE,WHITE,MSO_SHAPE.ROUNDED_RECTANGLE)
    badge.adjustments[0]=.5
    p.text("presenter","汇报人署名",480,452,320,47,"汇报人：XXX",24,WHITE,True,True)


def quadrants(p,bullets):
    p.region("lead","共同判断","在四类因素上方概括整体判断")
    p.text("summary","总体判断",50,101,1180,92,"具备实施基础，提前制定预案，\n从四个方面分析项目风险",25,RED,True,True,PALE)
    data=[
        ("成本效益","实际应用需要考虑经济性和成本效益\n制备成本、能耗和维护成本可能**影响商业化前景**"),
        ("安全问题","需要考虑环境和人体健康安全风险\n材料制备和应用过程可能**涉及有害化学品**"),
        ("技术更新","项目成果可能面临**快速过时的风险**\n技术领域竞争激烈，新技术和新材料不断涌现"),
        ("标准限制","实施和市场准入受相关要求影响\n材料应用需要遵守**环保、安全和质量标准**")]
    for i,(name,points) in enumerate(data):
        x=50+602*(i%2); y=209+245*(i//2)
        p.region(f"group{i+1}",f"第{i+1}类因素","用短标签命名类别，下方列表解释该类因素及其影响")
        clear_effects(p.rect(x,y+15,576,216,WHITE,BLUE))
        p.band("labels","类别名称",x,y,185,46,name)
        corner_tab(p.slide.shapes[-1])
        listing(p,f"items{i+1}","本类独立列表",x+18,y+57,540,159,points,19,not bullets)


def outcomes(p):
    p.region("lead","核心发现","以一句判断连接案例结果与后续监测")
    p.text("summary","核心判断",46,99,1188,64,"稻米安全改善的同时，工程持续监测土壤、产量与成本",23,BLUE,True,True)
    summary_gradient(p.slide.shapes[-1])
    p.region("results","图示与案例结果","主图下列出不同地点的指标与观察期，再用一句话概括")
    p.pic("figure","结果图或实测图表",45,172,591,307,(64,213,664,355))
    listing(p,"observations","地点、指标及观察期",49,487,586,146,"宁乡：稻米镉达标率由**12.5%**升至**100%**，图示对应2019年至2021年。\n资兴：稻米镉含量降低**75.5%**，图示对应2019年至2024年。",17,mark="✓")
    p.text("takeaway","图旁小结",47,641,584,56,"长期监测显示稻米镉含量**稳定符合国家标准**",17,bold=True,center=True)
    p.region("monitoring","指标与稳定性","列出指标、数值、时间跨度与适用条件")
    p.text("headings","分析组标题",655,172,570,49,"土壤与产量",23,BLUE,True)
    listing(p,"indicators","监测指标",658,223,569,202,"活性镉：交换态与水溶态镉均报告降低**36%–47%**。\npH：由4.5–5.5调节至6.0–7.0。\n有机质：土壤有机质含量提升15%–20%。\n稳定性：持续跟踪镉活性及水稻产量。",17,mark="✓")
    p.region("adoption","成本与推广","解释实施投入、覆盖范围和推广条件")
    p.text("headings","分析组标题",655,445,570,56,"成本、认可与推广",23,BLUE,True)
    listing(p,"adoption","成本及推广要点",658,505,569,174,"成本：约2000元/亩，含5年跟踪监测。\n推广：在湖南、江西、湖北等水稻主产区应用，累计修复面积超**3万亩**。",18,mark="✓")


def implementation(p):
    p.region("lead","案例总述","先明确工程同时要实现的目标")
    p.text("summary","双重目标",48,99,1184,65,"湖南酸性稻田治理同时追求**降镉与保产**",24,BLUE,True,True)
    summary_gradient(p.slide.shapes[-1])
    sections=[("酸性稻田与双重目标","湖南稻田镉污染修复技术研发始于2013年。\n稻田多为酸性红壤，治理需兼顾**降镉与保产**。\n湖南省地球物理地球化学调查所开展专项攻关，历经11年研发。"),
              ("示范与配方优化","2017年宁乡建立**2680亩**示范基地，2019年扩展至资兴**6750亩**试验田。\n在多地布设试验区块，开展采样、配方优化与效果验证。\n积累不同土壤类型的技术参数。")]
    for i,(name,items) in enumerate(sections):
        x=45+610*i
        p.region(f"analysis{i+1}",name,"用蓝色分标题划分背景目标与实施迭代，独立要点用无序符号排列")
        p.text("headings","分析维度",x,174,576,56,name,23,BLUE,True)
        listing(p,"details","分点分析",x+2,235,568,228,items,18,mark="➢")
    p.region("scene","现场配图","附场景照片与图注，明确照片的证据用途")
    p.pic("photo","场景照片",46,479,261,188,(52,568,285,213))
    p.text("caption","图注意义",46,669,290,31,"农田场景参考图（非成效证据）",12)
    p.region("steps","四步实施流程","依照实际执行顺序填写四个可操作环节")
    p.text("flowTitle","流程总领",364,493,850,55,"闭环治理的四个环节",22,BLUE,True)
    for i,title in enumerate(("土壤检测诊断","定制化药剂配比","精准施用","动态监测")):
        x=364+222*i
        p.rect(x,565,193,81,WHITE,BLUE)
        p.text("steps","实施环节",x+2,567,189,77,title,18,BLUE,True,True)
        if i<3:p.arrow(x+193,591,29,23,True)


def draw(p,index,variant):
    if index<2:bookend(p,index==1)
    elif index==2:quadrants(p,variant=="bullets")
    elif index==3:outcomes(p)
    else:implementation(p)
