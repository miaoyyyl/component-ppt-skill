from io import BytesIO
import hashlib
import json
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.util import Pt

from native_pptx import build_single_slide

LIB = Path(__file__).resolve().parents[1] / "library"
ARCHIVED_REFERENCES = LIB.parent / "materials" / "component-references"
BLUE, PALE, RED, INK, WHITE = "0046A8", "F2F6FA", "C00000", "202020", "FFFFFF"
NAMES = [
    ("FOUR-ISSUE-CASE", "四类问题与对应案例", "四项问题各有表现、要点和独立案例"),
    ("TRIPLE-EVIDENCE", "三栏图证与底部总述", "三组证据分别解释后，在页底归纳共同判断"),
    ("DUAL-METHOD-CASE", "双路径核心手段与案例", "两条路径按同一结构说明手段、案例、分析与结果"),
    ("SPLIT-EVIDENCE", "单框虚线双栏分析", "围绕同一对象，从两个维度分别解释图证"),
    ("REFERENCES", "编号参考文献", "按正文引用顺序列出完整文献条目"),
    ("CASE-BRANCH", "案例背景与行动分支", "由案例背景分出行动与效果，或两种行动"),
    ("CASE-CHAIN", "案例背景、技术与效果链", "沿单一案例的背景、实施技术和观测效果逐层展开"),
    ("COVER","汇报首页","用居中大标题介绍汇报主题，在蓝白交界标明汇报人"),
    ("CLOSING","汇报结束页","回扣汇报题目，以居中致谢或讨论邀请结束汇报"),
    ("FOUR-TEXT","四类并列纯文字","先给总判断，再以四组有序或无序要点并列展开"),
    ("CASE-OUTCOME","无框案例 · 成效与推广","左侧用图与地区数据交代结果，右侧分组说明稳定性、成本与推广"),
    ("CASE-IMPLEMENTATION","无框案例 · 实施与迭代","上方双栏梳理背景与实施，下方用场景图和四步流程说明如何执行"),
]


def px(value):
    return round(value * 9525)


class Layout:
    def __init__(self, index):
        self.deck = Presentation()
        self.deck.slide_width, self.deck.slide_height = px(1280), px(720)
        self.slide = self.deck.slides.add_slide(self.deck.slide_layouts[6])
        name = f"reference-{17+index}.png"
        archived = ARCHIVED_REFERENCES / name
        self.reference = archived if archived.is_file() else LIB / "references" / name
        self.regions = []
        self.slots = {}
        self.region("header", "统一页标题", "用一句短标题标明本页主题")
        self.regions[-1]["shapeIds"] = [61000, 61001]
        self.bind("pageTitle", 61000, "text", "本页短标题", "page-title")

    def region(self, key, name, purpose):
        self.regions.append(dict(id=key, name=name, function=purpose, replace=purpose, shapeIds=[]))

    def bind(self, key, ident, kind, instruction, role=None):
        slot = self.slots.setdefault(key, dict(key=key, shapeIds=[], kind=kind, role=role or key, replace=instruction))
        slot["shapeIds"].append(ident)

    def record(self, shape):
        self.regions[-1]["shapeIds"].append(shape.shape_id)
        return shape

    def rect(self, x,y,w,h, fill=WHITE, border=None, shape_type=MSO_SHAPE.RECTANGLE):
        s = self.record(self.slide.shapes.add_shape(shape_type,*map(px,(x,y,w,h))))
        s.fill.solid()
        s.fill.fore_color.rgb = RGBColor.from_string(fill)
        if border:
            s.line.color.rgb = RGBColor.from_string(border)
            s.line.width = Pt(.65)
        else:
            s.line.fill.background()
        return s

    def text(self,key,label,x,y,w,h,value,size=18,color=INK,bold=False,center=False,fill=None):
        s = self.rect(x,y,w,h,fill) if fill else self.record(self.slide.shapes.add_textbox(*map(px,(x,y,w,h))))
        s.name = label
        self.bind(key,s.shape_id,"text-group",label)
        t = s.text_frame
        t.clear()
        t.word_wrap = True
        t.auto_size = MSO_AUTO_SIZE.NONE
        t.margin_left = t.margin_right = px(5)
        t.margin_top = t.margin_bottom = 0
        t.vertical_anchor = MSO_ANCHOR.MIDDLE
        for i,line in enumerate(value.split("\n")):
            p = t.add_paragraph() if i else t.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if center else PP_ALIGN.LEFT
            for j,part in enumerate(line.split("**")):
                if not part: continue
                r=p.add_run(); r.text=part
                r.font.name="微软雅黑"; r.font.size=Pt(size); r.font.bold=bold or j%2==1
                r.font.color.rgb=RGBColor.from_string(RED if j%2 else color)

    def pic(self,key,label,x,y,w,h,crop,lock_frame=False):
        with Image.open(self.reference) as im:
            a,b,c,d=crop
            image=im.crop((a,b,a+c,b+d))
            if lock_frame:
                canvas=Image.new("RGB",(w,h),"white")
                image.thumbnail((w,h),Image.Resampling.LANCZOS)
                canvas.paste(image,((w-image.width)//2,(h-image.height)//2))
                image=canvas
                c,d=w,h
            buf=BytesIO(); image.save(buf,format="PNG"); buf.seek(0)
        scale=min(w/c,h/d)
        s=self.record(self.slide.shapes.add_picture(buf,*map(px,(x+(w-c*scale)/2,y+(h-d*scale)/2,c*scale,d*scale))))
        s.name=label
        self.bind(key,s.shape_id,"image-group",label)

    def arrow(self,x,y,w,h,right=False):
        self.rect(x,y,w,h,BLUE,shape_type=MSO_SHAPE.RIGHT_ARROW if right else MSO_SHAPE.DOWN_ARROW)

    def band(self,key,label,x,y,w,h,value):
        self.text(key,label,x,y,w,h,value,21,WHITE,True,True,BLUE)


def four(p):
    from layout_styles import case_marker
    from pptx.oxml.xmlchemy import OxmlElement
    rows=[
        ("空间割裂","城市与水的硬隔绝","水体被视为交通和城市建设的障碍","高架桥沿河而建\n连续护栏阻断可达性\n滨水区缺乏生态连通性","黄浦滨江变为工业锈带"),
        ("生态硬化","渠道化与生态承载力衰退","传统水规划重工程、轻生态","硬化护岸覆盖混凝土\n切断水土交换\n水生生物栖息地丧失","长沙圭塘河水体黑臭"),
        ("文化失落","水文化遗产的系统性抹除","水体承载的城市记忆被忽视","码头与桥梁遗存被拆除\n现代建筑取代历史风貌\n场所精神与地域特色丧失","苏州河里分建筑消失"),
        ("管理滞后","多部门分割与制度性障碍","分割管理，缺乏顶层规划","各部门缺乏统一协调\n技术标准不统一\n信息共享不足","成都锦江排水管理分散")]
    for i,(name,issue,why,points,case) in enumerate(rows):
        x=25+310*i
        p.region(f"column{i+1}",name,"先用现象图和要点解释一类问题，再用下方具名案例及照片举证")
        p.rect(x,100,300,598,PALE,BLUE)
        p.band("topics","问题类别",x,100,300,48,name)
        p.text("issues","问题体现",x+5,153,290,42,issue,16,RED,True,True)
        p.text("descriptions","问题说明",x+7,205,115,116,why,15,bold=True)
        p.pic("issuePhotos","问题现场图",x+125,202,165,125,([152,494,835,1177][i],219,175,131))
        p.text("points","三个具体表现",x+10,341,282,99,points,15)
        for paragraph in p.slide.shapes[-1].text_frame.paragraphs:
            props = paragraph._p.get_or_add_pPr()
            props.set("marL",str(px(20)))
            props.set("indent",str(px(-18)))
            color = OxmlElement("a:buClr")
            rgb = OxmlElement("a:srgbClr")
            rgb.set("val",BLUE)
            color.append(rgb)
            props.append(color)
            marker = OxmlElement("a:buChar")
            marker.set("char","•")
            props.append(marker)
        p.rect(x+10,448,280,1,BLUE)
        p.text("caseLabels","案例标签",x+51,454,239,35,"典型案例",16,BLUE,True)
        p.text("caseNames","案例名称",x+6,493,288,39,case,16,bold=True,center=True)
        p.pic("casePhotos","对应案例照片",x+20,538,260,148,([56,397,739,1081][i],574,243,139))
    for i in range(4):
        case_marker(p,25+310*i+4,460,38,23,f"column{i+1}")


def triple(p):
    rows=[
        ("城市韧性量化评估框架",(84,185,330,257),"城市韧性三维度耦合评估框架","以水资源、生态环境与社会经济为维度\n建立城市韧性量化评估框架\n识别更新规划的承载力阈值","J. Environ. Manage., 2026, 401: 128882."),
        ("滨水缓冲区空间评价",(499,186,381,249),"基于多源数据识别保护与修复区域","结合水文与土地利用数据划定区域\n将坡度与土壤渗透性纳入评价","STOTEN, 2021, 757: 143982; 798: 149238."),
        ("城市紧凑度与生态系统服务",(945,184,376,254),"城市紧凑性与生态系统服务的空间格局","分析紧凑度影响的空间异质性\n区分中心区与边缘区的生态响应","Environ. Impact Assess. Rev., 2023, 98: 106975.")]
    for i,(name,crop,cap,pts,src) in enumerate(rows):
        x=40+405*i
        p.region(f"evidence{i+1}",name,"短总述限定该栏观点；主图和图注说明依据；要点解释证据，栏尾保留出处")
        p.rect(x,105,390,526,WHITE,BLUE)
        p.text("overviews","各栏简短总述",x,105,390,54,name,20,BLUE,True,True,PALE)
        p.pic("figures","支持该栏观点的图表",x+8,167,374,225,crop)
        p.text("captions","图意说明",x+8,397,374,45,cap,15,bold=True,center=True)
        p.text("analysis","分点分析",x+8,446,374,140,"▪ "+"\n▪ ".join(pts.splitlines()),17,fill=PALE)
        p.text("sources","图表或研究出处",x+6,591,378,36,src,11)
    p.region("synthesis","全页底部总述","用一句话归纳三栏证据共同支持的判断")
    p.text("takeaway","整页共同结论",40,648,1200,51,"通过前置的**生态敏感性评估**，避免先破坏后补救的高风险开发模式",21,BLUE,True,True,PALE)


def dual(p,progress):
    from layout_styles import summary_gradient,case_marker
    rows=[("滨水工业遗存的织补与更新","以功能置换与空间织补转化旧仓库和码头","广州太古仓码头的功能活化","红砖仓库转化为艺术展览与餐饮空间\n结合历史记忆、消费体验与夜间经济\n通过空间叠合呈现历史记忆","由仓储空间转向复合型文化消费场所",(112,378,470,136)),
          ("非物质文化活态传承","将水乡生活习俗和水上交通融入滨水公共生活","郑州市金水河综合整治工程","结合地域文化塑造故事文化空间\n因地制宜连接水生态与水景观\n重现山水格局并划分滨水功能片区","将地域文化融入现代滨水公共生活",(705,449,285,160))]
    for i,(name,means,case,pts,end,crop) in enumerate(rows):
        x=34+i*634
        p.region(f"path{i+1}",name,"两栏均依次填入总领手段、图文案例、三条编号分析与局部结论")
        p.rect(x,101,578,598,WHITE,BLUE)
        p.text("routeNames","路径名称",x+6,108,566,54,name,22,BLUE,True,True)
        summary_gradient(p.slide.shapes[-1])
        p.band("methodLabels","手段标签",x+8,172,120,77,"核心\n手段")
        p.text("means","本栏总领性的核心手段",x+128,172,442,77,means,18,bold=True,fill=PALE)
        p.text("caseLabels","案例标签",x+70,259,494,31,"典型案例",17,BLUE,True)
        p.pic("examples","案例图片",x+40,295,498,139,crop,lock_frame=True)
        p.text("caseNames","案例名称",x+10,437,558,39,case,16,bold=True,center=True)
        for j,line in enumerate(pts.splitlines()):
            p.text("listNumbers","分析项序号",x+14,487+j*45,34,32,str(j+1),16,WHITE,True,True,BLUE)
            p.text("points","编号分析",x+54,483+j*45,509,42,line,16)
        p.text("localResults","本栏结论",x+8,638,562,52,end,17,RED,True,True,PALE)
    if progress: p.arrow(616,366,48,30,True)
    for i in range(2):
        case_marker(p,34+634*i+2,264,54,25,f"path{i+1}")


def split(p):
    p.region("shared","共同对象和单一外框","点明两栏共同分析的对象；保留一个连续边框和中央虚线")
    p.text("subject","共同分析对象",40,98,1200,45,"PFOS赋存特征——木鼠",23,BLUE,True,True)
    p.rect(36,150,1208,535,WHITE,BLUE)
    line=p.record(p.slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,*map(px,(640,167,640,670))))
    line.line.color.rgb=RGBColor.from_string(BLUE);line.line.dash_style=MSO_LINE_DASH_STYLE.DASH
    for i,(title,points,cap,end) in enumerate([
        ("年龄与暴露时长","污染组浓度高于对照组\n污染组肝脏浓度随年龄增长而上升\n对照组未见年龄效应或性别差异","图4  肝脏浓度与年龄的关系","摄入与消除速率的差异有助于解释长期蓄积"),
        ("不同靶器官","比较各器官中的浓度\n肝脏为主要靶器官\n进一步区分性别、地点和器官差异","图5  不同器官中的PFOS浓度","蛋白质结合影响其在肝脏和血液中的分布")]):
        x=50+605*i
        p.region(f"axis{i+1}",title,"依次替换分析维度、要点、图表、图注和本栏结论")
        p.text("axes","分析维度",x,161,575,43,title,21,BLUE,True)
        p.text("findings","分点发现",x,211,575,110,"✓ "+"\n✓ ".join(points.splitlines()),17)
        if i==0:p.pic("leftChart","第一维度图表",x+40,332,492,213,(177,350,405,276))
        else:
            p.pic("rightCharts","第二维度对照图",x,332,280,213,(725,418,338,193))
            p.pic("rightCharts","第二维度对照图",x+287,332,280,213,(1079,418,318,193))
        p.text("captions","图注",x,548,565,41,cap,14,center=True)
        p.text("conclusions","本维度解释",x+5,598,560,69,end,18,BLUE,True,True,PALE)
    p.text("sources","研究来源",50,685,1170,19,"D’ Hollander et al., 2014；Bogdanska et al., 2011",10)


def references(p):
    p.region("entries","连续编号的文献表","每段一条完整引用；按正文编号排列，超出本页时复制页面并续号")
    entries=[
        "KRAFFT T, RIESS M. Per- and polyfluorinated substances (PFASs): Environmental challenges [J/OL]. Current Opinion in Colloid & Interface Science, 2015.",
        "P. Z, J. S, M. H, W. A, et al. Perfluorooctanoic acid (PFOA) and perfluorooctanesulfonic acid (PFOS) in surface waters, sediments, soils and wastewater – A review on concentrations and distribution coefficients[J]. Chemosphere, 2013, 91(6): 725–732.",
        "Changlong W, Xin S, Qing W, Zhihao H, et al. Sorption Kinetics, Isotherms and Mechanisms of PFOS on Soils with Different Physicochemical Properties[J]. Ecotoxicology and Environmental Safety, 2017, 142: 40–50.",
        "Tutan D A, Thomas B, Graham D, Stephanie L W, et al. Phase Separation and Passive Diffusion of Perfluorooctane Sulfonic Acid (PFOS) in Multilamellar and Unilamellar Vesicles[J]. crossref, 2024.",
        "Paola B O, Ana L R, Reyna F, Aidee S L, et al. Perfluorooctane Sulfonate and Perfluorooctanoic Acid Induce Plasma Membrane Dysfunction in Boar Spermatozoa During in Vitro Capacitation[J]. Reproductive Toxicology, 2022, 110: 85–96.",
        "Wen Z, Jeremiah D Z, Yi W, et al. Organic Anion Transporting Polypeptides Contribute to the Disposition of Perfluoroalkyl Acids in Humans and Rats[J]. Toxicological Sciences, 2016: 236.",
        "Wen Z, Jeremiah D Z, Yi W, et al. Na+/Taurocholate Cotransporting Polypeptide and Apical Sodium-Dependent Bile Acid Transporter Are Involved in the Disposition of Perfluoroalkyl Sulfonates in Humans and Rats[J]. Toxicological Sciences, 2015, 146(2): 363–73.",
        "D’ Hollander W, De Bruyn L, Hagenaars A, et al. Characterisation of perfluorooctane sulfonate (PFOS) in a terrestrial ecosystem near a fluorochemical plant in Flanders, Belgium. Environ Sci Pollut Res 21, 11856–11866 (2014).",
        "Bogdanska J, Borg D, Sundström M, et al. Tissue distribution of 35S-labelled perfluorooctane sulfonate in adult mice after oral exposure to a low environmentally relevant dose or a high experimental dose. Toxicology 284: 54–62."]
    p.text("bibliography","按正文顺序填写已核实文献，每段以方括号编号开头",46,102,1190,586,"\n".join(f"[{i}] {e}" for i,e in enumerate(entries,1)),14,bold=True)
    for paragraph in p.slide.shapes[-1].text_frame.paragraphs:paragraph.space_after=Pt(7)


def branch(p,two):
    p.region("subject","技术或政策","点明案例所采用的技术、政策或行动主题")
    p.text("topic","技术或政策名称",30,97,1220,43,"生物修复技术——微生物降解",23,BLUE,True,True)
    p.region("background","案例背景","并列说明事件和约束条件，用向下箭头引出下方行动")
    p.rect(30,150,830,245,PALE,BLUE)
    p.band("backgroundLabel","背景标题",30,150,830,53,"案例背景")
    p.text("context","背景与条件",52,216,436,160,"原油泄漏\n• 当地石油设施遭到焚烧\n• 原油覆盖城镇及周边土地\n• 形成油污层",18)
    p.text("context","背景与条件",491,216,351,160,"气候条件加剧污染\n• 夏季高温加速挥发与渗透\n• 污染土壤释放有害气体",18)
    for x in (219,625):p.arrow(x,399,52,45)
    for i in range(2):
        x=30+439*i
        p.region("action" if i==0 else "branch2","行动一" if i==0 else "效果或第二项行动","填写实施手段或可追溯的效果；两行动变体在此并列不同手段")
        p.rect(x,454,391,234,PALE,BLUE)
        name=("菌剂处理" if i==0 else "覆盖保湿") if two else ("生物技术" if i==0 else "修复效果")
        p.band("branchNames","分支名称",x,454,391,54,name)
        value=("筛选当地耐高温的石油降解菌，制成菌剂喷洒于污染土壤" if i==0 else "覆盖秸秆保湿，提高微生物活性") if two else ("• 筛选耐高温的石油降解菌并喷洒菌剂\n• 覆盖秸秆保湿，提高微生物活性" if i==0 else "• 修复后区域可种植耐旱作物\n• 为中轻度污染土壤提供修复方案")
        p.text("branches","分支具体内容",x+12,519,367,151,value,18)
    if not two:p.arrow(425,546,39,46,True)
    p.region("evidence","右侧案例材料","行动与效果形式配现场图；两项行动形式在右侧集中说明效果")
    if two:
        p.band("outcomeHeading","右侧效果标题",885,204,357,56,"修复效果")
        p.text("outcome","有材料支持的结果",885,268,357,233,"修复后区域已可种植耐旱作物。\n\n为中轻度石油污染土壤提供修复方案。",20,fill=PALE)
    else:
        p.text("photoHeading","图片上方说明",888,168,353,84,"低成本生物修复技术",21,bold=True,center=True)
        p.pic("photo","现场照片",891,266,351,241,(1028,301,382,255))
        p.text("photoCaption","现场图注",890,514,350,84,"专家研究受油污染土壤的微生物处理过程",16,center=True)
    p.text("sources","案例出处",888,622,354,63,"Harari and Annesi-Maesano, 2023\nRobert, 2015",13)


def chain(p):
    from layout_styles import clear_effects
    p.region("case","案例名称","标明案例及具体技术主题")
    p.text("caseName","案例和技术主题",40,97,1200,48,"未爆弹药处理——智能探测与协同修复",22,BLUE,True,True)
    p.region("background","事件背景","给出时间、地点、影响范围与起始问题")
    p.band("backgroundLabel","背景标题",39,151,803,46,"案例背景")
    clear_effects(p.slide.shapes[-1])
    p.text("context","案例起始情况",75,202,731,77,"波黑战争遗留未爆弹药，分布于农田和森林。弹药锈蚀使铅、汞等重金属进入土壤。",18,bold=True,fill=PALE)
    clear_effects(p.slide.shapes[-1])
    p.arrow(413,289,55,39)
    p.region("technique","修复技术","登记项目及其连续实施步骤")
    p.band("techLabel","行动标题",39,330,803,47,"修复技术")
    clear_effects(p.slide.shapes[-1])
    p.text("project","项目名称",39,389,139,109,"安全土地\n项目",18,WHITE,True,True,BLUE)
    p.text("actions","实施技术与步骤",187,383,655,123,"✓ 智能探测：使用探地雷达与金属探测器定位\n✓ 机器人排爆：拆弹引信并集中销毁\n✓ 土壤修复：化学淋洗处理污染土壤",17,fill=PALE)
    clear_effects(p.slide.shapes[-1])
    p.arrow(413,512,55,41)
    p.region("outcomes","修复效果","填写有依据的指标、观察窗口及结果配图")
    p.rect(39,560,803,137,PALE)
    p.text("resultLabel","效果标题",46,567,100,120,"修复\n效果",20,BLUE,True,True)
    p.text("results","案例效果指标",152,566,452,124,"✓ 污染区域清理：60%（5年内完成）\n✓ 耕地恢复：30万公顷\n✓ 农业产值提升：40%\n✓ 伤亡事故下降：95%",15)
    p.pic("effectPhoto","效果配图",610,566,220,124,(653,648,231,153))
    p.region("scene","右侧现场图","用大图展示案例现场，在下方注明图意及来源")
    p.pic("scenePhoto","相关现场图片",882,151,352,445,(1058,159,302,474))
    p.text("caption","照片说明",875,603,368,48,"巴干半岛弹药清理",16,center=True)
    p.text("source","案例出处",887,657,350,38,"Robert, 2015",13)
    # 外框后创建、移至底层，正文和槽位的既有对象编号保持稳定。
    for region_id, y, height in (("background",151,134), ("technique",330,180)):
        region = next(r for r in p.regions if r["id"] == region_id)
        frame = p.rect(39,y,803,height,WHITE,BLUE)
        frame.name = region_id + "-blue-white-frame"
        clear_effects(frame)
        p.regions[-1]["shapeIds"].remove(frame.shape_id)
        region["shapeIds"].append(frame.shape_id)
        p.slide.shapes._spTree.insert(2,frame._element)


def layout(index,variant=None):
    p=Layout(index)
    if index>=7:
        from fifth_layouts import draw
        draw(p,index-7,variant)
    elif index==2:dual(p,variant=="progression")
    elif index==5:branch(p,variant=="two-actions-effects")
    else:[four,triple,None,split,references,None,chain][index](p)
    return p


def native(component,theme,target,variant=None):
    p=layout(component["referenceNumber"]-17,variant)
    stream=BytesIO();p.deck.save(stream);stream.seek(0)
    spec=dict(component, slots=list(p.slots.values()))
    return build_single_slide(stream,spec,theme,target),p


def variant_form(component,spec,theme,target):
    if spec not in component.get("variants",[]):raise ValueError("此组件未登记该变体")
    detail,p=native(component,theme,target,spec["id"])
    return dict(id=spec["id"],name=spec["name"],recipe=spec["recipe"],pageFrame=detail["pageFrame"],
                templateElements=detail["templateElements"],regionShapeIds={r["id"]:r["shapeIds"] for r in p.regions},
                replacedSlotKeys=[s["key"] for s in component["slots"]],addedSlots=list(p.slots.values()),sampleContent=spec["name"])


def register():
    def write(path,value):path.write_text(json.dumps(value,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    path=LIB/"components.json";catalog=json.loads(path.read_text(encoding="utf-8"))
    manifest=LIB/"references/manifest.json";refs=json.loads(manifest.read_text(encoding="utf-8"))
    batch_numbers=range(17,17+len(NAMES))
    catalog["components"]=[c for c in catalog["components"] if c["referenceNumber"] not in batch_numbers]
    refs=[r for r in refs if r["referenceNumber"] not in batch_numbers]
    titles=["传统城市更新中水规划的痛点","前置规划与生态底线","水文化赋能城市更新的路径","陆地生态系统中的富集与传递案例","参考文献","修复核心技术与典型案例","修复核心技术与典型案例"]
    limits=["四项缺少对应案例","只有并列标题而没有可综合的证据","两条路线缺乏可比的案例材料","两项分析没有共同对象","缺少来源信息","行动缺少背景依据","尚无可报告效果"]
    for i,(code,name,task) in enumerate(NAMES):
        p=layout(i)
        variants=[]
        if i==9:variants=[dict(id="bullets",name="无序列表",recipe=dict(kind="quadrant-bullets"))]
        if i==2:variants=[dict(id="progression",name="两条路径递进",recipe=dict(kind="case-progression"))]
        if i==5:variants=[dict(id="two-actions-effects",name="左侧两项行动，右侧展示效果",recipe=dict(kind="case-two-actions"))]
        published = p.reference.parent == LIB / "references"
        c=dict(id=f"BW-{code}-01",name=name,referenceNumber=17+i,referenceLabel=f"第四批图{i+1}",referenceFile=p.reference.name if published else None,
               sourceFile=p.reference.name,sourceKind="reference-reconstruction",sourceSlide=1,titleShapeId=61000,
               task=task,argument=task,fits=[task],doesNotFit=[(limits+["正文论证","新增论据","各项必须配图","缺少结果依据","没有明确实施环节"])[i]],fixed=["统一白底蓝色页框，无页码",*[r["name"] for r in p.regions[1:]]],
               adaptationRules=["按槽位替换文字和图像，保留字体、位置和图片比例。","示例文字取自参考截图；制作时用本次材料核实并替换结论、数据及引文。"],
               regions=p.regions,slots=list(p.slots.values()),variants=variants,
               pageFrame=dict(titleText=(titles+["","","项目风险分析","修复成效、稳定性与推广","工程实施与迭代"])[i],sourceTitleShapeId=None,removeShapeIds=[],translateY=0))
        if i>=7:c["referenceLabel"]=f"第五批图{i-6}"
        if i in (7,8):
            del c["pageFrame"]
            c["titleShapeId"]=None
            c["fixed"]=["蓝白交界与汇报人胶囊","中心大字，首页为题目，结束页为致谢"]
            c["adaptationRules"].append("汇报人、题目和致谢文字按材料替换。")
        if i==9:c["adaptationRules"].append("每个象限均为独立的原生列表，可单独切换编号或项目符号；编号从本象限第一项开始。")
        if i==4:c["adaptationRules"].append("条目数量按实际引用决定；超出页面容量时续页并延续编号。")
        catalog["components"].append(c)
        with Image.open(p.reference) as image:size=list(image.size)
        record = dict(referenceNumber=17+i,referenceBatch="batch-04" if i<7 else "batch-05",batchFigure=i+1 if i<7 else i-6,file=p.reference.name,
                      pixels=size,sha256=hashlib.sha256(p.reference.read_bytes()).hexdigest(),status="用户原图，像素未改写")
        if not published:
            record["sourcePath"] = p.reference.relative_to(LIB.parent).as_posix()
        refs.append(record)
    first={"BW-COVER-01":-3,"BW-CLOSING-01":-2,"BW-TOC-01":-1}
    catalog["components"].sort(key=lambda c:first.get(c["id"],c["referenceNumber"]))
    refs.sort(key=lambda r:r["referenceNumber"])
    if catalog["version"]=="0.4.0":
        catalog["version"]="0.5.0"
        catalog["compatibleTaskVersions"].append("0.4.0")
    if catalog["version"]=="0.5.0":
        catalog["version"]="0.6.0"
        catalog["compatibleTaskVersions"].append("0.5.0")
    write(path,catalog);write(manifest,refs)


if __name__=="__main__":register()
