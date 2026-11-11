"""统一内容页页框：显式白底、原生页标题、双层蓝隔断和底条；正文不缩字。"""
from __future__ import annotations
import copy
import functools
import zipfile
from pathlib import Path
from lxml import etree as E

from native_pptx import (NS,SHAPES,tag,xml,shape_id,shape_inventory,slide_part,rels_path,
                         resolve_target,xfrm,set_box,set_fonts,solid_text_color)

ROOT=Path(__file__).resolve().parents[1]
TITLE_ID=61000
DIVIDER_ID=61001
FOOTER_ID=61002
FRAME_IDS={TITLE_ID,DIVIDER_ID,FOOTER_ID,61003,61004}


@functools.lru_cache(maxsize=4)
def frame_seeds(source_file,source_slide):
    with zipfile.ZipFile(ROOT/'materials'/source_file) as z:data={n:z.read(n) for n in z.namelist()}
    part,_=slide_part(data,source_slide)
    layout=resolve_target(part,next(r for r in xml(data[rels_path(part)]) if r.get('Type','').endswith('/slideLayout')).get('Target'))
    tree=xml(data[layout]).find('p:cSld/p:spTree',NS)
    return tuple(copy.deepcopy(next(n for n in tree if shape_id(n)==ident)) for ident in (7,12))


def _shift(node,dy):
    trans=xfrm(node)
    if trans is None:return
    off=trans.find('a:off',NS)
    if off is not None:off.set('y',str(int(off.get('y'))+round(dy*9525)))


def _translation(config,ident):
    return config.get('translateY',0)+config.get('extraTranslateY',{}).get(str(ident),0)


def _resize(node,size):
    ext=xfrm(node).find('a:ext',NS)
    ext.set('cx',str(round(size[0]*9525)));ext.set('cy',str(round(size[1]*9525)))


def restore_content_coordinates(root,component):
    """变体在源内容坐标中生成，再重施页框，避免固定配方坐标被重复平移。"""
    tree=root.find('p:cSld/p:spTree',NS)
    if not any(shape_id(n)==TITLE_ID for n in tree):return
    config=component['pageFrame'];sizes={r['shapeId']:r for r in config.get('sizeOverrides',[])}
    for node in list(tree):
        if E.QName(node).localname not in SHAPES:continue
        ident=shape_id(node)
        if ident in FRAME_IDS:tree.remove(node);continue
        _shift(node,-_translation(config,ident))
        if ident in sizes:_resize(node,sizes[ident]['originalSize'])


def _page_number(node):
    return bool(node.xpath('.//p:ph[@type="sldNum"] | .//a:fld[@type="slidenum"]',namespaces=NS))


def clean_inherited_graphics(root):
    """保留母版文字样式占位，仅移除会覆盖统一页框的图形和页脚占位。"""
    tree=root.find('p:cSld/p:spTree',NS)
    if tree is None:return
    for node in list(tree):
        if E.QName(node).localname not in SHAPES:continue
        ph=node.find('.//p:ph',NS)
        if ph is None or ph.get('type') in {'sldNum','dt','ftr'}:tree.remove(node)


def _literal_blue(node,primary):
    for style in list(node.findall('.//p:style',NS)):style.getparent().remove(style)
    for color in node.findall('.//a:schemeClr',NS):
        color.tag=tag('a','srgbClr');color.set('val',primary)
    for color in node.findall('.//p:spPr/a:solidFill/a:srgbClr',NS):color.set('val',primary)


def title_shape(text,theme):
    node=E.Element(tag('p','sp'));nv=E.SubElement(node,tag('p','nvSpPr'))
    E.SubElement(nv,tag('p','cNvPr'),id=str(TITLE_ID),name='page-frame-title')
    E.SubElement(nv,tag('p','cNvSpPr'),txBox='1');E.SubElement(nv,tag('p','nvPr'))
    props=E.SubElement(node,tag('p','spPr'));trans=E.SubElement(props,tag('a','xfrm'))
    E.SubElement(trans,tag('a','off'),x='0',y='0');E.SubElement(trans,tag('a','ext'),cx='0',cy='0')
    E.SubElement(E.SubElement(props,tag('a','prstGeom'),prst='rect'),tag('a','avLst'))
    E.SubElement(props,tag('a','noFill'));E.SubElement(E.SubElement(props,tag('a','ln')),tag('a','noFill'))
    set_box(node,*theme['contentPageFrame']['titleBox'])
    body=E.SubElement(node,tag('p','txBody'))
    E.SubElement(E.SubElement(body,tag('a','bodyPr'),wrap='none',lIns='0',rIns='0',tIns='0',bIns='0',anchor='ctr'),tag('a','noAutofit'))
    E.SubElement(body,tag('a','lstStyle'));paragraph=E.SubElement(body,tag('a','p'))
    E.SubElement(paragraph,tag('a','pPr'),algn='l')
    run=E.SubElement(paragraph,tag('a','r'));pr=E.SubElement(run,tag('a','rPr'),sz=str(theme['typography']['titlePt']*100),b='1')
    solid_text_color(pr,theme['palette']['primary']);set_fonts(pr,theme['typography']['family'])
    E.SubElement(run,tag('a','t')).text=text
    return node


def append_content_summary(root,config,theme):
    """在原生内容下方增加用户指定的总述槽位，独立于页标题和底条。"""
    node=title_shape(config["text"],theme)
    nv=node.find("p:nvSpPr/p:cNvPr",NS)
    nv.set("id",str(config["shapeId"]))
    nv.set("name","content-summary")
    set_box(node,*config["box"])
    props=node.find("p:spPr",NS)
    props.remove(props.find("a:noFill",NS))
    fill=E.Element(tag("a","solidFill"))
    E.SubElement(fill,tag("a","srgbClr"),val=theme["palette"]["surface"])
    props.insert(2,fill)
    line=props.find("a:ln",NS)
    line.remove(line.find("a:noFill",NS))
    line.set("w","9525")
    E.SubElement(E.SubElement(line,tag("a","solidFill")),tag("a","srgbClr"),val=theme["palette"]["primary"])
    body=node.find("p:txBody/a:bodyPr",NS)
    body.set("wrap","square")
    body.set("lIns","95250")
    body.set("rIns","95250")
    node.find("p:txBody/a:p/a:pPr",NS).set("algn","ctr")
    node.find("p:txBody/a:p/a:r/a:rPr",NS).set("sz",str(config["fontPt"]*100))
    root.find("p:cSld/p:spTree",NS).append(node)
    return config["shapeId"]


def apply_page_frame(root,component,theme):
    if not component.get('pageFrame'):return None
    config=component['pageFrame'];tree=root.find('p:cSld/p:spTree',NS)
    if any(shape_id(n)==TITLE_ID for n in tree):restore_content_coordinates(root,component)
    before={n['id']:n for n in shape_inventory(root)};removed=[]
    remove_ids=set(config.get('removeShapeIds',[]))
    for node in list(tree):
        if E.QName(node).localname not in SHAPES:continue
        ident=shape_id(node)
        if ident in remove_ids or _page_number(node):tree.remove(node);removed.append(ident)
        else:_shift(node,_translation(config,ident))
    for rule in config.get('sizeOverrides',[]):
        node=next((n for n in tree if shape_id(n)==rule['shapeId']),None)
        if node is not None:_resize(node,rule['size'])
    common=root.find('p:cSld',NS);old=common.find('p:bg',NS)
    if old is not None:common.remove(old)
    bg=E.Element(tag('p','bg'));props=E.SubElement(bg,tag('p','bgPr'))
    E.SubElement(E.SubElement(props,tag('a','solidFill')),tag('a','srgbClr'),val=theme['palette']['white']);E.SubElement(props,tag('a','effectLst'))
    common.insert(0,bg);root.set('showMasterSp','0')
    frame=theme['contentPageFrame'];divider,footer=(copy.deepcopy(n) for n in frame_seeds(frame['sourceFile'],frame['sourceSlide']))
    for ident,node in [(DIVIDER_ID,divider),(FOOTER_ID,footer)]:
        for index,nv in enumerate(node.findall('.//p:cNvPr',NS)):
            nv.set('id',str(ident if index==0 else 61002+index));nv.set('name',('page-frame-divider' if ident==DIVIDER_ID else 'page-frame-footer')+(f'-{index}' if index else ''))
        _literal_blue(node,theme['palette']['primary']);tree.append(node)
    tree.append(title_shape(config['titleText'],theme))
    after={n['id']:n for n in shape_inventory(root)}
    changes=[{'shapeId':ident,'before':old['boxPx'],'after':after[ident]['boxPx']} for ident,old in before.items() if ident in after and old['boxPx']!=after[ident]['boxPx']]
    return {'id':frame['id'],'titleShapeId':TITLE_ID,'removedShapeIds':removed,'geometryChanges':changes,'bodyFontSizesPreserved':True,'background':'FFFFFF','pageNumbers':False}
