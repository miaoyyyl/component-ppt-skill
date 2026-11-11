"""选择用户截图对应的局部结构；只平移，不改变其内部大小和对齐。"""
from lxml import etree as E

from native_pptx import NS, SHAPES, shape_id, xfrm


def select_source_module(root, component):
    rule = component.get("sourceSelection")
    if not rule:
        return None
    tree = root.find("p:cSld/p:spTree", NS)
    selected = set(rule["topLevelShapeIds"])
    found = set()
    for node in list(tree):
        if E.QName(node).localname not in SHAPES:
            continue
        ident = shape_id(node)
        if ident not in selected:
            tree.remove(node)
        else:
            found.add(ident)
    if selected != found:
        raise ValueError(f"截图区域的原生对象缺失：{selected-found}")
    if rule.get("hideMasterGraphics"):
        root.set("showMasterSp", "0")
    return rule


def translate_source_module(root, rule):
    if not rule:
        return
    dx, dy = rule.get("translateX", 0), rule.get("translateY", 0)
    for node in root.find("p:cSld/p:spTree", NS):
        if E.QName(node).localname not in SHAPES:
            continue
        trans = xfrm(node)
        if trans is None:
            raise ValueError(f"局部组件对象没有位移属性：{shape_id(node)}")
        off = trans.find("a:off", NS)
        off.set("x", str(int(off.get("x")) + round(dx * 9525)))
        off.set("y", str(int(off.get("y")) + round(dy * 9525)))
