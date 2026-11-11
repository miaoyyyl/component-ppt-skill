"""从用户原始 PPT 提取单页组件，保留原生对象及其依赖，不使用外部 skill。"""
from __future__ import annotations

import copy
import colorsys
import posixpath
from pathlib import Path
from urllib.parse import unquote
import zipfile

from lxml import etree as E

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
SHAPES = {"sp", "pic", "grpSp", "cxnSp", "graphicFrame"}
FIXED_TEXT_PREFIX = "component-fixed-text-"


def tag(prefix, name):
    return "{" + NS[prefix] + "}" + name


def xml(data):
    return E.fromstring(data)


def serialize(node):
    return E.tostring(node, encoding="UTF-8", xml_declaration=True, standalone=True)


def shape_id(node):
    n = node.find(".//p:cNvPr", NS)
    return int(n.get("id")) if n is not None else None


def is_fixed_text(node):
    props = node.find("p:nvSpPr/p:cNvPr", NS)
    return props is not None and props.get("name", "").startswith(FIXED_TEXT_PREFIX)


def rels_path(part):
    if not part:
        return "_rels/.rels"
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def resolve_target(part, target):
    path = unquote(target.split("#", 1)[0])
    if path.startswith("/"):
        return posixpath.normpath(path.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), path))


def slide_part(data, slide_number):
    presentation = xml(data["ppt/presentation.xml"])
    entries = presentation.find("p:sldIdLst", NS)
    if slide_number < 1 or slide_number > len(entries):
        raise ValueError(f"源页码越界：{slide_number}")
    rel_id = entries[slide_number - 1].get(tag("r", "id"))
    rels = xml(data[rels_path("ppt/presentation.xml")])
    rel = next(r for r in rels if r.get("Id") == rel_id)
    return resolve_target("ppt/presentation.xml", rel.get("Target")), rel_id


def set_fonts(pr, family):
    bold_face = any(
        any(word in n.get("typeface", "").lower() for word in ("bold", "heavy", "black"))
        for n in pr if E.QName(n).localname in {"latin", "ea", "cs"}
    )
    if bold_face and pr.get("b") is None:
        pr.set("b", "1")
    for n in list(pr):
        if E.QName(n).localname in {"latin", "ea", "cs"}:
            pr.remove(n)
    insert_at = next((i for i, n in enumerate(pr) if E.QName(n).localname in
                      {"sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"}), len(pr))
    for offset, name in enumerate(("latin", "ea", "cs")):
        n = E.Element(tag("a", name), typeface=family)
        pr.insert(insert_at + offset, n)


def solid_text_color(pr, color):
    for n in list(pr):
        if E.QName(n).localname in {"noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"}:
            pr.remove(n)
    fill = E.Element(tag("a", "solidFill"))
    E.SubElement(fill, tag("a", "srgbClr"), val=color)
    at = 1 if len(pr) and E.QName(pr[0]).localname == "ln" else 0
    pr.insert(at, fill)


def normalize_xml(root, theme, *, preserve_colors=False):
    family = theme["typography"]["family"]
    primary = theme["palette"]["primary"]
    for n in root.iter():
        q = E.QName(n)
        if q.namespace != NS["a"]:
            continue
        if q.localname in {"rPr", "defRPr", "endParaRPr"}:
            set_fonts(n, family)
        elif q.localname == "srgbClr" and not preserve_colors:
            val = n.get("val", "")
            if len(val) == 6:
                rgb = tuple(int(val[i:i+2], 16) / 255 for i in (0, 2, 4))
                h, s, v = colorsys.rgb_to_hsv(*rgb)
                if .50 <= h <= .75 and s > .25:
                    n.set("val", primary)
    for run in root.findall(".//a:r", NS):
        pr = run.find("a:rPr", NS)
        if pr is None:
            pr = E.Element(tag("a", "rPr"))
            run.insert(0, pr)
        set_fonts(pr, family)
    for scheme in root.findall(".//a:fontScheme", NS):
        for font in scheme.iter():
            if "typeface" in font.attrib:
                font.set("typeface", family)


def toc_positions(count, theme):
    """净间距包括面板顶部、底部和每两条之间；不把数量固定为6。"""
    if not isinstance(count, int) or isinstance(count, bool) or count < 1:
        raise ValueError("目录必须至少包含1个章节")
    rule = theme["toc"]
    height = min(rule["preferredRowHeight"],
                 (rule["panelHeight"] - (count + 1) * rule["minimumGap"]) / count)
    if height < rule["minimumRowHeight"]:
        raise ValueError(f"{count}项目录超过当前面板在固定字号下的容量，请人工决定分栏或分页")
    gap = (rule["panelHeight"] - count * height) / (count + 1)
    return [{"y": rule["panelY"] + gap + i * (height + gap),
             "height": height, "gap": gap} for i in range(count)]


def xfrm(shape):
    for path in ("p:spPr/a:xfrm", "p:grpSpPr/a:xfrm", "p:xfrm"):
        n = shape.find(path, NS)
        if n is not None:
            return n
    return None


def set_box(shape, x, y, width, height):
    trans = xfrm(shape)
    if trans is None:
        raise ValueError(f"对象{shape_id(shape)}没有原生变换")
    off = trans.find("a:off", NS)
    ext = trans.find("a:ext", NS)
    off.set("x", str(round(x * 9525)))
    off.set("y", str(round(y * 9525)))
    ext.set("cx", str(round(width * 9525)))
    ext.set("cy", str(round(height * 9525)))


def assign_id(shape, new_id, name):
    for node in shape.findall(".//p:cNvPr", NS):
        node.set("id", str(new_id))
        node.set("name", name)
    # 克隆行中的创建标识不是可见内容；移除后避免重复元数据。
    for node in shape.xpath('.//*[local-name()="creationId"]'):
        node.getparent().remove(node)
    # 源文件编辑器的私有标签按单个对象登记，克隆后不能共享同一tags部件。
    for node in shape.findall(".//p:custDataLst", NS):
        node.getparent().remove(node)


def toc_text(shape, value, size_pt, color, family, align):
    body = shape.find("p:txBody", NS)
    props = body.find("a:bodyPr", NS)
    for name in ("lIns", "rIns", "tIns", "bIns"):
        props.set(name, "0")
    props.set("anchor", "ctr")
    props.set("wrap", "none")
    for n in list(props):
        if E.QName(n).localname in {"spAutoFit", "normAutofit", "noAutofit"}:
            props.remove(n)
    props.append(E.Element(tag("a", "noAutofit")))
    for p in list(body.findall("a:p", NS)):
        body.remove(p)
    p = E.SubElement(body, tag("a", "p"))
    E.SubElement(p, tag("a", "pPr"), algn=align)
    run = E.SubElement(p, tag("a", "r"))
    pr = E.SubElement(run, tag("a", "rPr"), sz=str(round(size_pt * 100)), b="1")
    solid_text_color(pr, color)
    set_fonts(pr, family)
    E.SubElement(run, tag("a", "t")).text = value


def validate_toc_labels(labels, active_index):
    if not isinstance(labels, (list, tuple)) or not labels:
        raise ValueError("目录必须至少包含1个章节")
    if not isinstance(active_index, int) or isinstance(active_index, bool) or not 0 <= active_index < len(labels):
        raise ValueError("当前章节索引不在目录范围内")
    if any(not isinstance(s, str) or not s.strip() or "\n" in s or "\r" in s for s in labels):
        raise ValueError("章名必须是非空单行文本")


def rebuild_toc(root, labels, active_index, theme):
    validate_toc_labels(labels, active_index)
    rows = toc_positions(len(labels), theme)
    tree = root.find("p:cSld/p:spTree", NS)
    by_id = {shape_id(n): n for n in tree if E.QName(n).localname in SHAPES}
    # 从用户原生页取当前行和浅灰行，不重新绘制轮廓。
    seeds = {True: [copy.deepcopy(by_id[i]) for i in (166, 168, 169, 167)],
             False: [copy.deepcopy(by_id[i]) for i in (170, 172, 173, 171)]}
    old_ids = {166,168,169,167,170,172,173,171,174,176,177,175,178,180,181,179,
               182,184,185,183,192,194,195,193}
    for node in list(tree):
        if E.QName(node).localname in SHAPES and shape_id(node) in old_ids:
            tree.remove(node)
    palette = theme["palette"]
    for i, (label, row) in enumerate(zip(labels, rows)):
        current = i == active_index
        color = palette["primary"] if current else palette["inactive"]
        background, title, number, line = copy.deepcopy(seeds[current])
        y, h = row["y"], row["height"]
        set_box(background, 500.27, y, 595.74, h)
        set_box(title, 641.95, y, 444.06, h)
        set_box(number, 526.03, y, 96.54, h)
        line_height = min(50.64, h - 16)
        set_box(line, 627.63, y + (h - line_height) / 2, 0, line_height)
        for j, node in enumerate((background, title, number, line)):
            assign_id(node, 20000 + i * 10 + j, f"toc-{i+1:02d}-{('frame','label','number','separator')[j]}")
            tree.append(node)
        line_props = line.find("p:spPr/a:ln", NS)
        if line_props is not None:
            solid_text_color(line_props, color)
        toc_text(title, label, theme["typography"]["tocTitlePt"], color,
                 theme["typography"]["family"], "l")
        toc_text(number, f"{i+1:02d}", theme["typography"]["tocNumberPt"], color,
                 theme["typography"]["family"], "ctr")
    return rows


def shape_inventory(root):
    result = []

    def visit(node, parent_id=None, mapping=(0, 0, 1, 1)):
        kind = E.QName(node).localname
        if kind not in SHAPES:
            return
        ident = shape_id(node)
        trans = xfrm(node)
        box = None
        next_mapping = mapping
        rotation = 0
        if trans is not None:
            off, ext = trans.find("a:off", NS), trans.find("a:ext", NS)
            if off is not None and ext is not None:
                ox, oy, sx, sy = mapping
                x, y = int(off.get("x")), int(off.get("y"))
                w, h = int(ext.get("cx")), int(ext.get("cy"))
                box = [round((ox + x * sx) / 9525, 4), round((oy + y * sy) / 9525, 4),
                       round(w * sx / 9525, 4), round(h * sy / 9525, 4)]
                rotation = int(trans.get("rot", "0")) / 60000
                if kind == "grpSp":
                    co, ce = trans.find("a:chOff", NS), trans.find("a:chExt", NS)
                    if co is not None and ce is not None:
                        nsx = sx * w / max(int(ce.get("cx")), 1)
                        nsy = sy * h / max(int(ce.get("cy")), 1)
                        next_mapping = (ox + x * sx - int(co.get("x")) * nsx,
                                        oy + y * sy - int(co.get("y")) * nsy, nsx, nsy)
        text = "\n".join("".join(t.text or "" for t in p.findall(".//a:t", NS))
                         for p in node.findall("p:txBody/a:p", NS))
        body = node.find("p:txBody/a:bodyPr", NS)
        result.append({"id": ident, "parentId": parent_id, "kind": kind, "boxPx": box,
                       "rotationDegrees": rotation, "text": text,
                       "fontSizesPt": sorted({int(n.get("sz")) / 100 for n in node.findall(".//a:rPr", NS) if n.get("sz")}),
                       "textFrame": dict(body.attrib) if body is not None else None,
                       "descendantTextShapeIds": [shape_id(n) for n in node.findall(".//p:sp", NS) if n.findall("p:txBody/a:p/a:r/a:t", NS)]})
        if kind == "grpSp":
            for child in node:
                visit(child, ident, next_mapping)

    for child in root.find("p:cSld/p:spTree", NS):
        visit(child)
    return result


def build_single_slide(source, component, theme, output, labels=None, active_index=0):
    with zipfile.ZipFile(source) as z:
        data = {name: z.read(name) for name in z.namelist()}
    selected, rel_id = slide_part(data, component["sourceSlide"])
    source_slide = xml(data[selected])
    from source_modules import select_source_module, translate_source_module
    module_rule = select_source_module(source_slide, component)
    original_inventory = shape_inventory(source_slide)
    translate_source_module(source_slide, module_rule)
    layout_inventory = shape_inventory(source_slide) if module_rule else None
    if component.get("nativeAdditions"):
        from component_variants import add_source_copies
        add_source_copies(source_slide, component)
    toc_rows = None
    if component["id"] == "BW-TOC-01":
        if labels is None:
            original_map = {s["id"]: s for s in original_inventory}
            labels = [original_map[i]["text"] for i in (168,172,176,180,184,194)]
        toc_rows = rebuild_toc(source_slide, labels, active_index, theme)
    elif component.get("tocLayout"):
        from suite_toc import rebuild_suite_toc
        toc_rows, labels = rebuild_suite_toc(source_slide, component["tocLayout"], theme, labels, active_index, source=True)
    for shape in source_slide.findall(".//p:sp", NS):
        if shape_id(shape) in component.get("fixedTextShapeIds", []):
            shape.find("p:nvSpPr/p:cNvPr", NS).set("name", FIXED_TEXT_PREFIX + str(shape_id(shape)))
        if shape_id(shape) == component.get('pageFrame',{}).get('sourceTitleShapeId',component.get("titleShapeId")):
            for run in shape.findall(".//a:r", NS):
                pr = run.find("a:rPr", NS)
                if pr is None:
                    pr = E.Element(tag("a", "rPr"))
                    run.insert(0, pr)
                pr.set("sz", str(theme["typography"]["titlePt"] * 100))
    preserve_colors = component.get("preserveArtworkColors", False)
    normalize_xml(source_slide, theme, preserve_colors=preserve_colors)
    for rule in component.get("colorOverrides", []):
        node = next(n for n in source_slide.iter() if E.QName(n).localname in SHAPES and shape_id(n) == rule["shapeId"])
        if rule["mode"] == "blue-structure":
            for color in node.findall(".//a:srgbClr", NS):
                color.set("val", theme["palette"]["primary"])
            for color in node.findall(".//a:schemeClr", NS):
                color.tag = tag("a", "srgbClr")
                color.set("val", theme["palette"]["primary"])
        elif rule["mode"] in {"yellow-text","blue-text"}:
            for pr in node.findall(".//a:rPr", NS):
                solid_text_color(pr, theme["palette"]["emphasisYellow" if rule["mode"]=="yellow-text" else "primary"])
    for rule in component.get("textFrameOverrides",[]):
        for ident in rule["shapeIds"]:
            node=next(n for n in source_slide.findall(".//p:sp",NS) if shape_id(n)==ident)
            props=node.find("p:txBody/a:bodyPr",NS)
            if "centeredWidthScale" in rule:
                trans=xfrm(node)
                offset=trans.find("a:off",NS)
                extent=trans.find("a:ext",NS)
                width=int(extent.get("cx"))
                new_width=round(width*rule["centeredWidthScale"])
                offset.set("x",str(int(offset.get("x"))+round((width-new_width)/2)))
                extent.set("cx",str(new_width))
            if rule.get("zeroInsets"):
                for attribute in ("lIns","rIns","tIns","bIns"):props.set(attribute,"0")
            if rule.get("singleLine"):
                props.set("wrap","none");props.set("anchor","ctr")
                for child in list(props):
                    if E.QName(child).localname in {"normAutofit","spAutoFit","noAutofit"}:props.remove(child)
                props.append(E.Element(tag("a","noAutofit")))
            for paragraph in node.findall("p:txBody/a:p",NS):
                pp=paragraph.find("a:pPr",NS)
                if pp is None:pp=E.Element(tag("a","pPr"));paragraph.insert(0,pp)
                pp.set("algn",rule.get("align","ctr"))
                if rule.get("zeroInsets"):
                    pp.set("marL","0");pp.set("indent","0")
    strong_roles = {"page-title", "lead", "lead-definition", "yellow-label", "red-subtitle",
                    "local-conclusion", "module-title", "page-conclusion", "method-title",
                    "figure-title", "model-class", "mechanism-title", "diagnostic-takeaway", "development-nodes"}
    strong_ids = {i for slot in component["slots"] if slot["role"] in strong_roles for i in slot["shapeIds"]}
    strong_ids.update(i for slot in component["slots"] if slot["key"] in component.get("strongSlotKeys",[]) for i in slot["shapeIds"])
    for node in source_slide.iter():
        if E.QName(node).localname in SHAPES and shape_id(node) in strong_ids:
            for pr in node.findall(".//a:rPr", NS):
                pr.set("b", "1")
    for pr in source_slide.findall(".//a:rPr", NS):
        color = pr.find("a:solidFill/a:srgbClr", NS)
        if color is not None and color.get("val") in {theme["palette"]["emphasisRed"], theme["palette"]["emphasisYellow"]}:
            pr.set("b", "1")
    frame_layout=None
    if component.get('pageFrame'):
        from page_frame import apply_page_frame
        frame_layout=apply_page_frame(source_slide,component,theme)
    generated_ids=[]
    if component.get("contentSummary"):
        from page_frame import append_content_summary
        generated_ids.append(append_content_summary(source_slide,component["contentSummary"],theme))
    for slot in component["slots"]:
        if slot.get("chartShapeIds"):
            from component_variants import find_shape
            from progress_rings import set_ring_percentage
            for text_id, chart_id in zip(slot["shapeIds"], slot["chartShapeIds"], strict=True):
                value = "".join(find_shape(source_slide, text_id).xpath(".//a:t/text()", namespaces=NS))
                set_ring_percentage(data, selected, find_shape(source_slide, chart_id), value)
    data[selected] = serialize(source_slide)
    presentation = xml(data["ppt/presentation.xml"])
    ids = presentation.find("p:sldIdLst", NS)
    for entry in list(ids):
        if entry.get(tag("r", "id")) != rel_id:
            ids.remove(entry)
    for child in list(presentation):
        if E.QName(child).localname in {"custShowLst", "extLst", "embeddedFontLst"}:
            presentation.remove(child)
    presentation.set("embedTrueTypeFonts", "0")
    data["ppt/presentation.xml"] = serialize(presentation)
    kept = {}

    def walk(part):
        if part in kept:
            return
        if part:
            if part not in data:
                raise ValueError(f"PPT依赖缺失：{part}")
            payload = data[part]
            if part.endswith(".xml") and part != selected:
                node = xml(payload)
                if component.get('pageFrame') and ('/slideMasters/' in part or '/slideLayouts/' in part):
                    from page_frame import clean_inherited_graphics
                    clean_inherited_graphics(node)
                normalize_xml(node, theme, preserve_colors=preserve_colors)
                if part == "docProps/app.xml":
                    for item in node:
                        if E.QName(item).localname == "Slides":
                            item.text = "1"
                payload = serialize(node)
            kept[part] = payload
        relation_path = rels_path(part)
        if relation_path not in data:
            return
        relations = xml(data[relation_path])
        targets = []
        for relation in list(relations):
            typ = relation.get("Type", "")
            if typ.endswith(("/thumbnail", "/font")):
                relations.remove(relation)
                continue
            if relation.get("TargetMode") == "External":
                continue
            target = resolve_target(part, relation.get("Target", ""))
            if typ.endswith("/slide") and target != selected:
                relations.remove(relation)
                continue
            targets.append(target)
        kept[relation_path] = serialize(relations)
        for target in targets:
            walk(target)

    walk("")
    types = xml(data["[Content_Types].xml"])
    for entry in list(types):
        if E.QName(entry).localname == "Override" and entry.get("PartName", "").lstrip("/") not in kept:
            types.remove(entry)
    kept["[Content_Types].xml"] = serialize(types)
    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(kept):
            info = zipfile.ZipInfo(name, date_time=(2026, 11, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, kept[name])
    final_inventory = shape_inventory(source_slide)
    result = {"sourcePart": selected, "sourceElements": original_inventory,
            "templateElements": final_inventory, "tocRows": toc_rows,
            "tocLabels": labels, "activeIndex": active_index if toc_rows else None,'pageFrame':frame_layout}
    if generated_ids:result["generatedShapeIds"]=generated_ids
    if module_rule:
        result.update({"sourceSelection": module_rule, "layoutElements": layout_inventory})
    return result
