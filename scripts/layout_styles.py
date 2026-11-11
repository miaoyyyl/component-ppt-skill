"""保留参考组件的渐变、单角标签和蓝框白底结构。"""
from pptx.oxml.xmlchemy import OxmlElement


def clear_effects(shape):
    props = shape._element.spPr
    for node in list(props):
        if node.tag.rsplit("}", 1)[-1] in {"effectLst", "effectDag"}:
            props.remove(node)
    props.append(OxmlElement("a:effectLst"))


def case_marker(layout,x,y,width,height,region_id):
    from pptx.enum.shapes import MSO_SHAPE
    from reference_layouts import BLUE
    marker = layout.rect(x,y,width,height,BLUE,shape_type=MSO_SHAPE.PENTAGON)
    marker.name = "typical-case-marker"
    marker.adjustments[0] = .25
    clear_effects(marker)
    layout.regions[-1]["shapeIds"].remove(marker.shape_id)
    next(r for r in layout.regions if r["id"]==region_id)["shapeIds"].append(marker.shape_id)


def summary_gradient(shape):
    shape.fill.gradient()
    shape.fill.gradient_angle = 0
    fill = shape._element.spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}gradFill")
    stops = fill.find("{http://schemas.openxmlformats.org/drawingml/2006/main}gsLst")
    for child in list(stops):
        stops.remove(child)
    for position, color in ((0, "FFFFFF"), (50000, "E6EFFB"), (100000, "FFFFFF")):
        stop = OxmlElement("a:gs")
        stop.set("pos", str(position))
        value = OxmlElement("a:srgbClr")
        value.set("val", color)
        stop.append(value)
        stops.append(stop)
    shape.line.fill.background()
    clear_effects(shape)


def corner_tab(shape):
    """原图标签左侧为直角，仅右上角圆弧。"""
    props = shape._element.spPr
    old = props.find("{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom")
    geom = OxmlElement("a:custGeom")
    for name in ("avLst", "gdLst", "ahLst", "cxnLst"):
        geom.append(OxmlElement("a:" + name))
    rect = OxmlElement("a:rect")
    for name, value in (("l", "0"), ("t", "0"), ("r", "r"), ("b", "b")):
        rect.set(name, value)
    geom.append(rect)
    paths = OxmlElement("a:pathLst")
    path = OxmlElement("a:path")
    path.set("w", "185000")
    path.set("h", "46000")
    for command, points in (
        ("moveTo", [(0, 0)]),
        ("lnTo", [(157000, 0)]),
        ("cubicBezTo", [(172464, 0), (185000, 12536), (185000, 28000)]),
        ("lnTo", [(185000, 46000)]),
        ("lnTo", [(0, 46000)]),
        ("close", []),
    ):
        node = OxmlElement("a:" + command)
        for x, y in points:
            point = OxmlElement("a:pt")
            point.set("x", str(x))
            point.set("y", str(y))
            node.append(point)
        path.append(node)
    paths.append(path)
    geom.append(paths)
    props.replace(old, geom)
    clear_effects(shape)
    effects = props.find("{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst")
    shadow = OxmlElement("a:outerShdw")
    for name, value in (("blurRad", "76200"), ("dist", "38100"), ("dir", "2700000"), ("rotWithShape", "0")):
        shadow.set(name, value)
    color = OxmlElement("a:srgbClr")
    color.set("val", "000000")
    alpha = OxmlElement("a:alpha")
    alpha.set("val", "18000")
    color.append(alpha)
    shadow.append(color)
    effects.append(shadow)
