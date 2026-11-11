"""组件母版配色替换和主题资产定位。"""
from colorsys import hls_to_rgb, rgb_to_hls
from functools import lru_cache
import json
from pathlib import Path

BASE_PALETTE = "blue"
BLUE = "0046A5"
MOUNTAIN_ACTIVE = "C00000"
RED = "A40808"
INACTIVE = "D9D9D9"

# 只处理导航页中直接嵌有蓝调的装饰图。林景已有可换色的形状蒙版。
TINTED_PICTURES = {
    "BW-COVER-01": {"slides": {8}},
    "BW-TOC-01": {"slides": {2058}, "slideMasters": {3076}},
    "BW-CLOSING-01": {"slides": {9}},
    "BW-COVER-MOUNTAIN-01": {"slides": {4, 15}},
    "BW-TOC-MOUNTAIN-01": {"slides": {214, 216}},
    "BW-CLOSING-MOUNTAIN-01": {"slides": {4, 15}},
    "BW-COVER-TECH-01": {"slideMasters": {9, 15}},
    "BW-TOC-TECH-01": {"slides": {57}, "slideMasters": {9, 15}},
    "BW-CLOSING-TECH-01": {"slideMasters": {9, 15}},
    "BW-COVER-WAVE-01": {"slides": {34}},
    "BW-TOC-WAVE-01": {"slides": {35}},
    "BW-CLOSING-WAVE-01": {"slides": {34}},
}
DRAWING_FOLDERS = {"slides", "slideLayouts", "slideMasters", "theme", "charts"}


def selected_palette(task, catalog):
    ident = task.get("paletteId", BASE_PALETTE)
    if ident != BASE_PALETTE and not any(ident == item["id"] for item in catalog.get("palettes", [])):
        raise ValueError("任务配色未在组件库中登记")
    return ident


def theme_path(library, ident):
    return Path(library) / ("theme.json" if ident == BASE_PALETTE else f"palettes/{ident}/theme.json")


def load_scheme(library, ident):
    if ident == BASE_PALETTE:
        return None
    manifest = json.loads((Path(library) / "palettes/manifest.json").read_text(encoding="utf-8"))
    return manifest["schemes"][ident]


def palette_form(scheme, key, mode, source):
    form = scheme["forms"].get(key, {}).get(mode)
    if not form or form["sourceTemplateSha256"] != source["templateSha256"]:
        raise ValueError(f"配色母版缺失或需要随原版更新：{key} / {mode}")
    return form


def as_hls(value):
    return rgb_to_hls(*(int(value[i:i + 2], 16) / 255 for i in (0, 2, 4)))


@lru_cache(maxsize=4096)
def recolor(value, primary):
    if primary == BLUE or len(value) != 6:
        return value
    hue, light, saturation = as_hls(value)
    if saturation == 0 or not 186 <= hue * 360 <= 262:
        return value
    _, blue_light, _ = as_hls(BLUE)
    target_hue, target_light, target_saturation = as_hls(primary)
    if light < blue_light:
        light = light * target_light / blue_light
    else:
        light = target_light + (light - blue_light) * (1 - target_light) / (1 - blue_light)
    channels = hls_to_rgb(target_hue, light, saturation * target_saturation)
    return "".join(f"{round(channel * 255):02X}" for channel in channels)


def tint_decoration(blip, primary):
    from lxml import etree
    from native_pptx import tag

    # 用原生双色效果保留背景纹理、原图字节、裁切和透明度。
    channels = tuple(int(primary[i:i + 2], 16) for i in (0, 2, 4))
    dark = "".join(f"{round(value * 0.55):02X}" for value in channels)
    light = "".join(f"{round(value * 0.45 + 255 * 0.55):02X}" for value in channels)
    effect = etree.Element(tag("a", "duotone"))
    etree.SubElement(effect, tag("a", "srgbClr"), val=dark)
    etree.SubElement(effect, tag("a", "srgbClr"), val=light)
    blip.insert(0, effect)


def fade_mountain_rows(root):
    from native_pptx import shape_id, tag

    rows = {}
    for shape in root.iter(tag("p", "sp")):
        ident = shape_id(shape)
        if ident is not None and 20000 <= ident < 20100:
            rows.setdefault((ident - 20000) // 10, []).append(shape)
    active = [key for key, shapes in rows.items() if any(
        node.get("val") == MOUNTAIN_ACTIVE
        for shape in shapes for node in shape.iter(tag("a", "srgbClr"))
    )]
    if len(active) != 1:
        raise ValueError("山景目录缺少唯一的当前章节")
    for key, shapes in rows.items():
        if key == active[0]:
            continue
        for shape in shapes:
            for node in shape.iter(tag("a", "srgbClr")):
                node.set("val", recolor(node.get("val"), INACTIVE))
            for node in shape.iter(tag("a", "schemeClr")):
                if node.get("val") == "accent6":
                    node.tag = tag("a", "srgbClr")
                    node.set("val", INACTIVE)


def recolor_package(package, component_id, primary):
    from native_pptx import NS, serialize, shape_id, tag, xml

    tint_shapes = TINTED_PICTURES.get(component_id, {})
    for name, payload in package.items():
        parts = name.split("/")
        if len(parts) != 3 or parts[0] != "ppt" or parts[1] not in DRAWING_FOLDERS or not name.endswith(".xml"):
            continue
        root = xml(payload)
        gray_mountain_rows = primary == RED and component_id == "BW-TOC-MOUNTAIN-01" and parts[1] == "slides"
        if gray_mountain_rows:
            fade_mountain_rows(root)
        changed = gray_mountain_rows
        for node in root.iter(tag("a", "srgbClr")):
            original = node.get("val")
            color = recolor(original, primary)
            if color != original:
                node.set("val", color)
                changed = True
        for picture in root.iter(tag("p", "pic")):
            if shape_id(picture) in tint_shapes.get(parts[1], ()):
                blip = picture.find("p:blipFill/a:blip", NS)
                if blip is not None:
                    tint_decoration(blip, primary)
                    changed = True
        if changed:
            package[name] = serialize(root)
    return package


def suite_label_color(layout, theme, current):
    primary = theme["palette"]["primary"]
    colors = layout["labelColors"]
    # 红白山景目录用浅灰弱化其余章节，保持当前红色章节可辨认。
    if primary == RED and colors[0] == MOUNTAIN_ACTIVE and not current:
        primary = theme["palette"]["inactive"]
    return recolor(colors[0 if current else 1], primary)
