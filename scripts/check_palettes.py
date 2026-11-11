"""核对三色母版的文件哈希、原生结构、图片字节与渲染尺寸。"""
from zipfile import ZipFile

from lxml import etree
from PIL import Image, ImageChops

from build_palettes import LIB, registered_forms
from native_pptx import NS, shape_id, tag, xml
from palette_assets import BASE_PALETTE, DRAWING_FOLDERS, RED, TINTED_PICTURES, load_scheme, palette_form, theme_path
from production_tasks import digest, read_json


def drawing_signature(payload, component, folder, primary):
    root = xml(payload)
    pictures = TINTED_PICTURES.get(component, {}).get(folder, ())
    for picture in root.iter(tag("p", "pic")):
        if shape_id(picture) in pictures:
            for tint in picture.findall("p:blipFill/a:blip/a:duotone", NS):
                tint.getparent().remove(tint)
    for node in root.iter(tag("a", "srgbClr"), tag("a", "schemeClr")):
        value = node.get("val")
        if component == "BW-TOC-MOUNTAIN-01" and folder == "slides" and primary == RED and value == "accent6":
            shape = next(node.iterancestors(tag("p", "sp")), None)
            ident = shape_id(shape) if shape is not None else None
            if ident is not None and 20010 <= ident < 20100:
                node.tag = tag("a", "srgbClr")
        if node.tag == tag("a", "srgbClr") and value not in {"FFFFFF", "000000", "C00000", "FFFF00"}:
            # 白黑、红黄强调及其余 XML 属性与关系保持原样。
            node.set("val", "drawing-color")
    return etree.tostring(root, method="c14n")


def check():
    catalog = read_json(LIB / "components.json")
    sources = list(registered_forms())
    canvas = read_json(LIB / "theme.json")["canvas"]
    dimensions = canvas["width"], canvas["height"]
    count = 0
    for spec in catalog["palettes"]:
        palette = spec["id"]
        if palette == BASE_PALETTE:
            continue
        scheme = load_scheme(LIB, palette)
        theme = read_json(theme_path(LIB, palette))
        assert theme["palette"]["primary"] == spec["primary"], palette
        assert set(scheme["forms"]) == {key for _, key, _, _, _ in sources}, palette
        for component, key, mode, source, source_sha in sources:
            asset = palette_form(scheme, key, mode, {"templateSha256": source_sha})
            pptx, png = LIB / asset["templatePath"], LIB / asset["previewPath"]
            assert digest(pptx.read_bytes()) == asset["templateSha256"], pptx
            assert digest(png.read_bytes()) == asset["previewSha256"], png
            with ZipFile(source) as before, ZipFile(pptx) as after:
                assert set(before.namelist()) == set(after.namelist()), pptx
                for part in before.namelist():
                    original, themed = before.read(part), after.read(part)
                    if original == themed:
                        continue
                    sections = part.split("/")
                    assert len(sections) == 3 and sections[0] == "ppt" and sections[1] in DRAWING_FOLDERS and part.endswith(".xml"), (pptx, part)
                    assert drawing_signature(original, component, sections[1], spec["primary"]) == drawing_signature(
                        themed, component, sections[1], spec["primary"]), (pptx, part)
            base_png = LIB / ("previews" if mode == "case" else "placeholders/previews") / (key + ".png")
            with Image.open(png) as colored, Image.open(base_png) as blue:
                assert colored.size == blue.size == dimensions, png
                assert ImageChops.difference(colored.convert("RGB"), blue.convert("RGB")).getbbox(), png
            count += 1
    return count


if __name__ == "__main__":
    print(f"配色检查通过。{check()} 份母版与预览的哈希、结构和尺寸已核对")
