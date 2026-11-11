"""重建组件母版、授权变体、原生对象清单与离线目录数据。"""
import argparse
import hashlib
import json
from pathlib import Path

from native_pptx import build_single_slide
from component_variants import build_variant
from palette_assets import BASE_PALETTE, recolor_package

ROOT = Path(__file__).resolve().parents[1]
LIB = ROOT / "library"


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def bindings_for(slots, elements):
    bindings = []
    for slot in slots:
        ids = slot["shapeIds"]
        bound = [s for s in elements if s["id"] in ids]
        bindings.append({"key":slot["key"],"templateShapeIds":ids,
            "textShapeIds":[i for s in bound for i in ([s["id"]] if s["text"] else s["descendantTextShapeIds"])],
            "boxesPx":[s["boxPx"] for s in bound],
            "fontSizesPt":sorted({pt for s in bound for pt in s["fontSizesPt"]})})
    return bindings


def build(only=None):
    catalog = read_json(LIB / "components.json")
    theme = read_json(LIB / "theme.json")
    inventories = []
    previous = {}
    if only:
        previous = {c["id"]:c for c in read_json(LIB / "inventory.json")["components"]}
        unknown = set(only) - {c["id"] for c in catalog["components"]}
        if unknown:
            raise ValueError(f"组件未登记：{sorted(unknown)}")
    for component in catalog["components"]:
        ident = component["id"]
        if only and ident not in only:
            inventories.append(previous[ident])
            continue
        source = ROOT / "materials" / component["sourceFile"]
        target = LIB / "templates" / f"{ident}.pptx"
        if component.get("sourceKind") == "reference-reconstruction":
            from reference_layouts import native
            detail, layout = native(component, theme, target)
            source = layout.reference
            component["regions"] = layout.regions
            component["slots"] = list(layout.slots.values())
        else:
            detail = build_single_slide(source, component, theme, target)
        detail["slotBindings"] = []
        for slot in component["slots"]:
            ids = slot["shapeIds"]
            if detail["tocRows"] and slot["key"] in {"chapters", "chapterNumbers"}:
                offset = 1 if slot["key"] == "chapters" else 2
                ids = [20000 + i * 10 + offset for i in range(len(detail["tocRows"]))]
            bound = [s for s in detail["templateElements"] if s["id"] in ids]
            detail["slotBindings"].append({"key": slot["key"], "templateShapeIds": ids,
                "textShapeIds": [i for s in bound for i in ([s["id"]] if s["text"] else s["descendantTextShapeIds"])],
                "boxesPx": [s["boxPx"] for s in bound],
                "fontSizesPt": sorted({pt for s in bound for pt in s["fontSizesPt"]})})
        detail.update({"id": ident, "sourceSha256": sha256(source),
                       "templateSha256": sha256(target),
                       "sourcePath": source.relative_to(ROOT).as_posix(),
                       "templatePath": target.relative_to(LIB).as_posix()})
        if component.get("variants"):
            detail["variants"] = []
            for spec in component["variants"]:
                file_id = ident + "__" + spec["id"]
                variant_target = LIB / "templates" / f"{file_id}.pptx"
                variant = build_variant(target,component,spec,theme,LIB,variant_target)
                slots = [s for s in component["slots"] if s["key"] not in variant["replacedSlotKeys"]] + variant["addedSlots"]
                variant.update({"templatePath":variant_target.relative_to(LIB).as_posix(),
                    "previewPath":f"previews/{file_id}.png","templateSha256":sha256(variant_target),
                    "slotBindings":bindings_for(slots,variant["templateElements"])})
                detail["variants"].append(variant)
        inventories.append(detail)
        print(f"已构建 {ident}：{target.stat().st_size:,} bytes")
    inventory = {"version": catalog["version"], "generatedBy": "scripts/build_library.py",
                 "components": inventories}
    (LIB / "inventory.json").write_text(json.dumps(inventory, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (LIB / "components.json").write_text(json.dumps(catalog, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    write_viewer_data(catalog, theme, inventory)


def write_viewer_data(catalog=None, theme=None, inventory=None):
    catalog = catalog or read_json(LIB / "components.json")
    theme = theme or read_json(LIB / "theme.json")
    inventory = inventory or read_json(LIB / "inventory.json")
    preview_paths = {c["id"]:f'previews/{c["id"]}.png' for c in catalog["components"]}
    for c in inventory["components"]:
        for variant in c.get("variants",[]):
            preview_paths[c["id"]+"__"+variant["id"]] = variant["previewPath"]
    report_path=ROOT/"docs/powerpoint-render.json"
    reports=read_json(report_path).get("items",{}) if report_path.exists() else {}
    def preview_current(key,value):
        record=reports.get(key)
        if not record or not (LIB/value).exists():
            return False
        return (ROOT/record["input"]).exists() and sha256(ROOT/record["input"])==record["inputSha256"] and sha256(LIB/value)==record["previewSha256"]
    payload = {"catalog": catalog, "theme": theme, "inventory": inventory,
               "renderedPreviews": {key:preview_current(key,value) for key,value in preview_paths.items()}}
    placeholder_path=LIB/'placeholders/manifest.json';placeholder_report=ROOT/'docs/placeholder-render.json'
    placeholders=read_json(placeholder_path).get('items',[]) if placeholder_path.exists() else []
    placeholder_records=read_json(placeholder_report).get('items',{}) if placeholder_report.exists() else {}
    source_hashes={c['id']:c['templateSha256'] for c in inventory['components']}
    for c in inventory['components']:
        for v in c.get('variants',[]):source_hashes[c['id']+'__'+v['id']]=v['templateSha256']
    payload['placeholderForms']={p['id']:p for p in placeholders}
    payload['placeholderPreviews']={p['id']:bool(p['sourceTemplateSha256']==source_hashes.get(p['id']) and (LIB/p['templatePath']).exists() and (LIB/p['previewPath']).exists()
        and sha256(LIB/p['templatePath'])==p['templateSha256'] and placeholder_records.get(p['id'],{}).get('inputSha256')==p['templateSha256']
        and placeholder_records.get(p['id'],{}).get('previewSha256')==sha256(LIB/p['previewPath'])) for p in placeholders}
    (LIB / "data.js").write_text("window.COMPONENT_LIBRARY = " + json.dumps(payload, ensure_ascii=False) + ";\n", encoding="utf-8")


def main():
    catalog = read_json(LIB / "components.json")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--toc-labels", type=Path, help="包含章名数组的JSON文件")
    parser.add_argument("--active", type=int, default=1, help="当前章序号，从1开始")
    parser.add_argument("--output", type=Path, help="自定义目录页的输出PPTX")
    parser.add_argument("--palette", choices=[scheme["id"] for scheme in catalog["palettes"]],
                        default=BASE_PALETTE, help="自定义目录或变体的配色")
    parser.add_argument("--viewer-only", action="store_true")
    parser.add_argument("--only", nargs="+", help="仅重建这些组件，保留其他母版及清单")
    parser.add_argument("--component", help="自定义目录或导出变体所使用的组件ID")
    parser.add_argument("--variant", help="登记的变体ID")
    parser.add_argument("--recipe-file",type=Path,help="本组件已授权的内容形式JSON，例如四栏图文选择")
    parser.add_argument("--panel-modes", help="三栏形式，逗号分隔的image-text/nested-text")
    parser.add_argument("--panel-images",type=Path,help="以栏序号1/2/3为键、真实图片路径为值的JSON")
    parser.add_argument("--connection-mode",choices=["related","parallel"],default="related")
    parser.add_argument("--bullet",choices=["◆","•","▪","–"],help="多层列表的无序符号")
    args = parser.parse_args()
    scheme = next(s for s in catalog["palettes"] if s["id"] == args.palette)
    if args.palette != BASE_PALETTE and not (args.toc_labels or args.component):
        parser.error("--palette 用于导出自定义目录或变体")
    if args.toc_labels:
        if not args.output:
            parser.error("--toc-labels 必须同时提供 --output")
        ident = args.component or "BW-TOC-01"
        component = next((c for c in catalog["components"] if c["id"] == ident and c.get("pageRole") == "toc"), None)
        if component is None:
            parser.error(f"{ident} 不是已登记的目录组件")
        from deck_pipeline import reflow_toc
        reflow_toc(LIB / "templates" / f"{ident}.pptx", read_json(args.toc_labels), args.active - 1,
                   read_json(LIB / "theme.json"), args.output, layout=component.get("tocLayout"))
        apply_output_palette(args.output, ident, scheme)
        print(f"已构建自定义目录：{args.output}")
    elif args.component:
        if not args.output or sum(bool(v) for v in (args.variant,args.panel_modes,args.recipe_file))!=1:
            parser.error("组件导出需提供--output，并在--variant、--panel-modes、--recipe-file中选择一项")
        component = next(c for c in catalog["components"] if c["id"] == args.component)
        if args.panel_modes:
            if not component.get("panelChoices"):
                parser.error("此组件没有授权按栏切换图文/小矩形")
            spec={"id":"custom","name":"按内容选择的三栏形式","recipe":{"kind":"panel-formats",
                "panelModes":args.panel_modes.split(","),"connectionMode":args.connection_mode}}
            if args.panel_images:
                images=read_json(args.panel_images)
                if not isinstance(images,dict) or not set(images)<={"1","2","3"}:
                    parser.error("图片映射只能使用字符串栏序号1、2、3")
                spec["recipe"]["images"]={key:str((args.panel_images.parent/Path(value)).resolve()) for key,value in images.items()}
        elif args.recipe_file:
            recipe=read_json(args.recipe_file)
            if recipe.get("images"):
                recipe["images"]={key:str((args.recipe_file.parent/Path(value)).resolve()) for key,value in recipe["images"].items()}
            spec={"id":"custom","name":"按本次内容选择的形式","recipe":recipe}
        else:
            spec=next(v for v in component.get("variants",[]) if v["id"]==args.variant)
            if args.bullet:
                if spec["recipe"]["kind"]!="decompose-list":
                    parser.error("只有多层列表变体可以更换无序符号")
                spec["recipe"]["bullet"]=args.bullet
        result=build_variant(LIB/"templates"/f'{component["id"]}.pptx',component,spec,
            read_json(LIB/"theme.json"),LIB,args.output)
        apply_output_palette(args.output, component["id"], scheme)
        decision={"componentId":component["id"],"variant":spec,"structureLocked":True,"paletteId":args.palette,
            "selectionScope":"用户授权的栏内或标签选择；未进行Word内容自动制作",
            "addedSlots":result["addedSlots"],"replacedSlotKeys":result["replacedSlotKeys"]}
        args.output.with_suffix(".selection.json").write_text(json.dumps(decision,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
        print(f"已导出组件变体及选择记录：{args.output}")
    elif args.viewer_only:
        write_viewer_data()
    else:
        build(args.only)


def apply_output_palette(path, component_id, scheme):
    if scheme["id"] != BASE_PALETTE:
        from deck_native import read_package, save_package

        save_package(recolor_package(read_package(path), component_id, scheme["primary"]), path)


if __name__ == "__main__":
    main()
