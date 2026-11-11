"""以已验收的母版为基准，构建并实际渲染全部备用配色。"""
import json
import shutil
from pathlib import Path

from deck_native import read_package, save_package
from deck_pipeline import run_powerpoint
from palette_assets import BASE_PALETTE, BLUE, recolor, recolor_package
from production_tasks import ROOT, digest, read_json, write_json

LIB = ROOT / "library"


def make_theme(base, scheme):
    primary = scheme["primary"]
    theme = {
        **base,
        "id": f'{scheme["id"]}-white-human-selected',
        "paletteId": scheme["id"],
        "paletteName": scheme["name"],
        "palette": {key: recolor(value, primary) for key, value in base["palette"].items()},
        "colorPolicy": {
            **base["colorPolicy"],
            "base": "白色画布、所选主色页标题、双层隔断和底条。沿用原组件的色阶关系。",
            "yellow": "保留黄色标签，用于主色底板上的短标签",
        },
        "contentPageFrame": {
            **base["contentPageFrame"],
            "id": "white-title-divider-footer-v1",
            "policy": "标题32pt主色加粗。白色画布、双层隔断和底条。保持原母版的字号、正文层级和图框。",
        },
    }
    return theme


def registered_forms():
    inventory = read_json(LIB / "inventory.json")["components"]
    placeholders = {item["id"]: item for item in read_json(LIB / "placeholders/manifest.json")["items"]}
    for component in inventory:
        ident = component["id"]
        for form in [component, *component.get("variants", [])]:
            key = ident if form is component else f'{ident}__{form["id"]}'
            placeholder = placeholders[key]
            if placeholder["sourceTemplateSha256"] != form["templateSha256"]:
                raise ValueError(f"先更新 X 占位母版：{key}")
            for mode, source in (("case", form), ("placeholder", placeholder)):
                path = LIB / source["templatePath"]
                if digest(path.read_bytes()) != source["templateSha256"]:
                    raise ValueError(f"母版与登记哈希不符：{path}")
                yield ident, key, mode, path, source["templateSha256"]


def build():
    catalog = read_json(LIB / "components.json")
    base = read_json(LIB / "theme.json")
    if base["palette"]["primary"] != BLUE:
        raise ValueError("主色映射须与蓝白基准母版一致")
    sources = list(registered_forms())
    stage = ROOT / "_qa/palette-build"
    stage.mkdir(parents=True, exist_ok=True)
    manifest_path = LIB / "palettes/manifest.json"
    previous = read_json(manifest_path)["schemes"] if manifest_path.exists() else {}
    manifest = {"schemaVersion": 1, "libraryVersion": catalog["version"], "schemes": {}}
    data, publish = {}, []
    for scheme in catalog["palettes"]:
        ident = scheme["id"]
        if ident == BASE_PALETTE:
            continue
        theme = make_theme(base, scheme)
        theme_path = f"palettes/{ident}/theme.json"
        write_json(stage / theme_path, theme)
        publish.append(theme_path)
        old_scheme = previous.get(ident, {})
        forms, pending = {}, []
        for component_id, key, mode, source, source_hash in sources:
            suffix = "-X" if mode == "placeholder" else ""
            template = f"palettes/{ident}/{mode}/templates/{key}--{ident}{suffix}.pptx"
            preview = f"palettes/{ident}/{mode}/previews/{key}.png"
            package = recolor_package(read_package(source), component_id, scheme["primary"])
            save_package(package, stage / template)
            entry = {
                "templatePath": template,
                "templateSha256": digest((stage / template).read_bytes()),
                "sourceTemplateSha256": source_hash,
                "previewPath": preview,
            }
            forms.setdefault(key, {})[mode] = entry
            old = old_scheme.get("forms", {}).get(key, {}).get(mode)
            if (old and all(old.get(k) == v for k, v in entry.items())
                    and (LIB / template).is_file() and digest((LIB / template).read_bytes()) == entry["templateSha256"]
                    and (LIB / preview).is_file() and digest((LIB / preview).read_bytes()) == old.get("previewSha256")):
                entry["previewSha256"] = old["previewSha256"]
                continue
            publish.extend((template, preview))
            pending.append((entry, {
                "id": f"{key}--{mode}",
                "input": str(stage / template),
                "image": str(stage / preview),
                "output": str(stage / "roundtrip" / ident / f"{key}--{mode}"),
            }))
        renderer = old_scheme.get("renderer")
        for start in range(0, len(pending), 20):
            batch = pending[start:start + 20]
            report = run_powerpoint(stage, {"operation": "placeholder-batch", "items": [item for _, item in batch]},
                                    f"palette-{ident}-{start // 20 + 1:02d}")
            renderer = report["application"]
            for entry, item in batch:
                record = report["items"][item["id"]]
                if record["inputSha256"] != entry["templateSha256"] or not record["savedCopyAndReopened"]:
                    raise ValueError(f'PowerPoint 检查失败：{item["id"]}')
                entry["previewSha256"] = record["previewSha256"]
            print(f'{scheme["name"]}：PowerPoint 已渲染并重开 {start + len(batch)}/{len(pending)} 页', flush=True)
        manifest["schemes"][ident] = {"themePath": theme_path, "forms": forms, "renderer": renderer}
        data[ident] = {"theme": theme, "forms": forms}

    for relative in publish:
        target = LIB / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(stage / relative, target)
    write_json(manifest_path, manifest)
    script = "window.COMPONENT_LIBRARY.paletteAssets = " + json.dumps(data, ensure_ascii=False) + ";\n"
    (LIB / "palettes/data.js").write_text(script, encoding="utf-8")
    from build_library import write_viewer_data
    write_viewer_data()


if __name__ == "__main__":
    build()
