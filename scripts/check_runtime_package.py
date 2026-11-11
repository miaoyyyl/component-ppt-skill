"""校验运行包ZIP或解压目录的完整性与边界，只需Python标准库。"""
from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path, PurePosixPath


def check(package):
    package = Path(package).resolve()
    archive = None
    if package.is_file():
        archive = zipfile.ZipFile(package)
        names = [n.filename for n in archive.infolist() if not n.is_dir()]
        if len(names) != len(set(names)) or any("\\" in n or PurePosixPath(n).is_absolute() or ".." in PurePosixPath(n).parts for n in names):
            raise ValueError("ZIP文件重名或存在越界路径")
        prefixes = {n.split("/", 1)[0] for n in names}
        if len(prefixes) != 1 or any("/" not in n for n in names):
            raise ValueError("ZIP应只含一个运行包根目录")
        prefix = next(iter(prefixes)) + "/"
        actual = {n[len(prefix):] for n in names}

        def read(name):
            return archive.read(prefix + name)
    else:
        if not package.is_dir():
            raise ValueError("ZIP或解压目录不存在")
        actual = {p.relative_to(package).as_posix() for p in package.rglob("*") if p.is_file()}

        def read(name):
            path = (package / name).resolve()
            if not path.is_relative_to(package):
                raise ValueError("文件越出运行包目录")
            return path.read_bytes()
    try:
        manifest = json.loads(read("MANIFEST.json"))
        if manifest.get("schema") != "ppt-skill-runtime-package" or manifest.get("schemaVersion") != 1:
            raise ValueError("清单格式不正确")
        expected = {f["path"] for f in manifest["files"]}
        if len(expected) != len(manifest["files"]):
            raise ValueError("清单路径重复")
        if not expected | {"MANIFEST.json"} <= actual:
            raise ValueError("包内文件缺失")
        extras = actual - expected - {"MANIFEST.json"}
        if extras:
            raise ValueError("目录包含清单外文件，请在未运行任务的干净解压目录校验：" + str(sorted(extras)[:3]))
        total = 0
        for entry in manifest["files"]:
            name = entry["path"]
            parts = PurePosixPath(name).parts
            if PurePosixPath(name).is_absolute() or ".." in parts or ":" in name or "\\" in name:
                raise ValueError("清单路径越界")
            if any(p in {"outputs", "output", ".git", ".runtime", "_qa", "by_human", "materials", "__pycache__", "node_modules"} for p in parts):
                raise ValueError("意外收录主题资料、输出或缓存：" + name)
            if parts[-1].startswith((".env", "~$")):
                raise ValueError("意外收录配置或Office锁文件")
            payload = read(name)
            if len(payload) != entry["bytes"] or hashlib.sha256(payload).hexdigest() != entry["sha256"]:
                raise ValueError("文件大小或哈希不匹配：" + name)
            total += len(payload)
        for skill in manifest["researchSkills"]:
            if "skills/" + skill + "/SKILL.md" not in expected:
                raise ValueError("科研skill入口缺失")
        for name in ("README.md", "requirements-ppt.txt", "requirements-research.lock.txt", "project/library/index.html",
                     "project/scripts/research_bundle.py", "project/docs/PDF文献驱动的Agent制作流程.md",
                     "optional-skills/academic-blue-ppt-design/SKILL.md"):
            if name not in expected:
                raise ValueError("入口缺失：" + name)
        catalog = json.loads(read("project/library/components.json"))
        inventory = json.loads(read("project/library/inventory.json"))
        placeholders = json.loads(read("project/library/placeholders/manifest.json"))
        if catalog["version"] != manifest["libraryVersion"] or len(catalog["components"]) != manifest["components"]:
            raise ValueError("组件版本或数量不一致")
        if sum(len(c.get("variants", [])) for c in catalog["components"]) != manifest["variants"]:
            raise ValueError("变体数量不一致")
        forms = 0
        for component in inventory["components"]:
            for form in [component, *component.get("variants", [])]:
                target = "project/library/" + form["templatePath"]
                if hashlib.sha256(read(target)).hexdigest() != form["templateSha256"]:
                    raise ValueError("原生母版哈希不符：" + target)
                forms += 1
        if len(placeholders["items"]) != forms:
            raise ValueError("X占位形式数量不一致")
        for form in placeholders["items"]:
            target = "project/library/" + form["templatePath"]
            if hashlib.sha256(read(target)).hexdigest() != form["templateSha256"]:
                raise ValueError("X占位母版哈希不符")
            read("project/library/" + form["previewPath"])
        for component in catalog["components"]:
            if component.get("referenceFile"):
                read("project/library/references/" + component["referenceFile"])
        return {"passed": True, "checkedFiles": len(manifest["files"]), "uncompressedBytes": total,
                "libraryVersion": catalog["version"], "components": manifest["components"], "variants": manifest["variants"],
                "nativeForms": forms, "placeholderForms": len(placeholders["items"]),
                "researchSkills": len(manifest["researchSkills"]), "topicOutputsIncluded": False,
                "scientificOrPowerPointReview": "not-performed-by-this-integrity-check"}
    finally:
        if archive:
            archive.close()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("package", type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(check(args.package), ensure_ascii=False, indent=2))
    except (ValueError, KeyError, OSError, zipfile.BadZipFile) as error:
        parser.exit(1, "运行包检查未通过：" + str(error) + "\n")


if __name__ == "__main__":
    main()
