"""打包运行资料，白名单收录项目和skill，不收录主题输出、原稿、缓存和凭据。"""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import zipfile
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RESEARCH_SKILLS = ("research-literature", "research-reading", "research-writing", "research-figures", "research-design")
OMIT = {"__pycache__", ".git", ".runtime", "_qa", "node_modules", ".venv", "venv", "outputs", "output"}


def sha(payload):
    return hashlib.sha256(payload).hexdigest()


def enumerate_files(root):
    root = root.resolve()
    if not root.is_dir():
        raise ValueError("需要的运行目录不存在：" + str(root))
    for path in sorted(root.rglob("*")):
        relative = path.relative_to(root)
        if any(p in OMIT for p in relative.parts) or path.suffix.lower() in {".pyc", ".pyo", ".log", ".tmp", ".temp"} or path.name.startswith(("~$", ".env")):
            continue
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError("运行目录存在指向目录外的链接：" + str(path))
        if path.is_file():
            yield path, relative


def build(output, skills_root, optional_ppt):
    output = output.resolve()
    if not output.is_relative_to(ROOT) or output.suffix.lower() != ".zip":
        raise ValueError("输出必须为当前项目内的ZIP文件")
    if output.exists() or output.with_suffix(".zip.part").exists():
        raise ValueError("输出已存在，请选择新文件名，避免覆盖历史包")
    entries = {}

    def add(path, target, category):
        target = Path(target).as_posix()
        if target in entries:
            raise ValueError("包内文件重名：" + target)
        if not path.is_file():
            raise ValueError("缺少运行文件：" + str(path))
        entries[target] = (path, category)

    # 当前运行项目；materials/by_human/outputs从未进入候选集合
    for directory in ("library", "scripts", "templates", "demos", "packaging"):
        for path, relative in enumerate_files(ROOT / directory):
            add(path, "project/" + directory + "/" + relative.as_posix(), "project")
    for name in ("AGENTS.md", "README.md", "requirements.txt", ".gitignore", ".gitattributes", "启动制作任务.ps1"):
        add(ROOT / name, "project/" + name, "project")
    # 说明保留；排除检查图片、生成样例、JSON报告和安装历史补丁
    for path in sorted((ROOT / "docs").glob("*.md")):
        add(path, "project/docs/" + path.name, "project-docs")
    for skill in RESEARCH_SKILLS:
        for path, relative in enumerate_files(skills_root / skill):
            add(path, "skills/" + skill + "/" + relative.as_posix(), "research-skill")
    for path, relative in enumerate_files(optional_ppt):
        add(path, "optional-skills/academic-blue-ppt-design/" + relative.as_posix(), "optional-skill")
    for name in ("README.md", "requirements-ppt.txt", "requirements-research.lock.txt"):
        add(ROOT / "packaging" / name, name, "package-support")

    catalog = json.loads((ROOT / "library/components.json").read_text(encoding="utf-8"))
    version = subprocess.run(["git", "rev-parse", "HEAD"], cwd=ROOT, capture_output=True, text=True, check=True).stdout.strip()
    timestamp = datetime.now(timezone(timedelta(hours=8)))
    prefix = "ppt-research-runtime-" + timestamp.strftime("%Y%m%d")
    manifest = {"schema": "ppt-skill-runtime-package", "schemaVersion": 1, "createdAt": timestamp.isoformat(),
                "projectGitRevisionBeforePackaging": version, "workingFilesCapturedBySha256": True,
                "libraryVersion": catalog["version"], "components": len(catalog["components"]),
                "variants": sum(len(c.get("variants", [])) for c in catalog["components"]),
                "researchSkills": list(RESEARCH_SKILLS), "optionalSkills": ["academic-blue-ppt-design"],
                "sourceRebuildAvailable": False, "executablesAndHostToolsIncluded": False,
                "excluded": ["outputs/", "by_human/", "materials/", ".runtime/", "_qa/", ".git/", "docs/qa/ generated results", "docs/*.json historical reports", "virtualenv/node_modules/credentials"],
                "manifestSelfHashIncluded": False, "files": []}
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(".zip.part")
    with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for target, (path, category) in sorted(entries.items()):
            payload = path.read_bytes()
            archive.writestr(prefix + "/" + target, payload)
            manifest["files"].append({"path": target, "bytes": len(payload), "sha256": sha(payload), "category": category})
        # 浏览器检查写预览时需要目录；不携带旧检查截图
        target = "project/docs/qa/README.md"
        payload = "本目录用于目标机重新执行浏览器检查后保存截图；运行包不携带历史测试成品\n".encode("utf-8")
        archive.writestr(prefix + "/" + target, payload)
        manifest["files"].append({"path": target, "bytes": len(payload), "sha256": sha(payload), "category": "package-support"})
        archive.writestr(prefix + "/MANIFEST.json", json.dumps(manifest, ensure_ascii=False, indent=2) + "\n")
    temporary.rename(output)
    checksum = sha(output.read_bytes())
    output.with_suffix(".zip.sha256").write_text(checksum + "  " + output.name + "\n", encoding="utf-8")
    return {"archive": str(output), "bytes": output.stat().st_size, "sha256": checksum,
            "files": len(manifest["files"]) + 1, "components": manifest["components"], "variants": manifest["variants"]}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--skills-root", type=Path, default=Path.home() / ".agents/skills")
    parser.add_argument("--optional-ppt-skill", type=Path, default=Path.home() / ".codex/skills/academic-blue-ppt-design")
    args = parser.parse_args()
    try:
        result = build(args.output, args.skills_root, args.optional_ppt_skill)
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        parser.exit(1, "打包未完成：" + str(error) + "\n")
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
