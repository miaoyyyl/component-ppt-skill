"""用本机PowerPoint后台渲染组件；不开编辑窗口，不操作鼠标或键盘。"""
import json
from pathlib import Path
import subprocess
import sys
import hashlib

ROOT = Path(__file__).resolve().parents[1]


def render_worker():
    import pythoncom
    from win32com.client import DispatchEx
    pythoncom.CoInitialize()
    application = None
    presentation = None
    try:
        application = DispatchEx("PowerPoint.Application")
        components = json.loads((ROOT / "library/components.json").read_text(encoding="utf-8"))["components"]
        output = ROOT / "library/previews"
        output.mkdir(exist_ok=True)
        tasks = [(item["id"], ROOT / "library/templates" / f'{item["id"]}.pptx', output / f'{item["id"]}.png') for item in components]
        if "--with-variants" in sys.argv:
            inventory=json.loads((ROOT/"library/inventory.json").read_text(encoding="utf-8"))
            for item in inventory["components"]:
                for variant in item.get("variants",[]):
                    tasks.append((item["id"]+"__"+variant["id"],ROOT/"library"/variant["templatePath"],ROOT/"library"/variant["previewPath"]))
        if "--only" in sys.argv:
            chosen = {value for value in sys.argv[sys.argv.index("--only") + 1:] if not value.startswith("--")}
            tasks = [t for t in tasks if t[0] in chosen or t[0].split("__")[0] in chosen]
            if not tasks:
                raise ValueError("没有匹配的待渲染组件")
        if "--qa-only" in sys.argv:
            qa_output = ROOT / "docs/qa"
            qa_output.mkdir(exist_ok=True)
            tasks = [(f"toc-check-{n:02d}", ROOT / "_qa" / f"toc-check-{n:02d}.pptx", qa_output / f"toc-{n:02d}.png") for n in (2,10)]
        if "--variant-checks" in sys.argv:
            qa_output=ROOT/"docs/qa"
            qa_output.mkdir(exist_ok=True)
            tasks=[("column-image-check",ROOT/"_qa/middle-image.pptx",qa_output/"column-image.png"),
                   ("list-square-check",ROOT/"_qa/list-square-bullet.pptx",qa_output/"list-square.png")]
        if "--four-check" in sys.argv:
            tasks=[("four-mixed-check",ROOT/"_qa/four-mixed.pptx",ROOT/"docs/qa/four-mixed.png")]
        report_path = ROOT / "docs/powerpoint-render.json"
        report = json.loads(report_path.read_text(encoding="utf-8")) if report_path.exists() else {"application":"Microsoft PowerPoint","items":{}}
        report["version"] = str(application.Version)
        for ident, source, image_target in tasks:
            presentation = application.Presentations.Open(str(source), True, False, False)
            if presentation.Slides.Count != 1:
                raise RuntimeError(f"组件不是单页：{ident}")
            presentation.Slides(1).Export(str(image_target), "PNG", 1280, 720)
            roundtrip = ROOT / "_qa" / f"roundtrip-{ident}.pptx"
            roundtrip.parent.mkdir(exist_ok=True)
            if roundtrip.exists():
                roundtrip.unlink()
            presentation.SaveCopyAs(str(roundtrip), 24)
            presentation.Close()
            presentation = application.Presentations.Open(str(roundtrip), True, False, False)
            if presentation.Slides.Count != 1:
                raise RuntimeError(f"保存后重开异常：{ident}")
            presentation.Close()
            presentation = None
            report["items"][ident] = {"input":source.relative_to(ROOT).as_posix(),
                "inputSha256":hashlib.sha256(source.read_bytes()).hexdigest(),
                "preview":image_target.relative_to(ROOT).as_posix(),
                "previewSha256":hashlib.sha256(image_target.read_bytes()).hexdigest(),
                "opened":True,"rendered":True,"savedCopyAndReopened":True}
            print(f"已渲染并保存重开 {ident}", flush=True)
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    finally:
        if presentation is not None:
            presentation.Close()
        # DispatchEx仅创建本次后台实例。出现用户新打开的文档时保留该实例。
        if application is not None and application.Presentations.Count == 0:
            application.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if "--worker" in sys.argv:
        render_worker()
    else:
        try:
            subprocess.run([sys.executable, __file__, "--worker", *sys.argv[1:]], check=True, timeout=240)
        except subprocess.TimeoutExpired:
            raise SystemExit("后台渲染超时；未执行桌面操作，请检查本次后台PowerPoint实例")
