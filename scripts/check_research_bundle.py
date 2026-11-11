"""无网络的PDF任务验收：验证输入、原页裁切、来源/组件约束，不冒充科学精读。"""
from __future__ import annotations

import copy
import io
import shutil
import tempfile
import unittest
from pathlib import Path

from PIL import Image

from research_bundle import (ROOT, check_bundle, crop_figure, digest, initialize, normalize_task,
                             pymupdf, read_json, write_json)
from production_tasks import CONSTRAINTS, WORKFLOW


class ResearchBundleChecks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        (ROOT / ".runtime").mkdir(exist_ok=True)
        cls.temp = tempfile.TemporaryDirectory(dir=ROOT / ".runtime", prefix="pdf-research-check-")
        cls.root = Path(cls.temp.name)
        fitz = pymupdf()
        image = Image.new("RGBA", (100, 50), (255, 0, 0, 160))
        buffer = io.BytesIO()
        image.save(buffer, format="PNG")
        doc = fitz.open()
        page = doc.new_page(width=400, height=600)
        page.insert_text((25, 35), "TEST FIXTURE. Not scientific evidence. Figure 1 shows a vector box.")
        page.draw_rect(fitz.Rect(30, 100, 200, 220), color=(0, 0, 1), fill=(0.8, 0.9, 1))
        page.insert_image(fitz.Rect(50, 250, 250, 350), stream=buffer.getvalue())
        page = doc.new_page(width=400, height=600)
        page.insert_image(fitz.Rect(50, 100, 250, 200), stream=buffer.getvalue())
        page = doc.new_page(width=400, height=600)
        page.draw_rect(fitz.Rect(10, 10, 110, 60), color=(1, 0, 0), fill=(1, 0, 0))
        page.set_rotation(90)
        cls.pdf = cls.root / "paper.pdf"
        doc.save(cls.pdf)
        doc.close()
        cls.raw = {"schema": "ppt-production-task", "schemaVersion": 1, "id": "pdf-input-test-001",
                   "libraryVersion": read_json(ROOT / "library/components.json")["version"],
                   "createdAt": "2026-11-02T00:00:00Z", "updatedAt": "2026-11-02T00:00:00Z", "savedAt": None,
                   "brief": {"title": "功能验证", "audience": "验收", "durationMinutes": 5,
                             "pageRange": "2-3", "mustKeep": "测试重点", "sourceNotes": "无Word", "researchScope": "不联网"},
                   "preferredComponents": [{"componentId": "BW-FOUR-TEXT-01", "variantId": "base", "priority": 1}],
                   "documents": [], "workflow": WORKFLOW, "constraints": CONSTRAINTS,
                   "fallback": {"mode": "unselected-library-components", "variantPolicy": "model-auto", "componentIds": []},
                   "handoff": {"components": [{"id": "fake", "forms": []}]}}
        cls.selection = cls.root / "selection.json"
        write_json(cls.selection, cls.raw)
        cls.base = cls.root / "baseline"
        initialize(cls.selection, [cls.pdf, cls.pdf], "只用于功能验证，不形成科研结论", cls.base, dpi=72)

    @classmethod
    def tearDownClass(cls):
        cls.temp.cleanup()

    def setUp(self):
        self.directory = self.root / self._testMethodName
        shutil.copytree(self.base, self.directory)

    def plan(self):
        sources = read_json(self.directory / "sources.json")
        sources["sources"][0]["readScope"] = "selected-sections"
        write_json(self.directory / "sources.json", sources)
        write_json(self.directory / "focus-plan.json", {"questions": ["验证来源对应"], "storyline": ["验证输入到方案的关联"]})
        write_json(self.directory / "evidence.json", {"claims": [{"id": "C001", "statement": "这是功能测试材料",
                   "sourceRefs": ["P001:p0001:b001"], "type": "observation", "status": "reviewed"}]})
        component = next(c for c in read_json(self.directory / "component-options.json") if c["id"] == "BW-FOUR-TEXT-01")
        page = {"pageNumber": 1, "title": "输入与证据关联验证", "purpose": "验证记录", "section": "测试",
                "componentId": component["id"], "variantId": "base", "selectionReason": "人工优先且逻辑适配",
                "claimIds": ["C001"], "slots": [{"key": s["key"], "text": "测试", "claimIds": ["C001"]}
                    for s in component["forms"][0]["slots"] if s["kind"] != "state"]}
        write_json(self.directory / "page-plan.json", {"pages": [page]})
        return page

    def test_word_is_optional_and_legacy_service_unchanged(self):
        self.assertTrue(check_bundle(self.directory)["passed"])
        self.assertFalse(read_json(self.directory / "research-task.json")["wordRequired"])
        with self.assertRaises(ValueError):
            normalize_task(self.raw, read_json(ROOT / "library/components.json"))

    def test_selected_palette_reaches_task_theme_and_templates(self):
        raw = {**self.raw, "paletteId": "purple"}
        selection = self.root / "purple-selection.json"
        write_json(selection, raw)
        output = self.root / "purple-task"
        initialize(selection, [self.pdf], "配色交接验证", output, dpi=72, max_pages=1)
        self.assertEqual(read_json(output / "selection.json")["paletteId"], "purple")
        self.assertEqual(read_json(output / "theme.json")["palette"]["primary"], "640872")
        options = read_json(output / "component-options.json")
        self.assertTrue(all(form["template"]["path"].startswith("library/palettes/purple/placeholder/")
                            for component in options for form in component["forms"]))
        self.assertTrue(check_bundle(output)["passed"])

    def test_originals_duplicates_and_unread_state(self):
        sources = read_json(self.directory / "sources.json")["sources"]
        self.assertEqual(len(sources), 1)
        self.assertEqual(len(sources[0]["aliases"]), 2)
        self.assertEqual(sources[0]["readScope"], "extracted-not-read")
        self.assertEqual(digest(self.pdf.read_bytes()), digest((self.directory / "sources/P001.pdf").read_bytes()))
        self.assertEqual(sources[0]["pageCount"], 3)
        self.assertFalse(sources[0]["partialExtraction"])

    def test_text_images_mask_and_scan_are_distinct(self):
        pages = read_json(self.directory / "extracted/P001/pages.json")["pages"]
        self.assertIn("Not scientific evidence", pages[0]["blocks"][0]["text"])
        self.assertEqual(pages[1]["textStatus"], "needs-visual-or-ocr")
        figures = read_json(self.directory / "figures.json")["figures"]
        self.assertTrue(figures and figures[0]["maskFile"])
        self.assertTrue(all(f["review"]["status"] == "pending" for f in figures))
        self.assertNotEqual(len(figures), len(pages))

    def test_crops_vector_and_rotated_page(self):
        vector = crop_figure(self.directory, "P001", 1, [30, 100, 200, 220], "vector", 72)
        with Image.open(self.directory / vector["file"]["path"]) as img:
            self.assertEqual(img.size, (170, 120))
            self.assertGreater(img.getpixel((80, 60))[2], 220)
        # 原页左上红框在90度预览中变为右上，crop须得到同一图形
        rotated = crop_figure(self.directory, "P001", 3, [540, 10, 590, 110], "rotated", 72)
        with Image.open(self.directory / rotated["file"]["path"]) as img:
            self.assertEqual(img.size, (50, 100))
            self.assertLess(img.getpixel((25, 50))[1], 10)
        self.assertEqual(rotated["sourceSha256"], digest(self.pdf.read_bytes()))
        self.assertTrue(check_bundle(self.directory)["passed"])

    def test_invalid_crops_and_collision_do_not_overwrite(self):
        for source, number, bbox, ident in [("unknown", 1, [1, 1, 20, 20], "bad"), ("P001", 0, [1, 1, 20, 20], "bad"),
                                             ("P001", 1, [-1, 1, 20, 20], "bad"), ("P001", 1, [1, 1, 1, 20], "bad"),
                                             ("P001", 1, [1, 1, 20, 20], "../escape")]:
            with self.assertRaises(ValueError):
                crop_figure(self.directory, source, number, bbox, ident)
        crop_figure(self.directory, "P001", 1, [1, 1, 20, 20], "one")
        with self.assertRaises(ValueError):
            crop_figure(self.directory, "P001", 1, [1, 1, 20, 20], "one")

    def test_unchosen_fallback_and_untrusted_handoff(self):
        options = read_json(self.directory / "component-options.json")
        self.assertNotIn("fake", {c["id"] for c in options})
        self.assertEqual(next(c for c in options if c["id"] == "BW-FOUR-TEXT-01")["selectedVariant"], "base")
        self.assertEqual(sum(c["selectionTier"] == "preferred" for c in options), 1)
        empty = copy.deepcopy(self.raw)
        empty["preferredComponents"] = []
        write_json(self.root / "empty-selection.json", empty)
        out = self.root / "all-fallback"
        initialize(self.root / "empty-selection.json", [self.pdf], "全部作为备用", out, dpi=72, max_pages=1)
        self.assertTrue(all(c["selectionTier"] == "fallback" for c in read_json(out / "component-options.json")))
        source = read_json(out / "sources.json")["sources"][0]
        self.assertTrue(source["partialExtraction"])
        self.assertEqual(source["unreadPages"], [2, 3])

    def test_rejects_existing_directory_and_blank_direction(self):
        with self.assertRaises(ValueError):
            initialize(self.selection, [self.pdf], "重点", self.directory)
        with self.assertRaises(ValueError):
            initialize(self.selection, [self.pdf], " ", self.root / "blank")

    def test_tampering_source_and_text_index_is_rejected(self):
        for relative in ["sources/P001.pdf", "extracted/P001/pages.json"]:
            path = self.directory / relative
            payload = path.read_bytes()
            path.write_bytes(payload + b"changed")
            with self.assertRaises(ValueError):
                check_bundle(self.directory)
            path.write_bytes(payload)

    def test_plan_requires_actual_agent_work(self):
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        self.plan()
        result = check_bundle(self.directory, "plan")
        self.assertTrue(result["passed"])
        self.assertEqual(result["powerpointVisualReview"], "not-checked")

    def test_unknown_location_unread_source_and_missing_slot_rejected(self):
        page = self.plan()
        evidence = read_json(self.directory / "evidence.json")
        evidence["claims"][0]["sourceRefs"] = ["P001:p9999"]
        write_json(self.directory / "evidence.json", evidence)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        self.plan()
        sources = read_json(self.directory / "sources.json")
        sources["sources"][0]["readScope"] = "extracted-not-read"
        write_json(self.directory / "sources.json", sources)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        self.plan()
        page["slots"].pop()
        write_json(self.directory / "page-plan.json", {"pages": [page]})
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")

    def test_fixed_variant_and_unknown_claim_rejected(self):
        page = self.plan()
        page["variantId"] = "numbered"
        write_json(self.directory / "page-plan.json", {"pages": [page]})
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        page = self.plan()
        page["claimIds"] = ["invented"]
        write_json(self.directory / "page-plan.json", {"pages": [page]})
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")

    def image_plan(self, figure_id):
        self.plan()
        component = next(c for c in read_json(self.directory / "component-options.json") if c["id"] == "BW-COVER-01")
        page = {"pageNumber": 1, "title": "功能验证", "purpose": "验证图像约束", "section": "测试",
                "componentId": component["id"], "variantId": "base", "selectionReason": "备用组件需要图像槽位",
                "claimIds": ["C001"], "slots": []}
        for slot in component["forms"][0]["slots"]:
            if slot["kind"] == "state":
                continue
            binding = {"key": slot["key"], "claimIds": []}
            binding.update({"figureIds": [figure_id]} if slot["kind"].startswith("image") else {"text": "测试"})
            page["slots"].append(binding)
        write_json(self.directory / "page-plan.json", {"pages": [page]})

    def test_image_needs_review_and_unknown_figures_rejected(self):
        figures = read_json(self.directory / "figures.json")
        ident = figures["figures"][0]["id"]
        self.image_plan(ident)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        figures["figures"][0]["review"] = {"status": "reviewed", "decision": "use", "reason": "仅测试图像绑定"}
        write_json(self.directory / "figures.json", figures)
        self.assertTrue(check_bundle(self.directory, "plan")["passed"])
        self.image_plan("invented-image")
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")

    def test_ai_is_concept_with_record_and_existing_style_reference(self):
        figures = read_json(self.directory / "figures.json")
        reference = figures["figures"][0]
        ai = copy.deepcopy(reference)
        ai.update({"id": "G001", "origin": "ai-generated", "role": "quantitative-result",
                   "review": {"status": "reviewed", "decision": "use", "reason": "仅测试生成记录约束"},
                   "generationRecord": {"tool": "test-fixture", "prompt": "FUNCTIONAL TEST ONLY", "outputSha256": ai["file"]["sha256"]},
                   "styleReferenceIds": [reference["id"]]})
        figures["figures"].append(ai)
        self.image_plan("G001")
        write_json(self.directory / "figures.json", figures)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        ai["role"] = "concept"
        write_json(self.directory / "figures.json", figures)
        self.assertTrue(check_bundle(self.directory, "plan")["passed"])
        ai["styleReferenceIds"] = ["unavailable-reference"]
        write_json(self.directory / "figures.json", figures)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")

    def test_external_claim_requires_source_locator_and_read_scope(self):
        self.plan()
        sources = read_json(self.directory / "sources.json")
        external = {"id": "E001", "kind": "literature", "url": "https://example.org/test-fixture",
                    "accessedAt": "2026-11-02", "readScope": "metadata", "locators": ["E001:results"]}
        sources["sources"].append(external)
        write_json(self.directory / "sources.json", sources)
        evidence = read_json(self.directory / "evidence.json")
        evidence["claims"][0]["sourceRefs"] = ["E001:results"]
        write_json(self.directory / "evidence.json", evidence)
        with self.assertRaises(ValueError):
            check_bundle(self.directory, "plan")
        external["readScope"] = "full-text"
        write_json(self.directory / "sources.json", sources)
        self.assertTrue(check_bundle(self.directory, "plan")["passed"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
