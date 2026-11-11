"""从原生目录行复制当前和常规样式，按章节数量等距重排。"""
import copy

from native_pptx import NS, assign_id, set_box, shape_id, shape_inventory, tag, toc_positions, toc_text, validate_toc_labels
from palette_assets import suite_label_color


ROW_ROLES = ("frame", "label", "number", "badge")
ROW_START = 20000


def suite_number(index, style):
    if style == "chinese":
        return "一二三四五六七八九十"[index]
    return f"{index + 1:02d}" if style == "two-digit" else str(index + 1)


def rebuild_suite_toc(root, layout, theme, labels=None, active_index=0, *, source=False):
    tree = root.find("p:cSld/p:spTree", NS)
    nodes = {shape_id(n): n for n in tree.iter(tag("p", "sp"))}
    boxes = {s["id"]: s["boxPx"] for s in shape_inventory(root)}
    if source:
        rows = layout["sourceRows"]
        remove_ids = set(layout["sourceRowRoots"])
        if labels is None:
            labels = ["".join(nodes[row[1]].xpath(".//a:t/text()", namespaces=NS)) for row in rows]
    else:
        rows = [[ROW_START + i * 10 + j for j in range(len(ROW_ROLES))] for i in range(2)]
        remove_ids = {i for i in nodes if ROW_START <= i < ROW_START + 100}
    validate_toc_labels(labels, active_index)
    positions = toc_positions(len(labels), {"toc": layout["spacing"]})
    prototypes = [[(nodes[i], boxes[i]) for i in row] for row in rows[:2]]
    for node in list(tree):
        if shape_id(node) in remove_ids:
            tree.remove(node)

    for index, (label, pos) in enumerate(zip(labels, positions)):
        current = index == active_index
        prototype = prototypes[0 if current else 1]
        top = min(box[1] for _, box in prototype)
        height = max(box[1] + box[3] for _, box in prototype) - top
        scale = pos["height"] / height
        # 底框和色块先绘制，数字与章名位于上层。
        for offset in (0, 3, 1, 2):
            role = ROW_ROLES[offset]
            seed, box = prototype[offset]
            node = copy.deepcopy(seed)
            x, y, w, h = box
            row_y = pos["y"] + (y - top) * scale
            row_height = h * scale
            geom = node.find("p:spPr/a:prstGeom", NS)
            if role == "badge" and geom is not None and geom.get("prst") == "ellipse":
                # 方形中文数字在密集布局下仍需完整落在圆牌内。
                diameter = min(h, pos["height"], layout["numberPt"] * 2)
                if diameter > row_height:
                    row_height = diameter
                    row_y = pos["y"] + (pos["height"] - diameter) / 2
                width = max(w * scale, diameter)
                x, w = x + (w - width) / 2, width
            set_box(node, x, row_y, w, row_height)
            assign_id(node, ROW_START + index * 10 + offset, f"toc-{index + 1:02d}-{role}")
            if role in {"label", "number"}:
                color = suite_label_color(layout, theme, current) if role == "label" else theme["palette"]["white"]
                value = label if role == "label" else suite_number(index, layout["numberStyle"])
                size = layout["labelPt"] if role == "label" else layout["numberPt"]
                toc_text(node, value, size, color, theme["typography"]["family"], "l" if role == "label" else "ctr")
            tree.append(node)
    return positions, labels
