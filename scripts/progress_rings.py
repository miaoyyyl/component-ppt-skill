"""同步研究进展组件的环形图数据与图表内嵌工作簿。"""
import re
from datetime import datetime
from decimal import Decimal
from io import BytesIO

from xlsxwriter import Workbook

from native_pptx import NS, rels_path, resolve_target, serialize, tag, xml


CHART_NS = {**NS, "c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}


def related_part(data, source, rid):
    relation = next(r for r in xml(data[rels_path(source)]) if r.get("Id") == rid)
    return resolve_target(source, relation.get("Target"))


def set_ring_percentage(data, slide_part, shape, value):
    match = re.fullmatch(r"([0-9]{1,3}(?:\.[0-9])?)[%％]", value.strip())
    if match is None or Decimal(match[1]) > 100:
        raise ValueError("完成进度须为 0% 至 100% 的单个百分比，最多保留一位小数")
    completed = Decimal(match[1])
    amounts = [completed, 100 - completed]
    labels = ["已完成", "未完成"]
    series_name = "完成进度"

    chart_ref = shape.find(".//c:chart", CHART_NS)
    chart_part = related_part(data, slide_part, chart_ref.get(tag("r", "id")))
    chart = xml(data[chart_part])
    series = chart.find(".//c:pieChart/c:ser", CHART_NS)
    for path, ref, items in [
        ("c:tx/c:strRef", "Sheet1!$B$1", [series_name]),
        ("c:cat/c:strRef", "Sheet1!$A$2:$A$3", labels),
        ("c:val/c:numRef", "Sheet1!$B$2:$B$3", [format(n.normalize(), "f") for n in amounts]),
    ]:
        source = series.find(path, CHART_NS)
        source.find("c:f", CHART_NS).text = ref
        points = source.xpath("*/c:pt/c:v", namespaces=CHART_NS)
        for node, item in zip(points, items, strict=True):
            node.text = item
    data[chart_part] = serialize(chart)

    embedded = chart.find("c:externalData", CHART_NS)
    workbook_part = related_part(data, chart_part, embedded.get(tag("r", "id")))
    buffer = BytesIO()
    with Workbook(buffer, {"in_memory": True}) as workbook:
        workbook.set_properties({"created": datetime(2026, 11, 1)})
        sheet = workbook.add_worksheet("Sheet1")
        sheet.write_row(0, 0, ["状态", series_name])
        for row, (label, amount) in enumerate(zip(labels, amounts), 1):
            sheet.write_row(row, 0, [label, float(amount)])
    data[workbook_part] = buffer.getvalue()
    return f"{completed.normalize():f}%"
