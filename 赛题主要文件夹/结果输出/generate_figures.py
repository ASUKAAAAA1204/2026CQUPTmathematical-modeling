# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
import math
import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from xml.sax.saxutils import escape


BASE = Path(__file__).resolve().parent.parent
OUT = BASE / "B题_solution_outputs"
FIG = OUT / "figures"
FIG.mkdir(exist_ok=True)


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col:
        if ch.isalpha():
            n = n * 26 + ord(ch.upper()) - 64
    return n - 1


def read_experiment(csv_path: Path):
    rows = []
    types_order = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        next(reader)
        for r in reader:
            sample_type = r[0]
            if sample_type not in types_order:
                types_order.append(sample_type)
            rows.append(
                {
                    "type": sample_type,
                    "id": int(r[1]),
                    "day": int(r[2]),
                    "mag": float(r[3]),
                    "temp": float(r[4]),
                    "hum": float(r[5]),
                }
            )
    return rows, types_order


def read_xlsx_weather(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = workbook.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        first_sheet = sheets[0]
        sheet_id = first_sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        target = relmap[sheet_id].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target

        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
            for si in shared_xml:
                shared.append("".join(t.text or "" for t in si.iter(ns_t)))

        rows = []
        worksheet = ET.fromstring(zf.read(target))
        ns_c = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
        ns_v = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v"
        ns_is = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is"
        ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
        for row in worksheet.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
            values = []
            for cell in row.iter(ns_c):
                ref = cell.attrib.get("r", "")
                idx = col_to_idx("".join(filter(str.isalpha, ref))) if ref else len(values)
                cell_type = cell.attrib.get("t")
                v = cell.find(ns_v)
                inline = cell.find(ns_is)
                value = ""
                if v is not None:
                    raw = v.text or ""
                    value = shared[int(raw)] if cell_type == "s" else raw
                elif inline is not None:
                    value = "".join(t.text or "" for t in inline.iter(ns_t))
                while len(values) <= idx:
                    values.append("")
                values[idx] = value
            if values and values[0].isdigit():
                rows.append(
                    {
                        "seq": int(values[0]),
                        "year": int(values[1]),
                        "month": int(values[2]),
                        "date": int(values[3]),
                        "weather": values[4],
                        "temp": float(values[5]),
                        "hum": float(values[6]),
                    }
                )
        return rows


def read_xlsx_table(xlsx_path: Path):
    with zipfile.ZipFile(xlsx_path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = workbook.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheets")
        first_sheet = sheets[0]
        sheet_id = first_sheet.attrib[
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        ]
        target = relmap[sheet_id].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target

        shared = []
        if "xl/sharedStrings.xml" in zf.namelist():
            shared_xml = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
            for si in shared_xml:
                shared.append("".join(t.text or "" for t in si.iter(ns_t)))

        rows = []
        worksheet = ET.fromstring(zf.read(target))
        ns_c = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"
        ns_v = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v"
        ns_is = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}is"
        ns_t = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
        for row in worksheet.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row"):
            values = []
            for cell in row.iter(ns_c):
                ref = cell.attrib.get("r", "")
                idx = col_to_idx("".join(filter(str.isalpha, ref))) if ref else len(values)
                cell_type = cell.attrib.get("t")
                v = cell.find(ns_v)
                inline = cell.find(ns_is)
                value = ""
                if v is not None:
                    raw = v.text or ""
                    value = shared[int(raw)] if cell_type == "s" else raw
                elif inline is not None:
                    value = "".join(t.text or "" for t in inline.iter(ns_t))
                while len(values) <= idx:
                    values.append("")
                values[idx] = value
            if any(v != "" for v in values):
                rows.append(values)
        return rows


def read_csv_rows(path: Path):
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def color_for_type(sample_type: str) -> str:
    return {
        "小号铁钉": "#1f77b4",
        "小号铁夹": "#ff7f0e",
        "普通钢筋": "#2ca02c",
        "锈蚀钢筋": "#d62728",
    }.get(sample_type, "#333333")


def svg_wrap(width, height, body):
    return f'''<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">
<rect width="100%" height="100%" fill="white"/>
{body}
</svg>'''


def txt(x, y, s, size=12, anchor="start", fill="#111", weight="normal"):
    return f'<text x="{x}" y="{y}" font-size="{size}" text-anchor="{anchor}" fill="{fill}" font-family="SimSun, Microsoft YaHei, Arial" font-weight="{weight}">{escape(str(s))}</text>'


def line(x1, y1, x2, y2, stroke="#333", width=1, dash=None):
    dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" stroke-width="{width}"{dash_attr}/>'


def circle(cx, cy, r, fill, stroke="none", opacity=1.0):
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{fill}" stroke="{stroke}" opacity="{opacity}"/>'


def rect(x, y, w, h, fill="none", stroke="#333", width=1, rx=0, opacity=1.0):
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{width}" opacity="{opacity}"/>'


def polyline(points, stroke, width=2, fill="none", opacity=1.0, dash=None):
    dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
    pts = " ".join(f"{x},{y}" for x, y in points)
    return f'<polyline points="{pts}" fill="{fill}" stroke="{stroke}" stroke-width="{width}" opacity="{opacity}"{dash_attr}/>'


def to_xy(value, vmin, vmax, left, top, width, height):
    if vmax == vmin:
        return left, top + height
    x = left + (value[0] - vmin[0]) / (vmax[0] - vmin[0]) * width
    y = top + height - (value[1] - vmin[1]) / (vmax[1] - vmin[1]) * height
    return x, y


def draw_axes(x, y, w, h, x_ticks=None, y_ticks=None, x_label="", y_label="", title=""):
    parts = [rect(x, y, w, h, fill="white", stroke="#333", width=1)]
    parts.append(line(x, y + h, x + w, y + h, width=1.2))
    parts.append(line(x, y, x, y + h, width=1.2))
    if title:
        parts.append(txt(x + w / 2, y - 12, title, size=16, anchor="middle", weight="bold"))
    if x_ticks:
        for xt, label in x_ticks:
            parts.append(line(xt, y + h, xt, y + h + 5))
            parts.append(txt(xt, y + h + 20, label, size=10, anchor="middle"))
    if y_ticks:
        for yt, label in y_ticks:
            parts.append(line(x - 5, yt, x, yt))
            parts.append(txt(x - 8, yt + 4, label, size=10, anchor="end"))
    if x_label:
        parts.append(txt(x + w / 2, y + h + 40, x_label, size=12, anchor="middle"))
    if y_label:
        parts.append(txt(x - 50, y + h / 2, y_label, size=12, anchor="middle"))
    return parts


def save_svg(name, body, width=1200, height=800):
    path = FIG / name
    path.write_text(svg_wrap(width, height, body), encoding="utf-8")
    return path


def make_fig1_raw_decay(obs, types_order):
    W, H = 1200, 800
    left, top, width, height = 80, 60, 980, 620
    parts = [txt(W / 2, 28, "四类样品原始剩磁衰减曲线", size=20, anchor="middle", weight="bold")]
    days = [r["day"] for r in obs]
    mags = [r["mag"] for r in obs]
    xmin, xmax = 0, 90
    ymin, ymax = 0, max(mags) * 1.05
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * d / 90, str(d)) for d in [0, 15, 30, 45, 60, 75, 90]],
        y_ticks=[(top + height - height * v / ymax, f"{v:.1f}") for v in [0, 1, 2, 3, 4, 5]],
        x_label="测量天数",
        y_label="剩磁(mT)",
    )
    for t in types_order:
        rows = [r for r in obs if r["type"] == t and r["day"] > 0]
        points = []
        for r in rows:
            x = left + width * r["day"] / 90
            y = top + height - height * r["mag"] / ymax
            points.append((x, y))
        parts.append(polyline(points, color_for_type(t), width=2.5))
        for x, y in points[::6]:
            parts.append(circle(x, y, 2.5, color_for_type(t), opacity=0.8))
    legend_x, legend_y = 1030, 100
    for i, t in enumerate(types_order):
        parts.append(rect(legend_x, legend_y + i * 32, 16, 16, fill=color_for_type(t), stroke=color_for_type(t)))
        parts.append(txt(legend_x + 24, legend_y + 12 + i * 32, t, size=12))
    return save_svg("fig1_raw_decay.svg", "\n".join(parts), W, H)


def make_fig2_normalized_overlay(obs, types_order):
    W, H = 1200, 800
    left, top, width, height = 80, 60, 980, 620
    parts = [txt(W / 2, 28, "四类样品归一化曲线叠加图", size=20, anchor="middle", weight="bold")]
    lookup = {(r["type"], r["id"], r["day"]): r["mag"] for r in obs}
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * d / 90, str(d)) for d in [0, 15, 30, 45, 60, 75, 90]],
        y_ticks=[(top + height - height * v / 1.0, f"{v:.1f}") for v in [0, 0.2, 0.4, 0.6, 0.8, 1.0]],
        x_label="测量天数",
        y_label="归一化保留率",
    )
    for t in types_order:
        sample_ids = sorted({r["id"] for r in obs if r["type"] == t and r["day"] == 0})
        series = {}
        for sid in sample_ids:
            rows = [r for r in obs if r["type"] == t and r["id"] == sid]
            m0 = next(r["mag"] for r in rows if r["day"] == 0)
            pts = []
            for r in rows:
                if r["day"] == 0:
                    continue
                x = left + width * r["day"] / 90
                y = top + height - height * (r["mag"] / m0)
                pts.append((x, y))
            series[sid] = pts
            parts.append(polyline(pts, color_for_type(t), width=1.2, opacity=0.22))
        mean_pts = []
        for day in range(1, 91):
            vals = []
            for sid in sample_ids:
                m0_key = (t, sid, 0)
                day_key = (t, sid, day)
                if m0_key in lookup and day_key in lookup:
                    vals.append(lookup[day_key] / lookup[m0_key])
            x = left + width * day / 90
            if vals:
                y = top + height - height * (sum(vals) / len(vals))
            else:
                y = top + height
            mean_pts.append((x, y))
        parts.append(polyline(mean_pts, color_for_type(t), width=3.0))
    legend_x, legend_y = 1030, 100
    for i, t in enumerate(types_order):
        parts.append(rect(legend_x, legend_y + i * 32, 16, 16, fill=color_for_type(t), stroke=color_for_type(t)))
        parts.append(txt(legend_x + 24, legend_y + 12 + i * 32, t, size=12))
    return save_svg("fig2_normalized_overlay.svg", "\n".join(parts), W, H)


def make_fig3_m0_vs_m90(obs, types_order):
    W, H = 1200, 800
    left, top, width, height = 90, 60, 920, 620
    parts = [txt(W / 2, 28, "M0 与 M90/M0 散点图", size=20, anchor="middle", weight="bold")]
    pts_all = []
    for t in types_order:
        sample_ids = sorted({r["id"] for r in obs if r["type"] == t and r["day"] == 0})
        for sid in sample_ids:
            m0 = next(r["mag"] for r in obs if r["type"] == t and r["id"] == sid and r["day"] == 0)
            m90 = next(r["mag"] for r in obs if r["type"] == t and r["id"] == sid and r["day"] == 90)
            pts_all.append((m0, m90 / m0, t))
    xmin = min(p[0] for p in pts_all) * 0.95
    xmax = max(p[0] for p in pts_all) * 1.05
    ymin = min(p[1] for p in pts_all) * 0.95
    ymax = max(p[1] for p in pts_all) * 1.05
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * (v - xmin) / (xmax - xmin), f"{v:.1f}") for v in [xmin, (xmin + xmax) / 2, xmax]],
        y_ticks=[(top + height - height * (v - ymin) / (ymax - ymin), f"{v:.2f}") for v in [ymin, (ymin + ymax) / 2, ymax]],
        x_label="第0天初值 M0",
        y_label="M90/M0",
    )
    for x, y, t in pts_all:
        sx = left + width * (x - xmin) / (xmax - xmin)
        sy = top + height - height * (y - ymin) / (ymax - ymin)
        parts.append(circle(sx, sy, 5, color_for_type(t), opacity=0.85))
    for t in types_order:
        c = color_for_type(t)
        parts.append(rect(1030, 100 + 32 * types_order.index(t), 16, 16, fill=c, stroke=c))
        parts.append(txt(1054, 112 + 32 * types_order.index(t), t, size=12))
    return save_svg("fig3_m0_vs_m90_ratio.svg", "\n".join(parts), W, H)


def make_fig4_problem2_coef():
    rows = read_csv_rows(OUT / "问题2_主效应与交互效应.csv")
    data = [r for r in rows if r["变量"] not in ("") and r["系数"]]
    data = [r for r in data if r["变量"] not in ("截距",)]
    top_rows = data[:8]
    W, H = 1200, 800
    left, top, width, height = 320, 60, 820, 620
    parts = [txt(W / 2, 28, "问题2主效应与交互效应系数图", size=20, anchor="middle", weight="bold")]
    vals = [float(r["系数"]) for r in top_rows]
    max_abs = max(abs(v) for v in vals) * 1.1
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * (v + max_abs) / (2 * max_abs), f"{v:.1f}") for v in [-max_abs, 0, max_abs]],
        y_ticks=[(top + height - height * (i + 0.5) / len(top_rows), r["变量"]) for i, r in enumerate(top_rows)],
        x_label="系数",
        y_label="变量",
    )
    zero_x = left + width * (0 + max_abs) / (2 * max_abs)
    parts.append(line(zero_x, top, zero_x, top + height, stroke="#888", width=1, dash="5,4"))
    for i, r in enumerate(top_rows):
        coef = float(r["系数"])
        y = top + height * (i + 0.5) / len(top_rows)
        x = left + width * (coef + max_abs) / (2 * max_abs)
        if coef >= 0:
            parts.append(rect(zero_x, y - 8, x - zero_x, 16, fill="#2ca02c", stroke="#2ca02c"))
        else:
            parts.append(rect(x, y - 8, zero_x - x, 16, fill="#d62728", stroke="#d62728"))
        parts.append(txt(20, y + 4, r["变量"], size=12))
    return save_svg("fig4_problem2_coefficients.svg", "\n".join(parts), W, H)


def make_fig5_problem2_ablation():
    rows = read_csv_rows(OUT / "问题2_消融分析.csv")
    versions = [r for r in rows if r["模型版本"]]
    W, H = 1200, 800
    left, top, width, height = 90, 60, 980, 620
    parts = [txt(W / 2, 28, "问题2三种交互构造消融对比", size=20, anchor="middle", weight="bold")]
    max_v = max(float(r["RMSE_m"]) for r in versions) * 1.2
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * i / 3, label) for i, label in enumerate(["无交互", "对称交互", "门槛交互"])],
        y_ticks=[(top + height - height * v / max_v, f"{v:.3f}") for v in [0, max_v / 4, max_v / 2, 3 * max_v / 4, max_v]],
        x_label="模型版本",
        y_label="RMSE_m",
    )
    bar_w = 160
    for i, r in enumerate(versions):
        x = left + 120 + i * 260
        v = float(r["RMSE_m"])
        bar_h = height * v / max_v
        parts.append(rect(x, top + height - bar_h, bar_w, bar_h, fill=["#1f77b4", "#ff7f0e", "#2ca02c"][i], stroke="none", rx=6, opacity=0.85))
        parts.append(txt(x + bar_w / 2, top + height - bar_h - 10, f"{v:.4f}", size=12, anchor="middle"))
        parts.append(txt(x + bar_w / 2, top + height + 20, r["模型版本"], size=12, anchor="middle"))
    return save_svg("fig5_problem2_ablation.svg", "\n".join(parts), W, H)


def make_fig6_dynamic_thresholds():
    rows = read_xlsx_table(OUT / "问题3_动态阈值修正表.xlsx")
    header = rows[0]
    data = rows[1:]
    idx_map = {name: i for i, name in enumerate(header)}
    series = {
        "小号铁钉": [],
        "小号铁夹": [],
        "普通钢筋": [],
        "锈蚀钢筋": [],
    }
    for row in data:
        day = int(float(row[idx_map["序号"]]))
        for t in series:
            val = float(row[idx_map[f"{t}动态阈值(mT)"]])
            series[t].append((day, val))
    W, H = 1200, 800
    left, top, width, height = 90, 60, 980, 620
    parts = [txt(W / 2, 28, "1-90天动态阈值曲线图", size=20, anchor="middle", weight="bold")]
    ymax = max(v for pts in series.values() for _, v in pts) * 1.08
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * d / 90, str(d)) for d in [1, 22, 45, 68, 90]],
        y_ticks=[(top + height - height * v / ymax, f"{v:.2f}") for v in [0, ymax / 4, ymax / 2, 3 * ymax / 4, ymax]],
        x_label="天数",
        y_label="动态阈值(mT)",
    )
    for t, pts in series.items():
        points = [(left + width * d / 90, top + height - height * v / ymax) for d, v in pts]
        parts.append(polyline(points, color_for_type(t), width=2.5))
    legend_x, legend_y = 1030, 100
    for i, t in enumerate(series):
        parts.append(rect(legend_x, legend_y + i * 32, 16, 16, fill=color_for_type(t), stroke=color_for_type(t)))
        parts.append(txt(legend_x + 24, legend_y + 12 + i * 32, t, size=12))
    return save_svg("fig6_dynamic_thresholds.svg", "\n".join(parts), W, H)


def make_fig7_interval_width():
    rows = read_csv_rows(OUT / "问题4_三区域判定结果.csv")
    groups = {}
    for r in rows:
        day = int(r["测量天数"])
        if not r["反推初值下界(mT)"] or not r["反推初值上界(mT)"]:
            continue
        width = float(r["反推初值上界(mT)"]) - float(r["反推初值下界(mT)"])
        groups.setdefault(r["样品类型"], []).append((day, width))
    W, H = 1200, 800
    left, top, width, height = 90, 60, 980, 620
    parts = [txt(W / 2, 28, "反推初值区间宽度-延迟天数曲线图", size=20, anchor="middle", weight="bold")]
    ymax = max(v for pts in groups.values() for _, v in pts) * 1.1
    parts += draw_axes(
        left,
        top,
        width,
        height,
        x_ticks=[(left + width * d / 90, str(d)) for d in [1, 22, 45, 68, 90]],
        y_ticks=[(top + height - height * v / ymax, f"{v:.3f}") for v in [0, ymax / 4, ymax / 2, 3 * ymax / 4, ymax]],
        x_label="延迟天数",
        y_label="区间宽度(mT)",
    )
    for t, pts in groups.items():
        points = [(left + width * d / 90, top + height - height * v / ymax) for d, v in pts]
        parts.append(polyline(points, color_for_type(t), width=2.2))
    return save_svg("fig7_interval_width_delay.svg", "\n".join(parts), W, H)


def make_fig8_flowchart():
    W, H = 1200, 800
    parts = [txt(W / 2, 28, "现场判定流程图", size=20, anchor="middle", weight="bold")]
    boxes = [
        (70, 110, 180, 70, "输入\n样品类型/延迟/天气/剩磁"),
        (300, 110, 180, 70, "预处理\n清洗/校验/匹配"),
        (530, 110, 180, 70, "预测\nR_s(t)"),
        (760, 110, 180, 70, "阈值映射\nΘ_s(t)"),
        (990, 110, 180, 70, "Bootstrap\n区间估计"),
        (300, 300, 180, 70, "三区域判定\n支持/灰区/不支持"),
        (760, 300, 180, 70, "复检建议\n补充证据"),
    ]
    for x, y, w, h, text in boxes:
        parts.append(rect(x, y, w, h, fill="#f7f7f7", stroke="#333", width=1.2, rx=12))
        lines = text.split("\n")
        parts.append(txt(x + w / 2, y + 28, lines[0], size=14, anchor="middle", weight="bold"))
        if len(lines) > 1:
            parts.append(txt(x + w / 2, y + 50, lines[1], size=12, anchor="middle"))
    arrows = [
        (250, 145, 300, 145),
        (480, 145, 530, 145),
        (710, 145, 760, 145),
        (940, 145, 990, 145),
        (1080, 180, 1080, 250),
        (1080, 250, 480, 250),
        (480, 250, 480, 300),
        (480, 335, 760, 335),
        (850, 370, 850, 430),
    ]
    for x1, y1, x2, y2 in arrows:
        parts.append(line(x1, y1, x2, y2, stroke="#555", width=2))
    parts.append(txt(640, 250, "判定输出", size=12, anchor="middle"))
    return save_svg("fig8_flowchart.svg", "\n".join(parts), W, H)


def main():
    csv_path = next(p for p in BASE.iterdir() if p.suffix.lower() == ".csv")
    xlsx_path = next(p for p in BASE.iterdir() if p.suffix.lower() == ".xlsx")
    obs, types_order = read_experiment(csv_path)
    weather_rows = read_xlsx_weather(xlsx_path)
    make_fig1_raw_decay(obs, types_order)
    make_fig2_normalized_overlay(obs, types_order)
    make_fig3_m0_vs_m90(obs, types_order)
    make_fig4_problem2_coef()
    make_fig5_problem2_ablation()
    make_fig6_dynamic_thresholds()
    make_fig7_interval_width()
    make_fig8_flowchart()


if __name__ == "__main__":
    main()
