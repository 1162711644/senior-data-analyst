#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
根据 JSON 规格生成一份带图表的 Word 分析报告（A4，蓝白/红白两套色系）。

用法：
    python3 build_report_docx.py --spec report.json --out /path/报告.docx [--theme blue|red]

规格示例见 references/docx-spec.md。脚本自动完成：
  - 标题区（主标题 + 副标题 + 分隔线）
  - 读表摘要框
  - 五段正文（固定命名由 spec 提供）
  - 表格（表题在上）、图表（折线/条形/环形，图题在下并自动补单位/时间范围/来源）
  - 项目符号、提示框、页脚页码

依赖：python-docx、matplotlib（缺则 pip3 install python-docx matplotlib）
"""
import argparse
import json
import os
import shutil
import sys
import tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ---------------------------------------------------------------- 主题与常量

THEMES = {
    "blue": {
        "primary": "1F4E79",      # 深蓝：标题、表头
        "accent": "2E75B6",       # 辅助蓝
        "light": "EAF1F8",        # 浅蓝：提示框底色
        "chart": ["1F4E79", "2E75B6", "5B9BD5", "A6C9E8", "BFD7ED", "8EA9C6"],
        "up": "C00000",           # 涨=红
        "down": "1E7B34",         # 跌=绿
    },
    "red": {
        "primary": "B42318",      # 砖红：标题、表头
        "accent": "D9453D",
        "light": "FBEAE7",
        "chart": ["8E1B12", "B42318", "D9453D", "EA8A84", "F5C9C5", "C77B74"],
        "up": "C00000",
        "down": "1E7B34",
    },
}

BODY_FONT = "宋体"
ASCII_FONT = "Times New Roman"
HEAD_FONT = "微软雅黑"
BODY_SIZE = 12          # 小四
CONTENT_WIDTH_CM = 15.6 # A4 减去左右页边距

CJK_CANDIDATES = [
    "Noto Sans CJK SC", "Noto Sans CJK JP", "Source Han Sans SC",
    "WenQuanYi Zen Hei", "WenQuanYi Micro Hei", "AR PL UMing CN",
]


def setup_matplotlib_font():
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in CJK_CANDIDATES:
        if name in installed:
            plt.rcParams["font.sans-serif"] = [name, "DejaVu Sans"]
            break
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 200


# ---------------------------------------------------------------- 基础工具

def set_run(run, size=BODY_SIZE, bold=False, color=None, font=BODY_FONT,
            ascii_font=ASCII_FONT, italic=False):
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), font)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def rich_runs(paragraph, text, size=BODY_SIZE, color=None, font=BODY_FONT):
    """支持 **加粗** 的简易富文本。"""
    parts = text.split("**")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = paragraph.add_run(part)
        set_run(run, size=size, bold=(i % 2 == 1), color=color, font=font)


def para_format(p, before=0, after=6, line=1.35, indent_chars=0, align=None):
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if indent_chars:
        pf.first_line_indent = Pt(BODY_SIZE * indent_chars)
    if align is not None:
        p.alignment = align
    return p


def add_bottom_border(paragraph, color, sz=12, space=2):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(sz))
    bottom.set(qn("w:space"), str(space))
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def shade_cell(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hexcolor)
    tcPr.append(shd)


def set_table_borders(table, color="BFBFBF", sz=4, inside=True):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    edges = ["top", "left", "bottom", "right"]
    if inside:
        edges += ["insideH", "insideV"]
    for edge in edges:
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def set_cell_margins(table, top=60, bottom=60, left=100, right=100):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for name, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def add_field(paragraph, field="PAGE"):
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)
    return run


# ---------------------------------------------------------------- 图表

def fmt_num(v):
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(v) >= 1e8:
        return f"{v / 1e8:.2f}亿"
    if abs(v) >= 1e4:
        return f"{v / 1e4:.1f}万"
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}"
    return f"{v:,.1f}"


def value_label(v, unit=""):
    if "%" in unit:
        return f"{v:.1f}%"
    return fmt_num(v)


def render_chart(spec, theme, outdir, idx):
    kind = spec.get("kind", "line")
    unit = spec.get("unit", "")
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    palette = ["#" + h.lstrip("#") for h in theme["chart"]]

    if kind == "line":
        x = spec.get("x", [])
        for i, s in enumerate(spec.get("series", [])):
            ax.plot(x, s.get("values", []), marker="o", markersize=4.5,
                    linewidth=2, color=palette[i % len(palette)],
                    label=s.get("name", ""))
        ax.set_xticks(range(len(x)))
        ax.set_xticklabels([str(t) for t in x],
                           fontsize=9,
                           rotation=30 if len(x) > 8 else 0,
                           ha="right" if len(x) > 8 else "center")
        ax.grid(axis="y", color="#E5E5E5", linewidth=0.8)
        if len(spec.get("series", [])) > 1:
            ax.legend(frameon=False, fontsize=9, ncol=min(4, len(spec["series"])))
        if unit and "%" not in unit:
            ax.set_ylabel(f"单位：{unit}", fontsize=9, color="#595959")

    elif kind in ("bar", "column"):
        cats = spec.get("categories", [])
        vals = spec.get("values", [])
        colors = [palette[i % len(palette)] for i in range(len(vals))]
        if spec.get("orientation", "v") == "h":
            bars = ax.barh([str(c) for c in cats], vals, color=colors, height=0.62)
            ax.invert_yaxis()
            ax.grid(axis="x", color="#E5E5E5", linewidth=0.8)
            ax.set_axisbelow(True)
            vmax = max([abs(v) for v in vals] + [1])
            for b, v in zip(bars, vals):
                ax.text(b.get_width() + vmax * 0.02, b.get_y() + b.get_height() / 2,
                        value_label(v, unit), va="center", fontsize=9, color="#595959")
            ax.set_xlim(0, vmax * 1.22)
            if unit and "%" not in unit:
                ax.set_xlabel(f"单位：{unit}", fontsize=9, color="#595959")
        else:
            bars = ax.bar([str(c) for c in cats], vals, color=colors, width=0.55)
            ax.grid(axis="y", color="#E5E5E5", linewidth=0.8)
            ax.set_axisbelow(True)
            vmax = max([abs(v) for v in vals] + [1])
            for b, v in zip(bars, vals):
                ax.text(b.get_x() + b.get_width() / 2,
                        b.get_height() + vmax * 0.015,
                        value_label(v, unit), ha="center", fontsize=9, color="#595959")
            ax.set_ylim(min(0, min(vals)), vmax * 1.18)
            ax.set_xticklabels([str(c) for c in cats], fontsize=9,
                               rotation=20 if len(cats) > 6 else 0)
            if unit and "%" not in unit:
                ax.set_ylabel(f"单位：{unit}", fontsize=9, color="#595959")

    elif kind == "donut":
        cats = spec.get("categories", [])
        vals = spec.get("values", [])
        wedges, _, autotexts = ax.pie(
            vals, colors=[palette[i % len(palette)] for i in range(len(vals))],
            startangle=90, counterclock=False,
            wedgeprops=dict(width=0.42, edgecolor="white", linewidth=1.5),
            autopct="%1.1f%%", pctdistance=0.79,
            textprops=dict(fontsize=8.5, color="white"),
        )
        for t in autotexts:
            t.set_color("#FFFFFF")
        ax.legend(wedges,
                  [f"{c}  {value_label(v, unit)}" for c, v in zip(cats, vals)],
                  loc="center left", bbox_to_anchor=(1.0, 0.5),
                  frameon=False, fontsize=9)
        ax.set_aspect("equal")
    else:
        raise ValueError(f"不支持的图表类型：{kind}")

    for side in ("top", "right"):
        ax.spines[side].set_visible(False)
    if kind != "donut":
        for side in ("left", "bottom"):
            ax.spines[side].set_color("#BFBFBF")
        ax.tick_params(colors="#595959", labelsize=9)
    fig.tight_layout()
    path = os.path.join(outdir, f"chart_{idx}.png")
    fig.savefig(path, dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return path


# ---------------------------------------------------------------- 文档部件

def add_title_block(doc, spec, theme):
    t = doc.add_paragraph()
    para_format(t, before=0, after=4, line=1.15)
    set_run(t.add_run(spec.get("title", "业务数据分析报告")),
            size=22, bold=True, color=theme["primary"], font=HEAD_FONT)

    sub = spec.get("subtitle") or spec.get("meta_line")
    if sub:
        p = doc.add_paragraph()
        para_format(p, before=0, after=10, line=1.2)
        set_run(p.add_run(sub), size=10.5, color="808080", font=HEAD_FONT)

    rule = doc.add_paragraph()
    para_format(rule, before=0, after=10, line=1)
    add_bottom_border(rule, theme["primary"], sz=14, space=1)


def add_callout(doc, label, text, theme, size=10.5):
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(CONTENT_WIDTH_CM)
    shade_cell(cell, theme["light"])
    set_table_borders(table, color=theme["light"], sz=4, inside=False)
    set_cell_margins(table, top=140, bottom=140, left=200, right=200)

    p = cell.paragraphs[0]
    para_format(p, before=0, after=0, line=1.3)
    if label:
        set_run(p.add_run(label), size=size, bold=True, color=theme["primary"], font=HEAD_FONT)
    rich_runs(p, text, size=size, color="262626")

    spacer = doc.add_paragraph()
    para_format(spacer, before=0, after=0, line=0.6)


def add_heading1(doc, text, theme, page_break=False):
    p = doc.add_paragraph()
    if page_break:
        p.paragraph_format.page_break_before = True
    para_format(p, before=18, after=8, line=1.2)
    set_run(p.add_run(text), size=16, bold=True, color=theme["primary"], font=HEAD_FONT)
    add_bottom_border(p, theme["light"], sz=8, space=4)
    return p


def add_heading2(doc, text, theme):
    p = doc.add_paragraph()
    para_format(p, before=10, after=4, line=1.25)
    set_run(p.add_run(text), size=13, bold=True, color="262626", font=HEAD_FONT)
    return p


def add_body(doc, text, theme, indent=True, size=BODY_SIZE, color=None, align=None):
    p = doc.add_paragraph()
    para_format(p, before=0, after=6, line=1.4,
                indent_chars=2 if indent else 0,
                align=align if align is not None else WD_ALIGN_PARAGRAPH.JUSTIFY)
    rich_runs(p, text, size=size, color=color)
    return p


def add_bullets(doc, items, theme, size=BODY_SIZE):
    for item in items:
        p = doc.add_paragraph()
        para_format(p, before=0, after=5, line=1.35)
        pf = p.paragraph_format
        pf.left_indent = Cm(0.75)
        pf.first_line_indent = Cm(-0.4)
        set_run(p.add_run("•  "), size=size, color=theme["primary"], bold=True)
        rich_runs(p, item, size=size)


def add_table(doc, block, theme):
    caption = block.get("caption")
    if caption:
        p = doc.add_paragraph()
        para_format(p, before=10, after=4, line=1.2, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_run(p.add_run(caption), size=10, bold=True, color="262626", font=HEAD_FONT)

    header = block.get("header", [])
    rows = block.get("rows", [])
    ncol = max([len(header)] + [len(r) for r in rows]) if rows or header else 1
    table = doc.add_table(rows=0, cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    set_table_borders(table, color="BFBFBF", sz=4)
    set_cell_margins(table)

    if header:
        hr = table.add_row()
        repeat_header(hr)
        for j, h in enumerate(header):
            cell = hr.cells[j]
            shade_cell(cell, theme["primary"])
            p = cell.paragraphs[0]
            para_format(p, before=0, after=0, line=1.1, align=WD_ALIGN_PARAGRAPH.CENTER)
            set_run(p.add_run(str(h)), size=10, bold=True, color="FFFFFF", font=HEAD_FONT)

    for r in rows:
        row = table.add_row()
        for j in range(ncol):
            cell = row.cells[j]
            val = str(r[j]) if j < len(r) else ""
            p = cell.paragraphs[0]
            para_format(p, before=0, after=0, line=1.1,
                        align=WD_ALIGN_PARAGRAPH.LEFT if j == 0 else WD_ALIGN_PARAGRAPH.CENTER)
            set_run(p.add_run(val), size=10, color="262626")

    widths = block.get("col_widths")
    if widths:
        for j, w in enumerate(widths[:ncol]):
            for row in table.rows:
                row.cells[j].width = Cm(w)

    note = block.get("note")
    if note:
        p = doc.add_paragraph()
        para_format(p, before=3, after=8, line=1.15)
        set_run(p.add_run(note), size=8.5, color="808080")
    else:
        spacer = doc.add_paragraph()
        para_format(spacer, before=0, after=4, line=0.6)


def add_chart(doc, block, theme, outdir, idx):
    chart_spec = block.get("chart", block)
    path = render_chart(chart_spec, theme, outdir, idx)
    p = doc.add_paragraph()
    para_format(p, before=8, after=4, line=1.0, align=WD_ALIGN_PARAGRAPH.CENTER)
    p.add_run().add_picture(path, width=Cm(CONTENT_WIDTH_CM))

    caption = block.get("caption") or chart_spec.get("caption")
    if caption:
        cp = doc.add_paragraph()
        para_format(cp, before=2, after=2, line=1.2, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_run(cp.add_run(caption), size=9.5, bold=True, color="404040", font=HEAD_FONT)

    bits = []
    if chart_spec.get("unit"):
        bits.append(f"单位：{chart_spec['unit']}")
    if chart_spec.get("time_range"):
        bits.append(f"时间范围：{chart_spec['time_range']}")
    if chart_spec.get("source"):
        bits.append(f"数据来源：{chart_spec['source']}")
    if bits:
        np_ = doc.add_paragraph()
        para_format(np_, before=0, after=10, line=1.15, align=WD_ALIGN_PARAGRAPH.CENTER)
        set_run(np_.add_run("｜".join(bits)), size=8.5, color="808080")


def build_footer(doc, theme, footer_text):
    section = doc.sections[0]
    fp = section.footer.paragraphs[0]
    para_format(fp, before=0, after=0, line=1.0, align=WD_ALIGN_PARAGRAPH.RIGHT)
    set_run(fp.add_run(footer_text or "资深数据分析师 · 数据分析报告"), size=8.5, color="808080")
    set_run(fp.add_run("　｜　第 "), size=8.5, color="808080")
    add_field(fp, "PAGE")
    set_run(fp.add_run(" 页"), size=8.5, color="808080")


def add_block(doc, block, theme, outdir, counter):
    btype = block.get("type")
    if btype == "paragraph":
        add_body(doc, block.get("text", ""), theme,
                 indent=block.get("indent", True))
    elif btype == "bullets":
        add_bullets(doc, block.get("items", []), theme)
    elif btype == "subheading":
        add_heading2(doc, block.get("text", ""), theme)
    elif btype == "table":
        add_table(doc, block, theme)
    elif btype == "chart":
        counter["chart"] += 1
        add_chart(doc, block, theme, outdir, counter["chart"])
    elif btype == "callout":
        add_callout(doc, block.get("label", ""), block.get("text", ""), theme)
    else:
        raise ValueError(f"不支持的 block 类型：{btype}")


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec", required=True, help="报告 JSON 规格文件")
    ap.add_argument("--out", required=True, help="输出 .docx 路径")
    ap.add_argument("--theme", choices=["blue", "red"], default=None)
    args = ap.parse_args()

    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)

    theme_name = args.theme or spec.get("theme") or "blue"
    if theme_name not in THEMES:
        theme_name = "blue"
    theme = THEMES[theme_name]

    setup_matplotlib_font()
    doc = Document()

    # 页面设置
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin, sec.bottom_margin = Cm(2.4), Cm(2.3)
    sec.left_margin, sec.right_margin = Cm(2.7), Cm(2.7)

    # 默认样式
    normal = doc.styles["Normal"]
    normal.font.name = ASCII_FONT
    normal.font.size = Pt(BODY_SIZE)
    rPr = normal.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), BODY_FONT)

    build_footer(doc, theme, spec.get("footer"))

    # 标题区
    add_title_block(doc, spec, theme)

    # 读表摘要
    summary = spec.get("summary")
    if summary:
        add_callout(doc, "读表摘要｜", summary, theme)

    counter = {"chart": 0}
    outdir = tempfile.mkdtemp(prefix="lm_report_")

    try:
        for i, section in enumerate(spec.get("sections", [])):
            add_heading1(doc, section.get("heading", ""), theme,
                         page_break=(spec.get("page_break_before_body", False) and i == 0))
            for text in section.get("paragraphs", []):
                add_body(doc, text, theme)
            for block in section.get("blocks", []):
                add_block(doc, block, theme, outdir, counter)

        appendix = spec.get("appendix")
        if appendix:
            add_heading1(doc, "附：数据说明", theme)
            for text in appendix.get("paragraphs", []):
                add_body(doc, text, theme)
            for block in appendix.get("blocks", []):
                add_block(doc, block, theme, outdir, counter)

        note = spec.get("usage_note")
        if note:
            add_heading1(doc, "使用说明", theme)
            add_callout(doc, "", note, theme)

        out = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out), exist_ok=True)
        doc.save(out)
        print(f"OK 已生成：{out}")
        print(f"    色系：{'蓝白' if theme_name == 'blue' else '红白'}｜段落：{len(spec.get('sections', []))} 段｜图表：{counter['chart']} 张")
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
