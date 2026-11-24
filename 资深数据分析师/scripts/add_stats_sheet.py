#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在原表副本中新增「公式驱动」的统计结果 sheet。

用法:
    python3 add_stats_sheet.py spec.json          # 按 spec 生成
    python3 add_stats_sheet.py --check out.xlsx   # 自检：sheet/公式数/重算标记

设计原则:
    1. 原表数据区默认整列引用（如 '明细'!$F:$F），数据增删后自动重算；
       若原表底部有合计行等非明细行，必须在 spec 里给 data_range（如 "A2:F5000"）。
    2. 只写 output 指定的副本；source 文件只读不写。
    3. 结果 sheet 里所有指标都是 Excel 公式，不落死值，Excel/WPS 打开即自动计算。
"""
import argparse
import json
import os
import shutil
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True)
NOTE_FONT = Font(size=9, color="595959")
GROUP_FILL = PatternFill("solid", fgColor="EAF1F8")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CN_NUMFMT = "#,##0.00"
PCT_FMT = "0.00%"


# ---------- 基础工具 ----------
def qname(sheet):
    return "'" + sheet.replace("'", "''") + "'"


def split_range(data_range):
    start, end = data_range.split(":")
    sc = "".join(ch for ch in start if ch.isalpha())
    ec = "".join(ch for ch in end if ch.isalpha())
    sr = "".join(ch for ch in start if ch.isdigit())
    er = "".join(ch for ch in end if ch.isdigit())
    return sc, sr, ec, er


def col_ref(sheet, col_letter, data_range=None):
    """带表名的整列或数据区引用，如 '明细'!$F:$F。"""
    if data_range:
        _, sr, _, er = split_range(data_range)
        return f"{qname(sheet)}!${col_letter}${sr}:${col_letter}${er}"
    return f"{qname(sheet)}!${col_letter}:${col_letter}"


def header_map(ws, header_row):
    mapping, dup = {}, []
    for cell in ws[header_row]:
        name = str(cell.value).strip() if cell.value is not None else ""
        if not name:
            continue
        if name in mapping:
            dup.append(name)
        mapping[name] = get_column_letter(cell.column)
    if dup:
        raise SystemExit(f"表头重名，无法定位字段：{dup}。先把表头改成唯一名称再生成。")
    if not mapping:
        raise SystemExit(f"第 {header_row} 行没有读到表头，请确认 header_row 设置。")
    return mapping


def need_field(headers, field, sheet=""):
    if field not in headers:
        raise SystemExit(f"字段“{field}”不在 {sheet or '表'} 的表头里；可用字段：{list(headers)}")
    return headers[field]


def table_width(headers):
    return max(column_index_from_string(c) for c in headers.values())


def style_header(ws, row, ncol, widths=None):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 26
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def write_notes(ws, row, lines):
    for i, line in enumerate(lines):
        ws.cell(row=row + i, column=1, value=line).font = NOTE_FONT


def collect_groups(ws, header_row, group_col_letters, keep_blank=False, data_range=None):
    rows, seen = [], set()
    end_row = ws.max_row
    if data_range:
        end_row = int("".join(ch for ch in data_range.split(":")[1] if ch.isdigit()))
    for row in ws.iter_rows(min_row=header_row + 1, max_row=end_row, values_only=True):
        vals = []
        for c in group_col_letters:
            idx = column_index_from_string(c) - 1
            v = row[idx] if idx < len(row) else None
            if isinstance(v, str):
                v = v.strip()
            vals.append(v)
        if not keep_blank and all(v in (None, "") for v in vals):
            continue
        key = tuple("" if v is None else v for v in vals)
        if key in seen:
            continue
        seen.add(key)
        rows.append(list(key))
    return rows


# ---------- total：整表汇总 ----------
def add_total_sheet(wb, spec):
    src = spec["source"]
    sh, hr, dr = src["sheet"], src.get("header_row", 1), src.get("data_range")
    headers = header_map(wb[sh], hr)
    ws = wb.create_sheet(spec["name"])
    ws["A1"], ws["B1"], ws["C1"] = "指标", "统计结果", "口径"
    style_header(ws, 1, 3, [26, 18, 56])

    r = 2
    for m in spec["metrics"]:
        kind, label = m["type"], m["label"]
        if kind in ("sum", "average", "max", "min"):
            col = need_field(headers, m["field"], sh)
            rng = col_ref(sh, col, dr)
            fn = {"sum": "SUM", "average": "AVERAGE", "max": "MAX", "min": "MIN"}[kind]
            f = f"={fn}({rng})" if kind in ("sum", "max", "min") else f'=IFERROR(AVERAGE({rng}),"")'
            desc = f"{sh}.{m['field']} 的{'求和' if kind == 'sum' else {'average': '平均', 'max': '最大值', 'min': '最小值'}[kind]}"
        elif kind == "count":
            col = need_field(headers, m["field"], sh) if m.get("field") else list(headers.values())[0]
            f = f"=COUNTA({col_ref(sh, col, dr)})" + ("" if dr else "-1")
            desc = "行数（不含表头）"
        else:
            raise SystemExit(f"total 版不支持指标类型：{kind}（可用 sum/average/max/min/count）")
        ws.cell(row=r, column=1, value=label).border = BORDER
        c = ws.cell(row=r, column=2, value=f)
        c.number_format = "0" if kind == "count" else (m.get("numfmt") or CN_NUMFMT)
        c.border = BORDER
        ws.cell(row=r, column=3, value=desc).font = NOTE_FONT
        r += 1

    write_notes(ws, r + 1, [
        f"数据源：{sh}（表头第 {hr} 行，数据区：{dr or '整列'}）",
        "口径：" + "、".join(m["label"] for m in spec["metrics"]),
        "生成：" + datetime.now().strftime("%Y-%m-%d %H:%M") + "；结果为公式，原表数据增减后自动重算。",
    ] + ([str(spec["notes"])] if spec.get("notes") else []))
    return ws


# ---------- group：分组汇总 ----------
def group_metric_formula(kind, m, row, headers, sh, dr, crit_tmpl):
    crit = ",".join(f"{rng},{cell.replace('{row}', str(row))}" for rng, cell in crit_tmpl)
    if kind in ("sum", "average", "max", "min"):
        col = need_field(headers, m["field"], sh)
        rng = col_ref(sh, col, dr)
        if kind == "sum":
            return f"=SUMIFS({rng},{crit})"
        fn = {"average": "AVERAGEIFS", "max": "_xlfn.MAXIFS", "min": "_xlfn.MINIFS"}[kind]
        return f'=IFERROR({fn}({rng},{crit}),"")'
    if kind in ("count", "count_field"):
        if m.get("field"):
            rng = col_ref(sh, need_field(headers, m["field"], sh), dr)
            return f'=COUNTIFS({crit},{rng},"<>")'
        rng = col_ref(sh, list(headers.values())[0], dr)
        return f'=COUNTIFS({crit},{rng},"<>")'
    raise SystemExit(f"group 版不支持指标类型：{kind}")


def add_group_sheet(wb, spec):
    src = spec["source"]
    sh, hr, dr = src["sheet"], src.get("header_row", 1), src.get("data_range")
    headers = header_map(wb[sh], hr)
    groups = spec["group_by"]
    for g in groups:
        need_field(headers, g, sh)
    labels = collect_groups(wb[sh], hr, [headers[g] for g in groups], spec.get("keep_blank", False), dr)
    if not labels:
        raise SystemExit("分组维度没有取到任何数据，检查 header_row 和 group_by 是否对。")
    if spec.get("sort_groups"):
        labels.sort(key=lambda x: tuple((v is None, str(v)) for v in x))

    ws = wb.create_sheet(spec["name"])
    col = 1
    group_cols = []
    for g in groups:
        ws.cell(row=1, column=col, value=g)
        group_cols.append(col)
        col += 1
    metric_idx = {}
    for m in spec["metrics"]:
        if m["label"] in metric_idx:
            raise SystemExit(f"指标名重复：{m['label']}，改个名字再生成。")
        ws.cell(row=1, column=col, value=m["label"])
        metric_idx[m["label"]] = col
        col += 1
    style_header(ws, 1, col - 1, [18] * len(groups) + [16] * len(spec["metrics"]))

    first, last = 2, 1 + len(labels)
    for i, vals in enumerate(labels):
        for j, v in enumerate(vals):
            cell = ws.cell(row=first + i, column=group_cols[j], value=v)
            cell.fill = GROUP_FILL
            cell.border = BORDER
        for cidx in metric_idx.values():
            ws.cell(row=first + i, column=cidx).border = BORDER

    crit_tmpl = [(col_ref(sh, headers[g], dr), get_column_letter(group_cols[j]) + "{row}")
                 for j, g in enumerate(groups)]

    def ref(label, row):
        return f"${get_column_letter(metric_idx[label])}${row}"

    for m in spec["metrics"]:
        cidx, kind, label = metric_idx[m["label"]], m["type"], m["label"]
        fmt = m.get("numfmt")
        for i in range(len(labels)):
            row = first + i
            if kind in ("sum", "average", "max", "min", "count", "count_field"):
                f = group_metric_formula(kind, m, row, headers, sh, dr, crit_tmpl)
            elif kind == "share_of":
                f = f'=IFERROR({ref(m["base"], row)}/SUM({ref(m["base"], first)}:{ref(m["base"], last)}),"")'
                fmt = fmt or PCT_FMT
            elif kind == "rank_of":
                f = (f'=IFERROR(RANK({ref(m["base"], row)},'
                     f'{ref(m["base"], first)}:{ref(m["base"], last)},{m.get("order", 0)}),"")')
                fmt = fmt or "0"
            elif kind == "ratio_of":
                f = f'=IFERROR({ref(m["numerator"], row)}/{ref(m["denominator"], row)},"")'
                fmt = fmt or PCT_FMT
            elif kind == "diff_of":
                f = f'={ref(m["a"], row)}-{ref(m["b"], row)}'
                fmt = fmt or CN_NUMFMT
            elif kind == "sum_of":
                f = f'={ref(m["a"], row)}+{ref(m["b"], row)}'
                fmt = fmt or CN_NUMFMT
            elif kind == "product_of":
                f = f'={ref(m["a"], row)}*{ref(m["b"], row)}'
                fmt = fmt or CN_NUMFMT
            elif kind == "custom":
                f = m["formula_tpl"].replace("{row}", str(row))
                fmt = fmt or CN_NUMFMT
            else:
                raise SystemExit(f"group 版不支持指标类型：{kind}")
            c = ws.cell(row=row, column=cidx, value=f)
            if not fmt:
                fmt = CN_NUMFMT if kind in ("sum", "average", "max", "min", "count", "count_field") else "General"
            if kind == "count" and not m.get("numfmt"):
                fmt = "0"
            c.number_format = fmt

    if spec.get("freeze_header", True):
        ws.freeze_panes = ws.cell(row=2, column=1)

    write_notes(ws, last + 2, [
        f"数据源：{sh}（表头第 {hr} 行，数据区：{dr or '整列'}）；分组：{' + '.join(groups)}，共 {len(labels)} 组",
        "口径：" + "、".join(f"{m['label']}={m['type']}" for m in spec["metrics"]),
        "生成：" + datetime.now().strftime("%Y-%m-%d %H:%M") + "；指标为公式，原表数据增减后自动重算。",
    ] + ([str(spec["notes"])] if spec.get("notes") else []))
    return ws


# ---------- join：跨表关联 ----------
def add_join_sheet(wb, spec):
    left, right = spec["left"], spec["right"]
    ws_l, ws_r = wb[left["sheet"]], wb[right["sheet"]]
    hr_l, hr_r = left.get("header_row", 1), right.get("header_row", 1)
    h_l, h_r = header_map(ws_l, hr_l), header_map(ws_r, hr_r)
    keys = left["key"] if isinstance(left["key"], list) else [left["key"]]
    rkeys = right["key"] if isinstance(right["key"], list) else [right["key"]]
    if len(keys) != len(rkeys):
        raise SystemExit("左右关联键数量不一致。")
    for k in keys:
        need_field(h_l, k, left["sheet"])
    for k in rkeys:
        need_field(h_r, k, right["sheet"])

    keep = list(left.get("keep", []))
    for k in keys:
        if k not in keep:
            keep.insert(0, k)
    for k in keep:
        need_field(h_l, k, left["sheet"])
    bring = right.get("bring", [])
    for b in bring:
        need_field(h_r, b["field"], right["sheet"])

    out_cols = keep + [b["label"] for b in bring] + [m["label"] for m in spec.get("metrics", [])]
    ws = wb.create_sheet(spec["name"])
    for i, name in enumerate(out_cols, start=1):
        ws.cell(row=1, column=i, value=name)
    style_header(ws, 1, len(out_cols), [16] * len(out_cols))

    rows = []
    for src_row, row in enumerate(ws_l.iter_rows(min_row=hr_l + 1, values_only=True), start=hr_l + 1):
        vals = []
        for name in keep:
            idx = column_index_from_string(h_l[name]) - 1
            v = row[idx] if idx < len(row) else None
            if isinstance(v, str):
                v = v.strip()
            vals.append(v)
        if all(v in (None, "") for v in vals):
            continue
        rows.append((src_row, vals))
    if spec.get("distinct_rows", True):
        seen, uniq = set(), []
        for src_row, v in rows:
            k = tuple("" if x is None else x for x in v)
            if k not in seen:
                seen.add(k)
                uniq.append((src_row, v))
        rows = uniq
    if not rows:
        raise SystemExit(f"左表 {left['sheet']} 没有抽取到数据行。")

    for ri, (src_row, vals) in enumerate(rows, start=2):
        for ci, name in enumerate(keep, start=1):
            f = f"={qname(left['sheet'])}!${h_l[name]}${src_row}"
            ws.cell(row=ri, column=ci, value=f).border = BORDER
    last = 1 + len(rows)
    letter = {name: get_column_letter(i) for i, name in enumerate(out_cols, start=1)}
    key_letters = [letter[k] for k in keys]

    for b in bring:
        cidx = out_cols.index(b["label"]) + 1
        rcol = h_r[b["field"]]
        for ri in range(2, last + 1):
            if len(keys) == 1:
                f = (f'=IFERROR(INDEX({col_ref(right["sheet"], rcol)},'
                     f'MATCH(${key_letters[0]}{ri},{col_ref(right["sheet"], h_r[rkeys[0]])},0)),"")')
            else:
                cond = "*".join(f"({col_ref(right['sheet'], h_r[k])}=${kl}{ri})" for k, kl in zip(rkeys, key_letters))
                f = f'=IFERROR(INDEX({col_ref(right["sheet"], rcol)},MATCH(1,INDEX(({cond}),0),0)),"")'
            ws.cell(row=ri, column=cidx, value=f).border = BORDER

    for m in spec.get("metrics", []):
        cidx, kind = out_cols.index(m["label"]) + 1, m["type"]
        fmt = m.get("numfmt")
        for ri in range(2, last + 1):
            if kind == "ratio_of":
                f = f'=IFERROR({letter[m["numerator"]]}{ri}/{letter[m["denominator"]]}{ri},"")'
                fmt = fmt or PCT_FMT
            elif kind == "diff_of":
                f = f'={letter[m["a"]]}{ri}-{letter[m["b"]]}{ri}'
                fmt = fmt or CN_NUMFMT
            elif kind == "sum_of":
                f = f'={letter[m["a"]]}{ri}+{letter[m["b"]]}{ri}'
                fmt = fmt or CN_NUMFMT
            elif kind == "product_of":
                f = f'={letter[m["a"]]}{ri}*{letter[m["b"]]}{ri}'
                fmt = fmt or CN_NUMFMT
            elif kind == "custom":
                f = m["formula_tpl"].replace("{row}", str(ri))
                fmt = fmt or CN_NUMFMT
            else:
                raise SystemExit(f"join 版不支持指标类型：{kind}")
            c = ws.cell(row=ri, column=cidx, value=f)
            c.number_format = fmt
            c.border = BORDER

    ws.freeze_panes = "A2"
    write_notes(ws, last + 2, [
        f"左表：{left['sheet']}（输出 {len(rows)} 行）　右表：{right['sheet']}",
        "关联键：" + " + ".join(keys) + "；右表字段用 INDEX/MATCH 取数，匹配不上留空。",
        "生成：" + datetime.now().strftime("%Y-%m-%d %H:%M") + "；公式随两表数据变化重算。",
    ] + ([str(spec["notes"])] if spec.get("notes") else []))
    return ws


# ---------- combine：多表同字段合计 ----------
def add_combine_sheet(wb, spec):
    ws = wb.create_sheet(spec["name"])
    ws["A1"], ws["B1"], ws["C1"] = "指标", "统计结果", "口径"
    style_header(ws, 1, 3, [26, 18, 64])
    label_row = {}
    rows = []
    for i, m in enumerate(spec["metrics"], start=2):
        label_row[m["label"]] = i
        rows.append(m)
    for i, m in enumerate(spec["metrics"], start=2):
        kind = m["type"]
        sums, cnts = [], []
        for s in spec["sources"]:
            wsrc = wb[s["sheet"]]
            headers = header_map(wsrc, s.get("header_row", 1))
            fld = need_field(headers, s["field"], s["sheet"])
            sums.append(f"SUM({col_ref(s['sheet'], fld, s.get('data_range') or spec.get('data_range'))})")
            cfld = headers[s.get("count_field", s["field"])]
            cnts.append(f"COUNTA({col_ref(s['sheet'], cfld, s.get('data_range') or spec.get('data_range'))})-1")
        if kind == "sum_multi":
            f = "=" + "+".join(sums)
            fmt, desc = CN_NUMFMT, " + ".join(f"{s['sheet']}.{s['field']}" for s in spec["sources"])
        elif kind == "count_multi":
            f = "=" + "+".join(cnts)
            fmt, desc = "0", " + ".join(f"{s['sheet']} 行数" for s in spec["sources"])
        elif kind == "average_multi":
            f = f'=IFERROR(({"+".join(sums)})/({"+".join(cnts)}),"")'
            fmt, desc = CN_NUMFMT, "多表加权平均（合计/行数）"
        elif kind == "ratio_of":
            num, den = label_row[m["numerator"]], label_row[m["denominator"]]
            f = f'=IFERROR(B{num}/B{den},"")'
            fmt, desc = PCT_FMT, f"{m['numerator']} ÷ {m['denominator']}"
        else:
            raise SystemExit(f"combine 版不支持指标类型：{kind}")
        ws.cell(row=i, column=1, value=m["label"]).border = BORDER
        c = ws.cell(row=i, column=2, value=f)
        c.number_format = m.get("numfmt") or fmt
        c.border = BORDER
        ws.cell(row=i, column=3, value=desc).font = NOTE_FONT

    write_notes(ws, len(spec["metrics"]) + 3, [
        "多表合并：各表同一字段相加，整列引用，自动重算。",
        "生成：" + datetime.now().strftime("%Y-%m-%d %H:%M"),
    ] + ([str(spec["notes"])] if spec.get("notes") else []))
    return ws


LAYOUTS = {"total": add_total_sheet, "group": add_group_sheet, "join": add_join_sheet, "combine": add_combine_sheet}


def check(path):
    wb = openpyxl.load_workbook(path)
    print("sheets:", wb.sheetnames)
    total = 0
    for ws in wb.worksheets:
        n = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
        total += n
        print(f"  {ws.title}: {ws.max_row} 行 x {ws.max_column} 列，公式 {n} 个")
    print("fullCalcOnLoad:", getattr(getattr(wb, "calculation", None), "fullCalcOnLoad", None))
    print("公式总数:", total)


def main():
    ap = argparse.ArgumentParser(description="在原表副本里新增公式驱动的统计结果 sheet")
    ap.add_argument("spec", nargs="?", help="spec.json 路径")
    ap.add_argument("--check", help="检查工作簿：sheet 列表、公式数量、重算标记")
    args = ap.parse_args()
    if args.check:
        check(args.check)
        return
    if not args.spec:
        ap.error("需要 spec.json 或 --check")

    with open(args.spec, encoding="utf-8") as fh:
        spec = json.load(fh)

    src_path, out_path = spec["source"]["file"], spec["output"]["file"]
    if not os.path.exists(src_path):
        raise SystemExit(f"源文件不存在：{src_path}")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    if os.path.abspath(src_path) != os.path.abspath(out_path):
        shutil.copyfile(src_path, out_path)

    wb = openpyxl.load_workbook(out_path)
    if getattr(wb, "calculation", None) is None:
        wb.calculation = openpyxl.workbook.properties.CalcProperties()
    wb.calculation.fullCalcOnLoad = True

    for s in spec["sheets"]:
        sheet_spec = dict(s)
        if "source" not in sheet_spec and "source" in spec:
            sheet_spec["source"] = spec["source"]
        layout = sheet_spec.get("layout", "group")
        if layout not in LAYOUTS:
            raise SystemExit(f"未知 layout：{layout}（可选 {list(LAYOUTS)}）")
        if sheet_spec["name"] in wb.sheetnames:
            del wb[sheet_spec["name"]]
        LAYOUTS[layout](wb, sheet_spec)
        print(f"[ok] 新增 sheet：{sheet_spec['name']}（{layout}）")

    wb.save(out_path)
    print("已保存：", out_path)
    print("提示：用 Excel/WPS 打开会自动计算；如显示 0，按 Ctrl+Alt+F9 强制重算。")


if __name__ == "__main__":
    main()
