#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""读表体检：列出 sheet、字段、类型、空值、重复、疑似关联键，供统计需求挖掘使用。

用法:
    python3 inspect_workbook.py 数据表.xlsx
    python3 inspect_workbook.py 数据表.xlsx --json      # 机器可读
    python3 inspect_workbook.py 数据表.xlsx --rows 3000 # 采样行数
"""
import argparse
import json
import os
import sys
from collections import Counter
from datetime import date, datetime

import openpyxl


def infer_type(values):
    kinds = set()
    for v in values:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        if isinstance(v, bool):
            kinds.add("bool")
        elif isinstance(v, (datetime, date)):
            kinds.add("date")
        elif isinstance(v, (int, float)):
            kinds.add("number")
        else:
            s = str(v).strip()
            # 纯数字文本 / 日期文本单独识别，避免把文本数字当数值
            try:
                float(s.replace(",", "").replace("，", ""))
                kinds.add("number_text")
                continue
            except ValueError:
                pass
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y年%m月%d日", "%Y-%m", "%m/%d/%Y"):
                try:
                    datetime.strptime(s, fmt)
                    kinds.add("date_text")
                    break
                except ValueError:
                    continue
            else:
                kinds.add("text")
    if not kinds:
        return "空列"
    if len(kinds) > 1:
        return "混合(" + "/".join(sorted(kinds)) + ")"
    return {"bool": "布尔", "date": "日期", "number": "数值", "text": "文本",
            "number_text": "数字文本", "date_text": "日期文本"}[kinds.pop()]


def find_header_row(ws, scan=10):
    best_row, best_score = 1, -1
    for r in range(1, min(scan, ws.max_row or 1) + 1):
        vals = [c.value for c in ws[r]]
        filled = [v for v in vals if v is not None and str(v).strip() != ""]
        texty = [v for v in filled if isinstance(v, str)]
        score = len(filled) + len(texty)
        if len(filled) >= 2 and score > best_score:
            best_row, best_score = r, score
    return best_row


def inspect(path, sample_rows=2000):
    wb = openpyxl.load_workbook(path, data_only=True)
    report = {"file": os.path.abspath(path), "sheets": []}
    for ws in wb.worksheets:
        hr = find_header_row(ws)
        header_row_cells = list(ws[hr])
        headers = []
        for i, c in enumerate(header_row_cells, start=1):
            name = str(c.value).strip() if c.value is not None else f"__第{i}列(无表头)__"
            headers.append((i, name))
        col_data = {i: [] for i, _ in headers}
        dup_rows = Counter()
        tail_total = False
        last_data_row = hr
        for r in range(hr + 1, min(ws.max_row, hr + sample_rows) + 1):
            row_vals = [ws.cell(row=r, column=i).value for i, _ in headers]
            if all(v is None or (isinstance(v, str) and not v.strip()) for v in row_vals):
                continue
            first = row_vals[0]
            if isinstance(first, str) and first.strip() in ("合计", "总计", "小计", "汇总"):
                tail_total = True
                continue
            last_data_row = r
            key = tuple("" if v is None else str(v).strip() for v in row_vals)
            dup_rows[key] += 1
            for i, name in headers:
                col_data[i].append(ws.cell(row=r, column=i).value)

        cols = []
        for i, name in headers:
            vals = col_data[i]
            nonblank = [v for v in vals if v is not None and str(v).strip() != ""]
            t = infer_type(vals)
            stat = {"col": openpyxl.utils.get_column_letter(i), "name": name, "type": t,
                    "nonblank": len(nonblank), "blank": len(vals) - len(nonblank)}
            if t in ("数值",) or t.startswith("混合"):
                nums = [v for v in nonblank if isinstance(v, (int, float)) and not isinstance(v, bool)]
                if nums:
                    stat["min"], stat["max"] = min(nums), max(nums)
                    stat["sum"] = round(sum(nums), 2)
            distinct = len({("" if v is None else str(v).strip()) for v in vals})
            stat["distinct"] = distinct
            if distinct <= 12 and len(vals):
                top = Counter("" if v is None else str(v).strip() for v in vals).most_common(8)
                stat["sample_values"] = [k for k, _ in top]
            if distinct == len(vals) and len(nonblank) == len(vals) and len(vals) > 0:
                stat["unique_candidate"] = True
            if len(nonblank) == 0:
                stat["empty"] = True
            cols.append(stat)

        report["sheets"].append({
            "sheet": ws.title,
            "header_row": hr,
            "data_rows_sampled": max(0, last_data_row - hr),
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "has_total_row": tail_total,
            "duplicate_row_count": sum(n - 1 for n in dup_rows.values() if n > 1),
            "columns": cols,
        })
    return report


def render(report):
    out = [f"文件：{report['file']}", f"工作表数量：{len(report['sheets'])}"]
    if report.get("note"):
        out.insert(1, "提示：" + report["note"])
    for s in report["sheets"]:
        out.append("")
        out.append(f"■ {s['sheet']}（表头第 {s['header_row']} 行，采样 {s['data_rows_sampled']} 行数据，"
                   f"{len(s['columns'])} 个字段）")
        if s["has_total_row"]:
            out.append("  ! 表底部疑似有“合计/总计”行——统计时须限定数据区，不能整列求和。")
        if s["duplicate_row_count"]:
            out.append(f"  · 完全重复行：{s['duplicate_row_count']} 行")
        for c in s["columns"]:
            bits = [f"{c['col']}列 {c['name']}", c["type"], f"非空 {c['nonblank']}"]
            if c.get("blank"):
                bits.append(f"空 {c['blank']}")
            if "min" in c:
                bits.append(f"范围 {c['min']}~{c['max']}")
            bits.append(f"不同值 {c['distinct']}")
            if c.get("unique_candidate"):
                bits.append("★可当关联键")
            if c.get("sample_values"):
                bits.append("取值：" + "、".join(c["sample_values"]))
            if c.get("empty"):
                bits.append("空列，建议忽略")
            out.append("  - " + " | ".join(bits))
    return "\n".join(out)


def csv_to_temp_xlsx(path):
    """CSV 先转成临时 xlsx 再体检；数字列尝试转成数值。"""
    import csv
    import tempfile
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = os.path.splitext(os.path.basename(path))[0][:31] or "Sheet1"
    with open(path, encoding="utf-8-sig", newline="") as fh:
        for r, row in enumerate(csv.reader(fh), start=1):
            for c, v in enumerate(row, start=1):
                s = v.strip()
                if s == "":
                    continue
                try:
                    num = float(s.replace(",", ""))
                    ws.cell(row=r, column=c, value=int(num) if num.is_integer() and "." not in s else num)
                except ValueError:
                    ws.cell(row=r, column=c, value=v)
    tmp = os.path.join(tempfile.mkdtemp(), "csv_inspect.xlsx")
    wb.save(tmp)
    return tmp


def main():
    ap = argparse.ArgumentParser(description="读表体检，输出字段/类型/关联键线索")
    ap.add_argument("path")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--rows", type=int, default=2000)
    args = ap.parse_args()
    if not os.path.exists(args.path):
        sys.exit(f"文件不存在：{args.path}")
    target = args.path
    if args.path.lower().endswith(".csv"):
        target = csv_to_temp_xlsx(args.path)
    rep = inspect(target, args.rows)
    rep["file"] = os.path.abspath(args.path)
    rep["note"] = "CSV 已转临时工作簿体检，正式统计前请先转成 xlsx。" if target != args.path else ""
    print(json.dumps(rep, ensure_ascii=False, indent=2, default=str) if args.json else render(rep))


if __name__ == "__main__":
    main()
