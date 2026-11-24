#!/usr/bin/env python3
"""资深数据分析师·数据整理 —— 表格修复辅助脚本

只做四类确定性操作：复制副本、拆合并并向下填充、标黄加批注、交付前自检。
业务判断（单位口径、去重主键、是否拆列）仍由调用方问用户确认后自行处理。

用法：
  python3 sheet_fix.py clone <src.xlsx> <dst.xlsx>
  python3 sheet_fix.py inspect <file.xlsx> [--rows 8] [--sheet Sheet1]
  python3 sheet_fix.py unmerge-fill <file.xlsx> --range A2:A10 [--sheet Sheet1]
  python3 sheet_fix.py mark <file.xlsx> --cells D7,E8 --note "原值：约1200；原因：含估算词；待您确认：金额" [--sheet Sheet1]
  python3 sheet_fix.py audit <file.xlsx>
"""

import argparse
import shutil
import sys
from pathlib import Path

YELLOW = "FFFFFF00"


def _rgb(color):
    if color is None:
        return ""
    rgb = getattr(color, "rgb", None)
    return rgb if isinstance(rgb, str) else ""


def _is_yellow(cell):
    return _rgb(cell.fill.fgColor) in ("FFFFFF00", "00FFFF00", "FFFF00") and cell.fill.fill_type == "solid"


def _load(path, read_only=False):
    from openpyxl import load_workbook

    keep_vba = str(path).lower().endswith(".xlsm")
    return load_workbook(path, keep_vba=keep_vba, read_only=read_only)


def _save(wb, path):
    wb.save(path)


def _sheet(wb, name):
    if name:
        if name not in wb.sheetnames:
            sys.exit(f"错误：找不到工作表 {name}，现有：{wb.sheetnames}")
        return wb[name]
    return wb[wb.sheetnames[0]]


def cmd_clone(args):
    src, dst = Path(args.src), Path(args.dst)
    if not src.exists():
        sys.exit(f"错误：源文件不存在 {src}")
    if dst.exists():
        print(f"提示：目标已存在，覆盖 {dst}")
    shutil.copy2(src, dst)
    print(f"已复制：{src} -> {dst}")


def cmd_inspect(args):
    wb = _load(args.file)
    print(f"文件：{args.file}")
    print(f"工作表：{wb.sheetnames}")
    for ws in wb.worksheets:
        print(f"\n[{ws.title}] 维度 {ws.max_row} 行 x {ws.max_column} 列")
        merged = [str(r) for r in ws.merged_cells.ranges]
        print(f"  合并区域（{len(merged)}）：{', '.join(merged[:30]) or '无'}")
        for row in ws.iter_rows(min_row=1, max_row=min(args.rows, ws.max_row), values_only=True):
            cells = ["" if v is None else str(v)[:28] for v in row]
            print("  | " + " | ".join(cells))
    wb.close()


def cmd_unmerge_fill(args):
    wb = _load(args.file)
    ws = _sheet(wb, args.sheet)
    rng_list = [r.strip() for r in args.range.split(",") if r.strip()]
    if not rng_list:
        sys.exit("错误：--range 不能为空，例如 A2:A10 或 A2:A10,C2:C8")
    total = 0
    for rng in rng_list:
        from openpyxl.utils.cell import range_boundaries

        try:
            min_col, min_row, max_col, max_row = range_boundaries(rng)
        except ValueError:
            sys.exit(f"错误：无法解析区域 {rng}")
        targets = [m for m in list(ws.merged_cells.ranges)
                   if m.min_row == min_row and m.max_row == max_row
                   and m.min_col == min_col and m.max_col == max_col]
        if not targets:
            print(f"跳过 {rng}：该区域不是合并单元格（检查项1 无需处理）")
            continue
        top = ws.cell(row=min_row, column=min_col).value
        for m in targets:
            ws.unmerge_cells(str(m))
        if max_col - min_col == 0:  # 单列合并：向下填充
            for row in range(min_row, max_row + 1):
                c = ws.cell(row=row, column=min_col)
                if c.value in (None, ""):
                    c.value = top
                    total += 1
            print(f"{rng}：已拆分并向下填充，值={top!r}")
        else:
            print(f"{rng}：跨列合并（疑似标题行），已拆分但未填充，请按检查项9 人工处理")
    wb.save(args.file)
    print(f"已保存 {args.file}，共填充 {total} 个空单元格")


def cmd_mark(args):
    wb = _load(args.file)
    ws = _sheet(wb, args.sheet)
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill

    cells = [c.strip() for c in args.cells.split(",") if c.strip()]
    if not cells:
        sys.exit("错误：--cells 不能为空，例如 D7,E8 或 D7:D9")
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    marked = 0
    for spec in cells:
        target = ws[spec]
        if not isinstance(target, tuple):
            target = ((target,),)
        for row in target:
            for cell in row:
                cell.fill = fill
                cell.comment = Comment(f"【资深数据分析师·数据整理】{args.note}", "资深数据分析师·数据整理")
                marked += 1
    wb.save(args.file)
    print(f"已标黄加批注 {marked} 个单元格：{args.cells}")


def cmd_audit(args):
    wb = _load(args.file, read_only=False)
    total_yellow = 0
    missing_comment = []
    for ws in wb.worksheets:
        count = 0
        for row in ws.iter_rows():
            for cell in row:
                if _is_yellow(cell):
                    count += 1
                    total_yellow += 1
                    if not cell.comment or not cell.comment.text.strip():
                        missing_comment.append(f"{ws.title}!{cell.coordinate}")
        print(f"[{ws.title}] 标黄单元格：{count}")
    if missing_comment:
        print(f"缺少批注的标黄单元格（{len(missing_comment)}）：{', '.join(missing_comment[:50])}")
    print(f"合计标黄：{total_yellow}，待人工确认项应与此数一致（或按行/区域折算）")
    wb.close()


def main():
    parser = argparse.ArgumentParser(description="资深数据分析师·数据整理辅助脚本")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("clone", help="复制原表为已整理副本")
    p.add_argument("src")
    p.add_argument("dst")
    p.set_defaults(func=cmd_clone)

    p = sub.add_parser("inspect", help="查看工作表结构与前若干行")
    p.add_argument("file")
    p.add_argument("--rows", type=int, default=8)
    p.add_argument("--sheet", default=None)
    p.set_defaults(func=cmd_inspect)

    p = sub.add_parser("unmerge-fill", help="拆分合并单元格并向下填充")
    p.add_argument("file")
    p.add_argument("--range", required=True)
    p.add_argument("--sheet", default=None)
    p.set_defaults(func=cmd_unmerge_fill)

    p = sub.add_parser("mark", help="标黄并加批注")
    p.add_argument("file")
    p.add_argument("--cells", required=True)
    p.add_argument("--note", required=True)
    p.add_argument("--sheet", default=None)
    p.set_defaults(func=cmd_mark)

    p = sub.add_parser("audit", help="自检：统计标黄单元格与缺失批注")
    p.add_argument("file")
    p.set_defaults(func=cmd_audit)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
