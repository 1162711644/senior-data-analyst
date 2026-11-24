#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""资深数据分析师·数据体检 —— 自动体检 + 生成标注副本 + 生成体检报告。

用法：
    python3 table_check.py <表格1.xlsx> [表格2.csv ...] [--outdir 输出目录]

产出：
    <原文件名>_体检标注.xlsx   问题单元格填色 + 批注
    <原文件名>_体检报告.md     按检查项汇总的体检报告
    多文件时额外产出：跨表体检总览.md

原文件只读，绝不修改。
"""

import argparse
import csv
import datetime
import difflib
import os
import re
import sys
from collections import Counter, defaultdict

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 检查项元数据
# 颜色与批注口径来自技能规范，不要随意改动
CHECK_META = {
    1: {"name": "合并单元格/分类列只在首行有值", "level": "高", "color": "FFC000", "dark": False},
    2: {"name": "单位混着写", "level": "中", "color": "BF8F00", "dark": True},
    3: {"name": "日期格式乱", "level": "高", "color": "BDD7EE", "dark": False},
    4: {"name": "数字夹字", "level": "高", "color": "FFC7CE", "dark": False},
    5: {"name": "一格装多信息", "level": "中", "color": "D9C2E9", "dark": False},
    6: {"name": "空格子", "level": "中", "color": "D9D9D9", "dark": False},
    7: {"name": "多余空行空列", "level": "中", "color": "808080", "dark": True},
    8: {"name": "重复记录", "level": "高", "color": "FF0000", "dark": True},
    9: {"name": "表头问题", "level": "中", "color": "A97142", "dark": True},
    10: {"name": "跨表命名不统一", "level": "低", "color": "7FD6D6", "dark": True},
}

LEVEL_RANK = {"高": 3, "中": 2, "低": 1}
# 同等级冲突时的批注优先级
CHECK_ORDER = [8, 1, 3, 4, 6, 7, 9, 5, 2, 10]

RISK_TEXT = {1: "高", 2: "中", 3: "高", 4: "高", 5: "中", 6: "中", 7: "中", 8: "高", 9: "中", 10: "低"}

HURT = {
    1: "按分类字段汇总、透视时，空白组会被整组漏掉或错归到上一组",
    2: "同列口径不一，求和会算错、甚至直接跳过这些单元格",
    3: "无法排序、无法按月份/季度统计，图表时间轴会断",
    4: "公式算不了、透视识别成文本，求和结果悄悄变少",
    5: "无法单独筛选或拆分统计，一条记录当两条用",
    6: "该计入的数值缺一块，汇总、均值、占比全部失真",
    7: "统计范围被撑大，行数/列数虚高，透视出现空项",
    8: "重复行会让所有统计虚高（金额、单量、人数翻倍）",
    9: "合并大标题和含糊/重名列会让程序取错列，后续自动化全部错位",
    10: "合并多表时字段对不上，出现两列同类数据或大量空值",
}

FIX = {
    1: "拆开合并单元格，每组每一行都补上对应值（或保留合并但导出时向下填充）",
    2: "统一单位后只保留数字，单位写进列名（如「金额（元）」）",
    3: "统一为同一写法（推荐 YYYY-MM-DD 或真正的日期型），再批量转换",
    4: "去掉约/左右/暂定/未定等描述与货币符号，只留纯数字，备注另开一列",
    5: "按语义拆成多列（姓名 / 电话 / 地址），一列只放一种信息",
    6: "逐格补齐或确认空值含义（未填/不适用），不要留空参与统计",
    7: "删除整行整列全空的行列，合计行、说明行移出数据区",
    8: "按主键去重，保留最新一条，并在报告里记录被删条数",
    9: "首行只放列名，大标题删掉或移到表外；含糊列名改成明确业务名；重名列改名",
    10: "建立字段字典，多表统一同一个字段名（如统一「客户名称」）",
}

VAGUE_HEADERS = {
    "数据", "信息", "值", "内容", "备注", "其他", "详情", "列1", "column1",
    "字段", "数值", "项目", "名称", "说明", "data", "info", "value", "remark",
}

PRIMARY_KEY_RE = re.compile(
    r"(订单号|订单编号|工号|员工号|编号|单号|流水号|序号|id|ID|Id|身份证|证件号|手机号|电话|合同号|发票号)"
)

UNIT_RE = re.compile(
    r"(\d[\d,\.]*)\s*(公斤|千克|kg|KG|斤|克|g|G|件|个|台|套|箱|包|瓶|吨|米|m|平方米|平米|%|％|小时|天)"
)
CURRENCY_RE = re.compile(r"[￥¥]|(\d[\d,\.]*)\s*元")
FUZZY_RE = re.compile(r"(大约|约|左右|暂定|未定|待定|待确认|不确定|估计|大概|待补|TBD|tbd)")
PHONE_RE = re.compile(r"1[3-9]\d{9}")
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
IDCARD_RE = re.compile(r"\d{17}[\dXx]")

DATE_PATTERNS = [
    ("YYYY-MM-DD", re.compile(r"^\d{4}-\d{1,2}-\d{1,2}$")),
    ("YYYY/MM/DD", re.compile(r"^\d{4}/\d{1,2}/\d{1,2}$")),
    ("YYYY.MM.DD", re.compile(r"^\d{4}\.\d{1,2}\.\d{1,2}$")),
    ("YYYY年M月D日", re.compile(r"^\d{4}年\d{1,2}月\d{1,2}日?$")),
    ("M月D日", re.compile(r"^\d{1,2}月\d{1,2}日$")),
    ("MM/DD", re.compile(r"^\d{1,2}/\d{1,2}$")),
    ("MM-DD", re.compile(r"^\d{1,2}-\d{1,2}$")),
    ("YYYYMMDD", re.compile(r"^(19|20)\d{6}$")),
    ("MM/DD/YYYY", re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")),
]
DATE_HINT_RE = re.compile(r"(日期|时间|年月|月份|date|day)", re.I)


def is_blank(v):
    return v is None or (isinstance(v, str) and v.strip() == "")


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def date_pattern_of(v):
    """返回该单元格的日期写法标签；不是日期返回 None。"""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return "Excel日期型"
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s:
        return None
    for label, pat in DATE_PATTERNS:
        if pat.match(s):
            return label
    return None


def load_workbook_any(path):
    """xlsx/xlsm 直接读取；csv 转成内存工作簿。返回 (workbook, 是否由csv转换)。"""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        return openpyxl.load_workbook(path, data_only=False), False
    if ext == ".csv":
        rows = None
        for enc in ("utf-8-sig", "utf-8", "gbk", "gb18030"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    rows = list(csv.reader(f))
                break
            except (UnicodeDecodeError, LookupError):
                continue
        if rows is None:
            raise RuntimeError("CSV 编码无法识别：" + path)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        for r in rows:
            ws.append([None if c == "" else c for c in r])
        return wb, True
    raise RuntimeError("仅支持 .xlsx/.xlsm/.csv，遇到：" + path)


class SheetScan:
    """单个 sheet 的扫描结果。"""

    def __init__(self, sheet_name):
        self.sheet_name = sheet_name
        self.header_row = 1
        self.header_names = {}
        self.data_start = 2
        self.max_row = 0
        self.max_col = 0
        self.issues = []            # 每条：dict(check, cell(s), msg, detail, row, col)
        self.item_notes = defaultdict(list)   # 检查项 -> 说明文字（结构性说明）
        self.sub_checks = []        # 建议人工复核项描述

    def add(self, check, cell, row, col, where, why, fix, extra=""):
        self.issues.append(dict(
            check=check, cell=cell, row=row, col=col,
            level=CHECK_META[check]["level"],
            where=where, why=why, fix=fix, extra=extra,
        ))


def detect_header(ws):
    """在前 8 行里找表头行：非空且不像纯数字的单元格最多的那一行。"""
    best_row, best_score = 1, -1
    limit = min(ws.max_row or 1, 8)
    for r in range(1, limit + 1):
        cells = list(ws.iter_rows(min_row=r, max_row=r))
        vals = [c.value for c in cells[0]] if cells else []
        score = 0
        for v in vals:
            if is_blank(v):
                continue
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                continue
            if isinstance(v, (datetime.date, datetime.datetime)):
                continue
            score += 1
        if score > best_score:
            best_row, best_score = r, score
    return best_row if best_score > 0 else 1


def scan_sheet(ws, sheet_name):
    scan = SheetScan(sheet_name)
    scan.max_row = ws.max_row or 0
    scan.max_col = ws.max_column or 0
    header_row = detect_header(ws)
    scan.header_row = header_row
    scan.data_start = header_row + 1

    # ---------- 表头名称（合并单元格向右传播） ----------
    merged = list(ws.merged_cells.ranges)
    big_title_rows = set()
    for mr in merged:
        if mr.max_row < header_row and mr.max_col > mr.min_col:
            big_title_rows.add(mr.min_row)
            scan.add(
                9, mr.coord, mr.min_row, mr.min_col,
                where=f"表头上方第 {mr.min_row} 行有横跨 {mr.max_col - mr.min_col + 1} 列的大标题（{mr.coord}）",
                why="程序默认首行为表头，大标题会被当成列名，整表列定位错位",
                fix=FIX[9], extra="建议人工复核",
            )
            continue
        if mr.min_row <= header_row <= mr.max_row and mr.max_col > mr.min_col:
            scan.add(
                9, f"{mr.coord}", mr.min_row, mr.min_col,
                where=f"表头区域存在合并单元格 {mr.coord}（跨 {mr.max_col - mr.min_col + 1} 列）",
                why="合并大标题/合并列名会让程序只认到第一列，其余列变成空列名，取列必错",
                fix=FIX[9], extra="建议人工复核",
            )
    names = {}
    for c in range(1, scan.max_col + 1):
        v = ws.cell(row=header_row, column=c).value
        names[c] = cell_text(v)
    # 合并区域覆盖的表头列，沿用左上角名字
    for mr in merged:
        if mr.min_row <= header_row <= mr.max_row:
            left_val = cell_text(ws.cell(row=mr.min_row, column=mr.min_col).value)
            for c in range(mr.min_col, mr.max_col + 1):
                if not names.get(c):
                    names[c] = left_val
    scan.header_names = names

    # 表头含糊 / 重名
    norm = {}
    for c, name in names.items():
        if not name:
            continue
        key = re.sub(r"\s+", "", name).lower()
        norm.setdefault(key, []).append(c)
    for key, cols in norm.items():
        if len(cols) > 1:
            for c in cols:
                scan.add(
                    9, f"{get_column_letter(c)}{header_row}", header_row, c,
                    where=f"列名「{names[c]}」重复出现（{', '.join(get_column_letter(x) for x in cols)} 列）",
                    why="重名列会让筛选、透视、公式引用指向错误的列",
                    fix=FIX[9], extra="建议人工复核",
                )
    for c, name in names.items():
        if name and re.sub(r"\s+", "", name).lower() in VAGUE_HEADERS:
            scan.add(
                9, f"{get_column_letter(c)}{header_row}", header_row, c,
                where=f"列名「{name}」含义含糊，看不出放的是什么",
                why="列名不明确，后续接手的人和分析程序都不知道该拿它算什么",
                fix=FIX[9], extra="建议人工复核",
            )
    # 表头上方还有整行大标题
    for r in range(1, header_row):
        vals = [ws.cell(row=r, column=c).value for c in range(1, scan.max_col + 1)]
        non_blank = sum(1 for v in vals if not is_blank(v))
        if non_blank == 1 and scan.max_col > 2 and r not in big_title_rows:
            scan.add(
                9, f"A{r}", r, 1,
                where=f"第 {r} 行是横跨多列的大标题（真实表头在第 {header_row} 行）",
                why="程序默认首行为表头，会把标题当列名，整张表列定位错位",
                fix=FIX[9], extra="建议人工复核",
            )

    # ---------- 空行 / 空列 ----------
    last_data_row = scan.data_start - 1
    last_data_col = 0
    for r in range(scan.data_start, scan.max_row + 1):
        for c in range(1, scan.max_col + 1):
            if not is_blank(ws.cell(row=r, column=c).value):
                last_data_row = max(last_data_row, r)
                last_data_col = max(last_data_col, c)
    for r in range(scan.data_start, scan.max_row + 1):
        if all(is_blank(ws.cell(row=r, column=c).value) for c in range(1, (last_data_col or 1) + 1)):
            scan.add(
                7, f"A{r}", r, 1,
                where=f"第 {r} 行整行为空（表内空行或尾部空行）",
                why="空行会把统计范围撑大，透视出现空项、行数虚高",
                fix=FIX[7],
            )
    for c in range(1, scan.max_col + 1):
        if all(is_blank(ws.cell(row=r, column=c).value) for r in range(scan.data_start, last_data_row + 1)):
            scan.add(
                7, f"{get_column_letter(c)}{scan.data_start or 1}", scan.data_start, c,
                where=f"{get_column_letter(c)} 列从数据区开始整列为空",
                why="空列参与统计会多出无数空值列，透视和图表选中区域被撑大",
                fix=FIX[7],
            )
    scan.last_data_row = last_data_row
    scan.last_data_col = last_data_col or scan.max_col

    # ---------- 汇总行 / 说明行标记（这类行的空格不算检查项6） ----------
    summary_rows = set()
    for r in range(scan.data_start, last_data_row + 1):
        head = " ".join(cell_text(ws.cell(row=r, column=cc).value)
                        for cc in range(1, min(3, (last_data_col or scan.max_col)) + 1))
        filled_in_row = sum(1 for cc in range(1, (last_data_col or scan.max_col) + 1)
                            if not is_blank(ws.cell(row=r, column=cc).value))
        if re.search(r"(合计|小计|总计|累计|说明|注：|备注：)", head):
            summary_rows.add(r)
        if filled_in_row <= 2:
            summary_rows.add(r)
            if filled_in_row > 0:
                scan.add(
                    7, f"A{r}", r, 1,
                    where=f"第 {r} 行只有 {filled_in_row} 个单元格有值，疑似游离的合计/说明行",
                    why="游离行混在数据区里，会被统计成一条业务记录",
                    fix=FIX[7], extra="建议人工复核",
                )

    # ---------- 逐列扫描 ----------
    # 合并区域内部（非锚点）单元格实际继承了锚点值，不能当成空格子
    merged_cfg = {}
    for mr in merged:
        anchor = ws.cell(row=mr.min_row, column=mr.min_col).value
        for rr in range(mr.min_row, mr.max_row + 1):
            for cc in range(mr.min_col, mr.max_col + 1):
                if (rr, cc) != (mr.min_row, mr.min_col):
                    merged_cfg[(rr, cc)] = anchor
    for c in range(1, (last_data_col or scan.max_col) + 1):
        col_letter = get_column_letter(c)
        name = names.get(c, "")
        rows = list(range(scan.data_start, last_data_row + 1))
        values = []
        for r in rows:
            v = ws.cell(row=r, column=c).value
            if is_blank(v) and (r, c) in merged_cfg:
                v = merged_cfg[(r, c)]     # 合并区域内部按锚点值算
            values.append((r, v))
        filled = [(r, v) for r, v in values if not is_blank(v)]
        blanks = [r for r, v in values if is_blank(v)]
        if not values:
            continue

        # 检查项1：合并单元格 + 分类列悬空
        for mr in merged:
            if mr.min_col <= c <= mr.max_col and mr.max_row >= scan.data_start:
                scan.add(
                    1, mr.coord, mr.min_row, mr.min_col,
                    where=f"{col_letter} 列存在合并单元格 {mr.coord}（{mr.max_row - mr.min_row + 1} 行合成一格）",
                    why=HURT[1],
                    fix=FIX[1],
                )
        if blanks and len(filled) >= 2:
            # 分类列特征：非空值重复出现，中间夹着连续空档
            texts = [cell_text(v) for _, v in filled]
            uniq_ratio = len(set(texts)) / max(len(texts), 1)
            gaps = []
            prev, prev_text = None, ""
            values_data = [(r, v) for r, v in values if r not in summary_rows]
            for r, v in values_data:
                if not is_blank(v):
                    if prev is not None:
                        gaps.append((prev + 1, r - 1, prev_text))
                    prev, prev_text = r, cell_text(v)
            if prev is not None and values_data and values_data[-1][0] > prev:
                gaps.append((prev + 1, values_data[-1][0], prev_text))   # 末尾连续空档也算一组
            gap_blanks = sum(1 for a, b, _ in gaps for r in range(a, b + 1) if r in blanks)
            gap_cover = gap_blanks / max(len(blanks), 1)
            # 只有纯文本的分类字段才按「分类列悬空」判定，数值列不适用
            numeric_col = any(isinstance(v, (int, float)) and not isinstance(v, bool)
                              or (isinstance(v, str) and re.fullmatch(r"-?\d+(?:\.\d+)?", v.strip()))
                              for _, v in filled)
            looks_categorical = (not numeric_col) and 2 <= len(set(texts)) <= 6 \
                and (uniq_ratio <= 0.75 or len(filled) <= 3)
            if gaps and len(gaps) >= 2 and gap_cover >= 0.7 and looks_categorical:
                for a, b, up in gaps:
                    for r in range(a, b + 1):
                        scan.add(
                            1, f"{col_letter}{r}", r, c,
                            where=f"{col_letter} 列（列名「{name}」）第 {r} 行为空，上方是分类值「{up}」",
                            why=HURT[1],
                            fix=FIX[1], extra="建议人工复核（该列并非合并单元格，按「分类列悬空」模式判定）",
                        )
                # 这些格子不再算检查项6
                blanks = [r for r in blanks
                          if not any(a <= r <= b for a, b, _ in gaps)]

        # 检查项6：零散空格子（汇总行、说明行的空不算）
        blanks = [r for r in blanks if r not in summary_rows]
        if blanks and len(filled) >= 3:
            blank_ratio = len(blanks) / len(values)
            if 0 < blank_ratio <= 0.35:
                for r in blanks:
                    scan.add(
                        6, f"{col_letter}{r}", r, c,
                        where=f"{col_letter} 列（列名「{name}」）第 {r} 行为空，同列其他行大多有值",
                        why=HURT[6],
                        fix=FIX[6],
                    )

        # 检查项3：日期写法
        pats = {}
        for r, v in filled:
            p = date_pattern_of(v)
            if p:
                pats[r] = p
        if len(pats) >= 2:
            counter = Counter(pats.values())
            main_pat, _ = counter.most_common(1)[0]
            if len(counter) >= 2:
                for r, p in pats.items():
                    if p != main_pat:
                        scan.add(
                            3, f"{col_letter}{r}", r, c,
                            where=f"日期写法是「{p}」，同列主流写法是「{main_pat}」",
                            why=HURT[3],
                            fix=FIX[3],
                        )

        # 检查项4 / 2：数字夹字、单位混写
        num_like = 0
        for r, v in filled:
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                num_like += 1
            elif isinstance(v, str):
                s = v.strip()
                if re.fullmatch(r"-?\d+(\.\d+)?", s):
                    num_like += 1
        unit_map, unit_rows = {}, {}
        for r, v in filled:
            if not isinstance(v, str):
                continue
            s = v.strip()
            # 检查项4：货币符号 / 约左右
            if CURRENCY_RE.search(s) or (FUZZY_RE.search(s) and re.search(r"\d", s)):
                reason = "带货币符号（元/￥）" if CURRENCY_RE.search(s) else "数字里夹了约/左右/暂定/未定等描述词"
                scan.add(
                    4, f"{col_letter}{r}", r, c,
                    where=f"「{s}」{reason}",
                    why=HURT[4],
                    fix=FIX[4],
                )
                continue
            # 检查项2：其它单位
            m = UNIT_RE.search(s)
            if m:
                unit = m.group(2)
                unit_map.setdefault(unit, []).append(r)
                unit_rows[r] = unit
        if unit_map:
            kinds = len(unit_map)
            has_bare = num_like > 0
            if kinds >= 2 or (kinds == 1 and has_bare):
                if kinds >= 2:
                    main_unit = Counter({u: len(rs) for u, rs in unit_map.items()}).most_common(1)[0][0]
                    minority = [r for r, u in unit_rows.items() if u != main_unit]
                    desc = f"同列出现多种单位：{'、'.join(f'{u}({len(rs)}处)' for u, rs in unit_map.items())}"
                else:
                    minority = list(unit_rows.keys())
                    desc = "同列有的单元格带单位、有的只有数字"
                for r in minority:
                    scan.add(
                        2, f"{col_letter}{r}", r, c,
                        where=f"「{cell_text(values[[x[0] for x in values].index(r)][1])}」的单位写法与全列不一致（{desc}）",
                        why=HURT[2],
                        fix=FIX[2], extra="建议人工复核",
                    )
                scan.sub_checks.append(
                    f"{scan.sheet_name} · {col_letter} 列（{name}）：{desc}，请确认是否同一单位口径，确认后统一改写。"
                )

        # 检查项5：一格多信息
        for r, v in filled:
            if not isinstance(v, str):
                continue
            s = v.strip()
            if len(s) < 8:
                continue
            signals = []
            if PHONE_RE.search(s):
                signals.append("手机号")
            if EMAIL_RE.search(s):
                signals.append("邮箱")
            if IDCARD_RE.search(s):
                signals.append("证件号")
            seps = len(re.findall(r"[，,;；、/|]", s))
            hint_col = re.search(r"(姓名|电话|手机|地址|联系|客户信息|联系人|收件|收货|人员)", name)
            # 列名本身就暗示「一格装多种信息」时，只出现电话/邮箱就可判定
            if signals and (len(re.findall(r"[^\d\s]", s)) >= 6 or seps >= 1 or (hint_col and len(s) >= 10)):
                scan.add(
                    5, f"{col_letter}{r}", r, c,
                    where=f"「{s[:24]}{'…' if len(s) > 24 else ''}」一个格子里同时装有{'、'.join(signals)}和文字信息",
                    why=HURT[5],
                    fix=FIX[5], extra="建议人工复核",
                )
            elif seps >= 2 and len(s) >= 15 and re.search(r"\d", s) and re.search(r"[\u4e00-\u9fa5]", s):
                scan.add(
                    5, f"{col_letter}{r}", r, c,
                    where=f"「{s[:24]}{'…' if len(s) > 24 else ''}」用分隔符塞了多段信息",
                    why=HURT[5],
                    fix=FIX[5], extra="建议人工复核",
                )

    # ---------- 检查项8：重复记录 ----------
    key_cols = [c for c in range(1, (last_data_col or scan.max_col) + 1)
                if PRIMARY_KEY_RE.search(names.get(c, ""))]
    row_sig = defaultdict(list)
    for r in range(scan.data_start, last_data_row + 1):
        sig = tuple(cell_text(ws.cell(row=r, column=c).value) for c in
                    range(1, (last_data_col or scan.max_col) + 1))
        if all(x == "" for x in sig):
            continue
        row_sig[sig].append(r)
    for sig, rs in row_sig.items():
        if len(rs) > 1:
            for r in rs:
                scan.add(
                    8, f"A{r}", r, 1,
                    where=f"第 {r} 行与第 {'、'.join(str(x) for x in rs if x != r)} 行内容完全重复",
                    why=HURT[8],
                    fix=FIX[8],
                )
    if key_cols:
        for c in key_cols:
            col_letter = get_column_letter(c)
            key_map = defaultdict(list)
            for r in range(scan.data_start, last_data_row + 1):
                k = cell_text(ws.cell(row=r, column=c).value)
                if k:
                    key_map[k].append(r)
            for k, rs in key_map.items():
                if len(rs) > 1:
                    for r in rs:
                        scan.add(
                            8, f"{col_letter}{r}", r, c,
                            where=f"主键「{names.get(c, col_letter)}」的值「{k}」在第 {'、'.join(str(x) for x in rs)} 行重复出现",
                            why=HURT[8],
                            fix=FIX[8],
                        )
    return scan


def normalize_field(name):
    return re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]", "", name).lower()


def cross_table_check(scans, file_label=""):
    """检查项10：多 sheet / 多文件同义字段命名是否统一。返回 issue 列表 + 提示文字。"""
    variants = defaultdict(set)   # 归一化后的主干 -> 写法集合
    location = defaultdict(list)  # 写法 -> [位置]
    for scan in scans:
        for c, name in scan.header_names.items():
            clean = re.sub(r"\s+", "", name)
            if not clean or len(clean) < 2:
                continue
            key = normalize_field(clean)
            if not key:
                continue
            variants[key].add(clean)
            location[clean].append(f"{scan.sheet_name}·{get_column_letter(c)}列")
    # 近似字段聚类（客户名称 / 客户名 / 客户名称(全称) 等）
    keys = list(variants.keys())
    groups = []
    used = set()
    for i, k1 in enumerate(keys):
        if k1 in used:
            continue
        grp = {k1}
        used.add(k1)
        for k2 in keys[i + 1:]:
            if k2 in used:
                continue
            if difflib.SequenceMatcher(None, k1, k2).ratio() >= 0.82:
                grp.add(k2)
                used.add(k2)
        if len(grp) > 1:
            groups.append(grp)
    issues, notes = [], []
    for grp in groups:
        all_names = sorted({n for k in grp for n in variants[k]})
        if len(all_names) < 2:
            continue
        counter = Counter()
        for n in all_names:
            counter[n] = sum(len(re.findall(r"·", loc)) or 1 for loc in location[n])
        main = counter.most_common(1)[0][0]
        desc = "、".join(f"「{n}」" for n in all_names)
        notes.append(f"{desc} 指向同一含义字段，建议统一为「{main}」。出现位置：" +
                     "；".join(f"{n}→{', '.join(location[n])}" for n in all_names))
        for n in all_names:
            if n == main:
                continue
            for loc in location[n]:
                issues.append(dict(level="低", where=f"{loc} 的列名写作「{n}」，其它表写作「{main}」",
                                   why=HURT[10], fix=FIX[10], extra="建议人工复核",
                                   loc=loc, name=n))
    return issues, notes


def annotate(ws, scan, extra_issues=None):
    """把问题写回副本：填充色 + 批注。"""
    bucket = defaultdict(list)

    def _targets(coord):
        """区域坐标（A3:A5）只取左上角锚点：合并区域上色时锚点会铺满整块。"""
        if ":" in coord:
            return [coord.split(":")[0]]
        try:
            cell = ws[coord]
        except Exception:
            return []
        if cell.__class__.__name__ == "MergedCell":
            return []   # 非锚点的合并格不可写样式/批注，跳过
        return [cell.coordinate]

    for it in scan.issues:
        targets = _targets(it["cell"])
        for coord in targets:
            bucket[coord].append(it)
    # 检查项10 的标注落在表头单元格
    if extra_issues:
        for it in extra_issues:
            loc = it.get("loc", "")
            m = re.match(r"(.+)·([A-Z]+)列", loc)
            if m and m.group(1) == scan.sheet_name:
                coord = f"{m.group(2)}{scan.header_row}"
                it = dict(it, cell=coord, row=scan.header_row, col=0, check=10)
                bucket[coord].append(it)

    for coord, items in bucket.items():
        # 只取单元格内最严重的一条做主批注，同格其它检查项追加提示
        def sort_key(it):
            return (-LEVEL_RANK[it["level"]], CHECK_ORDER.index(it.get("check", 10))
                    if it.get("check") in CHECK_ORDER else 99)
        items_sorted = sorted(items, key=sort_key)
        main = items_sorted[0]
        check = main.get("check", 10)
        meta = CHECK_META[check]
        fill = PatternFill("solid", fgColor=meta["color"])
        cell = ws[coord]
        cell.fill = fill
        if meta["dark"]:
            cell.font = Font(color="FFFFFF")
        lines = [f"【检查项{check}·{meta['name']}｜风险：{meta['level']}】"]
        for it in items_sorted:
            lines.append(f"· 问题：{it['where']}")
        lines.append(f"· 为何碍事：{main['why']}")
        lines.append(f"· 建议：{main['fix']}")
        others = [f"检查项{i.get('check')}" for i in items_sorted[1:] if i.get("check")]
        if others:
            lines.append("· 该格还涉及：" + "、".join(dict.fromkeys(others)))
        if any(i.get("extra") == "建议人工复核" for i in items_sorted):
            lines.append("· 判定方式：机器初筛，建议人工复核后再改")
        cm = Comment("\n".join(lines), "资深数据分析师·数据体检")
        cm.width = 380
        cm.height = max(120, 30 * len(lines))
        cell.comment = cm


def build_report(scan, filename, csv_note="", xlsx_name=""):
    """生成单表体检报告 markdown。"""
    by_check = defaultdict(list)
    for it in scan.issues:
        by_check[it["check"]].append(it)

    total_rows = max(0, getattr(scan, "last_data_row", 0) - scan.data_start + 1)
    total_cols = scan.last_data_col or scan.max_col
    lines = [f"# 数据体检报告 — {filename}", ""]
    if csv_note:
        lines += [f"> {csv_note}", ""]
    lines += ["## 一、总览", ""]
    lines.append(f"- 数据量：{total_rows} 行 × {total_cols} 列（sheet「{scan.sheet_name}」，表头在第 {scan.header_row} 行）")
    if by_check:
        summary = "、".join(f"检查项{k} ×{len(by_check[k])}" for k in sorted(by_check))
    else:
        summary = "未发现明显问题"
    lines.append(f"- 发现问题：{summary}")
    high = sum(1 for it in scan.issues if it["level"] == "高")
    mid = sum(1 for it in scan.issues if it["level"] == "中")
    low = sum(1 for it in scan.issues if it["level"] == "低")
    lines.append(f"- 风险分布：高 {high} 处 / 中 {mid} 处 / 低 {low} 处")
    lines.append("- 优先处理：" + priority_text(by_check))
    lines.append("")

    lines += ["## 二、分项明细", ""]
    for k in range(1, 11):
        if k not in by_check:
            continue
        items = by_check[k]
        meta = CHECK_META[k]
        positions = summarize_positions(items)
        where_sample = items[0]["where"]
        lines.append(f"### 检查项{k} {meta['name']}（风险：{meta['level']}，共 {len(items)} 处）")
        lines.append("")
        lines.append(f"- 位置：{positions}")
        lines.append(f"- 问题：{where_sample}" + ("" if len(items) == 1 else f"（示例，共 {len(items)} 处同类）"))
        lines.append(f"- 碍事：{HURT[k]}")
        lines.append(f"- 建议：{FIX[k]}")
        lines.append("")
    if not by_check:
        lines += ["本次扫描未命中任何检查项，表格结构较规范；仍建议抽查金额列与日期列口径。", ""]

    lines += ["## 三、建议人工复核项", ""]
    manual = list(dict.fromkeys(scan.sub_checks))
    manual_items = defaultdict(list)
    for it in scan.issues:
        if str(it.get("extra", "")).startswith("建议人工复核"):
            manual_items[it["check"]].append(it)
    for k in sorted(manual_items):
        items = manual_items[k]
        if len(items) <= 2:
            for it in items:
                coord = "" if it["cell"] in it["where"] else f"（{it['cell']}）"
                manual.append(f"{it['where']}{coord}")
        else:
            item_texts = list(dict.fromkeys(i["where"] for i in items))
            sample = item_texts[0]
            manual.append(
                f"检查项{k}·{CHECK_META[k]['name']}：共 {len(items)} 处同类情况，位置 {summarize_positions(items)}"
                f"；示例：{sample}；请人工确认后统一处理"
            )
    if manual:
        for m in manual[:20]:
            lines.append(f"- {m}")
        if len(manual) > 20:
            lines.append(f"- ……其余 {len(manual) - 20} 条见标注副本批注")
    else:
        lines.append("- 无")
    lines.append("")
    lines += ["## 四、下一步", "",
              f"- 标注副本：`{xlsx_name or filename + '_体检标注.xlsx'}`，颜色图例见文末；批注里写了「哪里不对 / 为何碍事 / 怎么改」。",
              "- 改完硬伤（检查项 8 / 1 / 3 / 4）之后再开始求和、透视、作图，否则数字会悄悄错。",
              "- 语义类问题（检查项 2 / 5 / 9 / 10）建议人工确认，不要机器硬改。", ""]
    lines += ["### 颜色图例", ""]
    lines += ["| 检查项 | 颜色 | 风险 |", "| --- | --- | --- |"]
    COLOR_CN = {1: "橙色", 2: "深黄", 3: "浅蓝", 4: "粉红", 5: "浅紫",
                6: "灰色", 7: "深灰", 8: "红色", 9: "棕色", 10: "青色"}
    for k in range(1, 11):
        lines.append(f"| {k}·{CHECK_META[k]['name']} | {COLOR_CN[k]} | {CHECK_META[k]['level']} |")
    return "\n".join(lines) + "\n"


def priority_text(by_check):
    steps = []
    if 8 in by_check:
        steps.append("先改检查项8（重复记录会让所有统计虚高）")
    hard = [k for k in (1, 3, 4, 6, 7) if k in by_check]
    if hard:
        steps.append("再改 " + "/".join(f"检查项{k}" for k in hard) + "（影响求和、排序、统计范围）")
    soft = [k for k in (9, 5, 2, 10) if k in by_check]
    if soft:
        steps.append("最后人工复核 " + "/".join(f"检查项{k}" for k in soft) + "（语义类，需人判断）")
    return "；".join(steps) if steps else "无硬伤，可直接进入分析"


def summarize_positions(items, limit=8):
    coords = [it["cell"] for it in items]
    if len(coords) <= limit:
        return "、".join(coords)
    return "、".join(coords[:limit]) + f" 等共 {len(coords)} 个单元格"


def process_file(path, outdir):
    wb, from_csv = load_workbook_any(path)
    base = os.path.splitext(os.path.basename(path))[0]
    scans = []
    for ws in wb.worksheets:
        if ws.max_row is None or ws.max_row < 1:
            continue
        scans.append(scan_sheet(ws, ws.title))
    # 表内跨 sheet 的字段一致性
    intra_issues, intra_notes = cross_table_check(scans)
    for scan in scans:
        scoped = [i for i in intra_issues if i["loc"].startswith(scan.sheet_name + "·")]
        annotate(wb[scan.sheet_name], scan, scoped if len(scans) > 1 else None)
        if len(scans) > 1 and scoped:
            scan.item_notes[10] = intra_notes
            for it in scoped:
                scan.issues.append(dict(check=10, cell="", row=0, col=0, level=it["level"],
                                        where=it["where"], why=it["why"], fix=it["fix"],
                                        extra=it["extra"]))
    out_xlsx = os.path.join(outdir, f"{base}_体检标注.xlsx")
    wb.save(out_xlsx)
    note = "源文件为 CSV，已转为 xlsx 副本（原 CSV 未改动），标注写在转换后的副本上。" if from_csv else ""
    reports = []
    for scan in scans:
        if len(scans) == 1 or scan.sheet_name == base or scan.sheet_name in base:
            report_path = os.path.join(outdir, f"{base}_体检报告.md")
        else:
            report_path = os.path.join(outdir, f"{base}_{scan.sheet_name}_体检报告.md")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(build_report(scan, os.path.basename(path), note,
                                 os.path.basename(out_xlsx)))
        reports.append(report_path)
    return dict(path=path, base=base, scans=scans, xlsx=out_xlsx, reports=reports,
                intra_notes=intra_notes, from_csv=from_csv)


def build_overview(results, outdir):
    """多文件总览 + 跨文件一致性检查。"""
    all_scans = [s for r in results for s in r["scans"]]
    cross_issues, cross_notes = cross_table_check(all_scans)
    lines = ["# 跨表体检总览", ""]
    lines.append(f"- 本次共 {len(results)} 个文件、{len(all_scans)} 个 sheet。")
    lines.append("")
    lines += ["## 各文件问题概览", "",
              "| 文件 | sheet | 数据量 | 高 | 中 | 低 | 标注副本 |",
              "| --- | --- | --- | --- | --- | --- | --- |"]
    for r in results:
        for s in r["scans"]:
            h = sum(1 for i in s.issues if i["level"] == "高")
            m = sum(1 for i in s.issues if i["level"] == "中")
            l_ = sum(1 for i in s.issues if i["level"] == "低")
            rows = max(0, getattr(s, "last_data_row", 0) - s.data_start + 1)
            lines.append(f"| {r['base']} | {s.sheet_name} | {rows} 行 × {s.last_data_col or s.max_col} 列 | {h} | {m} | {l_} | {os.path.basename(r['xlsx'])} |")
    lines.append("")
    lines += ["## 跨文件字段一致性（检查项10）", ""]
    if cross_notes:
        for n in cross_notes:
            lines.append(f"- {n}")
    else:
        lines.append("- 未发现明显的同义字段命名分歧。")
    lines.append("")
    lines += ["## 建议处理顺序", ""]
    # 汇总所有检查项
    total = Counter()
    for s in all_scans:
        for it in s.issues:
            total[it["check"]] += 1
    if total:
        for k in sorted(total):
            lines.append(f"{k}. 检查项{k}·{CHECK_META[k]['name']}（{CHECK_META[k]['level']}风险）共 {total[k]} 处 —— {FIX[k]}")
    else:
        lines.append("未命中问题，可直接进入分析。")
    path = os.path.join(outdir, "跨表体检总览.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    return path


def main():
    parser = argparse.ArgumentParser(description="资深数据分析师·数据体检：体检表格并生成标注副本与报告")
    parser.add_argument("files", nargs="+", help="待体检的表格文件（.xlsx/.xlsm/.csv），可多个")
    parser.add_argument("--outdir", default=None, help="输出目录，默认与原文件同目录")
    args = parser.parse_args()

    results = []
    for path in args.files:
        if not os.path.exists(path):
            print(f"[跳过] 文件不存在：{path}", file=sys.stderr)
            continue
        outdir = args.outdir or os.path.dirname(os.path.abspath(path))
        os.makedirs(outdir, exist_ok=True)
        try:
            results.append(process_file(path, outdir))
        except Exception as e:
            print(f"[失败] {path}：{e}", file=sys.stderr)
    if not results:
        sys.exit(2)

    print("体检完成：")
    for r in results:
        issues = sum(len(s.issues) for s in r["scans"])
        print(f"- {os.path.basename(r['path'])}：{issues} 处问题")
        print(f"  标注副本：{r['xlsx']}")
        for p in r["reports"]:
            print(f"  体检报告：{p}")
    if len(results) > 1:
        overview = build_overview(results, args.outdir or os.path.dirname(os.path.abspath(results[0]["path"])))
        print(f"  跨表总览：{overview}")
    else:
        r = results[0]
        cross_issues, cross_notes = cross_table_check([s for s in r["scans"]])
        if cross_notes and len(r["scans"]) > 1:
            print("  跨 sheet 字段一致性提示：")
            for n in cross_notes:
                print(f"   - {n}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
