# openpyxl 操作细则与脚本用法

## 辅助脚本 `scripts/sheet_fix.py`

在技能目录下运行，路径按实际环境调整。所有命令都直接改传入的 `_已整理.xlsx`，不会碰原文件。

```bash
S=scripts/sheet_fix.py

# 1. 复制副本（先复制再改，永不原地编辑）
python3 $S clone "订单.xlsx" "订单_已整理.xlsx"

# 2. 看结构：工作表、维度、合并区域、前 N 行
python3 $S inspect "订单_已整理.xlsx" --rows 10

# 3. 拆合并并向下填充（检查项1）
python3 $S unmerge-fill "订单_已整理.xlsx" --range B2:B10 --sheet 订单

# 4. 标黄 + 批注（人工复核可视化）
python3 $S mark "订单_已整理.xlsx" --cells D3,E9,D12:D13 \
  --note "原值：01/21；原因：缺年份；待您确认：补全年份" --sheet 订单

# 5. 交付前自检：统计标黄数量、找出漏批注的单元格
python3 $S audit "订单_已整理.xlsx"
```

要点：

- `unmerge-fill` 只处理单列合并（如 `B2:B10`）；跨列合并（如 `A1:D1` 的表名标题）会拆分但不填充，交给检查项9 处理。
- `mark` 可传单元格或区域，逗号分隔；重复调用会覆盖原批注，同一格的批注请一次写全。
- `audit` 的"合计标黄"应与变更日志的"待人工确认"数量对得上；差太多说明有漏标或误标。

## 常写代码片段

### 读取体检报告

报告是表格时，按"检查项 / 工作表 / 单元格 / 原值 / 结论 / 是否人工复核"的列名定位；列名不标准就按内容猜列，并在回复中说明你的理解。报告是 md/docx/pdf 时先转文本再按同样的字段抽取。

### 日期统一（检查项3）

```python
import datetime, re

def to_iso(v):
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    m = re.match(r"^(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?$", s)
    if m:
        y, mo, d = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d}"
    m = re.match(r"^(\d{4})(\d{2})(\d{2})$", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return None  # 无法识别或缺年份 → 保留原值 + 标黄
```

### 剥单位（检查项2 / 检查项4）

```python
import re

ESTIMATE = re.compile(r"(约|大约|左右|大概|近似|暂定|未定)")
def strip_unit(v, divisor=1.0):
    s = str(v).strip().replace(",", "").replace("，", "").replace("＋", "").replace("+", "")
    approx = bool(ESTIMATE.search(s))
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    if not m:
        return "【待确认】", approx
    return round(float(m.group()) / divisor, 4), approx
```

换算（如万元→元）必须在日志注明系数，单位写进表头：`金额（元）`。

### 删除空行空列（检查项7）

```python
def is_blank_row(ws, r):
    return all((ws.cell(row=r, column=c).value in (None, "")) for c in range(1, ws.max_column + 1))

for r in range(ws.max_row, 0, -1):
    if is_blank_row(ws, r):
        ws.delete_rows(r)
```

删列同理用 `ws.delete_cols`，倒序处理，删前确认该列确实无内容、无表头意义。

### 去重（检查项8）

```python
import pandas as pd

df = pd.read_excel("订单_已整理.xlsx", sheet_name="订单", dtype=str)
df = df.drop_duplicates()                       # 完全重复
# 主键去重（主键已由用户确认）
df = df.sort_values("更新时间").drop_duplicates("订单号", keep="last")
```

pandas 读入会把数字、日期转成文本或推断类型，**写回前不要整表覆盖**：优先只回写被修改的列，或用 `dtype=object` + 写回后核对格式。改动面大时，可用 pandas 处理完直接生成新表，但要保留原表头与列顺序。

### 合并单元格取值

```python
for m in list(ws.merged_cells.ranges):
    if m.max_col - m.min_col == 0:      # 纵向合并
        top = ws.cell(row=m.min_row, column=m.min_col).value
        ws.unmerge_cells(str(m))
        for r in range(m.min_row, m.max_row + 1):
            ws.cell(row=r, column=m.min_col).value = top
```

## 保留格式的注意事项

- `load_workbook` 默认保留样式、列宽、合并、批注；另存会重写文件，图表/图片/透视表可能丢失，先提醒用户。
- `.xlsm` 用 `load_workbook(path, keep_vba=True)`，否则宏被剥离。
- `.xls` 老旧格式先用 `libreoffice --headless --convert-to xlsx` 转换后再处理。
- 原表有筛选、冻结窗格、条件格式时，改动区域尽量集中在单元格值，不整表重建。
