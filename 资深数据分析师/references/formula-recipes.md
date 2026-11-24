# spec 结构与公式速查

脚本：`scripts/add_stats_sheet.py`。读取 JSON spec → 复制源表 → 新增 sheet → 写公式 → 保存。

## 一、spec 顶层

```json
{
  "source": {"file": "原表.xlsx", "sheet": "明细", "header_row": 3, "data_range": "A4:G23"},
  "output": {"file": "原表_统计结果.xlsx"},
  "sheets": [ {"name": "按城市统计", "layout": "group", "group_by": ["城市"], "metrics": [...]} ]
}
```

- `source` 顶层默认值，可被单个 sheet 里的 `source` 覆盖（多表场景常用）。
- `file` 用绝对路径最稳。
- `header_row`：表头所在行，默认 1。表上方有标题/导出时间时必须改。
- `data_range`：**省略即整列引用**（能跟着数据长）；表底有"合计/总计/小计"行、或上面有标题行时，必须写明确数据区（如 `A4:G23`），否则合计行会被重复计入。
- `output.file`：结果副本路径。源文件不会被改写。

每个 sheet 通用字段：

| 字段 | 说明 |
| --- | --- |
| `name` | 新 sheet 名，用业务语言，如"按城市统计" |
| `layout` | `group` / `total` / `join` / `combine` |
| `notes` | 追加到 sheet 底部备注的一行说明（口径、单位、去重规则） |
| `freeze_header` | 默认 true，冻结首行 |

## 二、layout: group（分组汇总，最常用）

```json
{
  "name": "按城市统计", "layout": "group",
  "group_by": ["城市"],
  "metrics": [
    {"label": "销售额合计", "type": "sum", "field": "销售额"},
    {"label": "订单数", "type": "count"},
    {"label": "平均销售额", "type": "average", "field": "销售额"},
    {"label": "销售额占比", "type": "share_of", "base": "销售额合计"},
    {"label": "销售额排名", "type": "rank_of", "base": "销售额合计"},
    {"label": "成本合计", "type": "sum", "field": "成本"},
    {"label": "销售额成本比", "type": "ratio_of", "numerator": "销售额合计", "denominator": "成本合计"}
  ],
  "sort_groups": false, "keep_blank": false
}
```

支持的 `type`：

| type | 必填 | 公式 | 说明 |
| --- | --- | --- | --- |
| `sum` | `field` | `SUMIFS(值列, 条件列, 本行维度...)` | 分组求和 |
| `count` | 可选 `field` | `COUNTIFS(条件列, 本行维度, 值列,"<>")` | 分组行数；给 field 即该字段非空条数 |
| `count_field` | `field` | 同上 | 明确"非空计数" |
| `average` | `field` | `IFERROR(AVERAGEIFS(...),"")` | 分组平均 |
| `max` / `min` | `field` | `_xlfn.MAXIFS` / `_xlfn.MINIFS` | 分组最大/最小 |
| `share_of` | `base` | 本行指标 ÷ 该指标列合计 | 占比（百分比格式） |
| `rank_of` | `base`、可选 `order`(0 降序/1 升序) | `RANK(本行, 该指标整列, order)` | 排名 |
| `ratio_of` | `numerator`、`denominator`（均为本 sheet 指标 label） | 两列相除 | 比值（百分比格式） |
| `diff_of` | `a`、`b` | 两列相减 | 差值 |
| `sum_of` | `a`、`b` | 两列相加 | 合计 |
| `product_of` | `a`、`b` | 两列相乘 | 乘积 |
| `custom` | `formula_tpl` | 原文，`{row}` 会被替换为当前行号 | 兜底手写 |

多维度（交叉表）直接 `"group_by": ["城市", "门店"]`，公式自动加多组条件；占比/排名的 `base` 列是明细组合后的结果，不是小计层级，要小计请另开一张 sheet。

## 三、layout: total（整表一个数）

```json
{
  "name": "整体汇总", "layout": "total",
  "metrics": [
    {"label": "销售额总计", "type": "sum", "field": "销售额"},
    {"label": "订单总数", "type": "count", "field": "订单号"},
    {"label": "平均销售额", "type": "average", "field": "销售额"},
    {"label": "最高销售额", "type": "max", "field": "销售额"},
    {"label": "最低销售额", "type": "min", "field": "销售额"}
  ]
}
```

只支持 `sum / average / max / min / count`；输出为"指标 | 统计结果 | 口径"三列，适合做汇总页。

## 四、layout: join（两表按关联键对齐）

```json
{
  "name": "人员产出与工资对比", "layout": "join",
  "left":  {"sheet": "销售汇总", "key": ["员工号"], "keep": ["员工号", "销售额"]},
  "right": {"sheet": "工资表",   "key": ["员工号"],
            "bring": [{"label": "姓名", "field": "姓名"}, {"label": "实发工资", "field": "实发工资"}]},
  "metrics": [
    {"label": "销售额工资比", "type": "ratio_of", "numerator": "销售额", "denominator": "实发工资"},
    {"label": "销售额减工资", "type": "diff_of", "a": "销售额", "b": "实发工资"}
  ],
  "distinct_rows": true
}
```

- `left.keep`：左表要输出的字段（含键）。左表字段写成 `='销售汇总'!$B$2` 这样的**单元格公式**，不落死值。
- `right.bring`：右表搬过来的字段，用 `INDEX/MATCH` 按键匹配，匹配不上留空。
- **多键关联**：`key: ["门店","月份"]` 与右表同序，脚本自动用 `MATCH(1,INDEX((右键1=左键1)*(右键2=左键2),0),0)`。
- `distinct_rows`：左表同键重复时只保留首行（默认 true）；要做"一对多明细"改成 false。
- `metrics` 只支持行级运算：`ratio_of / diff_of / sum_of / product_of / custom`；分组汇总请再开一张 `group` sheet。

## 五、layout: combine（多表同字段纵向合计）

```json
{
  "name": "全平台阅读量", "layout": "combine",
  "sources": [
    {"sheet": "抖音", "field": "阅读量", "data_range": "A2:F200"},
    {"sheet": "小红书", "field": "阅读量"}
  ],
  "metrics": [
    {"label": "阅读量合计", "type": "sum_multi"},
    {"label": "篇数合计", "type": "count_multi"},
    {"label": "平均阅读量", "type": "average_multi"},
    {"label": "抖音占比", "type": "ratio_of", "numerator": "阅读量合计", "denominator": "阅读量合计"}
  ]
}
```

只用于"结构相同、字段同义"的多表相加。字段名不一致时，先用 Python 统一列名，或在 `sources` 里各自指定 `field`。

## 六、常用公式对照

| 目的 | 公式 |
| --- | --- |
| 分组求和 | `=SUMIFS('明细'!$D:$D,'明细'!$A:$A,$A2)` |
| 多条件求和 | `=SUMIFS('明细'!$F:$F,'明细'!$A:$A,$A2,'明细'!$B:$B,$B2)` |
| 分组计数 | `=COUNTIFS('明细'!$A:$A,$A2,'明细'!$F:$F,"<>")` |
| 分组平均 | `=IFERROR(AVERAGEIFS('明细'!$D:$D,'明细'!$A:$A,$A2),"")` |
| 分组最大/最小 | `=IFERROR(_xlfn.MAXIFS('明细'!$D:$D,'明细'!$A:$A,$A2),"")` |
| 占比 | `=IFERROR($B2/SUM($B$2:$B$9),"")` |
| 排名（降序/升序） | `=IFERROR(RANK($B2,$B$2:$B$9,0),"")` / `...,1` |
| 两指标比值 | `=IFERROR($B2/$G2,"")` |
| 跨表搬字段（单键） | `=IFERROR(INDEX('工资'!$D:$D,MATCH($A2,'工资'!$A:$A,0)),"")` |
| 跨表搬字段（多键） | `=IFERROR(INDEX('工资'!$D:$D,MATCH(1,INDEX(('工资'!$A:$A=$A2)*('工资'!$B:$B=$B2),0),0)),"")` |
| 按月分组（辅助列法） | 原表旁加 `=TEXT(C2,"yyyy-mm")` 辅助列，再按辅助列 SUMIFS |
| 环比/同比 | 结果表加"上期"列：`=SUMIFS(值,期间列,上期)`，环比 `=本期/上期-1` |
| 去重计数（无 Excel 365） | `=SUMPRODUCT((条件)*(1/COUNTIFS(键列,键列&"")))`，先确保键列无空值 |
| 文本数字转数值 | `=VALUE(TRIM(A2))` 或 `=SUMPRODUCT(--(条件)*(值列))` |

## 七、常见坑（照着排）

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| 结果全是 0 | 文本型数字（"1,200"）、带空格、或字段名取错 | 体检看"数字文本"，加 `VALUE`/`--` 转换；核对字段 |
| 数字翻倍 | 整列引用把底部"合计"行也算了 | spec 加 `data_range` |
| 占比合计不等于 100% | 有未匹配/空维度行，或分子分母口径不同 | 检查 `keep_blank` 与 base 指标 |
| `#NAME?` | 用了 `XLOOKUP`、`MAXIFS` 未加前缀等新函数 | 换 `INDEX/MATCH`、`_xlfn.MAXIFS` |
| 排名出现 `#N/A` | 被排的指标是空文本 | 用 `IFERROR(RANK(...),"")` |
| 日期分组不对 | 日期是文本 | 辅助列 `=DATEVALUE(...)` 或 `=TEXT(...,"yyyy-mm")` |
| 打开显示 0 但点进去有公式 | Excel 未重算 | 脚本已设 `fullCalcOnLoad`；仍不行按 `Ctrl+Alt+F9` |
| 图表/透视没了 | openpyxl 另存丢弃对象 | 生成前提醒用户，或改成交独立结果工作簿 |

## 八、交付前自检清单

1. `--check`：公式数 > 0、sheet 名对、`fullCalcOnLoad: True`。
2. 打开结果 sheet：分组维度是否含"合计"、维度是否漏（如空格/大小写差异）。
3. 抽 1–2 组手工对一遍：SUMIFS 结果 = 该组明细求和。
4. 每个 sheet 底部有口径备注；百分比列是百分比格式。
5. 用 `recalc_values.sh` 刷新缓存值（可选），再交付 xlsx。
