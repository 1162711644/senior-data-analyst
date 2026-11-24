# 资深数据分析师 · 麦肯锡商业级数据分析

> 用户丢来一个（或多个）Excel，从「这表有问题」一路做到「开会能投屏」：**体检 → 整理 → 统计 → 分析报告 → 数据看板**，五个工位一条流水线，可全跑，也可只做其中一步。

一个 Agent Skill（豆包 / Claude 通用），把五件常被拆开的事串成一条可复核的证据链：先找脏数据，再修成能算的表，再挖出值得算的指标，再写成结论先行的 Word 报告，最后上一块单文件 HTML 看板。

## 五个工位

| 工位 | 做什么 | 交付 |
| --- | --- | --- |
| ① 体检 | 十项扫描：合并单元格、日期乱、数字夹字、一格多信息、空格、重复行、表头问题、跨表命名 | `_体检标注.xlsx`（填色+批注）、`_体检报告.md` |
| ② 整理 | 按体检结果逐项修复，只改副本；业务口径先问再动 | `_已整理.xlsx`、`变更日志.md` |
| ③ 统计 | 挖出该统计什么 → 写一句可复制的统计指令 → 公式落地到新 sheet | `_统计结果.xlsx`（SUMIFS/RANK/SHARE 全公式，改数自动重算） |
| ④ 分析报告 | 五段固定结构 + 折线/条形/环形图表，结论先行 | `.docx`（蓝白 / 红白两套色系） |
| ⑤ 数据看板 | KPI 卡 + 分区图表 + 结论/证据/边界三句解读 | 单文件 `.html`（零依赖、离线可开、投屏可读） |

## 安装

把 `资深数据分析师/` 整个文件夹放进你的 Skills 目录：

- 豆包 / 办公模式：`<workspace>/.user_skills/资深数据分析师/`
- Claude Code：`~/.claude/skills/资深数据分析师/`
- 其他兼容 Agent Skills 规范的客户端：放进对应 skills 目录即可，`SKILL.md` 是唯一入口

```bash
git clone https://github.com/1162711644/senior-data-analyst.git
cp -r senior-data-analyst/资深数据分析师 <你的 skills 目录>/
```

依赖（只在跑脚本时需要）：

```bash
pip3 install openpyxl python-docx matplotlib
```

## 怎么用

丢一张表，说一句话：

> 帮我把这张销售表整理一下，出个分析报告和看板。

或者点名某一步：

> 先体检，别动数据。 / 按订单号去重，日期统一格式。 / 只看各城市销售额排名，用公式。 / 报告用红白。

技能会先报菜单（五件事能做什么），再按你的选择推进；四个业务口径点（单位、去重主键、是否拆分、关联键）会先问，其余自己判断并写在报告的口径说明里。

## 目录结构

```
资深数据分析师/
├── SKILL.md                      # 总控：流水线、必问项、铁规矩、自检
├── references/
│   ├── 01-体检规则.md             # 十项检查项判定 + 颜色/批注体系 + 报告结构
│   ├── 02-整理规则.md             # 修复规则 + 三类必问 + 标黄规范 + 变更日志
│   ├── 02a-openpyxl-recipes.md
│   ├── 02b-变更日志模板.md
│   ├── 03-统计需求挖掘.md         # 单表/多表分流 + 一句指令 + 公式落地
│   ├── single-table-prompt.md
│   ├── multi-table-prompt.md
│   ├── formula-recipes.md        # spec 结构、指标类型、公式速查
│   ├── analysis-playbook.md      # 报告五段怎么写
│   ├── docx-spec.md              # 报告 JSON 规格与排版
│   ├── chart-and-layout-rules.md # 图表选型、配色禁忌、投屏标准
│   ├── style-presets.md          # 四套看板风格
│   ├── spec-format.md            # 看板 spec.json 字段
│   └── 06-麦肯锡方法论.md         # 结论先行 / MECE / 假设驱动 / 建议落地 / 七条禁语
├── scripts/
│   ├── inspect_workbook.py       # 读表体检：字段/类型/空值/唯一键/合计行
│   ├── table_check.py            # 十项扫描 + 标注副本 + 体检报告
│   ├── sheet_fix.py              # clone / unmerge-fill / mark / audit
│   ├─��� add_stats_sheet.py        # 按 spec 生成公式版统计 sheet（--check 自检）
│   ├── recalc_values.sh          # LibreOffice 把公式结果缓存进文件
│   ├── build_report_docx.py      # JSON → Word 报告
│   └── build_dashboard.py        # JSON → 单文件 HTML 看板
└── assets/
    └── dashboard-template.html   # 零依赖看板模板（SVG + 纯 CSS）
```

## 设计取舍

- **原表永不改**：所有产物写在副本上（`_体检标注` / `_已整理` / `_统计结果`），体检报告与变更日志让每处改动可复核。
- **不编数据**：表里没有的字段、期间、口径写「数据不足」，不估算、不占位。
- **只问四处**：单位口径、去重主键、是否拆分、多表关联键。其余不打扰，判断依据写进报告。
- **统计用公式不用死值**：改了源数据，结果 sheet 自动重算。
- **麦肯锡标准**：标题写判断句、指标按 MECE 拆、结论配证据与边界、建议含「动作 + 预期 + 衡量」。

## 许可

MIT
