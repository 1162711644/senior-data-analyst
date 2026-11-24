# spec.json 格式

一行一个字段，缺省自动补默认值。跑 `python3 scripts/build_dashboard.py spec.json 输出.html` 即可出成品。

## 顶层

| 字段 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | str | ✅ | 页面主标题，**写结论** |
| `subtitle` | str | | 一行副标题（口径 / 范围） |
| `asOf` | str | | 数据截至，如 `数据截至 2025-10-31` |
| `source` | str | | 数据来源，如 `经营数据表.xlsx + 2025 年经营分析报告` |
| `theme` | str | | `blue-orange`(默认) / `ink-green` / `graphite-violet` / `slate-mono` |
| `themeCustom` | obj | | 覆写色值，键：`bg` `surface` `text` `muted` `primary` `secondary` `accent` `line` |
| `footer` | str | | 页脚，默认 `数据快照 · 更新请重新生成` |
| `kpis` | array | | KPI 卡，3–5 个 |
| `sections` | array | ✅ | 图表区块 |

## kpis[]

```json
{
  "label": "营业收入（万元）",
  "value": "8,420",
  "unit": "万元",
  "delta": "+23.4%",
  "trend": "up",            // up | down | flat | none
  "note": "对比 2024 年同期，超年度目标 5.2%",
  "highlight": true          // 可选，标为重点卡（用强调色边线）
}
```

## sections[]

```json
{
  "title": "营收 Q3 起加速，全年同比 +23%",   // 结论式标题
  "subtitle": "月度营收（万元）",
  "span": 2,                                  // 1 | 2，占几列（默认 1）
  "chart": { ... },
  "readout": {
    "conclusion": "Q3 起营收加速，全年同比 +23%。",
    "evidence": "7–9 月连续三月环比正增长，9 月为全年峰值。",
    "boundary": "含 10 月预估值，未剔除退货，绝对值仅供趋势参考。"
  }
}
```

### chart 五种

```json
// 折线：categories + series[]
{"type":"line","unit":"万元","categories":["1月","2月","3月"],
 "series":[{"name":"2025","data":[620,680,750]},{"name":"2024","data":[540,560,600]}],
 "highlightSeries": 0}

// 柱形（纵向）：同上结构
{"type":"bar","unit":"万元","categories":["华东","华南","华北"],"series":[{"name":"营收","data":[320,210,180]}],"highlightIndex": 0}

// 条形（横向排名，适合长名称/多类别）
{"type":"hbar","unit":"万元","categories":["上海直营","杭州直营","南京加盟"],"series":[{"name":"营收","data":[180,120,96]}]}

// 环形占比：≤5 类
{"type":"donut","unit":"%","categories":["直营","加盟","线上"],"series":[{"name":"占比","data":[52,30,18]}]}

// 漏斗：3–5 层
{"type":"funnel","unit":"人","categories":["曝光","点击","加购","下单"],"series":[{"name":"转化","data":[120000,24000,7200,1800]}]}
```

- `data` 里允许 `null`（断开点），`line` 会跳过。
- 颜色按 `primary → secondary → accent → muted` 顺序循环，第 4 条起用灰阶。
- `highlightSeries` / `highlightIndex` 指定的那条用强调色，其余用辅助色。

## 最小示例

```json
{
  "title": "营收同比增长 23%，Q3 起加速",
  "subtitle": "2025 年 1–10 月 · 华东大区",
  "asOf": "数据截至 2025-10-31",
  "source": "经营数据表.xlsx",
  "theme": "blue-orange",
  "kpis": [
    {"label":"营业收入（万元）","value":"8,420","delta":"+23.4%","trend":"up","note":"同比","highlight":true},
    {"label":"毛利率","value":"32.1","unit":"%","delta":"+1.8pp","trend":"up","note":"同比"},
    {"label":"年度目标达成","value":"105","unit":"%","delta":"+5.2pp","trend":"up","note":"对比 100% 目标"}
  ],
  "sections": [
    {"title":"Q3 起营收加速，9 月为全年峰值","span":2,
     "chart":{"type":"line","unit":"万元","categories":["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月"],
       "series":[{"name":"2025","data":[620,680,750,780,760,820,900,960,1080,1050]}]},
     "readout":{"conclusion":"Q3 起营收加速，全年同比 +23%。","evidence":"7–9 月连续三月环比正增长，9 月达 1,080 万。","boundary":"含 10 月预估值，绝对值仅供趋势参考。"}},
    {"title":"直营贡献过半营收","chart":{"type":"donut","categories":["直营","加盟","线上"],"series":[{"name":"占比","data":[52,30,18]}]},
     "readout":{"conclusion":"直营占比 52%，是营收主力。","evidence":"直营 4,378 万，加盟 2,526 万，线上 1,516 万。","boundary":"按开票主体归类，加盟转直营口径未单独拆分。"}},
    {"title":"上海直营单店产出最高","chart":{"type":"hbar","unit":"万元","categories":["上海直营","杭州直营","南京加盟","苏州加盟"],"series":[{"name":"营收","data":[180,120,96,72]}]},
     "readout":{"conclusion":"上海直营单店产出领先。","evidence":"180 万，为末位的 2.5 倍。","boundary":"仅 4 家样本，不宜外推到全区。"}}
  ]
}
```

## 校验与常见坑
- `categories` 与每条 `data` 长度必须一致；`donut` 至少 2 类、最多 5 类（超了合并「其他」）。
- `funnel` 数据必须逐层递减。
- 数字写字符串时不要带单位，单位放 `unit`。
- 生成脚本报错会直接指出字段路径，照改即可。
