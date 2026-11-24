#!/usr/bin/env python3
"""把 spec.json 注入模板，生成单文件 HTML 仪表盘。

用法:
    python3 build_dashboard.py spec.json 输出.html
    python3 build_dashboard.py spec.json 输出.html --template 自定义模板.html

退出码非 0 时会指出出错字段路径，照提示改 spec 即可。
"""
import argparse
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_TEMPLATE = HERE.parent / "assets" / "dashboard-template.html"
VALID_TYPES = {"line", "bar", "hbar", "donut", "funnel"}
THEMES = {"blue-orange", "ink-green", "graphite-violet", "slate-mono"}


def fail(msg):
    print("✗ " + msg, file=sys.stderr)
    sys.exit(1)


def check(spec):
    warn = []
    if not isinstance(spec, dict):
        fail("spec 顶层必须是 JSON 对象")
    if not str(spec.get("title") or "").strip():
        fail("缺 title（页面主标题，写结论句）")
    if not isinstance(spec.get("sections"), list) or not spec["sections"]:
        fail("sections 不能为空，至少一节图表")

    for i, k in enumerate(spec.get("kpis") or []):
        if not k.get("label") or k.get("value") in (None, ""):
            fail(f"kpis[{i}] 缺 label 或 value")
        if k.get("trend") not in (None, "up", "down", "flat", "none"):
            fail(f"kpis[{i}].trend 只能是 up/down/flat/none")

    if spec.get("theme") and spec["theme"] not in THEMES:
        fail(f"theme='{spec['theme']}' 不在 {'/'.join(sorted(THEMES))} 里")
    if (len(spec.get("kpis") or [])) > 6:
        warn.append("KPI 卡超过 6 张，建议压到 3–5 张")

    for i, sec in enumerate(spec["sections"]):
        prefix = f"sections[{i}]"
        if not str(sec.get("title") or "").strip():
            fail(f"{prefix} 缺 title（写结论，不写描述）")
        ch = sec.get("chart") or {}
        if ch.get("type") not in VALID_TYPES:
            fail(f"{prefix}.chart.type='{ch.get('type')}'，只支持 {'/'.join(sorted(VALID_TYPES))}")
        cats = ch.get("categories") or []
        series = ch.get("series") or []
        if not cats or not series:
            fail(f"{prefix}.chart 缺 categories 或 series")
        for s in series:
            if len(s.get("data") or []) != len(cats):
                fail(f"{prefix}.chart 系列「{s.get('name','')}」data 长度 {len(s.get('data') or [])} 与 categories {len(cats)} 不一致")
            for v in s.get("data") or []:
                if v is not None and not isinstance(v, (int, float)):
                    fail(f"{prefix}.chart 系列「{s.get('name','')}」里有非数字值 {v!r}")
        if ch["type"] == "donut":
            if len(cats) > 5:
                fail(f"{prefix} 是环形图但类别 {len(cats)} 个，超过 5 个请合并「其他」或改 hbar")
            if len(cats) < 2:
                fail(f"{prefix} 环形图至少 2 类")
        if ch["type"] == "funnel":
            nums = [v for v in series[0].get("data") or [] if isinstance(v, (int, float))]
            if any(nums[i] < nums[i + 1] for i in range(len(nums) - 1)):
                fail(f"{prefix} 漏斗各层必须逐层递减")
        if not (sec.get("readout") or {}).get("conclusion"):
            warn.append(f"{prefix} 缺 readout.conclusion，图上少一句结论")
        if len(spec["sections"]) > 8 and sec.get("span") == 2:
            warn.append("区块较多且有大区块，投屏可能要滚动，建议压缩")
    return warn


def js_safe(spec):
    text = json.dumps(spec, ensure_ascii=False, indent=2)
    # 防止内容提前闭合脚本标签，也避免 JS 里的行/段分隔符
    text = (text.replace("</", "<\\/")
                .replace("\u2028", "\\u2028")
                .replace("\u2029", "\\u2029"))
    return text


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec", help="spec.json 路径")
    ap.add_argument("out", help="输出的 html 路径")
    ap.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    args = ap.parse_args()

    spec_path = Path(args.spec).expanduser()
    if not spec_path.exists():
        fail(f"找不到 spec 文件：{spec_path}")
    try:
        spec = json.loads(spec_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        fail(f"spec.json 不是合法 JSON：{e}")

    warns = check(spec)

    tpl_path = Path(args.template).expanduser()
    if not tpl_path.exists():
        fail(f"找不到模板：{tpl_path}")
    tpl = tpl_path.read_text(encoding="utf-8")
    if "__DASHBOARD_DATA__" not in tpl:
        fail("模板里没有 __DASHBOARD_DATA__ 占位")

    html = tpl.replace("__DASHBOARD_DATA__", js_safe(spec)).replace("__TITLE__", str(spec.get("title", "数据仪表盘")).replace("<", "&lt;"))
    out = Path(args.out).expanduser()
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(html, encoding="utf-8")

    print(f"✓ 已生成 {out}（{out.stat().st_size / 1024:.1f} KB，{len(spec['sections'])} 个图表区块）")
    for w in warns:
        print("⚠ " + w)


if __name__ == "__main__":
    main()
