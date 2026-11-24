#!/usr/bin/env bash
# 用 LibreOffice 重算并缓存公式结果，方便在聊天/预览/手机上直接看到数字。
# 公式仍然保留，只是把计算结果写进缓存值。样式可能被 LibreOffice 微调。
# 用法：bash recalc_values.sh "原表_统计结果.xlsx"
set -euo pipefail

if [ $# -lt 1 ]; then
  echo "用法：bash recalc_values.sh <工作簿路径>" >&2
  exit 1
fi

f="$1"
[ -f "$f" ] || { echo "文件不存在：$f" >&2; exit 1; }
command -v soffice >/dev/null 2>&1 || { echo "未安装 LibreOffice，跳过重算（Excel 打开时会自动算）" >&2; exit 0; }

tmp=$(mktemp -d)
trap 'rm -rf "$tmp"' EXIT
soffice --headless --calc --convert-to xlsx --outdir "$tmp" "$f" >/dev/null 2>&1
base=$(basename "$f")
[ -f "$tmp/$base" ] || { echo "重算失败，保留原文件" >&2; exit 1; }
cp "$tmp/$base" "$f"
echo "已重算并写回缓存值：$f"
