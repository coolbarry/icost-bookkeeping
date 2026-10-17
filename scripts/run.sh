#!/bin/bash
# iCost 智能记账 —— 一键流程（账单 → 归类 → 确认页）
#
# 用法: ./run.sh <微信账单.xlsx> [支付宝账单.csv] [iCost历史导出.xlsx]
#   不传第 3 个参数时，默认找 ../iCost历史导出.xlsx
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

PY=$(pick_py)
WX="$1"; AL="${2:-}"; HIS="${3:-$HERE/../iCost历史导出.xlsx}"
OUT="$HERE/../output"
mkdir -p "$OUT"

echo "使用 Python: $PY"
"$PY" "$HERE/core/engine.py" --wechat "$WX" ${AL:+--alipay "$AL"} \
  ${HIS:+--history "$HIS"} --out "$OUT"

if [ -f "$OUT/记账明细.csv" ]; then
  "$PY" "$HERE/core/make_confirm.py" --rows "$OUT/记账明细.csv" \
    --out "$OUT/记账确认.html"
  echo ""
  echo "✅ 完成。请打开: $OUT/记账确认.html"
  echo "   1) 处理「待确认」的几笔   2) 点「复制全部」   3) 把结果发回给 AI"
else
  echo "❌ 引擎没有产出 记账明细.csv，请检查上面的报错"
  exit 1
fi
