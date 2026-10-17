#!/bin/bash
# 把「已确认的明细」转成 iCost 真正能导入的文件：按账本拆分 + 对齐官方模板 XML
#
# 用法: ./split.sh <已确认明细.xlsx> [输出目录]
#   <已确认明细.xlsx> 需含 21 列（iCost 导出格式），由 AI 在第 4 步产出
# 产出:
#   输出目录/NN_iCost导入_<账本>.xlsx   —— 主用（10 列，sheet=icost_template）
#   输出目录/NN_iCost导入_<账本>.csv    —— 备用（xlsx 导入失败时改试）
#   输出目录/手工补录清单.md            —— AI 导不进去的退款/报销
#   输出目录/导入校验报告.md            —— 逐字段校验结果
#   输出目录/00_诊断_*.xlsx             —— 导入失败时的排查文件
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

PY=$(pick_py)
SRC="$1"; OUT="${2:-$HERE/../导入文件}"
if [ -z "$SRC" ]; then
  echo "用法: ./split.sh <已确认明细.xlsx> [输出目录]"; exit 1
fi
mkdir -p "$OUT"

echo "使用 Python: $PY"
"$PY" "$HERE/core/export_split.py" "$SRC" "$OUT"
echo ""
echo "✅ 完成 → $OUT"
echo "   导入时注意：先切到对应账本，一次只导一个文件。"
