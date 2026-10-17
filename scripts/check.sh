#!/bin/bash
# 配置体检：判断系统是「未配置 / 有问题 / 可以直接跑」
#
# 用法:
#   ./check.sh
#
# 退出码: 0 就绪  1 有问题  2 未配置（需先按 初次引导.md 配置）
HERE="$(cd "$(dirname "$0")" && pwd)"
# shellcheck source=lib.sh
. "$HERE/lib.sh"

# 即使没装 openpyxl 也要能跑体检（脚本内部会提示缺少依赖）
PY=$(pick_py 2>/dev/null) || PY=python3

"$PY" "$HERE/core/check_config.py" "$HERE"
exit $?
