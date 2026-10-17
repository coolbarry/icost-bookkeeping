#!/bin/bash
# 公共函数：自动挑一个装了 openpyxl 的 Python
#
# 为什么要这个：macOS 上 `python3` 可能指向没装 openpyxl 的解释器，
# 直接跑会报 ModuleNotFoundError。这里按优先级逐个试，挑到能 import 的就停下。

pick_py() {
  local cands=(
    "${PYTHON:-}"
    "$HOME/.workbuddy/binaries/python/envs/default/bin/python"
    "$HOME/.workbuddy/binaries/python/versions/3.13.12/bin/python3"
    "$HOME/venv/bin/python"
    "./.venv/bin/python"
    "python3"
    "python"
  )
  local c
  for c in "${cands[@]}"; do
    [ -n "$c" ] || continue
    command -v "$c" >/dev/null 2>&1 || continue
    "$c" -c "import openpyxl" >/dev/null 2>&1 && { echo "$c"; return 0; }
  done
  # 全都不可用：给最后一次机会，装上再试
  local fallback="${PYTHON:-python3}"
  echo ""
  echo "❌ 没找到装了 openpyxl 的 Python，请先执行一次："  >&2
  echo "   $fallback -m pip install --user openpyxl"          >&2
  echo "   或者：export PYTHON=/你的/python/路径"               >&2
  return 1
}
