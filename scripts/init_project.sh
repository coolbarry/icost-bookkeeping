#!/bin/bash
# 从技能包铺出一个工作目录（首次使用）／把脚本升级到最新版（已用过）
#
# 用法:
#   ./init_project.sh <工作目录>              首次：铺出完整目录（含 config 模板）
#   ./init_project.sh --upgrade <工作目录>    升级：只换脚本和文档，保留 config/ 和 example/
#
# 铺出来的结构:
#   <工作目录>/
#   ├── 账单/          放微信 / 支付宝 / iCost 导出的原始账单
#   ├── 导入文件/      split.sh 的产出，拿去导入 iCost
#   ├── output/        run.sh 的中间产物（⚠️ 不能直接导入）
#   └── icost-system/  系统本体（脚本 + config + docs + example）
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
SKILL_ROOT="$(cd "$HERE/.." && pwd)"

MODE="init"
TARGET=""
for a in "$@"; do
  case "$a" in
    --upgrade) MODE="upgrade" ;;
    -h|--help) sed -n '2,20p' "$0"; exit 0 ;;
    *) TARGET="$a" ;;
  esac
done

if [ -z "$TARGET" ]; then
  echo "用法: ./init_project.sh [--upgrade] <工作目录>"
  exit 1
fi

# 校验技能包自身结构完整
for d in scripts/core assets/config references; do
  [ -d "$SKILL_ROOT/$d" ] || { echo "❌ 技能包不完整，缺 $d/（$SKILL_ROOT）"; exit 1; }
done

SYS="$TARGET/icost-system"
mkdir -p "$TARGET/账单" "$TARGET/导入文件" "$TARGET/output" "$SYS/core"

# ---- 脚本（每次都覆盖，这是唯一的升级通道）----
cp "$SKILL_ROOT/scripts/lib.sh"   "$SYS/lib.sh"
cp "$SKILL_ROOT/scripts/run.sh"   "$SYS/run.sh"
cp "$SKILL_ROOT/scripts/split.sh" "$SYS/split.sh"
cp "$SKILL_ROOT/scripts/check.sh" "$SYS/check.sh"
for f in engine.py make_confirm.py export_split.py check_config.py; do
  cp "$SKILL_ROOT/scripts/core/$f" "$SYS/core/$f"
done
chmod +x "$SYS"/*.sh

# ---- 文档（references → docs）----
mkdir -p "$SYS/docs"
cp "$SKILL_ROOT"/references/*.md "$SYS/docs/"
# 这两份是流程入口，额外放一份在根目录方便查找
cp "$SKILL_ROOT/references/初次引导.md"   "$SYS/初次引导.md"   2>/dev/null || true
cp "$SKILL_ROOT/references/给AI的指令.md" "$SYS/给AI的指令.md" 2>/dev/null || true

# ---- config：升级模式必须保住用户的真实配置 ----
if [ "$MODE" = "upgrade" ]; then
  if [ -d "$SYS/config" ]; then
    echo "  config/ 已存在 → 保留（你的真实配置不会被覆盖）"
  else
    cp -R "$SKILL_ROOT/assets/config" "$SYS/config"
  fi
  if [ -d "$SYS/example" ]; then
    echo "  example/ 已存在 → 保留"
  else
    cp -R "$SKILL_ROOT/assets/example" "$SYS/example"
  fi
else
  cp -R "$SKILL_ROOT/assets/config"  "$SYS/config"
  cp -R "$SKILL_ROOT/assets/example" "$SYS/example"
fi

# ---- 账单投放说明 ----
cat > "$TARGET/账单/放这里.md" << 'EOF'
# 把导出的账单放这里

每次记账需要三份原始账单：

| 文件 | 从哪导出 |
|---|---|
| 微信账单 `.xlsx` | 微信 → 我 → 服务 → 钱包 → 账单 → 常见问题 → 下载账单（用于个人对账） |
| 支付宝账单 `.csv` | 支付宝 → 我的 → 账单 → 右上角 → 开具交易流水证明 → 用于个人对账 |
| **iCost 历史导出 `.xlsx`** | iCost → 设置 → 账单导出 |

⚠️ 第三份**不能省**。它是去重的依据，少了会把已经记过的账再导一遍。
EOF

echo ""
echo "✅ 完成（模式：$MODE）→ $TARGET"
echo ""
echo "下一步："
echo "  cd $SYS && ./check.sh"
echo ""
echo "  退出码 2 = 未配置，按 docs/初次引导.md 引导用户填 config/"
echo "  退出码 1 = 配置有错，按输出逐条修"
echo "  退出码 0 = 就绪，可以跑 run.sh 了"
