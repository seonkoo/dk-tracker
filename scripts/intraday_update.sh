#!/usr/bin/env bash
# ============================================================
#  盘中更新脚本（dk-tracker）
#  用途：交易日盘中（如 10:30 / 14:30）重新生成站点，刷新
#        - 资金流向 tab 的 earnings-radar 真实盘中数据
#        - 持仓·回踩 / 回踩监测 的腾讯行情盘中价（K线）
#  并推送到 GitHub Pages。
#
#  调用方式：
#    本地/沙箱：  bash scripts/intraday_update.sh
#    GitHub Actions：设置 secret GH_TOKEN 后由 workflow 调用
# ============================================================
set -euo pipefail

REPO="/root/.codebuddy/artifact/dk-tracker"
cd "$REPO"

# ---- Token：优先环境变量（GitHub Actions），回退本地文件（沙箱）----
TOKEN="${GH_TOKEN:-}"
if [ -z "$TOKEN" ] && [ -f "$REPO/.ghtoken" ]; then
  TOKEN="$(cat "$REPO/.ghtoken")"
fi
if [ -z "$TOKEN" ]; then
  echo "[错误] 缺少 Token：请设置 GH_TOKEN 环境变量，或放 .ghtoken 到仓库目录。"
  exit 1
fi

# ---- 清旧切片，重新生成（拉取盘中最新数据）----
rm -f seed_p*.js app_p*.js seed_load.js app_load.js 2>/dev/null || true
echo ">> 重新生成站点（盘中数据）..."
python3.11 scripts/process.py --regen

# ---- 提交 + 推送 ----
git config user.email "bot@workbuddy.local" 2>/dev/null || true
git config user.name "WorkBuddy" 2>/dev/null || true
git add -A
if git diff --cached --quiet; then
  echo ">> 无变更，跳过。"
  exit 0
fi
git commit -q -m "data: 盘中更新（$(date '+%Y-%m-%d %H:%M')）"
BASIC=$(printf 'x-access-token:%s' "$TOKEN" | base64 | tr -d '\n')
git remote set-url origin "https://ghproxy.net/https://github.com/seonkoo/dk-tracker.git"
git -c "http.extraHeader=Authorization: Basic $BASIC" push origin main 2>&1 | tail -5
echo ">> 盘中更新完成，GitHub Pages 通常 1 分钟内重建。"
