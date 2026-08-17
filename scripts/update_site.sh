#!/usr/bin/env bash
# ============================================================
#  DK 记录表 —— 收到手机发来的 K表/D表 后，一键更新网站
#  用法:
#    bash scripts/update_site.sh <K表.xlsx> <D表.xlsx> [YYYY-MM-DD]
#    bash scripts/update_site.sh            # 自动扫描 /root/uploads 里最新的 K点/D点 文件
#  依赖: process.py (openpyxl), node (JS 语法校验), git, PAT (见下)
# ============================================================
set -euo pipefail

REPO="/root/.codebuddy/artifact/dk-tracker"
cd "$REPO"

# ---- 1) 定位 K表 / D表 ----
K="" ; D="" ; DATE=""
if [ $# -ge 2 ]; then
  K="$1" ; D="$2" ; DATE="${3:-}"
else
  echo ">> 自动扫描 /root/uploads 最新 K点/D点 文件 ..."
  K=$(ls -t /root/uploads/*K点*.xlsx /root/uploads/*K点*.xls 2>/dev/null | head -1)
  D=$(ls -t /root/uploads/*D点*.xlsx /root/uploads/*D点*.xls 2>/dev/null | head -1)
fi
[ -f "$K" ] || { echo "[错误] 找不到 K表: $K"; exit 1; }
[ -f "$D" ] || { echo "[错误] 找不到 D表: $D"; exit 1; }
echo ">> K表: $K"
echo ">> D表: $D"

# ---- 2) 清掉旧切片，避免残留 ----
rm -f seed_p*.js app_p*.js seed_load.js app_load.js 2>/dev/null || true

# ---- 3) 合并当天 + 重建静态站点 ----
echo ">> 合并并重建 ..."
if [ -n "$DATE" ]; then
  python3.11 scripts/process.py --k "$K" --d "$D" --date "$DATE"
else
  python3.11 scripts/process.py --k "$K" --d "$D"
fi

# ---- 4) 观察池：现由腾讯行情实时计算（沙箱/桌面均可），无需还原缓存 ----

# ---- 5) JS 语法校验（解码模板内联脚本）----
echo ">> 校验应用 JS 语法 ..."
python3.11 - <<'PY'
import re
html=open('app_template.html',encoding='utf-8').read()
app=max(re.findall(r'<script>(.*?)</script>', html, re.S), key=len)
open('/tmp/_app_check.js','w').write(app)
PY
node --check /tmp/_app_check.js && echo "   JS 语法 OK"

# ---- 6) 提交 + 推送 ----
TOKEN_FILE="/root/.codebuddy/artifact/.ghtoken"
if [ ! -f "$TOKEN_FILE" ]; then
  echo "[错误] 缺少 PAT：请把 GitHub Token 再次发我，或设置 GH_TOKEN 环境变量。"
  exit 1
fi
TOKEN="$(cat "$TOKEN_FILE")"
git config user.email "bot@workbuddy.local" 2>/dev/null || true
git config user.name "WorkBuddy" 2>/dev/null || true

git add -A
if git diff --cached --quiet; then
  echo ">> 无变更，跳过提交。"
else
  MSG="data: 合并手机上传的 DK 表（$(date +%Y-%m-%d)）"
  git commit -q -m "$MSG"
  echo ">> 已提交: $MSG"
fi

echo ">> 推送到 main（经 ghproxy）..."
BASIC=$(printf 'x-access-token:%s' "$TOKEN" | base64 | tr -d '\n')
git remote set-url origin "https://ghproxy.net/https://github.com/seonkoo/dk-tracker.git"
git -c "http.extraHeader=Authorization: Basic $BASIC" push origin main 2>&1 | tail -8
echo "PUSH_EXIT:${PIPESTATUS[0]}"
echo ">> 完成。GitHub Pages 通常 1 分钟内自动重建，刷新网站即可看到更新。"
