#!/bin/bash
# 抖音达播看板自动同步脚本：由 launchd WatchPaths 触发
# Excel 保存 → 消抖 → 同步(通过 osascript 获取桌面权限) → 构建 → 推送
# 独立 git 仓库，部署到 GitHub Pages（docs/ 目录）

set -e

PROJECT_DIR="/Users/xiaocao/CC/达播组近 30 日业绩看板"
LOG_FILE="$PROJECT_DIR/data/sync_dabogroup.log"
EXCEL_PATH="/Users/xiaocao/Desktop/蕉下文件/业绩追击/by月业绩/6月业绩/6月业绩追击（纯直播）.xlsx"
DEBOUNCE_SEC=5

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

# --- 消抖：等待文件写入完成 ---
BEFORE=$(stat -f "%m" "$EXCEL_PATH" 2>/dev/null || echo 0)
sleep $DEBOUNCE_SEC
AFTER=$(stat -f "%m" "$EXCEL_PATH" 2>/dev/null || echo 0)
if [ "$BEFORE" != "$AFTER" ]; then
    log "文件仍在变化中，再等 ${DEBOUNCE_SEC}s..."
    sleep $DEBOUNCE_SEC
fi

log "======== 检测到 Excel 更新，开始同步 ========"

cd "$PROJECT_DIR"

# Step 1: 数据同步（osascript 获取用户会话权限，绕过 macOS 桌面文件 TCC 限制）
log "--- 同步数据 ---"
osascript <<EOF
do shell script "cd '$PROJECT_DIR' && /usr/bin/python3 scripts/sync_dabogroup.py >> '$LOG_FILE' 2>&1"
EOF
if [ $? -ne 0 ]; then
    log "✗ 数据同步失败，终止"
    exit 1
fi

# Step 2: 构建独立页面
log "--- 构建页面 ---"
python3 scripts/build_dabogroup_standalone.py >> "$LOG_FILE" 2>&1 || {
    log "✗ 页面构建失败，终止"
    exit 1
}

# Step 3: Git 推送（需先配置 remote，参考：git remote add origin git@github.com:czcaizjy-lang/dabogroup.git）
log "--- 推送部署 ---"
git add data/dabogroup_data.json dabogroup.html docs/dabogroup.html 2>/dev/null

if git diff --cached --quiet 2>/dev/null; then
    log "无数据变更，跳过 git push"
else
    if git remote -v | grep -q origin; then
        git commit -m "📊 达播组数据自动更新 $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
        git push origin main >> "$LOG_FILE" 2>&1 && log "✓ 已推送到 GitHub Pages" || log "✗ git push 失败"
    else
        log "⚠ 未配置 git remote，仅本地提交"
        git commit -m "📊 达播组数据自动更新 $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1
    fi
fi

log "======== 同步完成 ========"
