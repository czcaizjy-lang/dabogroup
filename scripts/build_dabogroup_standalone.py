#!/usr/bin/env python3
"""
构建抖音达播看板独立 HTML 文件
读取 dabogroup.html（模板）和 dabogroup_data.json（数据），
生成内嵌数据的 docs/dabogroup.html（GitHub Pages 部署入口）。
"""

import json
import os
import sys

# 路径配置
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'dabogroup.template.html')
DATA_PATH = os.path.join(BASE_DIR, 'data', 'dabogroup_data.json')
OUTPUT_PATHS = [
    os.path.join(BASE_DIR, 'dabogroup.html'),            # 根目录：双击即开（内嵌数据）
    os.path.join(BASE_DIR, 'docs', 'index.html'),        # GitHub Pages 部署入口（docs/ 需有 index.html，根路径才可直接打开）
]

# ---- fetch 替换文本 ----
OLD_FETCH = """fetch('data/dabogroup_data.json?t=' + Date.now()).then(r => r.json()).then(d => {
  DATA = d;
  render();
}).catch(e => { console.error('Failed to load dashboard data:', e); });"""


def build():
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        template = f.read()

    if OLD_FETCH not in template:
        print('✗ 模板中未找到 fetch 语句，可能已经变更', file=sys.stderr)
        return False

    with open(DATA_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    data_str = json.dumps(data, ensure_ascii=False, separators=(',', ':'))

    # 拆分为三个脚本块：变量声明 → 内嵌数据 → 业务逻辑
    parts = template.split(OLD_FETCH, 1)
    before = parts[0]   # 含 <script> 开头 + 变量/主题/switchDataSource + IIFE
    after = parts[1]    # fmt / render / 图表 / 表格 / 下钻等所有业务逻辑

    inline_block = (
        '</script>\n'
        '<script id="inline-data" type="application/json">' + data_str + '</script>\n'
        '<script>\n'
        'DATA = JSON.parse(document.getElementById("inline-data").textContent);\n'
        'setTimeout(render, 0);'
    )

    output = before + inline_block + after

    # 写入所有输出路径（根目录 + docs 各一份内嵌数据版）
    for path in OUTPUT_PATHS:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(output)

    db = data.get('dabogroup', {})
    zb_all = db.get('summary_by_zb', {}).get('全部', {})
    size_kb = len(output) / 1024

    print(f'✓ 已生成 {len(OUTPUT_PATHS)} 个文件')
    for p in OUTPUT_PATHS:
        print(f'  → {os.path.relpath(p, BASE_DIR)}')
    print(f'  模板: {os.path.basename(TEMPLATE_PATH)} ({len(template)} 字符)')
    print(f'  数据: {os.path.basename(DATA_PATH)} ({len(data_str)} 字符)')
    print(f'  全量达人: {len(db.get("anchors", []))} | 全部GMV ¥{(zb_all.get("直播GMV", 0) or 0):,.2f}')
    return True


if __name__ == '__main__':
    ok = build()
    sys.exit(0 if ok else 1)
