#!/usr/bin/env bash
# 把 web/ 內容推到 gh-pages 分支(GitHub Pages 的發佈來源)
set -euo pipefail
cd "$(dirname "$0")/.."
git subtree split --prefix web -b gh-pages-tmp >/dev/null
git push -f origin gh-pages-tmp:gh-pages
git branch -D gh-pages-tmp >/dev/null
echo "已推送 gh-pages"
