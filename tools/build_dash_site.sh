#!/usr/bin/env bash
# 儀表板靜態網站（diybc-dash）建置腳本（2026-11-08）
# Render 靜態網站每次上傳時執行：把 static/ 整包複製到 _site/static/，
# 並列出所有儀表板用到的 Google Apps Script 部署 ID（gas-ids.txt），
# 讓訓練站伺服器的資料代轉（gas_proxy.py）不必重開也能放行新加的 ID。
set -euo pipefail
rm -rf _site
mkdir -p _site
cp -R static _site/static
grep -ohE 'script\.google\.com/macros/s/AKfycb[A-Za-z0-9_-]+/exec' static/*.html \
  | sed -E 's#^.*/s/(AKfycb[A-Za-z0-9_-]+)/exec$#\1#' | sort -u > _site/static/gas-ids.txt
echo "built: $(find _site -type f | wc -l) files, $(wc -l < _site/static/gas-ids.txt) gas ids"
