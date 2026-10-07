# 主檔自動草稿（GAS「diybc-sku-autodraft」）備份

- 生產版在 Google 端：script id `1g_xlp7Fpo-vTsLxgzOYrDrZxrAY90gBHynR8cWfR0e-8ZfOlE_KO7hmG`；每日 06:00–07:00 觸發器 `autoDraftNow`（跑 HEAD，沒有網頁部署要更新）。
- 做四件事：A 補 BOM別名 B 主檔沒有的料建草稿 C 上架完成清標記 D 算 dim_sku「最後有效日」（＝用到它的甜點在產品名稱對照表的最晚結束日；採購系統據此把檔期結束的品項先放「🧹 清貨尾」14 天、再收進「🗓 過季」）。
- `semiLastValid.gs` 包裝 D：先把 BOM 照 dim_sku 店製配方（dim_semi）展開，讓只出現在配方裡的原料也有最後有效日。v1.1（2026-10-07）起原 BOM 列也保留，店製品項本身（例：啾啾鳥材料包）也會在檔期結束後退場。
- ⚠ **檔案載入順序**：`semiLastValid.gs` 一定要排在 `skuAutoDraft.gs` 後面（它一載入就抓 `adLastValid_`）。用 clasp 推版時 `.clasp.json` 要有 `"filePushOrder": ["skuAutoDraft.js", "semiLastValid.js"]`，否則 clasp 照字母排序會把它排到前面、包裝失效（2026-10-07 實際發生，已修）。
- 試跑：編輯器選 `autoDraftDryRun`（不寫入，執行記錄最後一行列出「最後有效日 N 筆有日期／M 筆已過季」，並有「🧁 最後有效日已含店製…」那行＝包裝有接上）。
