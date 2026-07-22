# gas/pos — 「POS_折扣聚合」原始碼備份

- 生產專案：**POS_折扣聚合**（projectId 開頭 `1jsCg_m0…`）
- 生產部署：`AKfycbxv…`（永遠「管理部署 → 編輯現有部署 → 新版本」，不可新增部署）
- ⚠️ 廢棄分身「POS 儀表板」（`AKfycbze…`）**絕對不要編輯**
- 目前版本：`discount_aggregator.gs` **v19**（2026-07-22，檔期 detail 立方體＋加購餅乾源頭正名＋快取 V19）
- `bq_connector.gs`：SA 金鑰存於指令碼屬性 `SA_KEY_JSON`，**不在原始碼內**，本備份可公開
- 修改 SOP／地雷／迴歸錨點一律見 `dashboard-pos` skill

本資料夾為 2026-07-22 自 Drive 匯出＋v19 部署版快照；**日後每次改 GAS 部署後，須同步更新本資料夾**（鐵則 5）。
