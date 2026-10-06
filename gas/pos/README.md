# gas/pos — 「POS_折扣聚合」原始碼備份

- 生產專案：**POS_折扣聚合**（projectId 開頭 `1jsCg_m0…`）
- 生產部署：`AKfycbxv…`（永遠「管理部署 → 編輯現有部署 → 新版本」，不可新增部署）
- ⚠️ 廢棄分身「POS 儀表板」（`AKfycbze…`）**絕對不要編輯**
- 目前版本：**部署 @20（v21.1，2026-10-06）**：品項／件數／來客數排除 dudooGuard 補的「肚肚對帳調整」金額調整列（營收照算）；品項月報快取換 `POS_PRODUCT_MONTHLY_V2_`；月表 2025-03 店1＋2025-04 全店 13 個店月重建。回滾＝部署選第 19 版
- `bq_connector.gs`（線上檔名「直連 BigQuery v_daily_net」）：SA 金鑰存於指令碼屬性 `SA_KEY_JSON`，**不在原始碼內**，本備份可公開
- 修改 SOP／地雷／迴歸錨點一律見 `dashboard-pos` skill

| 檔案 | 說明 |
|---|---|
| `discount_aggregator.gs` | 主聚合（kpi／monthlyByStore／品項排行／daily_by_store／product_monthly） |
| `posRolling.gs` | 滾動視窗：月表 `agg_pos_month`／`agg_pos_item_month` 掃描與歷史合併層 |
| `product_monthly.gs` | 品項月報（線上檔名 `product_monthly.gs`，與主檔同名函式重複，兩邊要一起改） |
| `seasonal_detail.gs` | 檔期 detail 立方體 |
| `fix_posMonthAgg_20261006.gs` | 一次性：月表 13 個店月重建（已執行）。**公開版已移除內嵌的各店月銷售資料**，不可執行 |
| `bq_connector.gs` | BigQuery 直連 |

線上另有 `diag_tmp`、`_pr_setup_tmp` 兩個舊暫存檔，未備份。

本資料夾最早為 2026-07-22 自 Drive 匯出＋v19 部署版快照；2026-10-06 依線上 @20 同步。**日後每次改 GAS 部署後，須同步更新本資料夾**（鐵則 5）。
