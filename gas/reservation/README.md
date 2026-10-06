# 訂位資料管線（GAS「預約系統自動抓資料」）備份

- 綁定在「DIYBC 訂位資料」試算表；script id `1REpnj4Ep5zjRUeVPxnhp72JOMU8eaIcAPFLMbGfWNp75SKKRUJja4bLz`
- 網頁應用程式（訂位儀表板、決策中心讀這支）：部署 `AKfycbzjvgoj6MQoDizE-322wOLyOvq4P3HVbOlvyHBszdyfuOrYPaSnDa48vQb_34LLgEJh`（2026-10-06 起 @4）
- 本資料夾＝2026-10-06 線上 HEAD 全部檔案（程式碼.gs、forecast_pipeline.gs、guardian_v2.gs、appsscript.json）。`fetchRange_retry.gs` 是 9/29 只備份那一段的舊檔，以 `程式碼.gs` 為準。
- 改版一律先 clasp pull 線上版為底；時間驅動的觸發器跑的是 HEAD（存檔就生效），網頁讀取（doGet）要更新部署版本才生效。
- 帳密在指令碼屬性（DIYBC_EMAIL／DIYBC_PASSWORD），不在程式碼裡。細節見 `dashboard-reservation` skill。
