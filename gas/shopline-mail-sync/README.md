# diybc-shopline-mail-sync（Shopline 門市訂單 → 採購系統 fact_shopline，每天自動）

- Apps Script 獨立專案：`diybc-shopline-mail-sync`（script ID `1r2-X3ouI20eItCEKyQ5UkWVisfG1KvhepONrCOsjytp8GwH9ULBxaUID`），帳號 mydiybc@gmail.com。
- 每天 06:00（台北）`slmDaily`：讀 Gmail「[新訂單]」通知信（唯讀，Gmail 進階服務）→ 門市單（收件人含「N店」）→ 只新增 fact_shopline 還沒有的訂單。2026-11-05 後自動刪排程；11-04 寄提醒信「請 Claude 做 Shopline 最後補單」。
- 每次執行寫一列到月檔 `shopline_sync_log`（數字、單號、品名，不含顧客資料）。
- 已知限制（2026-10-07 實測）：出貨中心下單後才加的品項／改的數量不會反映（漏約 15% 金額）；總部後台代建的訂單（大貨到店）不寄信，抓不到。到期前用 Shopline 後台 API 整批補最終版本（見 dashboard-purchase skill）。
- 手動函式：`slmDryRun`（只算不寫）、`slmRunNow`（立即補跑）、`slmCheck`（拿既有資料比對）、`slmDiag`（只看一封信的品項區）、`slmSetup`（重裝排程）。
- `slmCatalog.js`：Shopline 商品目錄（2026-10-05 匯出）去空白鍵 → [正式品名, 分類]，由 Claude 產生。
- 改程式：本資料夾 `clasp push`（.clasp.json 不進版控，用 `clasp clone <scriptId>` 取得）。
