# dudooPOS_自動匯入.gs（只在 Google 端，未放進公開 repo）

專案「一鍵追加新品」（綁 BOM 本，script ID `1PkavxV6r3GZbEbS_B5ZWGL1ovm30T55XhllAuI1aglBe_RPERdP-DJnJ`）的每日 04:50 匯入程式。
原檔含通知信箱，所以不放公開 repo；完整改前／改後版本在經營者本機 `資料交換/輸出/gas-backup/一鍵追加新品/`。

## 2026-10-05 唯一改動（第 152 行，`dudooPOS_loginAndImportDate_` 內）

```js
  var allRows = dudooPOS_parseCSV_(content);
  allRows = dudooGuard_apply_(allRows, dateStr, token);   // 2026-10-05：跨日退款改記退款日＋剔除匯出檔幽靈品項，對齊肚肚業績概況（dudooGuard.gs）
```

`dudooGuard_apply_` 在 `gas/bom/dudooGuard.gs`；它出錯不會擋匯入（照原資料寫入並寄信）。
回滾：刪掉這一行即可（dudooGuard.gs 留著不影響任何東西）；每日 07:30 對帳觸發器 `dudooCheck_daily` 在「觸發條件」頁刪除。
