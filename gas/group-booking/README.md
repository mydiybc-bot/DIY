# 團體訂位追蹤（GAS）

- 專案：「DIYBC 團體訂位」（綁試算表 `1NrMGw6BOvMJigKfdAFNM1Np6KCa5cd6SEIBqzrmFLAQ`），script id `16eniSjM-zP1DucmOlmlyOG6zjHp6AgJp1sod6cuCmaPE_0g2dlvK_8eI`
- 網頁應用程式部署（訂位儀表板團體分頁、決策中心讀這支；只更新這個部署，不新增）：`AKfycbwCisp7PrQFsu_WJsv9_KNsEYAs2Ki_B13WSBFrjNsxq8fREamX27D6IABk26Iy5E9FpQ`
- 線上檔名是「Group Booking」，這裡存成 `group_booking.gs`。改版：`clasp pull` 抓線上版 → 改 → `clasp push` → `clasp create-version` → `clasp update-deployment <上面部署 ID> -V <版本>`。
- 2026-10-07 第 4 版＝v1.4（客服門市卡位不算問題；一般訂位散客不必先選甜點）。第一次備份進 repo。
- 2026-10-07 第 5 版＝v1.4.1（經營者更正：8 人以上（含）都算團體、都要先選甜點；門市卡位規則不變）。
