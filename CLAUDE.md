# DIY — 自己做烘焙聚樂部 營運系統

Flask 單一服務（`server.py`），部署在 Render，同時提供：員工口語訓練站、10 個營運儀表板（純前端 HTML，讀 Google Apps Script JSONP API）。

## 專案結構
- `server.py` — Flask 入口（路由、靜態檔、訓練站 API）。核心檔，改之前先提醒。
- `static/hub.html` — 儀表板主入口（10 張卡）。
- `static/dashboard-*.html`、`pos-dashboard.html`、`zijiren.html`、`google-reviews.html` — 各儀表板，**每頁自包含**（HTML+CSS+JS 同一檔，React/Recharts 走 CDN）。
- `static/index.html`、`static/app.js`、`static/styles.css` — 訓練站前端。核心檔，改之前先提醒。
- `gas/` — Google Apps Script 原始碼備份（生產版在 Google 端；改完 GAS 部署後必須同步這裡）。
- `training_*.py`、`*_store.py` — 訓練站後端邏輯與資料存取。
- `test_*.py` — 測試。

哪個儀表板對應哪個檔、哪個 skill、哪些檔已廢棄：見 `dashboard-overview` skill。動任何儀表板之前先讀它對應的 skill。

## 常用指令
- 本機啟動：`python3 server.py`（預設埠見 server.py）
- 測試：`python3 -m pytest test_training_logic.py test_integration.py -q`
- 改 HTML/JS 後：本機 `python3 server.py` 開頁實際點過一遍，看瀏覽器 console 有無紅字。
- 部署：推到 `main` 分支後 Render 自動部署（約 1–3 分鐘）。驗證用 `curl -sI https://diybc-training.onrender.com/static/<檔名>` 看 200 與 content-length 是否變動。

## 工作流程（每次任務）
1. **先計畫**：讀對應 skill → 讀要改的檔 → 列出「要改哪幾個檔、哪幾段、為什麼」→ 等我確認再動。
2. **改**：一次一個問題。完整輸出，不片段。
3. **驗**：重新讀檔對照；能跑的就跑（pytest / 本機啟動 / curl）；改 A 有沒有弄壞 B。
4. **推**：`git add` 指定檔 → commit 訊息寫清楚改了什麼 → 第 3 步驗證通過就直接 push main，不必問我；push 後用 curl 確認線上已更新。核心檔案先報備、禁 force push 等紅線照舊。
5. **收尾**：更新對應 skill 的「最後動過」與「最近改了什麼」（commit 編號、函式名）。

## 全站 UI 慣例（維護時不得移除）
- 每頁頂部資料狀態列（`.dsb`）
- 字級 ≥ 12px
- 「⬅ 回首頁」返回鍵
- 密碼閘門（`dim_auth`，各頁 `DASH_ID` 不同）
- 寫入類按鈕一律兩層 confirm

## 紅線
- 不動未指定的檔案。需要動到範圍外 → 停下問。
- 廢棄檔不改（清單見 `dashboard-overview` skill）。
- 資料表寫入類改動先備份、回報前後筆數。
- 其餘紅線與座標在 `~/.claude/rules/diybc-context.md`（本機）。
