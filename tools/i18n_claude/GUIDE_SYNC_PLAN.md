# 交辦書：到店須知（客人掃 QR 的頁面）甜點清單跟平板外語食譜自動同步

建立：2026-11-26（雲端 Code 規劃；雲端沒有 clasp 登入、secrets、官網連線，**要在 Mac 本機 Code 執行**）

## 經營者要的
- 客人掃 QR（`static/store-guide.html`）看到的甜點，要跟平板外語食譜一樣多（目前平板約 89 款，QR 頁只有 26 款）。
- 以後平板多一款，QR 頁自動跟著多，不用人記得去改。
- 2026-11-26 經營者選「再加完整介紹」：多出來的甜點也要有 7 語的主體／內餡／裝飾介紹、售價、尺寸，跟原本 26 款一樣完整。

## 現況（2026-11-26 雲端查到）
- QR 頁資料＝「外國客人到店須知 API」GAS（`gas/store-guide/Code.gs`，部署 `AKfycbzKmiZv95…`），每天 23:00 只讀官網 **English Menu**（`MENU_URL` 類別 6c236c58…），所以只有 26 款。
- 2026-11-26 13:19「體驗中心按鈕」手動同步失敗：`Service invoked too many times: translate`（LanguageApp 額度）；客人目前看到 11/25 23:02 那一版 26 款。
- 平板清單＝申請單 GAS `gas/recipe-req/i18n.gs` 的 `i18nListData_()`（要平板密碼）；每款有 id、名稱、封面、各語言品名（`local`）、類別（`cats`）。**沒有售價、尺寸、產品特色**。
- 平板反過來已經會讀 QR 頁的品名（`i18nGuideNames_`），所以兩邊品名要一致。

## 做法（照順序，每步驗完才做下一步）
0. PRE-CHECK：`git pull`；兩個 GAS 專案都先 `clasp pull` 抓線上整份跟 repo 比對（線上為準）；備份 store-guide 試算表的 items／i18n 分頁（記筆數）。
1. **申請單 GAS（recipe-req）新增公開唯讀動作 `guideList`**（NOAUTH＋READS，只給品名／照片／售價／介紹，不給食譜步驟）：
   - 名單＝`i18nListData_()` 的每一款（平板有的才列）。
   - 每款對到台灣**繁中啟用中**那支（中文原文的就是自己；舊英文原文的用「英文名｜中文名」或 kuan 對），從後台編輯頁讀 `Price`、`Size`、`Content`（產品特色）、封面。
   - 每天 06:30 `i18nDaily` 後順便更新，存新分頁 `guide_meta`（id、中文名、英文名、各語言品名、售價、尺寸、產品特色原文、雜湊、cats、更新時間）；`guideList` 只讀這個分頁（快）。
2. **到店須知 GAS（store-guide）`sync_` 改成兩個來源合併**：
   - 官網 English Menu 有的 26 款：照舊（完整資料、人工譯文不動）。
   - 平板有、English Menu 沒有的：從 `guideList` 補進來；英文名＝平板英文品名，照片＝封面，售價／尺寸＝後台；葷素、含酒、過敏原等標記從中文產品特色解析。
   - 介紹譯文讀 i18n 分頁 `how='claude'` 的列（鍵＝品名，src＝中文產品特色雜湊）；**這些款不用 LanguageApp**（經營者 11/23：外語食譜不用 API）。還沒翻好的照樣列出（品名用平板已翻好的 7 語，介紹先空），狀態列「待翻介紹」。
   - 調整 `check_`：「款數驟減一半」照留；介紹空白對「待翻介紹」的款不算錯；照片檢查一次 89 張要分批（fetchAll 每批 ≤30）。
   - 新增寫入動作 `putIntro`（投遞通行碼放指令碼屬性，Claude 從 secrets.env 讀，不回顯）：寫入 how='claude' 譯文。
3. **翻譯（Claude Code，不花 API）**：`tools/i18n_claude/` 加 `guide_pull.py`（列出產品特色有改或還沒翻的款）與 `guide_put.py`（投遞＋讀回）。翻譯員／回讀員規則照 `TRANSLATE.md`、`REVIEW.md` 的「中文原文」段：逐字照台灣中文、品名照後台英文名詞典、不自行增改。分成主體／內餡／裝飾／備註四欄（照原 26 款格式）。
   - 併進週一到週五 10:00 排程 `weekday-recipe-i18n-check`：平板食譜翻完後順便跑 guide_pull，有就翻完。
4. **QR 頁 `static/store-guide.html`**：89 款太長 → 加類別按鈕（沿用平板 `GROUPS` 的 7 語大類名稱，固定字、不呼叫翻譯）；其他畫面不動。總部預覽（`tablet-recipe-review.html` 到店須知分頁）的款數與「待翻介紹」清單一起顯示。
5. **驗收**：
   - QR 頁款數＝平板清單款數（扣掉經營者指定不給外國客人選的，若有）；原 26 款內容與改前逐款相同。
   - 抽 10 款售價／尺寸跟後台一致；7 語介紹逐款有字；照片全部 200。
   - 選 3 款（含壽星款）走到結帳頁，合計正確；陪同費規則照舊。
   - 隔天 23:00 排程跑完 `status.result=ok`，log 有新增筆數。
   - 手機實際掃一次印好的 QR 卡（網址不變，卡片不用重印）。
6. 收尾：GAS 都是「編輯現有部署 → 新版本」，改完同步回 repo `gas/`；更新對應 skill（recipe-cost／tablet 外語食譜相關段落）與 RUNBOOK；回報附前後款數。

## 不用動
- 印好的 QR 卡與指示牌：網址不變，不用重印。
- 官網 English Menu 那 26 款的人工譯文。
