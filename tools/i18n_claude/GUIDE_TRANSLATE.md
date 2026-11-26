# 到店須知（QR Code 頁）甜點介紹：翻譯員／回讀員說明（2026-11-26）

外國客人掃店內 QR Code → `store-guide.html` → 點甜點看介紹。介紹＝門市食譜後台每支中文食譜的「產品特色」與「尺寸」，**逐行**翻成 en／ja／ko／vi／id／th。
工作資料夾 `WORK`＝`/Users/diybc/Desktop/Claude/Code/資料交換/輸出/暫存/i18n_claude/guide/`。

## 免按允許的鐵則
- 看檔案一律用 **Read 工具**；寫檔用 **Write／Edit 工具**。不要用 Bash（不要 cat、ls、grep、python3 -c）。
- 品名詞典用 Google 試算表工具讀（`get_values`，spreadsheetId `10K3j3mzGz8uFlIUQI7PXz11I-3W8aTCVRnA_6jgUpIs`，range `i18n_zitems!A1:H400`；欄位 kind, zh, en, ja, ko, vi, id, th；只看 kind＝item）。工具沒有就略過，照後台英文名的習慣翻。

## 翻譯員
輸入：`WORK/todo_<語言>.json`＝`{行代號: 中文一行}`。對照用：`WORK/names.json`＝`{中文食譜名: {en, ja, ko, vi, id, th}}`（平板已放行的外語品名）、`WORK/ctx.json`＝`[{zh: 中文食譜名, content: 整段產品特色}]`（看上下文用）。`src.json` 太大，不要讀。
輸出：`WORK/tr_<語言>.json`＝`{行代號: 譯文}`，**todo 裡每個代號都要有**，不多不少；值是一行字串（不可有換行）。

規則（經營者 2026-11-26：「翻譯唯一標準＝台灣區中文食譜，逐字逐句翻，不可自己發想增改」）：
1. 逐字逐句照中文翻：中文寫的不刪不改、中文沒寫的不加（不補說明、不補括號解釋、不合併、不拆句）。看不懂或中文明顯有誤，照字面翻，並在回報列出代號與原因。
2. 數字、價格（$199、69 折＝照當地習慣寫「31% off」之類要看清楚：**69 折＝付原價的 69%**）、日期（11/26~12/29 照原格式）、時間（11:00）、尺寸（cm、×、*）、數量一律照原字，阿拉伯數字（泰文不可用泰文數字）。
3. 符號照抄、位置不動：● 🏠 🎁 ⭐ ❗ ⚠ ※ ・ ① ② 《》【】［］｜ / ＋ 等 emoji 與括號；項目名稱「●主體：」「●內餡：」「●裝飾：」「●上層：」「●尺寸約：」譯成當地語言的短標題、冒號保留。
4. 甜點名：行內提到的甜點（例「奧利奧：4吋｜加購價 $199」「選「另一半」」）用 names.json 該甜點的該語言品名（找中文名相同、或去掉括號／檔期標籤後相同的那支；去掉品名裡的全素等括號註記只留品名本身）；找不到才自己翻，回報列出。
5. 食材、工具、包裝名：照詞典（kind＝item 的 zh → 該語言）；詞典沒有的照常用說法。克林姆＝Whipping cream（ja 生クリーム、ko 생크림、vi Kem tươi、id Whipping cream、th วิปปิ้งครีม）。鮮奶油照詞典。
6. 店名固定寫法（全部語言都用這組拉丁字，客人拿給店員看）：南京 Nanjing、士林 Shilin、信義 A13 Xinyi A13、新店 Xindian、板橋 Banqiao、中壢 Zhongli、桃園藝文 Taoyuan Arts Center、新竹文化 Hsinchu Culture、台中精明 Taichung Jingming、台中草悟道 Taichung Calligraphy Greenway、台南 Focus Tainan Focus、高雄 SKM Park Kaohsiung SKM Park；行內只寫「信義、高雄」＝Xinyi A13、Kaohsiung SKM Park；吳寶春自己做＝Wu Pao Chun × DIYBC；自己做（品牌）＝DIYBC。
7. 尺寸行（例 6吋、4吋×4個、35片、5吋加高、18X10cm (2個)）：照字翻成當地寫法（en「6-inch」「4-inch × 4」「35 pieces」）。
8. 標籤固定譯法要整份一致（同一個中文標籤每行都同一個譯法）：［蛋奶素］［奶素］［蛋素］［全素］［含酒］［含酒，可不加］［⚠ 過敏原｜花生／堅果］［預製品=現成品］。
9. 語氣：en 簡單清楚；ja です・ます；ko 존댓말（-세요／-습니다）；vi、id、th 簡短禮貌。佔位字（例 OOO）照抄。
10. 寫完自己再逐行對一次中文：數字個數、符號、項目有沒有漏。回報：寫了幾行、自己翻的甜點名／店名、看不懂的代號。

## 回讀員（沒參與翻譯的另一位）
讀 `todo_<語言>.json`、`tr_<語言>.json`、`names.json`、`ctx.json`，逐行對照中文，寫 `WORK/review_<語言>.json`：
`{"verdict":"pass"|"fix", "issues":[{"k":"行代號","why":"錯在哪（中文）","fix":"改好的整行譯文"}]}`
只抓真錯（漏譯、多加、意思錯、數字／符號／日期不符、甜點名或店名沒照規則、同一標籤前後不一致、還有中文字（ja 漢字除外）、69 折算錯）；用字偏好不算錯。沒有錯 → `{"verdict":"pass","issues":[]}`。
