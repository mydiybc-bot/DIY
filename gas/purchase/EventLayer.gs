// ================================================================
// EventLayer.gs — 節慶事件層 v1.9（2026-10-01）
// v1.9：★ 新品申請表「預估銷售數」→ 開賣前首批量與檔期總需求（經營者「20261001 待處理事項」#6）
//   經營者裁示：預估銷售數＝「整個檔期、12 店合計份數」。
//   ① 每天唯讀「食譜系統_申請單」fact_recipe_req：狀態＝已核准／已核准（採購寫入失敗）／已寫入採購、預估銷售數 > 0、檔期不是封存
//      → plan[甜點名]＝{qty, 品牌別}（同名取最新一張）。讀不到 → plan 空、全部照原算法（try/catch）。
//   ② 拆店：品牌別「自己做」＝1–10 店、「吳寶春自己做」＝11、12 店、「自己做＆吳寶春自己做」＝12 店；
//      份額＝去年同檔期各店份數（有的話）→ 否則近 28 天各店甜點份數 → 否則平均。本店整檔推估 T＝預估×份額。
//   ③ 非曲線模式（開賣前、或沒有去年曲線）且有預估的甜點：首批目標＝T×首批比例（同原級距 20–70%）−已售；
//      開賣滿 7 天 POS 後照原規則歸零、交給實際用量（近 4 週均）。曲線模式（已開賣且有去年曲線）照原算法＝實際銷量推估下週。
//   ④ agg_evneed 開賣前也寫（有預估的甜點）：推估總份數＝max(T, 已售)、剩餘＝推估−已售 → 檔期備貨頁開賣前就看得到 12 店總需求。
//   沒有任何已核准預估時，所有輸出與 v1.8.5 逐字相同。
// ================================================================
// ================================================================
// EventLayer.gs — 節慶事件層 v1.8.5（2026-09-23）
// v1.8.5：★ 單位換算改成與主引擎（程式碼.gs 的 _conv）完全同一套規則
//   問題（中秋進貨健檢發現）：toUse／perUse 遇到「個 包 顆 滴 片 碗…」這類計數單位時，
//     直接當 1 回傳，沒有先查 dim_unitconv。結果：
//       色膏 BOM 寫「滴」（1 滴＝0.05 g）→ 被當 1 g，agg_evneed／agg_newitem 放大 20 倍
//       栗子 BOM 寫「顆」（1 顆＝15 g）→ 少算 15 倍；原味麻糬丁「包」（1 包＝20 g）→ 少算 20 倍
//     共 22 組「料｜單位」受影響（色膏 11 款、麻糬、麻糬丁、大甲芋泥餡、地瓜、起司絲、玫瑰花瓣、藍莓、栗子、荔枝玫瑰餡、鳳梨花乾、蓮花餅乾）。
//     主引擎 agg_purchase 一直是對的，所以「近4週均」與「檔期增量」兩邊單位不一致。
//   修法：新增 evConvFactor_，查找順序＝BOM 原名 → 主檔品名 →（通用）→ 1，與 _conv 逐字同規則；
//     toUse（agg_newitem／agg_evneed）與 perUse（agg_endcut）都改用它。計數單位查不到換算時仍以 1 計、不提醒（與主引擎相同）。
//   不影響：dim_event、agg_evcurve（只用份數）、器具模具需求（用 BOM 數量，不經換算）、表結構與欄位。
// ================================================================
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.8.4（2026-09-17）
 * v1.8.4：agg_endcut 加第 10 欄「近7天比」＝該店該料 POS 近 7 天用量 × 4 ÷ 近 28 天用量（上限 1，與收尾同一 BOM 口徑）
 *   原因（10 輪檢查第 10 輪上線驗證）：前端用 agg_purchase「上週÷近4週均」驗證已下架批次，但主引擎的上週用量比 POS 實況落後
 *     （1 號店杏仁片：上週 6／近4週均 21＝29%，POS 近 7 天 18／週均 22＝82%）⇒ 仍有 6 個店×料照常在用卻被扣 31～43%。
 *   前端改取 max(主引擎比例, 近7天比)，只會少扣；讀不到這欄（GAS 未更新）時自動退回只用主引擎比例。其餘與 v1.8.3 零改動。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.8.3（2026-09-19）
 * v1.8.3：「在售甜點」改為逐列判斷（上線後 10 輪檢查第 5 輪：中秋收尾 × 萬聖開賣重疊）；
 *         收尾計算包 try（第 8 輪：出例外不再讓整個重算中斷，agg_endcut 寫空＝前端不扣）
 *   問題：v1.8.1 的在售甜點＝「今天有效、且不在下架觀察窗（今天 +35 天）內」的甜點。
 *     ① 萬聖 10/1 才開賣 ⇒ 9/26 不算「今天有效」 ② 萬聖 11/4 下架在觀察窗內 ⇒ 10/3 也被排除
 *     ⇒ 包裝盒這類「本店近 4 週只有中秋在用、萬聖也要用」的料，被判貨到已結束、亮紅字「不建議叫」（建議量本身有萬聖備料撐著，但訊息誤導）。
 *   新判定（每一列各算）：這一列的下架日為 E，在售甜點＝用這個料、且
 *     「最早可賣日 ≤ 今天 + LEADIN_DAYS（14 天內會開賣或已在賣）」且「最晚下架日 > E + ENDCUT_OA_GAP_DAYS（7 天）」的甜點。
 *     中秋甜點（9/29）對中秋那列不算；萬聖（10/1～11/4）對中秋那列算；2099 的常態甜點永遠算。其餘與 v1.8.2 零改動。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.8.2（2026-09-18）
 * v1.8.2：★ 上線後實查修正——agg_endcut 每次寫完都設定欄位格式（BOM 本與專用檔都設）
 *   事故：v1.8.1 在舊表（7 欄）右邊插入 H、I 兩欄時，Google 試算表自動沿用左邊「更新日」欄的日期格式，
 *         「份數佔比」0.088 被顯示成 1899/12/30，gviz 讀出來是日期（連小時都被丟掉）⇒ 前端讀到 0
 *         ⇒ 替代原則失效、等於退回 v1.8 整份扣（1 號店擠花袋被扣 13%、廚房紙巾 17%）。
 *   修法：evFormatEndcut_ 每次設 結束日／更新日＝yyyy-mm-dd、佔比／份數佔比＝0.000、近4週週用量＝0.0、甜點／在售甜點＝純文字。
 *         只改顯示格式、不改儲存值（0.088 存成日期時底層仍是 0.088），重算後立即恢復正確。其餘與 v1.8.1 零改動。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.8.1（2026-09-17 晚）
 * v1.8.1：agg_endcut 加一欄「份數佔比」＝該店近 28 天甜點總份數中，這批下架甜點佔幾成（店層級、同結束日同值）
 *   原因（經營者要求再三檢查時發現）：v1.8 把下架甜點的用料整份扣掉，但客人不會因為檔期結束就不來，
 *   會改做其他甜點 ⇒ 擠花袋、廚房紙巾、包裝盒、奶油這類「大家都用」的料，用量大多會被其他甜點接走。
 *   實測 1 號店擠花袋被扣 13%、廚房紙巾 17%、7 號店奶油 6%——檔期後反而會缺貨。
 *   前端改用「扣掉會被替代的部分」：實際少算＝max(0, (用料佔比 − 份數佔比) ÷ (1 − 份數佔比))。
 *   專屬料（用料佔比 100%）照扣；一般耗材（用料佔比≈份數佔比）幾乎不扣；只在下架甜點特別用得兇的料才扣。
 *   ＋「在售甜點」欄（上線前模擬第 1 輪發現）：這個料還有哪些「對照表今天有效、且不在下架窗內」的甜點在用（最多列 3 款）。
 *     例：卡士達粉在 11、12 號店近 4 週只被夏季甜點用過，但鐵觀音達克瓦茲還在菜單上 ⇒ 前端不可判「貨到已結束，不建議叫」。
 *   ＋瘦身（上線前模擬第 4 輪發現：表每次開頁要下載約 560 KB，是第三大的表）：
 *     ①「影響不到 2%」的店×料整組不寫（擠花袋、砂糖、雞蛋這類替代後幾乎不扣的）②「在售甜點」只寫給專屬料
 *     （合計用料佔比 ≥95%，前端只在這種情況用得到）③前端只讀 7 欄（不讀近4週週用量、更新日）⇒ 約 200 KB。
 *   其餘與 v1.8 零改動；表頭每次重寫（舊表只有 7 欄）。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.8（2026-09-17）
 * v1.8：★ 新表 agg_endcut「收尾扣減」——根治「每個檔期最後一週照常叫貨、檔期後共用料多叫一個月」
 *   問題：前端訂貨點＝近4週均 × 涵蓋週數（1 週＋交期），完全不看甜點何時下架。
 *     ① 專屬料：最後一週仍照近4週均補 10 天份，貨到時檔期已結束（2026 中秋 9/26 叫貨估多叫約 $7,100／12 店）
 *     ② 共用料（奶油、白巧克力…）：檔期結束後近4週均還含 3 週檔期用量，接下來一個月持續多叫（中秋估約 $22,600）
 *   做法：每天用 POS 近 28 天 × BOM，算出「各店各料的用量中，有幾成來自即將下架／剛下架的甜點」，
 *     連同該甜點的結束日寫 agg_endcut。前端算訂貨點時，這幾成用量只算到結束日（見 dashboard-purchase bx 批）。
 *   甜點結束日＝產品名稱對照表「結束有效日」（同名多列取最晚；有任一列延續到觀察窗之後＝不算下架）
 *     觀察窗：結束日在「今天 −28 天 ～ 今天 +ENDCUT_AHEAD_DAYS 天」之間
 *     佔位日期保險：結束日已過超過 ENDCUT_TAIL_DAYS 天、POS 卻還在賣 ⇒ 視為對照表日期沒更新，不扣，log 提醒
 *   與檔期類別無關：一般甜點下架（對照表寫了結束日）同樣適用 ⇒ 以後任何檔期、任何下架都自動處理
 *   器具／模具不算（走標配）；不影響 agg_newitem／agg_evneed／agg_evcurve／dim_event（零改動）
 *   其他：鏡像補「列數不足先加列」保險（新表可能超過預設 1000 列）
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.7.1（2026-09-16 晚）
 * v1.7.1：兩件事，其餘與 v1.7 零改動
 *  ① 「專屬」判定改用 sku 編號比對（修「雞蛋被標中秋節·專屬」）
 *     事故：噠噠馬德蓮 BOM 寫「雞蛋 (請打到塑膠碗)」，其他 8 款一般甜點寫「雞蛋」，兩個寫法在 dim_sku 都是 F025，
 *           但舊判定拿 BOM 原字串比對 ⇒ 認不出是同一個料。
 *     新判定：其他產品用料以 sku_id 比對（對不到 dim_sku 的才用原字串）；「其他產品」＝
 *           (a) 對照表今天有效的列（手動對應空白時用 POS 品名，舊版會漏掉）
 *           (b) ∪ 近 EXCL_RECENT_DAYS 天 POS 有賣、且不屬於本檔期的甜點（對照表結束日過期但還在賣的保險）
 *     專屬影響：前端「🎯專屬」徽章與篩選、檔期備貨頁「已出貨」回推天數（專屬 90 天／共用 15 天）
 *  ② agg_evcurve「今年推估」改為「實際＋推估」：已過的日子＝POS 實際已售；今天起到對照表結束日＝推估
 *     （推估基準＝窗內已售 ÷ 去年同位置累積佔比，與各店合計同口徑）。整欄合計＝推估總量、合計−已售＝還沒賣
 *     背景：舊欄位是「推估總量 × 去年該週佔比」，已過的週也顯示模型值（8/14 週已售 327 卻顯示推估 160），
 *           看起來像沒跟著實際銷售調整。其實推估總量每天都有重算（9/5 3,574 → 9/16 3,803），只是欄位語意讓人誤會
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.7（2026-09-16）
 * v1.7：★ 各店各品項改用「自己的銷售」推估剩餘量（修 3 號店冰皮粉事故）
 *   事故：2026 中秋 3 號店冰皮月餅到 9/15 已賣 52 份，系統推估整檔 53 份 ⇒ 剩 1 份，還叫它調出 2,116 g。
 *   根因：v1.6 各店推估＝全公司推估 × 該店「全部限定品」份額 × 品項份額（假設每家店賣各品項的比例一樣）。
 *         3 號店每個料都被分 15.0%，但冰皮月餅實際佔全公司 22.4%；10 號店分 3.7%、實際 1.3%。
 *         再加上「剩餘＝推估−已售」，賣得比分配額快的店剩餘直接歸 0。全表 756 列有 49% 偏差 >30%。
 *   新算法（每天 07:31 用 POS 到昨天的資料重算，逐店逐品項）：
 *     A 累積速度＝該店該品項「窗內已售 ÷ 去年同位置累積佔比」（從該品項全公司首賣日起算）
 *     B 近 7 天速度＝該店該品項「近 7 天已售 ÷ 去年同位置 7 天佔比」（抓越賣越快的店），B 最多 A 的 PACE_B_CAP 倍（防包館單日爆量）
 *     等效整檔份數 R ＝ (A＋B)÷2；近 7 天 0 份但之前有賣 → 只用 A（不讓缺貨變成「不用備」）；只有一邊能算就用那一邊
 *     剩餘份數 ＝ R × 去年「今天～該品項檔期結束日」的佔比（★ 截在對照表結束日，檔期後的量不算）
 *     下週份數 ＝ R × 去年「今天起 7 天（不超過結束日）」的佔比 → agg_newitem 增量＝max(0, 下週 − 近28天均週份數)（增量語意不變）
 *     推估總份數 ＝ 已售 ＋ 剩餘份數（剩餘不再被已售吃掉）
 *     新品保底：這家店還沒賣過、且該品項全公司開賣 ≤ PACE_NEW_DAYS 天 → 暫用 v1.6 拆分（避免新上架品被 0 卡死）
 *   agg_evcurve「今年推估」也截在檔期結束日（結束日之後的天數不計）；表結構、欄位、鏡像、dim_event、
 *   開賣前首批級距、器具/模具日尖峰邏輯 全部零改動。前端不用改程式（只換說明文字）。
 *   Log 新增：POS 檔期銷量最新日期（不是昨天會警告）、各店合計 vs 全公司曲線、各方法筆數。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.6.2（2026-09-05）
 * v1.6.2：鏡像前檢查專用檔時區＝Asia/Taipei（新檔預設美國時區會讓日期早一天）
 * v1.6.1：分檔第二步——dim_event / agg_newitem / agg_evcurve / agg_evneed 寫完後整張鏡像到儀表板專用檔 EV_DASH_ID
 *         （前端已改讀專用檔）。鏡像失敗只記 Log，不影響 BOM 本。其餘與 v1.6 零改動。
 *****************************************************************/
/*****************************************************************
 * EventLayer.gs — 節慶事件層 v1.6（2026-09-05）
 * v1.6 變更（只動 agg_newitem 的「開賣後」路徑與檔期起始日校正；dim_event 節日週、係數演算法、
 *            開賣前首批級距、器具/模具日尖峰邏輯 零改動）：
 *  A) 進行中檔期的起始日改以 POS 實際首賣日校正（原本只校正已結束檔期）。
 *     背景：對照表起訖日多為佔位值（01-01～04-01 之類），2026 中秋對照表寫 8/20、實際 8/15 開賣。
 *     規則：開賣滿 5 個有銷售日才校正；結束日仍以對照表為準（還沒發生，POS 推不出）。
 *  B) ★ 去年檔期逐週曲線（全公司口徑）：
 *     去年主段逐日銷量 → 以「節日當天」為錨點（EV_FEST 有登錄者），否則以「主段結束日」為錨點，
 *     算出每一天佔全檔期的比例。主段＝第一段連續銷售，遇到 ≥GAP_DAYS 天零銷售即視為結束
 *     （2025 中秋 10/16 起斷 5 天，10/22~11/02 是半價清庫存 392 份，不算需求）。
 *  C) ★ 開賣後改「曲線推估」取代「7 天歸零交棒」：
 *     今年全公司總量 ＝ 至今已售 ÷ 去年到此為止的累積佔比
 *     各店份額 ＝ 今年至今各店實售份額（開賣滿 RESPLIT_MIN_DAYS 天）；否則用去年各店份額
 *     各品項份額 ＝ 既有 share[]（零改動）
 *     ⚠️ 需求量寫的是「增量」：max(0, 曲線推估下週份數 − 該品項該店近 4 週均週份數) × BOM 用量。
 *        前端訂貨點＝近4週均推出的量＋agg_newitem 需求量（相加），近4週均已含的部分不能再給一次，
 *        否則就是 8/31 退掉「逐週係數」的同一個重複計算。峰後曲線低於近4週均時需求量＝0（引擎本身會跟著降）。
 *     啟動條件：有去年樣本、去年主段 ≥ 14 天、開賣滿 RESPLIT_MIN_DAYS 天、累積佔比 ≥ CURVE_MIN_CUM。
 *     不滿足 → 走 v1.5 原路徑（首批級距／7 天歸零），一字未動。
 *  D) ★ 新增兩張小表給總部／出貨中心用（前端下一批接）：
 *     agg_evcurve：每檔期逐週（距節日週序）去年份數／佔比／今年已售／今年推估
 *     agg_evneed ：每檔期 × 店 × sku 的「檔期剩餘推估用量」「已用量」「已售份數」「推估總份數」
 *     兩表皆整張重建、建表時裁到 ≤300 列 × 表頭欄數（不吃 BOM 本 26 欄 × 1000 列的預設配額）。
 *  E) 新增唯讀 aaaEvCurvePreview：列出進行中檔期的曲線與推估，不寫任何表。改本檔前後先跑它。
 *  ★ 部署鐵則：改完任何 .gs 都要「部署 → 管理部署作業 → 編輯 → 新版本」。
 *    總覽頁「🔄 更新」走網頁應用程式版本，沒部署新版本就是跑舊碼（9/4 事故）。
 *****************************************************************/
/*****************************************************************
 * v1.5（2026-08-31）：節日週改「距檔期結束日的偏移」對齊（農曆節日）；無尖峰改寫停用說明列；
 *   尖峰週超出檔期改截在檔期內；新增唯讀 aaaEvPeakPreview。
 * v1.1（2026-07-27）：器具/模具改「日尖峰同時使用」＝單份用量 × ceil(去年週均份數×2÷7)，
 *   檔期內持續有效、不隨已售遞減；前端訂貨點取 max(標配, 檔期需求)。
 * v1（2026-07-14）：
 *  1) 自動維護 dim_event：短尖峰檔期（<30天）係數＝去年同檔期非限定品日均÷檔期前28天日均；
 *     長檔期（≥30天）不設整段係數，只偵測去年檔期內的「節日週」尖峰另寫一列；手動列永不覆蓋
 *  2) 自動維護 agg_newitem：去年同檔期限定品每店實銷為總池；開賣前均分、滿 3 天改實際份額；
 *     首批級距 ≤10天70%／11~29天50%／≥30天20%；需求量＝max(0,首批目標−已售)，7 天 POS 後歸零
 * 掛法：每日 07:31 觸發器指向 rebuildEventLayer（總覽頁「立即重算」亦會呼叫）。
 *****************************************************************/


/* ═══════════════════════════════════════════════════════════════
 * ★ GAS「執行」下拉選單專用入口（2026-08-27）
 * aaa 開頭排最前，即使選單錯位也只會跑到唯讀函式。
 * ═══════════════════════════════════════════════════════════════ */

/** 現況檢查：用 POS資料 目前真實的最舊日期跑。驗收預期＝警告 0 筆。 */
function aaaEvDryRunNow() { return ev2DryRun_(null); }

/** 裁切模擬：假設 POS資料 只保留 2025-03-01 以後（18 個月滾動視窗）。 */
function aaaEvDryRunAfterTrim() { return ev2DryRun_('2025-03-01'); }

/** ★ v1.5 新增（唯讀）：只算節日週會被放在哪裡，不寫任何表。驗收用。 */
function aaaEvPeakPreview() { return evPeakPreview_(); }

/** ★ v1.6 新增（唯讀）：列出進行中檔期的去年曲線、今年已售、推估總量與各店拆分。不寫任何表。 */
function aaaEvCurvePreview() { return evCurvePreview_(); }

var EV_BOM_ID = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';
/* ★ v1.9：新品申請表（食譜系統_申請單，私有、同擁有者）——只讀 */
var EV_REQ_ID = '10K3j3mzGz8uFlIUQI7PXz11I-3W8aTCVRnA_6jgUpIs';
var EV_PLAN_OK = { '已核准': 1, '已核准（採購寫入失敗）': 1, '已寫入採購': 1 };
var EV_BRAND_STORES = { '自己做': [1,2,3,4,5,6,7,8,9,10], '吳寶春自己做': [11,12], '自己做＆吳寶春自己做': [1,2,3,4,5,6,7,8,9,10,11,12] };
function evReadPlan_(log) {
  var out = {}, n = 0;
  try {
    var ss = SpreadsheetApp.openById(EV_REQ_ID);
    var rq = ss.getSheetByName('fact_recipe_req'), cp = ss.getSheetByName('fact_campaign');
    if (!rq || !cp || rq.getLastRow() < 2) { log.push('📋 新品申請表預估銷售數：沒有資料'); return out; }
    function tab(sh, cols) {
      var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); }), idx = {};
      cols.forEach(function (c) { var j = H.indexOf(c); if (j < 0) throw new Error(sh.getName() + ' 找不到欄位「' + c + '」'); idx[c] = j; });
      var last = Math.max.apply(null, cols.map(function (c) { return idx[c]; })) + 1;
      return { idx: idx, rows: sh.getRange(2, 1, sh.getLastRow() - 1, last).getValues() };
    }
    var C = tab(cp, ['campaign_id', '檔期名稱', '品牌別', 'status']), camp = {};
    C.rows.forEach(function (r) { camp[String(r[C.idx['campaign_id']]).trim()] = { name: String(r[C.idx['檔期名稱']]).trim(), brand: String(r[C.idx['品牌別']]).trim(), st: String(r[C.idx['status']]).trim() }; });
    var R = tab(rq, ['req_id', 'campaign_id', '商品暫定名稱', '商品正式名稱', '預估銷售數', 'status', 'updated_at']);
    R.rows.forEach(function (r) {
      var st = String(r[R.idx['status']]).trim(), q = Number(r[R.idx['預估銷售數']]) || 0;
      if (!EV_PLAN_OK[st] || !(q > 0)) return;
      var c = camp[String(r[R.idx['campaign_id']]).trim()] || {};
      if (c.st === '封存') return;
      var nm = String(r[R.idx['商品正式名稱']] || '').trim() || String(r[R.idx['商品暫定名稱']] || '').trim(); if (!nm) return;
      var up = String(r[R.idx['updated_at']] || '');
      if (out[nm] && out[nm].up > up) return;
      out[nm] = { qty: q, brand: c.brand || '', req: String(r[R.idx['req_id']]).trim(), camp: c.name || '', up: up };
    });
    n = Object.keys(out).length;
    log.push('📋 新品申請表預估銷售數：已核准 ' + n + ' 支' + (n ? '（' + Object.keys(out).slice(0, 10).map(function (k) { return k + ' ' + out[k].qty + ' 份'; }).join('、') + '）' : ''));
  } catch (e) { log.push('⚠️ 讀新品申請表預估銷售數失敗（本次照原算法）：' + e); return {}; }
  return out;
}
/* 近 28 天各店甜點份數（不含加價購／入場費／特約／冰淇淋／活動／陪同） */
function evRecentStoreQty_(pos, today) {
  var from = evYmd_(evAddDays_(today, -28)), to = evYmd_(today), EX = { '加價購': 1, '入場&共廚&其他': 1, '特約廠商': 1, '冰淇淋': 1, '活動': 1 }, o = {};
  pos.forEach(function (r) {
    var q = Number(r['數量']) || 0; if (q <= 0) return;
    var d = evDate_(r['建立日期']); if (!d) return; var y = evYmd_(d); if (y < from || y >= to) return;
    if (EX[String(r['主類別'] || '').trim()] || String(r['商品名稱'] || '').trim() === '陪同入場費') return;
    var st = String(r['分店代碼'] || '').trim(); if (st) o[st] = (o[st] || 0) + q;
  });
  return o;
}
/* 本店整檔推估 T＝預估×份額（品牌別限定店；份額＝去年同檔期 → 近 28 天 → 平均） */
function evPlanSplit_(pl, pool, recent) {
  var allow = {}; (EV_BRAND_STORES[pl.brand] || EV_BRAND_STORES['自己做＆吳寶春自己做']).forEach(function (s) { allow[String(s)] = 1; });
  var base = {}, tot = 0, src = '去年同檔期份額';
  Object.keys(pool || {}).forEach(function (st) { if (allow[st] && pool[st] > 0) { base[st] = pool[st]; tot += pool[st]; } });
  if (!(tot > 0)) { base = {}; tot = 0; src = '近 28 天甜點份額'; Object.keys(recent || {}).forEach(function (st) { if (allow[st] && recent[st] > 0) { base[st] = recent[st]; tot += recent[st]; } }); }
  if (!(tot > 0)) { base = {}; tot = 0; src = '平均'; Object.keys(allow).forEach(function (st) { base[st] = 1; tot += 1; }); }
  var t = {}; Object.keys(base).forEach(function (st) { t[st] = pl.qty * base[st] / tot; });
  return { t: t, src: src };
}
var EV_DASH_ID = '1FF7lW3JINR0-Id7MMYSRkoRzA1BYdqktO94NbzAFYG0';   // ★ v1.6.1 儀表板專用檔（與 Code.gs 的 DASH_ID 相同）
var EV_MIRROR_TABS = ['dim_event', 'agg_newitem', 'agg_evcurve', 'agg_evneed', 'agg_endcut'];   /* v1.8 ＋agg_endcut */

// 檔期別名：同一節慶不同年份叫法不同時，在此對應（雙向都列）
var EV_ALIAS = {
  '父親七夕': ['父親節'],
  '父親節': ['父親七夕'],
  '情人節': ['西洋情人節'],
  '西洋情人節': ['情人節']
};

/* ★ v1.6：農曆節日的國曆日期（曲線對齊錨點）。國曆節日不用登錄（同月同日對齊）。
   每年新增一列即可；沒登錄的檔期退回「距主段結束日」對齊。 */
var EV_FEST = {
  '中秋節':   { 2025: '2025-10-06', 2026: '2026-09-25', 2027: '2027-10-15' },
  '父親七夕': { 2025: '2025-08-29', 2026: '2026-08-19', 2027: '2027-08-08' },
  '父親節':   { 2025: '2025-08-29', 2026: '2026-08-19', 2027: '2027-08-08' },
  '端午節':   { 2025: '2025-05-31', 2026: '2026-06-19', 2027: '2027-06-09' }
};

var EV_CFG = {
  SHORT_MAX: 29,          // 檔期天數 <30 ＝短尖峰
  SUPER_SHORT_MAX: 10,    // ≤10天 ＝超短檔期
  FRAC_SUPER_SHORT: 0.7,  // 超短檔期首批比例
  FRAC_SHORT: 0.5,        // 11~29天首批比例
  FRAC_LONG: 0.2,         // ≥30天首批比例
  BASE_DAYS: 28,          // 係數分母：檔期開始前 N 天
  LEADIN_DAYS: 14,        // 開賣前 N 天啟動新品備料層
  TAKEOVER_DAYS: 7,       // 新品累積 N 天 POS 後備料層歸零（v1.6：曲線模式不適用）
  RESPLIT_MIN_DAYS: 3,    // 開賣滿 N 天改用實際份額
  PEAK_WIN: 7,            // 節日週視窗天數
  PEAK_MIN_COEF: 1.15,    // 節日週係數低於此視為雜訊不寫
  PEAK_MIN_DAYS: 3,       // ★ v1.5：節日週截到檔期內後至少要剩幾天才寫
  LOOKAHEAD_DAYS: 90,     // 只處理未來 N 天內會開始/進行中的檔期
  MAX_COEF: 3.0,          // 係數上限保險
  // ★ v1.6 曲線
  GAP_DAYS: 5,            // 去年檔期連續 N 天零銷售 ⇒ 主段結束（之後的視為清庫存）
  CURVE_MIN_DAYS: 14,     // 去年主段至少 N 天才用曲線
  CURVE_MIN_CUM: 0.10,    // 今年累積佔比至少到此才推估總量（太早推估會被前幾天雜訊放大）
  CURVE_MIN_CORRECT: 5,   // 進行中檔期起始日校正：至少 N 個有銷售日
  RECENT_DAYS: 28,        // 增量計算用的「近 N 天均」（對齊前端近 4 週均）
  CURVE_MAX_RATIO: 2.5,   // 推估總量／去年總量 上限保險（避免開賣爆量把整檔推到天上）
  // ★ v1.7 各店各品項速度
  PACE_DAYS: 7,           // 近 N 天速度（B）
  PACE_MIN_CUM: 0.05,     // A 的分母（去年累積佔比）至少要到此才算，避免剛開賣幾天被放大
  PACE_MIN_SHARE7: 0.03,  // B 的分母（去年同位置 7 天佔比）至少要到此才算
  PACE_B_CAP: 2.0,        // B 最多是 A 的幾倍（包館團體單日爆量保險）
  PACE_NEW_DAYS: 7,       // 品項全公司開賣 ≤N 天、這家店還沒賣過 → 暫用 v1.6 拆分保底
  // ★ v1.7.1 專屬判定
  EXCL_RECENT_DAYS: 7,    // 近 N 天 POS 有賣的非本檔期甜點，視為「其他產品還在用」（對照表結束日過期但還在賣的保險）
  // ★ v1.8 收尾扣減
  ENDCUT_AHEAD_DAYS: 35,  // 結束日在今天起 N 天內的甜點列入（涵蓋＝1 週＋交期，交期最長約 3 週）
  ENDCUT_BACK_DAYS: 28,   // 結束日在過去 N 天內的也列入（近4週均還含它們的用量）
  ENDCUT_TAIL_DAYS: 7,    // 結束後 N 天內還在賣＝清尾貨正常；超過還在賣＝對照表日期沒更新，不扣
  ENDCUT_MIN_SHARE: 0.005, // 佔比低於此不寫（雜訊）
  ENDCUT_MIN_EFFECT: 0.02, // ★ v1.8.1：替代後最多只少備不到 2% 的店×料整組不寫（瘦身）
  ENDCUT_OA_GAP_DAYS: 7   // ★ v1.8.3：在售甜點要比這一列多賣至少 N 天才算
};

function rebuildEventLayer() {
  var ss = SpreadsheetApp.openById(EV_BOM_ID);
  var log = ['rebuildEventLayer ' + new Date()];
  var today = evDateOnly_(new Date());

  // ---------- 讀取 ----------
  var map = evReadObjs_(ss, '產品名稱對照表');
  var pos = evReadObjs_(ss, 'POS資料');
  var bom = evReadObjs_(ss, 'BOM表');
  var sku = evReadObjs_(ss, 'dim_sku');
  var conv = evReadObjs_(ss, 'dim_unitconv');

  var plan = evReadPlan_(log), recentQty = evRecentStoreQty_(pos, today);   /* ★ v1.9 */
  var ctx = evBuildContext_(map, pos, today, log);
  var inst = ctx.inst, dayTotal = ctx.dayTotal, instDay = ctx.instDay, instItemStore = ctx.instItemStore, instItemFirst = ctx.instItemFirst;
  var instItemStoreDay = ctx.instItemStoreDay;

  // ---------- 名稱→sku、單位換算 ----------
  var name2sku = {};
  sku.forEach(function (r) {
    var id = String(r['sku_id'] || ''); if (!id) return;
    var u = { id: id, name: String(r['品名'] || '').trim(), useUnit: String(r['使用單位'] || 'g'), cat: String(r['品類別'] || '').trim() };
    var pn = String(r['品名'] || '').trim(); if (pn) name2sku[pn] = u;
    /* v1.6：分隔符補全形「；，」。⚠️ 前端 split 只認 ; , ， 、 ⇒ 資料一律用「、」分隔最安全 */
    String(r['BOM別名'] || '').split(/[,，、;；／\/]/).forEach(function (a) {
      a = a.trim(); if (a) name2sku[a] = u;
    });
  });
  var toUseWarned = {};
  function toUse(nm, qty, unit) { // BOM 量 → 使用單位量（v1.8.5：與主引擎 _conv 同規則，見 evConvFactor_）
    var cf = evConvFactor_(conv, name2sku, nm, unit);
    if (cf.warn && !toUseWarned[cf.key]) { toUseWarned[cf.key] = 1; log.push('⚠️ 未換算單位: ' + nm + ' / ' + cf.unit + '（暫以 1 計）'); }
    return qty * cf.f;
  }

  // ---------- BOM 索引 ----------
  var bomByDessert = {}; // 甜點名 -> [{mat,qty,unit}]
  bom.forEach(function (r) {
    var dn = String(r['甜點名稱'] || '').trim(); if (!dn) return;
    (bomByDessert[dn] = bomByDessert[dn] || []).push({
      mat: String(r['食材/器具名稱'] || '').trim(),
      qty: Number(r['數量']) || 0,
      unit: r['單位']
    });
  });

  // ---------- ★ v1.7.1 專屬判定用：用料 → sku 編號；近 N 天 POS 有賣的甜點 ----------
  function matKey(m) { var u = name2sku[m]; return u ? ('id:' + u.id) : ('name:' + m); }
  var posToCanon = {};
  map.forEach(function (r) {
    var p = String(r['POS 資料產品名稱'] || '').trim(), mm = String(r['手動對應產品名稱'] || '').trim();
    if (p) posToCanon[p] = mm || p;
  });
  var exclFrom = evYmd_(evAddDays_(today, -EV_CFG.EXCL_RECENT_DAYS)), exclTo = evYmd_(today);
  var recentSold = {};   /* 甜點名 → { 次類別: 1 } */
  pos.forEach(function (r) {
    if (!(Number(r['數量']) > 0)) return;
    var d = evDate_(r['建立日期']); if (!d) return;
    var y = evYmd_(d); if (y < exclFrom || y >= exclTo) return;
    var raw = String(r['商品名稱'] || '').trim(); if (!raw) return;
    var c = posToCanon[raw] || raw;
    (recentSold[c] = recentSold[c] || {})[String(r['次類別'] || '').trim().replace(/^(\d{4})\s*/, '$1 ')] = 1;
  });

  // ---------- 今日有效的「非本檔期」甜點用料集合（供專屬判定） ----------
  function validNow(r) {
    var st = evDate_(r['起始有效日']), en = evDate_(r['結束有效日']);
    if (st && st > today) return false;
    if (en && en < today) return false;
    return true;
  }

  // ---------- 逐檔期處理 ----------
  var evRows = [], niRows = [], curveRows = [], needRows = [];
  var lookEnd = evAddDays_(today, EV_CFG.LOOKAHEAD_DAYS);
  EV_ALERTS = [];
  var posMin = evPosMinDate_(dayTotal);
  var nowTs = evYmd_(today);
  Object.keys(inst).forEach(function (k) {
    var C = inst[k];
    if (C.end < evAddDays_(today, -35) || C.start > lookEnd) return; // 進行中/即將到來；已結束檔期保留35天
    var L = Math.round((C.end - C.start) / 864e5) + 1;
    var P = evFindPrev_(inst, C);

    // ===== dim_event 自動列（v1.6 零改動） =====
    var warnMsg = '';
    if (P) {
      if (L <= EV_CFG.SHORT_MAX) {
        EV_LAST_DIAG = null;
        var coef = evUplift_(P, dayTotal, instDay);
        if (coef) evRows.push([C.evName + '(自動)', evYmd_(C.start), evYmd_(C.end), coef, '全部', '是', '自動',
          '樣本:' + P.year + ' 非限定品 檔期日均/前' + EV_CFG.BASE_DAYS + '天日均']);
        else warnMsg = evWarnMsg_('係數', C, P, EV_LAST_DIAG, posMin);
      } else {
        var pk = evPeakWeek_(P, dayTotal);
        if (pk && pk.coef >= EV_CFG.PEAK_MIN_COEF) {
          var mp = evMapPeak_(pk, P, C);
          if (mp.ok) {
            evRows.push([C.evName + '節日週(自動)', evYmd_(mp.start), evYmd_(mp.end), pk.coef, '全部', '是', '自動',
              '樣本:' + P.year + ' 尖峰週 ' + evYmd_(pk.start) + '~' + evYmd_(pk.end) + '（檔期結束前 ' + mp.offset + ' 天）/檔期均週' + (mp.clipped ? '｜已截在檔期內' : '')]);
          } else {
            var mm = '⚠️ 節日週無法放入檔期：去年尖峰 ' + evYmd_(pk.start) + '~' + evYmd_(pk.end) + ' ×' + pk.coef + '，對齊後 ' + evYmd_(mp.start) + '~' + evYmd_(mp.end) + ' 不足 ' + EV_CFG.PEAK_MIN_DAYS + ' 天｜請用「手動」自填';
            evRows.push([C.evName + '節日週(自動)', evYmd_(C.start), evYmd_(C.end), 1, '全部', '否', '自動', mm]);
            EV_ALERTS.push({ camp: k, evName: C.evName, year: C.year, kind: 'dim_event', msg: mm });
          }
        } else if (!pk) {
          warnMsg = evWarnMsg_('尖峰週', C, P, null, posMin);
        } else {
          evRows.push([C.evName + '(自動)', evYmd_(C.start), evYmd_(C.end), 1, '全部', '否', '自動',
            'ℹ️ 長檔期不設整段係數（近4週均自動跟隨）；去年 ' + P.year + ' 檔期內最高週僅 ×' + pk.coef + '（門檻 ' + EV_CFG.PEAK_MIN_COEF + '）無明顯節日週。備料走 agg_newitem。需要加成請改「手動」填係數並啟用']);
        }
      }
    } else {
      warnMsg = '⚠️ 無歷史樣本，需人工係數（改「手動」並填係數後啟用）';
    }
    if (warnMsg) {
      evRows.push([C.evName + '(自動)', evYmd_(C.start), evYmd_(C.end), 1, '全部', '否', '自動', warnMsg]);
      EV_ALERTS.push({ camp: k, evName: C.evName, year: C.year, kind: 'dim_event', msg: warnMsg });
    }

    // ===== agg_newitem：新品備料層＋標示 =====
    var leadStart = evAddDays_(C.start, -EV_CFG.LEADIN_DAYS);
    if (today < leadStart || today > C.end) return;
    var frac = L <= EV_CFG.SUPER_SHORT_MAX ? EV_CFG.FRAC_SUPER_SHORT : (L <= EV_CFG.SHORT_MAX ? EV_CFG.FRAC_SHORT : EV_CFG.FRAC_LONG);
    var names = Object.keys(C.items); if (!names.length) return;

    var pool = {};
    if (P && instItemStore[(P.year) + ' ' + P.evName]) {
      var pis = instItemStore[P.year + ' ' + P.evName];
      Object.keys(pis).forEach(function (nm) {
        Object.keys(pis[nm]).forEach(function (stn) { pool[stn] = (pool[stn] || 0) + pis[nm][stn]; });
      });
    }
    var poolAll = 0; Object.keys(pool).forEach(function (stn) { poolAll += pool[stn]; });

    var soldTot = {}, soldStore = {}, soldAllStore = {};
    var cis = instItemStore[k] || {};
    names.forEach(function (nm) {
      soldTot[nm] = 0; soldStore[nm] = cis[nm] || {};
      Object.keys(soldStore[nm]).forEach(function (stn) { soldTot[nm] += soldStore[nm][stn]; soldAllStore[stn] = (soldAllStore[stn] || 0) + soldStore[nm][stn]; });
    });
    var daysOn = Math.floor((today - C.start) / 864e5) + 1;
    var sumSold = names.reduce(function (a, nm) { return a + soldTot[nm]; }, 0);
    var share = {};
    names.forEach(function (nm) {
      share[nm] = (daysOn >= EV_CFG.RESPLIT_MIN_DAYS && sumSold > 0) ? soldTot[nm] / sumSold : 1 / names.length;
    });

    /* ★ v1.7.1：其他產品用料以 sku 編號記（matKey），不再用 BOM 原字串 */
    var otherMat = {};
    map.forEach(function (r) {
      var nm = String(r['手動對應產品名稱'] || '').trim() || String(r['POS 資料產品名稱'] || '').trim();
      if (!nm || C.items[nm] || !validNow(r)) return;
      (bomByDessert[nm] || []).forEach(function (b) { otherMat[matKey(b.mat)] = 1; });
    });
    Object.keys(recentSold).forEach(function (nm) {
      if (C.items[nm] || recentSold[nm][k]) return;   /* 本檔期自己的品項不算 */
      (bomByDessert[nm] || []).forEach(function (b) { otherMat[matKey(b.mat)] = 1; });
    });

    // ===== ★ v1.6：去年曲線 → 今年推估（全公司口徑） =====
    var cv = null;
    if (P && poolAll > 0) {
      cv = evCurve_(P, C, instDay, today);
      var cvWhy = !cv.ok ? cv.why
        : (daysOn < EV_CFG.RESPLIT_MIN_DAYS ? '開賣未滿 ' + EV_CFG.RESPLIT_MIN_DAYS + ' 天'
        : (cv.cumToDate < EV_CFG.CURVE_MIN_CUM ? '累積佔比 ' + Math.round(cv.cumToDate * 100) + '% 未達 ' + Math.round(EV_CFG.CURVE_MIN_CUM * 100) + '%'
        : (sumSold <= 0 ? '今年尚無銷售' : '')));
      if (cvWhy) { log.push('ℹ️ ' + k + ' 曲線未啟用：' + cvWhy + '（走首批級距／7 天交棒原路徑）'); cv = null; }
    }
    var curveMode = !!cv;
    var estAll = 0, estRatio = 0;
    var storeShare = {};
    if (curveMode) {
      estAll = cv.est;                     /* ＝窗內已售÷累積佔比 ＋ 窗外（提早開賣）已售 */
      estRatio = estAll / cv.total;
      if (estRatio > EV_CFG.CURVE_MAX_RATIO) { estAll = cv.total * EV_CFG.CURVE_MAX_RATIO; log.push('⚠️ ' + k + ' 推估總量達去年 ' + estRatio.toFixed(2) + ' 倍，以上限 ' + EV_CFG.CURVE_MAX_RATIO + ' 倍計'); estRatio = EV_CFG.CURVE_MAX_RATIO; }
      var useThisYear = (daysOn >= EV_CFG.RESPLIT_MIN_DAYS && sumSold > 0);
      Object.keys(pool).forEach(function (stn) { storeShare[stn] = useThisYear ? ((soldAllStore[stn] || 0) / sumSold) : (pool[stn] / poolAll); });
      Object.keys(soldAllStore).forEach(function (stn) { if (storeShare[stn] == null) storeShare[stn] = useThisYear ? soldAllStore[stn] / sumSold : 0; });
      log.push('📈 ' + k + ' 曲線模式：去年主段 ' + evYmd_(P.start) + '~' + evYmd_(cv.mainEnd) + '（' + Math.round(cv.total) + ' 份）錨點=' + cv.anchorTxt
        + '｜今年已售 ' + Math.round(sumSold) + ' 份（窗內 ' + Math.round(cv.soldIn) + '＝去年同位置累積 ' + Math.round(cv.cumToDate * 100) + '%，提早開賣 ' + Math.round(cv.soldPre) + '）'
        + ' ⇒ 推估總量 ' + Math.round(estAll) + ' 份（去年 ' + estRatio.toFixed(2) + ' 倍）｜下週佔比 ' + Math.round(cv.nextWeekShare * 100) + '%');
      /* ★ v1.7：POS 檔期銷量最新日期（每天應是昨天；不是 ⇒ 今天推估用的是舊資料） */
      var lastSaleY = Object.keys(instDay[k] || {}).sort().pop() || '';
      var yestY = evYmd_(evAddDays_(today, -1));
      log.push((lastSaleY >= yestY ? '✅ ' : '⚠️ ') + k + ' POS 檔期銷量最新到 ' + (lastSaleY || '無') + (lastSaleY >= yestY ? '（＝昨天，推估已含最新銷售）' : '（不是昨天 ' + yestY + '：POS 同步沒跟上，今天各店推估用的是舊資料，請先查 POS 同步）'));
      // agg_evcurve：逐週列（★ v1.7：今年推估只算到檔期結束日）
      /* ★ v1.7.1：今年推估＝已過的日子用實際已售＋今天起到結束日用推估（evWeekEst_） */
      var rBase = evCurveBase_(cv, estAll);
      cv.weeks.forEach(function (w) {
        curveRows.push([C.evName, w.idx, w.pFrom, w.pTo, w.pQty, Math.round(w.share * 1000) / 10, w.cFrom, w.cTo, w.cSold, Math.round(evWeekEst_(cv, w, C.end, rBase)), w.rel, nowTs]);
      });
    }

    /* ★ v1.9：新品申請表預估銷售數（整檔 12 店合計）→ 本店整檔推估 T（只用在非曲線模式） */
    var planT = {};
    if (!curveMode) names.forEach(function (nm) {
      var pl = plan[nm]; if (!pl) return;
      var sp = evPlanSplit_(pl, pool, recentQty); planT[nm] = sp.t;
      log.push('📋 ' + k + '「' + nm + '」用新品申請表預估 ' + pl.qty + ' 份（' + pl.req + '，' + (pl.brand || '全店') + '）拆店依' + sp.src);
    });
    var planned = names.filter(function (nm) { return planT[nm]; });

    var acc = {}, need = {};
    if (!Object.keys(pool).length && planned.length < names.length) {
      var poolMsg = '⚠️ 無去年總池：agg_newitem 該檔期全 12 店 qty=0，備料建議為 0（非真的不用備料）' +
        '｜去年檔期=' + (P ? (P.year + ' ' + P.evName + ' ' + evYmd_(P.start) + '~' + evYmd_(P.end)) : '找不到') +
        '｜POS資料 最舊日期=' + (posMin || '未知');
      log.push('⚠️ ' + k + ' 無去年總池，備料層量=0 僅輸出標示列');
      evRows.push([C.evName + '(備料警示)', evYmd_(C.start), evYmd_(C.end), 1, '全部', '否', '自動', poolMsg]);
      EV_ALERTS.push({ camp: k, evName: C.evName, year: C.year, kind: 'agg_newitem', msg: poolMsg });
    }
    var recentFrom = evYmd_(evAddDays_(today, -EV_CFG.RECENT_DAYS));
    var paceStat = { AB: 0, A: 0, B: 0, zero: 0, newFloor: 0, capped: 0, remSum: 0, list: [] };   /* ★ v1.7 log 用 */
    names.forEach(function (nm) {
      var rows = bomByDessert[nm];
      if (!rows) { log.push('⚠️ 新品 BOM 對不上: ' + nm); return; }
      var first = (instItemFirst[k] || {})[nm];
      var taken = first ? Math.floor((today - first) / 864e5) + 1 : 0;
      var stores = curveMode ? Object.keys(storeShare) : (planT[nm] ? Object.keys(planT[nm]) : Object.keys(pool));   /* ★ v1.9 */
      stores.forEach(function (stn) {
        var sold = (soldStore[nm] || {})[stn] || 0;
        var remain, estTot = 0, nextWk = 0, recentWk = 0, remainServ = 0, planEst = 0;
        if (curveMode) {
          /* ★ v1.7 曲線路徑：各店各品項用「自己的銷售速度」× 去年曲線（取代 v1.6 的 全公司×店份額×品項份額）
             需求量仍是增量（下週 − 近4週均週份數）；剩餘份數給 agg_evneed。 */
          var dd = ((instItemStoreDay[k] || {})[nm] || {})[stn] || {};
          var pc = evStorePace_(cv, dd, first, (C.items[nm] && C.items[nm].end) || C.end);
          if (pc.soldAll <= 0 && first && taken <= EV_CFG.PACE_NEW_DAYS) {
            pc.R = estAll * (storeShare[stn] || 0) * share[nm]; pc.how = 'newFloor';   /* 新上架保底 */
          }
          remainServ = pc.R * pc.remShare;
          estTot = sold + remainServ;
          nextWk = pc.R * pc.nextShare;
          var rq = 0; Object.keys(dd).forEach(function (y) { if (y >= recentFrom && y < nowTs) rq += dd[y]; });
          recentWk = rq / (EV_CFG.RECENT_DAYS / 7);
          remain = Math.max(0, nextWk - recentWk);
          paceStat[pc.how === 'newFloor' ? 'newFloor' : (pc.how === '0' ? 'zero' : pc.how)]++;
          if (pc.capped) paceStat.capped++;
          paceStat.remSum += remainServ;
          if (sold > 0) paceStat.list.push({ st: stn, nm: nm, sold: sold, sold7: pc.sold7, rem: remainServ, nx: nextWk, how: pc.how });
        } else if (planT[nm]) {
          /* ★ v1.9：有申請表預估 → 首批目標＝本店整檔推估 T × 首批比例 − 已售；開賣滿 7 天 POS 後照原規則歸零，交給實際用量 */
          planEst = planT[nm][stn] || 0;
          remain = taken >= EV_CFG.TAKEOVER_DAYS ? 0 : Math.max(0, planEst * frac - sold);
        } else {
          /* v1.5 原路徑（零改動）：首批目標 − 已售，7 天 POS 後歸零 */
          var target = (pool[stn] || 0) * share[nm] * frac;
          remain = taken >= EV_CFG.TAKEOVER_DAYS ? 0 : Math.max(0, target - sold);
        }
        var weeklyServ = (planT[nm] ? planEst : (pool[stn] || 0) * share[nm]) / Math.max(1, L / 7);   /* 器具/模具維持 v1.5 輸入；v1.9 有預估改用 T */
        var durServ = Math.ceil(weeklyServ * 2 / 7);
        rows.forEach(function (b) {
          var u = name2sku[b.mat]; if (!u) { return; }
          var key = stn + '|' + u.id;
          var o = acc[key] = acc[key] || { st: stn, id: u.id, qty: 0, excl: !otherMat[matKey(b.mat)], who: {} };
          var isDur = (u.cat === '器具' || u.cat === '模具');
          if (isDur) o.qty += durServ * (Number(b.qty) || 0);
          else o.qty += remain * toUse(b.mat, b.qty, b.unit);
          o.who[nm] = 1;
          if (curveMode && !isDur) {
            /* agg_evneed：出貨中心用——檔期剩餘推估用量／已用量（皆使用單位） */
            var per = toUse(b.mat, b.qty, b.unit);
            var n = need[key] = need[key] || { st: stn, id: u.id, remainUse: 0, usedUse: 0, sold: 0, est: 0, who: {} };
            n.remainUse += remainServ * per;   /* ★ v1.7：剩餘份數直接來自各店速度，不再是 max(0, 推估−已售) */
            n.usedUse += sold * per;
            n.sold += sold; n.est += estTot; n.who[nm] = 1;
          } else if (!curveMode && planT[nm] && !isDur) {
            /* ★ v1.9：開賣前（非曲線）也寫 agg_evneed，出貨中心看得到 12 店總需求：推估總份數＝max(T, 已售) */
            var per2 = toUse(b.mat, b.qty, b.unit), est2 = Math.max(planEst, sold);
            var n2 = need[key] = need[key] || { st: stn, id: u.id, remainUse: 0, usedUse: 0, sold: 0, est: 0, who: {} };
            n2.remainUse += Math.max(0, est2 - sold) * per2;
            n2.usedUse += sold * per2;
            n2.sold += sold; n2.est += est2; n2.who[nm] = 1;
          }
        });
      });
      if (!Object.keys(pool).length) {
        for (var s = 1; s <= 12; s++) {
          rows.forEach(function (b) {
            var u = name2sku[b.mat]; if (!u) return;
            var key = s + '|' + u.id;
            var o = acc[key] = acc[key] || { st: String(s), id: u.id, qty: 0, excl: !otherMat[matKey(b.mat)], who: {} };
            o.who[nm] = 1;
          });
        }
      }
    });
    if (curveMode) {
      var compRem = (cv.cumToDate > 0 ? cv.soldIn / cv.cumToDate : 0) * evShareSum_(cv, cv.offToday, evOff_(cv, C.end));   /* 同口徑：不含提早開賣那段 */
      log.push('🏪 ' + k + ' 各店推估（v1.7 各店各品項自己的速度）：各店剩餘合計 ' + Math.round(paceStat.remSum) + ' 份 vs 全公司曲線剩餘 ' + Math.round(compRem) + ' 份'
        + '｜累積＋近7天 ' + paceStat.AB + '、只用累積 ' + paceStat.A + '、只用近7天 ' + paceStat.B + '、無資料 ' + paceStat.zero + '、新品保底 ' + paceStat.newFloor + '、近7天封頂 ' + paceStat.capped);
      paceStat.list.sort(function (a, b) { return b.rem - a.rem; }).slice(0, 8).forEach(function (x) {
        log.push('   ' + x.st + '店 ' + x.nm + '：已售 ' + Math.round(x.sold) + '（近7天 ' + Math.round(x.sold7) + '）→ 還會賣 ' + Math.round(x.rem) + ' 份、下週 ' + (Math.round(x.nx * 10) / 10) + ' 份');
      });
    }
    var unmatched = {};
    names.forEach(function (nm) { (bomByDessert[nm] || []).forEach(function (b) { if (!name2sku[b.mat]) unmatched[b.mat] = 1; }); });
    if (Object.keys(unmatched).length) log.push('⚠️ ' + k + ' 用料對不上 dim_sku: ' + Object.keys(unmatched).join('、'));
    Object.keys(acc).forEach(function (key) {
      var o = acc[key];
      niRows.push([o.st, o.id, Math.round(o.qty * 10) / 10, C.evName, o.excl ? '是' : '否', Object.keys(o.who).join('、')]);
    });
    Object.keys(need).forEach(function (key) {
      var n = need[key];
      needRows.push([C.evName, n.st, n.id, Math.round(n.remainUse * 10) / 10, Math.round(n.usedUse * 10) / 10, Math.round(n.sold), Math.round(n.est), Object.keys(n.who).join('、'), nowTs]);
    });
  });

  // ---------- ★ v1.8 收尾扣減 agg_endcut ----------
  /* v1.8.3（10 輪檢查第 8 輪）：收尾計算放在 dim_event／agg_newitem 寫入之前，出例外會讓整個重算中斷、當天檔期備料全不更新。
     包 try：失敗時 agg_endcut 寫空（前端＝不做收尾，安全方向），其他表照常 */
  var endCut;
  try { endCut = evEndCut_(map, pos, bomByDessert, name2sku, conv, today, log); }
  catch (eEc) { log.push('⚠️ 收尾扣減計算失敗：' + eEc + '｜agg_endcut 本次寫空（前端不做收尾扣減），其他表照常更新'); endCut = { rows: [] }; }

  // ---------- 寫回 dim_event（保留手動列） ----------
  var evSheet = evEnsure_(ss, 'dim_event', ['事件名', '開始日', '結束日', '加成係數', '適用品類', '啟用', '來源', '說明']);
  var keep = [];
  var old = evSheet.getDataRange().getValues();
  for (var i = 1; i < old.length; i++) {
    if (old[i][0] && String(old[i][6] || '') !== '自動') keep.push(old[i].slice(0, 8));
  }
  evSheet.getRange(2, 1, Math.max(1, evSheet.getLastRow()), 8).clearContent();
  var out = keep.concat(evRows);
  if (out.length) evSheet.getRange(2, 1, out.length, 8).setValues(out.map(function (r) { r = r.slice(0, 8); while (r.length < 8) r.push(''); return r; }));

  // ---------- 寫回 agg_newitem（全量重建） ----------
  var niSheet = evEnsure_(ss, 'agg_newitem', ['店號', 'sku_id', '需求量', '事件名', '專屬', '新品清單']);
  niSheet.getRange(2, 1, Math.max(1, niSheet.getLastRow()), 6).clearContent();
  if (niRows.length) niSheet.getRange(2, 1, niRows.length, 6).setValues(niRows);

  // ---------- ★ v1.6 寫回 agg_evcurve / agg_evneed（全量重建，小表） ----------
  var CV_H = ['事件名', '週序', '去年週起', '去年週訖', '去年份數', '佔比%', '今年週起', '今年週訖', '今年已售', '今年推估', '相對開賣週', '更新日'];
  var cvSheet = evEnsureSmall_(ss, 'agg_evcurve', CV_H);
  cvSheet.getRange(2, 1, Math.max(1, cvSheet.getLastRow()), CV_H.length).clearContent();
  if (curveRows.length) cvSheet.getRange(2, 1, curveRows.length, CV_H.length).setValues(curveRows);
  var ND_H = ['事件名', '店號', 'sku_id', '剩餘推估用量', '已用量', '已售份數', '推估總份數', '新品清單', '更新日'];
  var ndSheet = evEnsureSmall_(ss, 'agg_evneed', ND_H);
  ndSheet.getRange(2, 1, Math.max(1, ndSheet.getLastRow()), ND_H.length).clearContent();
  if (needRows.length) ndSheet.getRange(2, 1, needRows.length, ND_H.length).setValues(needRows);

  var EC_H = ['店號', 'sku_id', '結束日', '佔比', '近4週週用量', '甜點', '更新日', '份數佔比', '在售甜點', '近7天比'];   /* v1.8.1 ＋份數佔比、在售甜點；v1.8.4 ＋近7天比 */
  var ecSheet = evEnsureSmall_(ss, 'agg_endcut', EC_H);
  try { if (ecSheet.getMaxColumns() < EC_H.length) ecSheet.insertColumnsAfter(ecSheet.getMaxColumns(), EC_H.length - ecSheet.getMaxColumns()); } catch (e0) {}
  ecSheet.getRange(1, 1, 1, EC_H.length).setValues([EC_H]);   /* 舊表只有 7 欄表頭，每次重寫 */
  evEnsureRows_(ecSheet, endCut.rows.length + 1);
  ecSheet.getRange(2, 1, Math.max(1, ecSheet.getLastRow()), EC_H.length).clearContent();
  if (endCut.rows.length) ecSheet.getRange(2, 1, endCut.rows.length, EC_H.length).setValues(endCut.rows);
  evFormatEndcut_(ecSheet);   /* v1.8.2：鏡像前先設好格式，getValues 才會拿到數字而不是日期 */
  log.push('✂️ agg_endcut ' + endCut.rows.length + ' 列（收尾扣減：即將／剛下架甜點的用量只算到結束日；v1.8.1 附份數佔比，客人會改做其他甜點的部分前端不扣）');

  /* ★ v1.6.1：四張表鏡像到專用檔（v1.8 起五張） */
  log.push('專用檔鏡像：' + evMirror_(ss, EV_MIRROR_TABS).join('；'));
  try { var dstEc = SpreadsheetApp.openById(EV_DASH_ID).getSheetByName('agg_endcut'); if (dstEc) evFormatEndcut_(dstEc); } catch (eF) { log.push('⚠️ 專用檔 agg_endcut 設定格式失敗：' + eF); }   /* v1.8.2 */
  log.push('dim_event 自動列 ' + evRows.length + '、手動列保留 ' + keep.length + '；agg_newitem ' + niRows.length + ' 列；agg_evcurve ' + curveRows.length + ' 列；agg_evneed ' + needRows.length + ' 列');
  evRows.forEach(function (r) { log.push('   [dim_event] ' + r[0] + ' ' + r[1] + '~' + r[2] + ' ×' + r[3] + ' 啟用=' + r[5]); });

  if (EV_ALERTS.length) {
    log.push('');
    log.push('🔴 本次有 ' + EV_ALERTS.length + ' 筆警告（已寫入 dim_event，啟用=否）：');
    EV_ALERTS.forEach(function (a) { log.push('   [' + a.kind + '] ' + a.camp + '｜' + a.msg); });
    log.push('   POS資料 最舊日期 = ' + (posMin || '未知') +
             '｜若這些檔期的去年窗口已被滾動裁切，屬預期；否則要查對照表次類別是否漏建。');
  } else {
    log.push('✅ 無歷史樣本警告：0 筆');
  }

  Logger.log(log.join('\n'));
  return log.join('\n');
}

/* ═══════════════════════════════════════════════════════════════
 * ★ v1.6 曲線核心（純函式，不讀不寫試算表）
 * ═══════════════════════════════════════════════════════════════ */

/** 去年檔期主段結束日：從主段開始起，遇到連續 GAP_DAYS 天零銷售即結束。 */
function evMainEnd_(P, dayMap) {
  var last = null, gap = 0;
  for (var d = new Date(P.start); d <= P.end; d = evAddDays_(d, 1)) {
    var q = dayMap[evYmd_(d)] || 0;
    if (q > 0) { last = new Date(d); gap = 0; }
    else if (last) { gap++; if (gap >= EV_CFG.GAP_DAYS) break; }
  }
  return last || P.end;
}

/** 節日錨點：EV_FEST 有登錄→節日當天；否則 null。 */
function evFest_(evName, year) {
  var cand = [evName].concat(EV_ALIAS[evName] || []);
  for (var i = 0; i < cand.length; i++) {
    var t = EV_FEST[cand[i]]; if (t && t[year]) return evDate_(t[year]);
  }
  return null;
}

/**
 * 去年曲線 → 今年推估的所有中間量。
 * 回傳 {ok, why, total, mainEnd, anchorTxt, cumToDate, nextWeekShare, weeks:[{idx,pFrom,pTo,pQty,share,cFrom,cTo,cSold,rel}]}
 *  · 逐日佔比 dayShare[off]：off＝距錨點天數（負＝節前）
 *  · cumToDate＝今年「今天之前」對應位置的累積佔比
 *  · nextWeekShare＝今年今天起 7 天對應位置的佔比合計
 *  · weeks：以錨點所在週為第 0 週（週序＝負為節前），供 agg_evcurve 與畫面用
 */
function evCurve_(P, C, instDay, today) {
  var pk = P.year + ' ' + P.evName, ck = C.year + ' ' + C.evName;
  var dm = instDay[pk] || {}, cm = instDay[ck] || {};
  if (!Object.keys(dm).length) return { ok: false, why: '去年檔期 POS 無日銷資料' };
  var mainEnd = evMainEnd_(P, dm);
  var mainDays = Math.round((mainEnd - P.start) / 864e5) + 1;
  if (mainDays < EV_CFG.CURVE_MIN_DAYS) return { ok: false, why: '去年主段僅 ' + mainDays + ' 天（需≥' + EV_CFG.CURVE_MIN_DAYS + '）' };
  var fP = evFest_(P.evName, P.year), fC = evFest_(C.evName, C.year);
  var anchorP, anchorC, anchorTxt;
  if (fP && fC) { anchorP = fP; anchorC = fC; anchorTxt = '節日 ' + evYmd_(fP) + '→' + evYmd_(fC); }
  else { anchorP = mainEnd; anchorC = C.end; anchorTxt = '主段結束 ' + evYmd_(mainEnd) + '→檔期結束 ' + evYmd_(C.end); }
  var dayShare = {}, total = 0, offMin = 0, offMax = 0;
  for (var d = new Date(P.start); d <= mainEnd; d = evAddDays_(d, 1)) {
    var q = dm[evYmd_(d)] || 0; total += q;
    var off = Math.round((d - anchorP) / 864e5);
    dayShare[off] = (dayShare[off] || 0) + q;
    if (off < offMin) offMin = off; if (off > offMax) offMax = off;
  }
  if (!total) return { ok: false, why: '去年主段銷量為 0' };
  Object.keys(dayShare).forEach(function (o) { dayShare[o] = dayShare[o] / total; });
  var offToday = Math.round((today - anchorC) / 864e5);
  var cum = 0, next = 0;
  for (var o = offMin; o <= offMax; o++) {
    var s = dayShare[o] || 0;
    if (o < offToday) cum += s;
    else if (o < offToday + 7) next += s;
  }
  /* 今年已售拆兩段：去年同位置有資料的「窗內」（÷累積佔比放大）與今年提早開賣的「窗外」（原數加回，不放大）。
     2026 中秋 8/15 開賣、2025 對應位置 8/19 才開賣 ⇒ 前 4 天 201 份若一起放大會把總量高估 10%。 */
  var soldIn = 0, soldPre = 0;
  Object.keys(cm).forEach(function (y) {
    var off = Math.round((evDate_(y) - anchorC) / 864e5);
    if (off >= offToday) return;
    if (off < offMin) soldPre += cm[y]; else soldIn += cm[y];
  });
  // 逐週：錨點所在週＝第 0 週，週界＝錨點日往前後每 7 天
  var weeks = [], wMin = Math.floor(offMin / 7), wMax = Math.floor(offMax / 7), openQty = 0;
  for (var w = wMin; w <= wMax; w++) {
    var pq = 0, cs = 0;
    for (var o2 = w * 7; o2 < w * 7 + 7; o2++) {
      pq += (dayShare[o2] || 0) * total;
      var cd = evAddDays_(anchorC, o2); if (cd < today) cs += (cm[evYmd_(cd)] || 0);
    }
    pq = Math.round(pq);
    /* 「相對開賣週」的分母＝第一個「整週都落在主段內」的週。第一桶常只有 2 天有賣（2025 中秋 8/30~8/31＝181 份），
       拿它當分母會算出 ×4.96 這種假倍率；用第一個完整週（9/1~9/7＝737）尖峰週才是合理的 ×1.22。 */
    var full = (evAddDays_(anchorP, w * 7) >= P.start) && (evAddDays_(anchorP, w * 7 + 6) <= mainEnd);
    if (!openQty && full && pq > 0) openQty = pq;
    weeks.push({ idx: w, pFrom: evYmd_(evAddDays_(anchorP, w * 7)), pTo: evYmd_(evAddDays_(anchorP, w * 7 + 6)), pQty: pq, share: pq / total,
                 cFrom: evYmd_(evAddDays_(anchorC, w * 7)), cTo: evYmd_(evAddDays_(anchorC, w * 7 + 6)), cSold: cs, full: full, rel: 0 });
  }
  weeks.forEach(function (wk) { wk.rel = openQty ? Math.round(wk.pQty / openQty * 100) / 100 : 0; });
  var est = (cum > 0 ? soldIn / cum : 0) + soldPre;
  return { ok: true, total: total, mainEnd: mainEnd, anchorTxt: anchorTxt, cumToDate: cum, nextWeekShare: next, weeks: weeks,
           soldIn: soldIn, soldPre: soldPre, est: est,
           dayShare: dayShare, offMin: offMin, offMax: offMax, offToday: offToday, anchorC: anchorC };   /* ★ v1.7 供各店速度用 */
}

/** ★ v1.7.1：全公司推估基準（窗內已售 ÷ 累積佔比；推估總量觸上限時跟著封頂） */
function evCurveBase_(cv, estAll) {
  var r = cv.cumToDate > 0 ? cv.soldIn / cv.cumToDate : 0;
  return (estAll > 0 && r > estAll) ? estAll : r;
}

/** ★ v1.7.1：某週「實際＋推估」＝該週已過日子的實際已售（w.cSold）＋今天起到檔期結束日的推估 */
function evWeekEst_(cv, w, endDate, rBase) {
  var endOff = endDate ? evOff_(cv, endDate) : cv.offMax;
  var from = Math.max(w.idx * 7, cv.offToday), to = Math.min(w.idx * 7 + 6, endOff);
  return w.cSold + (to >= from ? rBase * evShareSum_(cv, from, to) : 0);
}

/** ★ v1.7：日期 → 距今年錨點天數 */
function evOff_(cv, d) { return Math.round((evDateOnly_(d) - cv.anchorC) / 864e5); }

/** ★ v1.7：去年曲線在 [from, to]（距錨點天數，含頭含尾）的佔比合計；超出主段範圍的天數算 0 */
function evShareSum_(cv, from, to) {
  var s = 0, a = Math.max(from, cv.offMin), b = Math.min(to, cv.offMax);
  for (var o = a; o <= b; o++) s += cv.dayShare[o] || 0;
  return s;
}

/**
 * ★ v1.7 核心（純函式）：某店某品項的「等效整檔份數」R 與剩餘／下週佔比。
 *  dd：{ 'YYYY-MM-DD': 份數 }（該店該品項逐日）；firstDate：該品項全公司首賣日；endDate：該品項檔期結束日
 *  A＝窗內已售 ÷ 去年同位置累積佔比（從 max(品項首賣日, 去年主段起點) 算到昨天）
 *  B＝近 PACE_DAYS 天已售 ÷ 去年同位置那幾天的佔比，最多 A × PACE_B_CAP
 *  R＝(A+B)/2；近 7 天 0 份但之前有賣 → 只用 A；只有一邊能算用那一邊；都不能算 → 0
 */
function evStorePace_(cv, dd, firstDate, endDate) {
  var offT = cv.offToday;
  var fromOff = Math.max(firstDate ? evOff_(cv, firstDate) : cv.offMin, cv.offMin);
  var endOff = endDate ? evOff_(cv, endDate) : cv.offMax;
  var w7 = Math.max(offT - EV_CFG.PACE_DAYS, fromOff);
  var soldAll = 0, soldA = 0, sold7 = 0;
  Object.keys(dd || {}).forEach(function (y) {
    var d = evDate_(y); if (!d) return;
    var o = evOff_(cv, d); if (o >= offT) return;
    var q = Number(dd[y]) || 0;
    soldAll += q;
    if (o >= fromOff) soldA += q;
    if (o >= w7) sold7 += q;
  });
  var cumA = evShareSum_(cv, fromOff, offT - 1), share7 = evShareSum_(cv, w7, offT - 1);
  var A = cumA >= EV_CFG.PACE_MIN_CUM ? soldA / cumA : null;
  var B = (share7 >= EV_CFG.PACE_MIN_SHARE7 && (offT - w7) >= 3) ? sold7 / share7 : null;
  var capped = false;
  if (A != null && B != null && A > 0 && B > A * EV_CFG.PACE_B_CAP) { B = A * EV_CFG.PACE_B_CAP; capped = true; }
  var R, how;
  if (A != null && B != null) {
    if (sold7 <= 0 && soldA > 0) { R = A; how = 'A'; } else { R = (A + B) / 2; how = 'AB'; }
  } else if (A != null) { R = A; how = 'A'; }
  else if (B != null) { R = B; how = 'B'; }
  else { R = 0; how = '0'; }
  var remShare = endOff < offT ? 0 : evShareSum_(cv, offT, endOff);
  var nextShare = endOff < offT ? 0 : evShareSum_(cv, offT, Math.min(offT + 6, endOff));
  return { R: Math.max(0, R), A: A, B: B, soldAll: soldAll, soldA: soldA, sold7: sold7, cumA: cumA, share7: share7,
           remShare: remShare, nextShare: nextShare, how: how, capped: capped };
}

/** ★ v1.6 唯讀預覽：每個進行中檔期的曲線、今年已售、推估總量、各店拆分。不寫任何表。 */
function evCurvePreview_() {
  var ss = SpreadsheetApp.openById(EV_BOM_ID);
  var today = evDateOnly_(new Date());
  var map = evReadObjs_(ss, '產品名稱對照表'), pos = evReadObjs_(ss, 'POS資料');
  var ctx = evBuildContext_(map, pos, today, []);
  var out = ['evCurvePreview ' + evYmd_(today)];
  Object.keys(ctx.inst).forEach(function (k) {
    var C = ctx.inst[k];
    if (today < evAddDays_(C.start, -EV_CFG.LEADIN_DAYS) || today > C.end) return;
    var P = evFindPrev_(ctx.inst, C);
    var line = k + '（' + evYmd_(C.start) + '~' + evYmd_(C.end) + (C.startCorrected ? '，起始日已依 POS 校正' : '') + '）';
    if (!P) { out.push(line + '：無去年樣本'); return; }
    var cv = evCurve_(P, C, ctx.instDay, today);
    if (!cv.ok) { out.push(line + '：曲線不可用—' + cv.why); return; }
    var cis = ctx.instItemStore[k] || {}, sumSold = 0, byStore = {};
    Object.keys(cis).forEach(function (nm) { Object.keys(cis[nm]).forEach(function (stn) { sumSold += cis[nm][stn]; byStore[stn] = (byStore[stn] || 0) + cis[nm][stn]; }); });
    var est = cv.est;
    out.push(line + '：去年主段 ' + evYmd_(P.start) + '~' + evYmd_(cv.mainEnd) + ' 共 ' + Math.round(cv.total) + ' 份｜錨點 ' + cv.anchorTxt);
    out.push('   今年已售 ' + Math.round(sumSold) + ' 份（窗內 ' + Math.round(cv.soldIn) + '＝去年同位置累積 ' + Math.round(cv.cumToDate * 100) + '%，提早開賣 ' + Math.round(cv.soldPre) + '）⇒ 推估總量 ' + Math.round(est) + ' 份（去年 ' + (cv.total ? (est / cv.total).toFixed(2) : '-') + ' 倍）；下週佔比 ' + Math.round(cv.nextWeekShare * 100) + '%'
      + (cv.cumToDate < EV_CFG.CURVE_MIN_CUM ? '　⚠️ 累積未達 ' + Math.round(EV_CFG.CURVE_MIN_CUM * 100) + '%，rebuild 會走原路徑' : '　✅ 會用曲線'));
    var rBaseP = evCurveBase_(cv, est);   /* ★ v1.7.1：與 agg_evcurve 同口徑（實際＋推估，算到對照表結束日） */
    cv.weeks.forEach(function (w) {
      out.push('   週' + (w.idx >= 0 ? '+' : '') + w.idx + '  去年 ' + w.pFrom + '~' + w.pTo + ' ' + w.pQty + ' 份（' + Math.round(w.share * 100) + '%，×' + w.rel + '）→ 今年 ' + w.cFrom + '~' + w.cTo + ' 已售 ' + w.cSold + ' 實際＋推估 ' + Math.round(evWeekEst_(cv, w, C.end, rBaseP)));
    });
    var ks = Object.keys(byStore).sort(function (a, b) { return (+a) - (+b); });
    out.push('   各店今年份額：' + ks.map(function (s) { return s + '店 ' + byStore[s] + '(' + Math.round(byStore[s] / Math.max(1, sumSold) * 100) + '%)'; }).join('、'));
  });
  if (out.length === 1) out.push('（目前沒有進行中或即將開賣的檔期）');
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ═══════════════════════════════════════════════════════════════
 * ★ v1.5 新增／抽出的共用區塊
 * ═══════════════════════════════════════════════════════════════ */

/** 從對照表＋POS 建立檔期 instance、日總量、檔期日量等（rebuild 與 preview 共用；純讀）。 */
function evBuildContext_(map, pos, today, log) {
  var inst = {};
  map.forEach(function (r) {
    var sub = String(r['次類別'] || '').trim();
    var m = sub.match(/^(\d{4})\s*(.+)$/);
    if (!m) return;
    var st = evDate_(r['起始有效日']), en = evDate_(r['結束有效日']);
    if (!st || !en) return;
    var k = m[1] + ' ' + m[2];
    var o = inst[k] = inst[k] || { year: +m[1], evName: m[2], start: st, end: en, items: {}, posNames: {} };
    if (st < o.start) o.start = st;
    if (en > o.end) o.end = en;
    var nm = String(r['手動對應產品名稱'] || r['POS 資料產品名稱'] || '').trim();
    if (nm) o.items[nm] = { start: st, end: en };
    var pn = String(r['POS 資料產品名稱'] || '').trim();
    if (pn) o.posNames[pn] = nm || pn;
  });
  var dayTotal = {}, instDay = {}, instItemStore = {}, instItemFirst = {}, instItemStoreDay = {};
  pos.forEach(function (r) {
    var d = evDate_(r['建立日期']); if (!d) return;
    var y = evYmd_(d), q = Number(r['數量']) || 0; if (!q) return;
    dayTotal[y] = (dayTotal[y] || 0) + q;
    var sub = String(r['次類別'] || '').trim();
    var m = sub.match(/^(\d{4})\s*(.+)$/); if (!m) return;
    var k = m[1] + ' ' + m[2], o = inst[k]; if (!o) return;
    (instDay[k] = instDay[k] || {})[y] = (instDay[k][y] || 0) + q;
    var raw = String(r['商品名稱'] || '').trim();
    var nm = o.posNames[raw] || raw;
    var stn = String(r['分店代碼'] || '');
    var ii = (instItemStore[k] = instItemStore[k] || {});
    (ii[nm] = ii[nm] || {})[stn] = (ii[nm][stn] || 0) + q;
    var ff = (instItemFirst[k] = instItemFirst[k] || {});
    if (!ff[nm] || d < ff[nm]) ff[nm] = d;
    /* ★ v1.6：限定品 逐店逐日（只有檔期品，體積很小），供曲線增量計算的近 N 天均 */
    var sd = (instItemStoreDay[k] = instItemStoreDay[k] || {});
    var sn = (sd[nm] = sd[nm] || {});
    (sn[stn] = sn[stn] || {})[y] = (sn[stn][y] || 0) + q;
  });
  // 歷史檔期窗口以 POS 實際售出日校正；★ v1.6：進行中檔期的「起始日」也校正（結束日仍以對照表為準）
  Object.keys(inst).forEach(function (k) {
    var o = inst[k], dd = instDay[k];
    if (!dd) return;
    var ys = Object.keys(dd).sort();
    if (o.end >= today) {
      if (o.start <= today && ys.length >= EV_CFG.CURVE_MIN_CORRECT) {
        var s0 = evDate_(ys[0]);
        if (s0 && evYmd_(s0) !== evYmd_(o.start)) { o.startRaw = o.start; o.start = s0; o.startCorrected = true; if (log) log.push('ℹ️ ' + k + ' 起始日依 POS 首賣日校正：' + evYmd_(o.startRaw) + ' → ' + evYmd_(s0)); }
      }
      return;
    }
    if (ys.length >= 5) {
      var s = evDate_(ys[0]), e = evDate_(ys[ys.length - 1]);
      if (s && e && e > s) { o.start = s; o.end = e; }
    }
  });
  return { inst: inst, dayTotal: dayTotal, instDay: instDay, instItemStore: instItemStore, instItemFirst: instItemFirst, instItemStoreDay: instItemStoreDay };
}

/** 找去年／前年同檔期（含別名） */
function evFindPrev_(inst, C) {
  var P = null;
  var cand = [C.evName].concat(EV_ALIAS[C.evName] || []);
  [1, 2].forEach(function (back) {
    if (P) return;
    cand.forEach(function (nm) { if (!P && inst[(C.year - back) + ' ' + nm]) P = inst[(C.year - back) + ' ' + nm]; });
  });
  return P;
}

/**
 * ★ v1.5 核心：去年尖峰週 → 今年檔期內的位置。
 * 對齊基準＝「距檔期結束日幾天」：offset = 去年檔期結束日 − 去年尖峰週結束日。
 * 今年尖峰週結束日 = 今年檔期結束日 − offset；開始日 = 結束日 − (PEAK_WIN−1)。
 * 若超出今年檔期則截在檔期內；截後不足 PEAK_MIN_DAYS 天視為放不進去（ok=false）。
 */
function evMapPeak_(pk, P, C) {
  var offset = Math.round((P.end - pk.end) / 864e5);
  if (offset < 0) offset = 0;
  var e2 = evAddDays_(C.end, -offset);
  var s2 = evAddDays_(e2, -(EV_CFG.PEAK_WIN - 1));
  var clipped = false;
  if (s2 < C.start) { s2 = C.start; clipped = true; }
  if (e2 > C.end) { e2 = C.end; clipped = true; }
  var days = Math.round((e2 - s2) / 864e5) + 1;
  return { start: s2, end: e2, offset: offset, clipped: clipped, ok: days >= EV_CFG.PEAK_MIN_DAYS };
}

/** ★ v1.5 唯讀預覽：列出每個進行中／即將到來的長檔期，節日週會放在哪裡。不寫任何表。 */
function evPeakPreview_() {
  var ss = SpreadsheetApp.openById(EV_BOM_ID);
  var today = evDateOnly_(new Date());
  var map = evReadObjs_(ss, '產品名稱對照表'), pos = evReadObjs_(ss, 'POS資料');
  var ctx = evBuildContext_(map, pos, today, []);
  var lookEnd = evAddDays_(today, EV_CFG.LOOKAHEAD_DAYS);
  var out = ['evPeakPreview ' + today];
  Object.keys(ctx.inst).forEach(function (k) {
    var C = ctx.inst[k];
    if (C.end < evAddDays_(today, -35) || C.start > lookEnd) return;
    var L = Math.round((C.end - C.start) / 864e5) + 1;
    var P = evFindPrev_(ctx.inst, C);
    var line = k + '（' + evYmd_(C.start) + '~' + evYmd_(C.end) + '，' + L + ' 天）';
    if (!P) { out.push(line + '：無去年樣本'); return; }
    if (L <= EV_CFG.SHORT_MAX) { out.push(line + '：短檔期，走整段係數'); return; }
    var pk = evPeakWeek_(P, ctx.dayTotal);
    if (!pk) { out.push(line + '：去年 ' + P.year + ' 檔期 POS 無資料'); return; }
    var mp = evMapPeak_(pk, P, C);
    out.push(line + '：去年 ' + P.year + ' 檔期 ' + evYmd_(P.start) + '~' + evYmd_(P.end) + '，尖峰週 ' + evYmd_(pk.start) + '~' + evYmd_(pk.end) + ' ×' + pk.coef
      + ' → 今年 ' + evYmd_(mp.start) + '~' + evYmd_(mp.end) + '（檔期結束前 ' + mp.offset + ' 天' + (mp.clipped ? '，已截' : '') + '）'
      + (pk.coef >= EV_CFG.PEAK_MIN_COEF ? (mp.ok ? ' ✅ 會寫入' : ' ⚠️ 放不進檔期') : ' ℹ️ 低於門檻 ' + EV_CFG.PEAK_MIN_COEF + '，寫停用說明列'));
  });
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ═══════════════════════════════════════════════════════════════
 * v1.4：無聲失效 → 有聲（保留）
 * ═══════════════════════════════════════════════════════════════ */

var EV_LAST_DIAG = null;
var EV_ALERTS = [];

function evPosMinDate_(dayTotal) {
  var ks = Object.keys(dayTotal || {});
  if (!ks.length) return null;
  ks.sort();
  return ks[0];
}

function evWarnMsg_(kind, C, P, diag, posMin) {
  var back = C.year - P.year;
  var backTxt = back === 1 ? '去年' : (back === 2 ? '前年' : back + ' 年前');
  var need = diag ? (diag.baseFrom + ' ~ ' + diag.pEnd)
                  : (evYmd_(evAddDays_(P.start, -EV_CFG.BASE_DAYS)) + ' ~ ' + evYmd_(P.end));
  var detail = diag
    ? '｜實得樣本 檔期內 ' + diag.inD + ' 天（需≥3）、基線 ' + diag.baseD + ' 天（需≥' + 7 + '）、基線量 ' + diag.baseQ
    : '｜去年檔期期間在 POS 完全查無日銷資料';
  return '⚠️ 無歷史樣本：加成係數未產出，agg_newitem 該檔期為 0'
       + '｜檔期=' + C.year + ' ' + C.evName + '（' + kind + '路徑）'
       + '｜缺的是「' + backTxt + '」窗口 ' + P.year + ' ' + P.evName
       + '｜需要 POS 涵蓋 ' + need
       + '｜POS資料 目前最舊日期=' + (posMin || '未知')
       + detail
       + '｜處置：確認是否為滾動裁切所致；若是且仍需此檔期加成，請改「手動」填係數並啟用。';
}

/* ===== helpers ===== */
// v1.8.5：BOM 單位 → 使用單位的倍數。與主引擎 程式碼.gs 的 _conv 同規則：
//   查找序＝BOM 原名 → 主檔品名 →（通用）→ 1。不論是不是計數單位，都先查換算表（v1.8.4 以前計數單位直接當 1，是錯的）。
//   warn：查不到換算、且單位不是 g／ml／空白／使用單位／計數單位 ⇒ 提醒（與主引擎「未換算單位」同一套豁免）。
var EV_COUNT_U = { '個': 1, '支': 1, '張': 1, '顆': 1, '片': 1, '包': 1, '罐': 1, '台': 1, '條': 1, '滴': 1, '碗': 1, '杯': 1, '份': 1, '組': 1 };
function evConvFactor_(conv, name2sku, nm, unit) {
  unit = String(unit == null ? '' : unit).trim(); nm = String(nm || '').trim();
  var u = name2sku[nm], alt = u ? String(u.name || '').trim() : '', i, c, ci, cu;
  function amtOf(c) { var a = Number(c['換算量']); return (a > 0) ? a : 1; }
  for (i = 0; i < conv.length; i++) { c = conv[i]; if (String(c['適用品項'] || '').trim() === nm && String(c['單位'] || '').trim() === unit) return { f: amtOf(c), warn: false, unit: unit, key: nm + '|' + unit }; }
  if (alt && alt !== nm) for (i = 0; i < conv.length; i++) { c = conv[i]; if (String(c['適用品項'] || '').trim() === alt && String(c['單位'] || '').trim() === unit) return { f: amtOf(c), warn: false, unit: unit, key: nm + '|' + unit }; }
  for (i = 0; i < conv.length; i++) { c = conv[i]; ci = String(c['適用品項'] || '').trim(); cu = String(c['單位'] || '').trim();
    if ((ci === '（通用）' || ci === '(通用)' || ci === '通用') && cu === unit) return { f: amtOf(c), warn: false, unit: unit, key: nm + '|' + unit }; }
  var ok = (unit === 'g' || unit === 'ml' || unit === '' || EV_COUNT_U[unit] || (u && unit === String(u.useUnit || '').trim()));
  return { f: 1, warn: !ok, unit: unit, key: nm + '|' + unit };
}
function evReadObjs_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) return [];
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var h = v[0].map(function (x) { return String(x || '').trim(); });
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, any = false;
    for (var j = 0; j < h.length; j++) { o[h[j]] = v[i][j]; if (v[i][j] !== '' && v[i][j] != null) any = true; }
    if (any) out.push(o);
  }
  return out;
}
function evEnsure_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, headers.length).setValues([headers]); }
  return sh;
}
/** ★ v1.6.1：整張複製到專用檔（值）。開不到專用檔或任一張失敗都只回文字，不丟例外。 */
function evMirror_(src, names) {
  var dst; try { dst = SpreadsheetApp.openById(EV_DASH_ID); } catch (e) { return ['⚠️ 專用檔開不到：' + e]; }
  try { if (dst.getSpreadsheetTimeZone && dst.getSpreadsheetTimeZone() !== 'Asia/Taipei') dst.setSpreadsheetTimeZone('Asia/Taipei'); } catch (e0) {}   /* v1.6.2：時區保險（見 Code.gs dashTz_） */
  var out = [];
  names.forEach(function (name) {
    try {
      var sh = src.getSheetByName(name); if (!sh) { out.push(name + '：來源無此分頁'); return; }
      var v = sh.getDataRange().getValues();
      var t = dst.getSheetByName(name) || dst.insertSheet(name);
      t.clearContents();
      evEnsureRows_(t, v.length); if (v.length && t.getMaxColumns() < v[0].length) t.insertColumnsAfter(t.getMaxColumns(), v[0].length - t.getMaxColumns());   /* v1.8：列數不足先加 */
      if (v.length && v[0].length) t.getRange(1, 1, v.length, v[0].length).setValues(v);
      out.push(name + '：' + (v.length - 1) + ' 列');
    } catch (e2) { out.push(name + '：失敗 ' + e2); }
  });
  return out;
}
/** ★ v1.8.2：agg_endcut 欄位格式（插欄時會沿用左邊日期格式，份數佔比會被當日期；每次寫完都重設） */
function evFormatEndcut_(sh) {
  try {
    var n = Math.max(1, sh.getMaxRows() - 1), mc = sh.getMaxColumns();
    var fm = [[3, 'yyyy-mm-dd'], [4, '0.000'], [5, '0.0'], [6, '@'], [7, 'yyyy-mm-dd'], [8, '0.000'], [9, '@'], [10, '0.000']];   /* v1.8.4 ＋J 近7天比 */
    fm.forEach(function (x) { if (x[0] <= mc) sh.getRange(2, x[0], n, 1).setNumberFormat(x[1]); });
  } catch (e) {}
}

/** ★ v1.8：確保分頁至少有 n 列（寫入前用；小表建表時會被裁到 300 列） */
function evEnsureRows_(sh, n) {
  try { var mr = sh.getMaxRows(); if (mr < n) sh.insertRowsAfter(mr, n - mr); } catch (e) {}
}

/**
 * ★ v1.8 收尾扣減（純計算，不讀寫試算表）
 * 回傳 { rows: [[店號, sku_id, 結束日, 佔比, 近4週週用量, 甜點, 更新日]], ending: {甜點: 結束日} }
 * 佔比＝近 28 天「該店該料」用量中，來自某結束日那批甜點的比例（同一料同一結束日多款甜點合併）
 */
function evEndCut_(map, pos, bomByDessert, name2sku, conv, today, log) {
  var nowTs = evYmd_(today);
  var backFrom = evAddDays_(today, -EV_CFG.ENDCUT_BACK_DAYS), aheadTo = evAddDays_(today, EV_CFG.ENDCUT_AHEAD_DAYS);
  var recentFrom = evYmd_(evAddDays_(today, -EV_CFG.RECENT_DAYS));
  // ① 甜點結束日（對照表；同名多列取最晚，任一列延續到觀察窗之後＝不算下架）
  var posToCanon = {}, endOf = {}, subOf = {};
  map.forEach(function (r) {
    var p = String(r['POS 資料產品名稱'] || '').trim(), m = String(r['手動對應產品名稱'] || '').trim();
    var nm = m || p; if (!nm) return;
    if (p) posToCanon[p] = nm;
    var st = evDate_(r['起始有效日']), en = evDate_(r['結束有效日']);
    if (st && st > aheadTo) return;                       /* 還沒開始的下一期，不影響目前 */
    var cur = endOf[nm];
    var e = en || new Date(2099, 11, 31);                 /* 沒填結束日＝永久 */
    if (!cur || e > cur) { endOf[nm] = e; subOf[nm] = String(r['次類別'] || '').trim(); }
  });
  // ② 近 28 天 POS × BOM（使用單位）；同時記每款甜點最後銷售日
  var convCache = {}, warned = {};
  function perUse(b) {   /* 與 rebuildEventLayer 的 toUse 同規則（v1.8.5：evConvFactor_），每個「料|單位」只算一次、只提醒一次 */
    var unit = String(b.unit || '').trim(), key = b.mat + '|' + unit;
    if (convCache[key] != null) return convCache[key] * (Number(b.qty) || 0);
    var cf = evConvFactor_(conv, name2sku, b.mat, unit);
    if (cf.warn && !warned[key]) { warned[key] = 1; log.push('⚠️ 收尾扣減 未換算單位: ' + b.mat + ' / ' + unit + '（暫以 1 計）'); }
    convCache[key] = cf.f;
    return cf.f * (Number(b.qty) || 0);
  }
  var tot = {}, endUse = {}, lastSale = {}, storePortions = {}, storeDessertQty = {};   /* v1.8.1：店層級份數 */
  var tot7 = {}, recent7From = evYmd_(evAddDays_(today, -7));   /* v1.8.4：近 7 天用量 */
  pos.forEach(function (r) {
    var q = Number(r['數量']) || 0; if (!q) return;
    var d = evDate_(r['建立日期']); if (!d) return;
    var y = evYmd_(d); if (y < recentFrom || y >= nowTs) return;
    var raw = String(r['商品名稱'] || '').trim(); if (!raw) return;
    var nm = posToCanon[raw] || raw;
    var rows = bomByDessert[nm]; if (!rows) return;
    if (!lastSale[nm] || y > lastSale[nm]) lastSale[nm] = y;
    var stn = String(r['分店代碼'] || ''); if (!stn) return;
    storePortions[stn] = (storePortions[stn] || 0) + q;
    (storeDessertQty[stn] = storeDessertQty[stn] || {})[nm] = ((storeDessertQty[stn] || {})[nm] || 0) + q;
    rows.forEach(function (b) {
      var u = name2sku[b.mat]; if (!u) return;
      if (u.cat === '器具' || u.cat === '模具') return;
      var use = q * perUse(b); if (!use) return;
      var k = stn + '|' + u.id;
      tot[k] = (tot[k] || 0) + use;
      if (y >= recent7From) tot7[k] = (tot7[k] || 0) + use;
      (endUse[k] = endUse[k] || {})[nm] = ((endUse[k] || {})[nm] || 0) + use;
    });
  });
  // ③ 哪些甜點算「即將／剛下架」
  var ending = {}, stale = [], soon = [];
  Object.keys(endOf).forEach(function (nm) {
    var e = endOf[nm]; if (e < backFrom || e > aheadTo) return;
    var ls = lastSale[nm] || '';
    if (e < today && ls && ls > evYmd_(evAddDays_(e, EV_CFG.ENDCUT_TAIL_DAYS))) { stale.push(nm + '（對照表 ' + evYmd_(e) + ' 結束，POS 賣到 ' + ls + '）'); return; }
    ending[nm] = evYmd_(e);
    if (e >= today) soon.push(nm + ' ' + evYmd_(e).slice(5) + (/^\d{4}/.test(subOf[nm] || '') ? '' : '（一般品）'));
  });
  // ③-2 在售甜點（v1.8.3 逐列）：每款甜點「最早可賣日」「最晚下架日」（只看還沒結束的對照表列）
  var availFrom = {}, availTo = {}, leadTo = evAddDays_(today, EV_CFG.LEADIN_DAYS);
  map.forEach(function (r) {
    var nm = String(r['手動對應產品名稱'] || '').trim() || String(r['POS 資料產品名稱'] || '').trim(); if (!nm) return;
    var st = evDate_(r['起始有效日']), en = evDate_(r['結束有效日']) || new Date(2099, 11, 31);
    if (en < today) return;
    var s0 = st || today;
    if (!availFrom[nm] || s0 < availFrom[nm]) availFrom[nm] = s0;
    if (!availTo[nm] || en > availTo[nm]) availTo[nm] = en;
  });
  var skuUsers = {};   /* sku_id → [甜點…]（有 BOM 的甜點） */
  Object.keys(bomByDessert).forEach(function (nm) {
    if (!availFrom[nm] || availFrom[nm] > leadTo) return;   /* 14 天內不會開賣、或已經沒有有效列 */
    bomByDessert[nm].forEach(function (b) {
      var u = name2sku[b.mat]; if (!u) return;
      var a = skuUsers[u.id] = skuUsers[u.id] || [];
      if (a.indexOf(nm) < 0) a.push(nm);
    });
  });
  function otherActiveFor(id, endYmd) {   /* 比這一列（下架日 endYmd）多賣至少 GAP 天的甜點 */
    var lim = evAddDays_(evDate_(endYmd), EV_CFG.ENDCUT_OA_GAP_DAYS);
    return (skuUsers[id] || []).filter(function (nm) { return availTo[nm] > lim; });
  }

  // ④ 輸出：店 × 料 × 結束日（v1.8.1：附店層級份數佔比、在售甜點）
  var portionShare = {};   /* 店 → { 結束日: 份數佔比 } */
  Object.keys(storeDessertQty).forEach(function (stn) {
    var T = storePortions[stn] || 0; if (!(T > 0)) return;
    Object.keys(storeDessertQty[stn]).forEach(function (nm) {
      var e = ending[nm]; if (!e) return;
      (portionShare[stn] = portionShare[stn] || {})[e] = ((portionShare[stn] || {})[e] || 0) + storeDessertQty[stn][nm] / T;
    });
  });
  var out = [];
  Object.keys(endUse).forEach(function (k) {
    var t = tot[k]; if (!(t > 0)) return;
    var byEnd = {};
    Object.keys(endUse[k]).forEach(function (nm) {
      var e = ending[nm]; if (!e) return;
      var o = byEnd[e] = byEnd[e] || { use: 0, who: [] };
      o.use += endUse[k][nm]; o.who.push(nm);
    });
    Object.keys(byEnd).forEach(function (e) {
      var sh = byEnd[e].use / t; if (sh < EV_CFG.ENDCUT_MIN_SHARE) return;
      var p = k.split('|');
      var ps = ((portionShare[p[0]] || {})[e]) || 0;
      var oa = otherActiveFor(p[1], e);   /* v1.8.3 逐列 */
      var r7 = Math.round(Math.min(1, (tot7[k] || 0) * 4 / t) * 1000) / 1000;   /* v1.8.4 近7天比 */
      out.push([p[0], p[1], e, Math.round(Math.min(1, sh) * 1000) / 1000, Math.round(byEnd[e].use / (EV_CFG.RECENT_DAYS / 7) * 10) / 10, byEnd[e].who.join('、'), nowTs, Math.round(Math.min(1, ps) * 1000) / 1000,
        oa.length ? (oa.slice(0, 3).join('、') + (oa.length > 3 ? ' 等 ' + oa.length + ' 款' : '')) : '', r7]);
    });
  });
  out.sort(function (a, b) { return (+a[0]) - (+b[0]) || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0) || (a[2] < b[2] ? -1 : 1); });
  /* ★ v1.8.1 瘦身：同一店×料依結束日累加，算「替代後最多會少備幾成」；不到 ENDCUT_MIN_EFFECT 整組不寫；
     在售甜點只留給合計用料佔比 ≥95% 的專屬料（其他情況前端用不到） */
  var kept = [], dropPairs = 0, grp = {};
  out.forEach(function (r) { var k = r[0] + '|' + r[1]; (grp[k] = grp[k] || []).push(r); });
  Object.keys(grp).forEach(function (k) {
    var L = grp[k], S = 0, P = 0, minF = 1;
    L.forEach(function (r) { S = Math.min(1, S + r[3]); P += r[7]; var f = P >= 0.95 ? Math.max(0, 1 - S) : Math.max(0, Math.min(1, (1 - S) / (1 - P))); if (f < minF) minF = f; });
    if (1 - minF < EV_CFG.ENDCUT_MIN_EFFECT) { dropPairs++; return; }
    L.forEach(function (r) { if (S < 0.95) r[8] = ''; kept.push(r); });
  });
  kept.sort(function (a, b) { return (+a[0]) - (+b[0]) || (a[1] < b[1] ? -1 : a[1] > b[1] ? 1 : 0) || (a[2] < b[2] ? -1 : 1); });
  log.push('✂️ 收尾扣減：替代後影響不到 ' + Math.round(EV_CFG.ENDCUT_MIN_EFFECT * 100) + '% 的店×料 ' + dropPairs + ' 組不寫（瘦身）');
  out = kept;
  log.push('✂️ 收尾扣減：即將下架 ' + soon.length + ' 款' + (soon.length ? '（' + soon.slice(0, 12).join('、') + (soon.length > 12 ? '…' : '') + '）' : '') + '；剛下架仍在近4週均內 ' + (Object.keys(ending).length - soon.length) + ' 款');
  if (stale.length) log.push('⚠️ 對照表結束日已過卻還在賣（不扣減，請更新對照表）：' + stale.slice(0, 10).join('、') + (stale.length > 10 ? '…共 ' + stale.length + ' 款' : ''));
  return { rows: out, ending: ending, stale: stale };
}

/** ★ v1.6：建小表——新分頁預設 1000 列 × 26 欄＝26,000 格；裁到 300 列 × 表頭欄數，省 BOM 本配額（地雷 3.5）。 */
function evEnsureSmall_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    try {
      var mc = sh.getMaxColumns(); if (mc > headers.length) sh.deleteColumns(headers.length + 1, mc - headers.length);
      var mr = sh.getMaxRows(); if (mr > 300) sh.deleteRows(301, mr - 300);
    } catch (e) {}
  }
  return sh;
}
function evDate_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : evDateOnly_(v);
  var d = new Date(String(v).replace(/\//g, '-'));
  return isNaN(d) ? null : evDateOnly_(d);
}
function evDateOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function evYmd_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function evAddDays_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function evShiftYear_(d, n) { return new Date(d.getFullYear() + n, d.getMonth(), d.getDate()); }  // v1.5 起不再用於節日週，保留供舊呼叫
function evUplift_(P, dayTotal, instDay) {
  var pk = P.year + ' ' + P.evName;
  var inQ = 0, inD = 0, baseQ = 0, baseD = 0;
  for (var d = new Date(P.start); d <= P.end; d = evAddDays_(d, 1)) {
    var y = evYmd_(d), t = dayTotal[y];
    if (t == null) continue;
    inQ += t - ((instDay[pk] || {})[y] || 0); inD++;
  }
  var b0 = evAddDays_(P.start, -EV_CFG.BASE_DAYS);
  for (var d2 = b0; d2 < P.start; d2 = evAddDays_(d2, 1)) {
    var y2 = evYmd_(d2), t2 = dayTotal[y2];
    if (t2 == null) continue;
    baseQ += t2; baseD++;
  }
  EV_LAST_DIAG = { inD: inD, baseD: baseD, baseQ: baseQ, inQ: inQ,
                   pYear: P.year, pName: P.evName,
                   pStart: evYmd_(P.start), pEnd: evYmd_(P.end),
                   baseFrom: evYmd_(evAddDays_(P.start, -EV_CFG.BASE_DAYS)) };
  if (inD < 3 || baseD < 7 || !baseQ) return null;
  var c = (inQ / inD) / (baseQ / baseD);
  c = Math.max(0.5, Math.min(EV_CFG.MAX_COEF, c));
  return Math.round(c * 100) / 100;
}
function evPeakWeek_(P, dayTotal) {
  var days = [], d;
  for (d = new Date(P.start); d <= P.end; d = evAddDays_(d, 1)) days.push({ d: new Date(d), q: dayTotal[evYmd_(d)] || 0 });
  if (days.length < EV_CFG.PEAK_WIN + 3) return null;
  var tot = days.reduce(function (a, x) { return a + x.q; }, 0);
  if (!tot) return null;
  var avg = tot / days.length, best = null;
  for (var i = 0; i + EV_CFG.PEAK_WIN <= days.length; i++) {
    var s = 0;
    for (var j = i; j < i + EV_CFG.PEAK_WIN; j++) s += days[j].q;
    if (!best || s > best.s) best = { s: s, i: i };
  }
  var coef = (best.s / EV_CFG.PEAK_WIN) / avg;
  coef = Math.max(1, Math.min(EV_CFG.MAX_COEF, coef));
  return { start: days[best.i].d, end: days[best.i + EV_CFG.PEAK_WIN - 1].d, coef: Math.round(coef * 100) / 100 };
}