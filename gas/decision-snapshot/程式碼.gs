/**
 * diybc-decision-snapshot ｜ DIYBC 經營快照（AI 決策分身 P0 階段 c＋d：每日經營早報＋d2：逾時修正／巡檢員／舊資料警示）
 * ------------------------------------------------------------------
 * 用途：把 5 類經營數字算好寫進「DIYBC 經營快照」Sheet（唯一可寫入的檔），供 Chat 讀取。
 * 規則：
 *   - 既有 Sheet 一律唯讀（只用 getValues）；既有 GAS 端點只打讀取 action。
 *   - 不存客人個資：團體 GAS 的回傳只在記憶體裡算各店計數，原始內容不寫入任何分頁、log、報告。
 *   - 團體通關碼只從「指令碼屬性」GB_PASSCODE 讀，不寫進程式、不寫進 log／meta／錯誤訊息。
 *   - 觸發器只建在本專案：每日 runAll（經營者／Chat 核准的時間）＋ 分段接續用的一次性 runAllContinue（跑完自刪）。
 *   - 照抄的前端函式：區塊前註明「來源檔名＋md5＋原行號」，邏輯一字不改。
 *   - 經營早報只寄 dim_rule BRIEF_TO（1 人，預設 mydiybc@gmail.com），不加 CC／BCC；內容只讀 alerts／snap_latest 排版，不改算法。
 *   - d2 時間管理：每類開始前看「已用秒數＋該類預估秒數（EST_A～EST_E）」是否超過 RUN_BUDGET_SEC；讀資料前超過 270 秒就「延後」到下一段重跑該類（單次讀取最長實測 61 秒，留 90 秒緩衝）；
 *     同一類連續延後 2 次改標失敗；接續段最多 8 段。被 Google 強制中止時程式無法寄信 → 每天 12:15 的巡檢員 dailyWatchdog 補跑／補寄／通報。
 * 入口：runAll（每日，完成後寄經營早報）｜dailyWatchdog（每日巡檢）／watchdogRecheck（巡檢 20 分鐘後複查，一次性）
 *       sendBriefNow（強制重寄）／previewBrief（只看不寄）／briefTestOn・briefTestOff｜snapA_staffing / snapB_campaign / snapC_booking / snapD_reviews / snapE_members（單類手動）
 *       setupSnapshotSheet（第一次）｜setupDailyTrigger（建每日 runAll＋巡檢員觸發器，可重跑）
 *       diagState（印出執行狀態值，不含通關碼）｜warmCacheLy（預先把去年訂位月資料存進 cache_ly）｜zzTest…（測試旗標，正式時全部關閉）
 * 產出：2026-09-25 階段d2（小修：檢查線 270／預算 240、D_last 措辭） ｜ 內容 sha256（不含本行）de65cdf981c4108f
 */

// ============================================================
// 0. 設定
// ============================================================
var TZ = 'Asia/Taipei';
var SNAP_NAME = 'DIYBC 經營快照';
var PROP_SNAP_ID = 'SNAP_SS_ID';

var SRC = {
  SCH_API: 'https://script.google.com/macros/s/AKfycbwTE1H-uh6mtsoZP_HKscF8XGgpv8Mxu6lH1M8Xrbp0kjd8M2OR1sUtPZrnSVBAJkF6Og/exec',
  POS_API: 'https://script.google.com/macros/s/AKfycbxvD8dQdWJdaK9Ih9ZZc3PRT8rICcqbCrk5cKQY3s6zF1MmZgWL2g803zrJTBsuE3p4pQ/exec',
  RSV_API: 'https://script.google.com/macros/s/AKfycbzjvgoj6MQoDizE-322wOLyOvq4P3HVbOlvyHBszdyfuOrYPaSnDa48vQb_34LLgEJh/exec',
  REV_API: 'https://script.google.com/macros/s/AKfycbwoOqVB03froYy4frO4dUDUEPN3YfURH4bx1H71PWIYyoN5C0Pk16rBX-TblKxWgN91/exec',
  GB_API:  'https://script.google.com/macros/s/AKfycbwCisp7PrQFsu_WJsv9_KNsEYAs2Ki_B13WSBFrjNsxq8fREamX27D6IABk26Iy5E9FpQ/exec',   // 團體訂位 GAS（只打 fn=group 讀取；通關碼讀指令碼屬性 GB_PASSCODE）
  GB_SS:   '1NrMGw6BOvMJigKfdAFNM1Np6KCa5cd6SEIBqzrmFLAQ',   // DIYBC 團體訂位（gb_bookings，只讀 date/store/size/status/scanned_at 5 欄）
  REV_SS:  '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw',   // Google評論_每日新增（只讀 C 欄評論時間，取資料最新日）
  MEM_SS:  '1hZ8hpzrCNycrpm64Ge0fVxWDV-L__bxiagXk0btrBBQ',   // 自己人消費記錄分析（2026 各店新增自己人、agg_store_kpi_365）
  PUR_DASH:'1FF7lW3JINR0-Id7MMYSRkoRzA1BYdqktO94NbzAFYG0',   // 採購儀表板專用檔
  PUR_MONTH:'19tG-AMYtiZPTYQcz85TGiGS2G4M5U1gZw9Mab2UFFU0'   // 採購月檔（fact_shopline）
};

var TAB = { LATEST:'snap_latest', HIST:'snap_history', ALERTS:'alerts', RULE:'dim_rule', MAP:'dim_store_map', META:'meta', STAGE:'_stage', AHIST:'alerts_history', LY:'cache_ly', TRACE:'_trace', SERIES:'snap_series' };
var MAIL_TO = 'mydiybc@gmail.com';            // M9 異常通知收件人（交辦書指定）
var HIST_KEEP_DAYS = 400;                      // M6 snap_history 保留天數
var PROP_GB_PASS = 'GB_PASSCODE';              // 團體通關碼（經營者親手填在指令碼屬性）
var PROP_RUN_STATE = 'SNAP_RUN_STATE';         // M8 分段接續的執行狀態
var CATS = [['A', '人力營收'], ['B', '檔期叫貨'], ['C', '訂位'], ['D', '評論'], ['E', '自己人']];
var HDR = {
  snap_history: ['快照日','店號','指標代碼','期間','數值','比較值','比較類型','資料最新日','燈號','執行批次'],
  alerts:       ['燈號','店號','類別','說明','數值','門檻'],
  dim_rule:     ['規則代碼','類別','說明','數值','單位'],
  dim_store_map:['店號','標準店名','區','POS','排班','訂位','評論','自己人年表','採購'],
  meta:         ['執行批次','類別','來源','讀取項目','耗時秒','筆數','資料最新日','結果','錯誤訊息'],
  alerts_history:['快照日','燈號','店號','類別','說明','數值','門檻','狀況鍵'],   // D5（隱藏分頁）
  cache_ly:     ['年月','日期','店號','有效筆數','有效人數','存入時間'],             // d2 D2-5（隱藏分頁）：去年訂位按日按店合計，只存計數、不存個資
  _trace:       ['時間','類別','步驟','開始/結束','耗時秒','執行批次'],             // d3 D3-1（隱藏分頁）：每次讀取的開始／結束，只記步驟名與秒數、不存回傳內容
  snap_series:  ['日期','店號','淨營收','去年同日淨營收','工時','更新批次']           // e1（分頁）：各店每日趨勢（A 類順便算，同日同店覆蓋；給儀表板 e3 用）
};
// ---- d2：時間管理與狀態值 ----
var RUN_HARD_SEC = 270;          // D2-2 讀資料前已用超過此秒數 → 延後（GAS 單次上限 360 秒；單次讀取最長實測 61 秒，270＋61＜360）
var RUN_FINAL_SEC = 45;          // 寫入快照＋寄早報預留秒數：已用＋此值 > RUN_HARD_SEC 就把寫入交給下一段
var RUN_MAX_SEGMENTS = 8;        // D2-2 接續段總數上限
var DEFER_MAX = 2;               // D2-2 同一類連續延後幾次改標失敗
var EST_KEEP = 7;                // D2-1 預估秒數＝近 7 次最大值 × 1.3
var DEFER_TAG = '【延後】';
var PROP_EST_HIST = 'SNAP_EST_HIST';      // 各類近 7 次實際秒數
var PROP_CAT_DONE = 'SNAP_CAT_DONE';      // 各類最後一次算完的時間 {A:'yyyy-MM-dd HH:mm:ss',…}
var PROP_DONE = 'SNAP_DONE';              // 最後一次 runAll 寫完快照 {date, at, ok:[], fail:[{cat,err}]}
var PROP_FRESH = 'SNAP_FRESH';            // 各類最後一次的資料新鮮度（補跑部分類別時，「0 資料」警示沿用）
var PROP_TEST = 'SNAP_TEST';              // 測試旗標（JSON），正式時不存在
var PROP_BRIEF_SENT_TEST = 'BRIEF_SENT_TEST_DATE';   // 測試模式下的「今天已寄早報」
var EXEC_T0 = 0;                          // 本次執行（本段）開始時間；0＝不做時間檢查（單類手動）
// ---- d3：保險絲與追查記錄 ----
var FUSE_MIN = 5;                         // D3-2 每段一開始就預約 5 分鐘後的接續段（保險絲）；本段正常結束時照舊刪除／改排
var STUCK_MSG = '讀取卡住超過 6 分鐘，已連續 2 次';   // D3-2 同一類卡住（或卡住＋延後）合計 2 次 → 標失敗
var TRACE_KEEP_DAYS = 14;                 // D3-1 _trace 保留天數
var REV_TAIL_ROWS = 500;                  // D3-3 評論 Sheet C 欄只讀最後 500 列

var RULE_DEFAULTS = [
  ['SCH_WD_HOURS','A 排班','平日工時上限（computeFlags：weekday_hours ÷ 5 超過即「平日工時過高」）',30,'小時'],
  ['SCH_WE_HOURS','A 排班','假日工時上限（weekend_hours ÷ 2 超過即「假日工時過高」）',50,'小時'],
  ['SCH_RPH_MIN','A 排班','週 rev/h 下限（低於即「人效偏低」）',950,'元/小時'],
  ['SCH_REV_MIN','A 排班','週營收下限（低於即「營收偏低」）',50000,'元'],
  ['YOY_ORANGE','A 同期成長','同期成長率低於此值 → 🟠',-10,'%'],
  ['YOY_RED','A 同期成長','同期成長率低於此值 → 🔴',-20,'%'],
  ['CAMP_SHIP_ORANGE','B 檔期','單一品項 已出貨÷推估總需求 高於此值 → 🟠',130,'%'],
  ['CAMP_SHIP_RED','B 檔期','單一品項 已出貨÷推估總需求 高於此值 → 🔴',180,'%'],
  ['CAMP_ZERO_NEED_RED','B 檔期','推估總需求 0 但已出貨 → 🔴（1＝啟用，0＝關閉；只看檔期專屬品，且已出貨金額 ≥ CAMP_MIN_EXCESS_NTD）',1,''],
  ['CAMP_MIN_EXCESS_NTD','B 檔期','超標門檻：超出金額（已出貨金額 − 推估需求金額）≥ 此值才算超標；只看檔期專屬品，共用品不亮燈',500,'元'],
  ['INV_DAYS_YELLOW','B 盤點','最後盤點超過幾天（或從未盤點）→ 🟡',14,'天'],
  ['RSV_CANCEL_UP_ORANGE','C 訂位','本月取消率高於該店上月幾個百分點 → 🟠',10,'百分點'],
  ['GB_URGENT_RED','C 團體','團體需立即處理 ≥ 此值 → 🔴（來源：團體 GAS，與團體頁「🚨 需立即處理」同口徑）',1,'項'],
  ['GB_NODESSERT_ORANGE','C 團體','團體未選甜點 ≥ 此值 → 🟠（來源：團體 GAS，與團體頁「還沒選甜點」同口徑）',1,'筆'],
  ['REV_BEHIND_ORANGE','D 評論','有效評論落後本月進度目標 ≥ 此則數 → 🟠',3,'則'],
  ['REV_LAST_DAYS','D 評論','月底前幾天內',7,'天'],
  ['REV_LAST_MIN','D 評論','月底前 N 天仍低於此則數 → 🔴',10,'則'],
  ['MEM_BEHIND_ORANGE','E 自己人','累計達成率低於年度進度線幾個百分點 → 🟠',10,'百分點'],
  ['MEM_BEHIND_RED','E 自己人','累計達成率低於年度進度線幾個百分點 → 🔴',20,'百分點'],
  ['MEM_STALE_DAYS','E 自己人','年表本月新增數連續幾天沒變 → 🟡「年表可能停更」（列在 0 資料）',2,'天'],
  ['FRESH_SALES_LAG_DAYS','0 資料','排班日營收或 POS 最新日，比今天早超過幾天 → 🟡（1＝早於昨天就亮）',1,'天'],
  ['FRESH_FUT_LAG_DAYS','0 資料','未來訂位資料最早日期，比今天早超過幾天 → 🟡「未來訂位停更」（0＝早於今天就亮）',0,'天'],
  ['FRESH_SHOP_LAG_DAYS','0 資料','大平台下單（fact_shopline）最新日落後超過幾天 → 🟡',3,'天'],
  ['FRESH_REV_LAG_DAYS','0 資料','評論最新日，比今天早超過幾天 → 🟡（2＝早於前天就亮）',2,'天'],
  ['RUN_SPLIT_SEC','系統','（d2 起停用，改用 RUN_BUDGET_SEC 與 EST_A～EST_E；保留列不刪）',270,'秒'],
  ['RUN_BUDGET_SEC','系統','每段時間預算：已用秒數＋下一類預估秒數超過此值就不開始，交給 1 分鐘後的接續段（讀資料前的檢查線 270 秒，GAS 上限 360 秒）',240,'秒'],
  ['EST_A','系統','A 人力營收 預估秒數（每次算完自動更新＝近 7 次最大值 × 1.3）',80,'秒'],
  ['EST_B','系統','B 檔期叫貨 預估秒數（每次算完自動更新＝近 7 次最大值 × 1.3）',30,'秒'],
  ['EST_C','系統','C 訂位 預估秒數（每次算完自動更新＝近 7 次最大值 × 1.3）',60,'秒'],
  ['EST_D','系統','D 評論 預估秒數（每次算完自動更新＝近 7 次最大值 × 1.3）',120,'秒'],
  ['EST_E','系統','E 自己人 預估秒數（每次算完自動更新＝近 7 次最大值 × 1.3）',40,'秒'],
  ['BRIEF_TO','早報','每日經營早報收件人（只能填 1 個 Email；格式不對會改寄 mydiybc@gmail.com）','mydiybc@gmail.com',''],
  ['BRIEF_ON','早報','1＝每天 runAll 完成後寄經營早報；0＝暫停（M9 異常信照舊）',1,''],
  ['BRIEF_MAX_RED','早報','早報 🔴 今天要處理 最多幾行（超過時最後一行寫「另有 N 則」）',6,'行'],
  ['BRIEF_MAX_ORANGE','早報','早報 🟠 本週留意 最多幾行',5,'行'],
  ['BRIEF_MAX_YELLOW','早報','早報 🟡 資料提醒 最多幾行',4,'行'],
  ['C3_LIGHT_START','C 訂位','C3 未來 14 天訂位（領先指標）開始亮燈的日期（yyyy-MM-dd）；空白＝不亮燈、不進 alerts（e1 起先累積 28 天基準）','',''],
  ['PAGE_STALE_RED_AFTER','儀表板','今日決策頁：快照不是今天時，幾點（HH:mm，台灣時間）以後改成紅色提示條；在這之前只顯示灰色小字「今天的資料 11:00 更新」','11:30','']
];
var RULE_TEXT = { BRIEF_TO: 1, C3_LIGHT_START: 1, PAGE_STALE_RED_AFTER: 1 };   // 文字型規則（其餘都是數字）

// 店名對照（依 2026-09-24 各來源實際店名逐一核對；經營者可在 dim_store_map 分頁修改）
var STORE_MAP_DEFAULTS = [
  [1,'台中精明店','中南區','台中精明店','精明店','台中精明店','台中精明店','台中精明店','1'],
  [2,'台中草悟道店','中南區','台中草悟道店','草悟道店','台中草悟道店','台中草悟道店','台中草悟道店','2'],
  [3,'台北南京店','北一區','台北南京店','南京店','台北南京店','台北南京店','台北南京店','3'],
  [4,'台北士林店','北一區','台北士林店','士林店','台北士林店','台北士林店','台北士林店','4'],
  [5,'台南Focus店','中南區','台南Focus店','Focus店','台南Focus店','台南focus店','台南Focus店','5'],
  [6,'新竹文化店','中南區','新竹文化店','新竹店','新竹文化店','新竹文化店','新竹文化店','6'],
  [7,'新北板橋店','北二區','新北板橋店','板橋店','新北板橋店','新北板橋店','新北板橋店','7'],
  [8,'新北新店店','北一區','新北新店店','新店店','新北新店店','新北新店店','新北新店店','8'],
  [9,'桃園中壢店','北三區','桃園中壢店','中壢店','桃園中壢店','桃園中壢店','桃園中壢店','9'],
  [10,'桃園藝文店','北三區','桃園藝文店','桃園店','桃園藝文店','桃園藝文店','桃園藝文店','10'],
  [11,'台北遠百信義A13店','北二區','吳寶春自己做台北信義A13店','信義A13店','台北遠百信義A13店','吳寶春(台北信義A13)','台北遠百信義A13店','11'],
  [12,'高雄SKM Park店','中南區','吳寶春自己做高雄SKM Park店','高雄SKM店','高雄SKM Park店','吳寶春(高雄SKM)','高雄SKM Park店','12']
];
var ALL_ROW = '全公司';

// ============================================================
// 1. 照抄的前端函式（邏輯一字不改；每段註明來源檔名＋md5＋原行號）
// ============================================================
// ---- 照抄：google-reviews.html ｜ md5 d6841306bc53326f6bd7e32ee46e9cf3 ｜ 原行號 L411–L420 ｜ 評論：去除 Google 翻譯（邏輯一字不改）----
function cleanGoogleTranslation(text) {
if (!text) return text;
// 模式 A: (Translated by Google) 翻譯 (Original) 原文 → 取原文
let m = text.match(/^\(Translated by Google\)[\s\S]*?\(Original\)\s*([\s\S]+)/);
if (m) return m[1].trim();
// 模式 B: 原文 (Translated by Google) 翻譯 → 取原文
m = text.match(/^([\s\S]+?)\s*\(Translated by Google\)/);
if (m) return m[1].trim();
return text;
}

// ---- 照抄：google-reviews.html ｜ md5 d6841306bc53326f6bd7e32ee46e9cf3 ｜ 原行號 L648、L653–L673 ｜ 評論：每月目標、SERVICE_KW、NAME_STOP、validContent、hasPersonName、isValidReview（邏輯一字不改）----
const MONTHLY_REVIEW_TARGET = 15;  // 各店每月「有效評論」目標則數（單月固定）
const SERVICE_KW = ['服務','服务','店員','店员','員工','员工','老闆','老板','店長','店长','小幫手','小帮手','工作人員','工作人员','人員','人员','師傅','师傅','同仁','店家','櫃檯','柜台','服務員','服务员','小姐','小哥','親切','亲切','熱心','热心','熱情','热情','熱忱','热忱','耐心','細心','细心','貼心','贴心','用心','友善','專業','专业','態度','态度','周到','大方','幫忙','帮忙','幫助','帮助','協助','协助','支援','教學','教学','教導','教导','講解','讲解','指導','指导','解說','解说','引導','引导','示範','示范','招待','人很好','人都很好','人超好','人真好','人蠻好','人满好','人不錯','人不错','nice','friendly','staff','service','patient','helpful','kind'];
// 大寫英文詞中「不是人名」的常見字，需排除（其餘大寫英文一律從寬視為人名）
const NAME_STOP = new Set(['THE','AND','THIS','THAT','THEY','WAS','WERE','ARE','OUR','YOU','YOUR','HER','HIS','ITS','FOR','HAD','HAS','HAVE','VERY','GOOD','NICE','GREAT','FUN','BUT','ALL','CAN','WILL','SHE','HIM','NOT','AMAZING','AWESOME','REALLY','SUPER','LOVE','LOVED','YES','WOW','THANK','THANKS','GOOGLE','ORIGINAL','TRANSLATED','WOULD','COULD','SHOULD','HERE','THERE','WHEN','WHAT','TIME','MADE','MAKE','RECOMMEND','HIGHLY','DEFINITELY','EXPERIENCE','STAFF','SERVICE','ALSO','EVEN','JUST','ONLY','PLACE','CAKE','DIY','SKM','WITH','FROM','GOT','AGAIN','BEEN','THEM','WHO','HOW','OUT','WALKED','EVENTUALLY','EASY','UNLIMITED','ACCOMODATING','ACCOMMODATING','ESPECIALLY','NEITHER','HOWEVE','HOWEVER','WHENEVER','SEEING','ANOTHER','SOMETIMES','OVERALL','WHETHER','TAIPEI','BAKERY','SMALL','PIE','TARTS','BLACK','CHIFFON','VALENTINE','DAY','TEAM','TODAY','FIRST','THANKYOU','INSTAGRAM','PAD','PADS']);
function validContent(r){ return cleanGoogleTranslation(String(r.text||'')).trim(); }   // 5星非空白判定用清理後內容
function hasPersonName(c, store){
  if (allowNames(c, store).length > 0) return true;                          // 分店夥伴綽號白名單（2026-08-17）
  if (/謝謝|谢谢|感謝|感谢|多謝|多谢|多虧|多亏|感恩/.test(c)) return true;   // 答謝某人＝有提到人
  const en = c.match(/[A-Z][a-z]{2,}|[A-Z]{3,}/g) || [];                    // 大寫英文名（從寬）
  if (en.some(t => !NAME_STOP.has(t.toUpperCase()) && !storeStopped(store, t))) return true;
  const dup = c.match(/([\u4e00-\u9fff])\1/g) || [];                        // 中文疊字名（與夥伴表同口徑）
  if (dup.some(w => !DUP_NAME_STOP.has(w) && !storeStopped(store, w))) return true;
  return nickCandidates(c, store).length > 0;                                      // 「小X／阿X」小名（2026-08-13 放寬，與夥伴表同口徑）
}
function isValidReview(r){
  if (r.stars !== 5) return false;
  const c = validContent(r);
  if (!c) return false;                                                    // 空白
  const low = c.toLowerCase();
  if (SERVICE_KW.some(k => low.includes(k.toLowerCase()))) return true;    // 提到服務/事項
  return hasPersonName(c, r.store);                                                 // 或有人名（從寬）
}

// ---- 照抄：google-reviews.html ｜ md5 d6841306bc53326f6bd7e32ee46e9cf3 ｜ 原行號 L812–L855 ｜ 評論：NICK_STOP、nickCandidates、DUP_NAME_STOP、NAME_STOP_BY_STORE、storeStopped、NAME_ALLOW_BY_STORE、allowNames（邏輯一字不改）----
const NICK_STOP = new Set(['小姐','小哥','小弟','小妹','小孩','小朋','小童','小幫','小編','小二','小一','小三','小四','小五','小六','小七','小八','小九','小十','小半','小兩','小時','小心','小小','小學','小人','小費','小吃','小點','小蛋','小餅','小杯','小盒','小包','小袋','小碗','小盤','小份','小塊','小巧','小型','小號','小店','小廚','小資','小貴','小累','小擠','小吵','小熱','小冷','小遠','小慢','小忙','小亂','小趕','小急','小雷','小失','小可','小遺','小驚','小確','小缺','小抱','小尷','小提','小建','小地','小插','小瑕','小知','小遊','小禮','小紅','小計','小結','小了','小的','小喔','小耶','小啊','小唷','小哦','小呢','小吧','小啦','小欸','阿姨','阿嬤','阿公','阿伯','阿婆','阿母','阿爸','阿祖','阿姑','阿舅','阿嫂','阿姊','阿姐','阿兵','阿彌']);
const NICK_PREV_BLOCK = '太很超較偏變縮略稍微挺蠻滿大極頗更最還嫌從自老點';  // 前一字是程度/尺寸詞 →「太小了」「有點小貴」類，非人名
function nickCandidates(t, store){
  const out = [];
  const re = /[小阿][\u4e00-\u9fff]/g;
  let m;
  while ((m = re.exec(t)) !== null) {
    const w = m[0];
    if (NICK_STOP.has(w)) continue;
    if (store && storeStopped(store, w)) continue;               // 分店排除表（2026-08-17）
    const prev = m.index > 0 ? t[m.index - 1] : '';
    if (prev && NICK_PREV_BLOCK.indexOf(prev) >= 0) continue;
    out.push(w);
  }
  return out;
}
const DUP_NAME_STOP = new Set(['謝謝','姐姐','姊姊','哥哥','弟弟','妹妹','爸爸','媽媽','叔叔','爺爺','奶奶','婆婆','公公','推推','好好','讚讚','棒棒','滿滿','哈哈','嘻嘻','呵呵','嘿嘿','啦啦','喔喔','哦哦','嗯嗯','會會','慢慢','剛剛','常常','天天','日日','人人','個個','種種','多多','快快','微微','漸漸','默默','偷偷','悄悄','紛紛','匆匆','輕輕','深深','遠遠','高高','長長','大大','小小','少少','早早','晚晚','久久','團團','整整','樣樣','件件','處處','時時','每每','再再','一一','看看','試試','問問','想想','做做','玩玩','走走','坐坐','聊聊','笑笑','等等','通通','統統','稍稍','明明','白白','空空','緩緩','徐徐','悠悠','款款','聲聲','步步','層層','片片','點點','絲絲','陣陣','串串','包包','盒盒','娃娃','毛毛','乖乖','呀呀','耶耶','蹦蹦','跳跳','體體','体体','谢谢','赞赞','满满','刚刚','种种','会会','个个','长长','远远','点点','层层','处处','时时','样样','试试','问问','团团','声声','纷纷','轻轻','渐渐','谈谈','缓缓','静静']);
// ===== 分店「非夥伴名」排除表（2026-08-17 經營者逐店確認）=====
// 只在該店排除；同一名稱在別店仍視為夥伴（例：Jelly 在精明店非夥伴、他店仍算）。
// 全站都不可能是人名的通用詞（英文句子字、品牌詞）請加全域表 NAME_STOP / EN_NAME_STOP / DUP_NAME_STOP / NICK_STOP。
// key 必須與評論 Sheet「門市」完全一致（半形括號、小寫 focus）。
const NAME_STOP_BY_STORE = {
  '台中精明店': new Set(['小卡','嗚嗚','小狗','熊熊','酥酥','小細','阿勇','抱抱','小帥','Dolce','Grazie','Jelly','小天']),
  '台中草悟道店': new Set(['小漏']),
  '台北南京店': new Set(['美美','小狗','小白','熊熊','重重','順順','利利','小家','往往']),
  '台北士林店': new Set(['美美','小狗','漂漂','亮亮','小卡','小技','小福','學學','小動','醜醜','約約','拍拍','妥妥','當當']),
  '台南focus店': new Set(['小白','可可','小狗','Oreo','小夥','眉眉','角角','小姊','小黑','妙妙','漂漂','亮亮','油油','怪怪','小美']),
  '新竹文化店': new Set(['小卡','小白','可可','舒舒','服服','成成','小黃','她她']),
  '新北板橋店': new Set(['小狗','嗚嗚','漂漂','亮亮','小卡','小老','狗狗','重重','畫畫','小細','平平','貼貼','小尖','小些','是是','小飲','鬆鬆','小手','小夥','嚐嚐']),
  '新北新店店': new Set(['小狗','小白','嗚嗚','小卡','小老','狗狗','熊熊','畫畫','可可','小細','啊啊','店店','小祝','抱抱','噴噴','誇誇','短短','他他','小女','愛愛']),
  '桃園中壢店': new Set(['颼颼','漂漂','亮亮','小卡','小白']),
  '桃園藝文店': new Set(['小細','抱抱','美美','小配','小空','Michael','Jackson','小白','指指','太太','小老']),
  '吳寶春(台北信義A13)': new Set(['小狗','小白','嗚嗚','邊邊','坑坑','疤疤','上上','小作','嘟嘟']),
  '吳寶春(高雄SKM)': new Set(['小卡','小白','可可','Oreo','小夥','小撇','小老','小創','小聲','小鳥','脆脆']),
};
function storeStopped(store, w){ const s = NAME_STOP_BY_STORE[store]; return !!(s && s.has(w)); }
// ===== 分店「夥伴綽號白名單」（2026-08-17）=====
// 三字以上／非疊字／非「小X阿X」的綽號（如 花栗鼠），三條既有人名規則一律漏抓 → 逐店登記在此。
// 命中＝直接視為人名：有效評論判定通過、夥伴排行榜列名（類型標「綽號」）。用「內容包含」比對。
// 新增前先確認確實是該店夥伴綽號；字串越短越容易誤傷（勿放兩字以下通用詞）。
const NAME_ALLOW_BY_STORE = {
  '新北新店店': ['花栗鼠'],
};
function allowNames(t, store){ const l = NAME_ALLOW_BY_STORE[store]; if(!l) return []; return l.filter(w => t.indexOf(w) >= 0); }

// ---- 照抄：dashboard-schedule.html ｜ md5 20a09c9cf1ac2b6c560f2b46f4ebeac1 ｜ 原行號 L763 ｜ 排班：閾值預設值（本專案再以 dim_rule 覆蓋前 4 個）（邏輯一字不改）----
let thresholds = { wdHours: 30, weHours: 50, rph: 950, rev: 50000, dayWd: 500, dayWe: 1000 };

// ---- 照抄：dashboard-schedule.html ｜ md5 20a09c9cf1ac2b6c560f2b46f4ebeac1 ｜ 原行號 L1107–L1142 ｜ 排班：computeFlags（邏輯一字不改）----
// ============ 紅旗判定(以滑桿閾值即時計算) ============
function computeFlags(row) {
  // 注意:row.flags 是後端原始判定(weekday_over_16 / weekend_over_24 / morning_solo / evening_solo / low_rev_per_hour 用固定值)
  // 這裡用滑桿值「重新」判定
  const flags = [];

  // 未完結週(營收=0 或空 但工時>0):班表已上但營收還沒進來的未來週,不該被判定紅旗
  // 直接返回空陣列,避免「人效偏低 / 營收偏低」誤判
  // 注意:用寬鬆比較 < 1(而非 === 0),因 total_revenue 可能是字串 "0"、null 或 undefined
  if ((!row.total_net_revenue || Number(row.total_net_revenue) < 1) && row.total_hours > 0) {
    return flags;
  }

  // 平日工時超標 = 後端 weekday_over_16,但 16 是固定的。
  // 為了讓滑桿真的影響,這裡用 row 的 weekday_hours 平均(週內最高天無法從聚合資料還原,
  // 只能用週平均當代理。這是已知限制)
  // 改用「週內平日平均」: weekday_hours / 5 工作天估算
  const wdAvg = row.weekday_hours / 5;
  if (wdAvg > thresholds.wdHours) flags.push({ k: 'wd_over', label: '平日工時過高', sev: 'red' });

  const weAvg = row.weekend_hours / 2;
  if (weAvg > thresholds.weHours) flags.push({ k: 'we_over', label: '假日工時過高', sev: 'red' });

  if (row.rev_per_hour_net < thresholds.rph && row.total_hours > 0) {
    flags.push({ k: 'low_rph', label: '人效偏低', sev: 'orange' });
  }
  if (row.total_net_revenue < thresholds.rev) {
    flags.push({ k: 'low_rev', label: '營收偏低', sev: 'yellow' });
  }

  // 注:早/晚班獨守(morning_solo / evening_solo)已於 2026-05-18 移除
  // 業務邏輯:DIY 烘焙 SOP 設計為 1 人服務 8-10 客,獨守是常態配置而非警示
  // 真正的人力健康度看 rev/h 即可(low_rph flag),人數本身無意義

  return flags;
}

// ---- 照抄：dashboard-reservation.html ｜ md5 788126396bd55e3e4643c4a16767abb5 ｜ 原行號 L721–L727 ｜ 訂位：remapRows（邏輯一字不改）----
function remapRows(srcCols, srcRows, dstCols) {
  if (!srcRows.length) return [];
  if (srcCols.join('\u0001') === dstCols.join('\u0001')) return srcRows;
  const map = dstCols.map(c => srcCols.indexOf(c));
  console.warn('[訂位儀表板] future 與 fact 欄位順序不同，已依欄名重排');
  return srcRows.map(r => map.map(i => (i < 0 ? '' : r[i])));
}

// ---- 照抄：dashboard-reservation.html ｜ md5 788126396bd55e3e4643c4a16767abb5 ｜ 原行號 L782–L783、L785–L796 ｜ 訂位：minDateOf、maxDateOf、todayStr、curYm、prevYm（邏輯一字不改）----
function minDateOf(rows, di) { let lo = ''; rows.forEach(r => { const d = r[di]; if (d && (!lo || d < lo)) lo = d; }); return lo; }
function maxDateOf(rows, di) { let hi = ''; rows.forEach(r => { const d = r[di]; if (d && d > hi) hi = d; }); return hi; }
function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}
function curYm() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}
function prevYm() {
  const n = new Date(); const d = new Date(n.getFullYear(), n.getMonth() - 1, 1);
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}

// ---- 照抄：dashboard-reservation.html ｜ md5 788126396bd55e3e4643c4a16767abb5 ｜ 原行號 L815–L819 ｜ 訂位：buildIdx（邏輯一字不改）----
function buildIdx(cols) {
  const idx = {};
  cols.forEach((c, i) => idx[c] = i);
  return idx;
}

// ---- 照抄：dashboard-reservation.html ｜ md5 788126396bd55e3e4643c4a16767abb5 ｜ 原行號 L934–L944 ｜ 訂位：isTestStore、dropTest（邏輯一字不改）----
function isTestStore(nm) { return /測試|測試用|^test/i.test(String(nm || '')); }
let EXCLUDED_TEST = 0;
function dropTest(cols, rows) {
  cols = cols || []; rows = rows || [];
  const idx = buildIdx(cols);
  let dropped = 0;
  const keep = (idx.store === undefined) ? rows
    : rows.filter(r => { if (isTestStore(r[idx.store])) { dropped++; return false; } return true; });
  return { cols: cols, rows: keep, idx: idx, dropped: dropped };
}


// ---- 照抄：dashboard-reservation.html ｜ md5 788126396bd55e3e4643c4a16767abb5 ｜ 原行號 L1107–L1108 ｜ 訂位：isValid、pct（邏輯一字不改）----
function isValid(r, idx) { return r[idx.status] !== '已取消'; }
function pct(n, d) { return d > 0 ? (n / d * 100) : null; }

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L324、L328、L427–L433 ｜ 採購：TODAY、num、gvDate、ymd、toBool（邏輯一字不改）----
var FWD_DAYS=14, TODAY=(typeof window!=="undefined"&&window.__TODAY__)?new Date(window.__TODAY__):new Date();
function num(x){return (x===null||x===undefined||x===""||isNaN(+x))?0:+x;}
function gvDate(v){if(v==null||v==="")return null;if(v instanceof Date)return isNaN(v)?null:v;
  if(typeof v==="string"){var m=v.match(/^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+),(\d+))?/);
    if(m)return new Date(+m[1],+m[2],+m[3],+(m[4]||0),+(m[5]||0),+(m[6]||0));
    var d=new Date(v.replace(/\//g,"-"));return isNaN(d)?null:d;}
  if(typeof v==="number"){var d2=new Date(v);return isNaN(d2)?null:d2;}return null;}
function ymd(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}
function toBool(x){return x===true||x==="TRUE"||x==="true"||x===1||x==="1";}

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L521、L527、L625、L715、L955–L957 ｜ 採購：isDur、pOf、isDC、vendorOf、pqEff、pqCust、pqOf（邏輯一字不改）----
function isDur(s){return s.cat==="器具"||s.cat==="模具";}
function pOf(st,s){return (PARAM[st]||{})[s.id]||{};}
function isDC(st,s){return /出貨中心/.test(String(vendorOf(st,s)||""));}
function vendorOf(st,s){var v=pOf(st,s).custVendor;return v||s.vendor;}
function pqEff(st,s){var c=num(pOf(st,s).custPackQty);return c>0?c:num(s.packQty);}
function pqCust(st,s){return num(pOf(st,s).custPackQty)>0;}
function pqOf(st,s){var q=pqEff(st,s);return q>0?q:1;}

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L1413、L1422、L1434–L1440 ｜ 採購：CAMPHQ_LEAD、CAMPHQ_LEAD_EXCL、buildNewAgg（邏輯一字不改）----
var CAMPHQ_LEAD=15;
var CAMPHQ_LEAD_EXCL=90;
function buildNewAgg(rows){var g={};(rows||[]).forEach(function(r){var st=String(r["店號"]==null?"":r["店號"]).trim(),id=String(r["sku_id"]||"").trim();if(!st||!id)return;
  var part={qty:num(r["需求量"]),ev:String(r["事件名"]||""),excl:(r["專屬"]==="是"||r["專屬"]===true),who:String(r["新品清單"]||"")};
  var s=(g[st]=g[st]||{}),m=s[id];
  if(!m){s[id]={qty:part.qty,ev:part.ev,excl:part.excl,who:part.who,parts:[part]};return;}
  m.parts.push(part);m.qty+=part.qty;if(part.ev&&m.ev.split("、").indexOf(part.ev)<0)m.ev+=(m.ev?"、":"")+part.ev;
  m.excl=m.excl&&part.excl;m.who=m.parts.map(function(p){return p.who;}).filter(Boolean).join("、");});
  return g;}

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L1525、L1537–L1546 ｜ 採購：campExcl、campLead、campHQShipped（邏輯一字不改）----
function campExcl(id){var k=Object.keys(NEWAGG);for(var i=0;i<k.length;i++){var m=NEWAGG[k[i]][id];if(m)return !!m.excl;}return false;}
function campLead(s){return campExcl(s.id)?CAMPHQ_LEAD_EXCL:CAMPHQ_LEAD;}
function campHQShipped(ev,s,rows){   /* 該店該品項在視窗內的大平台下單量（使用單位）與包數；另回報視窗外還有多少 */
  var ws=EVCV[ev]||[],lead=campLead(s);
  var from=ws.length?ymd(new Date(new Date(ws[0].cFrom+"T00:00:00").getTime()-lead*864e5)):"";
  var g=0,packs=0,names={},preG=0,prePacks=0,preFirst="";
  (rows||[]).forEach(function(r){if(!anPlatMatch(r.name,s))return;
    var u=shipUnitG(String(r.st),{unit:s.useUnit,sku:s},r.name);
    if(from&&r.date<from){preG+=num(r.qty)*u.g;prePacks+=num(r.qty);if(!preFirst||r.date<preFirst)preFirst=r.date;return;}
    g+=num(r.qty)*u.g;packs+=num(r.qty);names[r.name]=1;});
  return {g:g,packs:packs,from:from,lead:lead,preG:preG,prePacks:prePacks,preFirst:preFirst,names:Object.keys(names)};}

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L2094、L2109–L2179 ｜ 採購：anNorm、大平台品名配對（anPlatScore／anPlatOwn／anPlatMatch）（邏輯一字不改）----
function anNorm(x){return String(x||"").replace(/[\s（）()\[\]【】\/／\-－_·．.｜|,，、]/g,"").toLowerCase();}
var PLATOWN={},PLAT_MIN=0.6;
var PLAT_SZ=/(\d+(?:\.\d+)?)\s*(吋|cm|CM|公分|號)/g;
function platSizes(t){var o={},m;PLAT_SZ.lastIndex=0;while((m=PLAT_SZ.exec(String(t||""))))o[m[1]+m[2].toLowerCase()]=1;return o;}
function platVariant(n){var t=String(n||""),i=Math.max(t.lastIndexOf("｜"),t.lastIndexOf("|"));return i>=0?t.slice(i+1):"";}
/* 2026-09-03 ag 批：括號裡的字是「規格／適用對象」，不是商品本身。
   實例：「冰箱保鮮盒附蓋 (2.6L/1.8L/0.62L/5L(裝飾鮮奶油專用)) /組」整批下單被掛到「裝飾鮮奶油」名下——
   完整包含比對讀到括號內的「裝飾鮮奶油」還拿 0.9+ 高分，連 ⚠近似 都不會標。牛頭不對馬嘴。
   三個修正：
     ① 主名段 platHead ＝「第一個括號之前」＋「｜之後的變體段（同樣去括號）」。完整包含只在主名段成立才給高分；
        只出現在括號內的降為弱證據 PLAT_WEAK(0.5)，低於門檻 0.6。
        變體段必須算進主名段，否則「仙女棒蠟燭 (愛心/星星)｜星星」的主名段只剩「仙女棒蠟燭」，三個變體會全歸同一筆。
     ② 「〈品名〉專用」是修飾語不是該品項（檸檬專用刀 ≠ 檸檬），同樣降為弱證據。
     ③ 變體把關 platVarKey：主檔品名尾端括號內 1–3 字（布/木/紙、愛心/星星、金/銀、彎曲…）是變體識別字，
        大平台品名有變體段時必須出現同一個字，否則淘汰。與尺寸把關同一個道理。
   ⚠️ 弱證據**不可以 continue**——否則規格編號(codeHit)與雙字組比對永遠跑不到，
      「B08 (塔派專用粉) 10包/組」「4/6吋 抹面板…｜4吋」會因此配丟。要往下跑完再取最高分。 */
function platHead(n){var t=String(n||"");
  var i=Math.max(t.lastIndexOf("｜"),t.lastIndexOf("|"));
  var base=(i>=0?t.slice(0,i):t),vari=(i>=0?t.slice(i+1):"");
  function cut(x){var j=String(x).search(/[（(]/);return j>=0?String(x).slice(0,j):String(x);}
  return cut(base)+" "+cut(vari);}
function platVarKey(nm){var m=String(nm||"").match(/[（(]([^）)]{1,3})[）)]\s*$/);return m?m[1]:"";}
var PLAT_WEAK=0.5;
function anPlatScore(platName,sku){
  if(!sku||!sku.name)return 0;
  var nP=anNorm(platName);if(!nP)return 0;
  var pv=platVariant(platName);
  var sk=Object.keys(platSizes(sku.name));            /* 尺寸把關 */
  if(sk.length){var pSz=platSizes(pv||platName);
    if(Object.keys(pSz).length&&!sk.some(function(k){return pSz[k];}))return 0;}
  var vk=platVarKey(sku.name);                        /* ③ 變體把關（布/木/紙、愛心/星星、金/銀…）*/
  if(vk&&pv&&anNorm(pv).indexOf(anNorm(vk))<0)return 0;
  var code=(String(sku.name).match(/^([A-Za-z]\d{2})(?![0-9A-Za-z])/)||[])[1];
  var codeHit=!!(code&&new RegExp("^"+code+"(?![0-9])","i").test(String(platName).trim()));
  var nH=anNorm(platHead(platName));
  var cands=[sku.name].concat(sku.aliases||[]),best=0;
  for(var i=0;i<cands.length;i++){
    var n=anNorm(cands[i]).replace(/^[a-z]\d{2}/,"");if(n.length<2)continue;
    var at=nH.indexOf(n);
    if(at>=0){
      if(nH.substr(at+n.length,2)==="專用"){          /* ②「〈品名〉專用」是修飾語 → 只算弱證據 */
        if(PLAT_WEAK>best)best=PLAT_WEAK;
        continue;}
      var sc=0.9+Math.min(0.09,n.length/200);         /* ① 完整包含（限主名段）*/
      if(codeHit)sc=Math.min(0.99,sc+0.02);
      if(sc>best)best=sc;
      continue;}
    if(nP.indexOf(n)>=0&&PLAT_WEAK>best)best=PLAT_WEAK;  /* 只出現在括號內 → 弱證據 */
    /* 相似度**也只在主名段上算**：若拿整串比，括號內的「裝飾鮮奶油專用」照樣會讓雙字組全中，前面的把關就白做了 */
    var bg=[],hit=0;for(var k=0;k+1<n.length;k++)bg.push(n.substr(k,2));
    if(!bg.length)continue;
    bg.forEach(function(b){if(nH.indexOf(b)>=0)hit++;});
    var r=hit/bg.length;
    if(r>=PLAT_MIN){var s2=Math.min(0.8,r*0.8);if(s2>best)best=s2;}}
  if(codeHit&&0.85>best)best=0.85;                    /* 規格編號相同（B08…）本身就是強證據，與主名段無關 */
  return best;}
function anPlatOwn(platName){                          /* ① 唯一歸屬（含分數，供 UI 標近似） */
  var key=anNorm(platName);if(!key)return {id:"",sc:0};
  if(Object.prototype.hasOwnProperty.call(PLATOWN,key))return PLATOWN[key];
  var bestId="",bestSc=0,bestLen=0,bestWs=9;
  SKUS.forEach(function(s){
    var sc=anPlatScore(platName,s);if(!(sc>0))return;
    /* 2026-09-03 ai 批：長度以正規化後計（「寒天 QQ」與「寒天QQ」正規化後同長），避免多一個空格就搶走配對 */
    var ln=anNorm(s.name||"").length;
    var ws=/\s/.test(String(s.name||""))?1:0;                       /* 品名內有空白＝多半是重複建檔的錯字版 */
    if(sc>bestSc||(Math.abs(sc-bestSc)<1e-9&&(ln>bestLen||(ln===bestLen&&(ws<bestWs||(ws===bestWs&&s.id<bestId)))))){bestSc=sc;bestId=s.id;bestLen=ln;bestWs=ws;}});
  PLATOWN[key]=(bestSc>=PLAT_MIN?{id:bestId,sc:bestSc}:{id:"",sc:bestSc});
  return PLATOWN[key];}
function anPlatMatch(platName,sku){
  if(!sku||!sku.id)return false;
  return anPlatOwn(platName).id===sku.id;}

// ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L2256–L2284 ｜ 採購：platSpecG、platSpecCount、shipUnitG（邏輯一字不改）----
function platSpecG(name){             /* 從大平台品名解析每包多少 g/ml */
  var t=String(name||""),m,base=null;
  m=t.match(/(\d+(?:\.\d+)?)\s*(?:kg|KG|Kg|公斤)/);            if(m)base=+m[1]*1000;
  if(base==null){m=t.match(/(\d+(?:\.\d+)?)\s*公升/);           if(m)base=+m[1]*1000;}
  if(base==null){m=t.match(/(\d+(?:\.\d+)?)\s*(?:g|G|公克|克)(?![a-zA-Z])/); if(m)base=+m[1];}
  if(base==null){m=t.match(/(\d+(?:\.\d+)?)\s*(?:ml|ML|毫升|cc|CC)/);        if(m)base=+m[1];}
  if(base==null)return null;
  var mm=t.match(/(\d+(?:\.\d+)?)\s*[包入個條支片張]\s*[\/／]\s*[組箱盒]/);   /* 「10包/組」之類的組裝倍數 */
  if(mm&&+mm[1]>1)base=base*(+mm[1]);
  return base>0?base:null;}
/* 2026-09-03 ae 批修正（經營者實測抓到）：舊版只在使用單位是 g/ml 時解析大平台規格，
   計數單位（包/個/張…）一律退回主檔內容量。櫻花翻糖主檔是「包／包／內容量 1」，大平台是「10包/組」，
   於是 1 號店下單 2 組（＝20 包）被算成 2 包，扣掉用量 6 包後帳面剩 −4 包。**這是硬錯。**
   規格寫法的通則是「N<使用單位>／<售賣單位>」：10包/組＝1 組有 10 包、2kg/原包裝＝1 包裝有 2000 g、
   50組/箱＝1 箱有 50 組。所以只要拿「數字＋使用單位＋斜線」去比對就正確，不分重量或計數。 */
function platSpecCount(name,unit){
  var t=String(name||""),u=String(unit||"").trim();if(!u)return null;
  var esc=u.replace(/[.*+?^${}()|[\]\\]/g,"\\$&");
  var m=t.match(new RegExp("(\\d+(?:\\.\\d+)?)\\s*"+esc+"\\s*[\\/／]"));
  if(m&&+m[1]>0)return +m[1];
  if(/^(個|入|張|片|支|條|組|包|捲|卷|把|對|只|顆|件)$/.test(u)){      /* 入/個 是泛用計數詞，互通 */
    var m2=t.match(/(\d+(?:\.\d+)?)\s*(?:入|個)\s*[\/／]/);
    if(m2&&+m2[1]>0)return +m2[1];}
  return null;}
function shipUnitG(st,m,name){   /* 回 {g:每個平台單位含多少使用單位, src:來源, hq:主檔內容量, mis:是否不一致} */
  var uu=String(m.unit||"").trim(),pq=m.sku?pqOf(st,m.sku):0;
  var v=(uu==="g"||uu==="ml")?platSpecG(name):platSpecCount(name,uu);
  if(v>0)return {g:v,src:"plat",hq:pq,mis:(pq>0&&Math.abs(pq-v)>0.001)};
  return {g:(pq>0?pq:1),src:"hq",hq:pq,mis:false};}

// ---- 照抄：zijiren.html ｜ md5 24387ca360e4a05725a0801dac0f73f2 ｜ 原行號 L410、L414–L420、L422、L428–L441、L445–L450 ｜ 自己人：STORE_NAMES、TARGETS_2026、TOTAL_TARGET_2026、ELAPSED、CLOSED_IDX_2026、sumClosed2026、updateElapsedMonths（邏輯一字不改）----
const STORE_NAMES = ["台中精明店","台中草悟道店","台北南京店","台北士林店","台南Focus店","新竹文化店","新北板橋店","新北新店店","桃園中壢店","桃園藝文店","台北遠百信義A13店","高雄SKM Park店"];
const TARGETS_2026 = {
  "台中精明店":4580,"台中草悟道店":7190,"台北南京店":10116,
  "台北士林店":7409,"台南Focus店":6838,"新竹文化店":5434,
  "新北板橋店":7084,"新北新店店":3854,"桃園中壢店":4774,
  "桃園藝文店":3971,"台北遠百信義A13店":6693,"高雄SKM Park店":4584
};
const TOTAL_TARGET_2026 = 72528;
let ELAPSED_MONTHS_2026 = 6; // 預設值，由 updateElapsedMonths() 動態計算
const CLOSED_IDX_2026 = (() => {
  const n = new Date();
  if (n.getFullYear() > 2026) return 11;
  if (n.getFullYear() < 2026) return -1;
  return n.getMonth() - 1;   // 7月 → 5（＝6月）
})();

// 只加總「1月 ～ 已完結月」的新增人數（key 格式 `2026/0` ~ `2026/11`，0-indexed）
function sumClosed2026(monthMap) {
  return Object.entries(monthMap || {}).reduce((sum, [k, v]) => {
    const mo = parseInt(k.split('/')[1], 10);
    return (!isNaN(mo) && mo <= CLOSED_IDX_2026) ? sum + (Number(v) || 0) : sum;
  }, 0);
}
function updateElapsedMonths(lastDataDate) {
  const d = lastDataDate || new Date();
  if (d.getFullYear() < 2026) { ELAPSED_MONTHS_2026 = 0; return; }
  if (d.getFullYear() > 2026) { ELAPSED_MONTHS_2026 = 12; return; }
  ELAPSED_MONTHS_2026 = d.getMonth();   // 7月 → 6（＝已完結 6 個月 → 50%）
}

// ---- 照抄：zijiren.html ｜ md5 24387ca360e4a05725a0801dac0f73f2 ｜ 原行號 L609、L737–L751 ｜ 自己人：parseRows、parseNewMembersSheet（邏輯一字不改）----
function parseRows(res) { return (res && res.table && res.table.rows) ? res.table.rows : []; }
function parseNewMembersSheet(res, year) {
  const result = {};
  parseRows(res).forEach(r => {
    if (!r.c || !r.c[0] || !r.c[0].v) return;
    const storeName = String(r.c[0].v).trim();
    if (storeName === '合計' || storeName === '分店') return;
    if (!STORE_NAMES.includes(storeName)) return;
    result[storeName] = {};
    for (let mo = 1; mo <= 12; mo++) {
      const val = (r.c[mo] && r.c[mo].v != null) ? r.c[mo].v : 0;
      if (val > 0) result[storeName][`${year}/${mo - 1}`] = val; // gvizMonth 0-indexed
    }
  });
  return result;
}


// ============================================================
// 2. 共用工具（本專案自寫）
// ============================================================
var RUN = null;   // 本次執行批次 { ts, date, meta:[], t0, fresh:{} }

function newRun_() {
  var now = new Date();
  return { ts: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), date: Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), meta: [], t0: Date.now(), fresh: {} };
}
function d2s_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function ymOf_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM'); }
function r1_(x) { return (x === null || x === undefined || isNaN(x)) ? '' : Math.round(x * 10) / 10; }
function r0_(x) { return (x === null || x === undefined || isNaN(x)) ? '' : Math.round(x); }
function addDays_(ds, n) { var p = ds.split('-'); var d = new Date(+p[0], +p[1] - 1, +p[2] + n); return d2s_(d); }
function daysInMonth_(ym) { var p = ym.split('-'); return new Date(+p[0], +p[1], 0).getDate(); }
function prevYmOf_(ym) { var p = ym.split('-'); var d = new Date(+p[0], +p[1] - 2, 1); return ymOf_(d); }
function lyOf_(ds) { return (+ds.slice(0, 4) - 1) + ds.slice(4); }   // 去年同日（yyyy-MM-dd 或 yyyy-MM）
function daysBetween_(a, b) { return Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 864e5); }   // b − a（天）
function ntd_(x) { return 'NT$' + Math.round(Number(x) || 0).toLocaleString('en-US'); }
function dstr_(x) { return (x instanceof Date) ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd') : String(x || '').slice(0, 10); }

// ---- d2 D2-2：執行途中的時間檢查 ----
function usedSec_() { return EXEC_T0 ? (Date.now() - EXEC_T0) / 1000 : 0; }
/** 讀資料前呼叫：本段已用超過 RUN_HARD_SEC 秒 → 丟出「延後」訊號（不算失敗，由下一段從頭重跑這一類） */
function timeGuard_(where) {
  if (EXEC_T0 && usedSec_() > RUN_HARD_SEC) throw new Error(DEFER_TAG + '本段已用 ' + Math.round(usedSec_()) + ' 秒（上限 ' + RUN_HARD_SEC + ' 秒），「' + where + '」延到下一段');
}
function isDefer_(e) { return String(e && e.message || e).indexOf(DEFER_TAG) === 0; }
function testCfg_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_TEST) || '{}') || {}; } catch (e) { return {}; } }
function isTestMode_() { var p = PropertiesService.getScriptProperties(); return p.getProperty(PROP_BRIEF_TEST) === '1' || !!p.getProperty(PROP_TEST); }
/** 測試旗標 slow：讀資料時模擬多等 N 秒（每 5 秒檢查一次時間，超過就延後）；正式時沒有旗標，直接略過 */
function testSlow_(cat) {
  var t = testCfg_(), s = t.slow && Number(t.slow[cat]);
  if (!(s > 0)) return;
  var end = Date.now() + s * 1000;
  while (Date.now() < end) { timeGuard_('【測試】模擬慢讀'); Utilities.sleep(Math.max(1, Math.min(5000, end - Date.now()))); }
}

// ---- d3 D3-1：追查記錄（_trace 隱藏分頁）----
// 每次讀取「開始前」「結束後」各寫一列並立刻 flush：就算這次讀取卡死被 Google 中止，也看得到卡在哪一步（只有「開始」沒有「結束」）。
// 只記時間、類別、步驟名、秒數；不存任何回傳內容。寫入失敗只記 log，不影響快照。
var TRACE_SH_ = null;
function traceSheet_() {
  if (TRACE_SH_) return TRACE_SH_;
  var ss = snapSS_(), sh = ss.getSheetByName(TAB.TRACE);
  if (!sh) { sh = ss.insertSheet(TAB.TRACE, ss.getSheets().length); sh.getRange(1, 1, 1, HDR._trace.length).setValues([HDR._trace]); sh.hideSheet(); }
  TRACE_SH_ = sh;
  return sh;
}
function trace_(cat, step, phase, sec) {
  for (var k = 0; k < 2; k++) {
    try {
      var sh = traceSheet_();
      writeTyped_(sh.getRange(sh.getLastRow() + 1, 1, 1, HDR._trace.length),
        [[Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), cat, step, phase, sec === undefined ? '' : Math.round(sec * 10) / 10, (RUN && RUN.ts) || '']]);
      SpreadsheetApp.flush();
      return;
    } catch (e) { TRACE_SH_ = null; if (k) Logger.log('追查記錄寫入失敗（不影響快照）：' + shortErr_(e && e.message || e)); }
  }
}
/** 刪掉 14 天前的追查記錄（runAll 開始時呼叫） */
function tracePrune_() {
  try {
    var sh = traceSheet_(), last = sh.getLastRow();
    if (last < 2) return;
    var cut = addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), -TRACE_KEEP_DAYS);
    var v = sh.getRange(2, 1, last - 1, 1).getValues(), n = 0;
    while (n < v.length && String(v[n][0]).slice(0, 10) < cut) n++;
    if (n > 0) sh.deleteRows(2, n);
  } catch (e) { Logger.log('追查記錄清理失敗：' + shortErr_(e && e.message || e)); }
}
/** 測試旗標 hang：模擬單次讀取卡死 N 秒（中途不做任何時間檢查，模擬真實卡住）；正式時沒有旗標，直接略過 */
function testHang_(cat) {
  var t = testCfg_(), s = t.hang && Number(t.hang[cat]);
  if (!(s > 0)) return;
  var end = Date.now() + s * 1000;
  while (Date.now() < end) Utilities.sleep(Math.max(1, Math.min(290000, end - Date.now())));
}

/** 讀 GAS 端點：剝除 callback( ) 外殼；錯誤或非 JSON → 最多重試 2 次（等 5／10 秒）；每次讀取前、等待前都檢查時間 */
function fetchJson_(cat, label, url) {
  var delays = [5, 10], t0 = Date.now(), lastErr = null;
  for (var i = 0; i <= delays.length; i++) {
    timeGuard_(label);
    testSlow_(cat);
    var tr = (cat === 'A' || cat === 'C' || cat === 'D'), tA = Date.now(), trStep = label + (i ? '（第 ' + (i + 1) + ' 次）' : '');
    if (tr) { trace_(cat, trStep, '開始'); testHang_(cat); }
    try {
      var resp;
      try { resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }); }
      finally { if (tr) trace_(cat, trStep, '結束', (Date.now() - tA) / 1000); }
      var code = resp.getResponseCode();
      if (code !== 200) throw new Error('HTTP ' + code);
      var t = String(resp.getContentText('UTF-8')).trim();
      var m = t.match(/^[A-Za-z_$][\w$.]*\s*\(/);
      if (m) { t = t.slice(m[0].length).replace(/\)\s*;?\s*$/, ''); }
      var obj = JSON.parse(t);
      logMeta_(cat, label, url.replace(/^https:\/\/script\.google\.com\/macros\/s\/(.{10}).*?\/exec/, 'GAS $1…'), (Date.now() - t0) / 1000, '', '', '成功' + (i ? '（第 ' + (i + 1) + ' 次）' : ''), '');
      return obj;
    } catch (e) {
      if (isDefer_(e)) throw e;
      lastErr = e;
      if (i < delays.length) { timeGuard_(label + '（重試前）'); Utilities.sleep(delays[i] * 1000); }
    }
  }
  logMeta_(cat, label, url.replace(/^https:\/\/script\.google\.com\/macros\/s\/(.{10}).*?\/exec/, 'GAS $1…'), (Date.now() - t0) / 1000, '', '', '失敗', String(lastErr && lastErr.message || lastErr));
  throw new Error(label + ' 讀取失敗：' + (lastErr && lastErr.message || lastErr));
}

/** 讀既有 Sheet（唯讀 getValues）：回傳 [{欄名:值}]（第 1 列為表頭） */
function readTable_(cat, ssId, sheetName) {
  timeGuard_(sheetName);
  testSlow_(cat);
  var t0 = Date.now();
  var sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sh) { logMeta_(cat, sheetName, ssId, (Date.now() - t0) / 1000, 0, '', '失敗', '找不到分頁'); throw new Error('找不到分頁 ' + sheetName); }
  var v = sh.getDataRange().getValues();
  var h = (v[0] || []).map(function (x) { return String(x).trim(); });
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var o = {}, empty = true;
    for (var j = 0; j < h.length; j++) { if (!h[j]) continue; var x = v[i][j]; o[h[j]] = (x === '' ? null : x); if (x !== '' && x !== null) empty = false; }
    if (!empty) rows.push(o);
  }
  logMeta_(cat, sheetName, ssId.slice(0, 8) + '…', (Date.now() - t0) / 1000, rows.length, '', '成功', '');
  return rows;
}

function logMeta_(cat, source, item, sec, n, latest, result, err) {
  if (!RUN) RUN = newRun_();
  RUN.meta.push([RUN.ts, cat, source, item, Math.round(sec * 10) / 10, n, latest, result, err]);
}
function setMetaLatest_(cat, source, n, latest) {
  for (var i = RUN.meta.length - 1; i >= 0; i--) {
    if (RUN.meta[i][1] === cat && RUN.meta[i][2] === source) { if (n !== undefined && n !== null) RUN.meta[i][5] = n; if (latest) RUN.meta[i][6] = latest; return; }
  }
}

function snapSS_() {
  var id = PropertiesService.getScriptProperties().getProperty(PROP_SNAP_ID);
  if (!id) throw new Error('尚未建立快照 Sheet，請先執行 setupSnapshotSheet()');
  return SpreadsheetApp.openById(id);
}

/** dim_rule：缺的規則補列、既有規則更新「類別／說明／單位」（數值保留經營者設定），再讀成 {代碼:數值} */
function readRules_(ss) {
  var sh = ss.getSheetByName(TAB.RULE), v = sh.getDataRange().getValues(), r = {}, pos = {};
  for (var i = 1; i < v.length; i++) { var k = String(v[i][0]).trim(); if (k) { r[k] = RULE_TEXT[k] ? String(v[i][3]).trim() : Number(v[i][3]); pos[k] = i + 1; } }
  var add = [];
  RULE_DEFAULTS.forEach(function (d) {
    if (pos[d[0]]) {
      var row = v[pos[d[0]] - 1];
      if (String(row[1]) !== d[1] || String(row[2]) !== d[2] || String(row[4]) !== d[4]) { sh.getRange(pos[d[0]], 2, 1, 2).setValues([[d[1], d[2]]]); sh.getRange(pos[d[0]], 5).setValue(d[4]); }
    } else add.push(d);
    if (!(d[0] in r) || (RULE_TEXT[d[0]] ? r[d[0]] === '' : isNaN(r[d[0]])) || v[pos[d[0]] - 1] && v[pos[d[0]] - 1][3] === '') r[d[0]] = d[3];
  });
  if (add.length) writeTyped_(sh.getRange(sh.getLastRow() + 1, 1, add.length, 5), add);
  return r;
}

/** 店名對照：byCol[來源欄][店名] = 店號 */
function readStoreMap_(ss) {
  var v = ss.getSheetByName(TAB.MAP).getDataRange().getValues(), h = v[0].map(String), m = { list: [], byCol: {}, name: {}, zone: {} };
  h.forEach(function (c) { m.byCol[c] = {}; });
  for (var i = 1; i < v.length; i++) {
    var id = String(v[i][0]).trim(); if (!id) continue;
    m.list.push(id); m.name[id] = String(v[i][1]); m.zone[id] = String(v[i][2]);
    for (var j = 1; j < h.length; j++) { var nm = String(v[i][j] || '').trim(); if (nm) m.byCol[h[j]][nm] = id; }
  }
  return m;
}

// ---- 結果物件 ----
function newRes_(cat, catName) { return { cat: cat, catName: catName, recs: [], alerts: [] }; }
/** rec：store(店號或'全公司')、code、period、value、cmp、cmpType、latest、light、wide（snap_latest 欄名；空白＝不上寬表） */
function rec_(res, store, code, period, value, cmp, cmpType, latest, light, wide) {
  res.recs.push({ store: String(store), code: code, period: period || '', value: (value === undefined ? '' : value), cmp: (cmp === undefined ? '' : cmp), cmpType: cmpType || '', latest: latest || '', light: light || '', wide: wide || '' });
}
function alert_(res, light, store, desc, value, thr) { res.alerts.push([light, String(store), res.cat + ' ' + res.catName, desc, value === undefined ? '' : value, thr === undefined ? '' : thr]); }

var LIGHT_ORDER = { '🔴': 0, '🟠': 1, '🟡': 2, '🟢': 3, '': 4 };
function storeSort_(s) { var n = Number(s); return isNaN(n) ? 99 : n; }
function catOfCode_(code) { return String(code || '').charAt(0); }   // 'A1_hours' → 'A'
function catOfAlert_(row) { return String(row[2] || '').charAt(0); }  // 'A 人力營收' → 'A'；'0 資料' → '0'

/** 寫入時鎖住型別：文字格子先設「純文字」格式（避免 '2026-09-23'、'2026-09' 被 Sheets 自動轉成日期），數字格子維持自動格式 */
function writeTyped_(range, values) {
  var f = range.getNumberFormats();
  for (var i = 0; i < values.length; i++) for (var j = 0; j < values[i].length; j++) { if (typeof values[i][j] === 'string') f[i][j] = '@'; }
  range.setNumberFormats(f);
  range.setValues(values);
}

// ============================================================
// 3. 寫入快照 Sheet（全部類別算完才寫：snap_latest、alerts、snap_history、meta）
// ============================================================
/**
 * finalize_：把本次各類結果寫進快照 Sheet。
 *   results：{A:{ok,err,recs,alerts,meta,fresh}, …}（本次有跑的類別）
 *   isAll：true＝runAll（另產生「0 資料」新鮮度警示、meta 整張換新、必要時寄信）
 * 失敗的類別：snap_latest 與 alerts 保留上一次的數字，另加 🔴 失敗警示。
 */
function finalize_(results, isAll) {
  var ss = snapSS_(), rule = readRules_(ss), map = readStoreMap_(ss), tW = Date.now();
  // d2：carry＝今天稍早已算完、這次補跑沒重算的類別（snap_latest／alerts／history 保留，只借用它的資料新鮮度）
  var ran = Object.keys(results).filter(function (c) { return !results[c].carry; }), okCats = ran.filter(function (c) { return results[c].ok; });
  var failCats = ran.filter(function (c) { return !results[c].ok; });

  // --- snap_latest：成功的類別換新欄，其餘類別沿用舊欄；每次重建在全新分頁，放第 1 個 ---
  var shL = ss.getSheetByName(TAB.LATEST);
  var old = shL ? shL.getDataRange().getValues() : [['店號', '標準店名', '區']], oh = old[0] || [];
  var oldByKey = {}; for (var i = 1; i < old.length; i++) oldByKey[String(old[i][0])] = old[i];
  var rowKeys = map.list.concat([ALL_ROW]);
  var hdr = ['店號', '標準店名', '區'], getters = [];
  CATS.forEach(function (c) {
    var cat = c[0];
    if (okCats.indexOf(cat) >= 0) {
      var recs = results[cat].recs, cols = [], cell = {};
      recs.forEach(function (r) { if (r.wide) { if (cols.indexOf(r.wide) < 0) cols.push(r.wide); cell[r.store + '\u0001' + r.wide] = r; } });
      cols.forEach(function (w) {
        hdr.push(w); getters.push(function (k) { var r = cell[k + '\u0001' + w]; return r ? r.value : ''; });
        hdr.push(w + '｜資料最新日'); getters.push(function (k) { var r = cell[k + '\u0001' + w]; return r ? r.latest : ''; });
      });
    } else {
      oh.forEach(function (h, j) { h = String(h); if (j >= 3 && h.indexOf(cat + ' ') === 0) { hdr.push(h); getters.push(function (k) { var o = oldByKey[k]; return o ? o[j] : ''; }); } });
    }
  });
  var out = [hdr];
  rowKeys.forEach(function (k) {
    var row = [k === ALL_ROW ? ALL_ROW : Number(k), k === ALL_ROW ? '12 店合計' : map.name[k], k === ALL_ROW ? '' : map.zone[k]];
    getters.forEach(function (g) { var x = g(k); row.push(x instanceof Date ? dstr_(x) : x); });
    out.push(row);
  });
  var tmpName = TAB.LATEST + '_tmp', oldTmp = ss.getSheetByName(tmpName);
  if (oldTmp) ss.deleteSheet(oldTmp);
  var shN = ss.insertSheet(tmpName, 0);   // M5：永遠放第 1 個分頁
  writeTyped_(shN.getRange(1, 1, out.length, hdr.length), out);
  shN.setFrozenRows(1); shN.setFrozenColumns(2);
  if (shL) ss.deleteSheet(shL);
  shN.setName(TAB.LATEST);
  ss.setActiveSheet(shN); ss.moveActiveSheet(1);

  // --- 「0 資料」新鮮度警示（只在 runAll 產生） ---
  var alerts0 = isAll ? freshAlerts_(results, rule, map) : null;

  // --- alerts：成功類別換新；失敗類別保留舊列＋🔴；0 資料（runAll 才換） ---
  var shA = ss.getSheetByName(TAB.ALERTS), av = shA.getDataRange().getValues(), all = [];
  for (var a = 1; a < av.length; a++) {
    if (String(av[a][0]) === '') continue;
    var ac = catOfAlert_(av[a]);
    if (okCats.indexOf(ac) >= 0) continue;                 // 成功類別：換新
    if (ac === '0' && alerts0) continue;                    // 0 資料：runAll 換新
    if (failCats.indexOf(ac) >= 0 && String(av[a][3]).indexOf('本次快照失敗') >= 0) continue;   // 舊的失敗警示不重複
    all.push(av[a]);
  }
  okCats.forEach(function (c) { all = all.concat(results[c].alerts); });
  if (!isAll) failCats.forEach(function (c) {   // runAll 的失敗改列在「0 資料」，不重複
    var nm = CATS.filter(function (x) { return x[0] === c; })[0][1];
    all.push(['🔴', ALL_ROW, c + ' ' + nm, nm + ' 本次快照失敗，這一類的數字沿用上一次的結果。原因：' + shortErr_(results[c].err), '', '']);
  });
  if (alerts0) all = all.concat(alerts0);
  all.sort(function (x, y) { return (LIGHT_ORDER[x[0]] - LIGHT_ORDER[y[0]]) || String(x[2]).localeCompare(String(y[2])) || (storeSort_(x[1]) - storeSort_(y[1])); });
  shA.clear();
  shA.getRange(1, 1, 1, HDR.alerts.length).setValues([HDR.alerts]);
  if (all.length) writeTyped_(shA.getRange(2, 1, all.length, HDR.alerts.length), all);

  // --- snap_history：同日覆蓋（先刪今天、本次成功類別的舊列）＋ 保留最近 400 天 ---
  var hist = [];
  okCats.forEach(function (c) { results[c].recs.forEach(function (r) { hist.push([RUN.date, r.store, r.code, r.period, r.value, r.cmp, r.cmpType, r.latest, r.light, RUN.ts]); }); });
  writeHistory_(ss, okCats, hist);

  // --- meta：runAll 整張換新；單類只換該類 ---
  var metaRows = [];
  ran.forEach(function (c) { metaRows = metaRows.concat(results[c].meta || []); });
  if (isAll) {
    var segs = RUN.segments || 1, carried = Object.keys(results).filter(function (c) { return results[c].carry; });
    metaRows.push([RUN.ts, 'ALL', '寫入快照（snap_latest／alerts／snap_history）', '', Math.round((Date.now() - tW) / 100) / 10, '', '', '成功', '']);
    metaRows.push([RUN.ts, 'ALL', '【runAll 總計】', CATS.map(function (c) { return c[0] + (results[c[0]] && results[c[0]].carry ? '＝' : (results[c[0]] && results[c[0]].ok ? '✓' : '✗')); }).join(' '),
      Math.round((Date.now() - RUN.t0) / 100) / 10, '', '分 ' + segs + ' 段執行（每段預算 ' + rule.RUN_BUDGET_SEC + ' 秒）' + (carried.length ? '；只補跑沒完成的類別，' + carried.join('') + ' 沿用今天稍早的結果（＝）' : ''), failCats.length ? '部分失敗' : '成功', '']);
  }
  var shM = ss.getSheetByName(TAB.META), mv = shM.getDataRange().getValues(), keepM = [];
  if (!isAll) for (var m = 1; m < mv.length; m++) { if (String(mv[m][1]) && ran.indexOf(String(mv[m][1])) < 0) keepM.push(mv[m]); }
  var allM = keepM.concat(metaRows);
  shM.clear();
  shM.getRange(1, 1, 1, HDR.meta.length).setValues([HDR.meta]);
  if (allM.length) writeTyped_(shM.getRange(2, 1, allM.length, HDR.meta.length), allM);

  // --- d2：記錄各類算完時間（早報／儀表板判斷「今天有沒有算完」）；runAll 另記整份快照完成時間與資料新鮮度 ---
  var props = PropertiesService.getScriptProperties(), nowTs = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  var cd = jsonProp_(PROP_CAT_DONE); okCats.forEach(function (c) { cd[c] = nowTs; }); props.setProperty(PROP_CAT_DONE, JSON.stringify(cd));
  if (isAll) {
    var fr = jsonProp_(PROP_FRESH); okCats.forEach(function (c) { fr[c] = results[c].fresh || {}; }); props.setProperty(PROP_FRESH, JSON.stringify(fr));
    var okToday = CATS.map(function (c) { return c[0]; }).filter(function (c) { return String(cd[c] || '').slice(0, 10) === nowTs.slice(0, 10); });
    props.setProperty(PROP_DONE, JSON.stringify({ date: nowTs.slice(0, 10), at: nowTs.slice(0, 16), ok: okToday,
      fail: failCats.map(function (c) { return { cat: c, err: shortErr_(results[c].err) }; }) }));
  }

  // --- M9 判斷材料交給 runLoop_：早報寄出就不再寄 M9；早報沒寄出時，任一類失敗或 0 資料 🔴 → 寄 M9 信 ---
  var reds0 = (alerts0 || []).filter(function (x) { return x[0] === '🔴'; });
  return { ok: okCats, fail: failCats, reds0: reds0 };
}

function jsonProp_(k) { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(k) || '{}') || {}; } catch (e) { return {}; } }
function shortErr_(e) {
  var s = String(e || '').split('\n')[0];
  return s.length > 160 ? s.slice(0, 160) + '…' : s;
}

/** M6：同日覆蓋＋保留 400 天。今天的列一定在最後面（依時間追加），只讀 A、C 兩欄找範圍。 */
function writeHistory_(ss, cats, rows) {
  var sh = ss.getSheetByName(TAB.HIST), last = sh.getLastRow(), today = RUN.date;
  if (last >= 2) {
    var colA = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return dstr_(r[0]); });
    var start = -1;
    for (var i = colA.length - 1; i >= 0; i--) { if (colA[i] === today) start = i; else if (colA[i] < today) break; }
    if (start >= 0) {
      var n = colA.length - start, blk = sh.getRange(start + 2, 1, n, HDR.snap_history.length).getValues();
      var keep = blk.filter(function (r) { return cats.indexOf(catOfCode_(r[2])) < 0; }).map(function (r) { return r.map(function (x) { return x instanceof Date ? dstr_(x) : x; }); });
      sh.getRange(start + 2, 1, n, HDR.snap_history.length).clear();
      if (keep.length) writeHistTyped_(sh.getRange(start + 2, 1, keep.length, HDR.snap_history.length), keep);
    }
    // 保留最近 HIST_KEEP_DAYS 天：最舊的在最上面
    var cutoff = addDays_(today, -HIST_KEEP_DAYS), k = 0;
    while (k < colA.length && colA[k] && colA[k] < cutoff) k++;
    if (k > 0) {
      var lastNow = sh.getLastRow();
      if (k < lastNow - 1) sh.deleteRows(2, k); else sh.getRange(2, 1, k, HDR.snap_history.length).clear();
    }
  }
  if (rows.length) writeHistTyped_(sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR.snap_history.length), rows);
}

/** e4：snap_history 專用寫入。文字格子設「純文字」（同 writeTyped_）；非文字（數字）格子一律設 General，
 *  不沿用格子原本的格式——2026-09-26 查到「數值」欄有殘留的日期格式，數字寫進去會被讀成 1900-01-xx 的日期。 */
function writeHistTyped_(range, values) {
  var f = values.map(function (r) { return r.map(function (x) { return typeof x === 'string' ? '@' : 'General'; }); });
  range.setNumberFormats(f);
  range.setValues(values);
}

/** M4：「0 資料」新鮮度警示（門檻全在 dim_rule） */
function freshAlerts_(results, rule, map) {
  var out = [], today = RUN.date, L = '0 資料';
  var f = function (c) { return (results[c] && results[c].fresh) || {}; };
  var push = function (light, desc, val, thr) { out.push([light, ALL_ROW, L, desc, val, thr]); };
  CATS.forEach(function (c) {
    if (results[c[0]] && !results[c[0]].ok && !results[c[0]].carry) push('🔴', c[1] + '（' + c[0] + ' 類）這次讀取失敗，這一類的快照數字沒有更新。原因：' + shortErr_(results[c[0]].err), '失敗', '任一類讀取失敗');
  });
  var a = f('A');
  if (a.lastSales && daysBetween_(a.lastSales, today) > rule.FRESH_SALES_LAG_DAYS) push('🟡', '排班分析的每日營收只到 ' + a.lastSales + '（今天 ' + today + '），比預期晚；人力營收的數字少了最近幾天', a.lastSales, '早於今天 ' + rule.FRESH_SALES_LAG_DAYS + ' 天以上');
  if (a.posLatest && daysBetween_(a.posLatest, today) > rule.FRESH_SALES_LAG_DAYS) push('🟡', 'POS 銷售資料只到 ' + a.posLatest + '（今天 ' + today + '），POS 同步可能延遲', a.posLatest, '早於今天 ' + rule.FRESH_SALES_LAG_DAYS + ' 天以上');
  var b = f('B');
  if (b.shopLatest && daysBetween_(b.shopLatest, today) > rule.FRESH_SHOP_LAG_DAYS) push('🟡', '大平台下單資料（fact_shopline）只到 ' + b.shopLatest + '，已落後 ' + daysBetween_(b.shopLatest, today) + ' 天；檔期「已出貨」會偏少，請上傳最新的大平台報表', b.shopLatest, '落後超過 ' + rule.FRESH_SHOP_LAG_DAYS + ' 天');
  var c2 = f('C');
  if (c2.futMin && daysBetween_(c2.futMin, today) > rule.FRESH_FUT_LAG_DAYS) push('🟡', '未來訂位停更：未來訂位表最早日期是 ' + c2.futMin + '（應該是今天 ' + today + '），訂位頁的每日重建可能沒跑', c2.futMin, '早於今天');
  if (c2.gbState === '未設定') push('🟡', '團體訂位通關碼還沒設定：未選甜點／未付足訂金／需立即處理三欄寫「未設定」。請經營者到本專案「指令碼屬性」新增 GB_PASSCODE', '未設定', '');
  if (c2.gbState === '通關碼錯誤') push('🟡', '團體訂位通關碼不正確：團體三欄寫「通關碼錯誤」。請經營者到本專案「指令碼屬性」更新 GB_PASSCODE', '通關碼錯誤', '');
  if (c2.gbState === '讀取失敗') push('🔴', '團體訂位資料讀取失敗（團體 GAS 沒有回應）：團體三欄寫「讀取失敗」。原因：' + shortErr_(c2.gbErr), '讀取失敗', '任一來源讀取失敗');
  var d = f('D');
  if (d.latest && daysBetween_(d.latest, today) > rule.FRESH_REV_LAG_DAYS) push('🟡', 'Google 評論最新一則是 ' + d.latest + '，已經 ' + daysBetween_(d.latest, today) + ' 天沒有新評論進來，評論同步（Make）可能停了', d.latest, '早於今天 ' + rule.FRESH_REV_LAG_DAYS + ' 天以上');
  var e = f('E');
  if (e.stale) push('🟡', '自己人年表可能停更：全公司本月新增 ' + e.cur + ' 人，連續 ' + rule.MEM_STALE_DAYS + ' 天沒變（' + (e.prevVals || []).join('、') + '）', e.cur, '連續 ' + rule.MEM_STALE_DAYS + ' 天不變');
  return out;
}

/** M9：異常通知信（白話；不含個資、不含通關碼） */
function notifyMail_(ss, results, failCats, reds0) {
  var lines = ['經營快照 ' + RUN.ts + ' 的每日自動執行有異常：', ''];
  failCats.forEach(function (c) {
    var nm = CATS.filter(function (x) { return x[0] === c; })[0][1];
    lines.push('・「' + nm + '」這一類沒有算完，快照裡這一類的數字維持上一次的結果。原因：' + shortErr_(results[c].err));
  });
  reds0.forEach(function (x) { if (x[3].indexOf('這次讀取失敗') < 0) lines.push('・' + x[3]); });
  lines.push('', '詳細請看「DIYBC 經營快照」的 alerts 分頁（0 資料）與 meta 分頁：', ss.getUrl(), '', '這封信只有出問題時才會寄；正常時不寄。需要處理時請交給 Cowork。');
  try { MailApp.sendEmail({ to: MAIL_TO, subject: (isTestMode_() ? '【測試】' : '') + '【經營快照】異常', body: lines.join('\n') }); RUN.mailed = true; logMeta_('ALL', '異常通知信', MAIL_TO, 0, '', '', '已寄出', ''); }
  catch (e) { logMeta_('ALL', '異常通知信', MAIL_TO, 0, '', '', '寄信失敗', shortErr_(e && e.message || e)); }
}

// ---- M8：分段執行的暫存（_stage 分頁，隱藏；每類一筆 JSON，超過 45,000 字切段）----
function stageSheet_(ss) {
  var sh = ss.getSheetByName(TAB.STAGE);
  if (!sh) { sh = ss.insertSheet(TAB.STAGE, ss.getSheets().length); sh.hideSheet(); }
  return sh;
}
function stageClear_(ss) { var sh = ss.getSheetByName(TAB.STAGE); if (sh) sh.clear(); }
function stagePut_(ss, runTs, cat, obj) {
  var sh = stageSheet_(ss), s = JSON.stringify(obj), rows = [];
  for (var i = 0; i * 45000 < s.length; i++) rows.push([runTs, cat, i, s.slice(i * 45000, (i + 1) * 45000)]);
  if (!rows.length) rows.push([runTs, cat, 0, '{}']);
  writeTyped_(sh.getRange(sh.getLastRow() + 1, 1, rows.length, 4), rows);
}
function stageGetAll_(ss, runTs) {
  var sh = ss.getSheetByName(TAB.STAGE), out = {};
  if (!sh || sh.getLastRow() < 1) return out;
  var v = sh.getRange(1, 1, sh.getLastRow(), 4).getValues(), parts = {};
  v.forEach(function (r) { if (String(r[0]) !== runTs) return; (parts[r[1]] = parts[r[1]] || [])[Number(r[2])] = String(r[3]); });
  Object.keys(parts).forEach(function (c) { out[c] = JSON.parse(parts[c].join('')); });
  return out;
}

/** 跑一類：成功回 {ok, recs, alerts, meta, fresh}；失敗回 {ok:false, err, meta, fresh} */
function runOneCat_(cat) {
  var nm = CATS.filter(function (x) { return x[0] === cat; })[0][1];
  var fn = { A: calcA_, B: calcB_, C: calcC_, D: calcD_, E: calcE_ }[cat];
  var m0 = RUN.meta.length, t0 = Date.now(), res = newRes_(cat, nm), ok = true, err = '', deferred = false;
  RUN.fresh[cat] = {};
  try {
    var tc = testCfg_(); if (tc.fail && tc.fail.indexOf(cat) >= 0) throw new Error('【測試】模擬「' + nm + '」讀取失敗');
    var ss = snapSS_(); fn(res, readRules_(ss), readStoreMap_(ss));
  }
  catch (e) { ok = false; err = String(e && e.message || e); deferred = isDefer_(e); }
  var sec = (Date.now() - t0) / 1000;
  logMeta_(cat, '【' + cat + ' 類小計】', nm, sec, res.recs.length, '', ok ? '成功' : (deferred ? '延後' : '失敗'), ok ? '' : err);
  var r = { ok: ok, err: err, deferred: deferred, sec: Math.round(sec * 10) / 10, meta: RUN.meta.slice(m0), fresh: RUN.fresh[cat] };
  if (ok) { r.recs = res.recs; r.alerts = res.alerts; }
  return r;
}

function clearContinueTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'runAllContinue') ScriptApp.deleteTrigger(t); });
}

/**
 * runAll 的主迴圈（d2）：
 *   D2-1 每類開始前：本段已用秒數＋該類預估秒數（EST_x）> RUN_BUDGET_SEC → 不開始，交給 1 分鐘後的接續段
 *        （本段一類都還沒跑時照樣開始，避免預估值太大永遠不跑）
 *   D2-2 類別途中讀資料前超過 RUN_HARD_SEC → 延後：本類不存結果，下一段從頭重跑；同一類連續延後 DEFER_MAX 次 → 改標失敗
 *        接續段超過 RUN_MAX_SEGMENTS → 剩下的類別標失敗，寫入快照並寄信
 *   寫入快照前：已用＋RUN_FINAL_SEC > RUN_HARD_SEC → 寫入也交給下一段
 *   D3-2 保險絲：每段一開始先預約 5 分鐘後的 runAllContinue；本段若被 Google 中止（單次讀取卡死），保險絲會接手。
 *        接手時看到上一段「開始了某類但沒做完」（state.running）→ 算卡住 1 次（與延後合計），從該類重跑；
 *        合計 2 次 → 該類標失敗「讀取卡住超過 6 分鐘，已連續 2 次」，其餘類別照跑、照寫快照、照寄早報。保險絲接手也算 1 段（上限 8 段）。
 *   state：{ ts, date, t0, done:[], segments, defer:{A:次數}, late(巡檢員補跑), carry(只補跑部分類別), overCap,
 *            live(本段進行中；正常交棒／結束才清), running:{cat,at}(該類開始、還沒做完), stuck:{A:次數}, stuckMeta:{A:[meta 列]} }
 */
function runLoop_(state) {
  var props = PropertiesService.getScriptProperties();
  var save = function () { props.setProperty(PROP_RUN_STATE, JSON.stringify(state)); };
  // D3-2 保險絲：任何可能卡住的讀取之前，先預約下一段
  clearContinueTriggers_();
  ScriptApp.newTrigger('runAllContinue').timeBased().after(FUSE_MIN * 60 * 1000).create();
  state.defer = state.defer || {}; state.stuck = state.stuck || {}; state.stuckMeta = state.stuckMeta || {};
  var died = !!state.live;   // 上一段沒有正常交棒／結束就不見了 → 被 Google 中止，這一段是保險絲接手
  if (died) {
    state.segments++;
    if (state.segments > RUN_MAX_SEGMENTS + 1) {   // 連寫入都一直失敗：停止，不再無限接續
      props.deleteProperty(PROP_RUN_STATE); clearContinueTriggers_();
      throw new Error('分段接續已超過 ' + RUN_MAX_SEGMENTS + ' 段仍未完成（保險絲），這次停止；巡檢員會再檢查');
    }
  }
  state.live = 1;
  var ss = snapSS_(), rule = readRules_(ss), execStart = Date.now();
  EXEC_T0 = execStart;
  RUN = { ts: state.ts, date: state.date, meta: [], t0: state.t0, fresh: {}, segments: state.segments };
  var ranHere = 0;
  var handoff = function (why) {
    if (state.segments >= RUN_MAX_SEGMENTS) return false;
    state.segments++;
    state.live = 0;
    save();
    clearContinueTriggers_();
    ScriptApp.newTrigger('runAllContinue').timeBased().after(60 * 1000).create();
    Logger.log(why + '；第 ' + state.segments + ' 段 1 分鐘後接著跑');
    return true;
  };
  var failRest = function (why) {   // 接續段用完：剩下的類別標失敗
    state.overCap = why;
    CATS.forEach(function (c) {
      if (state.done.indexOf(c[0]) >= 0) return;
      stagePut_(ss, state.ts, c[0], { ok: false, err: why, meta: [[RUN.ts, c[0], '【' + c[0] + ' 類小計】', c[1], 0, '', '', '失敗', why]], fresh: {} });
      state.done.push(c[0]);
    });
    delete state.running;
    save();
  };
  if (died && state.running && state.done.indexOf(state.running.cat) < 0) {   // D3-2：上一段卡在這一類
    var sc = state.running.cat, snm = CATS.filter(function (x) { return x[0] === sc; })[0][1];
    var waited = Math.round((Date.now() - Number(state.running.t || Date.now())) / 1000);
    state.stuck[sc] = (state.stuck[sc] || 0) + 1;
    state.defer[sc] = (state.defer[sc] || 0) + 1;
    var sm = [RUN.ts, sc, '保險絲接手', snm, 0, '', '', '卡住', '「' + snm + '」' + state.running.at + ' 開始後沒有做完（被 Google 中止，第 ' + state.stuck[sc] + ' 次卡住），' + waited + ' 秒後由保險絲接手'];
    (state.stuckMeta[sc] = state.stuckMeta[sc] || []).push(sm);
    delete state.running;
    if (state.defer[sc] >= DEFER_MAX) {
      var sWhy = '「' + snm + '」' + STUCK_MSG;
      stagePut_(ss, state.ts, sc, { ok: false, err: sWhy, meta: state.stuckMeta[sc].concat([[RUN.ts, sc, '【' + sc + ' 類小計】', snm, 0, '', '', '失敗', sWhy]]), fresh: {} });
      state.done.push(sc);
    }
    save();
    Logger.log('保險絲接手：' + sc + ' 類卡住第 ' + state.stuck[sc] + ' 次');
  } else if (died) { delete state.running; save(); }
  else save();
  if (state.segments > RUN_MAX_SEGMENTS && CATS.some(function (c) { return state.done.indexOf(c[0]) < 0; })) failRest('接續段已達 ' + RUN_MAX_SEGMENTS + ' 段上限，還沒算的類別這次放棄');
  for (var i = 0; i < CATS.length; i++) {
    var cat = CATS[i][0], nm = CATS[i][1];
    if (state.done.indexOf(cat) >= 0) continue;
    var est = Number(rule['EST_' + cat]) || 60, used = (Date.now() - execStart) / 1000;
    if (ranHere > 0 && used + est > rule.RUN_BUDGET_SEC) {
      if (handoff('已用 ' + Math.round(used) + ' 秒＋' + cat + ' 類預估 ' + est + ' 秒 > 預算 ' + rule.RUN_BUDGET_SEC + ' 秒')) return false;
      failRest('接續段已達 ' + RUN_MAX_SEGMENTS + ' 段上限，還沒算的類別這次放棄');
      break;
    }
    state.running = { cat: cat, at: Utilities.formatDate(new Date(), TZ, 'HH:mm:ss'), t: Date.now() }; save();   // D3-2「X 類開始時間」
    var r = runOneCat_(cat); ranHere++;
    delete state.running;
    if (state.stuckMeta[cat]) r.meta = state.stuckMeta[cat].concat(r.meta);
    if (r.ok) { var er = estRecord_(ss, cat, r.sec); if (er) r.meta.push(er); }
    if (r.deferred) {
      state.defer[cat] = (state.defer[cat] || 0) + 1;
      if (state.defer[cat] < DEFER_MAX) {
        if (handoff(cat + ' 類讀資料時時間不夠（第 ' + state.defer[cat] + ' 次延後）')) return false;
        failRest('接續段已達 ' + RUN_MAX_SEGMENTS + ' 段上限，還沒算的類別這次放棄');
        break;
      }
      var why = state.stuck[cat] ? '「' + nm + '」' + STUCK_MSG + '（其中延後 ' + (state.defer[cat] - state.stuck[cat]) + ' 次）。最後一次：' + shortErr_(r.err)
                                 : '「' + nm + '」連續 ' + DEFER_MAX + ' 次時間不夠（延後），改標失敗。最後一次：' + shortErr_(r.err);
      r.meta.forEach(function (m) { if (m[2] === '【' + cat + ' 類小計】') { m[7] = '失敗'; m[8] = why; } });
      r = { ok: false, err: why, meta: r.meta, fresh: r.fresh };
    }
    stagePut_(ss, state.ts, cat, r);
    state.done.push(cat);
    save();
  }
  // 寫入快照前再看一次時間：不夠就把寫入交給下一段（寫入＋寄早報約 20～40 秒）
  if (ranHere > 0 && (Date.now() - execStart) / 1000 + RUN_FINAL_SEC > RUN_HARD_SEC) {
    if (handoff('類別都算完了，但本段剩下的時間不夠寫入快照')) return false;
  }
  if (state.running) { delete state.running; save(); }
  var results = stageGetAll_(ss, state.ts);
  if (state.carry) {   // 巡檢員只補跑沒完成的類別：其餘類別標 carry（沿用今天稍早的結果）
    var fr = jsonProp_(PROP_FRESH);
    CATS.forEach(function (c) { if (!results[c[0]]) results[c[0]] = { ok: true, carry: true, fresh: fr[c[0]] || {}, meta: [] }; });
  }
  RUN.segments = state.segments;
  var fin = finalize_(results, true);
  // finalize_ 換掉了 snap_latest 分頁；本函式開頭拿的 ss 物件還記得舊分頁，會報「Sheet … not found」→ 重新開一次再清暫存。
  // 清暫存只是收尾，失敗也不影響今天的快照，所以只記 log、不寄信。
  try { stageClear_(snapSS_()); } catch (e) { Logger.log('清 _stage 暫存失敗（不影響快照，下次 runAll 開始時會再清）：' + shortErr_(e && e.message || e)); }
  // D5：當天 alerts 存進 alerts_history（早報判斷 🆕／第 N 天用）
  try { var ssH = snapSS_(); writeAlertsHistory_(ssH, readStoreMap_(ssH), state.date); } catch (e) { Logger.log('alerts_history 寫入失敗：' + shortErr_(e && e.message || e)); }
  bundleCacheClear_();   // e1：快照與 alerts_history 都寫完 → 清掉儀表板讀取窗口的快取
  // D1：全部類別完成後寄早報；同一次已寄早報就不寄 M9 異常信，早報沒寄出才照舊寄 M9
  var briefSent = false;
  try { briefSent = sendBrief_(state.late ? 'late' : 'auto').sent; } catch (e) { Logger.log('早報寄送失敗：' + shortErr_(e && e.message || e)); }
  if (state.overCap || (!briefSent && (fin.fail.length || fin.reds0.length))) notifyMail_(snapSS_(), results, fin.fail, fin.reds0);
  props.deleteProperty(PROP_RUN_STATE);
  clearContinueTriggers_();
  bundlePrebuild_(execStart);   // _bundle 交辦書 2026-09-26：早報寄出後預先準備首頁資料（整段 try/catch，失敗只寫 meta；runAll 與巡檢員補跑都走這裡）
  return true;
}

/** D2-1：記錄這一類的實際秒數，預估秒數自動更新為「近 7 次最大值 × 1.3」（寫回 dim_rule；回傳一列 meta）；測試旗標開著時不更新 */
function estRecord_(ss, cat, sec) {
  if (testCfg_().noEst || testCfg_().slow) return null;
  try {
    var props = PropertiesService.getScriptProperties(), h = jsonProp_(PROP_EST_HIST);
    var arr = (h[cat] || []).concat([Number(sec) || 0]).slice(-EST_KEEP); h[cat] = arr;
    props.setProperty(PROP_EST_HIST, JSON.stringify(h));
    var est = Math.max(5, Math.ceil(Math.max.apply(null, arr) * 1.3));
    var sh = ss.getSheetByName(TAB.RULE), v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === 'EST_' + cat) { sh.getRange(i + 1, 4).setValue(est); break; }
    return [RUN.ts, cat, '預估秒數自動更新', 'EST_' + cat, 0, arr.length, '', '→ ' + est + ' 秒', '近 ' + arr.length + ' 次：' + arr.join('、') + ' 秒；最大 × 1.3'];
  } catch (e) { Logger.log('預估秒數更新失敗：' + shortErr_(e && e.message || e)); return null; }
}

// ============================================================
// 3c. 每日經營早報（階段 d：D1～D6）
//   只讀 alerts／snap_latest／alerts_history（＋snap_history 取自己人年度進度線），排版後寄出。
//   不改任何算法與門檻；數字一律照抄分頁上的值，不自行推算（合計取 snap_latest 全公司列）。
// ============================================================
var AHIST_KEEP_DAYS = 90;                                   // D5 alerts_history 保留天數
var PROP_BRIEF_SENT = 'BRIEF_SENT_DATE';                    // D1 今天是否已寄（yyyy-MM-dd）
var PROP_BRIEF_TEST = 'BRIEF_TEST';                         // '1'＝測試模式：主旨加【測試】、不記 BRIEF_SENT_DATE
var BRIEF_CAT_ORDER = ['C', 'A', 'D', 'E', 'B', '0'];       // D4 排序：團體 → 營收人力 → 評論 → 自己人 → 檔期 → 0 資料
var BRIEF_LINE_MAX = 60;                                    // D4 每行字數上限
var BRIEF_ASK = '想看原因或建議，到 Claude「決策分身」專案問我';
var SHORT_NAME_DEFAULTS = { 1: '精明', 2: '草悟道', 3: '南京', 4: '士林', 5: 'Focus', 6: '新竹', 7: '板橋', 8: '新店', 9: '中壢', 10: '藝文', 11: 'A13', 12: 'SKM' };
var WEEKDAY_ZH = ['日', '一', '二', '三', '四', '五', '六'];

function blen_(s) { return Array.from(String(s)).length; }
function bnum_(s) { var m = String(s).replace(/,/g, '').match(/-?\d+(\.\d+)?/); return m ? Number(m[0]) : NaN; }
function bfmt_(x) { var n = Number(x); return (x === '' || x === null || isNaN(n)) ? String(x) : (Math.abs(n) >= 1000 ? Math.round(n).toLocaleString('en-US') : String(n)); }
function bsign_(g) { var n = Number(g); return isNaN(n) ? String(g) : (n > 0 ? '+' : (n < 0 ? '−' : '')) + Math.abs(n) + '%'; }
function bescape_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/**
 * 狀況種類（D4 合併、D5 新／持續的依據）：同一種狀況＝同一個 id，不看數值。
 * 回傳 { id, title(全店合併標題函式), show(每店顯示字), sort(越大越嚴重), totalCol(snap_latest 全公司欄), single(不合併、整行文字) }
 */
function briefKindOf_(r, map, rule) {
  var cat = String(r[2] || '').charAt(0), d = String(r[3] || ''), v = r[4] instanceof Date ? dstr_(r[4]) : String(r[4] === undefined ? '' : r[4]), sid = String(r[1] || '');
  var nm = map.name[sid] || '', rest = nm && d.indexOf(nm) === 0 ? d.slice(nm.length).replace(/^\s+/, '') : d;
  var m;
  if (cat === 'C') {
    if (/團體訂位有 \d+ 項需要立即處理/.test(d)) return { id: 'C_gb_urgent', title: function (t) { return '團體需立即處理' + (t !== '' ? ' ' + bfmt_(t) + ' 件' : ''); }, show: bfmt_(bnum_(v)), sort: bnum_(v), totalCol: 'C 團體需立即處理' };
    if (/筆團體訂位還沒選甜點/.test(d)) return { id: 'C_gb_nodessert', title: function (t) { return '團體還沒選甜點' + (t !== '' ? ' ' + bfmt_(t) + ' 筆' : ''); }, show: bfmt_(bnum_(v)), sort: bnum_(v), totalCol: 'C 團體未選甜點' };
    if (/本月取消率/.test(d)) return { id: 'C_cancel_up', title: function () { return '本月取消率比上月高'; }, show: v, sort: bnum_(v) };
    if (/對照表沒有的店名/.test(d)) return { id: 'C_unknown_store', single: '團體資料有店名對不上對照表，請補 dim_store_map' };
  }
  if (cat === 'A') {
    if ((m = d.match(/本月（[^）]*）淨營收比去年同期 ([+\-]?[\d.]+)%/))) return { id: 'A_yoy_cur', title: function () { return '本月營收比去年同期'; }, show: bsign_(m[1]), sort: -Number(m[1]) };
    if ((m = d.match(/上月（[^）]*）淨營收比去年同月 ([+\-]?[\d.]+)%/))) return { id: 'A_yoy_prev', title: function () { return '上月營收比去年同月'; }, show: bsign_(m[1]), sort: -Number(m[1]) };
    if ((m = d.match(/上週（.*）紅旗：(.+)$/))) return { id: 'A_flag', title: function () { return '上週排班紅旗'; }, show: m[1], sort: -bnum_(v.replace(/^rev\/h\s*/, '')) };
    if (/本月班表未上傳/.test(d)) return { id: 'A_nosched', title: function () { return '本月班表未上傳，人效算不出來'; }, show: '', sort: bnum_(v) };
  }
  if (cat === 'D') {
    if ((m = d.match(/月底前 (\d+) 天，本月有效評論只有 (\d+) 則（目標 (\d+)）/))) { var mm = m, lastMin = (rule && rule.REV_LAST_MIN !== undefined && rule.REV_LAST_MIN !== '') ? rule.REV_LAST_MIN : ruleDefault_('REV_LAST_MIN'); return { id: 'D_last', title: function () { return '月底前 ' + mm[1] + ' 天有效評論不到 ' + lastMin + ' 則（目標 ' + mm[3] + '）'; }, show: mm[2], sort: -Number(mm[2]) }; }
    if ((m = d.match(/落後進度 ([\d.]+) 則/))) return { id: 'D_behind', title: function () { return '有效評論落後本月進度'; }, show: '落後 ' + m[1], sort: Number(m[1]) };
  }
  if (cat === 'E') {
    if ((m = d.match(/自己人累計達成率 ([\d.]+)%.*共 ([\d.]+) 個百分點/))) return { id: 'E_behind', title: function () { return '自己人累計達成率低於進度線'; }, show: m[1] + '%', sort: Number(m[2]) };
  }
  if (cat === 'B') {
    if ((m = rest.match(/^(.+?)：\d+ 項檔期專屬料已出貨超過推估需求/))) { var ev = m[1]; return { id: 'B_over_' + ev, title: function (t) { return ev.replace(/節$/, '') + '專屬料多訂' + (t !== '' ? ' NT$' + bfmt_(t) : ''); }, show: bfmt_(bnum_(v)), sort: bnum_(v), totalCol: 'B ' + ev + ' 超出金額合計' }; }
    if (/沒盤點|從未在採購系統盤點/.test(d)) return { id: 'B_inv_stale', title: function () { return '太久沒盤點（' + String(r[5] || '').replace(/^>/, '超過 ') + '）'; }, show: v === '從未' ? '從未盤點' : v, sort: v === '從未' ? 9999 : bnum_(v) };
  }
  if (cat === '0') {
    if ((m = d.match(/^(.+?)（([A-E]) 類）這次讀取失敗/))) return { id: '0_fail_' + m[2], system: '系統：' + m[1] + '（' + m[2] + ' 類）今天沒有算完，數字是上一次的' };
    if (/排班分析的每日營收只到/.test(d)) return { id: '0_sales', single: '排班每日營收只到 ' + v + '，人力營收少了最近幾天' };
    if (/POS 銷售資料只到/.test(d)) return { id: '0_pos', single: 'POS 銷售資料只到 ' + v + '，POS 同步可能延遲' };
    if ((m = d.match(/大平台下單資料.*只到 (\S+)，已落後 (\d+) 天/))) return { id: '0_shop', single: '大平台下單資料只到 ' + m[1] + '，落後 ' + m[2] + ' 天，請上傳最新報表' };
    if (/未來訂位停更/.test(d)) return { id: '0_fut', single: '未來訂位停更：最早日期是 ' + v + '，訂位頁每日重建可能沒跑' };
    if (/通關碼還沒設定/.test(d)) return { id: '0_gb_unset', single: '團體通關碼還沒設定，團體三欄沒有數字' };
    if (/通關碼不正確/.test(d)) return { id: '0_gb_wrong', single: '團體通關碼不正確，團體三欄沒有數字' };
    if (/團體訂位資料讀取失敗/.test(d)) return { id: '0_gb_fail', single: '團體訂位資料讀取失敗，團體三欄沒有數字' };
    if ((m = d.match(/Google 評論最新一則是 (\S+)，已經 (\d+) 天/))) return { id: '0_rev', single: 'Google 評論最新一則是 ' + m[1] + '，已 ' + m[2] + ' 天沒新評論' };
    if (/年表可能停更/.test(d)) return { id: '0_mem', single: '自己人年表可能停更：本月新增 ' + v + ' 人連續幾天沒變' };
  }
  return { id: 'X_' + cat + '_' + d.replace(/[\d.,%NT$]+/g, '#').slice(0, 16), single: d };
}
/** dim_rule 預設值（briefKindOf_ 在沒有傳 rule 時使用） */
function ruleDefault_(code) { for (var i = 0; i < RULE_DEFAULTS.length; i++) if (RULE_DEFAULTS[i][0] === code) return RULE_DEFAULTS[i][3]; return ''; }
function briefKey_(r, kind) { return String(r[2] || '').charAt(0) + '|' + String(r[1] || '') + '|' + kind.id; }

/** 讀 dim_store_map 短名；沒有「短名」欄時自動補欄（預設值，經營者可改） */
function briefShortNames_(ss, map) {
  var sh = ss.getSheetByName(TAB.MAP), v = sh.getDataRange().getValues(), h = v[0].map(String), j = h.indexOf('短名'), out = {};
  if (j < 0) {
    j = h.length;
    var col = [['短名']];
    for (var i = 1; i < v.length; i++) col.push([SHORT_NAME_DEFAULTS[Number(v[i][0])] || String(v[i][1] || '')]);
    sh.getRange(1, j + 1, col.length, 1).setValues(col);
    v = sh.getDataRange().getValues();
  }
  for (var k = 1; k < v.length; k++) { var id = String(v[k][0]).trim(); if (id) out[id] = String(v[k][j] || '').trim() || map.name[id] || id; }
  return out;
}

/** D5：當天 alerts 追加到隱藏分頁 alerts_history（同日覆蓋、保留 90 天） */
function writeAlertsHistory_(ss, map, date) {
  var sh = ss.getSheetByName(TAB.AHIST);
  if (!sh) { sh = ss.insertSheet(TAB.AHIST, ss.getSheets().length); sh.hideSheet(); }
  var old = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, HDR.alerts_history.length).getValues() : [];
  var cutoff = addDays_(date, -AHIST_KEEP_DAYS);
  var keep = old.filter(function (r) { var d = dstr_(r[0]); return d && d !== date && d >= cutoff; }).map(function (r) { return [dstr_(r[0])].concat(r.slice(1)); });
  var av = ss.getSheetByName(TAB.ALERTS).getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]) !== ''; });
  var add = av.map(function (r) { return [date, r[0], r[1], r[2], r[3], r[4], r[5], briefKey_(r, briefKindOf_(r, map))]; });
  var all = keep.concat(add);
  sh.clear();
  sh.getRange(1, 1, 1, HDR.alerts_history.length).setValues([HDR.alerts_history]);
  if (all.length) writeTyped_(sh.getRange(2, 1, all.length, HDR.alerts_history.length), all);
  return add.length;
}

/** 組早報：回傳 { subject, text, html, lines:[{sec, text, src}], counts } */
function buildBrief_(ss, rule, map, today) {
  var short = briefShortNames_(ss, map);
  var sn = function (sid) { return short[sid] || map.name[sid] || sid; };
  var av = ss.getSheetByName(TAB.ALERTS).getDataRange().getValues().slice(1).filter(function (r) { return String(r[0]) !== ''; });
  var lv = ss.getSheetByName(TAB.LATEST).getDataRange().getValues(), lh = lv[0].map(String);
  var allRow = lv.filter(function (r) { return String(r[0]) === ALL_ROW; })[0] || [];
  var L = function (col, row) { var j = lh.indexOf(col), x = j < 0 ? '' : (row || allRow)[j]; return x === undefined || x === null ? '' : (x instanceof Date ? dstr_(x) : x); };
  var counts = { '🔴': 0, '🟠': 0, '🟡': 0 };
  av.forEach(function (r) { if (r[0] in counts) counts[r[0]]++; });
  // --- D2-6：快照完成時間；快照不是今天的 → 🔴 舊資料提醒；今天某一類沒算完 → 該類每行標（上次資料） ---
  var snap = snapDoneInfo_(ss), oldSnap = snap.date !== today, stale = {};
  if (!oldSnap) CATS.forEach(function (c) { if (String(snap.catAt[c[0]] || '').slice(0, 10) !== today) stale[c[0]] = true; });

  // --- D5：昨天以前的 alerts_history ---
  var hsh = ss.getSheetByName(TAB.AHIST), hv = hsh && hsh.getLastRow() >= 2 ? hsh.getRange(2, 1, hsh.getLastRow() - 1, HDR.alerts_history.length).getValues() : [];
  var byDay = {};
  hv.forEach(function (r) {
    var d = dstr_(r[0]); if (!d || d >= today) return;
    var k = String(r[7] || ''), p = k.split('|');
    var o = byDay[d] || (byDay[d] = { keys: {}, lines: {} });
    o.keys[k] = 1; o.lines[p[0] + '|' + p[2]] = 1;
  });
  var yday = addDays_(today, -1), marksOn = !!byDay[yday];
  var streak = function (field, key) { var n = 1, d = yday; while (byDay[d] && byDay[d][field][key]) { n++; d = addDays_(d, -1); } return n; };

  // --- D4：依 燈號＋類別＋狀況種類 合併 ---
  var groups = {}, order = [], sysLines = [];
  av.forEach(function (r) {
    var kind = briefKindOf_(r, map, rule), cat = String(r[2] || '').charAt(0);
    if (kind.system) { sysLines.push({ text: '🔴 ' + kind.system, src: r }); return; }
    var gk = r[0] + '|' + cat + '|' + kind.id;
    if (!groups[gk]) { groups[gk] = { light: r[0], cat: cat, kind: kind, rows: [] }; order.push(gk); }
    groups[gk].rows.push({ r: r, kind: kind });
  });
  var lineOf = function (g) {
    var lineKey = g.cat + '|' + g.kind.id, tail = '';
    var lineNew = marksOn && streak('lines', lineKey) === 1;
    if (marksOn) { var n = streak('lines', lineKey); tail = n === 1 ? ' 🆕' : ' 第 ' + n + ' 天'; }
    if (stale[g.cat]) tail += '（上次資料）';
    if (g.kind.single) {
      var head = g.light + ' ' + g.kind.single;
      return fitLine_(head, [], tail);
    }
    var total = g.kind.totalCol ? L(g.kind.totalCol) : '';
    var items = g.rows.slice().sort(function (a, b) { return (Number(b.kind.sort) || 0) - (Number(a.kind.sort) || 0) || (storeSort_(a.r[1]) - storeSort_(b.r[1])); })
      .map(function (x) {
        var isNewStore = marksOn && !lineNew && streak('keys', briefKey_(x.r, x.kind)) === 1;
        return (sn(String(x.r[1])) + (x.kind.show !== '' ? ' ' + x.kind.show : '')) + (isNewStore ? '🆕' : '');
      });
    return fitLine_(g.light + ' ' + g.kind.title(total), items, tail);
  };
  var sortGroups = function (light) {
    return order.filter(function (k) { return groups[k].light === light; }).sort(function (a, b) {
      return BRIEF_CAT_ORDER.indexOf(groups[a].cat) - BRIEF_CAT_ORDER.indexOf(groups[b].cat) || order.indexOf(a) - order.indexOf(b);
    });
  };
  var section = function (light, max) {
    var gs = sortGroups(light), out = [];
    if (!gs.length) return [{ text: '（無）', src: null }];
    var show = gs.length > max ? max - 1 : gs.length;
    gs.slice(0, show).forEach(function (k) { out.push({ text: lineOf(groups[k]), src: groups[k] }); });
    if (gs.length > show) {
      var rest = gs.slice(show).reduce(function (s, k) { return s + groups[k].rows.length; }, 0);
      out.push({ text: '另有 ' + rest + ' 則，見快照 alerts 分頁', src: null });
    }
    return out;
  };

  // --- ④ 全公司一覽（固定 5 行，全部照抄 snap_latest 全公司列；進度線取 snap_history 當天 E2_rate 的比較值） ---
  var dataTo = L('A 本月到營收最後一天淨營收｜資料最新日') || L('A 本月 rev/h｜資料最新日');
  var noH = String(L('A 無工時資料店') || '').split(/[、,]/).filter(String).map(function (s) { return sn(s.trim()); });
  var dCol = lh.indexOf('D 有效評論（本月）'), dOk = 0;
  lv.slice(1).forEach(function (r) { if (String(r[0]) !== ALL_ROW && dCol >= 0 && r[dCol] !== '' && Number(r[dCol]) >= MONTHLY_REVIEW_TARGET) dOk++; });
  var memLine = briefMemLine_(ss, today);
  var ovTag = function (c) { return stale[c] ? '（上次資料）' : ''; };
  var ov = [
    '本月營收 NT$' + bfmt_(L('A 本月到營收最後一天淨營收')) + '，比去年同期 ' + bsign_(L('A 同期成長%')) + ovTag('A'),
    '本月人效 rev/h ' + bfmt_(L('A 本月 rev/h')) + (noH.length ? '（不含 ' + noH.join('、') + '：沒有班表）' : '') + ovTag('A'),
    '有效評論：達標 ' + dOk + '／12 店，全公司本月 ' + bfmt_(L('D 有效評論（本月）')) + ' 則' + ovTag('D'),
    '自己人：本月新增 ' + bfmt_(L('E 本月至今新增')) + ' 人，累計達成率 ' + L('E 累計達成率%') + '%' + (memLine !== '' ? '（年度進度線 ' + memLine + '%）' : '') + ovTag('E'),
    '團體：有效 ' + bfmt_(L('C 團體有效筆數')) + ' 筆，需立即處理 ' + bfmt_(L('C 團體需立即處理')) + ' 件' + ovTag('C')
  ];

  var red = section('🔴', Number(rule.BRIEF_MAX_RED) || 6), org = section('🟠', Number(rule.BRIEF_MAX_ORANGE) || 5), yel = section('🟡', Number(rule.BRIEF_MAX_YELLOW) || 4);
  var dp = today.split('-'), d = new Date(+dp[0], +dp[1] - 1, +dp[2]);
  var subject = '【DIYBC 經營早報】' + (d.getMonth() + 1) + '/' + d.getDate() + '（' + WEEKDAY_ZH[d.getDay()] + '）｜🔴' + counts['🔴'] + ' 🟠' + counts['🟠'] + ' 🟡' + counts['🟡'] + (oldSnap ? '（舊資料）' : '');
  var url = ss.getUrl();
  var head0 = [{ text: '快照完成時間：' + mmdd_(snap.at) + (snap.at ? ' ' + String(snap.at).slice(11, 16) : '') + '｜資料到 ' + mmdd_(dataTo), src: null }];
  if (oldSnap) head0.unshift({ text: '🔴 ' + oldSnapMsg_(snap), src: null });
  var secs = [
    { h: null, lines: head0.concat(sysLines) },
    { h: '① 🔴 今天要處理', lines: red },
    { h: '② 🟠 本週留意', lines: org },
    { h: '③ 🟡 資料提醒', lines: yel },
    { h: '④ 📊 全公司一覽', lines: ov.map(function (t) { return { text: t, src: 'snap_latest' }; }) },
    { h: null, end: true, lines: [{ text: BRIEF_ASK, src: null }, { text: url, src: null, url: url }] }
  ];
  var text = [], html = ['<div style="font-family:-apple-system,BlinkMacSystemFont,\'PingFang TC\',\'Noto Sans TC\',sans-serif;font-size:16px;line-height:1.65;color:#222;max-width:560px">'];
  var flat = [];
  secs.forEach(function (s) {
    if (s.h) { text.push('', s.h); html.push('<div style="font-size:17px;font-weight:700;margin:18px 0 6px">' + bescape_(s.h) + '</div>'); }
    else if (s.end) { text.push(''); html.push('<div style="height:14px"></div>'); }
    s.lines.forEach(function (l) {
      var t = briefScrub_(l.text);
      flat.push({ sec: s.h || '', text: t, src: l.src });
      text.push(t);
      html.push(l.url ? '<div style="font-size:15px;margin:2px 0"><a href="' + bescape_(t) + '">打開 DIYBC 經營快照</a></div>'
                      : '<div style="font-size:16px;margin:3px 0">' + bescape_(t) + '</div>');
    });
  });
  html.push('</div>');
  return { subject: briefScrub_(subject), text: text.join('\n').replace(/^\n/, ''), html: html.join(''), lines: flat, counts: counts, snap: snap, oldSnap: oldSnap, stale: Object.keys(stale) };
}

/** 'yyyy-MM-dd…' → 'MM/DD'（空白回「不明」） */
function mmdd_(s) { s = String(s || ''); return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(5, 7) + '/' + s.slice(8, 10) : '不明'; }
/** D2-6 舊資料提醒句（早報與儀表板共用同一句） */
function oldSnapMsg_(snap) { return '這份早報用的是 ' + mmdd_(snap.at) + ' 的快照，今天的快照還沒完成'; }
/**
 * 快照完成資訊：{ at:'yyyy-MM-dd HH:mm', date:'yyyy-MM-dd', catAt:{A:'yyyy-MM-dd HH:mm:ss'…}, src }
 * 先讀狀態值 SNAP_DONE／SNAP_CAT_DONE；沒有時（d2 上線前的快照）改讀 meta 分頁的【runAll 總計】（開始時間＋耗時）與各類小計。
 */
function snapDoneInfo_(ss) {
  var done = jsonProp_(PROP_DONE), cd = jsonProp_(PROP_CAT_DONE);
  if (done.at) return { at: done.at, date: done.date, catAt: cd, src: 'SNAP_DONE' };
  var out = { at: '', date: '', catAt: {}, src: 'meta' };
  try {
    var mv = ss.getSheetByName(TAB.META).getDataRange().getValues();
    for (var i = 1; i < mv.length; i++) {
      var ts = String(mv[i][0] instanceof Date ? Utilities.formatDate(mv[i][0], TZ, 'yyyy-MM-dd HH:mm:ss') : mv[i][0]);
      if (mv[i][2] === '【runAll 總計】') {
        var p = ts.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})/);
        if (p) { var f = new Date(+p[1], +p[2] - 1, +p[3], +p[4], +p[5], +p[6]); f = new Date(f.getTime() + (Number(mv[i][4]) || 0) * 1000); out.at = Utilities.formatDate(f, TZ, 'yyyy-MM-dd HH:mm'); out.date = out.at.slice(0, 10); }
      }
      var m = String(mv[i][2]).match(/^【([A-E]) 類小計】$/);
      if (m && String(mv[i][7]) === '成功') out.catAt[m[1]] = ts;
    }
  } catch (e) {}
  return out;
}

/** 自己人年度進度線：snap_history 當天（或最近一天）全公司 E2_rate 的「比較值」 */
function briefMemLine_(ss, today) {
  var sh = ss.getSheetByName(TAB.HIST), last = sh.getLastRow(); if (last < 2) return '';
  var from = Math.max(2, last - 3000), v = sh.getRange(from, 1, last - from + 1, 6).getValues();
  for (var i = v.length - 1; i >= 0; i--) if (String(v[i][1]) === ALL_ROW && String(v[i][2]) === 'E2_rate' && dstr_(v[i][0]) <= today) return v[i][5];
  return '';
}

/** 每行 ≤ 60 字：放不下時列前 5 家，後面寫「等 N 家」；仍放不下就再少列幾家，最後才截斷 */
function fitLine_(head, items, tail) {
  var n = items.length;
  var build = function (k) { return n ? head + '：' + items.slice(0, k).join('、') + (k < n ? ' 等 ' + (n - k) + ' 家' : '') + tail : head + tail; };
  var s = build(n);
  if (blen_(s) <= BRIEF_LINE_MAX) return s;
  for (var k = Math.min(n - 1, 5); k >= 1; k--) { s = build(k); if (blen_(s) <= BRIEF_LINE_MAX) return s; }
  return Array.from(n ? build(1) : s).slice(0, BRIEF_LINE_MAX - 1).join('') + '…';
}

/** 紅線 3：早報不可出現電話、Email、通關碼（防呆遮罩；正常情況 alerts 本來就沒有這些） */
function briefScrub_(s) {
  var out = String(s).replace(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g, '［已遮蔽］')
    .replace(/(?:\+?886[- ]?|0)9\d{2}[- ]?\d{3}[- ]?\d{3}/g, '［已遮蔽］')
    .replace(/\b0\d{1,2}-\d{6,8}\b/g, '［已遮蔽］');
  var key = PropertiesService.getScriptProperties().getProperty(PROP_GB_PASS);
  if (key && key.length >= 4) out = out.split(key).join('［已遮蔽］');
  return out;
}
function briefTo_(rule) {
  var to = String(rule.BRIEF_TO || '').trim();
  return (/^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$/.test(to)) ? to : MAIL_TO;   // 只准 1 個收件人；格式不對就退回 mydiybc
}

/** 今天的日期（測試旗標 fakeToday 可假裝成別天，只供 T5 舊資料測試） */
function briefToday_() { return testCfg_().fakeToday || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
/** 今天是否已寄早報（測試模式看測試用的紀錄） */
function briefSentToday_(today) {
  var p = PropertiesService.getScriptProperties();
  return p.getProperty(p.getProperty(PROP_BRIEF_TEST) === '1' ? PROP_BRIEF_SENT_TEST : PROP_BRIEF_SENT) === today;
}
/** 寄早報。mode：'auto'（runAll 完成後，受 BRIEF_ON 與一天一封限制）｜'late'（巡檢員補跑／補寄，主旨加「（延遲）」）｜'force'（sendBriefNow 強制重寄） */
function sendBrief_(mode) {
  var props = PropertiesService.getScriptProperties();
  var ss = snapSS_(), rule = readRules_(ss), map = readStoreMap_(ss);
  var test = props.getProperty(PROP_BRIEF_TEST) === '1';
  var today = briefToday_();
  if (mode === 'auto' || mode === 'late') {
    if (Number(rule.BRIEF_ON) !== 1) { Logger.log('早報暫停中（dim_rule BRIEF_ON≠1），這次不寄'); return { sent: false, why: 'BRIEF_ON' }; }
    if (!test && props.getProperty(PROP_BRIEF_SENT) === today) { Logger.log('今天已寄過早報，這次不再寄'); return { sent: false, why: 'sent' }; }
  }
  var b = buildBrief_(ss, rule, map, today);
  var subj = (test ? '【測試】' : '') + b.subject + (mode === 'force' && !test ? '（重寄）' : '') + (mode === 'late' ? '（延遲）' : '');
  MailApp.sendEmail({ to: briefTo_(rule), subject: subj, body: b.text, htmlBody: b.html, name: 'DIYBC 經營快照' });
  props.setProperty(test ? PROP_BRIEF_SENT_TEST : PROP_BRIEF_SENT, today);
  if (RUN) RUN.mailed = true;
  Logger.log('已寄出早報：' + subj);
  return { sent: true, subject: subj, brief: b };
}
/** 手動強制重寄今天的早報（主旨加「（重寄）」；測試模式時改加【測試】） */
function sendBriefNow() { sendBrief_('force'); }
/** 只產生不寄：把早報純文字印在執行記錄，方便檢查 */
function previewBrief() {
  var ss = snapSS_(), rule = readRules_(ss), map = readStoreMap_(ss);
  var b = buildBrief_(ss, rule, map, briefToday_());
  Logger.log(b.subject + '\n\n' + b.text);
}
/** 測試模式開／關：開啟時主旨加【測試】、不記 BRIEF_SENT_DATE（同一天可重複測）；關閉時一併清空 BRIEF_SENT_DATE */
function briefTestOn() { PropertiesService.getScriptProperties().setProperty(PROP_BRIEF_TEST, '1'); Logger.log('早報測試模式：開'); }
function briefTestOff() { var p = PropertiesService.getScriptProperties(); p.deleteProperty(PROP_BRIEF_TEST); p.deleteProperty(PROP_BRIEF_SENT); p.deleteProperty(PROP_BRIEF_SENT_TEST); Logger.log('早報測試模式：關（BRIEF_SENT_DATE 已清空）'); }

// ============================================================
// 3b. 入口
// ============================================================
/** 每日觸發器呼叫：A→E 依序算，全部完成才寫 snap_latest／alerts／history／meta */
function runAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { Logger.log('另一個 runAll 正在執行，這次略過'); return; }
  RUN = null;
  try {
    var now = new Date();
    var state = { ts: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), date: Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), t0: Date.now(), done: [], segments: 1 };
    var ss = snapSS_();
    clearContinueTriggers_();
    stageClear_(ss);
    PropertiesService.getScriptProperties().setProperty(PROP_RUN_STATE, JSON.stringify(state));
    tracePrune_();
    runLoop_(state);
  } catch (e) { crashMail_('每日自動執行', e); throw e; }
  finally { lock.releaseLock(); }
}
/** 一次性觸發器呼叫（M8）：接著跑剩下的類別；跑完自刪 */
function runAllContinue() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {   // d2：別的執行（例如巡檢員）正占用 → 1 分鐘後再試，接續不遺失
    clearContinueTriggers_();
    if (PropertiesService.getScriptProperties().getProperty(PROP_RUN_STATE)) ScriptApp.newTrigger('runAllContinue').timeBased().after(60 * 1000).create();
    return;
  }
  RUN = null;
  try {
    clearContinueTriggers_();
    var s = PropertiesService.getScriptProperties().getProperty(PROP_RUN_STATE);
    if (!s) { Logger.log('沒有待接續的 runAll'); return; }
    runLoop_(JSON.parse(s));
  } catch (e) { crashMail_('分段接續執行', e); throw e; }
  finally { lock.releaseLock(); }
}

/** M9 補強：runAll 中途停止（例如找不到快照 Sheet）時寄信；已寄過異常信就不重複。不含個資、不含通關碼。 */
function crashMail_(where, e) {
  if (RUN && RUN.mailed) return;
  var msg = shortErr_(e && e.message || e);
  var hint = msg.indexOf('尚未建立快照 Sheet') >= 0
    ? '「指令碼屬性」裡的 SNAP_SS_ID（快照 Sheet 的 ID）不見了：請到 Apps Script 專案設定 → 指令碼屬性補回（不要執行 setupSnapshotSheet，那會另外開一個新的空白 Sheet）。'
    : '需要處理時請交給 Cowork。';
  var body = ['經營快照 ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm') + ' 的' + where + '中途停止，這次快照沒有更新（snap_latest、alerts 維持上一次的結果）。', '',
              '原因：' + msg, '', hint, '', '這封信只有出問題時才會寄；正常時不寄。'].join('\n');
  try { MailApp.sendEmail({ to: MAIL_TO, subject: (isTestMode_() ? '【測試】' : '') + '【經營快照】異常', body: body }); if (RUN) RUN.mailed = true; }
  catch (x) { Logger.log('異常通知信寄送失敗：' + shortErr_(x && x.message || x)); }
}

/**
 * 每日觸發器（2026-09-25 經營者／Chat 核准）：runAll 每天 11:00（Asia/Taipei）。
 * Google 定時誤差 ±15 分鐘，所以設 11:00：實際在 10:45～11:15 之間觸發，
 * 確保在訂位頁「未來訂位」10:15 重建之後至少 30 分鐘，避免誤報「未來訂位停更」。
 * 可重跑：先刪掉本專案既有的 runAll 觸發器，永遠只留 1 個；不碰其他函式的觸發器。
 */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { var h = t.getHandlerFunction(); if (h === 'runAll' || h === 'dailyWatchdog') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('runAll').timeBased().everyDays(1).atHour(11).nearMinute(0).inTimezone(TZ).create();
  ScriptApp.newTrigger('dailyWatchdog').timeBased().everyDays(1).atHour(12).nearMinute(15).inTimezone(TZ).create();
  Logger.log('已建立每日觸發器：runAll 每天 11:00（實際 10:45～11:15）＋ 巡檢員 dailyWatchdog 每天 12:15（實際約 12:00～12:30）（' + TZ + '）');
}

// ============================================================
// 3d. 巡檢員（d2 D2-3）：主執行被 Google 強制中止時無法寄信 → 由第二個每日觸發器檢查、補跑、補寄、通報
// ============================================================
/** 今天的完成狀態：各類今天是否算完、快照是否今天寫完、早報是否今天寄出、是否有接續觸發器在等 */
function todayStatus_(today) {
  var cd = jsonProp_(PROP_CAT_DONE), done = jsonProp_(PROP_DONE), props = PropertiesService.getScriptProperties();
  var catsDone = CATS.map(function (c) { return c[0]; }).filter(function (c) { return String(cd[c] || '').slice(0, 10) === today; });
  var missing = CATS.map(function (c) { return c[0]; }).filter(function (c) { return catsDone.indexOf(c) < 0; });
  var pending = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'runAllContinue'; });
  var rs = null; try { rs = JSON.parse(props.getProperty(PROP_RUN_STATE) || 'null'); } catch (e) {}
  return { catsDone: catsDone, missing: missing, latestToday: done.date === today, complete: !missing.length && done.date === today,
           briefSent: briefSentToday_(today), pending: pending, runState: rs, done: done };
}
function catNames_(cats) { return cats.map(function (c) { return CATS.filter(function (x) { return x[0] === c; })[0][1] + '（' + c + '）'; }).join('、'); }
function metaAppend_(what, result, note) {
  try {
    var sh = snapSS_().getSheetByName(TAB.META);
    writeTyped_(sh.getRange(sh.getLastRow() + 1, 1, 1, HDR.meta.length), [[Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), 'WD', '巡檢員', what, 0, '', '', result, note || '']]);
  } catch (e) { Logger.log('meta 記錄失敗：' + shortErr_(e && e.message || e)); }
}
function scheduleRecheck_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'watchdogRecheck') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('watchdogRecheck').timeBased().after(20 * 60 * 1000).create();
}
/** 每日巡檢（觸發器 12:15）：今天沒完成且沒有接續在等 → 補跑沒完成的類別＋排 20 分鐘後複查；資料完成但早報沒寄 → 補寄（延遲）；全部正常 → 只記 meta */
function dailyWatchdog() { watchdog_('巡檢'); }
/** 一次性觸發器（巡檢 20 分鐘後）：仍沒完成 → 寄「今天沒有完成」信；跑完自刪 */
function watchdogRecheck() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'watchdogRecheck') ScriptApp.deleteTrigger(t); });
  watchdog_('複查');
}
function watchdog_(phase) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { metaAppend_(phase, '略過', '另一個執行正在跑，20 分鐘後複查'); if (phase === '巡檢') scheduleRecheck_(); return; }
  RUN = null;
  try {
    var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), st = todayStatus_(today), rule = readRules_(snapSS_());
    if (st.complete) {
      if (st.briefSent || Number(rule.BRIEF_ON) !== 1) { metaAppend_(phase, '正常', '5 類今天都已算完，早報' + (st.briefSent ? '已寄' : '暫停中（BRIEF_ON≠1）')); return; }
      var r = sendBrief_('late');
      metaAppend_(phase, r.sent ? '補寄早報' : '早報沒寄出', '資料今天已完成，但早報沒有寄出 → 補寄（主旨加「（延遲）」）');
      return;
    }
    if (phase === '複查') { mailNotDone_(st, today); metaAppend_(phase, '已寄「今天沒有完成」信', '還沒完成：' + catNames_(st.missing)); return; }
    if (st.pending) { scheduleRecheck_(); metaAppend_(phase, '等待接續', '主執行還在分段接續中（有一次性觸發器在等），20 分鐘後複查'); return; }
    scheduleRecheck_();   // 先排好複查：就算這次補跑又被中止，20 分鐘後也會通報
    var state;
    if (st.runState && st.runState.date === today) { state = st.runState; state.late = true; }   // 今天中途停掉的 runAll → 從停下的地方接著跑
    else {
      var now = new Date();
      state = { ts: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss'), date: today, t0: Date.now(), done: st.catsDone.slice(), segments: 1, defer: {}, late: true, carry: st.catsDone.length > 0 };
      stageClear_(snapSS_());
    }
    metaAppend_(phase, '補跑', '今天還沒完成：' + catNames_(st.missing) + '；' + (st.catsDone.length ? '已完成的 ' + st.catsDone.join('') + ' 不重跑' : '5 類全部重跑'));
    PropertiesService.getScriptProperties().setProperty(PROP_RUN_STATE, JSON.stringify(state));
    runLoop_(state);
  } catch (e) { crashMail_('巡檢員補跑', e); throw e; }
  finally { lock.releaseLock(); }
}
/** 「今天沒有完成」通知信（白話；不含個資、不含通關碼） */
function mailNotDone_(st, today) {
  var fails = (st.done.date === today ? (st.done.fail || []) : []), why = {};
  fails.forEach(function (f) { why[f.cat] = f.err; });
  var lines = ['經營快照 ' + today + ' 到現在還沒有全部算完（巡檢員補跑後 20 分鐘複查）。', ''];
  lines.push('・還沒算完的類別：' + catNames_(st.missing));
  st.missing.forEach(function (c) { if (why[c]) lines.push('　－' + catNames_([c]) + '原因：' + why[c]); });
  if (st.pending) lines.push('・目前還有接續中的分段在等，可能稍後會自己完成。');
  lines.push('・快照目前是 ' + (st.done.at ? st.done.at + ' 完成' : '（查不到完成時間）') + ' 的資料；今天的早報與儀表板會標示「上次資料」或「舊資料」。');
  lines.push('', '詳細請看「DIYBC 經營快照」的 meta 分頁：', snapSS_().getUrl(), '', '需要處理時請交給 Cowork。');
  MailApp.sendEmail({ to: MAIL_TO, subject: (isTestMode_() ? '【測試】' : '') + '【經營快照】今天沒有完成', body: lines.join('\n') });
}

// ============================================================
// 3e. 維運與測試小工具（d2）
// ============================================================
/** 印出執行狀態值（不含通關碼） */
function diagState() {
  var p = PropertiesService.getScriptProperties().getProperties(), out = [];
  Object.keys(p).sort().forEach(function (k) { if (k === PROP_GB_PASS || /PASS|KEY|TOKEN|SECRET/i.test(k)) { out.push(k + '＝（已設定，不顯示）'); return; } out.push(k + '＝' + p[k]); });
  out.push('觸發器：' + ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join('、'));
  Logger.log(out.join('\n'));
}
/** T1：預算暫改 60 秒／還原 240 秒 */
function zzTestBudget60() { setRuleValue_('RUN_BUDGET_SEC', 60); }
function zzTestBudgetRestore() { setRuleValue_('RUN_BUDGET_SEC', 240); }
/** T2：D 類每次讀資料模擬多等 200 秒 */
function zzTestSlowD() { PropertiesService.getScriptProperties().setProperty(PROP_TEST, JSON.stringify({ slow: { D: 200 }, noEst: 1 })); }
/** T4：模擬 D 類讀取失敗 */
function zzTestFailD() { PropertiesService.getScriptProperties().setProperty(PROP_TEST, JSON.stringify({ fail: ['D'], noEst: 1 })); }
/** T5：假裝今天是明天（用今天的快照寄早報＝舊資料） */
function zzTestFakeTomorrow() { PropertiesService.getScriptProperties().setProperty(PROP_TEST, JSON.stringify({ fakeToday: addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), 1), noEst: 1 })); }
/** T3：刪除今天的完成紀錄（只刪狀態值，不動分頁） */
function zzTestClearToday() { var p = PropertiesService.getScriptProperties(); [PROP_CAT_DONE, PROP_DONE, PROP_BRIEF_SENT_TEST].forEach(function (k) { p.deleteProperty(k); }); }
/** T7：D 類第一次讀取模擬卡死 400 秒（不做時間檢查，會被 Google 在 360 秒中止） */
function zzTestHangD() { PropertiesService.getScriptProperties().setProperty(PROP_TEST, JSON.stringify({ hang: { D: 400 }, noEst: 1 })); }
/** 印出並清掉 SNAP_RUN_STATE（中斷留下的接續狀態）；同時刪掉等待中的 runAllContinue */
function zzClearRunState() {
  var p = PropertiesService.getScriptProperties();
  Logger.log('SNAP_RUN_STATE（清除前）＝' + (p.getProperty(PROP_RUN_STATE) || '（不存在）'));
  p.deleteProperty(PROP_RUN_STATE); clearContinueTriggers_();
  Logger.log('已清除；觸發器：' + ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join('、'));
}
/** D3-3 核對：評論 Sheet C 欄「整欄」與「最後 500 列」的資料最新日（只讀） */
function zzCheckRevLatest() {
  var sh = SpreadsheetApp.openById(SRC.REV_SS).getSheetByName('工作表1'), last = sh.getLastRow();
  var mx = function (v) { var l = ''; v.forEach(function (r) { var x = r[0], s = (x instanceof Date) ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd') : String(x || '').slice(0, 10); if (/^\d{4}-\d{2}-\d{2}$/.test(s) && s > l) l = s; }); return l; };
  var t0 = Date.now(), full = mx(sh.getRange(2, 3, Math.max(1, last - 1), 1).getValues()), s1 = (Date.now() - t0) / 1000;
  var from = Math.max(2, last - REV_TAIL_ROWS + 1); t0 = Date.now();
  var tail = mx(sh.getRange(from, 3, Math.max(1, last - from + 1), 1).getValues()), s2 = (Date.now() - t0) / 1000;
  Logger.log('評論 Sheet 共 ' + last + ' 列｜整欄最新日＝' + full + '（' + s1 + ' 秒）｜最後 ' + REV_TAIL_ROWS + ' 列最新日＝' + tail + '（' + s2 + ' 秒）｜' + (full === tail ? '相同' : '不同！'));
}
/** 關閉所有測試旗標 */
function zzTestOff() { PropertiesService.getScriptProperties().deleteProperty(PROP_TEST); }
function setRuleValue_(code, val) {
  var ss = snapSS_(); readRules_(ss);   // 先補齊 dim_rule 缺的規則列
  var sh = ss.getSheetByName(TAB.RULE), v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === code) { sh.getRange(i + 1, 4).setValue(val); Logger.log(code + ' → ' + val); return; }
  throw new Error('dim_rule 找不到 ' + code);
}

// ============================================================
// 3f. D2-5 去年訂位月資料快取（cache_ly，隱藏分頁；給階段 e1 的 C3 用）
//   去年的月資料不會再變：每個月第一次需要時讀一次 fn=month，按日按店存「有效筆數／人數」，之後直接讀這張表。
//   只存計數，不存客人個資；有效＝照抄的 isValid（非已取消）＋ dropTest（排除測試店）。
// ============================================================
function lyDaily_(ym, map) {
  var ss = snapSS_(), sh = ss.getSheetByName(TAB.LY);
  if (!sh) { sh = ss.insertSheet(TAB.LY, ss.getSheets().length); sh.getRange(1, 1, 1, HDR.cache_ly.length).setValues([HDR.cache_ly]); sh.hideSheet(); }
  var v = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, HDR.cache_ly.length).getValues() : [];
  var rows = v.filter(function (r) { return String(r[0]) === ym; }).map(function (r) { return [String(r[0]), dstr_(r[1]), Number(r[2]), Number(r[3]), Number(r[4])]; });
  var src = 'cache_ly';
  if (!rows.length) {
    var M = fetchJson_('C', '訂位 fn=month ' + ym + '（去年，存入 cache_ly）', SRC.RSV_API + '?fn=month&ym=' + ym);
    var D = dropTest(M.cols, M.rows), idx = D.idx, agg = {}, nowTs = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    D.rows.forEach(function (r) {
      if (!isValid(r, idx)) return;
      var sid = map.byCol['訂位'][String(r[idx.store] || '').trim()]; if (!sid) return;
      var d = String(r[idx.date] || '').slice(0, 10); if (d.slice(0, 7) !== ym) return;
      var a = agg[d + '|' + sid] || (agg[d + '|' + sid] = { n: 0, p: 0 }); a.n++; a.p += Number(r[idx.party_size]) || 0;
    });
    M = null; D = null;
    rows = Object.keys(agg).sort().map(function (k) { var p = k.split('|'); return [ym, p[0], Number(p[1]), agg[k].n, agg[k].p]; });
    if (rows.length) writeTyped_(sh.getRange(sh.getLastRow() + 1, 1, rows.length, HDR.cache_ly.length), rows.map(function (r) { return r.concat([nowTs]); }));
    src = 'fn=month';
  }
  var out = {};
  rows.forEach(function (r) { (out[r[1]] = out[r[1]] || {})[String(r[2])] = { n: r[3], p: r[4] }; });
  return { byDate: out, src: src, rows: rows.length };
}
/** 手動：把「今天起 14 天」對應的去年月份先存進 cache_ly，並印出各店去年同期 14 天的有效筆數／人數 */
function warmCacheLy() {
  RUN = newRun_();
  var ss = snapSS_(), map = readStoreMap_(ss), today = RUN.date, from = lyOf_(today), to = lyOf_(addDays_(today, 13)), yms = [];
  [from.slice(0, 7), to.slice(0, 7)].forEach(function (y) { if (yms.indexOf(y) < 0) yms.push(y); });
  var tot = {}, log = [];
  yms.forEach(function (ym) {
    var c = lyDaily_(ym, map); log.push(ym + '：' + c.rows + ' 列（來源 ' + c.src + '）');
    Object.keys(c.byDate).forEach(function (d) { if (d < from || d > to) return; Object.keys(c.byDate[d]).forEach(function (sid) { var t = tot[sid] || (tot[sid] = { n: 0, p: 0 }); t.n += c.byDate[d][sid].n; t.p += c.byDate[d][sid].p; }); });
  });
  map.list.forEach(function (sid) { var t = tot[sid] || { n: 0, p: 0 }; log.push(sid + ' ' + map.name[sid] + '：' + t.n + ' 筆／' + t.p + ' 人'); });
  Logger.log('去年同期 ' + from + '～' + to + '\n' + log.join('\n'));
}

/** 單類手動執行：只換這一類的 snap_latest 欄、alerts、今天的 history、meta（不產生 0 資料、不寄信） */
function runSingle_(cat) {
  RUN = newRun_();
  var r = {}; r[cat] = runOneCat_(cat);
  finalize_(r, false);
  return r[cat].ok;
}
function snapA_staffing() { return runSingle_('A'); }
function snapB_campaign() { return runSingle_('B'); }
function snapC_booking()  { return runSingle_('C'); }
function snapD_reviews()  { return runSingle_('D'); }
function snapE_members()  { return runSingle_('E'); }

/** 第一次執行：建立「DIYBC 經營快照」Sheet（僅限本人），建 6 個分頁表頭，填門檻與店名對照 */
function setupSnapshotSheet() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty(PROP_SNAP_ID), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create(SNAP_NAME); props.setProperty(PROP_SNAP_ID, ss.getId()); }
  ss.setSpreadsheetTimeZone(TZ);
  var order = [TAB.LATEST, TAB.HIST, TAB.ALERTS, TAB.RULE, TAB.MAP, TAB.META];
  order.forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (!sh) { if (i === 0 && ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0) { sh = ss.getSheets()[0]; sh.setName(name); } else { sh = ss.insertSheet(name); } }
    if (sh.getLastRow() === 0) {
      if (name === TAB.LATEST) sh.getRange(1, 1, 1, 3).setValues([['店號', '標準店名', '區']]);
      else sh.getRange(1, 1, 1, HDR[name].length).setValues([HDR[name]]);
      if (name === TAB.RULE) sh.getRange(2, 1, RULE_DEFAULTS.length, 5).setValues(RULE_DEFAULTS);
      if (name === TAB.MAP) sh.getRange(2, 1, STORE_MAP_DEFAULTS.length, 9).setValues(STORE_MAP_DEFAULTS);
      sh.setFrozenRows(1);
    }
  });
  Logger.log('快照 Sheet：' + ss.getUrl());
  return ss.getId();
}

// ============================================================
// 4. A 人力×營收（排班 getWeekKpiData ＋ POS）
// ============================================================
function calcA_(res, rule, map) {
  var W = fetchJson_('A', '排班 getWeekKpiData', SRC.SCH_API + '?action=getWeekKpiData');
  if (!W || !W.ok || !W.week_kpi) throw new Error('排班端點回應異常');
  var P = fetchJson_('A', 'POS 主端點', SRC.POS_API);
  // 閾值：照抄的預設值（schedule L763）再以 dim_rule 覆蓋前 4 個（dayWd/dayWe 不進紅旗）
  thresholds.wdHours = rule.SCH_WD_HOURS; thresholds.weHours = rule.SCH_WE_HOURS; thresholds.rph = rule.SCH_RPH_MIN; thresholds.rev = rule.SCH_REV_MIN;

  var sales = W.daily_sales || [], hours = W.daily_hours || [], wk = W.week_kpi || [];
  var lastSales = sales.reduce(function (m, x) { return x.date > m ? x.date : m; }, '');
  var lastHours = hours.reduce(function (m, x) { return x.date > m ? x.date : m; }, '');
  setMetaLatest_('A', '排班 getWeekKpiData', wk.length + ' 週列／' + sales.length + ' 日營收／' + hours.length + ' 日工時', '營收 ' + lastSales + '／工時 ' + lastHours);
  var posLatest = (P.daily || []).reduce(function (m, x) { return x.date > m ? x.date : m; }, '');
  setMetaLatest_('A', 'POS 主端點', (P.monthlyByStore || []).length + ' 店月列', 'POS 到 ' + posLatest + '（lastUpdated ' + (P.lastUpdated || '') + '）');
  RUN.fresh.A = { lastSales: lastSales, posLatest: posLatest };
  var stores = map.list;   // '1'..'12'
  var sumBy = function (arr, key, sid, from, to) {
    var t = 0; arr.forEach(function (x) { if (String(Number(x.store_id)) !== sid) return; if (x.date < from || x.date > to) return; t += Number(x[key] != null ? x[key] : 0) || 0; }); return t;
  };
  var netOf = function (sid, from, to) {
    var t = 0; sales.forEach(function (d) { if (String(Number(d.store_id)) !== sid) return; if (d.date < from || d.date > to) return; t += Number(d.net_revenue != null ? d.net_revenue : d.revenue) || 0; }); return t;
  };
  var lastSalesOf = function (sid) { var m = ''; sales.forEach(function (d) { if (String(Number(d.store_id)) === sid && d.date > m) m = d.date; }); return m; };
  var yRule = function (g) { if (g === '' || g === null) return ''; return g < rule.YOY_RED ? '🔴' : (g < rule.YOY_ORANGE ? '🟠' : '🟢'); };

  // ---- A1 上週（最近一個完整週：週一～週日都 ≤ 營收最後一天） ----
  var weeks = []; wk.forEach(function (r) { if (weeks.indexOf(r.week_start) < 0) weeks.push(r.week_start); }); weeks.sort();
  var fullWk = ''; weeks.forEach(function (ws) { if (addDays_(ws, 6) <= lastSales) fullWk = ws; });
  var ywOf = function (ws) { var r = wk.filter(function (x) { return x.week_start === ws; })[0]; return r ? r.year_week : ws; };
  var a1p = fullWk ? ywOf(fullWk) + '（' + fullWk.slice(5) + '～' + addDays_(fullWk, 6).slice(5) + '）' : '';
  var tH = 0, tN = 0, tF = 0, tNall = 0, noH1 = [];
  stores.forEach(function (sid) {
    var row = wk.filter(function (x) { return x.week_start === fullWk && String(Number(x.store_id)) === sid; })[0];
    if (!row || !(Number(row.total_hours) > 0)) { noH1.push(sid); if (row) tNall += Number(row.total_net_revenue) || 0; }
    if (!row) return;
    var fl = computeFlags(row).map(function (f) { return f.label; });
    rec_(res, sid, 'A1_hours', a1p, r1_(Number(row.total_hours)), '', '', lastSales, '', 'A 上週工時');
    rec_(res, sid, 'A1_net', a1p, r0_(Number(row.total_net_revenue)), '', '', lastSales, '', 'A 上週淨營收');
    rec_(res, sid, 'A1_rph', a1p, r0_(Number(row.rev_per_hour_net)), '', '', lastSales, fl.indexOf('人效偏低') >= 0 ? '🟠' : '🟢', 'A 上週 rev/h');
    rec_(res, sid, 'A1_flags', a1p, fl.length, fl.join('、'), '紅旗項目', lastSales, fl.length ? '🟠' : '🟢', 'A 上週紅旗');
    if (fl.length) alert_(res, '🟠', sid, map.name[sid] + ' 上週（' + a1p + '）紅旗：' + fl.join('、'), 'rev/h ' + r0_(Number(row.rev_per_hour_net)) + '／營收 ' + r0_(Number(row.total_net_revenue)), '平日 ' + thresholds.wdHours + 'h／假日 ' + thresholds.weHours + 'h／rev/h ' + thresholds.rph + '／週營收 ' + thresholds.rev);
    if (Number(row.total_hours) > 0) { tH += Number(row.total_hours) || 0; tN += Number(row.total_net_revenue) || 0; }
    if (fl.length) tF++;
  });
  rec_(res, ALL_ROW, 'A1_hours', a1p, r1_(tH), '', '', lastSales, '', 'A 上週工時');
  rec_(res, ALL_ROW, 'A1_net', a1p, r0_(tN + tNall), '', '', lastSales, '', 'A 上週淨營收');
  rec_(res, ALL_ROW, 'A1_rph', a1p, tH ? r0_(tN / tH) : '', r0_(tN), noH1.length ? '只算有工時的店（營收 ÷ 工時），排除 ' + noH1.join('、') + ' 號店' : '只算有工時的店（營收 ÷ 工時）', lastSales, '', 'A 上週 rev/h');
  rec_(res, ALL_ROW, 'A1_nohours', a1p, noH1.join('、'), '', '上週沒有工時資料的店', lastSales, '', '');
  rec_(res, ALL_ROW, 'A1_flags', a1p, tF, '', '紅旗店數', lastSales, '', 'A 上週紅旗');

  // ---- 本月／上月 ----
  var today = RUN.date, curYm = today.slice(0, 7), prvYm = prevYmOf_(curYm);
  // A1m 紅旗週數（週一落在該月的週；週原值、computeFlags；本月含未完結週由 computeFlags 自行排除營收 0 週）
  [[prvYm, '上月'], [curYm, '本月']].forEach(function (pm) {
    var totF = 0, totW = 0;
    stores.forEach(function (sid) {
      var rows = wk.filter(function (x) { return String(Number(x.store_id)) === sid && String(x.week_start).slice(0, 7) === pm[0]; });
      var nf = rows.filter(function (x) { return computeFlags(x).length > 0; }).length;
      rec_(res, sid, 'A1_flagweeks', pm[0], nf, rows.length, '週數', lastSales, '', 'A 紅旗週（' + pm[1] + '）');
      totF += nf; totW += rows.length;
    });
    rec_(res, ALL_ROW, 'A1_flagweeks', pm[0], totF, totW, '店×週', lastSales, '', 'A 紅旗週（' + pm[1] + '）');
  });

  // ---- A2 本月到營收最後一天、A3 去年同期 ----
  var m1 = curYm + '-01';
  var tot = { h: 0, n: 0, ly: 0, pos: 0, nH: 0 }, cutAll = '', noH2 = [];
  stores.forEach(function (sid) {
    var cut = lastSalesOf(sid); if (cut > addDays_(today, -1)) cut = addDays_(today, -1);
    if (!cut || cut < m1) { rec_(res, sid, 'A2_net', curYm, '', '', '', cut, '', 'A 本月到昨天淨營收'); return; }
    if (cut > cutAll) cutAll = cut;
    var h = sumBy(hours, 'hours', sid, m1, cut), n = netOf(sid, m1, cut), ly = netOf(sid, lyOf_(m1), lyOf_(cut));
    var per = m1 + '～' + cut;
    rec_(res, sid, 'A2_hours', per, r1_(h), '', '', cut, '', 'A 本月到營收最後一天工時');
    rec_(res, sid, 'A2_net', per, r0_(n), '', '', cut, '', 'A 本月到營收最後一天淨營收');
    rec_(res, sid, 'A2_rph', per, h ? r0_(n / h) : '', '', '', cut, '', 'A 本月 rev/h');
    var pos = (P.monthlyByStore || []).filter(function (m) { return String(m.storeCode) === sid && m.yearMonth === curYm; }).reduce(function (s, m) { return s + (Number(m.revenue_net) || 0); }, 0);
    rec_(res, sid, 'A2_pos_net', curYm, r0_(pos), r0_(n), '排班日資料加總', P.lastUpdated || '', Math.round(pos) === Math.round(n) ? '🟢' : '🟡', '');
    var g = ly ? r1_((n - ly) / ly * 100) : '';
    rec_(res, sid, 'A3_ly_net', lyOf_(m1) + '～' + lyOf_(cut), r0_(ly), '', '', cut, '', 'A 去年同期淨營收');
    rec_(res, sid, 'A3_growth', per, g, r0_(ly), '去年同期（同一天截止）', cut, yRule(g), 'A 同期成長%');
    if (g !== '' && g < rule.YOY_ORANGE) alert_(res, yRule(g), sid, map.name[sid] + ' 本月（' + m1.slice(5) + '～' + cut.slice(5) + '）淨營收比去年同期 ' + (g > 0 ? '+' : '') + g + '%', r0_(n) + ' vs ' + r0_(ly), rule.YOY_ORANGE + '% 🟠／' + rule.YOY_RED + '% 🔴');
    if (h > 0) tot.nH += n;
    else { noH2.push(sid); if (n > 0) alert_(res, '🟡', sid, map.name[sid] + '（' + sid + ' 號店）本月班表未上傳：本月（' + m1.slice(5) + '～' + cut.slice(5) + '）有營收 ' + ntd_(n) + '，但排班工時是 0，所以這家店的人效算不出來', ntd_(n) + '／0 小時', '本月營收 > 0 且工時＝0'); }
    tot.h += h; tot.n += n; tot.ly += ly; tot.pos += pos;
  });
  var perAll = m1 + '～' + cutAll, gAll = tot.ly ? r1_((tot.n - tot.ly) / tot.ly * 100) : '';
  var noHtxt = noH2.length ? '只算有工時的店（營收 ÷ 工時），排除 ' + noH2.join('、') + ' 號店' : '只算有工時的店（營收 ÷ 工時）';
  rec_(res, ALL_ROW, 'A2_hours', perAll, r1_(tot.h), '', '', cutAll, '', 'A 本月到營收最後一天工時');
  rec_(res, ALL_ROW, 'A2_net', perAll, r0_(tot.n), '', '', cutAll, '', 'A 本月到營收最後一天淨營收');
  rec_(res, ALL_ROW, 'A2_rph', perAll, tot.h ? r0_(tot.nH / tot.h) : '', r0_(tot.nH), noHtxt, cutAll, '', 'A 本月 rev/h');
  rec_(res, ALL_ROW, 'A2_nohours', perAll, noH2.join('、'), '', '本月沒有工時資料的店', cutAll, '', 'A 無工時資料店');
  rec_(res, ALL_ROW, 'A2_pos_net', curYm, r0_(tot.pos), r0_(tot.n), '排班日資料加總', P.lastUpdated || '', '', '');
  rec_(res, ALL_ROW, 'A3_ly_net', lyOf_(m1) + '～' + lyOf_(cutAll), r0_(tot.ly), '', '', cutAll, '', 'A 去年同期淨營收');
  rec_(res, ALL_ROW, 'A3_growth', perAll, gAll, r0_(tot.ly), '去年同期（同一天截止）', cutAll, yRule(gAll), 'A 同期成長%');

  // ---- A4 上月整月 ----
  var pFrom = prvYm + '-01', pTo = prvYm + '-' + String(daysInMonth_(prvYm)).padStart(2, '0');
  var t4 = { h: 0, n: 0, ly: 0, nH: 0 }, noH4 = [];
  stores.forEach(function (sid) {
    var h = sumBy(hours, 'hours', sid, pFrom, pTo), n = netOf(sid, pFrom, pTo), ly = netOf(sid, lyOf_(pFrom), lyOf_(pTo));
    var g = ly ? r1_((n - ly) / ly * 100) : '';
    rec_(res, sid, 'A4_hours', prvYm, r1_(h), '', '', pTo, '', 'A 上月工時');
    rec_(res, sid, 'A4_net', prvYm, r0_(n), '', '', pTo, '', 'A 上月淨營收');
    rec_(res, sid, 'A4_rph', prvYm, h ? r0_(n / h) : '', '', '', pTo, '', 'A 上月 rev/h');
    rec_(res, sid, 'A4_ly_net', lyOf_(prvYm), r0_(ly), '', '', pTo, '', 'A 上月去年同月');
    rec_(res, sid, 'A4_growth', prvYm, g, r0_(ly), '去年同月', pTo, yRule(g), 'A 上月成長%');
    if (g !== '' && g < rule.YOY_ORANGE) alert_(res, yRule(g), sid, map.name[sid] + ' 上月（' + prvYm + '）淨營收比去年同月 ' + g + '%', r0_(n) + ' vs ' + r0_(ly), rule.YOY_ORANGE + '% 🟠／' + rule.YOY_RED + '% 🔴');
    if (h > 0) t4.nH += n; else noH4.push(sid);
    t4.h += h; t4.n += n; t4.ly += ly;
  });
  var g4 = t4.ly ? r1_((t4.n - t4.ly) / t4.ly * 100) : '';
  rec_(res, ALL_ROW, 'A4_hours', prvYm, r1_(t4.h), '', '', pTo, '', 'A 上月工時');
  rec_(res, ALL_ROW, 'A4_net', prvYm, r0_(t4.n), '', '', pTo, '', 'A 上月淨營收');
  rec_(res, ALL_ROW, 'A4_rph', prvYm, t4.h ? r0_(t4.nH / t4.h) : '', r0_(t4.nH), noH4.length ? '只算有工時的店（營收 ÷ 工時），排除 ' + noH4.join('、') + ' 號店' : '只算有工時的店（營收 ÷ 工時）', pTo, '', 'A 上月 rev/h');
  rec_(res, ALL_ROW, 'A4_nohours', prvYm, noH4.join('、'), '', '上月沒有工時資料的店', pTo, '', '');
  rec_(res, ALL_ROW, 'A4_ly_net', lyOf_(prvYm), r0_(t4.ly), '', '', pTo, '', 'A 上月去年同月');
  rec_(res, ALL_ROW, 'A4_growth', prvYm, g4, r0_(t4.ly), '去年同月', pTo, yRule(g4), 'A 上月成長%');

  // ---- e1 snap_series：各店近 35 天每日淨營收／去年同日淨營收／工時（資料同上面的排班 getWeekKpiData，不另外讀）----
  var tS = Date.now(), dN = {}, dH = {};
  sales.forEach(function (d) { var k = String(Number(d.store_id)) + '|' + d.date; dN[k] = (dN[k] || 0) + (Number(d.net_revenue != null ? d.net_revenue : d.revenue) || 0); });
  hours.forEach(function (x) { var k = String(Number(x.store_id)) + '|' + x.date; dH[k] = (dH[k] || 0) + (Number(x.hours != null ? x.hours : 0) || 0); });
  var sEnd = addDays_(today, -1); if (lastSales && lastSales < sEnd) sEnd = lastSales;
  var sRows = [];
  for (var di = 34; di >= 0; di--) {
    var dd = addDays_(sEnd, -di), ldd = lyOf_(dd);
    stores.forEach(function (sid) {
      var k = sid + '|' + dd, lk = sid + '|' + ldd;
      sRows.push([dd, Number(sid), (k in dN) ? r0_(dN[k]) : '', (lk in dN) ? r0_(dN[lk]) : '', (k in dH) ? r1_(dH[k]) : '', RUN.ts]);
    });
  }
  var nS = seriesUpsert_(sRows);
  logMeta_('A', 'e1 snap_series（近 35 天每日）', addDays_(sEnd, -34) + '～' + sEnd, (Date.now() - tS) / 1000, nS, sEnd, '成功', '');
}

/** e1：snap_series 同日同店覆蓋（其餘日期保留，保留 400 天） */
function seriesUpsert_(rows) {
  var ss = snapSS_(), sh = ss.getSheetByName(TAB.SERIES);
  if (!sh) { sh = ss.insertSheet(TAB.SERIES, ss.getSheets().length); sh.getRange(1, 1, 1, HDR.snap_series.length).setValues([HDR.snap_series]); }
  var old = sh.getLastRow() >= 2 ? sh.getRange(2, 1, sh.getLastRow() - 1, HDR.snap_series.length).getValues() : [];
  var put = {}, cut = addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), -HIST_KEEP_DAYS);
  rows.forEach(function (r) { put[r[0] + '|' + r[1]] = 1; });
  var keep = old.map(function (r) { return [dstr_(r[0])].concat(r.slice(1)); }).filter(function (r) { return r[0] && r[0] >= cut && !put[r[0] + '|' + Number(r[1])]; });
  var all = keep.concat(rows).sort(function (a, b) { return a[0] < b[0] ? -1 : (a[0] > b[0] ? 1 : a[1] - b[1]); });
  if (old.length) sh.getRange(2, 1, old.length, HDR.snap_series.length).clearContent();
  if (all.length) writeTyped_(sh.getRange(2, 1, all.length, HDR.snap_series.length), all);
  return rows.length;
}

// ============================================================
// 5. B 檔期叫貨（採購專用檔 agg_evneed／agg_evcurve／agg_endcut／fact_inventory ＋ 月檔 fact_shopline）
// ============================================================
// 採購頁的全域變數（照抄的函式會用到；由本函式填值，等同頁面 bootstrap）
var SKUS = [], SKU = {}, PARAM = {}, NEWAGG = {}, EVCV = {}, EVND = null, EVNDLD = null;

function calcB_(out, rule, map) {
  TODAY = new Date();
  var dSku = readTable_('B', SRC.PUR_DASH, 'dim_sku');
  var dPar = readTable_('B', SRC.PUR_DASH, 'dim_store_par');
  var dNi = readTable_('B', SRC.PUR_DASH, 'agg_newitem');
  var evc = readTable_('B', SRC.PUR_DASH, 'agg_evcurve');
  var rs = readTable_('B', SRC.PUR_DASH, 'agg_evneed');
  var endc = readTable_('B', SRC.PUR_DASH, 'agg_endcut');
  var fInv = readTable_('B', SRC.PUR_DASH, 'fact_inventory');
  var shop = readTable_('B', SRC.PUR_MONTH, 'fact_shopline');
  var storeSet = {}, res = [];   // 對應頁面 bootstrap 的 res[13]＝agg_evcurve
  res[13] = evc;
  // ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L463–L466、L487–L488、L492 ｜ 採購 bootstrap：EVCV、NEWAGG、SKUS、PARAM 解析（邏輯一字不改）----
    EVCV={};(res[13]||[]).forEach(function(r){if(r["週序"]==null||r["事件名"]==null||r["今年週起"]==null)return;var ev=String(r["事件名"]||"").trim();if(!ev)return;
      var f=gvDate(r["今年週起"]),t=gvDate(r["今年週訖"]),pf=gvDate(r["去年週起"]),pt=gvDate(r["去年週訖"]);
      (EVCV[ev]=EVCV[ev]||[]).push({idx:num(r["週序"]),pFrom:pf?ymd(pf):String(r["去年週起"]||""),pTo:pt?ymd(pt):String(r["去年週訖"]||""),pQty:num(r["去年份數"]),share:num(r["佔比%"]),cFrom:f?ymd(f):String(r["今年週起"]||""),cTo:t?ymd(t):String(r["今年週訖"]||""),cSold:num(r["今年已售"]),cEst:num(r["今年推估"]),rel:num(r["相對開賣週"])});});
    Object.keys(EVCV).forEach(function(ev){EVCV[ev].sort(function(a,b){return a.idx-b.idx;});});
    NEWAGG=buildNewAgg(dNi);   /* ca 批：多檔期同料加總保留（見 buildNewAgg） */
    SKUS=dSku.map(function(r){return {id:r["sku_id"],name:r["品名"],cat:r["品類別"]||"食材",vendor:r["廠商"]||"",useUnit:r["使用單位"]||"g",buyUnit:r["採購單位"]||"包",xfer:String(r["調撥方式"]||"").trim(),packQty:num(r["每採購單位內容量"])||1,price:num(r["單價"]),defZone:String(r["預設分區"]||"").trim(),src:String(r["來源"]||"").trim(),stores:String(r["適用店"]||"").split(/[,，、;\s]+/).map(function(x){return x.trim();}).filter(Boolean),status:String(r["狀態"]||"").trim(),lastValid:(function(v){var d=gvDate(v);return d?ymd(d):"";})(r["最後有效日"]),shelf:String(r["效期"]||"").trim(),aliases:String(r["BOM別名"]||"").split(/[;；,，、]/).map(function(a){return a.trim();}).filter(Boolean)};}).filter(function(s){return s.id&&s.name;});   /* 2026-09-05 ao 批：補全形「；」（與 EventLayer v1.6 同）；資料仍一律用「、」 */
    PARAM={};dPar.forEach(function(r){var st=String(r["店號"]);storeSet[st]=1;((PARAM[st]=PARAM[st]||{})[r["sku_id"]]={par:num(r["標配數量"]),weight:r["權數"]!=null&&r["權數"]!==""?num(r["權數"]):1,orderCycleWeeks:r["訂貨週期天"]?num(r["訂貨週期天"])/7:1,leadDays:(r["交期天"]!=null&&r["交期天"]!==""?num(r["交期天"]):null),manualWeekly:num(r["手動週用量"]),safety:num(r["安全庫存"]),zone:r["盤點分區"]||"未分區",ord:num(r["盤點順序"]),used:r["使用中"]==null?true:toBool(r["使用中"]),custName:String(r["店品名"]||"").trim(),custVendor:String(r["店廠商"]||"").trim(),custPackQty:num(r["店內容量"])});});
  SKU = {}; SKUS.forEach(function (s) { SKU[s.id] = s; });
  // ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L1530–L1533 ｜ 採購 loadEvNeed 內的 agg_evneed 解析（邏輯一字不改）----
      var m={};
      (rs||[]).forEach(function(r){if(r["剩餘推估用量"]==null||r["sku_id"]==null)return;   /* 分頁不存在會回第一個分頁 → 以欄名把關 */
        var ev=String(r["事件名"]||"").trim(),st=String(r["店號"]||"").trim(),id=String(r["sku_id"]||"").trim();if(!ev||!st||!id)return;
        ((m[ev]=m[ev]||{})[id]=m[ev][id]||{})[st]={rem:num(r["剩餘推估用量"]),used:num(r["已用量"]),sold:num(r["已售份數"]),est:num(r["推估總份數"]),who:String(r["新品清單"]||"")};});
  EVND = m;
  // SHIST.ALL（照 loadShist／loadShistAll 的欄位對應；排除 13 號店；依日期排序）
  var rowsAll = shop.map(function (r) {
    var d = gvDate(r["日期"]);
    return { st: String(r["店號"] == null ? "" : r["店號"]).trim(), date: d ? ymd(d) : String(r["日期"] || "").slice(0, 10), ord: String(r["訂單號碼"] || ""), name: String(r["品項"] || ""), cat: String(r["分類"] || ""), qty: num(r["數量"]), price: num(r["單價"]) };
  }).filter(function (r) { return r.date && r.ord && r.st && r.st !== "13"; });
  rowsAll.sort(function (a, b) { return a.date < b.date ? -1 : (a.date > b.date ? 1 : 0); });
  var shopLatest = rowsAll.length ? rowsAll[rowsAll.length - 1].date : '';
  setMetaLatest_('B', 'fact_shopline', shop.length, shopLatest);
  RUN.fresh.B = { shopLatest: shopLatest };
  var evUpd = rs.reduce(function (m, r) { var d = gvDate(r['更新日']); return d && ymd(d) > m ? ymd(d) : m; }, '');
  setMetaLatest_('B', 'agg_evneed', rs.length, evUpd);

  // ---- B1 各檔期 × 各店 ----
  var evs = Object.keys(EVCV);
  evs.forEach(function (ev) {
    var ws = EVCV[ev], nd = EVND[ev] || {};
    // ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L1559 ｜ viewCampHQ：曲線推估總量／已售（邏輯一字不改）----
      var estTot=0,soldTot=0,_td=ymd(TODAY);ws.forEach(function(w){estTot+=(w.cTo<_td?w.cSold:Math.max(w.cEst,w.cSold));soldTot+=w.cSold;});   /* bv 批：已過的週一律用實際已售 */
    rec_(out, ALL_ROW, 'B1_est_total', ev, Math.round(estTot), '', '', evUpd, '', 'B ' + ev + ' 推估總份數');
    rec_(out, ALL_ROW, 'B1_sold_total', ev, Math.round(soldTot), '', '', evUpd, '', 'B ' + ev + ' 已售份數');
    var agg = {};   // st -> {items, needItems, demNT, shipNT, overs（專屬且超出 ≥ 門檻，亮燈）, shared（共用品，不亮燈）, oldRule（修改前口徑，診斷用）}
    Object.keys(nd).forEach(function (id) {
      var s = SKU[id]; if (!s || isDur(s)) return;
      var per = {}, sumRem = 0, sumUsed = 0, sumShip = 0, sumOn = 0, sumNeed = 0, sumEst = 0, sumSold = 0, dc = false, sumPre = 0, sumPrePacks = 0, preFirst = "", nOver = 0;
      // ---- 照抄：dashboard-purchase.html ｜ md5 3a922ade73669ad4ebdff1ac34e7f045 ｜ 原行號 L1595–L1602 ｜ viewCampHQ：逐店 已出貨／估計在店／尚需（邏輯一字不改）----
          Object.keys(nd[id]).forEach(function(st){var x=nd[id][st];var isdc=isDC(st,s);if(isdc)dc=true;
            var sh=(rowsAll&&isdc)?campHQShipped(ev,s,rowsAll.filter(function(r){return String(r.st)===st;})):null;
            var shipped=sh?sh.g:0,onhand=sh?Math.max(0,shipped-x.used):0,gap=x.rem-onhand;   /* gap 正＝缺、負＝多 */
            per[st]={rem:x.rem,used:x.used,sold:x.sold,est:x.est,shipped:shipped,packs0:sh?sh.packs:0,onhand:onhand,gap:gap,need:Math.max(0,gap),surplus:Math.max(0,-gap),pq:pqEff(st,s),dc:isdc,noShip:!sh,xin:[],xout:[],buy:0,packs:0,
              preG:sh?sh.preG:0,prePacks:sh?sh.prePacks:0,preFirst:sh?sh.preFirst:"",win:sh?sh.from:""};
            if(sh&&sh.preG>0){sumPre+=sh.preG;sumPrePacks+=sh.prePacks;if(!preFirst||(sh.preFirst&&sh.preFirst<preFirst))preFirst=sh.preFirst;}
            if(isdc&&sh&&shipped-x.used<0)nOver++;   /* 已用 ＞ 期間內已出貨：檔期前就有庫存，或更早的出貨不在視窗 */
            sumRem+=x.rem;sumUsed+=x.used;sumShip+=shipped;sumOn+=onhand;sumNeed+=Math.max(0,gap);sumEst+=x.est;sumSold+=x.sold;});
      var unitNT = (s.price || 0) / Math.max(1, s.packQty || 1);   // 採購頁 cost()：單價 ÷ 每採購單位內容量
      var excl = campExcl(id);   // M3：專屬判定＝採購頁 campExcl（agg_newitem「專屬」欄，全部檔期都專屬才算）
      Object.keys(per).forEach(function (st) {
        if (st === '13') return;
        var p = per[st], a = agg[st] || (agg[st] = { items: 0, needItems: 0, demNT: 0, shipNT: 0, overs: [], shared: [], oldRule: [] });
        var dem = p.rem + p.used;
        if (dem > 0 || p.shipped > 0) a.items++;
        if (p.dc && !p.noShip) {
          if (p.need > 0) a.needItems++;
          a.demNT += dem * unitNT; a.shipNT += p.shipped * unitNT;
          var o = null;
          if (dem <= 0 && p.shipped > 0) { if (rule.CAMP_ZERO_NEED_RED) o = { name: s.name, pct: null, light: '🔴', nt: p.shipped * unitNT }; }
          else if (dem > 0) {
            var pc = p.shipped / dem * 100;
            if (pc > rule.CAMP_SHIP_RED) o = { name: s.name, pct: pc, light: '🔴', nt: (p.shipped - dem) * unitNT };
            else if (pc > rule.CAMP_SHIP_ORANGE) o = { name: s.name, pct: pc, light: '🟠', nt: (p.shipped - dem) * unitNT };
          }
          if (o) {
            a.oldRule.push(o);
            if (!excl) a.shared.push(o);
            else if (o.nt >= rule.CAMP_MIN_EXCESS_NTD) a.overs.push(o);
          }
        }
      });
    });
    var tot = { items: 0, needItems: 0, demNT: 0, shipNT: 0, overs: 0, excess: 0, shared: 0, oldN: 0, oldNT: 0 };
    var byNt = function (x, y) { return y.nt - x.nt; };
    var sumNt = function (arr) { return arr.reduce(function (t, o) { return t + o.nt; }, 0); };
    var nameNt = function (o) { return o.name + (o.pct == null ? '（需求 0）' : '') + ' 超出 ' + ntd_(o.nt); };
    map.list.forEach(function (sid) {
      var a = agg[sid]; if (!a) return;
      a.overs.sort(byNt); a.shared.sort(byNt); a.oldRule.sort(byNt);
      var excess = sumNt(a.overs);
      var top = a.overs.slice(0, 5).map(nameNt).join('、');
      var lt = a.overs.some(function (o) { return o.light === '🔴'; }) ? '🔴' : (a.overs.length ? '🟠' : '🟢');
      var pct = a.demNT > 0 ? r1_(a.shipNT / a.demNT * 100) : '';
      rec_(out, sid, 'B1_items', ev, a.items, '', '', evUpd, '', 'B ' + ev + ' 品項數');
      rec_(out, sid, 'B1_need_items', ev, a.needItems, '', '', evUpd, '', 'B ' + ev + ' 還需備品項');
      rec_(out, sid, 'B1_ship_pct', ev, pct, r0_(a.demNT), '推估總需求（NT$，依主檔單價換算）', shopLatest, '', 'B ' + ev + ' 已出貨÷需求%（金額）');
      rec_(out, sid, 'B1_over_items', ev, a.overs.length, top, '專屬品超標（超出金額由大到小，最多 5 個）', shopLatest, lt, 'B ' + ev + ' 超標品項數');
      rec_(out, sid, 'B1_over_excess_nt', ev, r0_(excess), a.overs.length, '專屬品超標項數', shopLatest, lt, 'B ' + ev + ' 超出金額合計');
      rec_(out, sid, 'B1_shared_over_items', ev, a.shared.length, a.shared.slice(0, 5).map(nameNt).join('、'), '共用品超過推估（不亮燈、不進 alerts）', shopLatest, '', '');
      rec_(out, sid, 'B1_over_items_oldrule', ev, a.oldRule.length, r0_(sumNt(a.oldRule)), '修改前口徑（所有品項、無金額門檻）的超出金額合計', shopLatest, '', '');
      if (a.overs.length) alert_(out, lt, sid, map.name[sid] + ' ' + ev + '：' + a.overs.length + ' 項檔期專屬料已出貨超過推估需求，共超出 ' + ntd_(excess) + '。' + top + (a.overs.length > 5 ? ' 等' : ''), ntd_(excess), '專屬品；超出 ≥ ' + ntd_(rule.CAMP_MIN_EXCESS_NTD) + ' 且 >' + rule.CAMP_SHIP_ORANGE + '% 🟠／>' + rule.CAMP_SHIP_RED + '% 或需求 0 🔴');
      tot.items += a.items; tot.needItems += a.needItems; tot.demNT += a.demNT; tot.shipNT += a.shipNT; tot.overs += a.overs.length; tot.excess += excess;
      tot.shared += a.shared.length; tot.oldN += a.oldRule.length; tot.oldNT += sumNt(a.oldRule);
    });
    rec_(out, ALL_ROW, 'B1_items', ev, tot.items, '', '店×品項', evUpd, '', 'B ' + ev + ' 品項數');
    rec_(out, ALL_ROW, 'B1_need_items', ev, tot.needItems, '', '店×品項', evUpd, '', 'B ' + ev + ' 還需備品項');
    rec_(out, ALL_ROW, 'B1_ship_pct', ev, tot.demNT > 0 ? r1_(tot.shipNT / tot.demNT * 100) : '', r0_(tot.demNT), '推估總需求（NT$，依主檔單價換算）', shopLatest, '', 'B ' + ev + ' 已出貨÷需求%（金額）');
    rec_(out, ALL_ROW, 'B1_over_items', ev, tot.overs, '', '店×品項（專屬品）', shopLatest, '', 'B ' + ev + ' 超標品項數');
    rec_(out, ALL_ROW, 'B1_over_excess_nt', ev, r0_(tot.excess), tot.overs, '專屬品超標項數（店×品項）', shopLatest, '', 'B ' + ev + ' 超出金額合計');
    rec_(out, ALL_ROW, 'B1_shared_over_items', ev, tot.shared, '', '共用品超過推估（店×品項，不亮燈）', shopLatest, '', '');
    rec_(out, ALL_ROW, 'B1_over_items_oldrule', ev, tot.oldN, r0_(tot.oldNT), '修改前口徑（店×品項）的超出金額合計', shopLatest, '', '');
  });

  // ---- B2 已過結束日仍在收尾表 ----
  var endUpd = endc.reduce(function (m, r) { var d = gvDate(r['更新日']); return d && ymd(d) > m ? ymd(d) : m; }, '');
  setMetaLatest_('B', 'agg_endcut', endc.length, endUpd);
  var b2 = {}, b2t = 0, tdy = ymd(TODAY);
  endc.forEach(function (r) { var d = gvDate(r['結束日']); if (!d) return; if (ymd(d) >= tdy) return; var st = String(r['店號']).trim(); b2[st] = (b2[st] || 0) + 1; b2t++; });
  map.list.forEach(function (sid) { rec_(out, sid, 'B2_ended_rows', tdy, b2[sid] || 0, '', '結束日 < 今天', endUpd, '', 'B 已過結束日仍在收尾表（列）'); });
  rec_(out, ALL_ROW, 'B2_ended_rows', tdy, b2t, '', '結束日 < 今天', endUpd, '', 'B 已過結束日仍在收尾表（列）');

  // ---- B3 各店最後盤點日 ----
  var inv = {}, invLatest = '';
  fInv.forEach(function (r) { var st = String(r['店號']).trim(); var d = gvDate(r['時間戳']) || gvDate(r['盤點日']); if (!d) return; var s = ymd(d); if (!inv[st] || s > inv[st]) inv[st] = s; if (s > invLatest) invLatest = s; });
  setMetaLatest_('B', 'fact_inventory', fInv.length, invLatest);
  var nStale = 0;
  map.list.forEach(function (sid) {
    var last = inv[sid] || '', days = last ? Math.round((new Date(tdy + 'T00:00:00') - new Date(last + 'T00:00:00')) / 864e5) : '';
    var lt = (!last || days > rule.INV_DAYS_YELLOW) ? '🟡' : '🟢';
    rec_(out, sid, 'B3_last_count', tdy, last || '從未盤點', days, '距今天數', invLatest, lt, 'B 最後盤點日');
    rec_(out, sid, 'B3_days_since', tdy, days, '', '', invLatest, lt, 'B 距上次盤點天數');
    if (lt === '🟡') { nStale++; alert_(out, '🟡', sid, map.name[sid] + (last ? ' 已 ' + days + ' 天沒盤點（最後 ' + last + '）' : ' 從未在採購系統盤點'), last ? days + ' 天' : '從未', '>' + rule.INV_DAYS_YELLOW + ' 天'); }
  });
  rec_(out, ALL_ROW, 'B3_stale_stores', tdy, nStale, '', '超過門檻或從未盤點的店數（全公司列的「距上次盤點天數」留空）', invLatest, '', '');

  // ---- B4 大平台下單資料最新日 ----
  var lag = shopLatest ? Math.round((new Date(tdy + 'T00:00:00') - new Date(shopLatest + 'T00:00:00')) / 864e5) : '';
  rec_(out, ALL_ROW, 'B4_shopline_latest', tdy, shopLatest, lag, '落後天數', shopLatest, '', 'B 大平台下單資料最新日');
}

// ============================================================
// 6. C 訂位（訂位 GAS fn=month／fn=future；gb_bookings 直讀 5 欄）
// ============================================================
function calcC_(res, rule, map) {
  var today = RUN.date, cur = today.slice(0, 7), prv = prevYmOf_(cur);
  var mCur = fetchJson_('C', '訂位 fn=month ' + cur, SRC.RSV_API + '?fn=month&ym=' + cur);
  var mPrv = fetchJson_('C', '訂位 fn=month ' + prv, SRC.RSV_API + '?fn=month&ym=' + prv);
  var fut = fetchJson_('C', '訂位 fn=future', SRC.RSV_API + '?fn=future');
  // 當月＝歷史表＋ future 表（月底那幾天）；照頁面 L1012–L1032 的作法
  var FUTURE = dropTest(fut.cols, fut.rows);
  var _fact = dropTest(mCur.cols, mCur.rows);
  var CURm;
  if (!_fact.rows.length) {   // 整個月都還沒發生（月初）→ 全部來自 future 表
    var rowsF = FUTURE.rows.filter(function (r) { return String(r[FUTURE.idx.date] || '').slice(0, 7) === cur; });
    CURm = { cols: FUTURE.cols, rows: rowsF, idx: FUTURE.idx };
  } else {
    var fMax = maxDateOf(_fact.rows, _fact.idx.date);
    var extraSrc = FUTURE.rows.filter(function (r) { var d = String(r[FUTURE.idx.date] || ''); return d.slice(0, 7) === cur && d > fMax; });
    var extra = remapRows(FUTURE.cols, extraSrc, _fact.cols);
    CURm = { cols: _fact.cols, rows: _fact.rows.concat(extra), idx: _fact.idx };
  }
  var PRVm = dropTest(mPrv.cols, mPrv.rows);
  var factMax = maxDateOf(_fact.rows, _fact.idx.date), futMax = maxDateOf(FUTURE.rows, FUTURE.idx.date);
  setMetaLatest_('C', '訂位 fn=month ' + cur, mCur.rows.length, '歷史到 ' + factMax + '（cached_at ' + (mCur.cached_at || '') + '）');
  setMetaLatest_('C', '訂位 fn=month ' + prv, mPrv.rows.length, 'cached_at ' + (mPrv.cached_at || ''));
  var futMin = FUTURE.rows.reduce(function (m, r) { var d = String(r[FUTURE.idx.date] || '').slice(0, 10); return d && (!m || d < m) ? d : m; }, '');
  setMetaLatest_('C', '訂位 fn=future', fut.rows.length, '未來 ' + futMin + '～' + futMax);
  RUN.fresh.C = { futMin: futMin };
  var tdy = todayStr();
  var statsOf = function (M, sid) {
    var idx = M.idx, rows = M.rows.filter(function (r) { return map.byCol['訂位'][String(r[idx.store]).trim()] === sid; });
    if (sid === ALL_ROW) rows = M.rows;
    var valid = rows.filter(function (r) { return isValid(r, idx); });
    var cancelled = rows.filter(function (r) { return !isValid(r, idx); });
    var cancelRate = pct(cancelled.length, rows.length);
    var pastBooked = rows.filter(function (r) { return r[idx.status] === '已訂位' && r[idx.date] < tdy; });
    var pastBookedNoAttend = pastBooked.filter(function (r) { return Number(r[idx.attended]) === 0; });
    var noAttendRate = pct(pastBookedNoAttend.length, pastBooked.length);
    return { valid: valid.length, total: rows.length, cancelRate: cancelRate, noAttendRate: noAttendRate, noAttend: pastBookedNoAttend.length, past: pastBooked.length };
  };
  var latestCur = factMax + (futMax ? '（含未來到 ' + futMax + '）' : '');
  map.list.concat([ALL_ROW]).forEach(function (sid) {
    var c = statsOf(CURm, sid), p = statsOf(PRVm, sid);
    var up = (c.cancelRate !== null && p.cancelRate !== null) ? c.cancelRate - p.cancelRate : null;
    var lt = (sid !== ALL_ROW && up !== null && up > rule.RSV_CANCEL_UP_ORANGE) ? '🟠' : '🟢';
    rec_(res, sid, 'C1_valid', cur, c.valid, c.total, '全部訂位（含取消）', latestCur, '', 'C 本月有效訂位');
    rec_(res, sid, 'C1_cancel_rate', cur, r1_(c.cancelRate), r1_(p.cancelRate), '上月取消率', latestCur, lt, 'C 本月取消率%');
    rec_(res, sid, 'C1_noattend_rate', cur, r1_(c.noAttendRate), c.noAttend + '/' + c.past, '未勾出席／已過日期的已訂位', latestCur, '', 'C 本月未勾出席率%');
    rec_(res, sid, 'C1_valid', prv, p.valid, p.total, '全部訂位（含取消）', maxDateOf(PRVm.rows, PRVm.idx.date), '', 'C 上月有效訂位');
    rec_(res, sid, 'C1_cancel_rate', prv, r1_(p.cancelRate), '', '', maxDateOf(PRVm.rows, PRVm.idx.date), '', 'C 上月取消率%');
    rec_(res, sid, 'C1_noattend_rate', prv, r1_(p.noAttendRate), p.noAttend + '/' + p.past, '未勾出席／已過日期的已訂位', maxDateOf(PRVm.rows, PRVm.idx.date), '', 'C 上月未勾出席率%');
    if (lt === '🟠') alert_(res, '🟠', sid, map.name[sid] + ' 本月取消率 ' + r1_(c.cancelRate) + '%，比上月（' + r1_(p.cancelRate) + '%）高 ' + r1_(up) + ' 個百分點', r1_(c.cancelRate) + '%', '上月 +' + rule.RSV_CANCEL_UP_ORANGE + ' 百分點');
  });

  // ---- e1 C3 未來 14 天訂位（領先指標）：今天起 14 天有效訂位（future 表，isValid＋dropTest 與訂位頁同口徑）vs 去年同期 14 天（一律讀 cache_ly）----
  //   前 28 天不亮燈、不進 alerts（C3_LIGHT_START 空白＝不亮燈）；「訂位進度」＝未來已訂 ÷ 去年同期實際，不是成長率（未來訂位還會再增加）
  timeGuard_('C3 未來 14 天訂位');
  var t3 = Date.now(), c3To = addDays_(today, 13), lyFrom = lyOf_(today), lyTo = lyOf_(c3To);
  var c3 = {}, c3ly = {}, z3 = function () { return { n: 0, p: 0 }; };
  map.list.concat([ALL_ROW]).forEach(function (sid) { c3[sid] = z3(); c3ly[sid] = z3(); });
  FUTURE.rows.forEach(function (r) {
    if (!isValid(r, FUTURE.idx)) return;
    var d = String(r[FUTURE.idx.date] || '').slice(0, 10); if (d < today || d > c3To) return;
    var sid = map.byCol['訂位'][String(r[FUTURE.idx.store] || '').trim()]; if (!sid) return;
    var p = Number(r[FUTURE.idx.party_size]) || 0;
    [sid, ALL_ROW].forEach(function (k) { c3[k].n++; c3[k].p += p; });
  });
  var lySrc = [];
  [lyFrom.slice(0, 7), lyTo.slice(0, 7)].filter(function (y, i, a) { return a.indexOf(y) === i; }).forEach(function (ym) {
    var c = lyDaily_(ym, map); lySrc.push(ym + ' ' + c.src + ' ' + c.rows + ' 列');
    Object.keys(c.byDate).forEach(function (d) {
      if (d < lyFrom || d > lyTo) return;
      Object.keys(c.byDate[d]).forEach(function (sid) { if (!c3ly[sid]) return; [sid, ALL_ROW].forEach(function (k) { c3ly[k].n += c.byDate[d][sid].n; c3ly[k].p += c.byDate[d][sid].p; }); });
    });
  });
  var c3Per = today + '～' + c3To, c3LyPer = lyFrom + '～' + lyTo, c3Latest = futMin + '～' + futMax;
  map.list.concat([ALL_ROW]).forEach(function (sid) {
    var f = c3[sid], l = c3ly[sid], ratio = l.p > 0 ? r1_(f.p / l.p * 100) : '';
    rec_(res, sid, 'C3_fut14_people', c3Per, f.p, '', '', c3Latest, '', 'C 未來14天訂位人數');
    rec_(res, sid, 'C3_fut14_n', c3Per, f.n, '', '', c3Latest, '', 'C 未來14天訂位筆數');
    rec_(res, sid, 'C3_ly14_people', c3LyPer, l.p, '', 'cache_ly', c3LyPer.slice(-10), '', 'C 去年同期14天人數');
    rec_(res, sid, 'C3_ly14_n', c3LyPer, l.n, '', 'cache_ly', c3LyPer.slice(-10), '', 'C 去年同期14天筆數');
    rec_(res, sid, 'C3_ratio', c3Per, ratio, l.p, '訂位進度＝未來已訂人數 ÷ 去年同期實際人數（不是成長率）', c3Latest, '', 'C 訂位進度%');
  });
  logMeta_('C', 'e1 C3 未來 14 天訂位', c3Per + '｜去年 ' + lySrc.join('；'), (Date.now() - t3) / 1000, c3[ALL_ROW].n, futMax, '成功', '');

  // ---- C2 團體 ①gb_bookings（只讀 date／store／size／status／scanned_at，不讀個資欄）：狀態筆數＋有效團體的對照值／備用口徑 ----
  timeGuard_('gb_bookings');
  var t0 = Date.now(), sh = SpreadsheetApp.openById(SRC.GB_SS).getSheetByName('gb_bookings');
  var lastRow = sh.getLastRow(), hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var col = function (name) { var j = hdr.indexOf(name); if (j < 0) throw new Error('gb_bookings 找不到欄位 ' + name); return lastRow > 1 ? sh.getRange(2, j + 1, lastRow - 1, 1).getValues().map(function (r) { return r[0]; }) : []; };
  var cDate = col('date'), cStore = col('store'), cSize = col('size'), cStatus = col('status'), cScan = col('scanned_at');
  var scanMax = cScan.reduce(function (m, x) { var s = (x instanceof Date) ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd HH:mm') : String(x || ''); return s > m ? s : m; }, '');
  logMeta_('C', 'gb_bookings（5 欄）', SRC.GB_SS.slice(0, 8) + '…', (Date.now() - t0) / 1000, cStore.length, scanMax, '成功', '');
  var gb = {}, statuses = [], dMin = '', dMax = '', unknown = {};
  for (var i = 0; i < cStore.length; i++) {
    var nm = String(cStore[i] || '').trim(); if (!nm) continue;
    var sid = map.byCol['訂位'][nm]; if (!sid) { unknown[nm] = 1; continue; }
    var st = String(cStatus[i] || '').trim() || '(空白)';
    var dd = (cDate[i] instanceof Date) ? Utilities.formatDate(cDate[i], TZ, 'yyyy-MM-dd') : String(cDate[i] || '').slice(0, 10);
    if (dd && (!dMin || dd < dMin)) dMin = dd; if (dd > dMax) dMax = dd;
    if (statuses.indexOf(st) < 0) statuses.push(st);
    var valid = dd >= today && st !== 'canceled';
    [sid, ALL_ROW].forEach(function (k) {
      var g = gb[k] || (gb[k] = { st: {}, vn: 0, vp: 0 }), x = g.st[st] || (g.st[st] = { n: 0, p: 0 });
      x.n++; x.p += Number(cSize[i]) || 0;
      if (valid) { g.vn++; g.vp += Number(cSize[i]) || 0; }
    });
  }
  statuses.sort();
  var perGb = dMin ? dMin + '～' + dMax : '', perValid = today + '～' + (dMax > today ? dMax : today);

  // ---- C2 團體 ②未選甜點／未付足訂金／需立即處理：團體 GAS fn=group（通關碼讀指令碼屬性；原始回傳只在記憶體算計數）----
  var G = readGroupCounts_(map);   // { state:'成功'|'未設定'|'通關碼錯誤'|'讀取失敗', by:{sid:{act,ppl,nd,unpN,unpNT,urg}}, last_done, check }
  RUN.fresh.C.gbState = G.state; RUN.fresh.C.gbErr = G.err || '';
  var gbLatest = G.state === '成功' ? (G.last_done || '') : scanMax;
  map.list.concat([ALL_ROW]).forEach(function (sid) {
    var g = gb[sid] || { st: {}, vn: 0, vp: 0 }, q = (G.by && G.by[sid]) || { act: 0, ppl: 0, nd: 0, unpN: 0, unpNT: 0, urg: 0 };
    statuses.forEach(function (st) {
      var x = g.st[st] || { n: 0, p: 0 };
      rec_(res, sid, 'C2_status_n', perGb, x.n, st, '狀態', scanMax, '', 'C 團體 ' + st + ' 筆數');
      rec_(res, sid, 'C2_status_people', perGb, x.p, st, '狀態', scanMax, '', 'C 團體 ' + st + ' 人數');
    });
    // 有效團體＝團體頁口徑（團體 GAS：active 且 is_group，卡位不算）；團體 GAS 讀不到時才用 gb_bookings 備用口徑（含卡位，會偏多）
    var useGas = G.state === '成功';
    var vNote = useGas ? '團體頁口徑：有效且為團體（卡位不算）；比較值＝gb_bookings 到店日 ≥ 今天且非 canceled（含卡位，僅供對照）'
                       : '備用口徑（團體 GAS ' + G.state + '）：gb_bookings 到店日 ≥ 今天且非 canceled，含卡位，會比團體頁多';
    rec_(res, sid, 'C2_valid_n', useGas ? (G.window || perValid) : perValid, useGas ? q.act : g.vn, useGas ? g.vn : '', vNote, useGas ? gbLatest : scanMax, '', 'C 團體有效筆數');
    rec_(res, sid, 'C2_valid_people', useGas ? (G.window || perValid) : perValid, useGas ? q.ppl : g.vp, useGas ? g.vp : '', vNote, useGas ? gbLatest : scanMax, '', 'C 團體有效人數');
    if (G.state !== '成功') {
      ['C2_no_dessert', 'C2_unpaid', 'C2_urgent'].forEach(function (code, n) {
        rec_(res, sid, code, '', G.state, '', G.state === '未設定' ? '指令碼屬性 GB_PASSCODE 尚未設定' : (G.state === '通關碼錯誤' ? 'GB_PASSCODE 不正確' : '團體 GAS 讀取失敗'), gbLatest, '', ['C 團體未選甜點', 'C 團體未付足訂金', 'C 團體需立即處理'][n]);
      });
      return;
    }
    var ltU = (sid !== ALL_ROW && q.urg >= rule.GB_URGENT_RED) ? '🔴' : '', ltN = (sid !== ALL_ROW && q.nd >= rule.GB_NODESSERT_ORANGE) ? '🟠' : '';
    rec_(res, sid, 'C2_no_dessert', G.window || '', q.nd, '', '有效團體中還沒選甜點的筆數（＝團體頁「還沒選甜點」）', gbLatest, ltN, 'C 團體未選甜點');
    rec_(res, sid, 'C2_unpaid', G.window || '', q.unpN, r0_(q.unpNT), '有效團體中已付 < 應付訂金的筆數（比較值＝差額 NT$；合計＝團體頁「應收未收」）', gbLatest, '', 'C 團體未付足訂金');
    rec_(res, sid, 'C2_urgent', G.window || '', q.urg, '', '🚨 需立即處理的項目數（＝團體頁「需立即處理」）', gbLatest, ltU, 'C 團體需立即處理');
    if (ltU) alert_(res, '🔴', sid, map.name[sid] + ' 團體訂位有 ' + q.urg + ' 項需要立即處理（例如已取消卻仍掛已付款、逾期未付訂金、7 天內到店還沒選甜點），請客服打開訂位分析的「團體訂位」頁處理', q.urg + ' 項', '≥' + rule.GB_URGENT_RED + ' 項 🔴');
    if (ltN) alert_(res, '🟠', sid, map.name[sid] + ' 有 ' + q.nd + ' 筆團體訂位還沒選甜點，請提醒客人選好，採購才來得及備料', q.nd + ' 筆', '≥' + rule.GB_NODESSERT_ORANGE + ' 筆 🟠');
  });
  if (G.check) setMetaLatest_('C', '團體 GAS fn=group', null, (G.last_done || '') + '｜' + G.check);
  Object.keys(unknown).forEach(function (k) { alert_(res, '🟡', ALL_ROW, '團體訂位表（gb_bookings）出現對照表沒有的店名「' + k + '」，這筆沒有算進任何一店；請在 dim_store_map 的「訂位」欄補上', '', ''); });
  if (G.unknown) G.unknown.forEach(function (k) { alert_(res, '🟡', ALL_ROW, '團體 GAS 出現對照表沒有的店名「' + k + '」，這筆沒有算進任何一店；請在 dim_store_map 的「訂位」欄補上', '', ''); });
}

/**
 * 團體 GAS fn=group：只取各店計數，回傳內容（含客人姓名電話）不寫入任何地方。
 * 各店口徑對照團體頁（dashboard-reservation.html md5 78812639…）：
 *   有效團體＝rows 中 active 且 is_group（＝L2310 團體總表 groups 的條件；全公司＝summary.active_groups，L2279）
 *   未選甜點＝有效團體中 menu 為空（＝summary.no_dessert，L2282；團體總表 L2380 顯示「未選」）
 *   未付足訂金＝有效團體中 deposit > paid 且非 exempt（差額合計＝summary.unpaid「應收未收」，L2280）
 *   需立即處理＝所有列 issues 中 sev＝'high' 的項目數（＝summary.high，L2276 pending）
 * 全公司數字會與 summary 逐項核對，不一致寫進 meta。
 */
function readGroupCounts_(map) {
  var key = PropertiesService.getScriptProperties().getProperty(PROP_GB_PASS);
  var label = '團體 GAS fn=group', item = 'GAS ' + SRC.GB_API.replace(/^https:\/\/script\.google\.com\/macros\/s\/(.{10}).*$/, '$1') + '…?fn=group';
  if (!key) { logMeta_('C', label, item, 0, '', '', '警告：未設定通關碼', '指令碼屬性 GB_PASSCODE 尚未設定；團體三欄寫「未設定」'); return { state: '未設定' }; }
  var clean = function (m) { return String(m || '').split(key).join('***').replace(/key=[^&\s]*/g, 'key=***'); };
  var delays = [5, 10], t0 = Date.now(), lastErr = '', d = null;
  for (var i = 0; i <= delays.length; i++) {
    timeGuard_(label);
    var tA = Date.now(), trStep = label + (i ? '（第 ' + (i + 1) + ' 次）' : '');
    trace_('C', trStep, '開始');
    try {
      var resp;
      try { resp = UrlFetchApp.fetch(SRC.GB_API + '?fn=group&key=' + encodeURIComponent(key), { muteHttpExceptions: true, followRedirects: true }); }
      finally { trace_('C', trStep, '結束', (Date.now() - tA) / 1000); }
      if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
      var t = String(resp.getContentText('UTF-8')).trim(), m = t.match(/^[A-Za-z_$][\w$.]*\s*\(/);
      if (m) t = t.slice(m[0].length).replace(/\)\s*;?\s*$/, '');
      d = JSON.parse(t); t = null;
      if (d && d.ok === false && d.error === 'passcode') { logMeta_('C', label, item, (Date.now() - t0) / 1000, '', '', '警告：通關碼錯誤', 'GB_PASSCODE 不正確；團體三欄寫「通關碼錯誤」'); return { state: '通關碼錯誤' }; }
      if (!d || !d.ok) throw new Error('團體 GAS 回應異常：' + ((d && (d.msg || d.error)) || '無內容'));
      break;
    } catch (e) {
      if (isDefer_(e)) throw e;
      lastErr = clean(e && e.message || e); d = null;
      if (i < delays.length) { timeGuard_(label + '（重試前）'); Utilities.sleep(delays[i] * 1000); }
    }
  }
  if (!d) { logMeta_('C', label, item, (Date.now() - t0) / 1000, '', '', '失敗', lastErr); return { state: '讀取失敗', err: lastErr }; }
  var by = {}, unknown = {}, z = function () { return { act: 0, ppl: 0, nd: 0, unpN: 0, unpNT: 0, urg: 0 }; };
  by[ALL_ROW] = z();
  (d.rows || []).forEach(function (x) {
    var sid = map.byCol['訂位'][String(x.store || '').trim()];
    if (!sid) { unknown[String(x.store || '')] = 1; return; }
    var b = by[sid] || (by[sid] = z()), A = by[ALL_ROW];
    var urg = (x.issues || []).filter(function (it) { return it && it.sev === 'high'; }).length;
    b.urg += urg; A.urg += urg;
    if (x.active && x.is_group) {
      var ppl = Number(x.size) || 0, dep = Number(x.deposit) || 0, paid = Number(x.paid) || 0;
      b.act++; A.act++; b.ppl += ppl; A.ppl += ppl;
      if (!(x.menu && x.menu.length)) { b.nd++; A.nd++; }
      if (dep > paid && !x.exempt) { b.unpN++; A.unpN++; b.unpNT += dep - paid; A.unpNT += dep - paid; }
    }
  });
  var s = d.summary || {}, A = by[ALL_ROW], diff = [];
  [['有效團體', A.act, s.active_groups], ['人數', A.ppl, s.people], ['未選甜點', A.nd, s.no_dessert], ['需立即處理', A.urg, s.high], ['應收未收', A.unpNT, s.unpaid]].forEach(function (c) { if (c[2] !== undefined && Number(c[1]) !== Number(c[2])) diff.push(c[0] + ' 各店加總 ' + c[1] + '≠頁面 ' + c[2]); });
  var check = diff.length ? '⚠ 與團體頁摘要不一致：' + diff.join('；') : '與團體頁摘要一致（有效 ' + A.act + '／未選甜點 ' + A.nd + '／需立即處理 ' + A.urg + '／應收未收 ' + A.unpNT + '）';
  logMeta_('C', label, item, (Date.now() - t0) / 1000, (d.rows || []).length, d.last_done || '', diff.length ? '成功（摘要不一致）' : '成功', diff.length ? check : '');
  var win = d.window ? d.window.from + '～' + d.window.to : '';
  var out = { state: '成功', by: by, last_done: d.last_done || '', check: check, window: win, unknown: Object.keys(unknown) };
  d = null;
  return out;
}

// ============================================================
// 7. D 評論（評論 GAS，不帶 callback 回純 JSON）
// ============================================================
function calcD_(res, rule, map) {
  var data = fetchJson_('D', '評論 GAS', SRC.REV_API);
  if (!data || !data.ok) throw new Error('評論端點回應異常');
  var tP = Date.now();
  // 欄位對應照 google-reviews.html initLoad（L444–L453）：0 門市 1 評論者 2 年月 3 星等 4 品牌 5 區域 6 標記 7 內容
  var all = (data.rows || []).map(function (c) {
    return { store: c[0] || '', author: c[1] || '', ym: c[2] || '', stars: (c[3] === '' || c[3] === null || c[3] === undefined) ? null : Number(c[3]), brand: c[4] || '', region: c[5] || '', tag: c[6] || '', text: c[7] || '' };
  }).filter(function (r) { return r.store; });
  data = null;
  logMeta_('D', 'D 步驟：整理評論資料', '', (Date.now() - tP) / 1000, all.length, '', '成功', '');   // d2 D2-4 分步耗時
  // 資料最新日：評論 Sheet C 欄（評論時間）最大值（唯讀）
  timeGuard_('評論 Sheet C 欄');
  testSlow_('D');
  var t0 = Date.now(), tS = Date.now();
  trace_('D', '評論 Sheet 開啟', '開始');
  var shR = SpreadsheetApp.openById(SRC.REV_SS).getSheetByName('工作表1'), lastR = shR.getLastRow();
  trace_('D', '評論 Sheet 開啟', '結束', (Date.now() - tS) / 1000);
  // d3 D3-3：評論依時間往下新增，只讀最後 500 列就找得到最新日（2026-09-25 已核對與整欄讀取相同）
  var fromR = Math.max(2, lastR - REV_TAIL_ROWS + 1);
  tS = Date.now(); trace_('D', '評論 Sheet C 欄（最後 ' + REV_TAIL_ROWS + ' 列）', '開始');
  var cc = shR.getRange(fromR, 3, Math.max(1, lastR - fromR + 1), 1).getValues(), latest = '';
  trace_('D', '評論 Sheet C 欄（最後 ' + REV_TAIL_ROWS + ' 列）', '結束', (Date.now() - tS) / 1000);
  cc.forEach(function (r) { var x = r[0], s = (x instanceof Date) ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd') : String(x || '').slice(0, 10); if (/^\d{4}-\d{2}-\d{2}$/.test(s) && s > latest) latest = s; });
  logMeta_('D', '評論 Sheet C 欄（評論時間，最後 ' + REV_TAIL_ROWS + ' 列）', SRC.REV_SS.slice(0, 8) + '…', (Date.now() - t0) / 1000, cc.length, latest, '成功', '');
  setMetaLatest_('D', '評論 GAS', all.length, latest);
  RUN.fresh.D = { latest: latest };
  var today = RUN.date, cur = today.slice(0, 7), prv = prevYmOf_(cur);
  var dim = daysInMonth_(cur), dayN = Number(today.slice(8, 10)), passed = dayN - 1, left = dim - dayN;
  var target = MONTHLY_REVIEW_TARGET, prog = target * passed / dim;
  var cnt = function (ym, storeName) {
    var rows = all.filter(function (r) { return r.ym === ym && (storeName === null || r.store === storeName); });
    var valid = rows.filter(isValidReview);
    return { total: rows.length, five: rows.filter(function (r) { return r.stars === 5; }).length, valid: valid.length,
             named: valid.filter(function (r) { return hasPersonName(validContent(r), r.store); }).length,
             low: rows.filter(function (r) { return r.stars !== null && r.stars <= 4; }).length };
  };
  var revName = {}; Object.keys(map.byCol['評論']).forEach(function (nm) { revName[map.byCol['評論'][nm]] = nm; });
  var tJ = Date.now();
  map.list.concat([ALL_ROW]).forEach(function (sid) {
    var nm = sid === ALL_ROW ? null : revName[sid];
    [[cur, '本月'], [prv, '上月']].forEach(function (pm) {
      var c = cnt(pm[0], nm);
      rec_(res, sid, 'D1_total', pm[0], c.total, '', '', latest, '', 'D 評論總則數（' + pm[1] + '）');
      rec_(res, sid, 'D1_five', pm[0], c.five, '', '', latest, '', 'D 5★（' + pm[1] + '）');
      rec_(res, sid, 'D1_valid', pm[0], c.valid, sid === ALL_ROW ? target * 12 : target, '每月目標', latest, '', 'D 有效評論（' + pm[1] + '）');
      rec_(res, sid, 'D1_named', pm[0], c.named, '', '有效評論中帶人名（hasPersonName）', latest, '', 'D 有效且帶人名（' + pm[1] + '）');
      rec_(res, sid, 'D1_low', pm[0], c.low, '', '1–4 星', latest, '', 'D 未達5星（' + pm[1] + '）');
      if (pm[0] === cur && sid !== ALL_ROW) {
        var behind = prog - c.valid, lt = '🟢';
        if (left < rule.REV_LAST_DAYS && c.valid < rule.REV_LAST_MIN) lt = '🔴';
        else if (behind >= rule.REV_BEHIND_ORANGE) lt = '🟠';
        rec_(res, sid, 'D2_target_to_date', cur, r1_(prog), passed + '/' + dim, '已過天數／當月天數', latest, '', 'D 本月進度目標');
        rec_(res, sid, 'D2_behind', cur, r1_(behind), c.valid, '有效評論', latest, lt, 'D 落後進度（則）');
        if (lt === '🔴') alert_(res, '🔴', sid, map.name[sid] + ' 月底前 ' + left + ' 天，本月有效評論只有 ' + c.valid + ' 則（目標 ' + target + '）', c.valid + ' 則', '月底前 ' + rule.REV_LAST_DAYS + ' 天 <' + rule.REV_LAST_MIN + ' 則');
        else if (lt === '🟠') alert_(res, '🟠', sid, map.name[sid] + ' 本月有效評論 ' + c.valid + ' 則，落後進度 ' + r1_(behind) + ' 則（到昨天應有 ' + r1_(prog) + ' 則）', c.valid + ' 則', '落後 ≥' + rule.REV_BEHIND_ORANGE + ' 則');
      }
    });
  });
  logMeta_('D', 'D 步驟：判定有效評論（isValidReview／hasPersonName）', '', (Date.now() - tJ) / 1000, res.recs.length, '', '成功', '');   // d2 D2-4
}

// ============================================================
// 8. E 自己人（2026 各店新增自己人 年表 ＋ TARGETS_2026 照抄）
// ============================================================
function calcE_(res, rule, map) {
  timeGuard_('自己人年表');
  testSlow_('E');
  var t0 = Date.now(), ss = SpreadsheetApp.openById(SRC.MEM_SS);
  var v = ss.getSheetByName('2026 各店新增自己人').getDataRange().getValues();
  // 轉成 gviz 形狀（表頭列不進 rows），讓照抄的 parseNewMembersSheet 原封不動可用
  var gv = { table: { rows: v.slice(1).map(function (row) { return { c: row.map(function (x) { return { v: (x === '' ? null : x) }; }) }; }) } };
  var kpi = ss.getSheetByName('agg_store_kpi_365').getDataRange().getValues(), kpiUpd = '';
  for (var i = 1; i < kpi.length; i++) { var u = kpi[i][5]; if (u) { kpiUpd = (u instanceof Date) ? Utilities.formatDate(u, TZ, 'yyyy-MM-dd') : String(u).replace(/\//g, '-'); break; } }
  logMeta_('E', '2026 各店新增自己人（年表）＋ agg_store_kpi_365 更新時間', SRC.MEM_SS.slice(0, 8) + '…', (Date.now() - t0) / 1000, v.length - 1, '年表無更新日欄；管線更新 ' + kpiUpd, '成功', '');
  var sheet2026 = parseNewMembersSheet(gv, 2026);
  updateElapsedMonths(new Date());
  var curIdx = CLOSED_IDX_2026 + 1;
  var line = Math.round(ELAPSED_MONTHS_2026 / 12 * 100);
  var memName = {}; Object.keys(map.byCol['自己人年表']).forEach(function (nm) { memName[map.byCol['自己人年表'][nm]] = nm; });
  var today = RUN.date, cur = today.slice(0, 7), prv = prevYmOf_(cur);
  var tot = { cur: 0, prv: 0, closed: 0, target: 0, need: 0 };
  var latest = '年表（管線 ' + kpiUpd + '）';
  map.list.forEach(function (sid) {
    var s = memName[sid], mm = sheet2026[s] || {};
    var c = curIdx >= 0 && curIdx <= 11 ? (mm['2026/' + curIdx] || 0) : '';
    var p = curIdx - 1 >= 0 ? (mm['2026/' + (curIdx - 1)] || 0) : '';
    var actual = sumClosed2026(mm), target = TARGETS_2026[s] || 0;
    var rate = target ? Math.round(actual / target * 100) : '';
    var need = Math.max(0, Math.round((TARGETS_2026[s] || 0) * (curIdx + 1) / 12 - (actual || 0)));
    var gap = rate === '' ? '' : line - rate, lt = gap === '' ? '' : (gap > rule.MEM_BEHIND_RED ? '🔴' : (gap > rule.MEM_BEHIND_ORANGE ? '🟠' : '🟢'));
    rec_(res, sid, 'E1_new', cur, c, '', '本月至今（年表當月欄＝月初～昨天）', latest, '', 'E 本月至今新增');
    rec_(res, sid, 'E1_new', prv, p, '', '', latest, '', 'E 上月新增');
    rec_(res, sid, 'E2_closed_sum', '2026-01～' + prv, actual, target, '年度目標', latest, '', 'E 已完結月累計');
    rec_(res, sid, 'E2_target', '2026', target, '', '', latest, '', 'E 年度目標');
    rec_(res, sid, 'E2_rate', '2026-01～' + prv, rate, line, '年度進度線%（已完結月數÷12）', latest, lt, 'E 累計達成率%');
    rec_(res, sid, 'E3_need', cur, need, '', '月底回到進度線還差（照頁面算法）', latest, '', 'E 本月需招');
    if (lt === '🟠' || lt === '🔴') alert_(res, lt, sid, map.name[sid] + ' 自己人累計達成率 ' + rate + '%，低於年度進度線 ' + line + '% 共 ' + gap + ' 個百分點；本月還需招 ' + need + ' 人', rate + '%', '低於進度線 ' + rule.MEM_BEHIND_ORANGE + ' 🟠／' + rule.MEM_BEHIND_RED + ' 🔴 百分點');
    tot.cur += Number(c) || 0; tot.prv += Number(p) || 0; tot.closed += actual; tot.target += target; tot.need += need;
  });
  rec_(res, ALL_ROW, 'E1_new', cur, tot.cur, '', '本月至今', latest, '', 'E 本月至今新增');
  rec_(res, ALL_ROW, 'E1_new', prv, tot.prv, '', '', latest, '', 'E 上月新增');
  rec_(res, ALL_ROW, 'E2_closed_sum', '2026-01～' + prv, tot.closed, TOTAL_TARGET_2026, '年度目標', latest, '', 'E 已完結月累計');
  rec_(res, ALL_ROW, 'E2_target', '2026', TOTAL_TARGET_2026, '', '', latest, '', 'E 年度目標');
  rec_(res, ALL_ROW, 'E2_rate', '2026-01～' + prv, Math.round(tot.closed / TOTAL_TARGET_2026 * 100), line, '年度進度線%', latest, '', 'E 累計達成率%');
  rec_(res, ALL_ROW, 'E3_need', cur, tot.need, '', '各店加總', latest, '', 'E 本月需招');

  // 地雷 12：年表停更偵測——本月新增（全公司）與前 N 個「不同快照日」都一樣 → 🟡
  var stale = false, prevVals = [];
  try {
    var hv = snapSS_().getSheetByName(TAB.HIST).getDataRange().getValues(), byDay = {};
    for (var h = 1; h < hv.length; h++) {
      var dstr = (hv[h][0] instanceof Date) ? Utilities.formatDate(hv[h][0], TZ, 'yyyy-MM-dd') : String(hv[h][0]);
      if (String(hv[h][1]) === ALL_ROW && hv[h][2] === 'E1_new' && String(hv[h][3]) === cur && dstr < today) byDay[dstr] = hv[h][4];
    }
    var days = Object.keys(byDay).sort().slice(-rule.MEM_STALE_DAYS);
    prevVals = days.map(function (d) { return d + '=' + byDay[d]; });
    stale = days.length >= rule.MEM_STALE_DAYS && days.every(function (d) { return Number(byDay[d]) === tot.cur; });
  } catch (e) {}
  // M4：停更警示改列在「0 資料」（由 freshAlerts_ 產生）
  RUN.fresh.E = { stale: stale || (Number(today.slice(8, 10)) > 3 && tot.cur === 0), cur: tot.cur, prevVals: prevVals };
}


// ============================================================
// 9. e1 儀表板讀取窗口（doGet，JSONP，只讀）
//   紅線：只讀；不寫 _trace／meta、不取 LockService；不回傳客人個資與通關碼。
//   例外（_bundle 交辦書 2026-09-26）：bundle 即時計算後可寫回隱藏分頁 _bundle（只寫這一頁；runAll 進行中不寫）。
//   例外（處理回報 2026-09-28；2026-10-06 改免密碼）：track_set 只 appendRow 到隱藏分頁 track_log。
//   action=ping｜bundle｜history&store=N&cat=A~E&days=30｜series&store=N&days=35｜track_list｜track_set｜live&store=N（2026-10-06 店長頁即時資料）｜fresh（2026-10-06 各儀表板最後更新時間）｜equip（2026-10-06 器具／模具紅燈）
// ============================================================
var WEB_VER = 'e6-2026-10-06.11';   // .11：器具紅燈的甜點資料時間 eq.at 改成 yyyy-MM-dd HH:mm:ss（原本是英文日期字串）；.10：器具／模具紅燈（live 多回 equip、新增 action=equip）；.9：action=fresh 各儀表板「最後更新」（首頁卡片與各頁狀態列讀）；.8：店長頁即時資料 action=live（訂位 7 天各時段人數＋團體 14 天訂金／甜點，暫存 5 分鐘）；處理回報免密碼（處理中只要名字）；.7：決策中心處理回報 track_list／track_set（寫 track_log）；.6：snap_history 數字不再被轉成日期（writeHistTyped_）、刪 zzBundleCacheOnlyClear；.5：首頁資料預先準備（_bundle）；.4：bundle alerts 加「狀況種類」欄；dim_rule 加 PAGE_STALE_RED_AFTER（.3：刪除 clearTestData）
var BUNDLE_CHUNK = 90000;             // 單一快取 key 上限 100KB → 超過 90KB 切塊
var BUNDLE_TTL = 21600;               // CacheService 最長 6 小時；runAll 寫完快照時另外主動清除

function doGet(e) {
  var p = (e && e.parameter) || {}, cb = String(p.callback || '');
  var out;
  try {
    var act = String(p.action || 'ping');
    if (act === 'ping') out = { ok: true, version: WEB_VER };
    else if (act === 'bundle') out = webBundle_();
    else if (act === 'history') out = webHistory_(p.store, p.cat, p.days);
    else if (act === 'series') out = webSeries_(p.store, p.days);
    else if (act === 'track_list') out = webTrackList_();
    else if (act === 'track_set') out = webTrackSet_(p);
    else if (act === 'live') out = webLive_(p.store);
    else if (act === 'fresh') out = webFresh_();
    else if (act === 'equip') out = webEquip_();
    else out = { ok: false, error: '不支援的 action' };
  } catch (err) { out = { ok: false, error: '讀取失敗：' + shortErr_(err && err.message || err) }; }
  var json = JSON.stringify(out);
  if (/^[A-Za-z_$][\w$.]{0,63}$/.test(cb)) return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// f. 決策中心處理回報（2026-09-28；2026-10-06 改免密碼）：action=track_list／track_set
//   - 紀錄寫在快照 Sheet 隱藏分頁 track_log（只新增，不改舊列）。
//   - 2026-10-06 經營者裁定（各儀表板調整 1006 #4）：「處理中」「已處理」都不用密碼，只要名字；
//     「處理中」可以不寫說明，「已處理」要寫一句做了什麼。原本比對權限中心 dim_auth monthly 列的 trackPwCheck_ 已刪除。
//   - 防亂寫：同一個名字 10 分鐘內最多 TRACK_WHO_LIMIT 次；全部人合計 10 分鐘內最多 TRACK_ALL_LIMIT 次。
//   - 不取 ScriptLock（runAll 在用）。任何錯誤只回 {ok:false,error}，不影響 bundle／history／series、runAll、早報。
// ============================================================
var TRACK_TAB = 'track_log';
var HDR_TRACK = ['寫入時間', 'key', 'store', 'cat', 'kind', 'light', 'desc', 'status', 'note', 'who', 'role', 'snap'];
var TRACK_ROLES = ['店長', '魔導師', '經營者'];
var TRACK_DAYS = 60;           // track_list 回傳範圍
var TRACK_NOTE_MAX = 120, TRACK_WHO_MAX = 20, TRACK_KEY_MAX = 100;
var TRACK_WHO_LIMIT = 30;      // 同一個名字 10 分鐘內最多寫入次數
var TRACK_ALL_LIMIT = 300;     // 全部人合計 10 分鐘內最多寫入次數（免密碼後的防亂寫）
var TRACK_WIN_SEC = 600;

function trackSheet_(create) {
  var ss = snapSS_(), sh = ss.getSheetByName(TRACK_TAB);
  if (!sh && create) {
    sh = ss.insertSheet(TRACK_TAB, ss.getSheets().length);
    sh.getRange(1, 1, sh.getMaxRows(), HDR_TRACK.length).setNumberFormat('@');   // 全部純文字：日期、店號不被自動轉型
    sh.getRange(1, 1, 1, HDR_TRACK.length).setValues([HDR_TRACK]);
    sh.hideSheet();
  }
  return sh;
}
/** 寫入值一律字串；開頭是 = + - @ 的加單引號，避免被當成公式 */
function trackCell_(x) {
  var s = String(x == null ? '' : x);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function trackLen_(s) { return Array.from ? Array.from(s).length : s.length; }

function webTrackList_() {
  var sh = trackSheet_(false);
  if (!sh || sh.getLastRow() < 2) return { ok: true, items: [] };
  var cut = new Date(Date.now() - TRACK_DAYS * 864e5).toISOString(), v = [], to = sh.getLastRow();
  while (to >= 2) {   // 從表底往上讀，讀到 60 天前就停
    var from = Math.max(2, to - 1999), blk = sh.getRange(from, 1, to - from + 1, HDR_TRACK.length).getDisplayValues();
    v = blk.concat(v); to = from - 1;
    if (String(blk[0][0]) < cut) break;
  }
  var pick = {};
  v.forEach(function (r) {
    var at = String(r[0]), k = String(r[1]);
    if (!k || !at || at < cut) return;
    if (!pick[k] || at >= pick[k].at) pick[k] = { key: k, status: String(r[7]), note: String(r[8]), who: String(r[9]), role: String(r[10]), at: at, store: String(r[2]), cat: String(r[3]) };
  });
  var items = Object.keys(pick).map(function (k) { return pick[k]; }).sort(function (a, b) { return a.at < b.at ? 1 : (a.at > b.at ? -1 : 0); });
  return { ok: true, items: items };
}

function webTrackSet_(p) {
  var bad = function (m) { return { ok: false, error: m }; };
  try {
    var s = function (x) { return String(x == null ? '' : x).trim(); };
    var status = s(p.status), note = s(p.note), who = s(p.who), role = s(p.role), store = s(p.store), key = s(p.key);
    if (status !== 'doing' && status !== 'done') return bad('狀態只能是「處理中」或「已處理」');
    if (!note && status === 'done') return bad('請寫一句做了什麼');
    if (trackLen_(note) > TRACK_NOTE_MAX) return bad('說明最多 ' + TRACK_NOTE_MAX + ' 字');
    if (/09\d{8}/.test(note.replace(/[\s\-－]/g, '')) || /[^\s@]+@[^\s@]+\.[^\s@]+/.test(note)) return bad('說明裡有電話或 Email，請拿掉客人個資');
    if (!who || trackLen_(who) > TRACK_WHO_MAX) return bad('名字請填 1～' + TRACK_WHO_MAX + ' 字');
    if (TRACK_ROLES.indexOf(role) < 0) return bad('角色只能是店長、魔導師或經營者');
    if (!/^([1-9]|1[0-2])$/.test(store)) return bad('店號要是 1～12');
    var kp = key.split('|');
    if (!key || key.length > TRACK_KEY_MAX || kp.length < 2 || kp[0] !== store || !kp.slice(1).join('|')) return bad('警示代碼格式不對（店號|狀況種類）');

    var cache = CacheService.getScriptCache(), ak = 'trk_all', ac = Number(cache.get(ak) || 0);
    if (ac >= TRACK_ALL_LIMIT) return bad('現在回報的人太多，請 10 分鐘後再試');
    var wk = 'trk_who_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, who, Utilities.Charset.UTF_8));
    var wc = Number(cache.get(wk) || 0);
    if (wc >= TRACK_WHO_LIMIT) return bad('同一個名字 10 分鐘內回報超過 ' + TRACK_WHO_LIMIT + ' 次，請稍後再試');


    var sh = trackSheet_(true), at = new Date().toISOString();
    if (sh.getMaxRows() - sh.getLastRow() < 20) {   // 預留列用完 → 補 1000 列並設純文字
      var mr = sh.getMaxRows(); sh.insertRowsAfter(mr, 1000); sh.getRange(mr + 1, 1, 1000, HDR_TRACK.length).setNumberFormat('@');
    }
    var row = [at, key, store, s(p.cat).slice(0, 10), s(p.kind).slice(0, 60), s(p.light).slice(0, 4), s(p.desc).slice(0, 120), status, note, who, role, s(p.snap).slice(0, 10)];
    sh.appendRow(row.map(trackCell_));
    cache.put(wk, String(wc + 1), TRACK_WIN_SEC);
    cache.put(ak, String(ac + 1), TRACK_WIN_SEC);
    return { ok: true, item: { key: key, status: status, note: note, who: who, role: role, at: at } };
  } catch (e) {
    return bad('回報失敗：' + shortErr_(e && e.message || e));
  }
}

// ============================================================
// g. 店長頁即時資料（2026-10-06 經營者「各儀表板調整 1006」#4／#5）：action=live&store=N
//   店長打開決策中心「店長」分頁時讀，數字跟著來源更新（不等每天 11:00 的快照）：
//   ① 訂位 GAS fn=future：今天起 LIVE_DAYS 天，每天各時段已訂人數（有效訂位＝狀態不是「已取消」；排除測試門市；人數＝party_size，與訂位頁同口徑）
//   ② 團體 GAS fn=group（通關碼讀指令碼屬性 GB_PASSCODE）：今天起 LIVE_DAYS 天的有效團體（日期、時段、類別、人數、訂金已付／應付、
//      甜點品名與數量、問題標題）、需立即處理的項目，以及全部未來團體的摘要（有效／未選甜點／未付足訂金／需立即處理，與快照 C2 同口徑）
//   - 只讀：不寫任何分頁、log、meta、_trace；不取 LockService
//   - 12 店一起算，結果存 CacheService 5 分鐘（LIVE_TTL；有一邊讀取失敗只存 1 分鐘）；暫存期間再打開直接用
//   - 不回傳客人姓名、電話、LINE UID、訂位編號；問題標題與甜點名稱經 webScrub_ 遮罩
//   - 每時段人數上限 LIVE_SLOT_CAP 照抄訂位儀表板 SLOT_CAP（v2.15）；訂位頁改上限時這裡要一起改。紅燈線 140% 由頁面判斷（回傳 redPct）
// ============================================================
var LIVE_TTL = 300;          // 暫存秒數
var LIVE_TTL_PART = 60;      // 有一邊讀取失敗時的暫存秒數
var LIVE_DAYS = 14;          // 今天起幾天（店長頁：7 天行事曆＋8～14 天團體）
var LIVE_SLOT_RED = 140;     // 同時段人數 ≥ 每時段上限 × 140% → 紅燈（經營者 1006 指定）
var LIVE_SLOT_CAP = { '1': 14, '2': 20, '3': 24, '4': 13, '5': 16, '6': 19, '7': 12, '8': 16, '9': 12, '10': 18, '11': 14, '12': 12 };
var LIVE_KEY = 'live|v3';   // v3：eq.at 改成 yyyy-MM-dd HH:mm:ss；v2：多了 equip

function webLive_(store) {
  var sid = String(store == null ? '' : store).trim();
  if (!/^([1-9]|1[0-2])$/.test(sid)) return { ok: false, error: '店號要是 1～12' };
  var c = CacheService.getScriptCache(), t0 = Date.now(), all = liveCacheGet_(c), src = 'cache';
  if (!all) {
    all = liveBuild_(); src = 'build';
    if (all.fut.ok || all.gb.ok) liveCachePut_(c, all, (all.fut.ok && all.gb.ok) ? LIVE_TTL : LIVE_TTL_PART);
  }
  return { ok: true, version: WEB_VER, builtAt: all.builtAt, today: all.today, fut: all.fut, gb: all.gb, eq: all.eq || { ok: false, error: '尚未計算' },
    cap: LIVE_SLOT_CAP[sid] || null, redPct: LIVE_SLOT_RED,
    store: (all.stores || {})[sid] || { days: {}, groups: [], urgent: [], gsum: null, equip: [] },
    cache: { src: src, sec: Math.round((Date.now() - t0) / 100) / 10 } };
}
function liveCacheGet_(c) {
  try {
    var n = Number(c.get(LIVE_KEY + '|n')) || 0; if (!n) return null;
    var ks = []; for (var i = 0; i < n; i++) ks.push(LIVE_KEY + '|' + i);
    var got = c.getAll(ks), s = '';
    for (var j = 0; j < n; j++) { if (got[ks[j]] == null) return null; s += got[ks[j]]; }
    return JSON.parse(s);
  } catch (e) { return null; }
}
function liveCachePut_(c, obj, ttl) {
  var js = JSON.stringify(obj), m = {}, parts = Math.ceil(js.length / BUNDLE_CHUNK);
  for (var q = 0; q < parts; q++) m[LIVE_KEY + '|' + q] = js.slice(q * BUNDLE_CHUNK, (q + 1) * BUNDLE_CHUNK);
  m[LIVE_KEY + '|n'] = String(parts);
  try { c.putAll(m, ttl); } catch (e) {}
}
function liveParse_(resp) {
  var code = resp.getResponseCode(); if (code !== 200) throw new Error('HTTP ' + code);
  var t = String(resp.getContentText('UTF-8')).trim(), m = t.match(/^[A-Za-z_$][\w$.]*\s*\(/);
  if (m) t = t.slice(m[0].length).replace(/\)\s*;?\s*$/, '');
  return JSON.parse(t);
}
function liveTitle_(i) { return webScrub_(String((i && (i.title || i.code)) || '')).slice(0, 80); }
function liveClean_(e, key) { var s = shortErr_(e && e.message || e); if (key) s = s.split(key).join('***'); return s.replace(/key=[^&\s]*/g, 'key=***'); }
function liveBuild_() {
  var now = new Date(), today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), last = addDays_(today, LIVE_DAYS - 1);
  var map = readStoreMap_(snapSS_()), byName = map.byCol['訂位'] || {};
  var out = { builtAt: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm'), today: today, fut: { ok: false }, gb: { ok: false }, stores: {} };
  map.list.forEach(function (id) { out.stores[id] = { days: {}, groups: [], urgent: [], gsum: null, equip: [] }; });
  // ③ 器具／模具紅燈（1006 #2）：讀訂位資料的甜點明細＋BOM＋採購主檔，跟①②分開，失敗只影響這一塊
  try {
    var E = eqBuild_(today, last, byName);
    out.eq = { ok: true, at: E.at, rows: E.rows, unmatched: E.unmatched, rules: EQ_RULES.map(function (r) { return { id: r.id, label: r.label, limit: r.limit, unit: r.unit }; }) };
    map.list.forEach(function (id) { out.stores[id].equip = E.stores[id] || []; });
  } catch (eE) { out.eq = { ok: false, error: '器具紅燈計算失敗：' + shortErr_(eE && eE.message || eE) }; }
  var key = PropertiesService.getScriptProperties().getProperty(PROP_GB_PASS) || '';
  var reqs = [{ url: SRC.RSV_API + '?fn=future', muteHttpExceptions: true, followRedirects: true }];
  if (key) reqs.push({ url: SRC.GB_API + '?fn=group&key=' + encodeURIComponent(key), muteHttpExceptions: true, followRedirects: true });
  var resps;
  try { resps = UrlFetchApp.fetchAll(reqs); }
  catch (e0) {
    var m0 = liveClean_(e0, key);
    out.fut = { ok: false, error: '訂位讀取失敗：' + m0 };
    out.gb = { ok: false, error: key ? '團體讀取失敗：' + m0 : '團體通關碼未設定' };
    return out;
  }
  // ① 訂位 future：今天～+13 天，各店每天 筆數 n／人數 p／各時段人數 s
  try {
    var F = liveParse_(resps[0]), C = F.cols || [], R = F.rows || [];
    var iD = C.indexOf('date'), iS = C.indexOf('store'), iT = C.indexOf('slot'), iP = C.indexOf('party_size'), iSt = C.indexOf('status');
    if (iD < 0 || iS < 0 || iT < 0 || iP < 0 || iSt < 0) throw new Error('未來訂位欄位對不上');
    var fmin = '', fmax = '', unk = {};
    R.forEach(function (r) {
      var d = dstr_(r[iD]); if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return;
      if (!fmin || d < fmin) fmin = d;
      if (d > fmax) fmax = d;
      if (d < today || d > last) return;
      var nm = String(r[iS] || '').trim();
      if (!nm || isTestStore(nm) || String(r[iSt]) === '已取消') return;
      var id = byName[nm]; if (!id) { unk[nm] = 1; return; }
      var day = out.stores[id].days[d] || (out.stores[id].days[d] = { n: 0, p: 0, s: {} });
      var pp = Number(r[iP]) || 0, t = String(r[iT] || '').trim() || '—';
      day.n++; day.p += pp; day.s[t] = (day.s[t] || 0) + pp;
    });
    out.fut = { ok: true, min: fmin, max: fmax, rows: R.length, at: String(F.built_at || F.cached_at || ''), unknown: Object.keys(unk).slice(0, 5) };
  } catch (e1) { out.fut = { ok: false, error: '訂位讀取失敗：' + liveClean_(e1, key) }; }
  // ② 團體 fn=group：今天～+13 天的有效團體明細（不含姓名電話）＋全部未來團體摘要（＝快照 C2 口徑）
  if (!key) { out.gb = { ok: false, error: '團體通關碼未設定' }; return out; }
  try {
    var G = liveParse_(resps[1]);
    if (G && G.ok === false && G.error === 'passcode') throw new Error('團體通關碼不正確');
    if (!G || !G.ok) throw new Error('團體資料回應異常');
    var gunk = {};
    (G.rows || []).forEach(function (x) {
      var id = byName[String(x.store || '').trim()];
      if (!id) { gunk[String(x.store || '')] = 1; return; }
      var S = out.stores[id], g = S.gsum || (S.gsum = { act: 0, ppl: 0, nd: 0, unpN: 0, urg: 0 });
      var iss = x.issues || [], hi = iss.filter(function (i) { return i && i.sev === 'high'; }), mid = iss.filter(function (i) { return i && i.sev === 'mid'; });
      var d = dstr_(x.date);
      g.urg += hi.length;
      if (hi.length) S.urgent.push({ d: d, t: String(x.slot || ''), size: Number(x.size) || 0, cat: String(x.cat || ''), active: !!x.active, titles: hi.map(liveTitle_) });
      if (!(x.active && x.is_group)) return;
      var dep = Number(x.deposit) || 0, paid = Number(x.paid) || 0, menu = x.menu || [];
      g.act++; g.ppl += Number(x.size) || 0;
      if (!menu.length) g.nd++;
      if (dep > paid && !x.exempt) g.unpN++;
      if (d < today || d > last) return;
      S.groups.push({ d: d, t: String(x.slot || ''), cat: String(x.cat || ''), size: Number(x.size) || 0, dep: dep, paid: paid, exempt: !!x.exempt,
        aw: !!(x.awaiting && !x.awaiting.overdue),
        menu: menu.map(function (mm) { return [webScrub_(String(mm && mm[0] || '')).slice(0, 40), Number(mm && mm[1]) || 0]; }),
        hi: hi.map(liveTitle_), mid: mid.map(liveTitle_) });
    });
    var srt = function (a, b) { var x = a.d + ' ' + a.t, y = b.d + ' ' + b.t; return x < y ? -1 : (x > y ? 1 : 0); };
    map.list.forEach(function (id) {
      var S = out.stores[id];
      S.groups.sort(srt); S.urgent.sort(srt); S.urgent = S.urgent.slice(0, 10);
      if (!S.gsum) S.gsum = { act: 0, ppl: 0, nd: 0, unpN: 0, urg: 0 };
    });
    out.gb = { ok: true, at: String(G.last_done || ''), window: G.window ? String(G.window.from || '') + '～' + String(G.window.to || '') : '', unknown: Object.keys(gunk).slice(0, 5) };
  } catch (e2) { out.gb = { ok: false, error: '團體讀取失敗：' + liveClean_(e2, key) }; }
  return out;
}
// ============================================================
// i. 器具／模具紅燈（2026-10-06 經營者「各儀表板調整 1006」#2）
//   同店、同日、同一時段（開始時間）：把散客＋團體選的甜點份數 × 每份 BOM 用量加總，算出每一種器具／模具同時要用幾個：
//     ① 同款模具（採購主檔品類別＝模具；不含烤盤、各種擠花嘴／花嘴）≥ 6 個 → 紅燈
//     ② 電磁爐、平底鍋 ≥ 3 台 → 紅燈
//     ③ 手持攪拌機、塔皮機 ≥ 5 台 → 紅燈
//   甜點明細：訂位資料 fact_future_menu（訂位 GAS 抓後台列表時順便記「食譜」欄的甜點與份數；散客團體都有；已取消不算；不含姓名電話）
//   甜點 → BOM：BOM 本「BOM表」（甜點名稱、食材/器具名稱、數量＝每份用量）。名稱比對：產品名稱對照表 → 原名 → 去掉結尾括號註記（限當日壽星、葷…）→ 雙語名取「｜」後面；
//              都對不到的列在 unmatched（不算進紅燈，畫面提醒去補 BOM）
//   器具 → 品類別：採購專用檔 dim_sku（品名，或 BOM別名，全形分號分隔）；主檔沒建、名稱有「模」的也當模具
//   結果放在 live（每家店 store.equip）與 action=equip（12 店一起，訂位儀表板用），跟 live 同一份暫存。
// ============================================================
var EQ_RULES = [
  { id: 'mold', label: '同款模具', limit: 6, unit: '個' },
  { id: 'heat', label: '電磁爐／平底鍋', names: ['電磁爐', '平底鍋'], limit: 3, unit: '台' },
  { id: 'mix', label: '手持攪拌機／塔皮機', names: ['手持攪拌機', '塔皮機'], limit: 5, unit: '台' }
];
var EQ_EXCL = /烤盤|擠花嘴|花嘴/;
var EQ_MENU_TAB = 'fact_future_menu';
function eqRuleOf_(item, cat) {
  for (var i = 1; i < EQ_RULES.length; i++) if (EQ_RULES[i].names.indexOf(item) >= 0) return EQ_RULES[i];
  if (cat === '模具' && !EQ_EXCL.test(item)) return EQ_RULES[0];
  return null;
}
function eqNorm_(s) { return String(s || '').replace(/[\s《》「」『』]/g, ''); }
function eqDecode_(s) {
  return String(s || '').replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return String.fromCharCode(Number(d)); }).replace(/&amp;/g, '&');
}
function eqMatch_(name, bom, bomN, nmap) {
  var n = eqDecode_(eqDecode_(name)).trim(), base = n, prev;
  do { prev = base; base = base.replace(/\s*[（(][^（）()]*[）)]\s*$/, ''); } while (base !== prev);
  var cs = [n, base];
  if (base.indexOf('｜') >= 0) cs.push(base.split('｜').pop().trim());
  for (var i = 0; i < cs.length; i++) {
    var c = cs[i], t = nmap[c];
    if (t) { if (bom[t]) return t; if (bomN[eqNorm_(t)]) return bomN[eqNorm_(t)]; }
    if (bom[c]) return c;
    if (bomN[eqNorm_(c)]) return bomN[eqNorm_(c)];
  }
  return null;
}
function eqBuild_(today, last, byName) {
  // 1) dim_sku：品名／BOM別名 → 品類別；產品名稱對照表：POS 名 → BOM 名（都在採購專用檔）
  var pss = SpreadsheetApp.openById(SRC.PUR_DASH), ds = pss.getSheetByName('dim_sku').getDataRange().getValues(), h = ds[0].map(function (x) { return String(x).trim(); });
  var iN = h.indexOf('品名'), iC = h.indexOf('品類別'), iA = h.indexOf('BOM別名'), catOf = {};
  if (iN < 0 || iC < 0) throw new Error('dim_sku 欄位對不上');
  for (var r = 1; r < ds.length; r++) {
    var nm = String(ds[r][iN] || '').trim(), ct = String(ds[r][iC] || '').trim();
    if (!nm) continue;
    catOf[nm] = ct;
    if (iA >= 0) String(ds[r][iA] || '').split(/[；;]/).forEach(function (a) { a = a.trim(); if (a && !catOf[a]) catOf[a] = ct; });
  }
  var nmap = {}, nsh = pss.getSheetByName('產品名稱對照表');
  if (nsh) nsh.getDataRange().getValues().slice(1).forEach(function (x) { var a = String(x[0] || '').trim(), b = String(x[1] || '').trim(); if (a && b) nmap[a] = b; });
  // 2) BOM表：甜點 → [[器具, 每份數量, 規則 id]]（只留三條規則用得到的）
  var bsh = SpreadsheetApp.openById(FRESH_SS.BOM).getSheetByName('BOM表');
  if (!bsh) throw new Error('找不到 BOM表');
  var bv = bsh.getRange(1, 1, bsh.getLastRow(), 3).getValues(), bom = {}, bomN = {};
  for (var i = 1; i < bv.length; i++) {
    var d = String(bv[i][0] || '').trim(), it = String(bv[i][1] || '').trim(), q = Number(bv[i][2]) || 0;
    if (!d) continue;
    if (!bom[d]) { bom[d] = []; bomN[eqNorm_(d)] = d; }
    var rule = it ? eqRuleOf_(it, catOf[it] || (/模/.test(it) ? '模具' : '')) : null;   // 主檔沒建的（例：小矽膠模、大矽膠模），名稱有「模」也算模具
    if (rule && q > 0) bom[d].push([it, q, rule.id]);
  }
  // 3) 甜點明細：今天～last，已取消不算；同店同日同時段加總
  var msh = SpreadsheetApp.openById(FRESH_SS.RSV).getSheetByName(EQ_MENU_TAB);
  var res = { at: '', rows: 0, stores: {}, unmatched: [] };
  if (!msh || msh.getLastRow() < 2) return res;
  var mv = msh.getRange(2, 1, msh.getLastRow() - 1, 8).getValues(), slots = {}, un = {}, ruleById = {};
  EQ_RULES.forEach(function (x) { ruleById[x.id] = x; });
  mv.forEach(function (row) {
    var d = dstr_(row[0]); if (!/^\d{4}-\d{2}-\d{2}$/.test(d) || d < today || d > last) return;
    if (String(row[5]).trim() === '已取消') return;
    var sid = byName[String(row[1] || '').trim()]; if (!sid) return;
    var fa = (row[7] instanceof Date) ? Utilities.formatDate(row[7], TZ, 'yyyy-MM-dd HH:mm:ss') : String(row[7] || '').trim();   // v14：試算表會把 fetched_at 自動轉成日期，原本 String() 變成英文字串、比大小也錯
    if (fa > res.at) res.at = fa;
    var t = (row[2] instanceof Date) ? Utilities.formatDate(row[2], TZ, 'HH:mm') : String(row[2] || '').trim();
    var menu = []; try { menu = JSON.parse(String(row[6] || '[]')) || []; } catch (e) { menu = []; }
    res.rows++;
    var k = sid + '|' + d + '|' + t, S = slots[k] || (slots[k] = { sid: sid, d: d, t: t, need: {}, from: {}, rule: {} });
    menu.forEach(function (m) {
      var qty = Number(m && m[1]) || 0; if (!qty) return;
      var bn = eqMatch_(String(m[0] || ''), bom, bomN, nmap);
      if (!bn) { var u = webScrub_(String(m[0] || '')).slice(0, 40); un[u] = (un[u] || 0) + qty; return; }
      bom[bn].forEach(function (b) {
        S.need[b[0]] = (S.need[b[0]] || 0) + qty * b[1];
        S.rule[b[0]] = b[2];
        S.from[b[0]] = S.from[b[0]] || {}; S.from[b[0]][bn] = (S.from[b[0]][bn] || 0) + qty;
      });
    });
  });
  Object.keys(slots).forEach(function (k) {
    var S = slots[k];
    Object.keys(S.need).forEach(function (it) {
      var rule = ruleById[S.rule[it]]; if (!rule) return;
      var need = Math.round(S.need[it] * 10) / 10;
      if (need < rule.limit) return;
      (res.stores[S.sid] = res.stores[S.sid] || []).push({ d: S.d, t: S.t, item: it, need: need, limit: rule.limit, unit: rule.unit, rule: rule.id, label: rule.label,
        from: Object.keys(S.from[it]).map(function (bn) { return [bn, S.from[it][bn]]; }).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 6) });
    });
  });
  Object.keys(res.stores).forEach(function (sid) { res.stores[sid].sort(function (a, b) { var x = a.d + ' ' + a.t, y = b.d + ' ' + b.t; return x < y ? -1 : (x > y ? 1 : (b.need - a.need)); }); });
  res.unmatched = Object.keys(un).map(function (n) { return [n, un[n]]; }).sort(function (a, b) { return b[1] - a[1]; }).slice(0, 15);
  return res;
}
/** action=equip：12 店的器具／模具紅燈（訂位儀表板用；與 live 同一份暫存） */
function webEquip_() {
  var c = CacheService.getScriptCache(), t0 = Date.now(), all = liveCacheGet_(c), src = 'cache';
  if (!all) { all = liveBuild_(); src = 'build'; if (all.fut.ok || all.gb.ok) liveCachePut_(c, all, (all.fut.ok && all.gb.ok) ? LIVE_TTL : LIVE_TTL_PART); }
  var st = {}; Object.keys(all.stores || {}).forEach(function (id) { st[id] = (all.stores[id] || {}).equip || []; });
  return { ok: true, version: WEB_VER, builtAt: all.builtAt, today: all.today, eq: all.eq || { ok: false, error: '尚未計算' }, stores: st, cache: { src: src, sec: Math.round((Date.now() - t0) / 100) / 10 } };
}

/** 驗收用（只讀）：在編輯器直接呼叫 action=live，印出回應大小、秒數、各店天數與團體數＋個資檢查 */
function zzLiveTest() {
  CacheService.getScriptCache().remove(LIVE_KEY + '|n');
  var t0 = Date.now(), o = doGet({ parameter: { action: 'live', store: '11', callback: 'cb' } }).getContent(), t1 = Date.now();
  var t2 = Date.now(), o2 = doGet({ parameter: { action: 'live', store: '11', callback: 'cb' } }).getContent(), t3 = Date.now();
  Logger.log('live 第 1 次（即時組）' + (t1 - t0) / 1000 + ' 秒｜' + o.length + ' 字元；第 2 次（暫存）' + (t3 - t2) / 1000 + ' 秒｜' + o2.length + ' 字元');
  Logger.log(o.slice(0, 600));
  Logger.log('個資檢查：Email ' + (o.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || []).length + ' 處｜手機 ' + (o.match(/(?:\+?886[- ]?|0)9\d{2}[- ]?\d{3}[- ]?\d{3}/g) || []).length + ' 處｜LINE UID ' + (o.match(/U[0-9a-f]{32}/g) || []).length + ' 處');
}

// ============================================================
// h. 各儀表板「最後更新」（2026-10-06 經營者「各儀表板調整 1006」#7：每個儀表板都要明確顯示最後更新時間）：action=fresh
//   一次回 11 個系統各自的資料時間（只讀各系統試算表的表尾一欄，或指令碼屬性），存 CacheService 10 分鐘。
//   首頁 hub 每張卡片、各儀表板頂部資料狀態列都讀這一支（key：reviews／pos／member／pnl／schedule／vendor／recipe／purchase／monthly／reservation／announce）。
//   每項回 {txt（顯示用一句話）, at（最後更新時間 yyyy-MM-dd HH:mm，沒有就空白）, to（資料到哪一天）, warn（是否比正常慢）}；讀不到該項回 {err}。
//   只讀；不回傳任何內容資料（只有日期時間）。
// ============================================================
var FRESH_TTL = 1800, FRESH_KEY = 'fresh|v2';   // 30 分鐘（第一次組要開 9 個試算表，約 20～60 秒）
var FRESH_SS = {
  BOM: '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao',      // BOM 本（POS資料）
  PNL: '1khrFp_AYp3mEsL02cTRrYu1LNiRTP68aowM5_ClsE_c',      // P&L 資料倉儲（fact_pnl）
  SCH: '19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM',      // 排班分析資料倉（fact_daily_sales、import_log）
  VENDOR: '1Pp0C7zLWWwhO7miAxDQmxfYmX0bbh6crOK7dk0a_Zyw',   // 廠商單價（fact_vendor_price）
  RSV: '13NI3vGV4MSsngeO_DecVYKrOzky-fXCsN9ActscJorQ',      // DIYBC 訂位資料（fact_reservations_future）
  ANN: '1GXyGp9Y79HDhvqe4ZJbuQQBmyWnmnLOmXVVTD-x8uYs'       // 公告及工作清單_資料庫（fact_announce）
};
function webFresh_() {
  var c = CacheService.getScriptCache(), hit = c.get(FRESH_KEY);
  if (hit) { try { var o = JSON.parse(hit); o.cache = 'cache'; return o; } catch (e) {} }
  var out = freshBuild_();
  try { c.put(FRESH_KEY, JSON.stringify(out), FRESH_TTL); } catch (e) {}
  out.cache = 'build';
  return out;
}
/** 日期／時間值 → {d:'yyyy-MM-dd', t:'HH:mm' 或 ''} */
function freshVal_(v) {
  if (v instanceof Date) { if (isNaN(v)) return null; return { d: Utilities.formatDate(v, TZ, 'yyyy-MM-dd'), t: Utilities.formatDate(v, TZ, 'HH:mm') }; }
  var s = String(v == null ? '' : v).trim(), m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})(?:[ T](\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return { d: m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2), t: m[4] ? ('0' + m[4]).slice(-2) + ':' + m[5] : '' };
}
/** 讀某分頁某欄的表尾 n 列，回最大的日期時間 {d,t}（字串比大小：d＋t） */
function freshTailMax_(ss, tab, col, n) {
  var sh = ss.getSheetByName(tab); if (!sh) throw new Error('找不到分頁 ' + tab);
  var last = sh.getLastRow(); if (last < 2) return null;
  var k = Math.min(n, last - 1), v = sh.getRange(last - k + 1, col, k, 1).getValues(), best = null;
  for (var i = 0; i < v.length; i++) { var x = freshVal_(v[i][0]); if (x && (!best || x.d + ' ' + x.t > best.d + ' ' + best.t)) best = x; }
  return best;
}
function freshMd_(d) { return d ? d.slice(5, 7) + '/' + d.slice(8, 10) : '—'; }
function freshBuild_() {
  var now = new Date(), today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), hm = Utilities.formatDate(now, TZ, 'HH:mm'), yest = addDays_(today, -1);
  var items = {}, run = function (key, fn) { try { items[key] = fn(); } catch (e) { items[key] = { err: shortErr_(e && e.message || e) }; } };
  var ssCache = {}, open = function (id) { return ssCache[id] || (ssCache[id] = SpreadsheetApp.openById(id)); };
  run('pos', function () {   // POS：BOM 本 POS資料 C 欄「建立日期」；每天約 05:00 匯入前一天
    var x = freshTailMax_(open(FRESH_SS.BOM), 'POS資料', 3, 3000);
    return { to: x && x.d, warn: !!(x && x.d < yest && hm >= '07:00'), txt: '資料到 ' + freshMd_(x && x.d) + '（每天約 05:00 自動匯入前一天）' };
  });
  run('reviews', function () {   // Google 評論：評論 Sheet C 欄「評論時間」最後 500 列
    var x = freshTailMax_(open(SRC.REV_SS), '工作表1', 3, 500);
    return { to: x && x.d, at: x ? x.d + ' ' + x.t : '', warn: !!(x && x.d < addDays_(today, -2)), txt: '最新評論 ' + freshMd_(x && x.d) + (x && x.t ? ' ' + x.t : '') + '（每天自動抓）' };
  });
  run('member', function () {   // 自己人：agg_store_kpi_365 F 欄「更新時間」（自己人儀表板「資料更新至」同一格；「自己人原始資料」分頁 5 月底後就沒再更新，不能用）
    var sh = open(SRC.MEM_SS).getSheetByName('agg_store_kpi_365'); if (!sh) throw new Error('找不到分頁 agg_store_kpi_365');
    var x = sh.getLastRow() >= 2 ? freshVal_(sh.getRange(2, 6).getValue()) : null;
    return { to: x && x.d, warn: !!(x && x.d < addDays_(today, -1) && hm >= '10:00'), txt: '會員資料更新到 ' + freshMd_(x && x.d) + '（每天 08:30 自動）' };
  });
  run('pnl', function () {   // P&L：fact_pnl A 欄 year_month（民國年月）、L 欄 load_date
    var sh = open(FRESH_SS.PNL).getSheetByName('fact_pnl'); if (!sh) throw new Error('找不到分頁 fact_pnl');
    var last = sh.getLastRow(); if (last < 2) return { txt: '還沒有損益資料' };
    var v = sh.getRange(2, 1, last - 1, 12).getValues(), ym = 0, ld = '';
    v.forEach(function (r) { var y = Number(r[0]) || 0; if (y > ym) ym = y; var x = freshVal_(r[11]); if (x && x.d > ld) ld = x.d; });
    var ad = ym ? (Math.floor(ym / 100) + 1911) + '-' + ('0' + (ym % 100)).slice(-2) : '';
    return { to: ad, at: ld, txt: '損益資料到 ' + (ad || '—') + '・最後匯入 ' + freshMd_(ld) + '（每月手動）' };
  });
  run('schedule', function () {   // 排班：fact_daily_sales A 欄日期（營收，每天自動）＋ import_log 最後一列（班表上傳時間）
    var ss = open(FRESH_SS.SCH), x = freshTailMax_(ss, 'fact_daily_sales', 1, 3000), u = null;
    try { u = freshTailMax_(ss, 'import_log', 1, 50); } catch (e) {}
    return { to: x && x.d, at: u ? u.d + ' ' + u.t : '', warn: !!(x && x.d < addDays_(today, -2)), txt: '營收到 ' + freshMd_(x && x.d) + '・班表最後上傳 ' + (u ? freshMd_(u.d) + ' ' + u.t : '—') };
  });
  run('vendor', function () {   // 原物料：fact_vendor_price A 欄季別、K 欄日期（表尾 3000 列）
    var sh = open(FRESH_SS.VENDOR).getSheetByName('fact_vendor_price'); if (!sh) throw new Error('找不到分頁 fact_vendor_price');
    var last = sh.getLastRow(); if (last < 2) return { txt: '還沒有單價資料' };
    var k = Math.min(3000, last - 1), v = sh.getRange(last - k + 1, 1, k, 11).getValues(), q = '', d = '';
    v.forEach(function (r) { var qq = String(r[0] || '').trim(); if (/^\d{4}Q[1-4]$/.test(qq) && qq > q) q = qq; var x = freshVal_(r[10]); if (x && x.d > d) d = x.d; });
    return { to: d, txt: '單價最新 ' + (q || '—') + '・最後一筆 ' + freshMd_(d) + '（每季手動）' };
  });
  run('recipe', function () { return { txt: '即時（開頁直接讀 BOM 本與採購主檔）' }; });
  run('purchase', function () {   // 採購：agg_purchase G 欄「重算時間」＋ agg_usage_day C 欄「日」
    var ss = open(SRC.PUR_DASH), sh = ss.getSheetByName('agg_purchase'), rb = null;
    if (sh && sh.getLastRow() >= 2) rb = freshVal_(sh.getRange(2, 7).getValue());
    var x = null; try { x = freshTailMax_(ss, 'agg_usage_day', 3, 3000); } catch (e) {}
    return { to: x && x.d, at: rb ? rb.d + (rb.t ? ' ' + rb.t : '') : '', warn: !!(rb && rb.d < today && hm >= '08:00'),
      txt: '建議量重算 ' + (rb ? freshMd_(rb.d) + (rb.t && rb.t !== '00:00' ? ' ' + rb.t : '') : '—') + '・用量資料到 ' + freshMd_(x && x.d) + '（每天 07:05 自動）' };
  });
  run('monthly', function () {   // 決策中心：本專案最後一次寫完快照
    var dn = jsonProp_(PROP_DONE), at = String(dn.at || ''), x = freshVal_(at);
    return { at: x ? x.d + ' ' + x.t : '', warn: !!(x && x.d < today && hm >= '11:30'), txt: '營運快照 ' + (x ? freshMd_(x.d) + ' ' + x.t : '—') + '（每天 11:00 自動）' };
  });
  run('reservation', function () {   // 訂位：future 表 U 欄 fetched_at（每筆抓取時間）
    var sh = open(FRESH_SS.RSV).getSheetByName('fact_reservations_future'); if (!sh) throw new Error('找不到分頁 fact_reservations_future');
    var last = sh.getLastRow(); if (last < 2) return { txt: '未來訂位表是空的', warn: true };
    var v = sh.getRange(2, 21, last - 1, 1).getValues(), best = null;
    v.forEach(function (r) { var x = freshVal_(r[0]); if (x && (!best || x.d + ' ' + x.t > best.d + ' ' + best.t)) best = x; });
    return { at: best ? best.d + ' ' + best.t : '', warn: !!(best && best.d < today && hm >= '11:00'), txt: '未來訂位 ' + (best ? freshMd_(best.d) + ' ' + best.t : '—') + ' 更新（白天約每 3 小時）' };
  });
  run('announce', function () {   // 公告：fact_announce M 欄 updated_at（最後一次發佈／修改）
    var x = freshTailMax_(open(FRESH_SS.ANN), 'fact_announce', 13, 2000);
    return { at: x ? x.d + ' ' + x.t : '', txt: '即時・最後一次發佈／修改 ' + (x ? freshMd_(x.d) + ' ' + x.t : '—') };
  });
  return { ok: true, version: WEB_VER, builtAt: Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm'), items: items };
}

/** dim_rule 唯讀版（readRules_ 會補列，讀取窗口不可寫入） */
function readRulesRO_(ss) {
  var v = ss.getSheetByName(TAB.RULE).getDataRange().getValues(), r = {}, rows = [];
  for (var i = 1; i < v.length; i++) { var k = String(v[i][0]).trim(); if (!k) continue; r[k] = RULE_TEXT[k] ? String(v[i][3]).trim() : Number(v[i][3]); rows.push([k, webScrub_(v[i][1]), webScrub_(v[i][2]), v[i][3] instanceof Date ? (k === 'PAGE_STALE_RED_AFTER' ? Utilities.formatDate(v[i][3], TZ, 'HH:mm') : dstr_(v[i][3])) : (typeof v[i][3] === 'string' ? webScrub_(v[i][3]) : v[i][3]), webScrub_(v[i][4])]); }   // 收件人 Email 也遮蔽（讀取窗口不可出現 Email）
  RULE_DEFAULTS.forEach(function (d) { if (!(d[0] in r) || (RULE_TEXT[d[0]] ? false : isNaN(r[d[0]]))) r[d[0]] = d[3]; });
  return { map: r, rows: rows };
}
/** 分頁讀成純值表（日期轉字串；字串做個資遮罩——與 briefScrub_ 同規則，通關碼只讀一次，避免每格讀一次指令碼屬性） */
var WEB_KEY_ = null;
function webScrub_(s) {
  if (WEB_KEY_ === null) WEB_KEY_ = PropertiesService.getScriptProperties().getProperty(PROP_GB_PASS) || '';
  var out = String(s).replace(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g, '［已遮蔽］')
    .replace(/(?:\+?886[- ]?|0)9\d{2}[- ]?\d{3}[- ]?\d{3}/g, '［已遮蔽］')
    .replace(/\b0\d{1,2}-\d{6,8}\b/g, '［已遮蔽］');
  if (WEB_KEY_.length >= 4) out = out.split(WEB_KEY_).join('［已遮蔽］');
  return out;
}
function webTable_(sh) {
  var v = sh.getDataRange().getValues();
  return v.map(function (row) { return row.map(function (x) { return x instanceof Date ? dstr_(x) : (typeof x === 'string' ? webScrub_(x) : x); }); });
}
function bundleKey_() { return 'bnd|' + WEB_VER + '|' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, PropertiesService.getScriptProperties().getProperty(PROP_DONE) || '')); }
function bundleCacheClear_() {
  try {
    var c = CacheService.getScriptCache(), k = bundleKey_(), n = Number(c.get(k + '|n')) || 0, ks = [k + '|n'];
    for (var i = 0; i < Math.max(n, 20); i++) ks.push(k + '|' + i);
    c.removeAll(ks);
  } catch (e) { Logger.log('讀取窗口快取清除失敗（不影響快照）：' + shortErr_(e && e.message || e)); }
}
function webBundle_() {
  var c = CacheService.getScriptCache(), k = bundleKey_(), t0 = Date.now(), data = null, cached = false, parts = 0, src = 'cache';
  var n = Number(c.get(k + '|n')) || 0;
  if (n > 0) {
    var ks = []; for (var i = 0; i < n; i++) ks.push(k + '|' + i);
    var got = c.getAll(ks), s = '';
    for (var j = 0; j < n; j++) { if (got[ks[j]] == null) { s = null; break; } s += got[ks[j]]; }
    if (s) { try { data = JSON.parse(s); cached = true; parts = n; } catch (e) { data = null; } }
  }
  if (!data) {   // _bundle 交辦書：快取過期 → 讀隱藏分頁 _bundle（讀取鍵＝目前快取鍵，也就是同一個 WEB_VER＋同一個快照完成紀錄才用），讀到就補回快取
    try {
      var sd = bundleSheetGet_(k);
      if (sd) { data = sd.data; src = 'sheet'; parts = bundleCachePut_(c, k, sd.js); data.sizeKB = Math.round(sd.js.length / 102.4) / 10; }
    } catch (eS) { data = null; }
  }
  if (!data) {
    src = 'build';
    try { data = bundleBuild_(); }
    catch (e1) { Utilities.sleep(2000); try { data = bundleBuild_(); } catch (e2) { return { ok: false, error: '快照更新中，請稍後重新整理' }; } }
    var js = JSON.stringify(data);
    parts = bundleCachePut_(c, k, js);
    try { if (!PropertiesService.getScriptProperties().getProperty(PROP_RUN_STATE)) bundleSheetPut_(k, data.snapAt, js, '讀取窗口即時計算'); } catch (eW) {}   // runAll 進行中不寫（避免寫進還沒收尾的中間狀態）
    data.sizeKB = Math.round(js.length / 102.4) / 10;
  }
  // 「快照是否為今天」每次即時判斷（不進快取）
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  data.ok = true; data.snapIsToday = data.snapDate === today; data.now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
  data.cache = { hit: cached, src: src, parts: parts, sec: Math.round((Date.now() - t0) / 100) / 10 };
  return data;
}

// ---- _bundle 交辦書 2026-09-26：首頁資料預先準備 ----
//   runAll／巡檢員補跑：早報寄出後用讀取窗口同一個函式 bundleBuild_ 產生 bundle → 存隱藏分頁 _bundle（照 _stage，每格 45,000 字切段）＋放進 CacheService。
//   讀取窗口：CacheService → _bundle（讀取鍵相同才用）→ 即時計算（算完寫回兩處）。
//   _bundle 每列：讀取鍵（＝快取鍵：WEB_VER＋快照完成紀錄雜湊）｜快照完成時間｜段號｜總段數｜內容｜寫入時間｜來源
var BUNDLE_TAB = '_bundle';
var BUNDLE_CELL = 45000;
var HDR_BUNDLE = ['讀取鍵', '快照完成時間', '段號', '總段數', '內容', '寫入時間', '來源'];
function bundleCachePut_(c, k, js) {
  var m = {}, parts = Math.ceil(js.length / BUNDLE_CHUNK);
  for (var q = 0; q < parts; q++) m[k + '|' + q] = js.slice(q * BUNDLE_CHUNK, (q + 1) * BUNDLE_CHUNK);
  m[k + '|n'] = String(parts);
  try { c.putAll(m, BUNDLE_TTL); } catch (e) {}
  return parts;
}
/** 讀 _bundle：讀取鍵相符、段數齊全才用；否則回 null（不寫入任何東西） */
function bundleSheetGet_(k) {
  var sh = snapSS_().getSheetByName(BUNDLE_TAB);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, HDR_BUNDLE.length).getValues(), rows = [];
  v.forEach(function (r) { if (String(r[0]) === k) rows.push(r); });
  if (!rows.length) return null;
  var n = Number(rows[0][3]); if (!(n > 0) || rows.length !== n) return null;
  rows.sort(function (a, b) { return Number(a[2]) - Number(b[2]); });
  for (var i = 0; i < n; i++) if (Number(rows[i][2]) !== i) return null;
  var js = rows.map(function (r) { return String(r[4]); }).join('');
  return { data: JSON.parse(js), js: js };
}
/** 寫 _bundle：只動 _bundle 這一頁（整頁換成這一份；沒有就建立並隱藏） */
function bundleSheetPut_(k, at, js, from) {
  var ss = snapSS_(), sh = ss.getSheetByName(BUNDLE_TAB);
  if (!sh) { sh = ss.insertSheet(BUNDLE_TAB); sh.hideSheet(); }
  var n = Math.ceil(js.length / BUNDLE_CELL), now = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), rows = [];
  for (var i = 0; i < n; i++) rows.push([k, String(at || ''), i, n, js.slice(i * BUNDLE_CELL, (i + 1) * BUNDLE_CELL), now, from]);
  sh.clear();
  writeTyped_(sh.getRange(1, 1, 1, HDR_BUNDLE.length), [HDR_BUNDLE]);
  writeTyped_(sh.getRange(2, 1, rows.length, HDR_BUNDLE.length), rows);
  return n;
}
/** runLoop_ 收尾呼叫（早報寄出、清狀態之後）：時間不夠就跳過；整段 try/catch，失敗只寫 meta，不影響快照／早報／M9／保險絲 */
function bundlePrebuild_(execStart) {
  var t0 = Date.now(), used = (t0 - execStart) / 1000, res = '', note = '', parts = '';
  try {
    if (used + RUN_FINAL_SEC > RUN_HARD_SEC) {
      res = '跳過'; note = '本段已用 ' + Math.round(used) + ' 秒（＋預留 ' + RUN_FINAL_SEC + ' 秒超過 ' + RUN_HARD_SEC + ' 秒）→ 首頁改由讀取窗口即時計算';
    } else {
      var data = bundleBuild_(), js = JSON.stringify(data), k = bundleKey_();
      bundleCachePut_(CacheService.getScriptCache(), k, js);
      parts = bundleSheetPut_(k, data.snapAt, js, 'runAll 預先準備');
      res = '完成'; note = Math.round(js.length / 102.4) / 10 + ' KB；對應快照完成時間 ' + data.snapAt;
    }
  } catch (e) { res = '失敗'; note = shortErr_(e && e.message || e) + '（首頁改由讀取窗口即時計算，不影響快照與早報）'; }
  try {
    var sec = Math.round((Date.now() - t0) / 100) / 10, shM = snapSS_().getSheetByName(TAB.META);
    writeTyped_(shM.getRange(shM.getLastRow() + 1, 1, 1, HDR.meta.length), [[(RUN && RUN.ts) || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), 'W', '首頁資料預先準備', '_bundle（讀取窗口 bundle）', sec, parts, '', res, note]]);
  } catch (e2) { Logger.log('首頁資料預先準備 meta 記錄失敗：' + shortErr_(e2 && e2.message || e2)); }
}

function bundleBuild_() {
  var ss = snapSS_(), shL = ss.getSheetByName(TAB.LATEST), shA = ss.getSheetByName(TAB.ALERTS);
  if (!shL || !shA) throw new Error('snap_latest 或 alerts 分頁暫時不存在');
  var R = readRulesRO_(ss), map = readStoreMap_(ss), snap = snapDoneInfo_(ss);
  var day = snap.date || Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  var b = buildBrief_(ss, R.map, map, day);   // 與當天早報同一個函式；「今天」＝快照執行日
  var latest = webTable_(shL), alerts = webTable_(shA), mapT = webTable_(ss.getSheetByName(TAB.MAP));
  // alerts 每列加「新／第 N 天」（與早報 D5 同規則：alerts_history 昨天以前連續出現幾天）
  var hsh = ss.getSheetByName(TAB.AHIST), hv = hsh && hsh.getLastRow() >= 2 ? hsh.getRange(2, 1, hsh.getLastRow() - 1, HDR.alerts_history.length).getValues() : [];
  var byDay = {};
  hv.forEach(function (r) { var d = dstr_(r[0]); if (!d || d >= day) return; (byDay[d] = byDay[d] || {})[String(r[7] || '')] = 1; });
  var yday = addDays_(day, -1), marksOn = !!byDay[yday];
  alerts[0] = alerts[0].concat(['連續天數', '新／第 N 天', '狀況種類']);   // e2 修：狀況種類＝briefKindOf_ 的 id（頁面用來判斷燈號是不是格子那個數字造成的）
  for (var i = 1; i < alerts.length; i++) {
    if (String(alerts[i][0]) === '') continue;
    var kind = briefKindOf_(alerts[i], map, R.map), key = briefKey_(alerts[i], kind), n = 1, d = yday;
    while (byDay[d] && byDay[d][key]) { n++; d = addDays_(d, -1); }
    alerts[i] = alerts[i].concat([marksOn ? n : '', marksOn ? (n === 1 ? '新' : '第 ' + n + ' 天') : '', kind.id]);
  }
  var counts = { '🔴': 0, '🟠': 0, '🟡': 0 };
  alerts.slice(1).forEach(function (r) { if (r[0] in counts) counts[r[0]]++; });
  var lh = latest[0], allRow = latest.filter(function (r) { return String(r[0]) === ALL_ROW; })[0] || [];
  var Lc = function (col) { var j = lh.indexOf(col); return j < 0 ? '' : allRow[j]; };
  var short = briefShortNames_(ss, map);
  return {
    version: WEB_VER,
    snapAt: snap.at, snapDate: snap.date, snapCatAt: snap.catAt, oldSnapMsg: oldSnapMsg_(snap),
    dataTo: Lc('A 本月到營收最後一天淨營收｜資料最新日') || Lc('A 本月 rev/h｜資料最新日'),
    fresh: jsonProp_(PROP_FRESH), counts: counts,
    brief: { day: day, subject: b.subject, text: b.text, lines: b.lines.map(function (l) { return { sec: l.sec, text: l.text }; }), stale: b.stale },
    latest: latest, alerts: alerts, rules: R.rows, storeMap: mapT,
    stores: map.list.map(function (sid) { return { id: sid, name: map.name[sid], short: short[sid] || map.name[sid], zone: map.zone[sid] }; })
  };
}
/** 讀 snap_history 最後幾天（從底部往上讀，不讀整張表） */
function webHistory_(store, cat, days) {
  store = String(store || ''); cat = String(cat || '').toUpperCase(); days = Math.min(400, Math.max(1, Number(days) || 30));
  if (!/^[A-E]$/.test(cat) || !store) return { ok: false, error: '參數錯誤（store、cat=A~E）' };
  var sh = snapSS_().getSheetByName(TAB.HIST), last = sh.getLastRow();
  var cut = addDays_(Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), -days), v = [], to = last;
  while (to >= 2) {
    var from = Math.max(2, to - 4999), blk = sh.getRange(from, 1, to - from + 1, 10).getValues();
    v = blk.concat(v); to = from - 1;
    if (dstr_(blk[0][0]) < cut) break;
  }
  var pick = {};   // 同一天多批次 → 取最後一批
  v.forEach(function (r) {
    var d = dstr_(r[0]); if (!d || d < cut) return;
    if (String(r[1]) !== store || String(r[2]).charAt(0) !== cat) return;
    var ts = String(r[9] instanceof Date ? Utilities.formatDate(r[9], TZ, 'yyyy-MM-dd HH:mm:ss') : r[9]);
    var b = pick[d] || (pick[d] = { ts: '', rows: [] });
    if (ts > b.ts) { b.ts = ts; b.rows = []; }
    if (ts === b.ts) b.rows.push([d, String(r[2]), r[3] instanceof Date ? dstr_(r[3]) : r[3], r[4], r[5] instanceof Date ? dstr_(r[5]) : r[5], String(r[6]), r[7] instanceof Date ? dstr_(r[7]) : r[7], String(r[8])]);
  });
  var rows = []; Object.keys(pick).sort().forEach(function (d) { rows = rows.concat(pick[d].rows); });
  return { ok: true, store: store, cat: cat, days: days, cols: ['快照日', '指標代碼', '期間', '數值', '比較值', '比較類型', '資料最新日', '燈號'], rows: rows };
}
function webSeries_(store, days) {
  store = Number(store); days = Math.min(400, Math.max(1, Number(days) || 35));
  if (!(store >= 1)) return { ok: false, error: '參數錯誤（store）' };
  var sh = snapSS_().getSheetByName(TAB.SERIES);
  if (!sh || sh.getLastRow() < 2) return { ok: true, store: store, days: days, cols: HDR.snap_series.slice(0, 5), rows: [] };
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  var maxD = v.reduce(function (m, r) { var d = dstr_(r[0]); return d > m ? d : m; }, '');
  var cut = addDays_(maxD, -(days - 1));   // 以資料最後一天往回算 N 天（不是以今天算）
  var rows = v.filter(function (r) { return Number(r[1]) === store && dstr_(r[0]) >= cut; }).map(function (r) { return [dstr_(r[0]), Number(r[1]), r[2], r[3], r[4]]; });
  return { ok: true, store: store, days: days, cols: HDR.snap_series.slice(0, 5), rows: rows };
}
/** e1 驗收用（只讀）：在編輯器裡直接呼叫讀取窗口，印出回應大小與秒數；另外量一次「不經快取」的組裝秒數 */
function zzWebTest() {
  var t0 = Date.now(), out = doGet({ parameter: { action: 'bundle', callback: 'cb' } }).getContent(), t1 = Date.now();
  Logger.log('bundle 回應 ' + out.length + ' 字元（' + Math.round(out.length / 102.4) / 10 + ' KB）｜' + (t1 - t0) / 1000 + ' 秒｜開頭：' + out.slice(0, 160));
  var key = PropertiesService.getScriptProperties().getProperty(PROP_GB_PASS) || '';
  Logger.log('個資檢查：Email ' + (out.match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || []).length + ' 處｜手機 ' + (out.match(/(?:\+?886[- ]?|0)9\d{2}[- ]?\d{3}[- ]?\d{3}/g) || []).length + ' 處｜市話 ' + (out.match(/\b0\d{1,2}-\d{6,8}\b/g) || []).length + ' 處｜通關碼 ' + (key && out.indexOf(key) >= 0 ? '有（！）' : '沒有'));
  var t2 = Date.now(), d = bundleBuild_(), js = JSON.stringify(d);
  Logger.log('不經快取組裝 ' + (Date.now() - t2) / 1000 + ' 秒｜JSON ' + Math.round(js.length / 102.4) / 10 + ' KB｜切塊 ' + Math.ceil(js.length / BUNDLE_CHUNK));
  ['history&store=8&cat=A&days=30', 'series&store=8&days=35', 'ping'].forEach(function (q) {
    var p = {}; q.split('&').forEach(function (kv, i) { var a = kv.split('='); if (i === 0 && a.length === 1) p.action = a[0]; else if (i === 0) { p.action = a[0]; } else p[a[0]] = a[1]; });
    p.action = q.split('&')[0];
    var t = Date.now(), o = doGet({ parameter: p }).getContent();
    Logger.log(q + '：' + o.length + ' 字元｜' + (Date.now() - t) / 1000 + ' 秒｜' + o.slice(0, 200));
  });
}
