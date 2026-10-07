/**
 * 團體訂位追蹤 group_booking.gs  v1.4（2026-10-07）
 *   v1.4（經營者 2026-10-07 裁定）：
 *         ①客服訂位＋備註有「卡位」兩字但沒寫時間（例「人力不足卡位」）＝客服配合門市狀況卡位、避免現場客人太多，不是真的客人
 *           → 不算團體，不列「未付訂金」「未選甜點」等任何問題；改列為「門市卡位」（hold.type＝'block'，儀表板預設收在「正常的卡位」）。
 *           寫了時間的「14:00卡位」照舊是保護團體的卡位；客服訂位但備註沒有「卡位」（例：替客人代訂的包館）照舊當真的團體檢查。
 *         ②一般訂位（散客）公司不規定先選甜點 → 未選甜點只檢查大組／包館／包場（輸出 need_dessert；摘要的「還沒選甜點」同口徑）。
 *   v1.3：團體甜點自動同步到採購系統（fact_booking＝採購系統「預約」需求）。
 *         每一輪抓完後，把「有效團體訂位 × 客人選的甜點」寫進採購系統，客人改甜點／改數量／取消都會每天跟著更新；
 *         甜點名稱用採購系統同一套比對規則（產品名稱對照表）轉成 BOM 名，對不到的保留後台名稱，由採購系統標紅請店長選擇；
 *         店長選過的名稱會被記住，之後同名甜點自動套用（優先於近似比對）。早上第一輪改為 06:00（採購系統 07:05 重算前完成）。
 *         採購系統試算表是公開可讀 → 同步備註只寫姓名＋電話末 4 碼；採購系統回「忙碌」自動等候重送；
 *         可用指令碼屬性 GB_SYNC_MIN_SIZE 只同步 N 人以上的場次（預設全部）。
 *   v1.3.1：寫入後的核對改到「下一次執行」才做（2026-09-22 08:53 線上實證：同一次執行內重讀採購系統試算表，
 *         若範圍大小沒變，Google 會給第一次讀的快取舊值 → 13 筆其實已取消成功，卻被誤判「讀回仍有效」而暫停）。
 *         新增的列只要求「存在且內容相符」（之後被店長改甜點而取消，不算寫入失敗）；取消的列要求「狀態＝取消」。
 *   v1.2：客服會用自己的會員帳號幫電話客人訂位（代訂帳號），電話配對會失效 →
 *         取消單掛錢時，另外找「同店、同日、同時段」的有效團體訂位當作接手的新單；
 *         未付訂金的團體若同時段有已取消、仍掛已付的訂位，提示「錢可能在那一筆」。
 *   v1.1：改單／轉單配對只認「同一位客人＋同一家店」；同一帳號 60 天內有 5 筆以上團體訂位（常見於測試或代訂帳號）
 *         改提示「先確認是不是測試資料」，不再武斷指出錢轉到哪一筆（實測：一個測試帳號 13 筆跨 5 店，被誤配成改單）。
 * ------------------------------------------------------------------
 * 用途：每天抓訂位後台「今天～60 天後」的團體訂位（8 人以上、或類別為大組／包館／包場）
 *       與客服卡位，找出：①卡位逾期／孤兒卡位 ②資料錯誤（取消單掛錢、訂金不足…）
 *       ③未選甜點 ④違反消費規則，提供給訂位儀表板「團體訂位追蹤」區塊，並每天寄提醒信。
 *
 * 架構（沿用訂位管線的教訓）：
 *   - 獨立 Apps Script 專案，綁在新試算表「DIYBC 團體訂位」，完全不碰既有訂位管線。
 *   - gbWorker_ 每 15 分鐘醒來：沒事 1 秒收工；每天 06:00、14:00 後各跑一輪完整抓取（每輪結束接著同步採購系統）。
 *   - 一輪分三段（抓列表 → 抓明細 → 定稿），每次最多跑 270 秒，做不完下一次接著做（游標存屬性）。
 *     絕不讓單次執行撞到 6 分鐘硬殺（硬殺不會寄信＝無聲死亡）。
 *   - 抓完才定稿，中途失敗不會留下半套資料。
 *   - 只讀不寫：只打「訂位列表頁」與「修改頁」的 GET。
 *     ⛔ 絕對禁止呼叫任何 /Cancel… 網址（後台的取消是一個連結，GET 一下就直接取消訂位）。
 *
 * 指令碼屬性（專案設定 → 指令碼屬性）：
 *   DIYBC_EMAIL      訂位後台帳號（跟「預約系統自動抓資料」專案同一組）
 *   DIYBC_PASSWORD   訂位後台密碼
 *   GB_PASSCODE      儀表板通關碼（自己訂，夥伴打開團體區塊時輸入一次）
 *   GB_NOTIFY        提醒信收件人（未設定時用 mydiybc@gmail.com）
 *   （以下為採購系統同步用，平常不用動）
 *   GB_SYNC          on／off／paused（由 zz_gbSyncOn／zz_gbSyncOff 設定；paused＝寫入後核對不符自動暫停）
 *   GB_SYNC_MIN_SIZE 只同步幾人以上的場次；不設＝所有團體都同步（例：填 20 ＝只同步 20 人以上的包館）
 *   GB_SYNC_FORCE    填 1 ＝下一次同步略過安全檢查（改門檻、大量取消時才用；用完自動清除）
 *   PUR_TOKEN        採購系統寫入 token（不設＝與採購系統網頁同一組）
 *
 * 手動可執行的函式（編輯器上方下拉選單）：
 *   gbStatus        唯讀：看目前狀態（刻意放在檔案第一個，誤按也安全）
 *   gbTestLogin     唯讀：測試登入後台＋抓 1 天列表
 *   gbRunNow        立刻跑一輪（跑完前可能要按 2～3 次，每次最多 4.5 分鐘）
 *   gbSyncDryRun    唯讀：列出「如果現在同步到採購系統」會新增／取消哪些預約（不寫入）
 *   zz_gbSyncOn     開啟採購系統同步，並立刻同步一次
 *   zz_gbSyncOff    關閉採購系統同步（已寫進採購系統的預約保持原樣）
 *   zz_gbSetup      安裝排程（只需執行一次）
 * ------------------------------------------------------------------
 */

// ⚠️ 第一個函式必須是唯讀的（編輯器有時會執行檔案第一個函式，而不是下拉選單選的那個）
function gbStatus() {
  const P = PropertiesService.getScriptProperties();
  const lines = [];
  lines.push('版本：' + GB.VERSION);
  lines.push('帳號設定：' + (P.getProperty('DIYBC_EMAIL') ? '✅' : '❌ 缺 DIYBC_EMAIL')
    + '／密碼：' + (P.getProperty('DIYBC_PASSWORD') ? '✅' : '❌ 缺 DIYBC_PASSWORD')
    + '／通關碼：' + (P.getProperty('GB_PASSCODE') ? '✅' : '❌ 缺 GB_PASSCODE'));
  lines.push('提醒信收件人：' + (P.getProperty('GB_NOTIFY') || GB.NOTIFY_DEFAULT));
  lines.push('最後完成一輪：' + (P.getProperty('GB_LAST_DONE') || '（尚未完成過）'));
  lines.push('今天已完成：' + (P.getProperty('GB_DONE_DATE') === gbToday_() ? (P.getProperty('GB_DONE_COUNT') || '0') : '0') + ' 輪');
  lines.push('上一輪統計：' + (P.getProperty('GB_LAST_STATS') || '—'));
  const cyc = gbGetCycle_(P);
  lines.push('進行中的一輪：' + (cyc ? JSON.stringify({ day: cyc.day, phase: cyc.phase, next: cyc.next, to: cyc.to, span: cyc.span, reqs: cyc.reqs, kept: cyc.kept }) : '無'));
  lines.push('最近錯誤：' + (P.getProperty('GB_LAST_ERROR') || '無'));
  lines.push('採購系統同步：' + ({ on: '✅ 開啟', paused: '⛔ 已自動暫停（核對不符，請通知 Chat）' }[P.getProperty('GB_SYNC')] || '⏸ 關閉（執行 zz_gbSyncOn 開啟）')
    + (P.getProperty('GB_SYNC_PENDING') ? '／有待同步資料' : '') + (P.getProperty('GB_SYNC_VERIFY') ? '／上次寫入待核對（下一次執行自動核對）' : '') + '／同步範圍：' + (Number(P.getProperty('GB_SYNC_MIN_SIZE')) ? P.getProperty('GB_SYNC_MIN_SIZE') + ' 人以上' : '所有團體') + '／上次：' + (P.getProperty('GB_SYNC_LAST') || '—'));
  const trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  lines.push('排程：' + (trig.length ? trig.join(', ') : '❌ 尚未安裝（請執行 zz_gbSetup）'));
  const sh = SpreadsheetApp.getActive().getSheetByName(GB.SHEET_BOOKINGS);
  lines.push('gb_bookings 列數：' + (sh ? Math.max(0, sh.getLastRow() - 1) : '（尚未建立）'));
  Logger.log(lines.join('\n'));
  return lines.join('\n');
}

// ============ 設定 ============
const GB = {
  VERSION: 'gb-v1.4-20261007',
  BASE: 'https://diybc.azurewebsites.net',
  TZ: 'Asia/Taipei',
  DAYS_AHEAD: 60,            // 抓今天～60 天後（未來訂位表只有 35 天，團體常更早訂）
  ROW_CAP: 900,              // 後台列表上限約 1000 列；超過 900 視為可能被截斷 → 視窗減半重抓
  SPAN_INIT: 2,              // 第一段視窗天數（近期每天 200～600 筆）
  SPAN_MAX: 30,
  BUDGET_MS: 270 * 1000,     // 單次執行最多 270 秒（硬殺是 360 秒）
  RUN_FROM_HOUR: 6,          // 06:00～21:59 執行（v1.3：採購系統 07:05 重算，早上第一輪要在那之前完成）
  RUN_TO_HOUR: 22,
  SECOND_RUN_HOUR: 14,       // 每天第二輪
  DETAIL_TTL_MS: 3 * 86400 * 1000,
  STALE_HOURS: 30,           // 超過 30 小時沒完成一輪 → 寄信
  CYCLE_MAX_HOURS: 20,       // 一輪卡住超過 20 小時 → 放棄重來
  NOTIFY_DEFAULT: 'mydiybc@gmail.com',
  DASH_URL: 'https://diybc-training.onrender.com/static/dashboard-reservation.html',
  SHEET_BOOKINGS: 'gb_bookings',
  SHEET_STAGE: 'gb_stage',
  SHEET_DETAIL: 'gb_detail_cache',
  SHEET_LOG: 'gb_log',
  // 規則常數
  HOLD_NAME: '客服訂位',
  GROUP_CAT_RE: /大組|包館|包場/,
  GROUP_MIN_SIZE: 8,          // 8 人（含）以上＝團體
  HOLD_DAYS: 3,               // 訂位後 3 天內要付訂金
  DESSERT_DAYS: 7,            // 到店前 7 天要選完甜點
  DESSERT_REMIND_DAYS: 14,    // 14 天內未選 → 提醒
  VERIFY_DAYS: 2,             // 「末五碼確認中」超過 2 天未核對 → 提醒
  ACCOUNT_MANY: 5,            // 同一帳號（電話）60 天內 ≥5 筆團體訂位 → 多半是客服代訂帳號或測試，不用電話配對
  // 後台列表表頭（改版就整批擋下，避免欄位錯位寫進錯資料）
  EXPECTED_TH: ['日期', '分店', '時段', '會員', '', '', '', '類別', '人數', '食譜', '陪同', '總計', '已付', '狀態', '出席', '目的', '自己人等級', '備註', '']
};

// gb_bookings / gb_stage 的欄位（全部以純文字存，避免 Sheets 把日期、電話、00000 轉型）
const GB_COLS = ['id', 'kind', 'date', 'slot', 'store', 'member', 'phone', 'category', 'size', 'recipes', 'companions',
  'total', 'paid', 'paid_online', 'paid_offline', 'five', 'ecpay_no', 'status', 'status_text', 'purpose', 'memo', 'menu',
  'modify', 'rtd_id', 'scanned_at'];
const GB_DETAIL_COLS = ['id', 'modify', 'fetched_at', 'deadline', 'deposit', 'min_price', 'prepay_time', 'modify_time', 'modify_user'];
const GB_OUT_COLS = GB_COLS.concat(['deadline', 'deposit', 'min_price', 'prepay_time', 'modify_time', 'modify_user']);

// ============ 手動工具 ============
function gbTestLogin() {
  CacheService.getScriptCache().remove('GB_COOKIE');
  const ck = gbLogin_();
  const d = gbToday_();
  const html = gbFetchPath_(ck, gbListPath_(d, d));
  const parsed = gbParseList_(html);
  const keep = parsed.rows.filter(gbKeep_);
  const msg = '✅ 登入成功。今天（' + d + '）列表共 ' + parsed.count + ' 筆，其中團體／卡位 ' + keep.length + ' 筆。';
  Logger.log(msg);
  return msg;
}

function gbRunNow() {
  const r = gbWorker_({ force: true });
  Logger.log('本次結果：' + r);
  Logger.log(gbStatus());
  return r;
}

function zz_gbSetup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'gbWorker_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('gbWorker_').timeBased().everyMinutes(15).create();
  const P = PropertiesService.getScriptProperties();
  if (!P.getProperty('GB_NOTIFY')) P.setProperty('GB_NOTIFY', GB.NOTIFY_DEFAULT);
  gbEnsureSheets_();
  const msg = '✅ 已安裝排程 gbWorker_（每 15 分鐘），並建立資料分頁。';
  Logger.log(msg);
  return msg;
}

// ============ 排程主程式 ============
function gbWorker_(opt) {
  opt = (opt && typeof opt === 'object' && !opt.triggerUid) ? opt : {};
  const P = PropertiesService.getScriptProperties();
  const now = new Date();
  const hour = Number(Utilities.formatDate(now, GB.TZ, 'H'));
  const today = gbToday_();
  if (!opt.force && (hour < GB.RUN_FROM_HOUR || hour >= GB.RUN_TO_HOUR)) return 'skip-hour';

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 'locked';
  const t0 = Date.now();
  let result = 'idle';
  try {
    try {
      gbStaleCheck_(P, today);
      let cyc = gbGetCycle_(P);
      if (cyc && (cyc.day !== today || (Date.now() - cyc.started) > GB.CYCLE_MAX_HOURS * 3600 * 1000)) {
        gbLog_('warn', '放棄未完成的一輪（' + cyc.day + ' ' + cyc.phase + '），重新開始');
        cyc = null;
      }
      if (!cyc && (opt.force || gbShouldStart_(P, today, hour))) {
        cyc = gbNewCycle_(today);
        gbStageClear_();
        gbSaveCycle_(P, cyc);
        gbLog_('info', '開始新的一輪：' + cyc.next + ' ～ ' + cyc.to);
      }
      if (cyc) {
        const ck = gbLogin_();
        if (cyc.phase === 'scan') gbScan_(P, cyc, ck, t0);
        if (cyc.phase === 'detail') gbDetail_(P, cyc, ck, t0);
        if (cyc.phase === 'final' && gbTimeLeft_(t0) > 20000) {
          gbFinalize_(P, cyc, today, t0);
          P.deleteProperty('GB_CYCLE');
          P.deleteProperty('GB_LAST_ERROR');
          result = 'done';
        } else {
          gbSaveCycle_(P, cyc);
          P.deleteProperty('GB_LAST_ERROR');
          result = cyc.phase + '（未完成，下一次接著做）';
        }
      }
    } catch (e) {
      const msg = String((e && e.message) || e);
      P.setProperty('GB_LAST_ERROR', gbNowStr_() + ' ' + msg);
      gbLog_('error', msg);
      if (/表頭|改版/.test(msg)) gbAlertOnce_(P, 'THEAD', today, '【團體訂位】後台列表改版，抓取已暫停', msg);
      if (/登入失敗|token/.test(msg)) gbAlertOnce_(P, 'LOGIN', today, '【團體訂位】登入訂位後台失敗', msg);
      result = 'error: ' + msg;
    }
    // v1.3：採購系統同步——每一輪完成後自動跑；時間不夠或中斷，15 分鐘後的下一次接著做（每次都重新比對，不會重複寫）
    if (P.getProperty('GB_SYNC') === 'on' && (P.getProperty('GB_SYNC_PENDING') || P.getProperty('GB_SYNC_VERIFY')) && gbTimeLeft_(t0) > 60000) {
      try {
        gbSyncRun_(t0, false);
      } catch (e2) {
        const m2 = String((e2 && e2.message) || e2);
        gbLog_('error', '採購系統同步失敗：' + m2);
        gbAlertOnce_(P, 'SYNC', today, '【團體訂位】同步到採購系統失敗', m2 + '\n\n15 分鐘後會自動重試；持續失敗請通知 Chat。');
      }
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

function gbShouldStart_(P, today, hour) {
  const doneDate = P.getProperty('GB_DONE_DATE');
  const doneCount = doneDate === today ? Number(P.getProperty('GB_DONE_COUNT') || 0) : 0;
  if (doneCount === 0) return true;
  if (doneCount === 1 && hour >= GB.SECOND_RUN_HOUR) return true;
  return false;
}

function gbNewCycle_(today) {
  return {
    day: today, phase: 'scan', started: Date.now(),
    next: today, to: gbAddDays_(today, GB.DAYS_AHEAD), span: GB.SPAN_INIT,
    reqs: 0, scanned: 0, kept: 0, truncated: 0, details: 0
  };
}

// ---- 第一段：抓列表 ----
function gbScan_(P, cyc, ck, t0) {
  while (cyc.next <= cyc.to) {
    if (gbTimeLeft_(t0) < 45000) return;
    const to = gbMinStr_(gbAddDays_(cyc.next, cyc.span - 1), cyc.to);
    const html = gbFetchPath_(ck, gbListPath_(cyc.next, to));
    const parsed = gbParseList_(html);
    cyc.reqs++;
    if (parsed.count >= GB.ROW_CAP && cyc.span > 1) {
      cyc.span = Math.max(1, Math.floor(cyc.span / 2));   // 可能被截斷 → 視窗減半重抓
      gbSaveCycle_(P, cyc);
      continue;
    }
    if (parsed.count >= GB.ROW_CAP) {
      cyc.truncated++;
      gbLog_('warn', cyc.next + ' 單日列表 ' + parsed.count + ' 筆，可能被後台截斷');
    }
    const keep = parsed.rows.filter(gbKeep_);
    gbStageAppend_(keep);
    cyc.scanned += parsed.count;
    cyc.kept += keep.length;
    cyc.next = gbAddDays_(to, 1);
    if (parsed.count < GB.ROW_CAP / 3) cyc.span = Math.min(cyc.span * 2, GB.SPAN_MAX);
    gbSaveCycle_(P, cyc);
  }
  cyc.phase = 'detail';
}

// ---- 第二段：抓修改頁（付款期限、應付訂金、低消、付款時間、最後修改）----
function gbDetail_(P, cyc, ck, t0) {
  const rows = gbDedupe_(gbStageRead_());
  const cache = gbDetailCacheRead_();
  const nowMs = Date.now();
  const need = rows.filter(function (r) {
    if (r.status === 'canceled' || r.date < cyc.day || !r.id || !r.kind) return false;
    if (!(gbIsGroupRaw_(r) || gbIsHoldName_(r))) return false;
    const c = cache[r.id];
    return !c || c.modify !== r.modify || (nowMs - Date.parse(c.fetched_at)) > GB.DETAIL_TTL_MS;
  });
  for (let i = 0; i < need.length; i++) {
    if (gbTimeLeft_(t0) < 30000) { gbDetailCacheWrite_(cache); return; }
    const r = need[i];
    const html = gbFetchPath_(ck, '/Reservations/' + r.kind + '/' + r.id);
    const d = gbParseEdit_(html);
    cache[r.id] = {
      id: r.id, modify: r.modify, fetched_at: new Date().toISOString(),
      deadline: d.deadline, deposit: d.deposit, min_price: d.min_price,
      prepay_time: d.prepay_time, modify_time: d.modify_time, modify_user: d.modify_user
    };
    cyc.details++;
  }
  gbDetailCacheWrite_(cache);
  cyc.phase = 'final';
}

// ---- 第三段：定稿（整份覆蓋 gb_bookings）＋每天第一輪寄提醒信 ----
function gbFinalize_(P, cyc, today, t0) {
  const rows = gbDedupe_(gbStageRead_());
  const cache = gbDetailCacheRead_();
  const out = rows.map(function (r) {
    const c = cache[r.id] || {};
    return GB_OUT_COLS.map(function (k) {
      if (k in r) return r[k];
      return c[k] == null ? '' : c[k];
    });
  });
  const sh = gbSheet_(GB.SHEET_BOOKINGS, GB_OUT_COLS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
  if (out.length) {
    const rg = sh.getRange(2, 1, out.length, GB_OUT_COLS.length);
    rg.setNumberFormat('@');
    rg.setValues(out.map(function (a) { return a.map(function (v) { return String(v == null ? '' : v); }); }));
  }
  // 明細快取只保留這一輪還看得到的訂位，避免越積越大
  const ids = {}; rows.forEach(function (r) { ids[r.id] = 1; });
  Object.keys(cache).forEach(function (k) { if (!ids[k]) delete cache[k]; });
  gbDetailCacheWrite_(cache);

  const doneDate = P.getProperty('GB_DONE_DATE');
  const count = (doneDate === today ? Number(P.getProperty('GB_DONE_COUNT') || 0) : 0) + 1;
  P.setProperty('GB_DONE_DATE', today);
  P.setProperty('GB_DONE_COUNT', String(count));
  P.setProperty('GB_LAST_DONE', gbNowStr_());
  P.setProperty('GB_SYNC_PENDING', gbNowStr_());   // v1.3：這一輪的新資料，交給採購系統同步
  const stats = {
    rows: out.length, scanned: cyc.scanned, reqs: cyc.reqs, details: cyc.details,
    truncated: cyc.truncated, minutes: Math.round((Date.now() - cyc.started) / 60000)
  };
  P.setProperty('GB_LAST_STATS', JSON.stringify(stats));
  gbLog_('info', '完成一輪：' + JSON.stringify(stats));

  if (P.getProperty('GB_MAIL_DATE') !== today) {
    P.setProperty('GB_MAIL_DATE', today);
    try { gbSendDaily_(today); } catch (e) { gbLog_('error', '寄信失敗：' + e.message); }
  }
}

// ============ 登入與抓取 ============
function gbLogin_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('GB_COOKIE');
  if (hit) return hit;
  const P = PropertiesService.getScriptProperties();
  const email = P.getProperty('DIYBC_EMAIL'), pwd = P.getProperty('DIYBC_PASSWORD');
  if (!email || !pwd) throw new Error('登入失敗：缺少指令碼屬性 DIYBC_EMAIL / DIYBC_PASSWORD');
  const url = GB.BASE + '/Identity/Account/Login';
  const jar = {};
  const r1 = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: false });
  gbJarPut_(jar, r1);
  const html = r1.getContentText();
  const m = html.match(/name="__RequestVerificationToken"[^>]*?value="([^"]+)"/) ||
            html.match(/value="([^"]+)"[^>]*?name="__RequestVerificationToken"/);
  if (!m) throw new Error('登入失敗：登入頁找不到驗證 token（後台登入頁可能改版）HTTP ' + r1.getResponseCode());
  const r2 = UrlFetchApp.fetch(url, {
    method: 'post', muteHttpExceptions: true, followRedirects: false,
    headers: { Cookie: gbJarStr_(jar) },
    payload: { 'Input.Email': email, 'Input.Password': pwd, 'Input.RememberMe': 'false', '__RequestVerificationToken': m[1] }
  });
  gbJarPut_(jar, r2);
  const ok = r2.getResponseCode() === 302 && Object.keys(jar).some(function (k) { return /Identity\.Application/.test(k); });
  if (!ok) throw new Error('登入失敗：HTTP ' + r2.getResponseCode() + '，請確認 DIYBC_EMAIL / DIYBC_PASSWORD');
  const ck = gbJarStr_(jar);
  cache.put('GB_COOKIE', ck, 20 * 60);
  return ck;
}

function gbJarPut_(jar, resp) {
  const h = resp.getAllHeaders() || {};
  let sc = h['Set-Cookie'] || h['set-cookie'];
  if (!sc) return;
  if (!Array.isArray(sc)) sc = [sc];
  sc.forEach(function (line) {
    const kv = String(line).split(';')[0];
    const i = kv.indexOf('=');
    if (i > 0) jar[kv.slice(0, i).trim()] = kv.slice(i + 1).trim();
  });
}
function gbJarStr_(jar) {
  return Object.keys(jar).map(function (k) { return k + '=' + jar[k]; }).join('; ');
}

function gbListPath_(from, to) {
  return '/Reservations?uid=&date=' + from + '&date2=' + to + '&storeId=&type=&status=&memo=&userName=';
}

// 只允許 GET 列表頁與修改頁；任何 Cancel 網址一律拒絕（後台的取消連結 GET 一下就生效）
function gbFetchPath_(ck, path) {
  if (/cancel/i.test(path)) throw new Error('安全閘門：禁止存取取消網址 ' + path);
  if (!/^\/Reservations(\?|\/Edit\/|\/EditHourlyAdmin\/)/.test(path)) throw new Error('安全閘門：不在允許清單的網址 ' + path);
  for (let attempt = 0; attempt < 2; attempt++) {
    const r = UrlFetchApp.fetch(GB.BASE + path, { muteHttpExceptions: true, followRedirects: false, headers: { Cookie: ck } });
    const code = r.getResponseCode();
    if (code === 200) {
      const html = r.getContentText();
      if (/name="Input\.Password"/.test(html)) {         // 被導回登入頁
        CacheService.getScriptCache().remove('GB_COOKIE');
        ck = gbLogin_();
        continue;
      }
      return html;
    }
    if (code === 302 || code === 401) {
      CacheService.getScriptCache().remove('GB_COOKIE');
      ck = gbLogin_();
      continue;
    }
    throw new Error('後台回應 HTTP ' + code + '（' + path.slice(0, 60) + '）' + (code === 503 ? '，後台暫時無法使用，下一次會自動重試' : ''));
  }
  throw new Error('登入失敗：重新登入後仍被導回登入頁');
}

// ============ 解析：列表頁 ============
function gbParseList_(html) {
  const thead = html.match(/<thead[\s\S]*?<\/thead>/i);
  if (!thead) throw new Error('抓到的不是訂位列表頁（找不到表頭）');
  const th = [];
  thead[0].replace(/<th\b[^>]*>([\s\S]*?)<\/th>/gi, function (_, inner) { th.push(gbText_(inner)); return _; });
  if (th.join('|') !== GB.EXPECTED_TH.join('|')) {
    throw new Error('後台列表表頭改版，已停止抓取以免欄位錯位。預期：' + GB.EXPECTED_TH.join('|') + '；實際：' + th.join('|'));
  }
  const rows = [];
  const trRe = /<tr\b([^>]*\bdata-date="\d{8}"[^>]*)>([\s\S]*?)<\/tr>/gi;
  let m, count = 0;
  const scannedAt = gbNowStr_();
  while ((m = trRe.exec(html))) {
    const at = gbAttrs_(m[1]);
    const tds = [];
    const tdRe = /<td\b([^>]*)>([\s\S]*?)<\/td>/gi;
    let t;
    while ((t = tdRe.exec(m[2]))) tds.push({ a: gbAttrs_(t[1]), h: t[2] });
    if (tds.length < 19) continue;
    count++;
    const act = tds[18].h;
    const lk = act.match(/\/Reservations\/(Edit|EditHourlyAdmin)\/([0-9a-fA-F-]{36})/);
    const recPop = (tds[9].h.match(/data-content="([^"]*)"/) || [])[1] || '';
    const payPop = (tds[12].h.match(/data-content="([^"]*)"/) || [])[1] || '';
    const payTxt = gbText_(payPop.replace(/<br\s*\/?>/gi, '\n'));
    const menu = gbParseMenu_(recPop);
    const dd = at['data-date'];
    const tm = at['data-time'] || '';
    const md = at['data-modify'] || '';
    rows.push({
      id: lk ? lk[2].toLowerCase() : '',
      kind: lk ? lk[1] : '',
      date: dd.slice(0, 4) + '-' + dd.slice(4, 6) + '-' + dd.slice(6, 8),
      slot: tm.length >= 3 ? (('0' + tm).slice(-4, -2) + ':' + tm.slice(-2)) : gbText_(tds[2].h),
      store: gbText_(tds[1].h),
      member: gbDecode_(tds[3].a['data-v'] || '') || gbText_(tds[3].h).split(' ')[0],
      phone: gbDecode_(tds[4].a['data-v'] || '').trim(),
      category: gbText_(tds[7].h),
      size: gbNum_(tds[8].h),
      recipes: gbNum_(tds[9].h),
      companions: gbNum_(tds[10].h),
      total: gbNum_(tds[11].h),
      paid: gbNum_(tds[12].h),
      paid_online: gbNum_((payTxt.match(/綠界付款[:：]\s*(\d+)/) || [])[1]),
      paid_offline: gbNum_((payTxt.match(/離線付款[:：]\s*(\d+)/) || [])[1]),
      five: ((payTxt.match(/末五碼[:：](\d*)/) || [])[1] || ''),
      ecpay_no: ((payTxt.match(/綠界交易編號[:：](\S*)/) || [])[1] || ''),
      status: at['data-status'] || '',
      status_text: gbText_(tds[13].h),
      purpose: gbText_(tds[15].h),
      memo: gbText_(tds[17].h),
      menu: menu.map(function (x) { return x[0] + '|' + x[1]; }).join(';'),
      modify: md.length === 8 ? md.slice(0, 4) + '-' + md.slice(4, 6) + '-' + md.slice(6, 8) : '',
      rtd_id: at['data-rtdid'] || '',
      scanned_at: scannedAt
    });
  }
  return { count: count, rows: rows };
}

function gbParseMenu_(pop) {
  if (!pop) return [];
  const out = [];
  gbDecode_(pop).replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').split('\n').forEach(function (line) {
    const s = line.replace(/\s+/g, ' ').trim();
    const m = s.match(/^(.*\S)\s+X\s+(\d+)$/);
    if (m) out.push([m[1].replace(/[|;]/g, '／'), Number(m[2])]);
  });
  return out;
}

// 團體（原始判定）：類別是大組／包館／包場，或 8 人以上
function gbIsGroupRaw_(r) { return GB.GROUP_CAT_RE.test(r.category) || Number(r.size) >= GB.GROUP_MIN_SIZE; }
function gbIsHoldName_(r) { return String(r.member).indexOf(GB.HOLD_NAME) >= 0; }
function gbKeep_(r) { return !/測試|^test/i.test(r.store) && (gbIsGroupRaw_(r) || gbIsHoldName_(r)); }

// ============ 解析：修改頁 ============
function gbParseEdit_(html) {
  const f = {};
  (html.match(/<input\b[^>]*>/gi) || []).forEach(function (tag) {
    const a = gbAttrs_(tag);
    if (a.name && !(a.name in f)) f[a.name] = gbDecode_(a.value || '');
  });
  if (!('CancelTime' in f) && !('ModifyTime' in f)) throw new Error('修改頁格式不符（找不到付款期限／修改時間欄位）');
  const dl = f.CancelTime || '';
  const pp = f.PrePayTime || '';
  return {
    deadline: /^000/.test(dl) ? '' : dl,
    deposit: f.Deposite === undefined ? '' : String(gbNum_(f.Deposite)),
    min_price: f.MinPrice === undefined ? '' : String(gbNum_(f.MinPrice)),
    prepay_time: /\/0001 /.test(pp) ? '' : pp,
    modify_time: f.ModifyTime || '',
    modify_user: f.ModifyUser || ''
  };
}

// ============ 報表：規則判定（儀表板與提醒信共用同一份）============
function gbBuildReport_(rows, today) {
  const EXEMPT_RE = /[【\[［]\s*免訂金\s*[】\]］]/;
  const HOLD_RE = /(\d{1,2})\s*[:：]\s*(\d{2})\s*卡位/;
  const BLOCK_RE = /卡位/;   // v1.4：客服訂位備註有「卡位」但沒寫時間（例「人力不足卡位」）＝門市卡位，不是真的客人
  const B = rows.filter(function (r) { return r.date >= today; }).map(function (r) {
    const x = {};
    Object.keys(r).forEach(function (k) { x[k] = r[k]; });
    ['size', 'recipes', 'companions', 'total', 'paid', 'paid_online', 'paid_offline'].forEach(function (k) { x[k] = Number(r[k]) || 0; });
    x.deposit = (r.deposit === '' || r.deposit == null) ? null : Number(r.deposit);
    x.min_price = (r.min_price === '' || r.min_price == null) ? null : Number(r.min_price);
    x.menu = String(r.menu || '').split(';').filter(String).map(function (s) { const p = s.split('|'); return [p[0], Number(p[1]) || 0]; });
    x.active = r.status !== 'canceled';
    x.exempt = EXEMPT_RE.test(String(r.memo || ''));
    const hm = String(r.memo || '').match(HOLD_RE);
    x.hold_anchor = (gbIsHoldName_(r) && hm) ? (('0' + hm[1]).slice(-2) + ':' + hm[2]) : '';
    x.hold_block = gbIsHoldName_(r) && !x.hold_anchor && BLOCK_RE.test(String(r.memo || ''));
    x.is_group = !x.hold_anchor && !x.hold_block && gbIsGroupRaw_(r);
    x.need_dessert = GB.GROUP_CAT_RE.test(String(r.category || ''));   // v1.4：一般訂位（散客）不規定先選甜點，只有大組／包館／包場要選
    x.cust_key = gbIsHoldName_(r) ? '' : (String(r.phone || '').replace(/\D/g, '').length >= 8 ? 'P' + String(r.phone).replace(/\D/g, '') : (r.member ? 'N' + r.member : ''));
    x.issues = [];
    x.hold = null; x.awaiting = null; x.changes = [];
    x.days_to = gbDiffDays_(today, r.date);
    x.modify_by = gbWho_(r.modify_user);
    return x;
  });
  const byId = {}; B.forEach(function (x) { byId[x.id] = x; });

  function awaitingOf(x) {
    const start = x.modify || today;
    const due = x.deadline || gbAddDays_(start, GB.HOLD_DAYS);
    const left = gbDiffDays_(today, due);
    return { start: start, due: due, days_left: left, overdue: left < 0, from_deadline: !!x.deadline };
  }
  function add(x, code, sev, title, fix) { x.issues.push({ code: code, sev: sev, title: title, fix: fix }); }
  function md(d) { return d ? Number(d.slice(5, 7)) + '/' + Number(d.slice(8, 10)) : ''; }
  function money(n) { return '$' + String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

  // 1) 卡位：備註「HH:MM卡位」的客服訂位 → 找它保護的那一場
  B.forEach(function (x) {
    if (!x.active || !x.hold_anchor) return;
    const cands = B.filter(function (y) { return y.store === x.store && y.date === x.date && y.slot === x.hold_anchor && !y.hold_anchor; });
    const act = cands.filter(function (y) { return y.active; });
    const anchor = act.filter(function (y) { return y.paid > 0 || y.status === 'prePaid' || y.exempt; })[0] || act[0] || null;
    if (!anchor) {
      const why = cands.length ? '它保護的 ' + x.hold_anchor + ' 團體已取消' : '找不到它要保護的 ' + x.hold_anchor + ' 團體';
      x.hold = { type: 'orphan', anchor: x.hold_anchor, anchor_id: cands.length ? cands[0].id : '', reason: why };
      add(x, 'hold_orphan', 'high', '孤兒卡位：' + why, '立刻取消這筆卡位，把時段釋放給其他客人');
    } else if (anchor.paid > 0 || anchor.status === 'prePaid' || anchor.exempt) {
      x.hold = { type: 'protect', anchor: x.hold_anchor, anchor_id: anchor.id, reason: '保護 ' + x.hold_anchor + ' 已付訂金的團體' };
    } else {
      const aw = awaitingOf(anchor);
      x.hold = { type: 'waiting', anchor: x.hold_anchor, anchor_id: anchor.id, reason: '它保護的 ' + x.hold_anchor + ' 團體還沒付訂金', awaiting: aw };
      if (aw.overdue) add(x, 'hold_overdue', 'high', '卡位逾期：主單 ' + x.hold_anchor + ' 已超過付款期限 ' + (-aw.days_left) + ' 天仍未付訂金',
        '聯絡客人確認；不訂了就把主單和這筆卡位一起取消');
    }
  });

  // 1b) v1.4 門市卡位：客服配合門市狀況卡位（不是真的客人）→ 不付訂金、不選甜點都不算問題，只列在「正常的卡位」
  B.forEach(function (x) {
    if (!x.active || !x.hold_block) return;
    const m = String(x.memo || '').replace(/\s+/g, ' ').trim();
    x.hold = { type: 'block', anchor: '', anchor_id: '', reason: '門市卡位（' + (m.length > 20 ? m.slice(0, 20) + '…' : m) + '）：客服配合門市狀況卡位，不是真的客人，不必付訂金、選甜點' };
  });

  // 2) 團體訂位的各項檢查
  function slotActive(x) {    // 同店、同日、同時段的有效團體訂位（不含卡位）
    return B.filter(function (y) { return y.active && y.is_group && y.id !== x.id && y.store === x.store && y.date === x.date && y.slot === x.slot; });
  }
  function slotCanceledPaid(x) {  // 同店、同日、同時段、已取消但仍掛已付的團體訂位
    return B.filter(function (y) { return !y.active && y.is_group && y.paid > 0 && y.store === x.store && y.date === x.date && y.slot === x.slot
      && !/已轉至|轉至|轉到|已退款|退款完成|測試/.test(String(y.memo || '')); });
  }
  const acctCount = {};
  B.forEach(function (y) { if (y.is_group && y.cust_key) acctCount[y.cust_key] = (acctCount[y.cust_key] || 0) + 1; });
  B.forEach(function (x) {
    if (!x.is_group) return;
    const acctN = x.cust_key ? (acctCount[x.cust_key] || 0) : 0;
    const manyAcct = acctN >= GB.ACCOUNT_MANY;
    if (!x.active) {
      if (x.paid > 0 && !/已轉至|轉至|轉到|已退款|退款完成/.test(String(x.memo || ''))) {
        const mine = B.filter(function (y) { return y.active && y.is_group && x.cust_key && y.cust_key === x.cust_key; });
        const same = manyAcct ? [] : mine.filter(function (y) { return y.store === x.store; });
        const other = manyAcct ? [] : mine.filter(function (y) { return y.store !== x.store; });
        const lbl = function (y) { return md(y.date) + ' ' + y.store.replace(/店$/, '') + ' ' + y.size + ' 人'; };
        const slotNew = slotActive(x);
        const slotLbl = function (y) { return md(y.date) + ' ' + y.slot + ' ' + y.store.replace(/店$/, '') + ' ' + (y.member === GB.HOLD_NAME ? GB.HOLD_NAME : y.member) + ' ' + y.size + ' 人（已付 ' + money(y.paid) + '）'; };
        x.pair_ids = (same.length ? same : slotNew).map(function (y) { return y.id; });
        let fix;
        if (/測試/.test(String(x.memo || ''))) fix = '備註寫「測試」，但已付金額還在 → 把已付改 0，避免灌水訂金與營收統計';
        else if (same.length) fix = '錢應已轉到同店 ' + same.map(lbl).join('、') + ' 那筆 → 這筆已付改 0（綠界付款改不了就免改），備註寫「訂金已轉至 ' + md(same[0].date) + ' 新單」';
        else if (slotNew.length) fix = '同一時段已有新的有效訂位：' + slotNew.map(slotLbl).join('、') + ' → 若是同一團改單：確認錢已轉到新單後，這筆已付改 0、備註寫「訂金已轉至新單」；若新單沒收到這筆錢，要把錢轉過去或確認退款';
        else if (manyAcct) fix = '這個帳號 60 天內有 ' + acctN + ' 筆團體訂位，應是客服代訂帳號 → 確認這筆的錢轉到客人哪一筆正式訂位：轉過去了就把這筆已付改 0、備註寫「訂金已轉至 M/D 新單」；找不到就確認是否退款';
        else if (other.length) fix = '同一位客人在其他店有有效訂位（' + other.map(lbl).join('、') + '）→ 確認錢是否轉過去：是的話這筆已付改 0、備註寫「訂金已轉至 M/D 新單」；不是就確認退款';
        else fix = '找不到這位客人的其他有效訂位 → 確認要退款，或這筆其實不該取消；處理完在備註寫明';
        add(x, 'cancel_paid', 'high', '已取消，卻還掛著已付 ' + money(x.paid), fix);
      }
      return;
    }
    // 有效團體：改單紀錄只認「同一位客人＋同一家店」（跨店、或測試／代訂帳號不自動配對）
    if (!manyAcct) B.forEach(function (y) {
      if (!y.active && y.is_group && x.cust_key && y.cust_key === x.cust_key && y.store === x.store && y.id !== x.id) x.changes.push(y.id);
    });
    if (x.paid <= 0 && !x.exempt && (x.status === 'open' || x.status === '')) {
      x.awaiting = awaitingOf(x);
      const orphanMoney = slotCanceledPaid(x);
      const moneyHint = orphanMoney.length ? '。⚠️ 同一時段有已取消的訂位仍掛已付 ' + orphanMoney.map(function (y) { return money(y.paid); }).join('、') + '，這團的訂金可能在那一筆，請先核對再聯絡客人' : '';
      if (x.awaiting.overdue) add(x, 'unpaid_overdue', 'high',
        '未付訂金，已超過期限 ' + (-x.awaiting.days_left) + ' 天' + (x.awaiting.from_deadline ? '' : '（後台沒填付款期限，以最後修改日起算 ' + GB.HOLD_DAYS + ' 天）'),
        '聯絡客人：今天內付款，否則取消訂位釋放時段；若是公司同意的免訂金專案，備註加上「【免訂金】」' + moneyHint);
      else if (orphanMoney.length) add(x, 'money_elsewhere', 'mid',
        '還沒付訂金，但同一時段有已取消的訂位仍掛已付 ' + orphanMoney.map(function (y) { return money(y.paid); }).join('、'),
        '先核對那筆錢是不是這團的訂金：是的話轉到這筆並改為已付訂金，舊單已付改 0、備註寫「訂金已轉至新單」');
    }
    if (x.status === 'fiveDigi') {
      const days = gbDiffDays_(x.modify || today, today);
      if (days > GB.VERIFY_DAYS) add(x, 'verify_pending', 'mid', '客人回報已匯款（末五碼確認中）已 ' + days + ' 天未核對', '核對銀行入帳後改為已付訂金');
    }
    if (x.deposit != null) {
      if (x.deposit > 0 && x.paid > 0 && x.paid < x.deposit) {
        const late = x.deadline && x.deadline < today;
        add(x, 'underpaid', late ? 'high' : 'mid',
          '訂金不足：應付 ' + money(x.deposit) + '、已付 ' + money(x.paid) + '，差 ' + money(x.deposit - x.paid) + (late ? '（付款期限 ' + md(x.deadline) + ' 已過）' : '') + '，但狀態顯示「' + x.status_text + '」',
          '向客人補收差額；或調整應付訂金，並在備註寫明原因');
      }
      if (x.deposit > 0 && x.paid > x.deposit && !/多收|溢收|退款|退還|保留/.test(String(x.memo || ''))) {
        add(x, 'overpaid', 'mid', '多收訂金：應付 ' + money(x.deposit) + '、已付 ' + money(x.paid) + '，多 ' + money(x.paid - x.deposit),
          '退還差額，或在備註寫明多收原因（例：人數調降、轉單）');
      }
      if (x.deposit === 0 && !x.exempt) {
        add(x, 'deposit_missing', 'mid', '後台「應付訂金」是 0', '在修改頁補填應付訂金金額（公司同意免訂金的，備註加上「【免訂金】」）');
      }
    }
    if (x.paid_offline > 0 && x.five === '00000' && !/轉/.test(String(x.memo || ''))) {
      add(x, 'five_placeholder', 'mid', '離線付款 ' + money(x.paid_offline) + ' 的末五碼是 00000，看不出錢從哪來',
        '若是從舊訂位轉過來，備註寫「由 M/D 舊單轉入 ' + money(x.paid_offline) + '」；若是新匯款，填真實末五碼');
    }
    const menuQty = x.menu.reduce(function (s, m) { return s + m[1]; }, 0);
    if (!x.menu.length && x.need_dessert) {
      if (x.days_to <= GB.DESSERT_DAYS) add(x, 'no_dessert', 'high', '到店前 ' + GB.DESSERT_DAYS + ' 天內仍未選甜點（' + x.days_to + ' 天後到店）', '今天聯絡客人選甜點，店裡才能備料');
      else if (x.days_to <= GB.DESSERT_REMIND_DAYS) add(x, 'no_dessert_soon', 'info', x.days_to + ' 天後到店，尚未選甜點', '提醒客人在到店前 ' + GB.DESSERT_DAYS + ' 天選完');
    }
    if ((x.recipes > 0 || menuQty > 0) && x.recipes < Math.ceil(x.size / 2)) {
      add(x, 'below_rule', 'high', '違反消費規則：' + x.size + ' 人只做 ' + x.recipes + ' 份（2 人至少 1 份，最少 ' + Math.ceil(x.size / 2) + ' 份）',
        '聯絡客人補足份數，或修正人數');
    }
  });

  // 3) 摘要
  const groups = B.filter(function (x) { return x.is_group && x.active; });
  const all = B.filter(function (x) { return x.issues.length || x.hold || x.awaiting; });
  const sum = {
    active_groups: groups.length,
    people: groups.reduce(function (s, x) { return s + x.size; }, 0),
    paid: groups.reduce(function (s, x) { return s + x.paid; }, 0),
    deposit_due: groups.reduce(function (s, x) { return s + (x.deposit || 0); }, 0),
    unpaid: groups.reduce(function (s, x) { return s + (x.deposit ? Math.max(x.deposit - x.paid, 0) : 0); }, 0),
    detail_missing: groups.filter(function (x) { return x.deposit == null; }).length,
    high: B.reduce(function (s, x) { return s + x.issues.filter(function (i) { return i.sev === 'high'; }).length; }, 0),
    mid: B.reduce(function (s, x) { return s + x.issues.filter(function (i) { return i.sev === 'mid'; }).length; }, 0),
    no_dessert: groups.filter(function (x) { return !x.menu.length && x.need_dessert; }).length,
    no_dessert_7: groups.filter(function (x) { return !x.menu.length && x.need_dessert && x.days_to <= GB.DESSERT_DAYS; }).length,
    holds_protect: B.filter(function (x) { return x.hold && x.hold.type === 'protect'; }).length,
    holds_orphan: B.filter(function (x) { return x.hold && x.hold.type === 'orphan'; }).length,
    holds_waiting: B.filter(function (x) { return x.hold && x.hold.type === 'waiting'; }).length,
    holds_block: B.filter(function (x) { return x.hold && x.hold.type === 'block'; }).length,
    awaiting: B.filter(function (x) { return x.awaiting; }).length,
    awaiting_overdue: B.filter(function (x) { return (x.awaiting && x.awaiting.overdue) || (x.hold && x.hold.type === 'waiting' && x.hold.awaiting.overdue); }).length,
    flagged_rows: all.length
  };
  return { rows: B, summary: sum };
}

function gbWho_(u) {
  u = String(u || '');
  if (!u) return '';
  if (/^U[0-9a-f]{20,}$/i.test(u)) return '客人（LINE）';
  if (/@mydiybc\.com$/i.test(u) || /^mydiybc\d*@gmail\.com$/i.test(u)) return '夥伴：' + u.split('@')[0];
  return '非公司帳號';
}

// ============ 採購系統同步（v1.3）============
// 把「有效團體訂位 × 客人選的甜點」寫進採購系統的 fact_booking（＝採購系統「🎉 預約」頁、補貨公式裡的「預約需求」）。
// 寫入走採購系統自己的 API（跟店長在預約頁按「加入清單」／「取消」完全相同的路徑），不直接改採購系統的試算表。
//
// 每一筆同步寫入的預約，備註固定格式：〔團體〕姓名 電話 #訂位碼前8碼 ⟨POS 後台甜點名稱⟩
//   ‧ 用「訂位碼＋後台甜點名稱」辨認是哪一筆，所以客人改數量、改甜點、取消，隔天都會自動更新
//   ‧ 甜點欄：對得到 BOM → 寫 BOM 名；對不到 → 寫後台名稱（採購系統會標紅，請店長選）
//   ‧ 店長選過的甜點（甜點欄 ≠ 後台名稱且採購系統認得）一律保留，不會被隔天的同步改回去；
//     同一個後台名稱之後再出現在別的訂位，也自動套用店長的選擇
//
// 絕對不做的事（程式逐條檢查）：
//   ① 不動今天以前的預約  ② 不動店長自己登記的預約——唯一例外：同店、同日、備註電話與團體訂位相同
//      （＝店長手動登記的同一團），改由同步接手，避免重複備料  ③ 不改回店長選的甜點
//   ④ 資料看起來異常（團體忽然全部消失、要取消超過一半）→ 停止並寄信，不寫入  ⑤ 重跑不會重複寫
const PUR = {
  SS_ID: '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao',   // BOM 本（採購系統主資料庫，只讀）
  API: 'https://script.google.com/macros/s/AKfycbzOzNZ3L9cYrdFyHkopMQ_5HecXKxyo8_fL9prmSjxuOEmXqJT3aMsLHcyA4vjcO-oX/exec',
  TOKEN_DEFAULT: 'dbc-p-Xq7mKe2Ta9Rw',                    // 與採購系統網頁同一組（可用指令碼屬性 PUR_TOKEN 覆蓋）
  TAB_BOOKING: 'fact_booking',
  TAB_MAP: '產品名稱對照表',
  TAB_DESSERT: 'dim_dessert',
  STORE_NO: {
    '台中精明店': 1, '台中草悟道店': 2, '台北南京店': 3, '台北士林店': 4, '台南Focus店': 5, '新竹文化店': 6,
    '新北板橋店': 7, '新北新店店': 8, '桃園中壢店': 9, '桃園藝文店': 10, '台北遠百信義A13店': 11, '高雄SKM Park店': 12
  },
  POST_RETRY: 3,           // 採購系統回「忙碌」（有人正在存盤點）時，等幾秒重送的次數
  MAX_OPS: 300,            // 單次同步最多寫入／取消幾筆（超過視為異常，停止並寄信）
  MAX_DROP_RATIO: 0.5,     // 要「整筆拿掉」的同步預約超過現有的一半 → 停止並寄信
  MIN_DROP_CHECK: 10,      // 拿掉 10 筆以內不做比例檢查（取消一場大團體很正常）
  SOON_DAYS: 14            // 採購系統只算 14 天內的預約；提醒信只列這段期間對不到的甜點
};
function gbPurStoreNo_(name) {
  const k = String(name || '').replace(/\s+/g, '').toLowerCase();
  let hit = 0;
  Object.keys(PUR.STORE_NO).forEach(function (n) { if (n.replace(/\s+/g, '').toLowerCase() === k) hit = PUR.STORE_NO[n]; });
  return hit;
}
function gbPhoneTail_(p) { const d = String(p || '').replace(/\D/g, ''); return d.length >= 4 ? '尾' + d.slice(-4) : ''; }
const PUR_KEY_RE = /〔團體〕[\s\S]*#([0-9a-f]{12})\s*⟨POS\s*([\s\S]*?)⟩\s*$/;   // 訂位碼取前 12 碼（去掉 -），避免兩團撞號

function gbSyncDryRun() {
  const r = gbSyncRun_(Date.now(), true);
  Logger.log(r.report);
  return r.report;
}
function zz_gbSyncOn() {
  const P = PropertiesService.getScriptProperties();
  P.setProperty('GB_SYNC', 'on');
  P.setProperty('GB_SYNC_PENDING', gbNowStr_());
  const r = gbSyncRun_(Date.now(), false);
  Logger.log('✅ 已開啟採購系統同步（之後每一輪抓完會自動同步）\n' + r.report);
  return r.report;
}
function zz_gbSyncOff() {
  PropertiesService.getScriptProperties().setProperty('GB_SYNC', 'off');
  Logger.log('⏸ 已關閉採購系統同步；已寫進採購系統的預約保持原樣。');
}

// ---- 讀採購系統（只讀）----
function gbPurRead_() {
  const ss = SpreadsheetApp.openById(PUR.SS_ID);
  const tz = ss.getSpreadsheetTimeZone() || GB.TZ;
  const shB = ss.getSheetByName(PUR.TAB_BOOKING);
  if (!shB) throw new Error('採購系統找不到分頁 ' + PUR.TAB_BOOKING);
  const vb = shB.getDataRange().getValues();
  const hb = vb[0].map(function (h) { return String(h).trim(); });
  const ix = {};
  ['預約ID', '店號', '日期', '甜點', '數量', '備註', '狀態'].forEach(function (k) {
    ix[k] = hb.indexOf(k);
    if (ix[k] < 0) throw new Error('採購系統 fact_booking 欄位改了，找不到「' + k + '」，同步停止');
  });
  const bookings = vb.slice(1).filter(function (r) { return String(r[ix['預約ID']] || '').trim(); }).map(function (r) {
    const d = r[ix['日期']];
    return {
      id: String(r[ix['預約ID']]).trim(), store: String(r[ix['店號']]).trim(),
      date: (d instanceof Date) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : String(d || '').slice(0, 10),
      dessert: String(r[ix['甜點']] == null ? '' : r[ix['甜點']]).trim(), qty: Number(r[ix['數量']]) || 0,
      note: String(r[ix['備註']] == null ? '' : r[ix['備註']]), status: String(r[ix['狀態']] == null ? '' : r[ix['狀態']]).trim()
    };
  });
  const shM = ss.getSheetByName(PUR.TAB_MAP);
  const mapRows = shM ? shM.getRange(2, 1, Math.max(shM.getLastRow() - 1, 1), 2).getValues() : [];
  const shD = ss.getSheetByName(PUR.TAB_DESSERT);
  const desserts = shD ? shD.getRange(2, 1, Math.max(shD.getLastRow() - 1, 1), 1).getValues().map(function (r) { return String(r[0] == null ? '' : r[0]).trim(); }).filter(String) : [];
  if (!desserts.length) throw new Error('採購系統 dim_dessert 是空的，無法比對甜點名稱，同步停止');
  return { bookings: bookings, mapRows: mapRows, desserts: desserts };
}

// ---- 甜點名稱比對：逐字移植採購系統網頁的 bkNorm／bkUniq／bkResolve（兩邊判斷必須一致）----
function gbPurResolver_(mapRows, desserts) {
  const BKMAP = {};
  mapRows.forEach(function (r) {
    const pn = String(r[0] == null ? '' : r[0]).trim(); if (!pn) return;
    const cn = String(r[1] == null ? '' : r[1]).trim(); if (cn || !BKMAP[pn]) BKMAP[pn] = cn || pn;
  });
  const DSET = {}; desserts.forEach(function (d) { DSET[d] = 1; });
  function norm(s, lvl) {
    let t = String(s == null ? '' : s).replace(/[\s\u3000]+/g, '').replace(/（/g, '(').replace(/）/g, ')').replace(/／/g, '/').toLowerCase();
    if (lvl >= 2) t = t.replace(/【[^】]*】/g, '').replace(/\((葷|素|全素|蛋奶素|奶素|蛋素|限)\)$/, '').replace(/-限$/, '').replace(/限定$/, '');
    return t;
  }
  const set = {}; Object.keys(BKMAP).forEach(function (k) { set[k] = 1; }); desserts.forEach(function (d) { if (d && !set[d]) set[d] = 1; });
  const opts = Object.keys(set).sort();
  const ix = { l1: {}, l2: {}, all: opts };
  opts.forEach(function (n) { const k1 = norm(n, 1), k2 = norm(n, 2); (ix.l1[k1] = ix.l1[k1] || []).push(n); (ix.l2[k2] = ix.l2[k2] || []).push(n); });
  function uniq(arr) {
    if (!arr || !arr.length) return null; const byC = {}, order = [];
    arr.forEach(function (n) { const c = BKMAP[n] || n; if (!byC[c]) { byC[c] = n; order.push(c); } else if (BKMAP[n] && BKMAP[n] !== n) byC[c] = n; });
    return order.length === 1 ? byC[order[0]] : null;
  }
  function known(name) { name = String(name || '').trim(); if (!name) return false; if (BKMAP[name]) return true; return !!DSET[name]; }
  function resolve(name) {
    name = String(name || '').trim();
    if (!name) return { name: name, ok: false, approx: false };
    if (known(name)) return { name: name, ok: true, approx: false };
    const k1 = norm(name, 1), k2 = norm(name, 2); let hit;
    hit = uniq(ix.l1[k1]); if (hit) return { name: hit, ok: true, approx: hit !== name };
    hit = uniq(ix.l2[k2]); if (hit) return { name: hit, ok: true, approx: true };
    if (k2.length >= 3) {
      const cand = ix.all.filter(function (n) { const nk = norm(n, 2); return nk.length >= 3 && (nk.indexOf(k2) >= 0 || k2.indexOf(nk) >= 0); });
      hit = uniq(cand); if (hit) return { name: hit, ok: true, approx: true };
    }
    return { name: name, ok: false, approx: false };
  }
  // 採購系統重算時的展開規則：posToCanon[名稱] || 名稱，再查 BOM；這裡檢查「寫進去的名字重算時真的展得開」
  function expands(name) { return !!DSET[BKMAP[name] || name]; }
  function canon(name) { return BKMAP[name] || name; }
  return { resolve: resolve, known: known, expands: expands, canon: canon };
}

// ---- 算出「應該要有的預約」與要做的動作（純計算，不寫入）----
function gbSyncPlan_(groups, pur, today) {
  const R = gbPurResolver_(pur.mapRows, pur.desserts);
  const digits = function (s) { return String(s || '').replace(/\D/g, ''); };
  const isAuto = function (b) { return PUR_KEY_RE.test(b.note); };
  const keyOf = function (b) { const m = b.note.match(PUR_KEY_RE); return m ? m[1] + '|' + m[2] : ''; };
  const future = pur.bookings.filter(function (b) { return b.date >= today; });

  // 店長選過的甜點（含已取消的舊列，取最新）：後台名稱 → 店長選的名稱
  const learned = {};
  pur.bookings.filter(isAuto).forEach(function (b) {
    const m = b.note.match(PUR_KEY_RE); const pos = m[2];
    if (b.dessert && b.dessert !== pos && R.known(b.dessert)) learned[pos] = b.dessert;
  });

  const desired = [], unmatched = [], skipped = [], warn = [];
  groups.forEach(function (g) {
    const storeNo = gbPurStoreNo_(g.store);
    if (!storeNo) { skipped.push(g.date + ' ' + g.store + '（採購系統沒有這家店的店號）'); return; }
    // 採購系統的試算表「知道網址就能讀」→ 備註只放姓名＋電話末 4 碼，不放完整電話
    const who = [(g.member === GB.HOLD_NAME ? String(g.memo || '').slice(0, 12) : g.member), gbPhoneTail_(g.phone)].filter(String).join(' ');
    const qtyBy = {}, order = [];
    (g.menu || []).forEach(function (m) { const pos = String(m[0] || '').trim(); const q = Number(m[1]) || 0; if (!pos || q <= 0) return; if (!(pos in qtyBy)) { qtyBy[pos] = 0; order.push(pos); } qtyBy[pos] += q; });
    const rid = String(g.id || '').replace(/-/g, '').toLowerCase().slice(0, 12);
    order.forEach(function (pos) {
      // 順序：完全相符 → 店長選過的 → 近似比對 → 對不到（店長改過近似結果，代表近似是錯的，之後同名甜點以店長為準）
      const r = R.resolve(pos);
      let name = '', how = '', apx = '';
      if (r.ok) {
        const nm = R.expands(R.canon(r.name)) ? R.canon(r.name) : (R.expands(r.name) ? r.name : '');
        if (!nm) warn.push(pos + '：對照表有這個名稱，但指向的配方不在 BOM（' + R.canon(r.name) + '）');
        else if (!r.approx) { name = nm; how = 'exact'; }
        else apx = nm;
      }
      if (!name && learned[pos]) { name = learned[pos]; how = 'learned'; }
      if (!name && apx) { name = apx; how = 'approx'; }
      if (!name) { name = pos; how = 'miss'; }
      const row = {
        key: rid + '|' + pos, rid: rid, pos: pos, store: String(storeNo), date: g.date, qty: qtyBy[pos], dessert: name, how: how,
        note: '〔團體〕' + who + ' #' + rid + ' ⟨POS ' + pos + '⟩', phone: digits(g.phone), gid: g.id
      };
      desired.push(row);
      if (how === 'miss') unmatched.push({ date: g.date, store: g.store, who: who, pos: pos, qty: qtyBy[pos] });
    });
  });

  const existing = future.filter(function (b) { return b.status !== '取消' && isAuto(b); });
  const byKey = {}; existing.forEach(function (b) { const k = keyOf(b); (byKey[k] = byKey[k] || []).push(b); });
  const adds = [], cancels = [], takeovers = [];
  let kept = 0, drops = 0;
  const seen = {};
  desired.forEach(function (d) {
    seen[d.key] = 1;
    const ex = byKey[d.key] || [];
    if (!ex.length) { adds.push(d); return; }
    const e = ex[0];
    // 店長選過的甜點（≠ 後台名稱、且採購系統認得）一律保留
    const target = (e.dessert && e.dessert !== d.pos && R.known(e.dessert)) ? e.dessert : d.dessert;
    if (e.store !== d.store || e.date !== d.date || e.qty !== d.qty || e.dessert !== target || e.note !== d.note) {
      adds.push(Object.assign({}, d, { dessert: target, replaces: e.id }));
      cancels.push({ id: e.id, why: '更新（' + [e.date !== d.date ? '日期' : '', e.qty !== d.qty ? '數量 ' + e.qty + '→' + d.qty : '', e.dessert !== target ? '甜點' : '', e.store !== d.store ? '門市' : ''].filter(String).join('、') + '）', b: e });
    } else kept++;
    ex.slice(1).forEach(function (x) { cancels.push({ id: x.id, why: '重複的同步列', b: x }); });
  });
  existing.forEach(function (b) {
    if (!seen[keyOf(b)]) { cancels.push({ id: b.id, why: '客人已取消這筆訂位或拿掉這項甜點', b: b }); drops++; }
  });
  // 店長手動登記的同一團（同店、同日、備註裡有同一支電話）→ 由同步接手
  const bookingsSeen = {};
  desired.forEach(function (d) {
    const k = d.gid; if (bookingsSeen[k]) return; bookingsSeen[k] = 1;
    if (d.phone.length < 8) return;
    future.forEach(function (b) {
      if (b.status === '取消' || isAuto(b) || b.store !== d.store || b.date !== d.date) return;
      if (digits(b.note).indexOf(d.phone) >= 0) { takeovers.push({ id: b.id, why: '店長手動登記的同一團，改由同步接手', b: b }); }
    });
  });
  return { desired: desired, adds: adds, cancels: cancels, takeovers: takeovers, kept: kept, drops: drops, existingAuto: existing.length,
    unmatched: unmatched, skipped: skipped, warn: warn, learnedN: Object.keys(learned).length };
}

// ---- 執行同步 ----
function gbSyncRun_(t0, dry) {
  const P = PropertiesService.getScriptProperties();
  const today = gbToday_();
  const rep = gbBuildReport_(gbBookingsRead_(), today);
  // 同步門檻（指令碼屬性 GB_SYNC_MIN_SIZE，預設 0＝所有團體都同步；設 20 則只同步 20 人以上的場次）
  const minSize = Number(P.getProperty('GB_SYNC_MIN_SIZE')) || 0;
  const groups = rep.rows.filter(function (x) { return x.is_group && x.active && x.menu && x.menu.length && Number(x.size) >= minSize; });
  const pur = gbPurRead_();   // 本次執行第一次讀採購系統＝全新資料（同一次執行內重讀可能拿到快取舊值，見 v1.3.1）
  if (!dry) {
    const vr = gbSyncVerifyPrev_(P, pur.bookings);
    if (vr.state === 'bad') {
      P.setProperty('GB_SYNC', 'paused');
      const msg = vr.text + '\n\n已自動暫停同步（GB_SYNC=paused）。請通知 Chat 檢查；確認後執行 zz_gbSyncOn 重新開啟。';
      gbLog_('error', msg.slice(0, 1900));
      gbAlertOnce_(P, 'SYNC_VERIFY', today, '【團體訂位】同步到採購系統後核對不符，已暫停同步', msg);
      return { state: 'verify_fail', report: vr.text };
    }
    if (vr.state === 'ok') {
      gbLog_('info', vr.text);
      if (!P.getProperty('GB_SYNC_PENDING')) return { state: 'verified', report: vr.text };   // 這次只是為了核對而執行
    }
  }
  const plan = gbSyncPlan_(groups, pur, today);
  const ops = plan.adds.length + plan.cancels.length + plan.takeovers.length;
  const lines = [];
  lines.push((dry ? '【試算，不寫入】' : '【同步】') + today + '　有效團體 ' + groups.length + ' 場' + (minSize ? '（只算 ' + minSize + ' 人以上）' : '') + '、甜點 ' + plan.desired.length + ' 項；採購系統現有同步預約 ' + plan.existingAuto + ' 筆');
  lines.push('新增 ' + plan.adds.filter(function (a) { return !a.replaces; }).length + '、更新 ' + plan.adds.filter(function (a) { return a.replaces; }).length
    + '、取消 ' + plan.drops + '、接手店長手動登記 ' + plan.takeovers.length + '、不變 ' + plan.kept + '；對不到甜點 ' + plan.unmatched.length + ' 項');
  plan.adds.forEach(function (a) { lines.push('  ＋ ' + a.store + '店 ' + a.date + ' ' + a.dessert + ' ×' + a.qty + (a.how === 'miss' ? '　⚠️對不到，請店長選' : (a.how === 'approx' ? '　≈ 由「' + a.pos + '」比對' : (a.how === 'learned' ? '　（套用店長上次的選擇）' : ''))) + (a.replaces ? '　（取代 ' + a.replaces + '）' : '')); });
  plan.cancels.forEach(function (c) { lines.push('  － ' + c.b.store + '店 ' + c.b.date + ' ' + c.b.dessert + ' ×' + c.b.qty + '　' + c.why); });
  plan.takeovers.forEach(function (c) { lines.push('  ⇄ ' + c.b.store + '店 ' + c.b.date + ' ' + c.b.dessert + ' ×' + c.b.qty + '　' + c.why); });
  plan.skipped.forEach(function (x) { lines.push('  略過：' + x); });
  plan.warn.forEach(function (x) { lines.push('  ⚠️ ' + x); });

  // 安全檢查（任一不過就停，不寫入）
  let stop = '';
  if (!plan.desired.length && plan.existingAuto >= 5) stop = '團體甜點資料是 0 筆，但採購系統有 ' + plan.existingAuto + ' 筆同步預約——可能是抓取異常，不敢全部取消';
  else if (plan.drops > PUR.MIN_DROP_CHECK && plan.drops > plan.existingAuto * PUR.MAX_DROP_RATIO) stop = '要取消 ' + plan.drops + ' 筆同步預約，超過現有 ' + plan.existingAuto + ' 筆的一半';
  else if (ops > PUR.MAX_OPS) stop = '這次要寫入／取消 ' + ops + ' 筆，超過上限 ' + PUR.MAX_OPS;
  if (stop && P.getProperty('GB_SYNC_FORCE') !== '1') {
    lines.push('⛔ 安全檢查未通過，停止同步：' + stop + '（確認沒問題可設指令碼屬性 GB_SYNC_FORCE=1 後再執行一次，執行完會自動清除）');
    if (!dry) { gbLog_('error', lines.join('\n')); gbAlertOnce_(P, 'SYNC_STOP', today, '【團體訂位】同步到採購系統已暫停（安全檢查）', lines.join('\n')); }
    return { state: 'stopped', report: lines.join('\n'), plan: plan };
  }
  if (dry) return { state: 'dry', report: lines.join('\n'), plan: plan };

  // 寫入：同一筆先新增再取消（中途被打斷也只會暫時多一筆，下一次同步會自動清掉，不會變成少備料）
  let done = 0, seq = 0, cut = false, mirrorFail = 0;
  const addedIds = [], cancelledIds = [];
  const newId = function () { seq++; return 'BKG' + Date.now() + ('00' + seq).slice(-3); };
  const replaced = {};
  for (let i = 0; i < plan.adds.length; i++) {
    if (gbTimeLeft_(t0) < 40000) { cut = true; break; }
    const a = plan.adds[i]; const id = newId();
    if (gbPurPost_({ type: 'booking', id: id, store: a.store, date: a.date, dessert: a.dessert, qty: a.qty, note: a.note }).mirrorFail) mirrorFail++;
    addedIds.push({ id: id, a: a }); done++;
    if (a.replaces) replaced[a.replaces] = 1;
  }
  const cancelList = plan.cancels.concat(plan.takeovers);
  for (let j = 0; j < cancelList.length && !cut; j++) {
    const c = cancelList[j];
    // 「更新」的舊列只有在新列已經寫進去後才取消
    if (/^更新/.test(c.why) && !replaced[c.id]) continue;
    if (gbTimeLeft_(t0) < 30000) { cut = true; break; }
    if (gbPurPost_({ type: 'bookingDel', id: c.id }).mirrorFail) mirrorFail++;
    cancelledIds.push(c.id); done++;
  }

  // 核對不在這裡做：同一次執行內重讀，Google 可能給快取的舊值（v1.3.1）。記下清單，下一次執行（約 15 分鐘後）用全新讀取核對
  if (addedIds.length || cancelledIds.length) gbSyncVerifySave_(P, addedIds, cancelledIds);

  const soonEnd = gbAddDays_(today, PUR.SOON_DAYS);
  const last = {
    time: gbNowStr_(), adds: addedIds.length, cancels: cancelledIds.length, kept: plan.kept, unmatched: plan.unmatched.length,
    unmatched_soon: plan.unmatched.filter(function (u) { return u.date <= soonEnd; }).map(function (u) { return u.date.slice(5) + ' ' + u.store.replace(/店$/, '') + ' ' + u.pos + ' ×' + u.qty; }).slice(0, 30),
    verify: (addedIds.length || cancelledIds.length) ? '下一次執行核對' : '—', pending: cut
  };
  P.setProperty('GB_SYNC_LAST', JSON.stringify(last).slice(0, 8000));
  if (!cut) P.deleteProperty('GB_SYNC_PENDING');
  if (P.getProperty('GB_SYNC_FORCE')) P.deleteProperty('GB_SYNC_FORCE');
  lines.push('完成：寫入 ' + addedIds.length + '、取消 ' + cancelledIds.length + (cut ? '（時間不夠，剩下的 15 分鐘後接著做）' : '') + ((addedIds.length || cancelledIds.length) ? '；下一次執行（約 15 分鐘後）會重新讀取核對' : '')
    + (mirrorFail ? '；店長畫面那份（專用檔）有 ' + mirrorFail + ' 筆沒同步到，明早 07:05 採購系統重算時會自動補上' : ''));
  gbLog_('info', lines.join('\n').slice(0, 1900));
  return { state: cut ? 'partial' : 'ok', report: lines.join('\n'), plan: plan, last: last };
}

// ---- 寫入核對（v1.3.1）：這次寫了什麼記下來，下一次執行用全新讀取比對 ----
function gbSyncVerifySave_(P, addedIds, cancelledIds) {
  const mk = function (full) {
    return JSON.stringify({ t: gbNowStr_(),
      a: addedIds.map(function (x) { return full ? [x.id, x.a.store, x.a.date, x.a.dessert, x.a.qty] : [x.id]; }),
      c: cancelledIds });
  };
  let v = mk(true);
  if (v.length > 8500) v = mk(false);                 // 太長（屬性上限約 9KB）→ 只核對「有沒有寫進去」
  if (v.length > 8500) v = JSON.stringify({ t: gbNowStr_(), a: addedIds.slice(0, 250).map(function (x) { return [x.id]; }), c: cancelledIds.slice(0, 150), part: 1 });
  P.setProperty('GB_SYNC_VERIFY', v);
}
function gbSyncVerifyPrev_(P, bookings) {
  const raw = P.getProperty('GB_SYNC_VERIFY');
  if (!raw) return { state: 'none', text: '' };
  let v; try { v = JSON.parse(raw); } catch (e) { P.deleteProperty('GB_SYNC_VERIFY'); return { state: 'none', text: '' }; }
  const byId = {}; bookings.forEach(function (b) { if (!byId[b.id]) byId[b.id] = b; });   // 與採購系統 upsert_ 一致：同編號以第一筆為準
  const bad = [];
  (v.a || []).forEach(function (x) {
    const b = byId[x[0]];
    if (!b) { bad.push('新增 ' + x[0] + ' 不存在'); return; }
    // 之後被店長「改甜點」而取消是正常的（狀態可以是取消），只核對寫入當下的內容
    if (x.length > 1 && (b.store !== String(x[1]) || b.date !== x[2] || b.dessert !== x[3] || b.qty !== Number(x[4]))) bad.push('新增 ' + x[0] + ' 內容不符');
  });
  (v.c || []).forEach(function (id) { if (!byId[id] || byId[id].status !== '取消') bad.push('取消 ' + id + ' 仍有效'); });
  P.deleteProperty('GB_SYNC_VERIFY');
  const head = '上次寫入核對（' + v.t + '，寫入 ' + (v.a || []).length + '、取消 ' + (v.c || []).length + (v.part ? '，只核對前面一部分' : '') + '）：';
  if (bad.length) return { state: 'bad', text: head + '⚠️ 不符 ' + bad.length + ' 筆：' + bad.slice(0, 8).join('、') };
  return { state: 'ok', text: head + '全部相符' };
}

function gbPurPost_(obj) {
  const tok = PropertiesService.getScriptProperties().getProperty('PUR_TOKEN') || PUR.TOKEN_DEFAULT;
  const payload = {}; Object.keys(obj).forEach(function (k) { payload[k] = obj[k]; }); payload.token = tok;
  // 採購系統同一時間只讓一個人寫；店長正在存盤點時會回「busy」（HTTP 仍是 200）→ 等幾秒重送，不當成失敗
  for (let i = 0; i <= PUR.POST_RETRY; i++) {
    const r = UrlFetchApp.fetch(PUR.API, { method: 'post', contentType: 'text/plain;charset=utf-8', payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: true });
    const code = r.getResponseCode();
    if (code >= 400) throw new Error('採購系統寫入失敗 HTTP ' + code + '（' + obj.type + ' ' + (obj.id || '') + '）');
    let txt = ''; try { txt = String(r.getContentText() || ''); } catch (e) { txt = ''; }
    if (!/busy/i.test(txt)) return { code: code, mirrorFail: /mirror\W{0,4}fail/i.test(txt) };
    if (i < PUR.POST_RETRY) Utilities.sleep(5000 * (i + 1));
  }
  throw new Error('採購系統持續忙碌（重送 ' + PUR.POST_RETRY + ' 次仍忙碌）：' + obj.type + ' ' + (obj.id || ''));
}

// ============ API（JSONP）============
function doGet(e) {
  const p = (e && e.parameter) || {};
  let out;
  try {
    if (p.fn === 'ping') out = { ok: true, version: GB.VERSION };
    else if (p.fn === 'group') {
      const pass = PropertiesService.getScriptProperties().getProperty('GB_PASSCODE');
      if (!pass) out = { ok: false, error: 'server_passcode_not_set', msg: '後端尚未設定通關碼（GB_PASSCODE）' };
      else if (String(p.key || '') !== pass) { Utilities.sleep(1200); out = { ok: false, error: 'passcode', msg: '通關碼錯誤' }; }
      else out = gbApiPayload_();
    } else out = { ok: false, error: 'unknown_fn' };
  } catch (err) {
    out = { ok: false, error: 'server', msg: String((err && err.message) || err) };
  }
  const body = JSON.stringify(out);
  const cb = p.callback;
  if (cb && /^[A-Za-z_$][\w$]{0,60}$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + body + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function gbApiPayload_() {
  const today = gbToday_();
  const P = PropertiesService.getScriptProperties();
  const rows = gbBookingsRead_();
  const rep = gbBuildReport_(rows, today);
  const cyc = gbGetCycle_(P);
  const keep = rep.rows.filter(function (x) { return x.is_group || x.hold; });
  return {
    ok: true, version: GB.VERSION, today: today, now: gbNowStr_(),
    last_done: P.getProperty('GB_LAST_DONE') || '',
    last_error: P.getProperty('GB_LAST_ERROR') || '',
    running: cyc ? { phase: cyc.phase, started: gbFmt_(new Date(cyc.started)) } : null,
    window: { from: today, to: gbAddDays_(today, GB.DAYS_AHEAD) },
    rules: { hold_days: GB.HOLD_DAYS, dessert_days: GB.DESSERT_DAYS, group_min: GB.GROUP_MIN_SIZE },
    sync: (function () { let l = null; try { l = JSON.parse(P.getProperty('GB_SYNC_LAST') || 'null'); } catch (e) {} return { state: P.getProperty('GB_SYNC') || 'off', last: l }; })(),
    summary: rep.summary,
    rows: keep.map(function (x) {
      return {
        id: x.id, kind: x.kind, date: x.date, slot: x.slot, store: x.store, member: x.member, phone: x.phone,
        cat: x.category, size: x.size, recipes: x.recipes, comp: x.companions, total: x.total,
        paid: x.paid, paid_on: x.paid_online, paid_off: x.paid_offline, five: x.five,
        status: x.status, status_text: x.status_text, memo: x.memo, menu: x.menu,
        modify: x.modify, modify_by: x.modify_by, deadline: x.deadline || '', deposit: x.deposit, min_price: x.min_price,
        prepay_time: x.prepay_time || '', active: x.active, is_group: x.is_group, exempt: x.exempt, need_dessert: x.need_dessert,
        days_to: x.days_to, hold: x.hold, awaiting: x.awaiting, changes: x.changes, pair_ids: x.pair_ids || [],
        issues: x.issues
      };
    })
  };
}

// ============ 每日提醒信 ============
function gbSendDaily_(today) {
  const P = PropertiesService.getScriptProperties();
  const rep = gbBuildReport_(gbBookingsRead_(), today);
  const holdRows = rep.rows.filter(function (x) {
    return (x.hold && (x.hold.type === 'orphan' || (x.hold.type === 'waiting' && x.hold.awaiting.overdue))) ||
           (x.awaiting && x.awaiting.overdue);
  });
  // v1.3：採購系統同步後仍對不到 BOM 的甜點（14 天內到店），還沒算進備料
  let syncLast = null; try { syncLast = JSON.parse(P.getProperty('GB_SYNC_LAST') || 'null'); } catch (e) {}
  const miss = (P.getProperty('GB_SYNC') === 'on' && syncLast && syncLast.unmatched_soon) ? syncLast.unmatched_soon : [];
  if (!holdRows.length && !miss.length) { gbLog_('info', '今日無逾期卡位、無待選甜點，不寄信'); return; }
  const orphan = holdRows.filter(function (x) { return x.hold && x.hold.type === 'orphan'; }).length;
  const overdue = holdRows.length - orphan;
  holdRows.sort(function (a, b) { return (a.date + a.slot).localeCompare(b.date + b.slot); });
  const esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  let html = '<div style="font-family:sans-serif;font-size:14px;color:#2D1A14">';
  if (holdRows.length) {
    html += '<p>以下訂位正在佔用時段、卻沒有收到訂金（或它保護的團體已不存在）。請確認後到後台處理：</p>'
      + '<table cellpadding="6" style="border-collapse:collapse;font-size:13px"><tr style="background:#F3EEF9">'
      + '<th align="left">到店</th><th align="left">店</th><th align="left">名義／客人</th><th align="left">電話</th><th>人數</th><th align="left">狀況</th><th align="left">後台</th></tr>';
    holdRows.forEach(function (x) {
      const iss = x.issues.filter(function (i) { return i.sev === 'high'; })[0];
      const who = x.member === GB.HOLD_NAME ? GB.HOLD_NAME + '（' + String(x.memo || '').slice(0, 24) + '）' : x.member;
      html += '<tr style="border-top:1px solid #eee"><td>' + esc(x.date.slice(5) + ' ' + x.slot) + '</td><td>' + esc(x.store) + '</td><td>' + esc(who)
        + '</td><td>' + esc(x.phone) + '</td><td align="center">' + x.size + '</td><td>' + esc(iss ? iss.title : '') + '<br><span style="color:#6F4E45">→ ' + esc(iss ? iss.fix : '') + '</span></td>'
        + '<td><a href="' + GB.BASE + '/Reservations/' + x.kind + '/' + x.id + '">前往修改</a></td></tr>';
    });
    html += '</table>';
  }
  if (miss.length) {
    html += '<p style="margin-top:16px"><b>採購系統：以下團體甜點的名稱對不到 BOM，還沒算進備料</b>（最近一次同步 ' + esc(syncLast.time) + '）。'
      + '請該店店長打開採購系統「🎉 預約」頁，從清單選出正確的甜點按「確定」，選過一次之後同名甜點會自動套用：</p><ul style="font-size:13px">'
      + miss.map(function (m) { return '<li>' + esc(m) + '</li>'; }).join('') + '</ul>';
  }
  html += '<p style="color:#6F4E45;font-size:12px">另有 ' + rep.summary.high + ' 項需立即處理、' + rep.summary.mid
    + ' 項待更正的資料，請看訂位儀表板「團體訂位追蹤」：<a href="' + GB.DASH_URL + '">' + GB.DASH_URL + '</a><br>'
    + '規則：團體（8 人以上）訂位後 ' + GB.HOLD_DAYS + ' 天內要付訂金；公司同意的免訂金專案，請在後台備註加上「【免訂金】」就不會再被提醒。'
    + '客服配合門市卡位的（會員名義「客服訂位」、備註寫「卡位」）不是真的客人，不會被提醒。</p></div>';
  const to = P.getProperty('GB_NOTIFY') || GB.NOTIFY_DEFAULT;
  const subject = holdRows.length
    ? '【團體訂位】卡位逾期 ' + overdue + ' 筆、孤兒卡位 ' + orphan + ' 筆' + (miss.length ? '、採購甜點待選 ' + miss.length + ' 項' : '') + '（' + today.slice(5) + '）'
    : '【團體訂位】採購系統有 ' + miss.length + ' 個團體甜點名稱待店長選擇（' + today.slice(5) + '）';
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html });
  gbLog_('info', '已寄提醒信：逾期 ' + overdue + '、孤兒 ' + orphan + '、待選甜點 ' + miss.length);
}

function gbStaleCheck_(P, today) {
  const last = P.getProperty('GB_LAST_DONE');
  if (!last) return;
  const ms = Date.parse(last.replace(' ', 'T') + ':00+08:00');
  if (isNaN(ms) || (Date.now() - ms) < GB.STALE_HOURS * 3600 * 1000) return;
  gbAlertOnce_(P, 'STALE', today, '【團體訂位】資料已超過 ' + GB.STALE_HOURS + ' 小時沒有更新',
    '最後完成時間：' + last + '\n最近錯誤：' + (P.getProperty('GB_LAST_ERROR') || '無') + '\n請在 Apps Script 執行 gbStatus 查看狀態。');
}

function gbAlertOnce_(P, kind, today, subject, body) {
  const k = 'GB_ALERT_' + kind;
  if (P.getProperty(k) === today) return;
  P.setProperty(k, today);
  const to = P.getProperty('GB_NOTIFY') || GB.NOTIFY_DEFAULT;
  try { MailApp.sendEmail(to, subject, body); } catch (e) { gbLog_('error', '警示信寄送失敗：' + e.message); }
}

// ============ 試算表存取 ============
function gbEnsureSheets_() {
  gbSheet_(GB.SHEET_BOOKINGS, GB_OUT_COLS);
  gbSheet_(GB.SHEET_STAGE, GB_COLS);
  gbSheet_(GB.SHEET_DETAIL, GB_DETAIL_COLS);
  gbSheet_(GB.SHEET_LOG, ['time', 'level', 'message']);
}
function gbSheet_(name, cols) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function gbReadObjs_(name, cols) {
  const sh = gbSheet_(name, cols);
  const n = sh.getLastRow() - 1;
  if (n <= 0) return [];
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  return sh.getRange(2, 1, n, head.length).getValues().map(function (a) {
    const o = {}; head.forEach(function (h, i) { o[h] = a[i] === null || a[i] === undefined ? '' : String(a[i]); }); return o;
  });
}
function gbStageClear_() {
  const sh = gbSheet_(GB.SHEET_STAGE, GB_COLS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
}
function gbStageAppend_(rows) {
  if (!rows.length) return;
  const sh = gbSheet_(GB.SHEET_STAGE, GB_COLS);
  const rg = sh.getRange(sh.getLastRow() + 1, 1, rows.length, GB_COLS.length);
  rg.setNumberFormat('@');
  rg.setValues(rows.map(function (r) { return GB_COLS.map(function (k) { return String(r[k] == null ? '' : r[k]); }); }));
}
function gbStageRead_() { return gbReadObjs_(GB.SHEET_STAGE, GB_COLS); }
function gbBookingsRead_() { return gbReadObjs_(GB.SHEET_BOOKINGS, GB_OUT_COLS); }
function gbDedupe_(rows) {
  const m = {}, order = [];
  rows.forEach(function (r) { const k = r.id || (r.date + r.slot + r.store + r.member + r.size); if (!(k in m)) order.push(k); m[k] = r; });
  return order.map(function (k) { return m[k]; });
}
function gbDetailCacheRead_() {
  const o = {};
  gbReadObjs_(GB.SHEET_DETAIL, GB_DETAIL_COLS).forEach(function (r) { if (r.id) o[r.id] = r; });
  return o;
}
function gbDetailCacheWrite_(cache) {
  const sh = gbSheet_(GB.SHEET_DETAIL, GB_DETAIL_COLS);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getMaxColumns()).clearContent();
  const vals = Object.keys(cache).map(function (k) { return GB_DETAIL_COLS.map(function (c) { return String(cache[k][c] == null ? '' : cache[k][c]); }); });
  if (vals.length) {
    const rg = sh.getRange(2, 1, vals.length, GB_DETAIL_COLS.length);
    rg.setNumberFormat('@');
    rg.setValues(vals);
  }
}
function gbLog_(level, msg) {
  try {
    const sh = gbSheet_(GB.SHEET_LOG, ['time', 'level', 'message']);
    sh.appendRow([gbNowStr_(), level, String(msg).slice(0, 2000)]);
    if (sh.getLastRow() > 600) sh.deleteRows(2, 100);
  } catch (e) { /* 記錄失敗不影響主流程 */ }
  Logger.log('[' + level + '] ' + msg);
}

// ============ 游標 ============
function gbGetCycle_(P) {
  try { return JSON.parse(P.getProperty('GB_CYCLE') || 'null'); } catch (e) { return null; }
}
function gbSaveCycle_(P, cyc) { P.setProperty('GB_CYCLE', JSON.stringify(cyc)); }

// ============ 小工具（純 JS，不用 Java 橋接）============
function gbToday_() { return Utilities.formatDate(new Date(), GB.TZ, 'yyyy-MM-dd'); }
function gbNowStr_() { return Utilities.formatDate(new Date(), GB.TZ, 'yyyy-MM-dd HH:mm'); }
function gbFmt_(d) { return Utilities.formatDate(d, GB.TZ, 'yyyy-MM-dd HH:mm'); }
function gbTimeLeft_(t0) { return GB.BUDGET_MS - (Date.now() - t0); }
function gbAddDays_(ds, n) {
  const d = new Date(Date.UTC(Number(ds.slice(0, 4)), Number(ds.slice(5, 7)) - 1, Number(ds.slice(8, 10)) + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function gbDiffDays_(a, b) {  // b - a（天）
  if (!a || !b) return 0;
  const ta = Date.UTC(Number(a.slice(0, 4)), Number(a.slice(5, 7)) - 1, Number(a.slice(8, 10)));
  const tb = Date.UTC(Number(b.slice(0, 4)), Number(b.slice(5, 7)) - 1, Number(b.slice(8, 10)));
  return Math.round((tb - ta) / 86400000);
}
function gbMinStr_(a, b) { return a < b ? a : b; }
function gbAttrs_(s) {
  const o = {};
  String(s || '').replace(/([\w:-]+)\s*=\s*"([^"]*)"/g, function (_, k, v) { o[k.toLowerCase()] = v; return _; });
  return o;
}
function gbDecode_(s) {
  return String(s || '')
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCharCode(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}
function gbText_(h) {
  return gbDecode_(String(h || '').replace(/data-content="[^"]*"/gi, '').replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]*>/g, ''))
    .replace(/\s+/g, ' ').trim();
}
function gbNum_(h) {
  const m = gbText_(h).replace(/,/g, '').match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : 0;
}