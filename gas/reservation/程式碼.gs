/**
 * DIYBC 訂位系統自動抓取 → fact_reservations
 * v1.6 2026-08-01
 * 變更：新增 cookie 快取（登入一次 20 分鐘內重複使用，登入態失效自動重登重試）；
 *       每日任務「至少做完一天才准中止」，避免慢速登入吃光時間導致整輪空轉
 * 鐵則：回補或手動貼資料後必須執行 buildAllMonthCaches()；程式修改後 doGet 要「部署新版本」才生效
 * 地雷：單次執行上限 6 分鐘、觸發器每日總額 90 分鐘；Azure 登入端點時快時慢（冷啟動）
 */

const CFG = {
  BASE: 'https://diybc.azurewebsites.net',
  TZ: 'Asia/Taipei',
  FACT: 'fact_reservations',
  FUTURE: 'fact_reservations_future',
  LOOKBACK_DAYS: 3,
  FUTURE_DAYS: 35,
  SESS_TTL_MS: 20 * 60 * 1000,
  UA: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36',
  HEADERS: ['res_id','date','store','slot','member_name','phone','line_uid','email','category',
            'party_size','recipes','companions','order_total','paid','status','attended',
            'purpose','member_level','note','group_id','fetched_at'],
  EXPECTED_TH: ['日期','分店','時段','會員','電話','LineUID','Email','類別','人數','食譜',
                '陪同','總計','已付','狀態','出席','目的','自己人等級','備註']
};

/* ========== 一次性設定 ========== */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CFG.FACT, CFG.FUTURE].forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, CFG.HEADERS.length).setValues([CFG.HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  Logger.log('分頁建立完成');
}

function createDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'dailyFetchReservations') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyFetchReservations').timeBased().everyDays(1).atHour(10).create();
  Logger.log('每日 10:00 觸發器已建立');
}

/* ========== 每日任務（動態回看＋cookie 快取＋快取刷新） ========== */

function dailyFetchReservations() {
  const t0 = Date.now();
  const touched = {};
  let done = 0;
  try {
    const yest = addDays_(new Date(), -1);
    let start = addDays_(new Date(), -CFG.LOOKBACK_DAYS);
    const last = lastFactDate_();
    if (last) {
      const next = addDays_(parseDate_(last), 1);
      if (next < start) { start = next; Logger.log('偵測到斷層，從 ' + fmtDate_(next) + ' 補起'); }
    }
    const cap = addDays_(new Date(), -14);
    if (start < cap) start = cap;

    for (let d = start; d <= yest; d = addDays_(d, 1)) {
      // 至少做完一天才准中止（避免慢速登入吃光時間、整輪零產出）
      if (done > 0 && Date.now() - t0 > 280000) {
        Logger.log('時間保險絲：補到 ' + fmtDate_(d) + ' 前中止（明日自動續補）');
        refreshCaches_(touched); return;
      }
      const ds = fmtDate_(d);
      writeFactDay_(ds, fetchDayAuto_(ds));
      touched[ds.slice(0, 7)] = true;
      done++;
      Utilities.sleep(150);
    }
    refreshCaches_(touched);
    Logger.log('已補 ' + done + ' 天，耗時 ' + Math.round((Date.now() - t0) / 1000) + ' 秒');
    if (Date.now() - t0 > 200000) { Logger.log('時間不足：future 本日沿用舊資料'); return; }
    rebuildFuture_(t0);
    Logger.log('每日抓取完成，共 ' + Math.round((Date.now() - t0) / 1000) + ' 秒');
  } catch (e) {
    refreshCaches_(touched);
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '【DIYBC 訂位抓取失敗】' + fmtDate_(new Date()),
      (e && e.message ? e.message : e) + '\n\n' + (e && e.stack ? e.stack : ''));
    throw e;
  }
}

function refreshCaches_(touched) {
  if (!Object.keys(touched).length) return;
  try { rebuildCaches_(touched); } catch (e) { Logger.log('快取刷新失敗：' + e.message); }
}

function lastFactDate_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const vals = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  let max = '';
  vals.forEach(r => { const d = normDate_(r[0]); if (d > max) max = d; });
  return max || null;
}

function rebuildFuture_(t0) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FUTURE);
  if (!sh) throw new Error('找不到分頁 ' + CFG.FUTURE);
  const all = [];
  for (let i = 0; i <= CFG.FUTURE_DAYS; i++) {
    if (t0 && Date.now() - t0 > 300000) { Logger.log('時間保險絲：future 抓取未完，本日沿用舊資料'); return; }
    const d = fmtDate_(addDays_(new Date(), i));
    all.push.apply(all, fetchDayAuto_(d));
    Utilities.sleep(200);
  }
  const last = sh.getLastRow(); // 全部抓完才動 Sheet，中斷不會留半套
  if (last > 1) sh.getRange(2, 1, last - 1, CFG.HEADERS.length).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, CFG.HEADERS.length).setValues(all);
  try { futureMenuWrite_(all, null); } catch (e) { Logger.log('甜點明細寫入失敗（不影響 future）：' + e.message); }   // 2026-10-06 1006 #2
}

/* ---- 2026-10-06 各儀表板調整 1006 #2：未來訂位的甜點明細（給決策中心算「同時段器具／模具夠不夠」）----
   分頁 fact_future_menu：date｜store｜slot｜category｜party_size｜status｜menu（JSON：[[甜點名, 份數], …]）｜fetched_at，只存有選甜點的訂位；
   不含姓名電話。daysSet＝null 整份換掉（每天 10:15 整份重建）；有 daysSet＝只換那幾天、其他今天以後的照舊（白天每 3 小時近 14 天）。 */
const FM_SHEET = 'fact_future_menu';
const FM_HEAD = ['date', 'store', 'slot', 'category', 'party_size', 'status', 'menu', 'fetched_at'];
function futureMenuWrite_(rows, daysSet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FM_SHEET);
  if (!sh) { sh = ss.insertSheet(FM_SHEET); sh.getRange(1, 1, 1, FM_HEAD.length).setValues([FM_HEAD]); sh.setFrozenRows(1); }
  const H = CFG.HEADERS, iD = H.indexOf('date'), iS = H.indexOf('store'), iT = H.indexOf('slot'), iC = H.indexOf('category'),
        iP = H.indexOf('party_size'), iSt = H.indexOf('status'), iF = H.indexOf('fetched_at');
  const fresh = rows.filter(r => r.menu && r.menu.length)
    .map(r => [wdDateStr_(r[iD]), String(r[iS] || ''), normSlot_(r[iT]), String(r[iC] || ''), Number(r[iP]) || 0, String(r[iSt] || ''), JSON.stringify(r.menu), String(r[iF] || '')]);
  const last = sh.getLastRow(), keep = [], today = fmtDate_(new Date());
  if (daysSet && last > 1) {
    sh.getRange(2, 1, last - 1, FM_HEAD.length).getValues().forEach(r => { const d = wdDateStr_(r[0]); if (d && d >= today && !daysSet[d]) keep.push(r); });
  }
  const all = fresh.concat(keep);
  if (last > 1) sh.getRange(2, 1, last - 1, FM_HEAD.length).clearContent();
  if (all.length) {
    sh.getRange(2, 1, all.length, 4).setNumberFormat('@');   // 日期、店、時段、類別保持文字（避免被自動轉成日期／時間）
    sh.getRange(2, 1, all.length, FM_HEAD.length).setValues(all);
  }
  return fresh.length;
}

/* ============ 歷史大回補（多日合併抓取＋自適應窗口＋游標自癒） ============ */

function startBackfill_v2() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('BF_NEXT')) props.setProperty('BF_NEXT', '2025-01-01');
  if (!props.getProperty('BF_END'))  props.setProperty('BF_END', '2026-06-30');
  props.deleteProperty('BF_ERR');
  clearBackfillTriggers_();
  ScriptApp.newTrigger('backfillWorker_').timeBased().everyMinutes(10).create();
  Logger.log('回補已啟動：從 ' + props.getProperty('BF_NEXT') + ' 到 ' +
             props.getProperty('BF_END') + '，每次合併抓最多 7 天');
}

function backfillWorker_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const props = PropertiesService.getScriptProperties();
  const touched = {};
  try {
    const h = Number(Utilities.formatDate(new Date(), CFG.TZ, 'H'));
    if (h >= 4 && h < 7) { Logger.log('04–07 時窗口，本輪跳過'); return; }
    const endS = props.getProperty('BF_END');
    const nextS = props.getProperty('BF_NEXT');
    if (!nextS || !endS) { clearBackfillTriggers_(); return; }
    const t0 = Date.now();
    let win = Number(props.getProperty('BF_WIN') || 7);
    let d = parseDate_(nextS);
    const end = parseDate_(endS);
    let done = 0;

    while (d <= end && (done === 0 || Date.now() - t0 < 180000)) {
      let d2 = addDays_(d, win - 1); if (d2 > end) d2 = end;
      const ds = fmtDate_(d), ds2 = fmtDate_(d2);
      const rows = fetchRangeAuto_(ds, ds2);
      if (rows.length >= 900 && ds !== ds2) {
        win = Math.max(1, Math.floor(win / 2));
        props.setProperty('BF_WIN', String(win));
        Logger.log(ds + '~' + ds2 + ' 回傳 ' + rows.length + ' 筆近上限，縮窗至 ' + win + ' 天重抓');
        continue;
      }
      if (rows.length >= 900) {
        Logger.log('⚠ 單日 ' + ds + ' 即達 ' + rows.length + ' 筆，可能被截斷，照收並標記');
      }
      writeFactRange_(ds, ds2, rows);
      let t = parseDate_(ds);
      while (t <= parseDate_(ds2)) { touched[fmtDate_(t).slice(0, 7)] = true; t = addDays_(t, 1); }
      props.setProperty('BF_NEXT', fmtDate_(addDays_(d2, 1)));
      Logger.log(ds + '~' + ds2 + '：' + rows.length + ' 筆');
      if (rows.length < 450 && win < 7) { win++; props.setProperty('BF_WIN', String(win)); }
      d = addDays_(d2, 1);
      done++;
      Utilities.sleep(150);
    }

    props.deleteProperty('BF_ERR');
    if (d > end) {
      clearBackfillTriggers_();
      props.deleteProperty('BF_NEXT'); props.deleteProperty('BF_END'); props.deleteProperty('BF_WIN');
      refreshCaches_(touched);
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【DIYBC 訂位回補完成】',
        '已寫入至 ' + endS + '。快取已刷新，可直接檢查儀表板各月筆數。');
    } else {
      refreshCaches_(touched);
    }
  } catch (e) {
    refreshCaches_(touched);
    const err = Number(props.getProperty('BF_ERR') || 0) + 1;
    props.setProperty('BF_ERR', String(err));
    Logger.log('連續第 ' + err + ' 次失敗：' + e.message + '（下一輪自動再試）');
    if (err === 30) {
      clearBackfillTriggers_();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【DIYBC 訂位回補暫停】連續失敗 30 次',
        '停在 ' + props.getProperty('BF_NEXT') + '\n最後錯誤：' + e.message +
        '\n\n排除後執行 startBackfill_v2() 從游標續跑。');
    }
  } finally {
    lock.releaseLock();
  }
}

function writeFactRange_(ds, ds2, rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  if (!sh) throw new Error('找不到分頁 ' + CFG.FACT);
  const dates = {};
  let d = parseDate_(ds);
  const end = parseDate_(ds2);
  while (d <= end) { dates[fmtDate_(d)] = true; d = addDays_(d, 1); }
  const clean = rows.filter(r => dates[r[1]]);
  if (clean.length !== rows.length) {
    Logger.log('⚠ ' + ds + '~' + ds2 + '：過濾掉 ' + (rows.length - clean.length) + ' 筆窗口外日期列');
  }
  deleteRowsForDates_(sh, dates);
  if (clean.length) {
    sh.getRange(sh.getLastRow() + 1, 1, clean.length, CFG.HEADERS.length).setValues(clean);
  }
}

function deleteRowsForDates_(sh, dateSet) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  const idx = [];
  for (let i = 0; i < vals.length; i++) {
    if (dateSet[normDate_(vals[i][0])]) idx.push(i + 2);
  }
  let i = idx.length - 1;
  while (i >= 0) {
    const end = idx[i];
    let start = end;
    while (i > 0 && idx[i - 1] === start - 1) { i--; start = idx[i]; }
    sh.deleteRows(start, end - start + 1);
    i--;
  }
}

function resumeBackfill() { startBackfill_v2(); }

function backfillStatus() {
  const p = PropertiesService.getScriptProperties();
  Logger.log('下一個待抓日：' + (p.getProperty('BF_NEXT') || '（無進行中回補）') +
             '｜終點：' + (p.getProperty('BF_END') || '-') +
             '｜目前窗口：' + (p.getProperty('BF_WIN') || '7') + ' 天' +
             '｜連續失敗：' + (p.getProperty('BF_ERR') || '0'));
}

function stopBackfill() {
  clearBackfillTriggers_();
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty('BF_NEXT'); p.deleteProperty('BF_END');
  p.deleteProperty('BF_ERR'); p.deleteProperty('BF_WIN');
  Logger.log('回補已停止並清除游標');
}

function clearBackfillTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'backfillWorker_') ScriptApp.deleteTrigger(t);
  });
}

/* ===== 任意區段補資料（改日期後執行即可） ===== */
function fixJulyGap() {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('BF_NEXT', '2026-07-31');
  props.setProperty('BF_END', fmtDate_(addDays_(new Date(), -1)));
  props.deleteProperty('BF_ERR');
  props.setProperty('BF_WIN', '5');
  clearBackfillTriggers_();
  ScriptApp.newTrigger('backfillWorker_').timeBased().everyMinutes(10).create();
  Logger.log('斷層修補已啟動：' + props.getProperty('BF_NEXT') + ' ~ 昨天，完成寄信。');
}

/* ===== 重複稽核（唯讀） ===== */
function checkDuplicates() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  const last = sh.getLastRow();
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const seen = {}, dup = {};
  vals.forEach(r => {
    const id = String(r[0]).trim();
    if (!id) return;
    if (seen[id]) dup[normDate_(r[1])] = (dup[normDate_(r[1])] || 0) + 1;
    seen[id] = true;
  });
  const dates = Object.keys(dup).sort();
  if (!dates.length) { Logger.log('✅ 無重複（res_id 全表唯一，共 ' + vals.length + ' 列）'); return; }
  Logger.log('⚠ 發現重複，按日期：' + JSON.stringify(dup) +
             '\n處理：對該日期範圍重跑回補即自動清除。');
}

/* ========== 驗收測試 ========== */

function testFetchOneDay() {
  const t0 = Date.now();
  const d = '2026-07-21';
  const rows = fetchDayAuto_(d);
  Logger.log(d + ' 共 ' + rows.length + ' 筆，耗時 ' + Math.round((Date.now() - t0) / 1000) + ' 秒');
  Logger.log('第一筆：' + JSON.stringify(rows[0]));
}

/* ========== 登入 ＋ cookie 快取 ========== */

function getCookie_() {
  const p = PropertiesService.getScriptProperties();
  const c = p.getProperty('SESS_COOKIE');
  const at = Number(p.getProperty('SESS_AT') || 0);
  if (c && Date.now() - at < CFG.SESS_TTL_MS) return c;
  const nc = login_();
  p.setProperty('SESS_COOKIE', nc);
  p.setProperty('SESS_AT', String(Date.now()));
  return nc;
}

function dropCookie_() {
  PropertiesService.getScriptProperties().deleteProperty('SESS_COOKIE');
}

function fetchDayAuto_(dateStr) {
  try {
    return fetchDay_(getCookie_(), dateStr);
  } catch (e) {
    if (String(e.message).indexOf('登入態失效') < 0) throw e;
    dropCookie_();
    return fetchDay_(getCookie_(), dateStr);
  }
}

function fetchRangeAuto_(ds, ds2) {
  try {
    return fetchRange_(getCookie_(), ds, ds2);
  } catch (e) {
    if (String(e.message).indexOf('登入態失效') < 0) throw e;
    dropCookie_();
    return fetchRange_(getCookie_(), ds, ds2);
  }
}

function login_() {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty('DIYBC_EMAIL');
  const pass = props.getProperty('DIYBC_PASSWORD');
  if (!email || !pass) throw new Error('請先設定指令碼屬性 DIYBC_EMAIL / DIYBC_PASSWORD');

  const loginUrl = CFG.BASE + '/Identity/Account/Login?ReturnUrl=%2FReservations';
  const jar = {};

  const r1 = UrlFetchApp.fetch(loginUrl, {
    muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': CFG.UA }
  });
  if (r1.getResponseCode() !== 200) throw new Error('登入頁載入失敗 HTTP ' + r1.getResponseCode());
  collectCookies_(r1, jar);
  const html = r1.getContentText();

  const tokenM = html.match(/name="__RequestVerificationToken"[^>]*value="([^"]+)"/);
  if (!tokenM) throw new Error('登入頁找不到 __RequestVerificationToken（頁面結構可能已變更）');

  const pwTag = html.match(/<input\b[^>]*type="password"[^>]*>/i);
  const pwName = pwTag ? (pwTag[0].match(/name="([^"]+)"/) || [])[1] || 'Input.Password' : 'Input.Password';
  let emName = null;
  const emTag = html.match(/<input\b[^>]*type="email"[^>]*>/i);
  if (emTag) emName = (emTag[0].match(/name="([^"]+)"/) || [])[1];
  if (!emName) {
    const m2 = html.match(/<input\b[^>]*name="([^"]*Email[^"]*)"[^>]*>/i);
    emName = m2 ? m2[1] : 'Input.Email';
  }

  const payload = { '__RequestVerificationToken': tokenM[1], 'Input.RememberMe': 'false' };
  payload[emName] = email;
  payload[pwName] = pass;

  const r2 = UrlFetchApp.fetch(loginUrl, {
    method: 'post', payload: payload,
    muteHttpExceptions: true, followRedirects: false,
    headers: { 'User-Agent': CFG.UA, 'Cookie': jarStr_(jar) }
  });
  collectCookies_(r2, jar);

  const authed = Object.keys(jar).some(k => k.indexOf('Identity.Application') >= 0);
  if (r2.getResponseCode() !== 302 || !authed) {
    throw new Error('登入失敗（HTTP ' + r2.getResponseCode() + '）：檢查帳密或表單結構');
  }
  return jarStr_(jar);
}

/* ========== 核心：抓取 + 解析 ========== */

function fetchDay_(cookie, dateStr) {
  return fetchRange_(cookie, dateStr, dateStr);
}

/* ========== 暫時性錯誤自動重試（2026-09-29 新增） ==========
 * 後台偶發 502／503／504 或連線逾時 → 等一下再試，最多 3 次。
 * 時間保護：預估重試做完會超過「本次執行開始後 330 秒」就不重試，直接照舊報錯
 *（GAS 360 秒硬殺，不可以為了重試把整輪拖死）。
 * 只重試伺服器暫時錯誤；302（登入失效）、404 等一律不重試，行為與舊版相同。
 */
const FETCH_RETRY_T0_ = Date.now();            // 本次執行開始時間（每次執行都會重算）
const FETCH_RETRY_WAITS_ = [10000, 30000];     // 第 2、3 次嘗試前各等幾毫秒
const FETCH_RETRY_CODES_ = [500, 502, 503, 504];
const FETCH_RETRY_LIMIT_MS_ = 330000;

function fetchRetry_(url, opts, label) {
  let r = null, err = null;
  for (let i = 0; i <= FETCH_RETRY_WAITS_.length; i++) {
    const t = Date.now();
    r = null; err = null;
    try {
      r = UrlFetchApp.fetch(url, opts);
      if (FETCH_RETRY_CODES_.indexOf(r.getResponseCode()) < 0) {
        if (i > 0) Logger.log(label + '：第 ' + (i + 1) + ' 次嘗試成功');
        return r;
      }
    } catch (e) {
      err = e;
    }
    if (i === FETCH_RETRY_WAITS_.length) break;
    const why = r ? 'HTTP ' + r.getResponseCode() : String(err && err.message || err);
    const wait = FETCH_RETRY_WAITS_[i];
    const cost = Date.now() - t;   // 用這次花的時間估下一次
    if (Date.now() - FETCH_RETRY_T0_ + wait + cost > FETCH_RETRY_LIMIT_MS_) {
      Logger.log(label + '：' + why + '，本次執行時間不夠，不重試');
      break;
    }
    Logger.log(label + '：' + why + '，' + (wait / 1000) + ' 秒後第 ' + (i + 2) + ' 次嘗試');
    Utilities.sleep(wait);
  }
  if (err) throw err;
  return r;
}

function fetchRange_(cookie, ds, ds2) {
  const url = CFG.BASE + '/Reservations?uid=&date=' + ds + '&date2=' + ds2 +
              '&storeId=&type=&status=&memo=&userName=';
  const r = fetchRetry_(url, {
    muteHttpExceptions: true, followRedirects: false,
    headers: { 'User-Agent': CFG.UA, 'Cookie': cookie }
  }, '查詢 ' + ds + '~' + ds2);
  if (r.getResponseCode() === 302) throw new Error('查詢被轉址回登入頁：登入態失效');
  if (r.getResponseCode() !== 200) throw new Error('查詢失敗 HTTP ' + r.getResponseCode() + '（' + ds + '~' + ds2 + '）');
  return parseTable_(r.getContentText(), ds);
}

function parseTable_(html, dateStr) {
  const tableM = html.match(/<table[^>]*id="table"[^>]*>([\s\S]*?)<\/table>/);
  if (!tableM) { Logger.log(dateStr + '：找不到表格（當日可能無訂位）'); return []; }
  const table = tableM[1];

  const theadM = table.match(/<thead[\s\S]*?<\/thead>/);
  if (theadM) {
    const ths = theadM[0].match(/<th\b[^>]*>[\s\S]*?<\/th>/g) || [];
    const names = ths.map(t => {
      const dv = t.match(/data-v="([^"]*)"/);
      return dv ? decodeEnt_(dv[1]) : decodeEnt_(stripTags_(t));
    });
    for (let i = 0; i < CFG.EXPECTED_TH.length; i++) {
      if ((names[i] || '') !== CFG.EXPECTED_TH[i]) {
        throw new Error('表格結構已變更！預期[' + i + ']=' + CFG.EXPECTED_TH[i] +
                        '，實際=' + names[i] + '。全部實際表頭：' + JSON.stringify(names));
      }
    }
  }

  const bodyM = table.match(/<tbody[\s\S]*$/);
  const body = bodyM ? bodyM[0] : table;
  const trs = body.match(/<tr\b[^>]*>[\s\S]*?<\/tr>/g) || [];
  const now = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
  const out = [];

  trs.forEach(tr => {
    if (/class="separator"/.test(tr) || /colspan=/.test(tr)) return;
    const tds = tr.match(/<td\b[^>]*>[\s\S]*?<\/td>/g) || [];
    if (tds.length < 18) return;

    const open = (tr.match(/^<tr\b[^>]*>/) || [''])[0];
    const dDate = attr_(open, 'data-date');
    const dTime = attr_(open, 'data-time');
    const groupId = attr_(open, 'data-rtdid');
    const editM = tr.match(/\/Reservations\/Edit\/([0-9a-fA-F-]{36})/);
    const resId = editM ? editM[1] : '';

    const val = i => {
      const dv = tds[i].match(/data-v="([^"]*)"/);
      if (dv) return decodeEnt_(dv[1]);
      return decodeEnt_(stripTags_(tds[i]));
    };
    const attended = (/class="attend"/.test(tds[14]) && /\bchecked\b/.test(tds[14])) ? '出席' : '';
    // 2026-10-06 各儀表板調整 1006 #2：「食譜」欄 popover 有客人選的甜點（「名稱 X 份數」，散客、團體都有）→ 掛在這一列的 menu 屬性；
    // 不進 CFG.HEADERS（歷史表、未來表的欄位都不變），只由 futureMenuWrite_ 另外寫進 fact_future_menu
    const menuM = tds[9].match(/data-content="([^"]*)"/);
    const menu = menuM ? menuParse_(decodeEnt_(menuM[1])) : [];

    const row = [
      resId,
      dDate ? dDate.slice(0, 4) + '-' + dDate.slice(4, 6) + '-' + dDate.slice(6, 8) : dateStr,
      val(1),
      dTime ? dTime.slice(0, -2) + ':' + dTime.slice(-2) : val(2),
      val(3), val(4), val(5), val(6), val(7),
      toNum_(val(8)), toNum_(val(9)), toNum_(val(10)),
      toNum_(val(11)), toNum_(val(12)),
      val(13), attended,
      val(15), val(16), val(17),
      groupId, now
    ];
    if (menu.length) row.menu = menu;
    out.push(row);
  });
  return out;
}

/** 「<span>小黑炭Oreo布朗尼 (可做全素) X 1</span><br />…」→ [['小黑炭Oreo布朗尼 (可做全素)', 1], …] */
function menuParse_(html) {
  const out = [];
  String(html || '').split(/<br\s*\/?>/i).forEach(function (seg) {
    const t = stripTags_(seg).replace(/\s+/g, ' ').trim();
    const m = t.match(/^(.+?)\s*[xX×＊*]\s*(\d+)\s*$/);
    if (m && m[1]) out.push([m[1].trim(), Number(m[2]) || 0]);
  });
  return out;
}

/* ========== 寫入（冪等＋跨日過濾） ========== */

function writeFactDay_(dateStr, rows) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  if (!sh) throw new Error('找不到分頁 ' + CFG.FACT + '，請先執行 setup()');
  const clean = rows.filter(r => r[1] === dateStr);
  if (clean.length !== rows.length) {
    Logger.log('⚠ ' + dateStr + '：伺服器回傳含 ' + (rows.length - clean.length) + ' 筆非目標日期列，已過濾');
  }
  deleteRowsForDate_(sh, dateStr);
  if (clean.length) {
    sh.getRange(sh.getLastRow() + 1, 1, clean.length, CFG.HEADERS.length).setValues(clean);
  }
}

function deleteRowsForDate_(sh, dateStr) {
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  const idx = [];
  for (let i = 0; i < vals.length; i++) {
    if (normDate_(vals[i][0]) === dateStr) idx.push(i + 2);
  }
  let i = idx.length - 1;
  while (i >= 0) {
    const end = idx[i];
    let start = end;
    while (i > 0 && idx[i - 1] === start - 1) { i--; start = idx[i]; }
    sh.deleteRows(start, end - start + 1);
    i--;
  }
}

/* ========== JSONP API（Drive 月快取；隱私：不輸出姓名/電話/Email/備註） ========== */

const API_COLS = ['date','store','slot','category','party_size','recipes','companions',
                  'order_total','paid','status','attended','purpose','member_level','line_uid'];
const CACHE_PREFIX = 'RESV_CACHE_';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = (p.callback || 'callback').replace(/[^\w.]/g, '');
  try {
    if (p.fn === 'month') {
      if (!/^\d{4}-\d{2}$/.test(p.ym || '')) return out_(cb, '{"error":"ym 格式須為 YYYY-MM"}');
      let json = readCache_(p.ym);
      if (!json) {
        json = JSON.stringify(apiSheet_(CFG.FACT, p.ym));
        publishJson_(p.ym, json);
      }
      return out_(cb, json);
    }
    if (p.fn === 'future') return out_(cb, futureJson_());   // 2026-10-06：走暫存（CacheService），每次重建 future 後立刻換新
    const m = readCache_('months');
    return out_(cb, m || JSON.stringify(apiMonthsLive_()));
  } catch (err) {
    return out_(cb, JSON.stringify({ error: String((err && err.message) || err) }));
  }
}

function out_(cb, json) {
  return ContentService.createTextOutput(cb + '(' + json + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ---- 快取重建（回補或手動貼資料後必須執行 buildAllMonthCaches） ---- */
function buildAllMonthCaches() { rebuildCaches_(null); }

function rebuildCaches_(onlyYms) { // null=全部月份；{'2026-07':true}=只重建指定月（清單一律更新）
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  const last = sh.getLastRow();
  if (last < 2) return;
  const vals = sh.getRange(2, 1, last - 1, CFG.HEADERS.length).getValues();
  const idx = {};
  CFG.HEADERS.forEach((h, i) => idx[h] = i);
  const byYm = {}, cntAll = {};
  vals.forEach(r => {
    const d = normDate_(r[idx.date]);
    const ym = d.slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(ym)) return;
    cntAll[ym] = (cntAll[ym] || 0) + 1;
    if (!onlyYms || onlyYms[ym]) (byYm[ym] = byYm[ym] || []).push(apiRow_(r, idx, d));
  });
  const now = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
  Object.keys(byYm).sort().forEach(ym => {
    publishJson_(ym, JSON.stringify({ cols: API_COLS, rows: byYm[ym], cached_at: now }));
  });
  const months = Object.keys(cntAll).sort().map(ym => ({ ym: ym, rows: cntAll[ym] }));
  publishJson_('months', JSON.stringify({ months: months, cached_at: now }));
  Logger.log('快取更新完成（' + Object.keys(byYm).length + ' 個月份檔＋月份清單）' +
             (onlyYms ? '' : '：' + months.map(m => m.ym + '=' + m.rows).join(', ')));
}

function apiRow_(r, idx, d) {
  return [
    d, r[idx.store], normSlot_(r[idx.slot]), r[idx.category],
    Number(r[idx.party_size]) || 0, Number(r[idx.recipes]) || 0, Number(r[idx.companions]) || 0,
    Number(r[idx.order_total]) || 0, Number(r[idx.paid]) || 0,
    r[idx.status], r[idx.attended] === '出席' ? 1 : 0,
    r[idx.purpose], r[idx.member_level], r[idx.line_uid]
  ];
}

/* ---- Drive 快取存取 ---- */
function cacheFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('RESV_CACHE_FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('DIYBC訂位API快取');
  props.setProperty('RESV_CACHE_FOLDER', f.getId());
  return f;
}
function publishJson_(key, json) {
  const props = PropertiesService.getScriptProperties();
  const pk = CACHE_PREFIX + key;
  const fid = props.getProperty(pk);
  if (fid) { try { DriveApp.getFileById(fid).setContent(json); return; } catch (e) {} }
  const f = cacheFolder_().createFile(pk + '.json', json, 'application/json');
  props.setProperty(pk, f.getId());
}
function readCache_(key) {
  const fid = PropertiesService.getScriptProperties().getProperty(CACHE_PREFIX + key);
  if (!fid) return null;
  try { return DriveApp.getFileById(fid).getBlob().getDataAsString(); } catch (e) { return null; }
}

/* ---- 備援路徑（快取缺時現算；future 資料量小走即時） ---- */
function apiMonthsLive_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  const last = sh.getLastRow();
  if (last < 2) return { months: [] };
  const vals = sh.getRange(2, 2, last - 1, 1).getValues();
  const cnt = {};
  vals.forEach(r => {
    const ym = normDate_(r[0]).slice(0, 7);
    if (ym) cnt[ym] = (cnt[ym] || 0) + 1;
  });
  return { months: Object.keys(cnt).sort().map(ym => ({ ym: ym, rows: cnt[ym] })) };
}

function apiSheet_(sheetName, ym) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return { cols: API_COLS, rows: [] };
  const vals = sh.getRange(2, 1, last - 1, CFG.HEADERS.length).getValues();
  const idx = {};
  CFG.HEADERS.forEach((h, i) => idx[h] = i);
  const rows = [];
  vals.forEach(r => {
    const d = normDate_(r[idx.date]);
    if (ym && d.slice(0, 7) !== ym) return;
    rows.push(apiRow_(r, idx, d));
  });
  return { cols: API_COLS, rows: rows };
}

/* ---- 2026-10-06 各儀表板調整 1006 #6：未來訂位 JSON 暫存 ----
   原本每次 fn=future 都現讀整張 future 表（實測 6～14 秒）。改成：重建 future（每天 10:15 整份、白天每 3 小時近 14 天）
   寫完就把 JSON 放進 CacheService（切塊，最長 6 小時），網頁讀暫存約 1 秒；暫存沒了才現讀一次再放回。
   回傳多一欄 built_at＝最後一次重建 future 的時間（指令碼屬性 FUT_BUILT_AT），各儀表板拿來標「訂位資料更新於」。 */
var FUT_CACHE_KEY = 'FUT_JSON_v1', FUT_CHUNK = 90000, FUT_TTL = 21600;
function futureJson_() {
  try {
    var c = CacheService.getScriptCache(), n = Number(c.get(FUT_CACHE_KEY + '|n')) || 0;
    if (n > 0) {
      var ks = []; for (var i = 0; i < n; i++) ks.push(FUT_CACHE_KEY + '|' + i);
      var got = c.getAll(ks), s = '';
      for (var j = 0; j < n; j++) { if (got[ks[j]] == null) { s = null; break; } s += got[ks[j]]; }
      if (s) return s;
    }
  } catch (e) {}
  return futureCachePut_();
}
function futureCachePut_() {
  var o = apiSheet_(CFG.FUTURE, null);
  o.built_at = PropertiesService.getScriptProperties().getProperty('FUT_BUILT_AT') || '';
  var js = JSON.stringify(o);
  try {
    var c = CacheService.getScriptCache(), m = {}, parts = Math.ceil(js.length / FUT_CHUNK);
    for (var q = 0; q < parts; q++) m[FUT_CACHE_KEY + '|' + q] = js.slice(q * FUT_CHUNK, (q + 1) * FUT_CHUNK);
    m[FUT_CACHE_KEY + '|n'] = String(parts);
    c.putAll(m, FUT_TTL);
  } catch (e) { Logger.log('future 暫存寫入失敗（不影響資料）：' + e.message); }
  return js;
}
/** 重建 future 之後呼叫：記下重建時間＋立刻換新暫存 */
function futureCacheRefresh_() {
  try {
    PropertiesService.getScriptProperties().setProperty('FUT_BUILT_AT', Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'));
    futureCachePut_();
  } catch (e) { Logger.log('future 暫存換新失敗（不影響資料）：' + e.message); }
}

/* ========== 小工具 ========== */

function collectCookies_(resp, jar) {
  let sc = resp.getAllHeaders()['Set-Cookie'];
  if (!sc) return;
  if (typeof sc === 'string') sc = [sc];
  sc.forEach(c => {
    const kv = c.split(';')[0];
    const eq = kv.indexOf('=');
    if (eq > 0) jar[kv.slice(0, eq).trim()] = kv.slice(eq + 1).trim();
  });
}
function jarStr_(jar) {
  return Object.keys(jar).map(k => k + '=' + jar[k]).join('; ');
}
function attr_(tag, name) {
  const m = tag.match(new RegExp(name + '="([^"]*)"'));
  return m ? m[1] : '';
}
function stripTags_(s) {
  return s.replace(/^<td\b[^>]*>/, '').replace(/<\/td>$/, '')
          .replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
}
function decodeEnt_(s) {
  return String(s).replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&nbsp;/g, ' ')
    .replace(/&#(\d+);/g, (m, d) => String.fromCharCode(Number(d))).trim();
}
function toNum_(v) { const n = Number(String(v).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
function fmtDate_(d) { return Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd'); }
function parseDate_(s) { const p = s.split('-'); return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])); }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd');
  return String(v).trim();
}
function normSlot_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CFG.TZ, 'HH:mm');
  const s = String(v).trim();
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (m) return ('0' + m[1]).slice(-2) + ':' + m[2]; // 9:30 → 09:30
  return s;
}

function diagnose() {
  const p = PropertiesService.getScriptProperties();
  const trg = ScriptApp.getProjectTriggers()
    .map(t => t.getHandlerFunction() + '(' + t.getEventType() + ')').join(', ');
  Logger.log('觸發器：' + (trg || '（無）'));
  Logger.log('fact 最後日期：' + lastFactDate_() + '｜今天：' + fmtDate_(new Date()));
  Logger.log('future 上次更新：' + (p.getProperty('FUTURE_AT') || '（無）'));
  Logger.log('cookie 存在：' + (p.getProperty('SESS_COOKIE') ? '是' : '否'));
  const t0 = Date.now();
  try {
    const rows = fetchDayAuto_(fmtDate_(addDays_(new Date(), -1)));
    Logger.log('✅ 抓昨天成功：' + rows.length + ' 筆，耗時 ' + Math.round((Date.now()-t0)/1000) + ' 秒');
  } catch (e) {
    Logger.log('❌ 抓取失敗（' + Math.round((Date.now()-t0)/1000) + ' 秒）：' + e.message);
  }
}

/** ================================================================
 *  DIYBC 後台唯讀探針 v1 —— 自己人資料自動化 PRE-CHECK
 *
 *  性質：完全唯讀。不寫任何試算表、不改任何資料、不建觸發器。
 *  目的：一次驗完五件事
 *        ① GAS 能不能登入 diybc.azurewebsites.net
 *        ② StoreIndex 的 HTML 裡到底有沒有 Line UID
 *        ③ 查詢參數名叫什麼（日期／店別）
 *        ④ 單日筆數是否撞後台 1,000 列上限
 *        ⑤ vipstats 能不能直接解析
 *
 *  用法：
 *    1. 貼進「訂位資料」那個 GAS 專案（它已有 DIYBC_EMAIL/DIYBC_PASSWORD）
 *       ── 或任一專案，但要自行到「專案設定 → 指令碼屬性」補這兩項
 *    2. 執行 probeAll()（第一次會跳授權，按同意）
 *    3. 把「執行記錄」整段複製，貼回對話給我
 *
 *  ⚠️ 這支不會動到訂位管線的任何函式，函式名全加 _p 後綴避免撞名。
 * ================================================================ */

const P_BASE  = 'https://diybc.azurewebsites.net';
const P_LOGIN = ['/Identity/Account/Login', '/Account/Login', '/Login'];
const P_STORE = '/VIPHis/StoreIndex';
const P_STATS = '/VIPHis/vipstats';
const P_TESTDAY = '2026-08-03';   // 抓昨天（完整一天），要換日期改這裡

var _buf_p = [];
function log_p(s) { _buf_p.push(s); Logger.log(s); }

function probeAll() {
  _buf_p = [];
  var t0 = Date.now();
  log_p('===== DIYBC 後台探針 v1 =====');
  log_p('測試日：' + P_TESTDAY);

  var cred = getCred_p();
  if (!cred) return dump_p(t0);

  var cookie = login_p(cred);
  if (!cookie) { log_p('❌ 登入失敗，中止'); return dump_p(t0); }

  probeStoreIndex_p(cookie);
  probeVipStats_p(cookie);
  dump_p(t0);
}

function dump_p(t0) {
  log_p('===== 完成，耗時 ' + ((Date.now() - t0) / 1000).toFixed(1) + ' 秒 =====');
}

/* ---------- 憑證 ---------- */
function getCred_p() {
  var sp = PropertiesService.getScriptProperties();
  var email = sp.getProperty('DIYBC_EMAIL');
  var pwd   = sp.getProperty('DIYBC_PASSWORD');
  if (!email || !pwd) {
    log_p('❌ 找不到指令碼屬性 DIYBC_EMAIL / DIYBC_PASSWORD');
    log_p('   → 專案設定 → 指令碼屬性 → 新增，或把本檔貼進訂位那個專案');
    return null;
  }
  log_p('✅ 憑證已讀取：' + email.replace(/(.{3}).*(@.*)/, '$1***$2'));
  return { email: email, pwd: pwd };
}

/* ---------- Cookie 工具 ---------- */
function pickCookies_p(res) {
  var h = res.getAllHeaders();
  var sc = h['Set-Cookie'] || h['set-cookie'] || [];
  if (!Array.isArray(sc)) sc = [sc];
  return sc.map(function (s) { return String(s).split(';')[0]; }).join('; ');
}
function mergeCookies_p(a, b) {
  var map = {};
  ((a || '') + '; ' + (b || '')).split('; ').forEach(function (p) {
    var i = p.indexOf('=');
    if (i > 0) map[p.slice(0, i)] = p.slice(i + 1);
  });
  return Object.keys(map).map(function (k) { return k + '=' + map[k]; }).join('; ');
}

/* ---------- 登入 ---------- */
function login_p(cred) {
  for (var i = 0; i < P_LOGIN.length; i++) {
    var path = P_LOGIN[i], url = P_BASE + path, res;
    try {
      res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: false });
    } catch (e) { log_p('· ' + path + ' → 例外 ' + e); continue; }

    var code = res.getResponseCode();
    if (code !== 200) { log_p('· ' + path + ' → HTTP ' + code + '，跳過'); continue; }

    var html = res.getContentText();
    if (!/type="password"/i.test(html)) { log_p('· ' + path + ' → 無密碼欄，跳過'); continue; }
    log_p('✅ 登入頁：' + path);

    var cookies = pickCookies_p(res);

    var tok = (html.match(/name="__RequestVerificationToken"[^>]*value="([^"]+)"/i) ||
               html.match(/value="([^"]+)"[^>]*name="__RequestVerificationToken"/i) || [])[1];
    log_p('   token：' + (tok ? '有（' + tok.length + ' 字）' : '❌ 無'));

    var fields = [], re = /<input\b[^>]*>/gi, m;
    while ((m = re.exec(html))) {
      var tag = m[0];
      var nm = (tag.match(/name="([^"]+)"/i) || [])[1];
      var tp = (tag.match(/type="([^"]+)"/i) || [])[1] || 'text';
      if (nm) fields.push({ name: nm, type: tp });
    }
    log_p('   表單欄位：' + JSON.stringify(fields.map(function (f) { return f.name + ':' + f.type; })));

    var uf = (fields.filter(function (f) { return /email|user|account/i.test(f.name); })[0] || {}).name;
    var pf = (fields.filter(function (f) { return f.type.toLowerCase() === 'password'; })[0] || {}).name;
    if (!uf || !pf) { log_p('   ⚠️ 找不到帳號/密碼欄位名'); continue; }
    log_p('   帳號欄=' + uf + '　密碼欄=' + pf);

    var payload = {};
    payload[uf] = cred.email;
    payload[pf] = cred.pwd;
    if (tok) payload['__RequestVerificationToken'] = tok;

    var post = UrlFetchApp.fetch(url, {
      method: 'post', payload: payload,
      headers: { Cookie: cookies },
      followRedirects: false, muteHttpExceptions: true
    });
    var pc = post.getResponseCode();
    var all = mergeCookies_p(cookies, pickCookies_p(post));
    log_p('   POST → HTTP ' + pc);

    if (/AspNetCore/i.test(all)) { log_p('✅ 取得 Identity cookie'); return all; }
    log_p('   ⚠️ 未見 Identity cookie：' + all.slice(0, 150));
  }
  return null;
}

/* ---------- A. StoreIndex ---------- */
function probeStoreIndex_p(cookie) {
  log_p('\n----- A. StoreIndex（消費明細）-----');

  var r0 = UrlFetchApp.fetch(P_BASE + P_STORE, { headers: { Cookie: cookie }, muteHttpExceptions: true });
  var h0 = r0.getContentText();
  log_p('空查詢 HTTP ' + r0.getResponseCode() + '，長度 ' + h0.length);
  if (/type="password"/i.test(h0)) { log_p('❌ 被導回登入頁 → cookie 未生效'); return; }

  // 所有查詢欄位
  var names = [], re = /<(input|select)\b[^>]*name="([^"]+)"[^>]*>/gi, m;
  while ((m = re.exec(h0))) names.push(m[2]);
  log_p('查詢參數：' + JSON.stringify(names.filter(function (v, i, a) { return a.indexOf(v) === i; })));

  // type=date 的欄位名
  var dateF = [], reIn = /<input\b[^>]*>/gi;
  while ((m = reIn.exec(h0))) {
    if (/type="date"/i.test(m[0])) {
      var n = (m[0].match(/name="([^"]+)"/i) || [])[1];
      if (n) dateF.push(n);
    }
  }
  log_p('日期欄位：' + JSON.stringify(dateF));

  // 店別 select 的 options（前 5 個）
  var sel = h0.match(/<select\b[\s\S]*?<\/select>/gi) || [];
  sel.slice(0, 2).forEach(function (s, i) {
    var nm = (s.match(/name="([^"]+)"/i) || [])[1];
    var op = (s.match(/<option[^>]*value="([^"]*)"/gi) || []).slice(0, 5)
             .map(function (o) { return (o.match(/value="([^"]*)"/i) || [])[1]; });
    log_p('select#' + i + ' name=' + nm + ' options(前5)=' + JSON.stringify(op));
  });

  // 用偵測到的日期欄位查單日
  if (dateF.length < 2) { log_p('⚠️ 日期欄位不足 2 個，跳過單日查詢'); return; }
  var q = P_BASE + P_STORE + '?' + encodeURIComponent(dateF[0]) + '=' + P_TESTDAY +
          '&' + encodeURIComponent(dateF[1]) + '=' + P_TESTDAY;
  log_p('\n查詢 URL：' + q);

  var r1 = UrlFetchApp.fetch(q, { headers: { Cookie: cookie }, muteHttpExceptions: true });
  var h1 = r1.getContentText();
  log_p('HTTP ' + r1.getResponseCode() + '，長度 ' + h1.length);

  var trs = h1.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  log_p('🔢 表格列數：' + trs.length + (trs.length >= 999 ? '　⚠️⚠️ 疑似撞 1,000 列上限' : ''));

  // thead
  var thead = (h1.match(/<thead[\s\S]*?<\/thead>/i) || [''])[0];
  var ths = (thead.match(/<th[\s\S]*?<\/th>/gi) || [])
            .map(function (t) { return t.replace(/<[^>]+>/g, '').trim(); });
  log_p('thead：' + JSON.stringify(ths));

  // 🎯 UID 關鍵字全文搜尋
  var hits = [];
  ['uid', 'lineuid', 'line_uid', 'lineId', 'U0', 'data-v'].forEach(function (k) {
    var c = (h1.match(new RegExp(k, 'gi')) || []).length;
    if (c) hits.push(k + '×' + c);
  });
  log_p('🎯 UID 關鍵字命中：' + (hits.length ? hits.join('、') : '❌ 完全沒有'));

  // Line UID 典型長相：U + 32 位十六進位
  var uidLike = h1.match(/\bU[0-9a-f]{32}\b/gi) || [];
  log_p('🎯 U+32碼 樣式命中：' + uidLike.length + ' 個' +
        (uidLike.length ? '　範例 ' + uidLike[0].slice(0, 8) + '…' : ''));

  // 內嵌 JS 資料陣列偵測
  var inline = h1.match(/(var|let|const)\s+\w+\s*=\s*\[\s*\{[\s\S]{0,200}/i);
  log_p('內嵌 JS 陣列：' + (inline ? '✅ 有\n   ' + inline[0].slice(0, 200) : '無'));

  // 第一列資料的原始 HTML
  var first = trs.length > 1 ? trs[1] : (trs[0] || '');
  log_p('\n📄 第一列原始 HTML（前 1500 字）：\n' + first.slice(0, 1500));
}

