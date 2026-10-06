/************************************************************
 * DIYBC 訂位管線修復 ＋ 營收預估基礎建設   v1.1  2026-08-01
 * ----------------------------------------------------------
 * ⚠️ 整檔取代：把 forecast_pipeline 內容全選刪除後貼上本檔。
 *    v1.0 已安裝者照做即可，setupForecastTriggers 會自動汰換舊觸發器。
 *
 * ── 實測診斷（2026-08-01）────────────────────────────────
 * 【v1.0 已修好】fact_reservations_future 自 7/23 停擺 9 天
 *   根因：dailyFetchReservations 在同一個 6 分鐘預算內做四件重活，
 *         150 秒保險絲天天把 rebuildFuture_ 跳過。
 *   修法：拆成獨立觸發器。實測結果 8/01~8/15 共 2,874 筆 ✅
 *
 * 【v1.1 本次修】fact_reservations 停在 2026-07-29
 *   證據：觸發器有跑（8/1 08:15）、錯誤率 50%、但一封失敗信都沒有。
 *         程式碼裡有 sendEmail，代表 catch 得到的錯誤都會寄信；
 *         既然沒信 → 不是一般錯誤，而是 (a) 6 分鐘硬殺逾時（無法 catch，
 *         所以不寄信），或 (b) 240 秒回看保險絲提前中止，函式「正常結束」
 *         卻一天都沒寫入。
 *   修法：歷史回補也拆成獨立觸發器，每 2 小時自我檢查，
 *         **一次抓一天、抓完立刻寫**，被硬殺也不會回退。
 *
 * ── 安裝步驟（約 90 秒）──────────────────────────────────
 *  A. 貼上本檔 → Ctrl+S
 *  B. 執行 setupForecastTriggers（自動汰換舊觸發器）
 *  C. 【改一個數字】編輯器搜尋 FUTURE_DAYS，把 14 改成 35 → Ctrl+S
 *       這是既有 CFG 設定。改完未來視窗涵蓋 36 天，月營收預估才有完整資料。
 *       實測 15 天需 50 秒，36 天約 120 秒，rebuildFuture_ 保險絲 300 秒，安全。
 *  D. 【刪舊觸發器】左側「觸發條件」→ dailyFetchReservations → ⋮ → 刪除
 *       它的工作已被 dailyHistoryCatchup ＋ dailyFutureAndSnapshot 完全接管；
 *       留著只會重複跑並繼續逾時。函式本身保留不刪，仍可手動執行。
 *  E. 執行 dailyHistoryCatchup（立刻補 7/30、7/31）
 *  F. 執行 dailyFutureAndSnapshot（用新的 36 天視窗重建）
 *  G. 執行 pipelineHealth → 看記錄檔驗收
 *  ※ 本檔不含 doGet，不需重新部署，Web App 網址不變。
 ************************************************************/

var FP_SS_ID    = '13NI3vGV4MSsngeO_DecVYKrOzky-fXCsN9ActscJorQ';
var FP_FUTURE   = 'fact_reservations_future';
var FP_FACT     = 'fact_reservations';
var FP_SNAPSHOT = 'fact_resv_snapshot';
var FP_SNAP_HEAD = ['snapshot_date','target_date','lead_days','store',
                    'bookings_valid','party_size','order_total','bookings_cancelled'];

var FP_CATCHUP_FUSE_MS = 240 * 1000;   // 回補保險絲：240 秒後不再開新的一天
var FP_MAX_GAP_DAYS    = 30;           // 單次最多補幾天（防呆）


/* ========== 歷史回補：一次一天，抓完立刻寫 ========== */
function dailyHistoryCatchup() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('另一個執行中，略過'); return '略過'; }
  try {
    var t0 = Date.now();
    var tz = Session.getScriptTimeZone();
    var yesterday = Utilities.formatDate(fpAddDays_(new Date(), -1), tz, 'yyyy-MM-dd');
    var last = fpNorm_(lastFactDate_(), tz);

    if (!last) { Logger.log('無法判斷 fact 最後日期，中止'); return '無法判斷'; }
    if (last >= yesterday) { Logger.log('無斷層（最後 ' + last + '）'); return '無斷層'; }

    var cur = fpNextDay_(last), touched = {}, done = [], hitFuse = false, guard = 0;
    while (cur <= yesterday && guard++ < FP_MAX_GAP_DAYS) {
      if (Date.now() - t0 > FP_CATCHUP_FUSE_MS) { hitFuse = true; break; }
      var rows = fetchDayAuto_(cur);      // 自行處理登入，只需日期字串
      writeFactDay_(cur, rows);           // 冪等：內部先刪該日再寫
      touched[cur.slice(0, 7)] = true;
      done.push(cur + '(' + (rows ? rows.length : 0) + ')');
      cur = fpNextDay_(cur);
      Utilities.sleep(400);
    }

    if (done.length) {
      try { refreshCaches_(touched); }
      catch (e) { Logger.log('快取重建失敗（不影響已寫入的資料）：' + e); }
    }

    var msg = '回補 ' + done.length + ' 天：' + (done.join(', ') || '無')
            + (hitFuse ? '\n⏱ 觸及保險絲，剩餘由下一輪（2 小時後）自動續補'
                       : '\n✅ 已補到昨天');
    Logger.log(msg);
    return msg;
  } catch (e) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '[DIYBC] 訂位歷史回補失敗', String(e && e.stack || e));
    throw e;
  } finally { lock.releaseLock(); }
}

/* ========== 每日：重建 future ＋ 存 pickup 快照 ========== */
function dailyFutureAndSnapshot() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('另一個執行中，略過'); return; }
  try {
    var t0 = Date.now();          // ★全新計時起點，保險絲不會一開始就跳
    var cookie = login_();
    rebuildFuture_(cookie, t0);   // 沿用既有函式，完全不修改
    futureCacheRefresh_();        // 2026-10-06：記下重建時間、換新 fn=future 暫存（失敗只記 log）
    var n = snapshotFuture_();
    Logger.log('future 重建完成，快照寫入 ' + n + ' 列');
  } catch (e) {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
      '[DIYBC] 訂位 future/快照 失敗', String(e && e.stack || e));
    throw e;
  } finally { lock.releaseLock(); }
}

/* ========== 2026-10-06 各儀表板調整 1006 #6：未來訂位白天每 3 小時更新（近 14 天） ==========
   經營者：訂位分析評估每 3 小時更新一次，但不可以拖累儀表板。做法：
   ・觸發器每 3 小時叫一次 refreshFutureNear；夜間（23～07 時，後台慢）與 10 點（10:15 有整份重建＋快照）跳過，白天約跑 5 次。
   ・只重抓今天起 FN_DAYS 天（14 次查詢，約 1 分鐘），其餘日期沿用 10:15 那份；抓完才寫，超過保險絲就整輪放棄（沿用上一版）。
   ・不寫 fact_resv_snapshot（進榜曲線是用每天 10:15 那一次校準的，白天多寫會讓營收預估偏掉）。
   ・寫完 futureCacheRefresh_()：網頁讀暫存（約 1 秒），不會因為更新變慢。
   安裝：執行一次 setupFutureNearTrigger（可重跑，會先刪舊的再建）。 */
var FN_DAYS = 14, FN_FUSE_MS = 240 * 1000;
function refreshFutureNear() {
  var h = Number(Utilities.formatDate(new Date(), CFG.TZ, 'H'));
  if (h < 7 || h >= 23 || h === 10) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Logger.log('另一個執行中，這輪略過'); return; }
  try {
    var t0 = Date.now(), fresh = [], days = {}, today = fmtDate_(new Date());
    for (var i = 0; i < FN_DAYS; i++) {
      if (Date.now() - t0 > FN_FUSE_MS) { Logger.log('保險絲：' + i + ' 天後時間不夠，這輪不寫（沿用上一版）'); return; }
      var d = fmtDate_(addDays_(new Date(), i));
      days[d] = 1;
      fresh.push.apply(fresh, fetchDayAuto_(d));
      Utilities.sleep(200);
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FUTURE);
    if (!sh) throw new Error('找不到分頁 ' + CFG.FUTURE);
    var last = sh.getLastRow(), keep = [], di = CFG.HEADERS.indexOf('date');
    if (last > 1) {
      sh.getRange(2, 1, last - 1, CFG.HEADERS.length).getValues().forEach(function (r) {
        var d = wdDateStr_(r[di]);
        if (d && d >= today && !days[d]) keep.push(r);   // 14 天以外的照舊；今天以前的丟掉；14 天內換成剛抓的
      });
    }
    var rows = fresh.concat(keep);
    if (last > 1) sh.getRange(2, 1, last - 1, CFG.HEADERS.length).clearContent();   // 全部抓完才動表，中斷不留半套
    if (rows.length) sh.getRange(2, 1, rows.length, CFG.HEADERS.length).setValues(rows);
    try { futureMenuWrite_(fresh, days); } catch (e) { Logger.log('甜點明細寫入失敗（不影響 future）：' + e.message); }   // 1006 #2
    futureCacheRefresh_();
    Logger.log('近 ' + FN_DAYS + ' 天重抓 ' + fresh.length + ' 筆＋沿用 ' + keep.length + ' 筆，耗時 ' + Math.round((Date.now() - t0) / 1000) + ' 秒');
  } finally { lock.releaseLock(); }
}
/** 守門員第一次跑到時呼叫：沒有 refreshFutureNear 觸發器就建一個，另外預約 1 分鐘後跑一次 refreshFutureNearOnce 當驗證；
    做過就記在指令碼屬性 FN_TRIGGER_AT，之後不再碰（經營者若刻意刪掉觸發器，不會被自動建回來） */
function fnEnsureTrigger_(p) {
  try {
    if (p.getProperty('FN_TRIGGER_AT')) return;
    var ts = ScriptApp.getProjectTriggers();
    if (!ts.some(function (t) { return t.getHandlerFunction() === 'refreshFutureNear'; })) ScriptApp.newTrigger('refreshFutureNear').timeBased().everyHours(3).create();
    if (!ts.some(function (t) { return t.getHandlerFunction() === 'refreshFutureNearOnce'; })) ScriptApp.newTrigger('refreshFutureNearOnce').timeBased().after(60 * 1000).create();
    p.setProperty('FN_TRIGGER_AT', Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'));
    Logger.log('已自動建立 refreshFutureNear 每 3 小時觸發器＋1 分鐘後驗證一次');
  } catch (e) { Logger.log('自動建立 refreshFutureNear 觸發器失敗：' + e.message); }
}
/** 2026-10-06 1006 #2：甜點明細分頁第一次上線，守門員跑到時預約 1 分鐘後跑一次近 14 天更新，先把 fact_future_menu 寫出來（只做一次，FM_INIT_AT） */
function fmEnsureInit_(p) {
  try {
    if (p.getProperty('FM_INIT_AT')) return;
    if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'refreshFutureNearOnce'; })) ScriptApp.newTrigger('refreshFutureNearOnce').timeBased().after(60 * 1000).create();
    p.setProperty('FM_INIT_AT', Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm'));
  } catch (e) { Logger.log('甜點明細首次更新預約失敗：' + e.message); }
}
/** 一次性：跑一次 refreshFutureNear，先把自己的觸發器刪掉（跑完不留） */
function refreshFutureNearOnce() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshFutureNearOnce') ScriptApp.deleteTrigger(t); });
  refreshFutureNear();
}
function setupFutureNearTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'refreshFutureNear') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('refreshFutureNear').timeBased().everyHours(3).create();
  Logger.log('已建立 refreshFutureNear 每 3 小時觸發器（夜間與 10 點自動跳過）');
}

/* ========== 把今天的 future 表聚合成快照 append 進去 ========== */
function snapshotFuture_() {
  var ss  = SpreadsheetApp.openById(FP_SS_ID);
  var fut = ss.getSheetByName(FP_FUTURE);
  if (!fut) throw new Error('找不到分頁 ' + FP_FUTURE);

  var vals = fut.getDataRange().getValues();
  if (vals.length < 2) { Logger.log('future 表是空的，不寫快照'); return 0; }

  var head = vals[0].map(function (x) { return String(x).trim(); });
  var iD  = head.indexOf('date'),  iS = head.indexOf('store'),
      iSt = head.indexOf('status'), iP = head.indexOf('party_size'),
      iO  = head.indexOf('order_total');
  if (iD < 0 || iS < 0 || iSt < 0) throw new Error('future 欄位結構改變，請人工確認');

  var tz    = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var agg = {};
  for (var r = 1; r < vals.length; r++) {
    var d  = fpNorm_(vals[r][iD], tz);
    var st = String(vals[r][iS] || '').trim();
    if (!d || !st) continue;
    var k = d + '|' + st;
    var o = agg[k] || (agg[k] = { v: 0, p: 0, ot: 0, c: 0 });
    if (String(vals[r][iSt]).trim() === '已取消') { o.c++; }
    else { o.v++; o.p += Number(vals[r][iP]) || 0; o.ot += Number(vals[r][iO]) || 0; }
  }

  var sh = ss.getSheetByName(FP_SNAPSHOT);
  if (!sh) {
    sh = ss.insertSheet(FP_SNAPSHOT);
    sh.getRange(1, 1, 1, FP_SNAP_HEAD.length).setValues([FP_SNAP_HEAD]);
    sh.setFrozenRows(1);
  }
  fpPurgeSnapshotDate_(sh, today);   // 同日重跑維持冪等

  var t0ms = new Date(today + 'T00:00:00').getTime();
  var rows = Object.keys(agg).map(function (k) {
    var p = k.split('|'), d = p[0], st = p[1], o = agg[k];
    var lead = Math.round((new Date(d + 'T00:00:00').getTime() - t0ms) / 86400000);
    return [today, d, lead, st, o.v, o.p, o.ot, o.c];
  }).filter(function (r) { return r[2] >= 0; })
    .sort(function (a, b) { return a[2] - b[2] || (a[3] < b[3] ? -1 : 1); });

  if (!rows.length) return 0;
  var start = sh.getLastRow() + 1;
  // ★ 文字欄先設格式，避免 setValues 把 "2026-08-01" 自動轉成 Date（既有地雷）
  sh.getRange(start, 1, rows.length, 2).setNumberFormat('@');
  sh.getRange(start, 4, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, FP_SNAP_HEAD.length).setValues(rows);
  return rows.length;
}

function fpPurgeSnapshotDate_(sh, ymd) {
  var last = sh.getLastRow();
  if (last < 2) return;
  var tz = Session.getScriptTimeZone();
  var all = sh.getRange(2, 1, last - 1, FP_SNAP_HEAD.length).getValues();
  var keep = all.filter(function (r) { return fpNorm_(r[0], tz) !== ymd; });
  if (keep.length === all.length) return;
  sh.getRange(2, 1, all.length, FP_SNAP_HEAD.length).clearContent();
  if (keep.length) {
    sh.getRange(2, 1, keep.length, 2).setNumberFormat('@');
    sh.getRange(2, 4, keep.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, keep.length, FP_SNAP_HEAD.length).setValues(keep);
  }
}

/* ========== 小工具（一律加 fp 前綴，不與既有函式衝突） ========== */
function fpNorm_(v, tz) {
  tz = tz || Session.getScriptTimeZone();
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  var s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}
function fpAddDays_(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function fpNextDay_(ymd) {
  return Utilities.formatDate(fpAddDays_(new Date(ymd + 'T00:00:00'), 1),
                              Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ========== 健康檢查（唯讀，隨時可跑） ========== */
function pipelineHealth() {
  var ss = SpreadsheetApp.openById(FP_SS_ID), tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var out = ['=== DIYBC 訂位管線健康檢查 ' + today + ' ==='];

  var last = fpNorm_(lastFactDate_(), tz);
  var gap  = Math.round((new Date(today) - new Date(last)) / 86400000) - 1;
  out.push('fact_reservations 最新：' + last + '（落後 ' + gap + ' 天）' + (gap > 0 ? '  ⚠️' : '  ✅'));

  var fut = ss.getSheetByName(FP_FUTURE);
  if (fut) {
    var fv = fut.getDataRange().getValues();
    var fi = fv[0].map(function (x) { return String(x).trim(); }).indexOf('date');
    var lo = '9999', hi = '';
    for (var i = 1; i < fv.length; i++) {
      var d = fpNorm_(fv[i][fi], tz); if (!d) continue;
      if (d < lo) lo = d; if (d > hi) hi = d;
    }
    var span  = Math.round((new Date(hi) - new Date(lo)) / 86400000) + 1;
    var stale = Math.round((new Date(today) - new Date(lo)) / 86400000);
    out.push('future：' + lo + ' ~ ' + hi + '（' + span + ' 天 / ' + (fv.length - 1) + ' 筆）');
    out.push('future 新鮮度：' + (stale > 0 ? '⚠️ 已停止更新 ' + stale + ' 天' : '✅ 今日已重建'));
    out.push('future 視窗：' + span + ' 天'
      + (span >= 30 ? '  ✅ 足以做月營收預估'
                    : '  ⚠️ 請把 CFG.FUTURE_DAYS 改成 35（步驟 C）'));
  }

  var sn = ss.getSheetByName(FP_SNAPSHOT);
  if (!sn) { out.push('fact_resv_snapshot：尚未建立'); }
  else {
    var days = {}, sv = sn.getRange(2, 1, Math.max(1, sn.getLastRow() - 1), 1).getValues();
    sv.forEach(function (r) { var d = fpNorm_(r[0], tz); if (d) days[d] = 1; });
    var n = Object.keys(days).length;
    out.push('快照累積：' + n + ' 天 / ' + (sn.getLastRow() - 1) + ' 列');
    out.push(n < 28 ? '  → 還需 ' + (28 - n) + ' 天才能校準分店級進榜曲線'
                    : '  ✅ 已可校準分店級進榜曲線');
  }

  var tg = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  out.push('觸發器：' + tg.join(', '));
  if (tg.indexOf('dailyFetchReservations') >= 0)
    out.push('  ⚠️ 舊觸發器 dailyFetchReservations 仍在，請刪除（步驟 D）');

  Logger.log(out.join('\n'));
  return out.join('\n');
}

function fixAugGap() {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('BF_NEXT', '2026-08-09');
  props.setProperty('BF_END', fmtDate_(addDays_(new Date(), -1)));
  props.deleteProperty('BF_ERR');
  props.setProperty('BF_WIN', '4');
  clearBackfillTriggers_();
  ScriptApp.newTrigger('backfillWorker_').timeBased().everyMinutes(10).create();
  Logger.log('8月斷層回補已啟動：2026-08-09 ~ 昨天，完成會寄信');
}