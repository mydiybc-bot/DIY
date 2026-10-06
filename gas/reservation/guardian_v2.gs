/************************************************************
 * DIYBC 訂位歷史守門員 v2.6   2026-10-06
 * ----------------------------------------------------------
 * 【v2.6 改了什麼】※ 2026-10-06 守門員每 15 分鐘跑 150～362 秒、多半被 6 分鐘硬殺
 *   現象：10/06 16:44 起 guardianWorker_ 幾乎每輪 300 秒以上、「逾時」一大串，但歷史表其實沒有斷層（10/05 已在）。
 *   根因：guardianWorker_ 每輪都呼叫主檔 lastFactDate_()——讀整個日期欄（約 25 萬列）＋逐列 Utilities.formatDate
 *        （JS→Java 橋接；v2.4 就量過這支單獨 161 秒，當時只修了看門狗、沒修守門員）。表越長越慢，終於撐爆 6 分鐘；
 *        每輪還整段占住腳本鎖，cacheWorker_／dailyFutureAndSnapshot 可能因拿不到鎖而略過，也白白吃掉每日觸發器時間。
 *   改法：新增 gdLastFactDate_(昨天)——先讀表尾 GD_TAIL_ROWS 列（新的一天一律 append 在最後），純 JS（wdDateStr_）找最大日期；
 *        表尾最大日期 ≥ 昨天＝一定沒斷層，直接收工（約 1～3 秒）。表尾沒到昨天才退回 wdScanFact_() 全欄掃描（純 JS，約 40 秒），
 *        所以不會因為只看表尾而漏補。lastFactDate_() 本身不動（其他手動函式還在用）。
 *
 * 【本次（v2.5）改了什麼】※ 看門狗第一次寄信，抓到一個之前看不見的故障
 *   現象：2026-08-29 11:00 收到警報「快取停在 8/26 20:23，Sheet 15043 vs
 *        快取 13674」，但③佇列積壓沒報 ⇒ 佇列是空的。
 *   根因：cacheWorker_ 舊寫法是「先 deleteProperty 整個佇列，再做事，
 *        失敗才 setProperty 放回」。但 rebuildCaches_ 是整表掃描，
 *        24 萬列已撐爆 6 分鐘 —— 而 GAS 硬殺不是 exception，catch 抓不到，
 *        「放回佇列」那行永遠不會執行 ⇒ 佇列永久蒸發、快取從此不再更新，
 *        而且因為佇列空了，③佇列積壓也不會示警。完全無聲。
 *        這是 7月/8月同款死法的第四次變形，這次躲在自癒機制自己身上。
 *   改法（三層）：
 *     ⓐ cacheWorker_ 改成「一次只處理一個月，做完才從佇列移除」。
 *        硬殺時佇列還在，下一輪自動重試。
 *     ⓑ 改用 rebuildOneMonth(ym)（分段讀取的輕量版）取代 rebuildCaches_，
 *        並把 rebuildOneMonth 內的 normDate_ 換成純 JS 的 wdDateStr_
 *        （實測：純 JS 掃 22 萬列只要 40 秒，normDate_ 逐列呼叫是主要慢因）。
 *     ⓒ 「先記失敗、成功才扣」的計數法：硬殺與例外都會被計到，
 *        連續 4 輪失敗就寄信並跳過該月，避免卡住整個佇列。
 *     ⓓ 看門狗發現快取落後時，自動 queueMonth() 排入重建 —— 能自己修的
 *        就別叫人動手；下次還收到同一封信，才代表需要人工介入。
 *
 * 【v2.4 改了什麼】※ v2.3 實跑 174 秒，且第②項數錯，這版修掉
 *   實測：①資料斷層 161.2 秒｜②快取 12.6 秒｜③0.1 秒｜④0.9 秒
 *   兩個問題：
 *     ⓐ 真正的重活是第①項——主檔的 lastFactDate_() 單獨就要 161 秒。
 *        v2.3 誤判成第②項，優化錯地方了。
 *     ⓑ v2.3 的「只讀表尾 6 萬列」假設錯了——fact 並非嚴格依日期排序
 *        （歷次回補、刪除重寫會打亂順序），導致 2026-07 只數到 4,329 筆，
 *        對上快取的 12,522 筆變成「差 189%」的假異常。
 *        原本的防呆（檢查讀到的最早日期）擋不住亂序。
 *   改法：
 *     ①② 合併成「一次全表掃描日期欄」，純 JS 迴圈同時算出最大日期（給①）
 *        與各月筆數（給②）。一次 getValues 解決兩件事，不再呼叫
 *        lastFactDate_()，也不再依賴「資料有排序」這個假設。
 *     ④ 加時間容忍：dailyFutureAndSnapshot 排 10:15，11:00 前手動測試
 *        容許 future 落後 1 天，避免假警報。
 *
 * 【v2.3 改了什麼】※ v2.2 裝上後 watchdogTest 逾時，這版修掉
 *   看門狗自己被 6 分鐘硬殺撐爆了——這正是它該防的那種死法。
 *   兩個重活：
 *     ⓐ 第②項把 fact 整個日期欄（24 萬列）讀進來逐列轉換
 *     ⓑ 為了數 rows.length，把兩個月的快取檔（各約 2MB）整個 JSON.parse
 *   改法：
 *     ⓐ 常態改看快取檔的 cached_at 是否新鮮（用正則抓，不 parse）；
 *        只有 cached_at 落後時，才退回筆數比對，且改成「只讀表尾」
 *        （fact 依日期遞增 append，近 2 個月必定在尾端），不夠再自動擴大
 *     ⓑ 日期轉換全部改用純 JS（wdDateStr_ / wdYm_），不呼叫
 *        Utilities.formatDate——那是 JS→Java 橋接，跑 24 萬次會死
 *   另加自保：總時間預算 240 秒，逼近就跳過剩餘項目並在信中明說，
 *   確保看門狗「寧可少檢查一項，也絕不無聲死亡」。
 *
 * 【v2.2 改了什麼】
 *   看門狗新增第 ④ 項：未來訂位（future）新鮮度檢查。
 *   起因：2026-08-27 發現 future 停擺 2 天，但看門狗完全沒示警——
 *        v2 的三項檢查全都只看歷史表與快取，沒有一項在看 future。
 *        這是 7 月停 9 天、8 月停 17 天之後的「第三個沒被監控的死角」。
 *
 *   ⚠️ 同時把看門狗執行時間從 09:00 移到 11:00（重要，不可省略）
 *      dailyFutureAndSnapshot 是 10:15 才跑。若看門狗留在 09:00，
 *      每天早上都會在 future 尚未重建時看到昨天的日期 → 天天寄假警報
 *      → 三天後這封信就會被當成垃圾信忽略，監控等於白做。
 *
 * 【安裝步驟】
 *   A. 開啟 Apps Script 專案「預約系統自動抓資料」
 *   B. 左側點 guardian_v2.gs → 全選舊內容 → 貼上本檔全部內容 → Ctrl+S
 *   C. 函式下拉選 watchdogTest → 執行，看記錄確認四項都跑得動且各項耗時
 *      （若已經跑過 v2.2 的 upgradeWatchdog，觸發器已是 11:00，不必重跑；
 *        不確定的話重跑 upgradeWatchdog 也無妨，它會先清舊的再建）
 *   ※ 本檔不含 doGet，不需重新部署，Web App 網址不變
 *
 * ----------------------------------------------------------
 * 根除目標（7月斷9天、8月斷17天的同款死法，不再發生）：
 *   ① 斷層永不滾大：每 15 分鐘檢查、有洞就補、逐日存進度
 *   ② 快取分離：刷快取拆到獨立執行緒（23萬列重建是硬殺元兇）
 *   ③ 看門狗：每日 11:00 四項體檢，任一異常寄信
 *
 * 依賴既有函式（都在主檔，勿刪）：
 *   fetchDayAuto_ / writeFactDay_ / lastFactDate_ / rebuildCaches_
 *   fmtDate_ / addDays_ / parseDate_ / normDate_ / readCache_ /
 *   publishJson_ / apiRow_ / API_COLS / CFG
 * 舊的 dailyHistoryCatchup 函式可留著不管（觸發器會被本檔清掉）
 ************************************************************/

var GD_FUSE_MS   = 150 * 1000;      // 抓資料保險絲：150 秒後不開新的一天
var GD_DIRTY_KEY = 'GD_DIRTY_YMS';  // 待刷快取的月份清單（守門員寫、快取工讀）

/* ========== 一次性安裝 ========== */
function setupGuardian() {
  var toKill = ['dailyHistoryCatchup', 'autoKeeper_',
                'guardianWorker_', 'cacheWorker_', 'gapWatchdog', 'healthWatchdog'];
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (toKill.indexOf(t.getHandlerFunction()) >= 0) {
      ScriptApp.deleteTrigger(t); removed++;
    }
  });
  ScriptApp.newTrigger('guardianWorker_').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('cacheWorker_').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('healthWatchdog').timeBased().atHour(11).everyDays(1).create();
  Logger.log('已清除舊觸發器 ' + removed + ' 個\n' +
    '已建立：guardianWorker_（每15分，無斷層時1秒收工）\n' +
    '　　　　cacheWorker_（每30分，無待刷月份時1秒收工）\n' +
    '　　　　healthWatchdog（每日11:00，四項體檢，異常才寄信）\n' +
    '注意：dailyFutureAndSnapshot（每日10:15）不受影響，維持原樣');
}

/* ========== ① 守門員：每15分鐘，有洞補洞，逐日存進度 ========== */
function guardianWorker_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;                 // 與回補/快取共用鎖，自動排隊
  try {
    var p = PropertiesService.getScriptProperties();
    fnEnsureTrigger_(p);                            // 2026-10-06：第一次跑到這裡時補建「未來訂位每 3 小時更新」觸發器（只做一次）
    if (p.getProperty('BF_NEXT')) return;          // 大回補進行中，讓路
    var h = Number(Utilities.formatDate(new Date(), CFG.TZ, 'H'));
    if (h >= 4 && h < 7) return;                   // 夜間慢速窗，跳過

    var t0 = Date.now();
    var yest = fmtDate_(addDays_(new Date(), -1));
    var last = gdLastFactDate_(yest);               // v2.6：原本 lastFactDate_()（整欄＋formatDate，150～360 秒）
    if (!last) return;
    var cur = fmtDate_(addDays_(parseDate_(last), 1));
    if (cur > yest) return;                        // 無斷層，收工

    var dirty = {};
    try {
      JSON.parse(p.getProperty(GD_DIRTY_KEY) || '[]')
        .forEach(function (m) { dirty[m] = 1; });
    } catch (e) {}

    var done = 0;
    while (cur <= yest) {
      if (done > 0 && Date.now() - t0 > GD_FUSE_MS) {
        Logger.log('保險絲：本輪補 ' + done + ' 天，剩餘 15 分鐘後續補');
        return;
      }
      var rows = fetchDayAuto_(cur);
      writeFactDay_(cur, rows);                    // 冪等：先刪該日再寫
      dirty[cur.slice(0, 7)] = 1;
      p.setProperty(GD_DIRTY_KEY, JSON.stringify(Object.keys(dirty)));
      Logger.log('補 ' + cur + '：' + rows.length + ' 筆');
      done++;
      cur = fmtDate_(addDays_(parseDate_(cur), 1));
      Utilities.sleep(200);
    }
    Logger.log('已補到昨天，共 ' + done + ' 天');
  } finally { lock.releaseLock(); }
}

/* v2.6：找歷史表最後日期（給守門員用）。先看表尾、純 JS；沒到昨天才全欄掃描。說明見檔頭【v2.6】 */
var GD_TAIL_ROWS = 20000;
function gdLastFactDate_(yest) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  var col = wdDateCol_(sh), n = Math.min(GD_TAIL_ROWS, lastRow - 1);
  var tail = sh.getRange(lastRow - n + 1, col, n, 1).getValues(), max = '';
  for (var i = 0; i < tail.length; i++) { var d = wdDateStr_(tail[i][0]); if (d > max) max = d; }
  if (max && max >= yest) return max;              // 常態：表尾就有昨天 → 沒斷層
  var all = wdScanFact_();                          // 表尾沒到昨天 → 全欄掃一次（純 JS）確認
  return (all && all.maxDate) ? all.maxDate : (max || null);
}

/* ========== ② 快取工：每30分鐘，一次刷一個月 ========== */
/* v2.5 重寫。舊版把整個佇列先領走再做事，撞到 6 分鐘硬殺時佇列會永久蒸發
   （硬殺不是 exception，catch 與 finally 都救不回來）。新版改成：
     · 一次只處理佇列中的第一個月
     · 用 rebuildOneMonth（分段讀取）取代 rebuildCaches_（整表掃描）
     · 做完才把該月從佇列移除 → 硬殺時佇列還在，下一輪自動重試
     · 「先記失敗、成功才扣」：硬殺與例外都會被計數，連續失敗會寄信並跳過   */
var CW_FAIL_KEY  = 'CW_FAIL';   // {ym: 連續失敗次數}
var CW_MAX_FAIL  = 4;           // 連續失敗幾輪就放棄該月並寄信（4 輪 = 2 小時）

function cacheWorker_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var p = PropertiesService.getScriptProperties();
    var yms;
    try { yms = JSON.parse(p.getProperty(GD_DIRTY_KEY) || '[]'); }
    catch (e) { yms = []; }
    if (!yms.length) return;                       // 沒事做，收工

    var ym = yms[0];
    var fails = {};
    try { fails = JSON.parse(p.getProperty(CW_FAIL_KEY) || '{}'); } catch (e) { fails = {}; }

    // 這個月已經連續失敗太多次：移出佇列並寄信，免得卡住排在後面的月份
    if ((fails[ym] || 0) >= CW_MAX_FAIL) {
      yms.shift();
      p.setProperty(GD_DIRTY_KEY, JSON.stringify(yms));
      delete fails[ym];
      p.setProperty(CW_FAIL_KEY, JSON.stringify(fails));
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
        '【DIYBC 訂位管線】' + ym + ' 快取重建連續失敗，已跳過',
        ym + ' 的快取重建連續 ' + CW_MAX_FAIL + ' 輪沒有成功（可能逾時或出錯），' +
        '已從待刷佇列移除，避免卡住其他月份。\n\n' +
        '儀表板上這個月會持續顯示舊數字。\n\n處理：' +
        '\n1. Apps Script 手動執行 rebuildOneMonth(\'' + ym + '\')，看是逾時還是報錯' +
        '\n2. 若是逾時，代表該月資料量已超過單次處理上限，把記錄貼給 Claude');
      Logger.log(ym + ' 連續失敗 ' + CW_MAX_FAIL + ' 次，已跳過並寄信');
      return;
    }

    // 先記一次失敗，成功才扣掉。這樣「被硬殺」也會被計數（硬殺後面的程式都不會跑）
    fails[ym] = (fails[ym] || 0) + 1;
    p.setProperty(CW_FAIL_KEY, JSON.stringify(fails));

    var t0 = Date.now();
    rebuildOneMonth(ym);                           // 輕量版：只讀日期欄定位 + 分段讀取

    yms.shift();
    p.setProperty(GD_DIRTY_KEY, JSON.stringify(yms));
    delete fails[ym];
    p.setProperty(CW_FAIL_KEY, JSON.stringify(fails));
    Logger.log('快取已更新：' + ym + '（耗時 ' + Math.round((Date.now() - t0) / 1000) +
               ' 秒，佇列剩 ' + yms.length + ' 個月）');
  } finally { lock.releaseLock(); }
}

/* ========== 手動急救：一鍵啟動大回補（斷層滾大時用） ========== */
function fixGapNow() {
  var p = PropertiesService.getScriptProperties();
  var last = lastFactDate_();
  if (!last) { Logger.log('讀不到 fact 最後日期，中止'); return; }
  var start = fmtDate_(addDays_(parseDate_(last), 1));
  var yest = fmtDate_(addDays_(new Date(), -1));
  if (start > yest) { Logger.log('無斷層，不需回補'); return; }
  p.setProperty('BF_NEXT', start);
  p.setProperty('BF_END', yest);
  p.deleteProperty('BF_ERR');
  p.setProperty('BF_WIN', '4');
  clearBackfillTriggers_();
  ScriptApp.newTrigger('backfillWorker_').timeBased().everyMinutes(10).create();
  Logger.log('大回補已啟動：' + start + ' ~ ' + yest + '，完成會寄信。' +
             '\n期間守門員自動讓路，回補完成後自動接手。');
}

/************************************************************
 * DIYBC 訂位快取急救工具 v1  2026-08-26
 * ----------------------------------------------------------
 * 用途：fact 已有資料但儀表板沒更新時，單月重建 Drive 快取。
 *      比 rebuildCaches_ 輕：只讀日期欄定位 → 分段讀取目標月 →
 *      避開 24 萬列 × 21 欄的整表掃描（那是逾時元兇）。
 *
 * 用法：
 *   rebuildOneMonth('2026-08')  → 重建任一月
 *   queueMonth('2026-08')       → 丟給 cacheWorker_ 背景處理（不想等的話）
 *   cacheStatus()               → 唯讀對帳各月 Sheet vs 快取筆數
 ************************************************************/

function fixAugCache() { rebuildOneMonth('2026-08'); }

/** 單月快取重建：分段讀取，只碰目標月的列 */
function rebuildOneMonth(ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('格式須為 YYYY-MM');
  var t0 = Date.now();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  var last = sh.getLastRow();
  if (last < 2) { Logger.log('fact 是空的'); return; }

  // 第一步：只讀日期欄（1 欄，很輕），定位目標月的列號＋順便算各月筆數
  // v2.5：日期解析改用純 JS 的 wdDateStr_。normDate_ 內含 Utilities.formatDate，
  // 那是 JS→Java 橋接，逐列呼叫 24 萬次會把 6 分鐘預算吃掉一大半
  // （實測：純 JS 掃 22 萬列只要 40 秒）。這是 cacheWorker_ 逾時的主因之一。
  var dcol = wdDateCol_(sh);
  var dates = sh.getRange(2, dcol, last - 1, 1).getValues();
  var hits = [], cnt = {};
  for (var i = 0; i < dates.length; i++) {
    var d = wdDateStr_(dates[i][0]);
    var m = d.slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(m)) continue;
    cnt[m] = (cnt[m] || 0) + 1;
    if (m === ym) hits.push(i + 2);
  }
  Logger.log(ym + ' 命中 ' + hits.length + ' 列（定位耗時 ' +
             Math.round((Date.now() - t0) / 1000) + ' 秒）');
  if (!hits.length) { Logger.log('該月無資料，中止'); return; }

  // 第二步：把命中的列合併成連續區塊，分段讀取（避開整表 getValues）
  var idx = {};
  CFG.HEADERS.forEach(function (h, i) { idx[h] = i; });
  var rows = [];
  var s = 0;
  while (s < hits.length) {
    var e = s;
    while (e + 1 < hits.length && hits[e + 1] === hits[e] + 1) e++;
    var start = hits[s], len = hits[e] - hits[s] + 1;
    var vals = sh.getRange(start, 1, len, CFG.HEADERS.length).getValues();
    for (var k = 0; k < vals.length; k++) {
      var d2 = wdDateStr_(vals[k][idx.date]);
      if (d2 && d2.slice(0, 7) === ym) rows.push(apiRow_(vals[k], idx, d2));
    }
    s = e + 1;
  }

  // 第三步：寫快取檔＋更新月份清單
  var now = Utilities.formatDate(new Date(), CFG.TZ, 'yyyy-MM-dd HH:mm:ss');
  publishJson_(ym, JSON.stringify({ cols: API_COLS, rows: rows, cached_at: now }));
  var months = Object.keys(cnt).sort().map(function (m) {
    return { ym: m, rows: cnt[m] };
  });
  publishJson_('months', JSON.stringify({ months: months, cached_at: now }));

  Logger.log('✅ ' + ym + ' 快取已更新：' + rows.length + ' 筆，總耗時 ' +
             Math.round((Date.now() - t0) / 1000) + ' 秒' +
             '\n月份清單已同步：' + months.map(function (m) {
               return m.ym + '=' + m.rows;
             }).join(', '));
}

/** 把月份丟進待刷佇列，交給 cacheWorker_ 背景處理 */
function queueMonth(ym) {
  var p = PropertiesService.getScriptProperties();
  var cur;
  try { cur = JSON.parse(p.getProperty(GD_DIRTY_KEY) || '[]'); } catch (e) { cur = []; }
  if (cur.indexOf(ym) < 0) cur.push(ym);
  p.setProperty(GD_DIRTY_KEY, JSON.stringify(cur));
  Logger.log('已排入待刷佇列：' + cur.join(', ') + '（30 分鐘內由 cacheWorker_ 處理）');
}

/** 唯讀：看目前各月快取的內容筆數與時間，跟 Sheet 實際筆數對照 */
function cacheStatus() {
  var raw = readCache_('months');
  if (!raw) { Logger.log('月份清單快取不存在'); return; }
  var m = JSON.parse(raw);
  Logger.log('月份清單快取時間：' + m.cached_at);
  var out = [];
  m.months.forEach(function (x) {
    var j = readCache_(x.ym);
    if (!j) { out.push(x.ym + ' 清單=' + x.rows + ' ｜快取檔=（無）'); return; }
    var o = JSON.parse(j);
    var flag = (o.rows.length === x.rows) ? '✅' : '⚠️ 不一致';
    out.push(x.ym + ' 清單=' + x.rows + ' ｜快取檔=' + o.rows.length +
             ' ｜' + o.cached_at + ' ' + flag);
  });
  Logger.log(out.join('\n'));
}

/************************************************************
 * DIYBC 訂位看門狗 v2.2  2026-08-27
 * ----------------------------------------------------------
 * 演進史（每一項都是被實際事故打出來的）：
 *   v1  只看「Sheet 有沒有資料」
 *       → 8/26 踩雷：資料補到 8/25 但快取停在 8/8，儀表板顯示舊數字，
 *         v1 完全不示警（因為 Sheet 是健康的）
 *   v2  加上快取新鮮度、佇列積壓
 *       → 8/27 踩雷：future 停擺 2 天，三項檢查沒有一項在看 future，
 *         唯一會講的只有儀表板上那行紅字，而且要人剛好打開才看得到
 *   v2.2 加上 future 新鮮度，並把執行時間 09:00 → 11:00
 *
 * 每日 11:00 檢查四件事，任一異常就寄信：
 *   ① 資料斷層：fact 最新日期落後 > 2 天
 *   ② 快取落後：近 2 個月 Sheet 筆數 vs 快取檔筆數差異 > 5%
 *   ③ 佇列卡住：待刷月份積壓超過 12 小時沒被 cacheWorker_ 領走
 *   ④ future 停擺：未來訂位表最早日期不是今天（＝當日 10:15 沒重建成功）
 *
 * ⚠️ 為什麼是 11:00 不是 09:00
 *   dailyFutureAndSnapshot 排在 10:15。看門狗若留在 09:00，
 *   每天都會在 future 尚未重建時看到昨天的日期 → 天天寄假警報 →
 *   信被當垃圾忽略 → 真的出事時沒人看。監控最怕的不是漏報，是狼來了。
 ************************************************************/

var WD_GAP_DAYS   = 2;      // 資料落後幾天算異常
var WD_CACHE_TOL  = 0.05;   // 快取筆數容許誤差（5%）
var WD_QUEUE_HRS  = 12;     // 待刷佇列積壓幾小時算卡住
var WD_QUEUE_AT   = 'GD_DIRTY_AT';   // 佇列最早排入時間
var WD_FUT_MIN_WIN = 30;    // future 視窗少於幾天就提醒（正常應為 35）
var WD_CACHE_MAX_HRS = 26;  // 快取 cached_at 落後幾小時才需要進一步比對筆數
var WD_BUDGET_MS = 240 * 1000;  // 整輪體檢的時間預算（GAS 硬殺是 360 秒，留餘裕）
var WD_FUT_READY_HOUR = 11; // dailyFutureAndSnapshot 排 10:15，這時間前容許 future 落後 1 天

/* ---------- 純 JS 日期工具（不呼叫 Utilities.formatDate）----------
   Utilities.formatDate 是 JS→Java 橋接，單次很快但跑 24 萬次會直接把
   6 分鐘預算吃光。體檢是唯讀掃描，用純 JS 解析就夠。 */
function wdDateStr_(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' +
           ('0' + (v.getMonth() + 1)).slice(-2) + '-' +
           ('0' + v.getDate()).slice(-2);
  }
  var s = String(v == null ? '' : v).trim();
  var m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/.exec(s);
  return m ? (m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2)) : '';
}
function wdYm_(v) {
  var d = wdDateStr_(v);
  return d ? d.slice(0, 7) : '';
}
/** 從快取檔原始字串直接抓 cached_at（不 JSON.parse，省下 2MB 解析）*/
function wdCachedAt_(raw) {
  var m = /"cached_at"\s*:\s*"([^"]+)"/.exec(raw || '');
  return m ? m[1] : '';
}
/** 只在必要時才 parse，數 rows 筆數 */
function wdRowsCount_(raw) {
  try { return ((JSON.parse(raw) || {}).rows || []).length; } catch (e) { return -1; }
}
/** 一次掃描 fact 的日期欄，同時取得「最大日期」與「各月筆數」。
 *  為什麼要全表掃：fact 並非嚴格依日期排序（歷次回補、刪除重寫會打亂順序），
 *  只讀表尾會漏算——v2.3 就是這樣把 2026-07 數成 4,329 筆（實際上萬筆）。
 *  為什麼還是夠快：只讀 1 欄 + 純 JS 解析，不呼叫 Utilities.formatDate，
 *  也不呼叫主檔的 lastFactDate_()（實測那支單獨就要 161 秒）。 */
function wdScanFact_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.FACT);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { maxDate: '', cnt: {}, rows: 0 };
  var col = wdDateCol_(sh);
  var vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  var cnt = {}, maxDate = '', n = 0;
  for (var i = 0; i < vals.length; i++) {
    var d = wdDateStr_(vals[i][0]);
    if (!d) continue;
    n++;
    if (d > maxDate) maxDate = d;
    var m = d.slice(0, 7);
    cnt[m] = (cnt[m] || 0) + 1;
  }
  return { maxDate: maxDate, cnt: cnt, rows: n };
}

/* ========== 安裝：建立／重建 11:00 的看門狗觸發器 ========== */
function upgradeWatchdog() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (f === 'gapWatchdog' || f === 'healthWatchdog') {
      ScriptApp.deleteTrigger(t); removed++;
    }
  });
  ScriptApp.newTrigger('healthWatchdog').timeBased().atHour(11).everyDays(1).create();
  Logger.log('已移除舊看門狗 ' + removed + ' 個\n' +
    '已建立 healthWatchdog（每日 11:00，四項體檢：資料斷層＋快取新鮮度＋佇列積壓＋future 停擺）\n' +
    '⚠️ 時間刻意設在 11:00：dailyFutureAndSnapshot 是 10:15 才跑，\n' +
    '   看門狗若排在它前面，每天都會誤報 future 停擺。\n' +
    '其餘觸發器不受影響：guardianWorker_ / cacheWorker_ / dailyFutureAndSnapshot\n' +
    '接著請執行 watchdogTest 確認四項檢查都跑得動。');
}

/* ========== 看門狗本體 ========== */
function healthWatchdog() {
  var t = {};
  var issues = watchdogScan_(t);
  Logger.log('體檢完成，耗時 ' + (t.__total / 1000).toFixed(1) + ' 秒，異常 ' + issues.length + ' 項');
  if (!issues.length) return;                 // 一切正常，不打擾
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    '【DIYBC 訂位管線警報】發現 ' + issues.length + ' 項異常',
    issues.join('\n\n') +
    '\n\n────────────' +
    '\n※ 標示「已自動排入佇列」的項目不需要你做任何事，30 分鐘內會自己修好。' +
    '\n　 只有「連續兩天收到同一項」才代表自癒失敗、需要人工介入。' +
    '\n\n需要人工判讀時：' +
    '\n1. Apps Script 執行 pipelineHealth（看整體狀態）' +
    '\n2. Apps Script 執行 watchdogTest（唯讀重跑四項檢查，不寄信，會印各項耗時）' +
    '\n3. Apps Script 執行 cacheStatus（看各月快取 vs Sheet 是否對得上）' +
    '\n4. 開執行紀錄，看 guardianWorker_ / cacheWorker_ / dailyFutureAndSnapshot 是否連續失敗' +
    '\n5. 把上述記錄貼給 Claude 判讀');
}

/** 找出未來訂位表的分頁名稱（CFG 常數名不確定時逐一嘗試） */
function wdFutureSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cands = [];
  try {
    if (typeof CFG !== 'undefined' && CFG) {
      ['FUTURE', 'FACT_FUTURE', 'FUT', 'FUTURE_SHEET', 'SHEET_FUTURE'].forEach(function (k) {
        if (CFG[k] && cands.indexOf(CFG[k]) < 0) cands.push(CFG[k]);
      });
    }
  } catch (e) {}
  cands.push('fact_reservations_future');
  for (var i = 0; i < cands.length; i++) {
    var sh = ss.getSheetByName(cands[i]);
    if (sh) return sh;
  }
  return null;
}

/** 找出某分頁的日期欄欄號（讀表頭找 'date'，找不到才退回慣例第 2 欄） */
function wdDateCol_(sh) {
  try {
    var lastCol = sh.getLastColumn();
    if (lastCol > 0) {
      var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      for (var i = 0; i < hdr.length; i++) {
        if (String(hdr[i]).trim() === 'date') return i + 1;
      }
    }
  } catch (e) {}
  try {
    if (typeof CFG !== 'undefined' && CFG && CFG.HEADERS) {
      var j = CFG.HEADERS.indexOf('date');
      if (j >= 0) return j + 1;
    }
  } catch (e) {}
  return 2;
}

/** 唯讀掃描，回傳問題清單（給看門狗與手動測試共用）
 *  timing：可選物件，會把各項耗時（毫秒）寫進去，供 watchdogTest 印出來 */
function watchdogScan_(timing) {
  var today = fmtDate_(new Date());
  var issues = [];
  var t0 = Date.now();
  timing = timing || {};
  function left() { return WD_BUDGET_MS - (Date.now() - t0); }
  function mark(k, st) { timing[k] = Date.now() - st; }

  // ④ future 停擺（只有幾千列，最輕，先跑，確保時間不夠時它一定跑得到）
  var s4 = Date.now();
  try {
    var hourNow = Number(Utilities.formatDate(new Date(), CFG.TZ, 'H'));
    // dailyFutureAndSnapshot 排在 10:15。在那之前，future 表理當還是昨天那份，
    // 此時要求「最早日期＝今天」就是製造假警報 → 11:00 前容許落後 1 天。
    var futTol = (hourNow < WD_FUT_READY_HOUR) ? 1 : 0;
    var fsh = wdFutureSheet_();
    if (!fsh) {
      issues.push('【④ 未來訂位】找不到未來訂位分頁（試過 CFG 常數與 fact_reservations_future）' +
        '\n→ 分頁可能被改名。請確認名稱後，改 wdFutureSheet_() 的候選清單。');
    } else {
      var fn = fsh.getLastRow();
      if (fn < 2) {
        issues.push('【④ 未來訂位】' + fsh.getName() + ' 是空的' +
          '\n→ 執行 dailyFutureAndSnapshot 重建。訂位壓力與營收預估目前全部失效。');
      } else {
        var fcol = wdDateCol_(fsh);
        var fvals = fsh.getRange(2, fcol, fn - 1, 1).getValues();
        var flo = '', fhi = '';
        for (var k = 0; k < fvals.length; k++) {
          var fd = wdDateStr_(fvals[k][0]);
          if (!fd) continue;
          if (!flo || fd < flo) flo = fd;
          if (fd > fhi) fhi = fd;
        }
        if (!flo) {
          issues.push('【④ 未來訂位】' + fsh.getName() + ' 讀不到任何有效日期（日期欄第 ' +
            fcol + ' 欄）\n→ 表格結構可能變了，請人工確認。');
        } else {
          var stale = Math.round((parseDate_(today) - parseDate_(flo)) / 86400000);
          if (stale > futTol) {
            issues.push('【④ 未來訂位】已停止更新 ' + stale + ' 天' +
              '（表內最早日期 ' + flo + '，重建後應為今天 ' + today + '）' +
              '\n→ dailyFutureAndSnapshot 連續 ' + stale + ' 天沒跑成功。' +
              '\n→ 影響：未來訂位壓力、未來 14 天營收預估、當月剩餘天數預估全部系統性低估；' +
              '\n   儀表板上「下個月」的訂位也會少算。' +
              '\n→ 處理：Apps Script 執行 dailyFutureAndSnapshot，再執行 pipelineHealth 確認。');
          } else {
            var win = Math.round((parseDate_(fhi) - parseDate_(today)) / 86400000) + 1;
            if (win < WD_FUT_MIN_WIN) {
              issues.push('【④ 未來訂位】視窗只有 ' + win + ' 天（' + flo + ' ~ ' + fhi +
                '，正常應為 35 天）' +
                '\n→ CFG.FUTURE_DAYS 可能被改小，當月營收預估會缺天數。');
            }
          }
        }
      }
    }
  } catch (e) {
    issues.push('【④ 未來訂位】檢查失敗：' + e.message);
  }
  mark('④未來訂位', s4);

  // ③ 待刷佇列積壓（只讀 Properties，極輕）
  var s3 = Date.now();
  try {
    var p = PropertiesService.getScriptProperties();
    var q = p.getProperty(GD_DIRTY_KEY);
    if (!q || q === '[]') {
      p.deleteProperty(WD_QUEUE_AT);            // 佇列已清空，重置計時
    } else {
      var at = Number(p.getProperty(WD_QUEUE_AT) || 0);
      if (!at) {
        p.setProperty(WD_QUEUE_AT, String(Date.now()));   // 首次觀測到，開始計時
      } else {
        var hrs = (Date.now() - at) / 3600000;
        if (hrs > WD_QUEUE_HRS) {
          issues.push('【③ 佇列卡住】待刷月份 ' + q + ' 已積壓 ' +
            Math.round(hrs) + ' 小時（容許 ' + WD_QUEUE_HRS + ' 小時）' +
            '\n→ cacheWorker_ 可能連續失敗，請查執行紀錄。');
        }
      }
    }
  } catch (e) {
    issues.push('【③ 佇列卡住】檢查失敗：' + e.message);
  }
  mark('③佇列積壓', s3);

  // ①② 合併：一次掃描 fact 日期欄，同時得到「最大日期」與「各月筆數」
  var sf = Date.now();
  var scan = null;
  try {
    if (left() < 90 * 1000) {
      issues.push('【①② fact 掃描】本次體檢時間不足，資料斷層與快取新鮮度未檢查' +
        '\n→ 前兩項耗時異常，請執行 watchdogTest 看各項耗時。');
    } else {
      scan = wdScanFact_();
    }
  } catch (e) {
    issues.push('【①② fact 掃描】失敗：' + e.message);
  }
  mark('①②fact掃描', sf);

  // ① 資料斷層
  var s1 = Date.now();
  try {
    if (scan) {
      var last = scan.maxDate;
      var gap = last
        ? Math.round((parseDate_(today) - parseDate_(last)) / 86400000) - 1
        : 999;
      if (gap > WD_GAP_DAYS) {
        issues.push('【① 資料斷層】fact_reservations 最新 = ' + (last || '讀不到') +
          '，落後 ' + gap + ' 天（容許 ' + WD_GAP_DAYS + ' 天）' +
          '\n→ 守門員可能卡住。斷層若已滾大，執行 fixGapNow() 啟動大回補。');
      }
    }
  } catch (e) {
    issues.push('【① 資料斷層】檢查失敗：' + e.message);
  }
  mark('①資料斷層', s1);

  // ② 快取新鮮度：拿上面掃出來的各月筆數，跟快取檔的筆數對照
  var s2 = Date.now();
  try {
    if (scan) {
      var thisYm = today.slice(0, 7);
      var prevYm = fmtDate_(addDays_(parseDate_(today.slice(0, 8) + '01'), -1)).slice(0, 7);
      [thisYm, prevYm].forEach(function (ym) {
        var want = scan.cnt[ym] || 0;
        if (!want) return;                        // 該月 Sheet 本來就沒資料，跳過
        var raw = readCache_(ym);
        if (!raw) {
          var q1 = false;
          try { queueMonth(ym); q1 = true; } catch (e) {}
          issues.push('【② 快取新鮮度】' + ym + ' 快取檔不存在，Sheet 有 ' + want + ' 筆' +
            (q1 ? '\n→ ✅ 已自動排入待刷佇列，30 分鐘內由 cacheWorker_ 重建。'
                : '\n→ ⚠️ 自動排隊失敗，請手動執行 rebuildOneMonth(\'' + ym + '\')。'));
          return;
        }
        var ca = wdCachedAt_(raw) || '（讀不到時間）';
        var n = wdRowsCount_(raw);
        if (n < 0) {
          issues.push('【② 快取新鮮度】' + ym + ' 快取檔解析失敗（快取時間 ' + ca + '）' +
            '\n→ 檔案可能寫壞了。執行 rebuildOneMonth(\'' + ym + '\') 重建。');
          return;
        }
        var diff = Math.abs(n - want) / want;
        if (diff > WD_CACHE_TOL) {
          // v2.5：能自己修的就別叫人動手 —— 直接排入待刷佇列，
          // cacheWorker_ 30 分鐘內會重建。信裡只是告知，不是待辦事項。
          var queued = false;
          try { queueMonth(ym); queued = true; } catch (e) {}
          issues.push('【② 快取新鮮度】' + ym + '：Sheet ' + want + ' 筆 vs 快取 ' + n +
            ' 筆（差 ' + Math.round(diff * 100) + '%），快取時間 ' + ca +
            '\n→ 儀表板正顯示' + (n > want ? '多餘的舊資料' : '舊數字') + '。' +
            (queued
              ? '\n→ ✅ 已自動排入待刷佇列，30 分鐘內由 cacheWorker_ 重建，你不需要做任何事。' +
                '\n→ 若明天還收到同一封信，才代表自癒失敗，需要人工執行 rebuildOneMonth(\'' + ym + '\')。'
              : '\n→ ⚠️ 自動排隊失敗，請手動執行 rebuildOneMonth(\'' + ym + '\')。'));
        }
      });
    }
  } catch (e) {
    issues.push('【② 快取新鮮度】檢查失敗：' + e.message);
  }
  mark('②快取新鮮度', s2);

  timing.__total = Date.now() - t0;
  return issues;
}

/* ========== 手動測試（唯讀，隨時可跑，不寄信） ========== */
function watchdogTest() {
  var t = {};
  var issues = watchdogScan_(t);
  var lines = [];
  Object.keys(t).forEach(function (k) {
    if (k !== '__total') lines.push('　' + k + '：' + (t[k] / 1000).toFixed(1) + ' 秒');
  });
  var timing = '\n\n──── 各項耗時（總計 ' + (t.__total / 1000).toFixed(1) + ' 秒 / 預算 ' +
               (WD_BUDGET_MS / 1000) + ' 秒）────\n' + lines.join('\n');
  if (!issues.length) {
    Logger.log('✅ 四項檢查全數通過：資料無斷層、快取新鮮、佇列暢通、未來訂位今日已重建\n' +
               '（看門狗今天不會寄信）' + timing);
  } else {
    Logger.log('⚠️ 發現 ' + issues.length + ' 項異常（實際執行時會寄信）：\n\n' +
               issues.join('\n\n') + timing);
  }
}

/* ========== 唯讀：只看第 ④ 項，方便單獨確認 future 狀態 ========== */
function futureStatus() {
  var today = fmtDate_(new Date());
  var fsh = wdFutureSheet_();
  if (!fsh) { Logger.log('找不到未來訂位分頁'); return; }
  var fn = fsh.getLastRow();
  if (fn < 2) { Logger.log(fsh.getName() + ' 是空的'); return; }
  var fcol = wdDateCol_(fsh);
  var vals = fsh.getRange(2, fcol, fn - 1, 1).getValues();
  var lo = '', hi = '', n = 0;
  for (var i = 0; i < vals.length; i++) {
    var d = wdDateStr_(vals[i][0]);
    if (!d) continue;
    n++;
    if (!lo || d < lo) lo = d;
    if (d > hi) hi = d;
  }
  var stale = Math.round((parseDate_(today) - parseDate_(lo)) / 86400000);
  var win = Math.round((parseDate_(hi) - parseDate_(today)) / 86400000) + 1;
  Logger.log('分頁：' + fsh.getName() + '（日期欄第 ' + fcol + ' 欄）' +
    '\n涵蓋：' + lo + ' ~ ' + hi + '（' + win + ' 天 / ' + n + ' 筆）' +
    '\n新鮮度：' + (stale > 0 ? '⚠️ 已停止更新 ' + stale + ' 天' : '✅ 今日已重建') +
    '\n今天：' + today);
}