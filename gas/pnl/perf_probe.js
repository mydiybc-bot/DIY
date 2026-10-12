/**
 * perf_probe.gs — P&L 損益 GAS 效能探針
 * 2026-08-12  目的:量出每支 API 的時間到底花在哪一段,不要再用猜的。
 *
 * ▍使用方式
 *   1. Apps Script 專案 → 左側「檔案 +」→ 指令碼 → 命名 perf_probe → 貼上本檔
 *   2. Ctrl+S 存檔（★ 不需要重新部署,這支不經過 doGet）
 *   3. 上方函式下拉選 probePnlPerf → 執行
 *   4. 「執行記錄」全選複製給 Claude
 *
 * ▍紅線
 *   · 全檔唯讀:沒有 setValue / appendRow / deleteRow / deleteSheet
 *   · 只寫入 CacheService 的 __probe_ 前綴測試鍵,且用完即刪,不碰任何正式快取鍵
 *   · 不修改任何既有檔案,純新增
 *
 * ▍寫法注意（本專案既有地雷）
 *   全域 const 會造成整個專案靜默 SyntaxError,所有函式一起死。
 *   因此本檔所有設定值一律包在 probeCfg_() 裡,全域不宣告任何變數。
 */

function probeCfg_() {
  return {
    SHEET_ID: '1khrFp_AYp3mEsL02cTRrYu1LNiRTP68aowM5_ClsE_c',
    FACT: 'fact_pnl',
    LEDGER: 'fact_ledger',
    // 模擬儀表板實際查詢:12 店、115/01~115/07
    YM_MIN: 11501,
    YM_MAX: 11507,
    COL_YM: 1,      // A 欄 year_month
    COL_STORE: 5,   // E 欄 store_id
    CACHE_LIMIT: 100 * 1024  // CacheService 單鍵上限約 100KB
  };
}

function probePnlPerf() {
  var C = probeCfg_();
  var t = Date.now();
  var out = [];
  function lap(label) { var d = Date.now() - t; t = Date.now(); out.push(pad_(d) + ' ms  ' + label); return d; }
  function say(s) { out.push('           ' + s); }

  out.push('===== P&L GAS 效能探針  ' + new Date().toLocaleString('zh-TW') + ' =====');
  out.push('');

  /* ---------- 1. 開檔 ---------- */
  var ss = SpreadsheetApp.openById(C.SHEET_ID);
  var tOpen = lap('SpreadsheetApp.openById()   ← 整個試算表的重量都算在這');

  /* ---------- 2. 分頁盤點:哪些分頁在拖累開檔 ---------- */
  var sheets = ss.getSheets();
  lap('getSheets()  共 ' + sheets.length + ' 個分頁');
  var totalCells = 0, rows = [];
  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var r = sh.getLastRow(), c = sh.getLastColumn();
    var cells = r * c;
    totalCells += cells;
    rows.push({ name: sh.getName(), r: r, c: c, cells: cells });
  }
  rows.sort(function (a, b) { return b.cells - a.cells; });
  say('全檔已用儲存格約 ' + fmtN_(totalCells) + ' 格（Google 上限 1,000 萬）');
  say('前 12 大分頁:');
  for (var j = 0; j < Math.min(12, rows.length); j++) {
    say('   ' + padR_(rows[j].name, 30) + fmtN_(rows[j].r) + ' 列 × ' + rows[j].c +
        ' 欄 = ' + fmtN_(rows[j].cells) + ' 格');
  }
  var repl = rows.filter(function (x) { return x.name.indexOf('fact_pnl_replaced') === 0; });
  say('fact_pnl_replaced_* 尚存 ' + repl.length + ' 個,合計 ' +
      fmtN_(repl.reduce(function (a, b) { return a + b.cells; }, 0)) + ' 格');
  lap('分頁盤點');

  /* ---------- 3. 全表讀取:目前 get_fact_data 的做法 ---------- */
  var fact = ss.getSheetByName(C.FACT);
  if (!fact) { out.push('⛔ 找不到分頁 ' + C.FACT); Logger.log(out.join('\n')); return; }
  var lastR = fact.getLastRow(), lastC = fact.getLastColumn();
  lap('fact_pnl getLastRow/getLastColumn  → ' + fmtN_(lastR) + ' 列 × ' + lastC + ' 欄');

  var all = fact.getDataRange().getValues();
  var tFull = lap('★ getDataRange().getValues()  全表讀進記憶體 ' + fmtN_(all.length) + ' 列');

  /* ---------- 4. 只讀必要欄位:能省多少 ---------- */
  var slim = fact.getRange(1, 1, lastR, Math.min(lastC, 12)).getValues();
  var tSlim = lap('對照組:getRange(1,1,' + fmtN_(lastR) + ',12).getValues()');
  say('→ 兩者差 ' + (tFull - tSlim) + ' ms。若差距很小,代表瓶頸不在欄數而在列數');

  /* ---------- 5. 篩選 + 序列化 ---------- */
  var picked = [];
  for (var k = 1; k < all.length; k++) {
    var ym = Number(all[k][C.COL_YM - 1]);
    if (ym >= C.YM_MIN && ym <= C.YM_MAX) picked.push(all[k]);
  }
  lap('JS 篩選 ' + C.YM_MIN + '~' + C.YM_MAX + ' → ' + fmtN_(picked.length) + ' 列');

  var payload = JSON.stringify(picked);
  var bytes = payload.length;
  lap('JSON.stringify  → ' + fmtN_(bytes) + ' bytes (' + (bytes / 1024).toFixed(1) + ' KB)');

  /* ---------- 6. ★ 關鍵:這包塞得進 CacheService 嗎 ---------- */
  out.push('');
  out.push('----- CacheService 可行性（既有 600 秒快取為何沒生效的關鍵） -----');
  var cache = CacheService.getScriptCache();
  var key = '__probe_' + Date.now();
  var putOk = true, putErr = '';
  try { cache.put(key, payload, 60); } catch (e) { putOk = false; putErr = e.message; }
  var back = null;
  try { back = cache.get(key); } catch (e2) { }
  try { cache.remove(key); } catch (e3) { }
  lap('CacheService put + get + remove 測試');

  say('payload ' + (bytes / 1024).toFixed(1) + ' KB / 單鍵上限約 100 KB');
  say('put 是否丟例外: ' + (putOk ? '否' : '是 → ' + putErr));
  say('取回是否成功  : ' + (back ? '是（長度 ' + fmtN_(back.length) + '）' : '★ 否 — 寫進去了但拿不回來'));
  if (bytes > C.CACHE_LIMIT) {
    say('★★ 超過 100KB 上限 → CacheService.put() 會「靜默失敗」,不報錯但存不進去。');
    say('    若既有程式碼是直接 put 整包,那個 600 秒快取等於從來沒生效過,每次都是 miss。');
    say('    解法:切塊存（每塊 90KB）+ 一個索引鍵記塊數,取用時再組回來。');
  } else {
    say('未超過上限 → 這包是可以被快取的,既有快取沒生效是別的原因（key 設計或 TTL）。');
  }

  /* ---------- 7. get_yms 的實際成本 ---------- */
  out.push('');
  out.push('----- get_yms 成本（它只需要回答「有哪些年月」） -----');
  var ymSet = {};
  for (var m = 1; m < all.length; m++) { ymSet[all[m][C.COL_YM - 1]] = 1; }
  var yms = Object.keys(ymSet);
  lap('從已讀進來的全表萃取相異年月 → ' + yms.length + ' 個');
  say('若 get_yms 自己也跑一次 openById + 全表讀取,它的真實成本 ≈ ' +
      fmtN_(tOpen + tFull) + ' ms,只為了回傳 ' + yms.length + ' 個數字。');

  /* ---------- 8. fact_ledger ---------- */
  out.push('');
  var led = ss.getSheetByName(C.LEDGER);
  if (led) {
    say('fact_ledger: ' + fmtN_(led.getLastRow()) + ' 列 × ' + led.getLastColumn() + ' 欄');
    say('（它與 fact_pnl 同檔,會一起算進 openById 的重量）');
  } else {
    say('找不到 fact_ledger 分頁');
  }

  out.push('');
  out.push('===== 結束 =====');
  Logger.log(out.join('\n'));
}

/* ---- 小工具 ---- */
function pad_(n) { var s = String(n); while (s.length < 7) s = ' ' + s; return s; }
function padR_(s, n) { s = String(s); while (s.length < n) s = s + ' '; return s; }
function fmtN_(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }