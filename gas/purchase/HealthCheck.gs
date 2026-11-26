/*****************************************************************
 * HealthCheck.gs — 採購系統每日健康檢查 v1.0（2026-09-17）
 *
 * 目的：每天早上 EventLayer（07:31）重算完，自動把這次 10 輪檢查抓到的問題類型再查一遍，
 *       有問題就寄信給經營者。不用再靠人工「再檢查一次」才發現錯。
 *
 * 只讀不寫：除了自己的 health_log 分頁（BOM 本，最多留 90 列），不改任何資料表。
 *
 * 入口（GAS「執行」下拉選單）：
 *   aaaHealthCheckNow()            立即檢查；有 FAIL 才寄信；回傳文字報告（看執行記錄）
 *   aaaHealthCheckTestMail()       立即檢查且一定寄信（第一次測試收信用）
 *   aaaHealthCheckInstallTrigger() 安裝每日 08 點檢查的觸發器（重複執行會先刪掉舊的，不會重複）
 *   dailyHealthCheck()             觸發器呼叫：有 FAIL 當天寄；只有 WARN 的話每週一寄一次摘要
 *
 * 檢查項目（FAIL＝叫貨量可能已經錯、要當天處理；WARN＝要留意、週一彙整）
 *   H1 資料新鮮度   POS 到昨天、用量彙總／收尾表／檔期剩餘推估是今天重算的                     （FAIL）
 *   H2 兩檔一致     EventLayer 五張表 BOM 本＝專用檔逐格相同；用量彙總／主檔／店別參數列數相同   （FAIL／WARN）
 *   H3 欄位型態     收尾表 10 欄表頭、日期欄是日期、比例欄是 0～1 的數字（份數佔比被當日期的事故）  （FAIL）
 *   H4 收尾合理性   直接抓 GitHub 上網頁「真正在用」的收尾算法（endCutOf）來算：係數在 0～1、
 *                   已下架超過一週且沒有其他批次快下架的料，POS 近 7 天照常在用卻被扣到 ×0.7 以下 ⇒ 網頁算法被改壞；
 *                   另與本檔內建算法比對，不同步提醒更新本檔                                        （FAIL／WARN）
 *   H5 前後端配對   GitHub 上的網頁有沒有對應目前後台版本的關鍵程式（新後台＋舊網頁的事故）        （FAIL）
 *   H6 收尾失效     有甜點即將下架，收尾表卻是空的（收尾計算可能出錯）                              （FAIL）
 *   H7 檔期建檔     去年這時候開賣的檔期，今年對照表還沒建；今年將開賣的甜點沒有 BOM                （WARN）
 *   H8 盤點覆蓋     從未盤點、或超過 14 天沒盤點的店                                               （WARN）
 *   H9 主檔品質     有效品項單價 0、採購單位＝使用單位但內容量 >1、沒填廠商                         （WARN）
 *
 * ★ H4 以網頁的 endCutOf 為準（每天從 GitHub 抓）；本檔 hcEndCut_ 只是讀不到網頁時的備援與比對基準。
 *   改網頁 endCutOf 後若 H4 出現「算法不同步」提醒，要同步更新 hcEndCut_（skill 地雷 165）。
 *
 * ★ 權限：第一次執行會要求授權「寄送電子郵件」與「連線到外部服務」（讀 GitHub 上的網頁）。
 *   若專案的 appsscript.json 有明列 oauthScopes，要加上
 *   https://www.googleapis.com/auth/script.send_mail 與 https://www.googleapis.com/auth/script.external_request
 *****************************************************************/

var HC_BOM_ID = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';
var HC_DASH_ID = '1FF7lW3JINR0-Id7MMYSRkoRzA1BYdqktO94NbzAFYG0';
var HC_FRONT_URL = 'https://raw.githubusercontent.com/mydiybc-bot/DIY/main/static/dashboard-purchase.html';

var HC_CFG = {
  MAIL_TO: '',                 // 空白＝寄給安裝觸發器的帳號
  TRIGGER_HOUR: 8,             // 每天 8 點檢查（EventLayer 07:31 之後）
  LOG_TAB: 'health_log',
  LOG_KEEP: 90,
  DEFAULT_LEAD: 3,             // 與網頁 DEFAULT_LEAD 相同
  OVERCUT_USE: 0.8,            // 近 7 天用量 ≥ 平時的 80%
  OVERCUT_FACTOR: 0.7,         // 卻被扣到 ×0.7 以下 ⇒ 照常在用卻被大扣
  SOURCE_GAP: 0.5,             // 主引擎上週比例與 POS 近 7 天比相差 ≥50 個百分點 ⇒ 兩個用量來源對不上
  PS_ZERO_MAX: 0.05,           // 份數佔比為 0 的列超過 5% ⇒ 欄位格式可能又被改成日期
  CAMP_LOOKAHEAD: 21,          // 去年這天之後 21 天內開賣的檔期要提醒建檔
  LEADIN_DAYS: 14,             // 與 EventLayer LEADIN_DAYS 相同
  INV_STALE_DAYS: 14,
  ENDCUT_HEAD: ['店號', 'sku_id', '結束日', '佔比', '近4週週用量', '甜點', '更新日', '份數佔比', '在售甜點', '近7天比'],
  MIRROR_STRICT: ['dim_event', 'agg_newitem', 'agg_evcurve', 'agg_evneed', 'agg_endcut'],
  MIRROR_COUNT: ['agg_purchase', 'dim_sku', 'dim_store_par'],
  FRONT_MARKERS: [
    ["select A,B,C,D,F,H,I,J", '收尾表讀「近7天比」（配對 EventLayer v1.8.4）'],
    ['function buildNewAgg', '多檔期新品備料加總（ca 批）'],
    ['function endCutOf', '收尾扣減（bx 批以後）'],
    ['本週訂貨單合計（已盤點）', '總覽金額與訂貨單同口徑（ca 批）']
  ]
};

/* ===================== 入口 ===================== */

function aaaHealthCheckNow() { return hcExecute_({ forceMail: false, weekly: false }); }
function aaaHealthCheckTestMail() { return hcExecute_({ forceMail: true, weekly: true }); }
function dailyHealthCheck() { return hcExecute_({ forceMail: false, weekly: new Date().getDay() === 1 }); }

function aaaHealthCheckInstallTrigger() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyHealthCheck') { ScriptApp.deleteTrigger(t); removed++; }
  });
  ScriptApp.newTrigger('dailyHealthCheck').timeBased().everyDays(1).atHour(HC_CFG.TRIGGER_HOUR).inTimezone('Asia/Taipei').create();
  var msg = '已安裝每日健康檢查觸發器（每天 ' + HC_CFG.TRIGGER_HOUR + ' 點～' + (HC_CFG.TRIGGER_HOUR + 1) + ' 點之間執行）；刪除舊觸發器 ' + removed + ' 個';
  Logger.log(msg);
  return msg;
}

function hcExecute_(opt) {
  var now = new Date();
  var res;
  try {
    res = hcRun_(hcGasReader_(), now);
  } catch (e) {
    res = { fails: [{ code: 'H0', msg: '健康檢查本身執行失敗：' + e + '（系統可能有表格被改名或刪除，請找 Claude 查）' }], warns: [], infos: [], stats: {} };
  }
  var text = hcReport_(res, now);
  var send = opt.forceMail || res.fails.length > 0 || (opt.weekly && res.warns.length > 0);
  if (send) {
    try {
      MailApp.sendEmail(hcMailTo_(), hcSubject_(res, now), text);
      text += '\n\n（已寄信）';
    } catch (eM) {
      text += '\n\n⚠️ 寄信失敗：' + eM;
    }
  }
  try { hcLog_(res, now); } catch (eL) { text += '\n⚠️ 寫 health_log 失敗：' + eL; }
  Logger.log(text);
  return text;
}

function hcMailTo_() {
  if (HC_CFG.MAIL_TO) return HC_CFG.MAIL_TO;
  var e = '';
  try { e = Session.getEffectiveUser().getEmail(); } catch (x) {}
  if (!e) { try { e = Session.getActiveUser().getEmail(); } catch (y) {} }
  return e;
}

function hcSubject_(res, now) {
  var d = (now.getMonth() + 1) + '/' + now.getDate();
  if (res.fails.length) return '【採購系統健康檢查】❌ ' + res.fails.length + ' 項需要處理（' + d + '）';
  if (res.warns.length) return '【採購系統健康檢查】⚠️ ' + res.warns.length + ' 項留意（' + d + ' 週報）';
  return '【採購系統健康檢查】✅ 全部正常（' + d + '）';
}

/* ===================== 讀取層（GAS） ===================== */

function hcGasReader_() {
  var files = {}, cache = {};
  function ss(id) { return files[id] = files[id] || SpreadsheetApp.openById(id); }
  return {
    values: function (id, tab) {
      var k = id + '|' + tab;
      if (cache[k] !== undefined) return cache[k];
      var sh = ss(id).getSheetByName(tab);
      cache[k] = sh ? sh.getDataRange().getValues() : null;
      return cache[k];
    },
    maxDate: function (id, tab, header) {
      var sh = ss(id).getSheetByName(tab); if (!sh) return null;
      var lc = sh.getLastColumn(), lr = sh.getLastRow(); if (lr < 2) return null;
      var h = sh.getRange(1, 1, 1, lc).getValues()[0].map(function (x) { return String(x || '').trim(); });
      var c = h.indexOf(header); if (c < 0) return null;
      var col = sh.getRange(2, c + 1, lr - 1, 1).getValues(), mx = null;
      for (var i = 0; i < col.length; i++) { var d = hcD_(col[i][0]); if (d && (!mx || d > mx)) mx = d; }
      return mx;
    },
    fetch: function (url) {
      try { var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true }); return r.getResponseCode() === 200 ? r.getContentText() : null; }
      catch (e) { return null; }
    }
  };
}

function hcLog_(res, now) {
  var ss = SpreadsheetApp.openById(HC_BOM_ID);
  var sh = ss.getSheetByName(HC_CFG.LOG_TAB);
  var H = ['檢查時間', '結果', 'FAIL 數', 'WARN 數', '摘要'];
  if (!sh) {
    sh = ss.insertSheet(HC_CFG.LOG_TAB);
    sh.getRange(1, 1, 1, H.length).setValues([H]);
    try {
      var mc = sh.getMaxColumns(); if (mc > H.length) sh.deleteColumns(H.length + 1, mc - H.length);
      var mr = sh.getMaxRows(); if (mr > HC_CFG.LOG_KEEP + 10) sh.deleteRows(HC_CFG.LOG_KEEP + 11, mr - HC_CFG.LOG_KEEP - 10);
    } catch (e) {}
  }
  var summary = res.fails.concat(res.warns).map(function (x) { return x.code + ' ' + x.msg; }).join('｜').slice(0, 45000);
  var status = res.fails.length ? 'FAIL' : (res.warns.length ? 'WARN' : 'OK');
  var last = sh.getLastRow();
  if (sh.getMaxRows() < last + 1) sh.insertRowsAfter(sh.getMaxRows(), 10);
  sh.getRange(last + 1, 1, 1, H.length).setValues([[now, status, res.fails.length, res.warns.length, summary]]);
  var extra = sh.getLastRow() - 1 - HC_CFG.LOG_KEEP;
  if (extra > 0) sh.deleteRows(2, extra);
}

/* ===================== 檢查本體（純計算，不直接呼叫 SpreadsheetApp，方便測試） ===================== */

function hcRun_(rd, now) {
  var today = hcDateOnly_(now), todayY = hcYmd_(today), yY = hcYmd_(hcAdd_(today, -1));
  var fails = [], warns = [], infos = [], stats = {};
  function F(code, msg) { fails.push({ code: code, msg: msg }); }
  function W(code, msg) { warns.push({ code: code, msg: msg }); }

  var ec = hcObjs_(rd.values(HC_DASH_ID, 'agg_endcut'));
  var ecRaw = rd.values(HC_DASH_ID, 'agg_endcut');
  var ap = hcObjs_(rd.values(HC_DASH_ID, 'agg_purchase'));
  var nd = hcObjs_(rd.values(HC_DASH_ID, 'agg_evneed'));
  var ni = hcObjs_(rd.values(HC_DASH_ID, 'agg_newitem'));
  var par = hcObjs_(rd.values(HC_DASH_ID, 'dim_store_par'));
  var sku = hcObjs_(rd.values(HC_DASH_ID, 'dim_sku'));
  var inv = hcObjs_(rd.values(HC_DASH_ID, 'fact_inventory'));
  var map = hcObjs_(rd.values(HC_BOM_ID, '產品名稱對照表'));
  var bom = hcObjs_(rd.values(HC_BOM_ID, 'BOM表'));
  stats.endcutRows = ec ? ec.length : null;

  /* ---------- H1 資料新鮮度 ---------- */
  var posMax = rd.maxDate(HC_BOM_ID, 'POS資料', '建立日期');
  if (!posMax) F('H1', 'POS資料 讀不到「建立日期」');
  else if (hcYmd_(posMax) < yY) F('H1', 'POS 銷售資料只到 ' + hcYmd_(posMax) + '（應到昨天 ' + yY + '）：POS 同步沒跟上，今天所有建議量都用舊銷量');
  stats.posMax = posMax ? hcYmd_(posMax) : '';
  var apMax = hcMaxDateOf_(ap, '重算時間');
  if (!ap) F('H1', '專用檔 agg_purchase 讀不到');
  else if (!apMax || hcYmd_(apMax) !== todayY) F('H1', '用量彙總 agg_purchase 最後重算 ' + (apMax ? hcYmd_(apMax) : '無') + '，今天 07:05 沒有重算成功');
  if (ec && ec.length) {
    var ecMax = hcMaxDateOf_(ec, '更新日');
    if (!ecMax || hcYmd_(ecMax) !== todayY) F('H1', '收尾表 agg_endcut 更新日 ' + (ecMax ? hcYmd_(ecMax) : '無') + '，今天 07:31 EventLayer 沒有重算成功');
  }
  if (nd && nd.length) {
    var ndMax = hcMaxDateOf_(nd, '更新日');
    if (!ndMax || hcYmd_(ndMax) !== todayY) F('H1', '檔期剩餘推估 agg_evneed 更新日 ' + (ndMax ? hcYmd_(ndMax) : '無') + '，不是今天');
  }

  /* ---------- H2 兩檔一致 ---------- */
  HC_CFG.MIRROR_STRICT.forEach(function (tab) {
    var a = rd.values(HC_BOM_ID, tab), b = rd.values(HC_DASH_ID, tab);
    if (!a || !b) { F('H2', tab + '：' + (!a ? 'BOM 本' : '專用檔') + '找不到這張表'); return; }
    var d = hcFirstDiff_(a, b);
    if (d) F('H2', tab + '：BOM 本與專用檔不一致（' + d + '），網頁讀到的是專用檔');
  });
  HC_CFG.MIRROR_COUNT.forEach(function (tab) {
    var a = rd.values(HC_BOM_ID, tab), b = rd.values(HC_DASH_ID, tab);
    if (!a || !b) { W('H2', tab + '：' + (!a ? 'BOM 本' : '專用檔') + '找不到這張表'); return; }
    var na = hcDataRows_(a), nb = hcDataRows_(b);
    if (na !== nb) W('H2', tab + '：BOM 本 ' + na + ' 列、專用檔 ' + nb + ' 列（若 08 點前有人正在寫入屬暫時現象，隔天仍不同再處理）');
  });

  /* ---------- H3 欄位型態 ---------- */
  if (!ecRaw) F('H3', '專用檔沒有 agg_endcut 分頁');
  else {
    var head = (ecRaw[0] || []).map(function (x) { return String(x || '').trim(); });
    var want = HC_CFG.ENDCUT_HEAD;
    if (head.slice(0, want.length).join('|') !== want.join('|')) F('H3', 'agg_endcut 表頭不對：現在是「' + head.filter(String).join('、') + '」，應為「' + want.join('、') + '」（EventLayer 版本可能被蓋回舊版）');
    if (ec && ec.length) {
      var badDate = 0, badNum = 0, psZero = 0, r7Null = 0, r7Zero = 0, ex = '';
      ec.forEach(function (r) {
        if (!(r['結束日'] instanceof Date) || !(r['更新日'] instanceof Date)) { badDate++; if (!ex) ex = '店' + r['店號'] + ' ' + r['sku_id'] + ' 日期欄不是日期'; }
        ['佔比', '份數佔比', '近7天比'].forEach(function (c) {
          var v = r[c];
          if (c === '近7天比' && (v === '' || v == null)) { r7Null++; return; }
          if (typeof v !== 'number' || !isFinite(v) || v < 0 || v > 1.0005) { badNum++; if (!ex) ex = '店' + r['店號'] + ' ' + r['sku_id'] + ' ' + c + '＝' + hcShow_(v); }
        });
        if (r['份數佔比'] === 0) psZero++;
        if (r['近7天比'] === 0) r7Zero++;
      });
      if (badDate) F('H3', 'agg_endcut 有 ' + badDate + ' 列日期欄不是日期（' + ex + '）');
      if (badNum) F('H3', 'agg_endcut 有 ' + badNum + ' 格比例不是 0～1 的數字（' + ex + '）：欄位格式可能被改成日期，網頁會讀成 0，收尾扣減會算錯');
      if (psZero > ec.length * HC_CFG.PS_ZERO_MAX) F('H3', 'agg_endcut「份數佔比」有 ' + psZero + '/' + ec.length + ' 列是 0：大家都用的耗材會被多扣（9/18 事故同型）');
      if (r7Null > 0) F('H3', 'agg_endcut「近7天比」有 ' + r7Null + ' 列空白：網頁會退回主引擎比例，照常在用的料可能被多扣');
      if (ec.length >= 50 && r7Zero === ec.length) F('H3', 'agg_endcut「近7天比」全部是 0：EventLayer 近 7 天用量計算可能出錯');
    }
  }
  if (ap) {
    var apBad = 0;
    ap.forEach(function (r) { ['上週用量', '近4週均'].forEach(function (c) { var v = r[c]; if (v !== '' && v != null && (typeof v !== 'number' || !isFinite(v))) apBad++; }); });
    if (apBad) F('H3', 'agg_purchase 有 ' + apBad + ' 格用量不是數字：網頁會當成 0，建議量會變少');
  }
  if (ni) {
    var niBad = ni.filter(function (r) { return typeof r['需求量'] !== 'number' || !isFinite(r['需求量']) || r['需求量'] < 0; }).length;
    if (niBad) F('H3', 'agg_newitem 有 ' + niBad + ' 列需求量不是 ≥0 的數字');
  }

  /* ---------- H4 收尾合理性（以網頁真正在用的 endCutOf 計算） ---------- */
  var front = rd.fetch(HC_FRONT_URL);
  var feCut = front ? hcFrontEndCut_(front) : null;
  if (ec && ec.length && ap) {
    var aggMap = {}; ap.forEach(function (r) { aggMap[String(r['店號']) + '|' + r['sku_id']] = r; });
    var lead = {}; (par || []).forEach(function (r) { lead[String(r['店號']) + '|' + r['sku_id']] = r['交期天']; });
    var grp = {};
    ec.forEach(function (r) {
      var k = String(r['店號']) + '|' + r['sku_id'];
      (grp[k] = grp[k] || []).push({ end: hcYmd_(hcD_(r['結束日']) || today), share: +r['佔比'] || 0,
        ps: (typeof r['份數佔比'] === 'number') ? r['份數佔比'] : null, oa: String(r['在售甜點'] || ''),
        r7p: (typeof r['近7天比'] === 'number') ? r['近7天比'] : null, who: String(r['甜點'] || '') });
    });
    var skuName = {}; (sku || []).forEach(function (r) { skuName[r['sku_id']] = r['品名']; });
    var over = [], badF = 0, gap = [], shareOver = [], mism = [], pastLim = hcYmd_(hcAdd_(today, -7));
    Object.keys(grp).forEach(function (k) {
      var L = grp[k], a = aggMap[k] || {}, p = k.split('|');
      var nm = (skuName[p[1]] || p[1]);
      var tot = 0; L.forEach(function (x) { tot += x.share; }); if (tot > 1.001) shareOver.push('店' + p[0] + ' ' + nm + ' ' + Math.round(tot * 100) + '%');
      var ld = (lead[k] === '' || lead[k] == null) ? HC_CFG.DEFAULT_LEAD : +lead[k];
      var week = +a['上週用量'] || 0, avg4 = +a['近4週均'] || 0;
      var mine = hcEndCut_(L, today, ld, 1, week, avg4, false);
      var e = mine, src = '本檔';
      if (feCut) {
        try { e = feCut(p[0], p[1], L, today, ld, week, avg4); src = '網頁'; } catch (eF) { e = mine; src = '本檔（網頁算法執行失敗：' + eF + '）'; }
        if (e && mine && Math.abs(e.factor - mine.factor) > 0.001) mism.push('店' + p[0] + ' ' + nm + ' 網頁 ×' + e.factor.toFixed(3) + ' vs 本檔 ×' + mine.factor.toFixed(3));
      }
      if (!e) return;
      if (!(e.factor >= 0 && e.factor <= 1)) badF++;
      /* 照常在用卻被大扣：已下架批次都超過 7 天、涵蓋期內沒有其他批次下架 ⇒ 扣減只能來自已下架批次，應該被 POS 近 7 天實況校正 */
      var cov = 7 + ld, covEnd = hcYmd_(hcAdd_(today, cov - 1));
      var onlyOldPast = L.every(function (x) { return x.end < pastLim || x.end > covEnd; }) && L.some(function (x) { return x.end < pastLim; });
      var rP = null; L.forEach(function (x) { if (x.r7p != null && (rP == null || x.r7p > rP)) rP = x.r7p; });
      if (onlyOldPast && rP != null && rP >= HC_CFG.OVERCUT_USE && e.factor <= HC_CFG.OVERCUT_FACTOR)
        over.push('店' + p[0] + ' ' + nm + '（POS 近 7 天是平時 ' + Math.round(rP * 100) + '%，網頁卻扣到 ×' + e.factor.toFixed(2) + '）');
      var rA = avg4 > 0 ? Math.min(1, week / avg4) : null;
      if (rA != null && rP != null && Math.abs(rA - rP) >= HC_CFG.SOURCE_GAP) gap.push('店' + p[0] + ' ' + nm + '（主引擎 ' + Math.round(rA * 100) + '% vs POS ' + Math.round(rP * 100) + '%）');
    });
    stats.endcutPairs = Object.keys(grp).length;
    stats.h4Source = feCut ? '網頁 endCutOf' : '本檔備援算法';
    if (!feCut) W('H4', front ? '從 GitHub 網頁找不到收尾算法 endCutOf，改用本檔備援算法檢查（網頁可能大改，請找 Claude 確認）' : '讀不到 GitHub 網頁，改用本檔備援算法檢查收尾合理性');
    if (badF) F('H4', '收尾係數有 ' + badF + ' 組不在 0～1');
    if (shareOver.length) F('H4', '有 ' + shareOver.length + ' 組店×料的下架用料佔比加總超過 100%（' + shareOver.slice(0, 3).join('、') + '）：收尾計算出錯');
    if (over.length) F('H4', '有 ' + over.length + ' 組「照常在用卻被大扣」：' + over.slice(0, 5).join('、') + (over.length > 5 ? '…' : '') + '。網頁收尾算法可能被改壞，或近7天比讀不到');
    if (mism.length) W('H4', '健康檢查內建算法與網頁不同步 ' + mism.length + ' 組（' + mism.slice(0, 2).join('、') + '）：檢查仍以網頁為準，請找 Claude 更新 HealthCheck.gs');
    if (gap.length) W('H4', '有 ' + gap.length + ' 組主引擎上週用量與 POS 近 7 天差很多（' + gap.slice(0, 3).join('、') + (gap.length > 3 ? '…' : '') + '）：多半是主引擎週界落後，網頁已取兩者較高者；持續出現再查 BOM／POS 對照');
  }

  /* ---------- H5 前後端配對 ---------- */
  if (!front) W('H5', '讀不到 GitHub 上的網頁，無法檢查前後端版本是否配對');
  else {
    var miss = HC_CFG.FRONT_MARKERS.filter(function (m) { return front.indexOf(m[0]) < 0; }).map(function (m) { return m[1]; });
    if (miss.length) F('H5', 'GitHub 上的網頁缺少：' + miss.join('、') + '。網頁可能被蓋回舊版，與目前後台不配對（舊網頁會讀錯新欄位）');
    stats.frontBytes = front.length;
  }

  /* ---------- H6 收尾失效 ---------- */
  var soon = [];
  (map || []).forEach(function (r) {
    var en = hcD_(r['結束有效日']); if (!en) return;
    if (en >= today && en <= hcAdd_(today, 35)) soon.push(String(r['手動對應產品名稱'] || r['POS 資料產品名稱'] || '').trim());
  });
  if (soon.length && ec && ec.length === 0) F('H6', '有 ' + soon.length + ' 款甜點 35 天內下架（' + soon.slice(0, 3).join('、') + '），收尾表卻是空的：EventLayer 收尾計算可能失敗（看執行記錄有沒有「收尾扣減計算失敗」）');

  /* ---------- H7 檔期建檔提醒 ---------- */
  var camps = {};
  (map || []).forEach(function (r) {
    var m = String(r['次類別'] || '').trim().match(/^(\d{4})\s*(.+)$/); if (!m) return;
    var st = hcD_(r['起始有效日']), en = hcD_(r['結束有效日']); if (!st || !en) return;
    var k = m[1] + ' ' + m[2], o = camps[k] = camps[k] || { year: +m[1], name: m[2], start: st, items: {} };
    if (st < o.start) o.start = st;
    var nm = String(r['手動對應產品名稱'] || r['POS 資料產品名稱'] || '').trim(); if (nm) o.items[nm] = st;
  });
  var thisY = today.getFullYear(), bomNames = {};
  (bom || []).forEach(function (r) { var d = String(r['甜點名稱'] || '').trim(); if (d) bomNames[d] = 1; });
  Object.keys(camps).forEach(function (k) {
    var o = camps[k];
    if (o.year === thisY - 1) {
      var shifted = new Date(thisY, o.start.getMonth(), o.start.getDate());
      if (shifted >= today && shifted <= hcAdd_(today, HC_CFG.CAMP_LOOKAHEAD) && !camps[thisY + ' ' + o.name])
        W('H7', '去年「' + o.name + '」' + (o.start.getMonth() + 1) + '/' + o.start.getDate() + ' 開賣，今年對照表還沒建：系統會在開賣前 ' + HC_CFG.LEADIN_DAYS + ' 天開始備料，沒建檔各店就不會預備這檔的料');
    }
    if (o.year === thisY && o.start > today && o.start <= hcAdd_(today, HC_CFG.LEADIN_DAYS)) {
      var noBom = Object.keys(o.items).filter(function (nm) { return !bomNames[nm]; });
      if (noBom.length) W('H7', '「' + o.name + '」即將開賣，但 ' + noBom.length + ' 款甜點沒有 BOM（' + noBom.slice(0, 3).join('、') + '）：這些甜點的料不會被備到');
    }
  });

  /* ---------- H8 盤點覆蓋 ---------- */
  if (inv) {
    var byStore = {};
    inv.forEach(function (r) {
      var st = String(r['店號']); var d = hcD_(r['時間戳']) || hcD_(r['盤點日']);
      if (!st || r['現有量'] === '' || r['現有量'] == null) return;
      var o = byStore[st] = byStore[st] || { n: 0, last: null };
      o.n++; if (d && (!o.last || d > o.last)) o.last = d;
    });
    var never = [], stale = [];
    for (var s = 1; s <= 12; s++) {
      var o2 = byStore[String(s)];
      if (!o2 || !o2.n) never.push(s + ' 號店');
      else if (o2.last && (today - hcDateOnly_(o2.last)) / 864e5 > HC_CFG.INV_STALE_DAYS) stale.push(s + ' 號店（' + Math.round((today - hcDateOnly_(o2.last)) / 864e5) + ' 天）');
    }
    if (never.length) W('H8', '從未盤點：' + never.join('、') + '。這些店的建議量不可信（現有量當 0）');
    if (stale.length) W('H8', '超過 ' + HC_CFG.INV_STALE_DAYS + ' 天沒盤點：' + stale.join('、'));
  }

  /* ---------- H9 主檔品質 ---------- */
  if (sku) {
    var active = sku.filter(function (r) {
      if (!r['sku_id'] || String(r['狀態'] || '').trim()) return false;
      var lv = hcD_(r['最後有效日']); return !(lv && lv < today);
    });
    var p0 = active.filter(function (r) { return !(+r['單價'] > 0) && !/半成品/.test(String(r['廠商'] || '')); });
    var unit = active.filter(function (r) { return String(r['使用單位'] || '').trim() === String(r['採購單位'] || '').trim() && +r['每採購單位內容量'] > 1; });
    var noV = active.filter(function (r) { return !String(r['廠商'] || '').trim(); });
    if (p0.length) W('H9', '有效品項單價為 0：' + p0.length + ' 項（' + p0.slice(0, 5).map(function (r) { return r['品名']; }).join('、') + '）：訂貨金額會少算');
    if (unit.length) W('H9', '採購單位＝使用單位但內容量 >1：' + unit.length + ' 項（' + unit.slice(0, 5).map(function (r) { return r['品名'] + ' ' + r['採購單位'] + '×' + r['每採購單位內容量']; }).join('、') + '）：建議量會被讀錯倍數');
    if (noV.length) W('H9', '有效品項沒填廠商：' + noV.length + ' 項（' + noV.slice(0, 5).map(function (r) { return r['品名']; }).join('、') + '）');
  }

  /* ---------- H10 收銀品名自動補對（2026-11-26，posNameFix.gs 每天重算前跑，寫 posname_fix_log）：今天有新增或需人工就列出 ---------- */
  var pnf = hcObjs_(rd.values(HC_BOM_ID, 'posname_fix_log'));
  if (pnf) {
    var pnfToday = pnf.filter(function (r) { var d = hcD_(r['時間']); return !!d && hcYmd_(d) === todayY; });
    var pAdd = pnfToday.filter(function (r) { return String(r['結果']).indexOf('已新增') >= 0; });
    var pMan = pnfToday.filter(function (r) { return String(r['結果']).indexOf('新增') < 0; });
    if (pAdd.length) W('H10', '對照表收銀品名自動補對 ' + pAdd.length + ' 列（' + pAdd.slice(0, 5).map(function (r) { return r['原 POS 品名(A)'] + ' → ' + r['POS 實際品名']; }).join('、') + '）：已自動新增對照列，用量會算進去；請確認對得沒錯');
    if (pMan.length) W('H10', '對照表有 ' + pMan.length + ' 個甜點收銀品名對不上、要人工處理（' + pMan.slice(0, 5).map(function (r) { return r['原 POS 品名(A)'] + '：' + r['結果']; }).join('、') + '）：沒處理前這些甜點的用料不會算進訂貨');
  }
  return { fails: fails, warns: warns, infos: infos, stats: stats };
}

/* 從網頁原始碼取出 endCutOf 放進沙盒執行（網頁是收尾算法的唯一真相）。回傳 function(店號, sku_id, rows, today, lead, week, avg4) → 網頁的計算結果 */
function hcFrontEndCut_(html) {
  var src = hcExtractFn_(html, 'endCutOf'); if (!src) return null;
  var make;
  try {
    make = new Function('env',
      'var ENDCUT={},TODAY=new Date(),ST={period:"近4週"};' +
      'function isDur(s){return false;}' +
      'function num(x){return (x===null||x===undefined||x===""||isNaN(+x))?0:+x;}' +
      'function ymd(d){return d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")+"-"+String(d.getDate()).padStart(2,"0");}' +
      'function aggOf(st,s){return env.agg;}' +
      src + '\n' +
      'return function(st,id,rows,today,lead){ENDCUT={};ENDCUT[st]={};ENDCUT[st][id]=rows;TODAY=today;return endCutOf(st,{id:id,cat:"食材"},lead,1);};');
  } catch (e) { return null; }
  return function (st, id, rows, today, lead, week, avg4) {
    var env = { agg: { week: week, avg4: avg4 } };
    var run = make(env);
    return run(String(st), String(id), rows, today, lead);
  };
}

/* 取出 function name(...){...} 完整原始碼（大括號配對，略過字串與註解） */
function hcExtractFn_(src, name) {
  var i = src.indexOf('function ' + name + '('); if (i < 0) return null;
  var j = src.indexOf('{', i), d = 0, q = null, esc = false;
  for (var k = j; k < src.length; k++) {
    var c = src.charAt(k);
    if (q) { if (esc) esc = false; else if (c === '\\') esc = true; else if (c === q) q = null; continue; }
    if (c === '"' || c === "'") { q = c; continue; }
    if (c === '/' && src.charAt(k + 1) === '*') { var z = src.indexOf('*/', k + 2); if (z < 0) return null; k = z + 1; continue; }
    if (c === '{') d++;
    else if (c === '}') { d--; if (d === 0) return src.slice(i, k + 1); }
  }
  return null;
}

/* 與 dashboard-purchase.html endCutOf 同算法（係數、貨到已結束、實況校正）；另回傳「只算已下架超過 7 天那批」的係數供 H4 用 */
function hcEndCut_(L, today, lead, cyc, week, avg4, weekMode) {
  if (!L || !L.length) return null;
  var ld = Math.max(0, +lead || 0), cov = Math.max(1, Math.round((+cyc || 1) * 7 + ld)), t0 = hcDateOnly_(today);
  var arrD = new Date(t0.getFullYear(), t0.getMonth(), t0.getDate() + ld), arr = hcYmd_(arrD), rows = [], oa = '';
  L.forEach(function (x) {
    var p = String(x.end).split('-'), e = new Date(+p[0], (+p[1]) - 1, +p[2]); if (isNaN(e)) return;
    rows.push({ end: x.end, e: e, share: x.share, ps: (x.ps == null ? null : x.ps), oa: String(x.oa || ''), r7p: (x.r7p == null ? null : x.r7p) });
  });
  if (!rows.length) return null;
  rows.sort(function (a, b) { return a.end < b.end ? -1 : 1; });
  var lastEnd = rows[rows.length - 1].end; rows.forEach(function (r) { if (r.end === lastEnd && r.oa && !oa) oa = r.oa; });
  var hasPs = rows.every(function (r) { return r.ps != null; });
  var shareTot = 0; rows.forEach(function (r) { shareTot += r.share; }); if (!hasPs && shareTot < 0.95) return null;
  function fAt(ended) {
    var S = 0, P = 0; rows.forEach(function (r) { if (ended(r)) { S += r.share; P += (r.ps || 0); } }); S = Math.min(1, S); if (S <= 0) return 1;
    if (!hasPs || P >= 0.95) return Math.max(0, 1 - S); return Math.max(0, Math.min(1, (1 - S) / (1 - P)));
  }
  var rA = (+avg4 > 0) ? Math.min(1, (+week || 0) / +avg4) : null, rP = null;
  rows.forEach(function (r) { if (r.r7p != null && (rP == null || r.r7p > rP)) rP = r.r7p; });
  var r7 = (rA == null && rP == null) ? null : Math.max(rA == null ? 0 : rA, rP == null ? 0 : rP);
  var pastLim = new Date(t0.getFullYear(), t0.getMonth(), t0.getDate() - 7);
  var isPast = function (r) { return r.e < pastLim; }, hasPast = rows.some(isPast), emp = false;
  var fPast = hasPast ? fAt(isPast) : 1, adj = 1;
  if (hasPast && r7 != null && r7 > fPast) { emp = true; adj = fPast > 0 ? (r7 / fPast) : null; }
  if (weekMode && hasPast) emp = false;
  var fEff = function (ended) {
    if (weekMode) return fAt(function (r) { return ended(r) && !isPast(r); });
    var f = fAt(ended); if (!emp) return f;
    var onlyPast = rows.every(function (r) { return !ended(r) || isPast(r); });
    if (onlyPast) return Math.max(f, r7);
    return adj == null ? f : Math.min(1, f * adj);
  };
  var sum = 0; for (var i = 0; i < cov; i++) { var d = new Date(t0.getFullYear(), t0.getMonth(), t0.getDate() + i); sum += fEff(function (r) { return r.e < d; }); }
  var fArr = fEff(function (r) { return ld > 0 ? (r.end <= arr) : (r.end < hcYmd_(t0)); });
  return { factor: Math.max(0, Math.min(1, sum / cov)), dead: (!oa && fArr <= 0.001), emp: emp, r7: r7, rP: rP, oa: oa,
           pastOnlyFactor: hasPast ? fEff(isPast) : null };
}

/* ===================== 報告 ===================== */

function hcReport_(res, now) {
  var lines = ['採購系統健康檢查　' + hcYmd_(now) + ' ' + ('0' + now.getHours()).slice(-2) + ':' + ('0' + now.getMinutes()).slice(-2)];
  lines.push(res.fails.length ? '❌ 需要處理 ' + res.fails.length + ' 項' : '✅ 沒有會讓叫貨量出錯的問題');
  res.fails.forEach(function (x, i) { lines.push('  ' + (i + 1) + '. [' + x.code + '] ' + x.msg); });
  if (res.warns.length) {
    lines.push('');
    lines.push('⚠️ 留意 ' + res.warns.length + ' 項');
    res.warns.forEach(function (x, i) { lines.push('  ' + (i + 1) + '. [' + x.code + '] ' + x.msg); });
  }
  var st = res.stats || {};
  lines.push('');
  lines.push('資料：POS 到 ' + (st.posMax || '—') + '｜收尾表 ' + (st.endcutRows == null ? '—' : st.endcutRows) + ' 列' + (st.endcutPairs ? '（' + st.endcutPairs + ' 組店×料）' : ''));
  lines.push('處理方式：把這封信整段貼給 Claude（採購系統專案），說「健康檢查出錯」。');
  return lines.join('\n');
}

/* ===================== 小工具 ===================== */

function hcObjs_(v) {
  if (!v) return null; if (v.length < 1) return [];
  var h = v[0].map(function (x) { return String(x || '').trim(); }), out = [];
  for (var i = 1; i < v.length; i++) {
    var row = v[i], any = false, o = {};
    for (var j = 0; j < h.length; j++) { if (!h[j]) continue; o[h[j]] = row[j]; if (row[j] !== '' && row[j] != null) any = true; }
    if (any) out.push(o);
  }
  return out;
}
function hcDataRows_(v) { var n = 0; for (var i = 1; i < v.length; i++) { if (v[i].some(function (x) { return x !== '' && x != null; })) n++; } return n; }
function hcCell_(x) { if (x instanceof Date) return 'D' + x.getTime(); if (x == null) return ''; return typeof x + ':' + String(x); }
function hcFirstDiff_(a, b) {
  var ra = a.filter(function (r, i) { return i === 0 || r.some(function (x) { return x !== '' && x != null; }); });
  var rb = b.filter(function (r, i) { return i === 0 || r.some(function (x) { return x !== '' && x != null; }); });
  if (ra.length !== rb.length) return '列數 ' + (ra.length - 1) + ' vs ' + (rb.length - 1);
  for (var i = 0; i < ra.length; i++) {
    var n = Math.max(ra[i].length, rb[i].length);
    for (var j = 0; j < n; j++) {
      var x = ra[i][j], y = rb[i][j];
      if (hcCell_(x === undefined ? '' : x) !== hcCell_(y === undefined ? '' : y)) return '第 ' + (i + 1) + ' 列第 ' + (j + 1) + ' 欄：「' + hcShow_(x) + '」vs「' + hcShow_(y) + '」';
    }
  }
  return '';
}
function hcShow_(x) { if (x instanceof Date) return '日期 ' + hcYmd_(x); if (x == null || x === '') return '空白'; return String(x).slice(0, 30); }
function hcMaxDateOf_(rows, col) { if (!rows) return null; var mx = null; rows.forEach(function (r) { var d = hcD_(r[col]); if (d && (!mx || d > mx)) mx = d; }); return mx; }
function hcD_(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  var m = String(v).match(/^(\d{4})\D(\d{1,2})\D(\d{1,2})/); if (!m) return null;
  var d = new Date(+m[1], +m[2] - 1, +m[3]); return isNaN(d) ? null : d;
}
function hcDateOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function hcAdd_(d, n) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n); }
function hcYmd_(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }