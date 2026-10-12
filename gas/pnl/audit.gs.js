/**
 * ============================================================
 * audit.gs — 總分類帳 vs 損益表「對帳照妖鏡」（2026-07-14）
 * ============================================================
 * 用途：一次跑完 12 店 × 39 科目，自動列出所有 fact_ledger ≠ fact_pnl 的科目，
 *       產出可直接交給會計的落差清單。
 *
 * 【為什麼要有這支】
 * 2026-07-14 手動點 1 店「其他收入」→ 發現總帳有「車位出租收入 30,000」但損益表沒有。
 * 12 店 × 39 科目 = 468 格，手動點不完。這支自動照完。
 *
 * 【安全性】
 * - auditLedgerVsPnl()            → 100% 唯讀，只輸出到執行記錄（Logger），零寫入
 * - auditLedgerVsPnl_writeSheet() → 額外寫到「_audit_ledger」獨立新分頁（方便截圖給會計）
 *                                   ⚠️ 絕不碰 fact_pnl / fact_ledger / dim_account
 *
 * 【一致性】直接呼叫 ledger.gs 的 getLedgerRule_() 與 matchSummary_()
 *           → 照妖鏡的判定 = 儀表板點下去實際會看到的，不會有兩套邏輯
 *
 * 【部署】新增此檔到「烘焙DIY損益分析系統」GAS 專案 → Ctrl+S → 直接在編輯器執行
 *         （不用重新部署 Web App，這支不對外開放）
 *
 * 【執行】函式下拉選 auditLedgerVsPnl → 執行 → 看「執行記錄」
 * ============================================================
 */

var AUDIT_TOL = 1;          // 容差（元），避免浮點誤差誤報
var AUDIT_TAB = '_audit_ledger';

/* ============================================================
 * 對外入口
 * ============================================================ */

/** 主用：115 上半年，純唯讀 */
function auditLedgerVsPnl() {
  return runLedgerAudit_({ ymMin: 11501, ymMax: 11506, writeSheet: false });
}

/** 同上，但把落差清單寫到 _audit_ledger 新分頁（方便截圖給會計） */
function auditLedgerVsPnl_writeSheet() {
  return runLedgerAudit_({ ymMin: 11501, ymMax: 11506, writeSheet: true });
}

/** 全期間（113~115 全掃，噪音較多，第一次可先不跑） */
function auditLedgerVsPnl_allPeriods() {
  return runLedgerAudit_({ ymMin: 0, ymMax: 999999, writeSheet: false });
}

/** 只跑 114 年度（歷史核對用） */
function auditLedgerVsPnl_114() {
  return runLedgerAudit_({ ymMin: 11401, ymMax: 11412, writeSheet: false });
}

/* ============================================================
 * 主邏輯
 * ============================================================ */
function runLedgerAudit_(opt) {
  var t0 = new Date();
  var ss = SpreadsheetApp.openById(SHEET_ID);

  /* ---------- 1. 讀 dim_account，建立 39 個科目的查詢規則 ---------- */
  var shAcc = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  if (!shAcc) { Logger.log('❌ 找不到 dim_account 分頁'); return; }
  var av = shAcc.getDataRange().getValues();
  var aci = {}; av[0].forEach(function(h, i){ aci[String(h).trim()] = i; });

  var aNameCol = pickCol_(aci, ['鼎新原始科目名', 'account_name']);
  var aCodeCol = pickCol_(aci, ['account_code']);
  if (aNameCol === undefined) {
    Logger.log('❌ dim_account 找不到科目名欄位。實際 header：' + av[0].join(' | '));
    return;
  }

  var accounts = [], codeToName = {};
  for (var i = 1; i < av.length; i++) {
    var nm = String(av[i][aNameCol]).trim();
    if (!nm) continue;
    var ac = (aCodeCol !== undefined) ? String(av[i][aCodeCol]).trim() : nm;
    accounts.push({ code: ac, name: nm, rule: getLedgerRule_(nm) });   // ← 重用 ledger.gs 的規則
    codeToName[ac] = nm;
  }
  var auditable = accounts.filter(function(a){ return a.rule.codes.length > 0; });
  Logger.log('dim_account 科目 ' + accounts.length + ' 個，其中有總帳代碼可對帳 ' + auditable.length + ' 個');

  /* ---------- 2. 讀 fact_ledger，逐列判定歸屬科目 ---------- */
  var shL = ss.getSheetByName(LEDGER_TAB);
  if (!shL) { Logger.log('❌ 找不到 fact_ledger 分頁'); return; }
  var lv = shL.getDataRange().getValues();
  var lci = {}; lv[0].forEach(function(h, i){ lci[String(h).trim()] = i; });
  var lStore = lci['store_id'], lYm = lci['ym'], lCode = lci['account_code'],
      lSmy = lci['summary'], lAmt = lci['amount'];
  if (lStore === undefined || lYm === undefined || lCode === undefined || lAmt === undefined) {
    Logger.log('❌ fact_ledger 欄位對不上。實際 header：' + lv[0].join(' | '));
    return;
  }

  // 反向索引：鼎新 4 位代碼 → 候選科目（一碼可能對多科目，如 6201 / 6210）
  var codeIndex = {};
  auditable.forEach(function(a){
    a.rule.codes.forEach(function(c){
      if (!codeIndex[c]) codeIndex[c] = [];
      codeIndex[c].push(a);
    });
  });

  var ledgerSum = {}, ledgerStores = {}, unclassified = {}, conflicts = [], ledgerRows = 0;

  for (var i = 1; i < lv.length; i++) {
    var r = lv[i];
    var ym = toYmNum_(r[lYm]);
    if (!ym || ym < opt.ymMin || ym > opt.ymMax) continue;
    ledgerRows++;

    var st = String(Number(r[lStore]));
    ledgerStores[st] = true;

    var code = String(r[lCode]).trim();
    var smy  = String(r[lSmy] || '');
    var amt  = Number(r[lAmt]) || 0;

    var cands = codeIndex[code] || [];
    var hits = cands.filter(function(a){
      return matchSummary_(smy, a.rule.include, a.rule.exclude);   // ← 重用 ledger.gs
    });

    // 沒有任何科目撈得到 → 對照表缺代碼／關鍵字切不乾淨（最重要的警訊）
    if (hits.length === 0) {
      var uk = st + '|' + code;
      if (!unclassified[uk]) unclassified[uk] = { store: st, code: code, n: 0, amt: 0, ex: '' };
      unclassified[uk].n++;
      unclassified[uk].amt += amt;
      if (!unclassified[uk].ex) unclassified[uk].ex = smy.substring(0, 28);
      continue;
    }

    // 同一筆同時符合多個科目 → 關鍵字規則衝突（會被重複計算）
    if (hits.length > 1 && conflicts.length < 25) {
      conflicts.push('店' + st + '  ' + code + '  「' + smy.substring(0, 28) + '」→ ' +
                     hits.map(function(a){ return a.name; }).join(' / '));
    }

    hits.forEach(function(a){
      var k = st + '|' + a.name;
      // 收入類：總帳貸方為負 → 翻正，才能跟損益表（絕對值）直接比
      ledgerSum[k] = (ledgerSum[k] || 0) + amt * (a.rule.isIncome ? -1 : 1);
    });
  }

  /* ---------- 3. 讀 fact_pnl ---------- */
  var shP = ss.getSheetByName('fact_pnl');
  if (!shP) { Logger.log('❌ 找不到 fact_pnl 分頁'); return; }
  var pv = shP.getDataRange().getValues();
  var pci = {}; pv[0].forEach(function(h, i){ pci[String(h).trim()] = i; });

  var pStore = pickCol_(pci, ['store_id', '店號', 'store']);
  var pYm    = pickCol_(pci, ['year_month', 'ym', '年月', 'period']);   // ← fact_pnl 用 year_month
  var pYear  = pickCol_(pci, ['year_roc', '年度']);
  var pMon   = pickCol_(pci, ['month', '月份', 'mm']);
  var pName  = pickCol_(pci, ['account_name', '科目名稱', '鼎新原始科目名']);
  var pCode  = pickCol_(pci, ['account_code', '科目代碼']);
  var pAmt   = pickCol_(pci, ['amount', '金額']);

  if (pStore === undefined || pAmt === undefined ||
      (pYm === undefined && (pYear === undefined || pMon === undefined)) ||
      (pName === undefined && pCode === undefined)) {
    Logger.log('❌ fact_pnl 欄位對不上，請把下面這行貼給 Claude：');
    Logger.log('   實際 header：' + pv[0].join(' | '));
    return;
  }

  var pnlSum = {}, pnlStores = {}, pnlRows = 0;
  for (var i = 1; i < pv.length; i++) {
    var r = pv[i];
    var ym = (pYm !== undefined) ? toYmNum_(r[pYm])
                                 : (Number(r[pYear]) * 100 + Number(r[pMon]));
    if (!ym || ym < opt.ymMin || ym > opt.ymMax) continue;

    var st = String(Number(r[pStore]));
    pnlStores[st] = true;

    var nm = (pName !== undefined) ? String(r[pName]).trim()
                                   : codeToName[String(r[pCode]).trim()];
    if (!nm) continue;

    var k = st + '|' + nm;
    pnlSum[k] = (pnlSum[k] || 0) + (Number(r[pAmt]) || 0);
    pnlRows++;
  }
  Logger.log('讀入：fact_ledger ' + (lv.length - 1) + ' 列（期間內 ' + ledgerRows + '）、fact_pnl ' +
             (pv.length - 1) + ' 列（期間內 ' + pnlRows + '）');
  if (pnlRows === 0) {
    Logger.log('❌ fact_pnl 期間內 0 列！year_month 格式可能非預期。前 3 列樣本：');
    for (var s = 1; s <= Math.min(3, pv.length - 1); s++) {
      Logger.log('   year_month = 「' + pv[s][pYm] + '」');
    }
    return;
  }

  /* ---------- 4. 比對 ---------- */
  var stores = Object.keys(ledgerStores)
    .filter(function(s){ return pnlStores[s]; })
    .sort(function(a, b){ return Number(a) - Number(b); });
  var skipped = Object.keys(pnlStores)
    .filter(function(s){ return !ledgerStores[s]; })
    .sort(function(a, b){ return Number(a) - Number(b); });

  var diffs = [], okCount = 0, checked = 0, diffTotal = 0;

  stores.forEach(function(st){
    auditable.forEach(function(a){
      checked++;
      var k = st + '|' + a.name;
      var L = Math.round(ledgerSum[k] || 0);
      var P = Math.round(pnlSum[k] || 0);
      var d = L - P;
      if (Math.abs(d) <= AUDIT_TOL) { okCount++; return; }
      var note = (P === 0) ? '🔥 總帳有、損益表無'
               : (L === 0) ? '總帳無、損益表有'
               : '';
      diffs.push({ store: st, name: a.name, L: L, P: P, d: d, note: note });
      diffTotal += d;
    });
  });

  /* ---------- 5. 輸出 ---------- */
  Logger.log('');
  Logger.log('══════════ 總分類帳 vs 損益表　對帳照妖鏡 ══════════');
  Logger.log('期間：' + opt.ymMin + ' ~ ' + opt.ymMax);
  if (skipped.length) Logger.log('⏭ 跳過（無總分類帳資料）：店 ' + skipped.join('、店 '));
  Logger.log('');

  diffs.sort(function(x, y){ return Math.abs(y.d) - Math.abs(x.d); });

  Logger.log('🔴 金額不符：' + diffs.length + ' 筆（依差額大小排序）');
  if (diffs.length === 0) {
    Logger.log('   （無 —— 全部相符）');
  } else {
    diffs.forEach(function(x){
      Logger.log('   店' + x.store + '　' + x.name +
                 '　總帳 ' + fmtN_(x.L) + '　損益表 ' + fmtN_(x.P) +
                 '　差 ' + fmtS_(x.d) + (x.note ? '　' + x.note : ''));
    });
  }

  var uks = Object.keys(unclassified);
  Logger.log('');
  Logger.log('⚠️ 未歸類分錄：' + uks.length + ' 組（總帳有分錄，但沒有任何科目撈得到 → 對照表缺代碼）');
  uks.sort(function(a, b){ return Math.abs(unclassified[b].amt) - Math.abs(unclassified[a].amt); });
  uks.forEach(function(k){
    var u = unclassified[k];
    Logger.log('   店' + u.store + '　代碼 ' + u.code + '　' + u.n + ' 筆　合計 ' +
               fmtN_(u.amt) + '　例：「' + u.ex + '」');
  });

  Logger.log('');
  Logger.log('⚠️ 規則衝突：' + conflicts.length + ' 筆（一筆分錄同時符合多科目 → 會被重複計算）');
  conflicts.forEach(function(c){ Logger.log('   ' + c); });

  Logger.log('');
  Logger.log('══════════ 摘要 ══════════');
  Logger.log('檢查範圍：' + stores.length + ' 店 × ' + auditable.length + ' 科目 = ' + checked + ' 組');
  Logger.log('✅ 完全相符：' + okCount + '（' + Math.round(okCount / checked * 100) + '%）');
  Logger.log('🔴 金額不符：' + diffs.length + '　　淨差額 ' + fmtS_(diffTotal));
  Logger.log('⚠️ 未歸類：' + uks.length + ' 組');
  Logger.log('⚠️ 規則衝突：' + conflicts.length + ' 筆');
  Logger.log('耗時 ' + Math.round((new Date() - t0) / 1000) + ' 秒');

  /* ---------- 6.（選用）寫到獨立新分頁 ---------- */
  if (opt.writeSheet) {
    var sh = ss.getSheetByName(AUDIT_TAB) || ss.insertSheet(AUDIT_TAB);
    sh.clear();
    var out = [['店', '科目', '總分類帳', '損益表', '差額', '說明']];
    diffs.forEach(function(x){
      out.push([Number(x.store), x.name, x.L, x.P, x.d, x.note]);
    });
    uks.forEach(function(k){
      var u = unclassified[k];
      out.push([Number(u.store), '【未歸類】代碼 ' + u.code, u.amt, '', u.amt,
                u.n + ' 筆　例：' + u.ex]);
    });
    if (out.length === 1) out.push(['', '（全部相符）', '', '', '', '']);
    sh.getRange(1, 1, out.length, 6).setValues(out);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#FDE9D9');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, 6);
    Logger.log('');
    Logger.log('📄 已寫入分頁「' + AUDIT_TAB + '」（' + (out.length - 1) + ' 列）');
  }

  return { diffs: diffs.length, unclassified: uks.length, conflicts: conflicts.length };
}

/* ============================================================
 * 小工具
 * ============================================================ */

/**
 * 年月正規化 → 一律回民國 YYYMM（如 11501）
 * 吃得下：11501 / "11501" / "115/01" / "115-01" / "2026-01" / 202601
 */
function toYmNum_(v) {
  var s = String(v).trim();
  if (!s) return 0;
  var n;
  if (s.indexOf('/') >= 0 || s.indexOf('-') >= 0) {
    var p = s.split(/[\/\-]/);
    n = Number(p[0]) * 100 + Number(p[1]);
  } else {
    n = Number(s);
  }
  if (isNaN(n) || !n) return 0;
  if (n >= 190000) n -= 191100;   // 西元(202601) → 民國(11501)
  return n;
}

function fmtN_(n) {
  return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function fmtS_(n) {
  return (n > 0 ? '+' : '') + fmtN_(n);
}

/* ============================================================
 * 逐月鑽取 — 找出某科目的差異集中在哪一個月
 * 用法：改下面 AUDIT_DRILL 的科目名 → 執行 auditByMonth
 * ============================================================ */
var AUDIT_DRILL = '水電費';        // ← 改這裡換科目

function auditByMonth() {
  var name = AUDIT_DRILL;
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rule = getLedgerRule_(name);
  if (!rule.codes.length) { Logger.log('❌ 科目「' + name + '」沒有總帳代碼'); return; }

  /* ---- fact_ledger ---- */
  var lv = ss.getSheetByName(LEDGER_TAB).getDataRange().getValues();
  var lc = {}; lv[0].forEach(function(h, i){ lc[String(h).trim()] = i; });
  var codeSet = {}; rule.codes.forEach(function(c){ codeSet[c] = true; });

  var L = {};
  for (var i = 1; i < lv.length; i++) {
    var r = lv[i];
    var ym = toYmNum_(r[lc['ym']]);
    if (ym < 11501 || ym > 11506) continue;
    if (!codeSet[String(r[lc['account_code']]).trim()]) continue;
    if (!matchSummary_(String(r[lc['summary']] || ''), rule.include, rule.exclude)) continue;
    var k = Number(r[lc['store_id']]) + '|' + ym;
    L[k] = (L[k] || 0) + (Number(r[lc['amount']]) || 0) * (rule.isIncome ? -1 : 1);
  }

  /* ---- fact_pnl ---- */
  var pv = ss.getSheetByName('fact_pnl').getDataRange().getValues();
  var pc = {}; pv[0].forEach(function(h, i){ pc[String(h).trim()] = i; });

  var P = {};
  for (var i = 1; i < pv.length; i++) {
    var r = pv[i];
    if (String(r[pc['account_name']]).trim() !== name) continue;
    var ym = toYmNum_(r[pc['year_month']]);
    if (ym < 11501 || ym > 11506) continue;
    var k = Number(r[pc['store_id']]) + '|' + ym;
    P[k] = (P[k] || 0) + (Number(r[pc['amount']]) || 0);
  }

  /* ---- 輸出（只列有差額的店）---- */
  var yms = [11501, 11502, 11503, 11504, 11505, 11506];
  var stores = [1, 2, 3, 4, 5, 6, 7, 8, 11, 12];

  Logger.log('══════ 逐月鑽取：' + name + '（總帳代碼 ' + rule.codes.join(',') + '）══════');
  var grand = 0;
  stores.forEach(function(st){
    var lines = [], tot = 0;
    yms.forEach(function(ym){
      var k = st + '|' + ym;
      var l = Math.round(L[k] || 0), p = Math.round(P[k] || 0), d = l - p;
      tot += d;
      lines.push('   ' + ym + '　總帳 ' + fmtN_(l) + '　損益表 ' + fmtN_(p) +
                 (Math.abs(d) > 1 ? ('　🔴 ' + fmtS_(d)) : '　✅'));
    });
    if (Math.abs(tot) > 1) {
      grand += tot;
      Logger.log('');
      Logger.log('店' + st + '　全期差額 ' + fmtS_(tot));
      lines.forEach(function(x){ Logger.log(x); });
    }
  });
  Logger.log('');
  Logger.log('（只列出有差額的店）　淨差額合計 ' + fmtS_(grand));
}

/* ============================================================
 * 差異月份剖析 — 全科目全店，把每一筆差異定位到「哪一個月」
 * 決定「該請會計重匯哪幾個月的損益表」的關鍵工具
 * ============================================================ */
function auditDiffByMonth() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  /* dim_account → 規則 */
  var av = ss.getSheetByName(SHEET_NAMES.ACCOUNT).getDataRange().getValues();
  var ac = {}; av[0].forEach(function(h, i){ ac[String(h).trim()] = i; });
  var nameCol = pickCol_(ac, ['鼎新原始科目名', 'account_name']);
  var accounts = [];
  for (var i = 1; i < av.length; i++) {
    var nm = String(av[i][nameCol]).trim();
    if (!nm) continue;
    var rule = getLedgerRule_(nm);
    if (rule.codes.length) accounts.push({ name: nm, rule: rule });
  }
  var codeIndex = {};
  accounts.forEach(function(a){
    a.rule.codes.forEach(function(c){
      if (!codeIndex[c]) codeIndex[c] = [];
      codeIndex[c].push(a);
    });
  });

  /* fact_ledger → store|ym|name */
  var lv = ss.getSheetByName(LEDGER_TAB).getDataRange().getValues();
  var lc = {}; lv[0].forEach(function(h, i){ lc[String(h).trim()] = i; });
  var L = {}, lStores = {};
  for (var i = 1; i < lv.length; i++) {
    var r = lv[i];
    var ym = toYmNum_(r[lc['ym']]);
    if (ym < 11501 || ym > 11506) continue;
    var st = String(Number(r[lc['store_id']]));
    lStores[st] = true;
    var smy = String(r[lc['summary']] || '');
    (codeIndex[String(r[lc['account_code']]).trim()] || []).forEach(function(a){
      if (!matchSummary_(smy, a.rule.include, a.rule.exclude)) return;
      var k = st + '|' + ym + '|' + a.name;
      L[k] = (L[k] || 0) + (Number(r[lc['amount']]) || 0) * (a.rule.isIncome ? -1 : 1);
    });
  }

  /* fact_pnl */
  var pv = ss.getSheetByName('fact_pnl').getDataRange().getValues();
  var pc = {}; pv[0].forEach(function(h, i){ pc[String(h).trim()] = i; });
  var P = {}, pStores = {};
  for (var i = 1; i < pv.length; i++) {
    var r = pv[i];
    var ym = toYmNum_(r[pc['year_month']]);
    if (ym < 11501 || ym > 11506) continue;
    var st = String(Number(r[pc['store_id']]));
    pStores[st] = true;
    var k = st + '|' + ym + '|' + String(r[pc['account_name']]).trim();
    P[k] = (P[k] || 0) + (Number(r[pc['amount']]) || 0);
  }

  var stores = Object.keys(lStores).filter(function(s){ return pStores[s]; })
                     .sort(function(a, b){ return Number(a) - Number(b); });
  var yms = [11501, 11502, 11503, 11504, 11505, 11506];
  var byMonth = {};
  yms.forEach(function(ym){ byMonth[ym] = { items: [], net: 0 }; });

  stores.forEach(function(st){
    accounts.forEach(function(a){
      yms.forEach(function(ym){
        var k = st + '|' + ym + '|' + a.name;
        var l = Math.round(L[k] || 0), p = Math.round(P[k] || 0), d = l - p;
        if (Math.abs(d) <= 1) return;
        byMonth[ym].items.push({ st: st, name: a.name, l: l, p: p, d: d });
        byMonth[ym].net += d;
      });
    });
  });

  Logger.log('══════ 差異的月份分布（115/01–06）══════');
  Logger.log('');
  var dirty = [];
  yms.forEach(function(ym){
    var m = byMonth[ym];
    if (!m.items.length) { Logger.log(ym + '　✅ 完全乾淨'); return; }
    dirty.push(ym);
    Logger.log(ym + '　🔴 ' + m.items.length + ' 筆　淨差額 ' + fmtS_(m.net));
    m.items.sort(function(x, y){ return Math.abs(y.d) - Math.abs(x.d); });
    m.items.forEach(function(x){
      Logger.log('     店' + x.st + '　' + x.name + '　總帳 ' + fmtN_(x.l) +
                 '　損益表 ' + fmtN_(x.p) + '　' + fmtS_(x.d));
    });
    Logger.log('');
  });
  Logger.log('══════ 結論 ══════');
  Logger.log('需要請會計重新匯出損益表的月份：' + (dirty.length ? dirty.join('、') : '無，全部乾淨'));
}