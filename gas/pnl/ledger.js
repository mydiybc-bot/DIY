/**
 * ============================================================
 * ledger.gs v2 — 總分類帳明細查詢（2026-07-14 修正 collapse 誤算）
 * ============================================================
 * 用途：dashboard-pnl「損益明細」點科目時，查 fact_ledger 該科目明細
 *
 * 【v2 改了什麼】
 * 1. 【核心】新增「摘要關鍵字」過濾 → 解決一個鼎新代碼對應多個損益科目的混算：
 *      6201 → 薪資-員工薪資 / 獎金 / 實作費用 / 加班費 / 資遣費（5 科目）
 *      6210 → 人險-勞保 / 健保 / 勞退（3 科目）
 *    v1 會把整個 6210 的 161,414 顯示成「勞保」的明細（正確值 67,630）。
 * 2. 【核心】對照表補齊 39 科目（v1 缺 5202 進貨耗材 / 6222 職訓費 / 6220 研發 /
 *    6249 其他費用 / 7304 其他損失；且 6203 6204 6205… 多筆代碼原本就對錯）。
 * 3. 【地雷】移除 `codes.push(c.split('-')[0])` —— 這行會把子科目切回 4 位，
 *    未來 parser 若保留子科目會被它悄悄破壞。
 * 4. 收入類科目（dim_account 正負號=收入）翻轉符號 → 明細加總可直接對損益表。
 * 5. 費用科目未設定對照時 → 明確警告，不再誤報「營收類，來自 POS」。
 * 6. dim_account 若補了第 6 欄「鼎新代碼」+ 第 7 欄「摘要關鍵字」→ 自動優先採用
 *    （以後會計開新科目，填兩格即可，不用再改程式）。沒補也能跑，用下方 fallback。
 *
 * 【紅線】只讀 fact_ledger / dim_account，不寫任何資料。
 * 【部署】覆蓋原 ledger.gs → Ctrl+S → 部署 → 管理部署作業 → 生產版-損益分析 ✏️ → 新版本
 * 【前端】dashboard-pnl.html 零改動（錯誤訊息措辭已配合既有前端分流邏輯）
 * ============================================================
 */

var LEDGER_TAB = 'fact_ledger';

/* ============================================================
 * 主查詢
 * ============================================================ */
function getLedger_(params) {
  try {
    var storeIds = String(params.store_ids || '')
      .split(',').map(function(s){ return s.trim(); })
      .filter(function(s){ return s !== ''; });
    var ymMin = String(params.ym_min || '');
    var ymMax = String(params.ym_max || '');
    var accountName = String(params.account_name || '').trim();

    if (storeIds.length === 0 || !ymMin || !ymMax || !accountName) {
      return { ok: false, error: '缺少必要參數' };
    }

    var rule = getLedgerRule_(accountName);

    if (rule.codes.length === 0) {
      if (rule.isRevenue) {
        // 營收類本來就不走鼎新總帳（金額來自 POS／開立發票）
        // ⚠️ 訊息含「找不到科目」→ 觸發前端既有的友善文案，這是正確行為
        return { ok: false, error: '找不到科目「' + accountName + '」對應的鼎新代碼' };
      }
      // 費用／營業外科目沒設定 → 這是設定缺漏，必須明確警告
      // ⚠️ 訊息刻意「不含」找不到科目／鼎新代碼 字樣，前端才會走警告分支而非 POS 文案
      return { ok: false, error: '⚠️ 科目「' + accountName + '」尚未設定總帳對照，請補 dim_account 第 6 欄後重試。' };
    }

    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LEDGER_TAB);
    if (!sh) return { ok: false, error: 'fact_ledger 分頁不存在' };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, rows: [], total: 0 };

    var values = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
    var head = values[0];
    var col = {};
    head.forEach(function(h, i){ col[String(h).trim()] = i; });

    var cStore = col['store_id'], cYm = col['ym'], cCode = col['account_code'],
        cDate = col['voucher_date'], cVoucher = col['voucher_no'],
        cSummary = col['summary'], cAmount = col['amount'], cDc = col['dc'];

    var ymMinN = Number(ymMin), ymMaxN = Number(ymMax);
    var storeSet = {}; storeIds.forEach(function(s){ storeSet[String(Number(s))] = true; });
    var codeSet  = {}; rule.codes.forEach(function(c){ codeSet[String(c)] = true; });

    // 收入類：總帳貸方為負 → 翻正，明細加總才能直接對損益表金額
    var sign = rule.isIncome ? -1 : 1;

    var rows = [], total = 0;
    for (var i = 1; i < values.length; i++) {
      var r = values[i];
      if (!storeSet[String(Number(r[cStore]))]) continue;

      var ymN = Number(r[cYm]);
      if (ymN < ymMinN || ymN > ymMaxN) continue;

      if (!codeSet[String(r[cCode]).trim()]) continue;

      var summary = String(r[cSummary] || '');
      if (!matchSummary_(summary, rule.include, rule.exclude)) continue;

      var amt = (Number(r[cAmount]) || 0) * sign;
      total += amt;
      rows.push({
        store_id: String(Number(r[cStore])),
        ym: String(r[cYm]),
        voucher_date: String(r[cDate]),
        voucher_no: String(r[cVoucher]),
        summary: summary,
        amount: amt,
        dc: String(r[cDc])
      });
    }

    rows.sort(function(a, b){
      if (a.store_id !== b.store_id) return Number(a.store_id) - Number(b.store_id);
      if (a.ym !== b.ym) return Number(a.ym) - Number(b.ym);
      return String(a.voucher_no).localeCompare(String(b.voucher_no));
    });

    return { ok: true, rows: rows, total: total };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/* ============================================================
 * 摘要關鍵字比對
 * 規則：( include 為空 OR 摘要含任一 include ) AND ( 摘要不含任何 exclude )
 * ============================================================ */
function matchSummary_(summary, include, exclude) {
  var i;
  for (i = 0; i < exclude.length; i++) {
    if (exclude[i] && summary.indexOf(exclude[i]) >= 0) return false;
  }
  if (include.length === 0) return true;
  for (i = 0; i < include.length; i++) {
    if (include[i] && summary.indexOf(include[i]) >= 0) return true;
  }
  return false;
}

/* ============================================================
 * 科目名 → 查詢規則
 * 優先 dim_account（第 6 欄「鼎新代碼」/ 第 7 欄「摘要關鍵字」），沒有才用寫死 fallback
 * 回傳 { codes:[], include:[], exclude:[], isIncome:bool, isRevenue:bool }
 * ============================================================ */
function getLedgerRule_(accountName) {
  var rule = { codes: [], include: [], exclude: [], isIncome: false, isRevenue: false };

  /* ---------- 方案 A：dim_account 動態對照 ---------- */
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.ACCOUNT);
    if (sh) {
      var v = sh.getDataRange().getValues();
      var ci = {};
      v[0].forEach(function(h, i){ ci[String(h).trim()] = i; });

      var nameCol = pickCol_(ci, ['鼎新原始科目名', 'account_name']);
      var codeCol = pickCol_(ci, ['鼎新代碼', 'ledger_code']);
      var kwCol   = pickCol_(ci, ['摘要關鍵字', 'ledger_keyword']);
      var l1Col   = pickCol_(ci, ['大類 L1', '大類L1', 'category_l1', 'L1']);
      var signCol = pickCol_(ci, ['正負號']);

      if (nameCol !== undefined) {
        for (var i = 1; i < v.length; i++) {
          if (String(v[i][nameCol]).trim() !== accountName) continue;

          if (l1Col   !== undefined) rule.isRevenue = (String(v[i][l1Col]).trim() === '營業收入');
          if (signCol !== undefined) rule.isIncome  = (String(v[i][signCol]).trim() === '收入');

          if (codeCol !== undefined) {
            String(v[i][codeCol]).split(',').forEach(function(c){
              c = c.trim();
              if (c) rule.codes.push(c);   // ⚠️ 絕不做 split('-')[0]
            });
          }
          if (kwCol !== undefined) {
            String(v[i][kwCol]).split(',').forEach(function(k){
              k = k.trim();
              if (!k) return;
              var c0 = k.charAt(0);
              if (c0 === '!' || c0 === '！') rule.exclude.push(k.substring(1).trim());
              else rule.include.push(k);
            });
          }
          break;
        }
        if (rule.codes.length) return rule;   // dim_account 已設定 → 直接採用
      }
    }
  } catch (e) { /* 掉到方案 B */ }

  /* ---------- 方案 B：寫死 fallback（dim_account 未補欄時仍完全可用） ---------- */
  var fb = LEDGER_FALLBACK_[accountName];
  if (fb) {
    rule.codes    = fb.c;
    rule.include  = fb.i || [];
    rule.exclude  = fb.x || [];
    rule.isIncome = !!fb.inc;
  } else if (accountName === '銷貨收入') {
    rule.isRevenue = true;
  }
  return rule;
}

function pickCol_(ci, names) {
  for (var i = 0; i < names.length; i++) {
    if (ci[names[i]] !== undefined) return ci[names[i]];
  }
  return undefined;
}

/* ============================================================
 * 寫死 fallback 對照表 — 39 科目 / 32 代碼
 * 2026-07-14 以 1 店 115/01–06 三方對帳（fact_ledger × fact_pnl × 全公司代碼地圖）逐筆驗證
 * ⚠️ 銷貨收入不列入：營收金額來自 POS／開立發票，鼎新總帳無 4xxx 分錄（正確行為）
 * ============================================================ */
var LEDGER_FALLBACK_ = {
  /* ---- 食材成本 ---- */
  '進貨食材':        { c: ['5201'] },
  '進貨耗材':        { c: ['5202'] },   // v1 缺漏

  /* ---- 人事成本（6201 一碼 5 科目 → 靠摘要切分）---- */
  '薪資-員工薪資':   { c: ['6201'], x: ['獎金', '將金', '紅包', '實作', '加班', '資遣'] },
  '薪資-獎金':       { c: ['6201'], i: ['獎金', '將金', '紅包'] },
  '薪資-實作費用':   { c: ['6201'], i: ['實作'] },
  '薪資-加班費':     { c: ['6201'], i: ['加班'] },
  '薪資-資遣費':     { c: ['6201'], i: ['資遣'] },

  /* ---- 人險（6210 一碼 3 科目）---- */
  '人險-勞保':       { c: ['6210'], i: ['勞保'] },
  '人險-健保':       { c: ['6210'], i: ['健保'] },
  '人險-勞退':       { c: ['6210'], i: ['勞退'] },

  /* ---- 營運及管理費用 ---- */
  '職工福利':        { c: ['6219'] },
  '職業訓練/考照費': { c: ['6222'] },   // v1 缺漏（本次事件起點）
  '租金費用':        { c: ['6202'] },
  '水電費':          { c: ['6209'] },   // 含瓦斯費（鼎新損益表無獨立瓦斯科目）
  '清潔費':          { c: ['6217'] },
  '修繕費':          { c: ['6207'] },
  '耗材費用':        { c: ['6229'] },
  '運費':            { c: ['6205'] },
  '郵電費':          { c: ['6206'] },
  '手續費':          { c: ['6230'] },
  '差旅交通費':      { c: ['6204'] },
  '文具用品':        { c: ['6203'] },
  '雜項購置':        { c: ['6227'] },
  '管理費':          { c: ['6232'] },
  '內部管理費':      { c: ['6234'] },
  '勞務費':          { c: ['6223'] },
  '軟體維護費用':    { c: ['6242'] },
  '產險':            { c: ['6241'] },
  '稅捐':            { c: ['6213'] },
  '雜費':            { c: ['6240'] },
  '折舊及折耗':      { c: ['6215'] },
  '各項攤銷':        { c: ['6216'] },
  '研發費用':        { c: ['6220'] },   // v1 缺漏
  '其他費用':        { c: ['6249'] },   // v1 缺漏
  '廣告費':          { c: ['6208'] },

  /* ---- 營業外 ---- */
  '其他收入':        { c: ['7130'], inc: true },   // 收入類 → 翻正號
  '租賃稅費':        { c: ['7302'] },
  '其他損失':        { c: ['7304'] }    // v1 缺漏
};

/* ============================================================
 * 驗收測試 — 部署前在編輯器直接執行，全綠才部署
 * 基準：1 店 115/01～115/06（對照損益明細表本期金額）
 * ============================================================ */
function testLedgerRules() {
  var cases = [
    ['進貨食材',        440474],
    ['薪資-員工薪資',   744704],   // v1 會錯給 848518
    ['薪資-獎金',        97614],   // v1 會錯給 848518
    ['薪資-實作費用',     3200],
    ['薪資-加班費',       3000],
    ['人險-勞保',        67630],   // v1 會錯給 161414
    ['人險-健保',        45048],   // v1 會錯給 161414
    ['人險-勞退',        48736],   // v1 會錯給 161414
    ['職工福利',         31686],
    ['職業訓練/考照費',   6129],   // v1 誤報「來自 POS」
    ['租金費用',        330000],
    ['清潔費',           26760],
    ['修繕費',           13260],
    ['耗材費用',         67500],
    ['稅捐',            100094],
    ['內部管理費',      191972],
    ['雜費',             91790]
  ];
  var pass = 0;
  cases.forEach(function(t){
    var r = getLedger_({ store_ids: '1', ym_min: '11501', ym_max: '11506', account_name: t[0] });
    var got = r.ok ? r.total : ('ERR: ' + r.error);
    var ok = (got === t[1]);
    if (ok) pass++;
    Logger.log((ok ? '✅' : '❌') + ' ' + t[0] + ' → ' + got + '（應 ' + t[1] + '）');
  });
  Logger.log('---');
  Logger.log('通過 ' + pass + ' / ' + cases.length);

  // 已知落差（非程式問題，是 fact_pnl 損益表與總分類帳版本不同步，需請會計核對）
  Logger.log('');
  Logger.log('【已知落差 — 需會計核對，非程式錯誤】');
  ['水電費', '手續費', '其他收入'].forEach(function(n){
    var r = getLedger_({ store_ids: '1', ym_min: '11501', ym_max: '11506', account_name: n });
    Logger.log('  ' + n + ' 總帳=' + (r.ok ? r.total : r.error));
  });
  Logger.log('  水電費   總帳 87,410 vs 損益表 77,135 → 差 10,275');
  Logger.log('  手續費   總帳 25,202 vs 損益表 25,766 → 差   -564');
  Logger.log('  其他收入 總帳 31,500 vs 損益表  1,500 → 差 30,000（車位出租收入-巴爾巴）');
}