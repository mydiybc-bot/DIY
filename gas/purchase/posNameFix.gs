/*** posNameFix.gs — 對照表「收銀品名」自動補對（2026-11-26，經營者交辦）
 *
 * 事故：2026 聖誕節「老公公蘋果生乳酪 (葷)」「堆雪人莓果慕斯 (葷)」對照表 A 欄（POS 資料產品名稱）多寫了「(葷)」，
 *       收銀機實際品名沒有 → 11/21 起銷售全部沒算進用量、檔期分配也變 0，12 店專屬料都沒出現在訂貨單。
 *
 * 做法（每天 rebuildPurchaseAgg 開頭自動跑；「🔄 更新」大更新也會跑）：
 *   對照表裡「現在或 28 天內要用」的列（起始有效日 ≤ 今天＋28、結束有效日空白或 ≥ 今天−14），
 *   如果 A 欄品名近 30 天 POS 一份都沒賣，但「去掉空白與括號標示」後，POS 近 30 天剛好只有一個品名對得上，
 *   而且那個 POS 品名還沒出現在對照表任何一列的 A 欄 →
 *   在對照表最後「新增一列」：A＝POS 實際品名，B～H 照原列（B＝BOM 甜點名不變）。
 *   原列一格都不動（不刪不改），所以永遠可以還原：刪掉新增的那列即可。
 *   每一筆新增、以及「找不到對得上的品名」都記在 BOM 本分頁 posname_fix_log。
 *
 * 不會做的：一個名字對到兩個以上 POS 品名（例：「X」與「X(半價)」都有賣）→ 只記「需人工」，不自動補。
 * 專用檔：rebuildPurchaseAgg 結尾的 mirrorTabs_ 會把整張對照表複製過去（MIRROR_TABS 含「產品名稱對照表」）。
 * 試跑（不寫入）：在編輯器執行 aaaPosNameFixDry，看執行紀錄。
 ***/
var PNF_POS_DAYS = 30;      // POS 看最近幾天
var PNF_AHEAD = 28;         // 起始有效日在今天後幾天內的列也檢查（開賣前就能先對好；開賣前 POS 沒賣過 → 通常找不到，只記 log）
var PNF_BEHIND = 14;        // 結束有效日在今天前幾天內的列也檢查
var PNF_LOG_TAB = 'posname_fix_log';

function pnfNorm_(s) {
  s = String(s == null ? '' : s).trim();
  s = s.replace(/[！-～]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });   // 全形轉半形（含全形括號）
  s = s.replace(/\([^)]*\)/g, '');   // 去掉所有括號標示：(葷)、(葷,可蛋奶素)、(限當日壽星)…
  s = s.replace(/\s+/g, '');
  return s.toLowerCase();
}

/** 核心（純函式，方便測試）：map＝[{row, a, b, from, to}]；posQty＝{POS品名: 近 N 天份數}；回傳 {add:[…], miss:[…], multi:[…]} */
function pnfPlan_(map, posQty, today) {
  var DAY = 86400000, t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  var aSet = {}; map.forEach(function (m) { var a = String(m.a || '').trim(); if (a) aSet[a] = 1; });
  var byNorm = {}; Object.keys(posQty).forEach(function (p) { if (!(posQty[p] > 0)) return; var n = pnfNorm_(p); if (n) (byNorm[n] = byNorm[n] || []).push(p); });
  var add = [], miss = [], multi = [], planned = {};
  map.forEach(function (m) {
    var a = String(m.a || '').trim(); if (!a || !String(m.b || '').trim()) return;
    var f = _toDate(m.from), t = _toDate(m.to);
    if (!f) return;                                                    // 沒有起始有效日的舊列不管
    if (f.getTime() > t0 + PNF_AHEAD * DAY) return;                     // 還很久才開賣
    if (t && t.getTime() < t0 - PNF_BEHIND * DAY) return;               // 已經結束
    if (posQty[a] > 0) return;                                          // A 欄品名有賣＝對得上
    var c = (byNorm[pnfNorm_(a)] || []).filter(function (p) { return p !== a; });
    var started = f.getTime() <= t0 - 2 * DAY && f.getTime() >= t0 - 60 * DAY;   // 只看近 60 天開賣、滿 2 天還沒對到的新品才算「查無」（老品項近 30 天沒賣＝停賣，不是對不上）
    if (!c.length) { if (started) miss.push({ row: m.row, a: a, b: m.b }); return; }
    var free = c.filter(function (p) { return !aSet[p] && !planned[p]; });
    if (c.length > 1) { multi.push({ row: m.row, a: a, b: m.b, cand: c }); return; }
    if (!free.length) return;                                           // 那個 POS 品名已經有別列在對（例：法式檸檬塔）→ 不用補
    planned[free[0]] = 1;
    add.push({ row: m.row, a: a, b: m.b, pos: free[0], qty: posQty[free[0]] });
  });
  return { add: add, miss: miss, multi: multi };
}

/** 實際執行：ss＝BOM 本；P＝已讀好的 POS 分頁（readTab_ 結果，省一次讀取；沒給就自己讀）；dry＝只看不寫 */
function posNameAutoFix_(ss, P, dry) {
  var today = new Date(), DAY = 86400000, cut = new Date(today.getTime() - PNF_POS_DAYS * DAY);
  P = P || readTab_(ss, POS_TAB);
  var iD = P.idx('建立日期'), iN = P.idx('商品名稱'), iQ = P.idx('數量'), posQty = {};
  P.rows.forEach(function (r) { var d = _toDate(r[iD]); if (!d || d < cut) return; var q = _num(r[iQ]); if (!q) return; var n = String(r[iN] || '').trim(); if (n) posQty[n] = (posQty[n] || 0) + q; });
  var sh = ss.getSheetByName(MAP_TAB); if (!sh) throw new Error('找不到分頁:' + MAP_TAB);
  var v = sh.getDataRange().getValues(), H = v[0].map(function (h) { return String(h).replace(/\s+/g, ''); });
  var cA = H.indexOf('POS資料產品名稱'), cB = H.indexOf('手動對應產品名稱'), cF = H.indexOf('起始有效日'), cT = H.indexOf('結束有效日');
  if (cA < 0 || cB < 0 || cF < 0 || cT < 0) throw new Error('對照表表頭對不上（要有 POS 資料產品名稱／手動對應產品名稱／起始有效日／結束有效日）');
  var map = []; for (var i = 1; i < v.length; i++) map.push({ row: i + 1, a: v[i][cA], b: v[i][cB], from: v[i][cF], to: v[i][cT], vals: v[i] });
  var plan = pnfPlan_(map, posQty, today);
  var stamp = Utilities.formatDate(today, 'Asia/Taipei', 'yyyy/MM/dd HH:mm');
  var logs = [];
  if (!dry && plan.add.length) {
    var width = v[0].length, newRows = plan.add.map(function (x) {
      var src = map[x.row - 2].vals.slice(0, width);
      src[cA] = x.pos;
      return src;
    });
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, newRows.length, width).setValues(newRows);
    plan.add.forEach(function (x, k) { x.newRow = start + k; });
  }
  plan.add.forEach(function (x) { logs.push([stamp, dry ? '試跑：會新增' : '已新增', x.row, x.a, x.b, x.pos, x.qty, dry ? '' : x.newRow]); });
  plan.multi.forEach(function (x) { logs.push([stamp, '需人工：對到多個 POS 品名', x.row, x.a, x.b, x.cand.join('、'), '', '']); });
  plan.miss.forEach(function (x) { logs.push([stamp, '查無：開賣 2 天以上 POS 沒有對得上的品名', x.row, x.a, x.b, '', '', '']); });
  var lg = ss.getSheetByName(PNF_LOG_TAB);
  if (lg && lg.getLastRow() >= 2) {   /* 「需人工／查無」同一列 7 天內記過就不再記（每天跑，避免洗版）；新增一定記 */
    var seen = {}, lv = lg.getRange(2, 1, lg.getLastRow() - 1, 4).getValues(), t7 = today.getTime() - 7 * DAY;
    lv.forEach(function (r) { var d = (r[0] instanceof Date) ? r[0] : _toDate(String(r[0]).slice(0, 10)); if (d && d.getTime() >= t7) seen[r[1] + '|' + r[2] + '|' + r[3]] = 1; });
    logs = logs.filter(function (l) { return String(l[1]).indexOf('新增') >= 0 || !seen[l[1] + '|' + l[2] + '|' + l[3]]; });
  }
  if (!dry && logs.length) {
    if (!lg) { lg = ss.insertSheet(PNF_LOG_TAB); lg.getRange(1, 1, 1, 8).setValues([['時間', '結果', '對照表原列', '原 POS 品名(A)', 'BOM 甜點名(B)', 'POS 實際品名', '近30天份數', '新增在第幾列']]); }
    lg.getRange(lg.getLastRow() + 1, 1, logs.length, 8).setValues(logs);
  }
  Logger.log('收銀品名自動補對：新增 ' + plan.add.length + '、需人工 ' + plan.multi.length + '、查無 ' + plan.miss.length + (dry ? '（試跑，沒寫入）' : ''));
  logs.forEach(function (l) { Logger.log(l.join('｜')); });
  return { added: dry ? 0 : plan.add.length, plan: plan };
}

function aaaPosNameFixDry() { return posNameAutoFix_(SpreadsheetApp.openById(SS_ID), null, true); }
function aaaPosNameFixNow() { return posNameAutoFix_(SpreadsheetApp.openById(SS_ID), null, false); }
