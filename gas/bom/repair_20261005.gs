/**
 * repair_20261005.gs — 一次性：讓 POS 儀表板對上肚肚業績概況（經營者 2026-10-05 核可）
 *   ① 刪 2026-07-29 10 店「檸檬雙重奏 680×1、實收 680、來客 2」幽靈列：肚肚銷售明細匯出檔多出來的，
 *      交易紀錄第 4 張單只收 644。POS資料 有兩列一模一樣（另一列是第 9 張單的真實品項），
 *      刪「下一列是貓貓雙色戀人餅乾」那列（第 4 張單的位置）。
 *   ② 5 店 2025-12-26 的兩列退款（洛可可 −1、陪同入場費 −1，共 −809）改記退款日 2026-01-05（＝肚肚算法）。
 *   ③ 用 syncSheetToBigQuery_batch_v2 重寫 BigQuery 這 3 個店日（imported_at＝現在 → 排班隔天清晨自動重算這 3 天）。
 * fix20261005_preview：唯讀，列出會改哪幾列（列號＋整列內容＝備份）。
 * fix20261005_apply：重新找一次，筆數、位置、金額全部對上才改；對不上就停、什麼都不改。
 */
var FIX1005 = {
  GHOST: { store: 10, date: '2026-07-29', name: '檸檬雙重奏', price: 680, qty: 1, total: 680, cust: 2, next: '貓貓雙色戀人餅乾' },
  REFUND: { store: 5, date: '2025-12-26', to: '2026-01-05', names: ['洛可可', '陪同入場費'], sum: -809 }
};

function fix1005Find_() {
  var sh = SpreadsheetApp.openById(RECON_POS_SS_ID).getSheetByName(RECON_POS_TAB);
  var last = sh.getLastRow();
  var vals = sh.getRange(2, 1, last - 1, 10).getValues();
  var G = FIX1005.GHOST, R = FIX1005.REFUND, cache = {}, ghost = [], refund = [], moved = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i], st = parseInt(r[0], 10);
    if (st !== G.store && st !== R.store) continue;
    var v = r[2], ds;
    if (v instanceof Date) { ds = cache[v.getTime()]; if (ds === undefined) ds = cache[v.getTime()] = Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd'); }
    else ds = dguardNormDate_(v);
    if (st === G.store && ds === G.date && String(r[3]) === G.name && Number(r[6]) === G.price && Number(r[7]) === G.qty &&
        Number(r[8]) === G.total && Number(r[9]) === G.cust)
      ghost.push({ row: i + 2, next: vals[i + 1] ? String(vals[i + 1][3]) : '', vals: r });
    if (st === R.store && Number(r[7]) < 0 && (ds === R.date || ds === R.to) && R.names.indexOf(String(r[3])) >= 0)
      (ds === R.date ? refund : moved).push({ row: i + 2, vals: r });
  }
  return { sh: sh, last: last, ghost: ghost, refund: refund, moved: moved };
}

function fix1005Line_(x) {
  return '第 ' + x.row + ' 列：' + x.vals.map(function (v) {
    return v instanceof Date ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd') : String(v);
  }).join(' | ');
}

/** 檢查找到的列是否完全符合預期；回 {ok, msg, target} */
function fix1005Check_(f) {
  var G = FIX1005.GHOST, R = FIX1005.REFUND;
  var tgt = f.ghost.filter(function (g) { return g.next === G.next; });
  if (f.ghost.length !== 2 || tgt.length !== 1)
    return { ok: false, msg: '幽靈列：預期 2 列一樣的檸檬雙重奏、其中 1 列下一列是' + G.next + '；實際 ' + f.ghost.length + ' 列、符合位置 ' + tgt.length + ' 列' };
  var names = f.refund.map(function (x) { return String(x.vals[3]); }).sort().join(',');
  var sum = f.refund.reduce(function (a, x) { return a + Number(x.vals[6]) * Number(x.vals[7]); }, 0);
  if (f.refund.length !== 2 || names !== R.names.slice().sort().join(',') || sum !== R.sum)
    return { ok: false, msg: '退款列：預期 ' + R.date + ' 兩列（' + R.names.join('、') + '，合計 ' + R.sum + '）；實際 ' + f.refund.length + ' 列（' + names + '，合計 ' + sum + '）' };
  return { ok: true, msg: '全部符合', target: tgt[0] };
}

/** 唯讀預覽 */
function fix20261005_preview() {
  var f = fix1005Find_();
  Logger.log('POS資料 最後一列：' + f.last);
  Logger.log('— 檸檬雙重奏 680（' + FIX1005.GHOST.date + ' 10 店）：' + f.ghost.length + ' 列');
  f.ghost.forEach(function (g) { Logger.log('  ' + fix1005Line_(g) + '（下一列：' + g.next + '）'); });
  Logger.log('— 退款列（' + FIX1005.REFUND.date + ' 5 店）：' + f.refund.length + ' 列；已在 ' + FIX1005.REFUND.to + '：' + f.moved.length + ' 列');
  f.refund.concat(f.moved).forEach(function (x) { Logger.log('  ' + fix1005Line_(x)); });
  var c = fix1005Check_(f);
  Logger.log(c.ok ? '✅ ' + c.msg + '：會刪第 ' + c.target.row + ' 列、改 ' + f.refund.length + ' 列日期' : '❌ ' + c.msg + '（apply 會停下）');
  return c.ok;
}

/** 正式執行 */
function fix20261005_apply() {
  var f = fix1005Find_();
  var c = fix1005Check_(f);
  if (!c.ok) { Logger.log('❌ 停止，沒有改任何東西：' + c.msg); return false; }
  Logger.log('— 備份（改之前的整列內容）');
  f.ghost.concat(f.refund).forEach(function (x) { Logger.log('  ' + fix1005Line_(x)); });

  var to = FIX1005.REFUND.to.split('-');
  f.refund.forEach(function (x) {
    var nv = x.vals[2] instanceof Date ? new Date(+to[0], +to[1] - 1, +to[2]) : FIX1005.REFUND.to;
    f.sh.getRange(x.row, 3).setValue(nv);
  });
  f.sh.deleteRow(c.target.row);   // 先改日期、最後才刪列：刪列會讓後面的列號往前移，順序不能反
  SpreadsheetApp.flush();
  Logger.log('✅ 已改 ' + f.refund.length + ' 列日期 → ' + FIX1005.REFUND.to + '，已刪第 ' + c.target.row + ' 列');

  var g = fix1005Find_();
  Logger.log('— 改後：檸檬雙重奏 680 ' + g.ghost.length + ' 列（應為 1）；' + FIX1005.REFUND.date + ' 退款 ' + g.refund.length +
             ' 列（應為 0）；' + FIX1005.REFUND.to + ' 退款 ' + g.moved.length + ' 列（應為 2）；最後一列 ' + f.last + ' → ' + g.last);
  if (g.ghost.length !== 1 || g.refund.length !== 0 || g.moved.length !== 2) { Logger.log('❌ 改後檢查不符，未同步 BigQuery，請人工確認'); return false; }

  syncSheetToBigQuery_batch_v2([
    { date: FIX1005.GHOST.date, store: FIX1005.GHOST.store },
    { date: FIX1005.REFUND.date, store: FIX1005.REFUND.store },
    { date: FIX1005.REFUND.to, store: FIX1005.REFUND.store }
  ]);
  Logger.log('✅ BigQuery 3 個店日已重寫');
  return true;
}
