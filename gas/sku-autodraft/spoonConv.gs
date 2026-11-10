/*******************************************************************
 * spoonConv.gs  v1.0  (2026-11-05)
 * 平匙規格同步：每天早上隨 autoDraftNow（06:00–07:00，在 07:05 rebuild 之前）把
 * 「🚀 專案推動」平匙專案的【已定案】規格寫進 BOM 本 dim_unitconv 的「平匙」列。
 *   經營者 11/05 裁定：規格只在專案推動改一次，採購的週用量、訂貨點自動跟著變。
 * 規則：
 *   - 來源＝專案推動試算表 spoon_final（沒刪的列，每種食材取最後一次確認的那組；與專案推動 spoonSpec_ 同算法）。
 *   - 1 平匙＝幾 g：spoon_set 有「真實克數 real_g」用它，沒有就用定案的「食譜用克數 g_per_spoon」。
 *   - 只動 適用品項＝該食材、單位＝「平匙」的列：有列且數字不同 → 改換算量＋備註；沒有列 → 新增一列。
 *     沒定案的食材、其他單位（匙、g…）一律不碰；從不刪列。
 *   - 基準單位必須＝主檔使用單位（地雷 25）：主檔找不到或使用單位不是 g → 不寫、列警示。
 *   - 讀不到專案推動、沒有任何定案 → 什麼都不寫。克數要在 0～100 g 之間。
 *   - 寫 BOM 本；專用檔由 07:05 rebuildPurchaseAgg 的 mirrorTabs_ 整張複製過去。
 * 回滾：把 SC_ON 改 false（或刪本檔並刪 adRun_ 的呼叫行），dim_unitconv 用備份貼回。
 *******************************************************************/
var SC_ON = true;
var SC_PROJECT_ID = '1VE1dsSPk0Bl2dyTWrNo3G_oasD4pCR3ihllNedcOOT4';   // 專案推動_資料庫
var SC_CONV_TAB = 'dim_unitconv', SC_UNIT = '平匙';

function spoonConvDryRun() { var ss = SpreadsheetApp.openById(AD_SHEET_ID); var r = spoonConvSync_(ss, true); Logger.log(r.msg); return r.msg; }
/* 只跑平匙同步（不跑別名／草稿），要馬上套用時在編輯器執行；之後 07:05 rebuild 自動鏡像到專用檔 */
function spoonConvNow() {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ss = SpreadsheetApp.openById(AD_SHEET_ID), r = spoonConvSync_(ss, false);
    if (r.changed) { var lg = ss.getSheetByName(AD_LOG_TAB); if (lg) lg.appendRow([new Date(), 0, 0, r.msg]); }
    Logger.log(r.msg); return r.msg;
  } finally { lock.releaseLock(); }
}

function scRows_(sh) {
  var v = sh.getDataRange().getValues(), h = v[0].map(function (x) { return String(x || '').trim(); });
  return v.slice(1).map(function (r) { var o = {}; h.forEach(function (k, i) { if (k) o[k] = String(r[i] == null ? '' : r[i]).trim(); }); return o; });
}

/* 回傳 {changed:筆數, msg:一行說明} */
function spoonConvSync_(ss, dry) {
  if (!SC_ON) return { changed: 0, msg: '🥄 平匙同步：已關閉' };
  var pj = SpreadsheetApp.openById(SC_PROJECT_ID), fSh = pj.getSheetByName('spoon_final'), sSh = pj.getSheetByName('spoon_set');
  if (!fSh) return { changed: 0, msg: '🥄 平匙同步：讀不到 spoon_final，不寫入' };
  var spec = {};
  scRows_(fSh).forEach(function (r) {
    if (r.deleted === 'Y' || !r.ingredient || !(Number(r.g_per_spoon) > 0)) return;
    var o = spec[r.ingredient];
    if (!o || r.confirmed_at >= o.at) spec[r.ingredient] = { at: r.confirmed_at, g: Number(r.g_per_spoon), ml: r.spoon_ml };
  });
  var real = {};
  if (sSh) scRows_(sSh).forEach(function (r) { if (r.ingredient && Number(r.real_g) > 0) real[r.ingredient] = Number(r.real_g); });
  var ings = Object.keys(spec).sort();
  if (!ings.length) return { changed: 0, msg: '🥄 平匙同步：沒有定案規格，不寫入' };

  // 主檔使用單位（品名或 BOM別名）
  var sv = ss.getSheetByName(AD_SKU_TAB).getDataRange().getValues(), sh0 = sv[0].map(function (x) { return String(x || '').trim(); });
  var iN = sh0.indexOf('品名'), iU = sh0.indexOf('使用單位'), iA = sh0.indexOf('BOM別名'), useU = {};
  sv.slice(1).forEach(function (r) {
    var u = String(r[iU] || '').trim(), names = [String(r[iN] || '').trim()].concat(iA >= 0 ? String(r[iA] || '').split(/[;；,，、]/) : []);
    names.forEach(function (n) { n = String(n).trim(); if (n && !useU[n]) useU[n] = u; });
  });

  var cSh = ss.getSheetByName(SC_CONV_TAB);
  if (!cSh) return { changed: 0, msg: '🥄 平匙同步：找不到 dim_unitconv，不寫入' };
  var cv = cSh.getDataRange().getValues(), at = {};
  for (var i = 1; i < cv.length; i++) {
    if (String(cv[i][1]).trim() === SC_UNIT) at[String(cv[i][0]).trim()] = i;
  }
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'), upd = [], add = [], warn = [];
  ings.forEach(function (ing) {
    var s = spec[ing], g = real[ing] || s.g, src = real[ing] ? '實測' : '定案';
    if (!(g > 0 && g <= 100)) { warn.push(ing + ' 克數 ' + g + ' 不合理'); return; }
    if (useU[ing] !== 'g') { warn.push(ing + ' 主檔使用單位＝' + (useU[ing] || '找不到') + '（不是 g）'); return; }
    var note = '專案推動平匙' + src + '同步 ' + today + '（' + (s.ml ? s.ml + ' ml，' : '') + '定案 ' + String(s.at).slice(0, 10) + '）';
    if (at[ing] != null) {
      var r = at[ing], old = Number(cv[r][3]);
      if (Math.abs(old - g) > 1e-9 || String(cv[r][2]).trim() !== 'g') upd.push({ row: r + 1, ing: ing, old: cv[r][3], g: g, note: note });
    } else add.push([ing, SC_UNIT, 'g', g, note]);
  });
  if (!dry) {
    upd.forEach(function (u) { cSh.getRange(u.row, 3, 1, 3).setValues([['g', u.g, u.note]]); });
    if (add.length) cSh.getRange(cSh.getLastRow() + 1, 1, add.length, 5).setValues(add);
  }
  var msg = '🥄 平匙同步' + (dry ? '【試跑】' : '') + '：定案 ' + ings.length + ' 種、改 ' + upd.length + ' 列'
    + (upd.length ? '（' + upd.map(function (u) { return u.ing + ' ' + u.old + '→' + u.g; }).join('、') + '）' : '')
    + '、新增 ' + add.length + ' 列' + (add.length ? '（' + add.map(function (a) { return a[0] + ' ' + a[3]; }).join('、') + '）' : '')
    + (warn.length ? '、⚠ 略過 ' + warn.join('；') : '');
  return { changed: upd.length + add.length, msg: msg };
}
