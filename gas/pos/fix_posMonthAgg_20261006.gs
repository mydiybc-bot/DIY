/**
 * fix_posMonthAgg_20261006.gs — 一次性（經營者 2026-10-06 核可「第 4 項：2025/3–4 月舊月表重建 13 個店月」）
 * 月表 agg_pos_item_month／agg_pos_month 的 2025-03（1 店）與 2025-04（12 店）是 POS資料 裁切前寫的，
 * 含 10/02 已在 BigQuery 修掉的重複匯入（2025-04-23 整天重複、2025-03-09 店1 重複），所以 POS 儀表板這 13 個店月
 * 銷售數量／來客數／甜點數偏高（例 2025-04 精明 1,302 件 vs 肚肚 1,265）。
 * 新列由本機以修正後 BigQuery pos_transactions 依 _prScan 同一規則算出（排除 $500券 與「肚肚對帳調整」），
 * 驗證：以修正前備份 pos_transactions_bak_20261002 同法重算＝現行月表 852 品項列／146 分類列逐格相同。
 * 修前整份月表已備份到經營者本機 資料交換/輸出/20261006_POS其他數據對肚肚/。
 * fixPosMonthAgg20261006_preview：唯讀；fixPosMonthAgg20261006_apply：舊列總數量／總銷售額對上才改。
 */
// 【公開 repo 版】原檔此行內嵌 852 品項列＋146 分類列的各店月銷售資料（店×月×品項），因 DIY repo 公開而移除；
// 完整版只在 GAS 專案執行過，修前月表備份在經營者本機 資料交換/輸出/20261006_POS其他數據對肚肚/。
var FIX_PMA_1006 = { items: [ /* 已移除 */ ], months: [ /* 已移除 */ ] };
var FIX_PMA_1006_EXPECT = { oldQty: 18218, oldGross: 7234010, newQty: 17802, newGross: 7109091 };

function fixPma1006Target_(store, ym) { ym = String(ym); return ym === '2025-04' || (ym === '2025-03' && Number(store) === 1); }

function fixPma1006Read_() {
  var ss = SpreadsheetApp.openById(_prSsId());
  var it = ss.getSheetByName(PR_CFG.T_ITEM), mo = ss.getSheetByName(PR_CFG.T_MONTH);
  var iv = it.getRange(2, 1, it.getLastRow() - 1, PR_HDR_ITEM.length).getValues();
  var mv = mo.getRange(2, 1, mo.getLastRow() - 1, PR_HDR_MONTH.length).getValues();
  var tq = 0, tg = 0, ti = 0, tm = 0;
  iv.forEach(function (r) { if (fixPma1006Target_(r[0], r[1])) { ti++; tq += _prNum(r[5]); tg += _prNum(r[6]); } });
  mv.forEach(function (r) { if (fixPma1006Target_(r[0], r[1])) tm++; });
  return { ss: ss, iv: iv, mv: mv, ti: ti, tm: tm, tq: tq, tg: tg };
}

function fixPosMonthAgg20261006_preview() {
  var f = fixPma1006Read_(), E = FIX_PMA_1006_EXPECT;
  Logger.log('品項表 ' + f.iv.length + ' 列（目標 ' + f.ti + ' 列，數量 ' + f.tq + '，銷售額 ' + f.tg + '）；分類表 ' + f.mv.length + ' 列（目標 ' + f.tm + ' 列）');
  Logger.log('預期舊值 數量 ' + E.oldQty + '、銷售額 ' + E.oldGross + ' → 新值 數量 ' + E.newQty + '、銷售額 ' + E.newGross +
             '；新品項列 ' + FIX_PMA_1006.items.length + '、新分類列 ' + FIX_PMA_1006.months.length);
  var ok = f.tq === E.oldQty && f.tg === E.oldGross;
  Logger.log(ok ? '✅ 現況符合，apply 會照做' : '❌ 現況不符（可能已改過），apply 會停下');
  return ok;
}

function fixPosMonthAgg20261006_apply() {
  if (!FIX_PMA_1006.items.length) { Logger.log('❌ 公開 repo 版不含資料，不可執行'); return false; }
  var f = fixPma1006Read_(), E = FIX_PMA_1006_EXPECT;
  if (f.tq !== E.oldQty || f.tg !== E.oldGross) { Logger.log('❌ 停止，沒有改任何東西：目標舊列 數量 ' + f.tq + '／銷售額 ' + f.tg + ' 與預期不符'); return false; }
  var keepI = f.iv.filter(function (r) { return !fixPma1006Target_(r[0], r[1]); });
  var keepM = f.mv.filter(function (r) { return !fixPma1006Target_(r[0], r[1]); });
  var newI = keepI.concat(FIX_PMA_1006.items), newM = keepM.concat(FIX_PMA_1006.months);
  var ni = _prWrite(f.ss, PR_CFG.T_ITEM, PR_HDR_ITEM, newI, PR_YM_COL_ITEM);
  var nm = _prWrite(f.ss, PR_CFG.T_MONTH, PR_HDR_MONTH, newM, PR_YM_COL_MONTH);
  var g = fixPma1006Read_();
  Logger.log('品項表 ' + f.iv.length + ' → ' + g.iv.length + ' 列；分類表 ' + f.mv.length + ' → ' + g.mv.length + ' 列');
  Logger.log('目標店月 數量 ' + f.tq + ' → ' + g.tq + '（預期 ' + E.newQty + '），銷售額 ' + f.tg + ' → ' + g.tg + '（預期 ' + E.newGross + '）');
  var ok = g.tq === E.newQty && g.tg === E.newGross && g.iv.length === newI.length && g.mv.length === newM.length;
  Logger.log(ok ? '✅ 月表 13 個店月已重建' : '❌ 寫入後檢查不符，請用本機備份還原');
  return ok;
}
