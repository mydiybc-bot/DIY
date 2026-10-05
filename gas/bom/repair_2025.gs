/**
 * repair_2025.gs — 一次性（2026-10-05 經營者：「完成後，再幫我確認 2025 也需要一致」）
 * 2025 年 147 個店月中 10 個跟肚肚業績概況不一致（2025-12 店5 已由 repair_20261005 修好），逐日查清：
 *   ① 2025-01-04 店7 +680：肚肚匯出檔幽靈品項（蜂蜜檸檬戚風 680；BQ 兩列一樣，刪一列）
 *   ② 肚肚帳單總額≠品項加總（肚肚自己兩份報表對不上）→ 補「肚肚對帳調整」列：
 *      2025-03-22 店4 +50、04-16 店7 +29（已裁切月份，只在 BigQuery）；
 *      05-17 店7 +1300、08-13 店4 +2、08-21 店5 +550、09-20 店5 +20、09-20 店6 +20、10-19 店12 +10（POS資料 還在）
 *   ③ 跨月退款：店12「加價購」−20 原記 2025-09-18，肚肚算在退款日 2025-10-02 → 改記 10-02（同新口徑）
 * fix2025_preview：唯讀，逐項確認現況與預期相同。fix2025_apply：全部對上才改；改完逐月逐店再比一次。
 * 修改前 BigQuery 先備份到 pos_transactions_autobak。
 */
var FIX2025 = {
  GHOST: { ds: '2025-01-04', st: 7, name: '蜂蜜檸檬戚風', price: 680, expectRows: 2 },
  BQ_ADJ: [{ ds: '2025-03-22', st: 4, amt: 50 }, { ds: '2025-04-16', st: 7, amt: 29 }],
  SHEET_ADJ: [{ ds: '2025-05-17', st: 7, amt: 1300 }, { ds: '2025-08-13', st: 4, amt: 2 }, { ds: '2025-08-21', st: 5, amt: 550 },
              { ds: '2025-09-20', st: 5, amt: 20 }, { ds: '2025-09-20', st: 6, amt: 20 }, { ds: '2025-10-19', st: 12, amt: 10 }],
  REDATE: { st: 12, from: '2025-09-18', to: '2025-10-02', name: '加價購', price: 20, qty: -1 },
  // 每個店日：修前 BQ 銷售總額 → 肚肚銷售總額（preview 檢查用）
  EXPECT: { '2025-01-04|7': [32176, 31496], '2025-03-22|4': [32799, 32849], '2025-04-16|7': [13261, 13290],
            '2025-05-17|7': [43910, 45210], '2025-08-13|4': [25024, 25026], '2025-08-21|5': [27664, 28214],
            '2025-09-18|12': [12325, 12345], '2025-09-20|5': [34065, 34085], '2025-09-20|6': [44331, 44351],
            '2025-10-02|12': [10444, 10424], '2025-10-19|12': [25520, 25530] }
};

function fix2025Gross_() {
  var keys = Object.keys(FIX2025.EXPECT);
  var cond = keys.map(function (k) { var p = k.split('|'); return "(sale_date = '" + p[0] + "' AND store_code = " + p[1] + ')'; }).join(' OR ');
  var rows = audit_bqAll_('SELECT CAST(sale_date AS STRING), store_code, CAST(SUM(unit_price * quantity) AS INT64) FROM ' +
    DGUARD.BQ + 'pos_transactions` WHERE ' + cond + ' GROUP BY 1, 2');
  var out = {};
  rows.forEach(function (r) { out[r[0] + '|' + parseInt(r[1], 10)] = Number(r[2]); });
  return out;
}

function fix2025FindRedate_() {
  var R = FIX2025.REDATE;
  var sh = SpreadsheetApp.openById(RECON_POS_SS_ID).getSheetByName(RECON_POS_TAB);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues(), hit = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (parseInt(r[0], 10) !== R.st || String(r[3]) !== R.name || Number(r[6]) !== R.price || Number(r[7]) !== R.qty) continue;
    var ds = r[2] instanceof Date ? Utilities.formatDate(r[2], DGUARD.TZ, 'yyyy-MM-dd') : dguardNormDate_(r[2]);
    if (ds === R.from) hit.push({ row: i + 2, vals: r });
  }
  return { sh: sh, hit: hit };
}

/** 檢查現況是否與預期相同；回 {ok, msgs, redate} */
function fix2025Check_() {
  var msgs = [], ok = true;
  var g = fix2025Gross_();
  Object.keys(FIX2025.EXPECT).forEach(function (k) {
    var want = FIX2025.EXPECT[k][0], got = g[k];
    if (got !== want) { ok = false; msgs.push('❌ ' + k + ' 目前 ' + got + '，預期修前 ' + want); }
    else msgs.push('✅ ' + k + ' 修前 ' + got + ' → 肚肚 ' + FIX2025.EXPECT[k][1]);
  });
  var G = FIX2025.GHOST;
  var n = Number(audit_bqAll_('SELECT COUNT(*) FROM ' + DGUARD.BQ + "pos_transactions` WHERE sale_date = '" + G.ds + "' AND store_code = " + G.st +
    " AND product_name = '" + G.name + "' AND unit_price = " + G.price + ' AND quantity = 1')[0][0]);
  if (n !== G.expectRows) { ok = false; msgs.push('❌ 幽靈品項 ' + G.name + ' 預期 ' + G.expectRows + ' 列，實際 ' + n); }
  else msgs.push('✅ ' + G.ds + ' 店' + G.st + ' ' + G.name + ' 680 有 ' + n + ' 列（刪 1 列）');
  var rd = fix2025FindRedate_();
  if (rd.hit.length !== 1) { ok = false; msgs.push('❌ 店12 加價購退款 ' + FIX2025.REDATE.from + ' 預期 1 列，實際 ' + rd.hit.length); }
  else msgs.push('✅ 店12 加價購 −20 在第 ' + rd.hit[0].row + ' 列（改記 ' + FIX2025.REDATE.to + '）');
  return { ok: ok, msgs: msgs, redate: rd };
}

function fix2025_preview() {
  var c = fix2025Check_();
  c.msgs.forEach(function (m) { Logger.log(m); });
  Logger.log(c.ok ? '✅ 全部符合，fix2025_apply 會照做' : '❌ 有不符，fix2025_apply 會停下');
  return c.ok;
}

function fix2025_apply() {
  var c = fix2025Check_();
  if (!c.ok) { c.msgs.forEach(function (m) { Logger.log(m); }); Logger.log('❌ 停止，沒有改任何東西'); return false; }
  var T = DGUARD.BQ + 'pos_transactions`', B = DGUARD.BQ + 'pos_transactions_autobak`', G = FIX2025.GHOST;
  var keys = Object.keys(FIX2025.EXPECT).map(function (k) { var p = k.split('|'); return "(sale_date = '" + p[0] + "' AND store_code = " + p[1] + ')'; });

  // 0) 備份
  dguardBqRun_('CREATE TABLE IF NOT EXISTS ' + B + ' AS SELECT *, CURRENT_TIMESTAMP() AS bak_at FROM ' + T + ' WHERE FALSE;\n' +
    'INSERT INTO ' + B + ' SELECT *, CURRENT_TIMESTAMP() FROM ' + T + ' WHERE ' + keys.join(' OR ') + ';');
  Logger.log('✅ BigQuery 已備份 ' + keys.length + ' 個店日到 pos_transactions_autobak');

  // ① 已裁切月份：刪 1 列幽靈品項（兩列一樣，留 1 列）
  dguardBqRun_('CREATE TEMP TABLE keep AS SELECT * EXCEPT(rn) FROM (\n' +
    '  SELECT *, ROW_NUMBER() OVER (PARTITION BY product_name, unit_price, quantity ORDER BY imported_at) AS rn\n' +
    '  FROM ' + T + " WHERE sale_date = '" + G.ds + "' AND store_code = " + G.st + ')\n' +
    "WHERE NOT (product_name = '" + G.name + "' AND unit_price = " + G.price + ' AND quantity = 1 AND rn = 1);\n' +
    'BEGIN TRANSACTION;\n' +
    'DELETE FROM ' + T + " WHERE sale_date = '" + G.ds + "' AND store_code = " + G.st + ';\n' +
    'INSERT INTO ' + T + ' SELECT * FROM keep;\n' +
    'COMMIT TRANSACTION;');
  Logger.log('✅ ' + G.ds + ' 店' + G.st + ' 刪 1 列 ' + G.name + ' 680');

  // ② 已裁切月份：BigQuery 直接補調整列
  FIX2025.BQ_ADJ.forEach(function (a) {
    dguardBqRun_('INSERT INTO ' + T + ' (store_code, store_name, sale_date, product_name, main_category, sub_category, unit_price, quantity, total_amount, customer_count, imported_at)\n' +
      'SELECT ' + a.st + ', ANY_VALUE(store_name), DATE \'' + a.ds + "', " + dguardSql_(DGUARD.ADJ_NAME) + ', ' + dguardSql_(DGUARD.ADJ_CATEGORY) +
      ", '無', " + a.amt + ', 1, ' + a.amt + ', 0, CURRENT_TIMESTAMP() FROM ' + T + " WHERE store_code = " + a.st + " AND sale_date = '" + a.ds + "'");
    Logger.log('✅ ' + a.ds + ' 店' + a.st + ' 補調整 +' + a.amt + '（BigQuery）');
  });

  // ③ 跨月退款改記退款日（試算表）
  var R = FIX2025.REDATE, h = c.redate.hit[0];
  Logger.log('— 備份（改之前）第 ' + h.row + ' 列：' + h.vals.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, DGUARD.TZ, 'yyyy-MM-dd') : String(v); }).join(' | '));
  var to = R.to.split('-');
  c.redate.sh.getRange(h.row, 3).setValue(h.vals[2] instanceof Date ? new Date(+to[0], +to[1] - 1, +to[2]) : R.to);
  SpreadsheetApp.flush();
  Logger.log('✅ 店12 加價購 −20 改記 ' + R.to);

  // ② 試算表月份：補調整列（沿用 dudooGuard 的修正流程：對照表、補列、重寫 BigQuery）
  var token = dguardLogin_(), dashCache = {};
  var dashOf = function (ds) { return dashCache[ds] || (dashCache[ds] = dguardDashboard_(token, ds, ds)); };
  var done = dguardFixItems_(token, FIX2025.SHEET_ADJ.map(function (a) { return { ds: a.ds, st: a.st, mode: 'adjust', amt: a.amt }; }), dashOf);
  done.forEach(function (s) { Logger.log('✅ ' + s); });
  syncSheetToBigQuery_batch_v2([{ date: R.from, store: R.st }, { date: R.to, store: R.st }]);

  // 驗證：逐店日、2025 逐月逐店
  var g = fix2025Gross_(), bad = 0;
  Object.keys(FIX2025.EXPECT).forEach(function (k) {
    var want = FIX2025.EXPECT[k][1];
    if (g[k] !== want) { bad++; Logger.log('❌ ' + k + ' 修後 ' + g[k] + '，肚肚 ' + want); }
  });
  Logger.log(bad ? '❌ 有 ' + bad + ' 個店日修後不符' : '✅ 11 個店日銷售總額修後全部＝肚肚');
  var mBad = [];
  for (var m = 1; m <= 12; m++) {
    var mm = ('0' + m).slice(-2), from = '2025-' + mm + '-01', last = dguardAddDays_(m === 12 ? '2026-01-01' : '2025-' + ('0' + (m + 1)).slice(-2) + '-01', -1);
    var dd = dguardDashboard_(token, from, last), oo = dguardOursRange_(from, last);
    DGUARD.STORES.concat([13]).forEach(function (st) {
      var x = dd[st] ? dd[st].amount : 0, o = oo[st] || 0;
      if (x !== o) mBad.push('2025-' + mm + ' 店' + st + ' 儀表板 ' + o + ' vs 肚肚 ' + x);
    });
  }
  Logger.log(mBad.length ? '❌ 2025 逐月逐店仍不一致：\n  ' + mBad.join('\n  ') : '✅ 2025 年 12 個月×各店 全部＝肚肚業績概況');
  return !bad && !mBad.length;
}
