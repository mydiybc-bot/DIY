/**
 * posRolling.gs — POS資料 滾動視窗支援層
 * 專案：POS_折扣聚合（1jsCg_m0KkOtqz1USKAM0iJAZcJaSDXDky41lpW1DaM75U8fh0ckjOb1a）
 *
 * v2（2026-08-26｜Phase 2）改了三件事：
 *   ① 修 bug：computeAggregation() 回傳的是 JSON 字串，verifyPosMonthAgg 需先 JSON.parse（經營者實測回報）
 *   ② 口徑修正：agg_pos_item_month 的鍵加入「次類別」。v1 用「第一個非空值」collapse，
 *      若同店同月同品項混有空白與「YYYY 檔期名」兩種次類別，會讓整月數量都被算成限定甜點
 *      （isLimited = subCat !== '' && subCat !== '無'，discount_aggregator.gs:263）。
 *      ⚠️ 因此 v2 上線後必須重跑一次 rebuildPosMonthAgg()。
 *   ③ 新增合併層三支：_prMergeHistory / _prFillDailyHistory / _prMergeStats
 *
 * ── 合併層核心設計：裁切前是 no-op ──────────────────────────────
 * _prMergeHistory 只補「月表有、但 Sheet 已經沒有」的月份。
 * 裁切還沒做時 Sheet 月份是完整的 → 補 0 列 → payload 與 v20 完全相同。
 * 所以 Phase 2 可以安心先部署、先跑 verifyV18() 確認 7 店對照組一字不變，
 * 之後 Phase 4 裁切，合併層自動接手，不需要再改任何程式。
 *
 * ── 設計原則（v1 沿用）──────────────────────────────────────
 * 1. 口徑單一真相：不重抄白名單，直接引用 discount_aggregator.gs 的
 *    DESSERT_CATEGORIES / SPREADSHEET_ID / _dim。讀不到就 throw，絕不靜默用空集合。
 * 2. 全域防撞名：變數一律 PR_ 前綴、私有函式一律 _pr 前綴。
 * 3. 地雷 3.2：ym 欄寫入前 setNumberFormat('@')，且只格式化實際資料列。
 * 4. 地雷 3.3：固定欄寬，右側不讀、不寫、不清。禁用 getDataRange() 寫入。
 * 5. 地雷 3.4：先刪後寫必有零列保護。
 * 6. 本檔對 POS資料 純唯讀。裁切函式 Phase 4 才加。
 */

// ══════════════════════════════════════════════════════════
// 設定
// ══════════════════════════════════════════════════════════

var PR_CFG = {
  POS_TAB:     'POS資料',
  T_MONTH:     'agg_pos_month',
  T_ITEM:      'agg_pos_item_month',
  T_LOG:       'trim_log',
  KEEP_MONTHS: 18,
  KEEP_MIN:    18,
  WRITE_CHUNK: 5000
};

var PR_HDR_MONTH = ['store_code', 'ym', 'main_category', 'dessert_qty', 'companion_qty',
                    'headcount', 'txn_count', 'revenue_gross', 'revenue_actual', 'row_count'];
var PR_HDR_ITEM  = ['store_code', 'ym', 'product', 'main_category', 'sub_category',
                    'qty', 'revenue_gross', 'revenue_actual', 'row_count'];
var PR_HDR_LOG   = ['執行日期', '裁切月份', '刪除列數', 'BQ對帳結果', '備份檔連結', '備註'];

var PR_YM_COL_MONTH = 1;
var PR_YM_COL_ITEM  = 1;

// 合併層快取（同一次 computeAggregation 內 _prMergeHistory 與 _prFillDailyHistory 共用）
var PR_MERGE_STATS = null;
var PR_MONTH_HC    = null;   // ym → {headcount, dessert, companion}
var PR_MERGED_YMS  = null;   // 本次補進來的月份清單

// POS資料 欄位：0 店號 1 店名 2 建立日期 3 商品名稱 4 主類別 5 次類別 6 單價 7 數量 8 實收 9 來客數

// ══════════════════════════════════════════════════════════
// 共用
// ══════════════════════════════════════════════════════════

function _prDessertSet() {
  if (typeof DESSERT_CATEGORIES === 'undefined' || !DESSERT_CATEGORIES) {
    throw new Error('❌ 讀不到全域 DESSERT_CATEGORIES（應在 discount_aggregator.gs 檔案最外層宣告）');
  }
  var set = {};
  if (Object.prototype.toString.call(DESSERT_CATEGORIES) === '[object Array]') {
    DESSERT_CATEGORIES.forEach(function (c) { set[String(c)] = true; });
  } else {
    Object.keys(DESSERT_CATEGORIES).forEach(function (k) { if (DESSERT_CATEGORIES[k]) set[String(k)] = true; });
  }
  if (Object.keys(set).length === 0) throw new Error('❌ DESSERT_CATEGORIES 解析後為空集合，中止');
  return set;
}

function _prSsId() {
  if (typeof SPREADSHEET_ID === 'undefined' || !SPREADSHEET_ID) {
    throw new Error('❌ 讀不到全域 SPREADSHEET_ID');
  }
  return SPREADSHEET_ID;
}

function _prYm(v) {
  if (v instanceof Date && !isNaN(v.getTime())) {
    return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
  }
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
  return null;
}

function _prNum(v) { var n = Number(v); return isNaN(n) ? 0 : n; }

/** 讀月表分頁為二維陣列（固定欄寬，右側不碰）。無資料回空陣列。 */
function _prReadTab(tabName, width) {
  var ss = SpreadsheetApp.openById(_prSsId());
  var sh = ss.getSheetByName(tabName);
  if (!sh) return [];
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, width).getValues();
}

// ══════════════════════════════════════════════════════════
// ★ 合併層（Phase 2 新增）
// ══════════════════════════════════════════════════════════

/**
 * 把「Sheet 已無、月表仍有」的月份，展開成 POS資料 形狀的虛擬列，接在 dataRaw 後面。
 * 裁切前 → 月表所有月份 Sheet 都還在 → 回傳 dataRaw 本身（同一個陣列參照，零成本）。
 *
 * 虛擬列的日期一律設為該月 1 日；單價 = revenue_gross / qty，故 單價×數量 = 原銷售額。
 * 由於現行聚合全部是 SUM(qty) 與 SUM(單價×數量)，一列 qty=100 與 100 列 qty=1 結果相同，
 * 所以 monthly / productRank / priceBands / seasonal / birthday / unknownCategories / kpi
 * 全部自動正確，不需要改任何一支下游函式。
 */
function _prMergeHistory(dataRaw) {
  PR_MERGE_STATS = { merged: false, virtualRows: 0, months: [], sheetMonths: 0, note: '' };
  PR_MONTH_HC = {};
  PR_MERGED_YMS = {};

  try {
    // 1. Sheet 內現有月份
    var sheetYms = {};
    for (var r = 1; r < dataRaw.length; r++) {
      var ym = _prYm(dataRaw[r][2]);
      if (ym) sheetYms[ym] = true;
    }
    PR_MERGE_STATS.sheetMonths = Object.keys(sheetYms).length;

    // 2. 月表（品項級）中 Sheet 已無的月份
    var itemRows = _prReadTab(PR_CFG.T_ITEM, PR_HDR_ITEM.length);
    if (itemRows.length === 0) {
      PR_MERGE_STATS.note = '月表無資料，未合併（若尚未裁切屬正常）';
      return dataRaw;
    }

    var virt = [];
    for (var i = 0; i < itemRows.length; i++) {
      var q = itemRows[i];
      var ymv = String(q[1]);
      if (!ymv || sheetYms[ymv]) continue;          // Sheet 還有 → 不補（裁切前全部落在這裡）

      var store = _prNum(q[0]);
      var qty   = _prNum(q[5]);
      var gross = _prNum(q[6]);
      var act   = _prNum(q[7]);
      var p = ymv.split('-');
      var dateObj = new Date(Number(p[0]), Number(p[1]) - 1, 1);
      var unitPrice = qty !== 0 ? (gross / qty) : 0;
      var storeName = '';
      try { storeName = _dim(store).name; } catch (e) { storeName = String(store); }

      virt.push([store, storeName, dateObj, String(q[2]), String(q[3]), String(q[4]),
                 unitPrice, qty, act, 0]);
      PR_MERGED_YMS[ymv] = true;
    }

    if (virt.length === 0) {
      PR_MERGE_STATS.note = 'Sheet 月份完整，合併層待命中（no-op）';
      return dataRaw;
    }

    // 3. 供 _prFillDailyHistory 用的「月份 → 人數」（取自表 1，與表 2 互為交叉驗證）
    var monRows = _prReadTab(PR_CFG.T_MONTH, PR_HDR_MONTH.length);
    for (var j = 0; j < monRows.length; j++) {
      var mr = monRows[j];
      var mym = String(mr[1]);
      if (!PR_MERGED_YMS[mym]) continue;
      if (!PR_MONTH_HC[mym]) PR_MONTH_HC[mym] = { headcount: 0, dessert: 0, companion: 0 };
      PR_MONTH_HC[mym].headcount += _prNum(mr[5]);
      PR_MONTH_HC[mym].dessert   += _prNum(mr[3]);
      PR_MONTH_HC[mym].companion += _prNum(mr[4]);
    }

    PR_MERGE_STATS.merged = true;
    PR_MERGE_STATS.virtualRows = virt.length;
    PR_MERGE_STATS.months = Object.keys(PR_MERGED_YMS).sort();
    PR_MERGE_STATS.note = '已併入 ' + PR_MERGE_STATS.months.length + ' 個歷史月份（來源：月表）';

    return dataRaw.concat(virt);

  } catch (err) {
    // 合併層絕不可讓主聚合掛掉：失敗就退回原始資料，並在 payload 留下痕跡
    PR_MERGE_STATS.note = '⚠️ 合併層失敗，已退回僅 Sheet 資料：' + err;
    Logger.log(PR_MERGE_STATS.note);
    return dataRaw;
  }
}

/**
 * 視窗外的日期：依「當日營收佔該月營收的比例」把月人數分攤下去。
 * 只填 headcount 仍為 0 的日（有真實列的日不動）。
 * 目的：避免 客單價 = 營收 ÷ 0 在營收趨勢圖上變成 Infinity。
 * 分攤結果在「月」的層級完全正確，日層級是估算——這是月表先天沒有日粒度的必然取捨。
 */
function _prFillDailyHistory(dailyMap) {
  try {
    if (!PR_MERGE_STATS || !PR_MERGE_STATS.merged) return;
    if (!PR_MONTH_HC || Object.keys(PR_MONTH_HC).length === 0) return;

    // 依月份分組出「需要補的日」
    var byYm = {};
    for (var k in dailyMap) {
      var d = dailyMap[k];
      var ym = String(d.date).substring(0, 7);
      if (!PR_MERGED_YMS[ym]) continue;      // 只處理被合併的月份
      if (_prNum(d.headcount) > 0) continue; // 已有真實人數 → 不覆蓋
      if (!byYm[ym]) byYm[ym] = [];
      byYm[ym].push(d);
    }

    Object.keys(byYm).forEach(function (ym) {
      var days = byYm[ym];
      var hc = PR_MONTH_HC[ym] ? PR_MONTH_HC[ym].headcount : 0;
      if (!hc || days.length === 0) return;

      var totalG = 0;
      days.forEach(function (d) { totalG += _prNum(d.revenue_gross); });

      var assigned = 0;
      if (totalG > 0) {
        for (var i = 0; i < days.length - 1; i++) {
          var share = Math.round(hc * (_prNum(days[i].revenue_gross) / totalG));
          days[i].headcount = share;
          days[i].headcountEstimated = true;
          assigned += share;
        }
      } else {
        // 該月完全沒有營收資料（極罕見）→ 平均分攤
        var per = Math.floor(hc / days.length);
        for (var j = 0; j < days.length - 1; j++) {
          days[j].headcount = per;
          days[j].headcountEstimated = true;
          assigned += per;
        }
      }
      // 最後一天吃掉四捨五入殘差，保證該月加總 = 月表人數（分毫不差）
      var lastD = days[days.length - 1];
      lastD.headcount = hc - assigned;
      lastD.headcountEstimated = true;
    });
  } catch (err) {
    Logger.log('⚠️ _prFillDailyHistory 失敗（不影響主聚合）：' + err);
  }
}

/** 合併層狀態，掛進 payload 供前端 .dsb 資料狀態列顯示 */
function _prMergeStats() {
  return PR_MERGE_STATS || { merged: false, virtualRows: 0, months: [], sheetMonths: 0, note: '未執行' };
}

// ══════════════════════════════════════════════════════════
// 核心掃描（唯讀）
// ══════════════════════════════════════════════════════════

function _prScan() {
  var t0 = Date.now();
  var ss = SpreadsheetApp.openById(_prSsId());
  var sh = ss.getSheetByName(PR_CFG.POS_TAB);
  if (!sh) throw new Error('❌ 找不到分頁「' + PR_CFG.POS_TAB + '」');

  var lastRow = sh.getLastRow();
  var data = sh.getRange(1, 1, lastRow, 10).getValues();
  Logger.log('讀 ' + PR_CFG.POS_TAB + ' ' + (lastRow - 1) + ' 列（' + ((Date.now() - t0) / 1000).toFixed(1) + 's）');

  var dessertSet = _prDessertSet();
  Logger.log('✔ 白名單引用成功：' + Object.keys(dessertSet).length + ' 類');

  var mMap = {}, iMap = {}, badRows = [], catStat = {}, ymRows = {};
  var skipped = { blank: 0, v500: 0, badDate: 0 };

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var storeCode = r[0];
    var product = String(r[3] == null ? '' : r[3]);

    if (!storeCode && !product) { skipped.blank++; continue; }
    if (product.indexOf('$500券') >= 0) { skipped.v500++; continue; }
    if (product === '肚肚對帳調整') continue;   // 2026-10-06：金額調整列不進月表（品項／件數／人數）

    var ym = _prYm(r[2]);
    if (!ym) {
      skipped.badDate++;
      if (badRows.length < 200) {
        badRows.push({ sheetRow: i + 1, store: storeCode, storeName: String(r[1] == null ? '' : r[1]),
                       rawDate: (r[2] instanceof Date) ? ('[Date]' + r[2]) : JSON.stringify(r[2]),
                       product: product, mainCat: String(r[4] == null ? '' : r[4]),
                       qty: r[7], unitPrice: r[6], actual: r[8] });
      }
      continue;
    }
    ymRows[ym] = (ymRows[ym] || 0) + 1;

    var store = parseInt(storeCode, 10); if (isNaN(store)) store = 0;
    var mainCat = String(r[4] == null ? '' : r[4]);
    var subCat  = String(r[5] == null ? '' : r[5]);
    var qty     = _prNum(r[7]);
    var gross   = _prNum(r[6]) * qty;
    var actual  = _prNum(r[8]);

    var isCompanion = (product === '陪同入場費');
    var isDessert   = !isCompanion && (dessertSet[mainCat] === true);

    if (!catStat[mainCat]) catStat[mainCat] = { qty: 0, gross: 0, rows: 0 };
    catStat[mainCat].qty += qty; catStat[mainCat].gross += gross; catStat[mainCat].rows++;

    var mk = store + '|' + ym + '|' + mainCat;
    var mo = mMap[mk];
    if (!mo) mo = mMap[mk] = { store: store, ym: ym, cat: mainCat, dq: 0, cq: 0, txn: 0, g: 0, a: 0, n: 0 };
    if (isDessert)   mo.dq += qty;
    if (isCompanion) mo.cq += qty;
    mo.txn += qty; mo.g += gross; mo.a += actual; mo.n++;

    // ★ v2：次類別進 key，不再 collapse（限定甜點判定靠它）
    var ik = store + '|' + ym + '|' + product + '|' + subCat;
    var io = iMap[ik];
    if (!io) io = iMap[ik] = { store: store, ym: ym, product: product, cat: mainCat, sub: subCat,
                               q: 0, g: 0, a: 0, n: 0 };
    io.q += qty; io.g += gross; io.a += actual; io.n++;
  }

  Logger.log('掃描完成（' + ((Date.now() - t0) / 1000).toFixed(1) + 's）｜跳過：空白 ' + skipped.blank +
             '、$500券 ' + skipped.v500 + '、日期無法解析 ' + skipped.badDate);

  return { mMap: mMap, iMap: iMap, badRows: badRows, catStat: catStat, ymRows: ymRows,
           skipped: skipped, totalRows: lastRow - 1, dessertSet: dessertSet };
}

// ══════════════════════════════════════════════════════════
// 診斷（唯讀）
// ══════════════════════════════════════════════════════════

function diagPosMonthAgg() {
  var s = _prScan();
  var mRows = Object.keys(s.mMap).length, iRows = Object.keys(s.iMap).length;
  var mCells = (mRows + 1) * (PR_HDR_MONTH.length + 2), iCells = (iRows + 1) * (PR_HDR_ITEM.length + 2);

  Logger.log('');
  Logger.log('════════ 月表規模預估 ════════');
  Logger.log('  ' + PR_CFG.T_MONTH + '：' + mRows + ' 列 = ' + mCells.toLocaleString() + ' cells');
  Logger.log('  ' + PR_CFG.T_ITEM  + '：' + iRows + ' 列 = ' + iCells.toLocaleString() + ' cells');
  Logger.log('  合計 ≈ ' + (mCells + iCells).toLocaleString() + ' cells（目標 < 500,000）');

  Logger.log('');
  Logger.log('════════ 逐月列數 ════════');
  Object.keys(s.ymRows).sort().forEach(function (ym) { Logger.log('  ' + ym + '：' + s.ymRows[ym].toLocaleString()); });

  Logger.log('');
  Logger.log('════════ 主類別清單 ════════');
  var known = s.dessertSet, exclSet = {};
  if (typeof EXCLUDED_CATEGORIES !== 'undefined' && EXCLUDED_CATEGORIES) {
    if (Object.prototype.toString.call(EXCLUDED_CATEGORIES) === '[object Array]') {
      EXCLUDED_CATEGORIES.forEach(function (c) { exclSet[String(c)] = true; });
    } else {
      Object.keys(EXCLUDED_CATEGORIES).forEach(function (k) { if (EXCLUDED_CATEGORIES[k]) exclSet[String(k)] = true; });
    }
  }
  Object.keys(s.catStat).sort(function (a, b) { return s.catStat[b].gross - s.catStat[a].gross; })
    .forEach(function (c) {
      var st = s.catStat[c];
      var tag = known[c] ? '✅計人' : (exclSet[c] ? '⬜排除' : '❓未知');
      Logger.log('  ' + tag + ' 「' + (c === '' ? '(空白)' : c) + '」 數量 ' + st.qty.toLocaleString() +
                 '｜銷售額 ' + Math.round(st.gross).toLocaleString() +
                 '｜均價 ' + (st.qty > 0 ? Math.round(st.gross / st.qty) : 0));
    });

  Logger.log('');
  Logger.log('════════ 日期無法解析的列 ════════');
  Logger.log('  共 ' + s.skipped.badDate + ' 列');
  s.badRows.forEach(function (b) {
    Logger.log('  第' + b.sheetRow + '列｜店' + b.store + ' ' + b.storeName + '｜原始日期=' + b.rawDate +
               '｜品項=' + b.product + '｜主類別=' + b.mainCat +
               '｜單價=' + b.unitPrice + ' 數量=' + b.qty + ' 實收=' + b.actual);
  });

  Logger.log('');
  Logger.log('════════ 保留視窗試算 ════════');
  var yms = Object.keys(s.ymRows).sort();
  var keep = yms.slice(-PR_CFG.KEEP_MONTHS);
  var cut = yms.slice(0, Math.max(0, yms.length - PR_CFG.KEEP_MONTHS));
  Logger.log('  KEEP_MONTHS = ' + PR_CFG.KEEP_MONTHS + '（硬下限 ' + PR_CFG.KEEP_MIN + '）');
  Logger.log('  保留：' + (keep[0] || '-') + ' ~ ' + (keep[keep.length - 1] || '-') + '（' + keep.length + ' 個月）');
  if (cut.length === 0) { Logger.log('  可裁切：無'); }
  else {
    var cutRows = 0; cut.forEach(function (ym) { cutRows += s.ymRows[ym]; });
    Logger.log('  可裁切：' + cut.join('、') + '（' + cutRows.toLocaleString() + ' 列 ≈ ' +
               (cutRows * 13).toLocaleString() + ' cells）');
  }
  Logger.log('');
  Logger.log('✅ 診斷完成。本次未寫入任何資料。');
  return { monthRows: mRows, itemRows: iRows, cells: mCells + iCells };
}

// ══════════════════════════════════════════════════════════
// 寫入
// ══════════════════════════════════════════════════════════

function _prWrite(ss, tabName, headers, rows, ymColIdx) {
  if (!rows || rows.length === 0) {
    Logger.log('🛑 ' + tabName + '：算出 0 列 → 保留原資料未寫入（零列保護）');
    return 0;
  }
  var sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('❌ 找不到分頁「' + tabName + '」');

  var W = headers.length, need = rows.length + 1;
  if (sh.getMaxColumns() < W) sh.insertColumnsAfter(sh.getMaxColumns(), W - sh.getMaxColumns());
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());

  sh.getRange(1, 1, sh.getMaxRows(), W).clearContent();
  if (ymColIdx != null) sh.getRange(2, ymColIdx + 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, 1, W).setValues([headers]);

  var written = 0;
  for (var i = 0; i < rows.length; i += PR_CFG.WRITE_CHUNK) {
    var chunk = rows.slice(i, i + PR_CFG.WRITE_CHUNK);
    sh.getRange(2 + i, 1, chunk.length, W).setValues(chunk);
    written += chunk.length;
  }
  SpreadsheetApp.flush();
  Logger.log('✅ ' + tabName + '：寫入 ' + written.toLocaleString() + ' 列 × ' + W + ' 欄');
  return written;
}

function rebuildPosMonthAgg(fromYm, toYm) {
  var t0 = Date.now();
  Logger.log('════════ rebuildPosMonthAgg（v2：次類別進 key）════════');

  var s = _prScan();
  var ss = SpreadsheetApp.openById(_prSsId());
  function inRange(ym) {
    if (fromYm && ym < fromYm) return false;
    if (toYm && ym > toYm) return false;
    return true;
  }

  var mRows = Object.keys(s.mMap).sort().filter(function (k) { return inRange(s.mMap[k].ym); })
    .map(function (k) { var o = s.mMap[k];
      return [o.store, o.ym, o.cat, o.dq, o.cq, o.dq + o.cq, o.txn, Math.round(o.g), Math.round(o.a), o.n]; });

  var iRows = Object.keys(s.iMap).sort().filter(function (k) { return inRange(s.iMap[k].ym); })
    .map(function (k) { var o = s.iMap[k];
      return [o.store, o.ym, o.product, o.cat, o.sub, o.q, Math.round(o.g), Math.round(o.a), o.n]; });

  if (mRows.length === 0 && iRows.length === 0) {
    Logger.log('🛑 POS資料 命中 0 列 → 整批中止，未寫入任何資料');
    return { month: 0, item: 0, aborted: true };
  }

  var nm = _prWrite(ss, PR_CFG.T_MONTH, PR_HDR_MONTH, mRows, PR_YM_COL_MONTH);
  var ni = _prWrite(ss, PR_CFG.T_ITEM,  PR_HDR_ITEM,  iRows, PR_YM_COL_ITEM);

  Logger.log('');
  Logger.log('════════ 完成（' + ((Date.now() - t0) / 1000).toFixed(1) + 's）════════');
  Logger.log('  ' + PR_CFG.T_MONTH + '：' + nm.toLocaleString() + ' 列');
  Logger.log('  ' + PR_CFG.T_ITEM  + '：' + ni.toLocaleString() + ' 列（v2 因次類別進 key，會略多於 v1）');
  return { month: nm, item: ni, aborted: false };
}

function initTrimLog() {
  var ss = SpreadsheetApp.openById(_prSsId());
  var sh = ss.getSheetByName(PR_CFG.T_LOG);
  if (!sh) throw new Error('❌ 找不到分頁「' + PR_CFG.T_LOG + '」');
  sh.getRange(1, 1, 1, PR_HDR_LOG.length).setValues([PR_HDR_LOG]);
  sh.getRange(2, 2, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat('@');
  Logger.log('✅ ' + PR_CFG.T_LOG + ' 表頭已建立');
}

// ══════════════════════════════════════════════════════════
// 抽驗
// ══════════════════════════════════════════════════════════

/**
 * 抽驗指定月份：月表 vs 現行 computeAggregation()。
 * ★ v2 修正：computeAggregation() 回傳 JSON 字串，必須先 JSON.parse（v1 的 bug）。
 */
function verifyPosMonthAgg(ym) {
  if (!ym || !/^\d{4}-\d{2}$/.test(ym)) throw new Error('格式 YYYY-MM，例：verifyPosMonthAgg(\'2025-06\')');

  var monRows = _prReadTab(PR_CFG.T_MONTH, PR_HDR_MONTH.length);
  if (monRows.length === 0) throw new Error('❌ ' + PR_CFG.T_MONTH + ' 沒有資料，請先跑 rebuildPosMonthAgg()');

  var byStore = {}, tot = { dq: 0, cq: 0, hc: 0, g: 0, a: 0 };
  monRows.forEach(function (r) {
    if (String(r[1]) !== ym) return;
    var st = _prNum(r[0]);
    if (!byStore[st]) byStore[st] = { dq: 0, cq: 0, hc: 0, g: 0, a: 0 };
    byStore[st].dq += _prNum(r[3]); byStore[st].cq += _prNum(r[4]); byStore[st].hc += _prNum(r[5]);
    byStore[st].g  += _prNum(r[7]); byStore[st].a  += _prNum(r[8]);
    tot.dq += _prNum(r[3]); tot.cq += _prNum(r[4]); tot.hc += _prNum(r[5]);
    tot.g  += _prNum(r[7]); tot.a  += _prNum(r[8]);
  });

  var stores = Object.keys(byStore).sort(function (a, b) { return a - b; });
  if (stores.length === 0) { Logger.log('⚠️ 月表中查無 ' + ym); return; }

  Logger.log('════════ 月表 ' + ym + ' 逐店 ════════');
  Logger.log('  店 ｜ 甜點數 ｜ 陪同數 ｜ 來客數 ｜ 銷售額(毛) ｜ 實收');
  stores.forEach(function (st) {
    var o = byStore[st];
    Logger.log('  ' + st + ' ｜ ' + o.dq.toLocaleString() + ' ｜ ' + o.cq.toLocaleString() + ' ｜ ' +
               o.hc.toLocaleString() + ' ｜ ' + o.g.toLocaleString() + ' ｜ ' + o.a.toLocaleString());
  });
  Logger.log('  合計 ｜ ' + tot.dq.toLocaleString() + ' ｜ ' + tot.cq.toLocaleString() + ' ｜ ' +
             tot.hc.toLocaleString() + ' ｜ ' + tot.g.toLocaleString() + ' ｜ ' + tot.a.toLocaleString());

  // 表1 vs 表2 交叉驗證（同一份原始資料的兩條獨立聚合路徑，數字必須一致）
  var itemRows = _prReadTab(PR_CFG.T_ITEM, PR_HDR_ITEM.length);
  var iG = 0, iA = 0;
  itemRows.forEach(function (r) { if (String(r[1]) === ym) { iG += _prNum(r[6]); iA += _prNum(r[7]); } });
  Logger.log('');
  Logger.log('  表1 vs 表2 交叉驗證：銷售額 ' + tot.g.toLocaleString() + ' vs ' + iG.toLocaleString() +
             (Math.abs(tot.g - iG) <= 1 ? ' ✅' : ' ❌') +
             '｜實收 ' + tot.a.toLocaleString() + ' vs ' + iA.toLocaleString() +
             (Math.abs(tot.a - iA) <= 1 ? ' ✅' : ' ❌'));

  Logger.log('');
  Logger.log('════════ 對照現行 computeAggregation() ════════');
  try {
    var raw = computeAggregation();
    var agg = (typeof raw === 'string') ? JSON.parse(raw) : raw;   // ★ v2 修正
    var mbs = agg && agg.monthlyByStore;
    if (!mbs) { Logger.log('⚠️ payload 無 monthlyByStore，請改用肉眼對照儀表板'); return; }

    var ref = {};
    mbs.forEach(function (m) {
      if (String(m.yearMonth) !== ym) return;
      var st = _prNum(m.storeCode);
      if (!ref[st]) ref[st] = { hc: 0, dq: 0, cq: 0 };
      ref[st].hc += _prNum(m.headcount);
      ref[st].dq += _prNum(m.dessertCount);
      ref[st].cq += _prNum(m.companionCount);
    });
    if (Object.keys(ref).length === 0) { Logger.log('⚠️ 現行聚合查無 ' + ym); return; }

    var fail = 0;
    stores.forEach(function (st) {
      var a = byStore[st], b = ref[st];
      if (!b) { Logger.log('  ❌ 店' + st + '：現行聚合無此店'); fail++; return; }
      if (a.hc === b.hc && a.dq === b.dq && a.cq === b.cq) {
        Logger.log('  ✅ 店' + st + '：來客 ' + a.hc + '、甜點 ' + a.dq + '、陪同 ' + a.cq);
      } else {
        fail++;
        Logger.log('  ❌ 店' + st + '：來客 ' + a.hc + ' vs ' + b.hc + '｜甜點 ' + a.dq + ' vs ' + b.dq +
                   '｜陪同 ' + a.cq + ' vs ' + b.cq);
      }
    });
    Logger.log('');
    Logger.log(fail === 0 ? '🎉 ' + ym + ' 逐店逐值完全相符，PASS'
                          : '🛑 ' + ym + ' 有 ' + fail + ' 家店不符 → 停下回報');
  } catch (e) {
    Logger.log('⚠️ 無法自動比對：' + e);
    Logger.log('   改用肉眼對照：儀表板日期篩到 ' + ym + '，看「分店比較」的來客數/甜點數。');
  }
}

function verifyPosMonthAgg_all3() {
  ['2024-12', '2025-06', '2026-02'].forEach(function (ym) {
    Logger.log(''); Logger.log('##################### ' + ym + ' #####################');
    verifyPosMonthAgg(ym);
  });
}

/**
 * ★ Phase 2 專用：合併層自檢。
 * 裁切前應顯示「no-op（待命中）」；裁切後應顯示補進來的月份與虛擬列數。
 * 另檢查 payload 是否仍算得出合理的全期數字。
 */
function verifyMergeLayer() {
  Logger.log('════════ 合併層自檢 ════════');
  var raw = computeAggregation();
  var agg = (typeof raw === 'string') ? JSON.parse(raw) : raw;
  var hm = agg.historyMerge || {};

  Logger.log('  merged      : ' + hm.merged);
  Logger.log('  虛擬列數    : ' + (hm.virtualRows || 0).toLocaleString());
  Logger.log('  補進來的月份: ' + ((hm.months && hm.months.length) ? hm.months.join('、') : '（無）'));
  Logger.log('  Sheet 月份數: ' + hm.sheetMonths);
  Logger.log('  說明        : ' + hm.note);
  Logger.log('');
  Logger.log('  全期 net    : ' + (agg.kpi.totalRevenueNet || 0).toLocaleString());
  Logger.log('  全期 來客   : ' + (agg.kpi.headcount || 0).toLocaleString());
  Logger.log('  全期 甜點   : ' + (agg.kpi.dessertCount || 0).toLocaleString());
  Logger.log('  全期 客單價 : ' + Math.round((agg.kpi.totalRevenueNet || 0) / (agg.kpi.headcount || 1)));
  Logger.log('  totalRows   : ' + (agg.totalRows || 0).toLocaleString() + '（Sheet 實際列數）');
  Logger.log('  哨兵        : ' + ((agg.unknownCategories || []).length === 0 ? '空 ✅' : '⚠️ ' + agg.unknownCategories.length + ' 類未分類'));

  // 日人數 ÷0 檢查
  var bad = 0, est = 0;
  (agg.daily || []).forEach(function (d) {
    if (d.headcountEstimated) est++;
    if (_prNum(d.revenue_net) > 0 && _prNum(d.headcount) === 0) bad++;
  });
  Logger.log('  日人數估算列: ' + est + '｜有營收但人數 0 的日: ' + bad + (bad === 0 ? ' ✅' : ' ⚠️ 客單價會爆表'));
  Logger.log('');
  Logger.log(hm.merged ? '結論：合併層已生效（裁切後狀態）' : '結論：合併層待命中（裁切前的預期狀態，payload 應與 v20 相同）');
}
