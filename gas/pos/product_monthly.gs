/**
 * product_monthly.gs — 品項 × 月 聚合（獨立 endpoint）
 * 2026-09-07 新增｜v1
 *
 * 目的：讓「品項分析」能吃期間篩選。
 *   原 productRank 是全期預聚合、無時間維度，
 *   問「2026 哪個品項賣最好」答不出來。
 *
 * 設計決定（2026-09-07）：
 *  1. 獨立 endpoint（action=product_monthly）+ 獨立快取，
 *     比照 daily_by_store 模式。
 *     → 不動 computeAggregation、不動 CACHE_KEY_PREFIX
 *     → net / 來客數 / 白名單 / 檔期立方體 全部不受影響
 *  2. 只掃 Sheet 真實列，不併 _prMergeHistory 的虛擬列。
 *     理由：虛擬列是否帶「商品名稱」未確認，硬併會產生空品項污染排名。
 *     回傳 coverage 讓前端標示實際涵蓋區間。
 *  3. 不切店別維度。ym×品項 約數千列，再乘 12 店會撐爆 payload。
 *     未來要店別請另開 endpoint，不要加進這支。
 *  4. 欄名縮短（ym/p/m/s/q/r）以壓 payload，前端自行對照。
 *
 * 口徑：r = 商品單價 × 數量（＝ monthly[].revenue 的 gross 口徑）
 *      ⚠️ 不是營業額。折扣是訂單層事件，切不到品項層。
 *
 * 依賴既有全域：SPREADSHEET_ID、computeAggregation()（僅驗收用）
 */

var PRODUCT_MONTHLY_CACHE_PREFIX = 'POS_PRODUCT_MONTHLY_V2_';   // 2026-10-06：排除「肚肚對帳調整」換版
var PRODUCT_MONTHLY_CACHE_TTL = 3600;

function getProductMonthlyCached() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(PRODUCT_MONTHLY_CACHE_PREFIX + 'chunks');
  if (chunkCount !== null) {
    var n = parseInt(chunkCount, 10);
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(PRODUCT_MONTHLY_CACHE_PREFIX + 'part_' + i);
    var parts = cache.getAll(keys);
    var allPresent = true;
    for (var a = 0; a < n; a++) {
      if (!parts[PRODUCT_MONTHLY_CACHE_PREFIX + 'part_' + a]) { allPresent = false; break; }
    }
    if (allPresent) {
      var assembled = '';
      for (var b = 0; b < n; b++) assembled += parts[PRODUCT_MONTHLY_CACHE_PREFIX + 'part_' + b];
      return assembled;
    }
  }
  var jsonStr = JSON.stringify(computeProductMonthly());
  var chunkSize = 90000;
  var chunks = [];
  for (var j = 0; j < jsonStr.length; j += chunkSize) chunks.push(jsonStr.substring(j, j + chunkSize));
  var cacheObj = {};
  for (var k = 0; k < chunks.length; k++) cacheObj[PRODUCT_MONTHLY_CACHE_PREFIX + 'part_' + k] = chunks[k];
  cacheObj[PRODUCT_MONTHLY_CACHE_PREFIX + 'chunks'] = String(chunks.length);
  cache.putAll(cacheObj, PRODUCT_MONTHLY_CACHE_TTL);
  return jsonStr;
}

function clearProductMonthlyCache() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(PRODUCT_MONTHLY_CACHE_PREFIX + 'chunks');
  var keys = [PRODUCT_MONTHLY_CACHE_PREFIX + 'chunks'];
  if (chunkCount !== null) {
    for (var i = 0; i < parseInt(chunkCount, 10); i++) {
      keys.push(PRODUCT_MONTHLY_CACHE_PREFIX + 'part_' + i);
    }
  }
  cache.removeAll(keys);
  Logger.log('product_monthly cache cleared');
}

/**
 * 掃 POS資料，聚合成 月 × 品項 × 主類別 × 次類別。
 * ⚠️ 三條過濾規則刻意與 computeAggregation 主迴圈逐條對齊（含順序），
 *    這是驗收 [1] 交叉對帳能成立的前提。改一邊要改兩邊。
 */
function computeProductMonthly() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('POS資料');
  var data = sheet.getDataRange().getValues();

  var map = {};
  var minYm = null, maxYm = null;
  var skipped = { noKey: 0, voucher500: 0, badDate: 0 };

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var storeCode = Number(row[0]);
    var dateRaw = row[2];
    var product = String(row[3] || '');
    var mainCat = String(row[4] || '');
    var subCat = String(row[5] || '');
    var unitPrice = Number(row[6]) || 0;
    var qty = Number(row[7]) || 0;

    if (!storeCode && !product) { skipped.noKey++; continue; }
    if (product.indexOf('$500券') >= 0) { skipped.voucher500++; continue; }
    if (product === '肚肚對帳調整') continue;   // 2026-10-06：一鍵追加新品 dudooGuard 補的金額調整列，營收照算（BQ），不算品項／件數／人數
    var dateObj = (dateRaw instanceof Date) ? dateRaw : new Date(String(dateRaw));
    if (isNaN(dateObj.getTime())) { skipped.badDate++; continue; }

    var mo = dateObj.getMonth() + 1;
    var ym = dateObj.getFullYear() + '-' + (mo < 10 ? '0' + mo : mo);
    if (minYm === null || ym < minYm) minYm = ym;
    if (maxYm === null || ym > maxYm) maxYm = ym;

    var pName = product || '(未命名品項)';
    var key = ym + '|' + pName + '|' + mainCat + '|' + subCat;
    var e = map[key];
    if (!e) e = map[key] = { ym: ym, p: pName, m: mainCat, s: subCat, q: 0, r: 0 };
    e.q += qty;
    e.r += unitPrice * qty;
  }

  var rows = [];
  for (var k in map) {
    var v = map[k];
    v.q = Math.round(v.q);
    v.r = Math.round(v.r);
    rows.push(v);
  }
  rows.sort(function (a, b) {
    if (a.ym !== b.ym) return a.ym < b.ym ? -1 : 1;
    return b.r - a.r;
  });

  return {
    rows: rows,
    coverage: { minYm: minYm, maxYm: maxYm, source: 'sheet_only_no_history_merge' },
    skipped: skipped,
    sheetRows: data.length - 1,
    lastUpdated: new Date().toISOString(),
    schema: 'ym=年月 p=品項 m=主類別 s=次類別 q=數量 r=牌價營收(單價×數量,非營業額)'
  };
}

/* ── 驗收：部署前必跑，全 PASS 才准部署 ──────────────────────
   [1] 交叉對帳：本表依 主類別×月 加總 == 現有 monthly[] 的 revenue
   [2] 迴歸錨點：2026 YTD 甜點 銷售額 61,706,224 / 份數 100,982
   [3] payload 量測
   ⚠️ 本函式會呼叫 computeAggregation()（含 BQ 查詢），約需 1~2 分鐘
   ────────────────────────────────────────────────────────── */
function verifyProductMonthly() {
  clearProductMonthlyCache();
  var pm = computeProductMonthly();
  var jsonStr = JSON.stringify(pm);
  var pass = true;

  Logger.log('══════════ productMonthly 驗收 ══════════');
  Logger.log('Sheet 列數=' + pm.sheetRows + '｜輸出列數=' + pm.rows.length);
  Logger.log('涵蓋區間: ' + pm.coverage.minYm + ' ~ ' + pm.coverage.maxYm);
  Logger.log('略過: 無鍵=' + pm.skipped.noKey + ' $500券=' + pm.skipped.voucher500 + ' 壞日期=' + pm.skipped.badDate);

  Logger.log('【3. payload】長度=' + jsonStr.length + ' → 快取塊數=' + Math.ceil(jsonStr.length / 90000));
  if (jsonStr.length > 900000) Logger.log('  ⚠️ 超過 900KB，前端會變慢，回報 Claude');

  var main = JSON.parse(computeAggregation());
  var ref = {};
  main.monthly.forEach(function (m) {
    var k = m.yearMonth + '|' + m.mainCategory;
    ref[k] = (ref[k] || 0) + m.revenue;
  });
  var mine = {};
  pm.rows.forEach(function (x) {
    var k = x.ym + '|' + x.m;
    mine[k] = (mine[k] || 0) + x.r;
  });
  var bad = 0, checked = 0, missing = 0;
  Object.keys(mine).forEach(function (k) {
    checked++;
    if (!(k in ref)) { missing++; return; }
    if (Math.abs(mine[k] - ref[k]) > 1) {
      bad++;
      if (bad <= 10) Logger.log('  ❌ ' + k + ' pm=' + mine[k] + ' monthly=' + ref[k] + ' 差=' + (mine[k] - ref[k]));
    }
  });
  Logger.log('【1. 交叉對帳】比對 ' + checked + ' 格｜不符 ' + bad + '｜monthly 無對應 ' + missing);
  if (bad > 0) { pass = false; Logger.log('  ★FAIL★ 過濾規則與主迴圈不一致'); }

  var EX = { '入場&共廚&其他': 1, '加價購': 1, '特約廠商': 1, '冰淇淋': 1, '活動': 1 };
  var rev26 = 0, qty26 = 0;
  pm.rows.forEach(function (x) {
    if (x.ym.indexOf('2026') !== 0) return;
    if (EX[x.m]) return;
    rev26 += x.r; qty26 += x.q;
  });
  var ok2 = (rev26 === 61706224 && qty26 === 100982);
  Logger.log('【2. 2026 YTD 錨點】銷售額=' + rev26 + '（應 61,706,224）份數=' + qty26 + '（應 100,982）' + (ok2 ? ' → PASS' : ' → 需確認'));
  Logger.log('  註：錨點取自 2026-09-07（含至 2026-09-06）。之後每天會長大，屆時確認「≥ 錨點且趨勢合理」即可。');
  if (rev26 < 61706224) pass = false;

  var t26 = pm.rows.filter(function (x) { return x.ym.indexOf('2026') === 0 && !EX[x.m]; });
  var byP = {};
  t26.forEach(function (x) {
    if (!byP[x.p]) byP[x.p] = { q: 0, r: 0 };
    byP[x.p].q += x.q; byP[x.p].r += x.r;
  });
  var top = Object.keys(byP).map(function (k) { return [k, byP[k].r, byP[k].q]; })
    .sort(function (a, b) { return b[1] - a[1]; }).slice(0, 15);
  Logger.log('【預覽】2026 品項銷售額 TOP 15');
  top.forEach(function (t, i) { Logger.log('  ' + (i + 1) + '. ' + t[0] + ' ｜ ' + t[1] + ' 元 ｜ ' + t[2] + ' 份'); });

  Logger.log(pass ? '══════ 全數 PASS，可部署 ══════' : '══════ ★有 FAIL，禁止部署，貼 log 給 Claude★ ══════');
}