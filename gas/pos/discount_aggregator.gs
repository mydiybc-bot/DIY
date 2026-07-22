/**
 * 自己做烘焙聚樂部 — POS 銷售動態儀表板 後端 API
 * 獨立 Apps Script 專案
 * 透過 openById() 讀取 Google Sheets：1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao
 *
 * ── v5.0 改造（金額單一真相：店/月/日層級金額全取自 BigQuery v_daily_net）──
 *  口徑（已用 1店2026-04 對帳 = 369,921）：
 *    net = view.net（= gross G×H − fact_discount H−L 去重）
 *  Y1 範圍：
 *    - kpi 金額總計、monthlyByStore、daily 的 gross/discount/net → 改讀 v_daily_net
 *    - monthly[]（按 mainCat/subCat 細分）的 revenue/actual → 維持 POS資料Sheet 算（view 無品項/類別粒度）
 *    - productRank / priceBands / seasonal / headcount → 維持 Sheet 算
 *    - 移除 fact_discount 分頁依賴（net 不再靠它）
 *  需搭配同專案 bq_connector.gs（queryDailyNet 已測通）。
 *
 * ── v12 新增（2026-06-17）────────────────────────────────────────
 *  新增 birthdayCake 與 birthdayItems 兩個聚合陣列，供「生日壽星」分頁使用
 *
 * ── v18 修正（2026-07-14）★ 來客數口徑修正 + 靜默丟棄治本 ★ ─────
 *
 *  【背景】11 店 2026-06 客單價 555 元追查，發現「吳寶春大獎麵包」（全期 4,067 份、
 *  均價 591 元）是 DIY 產品（一個麵包 = 一位客人），卻完全沒被計入來客數。
 *
 *  【根因】EXCLUDED_CATEGORIES 是死碼——excludedSet 建了但全檔從未被讀取。
 *  實際判定只有 `isDessert = dessertSet[mainCat] === true`，是純白名單。
 *  任何不在白名單的主類別 = 靜默歸零、不報錯、不留痕。
 *  （2026-07-03 南京店 225 份、本次麵包 4,067 份，都是同一個機制。）
 *
 *  【改動】
 *   1. DESSERT_CATEGORIES 新增：'吳寶春大獎麵包'、'吳寶春麵包'、''（主類別空白）
 *   2. isDessert / isCompanion 改為互斥（加入 '' 後若陪同列主類別空白會雙計，防呆）
 *   3. ★ 新增未分類哨兵 unknownCategories：不在白名單、也不在排除清單的主類別，
 *      會被聚合輸出到 result.unknownCategories，讓靜默丟棄「現形」。（治本）
 *   4. EXCLUDED_CATEGORIES 從死碼改為「哨兵的已知排除名單」（真正被讀取）
 *   5. 快取 V17 → V18
 *   6. 修好 rebuildCacheAndCheck() 寫死 'POS_DASHBOARD_V12_' 前綴的既有 bug
 *
 *  【預期影響】net 完全不變；來客數 +4,170（278,709 → 282,879，+1.5%）；
 *  客單價下修（全公司 478 → 約 471；11 店 2026-06：555 → 約 498）。
 *  只有 11、12 店（吳寶春共同品牌）會變動，其餘 10 店一字不變。
 *
 * ── v19 新增（2026-07-22）★ 檔期 detail 立方體 + 加購餅乾源頭正名 ★ ──
 *
 *  1. seasonal.campaigns[].detail：品項 × 店 × 週 數量立方體（週一為始日曆週）
 *     - 供前端「品項×分店」「分店×週次備料表」檢視；口徑 SUM(數量)，含負數退貨列
 *     - 由新函式 attachSeasonalDetail_() 掛載（重掃記憶體內 data，無額外 Sheet 讀取）
 *     - 無店號列不入 cube（campaign 營收/items 照舊），差額由 verifySeasonalV19 監控
 *  2. 加購餅乾源頭正名：classifyVoucher 類別 加價購 → 行銷折扣
 *     （經營者 2026-07-16 裁定：陪同客轉化讓利＝真行銷折扣；前端 v21 同步移除 remap）
 *  3. SEASONAL_DETAIL_MAX_ITEMS 品項封頂閘門（0=不限；payload 超標改 15）
 *  4. 快取 V18 → V19；API_VERSION v5.0 → v5.1
 *  5. 新增驗收函式 verifySeasonalV19()（部署前必跑，接著跑 verifyV18）
 *
 *  【口徑紅線】net／來客數／白名單完全不動；verifyV18 對照組必須一字不變。
 * ───────────────────────────────────────────────────────────────
 */

var SPREADSHEET_ID = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';

/**
 * 白名單：計入「甜點數」與「來客數」的主類別。
 * ⚠️ 這是唯一有效的判定依據。不在這張表裡的主類別 = 不計人。
 * ⚠️ 改動會回溯重算全店全歷史的甜點數／來客數／客單價／陪同率。
 */
var DESSERT_CATEGORIES = [
  '乳酪&奶蓋&慕斯', '巧克力', '裝飾蛋糕', '限定甜點',
  '水果', '生日蛋糕', '慶祝蛋糕', '其他',
  '主題活動＆群友限定', '雙層蛋糕',
  '蛋糕', '點心&餅乾', '塔派', '找不到',
  // ── v18 新增（2026-07-14，經營者確認）──
  '吳寶春大獎麵包',   // DIY 麵包，均價 591 元，一個麵包 = 一位客人（11/12 共同品牌店專有）
  '吳寶春麵包',       // 同上，均價 600 元
  ''                  // 主類別空白 = 甜點品項尚未對到類別，經營者確認：計入（同「找不到」精神）
];

/**
 * 已知排除名單：確認「不計人」的主類別。
 * ⚠️ v18 之前這是死碼（excludedSet 從未被讀取）。
 * v18 起改為「未分類哨兵」的比對依據：
 *   主類別 ∉ DESSERT_CATEGORIES 且 ∉ EXCLUDED_CATEGORIES
 *   → 記錄到 result.unknownCategories，避免再次靜默丟棄。
 */
var EXCLUDED_CATEGORIES = [
  '加價購',            // 均價 41 元，DIY 週邊加購（營收攤進客單價，不計人）
  '入場&共廚&其他',    // 陪同入場費另由 product === '陪同入場費' 判定
  '特約廠商',          // 全期營收 0
  '冰淇淋',            // 均價 80 元，加購品
  '活動'               // 自己人 500 元券（不計人數、不計金額）
];

var STORE_DIM = {
  1:  { name: '台中精明店',                   region: '中南區', closed: false },
  2:  { name: '台中草悟道店',                 region: '中南區', closed: false },
  3:  { name: '台北南京店',                   region: '北一區', closed: false },
  4:  { name: '台北士林店',                   region: '北一區', closed: false },
  5:  { name: '台南Focus店',                  region: '中南區', closed: false },
  6:  { name: '新竹文化店',                   region: '中南區', closed: false },
  7:  { name: '新北板橋店',                   region: '北二區', closed: false },
  8:  { name: '新北新店店',                   region: '北一區', closed: false },
  9:  { name: '桃園中壢店',                   region: '北三區', closed: false },
  10: { name: '桃園藝文店',                   region: '北三區', closed: false },
  11: { name: '吳寶春自己做台北信義A13店',     region: '北二區', closed: false },
  12: { name: '吳寶春自己做高雄SKM Park店',    region: '中南區', closed: false },
  13: { name: '台南西門店',                   region: '中南區', closed: true  }
};

var PRICE_BANDS = ['0-200', '200-400', '400-600', '600-800', '800+'];
var CACHE_TTL = 3600;
var CACHE_KEY_PREFIX = 'POS_DASHBOARD_V19_';   // v19：seasonal 新增 detail 立方體，結構變動強制換版（前版 V18）
var API_VERSION = 'v5.1';

function _dim(code) {
  var c = Number(code);
  if (STORE_DIM[c]) return STORE_DIM[c];
  return { name: '未知店(' + code + ')', region: '未分區', closed: false };
}

/* ============================================================
   doGet — Web App 進入點
   ============================================================ */
function doGet(e) {
  try {
    var callback = (e && e.parameter && e.parameter.callback) || '';
    if (e && e.parameter && e.parameter.action === 'daily_by_store') {
      var dbsResult = getDailyByStoreCached();
      if (callback) {
        return ContentService.createTextOutput(callback + '(' + dbsResult + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return ContentService.createTextOutput(dbsResult).setMimeType(ContentService.MimeType.JSON);
    }
    var result = getCachedOrCompute();
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + result + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(result)
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var errorJson = JSON.stringify({ error: err.message, stack: err.stack });
    var callback2 = e && e.parameter && e.parameter.callback;
    if (callback2) {
      return ContentService
        .createTextOutput(callback2 + '(' + errorJson + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(errorJson)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ============================================================
   快取管理
   ============================================================ */
function getCachedOrCompute() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(CACHE_KEY_PREFIX + 'chunks');
  if (chunkCount !== null) {
    var n = parseInt(chunkCount, 10);
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY_PREFIX + 'part_' + i);
    var parts = cache.getAll(keys);
    var allPresent = true;
    for (var i = 0; i < n; i++) {
      if (!parts[CACHE_KEY_PREFIX + 'part_' + i]) { allPresent = false; break; }
    }
    if (allPresent) {
      var assembled = '';
      for (var i = 0; i < n; i++) assembled += parts[CACHE_KEY_PREFIX + 'part_' + i];
      return assembled;
    }
  }
  var jsonStr = computeAggregation();
  var chunkSize = 90000;
  var chunks = [];
  for (var i = 0; i < jsonStr.length; i += chunkSize) chunks.push(jsonStr.substring(i, i + chunkSize));
  var cacheObj = {};
  for (var i = 0; i < chunks.length; i++) cacheObj[CACHE_KEY_PREFIX + 'part_' + i] = chunks[i];
  cacheObj[CACHE_KEY_PREFIX + 'chunks'] = String(chunks.length);
  cache.putAll(cacheObj, CACHE_TTL);
  return jsonStr;
}

/* ============================================================
   核心聚合邏輯
   v5.0：金額(店/月/日)取自 view；品項/類別/headcount 仍掃 Sheet
   v18：白名單修正 + 未分類哨兵
   ============================================================ */
function computeAggregation() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('POS資料');
  var data = sheet.getDataRange().getValues();
  var totalRows = data.length - 1;

  var dessertSet = {};
  for (var i = 0; i < DESSERT_CATEGORIES.length; i++) dessertSet[DESSERT_CATEGORIES[i]] = true;
  var excludedSet = {};
  for (var i = 0; i < EXCLUDED_CATEGORIES.length; i++) excludedSet[EXCLUDED_CATEGORIES[i]] = true;

  var totalDessertCount = 0;
  var totalCompanionCount = 0;
  var totalActual = 0;
  var limitedRevenue = 0;

  var monthlyMap = {};
  var productMap = {};
  var priceBandCount = [0, 0, 0, 0, 0];
  var priceBandRevenue = [0, 0, 0, 0, 0];

  var birthdayCakeMap = {};
  var birthdayItemsMap = {};

  var unknownCatMap = {};   // ★ v18 未分類哨兵

  var storeSet = {};
  var regionSet = {};
  var mainCatSet = {};
  var subCatSet = {};

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var storeCode = Number(row[0]);
    var dateRaw = row[2];
    var product = String(row[3] || '');
    var mainCat = String(row[4] || '');
    var subCat = String(row[5] || '');
    var unitPrice = Number(row[6]) || 0;
    var qty = Number(row[7]) || 0;
    var actual = Number(row[8]) || 0;
    var revenue = unitPrice * qty;

    if (!storeCode && !product) continue;
    if (product.indexOf('$500券') >= 0) continue;
    var dateObj = (dateRaw instanceof Date) ? dateRaw : new Date(String(dateRaw));
    if (isNaN(dateObj.getTime())) continue;

    var year = dateObj.getFullYear();
    var month = dateObj.getMonth() + 1;
    var yearMonth = year + '-' + (month < 10 ? '0' + month : month);

    var dim = _dim(storeCode);
    var storeName = dim.name;
    var region = dim.region;

    // ★ v18：先判陪同，再判甜點，兩者互斥（加入 '' 白名單後防止雙計）
    var isCompanion = (product === '陪同入場費');
    var isDessert   = !isCompanion && (dessertSet[mainCat] === true);
    var isLimited   = (subCat !== '' && subCat !== '無');

    // ★ v18 未分類哨兵：既不計人、也不在已知排除清單 → 記錄，避免靜默丟棄
    if (!isDessert && !isCompanion && excludedSet[mainCat] !== true) {
      var ukey = (mainCat === '') ? '(空白)' : mainCat;
      if (!unknownCatMap[ukey]) unknownCatMap[ukey] = { mainCategory: ukey, qty: 0, revenue: 0 };
      unknownCatMap[ukey].qty += qty;
      unknownCatMap[ukey].revenue += revenue;
    }

    totalActual += actual;
    if (isDessert) totalDessertCount += qty;
    if (isCompanion) totalCompanionCount += qty;
    if (isLimited) limitedRevenue += revenue;

    storeSet[storeCode] = true;
    regionSet[region] = true;
    if (mainCat) mainCatSet[mainCat] = true;
    if (subCat && subCat !== '無') subCatSet[subCat] = true;

    var mKey = yearMonth + '|' + storeCode + '|' + mainCat + '|' + subCat;
    if (!monthlyMap[mKey]) {
      monthlyMap[mKey] = {
        yearMonth: yearMonth, storeCode: storeCode, store: storeName, region: region,
        mainCategory: mainCat, subCategory: subCat,
        revenue: 0, actual: 0, dessertCount: 0, companionCount: 0,
        headcount: 0, txnCount: 0, avgPrice: 0, _priceSum: 0
      };
    }
    var m = monthlyMap[mKey];
    m.revenue += revenue;
    m.actual += actual;
    m.txnCount += qty;
    m._priceSum += unitPrice * qty;
    if (isDessert) { m.dessertCount += qty; m.headcount += qty; }
    if (isCompanion) { m.companionCount += qty; m.headcount += qty; }

    var pKey = product + '|' + mainCat + '|' + subCat;
    if (!productMap[pKey]) {
      productMap[pKey] = {
        product: product, mainCategory: mainCat, subCategory: subCat,
        totalQty: 0, totalRevenue: 0, avgPrice: 0, _priceSum: 0, _stores: {}
      };
    }
    var p = productMap[pKey];
    p.totalQty += qty;
    p.totalRevenue += revenue;
    p._priceSum += unitPrice * qty;
    p._stores[storeCode] = true;

    if (mainCat === '生日蛋糕') {
      var bcType = (product === '我的生日蛋糕') ? '免費' : '加價購';
      var bcKey = yearMonth + '|' + storeCode + '|' + bcType;
      if (!birthdayCakeMap[bcKey]) {
        birthdayCakeMap[bcKey] = {
          yearMonth: yearMonth,
          store: dim.name,
          region: dim.region,
          type: bcType,
          qty: 0,
          revenue: 0
        };
      }
      birthdayCakeMap[bcKey].qty += qty;
      if (bcType === '加價購') birthdayCakeMap[bcKey].revenue += unitPrice * qty;

      if (bcType === '加價購') {
        var biKey = year + '|' + product;
        if (!birthdayItemsMap[biKey]) {
          birthdayItemsMap[biKey] = { yr: year, product: product, qty: 0, revenue: 0 };
        }
        birthdayItemsMap[biKey].qty += qty;
        birthdayItemsMap[biKey].revenue += unitPrice * qty;
      }
    }

    if (isDessert && unitPrice > 0) {
      var bandIdx;
      if (unitPrice < 200) bandIdx = 0;
      else if (unitPrice < 400) bandIdx = 1;
      else if (unitPrice < 600) bandIdx = 2;
      else if (unitPrice < 800) bandIdx = 3;
      else bandIdx = 4;
      priceBandCount[bandIdx] += qty;
      priceBandRevenue[bandIdx] += revenue;
    }
  }

  // ★ v18：未分類哨兵輸出
  var unknownCatArr = [];
  for (var uk in unknownCatMap) {
    var u = unknownCatMap[uk];
    u.qty = Math.round(u.qty);
    u.revenue = Math.round(u.revenue);
    unknownCatArr.push(u);
  }
  unknownCatArr.sort(function(a, b) { return b.revenue - a.revenue; });

  var birthdayCakeArr = [];
  for (var key in birthdayCakeMap) {
    var bc = birthdayCakeMap[key];
    bc.qty = Math.round(bc.qty);
    bc.revenue = Math.round(bc.revenue);
    birthdayCakeArr.push(bc);
  }
  birthdayCakeArr.sort(function(a, b) {
    if (a.yearMonth !== b.yearMonth) return a.yearMonth < b.yearMonth ? -1 : 1;
    if (a.store !== b.store) return a.store < b.store ? -1 : 1;
    return a.type < b.type ? -1 : 1;
  });

  var birthdayItemsArr = [];
  for (var key in birthdayItemsMap) {
    var bi = birthdayItemsMap[key];
    bi.qty = Math.round(bi.qty);
    bi.revenue = Math.round(bi.revenue);
    birthdayItemsArr.push(bi);
  }
  birthdayItemsArr.sort(function(a, b) {
    if (a.yr !== b.yr) return b.yr - a.yr;
    return b.qty - a.qty;
  });

  var monthlyArr = [];
  for (var key in monthlyMap) {
    var mm = monthlyMap[key];
    mm.avgPrice = mm.txnCount > 0 ? Math.round(mm._priceSum / mm.txnCount) : 0;
    delete mm._priceSum;
    monthlyArr.push(mm);
  }

  var productArr = [];
  for (var key in productMap) {
    var pp = productMap[key];
    pp.avgPrice = pp.totalQty > 0 ? Math.round(pp._priceSum / pp.totalQty) : 0;
    pp.storeCount = Object.keys(pp._stores).length;
    delete pp._priceSum;
    delete pp._stores;
    productArr.push(pp);
  }
  productArr.sort(function(a, b) { return b.totalRevenue - a.totalRevenue; });

  var priceBands = [];
  for (var i = 0; i < PRICE_BANDS.length; i++) {
    priceBands.push({ band: PRICE_BANDS[i], count: priceBandCount[i], revenue: priceBandRevenue[i] });
  }

  var storeArr = [];
  for (var sc in storeSet) {
    var d = _dim(sc);
    storeArr.push({ storeCode: Number(sc), name: d.name, region: d.region, closed: d.closed });
  }
  storeArr.sort(function(a, b) { return a.storeCode - b.storeCode; });
  var regionArr = Object.keys(regionSet).sort();
  var mainCatArr = Object.keys(mainCatSet).sort();
  var subCatArr = Object.keys(subCatSet).sort();

  var headcount = totalDessertCount + totalCompanionCount;
  var avgSpendPerHead = headcount > 0 ? Math.round(totalActual / headcount) : 0;
  var companionRate = headcount > 0 ? Math.round(totalCompanionCount / headcount * 10000) / 100 : 0;

  var netRows = queryDailyNet();

  var swapRevRows = queryVoucherProductRevenue();
  var swapRevByDayStore = {};
  var swapRevByYMName = {};
  for (var svi = 0; svi < swapRevRows.length; svi++) {
    var svr = swapRevRows[svi];
    var dsKey = svr.date + '|' + svr.store_code;
    swapRevByDayStore[dsKey] = (swapRevByDayStore[dsKey] || 0) + (svr.revenue || 0);
    var ymKey = String(svr.date).substring(0, 7) + '|' + _dim(svr.store_code).name;
    swapRevByYMName[ymKey] = (swapRevByYMName[ymKey] || 0) + (svr.revenue || 0);
  }

  var dailyMap = {};
  var mbsMap = {};
  var totalGross = 0, totalDiscountRaw = 0, totalNet = 0;

  for (var i = 0; i < netRows.length; i++) {
    var nr = netRows[i];
    var dStr = nr.sale_date;
    var ym = dStr.substring(0, 7);
    var sc = nr.store_code;
    var g = nr.gross || 0;
    var disc = nr.total_discount || 0;
    var net = nr.net || 0;

    var ph = swapRevByDayStore[dStr + '|' + sc] || 0;
    g = g - ph;
    disc = disc - ph;

    totalGross += g;
    totalDiscountRaw += disc;
    totalNet += net;

    if (!dailyMap[dStr]) {
      dailyMap[dStr] = { date: dStr, headcount: 0, revenue_gross: 0, revenue_actual: 0, revenue_net: 0, discount: 0 };
    }
    dailyMap[dStr].revenue_gross += g;
    dailyMap[dStr].discount += disc;
    dailyMap[dStr].revenue_net += net;

    var mbsKey = ym + '|' + sc;
    if (!mbsMap[mbsKey]) {
      var dim2 = _dim(sc);
      mbsMap[mbsKey] = {
        yearMonth: ym, storeCode: sc, store: dim2.name, region: dim2.region,
        revenue_gross: 0, revenue_actual: 0, discount: 0, revenue_net: 0,
        dessertCount: 0, companionCount: 0, headcount: 0, txnCount: 0
      };
    }
    var b = mbsMap[mbsKey];
    b.revenue_gross += g;
    b.discount += disc;
    b.revenue_net += net;
  }

  for (var mi = 0; mi < monthlyArr.length; mi++) {
    var mObj = monthlyArr[mi];
    var mbsKey2 = mObj.yearMonth + '|' + mObj.storeCode;
    if (mbsMap[mbsKey2]) {
      mbsMap[mbsKey2].dessertCount += mObj.dessertCount || 0;
      mbsMap[mbsKey2].companionCount += mObj.companionCount || 0;
      mbsMap[mbsKey2].headcount += mObj.headcount || 0;
      mbsMap[mbsKey2].txnCount += mObj.txnCount || 0;
      mbsMap[mbsKey2].revenue_actual += mObj.actual || 0;
    }
  }
  var monthlyByStoreArr = Object.values(mbsMap);

  var dbpRows = queryDiscountByProgram();
  var dbpMap = {};
  for (var dpi = 0; dpi < dbpRows.length; dpi++) {
    var dpr = dbpRows[dpi];
    var cls = classifyVoucher(dpr.project_name);
    var dpDim = _dim(dpr.store_code);
    var dpKey = dpr.ym + '|' + dpr.store_code + '|' + cls.voucher;
    if (!dbpMap[dpKey]) {
      dbpMap[dpKey] = {
        yearMonth: dpr.ym,
        store: dpDim.name,
        region: dpDim.region,
        voucher: cls.voucher,
        category: cls.category,
        discount: 0, uses: 0, txns: 0
      };
    }
    dbpMap[dpKey].discount += (dpr.discount || 0);
    dbpMap[dpKey].uses += (dpr.uses || 0);
    dbpMap[dpKey].txns += (dpr.txns || 0);
  }
  var discountByProgramArr = Object.keys(dbpMap).map(function(k){
    var e = dbpMap[k];
    e.placeholderRev = (e.voucher === '自己人4人同行500') ? (swapRevByYMName[e.yearMonth + '|' + e.store] || 0) : 0;
    return e;
  });
  discountByProgramArr.sort(function(a, b) {
    if (a.yearMonth !== b.yearMonth) return a.yearMonth < b.yearMonth ? -1 : 1;
    return b.discount - a.discount;
  });

  _fillDailyHeadcount(data, dailyMap, dessertSet);

  var dailyArr = [];
  for (var key in dailyMap) dailyArr.push(dailyMap[key]);
  dailyArr.sort(function(a, b) { return a.date < b.date ? -1 : 1; });

  var limitedRevenueRatio = totalGross > 0 ? Math.round(limitedRevenue / totalGross * 10000) / 100 : 0;

  var result = {
    kpi: {
      totalRevenue: totalGross,
      totalRevenueGross: totalGross,
      totalRevenueActual: totalActual,
      totalRevenueNet: totalNet,
      totalDiscount: totalDiscountRaw,
      headcount: headcount,
      dessertCount: totalDessertCount,
      companionCount: totalCompanionCount,
      avgSpendPerHead: avgSpendPerHead,
      companionRate: companionRate,
      limitedRevenue: limitedRevenue,
      limitedRevenueRatio: limitedRevenueRatio
    },
    monthly: monthlyArr,
    productRank: productArr,
    daily: dailyArr,
    priceBands: priceBands,
    filters: {
      stores: storeArr,
      regions: regionArr,
      mainCategories: mainCatArr,
      subCategories: subCatArr
    },
    lastUpdated: new Date().toISOString(),
    totalRows: totalRows,
    seasonal: attachSeasonalDetail_(data, buildSeasonalAnalysis(data)),   // v19：掛載 品項×店×週 detail 立方體
    monthlyByStore: monthlyByStoreArr,
    discountUnmatched: [],
    birthdayCake: birthdayCakeArr,
    birthdayItems: birthdayItemsArr,
    discountByProgram: discountByProgramArr,
    unknownCategories: unknownCatArr,   // ★ v18 未分類哨兵
    apiVersion: API_VERSION
  };

  return JSON.stringify(result);
}

function _fillDailyHeadcount(data, dailyMap, dessertSet) {
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var dateRaw = row[2];
    var product = String(row[3] || '');
    if (product.indexOf('$500券') >= 0) continue;
    var mainCat = String(row[4] || '');
    var qty = Number(row[7]) || 0;
    var dateObj = (dateRaw instanceof Date) ? dateRaw : new Date(String(dateRaw));
    if (isNaN(dateObj.getTime())) continue;
    var year = dateObj.getFullYear();
    var month = dateObj.getMonth() + 1;
    var day = dateObj.getDate();
    var dateStr = year + '-' + (month < 10 ? '0' + month : month) + '-' + (day < 10 ? '0' + day : day);
    var isCompanion = (product === '陪同入場費');                        // v18：先判陪同
    var isDessert = !isCompanion && (dessertSet[mainCat] === true);      // v18：互斥
    if (dailyMap[dateStr] && (isDessert || isCompanion)) {
      dailyMap[dateStr].headcount += qty;
    }
  }
}

/* ============================================================
   手動清除快取 / 測試
   ============================================================ */
function clearCache() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(CACHE_KEY_PREFIX + 'chunks');
  if (chunkCount !== null) {
    var n = parseInt(chunkCount, 10);
    var keys = [CACHE_KEY_PREFIX + 'chunks'];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY_PREFIX + 'part_' + i);
    cache.removeAll(keys);
  }
  Logger.log('Cache cleared.');
}
function runClearCache() { clearCache(); }

/* ============================================================
   ★ v18 驗收函式（唯讀，跑完可留著當迴歸測試）
   ============================================================ */
function verifyV18() {
  clearCache();
  var j = JSON.parse(computeAggregation());

  Logger.log('════════ v18 驗收 ════════');
  Logger.log('【1. 全公司來客數】');
  Logger.log('  甜點數   = ' + j.kpi.dessertCount   + '   (v17 = 223,928 → 應約 228,098)');
  Logger.log('  陪同數   = ' + j.kpi.companionCount + '   (應完全不變 = 54,781)');
  Logger.log('  來客數   = ' + j.kpi.headcount      + '   (v17 = 278,709 → 應約 282,879)');
  Logger.log('  淨營收   = ' + j.kpi.totalRevenueNet + '   ★ 應與 v17 完全相同（net 不受影響）');

  Logger.log('【2. 未分類哨兵】← 這裡若出現新類別，就是下一顆地雷');
  if (!j.unknownCategories.length) {
    Logger.log('  ✅ 無未分類主類別');
  } else {
    j.unknownCategories.forEach(function(u){
      Logger.log('  ⚠️ ' + u.mainCategory + '  份數=' + u.qty + '  營收=' + u.revenue +
                 '  均價=' + (u.qty ? Math.round(u.revenue / u.qty) : 0));
    });
  }

  Logger.log('【3. 11 店 2026-06（實驗組，有麵包）】');
  var s11 = j.monthlyByStore.filter(function(m){ return m.yearMonth === '2026-06' && m.storeCode === 11; })[0];
  if (s11) {
    Logger.log('  甜點數=' + s11.dessertCount + '  陪同數=' + s11.companionCount + '  來客數=' + s11.headcount + '  (v17: 997 / 253 / 1,250)');
    Logger.log('  淨營收=' + s11.revenue_net + '  ★ 應恆為 693,903');
    Logger.log('  淨客單價=' + Math.round(s11.revenue_net / s11.headcount) + '  (v17 = 555 → 應約 498)');
  }

  Logger.log('【4. 7 店 2026-06（對照組，無麵包）】');
  var s7 = j.monthlyByStore.filter(function(m){ return m.yearMonth === '2026-06' && m.storeCode === 7; })[0];
  if (s7) {
    Logger.log('  來客數=' + s7.headcount + '  淨營收=' + s7.revenue_net +
               '  淨客單價=' + Math.round(s7.revenue_net / s7.headcount) + '  ★ 應完全不變（461）');
  }

  Logger.log('【5. 品項分析不得出現 $500券假商品】');
  var has500 = j.productRank.filter(function(p){ return String(p.product || '').indexOf('$500券') >= 0; });
  Logger.log('  $500券假商品筆數 = ' + has500.length + '（應為 0）');
  Logger.log('════════ 驗收結束 ════════');
}

function testAggregation() {
  var start = new Date();
  var result = computeAggregation();
  var elapsed = (new Date() - start) / 1000;
  var parsed = JSON.parse(result);
  Logger.log('執行時間: ' + elapsed + ' 秒');
  Logger.log('資料筆數: ' + parsed.totalRows);
  Logger.log('monthly 聚合筆數: ' + parsed.monthly.length);
  Logger.log('monthlyByStore 筆數: ' + parsed.monthlyByStore.length);
  Logger.log('商品種類數: ' + parsed.productRank.length);
  Logger.log('每日資料筆數: ' + parsed.daily.length);
  Logger.log('--- 三口徑(取自 view, v15 已扣$500券佔位) ---');
  Logger.log('gross:  ' + parsed.kpi.totalRevenueGross);
  Logger.log('折扣:   ' + parsed.kpi.totalDiscount);
  Logger.log('net:    ' + parsed.kpi.totalRevenueNet);
  Logger.log('--- v12 驗收 ---');
  Logger.log('birthdayCake 筆數: ' + parsed.birthdayCake.length);
  Logger.log('birthdayItems 筆數: ' + parsed.birthdayItems.length);
  var oreo2025 = parsed.birthdayItems.filter(function(x){ return x.yr===2025 && x.product==='Oreo巧克力'; });
  Logger.log('birthdayItems 2025 Oreo巧克力 qty: ' + (oreo2025.length ? oreo2025[0].qty : 'NOT FOUND'));
  Logger.log('--- v13 驗收 (折扣分券種) ---');
  Logger.log('discountByProgram 筆數: ' + parsed.discountByProgram.length);
  var mktSum = 0, allSum = 0, catSet = {};
  parsed.discountByProgram.forEach(function(x){
    allSum += x.discount; catSet[x.category] = true;
    if (x.category === '行銷折扣') mktSum += x.discount;
  });
  Logger.log('大類: ' + Object.keys(catSet).join(', '));
  Logger.log('行銷折扣加總: ' + mktSum + ' / 全部折扣加總(raw): ' + allSum + ' / kpi.totalDiscount(已扣$500券): ' + parsed.kpi.totalDiscount);
  var byV = {};
  parsed.discountByProgram.forEach(function(x){ if(x.category==='行銷折扣'){ byV[x.voucher]=(byV[x.voucher]||0)+x.discount; } });
  var topV = Object.keys(byV).map(function(k){return [k, byV[k]];}).sort(function(a,b){return b[1]-a[1];}).slice(0,3);
  Logger.log('行銷折扣前三: ' + topV.map(function(t){return t[0]+'='+t[1];}).join(' | '));
  Logger.log('--- v14 驗收 (4人同行真實折扣) ---');
  var totalSwapRev = 0, dbp4 = 0, dbp4ph = 0;
  parsed.discountByProgram.forEach(function(x){
    totalSwapRev += (x.placeholderRev || 0);
    if (x.voucher === '自己人4人同行500') { dbp4 += x.discount; dbp4ph += (x.placeholderRev || 0); }
  });
  Logger.log('$500券營收(placeholderRev)總額: ' + totalSwapRev + '（應≈1,049,500）');
  Logger.log('4人同行500 帳面折扣: ' + dbp4 + ' / 應扣$500券: ' + dbp4ph + ' / 真實折扣: ' + (dbp4 - dbp4ph));
  Logger.log('--- v15 驗收 (毛額/營收去虛胖) ---');
  Logger.log('全部折扣加總(raw) − $500券 = ' + (allSum - totalSwapRev) + '（應 = kpi.totalDiscount ' + parsed.kpi.totalDiscount + '）');
  Logger.log('net 錨點: ' + parsed.kpi.totalRevenueNet + '（舊錨點 133,309,665 已於 2026-07-13 作廢，改以線上實測為準）');
  var has500 = parsed.productRank.filter(function(p){ return String(p.product||'').indexOf('$500券') >= 0; });
  Logger.log('品項分析含$500券假商品筆數(應為0): ' + has500.length);
  Logger.log('--- v18 驗收 (來客數口徑) ---');
  Logger.log('甜點數: ' + parsed.kpi.dessertCount + ' / 陪同數: ' + parsed.kpi.companionCount + ' / 來客數: ' + parsed.kpi.headcount);
  Logger.log('未分類主類別筆數: ' + parsed.unknownCategories.length +
             (parsed.unknownCategories.length ? ' → ' + parsed.unknownCategories.map(function(u){return u.mainCategory+'('+u.qty+')';}).join(', ') : ' ✅'));
}

/* ============================================================
   v13：折扣分券種 — classifyVoucher + queryDiscountByProgram
   ============================================================ */

function classifyVoucher(raw) {
  var s = String(raw || '').trim();
  var MKT = '行銷折扣', ADD = '加價購', NON = '非行銷折扣';
  var rules = [
    [/壽星.*69折|當日壽星/, '當日壽星加價換款69折', MKT],
    [/職人.*生日禮|職人.*66折/, '職人級生日禮66折', MKT],
    [/達人.*生日禮|達人.*77折/, '達人級生日禮77折', MKT],
    [/素人.*生日禮|素人.*88折/, '素人級生日禮88折', MKT],
    [/12歲生日/, '12歲生日禮', MKT],
    [/免費.*體驗券/, '免費甜點體驗券', MKT],
    [/自己人.*免費.*甜點|自己人.*免費製作/, '自己人免費做一份甜點', MKT],
    [/體驗7折/, '自己人體驗7折', MKT],
    [/買一送一/, '自己人買一送一券', MKT],
    [/好客券/, '好客券', MKT],
    [/百貨員工/, '百貨員工優惠88折', MKT],
    [/員工.*7折/, '員工優惠7折', MKT],
    [/達人.*88折/, '達人級優惠88折', MKT],
    [/職人.*85折/, '職人級優惠85折', MKT],
    [/聚樂人.*9折|自己人.*素人.*9折|素人級.*9折|自己人.*9折/, '自己人/素人級優惠9折', MKT],
    [/特約廠商/, '特約廠商9折', MKT],
    [/貝登堡/, '貝登堡點數兌換券', MKT],
    [/湯姆熊/, '高雄SKM湯姆熊甜點兌換', MKT],
    [/京站/, '京站合作券', MKT],
    [/京華鑽石/, '京華鑽石合作', MKT],
    [/SKM.*999/, 'SKM蛋糕餅乾999', MKT],
    [/小島散步/, '小島散步合作', MKT],
    [/週年慶/, '週年慶檔期', MKT],
    [/第二[份件]半價/, '新品第二份半價', MKT],
    [/加購.*餅乾|加購優惠|餅乾加購|蛋糕加購/, '加購餅乾', MKT],   // v19：加價購→行銷折扣（經營者 2026-07-16 裁定，源頭正名；前端 v21 同步移除 remap）
    [/自己人.*100.*甜點券/, '自己人100甜點券', MKT],
    [/甜點券\$?1[05]0/, '甜點券100', MKT],
    [/任選.*甜點券|任選500/, '任選一份甜點券', MKT],
    [/1499/, '1499券', MKT],
    [/四人同行|4人同行|四位同行|4人.*500|四人.*500|每人500|每份.*500|500券|500卷|500元券|500團|500折/, '自己人4人同行500', MKT],
    [/手動折扣|客服備注/, '手動折扣', NON],
    [/客訴|招待|免單/, '服務補償', NON],
    [/馬達|鐵門|機器|移動客人|改場次|配合包館|配合包場/, '營運調整', NON],
    [/訂金|付訂|預付|預售|預收/, '訂金', NON],
    [/包館|包場|團體優惠|團$/, '包場團體', NON]
  ];
  for (var i = 0; i < rules.length; i++) {
    if (rules[i][0].test(s)) return { voucher: rules[i][1], category: rules[i][2] };
  }
  return { voucher: '其他未分類', category: '其他' };
}

function queryDiscountByProgram() {
  var token = _getBqAccessToken_();
  var sql = 'SELECT FORMAT_DATE("%Y-%m", sale_date) AS ym, ' +
            'store_code, project_name, ' +
            'COUNT(*) AS uses, ' +
            'COUNT(DISTINCT serial_no) AS txns, ' +
            'CAST(SUM(discount) AS INT64) AS discount ' +
            'FROM `diybc-make-sync.diybc_pos.pos_discounts` ' +
            'GROUP BY ym, store_code, project_name';
  var base = 'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries';
  var res = UrlFetchApp.fetch(base, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 60000, maxResults: 50000 }),
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('BQ discountByProgram 查詢錯誤: ' + JSON.stringify(data.error));

  var out = [];
  function _collect(d) {
    if (!d.rows) return;
    for (var i = 0; i < d.rows.length; i++) {
      var f = d.rows[i].f;
      out.push({
        ym: f[0].v,
        store_code: Number(f[1].v),
        project_name: f[2].v,
        uses: Number(f[3].v),
        txns: Number(f[4].v),
        discount: Number(f[5].v)
      });
    }
  }
  _collect(data);

  var jobId = data.jobReference && data.jobReference.jobId;
  var pageToken = data.pageToken;
  var guard = 0;
  while (pageToken && jobId && guard < 20) {
    guard++;
    var r2 = UrlFetchApp.fetch(
      base + '/' + jobId + '?pageToken=' + encodeURIComponent(pageToken) + '&maxResults=50000',
      { method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var d2 = JSON.parse(r2.getContentText());
    if (d2.error) throw new Error('BQ discountByProgram 分頁錯誤: ' + JSON.stringify(d2.error));
    _collect(d2);
    pageToken = d2.pageToken;
  }
  return out;
}

function queryVoucherProductRevenue() {
  var token = _getBqAccessToken_();
  var sql = 'SELECT FORMAT_DATE("%Y-%m-%d", sale_date) AS d, store_code, ' +
            'CAST(SUM(total_amount) AS INT64) AS revenue ' +
            'FROM `diybc-make-sync.diybc_pos.pos_transactions` ' +
            'WHERE product_name LIKE "%$500券%" ' +
            'GROUP BY d, store_code';
  var base = 'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries';
  var res = UrlFetchApp.fetch(base, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 60000, maxResults: 50000 }),
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('BQ voucherProductRevenue 查詢錯誤: ' + JSON.stringify(data.error));

  var out = [];
  function _collect(d) {
    if (!d.rows) return;
    for (var i = 0; i < d.rows.length; i++) {
      var f = d.rows[i].f;
      out.push({ date: f[0].v, store_code: Number(f[1].v), revenue: Number(f[2].v) });
    }
  }
  _collect(data);

  var jobId = data.jobReference && data.jobReference.jobId;
  var pageToken = data.pageToken;
  var guard = 0;
  while (pageToken && jobId && guard < 20) {
    guard++;
    var r2 = UrlFetchApp.fetch(
      base + '/' + jobId + '?pageToken=' + encodeURIComponent(pageToken) + '&maxResults=50000',
      { method: 'get', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );
    var d2 = JSON.parse(r2.getContentText());
    if (d2.error) throw new Error('BQ voucherProductRevenue 分頁錯誤: ' + JSON.stringify(d2.error));
    _collect(d2);
    pageToken = d2.pageToken;
  }
  return out;
}

function keepWarm() {
  const url = 'https://docs.google.com/spreadsheets/d/1hZ8hpzrCNycrpm64Ge0fVxWDV-L__bxiagXk0btrBBQ/gviz/tq?tqx=out:json&tq=' + encodeURIComponent("SELECT COUNT(A) LIMIT 1");
  UrlFetchApp.fetch(url);
}

/* ============================================================
   v5.0 對帳驗證（唯讀，跑完可刪）
   ============================================================ */
function verifyV5_POS() {
  var rows = queryDailyNet('2026-04-01', '2026-04-30');
  var net1 = 0, gross1 = 0;
  rows.forEach(function(r){
    if(r.store_code === 1){ net1 += r.net; gross1 += r.gross; }
  });
  Logger.log('1店2026-04 gross=' + gross1 + ' net=' + net1 + ' (應 gross=393258 net=369921)');
}

/**
 * ===========================================================
 * 限定甜點檔期分析（seasonal）— 維持 Sheet 來源，不變
 * ===========================================================
 */
function parseCampaignSubcat(subcat) {
  if (!subcat) return null;
  var s = String(subcat).trim();
  var m = s.match(/^(20\d{2})\s*(\S.*)$/);
  if (!m) return null;
  return { year: parseInt(m[1]), name: m[2].trim(), key: m[1] + ' ' + m[2].trim() };
}

function _extractDate(raw) {
  if (!raw) return null;
  if (raw instanceof Date) {
    var y = raw.getFullYear();
    var m = ('0' + (raw.getMonth() + 1)).slice(-2);
    var d = ('0' + raw.getDate()).slice(-2);
    return y + '-' + m + '-' + d;
  }
  var s = String(raw);
  var mm = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!mm) return null;
  return mm[1] + '-' + ('0' + mm[2]).slice(-2) + '-' + ('0' + mm[3]).slice(-2);
}

function _findCol(headers, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var idx = headers.indexOf(candidates[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

function _normalizeStore(s) {
  if (s === null || s === undefined) return '未知';
  var v = String(s).trim();
  v = v.replace(/^["'""「」『』]+/, '').replace(/["'""「」『』]+$/, '');
  v = v.replace(/^　+/, '').replace(/　+$/, '');
  return v.trim() || '未知';
}

function buildSeasonalAnalysis(data) {
  if (!data || data.length < 2) return { campaigns: [], asOf: new Date().toISOString() };

  var headers = data[0];
  var col = {
    date:      _findCol(headers, ['建立日期', '日期', '時間']),
    code:      _findCol(headers, ['分店代碼', '門市代碼', '店號']),
    store:     _findCol(headers, ['分店名稱', '門市', '分店']),
    subCat:    _findCol(headers, ['次類別', '子類別']),
    revenue:   _findCol(headers, ['實收總額', '實收金額', '銷售額']),
    product:   _findCol(headers, ['商品名稱', '品項', '商品']),
    qty:       _findCol(headers, ['數量', '銷售數量']),
    unitPrice: _findCol(headers, ['商品單價', '單價'])
  };
  if (col.date < 0 || col.subCat < 0 || col.revenue < 0) {
    throw new Error('buildSeasonalAnalysis 找不到必要欄位 ' + JSON.stringify(col));
  }

  var campaignMap = {};
  var today = new Date().toISOString().slice(0, 10);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row) continue;
    var camp = parseCampaignSubcat(row[col.subCat]);
    if (!camp) continue;
    var dateStr = _extractDate(row[col.date]);
    if (!dateStr) continue;
    var rev = parseFloat(row[col.revenue]) || 0;
    var store;
    if (col.code >= 0 && Number(row[col.code])) {
      store = _dim(Number(row[col.code])).name;
    } else {
      store = _normalizeStore(col.store >= 0 ? row[col.store] : '未知');
    }
    if (store === '未知') continue;

    if (!campaignMap[camp.key]) {
      campaignMap[camp.key] = { key: camp.key, year: camp.year, name: camp.name, dates: {}, revenue: 0, byStore: {}, itemMap: {} };
    }
    var c = campaignMap[camp.key];
    c.dates[dateStr] = true;
    c.revenue += rev;
    if (!c.byStore[store]) c.byStore[store] = 0;
    c.byStore[store] += rev;

    if (col.product >= 0) {
      var prodName = String(row[col.product] || '').trim() || '(未命名品項)';
      var prodQty = (col.qty >= 0) ? (parseFloat(row[col.qty]) || 0) : 0;
      var prodUnitPrice = (col.unitPrice >= 0) ? (parseFloat(row[col.unitPrice]) || 0) : 0;
      if (!c.itemMap[prodName]) c.itemMap[prodName] = { product: prodName, qty: 0, revenue: 0, grossRevenue: 0 };
      c.itemMap[prodName].qty += prodQty;
      c.itemMap[prodName].revenue += rev;
      c.itemMap[prodName].grossRevenue += prodUnitPrice * prodQty;
    }
  }

  var campaigns = [];
  for (var key in campaignMap) {
    var camp = campaignMap[key];
    var dateList = Object.keys(camp.dates).sort();
    if (dateList.length === 0) continue;
    var startDate = dateList[0];
    var endDate = dateList[dateList.length - 1];
    var days = dateList.length;

    var byStore = [];
    for (var store in camp.byStore) byStore.push({ store: store, revenue: Math.round(camp.byStore[store]) });
    byStore.sort(function (a, b) { return b.revenue - a.revenue; });

    var campTotalRev = camp.revenue;
    var items = [];
    for (var prod in camp.itemMap) {
      var it = camp.itemMap[prod];
      items.push({
        product: it.product, qty: Math.round(it.qty), revenue: Math.round(it.revenue),
        originalUnitPrice: it.qty > 0 ? Math.round(it.grossRevenue / it.qty) : 0,
        revenuePct: campTotalRev > 0 ? Math.round(it.revenue / campTotalRev * 1000) / 10 : 0
      });
    }
    items.sort(function (a, b) { return b.revenue - a.revenue; });

    campaigns.push({
      key: camp.key, year: camp.year, name: camp.name,
      dateRange: { start: startDate, end: endDate, days: days },
      revenue: Math.round(camp.revenue), yoy: null,
      ongoing: startDate <= today && today <= endDate,
      byStore: byStore, items: items
    });
  }

  var byKey = {};
  for (var i = 0; i < campaigns.length; i++) byKey[campaigns[i].key] = campaigns[i];
  for (var i = 0; i < campaigns.length; i++) {
    var c = campaigns[i];
    var prevKey = (c.year - 1) + ' ' + c.name;
    var prev = byKey[prevKey];
    if (prev && prev.revenue > 0) c.yoy = Math.round(((c.revenue - prev.revenue) / prev.revenue) * 1000) / 1000;
    if (prev) {
      var prevStoreRev = {};
      for (var j = 0; j < prev.byStore.length; j++) prevStoreRev[prev.byStore[j].store] = prev.byStore[j].revenue;
      for (var k = 0; k < c.byStore.length; k++) {
        var bs = c.byStore[k];
        var pr = prevStoreRev[bs.store];
        bs.yoy = (pr && pr > 0) ? Math.round(((bs.revenue - pr) / pr) * 1000) / 1000 : null;
      }
    } else {
      for (var k = 0; k < c.byStore.length; k++) c.byStore[k].yoy = null;
    }
  }

  campaigns.sort(function (a, b) {
    if (b.year !== a.year) return b.year - a.year;
    return b.revenue - a.revenue;
  });

  return { campaigns: campaigns, asOf: new Date().toISOString() };
}

/* ============================================================
   ★ v19：檔期 detail 立方體（品項 × 店 × 週，週一為始日曆週）
   attachSeasonalDetail_() 由 computeAggregation 掛載
   驗收：verifySeasonalV19()（部署前必跑）
   ============================================================ */
var SEASONAL_DETAIL_MAX_ITEMS = 0; // 品項封頂閘門：0=不限。若 payload 快取塊數比 v18 基準多 >5 塊 → 改 15，超出品項併入「其他」

/** 回傳該日期所屬週的「週一」日期字串（用年月日分量建 Date，避開時區陷阱） */
function _weekStartMonday_(dateStr) {
  var p = String(dateStr).split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  var diff = (d.getDay() + 6) % 7; // Mon=0 ... Sun=6
  d.setDate(d.getDate() - diff);
  var m = d.getMonth() + 1, dd = d.getDate();
  return d.getFullYear() + '-' + (m < 10 ? '0' + m : m) + '-' + (dd < 10 ? '0' + dd : dd);
}

/**
 * 對每個檔期掛上 detail 立方體。
 * 重掃同一份 data（記憶體內陣列，無額外 Sheet 讀取），沿用既有全域 helper：
 * _findCol / parseCampaignSubcat / _extractDate（皆在 discount_aggregator.gs）。
 */
function attachSeasonalDetail_(data, seasonal) {
  if (!seasonal || !seasonal.campaigns || !seasonal.campaigns.length) return seasonal;
  if (!data || data.length < 2) return seasonal;

  var headers = data[0];
  var col = {
    date:    _findCol(headers, ['建立日期', '日期', '時間']),
    code:    _findCol(headers, ['分店代碼', '門市代碼', '店號']),
    subCat:  _findCol(headers, ['次類別', '子類別']),
    product: _findCol(headers, ['商品名稱', '品項', '商品']),
    qty:     _findCol(headers, ['數量', '銷售數量'])
  };
  // 缺必要欄位：不掛 detail、不炸主 payload（前端顯示「無明細」）
  if (col.date < 0 || col.code < 0 || col.subCat < 0 || col.product < 0 || col.qty < 0) return seasonal;

  // camp.key → { cube:{店號:{週一:{品項:qty}}}, weekSet:{}, storeSet:{}, noStoreQty }
  var acc = {};
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row) continue;
    var camp = parseCampaignSubcat(row[col.subCat]);
    if (!camp) continue;
    var dateStr = _extractDate(row[col.date]);
    if (!dateStr) continue;
    var qty = parseFloat(row[col.qty]) || 0;
    if (!qty) continue; // 0 不影響加總；負數退貨列保留
    var prod = String(row[col.product] || '').trim() || '(未命名品項)'; // 與 buildSeasonalAnalysis 同一 fallback 名
    var a = acc[camp.key];
    if (!a) a = acc[camp.key] = { cube: {}, weekSet: {}, storeSet: {}, noStoreQty: 0 };
    var scode = Number(row[col.code]) || 0;
    if (!scode) { a.noStoreQty += qty; continue; }
    var wk = _weekStartMonday_(dateStr);
    a.weekSet[wk] = true;
    a.storeSet[scode] = true;
    if (!a.cube[scode]) a.cube[scode] = {};
    if (!a.cube[scode][wk]) a.cube[scode][wk] = {};
    a.cube[scode][wk][prod] = (a.cube[scode][wk][prod] || 0) + qty;
  }

  for (var ci = 0; ci < seasonal.campaigns.length; ci++) {
    var c = seasonal.campaigns[ci];
    var a2 = acc[c.key];
    if (!a2) continue;

    // 品項軸：沿用 items[] 排序（營收降冪）；封頂時尾端加「其他」
    var CAP = SEASONAL_DETAIL_MAX_ITEMS;
    var itemNames = [], otherIdx = -1, capped = false;
    if (CAP > 0 && c.items.length > CAP) {
      for (var t = 0; t < CAP; t++) itemNames.push(c.items[t].product);
      itemNames.push('其他');
      otherIdx = CAP;
      capped = true;
    } else {
      for (var t2 = 0; t2 < c.items.length; t2++) itemNames.push(c.items[t2].product);
    }
    var nameToIdx = {};
    for (var ni = 0; ni < itemNames.length; ni++) nameToIdx[itemNames[ni]] = ni;

    var weeks = Object.keys(a2.weekSet).sort();
    var stores = Object.keys(a2.storeSet).map(Number).sort(function (x, y) { return x - y; });

    var qtyCube = {};
    for (var si = 0; si < stores.length; si++) {
      var sc = stores[si];
      var mat = [];
      for (var wi = 0; wi < weeks.length; wi++) {
        var rowArr = [];
        for (var ii = 0; ii < itemNames.length; ii++) rowArr.push(0);
        mat.push(rowArr);
      }
      var srcStore = a2.cube[sc] || {};
      for (var wi2 = 0; wi2 < weeks.length; wi2++) {
        var wkObj = srcStore[weeks[wi2]];
        if (!wkObj) continue;
        for (var pn in wkObj) {
          var idx = nameToIdx.hasOwnProperty(pn) ? nameToIdx[pn] : otherIdx;
          if (idx < 0) continue;
          mat[wi2][idx] += wkObj[pn];
        }
      }
      for (var wi3 = 0; wi3 < mat.length; wi3++)
        for (var ii3 = 0; ii3 < mat[wi3].length; ii3++)
          mat[wi3][ii3] = Math.round(mat[wi3][ii3]);
      qtyCube[String(sc)] = mat;
    }

    c.detail = {
      weekStart: 'mon',
      weeks: weeks,
      stores: stores,
      items: itemNames,
      capped: capped,
      noStoreQty: Math.round(a2.noStoreQty),
      qty: qtyCube
    };
  }
  return seasonal;
}

/* ═══════════════════════════════════════════════════════════════
 * v19 驗收（部署前在編輯器跑，實測就是對）
 * 檢查：①四行定點修改是否生效 ②payload 量測 ③立方體不變量
 *       ④週標籤全是週一 ⑤口徑不變量提示（接著跑 verifyV18）
 * ═══════════════════════════════════════════════════════════════ */
function verifySeasonalV19() {
  clearCache();
  var jsonStr = computeAggregation();
  var j = JSON.parse(jsonStr);
  Logger.log('══════════ v19 驗收 ══════════');

  Logger.log('【0. 四行定點修改是否生效】');
  Logger.log('  快取前綴 = ' + CACHE_KEY_PREFIX + '（應 POS_DASHBOARD_V19_）');
  Logger.log('  apiVersion = ' + j.apiVersion + '（應 v5.1）');
  Logger.log('  classifyVoucher("蛋糕加購") = ' + JSON.stringify(classifyVoucher('蛋糕加購')) + '（應 category=行銷折扣）');
  var hasDetail = j.seasonal && j.seasonal.campaigns && j.seasonal.campaigns.length && j.seasonal.campaigns[0].detail;
  Logger.log('  campaigns[0].detail ' + (hasDetail ? '✅ 存在' : '❌ 不存在 → computeAggregation 的 seasonal 那行沒改到'));

  Logger.log('【A. payload 量測（閘門：快取塊數比 v18 基準多 >5 → SEASONAL_DETAIL_MAX_ITEMS 改 15 重跑）】');
  Logger.log('  總長度 = ' + jsonStr.length + ' 字元 → 快取塊數 = ' + Math.ceil(jsonStr.length / 90000));
  Logger.log('  seasonal 長度 = ' + JSON.stringify(j.seasonal).length + ' 字元');

  Logger.log('【B. 立方體不變量：cube 加總 + 無店號 ＝ items 加總】');
  var bad = 0, camps = j.seasonal.campaigns;
  for (var i = 0; i < camps.length; i++) {
    var c = camps[i];
    if (!c.detail) { Logger.log('  ⚠️ ' + c.key + ' 無 detail'); bad++; continue; }
    var cubeSum = 0;
    for (var sc in c.detail.qty) {
      var mat = c.detail.qty[sc];
      for (var w = 0; w < mat.length; w++)
        for (var it = 0; it < mat[w].length; it++) cubeSum += mat[w][it];
    }
    var itemSum = 0;
    for (var k = 0; k < c.items.length; k++) itemSum += (c.items[k].qty || 0);
    var diff = itemSum - cubeSum - (c.detail.noStoreQty || 0);
    var tol = Math.max(3, Math.round(itemSum * 0.001));
    var flag = (diff === 0) ? '✅' : (Math.abs(diff) <= tol ? '🟡' : '❌');
    if (flag === '❌') bad++;
    Logger.log('  ' + flag + ' ' + c.key + '｜items=' + itemSum + ' cube=' + cubeSum +
               ' 無店號=' + (c.detail.noStoreQty || 0) + ' 差=' + diff +
               '｜店' + c.detail.stores.length + '×週' + c.detail.weeks.length +
               '×品項' + c.detail.items.length + (c.detail.capped ? '(封頂)' : ''));
  }

  Logger.log('【C. 週一檢查（掃全部週標籤）】');
  var wkBad = 0;
  for (var i2 = 0; i2 < camps.length; i2++) {
    var d2 = camps[i2].detail;
    if (!d2) continue;
    for (var w2 = 0; w2 < d2.weeks.length; w2++) {
      var p = d2.weeks[w2].split('-');
      if (new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])).getDay() !== 1) wkBad++;
    }
  }
  Logger.log(wkBad === 0 ? '  ✅ 全部週標籤都是週一' : '  ❌ 有 ' + wkBad + ' 個週標籤不是週一');

  Logger.log('【D. 口徑不變量】接著跑 verifyV18()：net 級別不變／7店對照組一字不變／未分類哨兵為空');
  Logger.log((bad === 0 && wkBad === 0)
    ? '══════ v19 不變量全過 → 跑完 verifyV18 再部署 ══════'
    : '══════ ❌ 未過，別部署，把完整 log 貼給 Claude ══════');
}

/**
 * 清快取 + 重算 + 寫回快取（強制吃最新 view / 最新白名單）
 * ⚠️ v18 修正：原本寫死 'POS_DASHBOARD_V12_' 前綴，寫進去的快取永遠讀不到（既有 bug）。
 */
function rebuildCacheAndCheck() {
  clearCache();
  var jsonStr = computeAggregation();
  var cache = CacheService.getScriptCache();
  var chunkSize = 90000;
  var chunks = [];
  for (var i = 0; i < jsonStr.length; i += chunkSize) chunks.push(jsonStr.substring(i, i + chunkSize));
  var cacheObj = {};
  for (var i = 0; i < chunks.length; i++) cacheObj[CACHE_KEY_PREFIX + 'part_' + i] = chunks[i];   // v18 修正
  cacheObj[CACHE_KEY_PREFIX + 'chunks'] = String(chunks.length);                                  // v18 修正
  cache.putAll(cacheObj, CACHE_TTL);
  var j = JSON.parse(jsonStr);
  var bx = j.monthlyByStore.filter(function(m){ return m.yearMonth==='2026-05' && m.storeCode==7; });
  Logger.log('快取已重建（前綴 ' + CACHE_KEY_PREFIX + '，' + chunks.length + ' 塊）');
  Logger.log('重算後 板橋5月: ' + JSON.stringify(bx));
}

function whatDoesDoGetReturn() {
  var result = getCachedOrCompute();
  var j = JSON.parse(result);
  var bx = j.monthlyByStore.filter(function(m){ return m.yearMonth==='2026-05' && m.storeCode==7; })[0];
  Logger.log('getCachedOrCompute 板橋5月: gross=' + bx.revenue_gross + ' net=' + bx.revenue_net + ' actual=' + bx.revenue_actual);
  Logger.log('快取前綴 = ' + CACHE_KEY_PREFIX);
}

/* ============================================================
   daily_by_store 獨立 endpoint（日×店明細，2026-06-25 新增）
   ============================================================ */
var DAILY_BY_STORE_CACHE_PREFIX = 'POS_DAILY_BY_STORE_V2_';   // v18：白名單改動，同步換版
var DAILY_BY_STORE_CACHE_TTL = 3600;

function getDailyByStoreCached() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(DAILY_BY_STORE_CACHE_PREFIX + 'chunks');
  if (chunkCount) {
    var n = Number(chunkCount);
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(DAILY_BY_STORE_CACHE_PREFIX + 'part_' + i);
    var parts = cache.getAll(keys);
    var allPresent = true;
    for (var a = 0; a < n; a++) { if (!parts[DAILY_BY_STORE_CACHE_PREFIX + 'part_' + a]) { allPresent = false; break; } }
    if (allPresent) {
      var assembled = '';
      for (var b = 0; b < n; b++) assembled += parts[DAILY_BY_STORE_CACHE_PREFIX + 'part_' + b];
      return assembled;
    }
  }
  var rows = computeDailyByStore().rows;
  var jsonStr = JSON.stringify(rows);
  var chunkSize = 90000;
  var chunks = [];
  for (var j = 0; j < jsonStr.length; j += chunkSize) chunks.push(jsonStr.substring(j, j + chunkSize));
  var cacheObj = {};
  for (var k = 0; k < chunks.length; k++) cacheObj[DAILY_BY_STORE_CACHE_PREFIX + 'part_' + k] = chunks[k];
  cacheObj[DAILY_BY_STORE_CACHE_PREFIX + 'chunks'] = String(chunks.length);
  cache.putAll(cacheObj, DAILY_BY_STORE_CACHE_TTL);
  return jsonStr;
}

function clearDailyByStoreCache() {
  var cache = CacheService.getScriptCache();
  var chunkCount = cache.get(DAILY_BY_STORE_CACHE_PREFIX + 'chunks');
  var keys = [DAILY_BY_STORE_CACHE_PREFIX + 'chunks'];
  if (chunkCount) { for (var i = 0; i < Number(chunkCount); i++) keys.push(DAILY_BY_STORE_CACHE_PREFIX + 'part_' + i); }
  cache.removeAll(keys);
  Logger.log('daily_by_store cache cleared');
}

function computeDailyByStore() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('POS資料');
  var data = sheet.getDataRange().getValues();
  var dessertSet = {};
  for (var i = 0; i < DESSERT_CATEGORIES.length; i++) dessertSet[DESSERT_CATEGORIES[i]] = true;
  var itemMap = {};
  var blankStore = { rows: 0, qty: 0, revenue: 0 };
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var storeCode = Number(row[0]);
    var dateRaw = row[2];
    var product = String(row[3] || '');
    var mainCat = String(row[4] || '');
    var unitPrice = Number(row[6]) || 0;
    var qty = Number(row[7]) || 0;
    var revenue = unitPrice * qty;
    if (!product) continue;
    if (product.indexOf('$500券') >= 0) continue;
    var dateObj = (dateRaw instanceof Date) ? dateRaw : new Date(String(dateRaw));
    if (isNaN(dateObj.getTime())) continue;
    var y = dateObj.getFullYear(), mo = dateObj.getMonth() + 1, d = dateObj.getDate();
    var dateStr = y + '-' + (mo < 10 ? '0' + mo : mo) + '-' + (d < 10 ? '0' + d : d);
    var isCompanion = (product === '陪同入場費');                      // v18：先判陪同
    var isDessert = !isCompanion && (dessertSet[mainCat] === true);    // v18：互斥
    var isLimited = (mainCat === '限定甜點');
    if (!storeCode) {
      if (isDessert || isCompanion) { blankStore.rows++; blankStore.qty += qty; blankStore.revenue += revenue; }
      continue;
    }
    var key = dateStr + '|' + storeCode;
    var it = itemMap[key];
    if (!it) { it = itemMap[key] = { dessert_count: 0, companion_count: 0, limited_revenue: 0, item_revenue: 0 }; }
    it.item_revenue += revenue;
    if (isDessert) it.dessert_count += qty;
    if (isCompanion) it.companion_count += qty;
    if (isLimited) it.limited_revenue += revenue;
  }
  var netRows = queryDailyNet();
  var swapRevRows = queryVoucherProductRevenue();
  var swapByDayStore = {};
  for (var s = 0; s < swapRevRows.length; s++) {
    var sv = swapRevRows[s];
    var sk = sv.date + '|' + sv.store_code;
    swapByDayStore[sk] = (swapByDayStore[sk] || 0) + (sv.revenue || 0);
  }
  var out = [];
  for (var nn = 0; nn < netRows.length; nn++) {
    var nr = netRows[nn];
    var dStr = nr.sale_date, sc = nr.store_code;
    var g = (nr.gross || 0), disc = (nr.total_discount || 0), net = (nr.net || 0);
    var ph = swapByDayStore[dStr + '|' + sc] || 0;
    g = g - ph; disc = disc - ph;
    var dim = _dim(sc);
    var it2 = itemMap[dStr + '|' + sc] || { dessert_count: 0, companion_count: 0, limited_revenue: 0, item_revenue: 0 };
    out.push({
      date: dStr,
      store: dim.name,
      region: dim.region,
      headcount: it2.dessert_count + it2.companion_count,
      dessert_count: it2.dessert_count,
      companion_count: it2.companion_count,
      limited_revenue: Math.round(it2.limited_revenue),
      item_revenue: Math.round(it2.item_revenue),
      revenue_gross: Math.round(g),
      revenue_net: Math.round(net),
      discount: Math.round(disc)
    });
  }
  out.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.store < b.store ? -1 : 1;
  });
  return { rows: out, blankStore: blankStore };
}

function verifyDailyByStore() {
  var res = computeDailyByStore();
  var rows = res.rows;
  Logger.log('[watch] blankStore rows=' + res.blankStore.rows + ' qty=' + res.blankStore.qty + ' revenue=' + Math.round(res.blankStore.revenue));
  var json = JSON.stringify(rows);
  Logger.log('[size] rows=' + rows.length + ' jsonLen=' + json.length + ' chunks=' + Math.ceil(json.length / 90000));
  var bad = 0;
  for (var i = 0; i < rows.length; i++) { if (rows[i].headcount !== rows[i].dessert_count + rows[i].companion_count) bad++; }
  Logger.log('[chk3] headcount!==dessert+companion violations=' + bad);
  var D = '2026-06-20';
  var sumNet = 0, sumGross = 0, sumDisc = 0, sumHc = 0;
  for (var j = 0; j < rows.length; j++) { if (rows[j].date === D) { sumNet += rows[j].revenue_net; sumGross += rows[j].revenue_gross; sumDisc += rows[j].discount; sumHc += rows[j].headcount; } }
  Logger.log('[chk2-mine] ' + D + ' stores-sum net=' + sumNet + ' gross=' + sumGross + ' discount=' + sumDisc + ' headcount=' + sumHc);
  var netRows = queryDailyNet(D, D);
  var swap = queryVoucherProductRevenue();
  var phByDay = 0;
  for (var s = 0; s < swap.length; s++) { if (swap[s].date === D) phByDay += (swap[s].revenue || 0); }
  var vNet = 0, vGross = 0, vDisc = 0;
  for (var k = 0; k < netRows.length; k++) { vGross += (netRows[k].gross || 0); vDisc += (netRows[k].total_discount || 0); vNet += (netRows[k].net || 0); }
  vGross -= phByDay; vDisc -= phByDay;
  Logger.log('[chk2-view] ' + D + ' view-sum net=' + Math.round(vNet) + ' gross=' + Math.round(vGross) + ' discount=' + Math.round(vDisc));
  for (var m = 0; m < rows.length; m++) { if (rows[m].date === D && rows[m].store === _dim(7).name) { Logger.log('[chk4] store7 ' + D + ' ' + JSON.stringify(rows[m])); } }
}
