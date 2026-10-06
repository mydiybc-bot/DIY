/**
 * ═══════════════════════════════════════════════════════════════
 * seasonal_detail.gs — v19（2026-07-22）★ 全新檔案，整檔貼上 ★
 * 檔期分析 detail 立方體：品項 × 店 × 週（週一為始的日曆週）
 * ═══════════════════════════════════════════════════════════════
 *
 * 【掛載】discount_aggregator.gs 的 computeAggregation() 裡
 *     seasonal: buildSeasonalAnalysis(data),
 *   改為
 *     seasonal: attachSeasonalDetail_(data, buildSeasonalAnalysis(data)),
 *
 * 【產出】campaigns[i].detail = {
 *     weekStart: 'mon',
 *     weeks:  ['YYYY-MM-DD', ...],  // 各週「週一」日期，升冪
 *     stores: [1, 2, ...],          // 有資料的店號，升冪
 *     items:  ['品項名', ...],       // 與 items[] 同序（營收降冪）；封頂時尾端為「其他」
 *     capped: false,                // 是否觸發品項封頂
 *     noStoreQty: 0,                // 無店號、進不了 cube 的數量（應≈0）
 *     qty: { '店號': 週×品項 二維整數矩陣 }
 *   }
 *
 * 【口徑】SUM(數量)，含負數退貨列；qty=0 列跳過（不影響加總）。
 *   無店號列不入 cube（campaign 營收 / items 照舊計入），差額由驗收函式監控。
 * 【不變量】各檔期：cube 全加總 + noStoreQty ＝ items[].qty 加總（verifySeasonalV19 檢查）
 * 【欄位偵測】與 buildSeasonalAnalysis 完全同一套候選名，確保同一批列進同一個檔期。
 */

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