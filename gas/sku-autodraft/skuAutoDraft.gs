/*******************************************************************
 * diybc-sku-autodraft  v1.2  (2026-09-17)
 * v1.2（方案 A，經營者裁示「不可無中生有」）：只替「產品名稱對照表裡已上架或已排定上架」的甜點建草稿。
 *   事故：9/15 BOM 表出現「《吶喊》吧！餅乾」的配方（對照表沒有、POS 沒賣過、今年萬聖檔期也還沒建），
 *         v1.1 仍替它的「萬聖節餅乾紙墊」自動在主檔建了草稿 C096。
 *   規則：BOM 對不上主檔的新料，至少要有一款用到它的甜點「在對照表」且「沒有結束日／結束日未過／結束不到 30 天」才建草稿。
 *         否則列「⏸ 甜點還沒上架，先不建草稿」清單；對照表建好後，隔天早上自動建。
 *   讀不到對照表（分頁不見、欄名改掉）⇒ 一律不建草稿（寧可少建，不可亂建），加別名／清標記／最後有效日照常。
 *   既有草稿若用到它的甜點都還沒上架 ⇒ 只列提醒，絕不自動刪除（刪除前要先刪 BOM 列，否則隔天會再被建回來）。
 *   「暫緩清單」有變動才寫 sku_autodraft_log（避免每天重複一列）。其餘 A 別名／C 清標記／D 最後有效日與 v1.1 逐行相同。
 *******************************************************************/
/*******************************************************************
 * diybc-sku-autodraft  v1.1  (2026-09-01 晚)
 * v1.1 亦：BOM 名含「下架／已下架／測試」不建草稿，改列 ⚠ 請刪 BOM 列；試跑會列「過季但近 4 週仍有用量」的矛盾清單（對照表結束日可能填錯）。
 * v1.1 新增 D) 最後有效日：每個 SKU ＝ 用到它的甜點（BOM表）在「產品名稱對照表」的最晚「結束有效日」。
 *   有任一甜點沒有結束日（常態品／對照表沒列）→ 視為長期有效 → 寫空白。沒有任何甜點用到 → 空白。
 *   寫入 dim_sku 新欄「最後有效日」（沒有就在最右建；欄數不夠自動插欄）。前端：最後有效日 < 今天 ⇒ 盤點頁「🗓 過季」籤、建議/訂貨不列。
 *   純覆蓋欄，不動其他欄；試跑不寫。
 * 用途：BOM 表出現主檔 dim_sku 沒有的食材/器具名 → 自動處理，不再靠人工新增：
 *   A) 只是寫法不同（空格、全半形括號、主檔多了代碼「A02 」前綴）→ 自動加進該 SKU 的「BOM別名」
 *   B) 真的沒有 → 自動新增「草稿列」：品名＝BOM 原字串、使用單位＝BOM 單位、品類別依單位＋關鍵字推測、
 *      內容量 1、採購單位＝使用單位、編號自動續編，新欄「來源」寫「自動草稿 YYYY-MM-DD」
 *   C) 草稿列的 廠商／內容量(>1 或採購單位≠使用單位)／單價 三格都填了 → 自動清掉「來源」標記（＝上架完成）
 * 草稿當天就進引擎：用量立刻算得到；訂貨單先以使用單位顯示，填了內容量就變成採購單位。
 * 部署：⚠️ 獨立新 GAS 專案（第 5 個），不可加進其他專案
 *   1. script.google.com → 新專案 → 命名 diybc-sku-autodraft → 貼本檔 → 儲存
 *   2. 函式 autoDraftDryRun → 執行（授權）→ 記錄看「將加別名 / 將新增草稿」清單
 *   3. 函式 autoDraftNow → 執行 → 寫入
 *   4. 觸發條件：autoDraftNow、時間驅動、日計時器、上午 6～7 點（要在 07:05 rebuild 之前）
 *******************************************************************/
var AD_SHEET_ID = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';
var AD_BOM_TAB = 'BOM表', AD_SKU_TAB = 'dim_sku', AD_LOG_TAB = 'sku_autodraft_log';
var AD_SRC_COL = '來源';                       // dim_sku 新欄（最右邊自動建）
var AD_MAP_TAB = '產品名稱對照表', AD_LV_COL = '最後有效日';   // v1.1
var AD_SKIP_RE = /下架|測試/;                 // v1.1：BOM 名含這些字 → 不建草稿，提示刪 BOM 列
var AD_AGG_TAB = 'agg_purchase';
var AD_ENDED_GRACE_DAYS = 30;                 // v1.2：甜點結束不到 N 天仍算「在賣」（對照表結束日常比實際停賣早）
var AD_HELD_PROP = 'AD_HELD_SIG';             // v1.2：暫緩清單簽章（有變動才寫 log）
var AD_PREFIX = { '食材': 'F', '耗材': 'C', '器具': 'T', '模具': 'T' };
var AD_COUNT_UNITS = { '個': 1, '支': 1, '張': 1, '顆': 1, '片': 1, '包': 1, '罐': 1, '台': 1, '條': 1, '滴': 1, '碗': 1, '杯': 1, '份': 1, '組': 1, '盒': 1, '袋': 1, '瓶': 1, '卷': 1, '捲': 1 };

function autoDraftNow()    { return adRun_(false); }
function autoDraftDryRun() { return adRun_(true); }

/* ---------- 品類別推測（可調；猜錯的話在主檔改，之後不會被蓋） ---------- */
function adGuessCat_(name, unit) {
  var n = String(name || ''), u = String(unit || '').trim();
  if (/壓模|矽膠模|模\b|模具|模組|圈模|派盤|烤模|蛋糕模|慕斯模|餅乾模|模型/.test(n) && !/材料包|組合包/.test(n)) return '模具';
  if (/筆|匙$|刮刀|抹刀|攪拌|打蛋|篩|秤|溫度計|噴槍|噴水壺|剪刀|擀|量杯|夾|刷$|鏟|針|棒$|轉台|花嘴|裁|模擬|工具|機$|器$/.test(n) && !/材料包|組合包|色膏|糖珠/.test(n)) return '器具';
  if (/盒|袋|杯$|杯蓋|蓋$|紙$|烤盤紙|襯|緞帶|包裝|材料包|組合包|插牌|貼紙|繩|圍邊|底盤|托盤|吸管|叉|湯匙|餐具|手套|標籤|禮盒|提袋|封口|保鮮|紙巾|蠟燭|裝飾組/.test(n)) return '耗材';
  if (u === 'g' || u === 'ml' || u === '' || /粉|糖|醬|餡|油|奶|乳|蛋|果|茶|巧克力|粒|珠|片|絲|豆|米|麵|皮|凍|醋|鹽|酒|膏|香|精|汁|泥|漿|丁|碎|仁|花|葉|籽|子$/.test(n)) return '食材';
  return AD_COUNT_UNITS[u] ? '耗材' : '食材';
}
/* ---------- 名稱正規化（只用來判斷「其實是同一個」） ---------- */
function adNorm_(s) {
  return String(s || '').replace(/[\s\u3000]/g, '').replace(/（/g, '(').replace(/）/g, ')').replace(/［/g, '[').replace(/］/g, ']').replace(/／/g, '/').toLowerCase();
}
function adStripCode_(s) { return adNorm_(s).replace(/^[a-z]\d{2}/, ''); }

function adRun_(dry) {
  var lock = LockService.getScriptLock();
  var out = ['skuAutoDraft ' + (dry ? '【試跑，不寫入】' : '') + new Date()];
  try {
    lock.waitLock(20000);
    var ss = SpreadsheetApp.openById(AD_SHEET_ID);
    var bomSh = ss.getSheetByName(AD_BOM_TAB), skuSh = ss.getSheetByName(AD_SKU_TAB);
    if (!bomSh || !skuSh) throw new Error('找不到分頁');
    var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

    // ---- dim_sku ----
    var sv = skuSh.getDataRange().getValues();
    var h = sv[0].map(function (x) { return String(x || '').trim(); });
    function col(name) { return h.indexOf(name); }
    var cId = col('sku_id'), cName = col('品名'), cCat = col('品類別'), cVendor = col('廠商'), cUse = col('使用單位'), cBuy = col('採購單位'), cPq = col('每採購單位內容量'), cPrice = col('單價'), cAlias = col('BOM別名'), cSrc = col(AD_SRC_COL);
    if ([cId, cName, cCat, cUse, cAlias].some(function (x) { return x < 0; })) throw new Error('dim_sku 缺必要欄（sku_id/品名/品類別/使用單位/BOM別名）');
    if (cSrc < 0) { cSrc = h.length; if (!dry) skuSh.getRange(1, cSrc + 1).setValue(AD_SRC_COL); out.push('首次執行：新增欄「' + AD_SRC_COL + '」於第 ' + (cSrc + 1) + ' 欄'); }
    var exact = {}, normIdx = {}, stripIdx = {}, maxNum = { F: 0, C: 0, T: 0 };
    var rows = [];
    for (var i = 1; i < sv.length; i++) {
      var r = sv[i], id = String(r[cId] || '').trim(), nm = String(r[cName] || '').trim();
      if (!id || !nm) continue;
      var rec = { row: i + 1, id: id, name: nm, aliases: String(r[cAlias] || '').split(/[;,，、]/).map(function (a) { return a.trim(); }).filter(Boolean), src: cSrc < r.length ? String(r[cSrc] || '') : '' };
      rows.push(rec);
      exact[nm] = rec; rec.aliases.forEach(function (a) { exact[a] = rec; });
      var nn = adNorm_(nm); if (!normIdx[nn]) normIdx[nn] = rec;
      var sn = adStripCode_(nm); if (sn && sn !== nn && !stripIdx[sn]) stripIdx[sn] = rec;
      var m = id.match(/^([FCT])(\d+)$/); if (m && +m[2] > maxNum[m[1]]) maxNum[m[1]] = +m[2];
    }

    // ---- BOM 掃描：食材名 → 單位（取最常見） ----
    var bv = bomSh.getDataRange().getValues();
    var bh = bv[0].map(function (x) { return String(x || '').trim(); });
    var bName = bh.indexOf('食材/器具名稱'), bUnit = bh.indexOf('單位');
    if (bName < 0) bName = 1; if (bUnit < 0) bUnit = 3;
    var bDessCol = bh.indexOf('甜點名稱'); if (bDessCol < 0) bDessCol = 0;   // v1.2
    var bom = {};
    for (var j = 1; j < bv.length; j++) {
      var mn = String(bv[j][bName] || '').trim(); if (!mn) continue;
      var mu = String(bv[j][bUnit] || '').trim();
      var e = bom[mn] = bom[mn] || { units: {}, n: 0, dess: {} }; e.n++; e.units[mu] = (e.units[mu] || 0) + 1;
      var dn0 = String(bv[j][bDessCol] || '').trim(); if (dn0) e.dess[dn0] = 1;   // v1.2：誰用到這個料
    }

    // ---- v1.2：甜點上架狀態（null＝對照表讀不到 ⇒ 本次不建任何草稿） ----
    var act = adDessertStatus_(ss, out);

    // ---- 分流 ----
    var addAlias = [], drafts = [], skipped = [], held = [];   // v1.2 held＝甜點還沒上架、暫緩建草稿
    Object.keys(bom).forEach(function (mn) {
      if (exact[mn]) return;
      var nn = adNorm_(mn), hit = normIdx[nn] || stripIdx[nn] || normIdx[adStripCode_(mn)] || null;
      if (hit) { addAlias.push({ rec: hit, alias: mn }); return; }
      var units = bom[mn].units, best = '', bc = -1; Object.keys(units).forEach(function (u) { if (units[u] > bc) { bc = units[u]; best = u; } });
      if (AD_SKIP_RE.test(mn)) { skipped.push(mn); return; }
      var dl = Object.keys(bom[mn].dess);   /* v1.2：至少一款用到它的甜點已上架或排定上架，才建草稿 */
      if (!act) { held.push({ name: mn, why: '對照表讀不到，無法確認甜點是否上架' }); return; }
      if (!dl.some(function (dn) { return act.isLive(dn); })) { held.push({ name: mn, why: act.why(dl) }); return; }
      var unit = best || 'g', cat = adGuessCat_(mn, unit);
      drafts.push({ name: mn, unit: unit, cat: cat, n: bom[mn].n });
    });

    skipped.forEach(function (mn) { out.push('  ⚠ 不建草稿（名稱含下架/測試）：「' + mn + '」→ 請從 BOM 表刪列'); });
    if (held.length) {   /* v1.2 */
      out.push('  ⏸ 甜點還沒上架，先不建草稿 ' + held.length + ' 項（對照表建好該甜點後，隔天早上自動建）：');
      held.forEach(function (x) { out.push('    「' + x.name + '」← ' + x.why); });
    }
    if (act) {   /* v1.2：既有草稿，用到它的甜點都還沒上架 ⇒ 只提醒，不刪 */
      rows.forEach(function (rec) {
        if (!rec.src || rec.src.indexOf('自動草稿') !== 0) return;
        var ds = {};
        [rec.name].concat(rec.aliases).forEach(function (nm0) { var b0 = bom[nm0]; if (b0) Object.keys(b0.dess).forEach(function (d0) { ds[d0] = 1; }); });
        var dl2 = Object.keys(ds);
        if (dl2.length && !dl2.some(function (dn) { return act.isLive(dn); }))
          out.push('  ⚠ 既有草稿的甜點都還沒上架：' + rec.id + ' ' + rec.name + '（' + act.why(dl2) + '）→ 總部決定保留或刪除；要刪請先刪 BOM 表該甜點的列，否則會再被建回來');
      });
    }

    // ---- A) 加別名 ----
    var aliasWrites = {};
    addAlias.forEach(function (a) {
      var rec = a.rec; if (rec.aliases.indexOf(a.alias) >= 0) return;
      rec.aliases.push(a.alias); aliasWrites[rec.row] = rec.aliases.join(',');
      out.push('  🔗 別名：「' + a.alias + '」→ ' + rec.name + '（' + rec.id + '）');
    });
    if (!dry) Object.keys(aliasWrites).forEach(function (row) { skuSh.getRange(+row, cAlias + 1).setValue(aliasWrites[row]); });

    // ---- B) 新增草稿 ----
    drafts.sort(function (a, b) { return b.n - a.n; });
    var newRows = [];
    drafts.forEach(function (d) {
      var pf = AD_PREFIX[d.cat] || 'F'; maxNum[pf]++; var id = pf + ('000' + maxNum[pf]).slice(-3);
      var row = []; for (var k = 0; k <= cSrc; k++) row.push('');
      row[cId] = id; row[cName] = d.name; row[cCat] = d.cat; row[cUse] = d.unit;
      if (cBuy >= 0) row[cBuy] = d.unit; if (cPq >= 0) row[cPq] = 1; if (cPrice >= 0) row[cPrice] = 0;
      row[cSrc] = '自動草稿 ' + today + '（類別推測）';
      newRows.push(row);
      out.push('  🆕 草稿：' + id + '｜' + d.name + '｜' + d.cat + '｜' + d.unit + '｜BOM ' + d.n + ' 列');
    });
    if (!dry && newRows.length) skuSh.getRange(skuSh.getLastRow() + 1, 1, newRows.length, cSrc + 1).setValues(newRows);

    // ---- C) 上架完成 → 清標記 ----
    var cleared = [];
    rows.forEach(function (rec) {
      if (!rec.src || rec.src.indexOf('自動草稿') !== 0) return;
      var r = sv[rec.row - 1];
      var vendorOk = cVendor >= 0 && String(r[cVendor] || '').trim() !== '';
      var pq = cPq >= 0 ? Number(r[cPq]) || 0 : 0, buy = cBuy >= 0 ? String(r[cBuy] || '').trim() : '', use = String(r[cUse] || '').trim();
      var packOk = pq > 1 || (buy && buy !== use);
      var priceOk = cPrice >= 0 && (Number(r[cPrice]) || 0) > 0;
      if (vendorOk && packOk && priceOk) { cleared.push(rec.name); if (!dry) skuSh.getRange(rec.row, cSrc + 1).setValue(''); }
    });

    // ---- D) 最後有效日（v1.1）----
    var lvStat = adLastValid_(ss, skuSh, sv, h, bv, bh, bName, exact, dry, today, out);

    out.push('別名 ' + Object.keys(aliasWrites).length + ' 筆、新增草稿 ' + newRows.length + ' 筆、甜點未上架暫緩 ' + held.length + ' 筆、上架完成清標記 ' + cleared.length + ' 筆' + (cleared.length ? '（' + cleared.join('、') + '）' : '') + '、最後有效日 ' + lvStat.written + ' 筆有日期／' + lvStat.expired + ' 筆已過季');
    var heldChanged = false;   /* v1.2：暫緩清單有變動才記 log */
    if (!dry) {
      try {
        var props = PropertiesService.getScriptProperties(), sig = held.map(function (x) { return x.name; }).sort().join('|');
        if ((props.getProperty(AD_HELD_PROP) || '') !== sig) { heldChanged = true; props.setProperty(AD_HELD_PROP, sig); }
      } catch (eP) { heldChanged = held.length > 0; }
    }
    if (!dry && (newRows.length || Object.keys(aliasWrites).length || cleared.length || heldChanged)) {
      var lg = ss.getSheetByName(AD_LOG_TAB);
      if (!lg) { lg = ss.insertSheet(AD_LOG_TAB); lg.getRange(1, 1, 1, 4).setValues([['執行時間', '別名', '草稿', '明細']]); }
      lg.appendRow([new Date(), Object.keys(aliasWrites).length, newRows.length, out.slice(1).join(' / ')]);
    }
  } catch (err) { out.push('❌ ' + err); }
  finally { try { lock.releaseLock(); } catch (e2) {} }
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/* ---------- v1.2：甜點上架狀態 ----------
 * 對照表每一列：甜點名＝手動對應產品名稱、POS 資料產品名稱（兩個名字都登記），結束有效日。
 * 某甜點任一列「沒有結束日」或「結束日 ≥ 今天 − AD_ENDED_GRACE_DAYS」⇒ 在賣／排定上架（起始日在未來也算排定上架）。
 * 名字比對：先原字串，再 adNorm_（空白、全半形括號）。對照表讀不到或缺欄 ⇒ 回 null（呼叫端一律不建草稿）。 */
function adDessertStatus_(ss, out) {
  var sh = ss.getSheetByName(AD_MAP_TAB);
  if (!sh) { out.push('  ⚠ 找不到「' + AD_MAP_TAB + '」：無法確認甜點是否上架，本次不建任何草稿（別名／清標記／最後有效日照常）'); return null; }
  var v = sh.getDataRange().getValues();
  if (!v.length) { out.push('  ⚠ 「' + AD_MAP_TAB + '」是空的：本次不建任何草稿'); return null; }
  var hh = v[0].map(function (x) { return String(x || '').replace(/\s+/g, ''); });
  var cNm = hh.indexOf('手動對應產品名稱'), cPos = hh.indexOf('POS資料產品名稱'), cEnd = hh.indexOf('結束有效日');
  if (cEnd < 0 || (cNm < 0 && cPos < 0)) { out.push('  ⚠ 對照表缺「結束有效日／產品名稱」欄：本次不建任何草稿'); return null; }
  var tz = 'Asia/Taipei', grace = Utilities.formatDate(new Date(new Date().getTime() - AD_ENDED_GRACE_DAYS * 864e5), tz, 'yyyy-MM-dd');
  var st = {};
  function mark(name, live, endY) {
    name = String(name || '').trim(); if (!name) return;
    [name, adNorm_(name)].forEach(function (k) {
      var o = st[k] = st[k] || { live: false, lastEnd: '' };
      if (live) o.live = true;
      if (endY && endY > o.lastEnd) o.lastEnd = endY;
    });
  }
  for (var i = 1; i < v.length; i++) {
    var e = v[i][cEnd], d = (e instanceof Date) ? e : (e ? new Date(String(e).replace(/\//g, '-')) : null);
    var endY = (d && !isNaN(d)) ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : '';
    var live = !endY || endY >= grace;
    if (cNm >= 0) mark(v[i][cNm], live, endY);
    if (cPos >= 0) mark(v[i][cPos], live, endY);
  }
  function get(dn) { dn = String(dn || '').trim(); return st[dn] || st[adNorm_(dn)] || null; }
  return {
    isLive: function (dn) { var o = get(dn); return !!(o && o.live); },
    why: function (dl) {
      if (!dl.length) return 'BOM 表沒有甜點名稱';
      return dl.slice(0, 3).map(function (dn) { var o = get(dn); return dn + (o ? '（對照表最後結束 ' + o.lastEnd + '，已超過 ' + AD_ENDED_GRACE_DAYS + ' 天）' : '（對照表沒有）'); }).join('、') + (dl.length > 3 ? ' 等 ' + dl.length + ' 款' : '');
    }
  };
}

/* ---------- D) 最後有效日（v1.1）----------
 * 對照表：甜點名 ＝ 手動對應產品名稱（沒填則 POS 資料產品名稱），日期 ＝ 結束有效日。
 * 同一甜點多列 → 取最晚；任一列結束日空白 → 該甜點長期有效。
 * BOM 甜點名 → 對照表甜點名 直接相等比對（食譜名）。BOM 有、對照表沒有的甜點 → 長期有效（保守，寧可多盤不漏訂）。 */
function adLastValid_(ss, skuSh, sv, h, bv, bh, bName, exact, dry, today, out) {
  var res = { written: 0, expired: 0 };
  var mapSh = ss.getSheetByName(AD_MAP_TAB);
  if (!mapSh) { out.push('  ⚠ 找不到「' + AD_MAP_TAB + '」，最後有效日略過'); return res; }
  var mv = mapSh.getDataRange().getValues();
  var mh = mv[0].map(function (x) { return String(x || '').replace(/\s+/g, ''); });   // 表頭去空白（地雷 20）
  var cNm = mh.indexOf('手動對應產品名稱'), cPos = mh.indexOf('POS資料產品名稱'), cEnd = mh.indexOf('結束有效日');
  if (cEnd < 0 || (cNm < 0 && cPos < 0)) { out.push('  ⚠ 對照表缺 結束有效日／產品名稱 欄，最後有效日略過'); return res; }
  var dEnd = {};                       // 甜點名 -> Date（最晚）或 null（長期有效）
  for (var i = 1; i < mv.length; i++) {
    var nm = String((cNm >= 0 && mv[i][cNm]) || (cPos >= 0 && mv[i][cPos]) || '').trim(); if (!nm) continue;
    var e = mv[i][cEnd], d = (e instanceof Date) ? e : (e ? new Date(String(e).replace(/\//g, '-')) : null);
    if (!d || isNaN(d)) { dEnd[nm] = null; continue; }        // 任一列無結束日 → 長期有效
    if (dEnd[nm] === null) continue;
    if (!dEnd[nm] || d > dEnd[nm]) dEnd[nm] = d;
  }
  // BOM：食材 → 用到它的甜點
  var bDess = bh.indexOf('甜點名稱'); if (bDess < 0) bDess = 0;
  var byItem = {};                     // sku_id -> {max:Date|null, any:bool}
  for (var j = 1; j < bv.length; j++) {
    var mn = String(bv[j][bName] || '').trim(), dn = String(bv[j][bDess] || '').trim(); if (!mn || !dn) continue;
    var rec = exact[mn] || exact[adNorm_(mn)]; if (!rec) continue;
    var o = byItem[rec.id] = byItem[rec.id] || { max: null, any: false, open: false };
    o.any = true;
    if (!(dn in dEnd) || dEnd[dn] === null) { o.open = true; continue; }   // 對照表沒列或無結束日 → 長期
    if (!o.max || dEnd[dn] > o.max) o.max = dEnd[dn];
  }
  var cLv = h.indexOf(AD_LV_COL);
  if (cLv < 0) {
    cLv = h.length;
    if (!dry) {
      if (skuSh.getMaxColumns() < cLv + 1) skuSh.insertColumnsAfter(skuSh.getMaxColumns(), cLv + 1 - skuSh.getMaxColumns());
      skuSh.getRange(1, cLv + 1).setValue(AD_LV_COL);
    }
    out.push('首次執行：新增欄「' + AD_LV_COL + '」於第 ' + (cLv + 1) + ' 欄');
  }
  var cId = h.indexOf('sku_id'), cName = h.indexOf('品名'), rows = sv.length - 1, col = [];
  var tz = 'Asia/Taipei', used = adAggUsed_(ss), conflict = [], sample = [];
  for (var k = 1; k < sv.length; k++) {
    var id = String(sv[k][cId] || '').trim(), o2 = byItem[id], v = '';
    if (id && o2 && o2.any && !o2.open && o2.max) {
      v = Utilities.formatDate(o2.max, tz, 'yyyy-MM-dd'); res.written++;
      if (v < today) { res.expired++; var nm2 = String(sv[k][cName] || id); if (sample.length < 15) sample.push(nm2 + ' ' + v); if (used[id] > 0) conflict.push(nm2 + '（最後有效 ' + v + '，近4週均 ' + Math.round(used[id]) + '）'); }
    }
    col.push([v]);
  }
  if (sample.length) out.push('  🗓 過季範例：' + sample.join('；'));
  if (conflict.length) { out.push('  ⚠ 過季但近 4 週仍有用量 ' + conflict.length + ' 筆（對照表結束有效日可能填錯，或該甜點續賣）：'); conflict.slice(0, 40).forEach(function (x) { out.push('    ' + x); }); }
  res.conflict = conflict.length;
  if (!dry && rows > 0) {
    var rg = skuSh.getRange(2, cLv + 1, rows, 1);
    rg.setNumberFormat('@');           // 鎖文字，避免 Sheets 轉日期型別（地雷 3.2）
    rg.setValues(col);
  }
  return res;
}

/* agg_purchase：sku_id → 全店最大「近4週均」（只用來偵測「過季卻仍在用」的矛盾） */
function adAggUsed_(ss) {
  var used = {}, sh = ss.getSheetByName(AD_AGG_TAB); if (!sh) return used;
  var v = sh.getDataRange().getValues(); if (v.length < 2) return used;
  var hh = v[0].map(function (x) { return String(x || '').trim(); }), cI = hh.indexOf('sku_id'), cA = hh.indexOf('近4週均');
  if (cI < 0 || cA < 0) return used;
  for (var i = 1; i < v.length; i++) { var id = String(v[i][cI] || '').trim(), a = Number(v[i][cA]) || 0; if (id && a > (used[id] || 0)) used[id] = a; }
  return used;
}