/*** monthAgg.gs v1.3 — 月度預聚合（第 6 支專案檔）
 *   v1.3（2026-07-30）欄寬邊界修正【重要】：
 *     舊版用 getDataRange() 讀整張表 → 使用者放在 E/F 欄的驗算公式會被當成「第 5、6 個表頭」讀進來，
 *     再用 clearContents() 清掉整張表後寫回，公式就被凍結成死文字（實際發生於 2026-07-30 的 F1）。
 *     修正：本表欄寬固定為 defHeaders.length，右側任何內容一律【不讀、不寫、不清】。
 *          使用者可安全地在 E 欄以後放驗算公式，不會被程式破壞。
 *   v1.2（2026-07-30）月份欄型別修正【重要】：
 *     Google 試算表會把 setValues 寫入的 "2025-07" 自動判讀成「日期」。下次讀回來是 Date 物件，
 *     與文字 "2025-07" 比對永遠不相等 → 目標月份刪不掉 → 重算變成「疊加」而非「取代」，資料重複灌水。
 *     修正 (a) 讀取端 _mNormYm()：Date / "YYYY-MM" / "YYYY-MM-DD" 一律正規化成 "YYYY-MM" 再比對
 *          (b) 寫入端：寫值前先把「月」欄整欄格式設為純文字(@)，杜絕再被轉型
 *          (c) 保留下來的歷史列也一併正規化寫回，讓全欄型別一致（前端 gviz 以字串比較月份）
 *     ※ 修好後重跑同一個月，會把重複的舊列與新列一起清掉再寫一份乾淨的（自我修復）。
 *   v1.1（2026-07-30）安全性修正，回溯歷史月必須用此版：
 *     (a) 目標月「計算結果 0 列」時 → 保留原有資料不改寫（原 v1.0 會先刪後寫 → 該月資料永久消失）
 *     (b) POS 命中 0 列 → 直接中止，一個字都不寫（防對照表/分頁異常造成整批清空）
 *   背景：月檔 agg_usage_month / agg_sales_month 的寫入程式於 2026-07-21 隨 Code.gs 整檔替換而遺失，
 *         最後一次寫入 = 2026-07-21 07:07，之後 9 天停更。本檔為重建版，獨立成檔，不動 Code.gs。
 *
 *   讀：BOM 本(SS_ID) 的 POS資料 / BOM表 / 產品名稱對照表 / dim_sku / dim_unitconv
 *   寫：外部月檔(M_SS_ID) 的 agg_usage_month / agg_sales_month
 *
 *   ⚠️ 本檔重用 Code.gs 的全域：SS_ID / POS_TAB / BOM_TAB / MAP_TAB / readTab_ / cv_ / _num / _toDate / _conv / LOSS
 *      故一律以 M_ / _m 前綴命名自有變數與函式，避免覆蓋 Code.gs 全域（Apps Script 各檔共用同一全域範圍）。
 *      _conv 第 4 參數 altItem 為 v6 新增；若部署的是舊版 Code.gs，多傳的參數會被忽略，不會出錯。
 *
 *   設計：**增量 upsert**。每日只重算「本月＋上月」，其餘月份的歷史列原封保留。
 *        整表重建請改用 rebuildMonthAggRange('2026-01','2026-07')（一次性補洞用）。
 *
 *   部署：貼上後 → 存檔 → 執行一次 rebuildMonthAgg（授權）→ 執行 setupMonthTrigger（建每日 07:40 觸發器）
 ***/

var M_SS_ID        = '19tG-AMYtiZPTYQcz85TGiGS2G4M5U1gZw9Mab2UFFU0'; // 外部月檔（BOM 本已近儲存格上限，故月表外置）
var M_TAB_USAGE    = 'agg_usage_month';
var M_TAB_SALES    = 'agg_sales_month';
var M_HEAD_USAGE   = ['店號','sku_id','月','用量'];
var M_HEAD_SALES   = ['店號','商品','月','數量'];
var M_MONTHS_BACK  = 1;     // 每日重算範圍：本月 + 往前 N 個月（1 = 本月與上月）
var M_SALES_CANON  = true;  // true=銷售表用「手動對應產品名稱」合併別名（葷素後綴／員工命名／版本號）；false=用 POS 原始品名
var M_SALES_DESSERT_ONLY = true; // true=只計有 BOM 配方的品項（自動排除 入場費／加價購／訂金／券 等非甜點）
var M_TRIGGER_HOUR = 7;
var M_TRIGGER_MIN  = 40;    // 避開 rebuildPurchaseAgg(07:05) 與 rebuildEventLayer(07:31)

/* ========== 純函式核心（可離線測試，不碰 SpreadsheetApp） ========== */

function _mYm(d){ return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2); }

/** 月份值正規化：Date / 'YYYY-MM' / 'YYYY-MM-DD' → 'YYYY-MM' */
function _mNormYm(v){
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : _mYm(v);
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (/^\d{4}-\d{1,2}$/.test(s)) { var a = s.split('-'); return a[0] + '-' + ('0' + a[1]).slice(-2); }
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 7);
  if (/^\d{4}\/\d{1,2}/.test(s)) { var b = s.split('/'); return b[0] + '-' + ('0' + b[1]).slice(-2); }
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : _mYm(d);
}

/** 產生「本月往前 n 個月」的 YYYY-MM 清單（含本月） */
function _mRecentMonths(n, today){
  var d = today || new Date(), out = [];
  for (var i = 0; i <= (n||0); i++){
    var x = new Date(d.getFullYear(), d.getMonth() - i, 1);
    out.push(_mYm(x));
  }
  return out;
}

/** 產生 from~to 的 YYYY-MM 清單（含頭尾） */
function _mMonthRange(from, to){
  var a = String(from).slice(0,7), b = String(to).slice(0,7), out = [];
  if (a > b) { var t = a; a = b; b = t; }
  var y = +a.slice(0,4), m = +a.slice(5,7), guard = 0;
  while (guard++ < 600){
    var ym = y + '-' + ('0'+m).slice(-2);
    out.push(ym);
    if (ym >= b) break;
    m++; if (m > 12){ m = 1; y++; }
  }
  return out;
}

/**
 * 月度核心計算
 * @param months 只計算這些 YYYY-MM（其餘 POS 列直接跳過，省時）
 * @return {usageRows:[{store,key,ym,qty}], salesRows:[...], orphans:{}, skippedSales:n, posHit:n}
 */
function _mCore(pos, bom, map, sku, unitconv, months){
  var want = {}; (months||[]).forEach(function(m){ want[String(m)] = 1; });

  var posToCanon = {};
  map.forEach(function(m){ if (m && m.pos) posToCanon[m.pos] = (m.canon || m.pos); });

  var bomBy = {};
  bom.forEach(function(b){ if (!b || !b.dessert || !b.item) return; (bomBy[b.dessert] = bomBy[b.dessert] || []).push(b); });

  var skuByName = {};
  sku.forEach(function(s){ if (!s || !s.name) return; skuByName[s.name] = s;
    (s.aliases || []).forEach(function(a){ if (a) skuByName[a] = s; }); });

  var usage = {}, sales = {}, orphans = {}, skippedSales = 0, posHit = 0;

  pos.forEach(function(p){
    if (!p) return;
    var d = _toDate(p.date); if (!d) return;
    var ym = _mYm(d); if (!want[ym]) return;
    var qty = _num(p.qty); if (!qty) return;
    var store = String(p.store == null ? '' : p.store).trim(); if (!store) return;

    var canon  = posToCanon[p.product] || p.product;
    var recipe = bomBy[canon];

    // ---- 銷售表 ----
    if (!M_SALES_DESSERT_ONLY || recipe){
      var nm = M_SALES_CANON ? canon : p.product;
      nm = String(nm == null ? '' : nm).trim();
      if (nm){
        var ks = store + '\u0001' + nm + '\u0001' + ym;
        sales[ks] = (sales[ks] || 0) + qty;
      }
    } else {
      skippedSales += qty;
    }

    // ---- 用量表 ----
    if (!recipe) return;
    posHit++;
    recipe.forEach(function(line){
      var s = skuByName[line.item];
      if (!s){ orphans[line.item] = (orphans[line.item] || 0) + 1; return; }
      if (s.cat === '器具' || s.cat === '模具') return;          // 器具/模具無耗用（走標配路徑）
      var base     = _num(line.qty) * _conv(unitconv, line.item, line.unit, s.name); // 查找序：BOM原名 → 主檔品名 → (通用) → 1
      var consumed = base * qty * (LOSS[s.cat] || 1);
      var ku = store + '\u0001' + s.sku_id + '\u0001' + ym;
      usage[ku] = (usage[ku] || 0) + consumed;
    });
  });

  function toRows(obj, round){
    return Object.keys(obj).map(function(k){
      var a = k.split('\u0001');
      return { store: a[0], key: a[1], ym: a[2], qty: round ? Math.round(obj[k]) : obj[k] };
    }).filter(function(r){ return r.qty > 0; })
      .sort(function(x, y){
        if (x.ym !== y.ym) return x.ym < y.ym ? -1 : 1;
        var nx = +x.store, ny = +y.store;
        if (!isNaN(nx) && !isNaN(ny) && nx !== ny) return nx - ny;
        if (x.store !== y.store) return x.store < y.store ? -1 : 1;
        return x.key < y.key ? -1 : (x.key > y.key ? 1 : 0);
      });
  }

  return {
    usageRows: toRows(usage, true),
    salesRows: toRows(sales, false),
    orphans: orphans,
    skippedSales: skippedSales,
    posHit: posHit
  };
}

/** 把 {store,key,ym,qty} 依「實際表頭順序」落位（表頭被人動過也不會錯位） */
function _mRowToArr(headers, obj, keyHeader, qtyHeader){
  var arr = [], i;
  for (i = 0; i < headers.length; i++) arr.push('');
  var used = {};
  function put(name, val, fallbackIdx){
    var idx = headers.indexOf(name);
    if (idx < 0 && fallbackIdx != null && fallbackIdx < headers.length && !used[fallbackIdx]) idx = fallbackIdx;
    if (idx < 0) return;
    arr[idx] = val; used[idx] = 1;
  }
  put('店號',      obj.store, 0);
  put(keyHeader,   obj.key,   1);
  put('月',        obj.ym,    2);
  put(qtyHeader,   obj.qty,   3);
  return arr;
}

/* ========== 寫入（upsert：只換目標月份，其餘歷史列保留） ========== */

function _mUpsert(ss, tabName, defHeaders, keyHeader, qtyHeader, months, rows){
  var sh = ss.getSheetByName(tabName);
  if (!sh){ sh = ss.insertSheet(tabName); sh.getRange(1,1,1,defHeaders.length).setValues([defHeaders]); }

  // v1.3：只處理固定欄寬 W，右側（使用者的驗算公式等）完全不碰
  var W = defHeaders.length;
  var lastRow = sh.getLastRow();
  var v = lastRow ? sh.getRange(1, 1, lastRow, W).getValues() : [];
  var headers = (v.length && String(v[0][0]||'').trim()) ? v.shift().map(function(h){ return String(h).trim(); }) : defHeaders.slice();
  if (headers.length !== W) headers = defHeaders.slice();

  var mi = headers.indexOf('月');
  if (mi < 0){ mi = 2; Logger.log('⚠️ ' + tabName + ' 找不到「月」欄，改用第 3 欄；請確認表頭：' + JSON.stringify(headers)); }

  // v1.1：只清「本次真的算出資料」的月份。算出 0 列的月份原資料保留（POS 可能已無該月紀錄）
  var got = {}; rows.forEach(function(o){ got[o.ym] = 1; });
  var kill = {}, skipped = [];
  months.forEach(function(m){ var mm = _mNormYm(m); if (got[mm]) kill[mm] = 1; else skipped.push(mm); });
  if (skipped.length) Logger.log('⚠️ ' + tabName + '：月份 ' + skipped.join('、') + ' 計算結果 0 列 → 已保留原有資料不改寫（請確認 POS資料 是否還留有該月）');

  // v1.2：以正規化後的月份比對（相容 Date 型別），保留列的月份也一併正規化成文字寫回
  var kept = [];
  v.forEach(function(r){
    var blank = r.every(function(c){ return c === '' || c === null; });
    if (blank) return;
    var ym = _mNormYm(r[mi]);
    if (kill[ym]) return;
    r[mi] = ym;
    kept.push(r);
  });

  var fresh = rows.map(function(o){ return _mRowToArr(headers, o, keyHeader, qtyHeader); });
  var all   = kept.concat(fresh);

  // v1.3：只清 A..W 欄，不用 clearContents()（那會連使用者放在右側的公式一起清掉）
  sh.getRange(1, 1, sh.getMaxRows(), W).clearContent();
  sh.getRange(1, 1, 1, W).setValues([headers]);
  if (all.length){
    var norm = all.map(function(r){
      var out = r.slice(0, W);
      while (out.length < W) out.push('');
      return out;
    });
    // v1.2：先把「月」欄鎖成純文字，避免 "2025-07" 被試算表判讀成日期（會導致下次刪不掉舊列）
    sh.getRange(2, mi + 1, norm.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, norm.length, W).setValues(norm);
  }
  return { kept: kept.length, written: fresh.length, total: all.length, skipped: skipped };
}

/* ========== 主流程 ========== */

function _mRun(months){
  var t0 = new Date().getTime();
  var ss = SpreadsheetApp.openById(SS_ID);
  var P = readTab_(ss, POS_TAB), B = readTab_(ss, BOM_TAB), M = readTab_(ss, MAP_TAB),
      S = readTab_(ss, 'dim_sku'), U = readTab_(ss, 'dim_unitconv');

  var pos = P.rows.map(function(r){ return { store: cv_(P,r,'分店代碼'), date: cv_(P,r,'建立日期'), product: cv_(P,r,'商品名稱'), qty: cv_(P,r,'數量') }; });
  var bom = B.rows.map(function(r){ return { dessert: cv_(B,r,'甜點名稱'), item: cv_(B,r,'食材/器具名稱'), qty: cv_(B,r,'數量'), unit: cv_(B,r,'單位') }; });
  var map = M.rows.map(function(r){ return { pos: cv_(M,r,'POS資料產品名稱'), canon: cv_(M,r,'手動對應產品名稱') }; });
  var sku = S.rows.map(function(r){ return { sku_id: cv_(S,r,'sku_id'), name: cv_(S,r,'品名'), cat: cv_(S,r,'品類別'),
    aliases: String(cv_(S,r,'BOM別名')||'').split(/[;,，、]/).map(function(x){ return x.trim(); }).filter(Boolean) }; });
  var unitconv = U.rows.map(function(r){ return { item: cv_(U,r,'適用品項'), unit: cv_(U,r,'單位'), base: cv_(U,r,'基準單位'), amt: cv_(U,r,'換算量') }; });

  var out = _mCore(pos, bom, map, sku, unitconv, months);

  // v1.1：整體防呆——POS 命中 0 列代表讀表或對照表異常，此時任何寫入都可能清空資料
  if (!out.posHit){
    Logger.log('🛑 已中止，未寫入任何資料：POS 命中 0 列。請檢查 ①POS資料 是否有 ' + JSON.stringify(months)
      + ' 的資料 ②產品名稱對照表/BOM表 是否正常。');
    return { months: months, aborted: true, posHit: 0 };
  }

  var mss = SpreadsheetApp.openById(M_SS_ID);
  var ru = _mUpsert(mss, M_TAB_USAGE, M_HEAD_USAGE, 'sku_id', '用量', months, out.usageRows);
  var rs = _mUpsert(mss, M_TAB_SALES, M_HEAD_SALES, '商品',  '數量', months, out.salesRows);

  var secs = Math.round((new Date().getTime() - t0) / 1000);
  Logger.log('rebuildMonthAgg 完成（' + secs + ' 秒）。重算月份=' + JSON.stringify(months)
    + ' ｜ ' + M_TAB_USAGE + ' 新寫=' + ru.written + ' 保留歷史=' + ru.kept + ' 合計=' + ru.total
    + ' ｜ ' + M_TAB_SALES + ' 新寫=' + rs.written + ' 保留歷史=' + rs.kept + ' 合計=' + rs.total
    + ' ｜ POS命中列=' + out.posHit + ' 非甜點份數(未計入銷售)=' + out.skippedSales);
  if (Object.keys(out.orphans).length)
    Logger.log('孤兒(BOM有但dim_sku無對應)=' + Object.keys(out.orphans).length + ' ' + JSON.stringify(out.orphans));
  return { months: months, usage: ru, sales: rs, posHit: out.posHit, skipped: ru.skipped };
}

/** 每日觸發用：重算本月＋上月 */
function rebuildMonthAgg(){
  return _mRun(_mRecentMonths(M_MONTHS_BACK));
}

/** 一次性補洞／回溯用：rebuildMonthAggRange('2026-01','2026-07') */
function rebuildMonthAggRange(from, to){
  if (!from || !to) throw new Error("用法：rebuildMonthAggRange('2026-01','2026-07')");
  var months = _mMonthRange(from, to);
  if (months.length > 18) throw new Error('一次最多 18 個月（避免 6 分鐘逾時），請分批。目前要求 ' + months.length + ' 個月');
  return _mRun(months);
}

/** 建立每日 07:40 觸發器（重複執行會先清掉舊的同名觸發器） */
function setupMonthTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'rebuildMonthAgg') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rebuildMonthAgg').timeBased().everyDays(1).atHour(M_TRIGGER_HOUR).nearMinute(M_TRIGGER_MIN).create();
  Logger.log('已建立 rebuildMonthAgg 每日觸發器 ≈' + M_TRIGGER_HOUR + ':' + M_TRIGGER_MIN);
}

/* ===== 一次性回溯（全部跑完後整段可刪） ===== */
function backfill_batch1(){ return rebuildMonthAggRange('2024-12','2025-08'); }
function backfill_batch2(){ return rebuildMonthAggRange('2025-09','2026-05'); }
function backfill_batch3(){ return rebuildMonthAggRange('2026-06','2026-08'); }   // 2026-10-04 經營者核可：6～8 月改用今天的 BOM 重算