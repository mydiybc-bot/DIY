/**
 * diybc-dcorder-api  v1.9.1（2026-11-12 新竹按重量也一箱一張托運單）
 *
 * v1.9.1：新竹司機反映多件合成一個貨號、掃描機只跳 1 件 → hctCreate 按重量也改一箱一張（件數 1、總重量平均分），上限 20 箱。
 * 門市 → 出貨中心 訂單 API（取代 Shopline 門市叫貨）
 *
 * v1.9：新分頁 dc_invsess（一店一列）：盤點頁「延續／新的一週（重新盤點）」原本只存在裝置本機，
 *       手機按了重新盤點、電腦還停在舊輪次（11/12 7 號店：手機 0/206、電腦 196/206）。改存雲端，各裝置以最後一次設定為準。
 *       GET  invsess &store=店號 → {sess:{store,start,setAt}|null}（start：0＝延續歷史；>0＝本輪開始毫秒）
 *       POST invsessSet {store,start,setAt}：setAt 比雲端舊就不蓋（回傳雲端現值，stale:true）
*
 * v1.7.1：merge 擋下「新竹托運單已交寄」的訂單（貨已寄出，只是沒按出貨完成；11/09 9 家店誤合併的原因）。
 *
 * v1.7：POST unmerge {id:<併入的那張>, base?}：把合併拆回原本的幾張（被併入的那幾張回到合併前的狀態與備註；
 *       併入那張扣掉它們的品項數量）。合併後那張已出貨就不能拆。合併時 dc_log 多記每張合併前的狀態與備註（prev），
 *       拆回時優先用它；v1.7 以前的合併改從 dc_log 的建單／改狀態紀錄推回合併前狀態（查不到＝待處理）。
 *
 * v1.6：新分頁 dc_vmark（一列＝一次標記的一個品項）：門市在訂貨單按「✓ 已向某廠商叫貨」或「這週下單完成（剩下的不叫）」時記下來，
 *       平板、電腦看到同一份。出貨中心的單不用標（送出的訂單本身就是紀錄）。
 *       GET  vmarks &store=店號 &days=N（預設 45）→ {marks:[{mid,store,vendor,kind,time,status,undoTime,items:[{sku_id,name,qty,unit}]}]}
 *       POST vmark {store, vendor, kind:'叫貨'|'不叫', mid, items:[{sku_id,name,qty,unit}]}（同一個 mid 再送只回原標記）
 *       POST vunmark {store, mid}：復原（狀態改「已復原」，列保留備查）
 *
 * v1.5：照新竹 2026-10-27 寄來的 v16.0.4 規格改：SOAP 1.2；測試／正式網址分開（HCT_ENV）；多客代（HCT_CS，重量計費 W／總長計費 S，
 *       總長一箱一筆）；託運單號只用數字（訂單號＋序號 2 碼）；確認出貨一批 20；新增 POST hctDispatch／hctDispatchCancel／hctQuery；
 *       dc_hct 最後加 6 欄（客代、計費、商品種類、箱序、派遣、派遣時間，舊表第一次用時自動補標題）。
 *
 * v1.3：新竹物流（HCT）托運單。新分頁 dc_hct（一列＝一張托運單）＋ dc_hct_img（標籤圖 base64，留 60 天）
 *       POST hctCreate {id, boxes, weight, rcv:{name,tel,addr}}：向新竹取貨號＋標籤（同一天再送＝更正件數重量，貨號不變）
 *       POST hctVoid {id}：作廢還沒交寄的托運單（新竹沒有取消 API；不交寄＝不出貨）
 *       POST hctReport {date}：確認當天出貨（新竹規定 18:00 前），回傳當天托運單給前端印「託運總表」
 *       GET  hct &date=yyyyMMdd｜&ids=訂單號,… ｜&days=N ［&img=1 才回標籤圖］
 *       新竹帳號放指令碼屬性 HCT_COMPANY／HCT_PWD（＋可選 HCT_ESCSNO 客代 11 碼、HCT_ESSTNO 出貨站 4 碼、HCT_SENDER／HCT_SENDER_TEL／HCT_SENDER_ADDR）；
 *       沒設＝新竹公開測試帳號 test／test1（標籤有「提供測試用 不得寄件」浮水印）。ping 會回 hct:'test'|'live'。
 *       改正式帳號：專案設定 → 指令碼屬性填好即可，程式不用改。
 *
 * v1.2：①新分頁 dc_loc（出貨中心品項「位置」，出貨中心人員自己填）：GET stock 多回 locs、新 GET locs、POST locSet
 *       ②POST edit：出貨中心編輯整張訂單（門市、品項、數量、單價、實出、備註、新竹貨號）；已出貨的單改實出會補扣／退回記數量庫存
 *       ③POST merge：同一店多張未出貨訂單合併成一張（併入最早那張；其他張改「已取消」、備註寫「已併入 …」，明細保留備查）
 *       編輯／合併都帶 base（畫面上看到的更新時間），別人先改過就擋下（conflict），不會互蓋。
 *
 * v1.1：GET orders／stock 的「整表讀＋組訂單」結果存 CacheService 10 分鐘（>90,000 字元切塊）；
 *       key 帶版本號（指令碼屬性 DCO_CACHE_VER），每次 doPost 寫入結束就換版本號 → 舊快取一次失效。
 *       寫入函式一律直接讀試算表（不經快取）；快取任何環節失敗都退回 v1.0 直接讀表的路。回傳內容不變。
 *       手動清快取：編輯器執行 aaaClearReadCache()（例如在試算表手改資料之後）。
 *
 * 獨立 Apps Script 專案（沿用「每支寫入 API 各自一個專案」慣例，不動 diybc-purchase-agg 主程式）。
 * 資料放在獨立試算表「出貨中心訂單（採購系統）」：第一次執行 setup() 時自動建立，ID 記在指令碼屬性 DCO_SS_ID。
 *   dc_order_line  一列＝一張訂單的一個品項（訂單層欄位每列重複，月結、判讀直接加總）
 *   dc_stock       出貨中心庫存（只記「記數量」的品項；沒列在這裡＝無限量）
 *   dc_log         所有異動紀錄
 *
 * 讀：GET  ?action=ping｜orders｜stock｜locs｜vmarks &token=…（可加 callback= 走 JSONP）
 * 寫：POST text/plain JSON {token, action:create|cancel|status|ship|stockSet|stockAdj|locSet|edit|merge, …}
 *
 * 部署：部署 → 新增部署作業 → 網頁應用程式；執行身分＝我；存取權＝所有人。
 *       之後改程式一律「管理部署作業 → 編輯 → 新版本」，網址不變。
 */
var DCO_TOKEN = 'dbc-dco-Rw8pZ3';
var DCO_VER = 'v1.9.1-hct1box';
var DCO_TZ = 'Asia/Taipei';
var SH_LINE = 'dc_order_line', SH_STOCK = 'dc_stock', SH_LOG = 'dc_log', SH_LOC = 'dc_loc';
var H_LINE = ['訂單號', '店號', '下單時間', '狀態', '行號', 'sku_id', '品名', '單位', '訂購量', '單價', '實出量', '小計', '備註', '出貨時間', '新竹貨號', 'cid', '更新時間'];
var H_STOCK = ['sku_id', '品名', '管理方式', '庫存', '更新時間', '備註'];
var H_LOG = ['時間', '動作', '訂單號', '店號', '內容'];
var H_LOC = ['sku_id', '品名', '位置', '更新時間'];   /* v1.2：出貨中心品項位置（一個品項一列；位置清空＝刪列） */
var SH_HCT = 'dc_hct';   /* v1.3：新竹托運單 */
var H_HCT = ['訂單號', '店號', '託運單號', '新竹貨號', '到著站', '件數', '重量', '出貨日', '狀態', '收件人', '電話', '地址', '環境', '建立時間', '更新時間', '交寄時間', '客代', '計費', '商品種類', '箱序', '派遣', '派遣時間'];   /* v1.5 加最後 6 欄 */
var SH_HCTIMG = 'dc_hct_img';   /* v1.3：標籤圖另放一頁（一張約 15,000 字），平常讀托運單不用讀圖；只留 HCT_IMG_DAYS 天 */
var H_HCTIMG = ['託運單號', '新竹貨號', '標籤', '建立時間'];
var SH_VMARK = 'dc_vmark';   /* v1.6：門市「本輪已叫貨／不叫」標記 */
var H_VMARK = ['標記號', '店號', '廠商', '類型', '標記時間', 'sku_id', '品名', '數量', '單位', '狀態', '復原時間'];
var SH_INVSESS = 'dc_invsess';   /* v1.9：各店盤點輪次（一店一列） */
var H_INVSESS = ['店號', '盤點開始', '設定時間', '設定時間（看）', '狀態'];
var HCT_IMG_DAYS = 60;
var DCO_MERGED_PREFIX = '已併入 ';                    /* v1.2：被合併的訂單備註開頭（前端靠它顯示「已併入」） */
var OPEN_ST = { '待處理': 1, '備貨中': 1 };

/* ── 讀取快取（v1.1） ─────────────────────────────────────────── */
var DCO_CACHE_TTL = 600;            // 10 分鐘
var DCO_CACHE_CHUNK = 90000;        // 單 key 上限 100KB → 每塊 90,000 字元
var DCO_CACHE_MAX_CHUNKS = 50;      // 超過約 4.5MB 就不存（只回算好的結果，功能照常）
var DCO_CACHE_PREFIX = 'DCO_RD_';
var DCO_CACHE_VER_PROP = 'DCO_CACHE_VER';

function dcoCacheVer_() {
  try { return PropertiesService.getScriptProperties().getProperty(DCO_CACHE_VER_PROP) || '0'; } catch (e) { return '0'; }
}
/* 寫入後呼叫：換版本號 → 所有讀取快取立即失效 */
function dcoCacheBump_() {
  try { PropertiesService.getScriptProperties().setProperty(DCO_CACHE_VER_PROP, String(Date.now()) + '-' + Math.floor(Math.random() * 1e9)); } catch (e) { }
}
function dcoCacheKey_(name) { return DCO_CACHE_PREFIX + dcoCacheVer_() + '_' + name; }
/* 讀分塊：<key>_n = 塊數、<key>_0.._n-1 = 內容；任一塊缺就當未命中 */
function dcoCacheGet_(key) {
  var cache = CacheService.getScriptCache();
  var meta = cache.get(key + '_n');
  if (meta === null || meta === undefined) return null;
  var n = parseInt(meta, 10);
  if (!(n >= 1)) return null;
  var keys = [];
  for (var i = 0; i < n; i++) keys.push(key + '_' + i);
  var parts = cache.getAll(keys), out = '';
  for (var j = 0; j < n; j++) {
    var part = parts[key + '_' + j];
    if (part === undefined || part === null) return null;
    out += part;
  }
  return out;
}
function dcoCachePut_(key, str) {
  var n = Math.ceil(str.length / DCO_CACHE_CHUNK);
  if (n > DCO_CACHE_MAX_CHUNKS) return;
  var obj = {};
  for (var i = 0; i < n; i++) obj[key + '_' + i] = str.substring(i * DCO_CACHE_CHUNK, (i + 1) * DCO_CACHE_CHUNK);
  obj[key + '_n'] = String(n);
  CacheService.getScriptCache().putAll(obj, DCO_CACHE_TTL);
}
/* 讀取共用：命中回 JSON.parse 後的物件；未命中算一次並存。回傳與 computeFn() 同結構（全是字串／數字／null，JSON 來回不失真） */
function dcoCached_(name, computeFn) {
  var key = null;
  try {
    key = dcoCacheKey_(name);
    var hit = dcoCacheGet_(key);
    if (hit !== null) return JSON.parse(hit);
  } catch (e) { key = null; }
  var val = computeFn();
  if (key) { try { dcoCachePut_(key, JSON.stringify(val)); } catch (e2) { } }
  return val;
}
/* 全部訂單（已組成訂單物件、依明細行號排好）：快取的是這一份；篩選仍每次依參數做 */
function dcoOrdersAll_() { return dcoCached_('orders', function () { return dcoGroup_(dcoRead_(SH_LINE).rows); }); }
/* v1.2：品項位置（只回有填位置的） */
function dcoLocs_() {
  return dcoCached_('locs', function () {
    return dcoRead_(SH_LOC).rows.filter(function (r) { return dcoStr_(r['sku_id']) && dcoStr_(r['位置']); }).map(function (r) {
      return { sku_id: dcoStr_(r['sku_id']), loc: dcoStr_(r['位置']) };
    });
  });
}
/* 手動清掉所有讀取快取（編輯器直接執行） */
function aaaClearReadCache() { dcoCacheBump_(); }

/* ── 初始化（第一次手動執行一次；會要求授權） ─────────────────── */
function setup() {
  var ss = dcoSS_();
  dcoEnsure_(ss, SH_LINE, H_LINE, ['A', 'B', 'F', 'O', 'P']);
  dcoEnsure_(ss, SH_STOCK, H_STOCK, ['A']);
  dcoEnsure_(ss, SH_LOG, H_LOG, ['C', 'D']);
  dcoEnsure_(ss, SH_LOC, H_LOC, ['A', 'C']);   /* v1.2 */
  dcoEnsure_(ss, SH_HCT, H_HCT, ['A', 'B', 'C', 'D', 'E', 'H', 'K']);   /* v1.3 */
  dcoEnsure_(ss, SH_HCTIMG, H_HCTIMG, ['A', 'B']);                     /* v1.3 */
  dcoEnsure_(ss, SH_VMARK, H_VMARK, ['A', 'B', 'F']);                   /* v1.6 */
  Logger.log('✅ 出貨中心訂單試算表：' + ss.getUrl());
  return ss.getUrl();
}

function dcoSS_() {
  var p = PropertiesService.getScriptProperties();
  var id = p.getProperty('DCO_SS_ID');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) { } }
  var ss = SpreadsheetApp.create('出貨中心訂單（採購系統）');
  try { ss.setSpreadsheetTimeZone(DCO_TZ); } catch (e) { }
  p.setProperty('DCO_SS_ID', ss.getId());
  ss.getSheets()[0].setName(SH_LINE);
  return ss;
}

function dcoEnsure_(ss, name, hdr, textCols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var lc = sh.getLastColumn();
  var cur = lc ? sh.getRange(1, 1, 1, lc).getValues()[0] : [];
  if (!cur.length || !String(cur[0] || '').trim()) {
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    sh.setFrozenRows(1);
  }
  (textCols || []).forEach(function (c) { sh.getRange(c + ':' + c).setNumberFormat('@'); });   /* 訂單號、店號、sku_id 一律文字，免得被轉成數字或日期 */
  return sh;
}

function dcoSheet_(name) {
  var ss = dcoSS_();
  var sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}

/* ── 讀表工具 ───────────────────────────────────────────────── */
function dcoRead_(name) {
  var sh = dcoSheet_(name);
  var v = sh.getDataRange().getValues();
  var h = v[0] || [];
  var rows = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 };
    for (var j = 0; j < h.length; j++) o[h[j]] = v[i][j];
    rows.push(o);
  }
  return { sh: sh, h: h, rows: rows };
}
function dcoFmt_(d) { return (d instanceof Date) ? Utilities.formatDate(d, DCO_TZ, 'yyyy-MM-dd HH:mm:ss') : String(d || ''); }
function dcoNum_(v) { var n = Number(v); return isFinite(n) ? n : 0; }
function dcoStr_(v) { return String(v == null ? '' : v).trim(); }
function dcoLog_(action, id, store, obj) {
  try { dcoSheet_(SH_LOG).appendRow([new Date(), action, id || '', store || '', JSON.stringify(obj || {}).slice(0, 45000)]); } catch (e) { }
}
function dcoOut_(o, cb) {
  var s = JSON.stringify(o);
  if (cb && /^[A-Za-z_$][\w$]*$/.test(cb)) return ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

/* 把明細列組回訂單 */
function dcoGroup_(rows) {
  var by = {}, order = [];
  rows.forEach(function (r) {
    var id = dcoStr_(r['訂單號']); if (!id) return;
    var o = by[id];
    if (!o) {
      o = by[id] = { id: id, store: dcoStr_(r['店號']), time: dcoFmt_(r['下單時間']), status: dcoStr_(r['狀態']), note: dcoStr_(r['備註']),
        shipTime: dcoFmt_(r['出貨時間']), hct: dcoStr_(r['新竹貨號']), cid: dcoStr_(r['cid']), updated: dcoFmt_(r['更新時間']), total: 0, lines: [] };
      order.push(o);
    }
    var ln = { line: dcoNum_(r['行號']), sku_id: dcoStr_(r['sku_id']), name: dcoStr_(r['品名']), unit: dcoStr_(r['單位']),
      qty: dcoNum_(r['訂購量']), price: dcoNum_(r['單價']), ship: (r['實出量'] === '' || r['實出量'] == null) ? null : dcoNum_(r['實出量']), amount: dcoNum_(r['小計']) };
    o.lines.push(ln); o.total += ln.amount;
  });
  order.forEach(function (o) { o.lines.sort(function (a, b) { return a.line - b.line; }); o.total = Math.round(o.total); });
  return order;
}

/* ── GET ────────────────────────────────────────────────────── */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback;
  try {
    if (p.action === 'ping') return dcoOut_({ ok: true, ver: DCO_VER, hct: hctCfg_().env }, cb);
    if (p.token !== DCO_TOKEN) return dcoOut_({ ok: false, error: 'token' }, cb);
    if (p.action === 'orders') return dcoOut_({ ok: true, orders: dcoOrders_(p) }, cb);
    if (p.action === 'stock') return dcoOut_({ ok: true, stock: dcoStock_(), locs: dcoLocs_() }, cb);   /* v1.2：多回 locs，舊前端不讀這欄 */
    if (p.action === 'locs') return dcoOut_({ ok: true, locs: dcoLocs_() }, cb);
    if (p.action === 'vmarks') return dcoOut_({ ok: true, marks: vmList_(p) }, cb);   /* v1.6 */
    if (p.action === 'invsess') return dcoOut_({ ok: true, sess: isGet_(p.store) }, cb);   /* v1.9 */
    if (p.action === 'hct') { var hp = hctPub_(); return dcoOut_({ ok: true, env: hp.env, cfg: hp, labels: hctList_(p) }, cb); }   /* v1.3；v1.5 多回 cfg（客代清單、派遣，不含帳密） */
    return dcoOut_({ ok: false, error: '未知的 action' }, cb);
  } catch (err) {
    return dcoOut_({ ok: false, error: String(err && err.message || err) }, cb);
  }
}

/* 參數：store（店號）、status（逗號分隔）、from/to（下單日 yyyy-MM-dd）、sfrom/sto（出貨日）、id、cid、days（預設 60）
   沒給任何日期條件時只回近 days 天下單的 ＋ 還沒處理完的（待處理、備貨中不論多久都回） */
function dcoOrders_(p) {
  var sts = p.status ? String(p.status).split(',') : null;
  var hasDate = p.from || p.to || p.sfrom || p.sto;
  var since = new Date(Date.now() - (dcoNum_(p.days) || 60) * 86400000);
  var list = dcoOrdersAll_().filter(function (o) {   /* v1.1：整表讀＋組訂單走快取；篩選照舊 */
    if (p.id && o.id !== p.id) return false;
    if (p.cid && o.cid !== p.cid) return false;
    if (p.store && o.store !== String(p.store)) return false;
    if (sts && sts.indexOf(o.status) < 0) return false;
    var od = o.time.slice(0, 10), sd = o.shipTime.slice(0, 10);
    if (p.from && od < p.from) return false;
    if (p.to && od > p.to) return false;
    if (p.sfrom && (!sd || sd < p.sfrom)) return false;
    if (p.sto && (!sd || sd > p.sto)) return false;
    if (!hasDate && !p.id && !p.cid && !OPEN_ST[o.status] && new Date(o.time.replace(' ', 'T') + '+08:00') < since) return false;
    return true;
  });
  list.sort(function (a, b) { return a.time < b.time ? 1 : -1; });
  return list;
}

function dcoStock_() {
  return dcoCached_('stock', function () {   /* v1.1：走快取 */
    return dcoRead_(SH_STOCK).rows.filter(function (r) { return dcoStr_(r['sku_id']); }).map(function (r) {
      return { sku_id: dcoStr_(r['sku_id']), name: dcoStr_(r['品名']), mode: dcoStr_(r['管理方式']) || '記數量', qty: dcoNum_(r['庫存']), time: dcoFmt_(r['更新時間']), note: dcoStr_(r['備註']) };
    });
  });
}

/* ── v1.6：門市「本輪已叫貨／不叫」標記 ───────────────────────── */
/* 分頁不存在時（舊試算表）補建；dcoSheet_ 只在缺分頁時跑 setup，這裡直接補這一張 */
function vmTable_() {
  var ss = dcoSS_();
  if (!ss.getSheetByName(SH_VMARK)) dcoEnsure_(ss, SH_VMARK, H_VMARK, ['A', 'B', 'F']);
  return dcoRead_(SH_VMARK);
}
function vmList_(p) {
  var store = dcoStr_(p.store), since = new Date(Date.now() - (dcoNum_(p.days) || 45) * 86400000), by = {}, list = [];
  vmTable_().rows.forEach(function (r) {
    var mid = dcoStr_(r['標記號']); if (!mid) return;
    if (store && dcoStr_(r['店號']) !== store) return;
    var t = r['標記時間'];
    if (t instanceof Date && t < since) return;
    var m = by[mid];
    if (!m) {
      m = by[mid] = { mid: mid, store: dcoStr_(r['店號']), vendor: dcoStr_(r['廠商']), kind: dcoStr_(r['類型']), time: dcoFmt_(t),
        status: dcoStr_(r['狀態']) || '有效', undoTime: dcoFmt_(r['復原時間']), items: [] };
      list.push(m);
    }
    m.items.push({ sku_id: dcoStr_(r['sku_id']), name: dcoStr_(r['品名']), qty: dcoNum_(r['數量']), unit: dcoStr_(r['單位']) });
  });
  list.sort(function (a, b) { return a.time < b.time ? 1 : -1; });
  return list;
}
function vmMark_(b) {
  var store = dcoStr_(b.store), mid = dcoStr_(b.mid), kind = dcoStr_(b.kind) === '不叫' ? '不叫' : '叫貨';
  var items = (b.items || []).filter(function (it) { return dcoStr_(it.sku_id); });
  if (!store) return { ok: false, error: '缺少店號' };
  if (!mid) return { ok: false, error: '缺少標記號' };
  if (!items.length) return { ok: false, error: '沒有品項' };
  var t = vmTable_();
  for (var i = 0; i < t.rows.length; i++) if (dcoStr_(t.rows[i]['標記號']) === mid) return { ok: true, mid: mid, dup: true };
  var now = new Date(), vendor = dcoStr_(b.vendor).slice(0, 60);
  var out = items.map(function (it) {
    return [mid, store, vendor, kind, now, dcoStr_(it.sku_id), dcoStr_(it.name).slice(0, 100), dcoNum_(it.qty), dcoStr_(it.unit).slice(0, 20), '有效', ''];
  });
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_VMARK.length).setValues(out);
  dcoLog_('vmark', mid, store, { vendor: vendor, kind: kind, n: out.length });
  return { ok: true, mid: mid, time: dcoFmt_(now), n: out.length };
}
function vmUnmark_(b) {
  var store = dcoStr_(b.store), mid = dcoStr_(b.mid);
  if (!mid) return { ok: false, error: '缺少標記號' };
  var t = vmTable_(), rs = t.rows.filter(function (r) { return dcoStr_(r['標記號']) === mid && (!store || dcoStr_(r['店號']) === store); });
  if (!rs.length) return { ok: false, error: '找不到這筆標記' };
  var now = new Date(), cS = dcoCol_(t, '狀態'), cU = dcoCol_(t, '復原時間');
  rs.forEach(function (r) { t.sh.getRange(r._row, cS).setValue('已復原'); t.sh.getRange(r._row, cU).setValue(now); });
  dcoLog_('vunmark', mid, store, { n: rs.length });
  return { ok: true, mid: mid, n: rs.length };
}

/* ── v1.9：盤點輪次（一店一列；數字存文字免得被轉成科學記號或日期） ── */
function isTable_() {
  var ss = dcoSS_();
  if (!ss.getSheetByName(SH_INVSESS)) dcoEnsure_(ss, SH_INVSESS, H_INVSESS, ['A', 'B', 'C']);
  return dcoRead_(SH_INVSESS);
}
function isRow_(t, store) { for (var i = 0; i < t.rows.length; i++) if (dcoStr_(t.rows[i]['店號']) === store) return t.rows[i]; return null; }
function isOut_(r) { return r ? { store: dcoStr_(r['店號']), start: Number(dcoStr_(r['盤點開始'])) || 0, setAt: Number(dcoStr_(r['設定時間'])) || 0 } : null; }
function isGet_(store) { store = dcoStr_(store); if (!store) return null; return isOut_(isRow_(isTable_(), store)); }
function isSet_(b) {
  var store = dcoStr_(b.store), start = Math.max(0, Math.floor(Number(b.start) || 0)), setAt = Math.floor(Number(b.setAt) || 0);
  if (!store) return { ok: false, error: '缺少店號' };
  if (!(setAt > 0)) return { ok: false, error: '缺少設定時間' };
  var t = isTable_(), r = isRow_(t, store), cur = isOut_(r);
  if (cur && cur.setAt >= setAt) return { ok: true, stale: true, sess: cur };   /* 別台裝置後來又設過：不蓋 */
  var row = [store, String(start), String(setAt), Utilities.formatDate(new Date(setAt), DCO_TZ, 'yyyy-MM-dd HH:mm:ss'), start > 0 ? '新的一週' : '延續歷史'];
  if (r) t.sh.getRange(r._row, 1, 1, row.length).setValues([row]);
  else t.sh.getRange(t.sh.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  dcoLog_('invsessSet', '', store, { start: start, setAt: setAt });
  return { ok: true, sess: { store: store, start: start, setAt: setAt } };
}

/* ── POST ───────────────────────────────────────────────────── */
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { return dcoOut_({ ok: false, error: '格式錯誤' }); }
  if (body.token !== DCO_TOKEN) return dcoOut_({ ok: false, error: 'token' });
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return dcoOut_({ ok: false, error: 'busy' });
  try {
    var a = body.action, r;
    if (a === 'create') r = dcoCreate_(body);
    else if (a === 'cancel') r = dcoCancel_(body);
    else if (a === 'status') r = dcoStatus_(body);
    else if (a === 'ship') r = dcoShip_(body);
    else if (a === 'stockSet') r = dcoStockSet_(body);
    else if (a === 'stockAdj') r = dcoStockAdj_(body);
    else if (a === 'locSet') r = dcoLocSet_(body);   /* v1.2 */
    else if (a === 'edit') r = dcoEdit_(body);       /* v1.2 */
    else if (a === 'merge') r = dcoMerge_(body);     /* v1.2 */
    else if (a === 'unmerge') r = dcoUnmerge_(body); /* v1.7 */
    else if (a === 'hctCreate') r = hctCreate_(body);   /* v1.3 */
    else if (a === 'hctVoid') r = hctVoid_(body);       /* v1.3 */
    else if (a === 'hctReport') r = hctReport_(body);   /* v1.3 */
    else if (a === 'hctDispatch') r = hctDispatch_(body);             /* v1.5 */
    else if (a === 'hctDispatchCancel') r = hctDispatchCancel_(body); /* v1.5 */
    else if (a === 'hctQuery') r = hctQuery_(body);                   /* v1.5 */
    else if (a === 'vmark') r = vmMark_(body);                        /* v1.6 */
    else if (a === 'vunmark') r = vmUnmark_(body);                    /* v1.6 */
    else if (a === 'invsessSet') r = isSet_(body);                    /* v1.9 */
    else r = { ok: false, error: '未知的 action' };
    return dcoOut_(r);
  } catch (err) {
    return dcoOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
    dcoCacheBump_();   /* v1.1：任何寫入動作結束（不論成功、失敗、中途出錯）都讓讀取快取失效 */
    lock.releaseLock();
  }
}

function dcoPad_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }

/* 下單：{store, items:[{sku_id,name,unit,qty,price}], note, cid}
   cid＝前端產生的一次性編號；同一個 cid 再送一次只回原訂單號（防止網路慢時重按變兩張單） */
function dcoCreate_(b) {
  var store = dcoStr_(b.store);
  var items = (b.items || []).filter(function (it) { return dcoStr_(it.sku_id) && dcoNum_(it.qty) > 0; });
  if (!store) return { ok: false, error: '缺少店號' };
  if (!items.length) return { ok: false, error: '沒有品項' };
  var t = dcoRead_(SH_LINE);
  var cid = dcoStr_(b.cid);
  if (cid) {
    for (var i = 0; i < t.rows.length; i++) if (dcoStr_(t.rows[i]['cid']) === cid) return { ok: true, id: dcoStr_(t.rows[i]['訂單號']), dup: true };
  }
  var now = new Date();
  /* 2026-10-22（經營者 10/22）：訂單號改全數字＝年月日 8 碼＋店號 2 碼＋流水號 2 碼（例 202610180201），跟 Shopline 一樣；
     之前的 DC261018-02-01 舊單號不動。訂單號欄是文字格式（dcoEnsure_），不會被試算表轉成數字 */
  var prefix = Utilities.formatDate(now, DCO_TZ, 'yyyyMMdd') + dcoPad_(store);
  var seen = {};
  t.rows.forEach(function (r) { var id = dcoStr_(r['訂單號']); if (id.indexOf(prefix) === 0) seen[id] = 1; });
  var seq = Object.keys(seen).length + 1, id = prefix + dcoPad_(seq);
  while (seen[id]) { seq++; id = prefix + dcoPad_(seq); }
  var note = dcoStr_(b.note).slice(0, 200), total = 0;
  var out = items.map(function (it, k) {
    var q = dcoNum_(it.qty), pr = dcoNum_(it.price), amt = Math.round(q * pr);
    total += amt;
    return [id, store, now, '待處理', k + 1, dcoStr_(it.sku_id), dcoStr_(it.name), dcoStr_(it.unit), q, pr, '', amt, note, '', '', cid, now];
  });
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_LINE.length).setValues(out);
  dcoLog_('create', id, store, { n: out.length, total: total });
  return { ok: true, id: id, total: total, n: out.length };
}

function dcoRowsOf_(t, id) { return t.rows.filter(function (r) { return dcoStr_(r['訂單號']) === id; }); }
function dcoCol_(t, name) { return t.h.indexOf(name) + 1; }
function dcoSet_(t, row, name, val) { t.sh.getRange(row, dcoCol_(t, name)).setValue(val); }
/* 同一張訂單的明細是一次 setValues 寫進去的＝連續列；整段讀出、改完一次寫回（逐格寫一張 20 項的單要 100 次呼叫） */
function dcoWriteRows_(t, rs, fn) {
  var sorted = rs.slice().sort(function (a, b) { return a._row - b._row; }), blocks = [];
  sorted.forEach(function (r) {
    var last = blocks[blocks.length - 1];
    if (last && last.end + 1 === r._row) { last.items.push(r); last.end = r._row; }
    else blocks.push({ start: r._row, end: r._row, items: [r] });
  });
  blocks.forEach(function (bk) {
    var rng = t.sh.getRange(bk.start, 1, bk.items.length, t.h.length), vals = rng.getValues();
    bk.items.forEach(function (r, i) {
      var m = fn(r) || {};
      Object.keys(m).forEach(function (k) { var c = t.h.indexOf(k); if (c >= 0) vals[i][c] = m[k]; });
    });
    rng.setValues(vals);
  });
}

/* 取消：{id, by:'store'|'dc', store}。門市只能取消「待處理」的單；出貨中心可取消還沒出貨的 */
function dcoCancel_(b) {
  var id = dcoStr_(b.id), t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  var st = dcoStr_(rs[0]['狀態']);
  if (st === '已取消') return { ok: true, id: id, already: true };
  if (st === '已出貨') return { ok: false, error: '已出貨，不能取消' };
  if (b.by === 'store') {
    if (dcoStr_(b.store) !== dcoStr_(rs[0]['店號'])) return { ok: false, error: '不是本店的訂單' };
    if (st !== '待處理') return { ok: false, error: '出貨中心已開始備貨，請直接聯絡出貨中心' };
  }
  var now = new Date();
  dcoWriteRows_(t, rs, function () { return { '狀態': '已取消', '更新時間': now }; });
  var hv = hctVoidOpen_(id, now);   /* v1.8：還沒交寄的托運單一起作廢，不然「確認今日出貨」會把它送給新竹 */
  dcoLog_('cancel', id, rs[0]['店號'], { by: b.by || '', hctVoid: hv });
  return { ok: true, id: id };
}

/* 改狀態（目前只用在「開始備貨」）：{id, status:'備貨中'} */
function dcoStatus_(b) {
  var id = dcoStr_(b.id), to = dcoStr_(b.status), t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  if (to !== '備貨中' && to !== '待處理') return { ok: false, error: '這個狀態請用出貨或取消' };
  var st = dcoStr_(rs[0]['狀態']);
  if (!OPEN_ST[st]) return { ok: false, error: '訂單已是「' + st + '」' };
  var now = new Date();
  dcoWriteRows_(t, rs, function () { return { '狀態': to, '更新時間': now }; });
  dcoLog_('status', id, rs[0]['店號'], { to: to });
  return { ok: true, id: id, status: to };
}

/* 出貨：{id, lines:[{line, ship}], hct}。沒給的行＝照訂購量出。記數量的品項扣庫存。重送同一張已出貨的單不會重扣 */
function dcoShip_(b) {
  var id = dcoStr_(b.id), t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  var st = dcoStr_(rs[0]['狀態']);
  if (st === '已出貨') return { ok: true, id: id, already: true };
  if (!OPEN_ST[st]) return { ok: false, error: '訂單已是「' + st + '」' };
  /* v1.8：畫面上的版本跟現在不一樣（別人剛編輯／合併／取消合併過，行號已重編）就不出貨，避免實出量套到別的品項 */
  if (b.base && dcoStr_(b.base) !== dcoUpd_(rs)) return { ok: false, error: '這張訂單剛被改過（' + dcoUpd_(rs) + '），請按「🔄 重新整理」看最新內容再按出貨完成', code: 'conflict' };
  var ship = {};
  (b.lines || []).forEach(function (l) { ship[String(l.line)] = Math.max(0, dcoNum_(l.ship)); });
  var now = new Date(), used = {}, total = 0;
  dcoWriteRows_(t, rs, function (r) {
    var q = ship.hasOwnProperty(String(r['行號'])) ? ship[String(r['行號'])] : dcoNum_(r['訂購量']);
    var amt = Math.round(q * dcoNum_(r['單價']));
    total += amt;
    var sk = dcoStr_(r['sku_id']); used[sk] = (used[sk] || 0) + q;
    var vals = { '實出量': q, '小計': amt, '狀態': '已出貨', '出貨時間': now, '更新時間': now };
    if (b.hct) vals['新竹貨號'] = dcoStr_(b.hct);
    return vals;
  });
  var st2 = dcoRead_(SH_STOCK);
  st2.rows.forEach(function (s) {
    var sk = dcoStr_(s['sku_id']);
    if (used[sk] && (dcoStr_(s['管理方式']) || '記數量') === '記數量') {
      dcoSet_(st2, s._row, '庫存', dcoNum_(s['庫存']) - used[sk]);
      dcoSet_(st2, s._row, '更新時間', now);
    }
  });
  dcoLog_('ship', id, rs[0]['店號'], { total: total, used: used, hct: b.hct || '' });
  return { ok: true, id: id, total: total };
}

/* 設定庫存：{items:[{sku_id,name,mode:'記數量'|'無限量',qty,note}]}（整列覆蓋；mode＝無限量 會把這列刪掉） */
function dcoStockSet_(b) {
  var t = dcoRead_(SH_STOCK), now = new Date(), idx = {};
  t.rows.forEach(function (r) { idx[dcoStr_(r['sku_id'])] = r._row; });
  var del = [], n = 0;
  (b.items || []).forEach(function (it) {
    var sk = dcoStr_(it.sku_id); if (!sk) return;
    var mode = dcoStr_(it.mode) === '無限量' ? '無限量' : '記數量';
    if (mode === '無限量') { if (idx[sk]) { del.push(idx[sk]); delete idx[sk]; } n++; return; }
    var row = [sk, dcoStr_(it.name), mode, dcoNum_(it.qty), now, dcoStr_(it.note).slice(0, 200)];
    if (idx[sk]) t.sh.getRange(idx[sk], 1, 1, row.length).setValues([row]);
    else { t.sh.appendRow(row); idx[sk] = t.sh.getLastRow(); }
    n++;
  });
  del.sort(function (x, y) { return y - x; }).forEach(function (r) { t.sh.deleteRow(r); });
  dcoLog_('stockSet', '', '', { items: b.items });
  return { ok: true, n: n };
}

/* 庫存加減：{sku_id, name, delta, note}（進貨填正數；報廢填負數） */
function dcoStockAdj_(b) {
  var sk = dcoStr_(b.sku_id); if (!sk) return { ok: false, error: '缺少 sku_id' };
  var t = dcoRead_(SH_STOCK), now = new Date(), hit = null;
  t.rows.forEach(function (r) { if (dcoStr_(r['sku_id']) === sk) hit = r; });
  var d = dcoNum_(b.delta), q;
  if (hit) { q = dcoNum_(hit['庫存']) + d; dcoSet_(t, hit._row, '庫存', q); dcoSet_(t, hit._row, '更新時間', now); }
  else { q = d; t.sh.appendRow([sk, dcoStr_(b.name), '記數量', q, now, dcoStr_(b.note).slice(0, 200)]); }
  dcoLog_('stockAdj', '', '', { sku_id: sk, delta: d, note: b.note || '' });
  return { ok: true, sku_id: sk, qty: q };
}

/* ── v1.2：位置、編輯、合併 ─────────────────────────────────── */

/* 設定位置：{items:[{sku_id, name, loc}]}（位置清空＝刪這列）。一次最多 1000 項 */
function dcoLocSet_(b) {
  var items = (b.items || []).filter(function (it) { return dcoStr_(it.sku_id); }).slice(0, 1000);
  if (!items.length) return { ok: false, error: '沒有品項' };
  var t = dcoRead_(SH_LOC), now = new Date(), idx = {}, del = [], add = [], n = 0;
  t.rows.forEach(function (r) { var k = dcoStr_(r['sku_id']); if (k && !idx[k]) idx[k] = r._row; });
  items.forEach(function (it) {
    var sk = dcoStr_(it.sku_id), loc = dcoStr_(it.loc).slice(0, 30);
    if (!loc) { if (idx[sk]) { del.push(idx[sk]); delete idx[sk]; } n++; return; }
    var row = [sk, dcoStr_(it.name).slice(0, 80), loc, now];
    if (idx[sk]) t.sh.getRange(idx[sk], 1, 1, row.length).setValues([row]);
    else { add.push(row); idx[sk] = -1; }
    n++;
  });
  if (add.length) t.sh.getRange(t.sh.getLastRow() + 1, 1, add.length, H_LOC.length).setValues(add);
  del.sort(function (x, y) { return y - x; }).forEach(function (r) { t.sh.deleteRow(r); });
  dcoLog_('locSet', '', '', { n: n, items: items.map(function (it) { return [dcoStr_(it.sku_id), dcoStr_(it.loc)]; }) });
  return { ok: true, n: n };
}

/* 訂單的「更新時間」（前端拿來當 base；別人先改過就不一樣） */
function dcoUpd_(rs) { return rs.length ? dcoFmt_(rs[0]['更新時間']) : ''; }
/* 把這些列刪掉（由下往上刪，列號才不會跑掉） */
function dcoDelRows_(t, rs) {
  rs.map(function (r) { return r._row; }).sort(function (x, y) { return y - x; }).forEach(function (row) { t.sh.deleteRow(row); });
}
/* 依 H_LINE 欄序組一列 */
function dcoLineRow_(o) {
  return H_LINE.map(function (h) { return o.hasOwnProperty(h) ? o[h] : ''; });
}
function dcoLinesSum_(lines) {
  return lines.map(function (l) { return dcoStr_(l.name) + '×' + dcoNum_(l.qty) + (l.ship != null && l.ship !== '' ? '(出' + dcoNum_(l.ship) + ')' : '') + '@' + dcoNum_(l.price); }).join('、').slice(0, 3000);
}

/* 編輯整張訂單：{id, base, store?, note?, hct?, lines:[{sku_id, name, unit, qty, price, ship?}]}
   待處理／備貨中：實出留空，小計＝訂購×單價；已出貨：小計＝實出×單價，實出改了會補扣／退回「記數量」庫存。
   已取消的單不能編輯。lines 會整批取代（行號重新編 1、2、3…）。 */
function dcoEdit_(b) {
  var id = dcoStr_(b.id), t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  var st = dcoStr_(rs[0]['狀態']);
  if (st === '已取消') return { ok: false, error: '已取消的訂單不能編輯' };
  if (b.base && dcoStr_(b.base) !== dcoUpd_(rs)) return { ok: false, error: '這張訂單剛被改過（' + dcoUpd_(rs) + '），請按「🔄 重新整理」看最新內容再改', code: 'conflict' };
  var shipped = (st === '已出貨');
  var lines = (b.lines || []).filter(function (l) { return dcoStr_(l.sku_id) && dcoNum_(l.qty) > 0; });
  if (!lines.length) return { ok: false, error: '至少要留一個品項（整張不要請用「取消訂單」）' };
  if (lines.length > 300) return { ok: false, error: '品項太多（最多 300）' };
  var store = b.hasOwnProperty('store') && dcoStr_(b.store) ? dcoStr_(b.store) : dcoStr_(rs[0]['店號']);
  if (!/^\d{1,2}$/.test(store)) return { ok: false, error: '店號格式不對' };
  var note = b.hasOwnProperty('note') ? dcoStr_(b.note).slice(0, 300) : dcoStr_(rs[0]['備註']);
  var hct = b.hasOwnProperty('hct') ? dcoStr_(b.hct).slice(0, 400) : dcoStr_(rs[0]['新竹貨號']);
  /* v1.8：這張有系統開的托運單 → 貨號以試算表現值為準（開著編輯框時別人產生托運單，存檔不可把新貨號蓋回空白） */
  try { if (dcoSS_().getSheetByName(SH_HCT) && hctActiveAll_(hctTable_(), id, '').length) hct = dcoStr_(rs[0]['新竹貨號']); } catch (e) {}
  var now = new Date(), total = 0, oldUsed = {}, newUsed = {};
  rs.forEach(function (r) { if (shipped) { var k = dcoStr_(r['sku_id']); oldUsed[k] = (oldUsed[k] || 0) + dcoNum_(r['實出量']); } });
  var before = rs.map(function (r) { return { name: r['品名'], qty: r['訂購量'], ship: r['實出量'], price: r['單價'] }; });
  var out = lines.map(function (l, k) {
    var q = dcoNum_(l.qty), pr = Math.max(0, dcoNum_(l.price)), sh = '';
    if (shipped) { sh = (l.ship === '' || l.ship == null) ? q : Math.max(0, dcoNum_(l.ship)); var sk = dcoStr_(l.sku_id); newUsed[sk] = (newUsed[sk] || 0) + sh; }
    var amt = Math.round((shipped ? sh : q) * pr);
    total += amt;
    var o = {};
    o['訂單號'] = id; o['店號'] = store; o['下單時間'] = rs[0]['下單時間']; o['狀態'] = st; o['行號'] = k + 1;
    o['sku_id'] = dcoStr_(l.sku_id); o['品名'] = dcoStr_(l.name).slice(0, 120); o['單位'] = dcoStr_(l.unit).slice(0, 20);
    o['訂購量'] = q; o['單價'] = pr; o['實出量'] = sh; o['小計'] = amt; o['備註'] = note;
    o['出貨時間'] = rs[0]['出貨時間']; o['新竹貨號'] = hct; o['cid'] = rs[0]['cid']; o['更新時間'] = now;
    return dcoLineRow_(o);
  });
  dcoDelRows_(t, rs);
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_LINE.length).setValues(out);
  var adj = {};
  if (shipped) {   /* 已出貨：實出差多少，記數量的庫存就補扣／退回多少 */
    var keys = {};
    Object.keys(oldUsed).concat(Object.keys(newUsed)).forEach(function (k) { keys[k] = 1; });
    Object.keys(keys).forEach(function (k) { var d = (newUsed[k] || 0) - (oldUsed[k] || 0); if (d) adj[k] = d; });
    if (Object.keys(adj).length) {
      var st2 = dcoRead_(SH_STOCK);
      st2.rows.forEach(function (s) {
        var sk = dcoStr_(s['sku_id']);
        if (adj[sk] && (dcoStr_(s['管理方式']) || '記數量') === '記數量') {
          dcoSet_(st2, s._row, '庫存', dcoNum_(s['庫存']) - adj[sk]);
          dcoSet_(st2, s._row, '更新時間', now);
        }
      });
    }
  }
  dcoLog_('edit', id, store, { status: st, storeFrom: dcoStr_(rs[0]['店號']), before: before, after: dcoLinesSum_(lines), note: note, hct: hct, stockAdj: adj, total: total });
  return { ok: true, id: id, total: total, n: out.length, updated: dcoFmt_(now) };
}

/* 合併：{ids:[…], into?, base:{訂單號:更新時間}}。同一店、都還沒出貨（待處理／備貨中）才能合併。
   併入 into（沒給＝下單時間最早那張）；同品項同單位同單價的數量相加，其他照順序接在後面。
   其他張：狀態改「已取消」、備註改「已併入 <into>｜原備註」，明細保留備查（月結、各種統計本來就不算已取消）。 */
function dcoMerge_(b) {
  var ids = (b.ids || []).map(dcoStr_).filter(function (x, i, a) { return x && a.indexOf(x) === i; });
  if (ids.length < 2) return { ok: false, error: '至少要選兩張訂單' };
  var sent = dcoSentIds_(ids);   /* v1.7.1 */
  if (sent.length) return { ok: false, error: sent.join('、') + ' 已經用新竹寄出（托運單已交寄），不能合併；請先按「✅ 出貨完成」結案', code: 'sent' };
  var t = dcoRead_(SH_LINE), by = {}, store = null, base = b.base || {};
  for (var i = 0; i < ids.length; i++) {
    var rs = dcoRowsOf_(t, ids[i]);
    if (!rs.length) return { ok: false, error: '找不到訂單 ' + ids[i] };
    var st = dcoStr_(rs[0]['狀態']);
    if (!OPEN_ST[st]) return { ok: false, error: ids[i] + ' 已是「' + st + '」，只有待處理／備貨中的訂單能合併', code: st === '已取消' ? 'cancelled' : '' };
    var sto = dcoStr_(rs[0]['店號']);
    if (store !== null && sto !== store) return { ok: false, error: '不同門市的訂單不能合併' };
    store = sto;
    if (base[ids[i]] && dcoStr_(base[ids[i]]) !== dcoUpd_(rs)) return { ok: false, error: ids[i] + ' 剛被改過，請按「🔄 重新整理」再合併', code: 'conflict' };
    by[ids[i]] = rs;
  }
  function tm(id) { var v = by[id][0]['下單時間']; return v instanceof Date ? v.getTime() : String(v); }
  var into = dcoStr_(b.into);
  if (!into || !by[into]) into = ids.slice().sort(function (a, c) { var x = tm(a), y = tm(c); return x < y ? -1 : (x > y ? 1 : (a < c ? -1 : 1)); })[0];
  var others = ids.filter(function (x) { return x !== into; }).sort(function (a, c) { var x = tm(a), y = tm(c); return x < y ? -1 : (x > y ? 1 : 0); });
  var anyPrep = ids.some(function (x) { return dcoStr_(by[x][0]['狀態']) === '備貨中'; });
  var status = anyPrep ? '備貨中' : '待處理';
  /* 合併明細 */
  var merged = [], key = {};
  [into].concat(others).forEach(function (oid) {
    by[oid].slice().sort(function (a, c) { return dcoNum_(a['行號']) - dcoNum_(c['行號']); }).forEach(function (r) {
      var k = dcoStr_(r['sku_id']) + '\u0001' + dcoStr_(r['單位']) + '\u0001' + dcoNum_(r['單價']);
      if (key.hasOwnProperty(k)) { merged[key[k]].qty += dcoNum_(r['訂購量']); return; }
      key[k] = merged.length;
      merged.push({ sku_id: dcoStr_(r['sku_id']), name: dcoStr_(r['品名']), unit: dcoStr_(r['單位']), qty: dcoNum_(r['訂購量']), price: dcoNum_(r['單價']) });
    });
  });
  var notes = [];
  var n0 = dcoStr_(by[into][0]['備註']); if (n0) notes.push(n0);
  others.forEach(function (oid) { var n1 = dcoStr_(by[oid][0]['備註']); if (n1) notes.push('（' + oid + '）' + n1); });
  var note = notes.join('；').slice(0, 300);
  var hct = dcoStr_(by[into][0]['新竹貨號']);
  var now = new Date(), total = 0, r0 = by[into][0];
  var out = merged.map(function (l, k) {
    var amt = Math.round(l.qty * l.price); total += amt;
    var o = {};
    o['訂單號'] = into; o['店號'] = store; o['下單時間'] = r0['下單時間']; o['狀態'] = status; o['行號'] = k + 1;
    o['sku_id'] = l.sku_id; o['品名'] = l.name; o['單位'] = l.unit; o['訂購量'] = l.qty; o['單價'] = l.price; o['實出量'] = ''; o['小計'] = amt;
    o['備註'] = note; o['出貨時間'] = ''; o['新竹貨號'] = hct; o['cid'] = r0['cid']; o['更新時間'] = now;
    return dcoLineRow_(o);
  });
  /* 其他張：原地改狀態與備註（明細保留） */
  others.forEach(function (oid) {
    var on = dcoStr_(by[oid][0]['備註']);
    dcoWriteRows_(t, by[oid], function () { return { '狀態': '已取消', '備註': (DCO_MERGED_PREFIX + into + (on ? '｜' + on : '')).slice(0, 300), '更新時間': now }; });
    hctVoidOpen_(oid, now);   /* v1.8：被併掉那張的未交寄托運單作廢 */
  });
  dcoDelRows_(t, by[into]);
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_LINE.length).setValues(out);
  var prev = {};   /* v1.7：合併前每張的狀態與備註，取消合併時用 */
  ids.forEach(function (x) { prev[x] = { status: dcoStr_(by[x][0]['狀態']), note: dcoStr_(by[x][0]['備註']) }; });
  dcoLog_('merge', into, store, { merged: others, n: out.length, total: total, status: status, prev: prev });
  return { ok: true, id: into, merged: others, n: out.length, total: total, status: status, updated: dcoFmt_(now) };
}

/* v1.7.1：這些訂單裡，新竹托運單已交寄（貨已寄出）的有哪幾張；讀不到托運單表就當沒有 */
function dcoSentIds_(ids) {
  var out = [];
  try {
    hctTable_().rows.forEach(function (r) {
      var id = dcoStr_(r['訂單號']);
      if (ids.indexOf(id) >= 0 && dcoStr_(r['狀態']) === '已交寄' && out.indexOf(id) < 0) out.push(id);
    });
  } catch (e) { }
  return out;
}
/* v1.7：取消合併：{id:<併入的那張>, base?:畫面上那張的更新時間}
   找出備註是「已併入 <id>」的已取消訂單（同一店）→ 各自回到合併前的狀態與備註（明細本來就保留著）；
   併入那張扣掉它們的品項數量（同品項同單位同單價），扣完剩下的就是合併前的內容。
   合併後那張若又被編輯到扣不回去（數量比被併入的少），就擋下、請改用「✏️ 編輯」手動處理。 */
function dcoUnmerge_(b) {
  var into = dcoStr_(b.id), t = dcoRead_(SH_LINE), rsI = dcoRowsOf_(t, into);
  if (!rsI.length) return { ok: false, error: '找不到訂單 ' + into };
  var stI = dcoStr_(rsI[0]['狀態']);
  if (!OPEN_ST[stI]) return { ok: false, error: into + ' 已是「' + stI + '」，只有還沒出貨的訂單能取消合併' };
  if (b.base && dcoStr_(b.base) !== dcoUpd_(rsI)) return { ok: false, error: into + ' 剛被改過，請按「🔄 重新整理」再取消合併', code: 'conflict' };
  var store = dcoStr_(rsI[0]['店號']), tag = DCO_MERGED_PREFIX + into, by = {}, others = [];
  t.rows.forEach(function (r) {
    var id = dcoStr_(r['訂單號']), nt = dcoStr_(r['備註']);
    if (!id || id === into || dcoStr_(r['狀態']) !== '已取消' || dcoStr_(r['店號']) !== store) return;
    if (nt !== tag && nt.indexOf(tag + '｜') !== 0) return;
    if (!by[id]) { by[id] = []; others.push(id); }
    by[id].push(r);
  });
  if (!others.length) return { ok: false, error: into + ' 沒有被合併進來的訂單' };
  function key(r) { return dcoStr_(r['sku_id']) + '\u0001' + dcoStr_(r['單位']) + '\u0001' + dcoNum_(r['單價']); }
  /* 併入那張扣掉其他張的數量 */
  var need = {};
  others.forEach(function (oid) { by[oid].forEach(function (r) { var k = key(r); need[k] = (need[k] || 0) + dcoNum_(r['訂購量']); }); });
  var keep = [];
  rsI.slice().sort(function (a, c) { return dcoNum_(a['行號']) - dcoNum_(c['行號']); }).forEach(function (r) {
    var k = key(r), q = dcoNum_(r['訂購量']), take = Math.min(q, need[k] || 0);
    need[k] = (need[k] || 0) - take; q = Math.round((q - take) * 1e6) / 1e6;
    if (q > 0) keep.push({ r: r, qty: q });
  });
  var short = Object.keys(need).filter(function (k) { return need[k] > 1e-6; });
  if (short.length) return { ok: false, error: '合併後 ' + into + ' 有品項被改過，數量對不起來，無法自動拆回；請用「✏️ 編輯」手動調整' };
  if (!keep.length) return { ok: false, error: '扣掉被併入的品項後 ' + into + ' 沒有品項了，無法自動拆回' };
  /* 合併前的狀態與備註：v1.7 起合併時有記；以前的合併從 dc_log 推回 */
  var snap = dcoMergeSnap_(into, others);
  function origNote(oid) { var nt = dcoStr_(by[oid][0]['備註']); return nt === tag ? '' : nt.slice((tag + '｜').length); }
  var noteI = dcoStr_(rsI[0]['備註']);
  if (snap.prev && snap.prev[into]) noteI = dcoStr_(snap.prev[into].note);
  else others.forEach(function (oid) {
    var on = origNote(oid); if (!on) return;
    var seg = '（' + oid + '）' + on;
    if (noteI.indexOf('；' + seg) >= 0) noteI = noteI.replace('；' + seg, '');
    else if (noteI.indexOf(seg + '；') === 0) noteI = noteI.slice((seg + '；').length);
    else if (noteI === seg) noteI = '';
  });
  function prevSt(oid) {
    if (snap.prev && snap.prev[oid] && OPEN_ST[snap.prev[oid].status]) return snap.prev[oid].status;
    return dcoPrevStatus_(snap.logs, oid, snap.time);
  }
  var now = new Date(), stInto = prevSt(into), back = [];
  others.forEach(function (oid) {
    var s1 = prevSt(oid), on = (snap.prev && snap.prev[oid]) ? dcoStr_(snap.prev[oid].note) : origNote(oid);
    dcoWriteRows_(t, by[oid], function () { return { '狀態': s1, '備註': on, '更新時間': now }; });
    back.push({ id: oid, status: s1 });
  });
  var r0 = rsI[0], total = 0;
  var out = keep.map(function (x, k) {
    var r = x.r, amt = Math.round(x.qty * dcoNum_(r['單價'])); total += amt;
    var o = {};
    H_LINE.forEach(function (h) { o[h] = r[h]; });
    o['狀態'] = stInto; o['行號'] = k + 1; o['訂購量'] = x.qty; o['實出量'] = ''; o['小計'] = amt; o['備註'] = noteI;
    o['出貨時間'] = ''; o['新竹貨號'] = dcoStr_(r0['新竹貨號']); o['cid'] = r0['cid']; o['下單時間'] = r0['下單時間']; o['更新時間'] = now;
    return dcoLineRow_(o);
  });
  dcoDelRows_(t, rsI);
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_LINE.length).setValues(out);
  dcoLog_('unmerge', into, store, { back: back, status: stInto, n: out.length, total: total, snap: snap.prev ? 'log' : 'derived' });
  return { ok: true, id: into, status: stInto, back: back, n: out.length, total: total, updated: dcoFmt_(now) };
}
/* 找這次合併的 dc_log：時間＋合併前快照（v1.7 起才有 prev）；另回傳 dc_log 全部列給 dcoPrevStatus_ 用 */
function dcoMergeSnap_(into, others) {
  var logs = [];
  try { logs = dcoRead_(SH_LOG).rows; } catch (e) { }
  var hit = null;
  logs.forEach(function (r) {
    if (dcoStr_(r['動作']) !== 'merge' || dcoStr_(r['訂單號']) !== into) return;
    var c = {}; try { c = JSON.parse(r['內容'] || '{}'); } catch (e) { }
    if ((c.merged || []).some(function (x) { return others.indexOf(x) >= 0; })) hit = { time: r['時間'], c: c };
  });
  return { logs: logs, time: hit ? hit.time : null, prev: hit && hit.c.prev ? hit.c.prev : null };
}
/* 合併前這張是什麼狀態：看合併以前最後一筆建單／改狀態紀錄（查不到＝待處理） */
function dcoPrevStatus_(logs, id, before) {
  var lim = before instanceof Date ? before.getTime() : Infinity, st = '待處理';
  (logs || []).forEach(function (r) {
    if (dcoStr_(r['訂單號']) !== id) return;
    var tm = r['時間'] instanceof Date ? r['時間'].getTime() : 0;
    if (tm >= lim) return;
    var a = dcoStr_(r['動作']);
    if (a === 'create') st = '待處理';
    else if (a === 'status') { var c = {}; try { c = JSON.parse(r['內容'] || '{}'); } catch (e) { } if (OPEN_ST[c.to]) st = c.to; }
  });
  return st;
}

/* ── v1.3／v1.5：新竹物流托運單 ──────────────────────────────────
   規格：《新竹物流配送出貨文件介接規格文件 v16.0.4》（2026/10/12，2026-10-27 新竹寄來）。SOAP 1.2、JSON 版方法：
     TransData_Json（取號＋標籤；同 epino＋edelno 再送＝更正，整筆覆蓋）、TransReport_Json（確認出貨，參數 sCompany／sPassword／dsCusJson）、
     QueryEDELNO_Json（用訂單編號查貨號）、TransDispatchOrder_Json（呼叫派遣收貨；回 {rtn_code, rtn_data}）、TransDataCancel_Json（取消派遣）。
     一批：要標籤圖 ≤5、其他 ≤20。欄位不可有 & " > < '。訂單編號（epino）只能英文和數字、同一出貨日不可重複。
   測試環境 https://hctrt.hct.com.tw/edi_webservice2_test/service1.asmx；正式 https://hctrt.hct.com.tw/edi_webservice2/service1.asmx
   一個客代一種計費：重量（W：一筆＝N 箱＋總重量 kg）或總長（S：一筆只能 1 箱，eqamt＝60／90／120／150／151～999，兩箱要拆兩筆）。
   託運單號（epino）＝訂單號（12 碼數字）＋序號 2 碼；同一張訂單作廢重開會換序號。
   指令碼屬性（專案設定 → 指令碼屬性；都沒設＝新竹公開測試帳號 test／test1）：
     HCT_COMPANY、HCT_PWD：公司帳號、密碼
     HCT_ENV：test（測試環境）｜live（正式）；沒設＝有帳號就當 live
     HCT_CS：客代清單，分號隔開，每組「客代|W或S|顯示名稱|商品種類」，第一組是預設。例：05974350011|W|一般件（按重量）|001;05974350002|S|一般件（按尺寸）|001
     HCT_ESSTNO：出貨站 4 碼（選填）
     HCT_SENDER／HCT_SENDER_TEL／HCT_SENDER_ADDR：出貨人名稱、電話、地址（選填；沒填用新竹帳號的預設值）
     HCT_COMPANY_NAME：託運總表「客代」後面的公司名稱
     HCT_DISPATCH：Y＝確認出貨後自動呼叫新竹來收件；HCT_DISPATCH_MARK：派遣備註（【平台或公司】-出貨人，30 字內） */
var HCT_URL_LIVE = 'https://hctrt.hct.com.tw/edi_webservice2/service1.asmx';
var HCT_URL_TEST = 'https://hctrt.hct.com.tw/edi_webservice2_test/service1.asmx';
var HCT_OK = { '已建單': 1, '已交寄': 1 };
var HCT_ARGS = { TransReport_Json: ['sCompany', 'sPassword', 'dsCusJson'] };   /* 其他方法都是 company／password／json */
var HCT_SIZES = { 60: 1, 90: 1, 120: 1, 150: 1 };
var HCT_MORE_COLS = ['客代', '計費', '商品種類', '箱序', '派遣', '派遣時間'];   /* v1.5 加在 dc_hct 最後 */

function hctCfg_() {
  var p = PropertiesService.getScriptProperties();
  var c = p.getProperty('HCT_COMPANY'), w = p.getProperty('HCT_PWD');
  var has = !!(c && w), envP = String(p.getProperty('HCT_ENV') || '').toLowerCase();
  var env = !has ? 'test' : (envP === 'test' ? 'test' : 'live');
  var cs = String(p.getProperty('HCT_CS') || p.getProperty('HCT_ESCSNO') || '').split(';').map(function (s) {
    var a = s.split('|').map(dcoStr_);
    return a[0] || a.length > 1 ? { code: a[0] || '', mode: (a[1] || 'W').toUpperCase() === 'S' ? 'S' : 'W', name: a[2] || (a[1] && a[1].toUpperCase() === 'S' ? '按尺寸' : '按重量'), kind: a[3] || '001' } : null;
  }).filter(Boolean);
  if (!cs.length) cs = [{ code: '', mode: 'W', name: '一般件', kind: '001' }];
  return {
    env: env, url: (has && envP === 'test') ? HCT_URL_TEST : HCT_URL_LIVE,   /* 公開測試帳號 test／test1 用正式網址 */
    company: has ? c : 'test', pwd: has ? w : 'test1', cs: cs,
    esstno: p.getProperty('HCT_ESSTNO') || '',
    sender: p.getProperty('HCT_SENDER') || '', senderTel: p.getProperty('HCT_SENDER_TEL') || '', senderAddr: p.getProperty('HCT_SENDER_ADDR') || '',
    companyName: p.getProperty('HCT_COMPANY_NAME') || '',
    dispatch: String(p.getProperty('HCT_DISPATCH') || '').toUpperCase() === 'Y',
    dispatchMark: (p.getProperty('HCT_DISPATCH_MARK') || '自己做烘焙聚樂部-出貨中心').slice(0, 30)
  };
}
function hctCsOf_(cfg, code) {
  code = dcoStr_(code);
  for (var i = 0; i < cfg.cs.length; i++) if (cfg.cs[i].code === code) return cfg.cs[i];
  return null;
}
/* 前端看得到的設定（不含帳密） */
function hctPub_() {
  var c = hctCfg_();
  return { env: c.env, cs: c.cs, dispatch: c.dispatch, companyName: c.companyName };
}
function hctClean_(v, n) { return dcoStr_(v).replace(/[&"<>']/g, '').slice(0, n || 100); }
/* 呼叫新竹 SOAP 1.2：回傳 *Result 解析後的 JSON（陣列或物件） */
function hctCall_(method, arr) {
  var cfg = hctCfg_(), a = HCT_ARGS[method] || ['company', 'password', 'json'];
  function esc(x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function tag(k, v) { return '<' + k + '>' + esc(v) + '</' + k + '>'; }
  var env = '<?xml version="1.0" encoding="utf-8"?><soap12:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
    + '<soap12:Body><' + method + ' xmlns="http://tempuri.org/">' + tag(a[0], cfg.company) + tag(a[1], cfg.pwd) + tag(a[2], JSON.stringify(arr)) + '</' + method + '></soap12:Body></soap12:Envelope>';
  var res = UrlFetchApp.fetch(cfg.url, { method: 'post', contentType: 'application/soap+xml; charset=utf-8', payload: env, muteHttpExceptions: true });
  var code = res.getResponseCode(), txt = res.getContentText('UTF-8');
  if (code !== 200) throw new Error('新竹系統回應 ' + code + '：' + txt.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 200));
  var m = txt.match(new RegExp('<' + method + 'Result>([\\s\\S]*?)</' + method + 'Result>'));
  if (!m) throw new Error('看不懂新竹系統的回應');
  var body = m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#xD;/g, '\r').replace(/&#xA;/g, '\n').replace(/&amp;/g, '&');
  return JSON.parse(body);
}
function hctArr_(x) { return Array.isArray(x) ? x : (x ? [x] : []); }
function hctHexToB64_(hex) {
  hex = String(hex || '').replace(/\s+/g, '');
  if (!hex || hex.length % 2) return '';
  var bytes = new Array(hex.length / 2);
  for (var i = 0; i < bytes.length; i++) { var v = parseInt(hex.substr(i * 2, 2), 16); bytes[i] = v > 127 ? v - 256 : v; }
  return Utilities.base64Encode(bytes);
}
function hctToday_() { return Utilities.formatDate(new Date(), DCO_TZ, 'yyyyMMdd'); }
/* dc_hct 讀表（v1.5 先補欄位標題：舊表沒有的欄加在最後） */
function hctTable_() {
  var sh = dcoSheet_(SH_HCT), lc = Math.max(1, sh.getLastColumn()), h = sh.getRange(1, 1, 1, lc).getValues()[0].map(dcoStr_);
  var miss = H_HCT.filter(function (k) { return h.indexOf(k) < 0; });
  if (miss.length) { sh.getRange(1, h.length + 1, 1, miss.length).setValues([miss]); sh.getRange(1, h.length + 1, sh.getMaxRows(), miss.length).setNumberFormat('@'); }
  return dcoRead_(SH_HCT);
}
function hctRowVals_(t, o) { return t.h.map(function (k) { return o.hasOwnProperty(k) ? o[k] : ''; }); }
/* 標籤圖：依託運單號找列（只讀找到的那幾列，不整頁讀） */
function hctImgFind_(sh, epino) {
  var last = sh.getLastRow(); if (last < 2) return 0;
  var c = sh.getRange(2, 1, last - 1, 1).createTextFinder(epino).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}
function hctImgGet_(epino) {
  var sh = dcoSheet_(SH_HCTIMG), row = hctImgFind_(sh, epino);
  return row ? String(sh.getRange(row, 3).getValue() || '') : '';
}
function hctImgPut_(epino, edelno, b64) {
  var sh = dcoSheet_(SH_HCTIMG), row = hctImgFind_(sh, epino), vals = [[epino, edelno, b64, new Date()]];
  if (row) sh.getRange(row, 1, 1, H_HCTIMG.length).setValues(vals);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, H_HCTIMG.length).setValues(vals);
  /* 清掉超過 HCT_IMG_DAYS 天的圖（舊的在上面；一次最多看前 200 列） */
  var last = sh.getLastRow(); if (last < 3) return;
  var n = Math.min(200, last - 2), ts = sh.getRange(2, 4, n, 1).getValues(), cut = Date.now() - HCT_IMG_DAYS * 86400000, k = 0;
  while (k < n && ts[k][0] instanceof Date && ts[k][0].getTime() < cut) k++;
  if (k) sh.deleteRows(2, k);
}
function hctRow_(r, img) {
  var mode = dcoStr_(r['計費']) || 'W';
  var o = { id: dcoStr_(r['訂單號']), store: dcoStr_(r['店號']), epino: dcoStr_(r['託運單號']), edelno: dcoStr_(r['新竹貨號']), erstno: dcoStr_(r['到著站']),
    boxes: dcoNum_(r['件數']), weight: dcoNum_(r['重量']), date: dcoStr_(r['出貨日']), status: dcoStr_(r['狀態']), name: dcoStr_(r['收件人']),
    tel: dcoStr_(r['電話']), addr: dcoStr_(r['地址']), env: dcoStr_(r['環境']), created: dcoFmt_(r['建立時間']), updated: dcoFmt_(r['更新時間']), reported: dcoFmt_(r['交寄時間']),
    cs: dcoStr_(r['客代']), mode: mode, kind: dcoStr_(r['商品種類']) || '001', box: dcoNum_(r['箱序']) || 1, dispatch: dcoStr_(r['派遣']), dispTime: dcoFmt_(r['派遣時間']) };
  if (img) o.img = hctImgGet_(o.epino);
  return o;
}
/* 讀托運單：date＝出貨日 yyyyMMdd；ids＝訂單號（逗號分隔）；days＝近 N 天（預設 7）；img=1 才帶標籤圖 */
function hctList_(p) {
  var rows = hctTable_().rows.filter(function (r) { return dcoStr_(r['訂單號']); });
  var ids = p.ids ? String(p.ids).split(',').map(dcoStr_) : null;
  var since = Utilities.formatDate(new Date(Date.now() - (dcoNum_(p.days) || 7) * 86400000), DCO_TZ, 'yyyyMMdd');
  return rows.filter(function (r) {
    if (ids) return ids.indexOf(dcoStr_(r['訂單號'])) >= 0;
    if (p.date) return dcoStr_(r['出貨日']) === String(p.date);
    return dcoStr_(r['出貨日']) >= since;
  }).map(function (r) { return hctRow_(r, p.img === '1'); });
}
/* 這張訂單目前有效的托運單（已建單／已交寄、同環境），依箱序排 */
function hctActiveAll_(t, id, env) {
  return t.rows.filter(function (r) { return dcoStr_(r['訂單號']) === id && HCT_OK[dcoStr_(r['狀態'])] && (!env || (dcoStr_(r['環境']) || env) === env); })
    .sort(function (a, b) { return (dcoNum_(a['箱序']) || 1) - (dcoNum_(b['箱序']) || 1); });
}
/* v1.8：把這張訂單「已建單（還沒交寄）」的托運單改作廢；回作廢幾筆。dc_hct 還沒建過就不動 */
function hctVoidOpen_(id, now) {
  try {
    if (!dcoSS_().getSheetByName(SH_HCT)) return 0;
    var t = hctTable_(), n = 0;
    t.rows.forEach(function (r) {
      if (dcoStr_(r['訂單號']) === id && dcoStr_(r['狀態']) === '已建單') { dcoSet_(t, r._row, '狀態', '作廢'); dcoSet_(t, r._row, '更新時間', now || new Date()); n++; }
    });
    return n;
  } catch (e) { return 0; }
}
function hctSetOrderNo_(id, no) {
  var t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (rs.length) dcoWriteRows_(t, rs, function () { return { '新竹貨號': no }; });   /* 不動「更新時間」：前端編輯的 base 不受影響 */
}
function hctPad2_(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function hctEpino_(id, seq) { return String(id).replace(/[^A-Za-z0-9]/g, '') + hctPad2_(seq); }   /* 只能英文和數字 */

/* 產生（或同日更正）托運單：{id, cs, boxes, weight, sizes:[60,90…], rcv:{name, tel, addr}}
   W（重量）：v1.9 起也一箱一筆（總重量平均分）；S（總長）：一箱一筆，sizes 每箱尺寸。
   同一天、同客代、還沒交寄 → 第 i 箱沿用原託運單號＋貨號（新竹當更正）；多的新開、少的作廢。換客代或隔天 → 舊的全部作廢重開。 */
function hctCreate_(b) {
  var id = dcoStr_(b.id), cfg = hctCfg_(), cs = hctCsOf_(cfg, b.cs) || cfg.cs[0];
  var rcv = b.rcv || {}, name = hctClean_(rcv.name, 40), tel = dcoStr_(rcv.tel).replace(/[^\d#]/g, '').slice(0, 20), addr = hctClean_(rcv.addr, 100);
  if (!name || !tel || !addr) return { ok: false, error: '收件人、電話、地址都要有' };
  var items = [];
  if (cs.mode === 'S') {
    var sizes = (b.sizes || []).map(function (x) { return Math.ceil(dcoNum_(x)); });
    if (!sizes.length || sizes.length > 20) return { ok: false, error: '按尺寸計費：每箱都要選尺寸（1～20 箱）' };
    for (var i = 0; i < sizes.length; i++) if (!(HCT_SIZES[sizes[i]] || (sizes[i] >= 151 && sizes[i] <= 999))) return { ok: false, error: '第 ' + (i + 1) + ' 箱尺寸要是 60、90、120、150 或 151～999 公分' };
    items = sizes.map(function (s) { return { ejamt: 1, eqamt: s }; });
  } else {
    var boxes = Math.floor(dcoNum_(b.boxes)), weight = Math.ceil(dcoNum_(b.weight));
    if (!(boxes >= 1 && boxes <= 20)) return { ok: false, error: '件數要填 1～20 箱' };
    if (!(weight >= 1 && weight <= 99999)) return { ok: false, error: '重量要填 1 公斤以上（小數會進位）' };
    /* v1.9（2026-11-12）：新竹司機掃描機不會跳多件 → 改一箱一張托運單（件數 1），總重量平均分到每箱（整數、每箱至少 1 公斤） */
    for (var bi = 0; bi < boxes; bi++) items.push({ ejamt: 1, eqamt: Math.max(1, Math.floor(weight / boxes) + (bi < weight % boxes ? 1 : 0)) });
  }
  var tl = dcoRead_(SH_LINE), rs = dcoRowsOf_(tl, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  if (dcoStr_(rs[0]['狀態']) === '已取消') return { ok: false, error: '已取消的訂單不能開托運單' };
  var store = dcoStr_(rs[0]['店號']);
  var t = hctTable_(), today = hctToday_(), act = hctActiveAll_(t, id, cfg.env), now = new Date();
  for (var j = 0; j < act.length; j++) if (dcoStr_(act[j]['狀態']) === '已交寄') return { ok: false, error: '這張訂單 ' + dcoStr_(act[j]['出貨日']) + ' 已交寄（貨號 ' + dcoStr_(act[j]['新竹貨號']) + '），不能再開；要重寄請先聯絡新竹站所' };
  var fix = act.length > 0 && act.every(function (r) { return dcoStr_(r['出貨日']) === today && (dcoStr_(r['客代']) || cs.code) === cs.code; });
  var voids = fix ? act.slice(items.length) : act;
  var seq = 0;
  t.rows.forEach(function (r) { if (dcoStr_(r['訂單號']) === id) seq++; });
  var base = { esdate: today, eprdcl2: cs.kind || '001', emark: hctClean_('訂單 ' + id, 100) };
  if (cs.code) base.escsno = cs.code;
  if (cfg.esstno) base.esstno = cfg.esstno;
  if (cfg.sender) base.etcsig = hctClean_(cfg.sender, 40);
  if (cfg.senderTel) base.ettel1 = String(cfg.senderTel).replace(/[^\d#]/g, '').slice(0, 15);
  if (cfg.senderAddr) base.etaddr = hctClean_(cfg.senderAddr, 100);
  var recs = items.map(function (it, k) {
    var o = { ercsig: name, ertel1: tel, eraddr: addr, ejamt: String(it.ejamt), eqamt: String(it.eqamt) };
    Object.keys(base).forEach(function (x) { o[x] = base[x]; });
    var old = fix ? act[k] : null;
    if (old) { o.epino = dcoStr_(old['託運單號']); o.edelno = dcoStr_(old['新竹貨號']); }
    else o.epino = hctEpino_(id, ++seq);
    return { rec: o, old: old, box: k + 1 };
  });
  /* 送新竹：要標籤圖一批最多 5 筆 */
  var res = [];
  for (var s = 0; s < recs.length; s += 5) {
    var part = recs.slice(s, s + 5), out;
    /* v1.8.1：後面一批沒回應（逾時／504）時，前面已經成功的照樣記進表（原本整個丟例外，新竹那邊留下沒人知道的貨號） */
    try { out = hctArr_(hctCall_('TransData_Json', part.map(function (x) { return x.rec; }))); }
    catch (e) { for (var q = s; q < recs.length; q++) res.push({ ErrMsg: '新竹沒有回應（' + String(e && e.message || e).slice(0, 80) + '），這箱請再按一次' }); break; }
    part.forEach(function (x, i) { res.push(out[i] || {}); });
  }
  /* 「訂單編號重複」（新竹那邊今天已有這個號碼，例：試算表被手動刪列）→ 換下一個序號單筆再送，最多 9 次 */
  recs.forEach(function (x, i) {
    for (var tries = 0; !x.old && tries < 9 && res[i].success !== 'Y' && /訂單編號重複/.test(String(res[i].ErrMsg || '')); tries++) {
      x.rec.epino = hctEpino_(id, ++seq);
      try { res[i] = hctArr_(hctCall_('TransData_Json', [x.rec]))[0] || {}; } catch (e) { res[i] = { ErrMsg: '新竹沒有回應（' + String(e && e.message || e).slice(0, 80) + '），這箱請再按一次' }; break; }
    }
  });
  var labels = [], fails = [];
  recs.forEach(function (x, i) {
    var r = res[i];
    if (r.success !== 'Y' && r.success !== 'R') { fails.push('第 ' + x.box + ' 箱：' + (r.ErrMsg || '原因不明')); return; }
    var img = hctHexToB64_(r.image), no = dcoStr_(r.edelno), stn = dcoStr_(r.erstno);
    var vals = { '訂單號': id, '店號': store, '託運單號': x.rec.epino, '新竹貨號': no, '到著站': stn, '件數': Number(x.rec.ejamt), '重量': Number(x.rec.eqamt),
      '出貨日': today, '狀態': '已建單', '收件人': name, '電話': tel, '地址': addr, '環境': cfg.env, '建立時間': x.old ? x.old['建立時間'] : now, '更新時間': now, '交寄時間': '',
      '客代': cs.code, '計費': cs.mode, '商品種類': cs.kind || '001', '箱序': x.box, '派遣': '', '派遣時間': '' };
    if (x.old) t.sh.getRange(x.old._row, 1, 1, t.h.length).setValues([hctRowVals_(t, vals)]);
    else t.sh.getRange(t.sh.getLastRow() + 1, 1, 1, t.h.length).setValues([hctRowVals_(t, vals)]);
    if (img) hctImgPut_(x.rec.epino, no, img);
    labels.push({ id: id, store: store, epino: x.rec.epino, edelno: no, erstno: stn, boxes: Number(x.rec.ejamt), weight: Number(x.rec.eqamt), date: today, status: '已建單',
      env: cfg.env, cs: cs.code, mode: cs.mode, kind: cs.kind || '001', box: x.box, img: img, fix: !!x.old });
  });
  if (!labels.length) {   /* v1.8：新竹一箱都沒接受 → 舊托運單、訂單貨號都不動 */
    dcoLog_('hctCreate', id, store, { cs: cs.code, mode: cs.mode, items: items, env: cfg.env, fix: fix, ok: [], fails: fails, voids: 0 });
    return { ok: false, error: '新竹系統沒有接受：' + fails.join('；') };
  }
  voids.forEach(function (r) { dcoSet_(t, r._row, '狀態', '作廢'); dcoSet_(t, r._row, '更新時間', now); });
  var nos = hctActiveAll_(hctTable_(), id, cfg.env).map(function (r) { return dcoStr_(r['新竹貨號']); }).filter(Boolean);
  hctSetOrderNo_(id, nos.join(','));
  dcoLog_('hctCreate', id, store, { cs: cs.code, mode: cs.mode, items: items, env: cfg.env, fix: fix, ok: labels.map(function (l) { return l.epino + '/' + l.edelno; }), fails: fails, voids: voids.length });
  if (!labels.length) return { ok: false, error: '新竹系統沒有接受：' + fails.join('；') };
  return { ok: !fails.length, error: fails.length ? '有 ' + fails.length + ' 箱沒成功：' + fails.join('；') : '', id: id, env: cfg.env, fix: fix, labels: labels,
    edelno: labels[0].edelno, label: labels[0] };
}

/* 作廢：{id}。這張訂單還沒交寄的托運單全部作廢（新竹沒有取消託運的 API；不確認出貨、不交給司機就不會寄出） */
function hctVoid_(b) {
  var id = dcoStr_(b.id), t = hctTable_(), act = hctActiveAll_(t, id, ''), now = new Date();
  if (!act.length) return { ok: true, id: id, already: true };
  if (act.some(function (r) { return dcoStr_(r['狀態']) === '已交寄'; })) return { ok: false, error: '已交寄的托運單不能作廢，請聯絡新竹站所' };
  act.forEach(function (r) { dcoSet_(t, r._row, '狀態', '作廢'); dcoSet_(t, r._row, '更新時間', now); });
  hctSetOrderNo_(id, '');
  dcoLog_('hctVoid', id, act[0]['店號'], { edelno: act.map(function (r) { return dcoStr_(r['新竹貨號']); }) });
  return { ok: true, id: id, n: act.length };
}

/* 派遣（叫新竹來收件）：rows＝dc_hct 列。回 {done, fail:[…], msg} */
function hctDispatchRows_(t, rows, cfg) {
  var done = 0, fail = [], msg = '', now = new Date();
  for (var i = 0; i < rows.length; i += 20) {
    var part = rows.slice(i, i + 20);
    var r = hctCall_('TransDispatchOrder_Json', part.map(function (x) { return { epino: dcoStr_(x['託運單號']), edelno: dcoStr_(x['新竹貨號']), emark: cfg.dispatchMark }; })) || {};
    var data = hctArr_(r.rtn_data || r), by = {};
    data.forEach(function (d) { by[dcoStr_(d.edelno)] = d; });
    if (r.rtn_msg) msg = String(r.rtn_msg);
    part.forEach(function (x) {
      var d = by[dcoStr_(x['新竹貨號'])] || {}, ok = String(d.success || '').toUpperCase() === 'Y';
      if (ok) { dcoSet_(t, x._row, '派遣', 'Y'); dcoSet_(t, x._row, '派遣時間', now); done++; }
      else { var e = d.errMsg || d.ErrMsg || r.rtn_msg || '新竹沒有回這一筆'; dcoSet_(t, x._row, '派遣', ('N ' + e).slice(0, 100)); fail.push({ id: dcoStr_(x['訂單號']), edelno: dcoStr_(x['新竹貨號']), err: e }); }
    });
  }
  return { done: done, fail: fail, msg: msg, code: '' };
}

/* 確認出貨（印託運總表）：{date?=今天}。當天「已建單」送新竹 TransReport（一批 20），成功改「已交寄」；新竹回「已經上傳過此筆」也算成功。
   有開 HCT_DISPATCH＝Y 就接著呼叫派遣。回傳當天全部有效托運單（不含圖）給前端印總表 */
function hctReport_(b) {
  var date = dcoStr_(b.date) || hctToday_(), t = hctTable_(), cfg = hctCfg_(), now = new Date();
  var dead = {};   /* v1.8：訂單已取消（含被合併掉）的托運單不送新竹，改作廢 */
  dcoRead_(SH_LINE).rows.forEach(function (r) { if (dcoStr_(r['狀態']) === '已取消') dead[dcoStr_(r['訂單號'])] = 1; });
  var skipped = [];
  var todo = t.rows.filter(function (r) {
    if (!(dcoStr_(r['出貨日']) === date && dcoStr_(r['狀態']) === '已建單' && dcoStr_(r['環境']) === cfg.env)) return false;
    if (dead[dcoStr_(r['訂單號'])]) { dcoSet_(t, r._row, '狀態', '作廢'); dcoSet_(t, r._row, '更新時間', now); r['狀態'] = '作廢'; skipped.push(dcoStr_(r['訂單號']) + '/' + dcoStr_(r['新竹貨號'])); return false; }
    return true;
  });
  var fail = [], done = 0, okRows = [];
  for (var i = 0; i < todo.length; i += 20) {
    var part = todo.slice(i, i + 20);
    var res = hctArr_(hctCall_('TransReport_Json', part.map(function (r) { return { epino: dcoStr_(r['託運單號']), edelno: dcoStr_(r['新竹貨號']) }; })));
    var by = {};
    res.forEach(function (x) { by[dcoStr_(x.edelno)] = x; });
    part.forEach(function (r) {
      var x = by[dcoStr_(r['新竹貨號'])] || {};
      if (x.success === 'Y' || /上傳過/.test(String(x.ErrMsg || ''))) { dcoSet_(t, r._row, '狀態', '已交寄'); dcoSet_(t, r._row, '交寄時間', now); dcoSet_(t, r._row, '更新時間', now); done++; okRows.push(r); }
      else fail.push({ id: dcoStr_(r['訂單號']), edelno: dcoStr_(r['新竹貨號']), err: x.ErrMsg || '新竹沒有回這一筆' });
    });
  }
  var disp = null;
  if (cfg.dispatch && okRows.length) { try { disp = hctDispatchRows_(t, okRows, cfg); } catch (e) { disp = { done: 0, fail: [], msg: String(e && e.message || e) }; } }
  dcoLog_('hctReport', '', '', { date: date, done: done, fail: fail, env: cfg.env, dispatch: disp, skipped: skipped });
  var list = hctTable_().rows.filter(function (r) { return dcoStr_(r['出貨日']) === date && HCT_OK[dcoStr_(r['狀態'])] && dcoStr_(r['環境']) === cfg.env; }).map(function (r) { return hctRow_(r, false); });
  return { ok: true, date: date, done: done, fail: fail, env: cfg.env, cs: cfg.cs, companyName: cfg.companyName, dispatch: disp, labels: list };
}

/* 呼叫派遣（手動補叫）：{date?=今天}。當天已交寄、還沒派遣成功的 */
function hctDispatch_(b) {
  var date = dcoStr_(b.date) || hctToday_(), t = hctTable_(), cfg = hctCfg_();
  var rows = t.rows.filter(function (r) { return dcoStr_(r['出貨日']) === date && dcoStr_(r['狀態']) === '已交寄' && dcoStr_(r['環境']) === cfg.env && dcoStr_(r['派遣']) !== 'Y'; });
  if (!rows.length) return { ok: true, done: 0, fail: [], msg: '沒有要派遣的托運單' };
  var d = hctDispatchRows_(t, rows, cfg);
  dcoLog_('hctDispatch', '', '', { date: date, dispatch: d, env: cfg.env });
  return { ok: true, done: d.done, fail: d.fail, msg: d.msg };
}

/* 取消派遣：{id}。這張訂單已派遣的托運單（司機還沒收件才有效；取消後同一個貨號不能再派遣） */
function hctDispatchCancel_(b) {
  var id = dcoStr_(b.id), t = hctTable_(), cfg = hctCfg_();
  var rows = hctActiveAll_(t, id, cfg.env).filter(function (r) { return dcoStr_(r['派遣']) === 'Y'; });
  if (!rows.length) return { ok: false, error: '這張訂單沒有已派遣的托運單' };
  var res = hctArr_(hctCall_('TransDataCancel_Json', rows.map(function (r) { return { epino: dcoStr_(r['託運單號']), edelno: dcoStr_(r['新竹貨號']) }; }))), by = {}, fail = [];
  res.forEach(function (x) { by[dcoStr_(x.edelno)] = x; });
  rows.forEach(function (r) {
    var x = by[dcoStr_(r['新竹貨號'])] || {};
    if (String(x.success || '').toUpperCase() === 'Y') dcoSet_(t, r._row, '派遣', '取消');
    else fail.push(dcoStr_(r['新竹貨號']) + '：' + (x.ErrMsg || '新竹沒有回這一筆'));
  });
  dcoLog_('hctDispatchCancel', id, rows[0]['店號'], { fail: fail });
  return fail.length ? { ok: false, error: '取消派遣沒成功：' + fail.join('；') } : { ok: true, id: id, n: rows.length };
}

/* 向新竹查這張訂單的託運單號有沒有貨號（QueryEDELNO）：{id}。只讀新竹、不寫表 */
function hctQuery_(b) {
  var id = dcoStr_(b.id), cfg = hctCfg_(), t = hctTable_();
  var act = hctActiveAll_(t, id, cfg.env);
  var eps = (act.length ? act : t.rows.filter(function (r) { return dcoStr_(r['訂單號']) === id && (dcoStr_(r['環境']) || cfg.env) === cfg.env; })).map(function (r) { return dcoStr_(r['託運單號']); }).filter(Boolean);
  if (!eps.length) return { ok: false, error: '這張訂單沒有託運單號' };
  /* 2026-10-28 實測：一次送多筆，新竹只回第一筆 → 一筆一筆查 */
  var res = eps.slice(0, 20).map(function (e) { return hctArr_(hctCall_('QueryEDELNO_Json', [{ epino: e }]))[0] || { epino: e, success: 'N', ErrMsg: '新竹沒有回覆' }; });
  return { ok: true, id: id, result: res.map(function (x, i) { return { epino: dcoStr_(x.epino) || eps[i], edelno: dcoStr_(x.edelno), success: dcoStr_(x.success), err: dcoStr_(x.ErrMsg) }; }) };
}

/* 授權用（加了 UrlFetchApp 要經營者按一次「執行」授權）：只連新竹看服務在不在，不寫任何資料 */
function aaaHctAuthCheck() {
  var res = UrlFetchApp.fetch(hctCfg_().url + '?op=TransData_Json', { muteHttpExceptions: true });
  Logger.log('新竹服務回應 ' + res.getResponseCode() + '；目前環境：' + hctCfg_().env);
}
