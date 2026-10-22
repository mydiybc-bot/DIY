/**
 * diybc-dcorder-api  v1.3（2026-10-21 新竹物流托運單）
 * 門市 → 出貨中心 訂單 API（取代 Shopline 門市叫貨）
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
 * 讀：GET  ?action=ping｜orders｜stock｜locs &token=…（可加 callback= 走 JSONP）
 * 寫：POST text/plain JSON {token, action:create|cancel|status|ship|stockSet|stockAdj|locSet|edit|merge, …}
 *
 * 部署：部署 → 新增部署作業 → 網頁應用程式；執行身分＝我；存取權＝所有人。
 *       之後改程式一律「管理部署作業 → 編輯 → 新版本」，網址不變。
 */
var DCO_TOKEN = 'dbc-dco-Rw8pZ3';
var DCO_VER = 'v1.4-id';
var DCO_TZ = 'Asia/Taipei';
var SH_LINE = 'dc_order_line', SH_STOCK = 'dc_stock', SH_LOG = 'dc_log', SH_LOC = 'dc_loc';
var H_LINE = ['訂單號', '店號', '下單時間', '狀態', '行號', 'sku_id', '品名', '單位', '訂購量', '單價', '實出量', '小計', '備註', '出貨時間', '新竹貨號', 'cid', '更新時間'];
var H_STOCK = ['sku_id', '品名', '管理方式', '庫存', '更新時間', '備註'];
var H_LOG = ['時間', '動作', '訂單號', '店號', '內容'];
var H_LOC = ['sku_id', '品名', '位置', '更新時間'];   /* v1.2：出貨中心品項位置（一個品項一列；位置清空＝刪列） */
var SH_HCT = 'dc_hct';   /* v1.3：新竹托運單 */
var H_HCT = ['訂單號', '店號', '託運單號', '新竹貨號', '到著站', '件數', '重量', '出貨日', '狀態', '收件人', '電話', '地址', '環境', '建立時間', '更新時間', '交寄時間'];
var SH_HCTIMG = 'dc_hct_img';   /* v1.3：標籤圖另放一頁（一張約 15,000 字），平常讀托運單不用讀圖；只留 HCT_IMG_DAYS 天 */
var H_HCTIMG = ['託運單號', '新竹貨號', '標籤', '建立時間'];
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
    if (p.action === 'hct') return dcoOut_({ ok: true, env: hctCfg_().env, labels: hctList_(p) }, cb);   /* v1.3 */
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
    else if (a === 'hctCreate') r = hctCreate_(body);   /* v1.3 */
    else if (a === 'hctVoid') r = hctVoid_(body);       /* v1.3 */
    else if (a === 'hctReport') r = hctReport_(body);   /* v1.3 */
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
  dcoLog_('cancel', id, rs[0]['店號'], { by: b.by || '' });
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
    if (mode === '無限量') { if (idx[sk]) del.push(idx[sk]); n++; return; }
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
  var hct = b.hasOwnProperty('hct') ? dcoStr_(b.hct).slice(0, 40) : dcoStr_(rs[0]['新竹貨號']);
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
  });
  dcoDelRows_(t, by[into]);
  t.sh.getRange(t.sh.getLastRow() + 1, 1, out.length, H_LINE.length).setValues(out);
  dcoLog_('merge', into, store, { merged: others, n: out.length, total: total, status: status });
  return { ok: true, id: into, merged: others, n: out.length, total: total, status: status, updated: dcoFmt_(now) };
}

/* ── v1.3：新竹物流托運單 ──────────────────────────────────────
   新竹 EDI Web Service（SOAP）：https://hctrt.hct.com.tw/EDI_WebService2/Service1.asmx
   用 *_Json 版本；各方法參數名不同（打錯只會回「資料錯誤請確認」）：TransData_Json＝company／password／json，
   TransReport_Json＝sCompany／sPassword／dsCusJson（2026-10-21 照新竹服務頁實測）。
   一次只送一張訂單（回傳含標籤圖，新竹上限一批 5 張）；標籤圖＝PNG 的 16 進位字串，1000×1000（10×10 公分標籤）。
   託運單號（epino）＝訂單號-序號：同一張訂單作廢重開會換序號，避免「訂單編號重複」。 */
var HCT_URL = 'https://hctrt.hct.com.tw/EDI_WebService2/Service1.asmx';
var HCT_OK = { '已建單': 1, '已交寄': 1 };
var HCT_ARGS = { TransData_Json: ['company', 'password', 'json'], TransReport_Json: ['sCompany', 'sPassword', 'dsCusJson'] };

function hctCfg_() {
  var p = PropertiesService.getScriptProperties();
  var c = p.getProperty('HCT_COMPANY'), w = p.getProperty('HCT_PWD');
  var live = !!(c && w);
  return {
    env: live ? 'live' : 'test', company: live ? c : 'test', pwd: live ? w : 'test1',
    escsno: p.getProperty('HCT_ESCSNO') || '', esstno: p.getProperty('HCT_ESSTNO') || '',
    sender: p.getProperty('HCT_SENDER') || '', senderTel: p.getProperty('HCT_SENDER_TEL') || '', senderAddr: p.getProperty('HCT_SENDER_ADDR') || ''
  };
}
/* 呼叫新竹 SOAP：method＝TransData_Json／TransReport_Json…，回傳解析後的 JSON 陣列 */
function hctCall_(method, arr) {
  var cfg = hctCfg_();
  function esc(x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  var a = HCT_ARGS[method];
  function tag(k, v) { return '<' + k + '>' + esc(v) + '</' + k + '>'; }
  var env = '<?xml version="1.0" encoding="utf-8"?><soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
    + '<soap:Body><' + method + ' xmlns="http://tempuri.org/">' + tag(a[0], cfg.company) + tag(a[1], cfg.pwd) + tag(a[2], JSON.stringify(arr)) + '</' + method + '></soap:Body></soap:Envelope>';
  var res = UrlFetchApp.fetch(HCT_URL, { method: 'post', contentType: 'text/xml; charset=utf-8', headers: { SOAPAction: '"http://tempuri.org/' + method + '"' }, payload: env, muteHttpExceptions: true });
  var code = res.getResponseCode(), txt = res.getContentText('UTF-8');
  if (code !== 200) throw new Error('新竹系統回應 ' + code + '：' + txt.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 200));
  var m = txt.match(new RegExp('<' + method + 'Result>([\\s\\S]*?)</' + method + 'Result>'));
  if (!m) throw new Error('看不懂新竹系統的回應');
  var body = m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#xD;/g, '\r').replace(/&#xA;/g, '\n').replace(/&amp;/g, '&');
  var out = JSON.parse(body);
  return Array.isArray(out) ? out : [out];
}
function hctHexToB64_(hex) {
  hex = String(hex || '').replace(/\s+/g, '');
  if (!hex || hex.length % 2) return '';
  var bytes = new Array(hex.length / 2);
  for (var i = 0; i < bytes.length; i++) { var v = parseInt(hex.substr(i * 2, 2), 16); bytes[i] = v > 127 ? v - 256 : v; }
  return Utilities.base64Encode(bytes);
}
function hctToday_() { return Utilities.formatDate(new Date(), DCO_TZ, 'yyyyMMdd'); }
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
  var o = { id: dcoStr_(r['訂單號']), store: dcoStr_(r['店號']), epino: dcoStr_(r['託運單號']), edelno: dcoStr_(r['新竹貨號']), erstno: dcoStr_(r['到著站']),
    boxes: dcoNum_(r['件數']), weight: dcoNum_(r['重量']), date: dcoStr_(r['出貨日']), status: dcoStr_(r['狀態']), name: dcoStr_(r['收件人']),
    tel: dcoStr_(r['電話']), addr: dcoStr_(r['地址']), env: dcoStr_(r['環境']), created: dcoFmt_(r['建立時間']), updated: dcoFmt_(r['更新時間']), reported: dcoFmt_(r['交寄時間']) };
  if (img) o.img = hctImgGet_(o.epino);
  return o;
}
/* 讀托運單：date＝出貨日 yyyyMMdd；ids＝訂單號（逗號分隔）；days＝近 N 天（預設 7）；img=1 才帶標籤圖 */
function hctList_(p) {
  var rows = dcoRead_(SH_HCT).rows.filter(function (r) { return dcoStr_(r['訂單號']); });
  var ids = p.ids ? String(p.ids).split(',').map(dcoStr_) : null;
  var since = Utilities.formatDate(new Date(Date.now() - (dcoNum_(p.days) || 7) * 86400000), DCO_TZ, 'yyyyMMdd');
  return rows.filter(function (r) {
    if (ids) return ids.indexOf(dcoStr_(r['訂單號'])) >= 0;
    if (p.date) return dcoStr_(r['出貨日']) === String(p.date);
    return dcoStr_(r['出貨日']) >= since;
  }).map(function (r) { return hctRow_(r, p.img === '1'); });
}
/* 這張訂單目前有效的托運單（已建單／已交寄，取最後一張） */
function hctActive_(t, id) {
  var hit = null;
  t.rows.forEach(function (r) { if (dcoStr_(r['訂單號']) === id && HCT_OK[dcoStr_(r['狀態'])]) hit = r; });
  return hit;
}
function hctSetOrderNo_(id, no) {
  var t = dcoRead_(SH_LINE), rs = dcoRowsOf_(t, id);
  if (rs.length) dcoWriteRows_(t, rs, function () { return { '新竹貨號': no }; });   /* 不動「更新時間」：前端編輯的 base 不受影響 */
}

/* 產生托運單：{id, boxes, weight, rcv:{name, tel, addr}} */
function hctCreate_(b) {
  var id = dcoStr_(b.id), boxes = Math.floor(dcoNum_(b.boxes)), weight = Math.ceil(dcoNum_(b.weight));
  var rcv = b.rcv || {}, name = dcoStr_(rcv.name).slice(0, 40), tel = dcoStr_(rcv.tel).replace(/[^\d#-]/g, '').slice(0, 15), addr = dcoStr_(rcv.addr).slice(0, 100);
  if (!(boxes >= 1 && boxes <= 999)) return { ok: false, error: '件數要填 1～999' };
  if (!(weight >= 1 && weight <= 99999)) return { ok: false, error: '重量要填 1 公斤以上（小數會進位）' };
  if (!name || !tel || !addr) return { ok: false, error: '收件人、電話、地址都要有' };
  var tl = dcoRead_(SH_LINE), rs = dcoRowsOf_(tl, id);
  if (!rs.length) return { ok: false, error: '找不到訂單' };
  var st = dcoStr_(rs[0]['狀態']);
  if (st === '已取消') return { ok: false, error: '已取消的訂單不能開托運單' };
  var t = dcoRead_(SH_HCT), cfg = hctCfg_(), today = hctToday_(), act = hctActive_(t, id), now = new Date();
  var fix = act && dcoStr_(act['出貨日']) === today && dcoStr_(act['狀態']) === '已建單' && dcoStr_(act['環境']) === cfg.env;   /* 同一天、還沒交寄＝更正 */
  if (act && !fix) {
    if (dcoStr_(act['狀態']) === '已交寄') return { ok: false, error: '這張訂單 ' + dcoStr_(act['出貨日']) + ' 已交寄（貨號 ' + dcoStr_(act['新竹貨號']) + '），不能再開；要重寄請先聯絡新竹站所' };
    dcoSet_(t, act._row, '狀態', '作廢'); dcoSet_(t, act._row, '更新時間', now);   /* 前一天開了沒交寄（或換了環境）→ 作廢舊的、開新的 */
  }
  var seq = 1;
  t.rows.forEach(function (r) { if (dcoStr_(r['訂單號']) === id) seq++; });
  var epino = fix ? dcoStr_(act['託運單號']) : id + '-' + seq;
  var rec = { epino: epino, ercsig: name, ertel1: tel, eraddr: addr, ejamt: String(boxes), eqamt: String(weight), esdate: today, emark: ('訂單 ' + id).slice(0, 100) };
  if (fix) rec.edelno = dcoStr_(act['新竹貨號']);
  if (cfg.escsno) rec.escsno = cfg.escsno;
  if (cfg.esstno) rec.esstno = cfg.esstno;
  if (cfg.sender) rec.etcsig = cfg.sender;
  if (cfg.senderTel) rec.ettel1 = cfg.senderTel;
  if (cfg.senderAddr) rec.etaddr = cfg.senderAddr;
  var r = hctCall_('TransData_Json', [rec])[0] || {};
  for (var tries = 0; !fix && tries < 9 && r.success !== 'Y' && /訂單編號重複/.test(String(r.ErrMsg || '')); tries++) {   /* 新竹那邊今天已有這個託運單號（例：試算表被手動刪列）→ 換下一個序號再送 */
    seq++; epino = rec.epino = id + '-' + seq;
    r = hctCall_('TransData_Json', [rec])[0] || {};
  }
  if (r.success !== 'Y' && r.success !== 'R') {
    dcoLog_('hctFail', id, rs[0]['店號'], { rec: rec, err: r.ErrMsg || '' });
    return { ok: false, error: '新竹系統沒有接受：' + (r.ErrMsg || '原因不明') };
  }
  var img = hctHexToB64_(r.image), no = dcoStr_(r.edelno), stn = dcoStr_(r.erstno);
  var row = [id, dcoStr_(rs[0]['店號']), epino, no, stn, boxes, weight, today, '已建單', name, tel, addr, cfg.env, fix ? act['建立時間'] : now, now, ''];
  if (fix) t.sh.getRange(act._row, 1, 1, H_HCT.length).setValues([row]);
  else t.sh.getRange(t.sh.getLastRow() + 1, 1, 1, H_HCT.length).setValues([row]);
  if (img) hctImgPut_(epino, no, img);
  hctSetOrderNo_(id, no);
  dcoLog_('hctCreate', id, rs[0]['店號'], { epino: epino, edelno: no, erstno: stn, boxes: boxes, weight: weight, env: cfg.env, fix: !!fix, img: img.length });
  return { ok: true, id: id, edelno: no, erstno: stn, epino: epino, fix: !!fix, env: cfg.env, label: { id: id, store: dcoStr_(rs[0]['店號']), epino: epino, edelno: no, erstno: stn, boxes: boxes, weight: weight, date: today, status: '已建單', env: cfg.env, img: img } };
}

/* 作廢：{id}。只能作廢還沒交寄的（新竹沒有取消 API；不放進總表、不交給司機就不會寄出） */
function hctVoid_(b) {
  var id = dcoStr_(b.id), t = dcoRead_(SH_HCT), act = hctActive_(t, id);
  if (!act) return { ok: true, id: id, already: true };
  if (dcoStr_(act['狀態']) === '已交寄') return { ok: false, error: '已交寄的托運單不能作廢，請聯絡新竹站所' };
  dcoSet_(t, act._row, '狀態', '作廢'); dcoSet_(t, act._row, '更新時間', new Date());
  hctSetOrderNo_(id, '');
  dcoLog_('hctVoid', id, act['店號'], { edelno: dcoStr_(act['新竹貨號']) });
  return { ok: true, id: id };
}

/* 確認出貨（印託運總表）：{date?=今天}。把當天「已建單」的送新竹 TransReport，成功改「已交寄」。
   新竹回「已經上傳過此筆」也算成功。回傳當天全部有效托運單（不含圖）給前端印總表 */
function hctReport_(b) {
  var date = dcoStr_(b.date) || hctToday_(), t = dcoRead_(SH_HCT), cfg = hctCfg_(), now = new Date();
  var todo = t.rows.filter(function (r) { return dcoStr_(r['出貨日']) === date && dcoStr_(r['狀態']) === '已建單' && dcoStr_(r['環境']) === cfg.env; });
  var fail = [], done = 0;
  for (var i = 0; i < todo.length; i += 30) {
    var part = todo.slice(i, i + 30);
    var res = hctCall_('TransReport_Json', part.map(function (r) { return { epino: dcoStr_(r['託運單號']), edelno: dcoStr_(r['新竹貨號']) }; }));
    var by = {};
    res.forEach(function (x) { by[dcoStr_(x.edelno)] = x; });
    part.forEach(function (r) {
      var x = by[dcoStr_(r['新竹貨號'])] || {};
      if (x.success === 'Y' || /上傳過/.test(String(x.ErrMsg || ''))) { dcoSet_(t, r._row, '狀態', '已交寄'); dcoSet_(t, r._row, '交寄時間', now); dcoSet_(t, r._row, '更新時間', now); done++; }
      else fail.push({ id: dcoStr_(r['訂單號']), edelno: dcoStr_(r['新竹貨號']), err: x.ErrMsg || '新竹沒有回這一筆' });
    });
  }
  dcoLog_('hctReport', '', '', { date: date, done: done, fail: fail, env: cfg.env });
  var list = dcoRead_(SH_HCT).rows.filter(function (r) { return dcoStr_(r['出貨日']) === date && HCT_OK[dcoStr_(r['狀態'])] && dcoStr_(r['環境']) === cfg.env; }).map(function (r) { return hctRow_(r, false); });
  return { ok: true, date: date, done: done, fail: fail, env: cfg.env, escsno: cfg.escsno, labels: list };
}

/* 授權用（加了 UrlFetchApp 要經營者按一次「執行」授權）：只連新竹看服務在不在，不寫任何資料 */
function aaaHctAuthCheck() {
  var res = UrlFetchApp.fetch(HCT_URL + '?op=TransData_Json', { muteHttpExceptions: true });
  Logger.log('新竹服務回應 ' + res.getResponseCode() + '；目前環境：' + hctCfg_().env);
}
