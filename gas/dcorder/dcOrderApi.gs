/**
 * diybc-dcorder-api  v1.0（2026-09-28 cn 批）
 * 門市 → 出貨中心 訂單 API（取代 Shopline 門市叫貨）
 *
 * 獨立 Apps Script 專案（沿用「每支寫入 API 各自一個專案」慣例，不動 diybc-purchase-agg 主程式）。
 * 資料放在獨立試算表「出貨中心訂單（採購系統）」：第一次執行 setup() 時自動建立，ID 記在指令碼屬性 DCO_SS_ID。
 *   dc_order_line  一列＝一張訂單的一個品項（訂單層欄位每列重複，月結、判讀直接加總）
 *   dc_stock       出貨中心庫存（只記「記數量」的品項；沒列在這裡＝無限量）
 *   dc_log         所有異動紀錄
 *
 * 讀：GET  ?action=ping｜orders｜stock &token=…（可加 callback= 走 JSONP）
 * 寫：POST text/plain JSON {token, action:create|cancel|status|ship|stockSet|stockAdj, …}
 *
 * 部署：部署 → 新增部署作業 → 網頁應用程式；執行身分＝我；存取權＝所有人。
 *       之後改程式一律「管理部署作業 → 編輯 → 新版本」，網址不變。
 */
var DCO_TOKEN = 'dbc-dco-Rw8pZ3';
var DCO_VER = 'v1.0';
var DCO_TZ = 'Asia/Taipei';
var SH_LINE = 'dc_order_line', SH_STOCK = 'dc_stock', SH_LOG = 'dc_log';
var H_LINE = ['訂單號', '店號', '下單時間', '狀態', '行號', 'sku_id', '品名', '單位', '訂購量', '單價', '實出量', '小計', '備註', '出貨時間', '新竹貨號', 'cid', '更新時間'];
var H_STOCK = ['sku_id', '品名', '管理方式', '庫存', '更新時間', '備註'];
var H_LOG = ['時間', '動作', '訂單號', '店號', '內容'];
var OPEN_ST = { '待處理': 1, '備貨中': 1 };

/* ── 初始化（第一次手動執行一次；會要求授權） ─────────────────── */
function setup() {
  var ss = dcoSS_();
  dcoEnsure_(ss, SH_LINE, H_LINE, ['A', 'B', 'F', 'O', 'P']);
  dcoEnsure_(ss, SH_STOCK, H_STOCK, ['A']);
  dcoEnsure_(ss, SH_LOG, H_LOG, ['C', 'D']);
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
    if (p.action === 'ping') return dcoOut_({ ok: true, ver: DCO_VER }, cb);
    if (p.token !== DCO_TOKEN) return dcoOut_({ ok: false, error: 'token' }, cb);
    if (p.action === 'orders') return dcoOut_({ ok: true, orders: dcoOrders_(p) }, cb);
    if (p.action === 'stock') return dcoOut_({ ok: true, stock: dcoStock_() }, cb);
    return dcoOut_({ ok: false, error: '未知的 action' }, cb);
  } catch (err) {
    return dcoOut_({ ok: false, error: String(err && err.message || err) }, cb);
  }
}

/* 參數：store（店號）、status（逗號分隔）、from/to（下單日 yyyy-MM-dd）、sfrom/sto（出貨日）、id、cid、days（預設 60）
   沒給任何日期條件時只回近 days 天下單的 ＋ 還沒處理完的（待處理、備貨中不論多久都回） */
function dcoOrders_(p) {
  var rows = dcoRead_(SH_LINE).rows;
  var sts = p.status ? String(p.status).split(',') : null;
  var hasDate = p.from || p.to || p.sfrom || p.sto;
  var since = new Date(Date.now() - (dcoNum_(p.days) || 60) * 86400000);
  var list = dcoGroup_(rows).filter(function (o) {
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
  return dcoRead_(SH_STOCK).rows.filter(function (r) { return dcoStr_(r['sku_id']); }).map(function (r) {
    return { sku_id: dcoStr_(r['sku_id']), name: dcoStr_(r['品名']), mode: dcoStr_(r['管理方式']) || '記數量', qty: dcoNum_(r['庫存']), time: dcoFmt_(r['更新時間']), note: dcoStr_(r['備註']) };
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
    else r = { ok: false, error: '未知的 action' };
    return dcoOut_(r);
  } catch (err) {
    return dcoOut_({ ok: false, error: String(err && err.message || err) });
  } finally {
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
  var prefix = 'DC' + Utilities.formatDate(now, DCO_TZ, 'yyMMdd') + '-' + dcoPad_(store) + '-';
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
