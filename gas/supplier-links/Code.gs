/**
 * 海外代理商購買平台 API（原名「新品供應商查詢」；2026-10-19 新建，同日改名，經營者交辦：「把檔案中每一個產品使用的廠商、供貨的廠商，還有它的連結表列出來，讓我可以搜尋」）
 * 綁在試算表「新品供應商查詢_資料庫」；前端＝static/dashboard-supplier.html。
 *
 * 資料來源：Drive 資料夾「2026」（新品研發工作申請單，一個檔期一份 Google 試算表，含子資料夾）
 *   每份檔的「新品通知表」分頁 → 「各店新增食材、器具」表（甜點名稱、新增食材/器模具、供應商、購買連結…）
 *   品項欄是公式（='01 臭豆腐哪有臭'!A18）→ 順著公式回到該甜點分頁，補「成本、提供方式、進貨廠商、甜點正式名稱」。
 *   xlsx 檔讀不到（SpreadsheetApp 只開得了 Google 試算表）→ 列在 files 分頁「略過」。
 *
 * 讀（GET，JSONP，免密碼，跟其他儀表板一樣靠頁面閘門擋路人）：
 *   ?action=ping        版號
 *   ?action=items       彙整結果（items＋files＋最後彙整時間）
 *   ?action=rebuild     立刻重新彙整（10 分鐘內只跑一次；只讀來源檔、只寫本試算表）
 * 寫（POST）：
 *   {action:'authSetup'}（免密碼）權限中心 dim_auth 補「supplier」列，已有就只把名稱改成 AUTH_NAME（只補沒有的；密碼空白＝還不能用，經營者在權限管理後台設）
 * 排程：rebuild 每天 06:00（在編輯器「觸發條件」介面建，不加 script.scriptapp 權限）。
 * 所有儲存格一律存純文字。
 */
var VER = 'v1.1';
var ROOT_FOLDER_ID = '1EVJrOIhB69jRkKDNbOpifbVq7BHENd2o';   /* 新品研發工作申請單／2026 */
var AUTH_DASH = 'supplier', AUTH_NAME = '海外代理商購買平台';   /* 2026-10-19 經營者改名（原「新品供應商查詢」） */
var AUTHC_SS_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4', AUTHC_TAB = 'dim_auth';
var REBUILD_GAP_SEC = 600;

var ITEM_COLS = ['item', 'dessert', 'campaign', 'file', 'folder', 'supplier', 'vendor_src', 'provide', 'links', 'cost', 'spec', 'usage',
  'keep', 'take', 'zone', 'pack', 'split', 'order_way', 'first_order', 'sticker', 'note', 'src_tab', 'src_row', 'file_id', 'gid'];
var FILE_COLS = ['file', 'folder', 'file_id', 'modified', 'status', 'n_items', 'n_links'];

/* 新品通知表標題 → 欄位（標題去掉空白後比對） */
var HDR = {
  '甜點名稱': 'dessert', '新增食材/器模具': 'item', '保存方式': 'keep', '原包裝規格': 'spec', '貼紙名稱': 'sticker',
  '每份使用量': 'usage', '分裝方式': 'split', '叫貨方式': 'order_way', '第一單處理': 'first_order', '拿取方式': 'take',
  '區域': 'zone', '包裝類型': 'pack', '供應商': 'supplier', '購買連結': 'links', '備註': 'note'
};
var PAIR = { spec: 1, usage: 1 };   /* 合併兩格（數字＋單位），例：300｜g */

/* ---------------- 進入點 ---------------- */
function doGet(e) {
  var p = (e && e.parameter) || {}, res;
  try {
    var a = p.action || 'ping';
    if (a === 'ping') res = { ok: true, ver: VER };
    else if (a === 'items') res = { ok: true, data: items_() };
    else if (a === 'rebuild') res = rebuildThrottled_();
    else res = { ok: false, error: '不認得的動作：' + a };
  } catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res, p.callback);
}
function doPost(e) {
  var res;
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'authSetup') res = authSetup_();
    else res = { ok: false, error: '不認得的動作：' + b.action };
  } catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res);
}
function out_(res, cb) {
  var s = JSON.stringify(res);
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(String(cb))) return ContentService.createTextOutput(cb + '(' + s + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function clean_(v) { return String(v == null ? '' : v).replace(/\r/g, '').replace(/\s*\n\s*/g, ' ').trim(); }
function key_(s) { return String(s || '').replace(/\s+/g, ''); }

/* ---------------- 編輯器用 ---------------- */
function authorize() {   /* 第一次授權用：碰一下會用到的權限，不改任何東西 */
  DriveApp.getFolderById(ROOT_FOLDER_ID).getName();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  Logger.log('授權 OK，' + VER);
}
function rebuild() {   /* 排程每天 06:00 跑這個 */
  var lk = LockService.getScriptLock();
  if (!lk.tryLock(30000)) throw new Error('正在彙整中，請稍後再試');
  try { return rebuild_(); } finally { lk.releaseLock(); }
}

/* ---------------- 讀 ---------------- */
function items_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    items: table_(ss.getSheetByName('items'), ITEM_COLS),
    files: table_(ss.getSheetByName('files'), FILE_COLS),
    rebuilt_at: PropertiesService.getScriptProperties().getProperty('rebuilt_at') || '',
    ver: VER
  };
}
function table_(sh, cols) {
  if (!sh || sh.getLastRow() < 2) return [];
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  var idx = cols.map(function (c) { return hdr.indexOf(c); });
  return sh.getRange(2, 1, sh.getLastRow() - 1, hdr.length).getDisplayValues().map(function (r) {
    var o = {}; cols.forEach(function (c, j) { o[c] = idx[j] < 0 ? '' : r[idx[j]]; }); return o;
  });
}
function rebuildThrottled_() {
  var pr = PropertiesService.getScriptProperties(), last = Number(pr.getProperty('rebuilt_ms') || 0);
  var wait = Math.ceil((last + REBUILD_GAP_SEC * 1000 - Date.now()) / 1000);
  if (wait > 0) return { ok: true, skipped: true, msg: '剛彙整過（' + (pr.getProperty('rebuilt_at') || '') + '），' + Math.ceil(wait / 60) + ' 分鐘後才能再按' };
  var r = rebuild();
  return { ok: true, msg: '彙整完成：' + r.files + ' 份檔、' + r.items + ' 列、' + r.links + ' 個有連結', stat: r };
}

/* ---------------- 彙整 ---------------- */
function rebuild_() {
  var t0 = Date.now(), files = [], rows = [], meta = [];
  walk_(DriveApp.getFolderById(ROOT_FOLDER_ID), '', files, meta);
  files.forEach(function (f) {
    var r;
    try { r = readFile_(f); } catch (err) { r = { rows: [], status: '讀取失敗：' + String(err && err.message || err).slice(0, 120) }; }
    var nl = r.rows.filter(function (x) { return x.links; }).length;
    meta.push({ file: f.name, folder: f.folder, file_id: f.id, modified: f.modified, status: r.status, n_items: String(r.rows.length), n_links: String(nl) });
    rows = rows.concat(r.rows);
  });
  meta.sort(function (a, b) { return a.folder + a.file < b.folder + b.file ? -1 : 1; });
  write_('items', ITEM_COLS, rows);
  write_('files', FILE_COLS, meta);
  var at = now_(), pr = PropertiesService.getScriptProperties();
  pr.setProperty('rebuilt_at', at); pr.setProperty('rebuilt_ms', String(Date.now()));
  var stat = { files: files.length, items: rows.length, links: rows.filter(function (x) { return x.links; }).length, sec: Math.round((Date.now() - t0) / 1000), at: at };
  log_(stat);
  return stat;
}
function walk_(folder, path, files, meta) {
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next(), mt = f.getMimeType(), mod = Utilities.formatDate(f.getLastUpdated(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    if (mt === MimeType.GOOGLE_SHEETS) files.push({ id: f.getId(), name: f.getName(), folder: path, modified: mod });
    else if (/\.xlsx?$/i.test(f.getName()) || mt === MimeType.MICROSOFT_EXCEL)
      meta.push({ file: f.getName(), folder: path, file_id: f.getId(), modified: mod, status: '略過：xlsx 讀不到，要轉成 Google 試算表才會列進來', n_items: '0', n_links: '0' });
  }
  var sub = folder.getFolders();
  while (sub.hasNext()) { var s = sub.next(); walk_(s, path ? path + '／' + s.getName() : s.getName(), files, meta); }
}
function readFile_(f) {
  var ss = SpreadsheetApp.openById(f.id);
  var sh = ss.getSheets().filter(function (s) { return s.getName().indexOf('新品通知表') === 0; })[0];
  if (!sh) return { rows: [], status: '沒有「新品通知表」分頁' };
  var last = sh.getLastRow(), lastC = Math.min(sh.getLastColumn(), 40);
  if (last < 2) return { rows: [], status: '新品通知表是空的' };
  var disp = sh.getRange(1, 1, last, lastC).getDisplayValues();
  /* 檔期名稱：A 欄「檔期：」右邊那格 */
  var campaign = '';
  for (var i = 0; i < Math.min(15, last); i++) if (/檔期/.test(disp[i][0]) && !/自己人/.test(disp[i][0])) { campaign = clean_(disp[i][1]); break; }
  /* 標題列：同一列同時有「新增食材/器模具」和「供應商」 */
  var h = -1;
  for (i = 0; i < last; i++) { var ks = disp[i].map(key_); if (ks.indexOf('新增食材/器模具') >= 0 && ks.indexOf('供應商') >= 0) { h = i; break; } }
  if (h < 0) return { rows: [], status: '新品通知表找不到「新增食材/器模具」標題列' };
  var col = {};
  disp[h].forEach(function (v, j) { var k = HDR[key_(v)]; if (k && col[k] == null) col[k] = j; });
  var n = last - h - 1;
  if (n < 1) return { rows: [], status: 'OK（沒有品項）' };
  var body = disp.slice(h + 1);
  var fml = sh.getRange(h + 2, 1, n, lastC).getFormulas();
  var rich = col.links != null ? sh.getRange(h + 2, col.links + 1, n, 1).getRichTextValues() : [];
  var tabs = {}, rows = [], gid = String(sh.getSheetId());
  body.forEach(function (r, k) {
    var item = col.item != null ? clean_(r[col.item]) : '';
    if (!item || item === '0' || item === '#REF!') return;
    var o = {};
    Object.keys(col).forEach(function (c) {
      var v = clean_(r[col[c]]);
      if (PAIR[c]) { var u = clean_(r[col[c] + 1]); if (u && u !== '0' && v && v !== '0') v = v + ' ' + u; }
      o[c] = v === '0' ? '' : v;
    });
    o.links = links_(r[col.links], rich[k] && rich[k][0]);
    /* 順著品項公式回到甜點分頁 */
    var m = /^=\s*'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)\s*$/.exec(fml[k][col.item] || '');
    if (m) {
      var src = srcTab_(ss, tabs, m[1]), sr = Number(m[3]);
      o.src_tab = m[1]; o.src_row = String(sr);
      if (src) {
        if (!o.dessert) o.dessert = src.name;
        var row = src.rows[sr - 1] || [];
        o.cost = clean_(row[2]); o.provide = clean_(row[11]); o.vendor_src = clean_(row[12]);
        if (!o.supplier || o.supplier === '0') o.supplier = o.vendor_src === '0' ? '' : o.vendor_src;
      }
    }
    ['cost', 'provide', 'vendor_src'].forEach(function (c) { if (o[c] === '0') o[c] = ''; });
    o.campaign = campaign; o.file = f.name; o.folder = f.folder; o.file_id = f.id; o.gid = gid;
    rows.push(o);
  });
  return { rows: rows, status: 'OK' };
}
function srcTab_(ss, cache, name) {
  if (cache[name] !== undefined) return cache[name];
  var s = ss.getSheetByName(name);
  if (!s) return (cache[name] = null);
  var v = s.getRange(1, 1, Math.min(s.getLastRow(), 60), 13).getDisplayValues();
  var nm = '';
  for (var i = 0; i < Math.min(10, v.length); i++) {
    if (/商品正式名稱/.test(v[i][0]) && clean_(v[i][1])) { nm = clean_(v[i][1]); break; }
    if (/商品暫定名稱/.test(v[i][0]) && clean_(v[i][1])) nm = clean_(v[i][1]);
  }
  return (cache[name] = { name: nm || name.replace(/^\d+\s*/, ''), rows: v });
}
function links_(txt, rt) {
  var out = [];
  var add = function (u) { u = String(u || '').trim(); if (/^https?:\/\//i.test(u) && out.indexOf(u) < 0) out.push(u); };
  if (rt) {
    add(rt.getLinkUrl());
    (rt.getRuns() || []).forEach(function (run) { add(run.getLinkUrl()); });
  }
  (String(txt || '').match(/https?:\/\/[^\s，、]+/g) || []).forEach(add);
  return out.join('\n');
}
function write_(name, cols, objs) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var vals = [cols].concat(objs.map(function (o) { return cols.map(function (c) { return o[c] == null ? '' : String(o[c]).slice(0, 5000); }); }));
  sh.clearContents();
  sh.getRange(1, 1, vals.length, cols.length).setNumberFormat('@').setValues(vals);
  sh.setFrozenRows(1);
}
function log_(stat) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName('log') || ss.insertSheet('log');
  if (sh.getLastRow() < 1) sh.appendRow(['ts', 'files', 'items', 'links', 'sec']);
  sh.appendRow([stat.at, stat.files, stat.items, stat.links, stat.sec]);
}

/* ---------------- 權限中心補列 ---------------- */
function authSetup_() {
  var ss = SpreadsheetApp.openById(AUTHC_SS_ID), sh = ss.getSheetByName(AUTHC_TAB);
  if (!sh) throw new Error('找不到權限中心 dim_auth');
  var last = sh.getLastRow(), at = 0;
  if (last > 1) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r, i) { if (!at && String(r[0]).trim() === AUTH_DASH) at = i + 2; });
  var bk = 'dim_auth_備份_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  if (at) {   /* 已有這列：只在名稱不同時改 B 欄名稱（改名用；密碼、效期、啟用都不動） */
    var cur = String(sh.getRange(at, 2).getValue()).trim();
    if (cur === AUTH_NAME) return { ok: true, msg: '已經有 ' + AUTH_DASH + ' 列，名稱也一致，不用改' };
    if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
    sh.getRange(at, 2).setNumberFormat('@').setValue(AUTH_NAME);
    return { ok: true, msg: AUTH_DASH + ' 列名稱「' + cur + '」→「' + AUTH_NAME + '」；備份 ' + bk, row: at };
  }
  if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
  sh.getRange(last + 1, 1, 1, 7).setNumberFormat('@');
  sh.getRange(last + 1, 1, 1, 7).setValues([[AUTH_DASH, AUTH_NAME, '', '12', 'Y', now_(), '海外代理商購買平台儀表板（2026-10-19 補；密碼空白＝還不能用，請在權限管理後台設定）']]);
  return { ok: true, msg: '已補 ' + AUTH_DASH + ' 列；備份 ' + bk, before: last - 1 };
}
