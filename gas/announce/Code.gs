/**
 * 公告及工作清單 API（announce-api-v1，2026-10-01 經營者「20261001 待處理事項」#1）
 * 用途：取代 NUEIP（nuHRM）的「公告管理」「工作清單」——公告與工作清單存在 Google 雲端硬碟資料夾「自己做 公告及工作清單」
 *       （1MnFAKso03ERa8zSj9z0_ytDUVoF66JYg）裡的一份試算表（資料庫）＋附檔子資料夾；儀表板 dashboard-announce.html 發佈、查詢、查閱。
 * 資料：試算表「公告及工作清單_資料庫」（setup 建立，ID 存指令碼屬性 DB_ID；只有擁有者能開，不開連結分享）
 *   fact_announce（一則公告一列；內文 HTML 切 content_html_1～4，每格 ≤45,000 字；content_text＝純文字給全文搜尋）
 *   fact_announce_att（附檔索引：Drive 檔案 ID）、fact_work（工作清單）、fact_work_discuss（討論）、dim_option（分類選單）、log、meta_import
 * 權限：讀內容要「看公告」密碼（dim_auth：announce）；發佈／修改／刪除／上傳／匯入要「發佈者」密碼（dim_auth：announce_pub）。
 *       密碼由「儀表板權限中心」集中管理（UrlFetch verify），驗過結果快取 10 分鐘（只存雜湊）。repo 是公開的，所以內容一律要密碼才讀得到。
 *   doGet（JSONP，公開）只回 ping 與 stat（數字，不含內容）。
 * 安全：寫入包 LockService；刪除＝標記刪除＋附檔搬到「_已刪除」，不永久刪；所有 Drive 動作只限根資料夾底下（防 drive 範圍誤動其他檔）。
 * 部署：網頁應用程式／執行身分＝我／存取＝所有人。第一次：編輯器選 setup → 執行 → 授權（建資料庫試算表與子資料夾）。
 *       之後更新一律「管理部署作業 → 編輯 → 新版本」（網址不變）。
 * v2（2026-10-04 效能第 2 批，只動讀取側）：①POST ping（不用密碼，前端密碼框一出現就預熱）②bundle＝驗密碼＋一次回 list／workList／options
 *       ③密碼驗證結果快取 10 分鐘 → 2 小時（key 含密碼雜湊；改密碼或停用後最多 2 小時才對本頁生效）④公告清單快取 5 分鐘（任何寫入後失效）
 *       ⑤單則 get 只讀那一列的內文，不再整張（含全部內文）讀。回傳 JSON 欄位不變。
 */
var VERSION = 'announce-api-v2';
var AUTH_SEC = 7200, LIST_TTL = 300;
var TZ = 'Asia/Taipei';
var ROOT_FOLDER_ID = '1MnFAKso03ERa8zSj9z0_ytDUVoF66JYg';
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var DASH_VIEW = 'announce', DASH_PUB = 'announce_pub';
var SEG_MAX = 45000, SEG_N = 4, LOG_KEEP = 5000, MAX_FILE = 10 * 1024 * 1024, MAX_ATT = 10;
var TABS = {
  fact_announce: ['ann_id', 'src', 'nueip_id', 'category', 'type', 'title', 'start_date', 'end_date', 'pinned', 'audience', 'creator',
    'created_at', 'updated_at', 'updated_by', 'content_text', 'content_html_1', 'content_html_2', 'content_html_3', 'content_html_4', 'att_count', 'deleted', 'deleted_at'],
  fact_announce_att: ['att_id', 'ann_id', 'file_name', 'mime', 'size', 'drive_file_id', 'url', 'src_url', 'uploaded_at', 'deleted'],
  fact_work: ['work_id', 'src', 'nueip_id', 'item', 'owner', 'creator', 'date', 'due_date', 'done_date', 'progress', 'importance', 'origin',
    'content_text', 'content_html_1', 'content_html_2', 'created_at', 'updated_at', 'updated_by', 'deleted', 'deleted_at'],
  fact_work_discuss: ['work_id', 'author', 'ts', 'text'],
  dim_option: ['kind', 'value', 'sort'],
  log: ['ts', 'who', 'action', 'id', 'ok', 'msg'],
  meta_import: ['ts', 'batch', 'kind', 'count', 'note']
};
var NUM_COLS = { att_count: 1, size: 1, sort: 1, count: 1 };
var SEG_COLS = { fact_announce: ['content_html_1', 'content_html_2', 'content_html_3', 'content_html_4'], fact_work: ['content_html_1', 'content_html_2'] };

/* ================= 第一次設定（在編輯器執行；會跳 Google 授權） ================= */
function setup() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('DB_ID'), ss;
  var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('公告及工作清單_資料庫');
    var f = DriveApp.getFileById(ss.getId()); root.addFile(f); DriveApp.getRootFolder().removeFile(f);
    P.setProperty('DB_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(TZ);
  Object.keys(TABS).forEach(function (n) { ensureTab_(ss, n); });
  ss.getSheets().forEach(function (sh) { if (!TABS[sh.getName()] && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh); });
  ['附檔', '內文圖片', '_匯入原檔', '_已刪除'].forEach(function (n) { sub_(n); });
  var o = load_('dim_option');
  if (!o.rows.length) {
    [['category', '公告', 1], ['category', '一般通知', 2], ['type', '一般公告', 1], ['type', '活動訊息', 2],
     ['progress', '未開始', 1], ['progress', '進行中', 2], ['progress', '已完成', 3], ['progress', '暫停', 4],
     ['importance', '★', 1], ['importance', '★★', 2], ['importance', '★★★', 3]].forEach(function (r) { put_(o, nextRow_(o), { kind: r[0], value: r[1], sort: r[2] }); });
  }
  var msg = '資料庫 ' + ss.getUrl() + '（在資料夾「' + root.getName() + '」）';
  Logger.log(msg); return msg;
}
function ensureTab_(ss, name) {
  var H = TABS[name], sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#FFF3E0'); sh.setFrozenRows(1); }
  var cur = sh.getRange(1, 1, 1, H.length).getValues()[0].map(function (x) { return String(x).trim(); });
  for (var i = 0; i < H.length; i++) if (cur[i] !== H[i]) throw new Error(name + ' 表頭第 ' + (i + 1) + ' 欄是「' + cur[i] + '」，應為「' + H[i] + '」');
  var n = sh.getMaxRows() - 1;
  if (n > 0) H.forEach(function (h, j) { if (!NUM_COLS[h]) sh.getRange(2, j + 1, n, 1).setNumberFormat('@'); });
  return sh;
}
function sub_(name) {
  var root = DriveApp.getFolderById(ROOT_FOLDER_ID), it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}
function subOf_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
/* 防呆：檔案必須在根資料夾底下（drive 權限很大，只准碰這個資料夾） */
function inRoot_(file) {
  var seen = {}, q = [];
  var ps = file.getParents(); while (ps.hasNext()) q.push(ps.next());
  while (q.length) { var f = q.shift(), id = f.getId(); if (id === ROOT_FOLDER_ID) return true; if (seen[id]) continue; seen[id] = 1; var pp = f.getParents(); while (pp.hasNext()) q.push(pp.next()); }
  return false;
}

/* ================= 入口 ================= */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || 'ping'), res;
  try {
    if (a === 'ping') res = { ok: true, data: pingData_() };
    else if (a === 'stat') res = { ok: true, data: stat_() };
    else res = { ok: false, msg: '讀公告內容要用 POST 並帶密碼' };
  } catch (err) { res = { ok: false, msg: errMsg_(err) }; }
  res.act = a;
  return out_(res, p.callback);
}
function pingData_() { return { version: VERSION, now: now_(), ready: !!PropertiesService.getScriptProperties().getProperty('DB_ID') }; }
var ACTIONS = { whoami: whoami_, list: list_, get: get_, save: save_, del: del_, upload: upload_, uploadImg: uploadImg_, attDel: attDel_,
  workList: workList_, workSave: workSave_, workDel: workDel_, options: options_, optSave: optSave_, importBatch: importBatch_, bundle: bundle_ };
var WRITES = { save: 1, del: 1, upload: 1, uploadImg: 1, attDel: 1, workSave: 1, workDel: 1, optSave: 1, importBatch: 1 };
var PUB_ONLY = WRITES;
var RQ_SEC = 21600;
function doPost(e) {
  var p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, msg: '送來的資料不是 JSON' }); }
  var act = String(p.action || ''), fn = ACTIONS[act], res;
  if (act === 'ping') { res = { ok: true, data: pingData_() }; res.act = 'ping'; return out_(res); }   /* v2：預熱用（不用密碼、不碰試算表），回法同 GET ping */
  if (!fn) return out_({ ok: false, act: act, msg: '不支援的動作：' + act });
  var who;
  try { who = auth_(p.password, (PUB_ONLY[act] || (act === 'whoami' && p.role === 'pub')) ? DASH_PUB : DASH_VIEW); }
  catch (err) { Utilities.sleep(800); res = { ok: false, act: act, code: 'auth', msg: errMsg_(err) }; log_('', act, '', false, res.msg); return out_(res); }
  var lock = LockService.getScriptLock();
  if (WRITES[act] && !lock.tryLock(15000)) return out_({ ok: false, act: act, code: 'busy', msg: '系統忙碌中，請稍後再按一次' });
  var rqKey = (WRITES[act] && p.rq) ? 'rq:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(act + '|' + String(p.rq).slice(0, 120)).getBytes()) : '';
  try {
    var cached = rqKey ? CacheService.getScriptCache().get(rqKey) : null;
    if (cached) res = JSON.parse(cached);
    else {
      var r = fn(p, who) || {};
      res = { ok: true, act: act, data: r.data };
      if (rqKey) { try { CacheService.getScriptCache().put(rqKey, JSON.stringify(res), RQ_SEC); } catch (ce) { } }
      if (WRITES[act]) log_(who, act, r.id || '', true, r.msg || '');
    }
  } catch (err) {
    res = { ok: false, act: act, msg: errMsg_(err) }; if (err && err.code) res.code = err.code;
    log_(who, act, '', false, res.msg);
  } finally { if (WRITES[act]) { lock.releaseLock(); listCacheClear_(); } }   /* v2：任何寫入做完（成功或失敗）公告清單暫存就失效 */
  return out_(res);
}
function out_(res, cb) {
  var s = JSON.stringify(res);
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(String(cb))) return ContentService.createTextOutput(cb + '(' + s + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
/* 密碼：問「儀表板權限中心」verify（dim_auth 的 announce／announce_pub）；成功結果快取 AUTH_SEC（v2：2 小時；key＝雜湊，不存明文）。
   代價：在權限中心改密碼或 enabled=N 後，拿舊密碼的人最多再用 2 小時（只影響本頁；其他儀表板不經本 API） */
function auth_(pw, dash) {
  pw = String(pw || '').trim();
  if (!pw) throw fail_('請輸入密碼', 'auth');
  var cache = CacheService.getScriptCache(), key = 'au:' + dash + ':' + hash_(dash + '|' + pw);
  if (cache.get(key)) return dash === DASH_PUB ? '發佈者' : '讀者';
  var url = AUTH_API + '?action=verify&dashboard=' + encodeURIComponent(dash) + '&pwd=' + encodeURIComponent(pw) + '&callback=cb';
  var t = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getContentText();
  var m = t.match(/^\s*[A-Za-z_$][\w$]*\(([\s\S]*)\);?\s*$/), j = {};
  try { j = JSON.parse(m ? m[1] : t); } catch (e) { throw fail_('權限中心回應看不懂，請稍後再試', 'auth'); }
  if (!j || !j.ok) throw fail_(dash === DASH_PUB ? '發佈者密碼不對（' + (j && j.error || '') + '）' : '密碼不對（' + (j && j.error || '') + '）', 'auth');
  cache.put(key, '1', AUTH_SEC);
  return dash === DASH_PUB ? '發佈者' : '讀者';
}
function whoami_(p, who) { return { data: { who: who } }; }
/* v2：開頁一次拿齊（驗密碼＋公告清單＋工作清單＋選單），三份各自的欄位與原本 list／workList／options 相同 */
function bundle_(p, who) { return { data: { list: list_(p, who).data, work: workList_(p, who).data, options: options_(p, who).data } }; }

/* ================= 公告 ================= */
function stat_() {
  var P = PropertiesService.getScriptProperties(); if (!P.getProperty('DB_ID')) return { ready: false };
  var c = CacheService.getScriptCache(), hit = c.get('stat'); if (hit) return JSON.parse(hit);
  var t = today_(), a = load_('fact_announce', true).rows.filter(function (r) { return r.deleted !== 'Y'; });
  var act = a.filter(function (r) { return (!r.start_date || r.start_date <= t) && (!r.end_date || r.end_date >= t); }).length;
  var w = load_('fact_work', true).rows.filter(function (r) { return r.deleted !== 'Y' && r.progress !== '已完成'; }).length;
  var res = { ready: true, active: act, total: a.length, workOpen: w };
  c.put('stat', JSON.stringify(res), 300); return res;
}
function annView_(r, withHtml) {
  var o = { ann_id: r.ann_id, src: r.src, nueip_id: r.nueip_id, category: r.category, type: r.type, title: r.title, start_date: r.start_date, end_date: r.end_date,
    pinned: r.pinned === 'Y', audience: r.audience, creator: r.creator, created_at: r.created_at, updated_at: r.updated_at, att_count: Number(r.att_count) || 0,
    text: String(r.content_text || '').slice(0, withHtml ? 100000 : 20000) };
  if (withHtml) o.html = [r.content_html_1, r.content_html_2, r.content_html_3, r.content_html_4].map(str_).join('');
  return o;
}
/* v2 公告清單暫存：清單（不含內文）＋「ann_id → 列號」放 CacheService 5 分鐘，鍵尾帶版本號 ann:ver；任何寫入做完就換版本號＝舊暫存失效
   （比直接刪鍵保險：讀的人剛好拿到舊表、寫的人同時寫完清掉、讀的人再放回去——放回去的是舊鍵，沒人會再讀到）。
   有人直接改試算表（不經本 API）→ 清單最多 5 分鐘後才看到；get 用列號讀回要是同一個 ann_id 才算，不然掃一次最新清單 */
function listRows_() {
  var cache = CacheService.getScriptCache(), key = 'ann:list:' + (cache.get('ann:ver') || '0');
  var hit = cacheGetBig_(cache, key);
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* 壞掉就重讀 */ } }
  var t = load_('fact_announce', true), rows = [], idx = {};
  t.rows.forEach(function (r) { if (r.deleted !== 'Y' && r.ann_id) { rows.push(annView_(r, false)); idx[r.ann_id] = r._row; } });
  var v = { rows: rows, idx: idx };
  try { cachePutBig_(cache, key, JSON.stringify(v), LIST_TTL); } catch (e) { /* 放不進去就每次讀 */ }
  return v;
}
function listCacheClear_() {
  try { var c = CacheService.getScriptCache(); c.put('ann:ver', String((parseInt(c.get('ann:ver') || '0', 10) || 0) + 1), 21600); } catch (e) { /* 清不掉就等 5 分鐘自然過期 */ }
}
function list_(p, who) {
  return { data: { rows: listRows_().rows, today: today_(), who: who } };
}
/* 讀一列成物件（同 load_ 的欄位規則） */
function rowObj_(sh, H, rn) {
  var v = sh.getRange(rn, 1, 1, H.length).getValues()[0], o = { _row: rn };
  H.forEach(function (h, j) { o[h] = NUM_COLS[h] ? v[j] : str_(v[j]); });
  return o;
}
/* v2：找一則公告（含內文）。先用清單暫存的列號直接讀那一列（讀回 ann_id 要相同）；對不上→掃一次清單（不含內文）定位，再只讀那一列的 content_html_1～4 */
function annRow_(id) {
  if (!id) return null;
  var sh = ss_().getSheetByName('fact_announce'); if (!sh) throw fail_('找不到分頁 fact_announce', 'setup');
  var H = TABS.fact_announce, rn = 0;
  try { rn = listRows_().idx[id] || 0; } catch (e) { rn = 0; }
  if (rn >= 2 && rn <= sh.getLastRow()) { var r = rowObj_(sh, H, rn); if (r.ann_id === id) return r.deleted === 'Y' ? null : r; }
  var t = load_('fact_announce', true), hit = t.rows.filter(function (x) { return x.ann_id === id && x.deleted !== 'Y'; })[0];
  if (!hit) return null;
  var segs = SEG_COLS.fact_announce, v = sh.getRange(hit._row, H.indexOf(segs[0]) + 1, 1, segs.length).getValues()[0];
  segs.forEach(function (h, j) { hit[h] = str_(v[j]); });
  return hit;
}
function get_(p, who) {
  var id = str_(p.ann_id), r = annRow_(id);
  if (!r) throw fail_('找不到公告 ' + id, 'notfound');
  var att = load_('fact_announce_att').rows.filter(function (a) { return a.ann_id === id && a.deleted !== 'Y'; })
    .map(function (a) { return { att_id: a.att_id, name: a.file_name, mime: a.mime, size: Number(a.size) || 0, url: a.url }; });
  var o = annView_(r, true); o.att = att;
  return { data: o };
}
function cleanTxt_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').slice(0, max || 500); }
function save_(p, who) {
  var a = p.ann || {}, t = load_('fact_announce'), id = str_(a.ann_id), row = null;
  var title = cleanTxt_(a.title, 200).trim(); if (!title) throw fail_('標題必填', 'invalid');
  var sd = normDate_(a.start_date), ed = normDate_(a.end_date);
  if (sd && !/^\d{4}-\d{2}-\d{2}$/.test(sd)) throw fail_('公告開始日格式要 YYYY-MM-DD', 'invalid');
  if (ed && !/^\d{4}-\d{2}-\d{2}$/.test(ed)) throw fail_('公告結束日格式要 YYYY-MM-DD', 'invalid');
  if (sd && ed && sd > ed) throw fail_('開始日不能晚於結束日', 'invalid');
  var html = String(a.html || ''), text = cleanTxt_(a.text || htmlText_(html), 45000);
  var segs = split_(html, SEG_N);
  if (id) { row = t.rows.filter(function (x) { return x.ann_id === id; })[0]; if (!row) throw fail_('找不到公告 ' + id, 'notfound');
    if (a.base_updated && str_(row.updated_at) !== str_(a.base_updated)) throw fail_('這則公告剛被別人改過，請重新讀取再改', 'conflict'); }
  else id = 'A' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  var o = { ann_id: id, src: row ? row.src : (a.src || 'dash'), nueip_id: row ? row.nueip_id : str_(a.nueip_id), category: cleanTxt_(a.category, 50), type: cleanTxt_(a.type, 20) || '一般公告',
    title: title, start_date: sd, end_date: ed, pinned: a.pinned ? 'Y' : '', audience: cleanTxt_(a.audience, 300), creator: row ? row.creator : (cleanTxt_(a.creator, 50) || who),
    created_at: row ? row.created_at : ((a.src === 'nueip' && a.created_at) ? cleanTxt_(a.created_at, 30) : now_()), updated_at: now_(), updated_by: who, content_text: text,
    content_html_1: segs[0] || '', content_html_2: segs[1] || '', content_html_3: segs[2] || '', content_html_4: segs[3] || '',
    att_count: row ? row.att_count : 0, deleted: '', deleted_at: '' };
  put_(t, row ? row._row : nextRow_(t), o);
  CacheService.getScriptCache().remove('stat');
  return { id: id, data: { ann_id: id, updated_at: o.updated_at }, msg: (row ? '更新' : '新增') + '公告：' + title };
}
function del_(p, who) {
  var id = str_(p.ann_id), t = load_('fact_announce'), row = t.rows.filter(function (x) { return x.ann_id === id; })[0];
  if (!row) throw fail_('找不到公告 ' + id, 'notfound');
  putCols_(t, row._row, { deleted: 'Y', deleted_at: now_(), updated_by: who }, function (h) { return h === 'deleted' || h === 'deleted_at' || h === 'updated_by'; });
  var at = load_('fact_announce_att'), trash = sub_('_已刪除');
  at.rows.filter(function (a) { return a.ann_id === id && a.deleted !== 'Y'; }).forEach(function (a) {
    try { var f = DriveApp.getFileById(a.drive_file_id); if (inRoot_(f)) f.moveTo(trash); } catch (e) { }
    putCols_(at, a._row, { deleted: 'Y' }, function (h) { return h === 'deleted'; });
  });
  CacheService.getScriptCache().remove('stat');
  return { id: id, data: { ann_id: id }, msg: '刪除（標記）公告 ' + row.title };
}
function upload_(p, who) {
  var id = str_(p.ann_id), t = load_('fact_announce'), row = t.rows.filter(function (x) { return x.ann_id === id && x.deleted !== 'Y'; })[0];
  if (!row) throw fail_('找不到公告 ' + id + '（請先存公告再上傳附檔）', 'notfound');
  var at = load_('fact_announce_att'), have = at.rows.filter(function (a) { return a.ann_id === id && a.deleted !== 'Y'; });
  if (have.length >= MAX_ATT) throw fail_('每則公告最多 ' + MAX_ATT + ' 個附檔', 'invalid');
  var name = cleanTxt_(p.name, 150).replace(/[\\\/:*?"<>|]/g, '_') || 'file', mime = cleanTxt_(p.mime, 100) || 'application/octet-stream';
  var bytes = Utilities.base64Decode(String(p.b64 || ''));
  if (!bytes.length) throw fail_('檔案是空的', 'invalid');
  if (bytes.length > MAX_FILE) throw fail_('檔案超過 10MB', 'invalid');
  var folder = subOf_(sub_('附檔'), id), f = folder.createFile(Utilities.newBlob(bytes, mime, name));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);   /* 單檔「知道連結可檢視」；連結只經過要密碼的 API 給出去，資料夾本身不公開 */
  var aid = 'F' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100);
  put_(at, nextRow_(at), { att_id: aid, ann_id: id, file_name: name, mime: mime, size: bytes.length, drive_file_id: f.getId(), url: f.getUrl(), src_url: cleanTxt_(p.src_url, 500), uploaded_at: now_(), deleted: '' });
  putCols_(t, row._row, { att_count: have.length + 1 }, function (h) { return h === 'att_count'; });
  return { id: id, data: { att_id: aid, url: f.getUrl(), name: name }, msg: '上傳附檔 ' + name };
}
/* 內文圖片：貼進內文的圖片（data:）存成 Drive 檔，內文改用縮圖網址；不算附檔、不進 fact_announce_att */
function uploadImg_(p, who) {
  var id = str_(p.ann_id), t = load_('fact_announce', true);
  if (!t.rows.some(function (x) { return x.ann_id === id && x.deleted !== 'Y'; })) throw fail_('找不到公告 ' + id, 'notfound');
  var mime = cleanTxt_(p.mime, 50); if (!/^image\/(png|jpe?g|gif|webp)$/.test(mime)) throw fail_('只接受 png／jpg／gif／webp 圖片', 'invalid');
  var bytes = Utilities.base64Decode(String(p.b64 || ''));
  if (!bytes.length || bytes.length > MAX_FILE) throw fail_('圖片是空的或超過 10MB', 'invalid');
  var name = 'img-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100) + '.' + mime.split('/')[1].replace('jpeg', 'jpg');
  var f = subOf_(sub_('內文圖片'), id).createFile(Utilities.newBlob(bytes, mime, name));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: id, data: { file_id: f.getId(), src: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w1600' }, msg: '內文圖片 ' + name };
}
function attDel_(p, who) {
  var aid = str_(p.att_id), at = load_('fact_announce_att'), a = at.rows.filter(function (x) { return x.att_id === aid && x.deleted !== 'Y'; })[0];
  if (!a) throw fail_('找不到附檔', 'notfound');
  try { var f = DriveApp.getFileById(a.drive_file_id); if (inRoot_(f)) f.moveTo(sub_('_已刪除')); } catch (e) { }
  putCols_(at, a._row, { deleted: 'Y' }, function (h) { return h === 'deleted'; });
  var t = load_('fact_announce'), row = t.rows.filter(function (x) { return x.ann_id === a.ann_id; })[0];
  if (row) putCols_(t, row._row, { att_count: Math.max(0, (Number(row.att_count) || 1) - 1) }, function (h) { return h === 'att_count'; });
  return { id: a.ann_id, data: { att_id: aid }, msg: '刪除附檔 ' + a.file_name };
}

/* ================= 工作清單 ================= */
function workView_(r) {
  return { work_id: r.work_id, src: r.src, nueip_id: r.nueip_id, item: r.item, owner: r.owner, creator: r.creator, date: r.date, due_date: r.due_date, done_date: r.done_date,
    progress: r.progress, importance: r.importance, origin: r.origin, text: r.content_text, html: str_(r.content_html_1) + str_(r.content_html_2), updated_at: r.updated_at };
}
function workList_(p, who) {
  var t = load_('fact_work'), d = load_('fact_work_discuss'), by = {};
  d.rows.forEach(function (x) { (by[x.work_id] = by[x.work_id] || []).push({ author: x.author, ts: x.ts, text: x.text }); });
  return { data: { rows: t.rows.filter(function (r) { return r.deleted !== 'Y' && r.work_id; }).map(function (r) { var o = workView_(r); o.discuss = by[r.work_id] || []; return o; }) } };
}
function workSave_(p, who) {
  var w = p.work || {}, t = load_('fact_work'), id = str_(w.work_id), row = null;
  var item = cleanTxt_(w.item, 200).trim(); if (!item) throw fail_('工作項目必填', 'invalid');
  if (id) { row = t.rows.filter(function (x) { return x.work_id === id; })[0]; if (!row) throw fail_('找不到工作 ' + id, 'notfound'); }
  else id = 'W' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
  var html = String(w.html || ''), segs = split_(html, 2);
  var o = { work_id: id, src: row ? row.src : (w.src || 'dash'), nueip_id: row ? row.nueip_id : str_(w.nueip_id), item: item, owner: cleanTxt_(w.owner, 100), creator: row ? row.creator : (cleanTxt_(w.creator, 50) || who),
    date: normDate_(w.date), due_date: normDate_(w.due_date), done_date: normDate_(w.done_date), progress: cleanTxt_(w.progress, 20), importance: cleanTxt_(w.importance, 10), origin: cleanTxt_(w.origin, 20),
    content_text: cleanTxt_(w.text || htmlText_(html), 45000), content_html_1: segs[0] || '', content_html_2: segs[1] || '',
    created_at: row ? row.created_at : ((w.src === 'nueip' && w.created_at) ? cleanTxt_(w.created_at, 30) : now_()), updated_at: now_(), updated_by: who, deleted: '', deleted_at: '' };
  put_(t, row ? row._row : nextRow_(t), o);
  CacheService.getScriptCache().remove('stat');
  return { id: id, data: { work_id: id }, msg: (row ? '更新' : '新增') + '工作：' + item };
}
function workDel_(p, who) {
  var id = str_(p.work_id), t = load_('fact_work'), row = t.rows.filter(function (x) { return x.work_id === id; })[0];
  if (!row) throw fail_('找不到工作 ' + id, 'notfound');
  putCols_(t, row._row, { deleted: 'Y', deleted_at: now_(), updated_by: who }, function (h) { return h === 'deleted' || h === 'deleted_at' || h === 'updated_by'; });
  CacheService.getScriptCache().remove('stat');
  return { id: id, data: { work_id: id }, msg: '刪除（標記）工作 ' + row.item };
}
function options_(p, who) {
  var o = {}; load_('dim_option').rows.sort(function (a, b) { return (Number(a.sort) || 0) - (Number(b.sort) || 0); }).forEach(function (r) { (o[r.kind] = o[r.kind] || []).push(r.value); });
  return { data: o };
}
function optSave_(p, who) {
  var kind = str_(p.kind), value = cleanTxt_(p.value, 50).trim();
  if (['category', 'type', 'progress', 'importance'].indexOf(kind) < 0 || !value) throw fail_('選單種類或值不對', 'invalid');
  var t = load_('dim_option'); if (t.rows.some(function (r) { return r.kind === kind && r.value === value; })) return { data: { dup: true } };
  put_(t, nextRow_(t), { kind: kind, value: value, sort: t.rows.length + 1 });
  return { id: value, data: { kind: kind, value: value }, msg: '新增選單 ' + kind + '：' + value };
}
/* 一次性匯入（NUEIP 匯出檔，儀表板「匯入」分頁上傳）：以 nueip_id 判斷是否已匯入（重匯同一批不會重複） */
function importBatch_(p, who) {
  var kind = str_(p.kind), items = p.items || [], batch = str_(p.batch) || stamp_(), n = 0, skip = 0, ids = {};
  if (!Array.isArray(items) || items.length > 60) throw fail_('每批最多 60 筆', 'invalid');
  if (kind === 'announce') {
    var t = load_('fact_announce'), have = {}; t.rows.forEach(function (r) { if (r.nueip_id) have[r.nueip_id] = 1; });
    t.rows.forEach(function (r) { if (r.nueip_id) ids[r.nueip_id] = r.ann_id; });
    items.forEach(function (a) { a.nueip_id = str_(a.nueip_id); if (!a.nueip_id || have[a.nueip_id]) { skip++; return; } a.src = 'nueip'; delete a.ann_id; var r = save_({ ann: a }, who); ids[a.nueip_id] = r.id; have[a.nueip_id] = 1; n++; });
  } else if (kind === 'work') {
    var w = load_('fact_work'), hw = {}; w.rows.forEach(function (r) { if (r.nueip_id) hw[r.nueip_id] = 1; });
    var dt = load_('fact_work_discuss');
    w.rows.forEach(function (r) { if (r.nueip_id) ids[r.nueip_id] = r.work_id; });
    items.forEach(function (x) { x.nueip_id = str_(x.nueip_id); if (!x.nueip_id || hw[x.nueip_id]) { skip++; return; } x.src = 'nueip'; delete x.work_id; var r = workSave_({ work: x }, who); ids[x.nueip_id] = r.id; hw[x.nueip_id] = 1; n++;
      (x.discuss || []).forEach(function (d) { put_(dt, nextRow_(dt), { work_id: r.id, author: cleanTxt_(d.author, 50), ts: cleanTxt_(d.ts, 30), text: cleanTxt_(d.text, 5000) }); }); });
  } else throw fail_('kind 要是 announce 或 work', 'invalid');
  var m = load_('meta_import'); put_(m, nextRow_(m), { ts: now_(), batch: batch, kind: kind, count: n, note: '略過已匯入 ' + skip });
  return { id: batch, data: { imported: n, skipped: skip, ids: ids }, msg: '匯入 ' + kind + ' ' + n + ' 筆（略過 ' + skip + '）' };
}

/* ================= 試算表工具 ================= */
var _ss = null;
function ss_() {
  if (_ss) return _ss;
  var id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw fail_('資料庫還沒建立（請在編輯器執行 setup）', 'setup');
  _ss = SpreadsheetApp.openById(id); return _ss;
}
function load_(name, skipSeg) {
  var sh = ss_().getSheetByName(name); if (!sh) throw fail_('找不到分頁 ' + name, 'setup');
  var H = TABS[name], last = sh.getLastRow(), rows = [];
  if (last > 1) {
    var skip = skipSeg ? (SEG_COLS[name] || []) : [];
    var v = sh.getRange(2, 1, last - 1, H.length).getValues();
    v.forEach(function (r, i) { var o = { _row: i + 2 }; H.forEach(function (h, j) { if (skip.indexOf(h) < 0) o[h] = NUM_COLS[h] ? r[j] : str_(r[j]); }); rows.push(o); });
  }
  return { sh: sh, H: H, rows: rows };
}
function nextRow_(t) { var r = t.sh.getLastRow() + 1; if (r > t.sh.getMaxRows()) t.sh.insertRowsAfter(t.sh.getMaxRows(), 200); return r; }
function put_(t, rowNo, obj) { putCols_(t, rowNo, obj, function () { return true; }); }
function putCols_(t, rowNo, obj, pred) {
  var H = t.H;
  H.forEach(function (h, j) {
    if (!pred(h) || obj[h] === undefined) return;
    var rg = t.sh.getRange(rowNo, j + 1);
    if (!NUM_COLS[h]) rg.setNumberFormat('@');
    rg.setValue(cell_(h, obj[h]));
  });
}
function cell_(h, v) { if (v === null || v === undefined) return ''; if (NUM_COLS[h]) { if (v === '') return ''; var n = Number(v); return isFinite(n) ? n : 0; } return safe_(str_(v)); }
function safe_(s) { return /^[=+'@-]/.test(s) ? "'" + s : s; }
function split_(s, n) {
  var out = [], i = 0;
  while (i < s.length) { var end = Math.min(i + SEG_MAX, s.length); if (end < s.length) { var hi = s.charCodeAt(end - 1); if (hi >= 0xD800 && hi <= 0xDBFF) end--; } out.push(s.slice(i, end)); i = end; }
  if (out.length > n) throw fail_('內容太長（' + s.length + ' 字，上限 ' + (SEG_MAX * n) + ' 字），請把圖片改用附檔、或精簡內容', 'toobig');
  return out;
}
function htmlText_(h) { return String(h || '').replace(/<style[\s\S]*?<\/style>/gi, '').replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<br\s*\/?>/gi, '\n').replace(/<\/(p|div|tr|li|h\d)>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/\n{3,}/g, '\n\n').trim(); }
function log_(who, act, id, ok, msg) {
  try {
    var sh = ss_().getSheetByName('log'); var r = sh.getLastRow() + 1; if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 500);
    sh.getRange(r, 1, 1, 6).setNumberFormat('@').setValues([[now_(), safe_(String(who || '')), safe_(String(act || '')), safe_(String(id || '')), ok ? 'Y' : 'N', safe_(String(msg || '').slice(0, 500))]]);
    if (r > LOG_KEEP + 500) sh.deleteRows(2, r - 1 - LOG_KEEP);
  } catch (e) { }
}
function hash_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s)).slice(0, 32); }
/* CacheService 一個鍵最多 100KB：超過就切塊（每塊 ≤30,000 個字＝UTF-8 最多 90KB；不切在 emoji 代理對中間），主鍵只放「#塊數」；太大（>40 塊）就不暫存 */
var CACHE_CHUNK = 30000, CACHE_MAX_CHUNKS = 40;
function cachePutBig_(cache, key, s, ttl) {
  s = String(s);
  if (s.length <= CACHE_CHUNK) { cache.put(key, s, ttl); return; }
  var parts = {}, n = 0, i = 0;
  while (i < s.length) {
    var end = Math.min(i + CACHE_CHUNK, s.length), hi = s.charCodeAt(end - 1);
    if (end < s.length && hi >= 0xD800 && hi <= 0xDBFF) end--;
    parts[key + ':' + n] = s.slice(i, end); n++; i = end;
    if (n > CACHE_MAX_CHUNKS) return;
  }
  cache.putAll(parts, ttl);
  cache.put(key, '#' + n, ttl);
}
function cacheGetBig_(cache, key) {
  var head = cache.get(key);
  if (head === null || head === undefined) return null;
  if (head.charAt(0) !== '#') return head;
  var n = parseInt(head.slice(1), 10), keys = [];
  for (var i = 0; i < n; i++) keys.push(key + ':' + i);
  var all = cache.getAll(keys), s = '';
  for (i = 0; i < n; i++) { var c = all[key + ':' + i]; if (c === null || c === undefined) return null; s += c; }
  return s;
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function stamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm'); }
function str_(v) { if (v === null || v === undefined) return ''; if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss'); return String(v); }
function normDate_(v) { if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); var s = str_(v).trim(); if (!s) return ''; var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/); return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s; }
function fail_(msg, code) { var e = new Error(msg); if (code) e.code = code; return e; }
function errMsg_(err) { return String((err && err.message) || err || '未知錯誤'); }
