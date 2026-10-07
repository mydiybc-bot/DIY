/**
 * 公告及工作清單 API（announce-api-v1，2026-10-01 經營者「20261001 待處理事項」#1）
 * 用途：取代 NUEIP（nuHRM）的「公告管理」「工作清單」——公告與工作清單存在 Google 雲端硬碟資料夾「自己做 公告及工作清單」
 *       （1MnFAKso03ERa8zSj9z0_ytDUVoF66JYg）裡的一份試算表（資料庫）＋附檔子資料夾；儀表板 dashboard-announce.html 發佈、查詢、查閱。
 * 資料：試算表「公告及工作清單_資料庫」（setup 建立，ID 存指令碼屬性 DB_ID；只有擁有者能開，不開連結分享）
 *   fact_announce（一則公告一列；內文 HTML 切 content_html_1～4，每格 ≤45,000 字；content_text＝純文字給全文搜尋）
 *   fact_announce_att（附檔索引：Drive 檔案 ID）、fact_work（工作清單）、fact_work_discuss（討論）、dim_option（分類選單）、log、meta_import
 * 權限（2026-10-06 起，見 v3）：公告免密碼；工作清單要「看公告」密碼（dim_auth：announce）；發佈／修改／刪除／上傳／匯入要「發佈者」密碼（dim_auth：announce_pub）。
 *       密碼由「儀表板權限中心」集中管理（UrlFetch verify），驗過結果快取（只存雜湊）。
 *   doGet（JSONP，公開）回 ping、stat（數字）、active（有效期內公告）。
 * 安全：寫入包 LockService；刪除＝標記刪除＋附檔搬到「_已刪除」，不永久刪；所有 Drive 動作只限根資料夾底下（防 drive 範圍誤動其他檔）。
 * 部署：網頁應用程式／執行身分＝我／存取＝所有人。第一次：編輯器選 setup → 執行 → 授權（建資料庫試算表與子資料夾）。
 *       之後更新一律「管理部署作業 → 編輯 → 新版本」（網址不變）。
 * v2（2026-10-04 效能第 2 批，只動讀取側）：①POST ping（不用密碼，前端密碼框一出現就預熱）②bundle＝驗密碼＋一次回 list／workList／options
 *       ③密碼驗證結果快取 10 分鐘 → 2 小時（key 含密碼雜湊；改密碼或停用後最多 2 小時才對本頁生效）④公告清單快取 5 分鐘（任何寫入後失效）
 *       ⑤單則 get 只讀那一列的內文，不再整張（含全部內文）讀。回傳 JSON 欄位不變。
 * v3（2026-10-06 經營者裁定「看公告不需要密碼」，各儀表板調整 1006）：公告的讀取 list／get／options／active 免密碼；
 *       工作清單 workList 仍要「看公告」密碼（announce）；發佈／修改／刪除／上傳／匯入仍要管理者密碼（announce_pub）。
 *       bundle 沒帶密碼＝只回公告與選單（work:null、workLocked:true）；帶對的看公告密碼＝跟原本一樣三份都回。
 *       新增 active＝今天在有效期內的公告（含純文字內文與附檔清單；給決策中心店長頁）；GET（JSONP）也可讀 active。
 * v4（2026-10-06 晚，經營者「工作清單也不需要密碼就可以看」）：workList 免密碼；bundle 不帶密碼也回公告＋工作清單＋選單（work 不再 null）。
 *       寫入（發佈／修改／刪除／上傳／匯入／工作新增修改）照舊要管理者密碼（announce_pub）。「看公告」密碼（dim_auth：announce）從此沒有地方用到。
 * v5（2026-10-07，經營者：工作清單搬進 104 指定員工公告，參與人員點公告裡的連結回到本系統留言、可上傳檔案）：
 *       新增 discussAdd（工作留言，免密碼、要填名字；每則 ≤5 個檔、每個 ≤10MB、合計 ≤20MB、限常見格式；防灌：同件工作每分鐘 10 則、全站每分鐘 30 則、每天 500 則）
 *       與 discussDel（刪留言＝標記刪除＋附檔搬「_已刪除」，要管理者密碼）。fact_work_discuss 表頭在原 4 欄後面補 6 欄
 *       （disc_id／role／att（附檔 JSON）／src／deleted／deleted_at）；升級要在編輯器執行一次 setupV5_20261007（先備份分頁、再補表頭與舊討論的留言編號）。
 *       留言附檔存「工作討論附檔/<工作編號>/」。workList 的 discuss 多回 disc_id／role／att／src，依時間舊到新排、不含已刪除。
 * v5.1（2026-10-07 午，經營者「改成管理者也不能刪留言」）：拿掉 discussDel（API 不再提供刪留言，管理者密碼也刪不了；discussDel_ 與一次性
 *       selfTestV5_20261007 一併移除）。deleted／deleted_at 欄保留（v5 自我測試留下 1 列已標記刪除的測試留言，discBy_ 照樣濾掉）。
 * v5.2（2026-10-07 傍晚，經營者「從 104 點連結開留言區太久」）：workList 結果放 CacheService 5 分鐘（鍵帶 ann:ver，任何寫入含留言做完就換版本＝失效），
 *       bundle 也吃同一份。直接改試算表（不經本 API）→ 工作清單最多 5 分鐘後才看到。回傳欄位不變。
 */
var VERSION = 'announce-api-v5.2';   // 2026-10-07 v5.2：工作清單暫存 5 分鐘（加快 104 連過來）；v5.1：留言誰都不能刪（拿掉 discussDel）；v5：工作清單留言 discussAdd（免密碼、可附檔）；v4：工作清單 workList 也免密碼、bundle 不帶密碼回三份；v3：公告免密碼（list／get／options／active），工作清單仍要密碼；v2＝2026-10-04 併入 v1.1 修正（setSharing 被拒略過＋fixAttIndex20261004）
var AUTH_SEC = 7200, LIST_TTL = 300;
var TZ = 'Asia/Taipei';
var ROOT_FOLDER_ID = '1MnFAKso03ERa8zSj9z0_ytDUVoF66JYg';
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var DASH_VIEW = 'announce', DASH_PUB = 'announce_pub';
var SEG_MAX = 45000, SEG_N = 4, LOG_KEEP = 5000, MAX_FILE = 10 * 1024 * 1024, MAX_ATT = 10;
/* v5 工作留言：上限與可收的副檔名（前端 dashboard-announce.html 的 DISC_* 要跟這裡一致） */
var DISC_TEXT_MAX = 3000, DISC_MAX_FILES = 5, DISC_MAX_FILE = 10 * 1024 * 1024, DISC_MAX_TOTAL = 20 * 1024 * 1024;
var DISC_EXT = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv', 'mp4', 'mov', 'm4a', 'mp3'];
var TABS = {
  fact_announce: ['ann_id', 'src', 'nueip_id', 'category', 'type', 'title', 'start_date', 'end_date', 'pinned', 'audience', 'creator',
    'created_at', 'updated_at', 'updated_by', 'content_text', 'content_html_1', 'content_html_2', 'content_html_3', 'content_html_4', 'att_count', 'deleted', 'deleted_at'],
  fact_announce_att: ['att_id', 'ann_id', 'file_name', 'mime', 'size', 'drive_file_id', 'url', 'src_url', 'uploaded_at', 'deleted'],
  fact_work: ['work_id', 'src', 'nueip_id', 'item', 'owner', 'creator', 'date', 'due_date', 'done_date', 'progress', 'importance', 'origin',
    'content_text', 'content_html_1', 'content_html_2', 'created_at', 'updated_at', 'updated_by', 'deleted', 'deleted_at'],
  fact_work_discuss: ['work_id', 'author', 'ts', 'text', 'disc_id', 'role', 'att', 'src', 'deleted', 'deleted_at'],   /* v5：後 6 欄是 2026-10-07 補的（setupV5_20261007） */
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
    else if (a === 'active') res = { ok: true, data: active_({}, '訪客').data };
    else res = { ok: false, msg: '讀公告內容請用 POST' };
  } catch (err) { res = { ok: false, msg: errMsg_(err) }; }
  res.act = a;
  return out_(res, p.callback);
}
function pingData_() { return { version: VERSION, now: now_(), ready: !!PropertiesService.getScriptProperties().getProperty('DB_ID') }; }
var ACTIONS = { whoami: whoami_, list: list_, get: get_, save: save_, del: del_, upload: upload_, uploadImg: uploadImg_, attDel: attDel_,
  workList: workList_, workSave: workSave_, workDel: workDel_, options: options_, optSave: optSave_, importBatch: importBatch_, bundle: bundle_, active: active_,
  discussAdd: discussAdd_ };   /* v5.1：沒有刪留言的動作（經營者：管理者也不能刪） */
var OPEN = { list: 1, get: 1, options: 1, active: 1, workList: 1, discussAdd: 1 };   /* v5：工作留言免密碼（要填名字）；v4（2026-10-06 晚）：工作清單也免密碼；v3：公告免密碼。其他寫入照舊要管理者密碼 */
var WRITES = { save: 1, del: 1, upload: 1, uploadImg: 1, attDel: 1, workSave: 1, workDel: 1, optSave: 1, importBatch: 1, discussAdd: 1 };
var PUB_ONLY = { save: 1, del: 1, upload: 1, uploadImg: 1, attDel: 1, workSave: 1, workDel: 1, optSave: 1, importBatch: 1 };   /* v5 起與 WRITES 分開：discussAdd 是寫入（要鎖、要記 log）但免密碼 */
var RQ_SEC = 21600;
function doPost(e) {
  var p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, msg: '送來的資料不是 JSON' }); }
  var act = String(p.action || ''), fn = ACTIONS[act], res;
  if (act === 'ping') { res = { ok: true, data: pingData_() }; res.act = 'ping'; return out_(res); }   /* v2：預熱用（不用密碼、不碰試算表），回法同 GET ping */
  if (!fn) return out_({ ok: false, act: act, msg: '不支援的動作：' + act });
  var who;
  if (OPEN[act] || (act === 'bundle' && !String(p.password || '').trim())) who = '訪客';   /* v3：公告免密碼；bundle 沒帶密碼只回公告＋選單 */
  else {
    try { who = auth_(p.password, (PUB_ONLY[act] || (act === 'whoami' && p.role === 'pub')) ? DASH_PUB : DASH_VIEW); }
    catch (err) { Utilities.sleep(800); res = { ok: false, act: act, code: 'auth', msg: errMsg_(err) }; log_('', act, '', false, res.msg); return out_(res); }
  }
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
function bundle_(p, who) {
  /* v4（2026-10-06 晚）：不帶密碼也回三份（工作清單免密碼）；v3 時沒帶密碼只回公告與選單 */
  return { data: { list: list_(p, who).data, work: workList_(p, who).data, options: options_(p, who).data } };
}
/* v3（2026-10-06）：今天在有效期內的公告（免密碼；決策中心店長頁用）。置頂優先，再依開始日新到舊；附檔清單一起回（只讀有附檔時才讀附檔表） */
function active_(p, who) {
  var t = today_(), rows = listRows_().rows.filter(function (r) { return (!r.start_date || r.start_date <= t) && (!r.end_date || r.end_date >= t); });
  rows.sort(function (a, b) { return (b.pinned ? 1 : 0) - (a.pinned ? 1 : 0) || String(b.start_date || '').localeCompare(String(a.start_date || '')) || String(b.created_at || '').localeCompare(String(a.created_at || '')); });
  var att = {};
  rows.forEach(function (r) { att[r.ann_id] = []; });
  if (rows.some(function (r) { return r.att_count > 0; })) {
    load_('fact_announce_att').rows.forEach(function (x) { if (att[x.ann_id] && x.deleted !== 'Y') att[x.ann_id].push({ name: x.file_name, url: x.url }); });
  }
  return { data: { today: t, at: now_(), rows: rows.map(function (r) {
    return { ann_id: r.ann_id, title: r.title, category: r.category, type: r.type, start_date: r.start_date, end_date: r.end_date, pinned: r.pinned, att_count: r.att_count, text: r.text, att: att[r.ann_id] || [] };
  }) } };
}

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
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (se) { }   /* 單檔「知道連結可檢視」；v1.1：資料夾本身已是「知道連結」時 Drive 不准單檔設得更窄（存取遭拒），略過即可，檔案照樣能經資料夾權限開啟 */
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
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (se) { }   /* 同 upload_：被拒就略過 */
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
/* 一次性（2026-10-04）：NUEIP 匯入時 v1 的 setSharing 被拒，139 個附檔已存進「附檔/<公告ID>/」但沒登記。
   本函式只補寫 fact_announce_att 與 att_count，不搬、不刪任何檔案；可重跑（已登記的 Drive 檔會略過）。在編輯器執行。 */
function fixAttIndex20261004() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var t = load_('fact_announce', true), at = load_('fact_announce_att');
    var byId = {}; t.rows.forEach(function (r) { if (r.deleted !== 'Y') byId[r.ann_id] = r; });
    var have = {}; at.rows.forEach(function (a) { if (a.drive_file_id) have[a.drive_file_id] = 1; });
    var before = at.rows.length, add = [], skipped = 0, orphan = [], folders = 0;
    var it = sub_('附檔').getFolders();
    while (it.hasNext()) {
      var fd = it.next(), id = fd.getName(), row = byId[id], fs = fd.getFiles(); folders++;
      while (fs.hasNext()) {
        var f = fs.next();
        if (!row) { orphan.push(id + '/' + f.getName()); continue; }
        if (have[f.getId()]) { skipped++; continue; }
        var made = f.getDateCreated();
        add.push({ att_id: 'F' + Utilities.formatDate(made, TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100), ann_id: id, file_name: f.getName(),
          mime: f.getMimeType(), size: f.getSize(), drive_file_id: f.getId(), url: f.getUrl(), src_url: row.nueip_id ? 'nueip:' + row.nueip_id : '',
          uploaded_at: Utilities.formatDate(made, TZ, 'yyyy-MM-dd HH:mm:ss'), deleted: '' });
        have[f.getId()] = 1;
      }
    }
    if (add.length) {
      var r0 = nextRow_(at), H = at.H;
      if (r0 + add.length - 1 > at.sh.getMaxRows()) at.sh.insertRowsAfter(at.sh.getMaxRows(), add.length + 50);
      var rg = at.sh.getRange(r0, 1, add.length, H.length);
      rg.setNumberFormats(add.map(function () { return H.map(function (h) { return NUM_COLS[h] ? '0' : '@'; }); }));
      rg.setValues(add.map(function (o) { return H.map(function (h) { return cell_(h, o[h]); }); }));
    }
    var cnt = {}; load_('fact_announce_att').rows.forEach(function (a) { if (a.deleted !== 'Y') cnt[a.ann_id] = (cnt[a.ann_id] || 0) + 1; });
    var fixed = 0;
    t.rows.forEach(function (r) {
      var n = cnt[r.ann_id] || 0;
      if ((Number(r.att_count) || 0) !== n) { putCols_(t, r._row, { att_count: n }, function (h) { return h === 'att_count'; }); fixed++; }
    });
    var msg = '附檔資料夾 ' + folders + ' 個；補登 ' + add.length + ' 個（已登記略過 ' + skipped + '）；附檔清單 ' + before + ' → ' + (before + add.length) + ' 列；更新附檔數 ' + fixed + ' 則；找不到公告的檔案 ' + orphan.length;
    log_('補登程式', 'fixAttIndex', '', true, msg);
    Logger.log(msg + (orphan.length ? '\n' + orphan.join('\n') : ''));
    return msg;
  } finally { lock.releaseLock(); }
}

function workView_(r) {
  return { work_id: r.work_id, src: r.src, nueip_id: r.nueip_id, item: r.item, owner: r.owner, creator: r.creator, date: r.date, due_date: r.due_date, done_date: r.done_date,
    progress: r.progress, importance: r.importance, origin: r.origin, text: r.content_text, html: str_(r.content_html_1) + str_(r.content_html_2), updated_at: r.updated_at };
}
function workList_(p, who) {
  /* v5.2：結果暫存 5 分鐘（同公告清單的 ann:ver 版本號；寫入後 listCacheClear_ 換號＝失效）；放不進去就照常每次讀 */
  var cache = CacheService.getScriptCache(), key = 'work:list:' + (cache.get('ann:ver') || '0'), hit = null;
  try { hit = cacheGetBig_(cache, key); } catch (e) { hit = null; }
  if (hit) { try { return { data: JSON.parse(hit) }; } catch (e) { /* 壞掉就重讀 */ } }
  var t = load_('fact_work'), d = load_('fact_work_discuss'), by = discBy_(d.rows);
  var data = { rows: t.rows.filter(function (r) { return r.deleted !== 'Y' && r.work_id; }).map(function (r) { var o = workView_(r); o.discuss = by[r.work_id] || []; return o; }) };
  try { cachePutBig_(cache, key, JSON.stringify(data), LIST_TTL); } catch (e) { }
  return { data: data };
}
/* v5：留言 → 給前端的樣子（附檔 JSON 解開；不回 Drive 檔案 ID 以外的內部欄位） */
function parseAtt_(s) { try { var a = JSON.parse(String(s || '[]')); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function discView_(x) {
  return { disc_id: x.disc_id || '', author: x.author, role: x.role || '', ts: x.ts, text: x.text, src: x.src || '',
    att: parseAtt_(x.att).map(function (a) { return { name: a.name, size: Number(a.size) || 0, mime: a.mime || '', url: a.url }; }) };
}
/* 依工作編號分組、不含已刪除、時間舊到新（NUEIP 舊討論的時間只有到分或到日，字串比照樣正確） */
function discBy_(rows) {
  var by = {};
  rows.forEach(function (x) { if (x.deleted === 'Y' || !x.work_id) return; (by[x.work_id] = by[x.work_id] || []).push(x); });
  Object.keys(by).forEach(function (k) {
    by[k] = by[k].map(function (x, i) { return { x: x, i: i }; })
      .sort(function (a, b) { return String(a.x.ts).localeCompare(String(b.x.ts)) || a.i - b.i; })
      .map(function (o) { return discView_(o.x); });
  });
  return by;
}
/* v5 防灌：CacheService 計數（呼叫端已拿到腳本鎖，不會互搶）；超過就擋 */
function rate_(key, max, sec, msg) {
  var c = CacheService.getScriptCache(), n = parseInt(c.get(key) || '0', 10) || 0;
  if (n >= max) throw fail_(msg, 'rate');
  c.put(key, String(n + 1), sec);
}
/* v5（2026-10-07）：工作留言（免密碼；104 指定員工公告裡的「點這裡留言」連過來）。要填名字；身分（店別／職稱）可空白。
   附檔：每則 ≤ DISC_MAX_FILES 個、每個 ≤10MB、合計 ≤20MB、副檔名要在 DISC_EXT；存「工作討論附檔/<工作編號>/」（只在根資料夾底下建）。
   先全部檢查完才建檔，避免檢查到一半失敗留下孤兒檔。 */
function discussAdd_(p, who) {
  var id = str_(p.work_id), w = load_('fact_work', true), row = w.rows.filter(function (x) { return x.work_id === id && x.deleted !== 'Y'; })[0];
  if (!id || !row) throw fail_('找不到這件工作（可能已刪除）', 'notfound');
  var author = cleanTxt_(p.author, 30).replace(/\s+/g, ' ').trim();
  if (!author) throw fail_('請填寫你的名字', 'invalid');
  var role = cleanTxt_(p.role, 30).replace(/\s+/g, ' ').trim();
  var text = cleanTxt_(p.text, DISC_TEXT_MAX + 50).replace(/\r\n?/g, '\n').trim();
  if (text.length > DISC_TEXT_MAX) throw fail_('留言最多 ' + DISC_TEXT_MAX + ' 字', 'invalid');
  var files = Array.isArray(p.files) ? p.files : [];
  if (!text && !files.length) throw fail_('請輸入留言內容或附上檔案', 'invalid');
  if (files.length > DISC_MAX_FILES) throw fail_('每則留言最多 ' + DISC_MAX_FILES + ' 個檔案', 'invalid');
  var blobs = [], total = 0;
  files.forEach(function (f) {
    var name = cleanTxt_(f && f.name, 150).replace(/[\\\/:*?"<>|]/g, '_').trim() || 'file', m = name.match(/\.([A-Za-z0-9]{1,5})$/), ext = m ? m[1].toLowerCase() : '';
    if (DISC_EXT.indexOf(ext) < 0) throw fail_('「' + name + '」的檔案類型不收（可上傳：' + DISC_EXT.join('、') + '）', 'invalid');
    var bytes = Utilities.base64Decode(String((f && f.b64) || ''));
    if (!bytes.length) throw fail_('「' + name + '」是空的', 'invalid');
    if (bytes.length > DISC_MAX_FILE) throw fail_('「' + name + '」超過 10MB', 'invalid');
    total += bytes.length;
    if (total > DISC_MAX_TOTAL) throw fail_('這次附的檔案合計超過 20MB，請分成幾則留言', 'invalid');
    blobs.push({ blob: Utilities.newBlob(bytes, cleanTxt_(f.mime, 100) || 'application/octet-stream', name), size: bytes.length });
  });
  rate_('disc:w:' + id, 10, 60, '這件工作 1 分鐘內留言太多次，請稍後再送');
  rate_('disc:all', 30, 60, '全站 1 分鐘內留言太多，請稍後再送');
  rate_('disc:day:' + today_(), 500, 86400, '今天的留言量已達上限（500 則），請聯絡管理者');
  var att = [];
  if (blobs.length) {
    var folder = subOf_(sub_('工作討論附檔'), id);
    blobs.forEach(function (b) {
      var f = folder.createFile(b.blob);
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (se) { }   /* 同 upload_：資料夾已是「知道連結」時會被拒，略過 */
      att.push({ name: b.blob.getName(), size: b.size, mime: b.blob.getContentType(), id: f.getId(), url: f.getUrl() });
    });
  }
  var d = load_('fact_work_discuss');
  var o = { work_id: id, author: author, ts: now_(), text: text, disc_id: 'D' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100),
    role: role, att: att.length ? JSON.stringify(att) : '', src: 'web', deleted: '', deleted_at: '' };
  put_(d, nextRow_(d), o);
  d.rows.push(o);
  return { id: id, data: { disc_id: o.disc_id, discuss: discBy_(d.rows)[id] || [] },
    msg: '留言 ' + o.disc_id + '：' + author + (role ? '（' + role + '）' : '') + '｜' + row.item + (att.length ? '｜附檔 ' + att.length + ' 個' : '') };
}
/* v5 一次性升級（2026-10-07，在編輯器執行；可重跑）：
   ①先把 fact_work_discuss 整張複製成「backup_fact_work_discuss_20261007」（已存在就不再複製）
   ②表頭第 5～10 欄補 disc_id／role／att／src／deleted／deleted_at（原 4 欄不動；已有字且不同就停）
   ③舊討論補留言編號 N＋列號、來源 nueip（只填空白格，已有的不改）④建「工作討論附檔」資料夾 */
function setupV5_20261007() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = ss_(), name = 'fact_work_discuss', sh = ss.getSheetByName(name), H = TABS[name];
    if (!sh) throw new Error('找不到分頁 ' + name);
    var bk = 'backup_' + name + '_20261007', before = Math.max(0, sh.getLastRow() - 1);
    if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
    var cur = sh.getRange(1, 1, 1, H.length).getValues()[0].map(function (x) { return String(x).trim(); });
    for (var i = 0; i < H.length; i++) {
      if (!cur[i]) sh.getRange(1, i + 1).setValue(H[i]).setFontWeight('bold').setBackground('#FFF3E0');
      else if (cur[i] !== H[i]) throw new Error(name + ' 表頭第 ' + (i + 1) + ' 欄是「' + cur[i] + '」，應為「' + H[i] + '」，停止升級');
    }
    ensureTab_(ss, name);
    var n = 0, last = sh.getLastRow();
    if (last > 1) {
      var ci = H.indexOf('disc_id') + 1, cs = H.indexOf('src') + 1;
      var ids = sh.getRange(2, ci, last - 1, 1).getValues(), src = sh.getRange(2, cs, last - 1, 1).getValues(), wid = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var r = 0; r < ids.length; r++) {
        if (!String(wid[r][0]).trim()) continue;
        if (!String(ids[r][0]).trim()) { ids[r][0] = 'N' + ('00000' + (r + 2)).slice(-6); n++; }
        if (!String(src[r][0]).trim()) src[r][0] = 'nueip';
      }
      sh.getRange(2, ci, last - 1, 1).setNumberFormat('@').setValues(ids);
      sh.getRange(2, cs, last - 1, 1).setNumberFormat('@').setValues(src);
    }
    sub_('工作討論附檔');
    var after = Math.max(0, sh.getLastRow() - 1);
    try { CacheService.getScriptCache().remove('stat'); } catch (ce) { }
    var msg = name + ' 升級完成：表頭 ' + H.length + ' 欄；舊討論補留言編號 ' + n + ' 列；資料列 ' + before + ' → ' + after + '（應相同）；備份分頁 ' + bk;
    log_('升級程式', 'setupV5', '', true, msg); Logger.log(msg); return msg;
  } finally { lock.releaseLock(); }
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
