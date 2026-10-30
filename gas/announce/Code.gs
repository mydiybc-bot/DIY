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
var VERSION = 'announce-api-v8';   // 2026-10-30 v8：門市雜記（noteList／noteAdd／noteDel，店密碼；可附圖片）；2026-10-22 v7.1：加一次性 closeWorkDue20261015（期限 10/15 以前的工作改已完成），API 動作未改；2026-10-17 v7：三頁合併成「自己做決策中心」——總部密碼（task_hq）可看任何一店、交辦全部、發佈公告／工作；店密碼（store_N）可交辦自己店；v6.1 2026-10-13：團體名單改讀近期訂位（打開就更新）；v6   // 2026-10-13 v6：門市今日（交辦事項＋團體聯絡資料，各店密碼）；
// 舊註解：v5.2   // 2026-10-07 v5.2：工作清單暫存 5 分鐘（加快 104 連過來）；v5.1：留言誰都不能刪（拿掉 discussDel）；v5：工作清單留言 discussAdd（免密碼、可附檔）；v4：工作清單 workList 也免密碼、bundle 不帶密碼回三份；v3：公告免密碼（list／get／options／active），工作清單仍要密碼；v2＝2026-10-04 併入 v1.1 修正（setSharing 被拒略過＋fixAttIndex20261004）
var AUTH_SEC = 7200, LIST_TTL = 300;
var TZ = 'Asia/Taipei';
var ROOT_FOLDER_ID = '1MnFAKso03ERa8zSj9z0_ytDUVoF66JYg';
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var DASH_VIEW = 'announce', DASH_PUB = 'announce_pub', DASH_HQ = 'task_hq';   /* v7：DASH_HQ＝決策中心「總部」密碼 */
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
    var needPub = PUB_ONLY[act] || (act === 'whoami' && p.role === 'pub');
    /* v7：決策中心用「總部」身分登入時帶 asHq，發佈／修改／刪除改驗總部密碼（task_hq）；沒帶照舊驗 announce_pub */
    try { if (needPub && p.asHq) { auth_(p.password, DASH_HQ); who = '發佈者（總部）'; } else who = auth_(p.password, needPub ? DASH_PUB : DASH_VIEW); }
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

/* =====================================================================
 * v6（2026-10-13，經營者「門市今日」：每店開店第一眼看的頁面 static/store-today.html）
 *   ① 交辦事項：店長／區主管／總部都能交辦；週期＝指定日期／每天／每週／每月／每季／每半年；夥伴打勾＋寫名字＋處理狀況。
 *   ② 團體訂位聯絡資料（姓名、電話、人數…）：只在「店密碼」驗過後才回（讀「DIYBC 團體訂位」gb_bookings）。
 *   密碼都在「儀表板權限中心」dim_auth：
 *     store_1～store_12＝各店夥伴密碼（看頁面、打勾）；task_store_1～12＝店長交辦密碼（只管自己的店）；
 *     task_zone_n1／n2／n3／cs＝區主管（北一／北二／北三／中南）；task_hq＝總部（全部 12 店）。
 *   storeAuthSetup：只補「還沒有的列」（密碼空白＝還不能用，經營者在權限管理後台設好密碼才生效）；第一次會先備份 dim_auth 分頁。
 * ===================================================================== */
var GB_SS_ID = '1NrMGw6BOvMJigKfdAFNM1Np6KCa5cd6SEIBqzrmFLAQ', GB_TAB = 'gb_bookings';
var AUTHC_SS_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4', AUTHC_TAB = 'dim_auth';
/* 店號｜標準店名｜短名｜區代號｜團體訂位表的店名（與決策中心 dim_store_map「訂位」欄一致） */
var STORE_DEF = [
  [1, '台中精明店', '精明', 'cs', '台中精明店'], [2, '台中草悟道店', '草悟道', 'cs', '台中草悟道店'],
  [3, '台北南京店', '南京', 'n1', '台北南京店'], [4, '台北士林店', '士林', 'n1', '台北士林店'],
  [5, '台南Focus店', 'Focus', 'cs', '台南Focus店'], [6, '新竹文化店', '新竹', 'cs', '新竹文化店'],
  [7, '新北板橋店', '板橋', 'n2', '新北板橋店'], [8, '新北新店店', '新店', 'n1', '新北新店店'],
  [9, '桃園中壢店', '中壢', 'n3', '桃園中壢店'], [10, '桃園藝文店', '藝文', 'n3', '桃園藝文店'],
  [11, '台北遠百信義A13店', 'A13', 'n2', '台北遠百信義A13店'], [12, '高雄SKM Park店', 'SKM', 'cs', '高雄SKM Park店']
];
var ZONE_DEF = { n1: '北一區', n2: '北二區', n3: '北三區', cs: '中南區' };
var TASK_FREQ = { once: '指定日期', daily: '每天', weekly: '每週', monthly: '每月', quarterly: '每季', halfyear: '每半年' };
var DONE_TAIL = 30000;   /* 打勾紀錄只讀最後 3 萬列（約 4～6 個月；週期最長半年，超過的早就換下一次了） */
var TASK_TITLE_MAX = 40, TASK_NOTE_MAX = 300, DONE_NOTE_MAX = 200, WHO_MAX = 20, GROUP_DAYS = 14;
TABS.dim_task = ['task_id', 'stores', 'title', 'note', 'freq', 'rule', 'start_date', 'end_date', 'creator_role', 'creator_name',
  'created_at', 'updated_at', 'updated_by', 'deleted', 'deleted_at'];
TABS.fact_task_done = ['task_id', 'store', 'occ', 'who', 'note', 'ts', 'deleted', 'deleted_at', 'deleted_by'];
ACTIONS.storeDay = storeDay_; ACTIONS.taskDone = taskDone_; ACTIONS.taskUndo = taskUndo_;
ACTIONS.taskAuth = taskAuth_; ACTIONS.taskList = taskList_; ACTIONS.taskSave = taskSave_; ACTIONS.taskDel = taskDel_;
ACTIONS.storeAuthSetup = storeAuthSetup_;
['storeDay', 'taskDone', 'taskUndo', 'taskAuth', 'taskList', 'taskSave', 'taskDel', 'storeAuthSetup'].forEach(function (a) { OPEN[a] = 1; });   /* 這些動作自己驗密碼（各店／各角色不同密碼），不走 announce 的密碼 */
['taskDone', 'taskUndo', 'taskSave', 'taskDel', 'storeAuthSetup'].forEach(function (a) { WRITES[a] = 1; });

function storeOf_(sid) { sid = Number(sid); for (var i = 0; i < STORE_DEF.length; i++) if (STORE_DEF[i][0] === sid) return { id: sid, name: STORE_DEF[i][1], short: STORE_DEF[i][2], zone: STORE_DEF[i][3], gb: STORE_DEF[i][4] }; throw fail_('店號不對', 'bad'); }
function zoneStores_(z) { return STORE_DEF.filter(function (s) { return s[3] === z; }).map(function (s) { return s[0]; }); }
/* v7：帶 hpw＝總部密碼（task_hq），可看任何一家店；沒帶照舊驗這家店的店密碼 */
function storeAuth_(p) {
  var st = storeOf_(p.store);
  if (String(p.hpw || '').trim()) {
    try { auth_(p.hpw, DASH_HQ); } catch (e) { Utilities.sleep(600); throw fail_(errMsg_(e).replace(/^密碼不對/, '總部密碼不對').replace('請輸入密碼', '請輸入總部密碼'), 'auth'); }
    return st;
  }
  try { auth_(p.spw, 'store_' + st.id); } catch (e) { Utilities.sleep(600); throw fail_(errMsg_(e).replace(/^密碼不對/, '店密碼不對').replace('請輸入密碼', '請輸入店密碼'), 'auth'); }
  return st;
}
/* 交辦身分：role＝store（店長，只管這家店）｜zone（區主管，管這家店所在的區）｜hq（總部，全部）。只驗選的那一種，避免猜錯把別的密碼鎖住 */
function mgrAuth_(p) {
  var st = storeOf_(p.store), role = String(p.role || ''), dash, scope, label;
  if (role === 'store') { dash = 'store_' + st.id; scope = [st.id]; label = st.short + '店'; }   /* v7：一店一組密碼（原 task_store_N 店長交辦密碼不再使用） */
  else if (role === 'zone') { dash = 'task_zone_' + st.zone; scope = zoneStores_(st.zone); label = ZONE_DEF[st.zone] + '區主管'; }
  else if (role === 'hq') { dash = 'task_hq'; scope = STORE_DEF.map(function (s) { return s[0]; }); label = '總部'; }
  else throw fail_('請選身分', 'bad');
  try { auth_(p.mpw, dash); } catch (e) { Utilities.sleep(600); throw fail_(errMsg_(e).replace(/^密碼不對/, '交辦密碼不對').replace('請輸入密碼', '請輸入交辦密碼'), 'auth'); }
  return { role: role, scope: scope, label: label, store: st };
}
function loadTail_(name, n) {
  var sh = ss_().getSheetByName(name), H = TABS[name], last = sh.getLastRow(), rows = [];
  if (last > 1) {
    var from = Math.max(2, last - n + 1), v = sh.getRange(from, 1, last - from + 1, H.length).getValues();
    v.forEach(function (r, i) { var o = { _row: from + i }; H.forEach(function (h, j) { o[h] = str_(r[j]); }); rows.push(o); });
  }
  return { sh: sh, H: H, rows: rows };
}
function ensureStoreTabs_() {
  var ss = ss_();
  ['dim_task', 'fact_task_done'].forEach(function (n) { if (!ss.getSheetByName(n)) ensureTab_(ss, n); });
}

/* ---------- 週期：算「最近一次該做的日子」（≤ 今天） ---------- */
function ymd_(y, m, d) { var dt = new Date(y, m - 1, d); return Utilities.formatDate(dt, TZ, 'yyyy-MM-dd'); }
function dim_(y, m) { return new Date(y, m, 0).getDate(); }
function dayIn_(y, m, d) { return d === 'L' ? dim_(y, m) : Math.min(Number(d), dim_(y, m)); }
function addMonths_(y, m, k) { var t = (y * 12 + (m - 1)) + k; return [Math.floor(t / 12), (t % 12) + 1]; }
function latestOcc_(t, today) {
  var y = Number(today.slice(0, 4)), m = Number(today.slice(5, 7)), rule = String(t.rule || ''), occ = '';
  if (t.freq === 'daily') occ = today;
  else if (t.freq === 'once') occ = (rule && rule <= today) ? rule : '';
  else if (t.freq === 'weekly') {
    var days = rule.split(',').map(Number).filter(function (x) { return x >= 1 && x <= 7; });
    var base = new Date(Number(today.slice(0, 4)), m - 1, Number(today.slice(8, 10)));
    for (var k = 0; k < 7 && !occ; k++) {
      var dt = new Date(base.getTime() - k * 86400000), wd = dt.getDay() || 7;
      if (days.indexOf(wd) >= 0) occ = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
    }
  } else if (t.freq === 'monthly') {
    var d0 = rule === 'L' ? 'L' : Number(rule) || 1, c = ymd_(y, m, dayIn_(y, m, d0));
    if (c <= today) occ = c; else { var pm = addMonths_(y, m, -1); occ = ymd_(pm[0], pm[1], dayIn_(pm[0], pm[1], d0)); }
  } else if (t.freq === 'quarterly' || t.freq === 'halfyear') {
    var len = t.freq === 'quarterly' ? 3 : 6, parts = rule.split('-'), mi = Math.min(Math.max(Number(parts[0]) || 1, 1), len), dd = parts[1] === 'L' ? 'L' : (Number(parts[1]) || 1);
    var ps = Math.floor((m - 1) / len) * len + 1;
    for (var back = 0; back < 3 && !occ; back++) {
      var st = addMonths_(y, ps, -len * back), tm = addMonths_(st[0], st[1], mi - 1), cand = ymd_(tm[0], tm[1], dayIn_(tm[0], tm[1], dd));
      if (cand <= today) occ = cand;
    }
  }
  if (occ && t.start_date && occ < t.start_date) occ = '';
  if (occ && t.end_date && occ > t.end_date) occ = '';
  return occ;
}
var WD_TXT = ['', '一', '二', '三', '四', '五', '六', '日'];
function taskWhen_(t) {
  var r = String(t.rule || ''), mi, dd;
  if (t.freq === 'daily') return '每天';
  if (t.freq === 'once') return r.slice(5).replace('-', '/');
  if (t.freq === 'weekly') return '每週' + r.split(',').map(function (x) { return WD_TXT[Number(x)] || ''; }).join('、');
  if (t.freq === 'monthly') return '每月' + (r === 'L' ? '月底' : r + '號');
  var parts = r.split('-'); mi = parts[0]; dd = parts[1] === 'L' ? '月底' : parts[1] + '號';
  if (t.freq === 'quarterly') return '每季第' + mi + '個月' + dd;
  return '每半年第' + mi + '個月' + dd;
}
/* 下一次要做的日子（≥ 今天；最多往後找 400 天）：給交辦清單顯示「下次：MM/DD」，新增在週期中間的交辦從下一次才開始 */
function nextOcc_(t, today) {
  var base = new Date(Number(today.slice(0, 4)), Number(today.slice(5, 7)) - 1, Number(today.slice(8, 10)));
  for (var k = 0; k < 400; k++) { var d = Utilities.formatDate(new Date(base.getTime() + k * 86400000), TZ, 'yyyy-MM-dd'); if (latestOcc_(t, d) === d) return d; }
  return '';
}
function tasksFor_(sid) {
  return load_('dim_task').rows.filter(function (t) {
    return t.deleted !== 'Y' && String(t.stores || '').split(',').map(Number).indexOf(Number(sid)) >= 0;
  });
}

/* ---------- 夥伴：今天這家店的全部資料 ---------- */
function storeDay_(p) {
  var st = storeAuth_(p); ensureStoreTabs_();
  var today = today_(), done = {};
  loadTail_('fact_task_done', DONE_TAIL).rows.forEach(function (d) { if (d.deleted !== 'Y' && Number(d.store) === st.id) done[d.task_id + '|' + d.occ] = d; });
  var tasks = [];
  tasksFor_(st.id).forEach(function (t) {
    var occ = latestOcc_(t, today); if (!occ) return;
    var d = done[t.task_id + '|' + occ];
    if (d && String(d.ts).slice(0, 10) !== today) return;   /* 之前就做完的不再顯示，只留今天打勾的讓大家看到 */
    tasks.push({ id: t.task_id, title: t.title, note: t.note, freq: t.freq, when: taskWhen_(t), occ: occ,
      st: d ? 'done' : (occ < today ? 'late' : 'today'), who: d ? d.who : '', doneNote: d ? d.note : '', doneAt: d ? String(d.ts).slice(11, 16) : '',
      by: t.creator_role === 'hq' ? '總部' : (t.creator_role === 'zone' ? '區主管' : '店長') });
  });
  var ord = { late: 0, today: 1, done: 2 };
  tasks.sort(function (a, b) { return ord[a.st] - ord[b.st] || String(a.occ).localeCompare(String(b.occ)) || String(a.title).localeCompare(String(b.title)); });
  return { data: { store: { id: st.id, name: st.name, short: st.short }, today: today, at: now_(), tasks: tasks, groups: groupsFor_(st, today) } };
}
/* 團體訂位（今天起 GROUP_DAYS 天、已取消不列）：姓名、電話、人數、訂金狀態、甜點。只回這家店。
   2026-10-13 (2)：名單改讀「DIYBC 訂位資料」近期訂位 fact_reservations_future（門市今日打開時每 10 分鐘可重抓近 7 天，比團體表一天 2 輪新），
   再從團體表 gb_bookings 補「應付訂金、備註、目的」（res_id＝團體表 id），甜點從 fact_future_menu（同一次重抓寫的）補；
   近期訂位讀不到才退回只用團體表。團體＝類別含大組／包館／包場、或 8 人以上（與團體 GAS 同規則）；
   會員是「客服訂位」且備註有「卡位」＝門市／客服卡位，不是真客人，不列。 */
var RSV_SS_ID = '13NI3vGV4MSsngeO_DecVYKrOzky-fXCsN9ActscJorQ', RSV_FUT_TAB = 'fact_reservations_future', RSV_MENU_TAB = 'fact_future_menu';
var GROUP_CAT_RE = /大組|包館|包場/, GROUP_MIN_SIZE = 8;
function hhmm_(v) { if (v instanceof Date) return Utilities.formatDate(v, TZ, 'HH:mm'); var m = String(v || '').match(/(\d{1,2}):(\d{2})/); return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : String(v || ''); }
/* 讀一張表。fixedHead＝寫入程式用的欄位順序（訂位管線照「位置」寫入，第一列標題不一定跟著改），有給就照位置對欄名，不看標題文字 */
var RSV_FUT_HEAD = ['res_id', 'date', 'store', 'slot', 'member_name', 'phone', 'line_uid', 'email', 'category', 'party_size', 'recipes', 'companions',
  'order_total', 'paid', 'status', 'attended', 'purpose', 'member_level', 'note', 'group_id', 'fetched_at'];   /* ＝訂位管線 CFG.HEADERS（改那邊要同步改這裡） */
var RSV_MENU_HEAD = ['date', 'store', 'slot', 'category', 'party_size', 'status', 'menu', 'fetched_at'];   /* ＝訂位管線 FM_HEAD */
function tabRows_(ssId, tab, fixedHead) {
  var sh = SpreadsheetApp.openById(ssId).getSheetByName(tab); if (!sh || sh.getLastRow() < 2) return { ix: {}, rows: [] };
  var v = sh.getDataRange().getValues(), ix = {};
  (fixedHead || v[0]).forEach(function (h, i) { ix[String(h).trim()] = i; });
  return { ix: ix, rows: v.slice(1) };
}
function menuList_(raw) {
  var a = raw; if (typeof a === 'string') { try { a = JSON.parse(a); } catch (e) { a = null; } }
  if (!Array.isArray(a)) return [];
  return a.map(function (x) { var n = Array.isArray(x) ? x[0] : (x && x.n), q = Array.isArray(x) ? x[1] : (x && x.q); n = String(n || '').replace(/^.*｜/, '').replace(/\s*\((葷|可做全素|全素|信義、高雄無法製作)\)/g, '').trim(); return n ? { n: n, q: Number(q) || 1 } : null; }).filter(Boolean);
}
function gbMenu_(s) { return menuList_(String(s || '').split(';').map(function (x) { var a = x.split('|'); return [a[0], a[1]]; })); }
function groupsFor_(st, today) {
  var endD = Utilities.formatDate(new Date(Date.now() + GROUP_DAYS * 86400000), TZ, 'yyyy-MM-dd');
  var gb = {}, gbList = [], gbAt = '';
  try {
    var G = tabRows_(GB_SS_ID, GB_TAB), gi = G.ix;
    G.rows.forEach(function (r) {
      if (String(r[gi.store]).trim() !== st.gb) return;
      var d = normDate_(r[gi.date]); if (!d || d < today || d > endD) return;
      var o = { id: String(r[gi.id] || ''), d: d, t: hhmm_(r[gi.slot]), name: String(r[gi.member] || ''), phone: String(r[gi.phone] || ''), size: Number(r[gi.size]) || 0,
        cat: String(r[gi.category] || ''), stTxt: String(r[gi.status_text] || ''), status: String(r[gi.status] || ''), purpose: String(r[gi.purpose] || ''), memo: String(r[gi.memo] || ''),
        menu: gbMenu_(r[gi.menu]), paid: Number(r[gi.paid]) || 0, deposit: gi.deposit === undefined ? 0 : (Number(r[gi.deposit]) || 0) };
      if (o.id) gb[o.id] = o;
      gbList.push(o);
      if (gi.scanned_at !== undefined && str_(r[gi.scanned_at]) > gbAt) gbAt = str_(r[gi.scanned_at]);
    });
  } catch (e) { }
  var rows = [], src = 'future', futAt = '';
  try {
    var F = tabRows_(RSV_SS_ID, RSV_FUT_TAB, RSV_FUT_HEAD), fi = F.ix, menus = {};
    if (fi.res_id === undefined || fi.member_name === undefined) throw new Error('近期訂位欄位不對');
    try {
      var M = tabRows_(RSV_SS_ID, RSV_MENU_TAB, RSV_MENU_HEAD), mi = M.ix;
      M.rows.forEach(function (r) { if (String(r[mi.store]).trim() === st.gb) menus[normDate_(r[mi.date]) + '|' + hhmm_(r[mi.slot]) + '|' + (Number(r[mi.party_size]) || 0)] = menuList_(r[mi.menu]); });
    } catch (eM) { }
    F.rows.forEach(function (r) {
      if (String(r[fi.store]).trim() !== st.gb) return;
      var at = str_(r[fi.fetched_at]); if (at > futAt) futAt = at;
      var d = normDate_(r[fi.date]); if (!d || d < today || d > endD) return;
      var status = String(r[fi.status] || ''); if (status === '已取消') return;
      var cat = String(r[fi.category] || ''), size = Number(r[fi.party_size]) || 0, name = String(r[fi.member_name] || ''), note = String(r[fi.note] || '');
      if (!(GROUP_CAT_RE.test(cat) || size >= GROUP_MIN_SIZE)) return;
      if (name.indexOf('客服訂位') >= 0 && /卡位/.test(note)) return;   // 客服／門市卡位
      var t = hhmm_(r[fi.slot]), g = gb[String(r[fi.res_id] || '')] || null, mk = d + '|' + t + '|' + size;
      rows.push({ d: d, t: t, name: name, phone: String(r[fi.phone] || ''), size: size, cat: cat, stTxt: status,
        purpose: String(r[fi.purpose] || '') || (g ? g.purpose : ''), memo: g ? g.memo : '', menu: menus[mk] || (g ? g.menu : []),
        paid: Number(r[fi.paid]) || 0, deposit: g ? g.deposit : 0 });
    });
  } catch (eF) {
    src = 'gb';   // 近期訂位讀不到 → 只用團體表（一天 2 輪）
    rows = gbList.filter(function (o) { return o.status !== 'canceled'; }).map(function (o) { return { d: o.d, t: o.t, name: o.name, phone: o.phone, size: o.size, cat: o.cat, stTxt: o.stTxt, purpose: o.purpose, memo: o.memo, menu: o.menu, paid: o.paid, deposit: o.deposit }; });
  }
  rows.sort(function (a, b) { return (a.d + a.t).localeCompare(b.d + b.t); });
  return { ok: true, rows: rows, src: src, at: src === 'future' ? futAt : gbAt };
}
function cleanWho_(s) { var w = cleanTxt_(s, WHO_MAX).trim(); if (!w) throw fail_('請寫名字', 'bad'); return w; }
function taskDone_(p) {
  var st = storeAuth_(p); ensureStoreTabs_();
  var who = cleanWho_(p.who), note = cleanTxt_(p.note, DONE_NOTE_MAX).trim(), today = today_();
  rate_('td:' + st.id, 40, 60, '按太快了，請等一下再按');
  var t = tasksFor_(st.id).filter(function (x) { return x.task_id === String(p.task_id); })[0];
  if (!t) throw fail_('這件交辦已經刪除了，請重新整理', 'gone');
  var occ = latestOcc_(t, today); if (!occ) throw fail_('這件今天不用做', 'bad');
  var tb = loadTail_('fact_task_done', DONE_TAIL);
  var dup = tb.rows.filter(function (d) { return d.deleted !== 'Y' && d.task_id === t.task_id && Number(d.store) === st.id && d.occ === occ; })[0];
  if (dup) return { id: t.task_id, msg: st.short + ' 已有人打勾（' + dup.who + '）', data: { already: true, who: dup.who } };
  put_(tb, nextRow_(tb), { task_id: t.task_id, store: String(st.id), occ: occ, who: who, note: note, ts: now_(), deleted: '' });
  return { id: t.task_id, msg: st.short + '｜' + who + '｜' + t.title, data: { ok: true } };
}
function taskUndo_(p) {
  var st = storeAuth_(p); ensureStoreTabs_();
  var who = cleanWho_(p.who), tb = loadTail_('fact_task_done', DONE_TAIL), today = today_(), hit = null;
  tb.rows.forEach(function (d) { if (d.deleted !== 'Y' && d.task_id === String(p.task_id) && Number(d.store) === st.id && d.occ === String(p.occ) && String(d.ts).slice(0, 10) === today) hit = d; });
  if (!hit) throw fail_('只能取消今天打的勾', 'bad');
  putCols_(tb, hit._row, { deleted: 'Y', deleted_at: now_(), deleted_by: who }, function (h) { return h === 'deleted' || h === 'deleted_at' || h === 'deleted_by'; });
  return { id: hit.task_id, msg: st.short + '｜取消打勾｜' + who, data: { ok: true } };
}

/* ---------- 交辦管理（店長／區主管／總部） ---------- */
function taskAuth_(p) { var a = mgrAuth_(p); return { data: { role: a.role, label: a.label, scope: a.scope, stores: STORE_DEF.filter(function (s) { return a.scope.indexOf(s[0]) >= 0; }).map(function (s) { return { id: s[0], short: s[2] }; }) } }; }
function taskView_(t) {
  return { id: t.task_id, stores: String(t.stores || '').split(',').map(Number).filter(Boolean), title: t.title, note: t.note, freq: t.freq, rule: t.rule,
    when: taskWhen_(t), next: nextOcc_(t, today_()), start: t.start_date, end: t.end_date, role: t.creator_role, by: t.creator_name, at: String(t.updated_at || t.created_at || '').slice(0, 16) };
}
function canEdit_(a, t) { return String(t.stores || '').split(',').map(Number).every(function (s) { return a.scope.indexOf(s) >= 0; }); }
function taskList_(p) {
  var a = mgrAuth_(p); ensureStoreTabs_();
  var rows = load_('dim_task').rows.filter(function (t) {
    return t.deleted !== 'Y' && String(t.stores || '').split(',').map(Number).some(function (s) { return a.scope.indexOf(s) >= 0; });
  }).map(function (t) { var v = taskView_(t); v.canEdit = canEdit_(a, t); return v; });
  rows.sort(function (x, y) { return String(y.at).localeCompare(String(x.at)); });
  return { data: { rows: rows } };
}
function checkRule_(freq, rule) {
  rule = String(rule || '').trim();
  if (!TASK_FREQ[freq]) throw fail_('請選多久做一次', 'bad');
  if (freq === 'daily') return '';
  if (freq === 'once') { if (!/^\d{4}-\d{2}-\d{2}$/.test(rule)) throw fail_('請選日期', 'bad'); return rule; }
  if (freq === 'weekly') { var w = rule.split(',').map(Number).filter(function (x) { return x >= 1 && x <= 7; }); if (!w.length) throw fail_('請選星期幾', 'bad'); return w.sort().join(','); }
  if (freq === 'monthly') { if (rule === 'L') return 'L'; var d = Number(rule); if (!(d >= 1 && d <= 31)) throw fail_('請選幾號', 'bad'); return String(d); }
  var parts = rule.split('-'), len = freq === 'quarterly' ? 3 : 6, mi = Number(parts[0]), dd = parts[1];
  if (!(mi >= 1 && mi <= len)) throw fail_('請選第幾個月', 'bad');
  if (dd !== 'L') { dd = Number(dd); if (!(dd >= 1 && dd <= 31)) throw fail_('請選幾號', 'bad'); }
  return mi + '-' + dd;
}
function taskSave_(p) {
  var a = mgrAuth_(p); ensureStoreTabs_();
  var who = cleanWho_(p.who), title = cleanTxt_(p.title, TASK_TITLE_MAX).trim(), note = cleanTxt_(p.note, TASK_NOTE_MAX).trim();
  if (!title) throw fail_('請寫要做什麼', 'bad');
  var freq = String(p.freq || ''), rule = checkRule_(freq, p.rule);
  var stores = (Array.isArray(p.stores) ? p.stores : String(p.stores || '').split(',')).map(Number).filter(function (s) { return a.scope.indexOf(s) >= 0; });
  stores = stores.filter(function (s, i) { return stores.indexOf(s) === i; }).sort(function (x, y) { return x - y; });
  if (!stores.length) throw fail_('請選要交辦給哪幾家店', 'bad');
  var start = /^\d{4}-\d{2}-\d{2}$/.test(String(p.start || '')) ? String(p.start) : today_();
  if (freq === 'once') start = rule;   /* 指定日期：開始日就是那天（不會因為今天才建而被擋掉） */
  var end = /^\d{4}-\d{2}-\d{2}$/.test(String(p.end || '')) ? String(p.end) : '';
  if (end && end < start) throw fail_('結束日不能早於開始日', 'bad');
  var tb = load_('dim_task'), id = String(p.id || '');
  if (id) {
    var t = tb.rows.filter(function (x) { return x.task_id === id && x.deleted !== 'Y'; })[0];
    if (!t) throw fail_('這件交辦已經刪除了', 'gone');
    if (!canEdit_(a, t)) throw fail_('這件是更上層交辦的，你不能改', 'deny');
    if (p.base && String(p.base) !== String(t.updated_at)) throw fail_('別人剛改過這件，請重新整理再改', 'conflict');
    putCols_(tb, t._row, { stores: stores.join(','), title: title, note: note, freq: freq, rule: rule, start_date: start, end_date: end, updated_at: now_(), updated_by: a.label + ' ' + who },
      function (h) { return ['stores', 'title', 'note', 'freq', 'rule', 'start_date', 'end_date', 'updated_at', 'updated_by'].indexOf(h) >= 0; });
    return { id: id, msg: a.label + '｜' + who + '｜修改｜' + title, data: { id: id } };
  }
  rate_('ts:' + a.role + ':' + a.store.id, 30, 600, '新增太多了，請 10 分鐘後再試');
  id = 'T' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
  put_(tb, nextRow_(tb), { task_id: id, stores: stores.join(','), title: title, note: note, freq: freq, rule: rule, start_date: start, end_date: end,
    creator_role: a.role, creator_name: a.label + ' ' + who, created_at: now_(), updated_at: now_(), updated_by: a.label + ' ' + who, deleted: '' });
  return { id: id, msg: a.label + '｜' + who + '｜新增｜' + title + '｜' + stores.join(','), data: { id: id } };
}
function taskDel_(p) {
  var a = mgrAuth_(p); ensureStoreTabs_();
  var who = cleanWho_(p.who), tb = load_('dim_task'), t = tb.rows.filter(function (x) { return x.task_id === String(p.id) && x.deleted !== 'Y'; })[0];
  if (!t) throw fail_('這件交辦已經刪除了', 'gone');
  if (!canEdit_(a, t)) throw fail_('這件是更上層交辦的，你不能刪', 'deny');
  putCols_(tb, t._row, { deleted: 'Y', deleted_at: now_(), updated_by: a.label + ' ' + who }, function (h) { return h === 'deleted' || h === 'deleted_at' || h === 'updated_by'; });
  return { id: t.task_id, msg: a.label + '｜' + who + '｜刪除｜' + t.title, data: { ok: true } };
}

/* ---------- 一次性設定：權限中心 dim_auth 補列（只補沒有的；密碼空白，經營者在權限管理後台設定） ---------- */
function storeAuthSetup_() {
  var ss = SpreadsheetApp.openById(AUTHC_SS_ID), sh = ss.getSheetByName(AUTHC_TAB);
  if (!sh) throw fail_('找不到權限中心 dim_auth', 'setup');
  var want = [];
  STORE_DEF.forEach(function (s) { want.push(['store_' + s[0], '門市今日｜' + s[2] + '（夥伴）', 720, '門市今日頁：看頁面、打勾']); });
  STORE_DEF.forEach(function (s) { want.push(['task_store_' + s[0], '門市今日｜' + s[2] + '店長交辦', 12, '只能交辦給自己的店']); });
  Object.keys(ZONE_DEF).forEach(function (z) { want.push(['task_zone_' + z, '門市今日｜' + ZONE_DEF[z] + '區主管交辦', 12, '交辦給 ' + zoneStores_(z).join('、') + ' 號店']); });
  want.push(['task_hq', '門市今日｜總部交辦', 12, '交辦給全部 12 店']);
  var last = sh.getLastRow(), have = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = 1; });
  var add = want.filter(function (w) { return !have[w[0]]; });
  if (!add.length) return { msg: '不用補', data: { added: 0 } };
  var bk = 'dim_auth_備份_' + stamp_();
  if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
  var rows = add.map(function (w) { return [w[0], w[1], '', w[2], 'Y', new Date(), w[3] + '（2026-10-13 補；密碼空白＝還不能用，請在權限管理後台設定）']; });
  sh.getRange(last + 1, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(last + 1, 3, rows.length, 1).setNumberFormat('@');
  sh.getRange(last + 1, 1, rows.length, 7).setValues(rows);
  return { msg: '補 ' + rows.length + ' 列；備份 ' + bk, data: { added: rows.length, backup: bk, before: last - 1 } };
}
/* v7.1 一次性（2026-10-22，經營者「期限在 2026/10/15 以前的工作清單都調整為已完成」；在編輯器執行，可重跑）：
   ①先把 fact_work 整張複製成「backup_fact_work_20261022」（已存在就不再複製）
   ②沒刪除、期限有填且 ≤ 2026-10-15、還沒完成（進度不是已完成且完成日空白）的工作 → 進度「已完成」、完成日 2026-10-22（批次調整日）
   ③只改 progress／done_date／updated_at／updated_by 四欄；資料列數前後要相同 ④清工作清單暫存（不然最多 5 分鐘後才看到） */
function closeWorkDue20261015() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = ss_(), name = 'fact_work', sh = ss.getSheetByName(name);
    if (!sh) throw new Error('找不到分頁 ' + name);
    var bk = 'backup_' + name + '_20261022', before = Math.max(0, sh.getLastRow() - 1);
    if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
    var t = load_(name, true), CUT = '2026-10-15', DONE = '2026-10-22', who = '批次調整（經營者 10/22 交辦）', ids = [];
    var isOpen = function (w) { return w.deleted !== 'Y' && w.progress !== '已完成' && !w.done_date; };
    var openBefore = t.rows.filter(isOpen).length;
    t.rows.forEach(function (w) {
      if (!isOpen(w) || !w.due_date || w.due_date > CUT) return;
      putCols_(t, w._row, { progress: '已完成', done_date: DONE, updated_at: now_(), updated_by: who }, function (h) { return h === 'progress' || h === 'done_date' || h === 'updated_at' || h === 'updated_by'; });
      ids.push(w.work_id);
    });
    var after = Math.max(0, sh.getLastRow() - 1), openAfter = load_(name, true).rows.filter(isOpen).length;
    listCacheClear_(); try { CacheService.getScriptCache().remove('stat'); } catch (ce) { }
    var msg = name + ' 批次改已完成 ' + ids.length + ' 件（期限 ≤ ' + CUT + '）；未完成 ' + openBefore + ' → ' + openAfter + '；資料列 ' + before + ' → ' + after + '（應相同）；備份分頁 ' + bk;
    log_('批次程式', 'closeWorkDue20261015', ids.join(',').slice(0, 300), true, msg); Logger.log(msg); Logger.log(ids.join('\n')); return msg;
  } finally { lock.releaseLock(); }
}


/* ================= v8（2026-10-30）門市雜記 =================
   決策中心「📝 門市雜記」分頁：夥伴自由記事＋上傳圖片。全部用店密碼（store_N）或總部密碼（hpw）驗，只看得到、寫得到自己那家店。
   - 分頁 fact_store_note（第一次用到時自動建）：只新增；刪除＝標記 deleted（列不真刪、圖片不刪，留在「門市雜記/<店號>/」）。
   - 圖片：每則 ≤ NOTE_MAX_FILES 張、每張 ≤10MB、合計 ≤20MB，只收圖片副檔名；檔案設「知道連結可檢視」，畫面用 Drive 縮圖網址顯示。
   - noteList 回最近 NOTE_LIST_MAX 則（讀最後 NOTE_TAIL 列）。 */
TABS.fact_store_note = ['note_id', 'store', 'author', 'text', 'att', 'ts', 'deleted', 'deleted_at', 'deleted_by'];
var NOTE_TEXT_MAX = 2000, NOTE_MAX_FILES = 6, NOTE_TAIL = 20000, NOTE_LIST_MAX = 300;
var NOTE_EXT = ['jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif'];
ACTIONS.noteList = noteList_; ACTIONS.noteAdd = noteAdd_; ACTIONS.noteDel = noteDel_;
['noteList', 'noteAdd', 'noteDel'].forEach(function (a) { OPEN[a] = 1; });   /* 自己驗店密碼 */
['noteAdd', 'noteDel'].forEach(function (a) { WRITES[a] = 1; });
function noteTab_() { var ss = ss_(); if (!ss.getSheetByName('fact_store_note')) ensureTab_(ss, 'fact_store_note'); }
function noteOut_(n) { var att = []; try { att = n.att ? JSON.parse(n.att) : []; } catch (e) { att = []; } return { id: n.note_id, author: n.author, text: n.text, ts: n.ts, att: att }; }
function noteList_(p) {
  var st = storeAuth_(p); noteTab_();
  var rows = loadTail_('fact_store_note', NOTE_TAIL).rows.filter(function (n) { return Number(n.store) === st.id && n.deleted !== 'Y'; });
  rows.sort(function (a, b) { return a.ts < b.ts ? 1 : (a.ts > b.ts ? -1 : 0); });
  return { data: { store: { id: st.id, short: st.short }, at: now_(), total: rows.length, notes: rows.slice(0, NOTE_LIST_MAX).map(noteOut_) } };
}
function noteAdd_(p) {
  var st = storeAuth_(p); noteTab_();
  var author = cleanTxt_(p.author, 30).replace(/\s+/g, ' ').trim();
  if (!author) throw fail_('請填寫你的名字', 'invalid');
  var text = cleanTxt_(p.text, NOTE_TEXT_MAX + 50).replace(/\r\n?/g, '\n').trim();
  if (text.length > NOTE_TEXT_MAX) throw fail_('每則最多 ' + NOTE_TEXT_MAX + ' 字', 'invalid');
  var files = Array.isArray(p.files) ? p.files : [];
  if (!text && !files.length) throw fail_('請寫一點內容或附上圖片', 'invalid');
  if (files.length > NOTE_MAX_FILES) throw fail_('每則最多 ' + NOTE_MAX_FILES + ' 張圖片', 'invalid');
  var blobs = [], total = 0;
  files.forEach(function (f, i) {
    var name = cleanTxt_(f && f.name, 150).replace(/[\\\/:*?"<>|]/g, '_').trim() || ('photo' + (i + 1) + '.jpg'), m = name.match(/\.([A-Za-z0-9]{1,5})$/), ext = m ? m[1].toLowerCase() : '';
    if (NOTE_EXT.indexOf(ext) < 0) throw fail_('「' + name + '」不是圖片（可上傳：' + NOTE_EXT.join('、') + '）', 'invalid');
    var bytes = Utilities.base64Decode(String((f && f.b64) || ''));
    if (!bytes.length) throw fail_('「' + name + '」是空的', 'invalid');
    if (bytes.length > DISC_MAX_FILE) throw fail_('「' + name + '」超過 10MB', 'invalid');
    total += bytes.length;
    if (total > DISC_MAX_TOTAL) throw fail_('這次的圖片合計超過 20MB，請分成幾則', 'invalid');
    blobs.push({ blob: Utilities.newBlob(bytes, cleanTxt_(f.mime, 100) || 'image/jpeg', name), size: bytes.length });
  });
  rate_('note:s:' + st.id, 20, 60, '這家店 1 分鐘內記太多則，請稍後再送');
  rate_('note:all', 60, 60, '全部門市 1 分鐘內記太多則，請稍後再送');
  rate_('note:day:' + today_(), 1000, 86400, '今天的門市雜記已達上限（1000 則），請聯絡管理者');
  var att = [];
  if (blobs.length) {
    var folder = subOf_(sub_('門市雜記'), String(st.id));
    blobs.forEach(function (b) {
      var f = folder.createFile(b.blob);
      try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (se) { }
      att.push({ name: b.blob.getName(), size: b.size, mime: b.blob.getContentType(), id: f.getId(), url: f.getUrl() });
    });
  }
  var t = loadTail_('fact_store_note', 1);
  var o = { note_id: 'N' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + Math.floor(Math.random() * 900 + 100), store: String(st.id), author: author, text: text,
    att: att.length ? JSON.stringify(att) : '', ts: now_(), deleted: '', deleted_at: '', deleted_by: '' };
  put_(t, nextRow_(t), o);
  return { id: o.note_id, data: { note: noteOut_(o) }, msg: '門市雜記 ' + o.note_id + '：' + st.short + '｜' + author + (att.length ? '｜圖片 ' + att.length + ' 張' : '') };
}
function noteDel_(p) {
  var st = storeAuth_(p); noteTab_();
  var who = cleanTxt_(p.who, 30).replace(/\s+/g, ' ').trim(); if (!who) throw fail_('請填寫你的名字', 'invalid');
  var id = String(p.id || ''), t = loadTail_('fact_store_note', NOTE_TAIL), n = t.rows.filter(function (x) { return x.note_id === id && Number(x.store) === st.id; })[0];
  if (!id || !n) throw fail_('找不到這則（可能已刪除）', 'notfound');
  if (n.deleted === 'Y') return { id: id, data: { id: id }, msg: '已刪除過' };
  putCols_(t, n._row, { deleted: 'Y', deleted_at: now_(), deleted_by: who }, function (h) { return h === 'deleted' || h === 'deleted_at' || h === 'deleted_by'; });
  return { id: id, data: { id: id }, msg: '刪除門市雜記 ' + id + '：' + st.short + '｜' + who };
}
