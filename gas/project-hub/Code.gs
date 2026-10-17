/**
 * 專案推動 API（2026-10-17 新建，經營者交辦：「做個專案推動儀表板，未來所有新專案都在這裡管理，從『食材改平匙』開始」）
 * 綁在試算表「專案推動_資料庫」（1VE1dsSPk0Bl2dyTWrNo3G_oasD4pCR3ihllNedcOOT4）；前端＝static/project-hub.html。
 *
 * 讀（GET，JSONP，免密碼，跟其他儀表板一樣靠頁面閘門擋路人）：
 *   ?action=ping          版號
 *   ?action=bundle        專案、階段、待決定事項、風險
 *   ?action=spoon         平匙換算：食譜用量快照＋各食材的匙設定
 * 寫（POST，body＝JSON 字串；都要帶 pw＝權限中心 dim_auth「project」那列的密碼）：
 *   {action:'auth'}                                   只驗密碼
 *   {action:'save', table, row}                      新增／修改一列（沒有 id＝新增）
 *   {action:'del', table, id}                        刪除（標記 deleted=Y，不真的刪）
 *   {action:'spoonSave', ingredient, g_per_spoon, max_spoons, tol_pct, note}
 *   {action:'authSetup'}（免密碼）                   權限中心 dim_auth 補「project」列（只補沒有的；密碼空白＝還不能用，經營者在權限管理後台設）
 * 所有儲存格一律存純文字（避免數字、日期被試算表自動轉格式）。
 */
var VER = 'v1.0';
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var AUTH_DASH = 'project', AUTH_SEC = 7200;
var AUTHC_SS_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4', AUTHC_TAB = 'dim_auth';

var T = {
  projects:  { id: 'project_id', cols: ['project_id', 'name', 'goal', 'owner', 'start', 'due', 'status', 'tools', 'sort', 'note', 'created_at', 'updated_at', 'deleted'] },
  phases:    { id: 'phase_id', cols: ['phase_id', 'project_id', 'no', 'title', 'start', 'end', 'owners', 'content', 'checkpoint', 'check_due', 'status', 'done_at', 'note', 'updated_at', 'deleted'] },
  decisions: { id: 'dec_id', cols: ['dec_id', 'project_id', 'no', 'question', 'suggestion', 'answer', 'status', 'decided_at', 'updated_at', 'deleted'] },
  risks:     { id: 'risk_id', cols: ['risk_id', 'project_id', 'no', 'risk', 'action', 'status', 'updated_at', 'deleted'] },
  spoon_lines: { id: '', cols: ['recipe', 'ingredient', 'grams', 'first_spoon', 'snap_date', 'source'] },
  spoon_set: { id: 'ingredient', cols: ['ingredient', 'g_per_spoon', 'max_spoons', 'tol_pct', 'note', 'saved_at'] },
  log:       { id: '', cols: ['ts', 'action', 'table', 'target', 'detail'] }
};
var EDITABLE = { projects: 1, phases: 1, decisions: 1, risks: 1 };

/* ---------------- 進入點 ---------------- */
function doGet(e) {
  var p = (e && e.parameter) || {}, res;
  try {
    var a = p.action || 'ping';
    if (a === 'ping') res = { ok: true, ver: VER };
    else if (a === 'bundle') res = { ok: true, data: bundle_(), at: now_() };
    else if (a === 'spoon') res = { ok: true, data: spoon_(), at: now_() };
    else res = { ok: false, error: '不認得的動作：' + a };
  } catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res, p.callback);
}
function doPost(e) {
  var res;
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}'), a = b.action;
    if (a === 'authSetup') res = authSetup_();
    else {
      auth_(b.pw);
      if (a === 'auth') res = { ok: true };
      else res = withLock_(function () {
        if (a === 'save') return save_(b.table, b.row || {});
        if (a === 'del') return del_(b.table, b.id);
        if (a === 'spoonSave') return spoonSave_(b);
        throw new Error('不認得的動作：' + a);
      });
    }
  } catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res);
}
function out_(res, cb) {
  var s = JSON.stringify(res);
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(String(cb))) return ContentService.createTextOutput(cb + '(' + s + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function withLock_(fn) {
  var lk = LockService.getScriptLock();
  if (!lk.tryLock(20000)) throw new Error('有人正在儲存，請再按一次');
  try { return fn(); } finally { lk.releaseLock(); }
}

/* ---------------- 試算表 ---------------- */
var _ss = null;
function ss_() { return _ss || (_ss = SpreadsheetApp.getActiveSpreadsheet()); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function sheet_(name) {
  var s = ss_(), sh = s.getSheetByName(name), cols = T[name].cols;
  if (!sh) {
    sh = s.insertSheet(name);
    sh.getRange(1, 1, 1000, cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    seedIfEmpty_(name, sh);
  }
  return sh;
}
function rows_(name) {
  var sh = sheet_(name), cols = T[name].cols, n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, cols.length).getDisplayValues().map(function (r, i) {
    var o = { _row: i + 2 }; cols.forEach(function (c, j) { o[c] = r[j]; }); return o;
  });
}
function clean_(o) { var x = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') x[k] = o[k]; }); return x; }
function live_(name) { return rows_(name).filter(function (r) { return r.deleted !== 'Y'; }).map(clean_); }
function append_(name, objs) {
  if (!objs.length) return;
  var sh = sheet_(name), cols = T[name].cols, start = sh.getLastRow() + 1;
  var vals = objs.map(function (o) { return cols.map(function (c) { return o[c] == null ? '' : String(o[c]); }); });
  sh.getRange(start, 1, vals.length, cols.length).setNumberFormat('@').setValues(vals);
}
function log_(action, table, target, detail) {
  try { append_('log', [{ ts: now_(), action: action, table: table || '', target: target || '', detail: String(detail || '').slice(0, 500) }]); } catch (e) {}
}

/* ---------------- 讀 ---------------- */
function bundle_() {
  return { projects: live_('projects'), phases: live_('phases'), decisions: live_('decisions'), risks: live_('risks') };
}
function spoon_() {
  var lines = rows_('spoon_lines').map(clean_);
  return { lines: lines, set: rows_('spoon_set').map(clean_), snap_date: lines.length ? lines[0].snap_date : '', source: lines.length ? lines[0].source : '' };
}

/* ---------------- 寫 ---------------- */
function save_(table, row) {
  if (!EDITABLE[table]) throw new Error('這張表不能改：' + table);
  var def = T[table], idc = def.id, all = rows_(table), id = String(row[idc] || '').trim(), t = now_();
  var o = {};
  def.cols.forEach(function (c) { if (row[c] != null) o[c] = String(row[c]).slice(0, 5000); });
  o.updated_at = t;
  if (table !== 'projects' && !o.project_id && !id) throw new Error('缺專案代號');
  if (!id) {
    id = { projects: 'P', phases: 'PH', decisions: 'D', risks: 'R' }[table] + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    o[idc] = id; o.deleted = '';
    if (table === 'projects') o.created_at = t;
    append_(table, [o]);
    log_('add', table, id, JSON.stringify(o));
    return { ok: true, data: o };
  }
  var hit = all.filter(function (r) { return r[idc] === id; })[0];
  if (!hit) throw new Error('找不到這一列：' + id + '（可能被別人刪了，請重新整理）');
  var cur = clean_(hit), sh = sheet_(table);
  def.cols.forEach(function (c) { if (o[c] != null && c !== idc && c !== 'created_at') cur[c] = o[c]; });
  sh.getRange(hit._row, 1, 1, def.cols.length).setNumberFormat('@').setValues([def.cols.map(function (c) { return cur[c] == null ? '' : String(cur[c]); })]);
  log_('edit', table, id, JSON.stringify(o));
  return { ok: true, data: cur };
}
function del_(table, id) {
  if (!EDITABLE[table]) throw new Error('這張表不能改：' + table);
  var def = T[table], hit = rows_(table).filter(function (r) { return r[def.id] === id; })[0];
  if (!hit) throw new Error('找不到這一列：' + id);
  var c = def.cols.indexOf('deleted') + 1;
  sheet_(table).getRange(hit._row, c).setValue('Y');
  sheet_(table).getRange(hit._row, def.cols.indexOf('updated_at') + 1).setValue(now_());
  log_('del', table, id, '');
  return { ok: true };
}
function spoonSave_(b) {
  var ing = String(b.ingredient || '').trim(), g = Number(b.g_per_spoon);
  if (!ing) throw new Error('缺食材');
  if (!(g > 0 && g < 1000)) throw new Error('一平匙克數要大於 0');
  var o = { ingredient: ing, g_per_spoon: String(g), max_spoons: String(Number(b.max_spoons) || 0), tol_pct: String(Number(b.tol_pct) || 0), note: String(b.note || '').slice(0, 300), saved_at: now_() };
  var hit = rows_('spoon_set').filter(function (r) { return r.ingredient === ing; })[0], cols = T.spoon_set.cols;
  if (hit) sheet_('spoon_set').getRange(hit._row, 1, 1, cols.length).setNumberFormat('@').setValues([cols.map(function (c) { return o[c]; })]);
  else append_('spoon_set', [o]);
  log_('spoonSave', 'spoon_set', ing, JSON.stringify(o));
  return { ok: true, data: o };
}

/* ---------------- 密碼：問權限中心 verify（dim_auth 的 project），成功快取 2 小時（key＝雜湊，不存明文） ---------------- */
function hash_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)).slice(0, 32); }
function auth_(pw) {
  pw = String(pw || '').trim();
  if (!pw) throw new Error('請輸入密碼');
  var cache = CacheService.getScriptCache(), key = 'au:' + hash_(AUTH_DASH + '|' + pw);
  if (cache.get(key)) return true;
  var url = AUTH_API + '?action=verify&dashboard=' + AUTH_DASH + '&pwd=' + encodeURIComponent(pw) + '&callback=cb';
  var t = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }).getContentText();
  var m = t.match(/^\s*[A-Za-z_$][\w$]*\(([\s\S]*)\);?\s*$/), j = {};
  try { j = JSON.parse(m ? m[1] : t); } catch (e) { throw new Error('權限中心回應看不懂，請稍後再試'); }
  if (!j || !j.ok) throw new Error('密碼不對（' + (j && j.error || '') + '）');
  cache.put(key, '1', AUTH_SEC);
  return true;
}

/* ---------------- 一次性：權限中心 dim_auth 補「project」列（只補沒有的；先備份） ---------------- */
function authSetup_() {
  var ss = SpreadsheetApp.openById(AUTHC_SS_ID), sh = ss.getSheetByName(AUTHC_TAB);
  if (!sh) throw new Error('找不到權限中心 dim_auth');
  var last = sh.getLastRow(), have = false;
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { if (String(r[0]).trim() === AUTH_DASH) have = true; });
  if (have) return { ok: true, msg: '已經有 project 列，不用補' };
  var bk = 'dim_auth_備份_' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
  sh.getRange(last + 1, 1, 1, 7).setNumberFormat('@');
  sh.getRange(last + 1, 1, 1, 7).setValues([[AUTH_DASH, '專案推動', '', '12', 'Y', now_(), '專案推動儀表板：看頁面＋編輯（2026-10-17 補；密碼空白＝還不能用，請在權限管理後台設定）']]);
  return { ok: true, msg: '已補 project 列；備份 ' + bk, before: last - 1 };
}

/* ---------------- 授權與自我檢查（經營者第一次在編輯器按「執行」用；只讀） ---------------- */
function authorize() {
  var n = ss_().getSheets().length;
  UrlFetchApp.fetch(AUTH_API + '?action=ping', { muteHttpExceptions: true });
  SpreadsheetApp.openById(AUTHC_SS_ID).getName();
  Logger.log('授權完成：試算表分頁 ' + n + ' 張、權限中心連得到');
}
/* 建好全部分頁並放入第一批資料（只在分頁不存在時建；已存在就不動） */
function setupTables() {
  Object.keys(T).forEach(function (k) { sheet_(k); });
  var d = ss_().getSheetByName('工作表1'); if (d && d.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(d);
  Logger.log(JSON.stringify({ projects: live_('projects').length, phases: live_('phases').length, decisions: live_('decisions').length, risks: live_('risks').length, spoon_lines: rows_('spoon_lines').length, spoon_set: rows_('spoon_set').length }));
}
/* 寫入流程自我檢查：新增測試專案 → 改 → 刪；跑完只留 deleted=Y 的一列（以及 log） */
function selfTest() {
  var a = save_('projects', { name: 'zz 測試專案（自我檢查，會自動刪）', status: '未開始' }).data;
  var b = save_('projects', { project_id: a.project_id, status: '進行中' }).data;
  del_('projects', a.project_id);
  var gone = live_('projects').every(function (r) { return r.project_id !== a.project_id; });
  Logger.log(JSON.stringify({ added: a.project_id, edited: b.status === '進行中', deleted: gone }));
}
