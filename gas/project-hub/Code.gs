/**
 * 專案推動 API（2026-10-17 新建，經營者交辦：「做個專案推動儀表板，未來所有新專案都在這裡管理，從『食材改平匙』開始」）
 * 綁在試算表「專案推動_資料庫」（1VE1dsSPk0Bl2dyTWrNo3G_oasD4pCR3ihllNedcOOT4）；前端＝static/project-hub.html。
 *
 * 讀（GET，JSONP，免密碼，跟其他儀表板一樣靠頁面閘門擋路人）：
 *   ?action=ping          版號
 *   ?action=bundle        專案、階段、待決定事項、風險、人員
 *   ?action=spoon         平匙換算：食譜用量快照＋各食材的匙設定
 *   ?action=remind_preview 今天提醒信會寄什麼（只回項目與收件人姓名，不寄、不回 email）
 * 寫（POST，body＝JSON 字串；都要帶 pw＝權限中心 dim_auth「project」那列的密碼）：
 *   {action:'auth'}                                   只驗密碼
 *   {action:'save', table, row}                      新增／修改一列（沒有 id＝新增）
 *   {action:'del', table, id}                        刪除（標記 deleted=Y，不真的刪）
 *   {action:'spoonSave', ingredient, g_per_spoon, max_spoons, tol_pct, note}
 *   {action:'authSetup'}（免密碼）                   權限中心 dim_auth 補「project」列（只補沒有的；密碼空白＝還不能用，經營者在權限管理後台設）
 * 所有儲存格一律存純文字（避免數字、日期被試算表自動轉格式）。
 *
 * v1.1（2026-10-18）：
 *   - 欄位改「照標題列對位」：T 裡新增的欄位會自動補在試算表最後一欄，舊資料不用搬。
 *   - phases 加 lead（負責人，一個人）、done_items（執行內容已勾完成的行，用換行分開存原文）；
 *     decisions 加 due（決定期限）、blocks（不決定會卡哪個階段）；新表 people（姓名、角色、Email，提醒信用）。
 *   - 每日提醒信 dailyReminder（排程每天 08:00）：過期檢查點、該開始沒開始、卡關、7 天內到的檢查點與階段、待決定期限 → 寄給負責人／窗口／經營者。
 */
var VER = 'v1.1';
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var AUTH_DASH = 'project', AUTH_SEC = 7200;
var AUTHC_SS_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4', AUTHC_TAB = 'dim_auth';
var DASH_URL = 'https://diybc-training.onrender.com/static/project-hub.html';
var REMIND_DAYS = 7;   /* 提醒信與「本週要做」看未來幾天 */

var T = {
  projects:  { id: 'project_id', cols: ['project_id', 'name', 'goal', 'owner', 'start', 'due', 'status', 'tools', 'sort', 'note', 'created_at', 'updated_at', 'deleted'] },
  phases:    { id: 'phase_id', cols: ['phase_id', 'project_id', 'no', 'title', 'start', 'end', 'owners', 'content', 'checkpoint', 'check_due', 'status', 'done_at', 'note', 'updated_at', 'deleted', 'lead', 'done_items'] },
  decisions: { id: 'dec_id', cols: ['dec_id', 'project_id', 'no', 'question', 'suggestion', 'answer', 'status', 'decided_at', 'updated_at', 'deleted', 'due', 'blocks'] },
  risks:     { id: 'risk_id', cols: ['risk_id', 'project_id', 'no', 'risk', 'action', 'status', 'updated_at', 'deleted'] },
  people:    { id: 'person_id', cols: ['person_id', 'name', 'role', 'email', 'note', 'updated_at', 'deleted'] },
  spoon_lines: { id: '', cols: ['recipe', 'ingredient', 'grams', 'first_spoon', 'snap_date', 'source'] },
  spoon_set: { id: 'ingredient', cols: ['ingredient', 'g_per_spoon', 'max_spoons', 'tol_pct', 'note', 'saved_at'] },
  log:       { id: '', cols: ['ts', 'action', 'table', 'target', 'detail'] }
};
var EDITABLE = { projects: 1, phases: 1, decisions: 1, risks: 1, people: 1 };
var ID_PREFIX = { projects: 'P', phases: 'PH', decisions: 'D', risks: 'R', people: 'U' };

/* ---------------- 進入點 ---------------- */
function doGet(e) {
  var p = (e && e.parameter) || {}, res;
  try {
    var a = p.action || 'ping';
    if (a === 'ping') res = { ok: true, ver: VER };
    else if (a === 'bundle') res = { ok: true, data: bundle_(), at: now_() };
    else if (a === 'spoon') res = { ok: true, data: spoon_(), at: now_() };
    else if (a === 'remind_preview') res = { ok: true, data: remindPublic_(remindBuild_()), at: now_() };
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

/* ---------------- 試算表（欄位照標題列對位；T 裡新增的欄位自動補在最後面） ---------------- */
var _ss = null, _hdr = {};
function ss_() { return _ss || (_ss = SpreadsheetApp.getActiveSpreadsheet()); }
function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); }
function sheet_(name) {
  var s = ss_(), sh = s.getSheetByName(name), cols = T[name].cols;
  if (!sh) {
    sh = s.insertSheet(name);
    sh.getRange(1, 1, 1000, cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    _hdr[name] = cols.slice();
    seedIfEmpty_(name, sh);
  } else if (!_hdr[name]) ensureCols_(name, sh);
  return sh;
}
/* 讀標題列；T 有、試算表沒有的欄位補在最後（整欄設純文字） */
function ensureCols_(name, sh) {
  var cols = T[name].cols, lastC = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastC).getDisplayValues()[0].map(function (h) { return String(h).trim(); });
  while (hdr.length && !hdr[hdr.length - 1]) hdr.pop();
  var missing = cols.filter(function (c) { return hdr.indexOf(c) < 0; });
  if (missing.length) {
    sh.getRange(1, hdr.length + 1, 1000, missing.length).setNumberFormat('@');
    sh.getRange(1, hdr.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    hdr = hdr.concat(missing);
    log_('addcol', name, missing.join(','), '');
  }
  _hdr[name] = hdr;
  return hdr;
}
function hdr_(name) { sheet_(name); return _hdr[name]; }
function rows_(name) {
  var sh = sheet_(name), hdr = hdr_(name), cols = T[name].cols, n = sh.getLastRow();
  if (n < 2) return [];
  var idx = cols.map(function (c) { return hdr.indexOf(c); });
  return sh.getRange(2, 1, n - 1, hdr.length).getDisplayValues().map(function (r, i) {
    var o = { _row: i + 2 }; cols.forEach(function (c, j) { o[c] = idx[j] < 0 ? '' : r[idx[j]]; }); return o;
  });
}
function clean_(o) { var x = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') x[k] = o[k]; }); return x; }
function live_(name) { return rows_(name).filter(function (r) { return r.deleted !== 'Y'; }).map(clean_); }
function rowVals_(name, o, base) {   /* base＝原本整列（保留 T 不認得的欄位） */
  var hdr = hdr_(name), cols = T[name].cols;
  return hdr.map(function (h, j) { return cols.indexOf(h) >= 0 ? (o[h] == null ? '' : String(o[h])) : (base ? base[j] : ''); });
}
function append_(name, objs) {
  if (!objs.length) return;
  var sh = sheet_(name), hdr = hdr_(name), start = sh.getLastRow() + 1;
  var vals = objs.map(function (o) { return rowVals_(name, o); });
  sh.getRange(start, 1, vals.length, hdr.length).setNumberFormat('@').setValues(vals);
}
function writeRow_(name, rowNo, o) {
  var sh = sheet_(name), hdr = hdr_(name), rg = sh.getRange(rowNo, 1, 1, hdr.length);
  rg.setNumberFormat('@').setValues([rowVals_(name, o, rg.getDisplayValues()[0])]);
}
function log_(action, table, target, detail) {
  try { append_('log', [{ ts: now_(), action: action, table: table || '', target: target || '', detail: String(detail || '').slice(0, 500) }]); } catch (e) {}
}

/* ---------------- 讀 ---------------- */
function bundle_() {
  return { projects: live_('projects'), phases: live_('phases'), decisions: live_('decisions'), risks: live_('risks'), people: live_('people') };
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
  if (table !== 'projects' && table !== 'people' && !o.project_id && !id) throw new Error('缺專案代號');
  if (table === 'people' && !id && !String(o.name || '').trim()) throw new Error('缺姓名');
  if (!id) {
    id = ID_PREFIX[table] + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    o[idc] = id; o.deleted = '';
    if (table === 'projects') o.created_at = t;
    append_(table, [o]);
    log_('add', table, id, JSON.stringify(o));
    return { ok: true, data: o };
  }
  var hit = all.filter(function (r) { return r[idc] === id; })[0];
  if (!hit) throw new Error('找不到這一列：' + id + '（可能被別人刪了，請重新整理）');
  var cur = clean_(hit);
  def.cols.forEach(function (c) { if (o[c] != null && c !== idc && c !== 'created_at') cur[c] = o[c]; });
  writeRow_(table, hit._row, cur);
  log_('edit', table, id, JSON.stringify(o));
  return { ok: true, data: cur };
}
function del_(table, id) {
  if (!EDITABLE[table]) throw new Error('這張表不能改：' + table);
  var def = T[table], hit = rows_(table).filter(function (r) { return r[def.id] === id; })[0];
  if (!hit) throw new Error('找不到這一列：' + id);
  var cur = clean_(hit); cur.deleted = 'Y'; cur.updated_at = now_();
  writeRow_(table, hit._row, cur);
  log_('del', table, id, '');
  return { ok: true };
}
function spoonSave_(b) {
  var ing = String(b.ingredient || '').trim(), g = Number(b.g_per_spoon);
  if (!ing) throw new Error('缺食材');
  if (!(g > 0 && g < 1000)) throw new Error('一平匙克數要大於 0');
  var o = { ingredient: ing, g_per_spoon: String(g), max_spoons: String(Number(b.max_spoons) || 0), tol_pct: String(Number(b.tol_pct) || 0), note: String(b.note || '').slice(0, 300), saved_at: now_() };
  var hit = rows_('spoon_set').filter(function (r) { return r.ingredient === ing; })[0];
  if (hit) writeRow_('spoon_set', hit._row, o);
  else append_('spoon_set', [o]);
  log_('spoonSave', 'spoon_set', ing, JSON.stringify(o));
  return { ok: true, data: o };
}

/* ---------------- 每日提醒（跟前端「本週要做」同一套規則） ---------------- */
function addDays_(ymd, n) {
  var d = new Date(Number(ymd.slice(0, 4)), Number(ymd.slice(5, 7)) - 1, Number(ymd.slice(8, 10)) + n);
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function md_(s) { return s ? Number(s.slice(5, 7)) + '/' + Number(s.slice(8, 10)) + '（' + '日一二三四五六'.charAt(new Date(Number(s.slice(0, 4)), Number(s.slice(5, 7)) - 1, Number(s.slice(8, 10))).getDay()) + '）' : ''; }
/* 回 { today, until, items:[{sev,date,pid,pname,tag,text,who}], noDue:[{pid,pname,n}] }；sev 0＝要處理、1＝檢查點／決定 7 天內到、2＝階段 7 天內開始 */
function remindBuild_() {
  var t = today_(), until = addDays_(t, REMIND_DAYS), b = bundle_(), items = [], noDue = [];
  b.projects.filter(function (p) { return p.status === '進行中' || p.status === '未開始'; }).forEach(function (p) {
    var ph = b.phases.filter(function (x) { return x.project_id === p.project_id; }).sort(function (a, c) { return (Number(a.no) || 0) - (Number(c.no) || 0); });
    ph.forEach(function (x) {
      if (x.status === '完成') return;
      var who = x.lead || x.owners || p.owner || '', nm = '第 ' + x.no + ' 階段「' + x.title + '」';
      if (x.status === '卡關') items.push({ sev: 0, date: x.end || x.start || t, pid: p.project_id, pname: p.name, tag: '卡關', text: nm + (x.note ? '：' + x.note : ''), who: who });
      if (x.status === '未開始' && x.start && x.start < t) items.push({ sev: 0, date: x.start, pid: p.project_id, pname: p.name, tag: '該開始了還沒開始', text: nm + '（原訂 ' + md_(x.start) + ' 開始）', who: who });
      if (x.check_due && x.check_due < t) items.push({ sev: 0, date: x.check_due, pid: p.project_id, pname: p.name, tag: '檢查點過期', text: nm + '：' + x.checkpoint, who: who });
      else if (x.check_due && x.check_due <= until) items.push({ sev: 1, date: x.check_due, pid: p.project_id, pname: p.name, tag: x.check_due === t ? '檢查點今天到' : '檢查點到期', text: nm + '：' + x.checkpoint, who: who });
      if (x.start && x.start >= t && x.start <= until) items.push({ sev: 2, date: x.start, pid: p.project_id, pname: p.name, tag: x.start === t ? '今天開始' : '階段開始', text: nm + (x.content ? '：' + String(x.content).split(/\n+/)[0] : ''), who: who });
    });
    var open = b.decisions.filter(function (d) { return d.project_id === p.project_id && d.status !== '已決定'; }), nd = 0;
    open.forEach(function (d) {
      var who = p.owner || '', txt = '待決定 #' + d.no + '：' + d.question + (d.blocks ? '（不決定會卡 ' + d.blocks + '）' : '');
      if (d.due && d.due < t) items.push({ sev: 0, date: d.due, pid: p.project_id, pname: p.name, tag: '待決定過期', text: txt, who: who });
      else if (d.due && d.due <= until) items.push({ sev: 1, date: d.due, pid: p.project_id, pname: p.name, tag: d.due === t ? '今天要決定' : '要決定了', text: txt, who: who });
      else if (!d.due) nd++;
    });
    if (nd) noDue.push({ pid: p.project_id, pname: p.name, n: nd, who: p.owner || '' });
  });
  items.sort(function (a, c) { return (a.sev - c.sev) || (a.date < c.date ? -1 : a.date > c.date ? 1 : 0); });
  return { today: t, until: until, items: items, noDue: noDue, people: b.people };
}
function remindPublic_(r) {   /* 給前端／curl 看的版本：不回 email */
  return { today: r.today, until: r.until, items: r.items, noDue: r.noDue, recipients: remindRecipients_(r).map(function (x) { return { name: x.name, role: x.role, n: x.items.length + x.noDue.length, hasEmail: !!x.email }; }) };
}
/* 誰收什麼：經營者（角色含「經營者」）收全部；其他人收「負責人／窗口」字串裡包含他姓名的項目；同一個 email 只寄一封 */
function remindRecipients_(r) {
  var out = [];
  (r.people || []).forEach(function (p) {
    var name = String(p.name || '').trim(); if (!name) return;
    var boss = /經營者/.test(String(p.role || ''));
    var mine = r.items.filter(function (it) { return boss || String(it.who).indexOf(name) >= 0; });
    var nd = r.noDue.filter(function (x) { return boss || String(x.who).indexOf(name) >= 0; });
    if (!mine.length && !nd.length) return;
    out.push({ name: name, role: p.role || '', email: String(p.email || '').trim(), boss: boss, items: mine, noDue: nd });
  });
  return out;
}
function remindHtml_(r, rc) {
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  var sec = function (title, color, list) {
    if (!list.length) return '';
    return '<h3 style="margin:16px 0 6px;color:' + color + ';font-size:15px">' + title + '（' + list.length + '）</h3><table style="border-collapse:collapse;font-size:14px">'
      + list.map(function (it) { return '<tr><td style="padding:4px 8px 4px 0;white-space:nowrap;color:#555">' + esc(md_(it.date)) + '</td><td style="padding:4px 8px;white-space:nowrap"><b>' + esc(it.tag) + '</b></td><td style="padding:4px 0">' + esc(it.text) + '<div style="color:#777;font-size:12.5px">' + esc(it.pname) + '　負責：' + esc(it.who || '未定') + '　<a href="' + DASH_URL + '#p=' + esc(it.pid) + '">開專案</a></div></td></tr>'; }).join('')
      + '</table>';
  };
  var s0 = rc.items.filter(function (i) { return i.sev === 0; }), s1 = rc.items.filter(function (i) { return i.sev === 1; }), s2 = rc.items.filter(function (i) { return i.sev === 2; });
  var h = '<div style="font-family:-apple-system,\'Noto Sans TC\',sans-serif;color:#2A2320;max-width:720px">'
    + '<p style="margin:0 0 4px">' + esc(rc.name) + ' 您好，這是 ' + esc(md_(r.today)) + ' 的專案進度提醒（看到 ' + esc(md_(r.until)) + '）。</p>'
    + sec('🔴 要處理', '#A3321F', s0) + sec('🟡 這 ' + REMIND_DAYS + ' 天內要到', '#8A5A00', s1) + sec('📅 這 ' + REMIND_DAYS + ' 天內開始', '#23436E', s2)
    + (rc.noDue.length ? '<p style="margin:14px 0 0;color:#555;font-size:13.5px">另外有還沒定期限的待決定事項：' + rc.noDue.map(function (x) { return esc(x.pname) + ' ' + x.n + ' 件'; }).join('、') + '，請在儀表板補上「決定期限」。</p>' : '')
    + '<p style="margin:18px 0 0;font-size:13px;color:#777">進度改了請到 <a href="' + DASH_URL + '">專案推動儀表板</a> 更新（階段狀態、執行內容打勾、待決定填答案），大家就都看得到。這封信每天 08:00 自動寄；不想收請跟經營者說。</p></div>';
  return h;
}
function remindText_(r, rc) {
  return rc.items.map(function (it) { return md_(it.date) + ' ' + it.tag + '：' + it.text + '（' + it.pname + '，負責 ' + (it.who || '未定') + '）'; }).join('\n') + '\n\n' + DASH_URL;
}
/* 排程每天 08:00 跑（setupTrigger 建立）。沒有東西要提醒的人不寄。 */
function dailyReminder() {
  var r = remindBuild_(), rcs = remindRecipients_(r), sent = [], skip = [], seen = {};
  rcs.forEach(function (rc) {
    if (!rc.email) { skip.push(rc.name + '（沒 email）'); return; }
    if (seen[rc.email]) { skip.push(rc.name + '（同 email 已寄）'); return; }
    seen[rc.email] = 1;
    var n0 = rc.items.filter(function (i) { return i.sev === 0; }).length, n1 = rc.items.length - n0;
    var subject = '🚀 專案推動 ' + md_(r.today) + '：要處理 ' + n0 + '、' + REMIND_DAYS + ' 天內 ' + n1 + (rc.boss ? '（全部專案）' : '');
    MailApp.sendEmail({ to: rc.email, subject: subject, htmlBody: remindHtml_(r, rc), body: remindText_(r, rc), name: '專案推動儀表板' });
    sent.push(rc.name);
  });
  log_('remind', '', r.today, '寄出 ' + sent.join('、') + (skip.length ? '；略過 ' + skip.join('、') : '') + '；項目 ' + r.items.length);
  Logger.log(JSON.stringify({ today: r.today, items: r.items.length, sent: sent, skip: skip }));
  return { sent: sent, skip: skip, items: r.items.length };
}
/* 編輯器用：看今天會寄什麼（不寄） */
function reminderPreview() {
  var r = remindBuild_();
  Logger.log(JSON.stringify({ today: r.today, until: r.until, items: r.items.map(function (i) { return [i.sev, i.date, i.tag, i.text, i.who]; }), noDue: r.noDue, recipients: remindRecipients_(r).map(function (x) { return x.name + (x.email ? '' : '（沒 email）') + ' ' + x.items.length + ' 項'; }) }));
}
/* 編輯器用：建立／重建每天 08:00 的排程（先刪舊的同名排程） */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(8).nearMinute(0).inTimezone('Asia/Taipei').create();
  Logger.log('已建立排程：dailyReminder 每天 08:00（Asia/Taipei）');
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
  var q = MailApp.getRemainingDailyQuota(), tr = ScriptApp.getProjectTriggers().length;
  Logger.log('授權完成：試算表分頁 ' + n + ' 張、權限中心連得到、今日還可寄 ' + q + ' 封信、排程 ' + tr + ' 個');
}
/* 建好全部分頁並放入第一批資料（只在分頁不存在時建；已存在就補缺的欄位） */
function setupTables() {
  Object.keys(T).forEach(function (k) { sheet_(k); });
  var d = ss_().getSheetByName('工作表1'); if (d && d.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(d);
  Logger.log(JSON.stringify({ projects: live_('projects').length, phases: live_('phases').length, decisions: live_('decisions').length, risks: live_('risks').length, people: live_('people').length, spoon_lines: rows_('spoon_lines').length, spoon_set: rows_('spoon_set').length, hdr: _hdr }));
}
/* 寫入流程自我檢查：新增測試專案 → 改 → 刪；跑完只留 deleted=Y 的一列（以及 log） */
function selfTest() {
  var a = save_('projects', { name: 'zz 測試專案（自我檢查，會自動刪）', status: '未開始' }).data;
  var b = save_('projects', { project_id: a.project_id, status: '進行中' }).data;
  del_('projects', a.project_id);
  var gone = live_('projects').every(function (r) { return r.project_id !== a.project_id; });
  Logger.log(JSON.stringify({ added: a.project_id, edited: b.status === '進行中', deleted: gone }));
}
