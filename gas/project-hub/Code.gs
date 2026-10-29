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
 *   {action:'specSave', ingredient, spoon_ml, g_per_spoon, real_g, note}      量匙規格（ml）＋食譜用克數＋真實克數；同時記一列 spoon_spec_log
 *                                                                              （2026-10-21 經營者：「儲存變更改成不需要密碼」→ 這個動作免 pw，其他寫入照舊要密碼）
 *   {action:'evoSnap', version, ver_date, rows:[…]}                           把頁面算好的換算記成一個版本（同一行同版本名稱會先標刪除再寫）
 *   {action:'finalConfirm', ingredient, g_per_spoon, spoon_ml, max_spoons, tol_pct, snap_date, uses:{line_key:數}}   v2.0：平匙換算頁「✅ 確認轉換到執行進度」
 *   {action:'finalEdit', fin_id, value('' 恢復系統算的｜數字｜'秤重'), reason}                               v2.0：已定案明細手動調整平匙數
 *   {action:'finalApplied', fin_ids:[…], batch, date, undo}                                                 v2.0：記錄已套用到門市食譜後台（Claude 改完後台用）
 *   {action:'authSetup'}（免密碼）                   權限中心 dim_auth 補「project」列（只補沒有的；密碼空白＝還不能用，經營者在權限管理後台設）
 * 所有儲存格一律存純文字（避免數字、日期被試算表自動轉格式）。
 *
 * v1.1（2026-10-18）：
 *   - 欄位改「照標題列對位」：T 裡新增的欄位會自動補在試算表最後一欄，舊資料不用搬。
 *   - phases 加 lead（負責人，一個人）、done_items（執行內容已勾完成的行，用換行分開存原文）；
 *     decisions 加 due（決定期限）、blocks（不決定會卡哪個階段）；新表 people（姓名、角色、Email，提醒信用）。
 *   - 每日提醒信 dailyReminder（排程每天 08:00）：過期檢查點、該開始沒開始、卡關、7 天內到的檢查點與階段、待決定期限 → 寄給負責人／窗口／經營者。
 */
/* v1.2（2026-10-21，經營者交辦「平匙專案加一頁：每個食材最後用的 1 平匙＝幾 g 與量匙規格，分食譜用克數／真實克數；
 *   換成平匙的食譜演進，可用食材或甜點篩選，並放入食譜後台『台灣匙版備份』第一版的舊平匙數做比較」）：
 *   - spoon_set 加 spoon_ml（量匙規格 ml）、real_g（1 平匙真實克數）；g_per_spoon 就是「食譜用克數」。
 *   - 新表 spoon_spec_log（規格每改一次記一列）、spoon_evo（每支甜點每一處乾料、每個版本的平匙數）。
 *   - 第一批資料在 seed.gs 的 EVO_SEED／SPEC_SEED（分頁第一次建立時放入）。 */
var VER = 'v2.0';   /* v2.0（2026-10-29）：spoon_final 定案明細（平匙換算頁按「✅ 確認轉換到執行進度」寫入、已定案明細可手動調平匙數、記錄套用到後台）；spoon_lines 多 recipe_id／step_no／step_title／seq（spoon_feed.gs 讀後台時一起記）；v1.9（2026-10-28）：item_usage／recipes_tw（🔎 品項使用查詢：後台每個品項被哪些食譜用，瀏覽器掃描後 itemScanSave 上傳，GET items 讀）；v1.8（2026-10-27）：spoonSave 可帶 spoon_ml；v1.7（2026-10-22）：spoon_lines 多 use_step／purpose（食譜系統 spoon_feed.gs 每天自動更新快照）；v1.6（2026-10-22）：spoon_set 多 kind，deco＝裝飾乾料（specSave 不要求食譜用克數）；v1.5（2026-10-22）：spoon 每項多 purpose／use_step（使用目的，資料在 purpose.gs）；v1.4（2026-10-22）：projects 多 files（📎 專案檔案）；v1.3：specSave 免密碼 */
var AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var AUTH_DASH = 'project', AUTH_SEC = 7200;
var AUTHC_SS_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4', AUTHC_TAB = 'dim_auth';
var DASH_URL = 'https://diybc-training.onrender.com/static/project-hub.html';
var REMIND_DAYS = 7;   /* 提醒信與「本週要做」看未來幾天 */

var T = {
  projects:  { id: 'project_id', cols: ['project_id', 'name', 'goal', 'owner', 'start', 'due', 'status', 'tools', 'sort', 'note', 'created_at', 'updated_at', 'deleted', 'files'] },   /* files（v1.4，2026-10-22）：📎 專案檔案，一行一個「名稱｜網址」 */
  phases:    { id: 'phase_id', cols: ['phase_id', 'project_id', 'no', 'title', 'start', 'end', 'owners', 'content', 'checkpoint', 'check_due', 'status', 'done_at', 'note', 'updated_at', 'deleted', 'lead', 'done_items'] },
  decisions: { id: 'dec_id', cols: ['dec_id', 'project_id', 'no', 'question', 'suggestion', 'answer', 'status', 'decided_at', 'updated_at', 'deleted', 'due', 'blocks'] },
  risks:     { id: 'risk_id', cols: ['risk_id', 'project_id', 'no', 'risk', 'action', 'status', 'updated_at', 'deleted'] },
  people:    { id: 'person_id', cols: ['person_id', 'name', 'role', 'email', 'note', 'updated_at', 'deleted'] },
  spoon_lines: { id: '', cols: ['recipe', 'ingredient', 'grams', 'first_spoon', 'snap_date', 'source', 'use_step', 'purpose', 'recipe_id', 'step_no', 'step_title', 'seq'] },   /* v2.0：recipe_id＝後台食譜編號、step_no＝第幾步、seq＝同一步同食材同克數的第幾個 */
  /* v2.0 定案明細（2026-10-29 經營者：「平匙換算頁按確認就把當下資料轉到執行進度；已定案明細的平匙數可手動調；正式食譜依這份調整」）
     一處一列；line_key＝recipe_id|step_no|食材|克數|seq。calc_*＝系統依當下規格算的；manual_spoons＝手動（數字或「秤重」）；final_spoons＝最後要用的（空白＝維持秤重）。
     重新確認同一食材：快照裡有的那幾處重算（舊列標 deleted=Y、replaced_by＝新批號），手動改過的照搬；快照裡已經沒有的（改完後台或停用）原樣保留。 */
  spoon_final: { id: 'fin_id', cols: ['fin_id', 'batch', 'ingredient', 'recipe', 'recipe_id', 'step_no', 'step_title', 'grams', 'seq', 'line_key', 'g_per_spoon', 'spoon_ml', 'max_spoons', 'tol_pct',
    'calc_spoons', 'calc_ok', 'calc_why', 'manual_spoons', 'manual_reason', 'manual_at', 'final_spoons', 'uses', 'purpose', 'use_step', 'source', 'confirmed_at', 'applied_at', 'applied_batch', 'updated_at', 'deleted', 'replaced_by'] },
  spoon_set: { id: 'ingredient', cols: ['ingredient', 'g_per_spoon', 'max_spoons', 'tol_pct', 'note', 'saved_at', 'spoon_ml', 'real_g', 'kind'] },   /* kind（v1.6）：deco＝裝飾乾料，只登記量匙與 1 匙實測克數，不做換算 */
  spoon_spec_log: { id: 'log_id', cols: ['log_id', 'ingredient', 'version', 'eff_date', 'spoon_ml', 'recipe_g', 'real_g', 'note', 'saved_at', 'deleted'] },
  spoon_evo: { id: 'evo_id', cols: ['evo_id', 'line_key', 'recipe', 'recipe_id', 'step', 'step_title', 'ingredient', 'grams', 'version', 'ver_date', 'spoons', 'spoon_ml', 'g_per_spoon', 'real_g', 'note', 'source', 'updated_at', 'deleted'] },
  item_usage: { id: '', cols: ['item_id', 'code', 'zone', 'name', 'unit', 'n_all', 'n_tw', 'n_zh', 'n_act', 'last_used', 'rids', 'scanned_at'] },   /* v1.9：每個後台品項被幾支食譜用；rids＝台灣區食譜 id 前 8 碼逗號分隔 */
  recipes_tw: { id: '', cols: ['rid', 'name', 'lang', 'status', 'updated', 'scanned_at'] },   /* v1.9：台灣區食譜清單（繁中＋英文），rid＝後台 id 前 8 碼 */
  log:       { id: '', cols: ['ts', 'action', 'table', 'target', 'detail'] }
};
var SCAN_TABLES = { item_usage: 1, recipes_tw: 1 };
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
    else if (a === 'items') res = { ok: true, data: items_(), at: now_() };
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
    else if (a === 'specSave') res = withLock_(function () { return specSave_(b); });   /* 平匙規格表「💾 儲存變更」免密碼 */
    else if (a === 'itemScanSave') res = withLock_(function () { return itemScanSave_(b); });   /* v1.9：瀏覽器掃描結果上傳，驗指令碼屬性 SCAN_TOKEN，不用儀表板密碼 */
    else {
      auth_(b.pw);
      if (a === 'auth') res = { ok: true };
      else res = withLock_(function () {
        if (a === 'save') return save_(b.table, b.row || {});
        if (a === 'del') return del_(b.table, b.id);
        if (a === 'spoonSave') return spoonSave_(b);
        if (a === 'evoSnap') return evoSnap_(b);
        if (a === 'finalConfirm') return finalConfirm_(b);
        if (a === 'finalEdit') return finalEdit_(b);
        if (a === 'finalApplied') return finalApplied_(b);
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
  if (sh.getMaxRows && start + vals.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + vals.length - 1 - sh.getMaxRows() + 200);   /* 新分頁只有 1,000 列，超過要先加列 */
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
/* ---------- v1.9 🔎 品項使用查詢 ---------- */
function items_() {
  var items = rows_('item_usage').map(clean_), recipes = rows_('recipes_tw').map(clean_);
  var at = items.length ? items[0].scanned_at : '';
  return { items: items, recipes: recipes, scanned_at: at };
}
function scanToken_() { try { return String(PropertiesService.getScriptProperties().getProperty('SCAN_TOKEN') || ''); } catch (e) { return ''; } }
/* {token, table:'item_usage'|'recipes_tw', reset:1(第一包先清掉舊資料), rows:[{…}]}；一包最多 600 列 */
function itemScanSave_(b) {
  var tk = scanToken_();
  if (!tk || String(b.token || '') !== tk) throw new Error('掃描上傳的通行碼不對');
  var table = String(b.table || ''), rows = b.rows || [];
  if (!SCAN_TABLES[table]) throw new Error('不認得的表：' + table);
  if (rows.length > 600) throw new Error('一次最多 600 列');
  var sh = sheet_(table), cols = hdr_(table);
  if (b.reset) { var n = sh.getLastRow(); if (n > 1) sh.getRange(2, 1, n - 1, Math.max(cols.length, sh.getLastColumn())).clearContent(); }
  var t = now_();
  append_(table, rows.map(function (r) { var o = {}; cols.forEach(function (c) { o[c] = r[c] == null ? '' : String(r[c]).slice(0, 40000); }); o.scanned_at = t; return o; }));
  log_('scan', table, b.reset ? 'reset' : 'append', rows.length + ' 列');
  return { ok: true, data: { table: table, added: rows.length, total: Math.max(sheet_(table).getLastRow() - 1, 0) } };
}
function spoon_() {
  var lines = rows_('spoon_lines').map(clean_);
  /* 使用目的（2026-10-22，資料在 purpose.gs）：同一支食譜同食材同克數出現多次，依快照順序對應 */
  var seen = {}, PU = typeof PURPOSE === 'undefined' ? {} : PURPOSE;
  lines.forEach(function (l) {
    var k = l.recipe + '|' + l.ingredient + '|' + l.grams, n = seen[k] = (seen[k] || 0), p = PU[k] || [];
    seen[k]++;
    var x = p.length ? p[Math.min(n, p.length - 1)] : null;
    /* v1.7：快照由食譜系統每天自動讀（spoon_feed.gs），列上自帶 use_step／purpose；purpose.gs 有人工修正的仍優先 */
    l.use_step = x ? x[0] : (l.use_step || ''); l.purpose = x ? x[1] : (l.purpose || '');
  });
  return { lines: lines, set: rows_('spoon_set').map(clean_), snap_date: lines.length ? lines[0].snap_date : '', source: lines.length ? lines[0].source : '',
    spec_log: live_('spoon_spec_log'), evo: live_('spoon_evo'), final: live_('spoon_final') };
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
  var hit = rows_('spoon_set').filter(function (r) { return r.ingredient === ing; })[0];
  var o = hit ? clean_(hit) : { ingredient: ing, spoon_ml: '', real_g: '' };   /* v1.2：保留量匙規格、真實克數 */
  if (b.spoon_ml !== undefined) { var ml = String(b.spoon_ml == null ? '' : b.spoon_ml).trim(); if (ml !== '' && !(Number(ml) >= 0)) throw new Error('量匙規格 ml 要是數字'); o.spoon_ml = ml === '' ? '' : String(Number(ml)); }   /* v1.8：換算頁也能填量匙 ml */
  o.g_per_spoon = String(g); o.max_spoons = String(Number(b.max_spoons) || 0); o.tol_pct = String(Number(b.tol_pct) || 0); o.note = String(b.note || '').slice(0, 300); o.saved_at = now_();
  if (hit) writeRow_('spoon_set', hit._row, o);
  else append_('spoon_set', [o]);
  specLog_(o, '換算頁存檔');
  log_('spoonSave', 'spoon_set', ing, JSON.stringify(o));
  return { ok: true, data: o };
}
function specLog_(o, version) {
  append_('spoon_spec_log', [{ log_id: 'SL' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10), ingredient: o.ingredient, version: version, eff_date: today_(),
    spoon_ml: o.spoon_ml || '', recipe_g: o.g_per_spoon || '', real_g: o.real_g || '', note: o.note || '', saved_at: o.saved_at || now_(), deleted: '' }]);
}
/* 量匙規格：spoon_ml＝量匙容量（ml），g_per_spoon＝食譜用克數（換算平匙數用），real_g＝真實克數（實測；平匙數 × 真實克數＝真正用了幾 g） */
function specSave_(b) {
  var ing = String(b.ingredient || '').trim(), ml = String(b.spoon_ml == null ? '' : b.spoon_ml).trim(), g = Number(b.g_per_spoon), rg = String(b.real_g == null ? '' : b.real_g).trim();
  if (!ing) throw new Error('缺食材');
  if (ml !== '' && !(Number(ml) > 0 && Number(ml) <= 500)) throw new Error('量匙規格要填 ml 數字');
  var hit = rows_('spoon_set').filter(function (r) { return r.ingredient === ing; })[0];
  var deco = hit ? hit.kind === 'deco' : String(b.kind || '') === 'deco';   /* 裝飾乾料：沒有食譜用克數 */
  if (!deco && !(g > 0 && g < 1000)) throw new Error('食譜用克數要大於 0');
  if (rg !== '' && !(Number(rg) > 0 && Number(rg) < 1000)) throw new Error('真實克數要大於 0（還沒實測就留空白）');
  var o = hit ? clean_(hit) : deco ? { ingredient: ing, max_spoons: '', tol_pct: '', kind: 'deco' } : { ingredient: ing, max_spoons: '6', tol_pct: '0' };
  o.spoon_ml = ml === '' ? '' : String(Number(ml)); o.g_per_spoon = deco ? '' : String(g); o.real_g = rg === '' ? '' : String(Number(rg));
  o.note = String(b.note || '').slice(0, 300); o.saved_at = now_();
  if (hit) writeRow_('spoon_set', hit._row, o);
  else append_('spoon_set', [o]);
  specLog_(o, String(b.version || '規格更新').slice(0, 40));
  log_('specSave', 'spoon_set', ing, JSON.stringify(o));
  return { ok: true, data: o };
}
/* 把頁面上「目前試算」記成一個版本（例：第二版 11/9 上線）；同一行（line_key）同一個版本名稱已經有的，先標刪除再寫，所以可以重記 */
function evoSnap_(b) {
  var ver = String(b.version || '').trim().slice(0, 40), vd = String(b.ver_date || '').trim(), rows = b.rows || [];
  if (!ver) throw new Error('缺版本名稱');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(vd)) throw new Error('生效日期格式要 yyyy-mm-dd');
  if (!rows.length) throw new Error('沒有要記的項目');
  if (rows.length > 600) throw new Error('一次最多 600 項');
  var t = now_(), keys = {}, del = 0;
  rows.forEach(function (r) { if (!r.line_key) throw new Error('缺 line_key'); keys[r.line_key] = 1; });
  rows_('spoon_evo').forEach(function (r) {
    if (r.deleted !== 'Y' && r.version === ver && keys[r.line_key]) { var c = clean_(r); c.deleted = 'Y'; c.updated_at = t; writeRow_('spoon_evo', r._row, c); del++; }
  });
  var base = 'E' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMddHHmmss'), F = ['line_key', 'recipe', 'recipe_id', 'step', 'step_title', 'ingredient', 'grams', 'spoons', 'spoon_ml', 'g_per_spoon', 'real_g', 'note'];
  var add = rows.map(function (r, i) {
    var o = { evo_id: base + ('000' + i).slice(-4), version: ver, ver_date: vd, source: '專案推動頁記錄', updated_at: t, deleted: '' };
    F.forEach(function (f) { o[f] = String(r[f] == null ? '' : r[f]).slice(0, 300); });
    return o;
  });
  append_('spoon_evo', add);
  log_('evoSnap', 'spoon_evo', ver, '寫入 ' + add.length + ' 列、取代 ' + del + ' 列，生效 ' + vd);
  return { ok: true, data: { added: add, replaced: del } };
}

/* ---------------- v2.0 定案明細 ---------------- */
/* 跟頁面 judge() 同一套：平匙數＝重量 ÷ 一平匙克數，取最近的整數或 .5；誤差內、不超過上限、至少半匙才可換 */
function judge_(grams, g, maxn, tol) {
  var n = grams / g, h = Math.round(n * 2) / 2, diff = h > 0 ? (h * g - grams) / grams : -1;
  var hit = Math.abs(n - h) < 1e-6 || (tol > 0 && Math.abs(diff) <= tol / 100 + 1e-9);
  if (h < 0.5) return { h: h, ok: false, why: '不到半平匙' };
  if (!hit) return { h: h, ok: false, why: '不是整數或半平匙' };
  if (maxn > 0 && h > maxn) return { h: h, ok: false, why: '超過 ' + maxn + ' 平匙，太多' };
  return { h: h, ok: true, why: '' };
}
function lineKey_(o) { return [o.recipe_id, o.step_no, o.ingredient, Number(o.grams), o.seq || 1].join('|'); }
function numStr_(x) { return String(Math.round(Number(x) * 1000) / 1000); }
function finalConfirm_(b) {
  var ing = String(b.ingredient || '').trim(), g = Number(b.g_per_spoon), maxn = Number(b.max_spoons) || 0, tol = Number(b.tol_pct) || 0;
  if (!ing) throw new Error('缺食材');
  if (!(g > 0 && g < 1000)) throw new Error('一平匙克數要大於 0');
  var lines = rows_('spoon_lines').filter(function (l) { return l.ingredient === ing; });
  if (!lines.length) throw new Error('食譜用量裡沒有「' + ing + '」');
  if (lines.some(function (l) { return !l.recipe_id || !l.step_no; })) throw new Error('食譜用量還沒記到後台編號和第幾步（要先按「🔄 更新食譜用量」一次），先不能確認');
  if (b.snap_date && lines[0].snap_date !== String(b.snap_date)) throw new Error('食譜用量在你打開頁面之後更新過（' + lines[0].snap_date + '），請重新整理頁面、看過再確認');
  var t = now_(), batch = 'F' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyMMdd-HHmmss'), uses = b.uses || {}, ml = String(b.spoon_ml == null ? '' : b.spoon_ml).trim();
  var old = rows_('spoon_final').filter(function (r) { return r.deleted !== 'Y' && r.ingredient === ing; }), oldBy = {};
  old.forEach(function (r) { (oldBy[r.line_key] = oldBy[r.line_key] || []).push(r); });
  var newKeys = {}, add = [], kept = 0, keptList = [];
  lines.forEach(function (l, i) {
    var lk = lineKey_(l), j = judge_(Number(l.grams), g, maxn, tol), o = (oldBy[lk] || [])[0];
    newKeys[lk] = 1;
    var row = { fin_id: batch + '-' + ('000' + i).slice(-4), batch: batch, ingredient: ing, recipe: l.recipe, recipe_id: l.recipe_id, step_no: l.step_no, step_title: l.step_title, grams: l.grams, seq: l.seq || '1', line_key: lk,
      g_per_spoon: numStr_(g), spoon_ml: ml, max_spoons: String(maxn), tol_pct: String(tol), calc_spoons: numStr_(j.h), calc_ok: j.ok ? 'Y' : 'N', calc_why: j.why,
      manual_spoons: '', manual_reason: '', manual_at: '', final_spoons: j.ok ? numStr_(j.h) : '', uses: uses[lk] == null ? (o ? o.uses : '') : String(uses[lk]),
      purpose: l.purpose || '', use_step: l.use_step || '', source: '平匙換算頁確認', confirmed_at: t, applied_at: '', applied_batch: '', updated_at: t, deleted: '', replaced_by: '' };
    if (o && o.manual_spoons !== '') {   /* 手動改過的照搬，不被重算蓋掉 */
      row.manual_spoons = o.manual_spoons; row.manual_reason = o.manual_reason; row.manual_at = o.manual_at; row.final_spoons = o.final_spoons;
      kept++; keptList.push(l.recipe + '：' + o.manual_spoons + (o.manual_spoons === '秤重' ? '' : ' 平匙'));
    }
    add.push(row);
  });
  var replaced = 0, untouched = 0;
  old.forEach(function (r) {
    if (newKeys[r.line_key]) { var c = clean_(r); c.deleted = 'Y'; c.replaced_by = batch; c.updated_at = t; writeRow_('spoon_final', r._row, c); replaced++; }
    else untouched++;
  });
  append_('spoon_final', add);
  /* 定案的規格同步存回平匙規格表（跟按「💾 存成設定」一樣，另記一列規格紀錄） */
  var sh = rows_('spoon_set').filter(function (r) { return r.ingredient === ing; })[0];
  var so = sh ? clean_(sh) : { ingredient: ing, real_g: '', kind: '' };
  so.g_per_spoon = numStr_(g); so.max_spoons = String(maxn); so.tol_pct = String(tol); if (ml !== '') so.spoon_ml = numStr_(ml);
  so.note = '確認轉換到執行進度（' + batch + '）'; so.saved_at = t;
  if (sh) writeRow_('spoon_set', sh._row, so); else append_('spoon_set', [so]);
  specLog_(so, '確認轉換到執行進度');
  var okN = add.filter(function (r) { return r.final_spoons !== ''; }).length;
  log_('finalConfirm', 'spoon_final', ing, batch + '：' + add.length + ' 處（換 ' + okN + '、秤重 ' + (add.length - okN) + '）、取代舊列 ' + replaced + '、手動照搬 ' + kept + '、快照外保留 ' + untouched);
  return { ok: true, data: { batch: batch, rows: add, total: add.length, ok: okN, keep: add.length - okN, replaced: replaced, manual_kept: kept, manual_list: keptList, untouched: untouched, set: so } };
}
/* 手動調平匙數：value＝數字（0.5 的倍數，0.5～20）、'秤重'（改成維持秤重）、''（恢復系統算的） */
function finalEdit_(b) {
  var id = String(b.fin_id || ''), v = String(b.value == null ? '' : b.value).trim(), why = String(b.reason || '').trim();
  var hit = rows_('spoon_final').filter(function (r) { return r.fin_id === id && r.deleted !== 'Y'; })[0];
  if (!hit) throw new Error('找不到這一處（可能剛被重新確認取代了，請重新整理）');
  if (hit.applied_at) throw new Error('這一處已經套用到門市食譜後台（' + hit.applied_batch + '），不能在這裡改；要改請跟 Claude 說');
  var o = clean_(hit), before = o.final_spoons === '' ? '秤重' : o.final_spoons, t = now_();
  if (v === '') { o.manual_spoons = ''; o.manual_reason = ''; o.manual_at = ''; o.final_spoons = o.calc_ok === 'Y' ? o.calc_spoons : ''; }
  else {
    if (why.length < 2) throw new Error('請寫手動調整的原因');
    if (v === '秤重') o.final_spoons = '';
    else {
      var n = Number(v);
      if (!(n > 0) || Math.abs(n * 2 - Math.round(n * 2)) > 1e-9) throw new Error('平匙數只能是 0.5 的倍數（例 1、1.5、2）');
      if (n < 0.5 || n > 20) throw new Error('平匙數要在 0.5～20 之間');
      v = numStr_(n); o.final_spoons = v;
    }
    o.manual_spoons = v; o.manual_reason = why.slice(0, 300); o.manual_at = t;
  }
  o.updated_at = t;
  writeRow_('spoon_final', hit._row, o);
  log_('finalEdit', 'spoon_final', id, o.recipe + '｜' + o.ingredient + ' ' + o.grams + ' g：' + before + ' → ' + (o.final_spoons === '' ? '秤重' : o.final_spoons) + (v === '' ? '（恢復系統算的）' : '（手動：' + why + '）'));
  return { ok: true, data: o };
}
/* 記錄哪幾處已套用到門市食譜後台（undo＝清掉，後台還原時用） */
function finalApplied_(b) {
  var ids = b.fin_ids || [], batch = String(b.batch || '').trim(), d = String(b.date || today_());
  if (!ids.length) throw new Error('沒有要記的項目');
  if (!b.undo && !batch) throw new Error('缺套用批號');
  var want = {}; ids.forEach(function (x) { want[x] = 1; });
  var n = 0, t = now_();
  rows_('spoon_final').forEach(function (r) {
    if (r.deleted === 'Y' || !want[r.fin_id]) return;
    var o = clean_(r); o.applied_at = b.undo ? '' : d; o.applied_batch = b.undo ? '' : batch; o.updated_at = t; writeRow_('spoon_final', r._row, o); n++;
  });
  if (n !== ids.length) throw new Error('只找到 ' + n + ' / ' + ids.length + ' 處（其餘可能被重新確認取代），已記錄找到的');
  log_('finalApplied', 'spoon_final', batch, (b.undo ? '取消 ' : '') + n + ' 處');
  return { ok: true, data: { n: n } };
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
/* 排程不用程式建（避免多要 script.scriptapp 權限）：Apps Script 編輯器左側「觸發條件」→ 新增 → 函式 dailyReminder、時間驅動、日計時器、上午 8 點至 9 點（2026-10-18 建）。 */

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
  var q = MailApp.getRemainingDailyQuota();
  Logger.log('授權完成：試算表分頁 ' + n + ' 張、權限中心連得到、今日還可寄 ' + q + ' 封信');
}
/* 建好全部分頁並放入第一批資料（只在分頁不存在時建；已存在就補缺的欄位） */
function setupTables() {
  Object.keys(T).forEach(function (k) { sheet_(k); });
  var d = ss_().getSheetByName('工作表1'); if (d && d.getLastRow() === 0 && ss_().getSheets().length > 1) ss_().deleteSheet(d);
  Logger.log(JSON.stringify({ projects: live_('projects').length, phases: live_('phases').length, decisions: live_('decisions').length, risks: live_('risks').length, people: live_('people').length, spoon_lines: rows_('spoon_lines').length, spoon_set: rows_('spoon_set').length, spoon_evo: rows_('spoon_evo').length, spoon_spec_log: rows_('spoon_spec_log').length, hdr: _hdr }));
}
/* 寫入流程自我檢查：新增測試專案 → 改 → 刪；跑完只留 deleted=Y 的一列（以及 log） */
function selfTest() {
  var a = save_('projects', { name: 'zz 測試專案（自我檢查，會自動刪）', status: '未開始' }).data;
  var b = save_('projects', { project_id: a.project_id, status: '進行中' }).data;
  del_('projects', a.project_id);
  var gone = live_('projects').every(function (r) { return r.project_id !== a.project_id; });
  Logger.log(JSON.stringify({ added: a.project_id, edited: b.status === '進行中', deleted: gone }));
}
