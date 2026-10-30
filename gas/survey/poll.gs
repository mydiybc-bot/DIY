/* =====================================================================
   🗳 門市內部調查（2026-10-30 加；跟顧客問卷共用同一個 GAS 與資料庫試算表，不必另外授權）
   用途：總部要各店回報某件事（例：擠花袋破損狀況），一份調查＝dim_poll 一列＋資料庫一張 poll_<代號> 分頁。
   前端：static/poll.html?p=<代號>（店員填）；&view=result（看結果：各店最新一筆＋誰還沒填）
         static/dashboard-poll.html（總部「🗳 門市問卷管理」：新增／編輯／結束收件／刪除／看結果；閘門 DASH_ID=poll）

   動作
     GET  ?action=poll_form&p=<代號>      → {ok, poll:{id,title,intro,status,fields}, stores}
     GET  ?action=poll_result&p=<代號>    → {ok, poll, stores, rows:[…全部回覆，舊→新…]}
     GET  ?action=poll_list               → {ok, polls:[{id,title,status,n,stores,last,created,updated,fields}…]}
     POST {action:'poll_submit', p, id, a:{欄位:答案}, ua}              → {ok, id, dup}（已結束的調查不收）
     POST {action:'poll_save', pwd, by, poll:{id?, title, intro, fields}} → {ok, id}（id 空白＝新增）
     POST {action:'poll_status', pwd, p, status:'open'|'closed'|'deleted'} → {ok}
   設定表 dim_poll（id, title, intro, status, fields_json, kmax, created, updated, by）；每次存檔前舊版記進 dim_poll_log。
   規則：同一家店可以重填，全部保留；結果以每店「最新一筆」為準。每份調查每天最多 POLL_CAP 筆。
   管理動作的密碼：權限中心 dim_auth 的 poll 列（pollAuthSetup 第一次建立，密碼先同 hub_hq）。
   ===================================================================== */
var POLL_CAP = 300;
var POLL_DASH = 'poll', POLL_DASH_NAME = '門市問卷管理';
var POLL_HEAD = ['id', 'title', 'intro', 'status', 'fields_json', 'kmax', 'created', 'updated', 'by'];
var POLL_TYPES = ['single', 'multi', 'text', 'number'];

/* 第一次讀 dim_poll 時種進去的調查（之後一律在儀表板改，改這裡沒有用） */
var POLL_SEED = {
  pipingbag: {
    title: '擠花袋使用狀況調查',
    intro: '想確認擠花袋破損是不是同一批產品的品質問題，以及各店實際受影響的狀況，後續會再統一評估處理方式。每家店填一份就好，資料有變可以重填（以最新一筆為準）。謝謝大家！',
    fields: [
      { k: 'store', type: 'single', title: '店別', options: 'STORES', required: true },
      { k: 'name', type: 'text', title: '填寫人姓名', required: true, max: 30 },
      { k: 'stock', type: 'number', title: '目前店內擠花袋庫存量', unit: '個', required: true, min: 0, max: 100000,
        hint: '還沒開封＋已開封的全部算進去；以「個」計' },
      { k: 'stock_note', type: 'text', title: '庫存備註（選填）', required: false, max: 200,
        hint: '例：還有 3 包未開、1 包已開；或不同尺寸各幾個' },
      { k: 'freq', type: 'single', title: '是否曾遇到擠花袋破損？', options: ['尚未遇到', '偶爾', '多次', '頻繁'], required: true },
      { k: 'pos', type: 'multi', title: '破損主要發生在哪個位置？（可複選）', required: true,
        options: ['接縫處（側邊黏合線）', '袋尖（剪口附近）', '袋身中段', '袋口（上方開口）', '其他'],
        cond: { k: 'freq', vals: ['偶爾', '多次', '頻繁'] }, hint: '請特別確認是不是「接縫處」裂開' },
      { k: 'pos_other', type: 'text', title: '「其他」是哪裡？', required: true, max: 100, cond: { k: 'pos', vals: ['其他'] } },
      { k: 'batch', type: 'text', title: '包裝上的批號／製造日期／進貨日期（選填）', required: false, max: 100,
        cond: { k: 'freq', vals: ['偶爾', '多次', '頻繁'] }, hint: '看得到就寫，幫助判斷是不是同一批' },
      { k: 'note', type: 'text', title: '其他補充（選填）', required: false, max: 500, long: true,
        hint: '例：大概從什麼時候開始、裝什麼材料時破、一天破幾個' }
    ]
  }
};

/* ---------------- 設定表 ---------------- */
function pollDim_() {
  var sh = sheet_('dim_poll', POLL_HEAD, [1, 7, 8]), P = PropertiesService.getScriptProperties();
  if (!P.getProperty('POLL_SEEDED')) {
    var have = {}; if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = 1; });
    Object.keys(POLL_SEED).forEach(function (id) {
      if (have[id]) return;
      var s = POLL_SEED[id], now = tz_(new Date(), 'yyyy-MM-dd HH:mm');
      sh.appendRow([id, s.title, s.intro, 'open', JSON.stringify(s.fields), 0, now, now, 'Claude（種子）']);
    });
    P.setProperty('POLL_SEEDED', '1');
  }
  return sh;
}
function pollAllRaw_() {
  var c = CacheService.getScriptCache(), hit = c.get('polls:v1');
  if (hit) return JSON.parse(hit);
  var sh = pollDim_(), out = [];
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, POLL_HEAD.length).getValues().forEach(function (r, i) {
    var fields = []; try { fields = JSON.parse(r[4] || '[]'); } catch (e) { }
    function ds(x) { return x instanceof Date ? tz_(x, 'yyyy-MM-dd HH:mm') : String(x); }   /* 試算表會把日期字串自動轉成日期 */
    out.push({ row: i + 2, id: String(r[0]), title: String(r[1]), intro: String(r[2]), status: String(r[3] || 'open'), fields: fields,
      kmax: Number(r[5]) || 0, created: ds(r[6]), updated: ds(r[7]), by: String(r[8]) });
  });
  try { c.put('polls:v1', JSON.stringify(out), 120); } catch (e) { }
  return out;
}
function pollBust_() { CacheService.getScriptCache().remove('polls:v1'); }
function pollFind_(p) {
  var id = String(p || ''), all = pollAllRaw_();
  for (var i = 0; i < all.length; i++) if (all[i].id === id && all[i].status !== 'deleted') return all[i];
  throw fail_('找不到這份調查（可能已被刪除）', 'poll');
}
function pollGet_(p) {
  var cfg = pollFind_(p), storeNames = STORES.map(function (s) { return s[1]; });
  return { id: cfg.id, title: cfg.title, intro: cfg.intro, status: cfg.status,
    fields: cfg.fields.map(function (f) { var o = {}; for (var x in f) o[x] = f[x]; if (o.options === 'STORES') o.options = storeNames; return o; }) };
}
function pollSheet_(pid, keys) {
  var ss = db_(), name = 'poll_' + pid, sh = ss.getSheetByName(name);
  var head = ['time', 'id', 'store_no', 'store'].concat(keys.filter(function (k) { return k !== 'store'; }), ['ua']);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1); }
  return sh;
}
function pollForm_(p) { return { ok: true, ver: VER, poll: pollGet_(p.p), stores: STORES }; }

/* ---------------- 店員送出 ---------------- */
function pollSubmit_(b) {
  var poll = pollGet_(b.p), id = clean_(b.id, 64);
  if (poll.status !== 'open') throw fail_('這份調查已經結束收件，謝謝！', 'closed');
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw fail_('缺少識別碼', 'input');
  var a = (b.a && typeof b.a === 'object') ? b.a : {}, ans = {}, missing = [];
  function shown(f) {
    if (!f.cond) return true;
    var pv = ans[f.cond.k]; if (!pv) return false;
    var arr = String(pv).split('|'); return f.cond.vals.some(function (x) { return arr.indexOf(x) >= 0; });
  }
  poll.fields.forEach(function (f) {
    if (!shown(f)) { ans[f.k] = ''; return; }
    var v = a[f.k];
    if (f.type === 'multi') {
      v = Array.isArray(v) ? v : (v == null || v === '' ? [] : [v]);
      v = v.map(function (x) { return clean_(x, 200); }).filter(function (x, i, arr) { return f.options.indexOf(x) >= 0 && arr.indexOf(x) === i; });
      if (f.required && !v.length) missing.push(f.title);
      ans[f.k] = v.join('|');
    } else if (f.type === 'number') {
      var s = clean_(v, 20);
      if (s === '') { if (f.required) missing.push(f.title); ans[f.k] = ''; return; }
      var n = Number(s), lo = f.min == null || f.min === '' ? 0 : Number(f.min), hi = f.max == null || f.max === '' ? 1e9 : Number(f.max);
      if (!isFinite(n) || n < lo || n > hi || Math.floor(n) !== n) throw fail_('「' + f.title + '」請填 ' + lo + '～' + hi + ' 的整數', 'input');
      ans[f.k] = n;
    } else if (f.type === 'text') {
      v = clean_(v, f.max || 500);
      if (f.required && !v) missing.push(f.title);
      ans[f.k] = v;
    } else {
      v = clean_(Array.isArray(v) ? v[0] : v, 200);
      if (v && f.options.indexOf(v) < 0) throw fail_('「' + f.title + '」的答案不在選項裡', 'input');
      if (f.required && !v) missing.push(f.title);
      ans[f.k] = v;
    }
  });
  if (missing.length) throw fail_('還沒填：' + missing.join('、'), 'missing');

  var c = CacheService.getScriptCache();
  if (c.get('pid:' + id)) return { id: id, dup: true };
  var keys = poll.fields.map(function (f) { return f.k; });
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var P = PropertiesService.getScriptProperties(), ck = 'pcnt_' + poll.id + '_' + today_(), n = Number(P.getProperty(ck) || 0);
    if (n >= POLL_CAP) throw fail_('今天這份調查的填寫次數已達上限，請明天再試或聯絡總部', 'cap');
    P.setProperty(ck, String(n + 1));
    var sh = pollSheet_(poll.id, keys), head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    var need = keys.filter(function (k) { return head.indexOf(k) < 0; });   /* 調查之後加了題目 → 補在 ua 前面 */
    if (need.length) {
      var uaAt = head.indexOf('ua'); if (uaAt < 0) { head.push('ua'); uaAt = head.length - 1; }
      head = head.slice(0, uaAt).concat(need, head.slice(uaAt));
      sh.getRange(1, 1, 1, head.length).setValues([head]);
    }
    var row = { time: new Date(), id: id, store_no: storeNo_(ans.store), store: ans.store, ua: clean_(b.ua, 160) };
    sh.appendRow(head.map(function (h) { return h in row ? row[h] : (h in ans ? ans[h] : ''); }));
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  c.put('pid:' + id, '1', 21600);
  c.remove('plist:v1');
  return { id: id, dup: false };
}

/* ---------------- 結果、清單 ---------------- */
function pollRows_(pid) {
  var sh = db_().getSheetByName('poll_' + pid), rows = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), head = v[0].map(String);
    for (var i = 1; i < v.length; i++) {
      var o = {};
      head.forEach(function (h, j) { if (h === 'ua') return; var x = v[i][j]; o[h] = x instanceof Date ? tz_(x, 'yyyy-MM-dd HH:mm') : x; });
      rows.push(o);
    }
  }
  return rows;
}
function pollResult_(p) {
  var poll = pollGet_(p.p);
  return { ok: true, ver: VER, poll: poll, stores: STORES, rows: pollRows_(poll.id), at: tz_(new Date(), 'yyyy-MM-dd HH:mm') };
}
function pollList_() {
  var c = CacheService.getScriptCache(), hit = c.get('plist:v1');
  if (hit) return JSON.parse(hit);
  var polls = pollAllRaw_().filter(function (x) { return x.status !== 'deleted'; }).map(function (x) {
    var rows = pollRows_(x.id), st = {};
    rows.forEach(function (r) { if (r.store) st[r.store] = 1; });
    return { id: x.id, title: x.title, intro: x.intro, status: x.status, fields: x.fields, n: rows.length, stores: Object.keys(st).length,
      last: rows.length ? rows[rows.length - 1].time : '', created: x.created, updated: x.updated, by: x.by };
  });
  polls.reverse();   /* 新的在前 */
  var o = { ok: true, ver: VER, polls: polls, n_stores: STORES.length, at: tz_(new Date(), 'yyyy-MM-dd HH:mm') };
  try { c.put('plist:v1', JSON.stringify(o), 60); } catch (e) { }
  return o;
}

/* ---------------- 管理（要密碼） ---------------- */
function pollAuth_(pwd) {
  var c = CacheService.getScriptCache(), bad = Number(c.get('ppwbad') || 0);
  if (bad >= 10) throw fail_('密碼錯太多次，請 10 分鐘後再試', 'locked');
  var v = SpreadsheetApp.openById(AUTH_SS).getSheetByName('dim_auth').getDataRange().getValues(), ok = false, found = false;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === POLL_DASH) {
    found = true;
    var real = String(v[i][2] == null ? '' : v[i][2]), on = String(v[i][4] || 'Y').toUpperCase() !== 'N';
    ok = on && real !== '' && String(pwd == null ? '' : pwd) === real;
  }
  if (!found) throw fail_('權限中心還沒有「門市問卷管理」的密碼列，請先執行 pollAuthSetup', 'auth');
  if (!ok) { c.put('ppwbad', String(bad + 1), 600); throw fail_('密碼不對（要用「門市問卷管理」這組密碼）', 'auth'); }
}
/* 檢查並整理題目；有問題整批擋下，訊息用白話 */
function pollClean_(poll, old) {
  var title = clean_(poll.title, 80), intro = clean_(poll.intro, 1000);
  if (!title) throw fail_('問卷名稱不能空白', 'input');
  var src = Array.isArray(poll.fields) ? poll.fields : [], kmax = old ? old.kmax : 0, used = {}, out = [];
  (old ? old.fields : []).forEach(function (f) { used[f.k] = 1; var m = /^f(\d+)$/.exec(f.k); if (m) kmax = Math.max(kmax, Number(m[1])); });
  out.push({ k: 'store', type: 'single', title: '店別', options: 'STORES', required: true });   /* 第一題固定，系統管 */
  var seen = { store: 1 };
  src.forEach(function (f, i) {
    if (!f || f.k === 'store') return;
    var n = out.length + 1, t = clean_(f.title, 200), type = String(f.type || '');
    if (!t) throw fail_('第 ' + n + ' 題沒有題目文字', 'input');
    if (POLL_TYPES.indexOf(type) < 0) throw fail_('第 ' + n + ' 題「' + t + '」題型不對', 'input');
    var k = String(f.k || '');
    if (!k || !/^[A-Za-z][\w]{0,30}$/.test(k) || seen[k] || (!used[k] && !/^f\d+$/.test(k))) { kmax++; k = 'f' + kmax; }
    else { var m = /^f(\d+)$/.exec(k); if (m) kmax = Math.max(kmax, Number(m[1])); }
    seen[k] = 1;
    var o = { k: k, type: type, title: t, required: !!f.required };
    var hint = clean_(f.hint, 200); if (hint) o.hint = hint;
    if (type === 'single' || type === 'multi') {
      var opts = (Array.isArray(f.options) ? f.options : []).map(function (x) { return clean_(x, 120); }).filter(function (x) { return x; });
      if (opts.length < 2) throw fail_('第 ' + n + ' 題「' + t + '」至少要 2 個選項', 'input');
      opts.forEach(function (x, j) {
        if (x.indexOf('|') >= 0) throw fail_('第 ' + n + ' 題的選項不能有「|」符號：' + x, 'input');
        if (opts.indexOf(x) !== j) throw fail_('第 ' + n + ' 題有重複的選項：' + x, 'input');
      });
      o.options = opts;
    } else if (type === 'number') {
      o.unit = clean_(f.unit, 10);
      o.min = f.min === '' || f.min == null ? 0 : Number(f.min);
      o.max = f.max === '' || f.max == null ? 1000000 : Number(f.max);
      if (!isFinite(o.min) || !isFinite(o.max) || o.min > o.max) throw fail_('第 ' + n + ' 題「' + t + '」的最小／最大值不對', 'input');
    } else {
      o.long = !!f.long; o.max = o.long ? 1000 : 200;
    }
    if (f.cond && f.cond.k) {
      var dep = out.filter(function (x) { return x.k === f.cond.k || x._src === f.cond.k; })[0];
      if (!dep || (dep.type !== 'single' && dep.type !== 'multi')) throw fail_('第 ' + n + ' 題「' + t + '」的顯示條件要看「前面的單選或複選題」', 'input');
      var depOpts = dep.options === 'STORES' ? STORES.map(function (s) { return s[1]; }) : dep.options;
      var vals = (Array.isArray(f.cond.vals) ? f.cond.vals : []).map(function (x) { return clean_(x, 120); }).filter(function (x) { return depOpts.indexOf(x) >= 0; });
      if (!vals.length) throw fail_('第 ' + n + ' 題「' + t + '」的顯示條件至少要勾一個答案', 'input');
      o.cond = { k: dep.k, vals: vals };
    }
    o._src = String(f.k || '');
    out.push(o);
  });
  if (out.length < 2) throw fail_('至少要有一題（店別以外）', 'input');
  out.forEach(function (o) { delete o._src; });
  return { title: title, intro: intro, fields: out, kmax: kmax };
}
function pollSave_(b) {
  pollAuth_(b.pwd);
  var poll = b.poll || {}, by = clean_(b.by, 30) || '總部', now = tz_(new Date(), 'yyyy-MM-dd HH:mm');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    pollBust_();
    var sh = pollDim_(), id = clean_(poll.id, 40), old = null;
    if (id) { old = pollFind_(id); }
    var c = pollClean_(poll, old);
    if (old) {
      if (String(b.updated || '') && old.updated !== String(b.updated)) throw fail_('這份問卷剛剛被別人改過（' + old.updated + '，' + old.by + '），請重新整理再改', 'conflict');
      sheet_('dim_poll_log', ['saved_at', 'id', 'by', 'json'], [1]).appendRow([now, id, by, JSON.stringify({ title: old.title, intro: old.intro, fields: old.fields })]);
      sh.getRange(old.row, 2, 1, POLL_HEAD.length - 1).setValues([[c.title, c.intro, old.status, JSON.stringify(c.fields), c.kmax, old.created, now, by]]);
    } else {
      id = 'p' + tz_(new Date(), 'yyMMddHHmmss');
      sh.appendRow([id, c.title, c.intro, 'open', JSON.stringify(c.fields), c.kmax, now, now, by]);
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); pollBust_(); CacheService.getScriptCache().remove('plist:v1'); }
  return { id: id, updated: now };
}
function pollStatus_(b) {
  pollAuth_(b.pwd);
  var st = String(b.status || '');
  if (['open', 'closed', 'deleted'].indexOf(st) < 0) throw fail_('狀態不對', 'input');
  pollBust_();
  var old = pollFind_(b.p), now = tz_(new Date(), 'yyyy-MM-dd HH:mm'), sh = pollDim_();
  sh.getRange(old.row, 4).setValue(st);
  sh.getRange(old.row, 8, 1, 2).setValues([[now, clean_(b.by, 30) || '總部']]);
  SpreadsheetApp.flush();
  pollBust_(); CacheService.getScriptCache().remove('plist:v1');
  return { id: old.id, status: st };   /* 刪除只是藏起來：設定列與回覆分頁都留著 */
}

/* 第一次上線執行一次：權限中心 dim_auth 補「poll」列（密碼先同 hub_hq），先備份 */
function pollAuthSetup() {
  var ss = SpreadsheetApp.openById(AUTH_SS), sh = ss.getSheetByName('dim_auth');
  var v = sh.getDataRange().getValues(), before = v.length;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === POLL_DASH) return '已有 ' + POLL_DASH + ' 列（第 ' + (i + 1) + ' 列），沒動';
  var hq = null; for (var k = 1; k < v.length; k++) if (String(v[k][0]) === 'hub_hq') hq = String(v[k][2] || '');
  if (!hq) throw fail_('dim_auth 找不到 hub_hq 的密碼，沒有新增', 'auth');
  var bk = sh.copyTo(ss); bk.setName('dim_auth_備份_' + tz_(new Date(), 'yyyyMMdd-HHmm'));
  sh.appendRow([POLL_DASH, POLL_DASH_NAME, hq, 12, 'Y', new Date(), '2026-10-30 Claude 新增（門市問卷管理儀表板）；密碼暫用總部 hub 密碼，要另設請改這格']);
  return '新增 ' + POLL_DASH + '（' + before + ' → ' + sh.getLastRow() + ' 列，備份分頁 ' + bk.getName() + '）';
}
