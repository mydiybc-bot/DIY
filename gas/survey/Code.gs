/* =====================================================================
   📝 顧客問卷 API（GAS 專案「顧客問卷_資料庫」，2026-10-28 建）
   取代門市食譜後台每支食譜最後一頁嵌的 SurveyCake 問卷（顧客體驗滿意度調查 6.0，https://www.surveycake.com/s/wrg1P）。
   前端：static/survey.html（客人在店內平板填；後台 iframe 嵌這一頁）
         static/dashboard-survey.html（總部看結果；閘門 DASH_ID=survey）

   動作
     GET  ?action=ping[&callback=]                     → {ok, ver, today}（today＝今天已收幾筆）
     GET  ?action=form[&callback=]                     → {ok, ver, questions:[…], stores:[[店號,店名]…]}（題目設定，快取 5 分鐘）
     POST {action:'submit', id, recipe, src, lang, sec, ua, a:{qid:答案}} → {ok, id, dup}
     GET  ?action=stats&from=yyyy-MM-dd&to=yyyy-MM-dd[&callback=]
                                                        → {ok, rows:[…期間內每筆…], months:[…每月摘要（全部期間）…], questions, n_all, db_url}（快取 3 分鐘）
   資料：authorize() 第一次執行時建試算表「顧客問卷_資料庫」（指令碼屬性 DB_SS）
         dim_question ＝ 題目設定（改這張就改問卷，不必改程式；qid 沿用 SurveyCake 6.0 的題號）
         fact_response ＝ 每筆回答一列，只新增（time, id, store_no, store, recipe, src, lang, sec, q1…, ua）
   防濫用：網址公開、沒有密碼 → 每天全站最多 SUBMIT_CAP 筆；同一個 id 6 小時內只收一次。
   ===================================================================== */
var VER = 'survey-2026-10-28.2';
var SUBMIT_CAP = 3000;                 /* 12 店一天最多約 1,200 位客人 */
var TEXT_MAX = 500;
var AUTH_SS = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4';   /* 儀表板權限中心（dim_auth） */
var DASH_ID = 'survey', DASH_NAME = '顧客問卷';

/* 店號 ↔ 店名（問卷 q1 的選項；順序照北→南給客人看，店號照公司編號） */
var STORES = [[3, '台北南京'], [4, '台北士林'], [11, '吳寶春自己做（台北信義A13）'], [7, '新北板橋'], [8, '新北新店'],
  [9, '桃園中壢'], [10, '桃園藝文'], [6, '新竹文化'], [1, '台中精明'], [2, '台中草悟道'], [5, '台南Focus'], [12, '吳寶春自己做（高雄SKM Park）']];

/* 題目種子（＝SurveyCake 6.0 原題；q3、q10 是 6.0 的條件題：前一題選特定答案才出現；q17 是新加的留言題）
   欄：qid, order, type(single／multi／text), title, options（| 分隔）, required(Y/N), enabled(Y/N), note,
       cond（出現條件，空白＝一定出現；寫法「q2=第1次」，多個答案用 | 分隔「q9=是|否」；條件不成立時這題不顯示、不必填） */
var Q_HEAD = ['qid', 'order', 'type', 'title', 'options', 'required', 'enabled', 'note', 'cond'];
var Q_SCHEMA = '2';   /* 2026-10-28.2：加 cond 欄＋補 q3、q10；舊表第一次讀取時自動補（qMigrate_） */
function seedQuestions_() {
  var storeOpts = STORES.map(function (s) { return s[1]; }).join('|');
  return [
    ['q1', 1, 'single', '您今日體驗的分店是？', storeOpts, 'Y', 'Y', '店別；平板會記住上一次選的店'],
    ['q2', 2, 'single', '請問這是您「一年內」第幾次來做甜點呢？（不限店別）', '第1次|第2次|第3次|第4次|第5次以上', 'Y', 'Y', ''],
    ['q3', 2.5, 'single', '請問您入座後，服務人員的解說：', '清楚明瞭，我知道接下來該如何完成甜點製作|解說太簡短，讓我不知該如何開始|解說太冗長，根本聽不懂說什麼', 'Y', 'Y', 'SurveyCake 6.0 條件題：第一次來的客人才問', 'q2=第1次'],
    ['q4', 3, 'single', '您今天選這款甜點時，第一個考慮的是？', '外型|價格|口味|不容易失敗|易分享|可否當生日蛋糕|看過別人分享|現在流行|當季限定|店內熱銷品|其他', 'Y', 'Y', ''],
    ['q5', 4, 'single', '您這次來的原因是？', '重要的人生日|自己生日|情侶約會|朋友聚會|家人同樂|喜歡烘焙|送禮|其他', 'Y', 'Y', ''],
    ['q6', 5, 'single', '您是否滿意今天使用的平板教學影片？', '是|否', 'Y', 'Y', '儀表板 KPI：平板教學滿意'],
    ['q7', 6, 'single', '您是否滿意今天店內的環境？', '是|否', 'Y', 'Y', '儀表板 KPI：環境滿意'],
    ['q8', 7, 'single', '您是否滿意今天服務人員的服務？', '是|否', 'Y', 'Y', '儀表板 KPI：服務滿意'],
    ['q9', 8, 'single', '在整體消費過程中，服務人員是否有提供您任何的協助？', '是|否', 'Y', 'Y', ''],
    ['q10', 8.5, 'multi', '服務人員協助您的地方是？（可複選）', '報到時熱情問候|協助甜點品項的選擇|製作過程時的協助，如補救、教學、設備使用、蛋糕抹面、找食材／器模具…|給予讚美鼓勵|其他', 'Y', 'Y', 'SurveyCake 6.0 條件題：上一題回答「是」才問', 'q9=是'],
    ['q11', 9, 'multi', '您今天做的甜點是？（可複選）', '乳酪類（名字有乳酪，包含巴斯克、提拉米蘇）|慕斯類|蛋糕類|塔派類|點心類|餅乾類|麵包類', 'Y', 'Y', ''],
    ['q12', 10, 'single', '您今天的甜點有製作成功嗎？', '是|否', 'Y', 'Y', '儀表板 KPI：製作成功'],
    ['q13', 11, 'single', '您今天製作甜點的時間？', '2 小時（含）以內|2～2.5 小時（含）|2.5～3 小時（含）|3 小時以上', 'Y', 'Y', ''],
    ['q14', 12, 'single', '您會推薦其他人來嗎？', '會|不會', 'Y', 'Y', '儀表板 KPI：推薦率'],
    ['q15', 13, 'single', '請問您的年齡？', '14 歲以下|15～19 歲|20～24 歲|25～29 歲|30～34 歲|35～39 歲|40～44 歲|45～49 歲|50 歲以上', 'Y', 'Y', ''],
    ['q16', 14, 'single', '請問您是透過哪個途徑得知自己做呢？', '自己做官方FB|自己做官方IG|自己做官方Line|朋友IG/FB分享|親朋好友介紹|Line社群|廣告|Google搜尋|部落客介紹|我是老朋友', 'Y', 'Y', ''],
    ['q17', 15, 'text', '還有什麼想跟我們說的嗎？（選填）', '', 'N', 'Y', '2026-10-28 新加；客人留言，儀表板「客人留言」分頁']
  ];
}
var R_FIXED = ['time', 'id', 'store_no', 'store', 'recipe', 'src', 'lang', 'sec'];   /* fact_response 固定欄；之後是各 qid；最後 ua */

function fail_(msg, code) { var e = new Error(msg); e.code = code || 'err'; return e; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function errMsg_(e) { return String((e && e.message) || e); }
function tz_(d, f) { return Utilities.formatDate(d, 'Asia/Taipei', f); }
function today_() { return tz_(new Date(), 'yyyyMMdd'); }

function doGet(e) {
  var p = (e && e.parameter) || {}, o;
  try {
    if (p.action === 'ping' || !p.action) o = { ok: true, ver: VER, today: countToday_() };
    else if (p.action === 'form') o = form_();
    else if (p.action === 'stats') o = stats_(p);
    else throw fail_('不認得的動作', 'act');
  } catch (err) { o = { ok: false, code: err.code || 'err', msg: errMsg_(err) }; }
  if (p.callback && /^[A-Za-z_$][\w$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(o) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return out_(o);
}

function doPost(e) {
  var o;
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (b.action === 'submit') o = submit_(b);
    else if (b.action === 'ping') o = { ver: VER, today: countToday_() };
    else throw fail_('不認得的動作', 'act');
    o.ok = true; o.act = b.action;
  } catch (err) { o = { ok: false, code: err.code || 'err', msg: errMsg_(err) }; }
  return out_(o);
}

/* ---------------- 資料表 ---------------- */
function db_() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('DB_SS'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (_) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('顧客問卷_資料庫'); ss.setSpreadsheetTimeZone('Asia/Taipei');
    P.setProperty('DB_SS', ss.getId());
  }
  return ss;
}
function qSheet_() {
  var ss = db_(), sh = ss.getSheetByName('dim_question');
  if (!sh) {
    sh = ss.insertSheet('dim_question');
    var rows = [Q_HEAD].concat(seedQuestions_());
    rows = rows.map(function (r) { r = r.slice(); while (r.length < Q_HEAD.length) r.push(''); return r; });
    sh.getRange(1, 1, rows.length, Q_HEAD.length).setValues(rows); sh.setFrozenRows(1);
  }
  var P = PropertiesService.getScriptProperties();
  if (P.getProperty('Q_SCHEMA') !== Q_SCHEMA) { qMigrate_(sh); P.setProperty('Q_SCHEMA', Q_SCHEMA); }
  return sh;
}
/* 舊題目表升級：補 cond 欄、補種子裡有但表上沒有的題（只新增，不改經營者已改過的題） */
function qMigrate_(sh) {
  var v = sh.getDataRange().getValues(), head = v[0].map(function (h) { return String(h).trim(); });
  if (head.indexOf('cond') < 0) {
    var col = head.length + 1; for (var i = 0; i < head.length; i++) if (!head[i]) { col = i + 1; break; }
    sh.getRange(1, col).setValue('cond'); head[col - 1] = 'cond';
  }
  var have = {}; for (var j = 1; j < v.length; j++) have[String(v[j][0] || '').trim()] = true;
  seedQuestions_().forEach(function (r) {
    if (have[r[0]]) return;
    var row = head.map(function (h) { var k = Q_HEAD.indexOf(h); return k >= 0 && k < r.length ? r[k] : ''; });
    sh.appendRow(row);
  });
  SpreadsheetApp.flush();
}
function parseCond_(s) {
  s = String(s || '').trim(); var k = s.indexOf('='); if (k < 1) return null;
  var vals = s.slice(k + 1).split('|').map(function (x) { return x.trim(); }).filter(Boolean);
  return vals.length ? { qid: s.slice(0, k).trim(), vals: vals } : null;
}
function rSheet_() {
  var ss = db_(), sh = ss.getSheetByName('fact_response');
  if (!sh) {
    sh = ss.insertSheet('fact_response');
    var head = R_FIXED.concat(questions_(true).map(function (q) { return q.qid; }), ['ua']);
    sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1);
    var d = ss.getSheetByName('工作表1') || ss.getSheetByName('Sheet1'); if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
  }
  return sh;
}
/* 讀題目（all=true 連停用的也回；儀表板標舊資料要用） */
function questions_(all) {
  var v = qSheet_().getDataRange().getValues(), out = [], ci = v[0].map(function (h) { return String(h).trim(); }).indexOf('cond');
  for (var i = 1; i < v.length; i++) {
    var r = v[i], qid = String(r[0] || '').trim(); if (!qid) continue;
    var q = { qid: qid, order: Number(r[1]) || 999, type: String(r[2] || 'single').trim(), title: String(r[3] || '').trim(),
      options: String(r[4] || '').split('|').map(function (s) { return s.trim(); }).filter(Boolean),
      required: String(r[5] || '').trim().toUpperCase() === 'Y', enabled: String(r[6] || 'Y').trim().toUpperCase() !== 'N',
      cond: ci >= 0 ? parseCond_(r[ci]) : null };
    if (all || q.enabled) out.push(q);
  }
  out.sort(function (a, b) { return a.order - b.order; });
  return out;
}
function form_() {
  var c = CacheService.getScriptCache(), hit = c.get('form:v2');
  if (hit) return JSON.parse(hit);
  var o = { ok: true, ver: VER, questions: questions_(false), stores: STORES, at: new Date().toISOString() };
  c.put('form:v2', JSON.stringify(o), 300);
  return o;
}

/* ---------------- 每日筆數／上限 ---------------- */
function cntKey_() { return 'cnt_' + today_(); }
function countToday_() { return Number(PropertiesService.getScriptProperties().getProperty(cntKey_()) || 0); }
function capTake_() {
  var P = PropertiesService.getScriptProperties(), k = cntKey_(), n = Number(P.getProperty(k) || 0);
  if (n >= SUBMIT_CAP) throw fail_('今天收到的問卷已達上限（' + SUBMIT_CAP + ' 筆），請明天再試', 'cap');
  P.setProperty(k, String(n + 1));
  P.getKeys().forEach(function (x) { if (/^cnt_\d{8}$/.test(x) && x < 'cnt_' + tz_(new Date(Date.now() - 3 * 864e5), 'yyyyMMdd')) P.deleteProperty(x); });
}

/* ---------------- 收一筆 ---------------- */
function clean_(s, max) { return String(s == null ? '' : s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max || 200); }
function storeNo_(name) { for (var i = 0; i < STORES.length; i++) if (STORES[i][1] === name) return STORES[i][0]; return ''; }

function submit_(b) {
  var id = clean_(b.id, 64); if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw fail_('缺少識別碼', 'input');
  var a = (b.a && typeof b.a === 'object') ? b.a : {};
  var qs = questions_(false), ans = {}, missing = [];
  /* 條件題：條件不成立 → 不收、不必填（照題目順序判斷，前一題的答案先清乾淨） */
  function shown(q) {
    if (!q.cond) return true;
    var pv = ans[q.cond.qid]; if (pv == null) return false;
    var arr = String(pv).split('|'); return q.cond.vals.some(function (x) { return arr.indexOf(x) >= 0; });
  }
  qs.forEach(function (q) {
    if (!shown(q)) { ans[q.qid] = ''; return; }
    var v = a[q.qid];
    if (q.type === 'multi') {
      v = Array.isArray(v) ? v : (v == null || v === '' ? [] : [v]);
      v = v.map(function (x) { return clean_(x, 200); }).filter(function (x) { return q.options.indexOf(x) >= 0; });
      v = v.filter(function (x, i) { return v.indexOf(x) === i; });
      if (q.required && !v.length) missing.push(q.qid);
      ans[q.qid] = v.join('|');
    } else if (q.type === 'text') {
      v = clean_(v, TEXT_MAX);
      if (q.required && !v) missing.push(q.qid);
      ans[q.qid] = v;
    } else {
      v = clean_(Array.isArray(v) ? v[0] : v, 200);
      if (v && q.options.length && q.options.indexOf(v) < 0) throw fail_('第 ' + q.qid + ' 題的答案不在選項裡', 'input');
      if (q.required && !v) missing.push(q.qid);
      ans[q.qid] = v;
    }
  });
  if (missing.length) throw fail_('還有必填的題目沒回答：' + missing.join('、'), 'missing');

  /* 同一個 id 6 小時內只收一次（客人按兩次送出、或網路重送） */
  var c = CacheService.getScriptCache();
  if (c.get('id:' + id)) return { id: id, dup: true };

  var store = ans.q1 || '', store_no = storeNo_(store) || (Number(b.store_no) || '');
  if (!store && store_no) STORES.forEach(function (s) { if (s[0] === store_no) store = s[1]; });
  var row = { time: new Date(), id: id, store_no: store_no, store: store, recipe: clean_(b.recipe, 120), src: clean_(b.src, 20) || 'tablet',
    lang: clean_(b.lang, 10) || 'zh', sec: Math.min(Math.max(Number(b.sec) || 0, 0), 86400) };
  var ua = clean_(b.ua, 160);

  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    capTake_();
    var sh = rSheet_(), head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    /* 題目表新增了題號 → fact_response 自動補欄（加在 ua 前面） */
    var need = Object.keys(ans).filter(function (q) { return head.indexOf(q) < 0; });
    if (need.length) {
      var uaAt = head.indexOf('ua'); if (uaAt < 0) { head.push('ua'); uaAt = head.length - 1; }
      head = head.slice(0, uaAt).concat(need, head.slice(uaAt));
      sh.getRange(1, 1, 1, head.length).setValues([head]);
    }
    var vals = head.map(function (h) { return h === 'ua' ? ua : (h in row ? row[h] : (h in ans ? ans[h] : '')); });
    sh.appendRow(vals);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  c.put('id:' + id, '1', 21600);
  return { id: id, dup: false };
}

/* ---------------- 儀表板讀取 ---------------- */
function stats_(p) {
  var from = /^\d{4}-\d{2}-\d{2}$/.test(p.from || '') ? p.from : tz_(new Date(), 'yyyy-MM') + '-01';
  var to = /^\d{4}-\d{2}-\d{2}$/.test(p.to || '') ? p.to : tz_(new Date(), 'yyyy-MM-dd');
  var c = CacheService.getScriptCache(), key = 'stats:' + from + ':' + to, hit = c.get(key);
  if (hit) return JSON.parse(hit);
  var ss = db_(), sh = rSheet_(), v = sh.getDataRange().getValues(), head = v[0].map(String);
  var iT = head.indexOf('time'), iS = head.indexOf('src'), qCols = head.map(function (h, i) { return /^q\d+$/.test(h) ? i : -1; }).filter(function (i) { return i >= 0; });
  var yes = { q6: '是', q7: '是', q8: '是', q9: '是', q12: '是', q14: '會' };
  var rows = [], months = {}, nAll = 0;
  for (var i = 1; i < v.length; i++) {
    var r = v[i], t = r[iT]; if (!(t instanceof Date)) { t = new Date(t); if (isNaN(t)) continue; }
    if (String(r[iS]) === 'test') continue;
    nAll++;
    var day = tz_(t, 'yyyy-MM-dd'), ym = day.slice(0, 7);
    var m = months[ym] || (months[ym] = { ym: ym, n: 0, q6: 0, q7: 0, q8: 0, q9: 0, q12: 0, q14: 0 });
    m.n++;
    for (var q in yes) { var ci = head.indexOf(q); if (ci >= 0 && String(r[ci]) === yes[q]) m[q]++; }
    if (day < from || day > to) continue;
    var a = {}; qCols.forEach(function (ci) { var x = r[ci]; if (x !== '' && x != null) a[head[ci]] = String(x); });
    rows.push({ t: t.toISOString(), d: day, sn: r[head.indexOf('store_no')] || '', s: String(r[head.indexOf('store')] || ''),
      r: String(r[head.indexOf('recipe')] || ''), sec: Number(r[head.indexOf('sec')]) || 0, a: a });
  }
  rows.reverse();
  var cut = rows.length > 20000; if (cut) rows = rows.slice(0, 20000);
  var o = { ok: true, ver: VER, from: from, to: to, rows: rows, cut: cut, n_all: nAll,
    months: Object.keys(months).sort().map(function (k) { return months[k]; }),
    questions: questions_(true), stores: STORES, db_url: ss.getUrl(), at: new Date().toISOString() };
  var s = JSON.stringify(o); if (s.length < 90000) c.put(key, s, 180);   /* CacheService 一筆最多 100KB */
  return o;
}

/* ---------------- 第一次：建表、題目種子、權限中心補列（在編輯器執行一次） ---------------- */
function authorize() {
  var ss = db_(), q = qSheet_(), r = rSheet_(), a = authSetup_();
  var msg = '試算表：' + ss.getUrl() + '\n題目 ' + (q.getLastRow() - 1) + ' 題；回答表 ' + (r.getLastRow() - 1) + ' 筆\n權限中心：' + a + '\nping：' + VER;
  Logger.log(msg); return msg;
}
/* 權限中心 dim_auth 補一列 survey（已有就不動）。密碼先沿用總部 hub（hub_hq）；要另設請到權限管理後台改。 */
function authSetup_() {
  var ss = SpreadsheetApp.openById(AUTH_SS), sh = ss.getSheetByName('dim_auth');
  var v = sh.getDataRange().getValues(), before = v.length;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === DASH_ID) return '已有 ' + DASH_ID + ' 列（第 ' + (i + 1) + ' 列），沒動';
  var hq = null; for (var k = 1; k < v.length; k++) if (String(v[k][0]) === 'hub_hq') hq = String(v[k][2] || '');
  if (!hq) throw fail_('dim_auth 找不到 hub_hq 的密碼，沒有新增', 'auth');
  var bk = sh.copyTo(ss); bk.setName('dim_auth_備份_' + tz_(new Date(), 'yyyyMMdd-HHmm'));
  sh.appendRow([DASH_ID, DASH_NAME, hq, 12, 'Y', new Date(), '2026-10-28 Claude 新增（顧客問卷結果儀表板）；密碼暫用總部 hub 密碼，要另設請改這格']);
  return '新增 ' + DASH_ID + '（' + before + ' → ' + sh.getLastRow() + ' 列，備份分頁 ' + bk.getName() + '）';
}
