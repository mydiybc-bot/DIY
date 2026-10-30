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
var VER = 'survey-2026-10-30.11';   /* .10：加門市內部調查（poll.gs）；.11：問卷設定搬進 dim_poll＋管理動作 */
var SUBMIT_CAP = 3000;                 /* 12 店一天最多約 1,200 位客人 */
var TEXT_MAX = 500;
var AGG_DAY_FROM = '2026-01-01';   /* 這天以前的歷史回覆用「月」彙總（agg_daily 一店一月一列），這天起用「日」彙總 */
var AGG_HEAD = ['day', 'grain', 'store_no', 'store', 'n', 'secSum', 'secN', 'counts'];
var AGGM_HEAD = ['ym', 'n', 'q6', 'q7', 'q8', 'q9', 'q12', 'q14'];
var CM_HEAD = ['time', 'day', 'id', 'store_no', 'store', 'recipe', 'q14', 'q12', 'text'];
var YES = { q6: '是', q7: '是', q8: '是', q9: '是', q12: '是', q14: '會' };   /* 儀表板 KPI 的「好」答案 */
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
    else if (p.action === 'stats2') o = stats2_(p);
    else if (p.action === 'qlist') o = qlist_();
    else if (p.action === 'poll_form') o = pollForm_(p);       /* 門市內部調查（poll.gs） */
    else if (p.action === 'poll_result') o = pollResult_(p);
    else if (p.action === 'poll_list') o = pollList_();
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
    else if (b.action === 'import') o = import_(b);   /* 歷史匯入用，匯完會拿掉 */
    else if (b.action === 'qsave') o = qsave_(b);
    else if (b.action === 'qrestore') o = qrestore_(b);
    else if (b.action === 'qauth') { authPwd_(b.pwd); o = { auth: true }; }
    else if (b.action === 'poll_submit') o = pollSubmit_(b);   /* 門市內部調查（poll.gs） */
    else if (b.action === 'poll_save') o = pollSave_(b);
    else if (b.action === 'poll_status') o = pollStatus_(b);
    else if (b.action === 'poll_auth') { pollAuth_(b.pwd); o = { auth: true }; }
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
  var c = CacheService.getScriptCache(), ck = 'qs:' + (all ? 1 : 0), hit = c.get(ck);   /* 快取 2 分鐘：送出時不必每筆讀題目表（改題目最慢 2～5 分鐘生效） */
  if (hit) return JSON.parse(hit);
  var out = questionsRead_(all); c.put(ck, JSON.stringify(out), 120); return out;
}
function questionsRead_(all) {
  var v = qSheet_().getDataRange().getValues(), out = [], ci = v[0].map(function (h) { return String(h).trim(); }).indexOf('cond');
  for (var i = 1; i < v.length; i++) {
    var r = v[i], qid = String(r[0] || '').trim(); if (!qid) continue;
    var q = { qid: qid, order: Number(r[1]) || 999, type: String(r[2] || 'single').trim(), title: String(r[3] || '').trim(),
      options: String(r[4] || '').split('|').map(function (s) { return s.trim(); }).filter(Boolean),
      required: String(r[5] || '').trim().toUpperCase() === 'Y', enabled: String(r[6] || 'Y').trim().toUpperCase() !== 'N',
      cond: ci >= 0 ? parseCond_(r[ci]) : null, note: String(r[7] || '') };
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
  if (n === 0) P.getKeys().forEach(function (x) { if (/^cnt_\d{8}$/.test(x) && x < 'cnt_' + tz_(new Date(Date.now() - 3 * 864e5), 'yyyyMMdd')) P.deleteProperty(x); });   /* 每天第一筆才清舊計數 */
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

  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    capTake_();
    var sh = rSheet_(), hc = CacheService.getScriptCache(), hh = hc.get('rhead'), head = hh ? JSON.parse(hh) : sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    /* 題目表新增了題號 → fact_response 自動補欄（加在 ua 前面） */
    var need = Object.keys(ans).filter(function (q) { return head.indexOf(q) < 0; });
    if (need.length) {
      var uaAt = head.indexOf('ua'); if (uaAt < 0) { head.push('ua'); uaAt = head.length - 1; }
      head = head.slice(0, uaAt).concat(need, head.slice(uaAt));
      sh.getRange(1, 1, 1, head.length).setValues([head]);
    }
    hc.put('rhead', JSON.stringify(head), 600);
    var vals = head.map(function (h) { return h === 'ua' ? ua : (h in row ? row[h] : (h in ans ? ans[h] : '')); });
    sh.appendRow(vals);
    if (row.src !== 'test') {
      try { aggMerge_([itemOf_(row.time, row.store_no, row.store, row.sec, ans)], false); if (ans.q17) cmAppend_(row, ans); }
      catch (ae) { PropertiesService.getScriptProperties().setProperty('agg_dirty', tz_(new Date(), 'yyyy-MM-dd HH:mm') + ' ' + errMsg_(ae)); }
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  c.put('id:' + id, '1', 21600);
  return { id: id, dup: false };
}

/* ---------------- 彙總表（儀表板讀這個，不讀十幾萬筆原始回答） ----------------
   agg_daily：一列＝一天（或歷史的一個月）×一家店；counts 是 JSON：{"q4|口味":3, "#q4":10}（"#qid"＝這題有幾人回答）
   agg_month：一列＝一個月，全公司，月趨勢用
   fact_comment：只放有留言的回答（src=test 不放） */
function sheet_(name, head, textCols) {
  var ss = db_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1);
    (textCols || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });   /* 日期欄當純文字，Sheets 才不會自動轉成日期 */
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  return sh;
}
function aggSheet_() { return sheet_('agg_daily', AGG_HEAD, [1, 2]); }
function aggMonthSheet_() { return sheet_('agg_month', AGGM_HEAD, [1]); }
function cmSheet_() { return sheet_('fact_comment', CM_HEAD, [2]); }
/* 儀表板快取版本號，每次彙總有變就加 1 */
function bumpStats_() { var P = PropertiesService.getScriptProperties(); P.setProperty('stats_ver', String((Number(P.getProperty('stats_ver')) || 0) + 1)); }
/* 表格列數不夠就先加（setValues 不會自動加列） */
function growTo_(sh, rows) { var mx = sh.getMaxRows(); if (rows > mx) sh.insertRowsAfter(mx, rows - mx); }
function dayOf_(v) { if (v instanceof Date) return tz_(v, 'yyyy-MM-dd'); return String(v == null ? '' : v).slice(0, 10); }

/* 一筆回答 → 彙總用的小物件 */
function itemOf_(time, sn, store, sec, ans) {
  var t = time instanceof Date ? time : new Date(time), day = tz_(t, 'yyyy-MM-dd');
  return { day: day, sn: sn === '' || sn == null ? 0 : Number(sn) || 0, s: String(store || ''), sec: Number(sec) || 0, a: ans };
}
function countAns_(c, qs, a) {
  qs.forEach(function (q) {
    var v = a[q.qid]; if (v == null || String(v) === '') return;
    c['#' + q.qid] = (c['#' + q.qid] || 0) + 1;
    if (q.type === 'text') return;
    String(v).split('|').forEach(function (o) { if (!o) return; var k = q.qid + '|' + o; c[k] = (c[k] || 0) + 1; });
  });
}
/* 併進彙總表。full=true 讀全表找舊列（匯入用）；false 只看最後 150 列（新回答一定落在最近幾列） */
function aggMerge_(items, full) {
  if (!items.length) return;
  bumpStats_();   /* 彙總有變 → 儀表板快取立刻作廢 */
  var qs = questions_(true), groups = {};
  items.forEach(function (it) {
    var old = it.day < AGG_DAY_FROM, gday = old ? it.day.slice(0, 7) + '-01' : it.day, grain = old ? 'm' : 'd';
    var key = gday + '|' + grain + '|' + it.sn;
    var g = groups[key] || (groups[key] = { day: gday, grain: grain, sn: it.sn, s: it.s, n: 0, secSum: 0, secN: 0, c: {} });
    if (it.s && !g.s) g.s = it.s;
    g.n++; if (it.sec > 0 && it.sec < 3600) { g.secSum += it.sec; g.secN++; }
    countAns_(g.c, qs, it.a);
  });
  /* 日／月彙總 */
  var sh = aggSheet_(), last = sh.getLastRow(), from = full ? 2 : Math.max(2, last - 149), idx = {};
  if (last >= from) sh.getRange(from, 1, last - from + 1, 3).getValues().forEach(function (r, i) { idx[dayOf_(r[0]) + '|' + String(r[1]) + '|' + (Number(r[2]) || 0)] = from + i; });
  var add = [];
  Object.keys(groups).forEach(function (k) {
    var g = groups[k], at = idx[k];
    if (at) {
      var cur = sh.getRange(at, 4, 1, 5).getValues()[0], c = {};
      try { c = JSON.parse(cur[4] || '{}'); } catch (e) { c = {}; }
      for (var ck in g.c) c[ck] = (c[ck] || 0) + g.c[ck];
      sh.getRange(at, 4, 1, 5).setValues([[cur[0] || g.s, (Number(cur[1]) || 0) + g.n, (Number(cur[2]) || 0) + g.secSum, (Number(cur[3]) || 0) + g.secN, JSON.stringify(c)]]);
    } else add.push([g.day, g.grain, g.sn, g.s, g.n, g.secSum, g.secN, JSON.stringify(g.c)]);
  });
  if (add.length) {
    add.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : (a[2] - b[2]); });
    growTo_(sh, sh.getLastRow() + add.length);
    sh.getRange(sh.getLastRow() + 1, 1, add.length, AGG_HEAD.length).setValues(add);
  }
  /* agg_month 不再寫：月趨勢由 stats2_ 從 agg_daily 現算（2026-10-28.8 起） */
}
function cmAppend_(row, ans) {
  var t = row.time instanceof Date ? row.time : new Date(row.time);
  cmSheet_().appendRow([t, tz_(t, 'yyyy-MM-dd'), row.id, row.store_no, row.store, row.recipe, ans.q14 || '', ans.q12 || '', ans.q17]);
}

/* 儀表板讀取（新）：期間內各店彙總＋全公司月趨勢＋留言。回應很小，不管歷史有多少筆都一樣快 */
function stats2_(p) {
  var from = /^\d{4}-\d{2}-\d{2}$/.test(p.from || '') ? p.from : tz_(new Date(), 'yyyy-MM') + '-01';
  var to = /^\d{4}-\d{2}-\d{2}$/.test(p.to || '') ? p.to : tz_(new Date(), 'yyyy-MM-dd');
  var c = CacheService.getScriptCache(), key = 'stats2:' + (PropertiesService.getScriptProperties().getProperty('stats_ver') || '0') + ':' + from + ':' + to, hit = c.get(key);
  if (hit) return JSON.parse(hit);
  var sh = aggSheet_(), last = sh.getLastRow(), v = last >= 2 ? sh.getRange(2, 1, last - 1, AGG_HEAD.length).getValues() : [], by = {}, hist = false, mo = {}, nAll = 0;
  v.forEach(function (r) {
    var day = dayOf_(r[0]), n = Number(r[4]) || 0;
    var cc = {}; try { cc = JSON.parse(r[7] || '{}'); } catch (e) { cc = {}; }
    /* 月趨勢（全公司、全期間）：分母用「這題有回答的人」（a_qX），舊問卷只有部分題目也不會拉低 */
    var ym = day.slice(0, 7), m = mo[ym] || (mo[ym] = { ym: ym, n: 0 });
    m.n += n; nAll += n;
    for (var q in YES) { m[q] = (m[q] || 0) + (cc[q + '|' + YES[q]] || 0); m['a_' + q] = (m['a_' + q] || 0) + (cc['#' + q] || 0); }
    if (day < from || day > to) return;
    var sn = Number(r[2]) || 0, g = by[sn] || (by[sn] = { sn: sn, s: String(r[3] || ''), n: 0, secSum: 0, secN: 0, c: {} });
    g.n += n; g.secSum += Number(r[5]) || 0; g.secN += Number(r[6]) || 0;
    for (var k in cc) g.c[k] = (g.c[k] || 0) + cc[k];
    if (String(r[1]) === 'm') hist = true;
  });
  var months = Object.keys(mo).sort().map(function (k) { return mo[k]; });
  var cs = cmSheet_(), cl = cs.getLastRow(), cv = cl >= 2 ? cs.getRange(2, 1, cl - 1, CM_HEAD.length).getValues() : [], cms = [], cmN = 0;
  for (var i = cv.length - 1; i >= 0; i--) {
    var r = cv[i], day = dayOf_(r[1]); if (day < from || day > to) continue;
    cmN++; if (cms.length < 500) cms.push({ d: day, sn: Number(r[3]) || 0, s: String(r[4] || ''), r: String(r[5] || ''), q14: String(r[6] || ''), q12: String(r[7] || ''), t: String(r[8] || '') });
  }
  var o = { ok: true, ver: VER, from: from, to: to, stores: Object.keys(by).map(function (k) { return by[k]; }), hist: hist, months: months, n_all: nAll,
    comments: cms, cm_n: cmN, questions: questions_(true), db_url: db_().getUrl(), at: new Date().toISOString() };
  var s = JSON.stringify(o); if (s.length < 90000) c.put(key, s, 300);
  return o;
}

/* 整個彙總表重建（修復用：在編輯器選 rebuildAgg → 執行）。從 fact_response 全部重算，src=test 不算 */
function rebuildAgg() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = db_(), cut = PropertiesService.getScriptProperties().getProperty('ARCH_CUT') || '';   /* 舊回答已搬到封存庫時，只重建切點以後；更早的彙總列保留 */
    var keep = [];
    if (cut) { var ag = ss.getSheetByName('agg_daily'); if (ag && ag.getLastRow() >= 2) keep = ag.getRange(2, 1, ag.getLastRow() - 1, AGG_HEAD.length).getValues().filter(function (r) { return dayOf_(r[0]) < cut; }); }
    ['agg_daily', 'agg_month'].forEach(function (n) { var x = ss.getSheetByName(n); if (x) ss.deleteSheet(x); });
    var agS = aggSheet_(); aggMonthSheet_();
    if (keep.length) { growTo_(agS, keep.length + 1); agS.getRange(2, 1, keep.length, AGG_HEAD.length).setValues(keep); }
    var sh = rSheet_(), last = sh.getLastRow(), head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String), done = 0;
    var iT = head.indexOf('time'), iS = head.indexOf('src'), iN = head.indexOf('store_no'), iSt = head.indexOf('store'), iSec = head.indexOf('sec'), qc = [];
    head.forEach(function (h, i) { if (/^q\d+$/.test(h)) qc.push([h, i]); });
    for (var from = 2; from <= last; from += 4000) {
      var n = Math.min(4000, last - from + 1), v = sh.getRange(from, 1, n, head.length).getValues(), items = [];
      v.forEach(function (r) {
        if (String(r[iS]) === 'test') return;
        var it0 = itemOf_(r[iT], r[iN], r[iSt], String(r[iS]) === 'surveycake' ? 0 : r[iSec], {});
        if (cut && it0.day < cut) return;
        var a = {}; qc.forEach(function (x) { var val = r[x[1]]; if (val !== '' && val != null) a[x[0]] = String(val); });
        it0.a = a; items.push(it0);
      });
      aggMerge_(items, true); done += items.length;
    }
    PropertiesService.getScriptProperties().deleteProperty('agg_dirty');
    bumpStats_();
    return '彙總表重建完成：' + done + ' 筆' + (cut ? '（' + cut + ' 以後；更早的彙總保留 ' + keep.length + ' 列）' : '');
  } finally { lock.releaseLock(); }
}

/* 匯入窗口：在編輯器執行 importOpen → 產生 60 分鐘有效的一次性隨機碼（寫在執行記錄），匯入的每個請求要帶這個碼（b.code）；importClose 或逾時就失效。不用任何真密碼 */
function importOpen() {
  var P = PropertiesService.getScriptProperties(), code = Utilities.getUuid().replace(/-/g, '').slice(0, 20), until = Date.now() + 60 * 60 * 1000;
  P.setProperty('IMPORT_CODE', code); P.setProperty('IMPORT_UNTIL', String(until));
  var msg = '匯入窗口開到 ' + tz_(new Date(until), 'HH:mm') + '，碼：' + code; Logger.log(msg); return msg;
}
function importClose() { var P = PropertiesService.getScriptProperties(); P.deleteProperty('IMPORT_CODE'); P.deleteProperty('IMPORT_UNTIL'); Logger.log('匯入窗口已關'); return '已關'; }
function importAuth_(b) {
  var P = PropertiesService.getScriptProperties(), code = P.getProperty('IMPORT_CODE'), until = Number(P.getProperty('IMPORT_UNTIL') || 0);
  if (!code || Date.now() > until || String(b.code || '') !== code) throw fail_('匯入窗口沒開或碼不對', 'auth');
}
/* 歷史回覆匯入（臨時，2026-10-28）：由本機 Claude 分批送來；匯完會拿掉這個動作
   op='rows'（預設）：b.rows＝[{t:ISO時間, id:'sc…', sn:店號, s:店名, sec:秒, a:{qid:答案}, ua:來源說明}]，同 id 已有就跳過（可重跑）
   op='raw'：原樣封存表（舊問卷題型不同、對不上新題目的，整張原樣收）：b.sheet、b.head、b.rows（二維陣列）、b.key（去重欄名，如 Hash）
   op='count'：各表筆數；op='backup'：把相關表各複製一份 backup_…_時間 */
function import_(b) {
  importAuth_(b);
  var op = String(b.op || 'rows');
  if (op === 'count') return counts_();
  if (op === 'backup') return backup_();
  if (op === 'raw') return importRaw_(b);
  var rows = Array.isArray(b.rows) ? b.rows : [];
  if (!rows.length || rows.length > 4000) throw fail_('rows 要 1～4000 筆', 'input');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = rSheet_(), last = sh.getLastRow(), head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String), have = {};
    var iId = head.indexOf('id'); if (last >= 2) sh.getRange(2, iId + 1, last - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = 1; });
    var qs = questions_(true), vals = [], items = [], skip = 0;
    rows.forEach(function (r) {
      var id = clean_(r.id, 64); if (!id || have[id]) { skip++; return; } have[id] = 1;
      var a = {}; qs.forEach(function (q) { var v = r.a && r.a[q.qid]; if (v != null && String(v) !== '') a[q.qid] = String(v); });
      var t = new Date(r.t), sec = Math.min(Math.max(Number(r.sec) || 0, 0), 86400);
      if (isNaN(t)) { skip++; return; }
      var row = { time: t, id: id, store_no: Number(r.sn) || 0, store: clean_(r.s, 60), recipe: '', src: 'surveycake', lang: 'zh', sec: sec };
      vals.push(head.map(function (h) { return h === 'ua' ? clean_(r.ua, 160) : (h in row ? row[h] : (h in a ? a[h] : '')); }));
      items.push(itemOf_(t, row.store_no, row.store, 0, a));   /* 彙總的「平均填答時間」不算舊問卷：舊的填答秒數是食譜頁開著到送出（約 2～3 小時），原值仍存在 fact_response.sec */
    });
    if (vals.length) { growTo_(sh, last + vals.length); sh.getRange(last + 1, 1, vals.length, head.length).setValues(vals); }
    aggMerge_(items, true);
    SpreadsheetApp.flush();
    return { added: vals.length, skipped: skip, total: sh.getLastRow() - 1 };
  } finally { lock.releaseLock(); }
}
function counts_() {
  var ss = db_(), o = {};
  ss.getSheets().forEach(function (sh) { var n = sh.getName(); if (/^(fact_|agg_|raw_|dim_|backup_)/.test(n)) o[n] = Math.max(sh.getLastRow() - 1, 0); });
  var aid = PropertiesService.getScriptProperties().getProperty('ARCHIVE_SS');
  if (aid) { try { SpreadsheetApp.openById(aid).getSheets().forEach(function (sh) { o['封存庫/' + sh.getName()] = Math.max(sh.getLastRow() - 1, 0); }); } catch (e) { o['封存庫'] = '打不開'; } }
  return { counts: o };
}
function backup_() {
  var ss = db_(), ts = tz_(new Date(), 'yyyyMMdd-HHmm'), made = {};
  ['fact_response', 'agg_daily', 'agg_month', 'fact_comment', 'dim_question'].forEach(function (n) {
    var sh = ss.getSheetByName(n); if (!sh) return;
    var c = sh.copyTo(ss); c.setName('backup_' + n + '_' + ts); made[c.getName()] = Math.max(c.getLastRow() - 1, 0);
  });
  return { backup: made, counts: counts_().counts };
}
function importRaw_(b) {
  var name = String(b.sheet || ''); if (!/^raw_[A-Za-z0-9_一-鿿]{1,40}$/.test(name)) throw fail_('封存表名稱要 raw_ 開頭', 'input');
  var rows = Array.isArray(b.rows) ? b.rows : [], head = Array.isArray(b.head) ? b.head.map(String) : [];
  if (!rows.length || rows.length > 4000 || !head.length) throw fail_('rows 要 1～4000 筆、要附 head', 'input');
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var ss = db_(), sh = ss.getSheetByName(name), nc = head.length;
    if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, nc).setValues([head]); sh.setFrozenRows(1); }
    var h0 = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    if (h0.join('¦') !== head.join('¦')) throw fail_('欄位和已存在的封存表不同', 'input');
    var last = sh.getLastRow(), have = {}, iK = b.key ? head.indexOf(String(b.key)) : -1;
    if (iK >= 0 && last >= 2) sh.getRange(2, iK + 1, last - 1, 1).getValues().forEach(function (r) { have[String(r[0])] = 1; });
    var vals = [], skip = 0;
    rows.forEach(function (r) {
      if (iK >= 0) { var k = String(r[iK] == null ? '' : r[iK]); if (k && have[k]) { skip++; return; } have[k] = 1; }
      var row = []; for (var j = 0; j < nc; j++) row.push(r[j] == null ? '' : String(r[j]).slice(0, 45000)); vals.push(row);
    });
    if (vals.length) {
      growTo_(sh, last + vals.length);
      var rg = sh.getRange(last + 1, 1, vals.length, nc); rg.setNumberFormat('@'); rg.setValues(vals);   /* 一律當純文字，不讓 Sheets 把時間／數字自動轉型 */
    }
    SpreadsheetApp.flush();
    return { sheet: name, added: vals.length, skipped: skip, total: sh.getLastRow() - 1 };
  } finally { lock.releaseLock(); }
}
/* 一次性（2026-10-28）：把已匯入歷史的彙總列秒數歸零（只動 2026-10-28 以前的列；之後現場回答的秒數是真的，不碰） */
function aggZeroSec() {
  var sh = aggSheet_(), last = sh.getLastRow(); if (last < 2) return '彙總表是空的';
  var days = sh.getRange(2, 1, last - 1, 1).getValues(), fg = sh.getRange(2, 6, last - 1, 2).getValues(), n = 0;
  for (var i = 0; i < days.length; i++) if (dayOf_(days[i][0]) < '2026-10-28' && (Number(fg[i][0]) || Number(fg[i][1]))) { fg[i] = [0, 0]; n++; }
  sh.getRange(2, 6, last - 1, 2).setValues(fg);
  bumpStats_();
  return '已把 ' + n + ' 列的秒數歸零';
}
/* ---------------- 分檔（2026-10-28）：讓主資料庫不會撐到 1,000 萬格上限 ----------------
   封存庫＝另一個試算表「顧客問卷_封存庫」（指令碼屬性 ARCHIVE_SS）。都在編輯器執行；時間不夠會說「再執行一次」，重跑安全。 */
function archiveSS_() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('ARCHIVE_SS'), ar = null;
  if (id) { try { ar = SpreadsheetApp.openById(id); } catch (_) { ar = null; } }
  if (!ar) { ar = SpreadsheetApp.create('顧客問卷_封存庫'); ar.setSpreadsheetTimeZone('Asia/Taipei'); P.setProperty('ARCHIVE_SS', ar.getId()); }
  return ar;
}
function dropDefault_(ar) { var d = ar.getSheetByName('工作表1') || ar.getSheetByName('Sheet1'); if (d && ar.getSheets().length > 1) ar.deleteSheet(d); }
/* 第 1 步：raw_*（舊問卷原樣）與 backup_*（匯入前備份）整張搬到封存庫；先複製、列數欄數一致才刪主檔那張 */
function archiveMoveSheets() {
  var ss = db_(), ar = archiveSS_(), t0 = Date.now(), out = [];
  var list = ss.getSheets().filter(function (sh) { return /^(raw_|backup_)/.test(sh.getName()); });
  for (var i = 0; i < list.length; i++) {
    if (Date.now() - t0 > 270000) { out.push('時間快到了，還有 ' + (list.length - i) + ' 張沒搬，請再執行一次'); break; }
    var sh = list[i], n = sh.getName(), rows = sh.getLastRow(), cols = sh.getLastColumn(), ex = ar.getSheetByName(n);
    if (!ex) { ex = sh.copyTo(ar); ex.setName(n); dropDefault_(ar); }
    if (ex.getLastRow() !== rows || ex.getLastColumn() !== cols) { out.push(n + '：封存庫那張列數不符（' + ex.getLastRow() + ' vs ' + rows + '），主檔沒刪'); continue; }
    ss.deleteSheet(sh); out.push(n + '：已搬 ' + (rows - 1) + ' 列');
  }
  var msg = out.join('\n') + '\n封存庫：' + ar.getUrl(); Logger.log(msg); return msg;
}
/* 第 2 步：fact_response 裡 cut（預設 2025-01-01）以前的原始回答搬到封存庫 fact_response_2025前；彙總表不動（儀表板照常） */
function archiveOldResponses() {
  var cut = '2025-01-01', P = PropertiesService.getScriptProperties(), stage = P.getProperty('ARCH_STAGE') || '';
  var ss = db_(), ar = archiveSS_(), sh = rSheet_(), name = 'fact_response_2025前', msg = [];
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var before = Number(P.getProperty('ARCH_BEFORE') || 0) || (sh.getLastRow() - 1);
    function oldRuns(target, wantOld) {   /* 回傳要刪的連續列段 [起列, 列數]；wantOld=true 找舊的、false 找新的 */
      var last = target.getLastRow(); if (last < 2) return [];
      var t = target.getRange(2, 1, last - 1, 1).getValues(), runs = [], st = -1;
      for (var i = 0; i <= t.length; i++) {
        var hit = i < t.length && ((dayOf_(t[i][0]) < cut) === wantOld);
        if (hit && st < 0) st = i; if (!hit && st >= 0) { runs.push([st + 2, i - st]); st = -1; }
      }
      return runs;
    }
    if (!stage) {
      P.setProperty('ARCH_BEFORE', String(before));
      var cp = ar.getSheetByName(name); if (cp) ar.deleteSheet(cp);
      cp = sh.copyTo(ar); cp.setName(name); dropDefault_(ar);
      if (cp.getLastRow() !== sh.getLastRow()) throw fail_('複製到封存庫的列數不符，停止');
      P.setProperty('ARCH_STAGE', 'copied'); stage = 'copied'; msg.push('已複製整張到封存庫（' + (cp.getLastRow() - 1) + ' 列）');
    }
    if (stage === 'copied') {
      var cp2 = ar.getSheetByName(name), r1 = oldRuns(cp2, false);
      for (var k = r1.length - 1; k >= 0; k--) cp2.deleteRows(r1[k][0], r1[k][1]);
      P.setProperty('ARCH_STAGE', 'trimmed'); stage = 'trimmed'; msg.push('封存庫只留 ' + cut + ' 以前：' + (cp2.getLastRow() - 1) + ' 列');
    }
    if (stage === 'trimmed') {
      var nOld = ar.getSheetByName(name).getLastRow() - 1, r2 = oldRuns(sh, true);
      for (var j = r2.length - 1; j >= 0; j--) sh.deleteRows(r2[j][0], r2[j][1]);
      var nNew = sh.getLastRow() - 1;
      if (nOld + nNew !== before) throw fail_('核對不符：封存 ' + nOld + '＋主檔 ' + nNew + ' ≠ 原本 ' + before + '（主檔可從封存庫補回）');
      P.setProperty('ARCH_STAGE', 'done'); P.setProperty('ARCH_CUT', cut); CacheService.getScriptCache().remove('rhead');
      msg.push('主檔刪掉 ' + cut + ' 以前的列：原本 ' + before + ' ＝ 封存 ' + nOld + ' ＋ 主檔 ' + nNew + ' ✔');
    } else if (stage === 'done') msg.push('之前已經做完（切點 ' + P.getProperty('ARCH_CUT') + '）');
  } finally { lock.releaseLock(); }
  var out = msg.join('\n') + '\n封存庫：' + ar.getUrl(); Logger.log(out); return out;
}

/* 還原匯入（編輯器選 rollbackImport → 執行）：把 src=surveycake 的列從 fact_response 拿掉，彙總表重建；raw_ 封存表另外手動刪 */
function rollbackImport() {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var sh = rSheet_(), last = sh.getLastRow(), nc = sh.getLastColumn(), v = last >= 2 ? sh.getRange(2, 1, last - 1, nc).getValues() : [];
    var head = sh.getRange(1, 1, 1, nc).getValues()[0].map(String), iS = head.indexOf('src');
    var keep = v.filter(function (r) { return String(r[iS]) !== 'surveycake'; });
    if (last >= 2) sh.getRange(2, 1, last - 1, nc).clearContent();
    if (keep.length) sh.getRange(2, 1, keep.length, nc).setValues(keep);
    var msg = '已拿掉匯入的 ' + (v.length - keep.length) + ' 筆，保留 ' + keep.length + ' 筆';
    lock.releaseLock(); lock = null;
    return msg + '；' + rebuildAgg();
  } finally { if (lock) lock.releaseLock(); }
}

/* ---------------- 儀表板「✏️ 編輯題目」（2026-10-28.9） ----------------
   讀：GET qlist → 全部題目（含停用）＋版本紀錄；寫：POST qsave／qrestore，要帶「顧客問卷」那組密碼（權限中心 dim_auth survey 列），
   後端直接讀 dim_auth 比對（明文存放，見 auth-center skill）；錯 10 次鎖 10 分鐘。每次存檔前把舊題目整份記進 dim_question_log。 */
var KPI_LOCK = { q6: '是', q7: '是', q8: '是', q9: '是', q12: '是', q14: '會' };   /* 儀表板指標題：題型固定單選、這個答案一定要在 */
function authPwd_(pwd) {
  var c = CacheService.getScriptCache(), bad = Number(c.get('pwbad') || 0);
  if (bad >= 10) throw fail_('密碼錯太多次，請 10 分鐘後再試', 'locked');
  var v = SpreadsheetApp.openById(AUTH_SS).getSheetByName('dim_auth').getDataRange().getValues(), ok = false;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === DASH_ID) {
    var real = String(v[i][2] == null ? '' : v[i][2]), on = String(v[i][4] || 'Y').toUpperCase() !== 'N';
    ok = on && real !== '' && String(pwd == null ? '' : pwd) === real;
  }
  if (!ok) { c.put('pwbad', String(bad + 1), 600); throw fail_('密碼不對（要用「顧客問卷」這組密碼）', 'auth'); }
}
function qLogSheet_() { return sheet_('dim_question_log', ['saved_at', 'by', 'n', 'json'], [1]); }
function qVer_() { return PropertiesService.getScriptProperties().getProperty('Q_VER') || ''; }
function qlist_() {
  var lg = qLogSheet_(), last = lg.getLastRow(), vers = [];
  if (last >= 2) lg.getRange(Math.max(2, last - 29), 1, Math.min(30, last - 1), 3).getValues().forEach(function (r) { vers.push({ at: String(r[0]), by: String(r[1]), n: Number(r[2]) || 0 }); });
  vers.reverse();
  return { ok: true, ver: VER, qver: qVer_(), next_qid: 'q' + (qMax_() + 1), questions: questionsRead_(true), stores: STORES.map(function (s) { return s[1]; }), kpi: KPI_LOCK, versions: vers };
}
/* 檢查並整理成要寫進表的列；有問題就整批擋下，訊息用白話 */
function qMax_() {   /* 用過的最大題號（目前題目＋所有舊版本），新題一律往上加，不重用 */
  var m = Number(PropertiesService.getScriptProperties().getProperty('Q_MAX') || 0);
  questionsRead_(true).forEach(function (q) { m = Math.max(m, Number(q.qid.slice(1)) || 0); });
  var lg = qLogSheet_(), last = lg.getLastRow();
  if (last >= 2) lg.getRange(2, 4, last - 1, 1).getValues().forEach(function (r) { try { JSON.parse(r[0] || '[]').forEach(function (x) { m = Math.max(m, Number(String(x[0]).slice(1)) || 0); }); } catch (e) { } });
  return m;
}
function qValidate_(list, isRestore) {
  if (!Array.isArray(list) || !list.length) throw fail_('沒有題目', 'input');
  var errs = [], seen = {}, pos = {}, out = [], curIds = {}, qmax = isRestore ? 0 : qMax_();
  if (!isRestore) questionsRead_(true).forEach(function (q) { curIds[q.qid] = 1; });
  list.forEach(function (q, i) {
    var qid = String(q.qid || ''), n = '「' + (clean_(q.title, 200).slice(0, 18) || '（沒寫題目）') + '」（' + (qid || '?') + '）';
    if (!/^q\d{1,3}$/.test(qid)) { errs.push(n + '：題號不對'); return; }
    if (seen[qid]) { errs.push(n + '：題號 ' + qid + ' 重複'); return; }
    seen[qid] = 1; pos[qid] = i;
    if (!isRestore && !curIds[qid] && Number(qid.slice(1)) <= qmax) errs.push(n + '：題號 ' + qid + ' 以前用過，新題要用 q' + (qmax + 1) + ' 以後（避免新舊答案混在一起）');
    var type = ['single', 'multi', 'text'].indexOf(q.type) >= 0 ? q.type : '';
    if (!type) errs.push(n + '：題型不對');
    var title = clean_(q.title, 200); if (!title) errs.push(n + '：題目文字是空的');
    var opts = (Array.isArray(q.options) ? q.options : []).map(function (o) { return clean_(o, 120); }).filter(Boolean);
    if (qid === 'q1') opts = STORES.map(function (s) { return s[1]; });   /* 店別選項由系統管，不吃畫面的改動 */
    if (type !== 'text') {
      if (opts.length < 2) errs.push(n + '：選項至少要 2 個');
      opts.forEach(function (o) { if (o.indexOf('|') >= 0) errs.push(n + '：選項不能有「|」符號（' + o + '）'); });
      var dup = opts.filter(function (o, k) { return opts.indexOf(o) !== k; }); if (dup.length) errs.push(n + '：選項重複（' + dup.join('、') + '）');
    } else opts = [];
    if (KPI_LOCK[qid]) {
      if (type !== 'single') errs.push(n + '是儀表板指標題，題型要維持單選');
      if (opts.indexOf(KPI_LOCK[qid]) < 0) errs.push(n + '是儀表板指標題，選項一定要有「' + KPI_LOCK[qid] + '」');
    }
    var cond = q.cond && q.cond.qid ? { qid: String(q.cond.qid), vals: (q.cond.vals || []).map(String).filter(Boolean) } : null;
    out.push({ qid: qid, type: type, title: title, opts: opts, req: q.required ? 'Y' : 'N', en: q.enabled === false ? 'N' : 'Y', note: clean_(q.note, 200), cond: cond, i: i });
  });
  if (!seen.q1) errs.push('店別題（q1）不能拿掉');
  out.forEach(function (q) {
    if (!q.cond) return;
    var p = out.filter(function (x) { return x.qid === q.cond.qid; })[0], n = '「' + q.title.slice(0, 18) + '」的跳題條件';
    if (!p) { errs.push(n + '：參考的題目不存在'); return; }
    if (p.i >= q.i) errs.push(n + '：只能參考排在它前面的題目');
    if (p.type === 'text') errs.push(n + '：不能參考文字題');
    if (!q.cond.vals.length) errs.push(n + '：至少勾一個答案');
    q.cond.vals.forEach(function (v) { if (p.opts.indexOf(v) < 0) errs.push(n + '：「' + v + '」不是「' + p.title.slice(0, 18) + '」的選項'); });
    if (p.en === 'N' && q.en === 'Y') errs.push(n + '：參考的題目已停用，這題永遠不會出現');
  });
  if (errs.length) throw fail_('還有 ' + errs.length + ' 個地方要改：\n' + errs.slice(0, 12).join('\n'), 'input');
  return out.map(function (q, k) { return [q.qid, k + 1, q.type, q.title, q.opts.join('|'), q.req, q.en, q.note, q.cond ? q.cond.qid + '=' + q.cond.vals.join('|') : '']; });
}
function qWrite_(rows, by) {
  var sh = qSheet_(), last = sh.getLastRow(), cur = last >= 2 ? sh.getRange(2, 1, last - 1, Q_HEAD.length).getValues() : [];
  var lg = qLogSheet_(); lg.appendRow([tz_(new Date(), 'yyyy-MM-dd HH:mm:ss'), by, cur.length, JSON.stringify(cur)]);   /* 存檔前留舊版 */
  if (lg.getLastRow() > 61) lg.deleteRows(2, lg.getLastRow() - 61);   /* 只留最近 60 版 */
  sh.getRange(1, 1, 1, Q_HEAD.length).setValues([Q_HEAD]);
  if (last >= 2) sh.getRange(2, 1, last - 1, Math.max(sh.getLastColumn(), Q_HEAD.length)).clearContent();
  growTo_(sh, rows.length + 1);
  sh.getRange(2, 1, rows.length, Q_HEAD.length).setValues(rows);
  SpreadsheetApp.flush();
  var c = CacheService.getScriptCache(); c.removeAll(['form:v2', 'qs:0', 'qs:1']);
  var mx = Number(PropertiesService.getScriptProperties().getProperty('Q_MAX') || 0); rows.forEach(function (r) { mx = Math.max(mx, Number(String(r[0]).slice(1)) || 0); });
  var v = tz_(new Date(), 'yyyyMMddHHmmss') + Math.floor(Math.random() * 1000); PropertiesService.getScriptProperties().setProperties({ Q_VER: v, Q_MAX: String(mx) }); bumpStats_();
  return v;
}
function qsave_(b) {
  authPwd_(b.pwd);
  var rows = qValidate_(b.questions);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (String(b.base || '') !== qVer_()) throw fail_('題目剛剛被別人（或另一個分頁）改過了，請按「重新讀取」再改', 'conflict');
    var v = qWrite_(rows, '儀表板');
    return { qver: v, n: rows.length };
  } finally { lock.releaseLock(); }
}
function qrestore_(b) {
  authPwd_(b.pwd);
  var lg = qLogSheet_(), last = lg.getLastRow(), at = String(b.at || ''), row = null;
  if (last >= 2) lg.getRange(2, 1, last - 1, 4).getValues().forEach(function (r) { if (String(r[0]) === at) row = r; });
  if (!row) throw fail_('找不到這個版本', 'input');
  var old = JSON.parse(row[3] || '[]').sort(function (a, b2) { return (Number(a[1]) || 0) - (Number(b2[1]) || 0); });
  var list = old.map(function (r) { var cd = parseCond_(r[8]); return { qid: String(r[0]), type: String(r[2]), title: String(r[3]), options: String(r[4] || '').split('|').filter(Boolean), required: String(r[5]).toUpperCase() === 'Y', enabled: String(r[6]).toUpperCase() !== 'N', note: String(r[7] || ''), cond: cd }; });
  var rows = qValidate_(list, true);
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    if (String(b.base || '') !== qVer_()) throw fail_('題目剛剛被別人改過了，請按「重新讀取」再試', 'conflict');
    return { qver: qWrite_(rows, '復原到 ' + at), n: rows.length };
  } finally { lock.releaseLock(); }
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
