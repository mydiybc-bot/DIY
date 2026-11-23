/* =====================================================================
   🎨 「測出你的蛋糕色」個性測驗紀錄（2026-11-23 加，跟顧客問卷同一個 GAS／資料庫）
   前端：static/color-quiz.html（客人用手機掃桌上 QR，?s=店號）
   方案：資料交換/輸出/成品/個性裝飾蛋糕測驗_20261123/方案_v1.md

   動作
     POST {action:'quiz_log', id, kind:'result'|'switch', store_no, main, sub, ans, lang, src, ua} → {ok, id, dup}
          result＝測完的結果（main 主色、sub 副色、ans 6 題答案的顏色字串，例 PYBGVP）
          switch＝客人在結果頁按「換成別的顏色」（main＝換成的顏色）
     GET  ?action=quiz_stats[&from=yyyy-MM-dd&to=yyyy-MM-dd][&callback=] → 各店×各色次數（不含 src=test）
   資料：分頁 quiz_color（第一筆時自動建）；不收任何個資（沒有姓名、電話、Email）。
   防濫用：每天最多 QUIZ_CAP 筆；同一個 id＋kind＋顏色 6 小時只收一次。
   ===================================================================== */
var QUIZ_CAP = 3000;
var QUIZ_COLORS = ['P', 'Y', 'B', 'G', 'V'];   /* 粉 黃 藍 綠 紫 */
var QUIZ_HEAD = ['time', 'day', 'id', 'kind', 'store_no', 'store', 'main', 'sub', 'ans', 'lang', 'src', 'ua'];

function quizSheet_() {
  var ss = db_(), sh = ss.getSheetByName('quiz_color');
  if (!sh) {
    sh = ss.insertSheet('quiz_color');
    sh.getRange(1, 1, 1, QUIZ_HEAD.length).setValues([QUIZ_HEAD]); sh.setFrozenRows(1);
  }
  return sh;
}

function quizLog_(b) {
  var id = clean_(b.id, 64), kind = b.kind === 'switch' ? 'switch' : 'result';
  if (!/^[A-Za-z0-9_-]{8,64}$/.test(id)) throw fail_('缺少識別碼', 'input');
  var main = String(b.main || ''), sub = String(b.sub || ''), ans = String(b.ans || '');
  if (QUIZ_COLORS.indexOf(main) < 0) throw fail_('顏色不對', 'input');
  if (sub && QUIZ_COLORS.indexOf(sub) < 0) throw fail_('副色不對', 'input');
  if (kind === 'result' && !/^[PYBGV]{6}$/.test(ans)) throw fail_('答案不完整', 'input');
  if (kind === 'switch') ans = '';
  var no = Number(b.store_no) || '', store = '';
  STORES.forEach(function (s) { if (s[0] === no) store = s[1]; });
  if (!store) no = '';

  var c = CacheService.getScriptCache(), dk = 'qz:' + id + ':' + kind + ':' + main;
  if (c.get(dk)) return { id: id, dup: true };
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var P = PropertiesService.getScriptProperties(), ck = 'qzcnt_' + today_(), n = Number(P.getProperty(ck) || 0);
    if (n >= QUIZ_CAP) throw fail_('今天的測驗次數已達上限', 'cap');
    if (n === 0) P.getKeys().forEach(function (k) { if (/^qzcnt_/.test(k) && k !== ck) P.deleteProperty(k); });
    P.setProperty(ck, String(n + 1));
    var now = new Date();
    quizSheet_().appendRow([now, tz_(now, 'yyyy-MM-dd'), id, kind, no, store, main, sub, ans,
      clean_(b.lang, 8), clean_(b.src, 20) || 'qr', clean_(b.ua, 160)]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  c.put(dk, '1', 21600);
  c.remove('qzstats');
  return { id: id, dup: false };
}

function quizStats_(p) {
  var from = /^\d{4}-\d{2}-\d{2}$/.test(p.from || '') ? p.from : '2000-01-01';
  var to = /^\d{4}-\d{2}-\d{2}$/.test(p.to || '') ? p.to : '2999-12-31';
  var c = CacheService.getScriptCache(), ck = 'qzstats', hit = c.get(ck), all;
  if (hit) all = JSON.parse(hit);
  else {
    var sh = quizSheet_(), last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, QUIZ_HEAD.length).getDisplayValues() : [];
    all = vals.map(function (r) { return [r[1], r[3], r[4], r[6], r[7], r[10]]; });   /* day, kind, store_no, main, sub, src */
    var s = JSON.stringify(all); if (s.length < 90000) c.put(ck, s, 180);
  }
  var by = {}, tot = { result: {}, sub: {}, switchTo: {} }, n = 0, sw = 0;
  all.forEach(function (r) {
    if (r[0] < from || r[0] > to || r[5] === 'test') return;
    var k = r[2] || '0';
    by[k] = by[k] || { result: {}, switchTo: {} };
    if (r[1] === 'result') {
      n++; by[k].result[r[3]] = (by[k].result[r[3]] || 0) + 1; tot.result[r[3]] = (tot.result[r[3]] || 0) + 1;
      if (r[4]) tot.sub[r[4]] = (tot.sub[r[4]] || 0) + 1;
    } else { sw++; by[k].switchTo[r[3]] = (by[k].switchTo[r[3]] || 0) + 1; tot.switchTo[r[3]] = (tot.switchTo[r[3]] || 0) + 1; }
  });
  return { ok: true, ver: VER, from: from, to: to, n: n, switches: sw, total: tot, byStore: by, stores: STORES };
}
