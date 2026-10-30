/* =====================================================================
   🗳 門市內部調查（2026-10-30 加；跟顧客問卷共用同一個 GAS 與資料庫試算表，不必另外授權）
   用途：總部要各店回報某件事（例：擠花袋破損狀況），一份調查＝POLLS 裡一筆設定＋資料庫一張 poll_<代號> 分頁。
   前端：static/poll.html?p=<代號>（店員填）；static/poll.html?p=<代號>&view=result（看結果：各店最新一筆＋誰還沒填）

   動作
     GET  ?action=poll_form&p=<代號>[&callback=]     → {ok, poll:{id,title,intro,fields}, stores}
     GET  ?action=poll_result&p=<代號>[&callback=]   → {ok, poll, stores, rows:[…全部回覆，舊→新…]}
     POST {action:'poll_submit', p, id, a:{欄位:答案}, ua} → {ok, id, dup}
   規則：同一家店可以重填，全部保留；結果頁以每店「最新一筆」為準。每份調查每天最多 POLL_CAP 筆。
   新增一份調查：在 POLLS 加一筆 → clasp push → update-deployment；不用動 poll.html。
   ===================================================================== */
var POLL_CAP = 300;

var POLLS = {
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

function pollGet_(p) {
  var cfg = POLLS[String(p || '')]; if (!cfg) throw fail_('找不到這份調查', 'poll');
  var storeNames = STORES.map(function (s) { return s[1]; });
  return { id: String(p), title: cfg.title, intro: cfg.intro || '',
    fields: cfg.fields.map(function (f) { var o = {}; for (var x in f) o[x] = f[x]; if (o.options === 'STORES') o.options = storeNames; return o; }) };
}
function pollSheet_(pid, keys) {
  var ss = db_(), name = 'poll_' + pid, sh = ss.getSheetByName(name);
  var head = ['time', 'id', 'store_no', 'store'].concat(keys.filter(function (k) { return k !== 'store'; }), ['ua']);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]); sh.setFrozenRows(1); }
  return sh;
}
function pollForm_(p) { return { ok: true, ver: VER, poll: pollGet_(p.p), stores: STORES }; }

function pollSubmit_(b) {
  var poll = pollGet_(b.p), id = clean_(b.id, 64);
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
      var n = Number(s);
      if (!isFinite(n) || n < (f.min || 0) || n > (f.max || 1e9) || Math.floor(n) !== n) throw fail_('「' + f.title + '」請填 0 以上的整數', 'input');
      ans[f.k] = n;
    } else if (f.type === 'text') {
      v = clean_(v, f.max || 200);
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
    var need = keys.filter(function (k) { return head.indexOf(k) < 0; });   /* 調查之後加了欄位 → 補在 ua 前面 */
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
  return { id: id, dup: false };
}

function pollResult_(p) {
  var poll = pollGet_(p.p), sh = db_().getSheetByName('poll_' + poll.id), rows = [];
  if (sh && sh.getLastRow() > 1) {
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues(), head = v[0].map(String);
    for (var i = 1; i < v.length; i++) {
      var o = {};
      head.forEach(function (h, j) { if (h === 'ua') return; var x = v[i][j]; o[h] = x instanceof Date ? tz_(x, 'yyyy-MM-dd HH:mm') : x; });
      rows.push(o);
    }
  }
  return { ok: true, ver: VER, poll: poll, stores: STORES, rows: rows, at: tz_(new Date(), 'yyyy-MM-dd HH:mm') };
}
