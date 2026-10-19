/* ================= 🌐 食譜翻譯：讀後台食譜原文＋三地食材名稱對照（2026-10-19） =================
   經營者 10/19：「把任何食譜自由轉換成澳門中文、菲律賓英文、越南文的食譜；三地對同一個食譜有不同用語、名稱（特別是食材、耗材、包裝盒、器具和模具）」。
   第一段只產出翻譯稿給人看，**不寫進後台**。翻譯本身在前端呼叫訓練站的 Claude 代理（/api/recipe/translate，10/19 起用身分密碼驗證，不用食譜小幫手密碼）；
   這裡只負責兩件只讀的事：
   ① trRecipe：讀一支食譜的「匯出食譜」頁（/steps/recipeprint/{id}，後台本來就有、只讀）→ 回傳內容 HTML，前端拆成欄位與步驟。
   ② trDict：讀後台「食材管理」三張清單 → 每個品項 10 種語言的名稱＋各語言單位（前端把食材工具清單直接換成當地名稱，不靠 AI 猜）：
      /IngredSources（編號、繁中名、繁中單位）、/IngredSources/NameList（10 語言名稱，與前一張同順序）、/ingredItems（各語言列的單位）。
   ⚠ 只讀：不改、不刪後台任何資料，不寫試算表。要身分密碼（任何身分），完整食譜內容不給只知道網址的人。
   ⚠ trDict 暫存 6 小時（後台三張清單約 7MB，讀一次 10 多秒）；帶 fresh:1 重讀。 */
var TR_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
var TR_DICT_KEY = 'trdict:v1', TR_DICT_TTL = 21600;

function trGet_(path) {
  var sess = rbSession_(), r = rbAuthedGet_(sess, path);
  if (r.code !== 200) throw rbErr_('食譜後台回應異常（HTTP ' + r.code + '）', 'rbnet');
  rbSaveSession_(sess);
  return r.text;
}
/* 表格每一列 → 每格的 HTML（不含 <td> 本身） */
function trRows_(html) {
  var out = [], tb = html.indexOf('<table'), re = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi, m;
  if (tb < 0) return out;
  re.lastIndex = tb;
  while ((m = re.exec(html))) {
    var cells = [], cr = /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi, c;
    while ((c = cr.exec(m[1]))) cells.push(c[1]);
    out.push(cells);
  }
  return out;
}

function trRecipe_(p) {
  var id = String(p.id || '').trim().toLowerCase();
  if (!TR_ID.test(id)) throw fail_('食譜編號不對（要後台網址最後那一串）', 'bad');
  var html = trGet_('/steps/recipeprint/' + id);
  var i = html.indexOf('id="exportContent"');
  if (i < 0) throw rbErr_('後台沒有這支食譜，或「匯出食譜」頁改版了', 'rbpage');
  var s = html.indexOf('>', i) + 1, e = html.lastIndexOf('</div>');
  var body = html.slice(s, e > s ? e : html.length);
  return { id: id, msg: id, data: { id: id, html: body, bytes: body.length } };
}

function trDict_(p) {
  var cache = CacheService.getScriptCache();
  if (!p.fresh) { var hit = cacheGetBig_(cache, TR_DICT_KEY); if (hit) return { data: JSON.parse(hit) }; }
  var t0 = Date.now();
  /* ① 品項清單：編號、繁中名、單位（後台排序＝品項 ID，與名稱對照表同順序） */
  var src = trRows_(trGet_('/IngredSources')).slice(1).map(function (c) {
    return { code: rbText_(c[0]), name: rbText_(c[2]), unit: rbText_(c[3]), cat: rbText_(c[1]) };
  }).filter(function (o) { return o.code || o.name; });
  /* ② 名稱對照表：第一列是 10 個語言 */
  var nl = trRows_(trGet_('/IngredSources/NameList'));
  var langs = (nl[0] || []).map(rbText_), names = nl.slice(1).map(function (c) { return c.map(rbText_); });
  /* ③ 各語言列的單位：語言＋名稱 → 單位（同名不同單位全留） */
  var units = {}, ir = trRows_(trGet_('/ingredItems')), ih = (ir[0] || []).map(rbText_);
  var col = function (re) { for (var j = 0; j < ih.length; j++) if (re.test(ih[j])) return j; return -1; };   /* 表頭是「Language<br>語言」這種中英兩行 */
  var iL = col(/Language/i), iN = col(/^Name/i), iU = col(/^Unit/i);
  if (iL < 0 || iN < 0 || iU < 0) throw rbErr_('後台「食材總表」欄位改版了，讀不到單位（表頭：' + ih.join('／') + '）', 'rbpage');
  ir.slice(1).forEach(function (c) {
    var L = rbText_(c[iL]), n = rbText_(c[iN]), u = rbText_(c[iU]);
    if (!L || !n || !u) return;
    var m = units[L] = units[L] || {}, k = rbNorm_(n);
    if (!m[k]) m[k] = [];
    if (m[k].indexOf(u) < 0) m[k].push(u);
  });
  /* 兩張表同順序：逐列核對繁中名，對不上的照名稱找 */
  var byName = {};
  src.forEach(function (o, i) { var k = rbNorm_(o.name); (byName[k] = byName[k] || []).push(i); });
  var used = {}, rows = names.map(function (n, i) {
    var o = src[i], k = rbNorm_(n[0]);
    if (!o || rbNorm_(o.name) !== k) { var c = (byName[k] || []).filter(function (j) { return !used[j]; }); o = c.length ? src[c[0]] : null; if (c.length) used[c[0]] = 1; }
    else used[i] = 1;
    return [o ? o.code : '', o ? o.unit : '', o ? o.cat : ''].concat(n);
  });
  var d = { at: now_(), ms: Date.now() - t0, langs: langs, head: ['code', 'unit', 'cat'].concat(langs), rows: rows, units: units, items: ir.length - 1 };
  try { cachePutBig_(cache, TR_DICT_KEY, JSON.stringify(d), TR_DICT_TTL); } catch (ce) { /* 放不進暫存只是下次重讀 */ }
  return { data: d };
}

ACTIONS.trRecipe = function (p, a) { return trRecipe_(p, a); };
ACTIONS.trDict = function (p, a) { return trDict_(p, a); };
READS.trRecipe = 1;   /* 只讀後台，不排鎖 */
READS.trDict = 1;

/* 編輯器手動驗證（只讀）：讀字典＋一支食譜，看大小與欄位 */
function trTestNow() {
  var d = trDict_({ fresh: 1 }).data;
  var r = trRecipe_({ id: '481cd8fc-5fee-49dd-a404-25b96d87a36a' }).data;
  Logger.log(JSON.stringify({ ms: d.ms, langs: d.langs, rows: d.rows.length, noCode: d.rows.filter(function (x) { return !x[0]; }).length,
    unitLangs: Object.keys(d.units).map(function (k) { return k + ':' + Object.keys(d.units[k]).length; }), items: d.items,
    dictKB: Math.round(JSON.stringify(d).length / 1024), recipeBytes: r.bytes, sample: d.rows.slice(0, 2) }));
}

/* ================= 🌐 三地確認用語（2026-10-19 第二批） =================
   經營者 10/19：全新的用語、食材、耗材、包裝盒、模具器具（以前沒用過的）→ AI 暫譯後，由主廚／代理商確認；確認過的要記住，下次照用。
   分頁 tr_confirmed（申請單試算表，第一次寫入時自動建）：lang（mo／ph／vn／me 澳門英文／en 英文）、kind（item＝品項名稱、term＝用語）、zh、tr、unit（品項的當地單位）、
   note、src（從哪支食譜／哪個功能確認的）、by（身分＋名字）、at、status（有效／作廢）。同一個 lang＋kind＋zh 只留一筆有效，改譯法＝舊的標作廢、加一筆新的（留歷史）。
   ⚠ 這份表**不會**自動寫進門市食譜後台的「食材名稱對照表」；翻譯分頁會優先用它，後台名稱要另外補（📤 新食譜上傳建新品項時可一起填三地名稱）。 */
var TR_CONF_TAB = 'tr_confirmed', TR_CONF_HEAD = ['lang', 'kind', 'zh', 'tr', 'unit', 'note', 'src', 'by', 'at', 'status'];
var TR_CONF_KEY = 'trconf:v1', TR_LANGS = { mo: 1, ph: 1, vn: 1, me: 1, en: 1 }, TR_CONF_MAX = 300;   /* 10/19 v33 加 me＝澳門英文、en＝英文（台灣門市英文版） */
/* 後台「新增食材」頁 10 個語言列的 LanguageId（10/19 讀頁面確認；跟食譜的語言 ID 相同） */
var TR_LANG_ID = { mo: 'ee184749-9860-48f4-a76f-14c80dd8fdb2', ph: '73743a68-2ef5-4e71-be40-3bebd8d97712', vn: '68d41688-4c57-44ed-b35d-0fcdbd1ee292' };

function trConfSheet_(create) {
  var ss = ss_(), sh = ss.getSheetByName(TR_CONF_TAB);
  if (!sh && create) {
    sh = ss.insertSheet(TR_CONF_TAB);
    sh.getRange(1, 1, sh.getMaxRows(), TR_CONF_HEAD.length).setNumberFormat('@');   /* 全部存文字（0.5、6吋這種不被轉格式） */
    sh.getRange(1, 1, 1, TR_CONF_HEAD.length).setValues([TR_CONF_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function trConfKey_(lang, kind, zh) { return lang + '|' + kind + '|' + rbNorm_(zh); }
function trConfRows_() {
  var sh = trConfSheet_(false);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, TR_CONF_HEAD.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 }; TR_CONF_HEAD.forEach(function (h, j) { o[h] = str_(r[j]); }); return o;
  });
}
function trConfList_(p) {
  var cache = CacheService.getScriptCache(), hit = !p.fresh && cacheGetBig_(cache, TR_CONF_KEY);
  if (hit) return { data: JSON.parse(hit) };
  var rows = trConfRows_().filter(function (o) { return o.status !== '作廢'; });
  var d = { at: now_(), head: TR_CONF_HEAD.concat(['row']), rows: rows.map(function (o) { return TR_CONF_HEAD.map(function (h) { return o[h]; }).concat([o._row]); }) };
  try { cachePutBig_(cache, TR_CONF_KEY, JSON.stringify(d), 600); } catch (e) { }
  return { data: d };
}
function trConfClean_(it) {
  it = it || {};
  var o = { lang: String(it.lang || ''), kind: String(it.kind || ''), zh: rbTrim_(it.zh).slice(0, 80), tr: rbTrim_(it.tr).slice(0, 160),
    unit: rbTrim_(it.unit).slice(0, 20), note: rbTrim_(it.note).slice(0, 200), src: rbTrim_(it.src).slice(0, 80) };
  if (!TR_LANGS[o.lang]) throw fail_('語言不對（只能是澳門中文、菲律賓英文、越南文、澳門英文、英文）', 'input');
  if (o.kind !== 'item' && o.kind !== 'term') throw fail_('種類不對（品項或用語）', 'input');
  if (!o.zh || !o.tr) throw fail_('「' + (o.zh || '（空白）') + '」的原文或譯法是空的', 'input');
  return o;
}
/* 寫入（呼叫端要先拿鎖：doPost 的寫入路徑、或 nrIngCreate_ 自己拿） */
function trConfWrite_(items, by) {
  var sh = trConfSheet_(true), rows = trConfRows_(), live = {}, add = [], kill = [], same = 0, seen = {};
  rows.forEach(function (o) { if (o.status !== '作廢') live[trConfKey_(o.lang, o.kind, o.zh)] = o; });
  items.forEach(function (o) {
    var k = trConfKey_(o.lang, o.kind, o.zh);
    if (seen[k]) return; seen[k] = 1;   /* 同一批重複的只取第一筆 */
    var old = live[k];
    if (old && old.tr === o.tr && (o.kind !== 'item' || old.unit === o.unit)) { same++; return; }
    if (old) kill.push(old._row);
    add.push([o.lang, o.kind, o.zh, o.tr, o.unit, o.note, o.src, by, now_(), '有效']);
  });
  kill.forEach(function (r) { sh.getRange(r, TR_CONF_HEAD.indexOf('status') + 1).setValue('作廢'); });
  if (add.length) {
    var at = sh.getLastRow() + 1;
    sh.getRange(at, 1, add.length, TR_CONF_HEAD.length).setNumberFormat('@').setValues(add);
    var back = sh.getRange(at, 1, add.length, TR_CONF_HEAD.length).getValues();   /* 讀回核對 */
    for (var i = 0; i < add.length; i++) if (str_(back[i][2]) !== add[i][2] || str_(back[i][3]) !== add[i][3]) throw fail_('寫進確認用語表後讀回不一致（第 ' + (at + i) + ' 列），請再按一次', 'verify');
  }
  CacheService.getScriptCache().remove(TR_CONF_KEY);
  return { added: add.length, replaced: kill.length, same: same };
}
function trConfirm_(p, auth) {
  if (!canPush_(auth)) throw fail_('確認譯法要「主廚」或「管理者」密碼', 'perm');
  var raw = Array.isArray(p.items) ? p.items : [];
  if (!raw.length) throw fail_('沒有要確認的譯法', 'input');
  if (raw.length > TR_CONF_MAX) throw fail_('一次最多確認 ' + TR_CONF_MAX + ' 筆', 'input');
  var items = raw.map(trConfClean_), by = auth.role + (rbTrim_(p.name) ? '：' + rbTrim_(p.name).slice(0, 20) : '');
  var r = trConfWrite_(items, by);
  return { id: '', data: r, msg: ('確認譯法 新增 ' + r.added + '（取代 ' + r.replaced + '）、相同 ' + r.same + '｜' + items.slice(0, 20).map(function (o) { return o.lang + ':' + o.zh + '→' + o.tr; }).join('、')).slice(0, 480) };
}
function trConfDel_(p, auth) {
  if (!canPush_(auth)) throw fail_('作廢譯法要「主廚」或「管理者」密碼', 'perm');
  var row = Number(p.row), sh = trConfSheet_(false);
  if (!sh || !(row >= 2) || row > sh.getLastRow()) throw fail_('找不到這一筆', 'input');
  var v = sh.getRange(row, 1, 1, TR_CONF_HEAD.length).getValues()[0];
  if (str_(v[2]) !== rbTrim_(p.zh) || str_(v[0]) !== String(p.lang || '')) throw fail_('這一筆已經被改過，請重新整理再試', 'stale');
  sh.getRange(row, TR_CONF_HEAD.indexOf('status') + 1).setValue('作廢');
  CacheService.getScriptCache().remove(TR_CONF_KEY);
  return { id: '', data: { row: row }, msg: '作廢 ' + v[0] + ':' + v[2] + '→' + v[3] };
}
ACTIONS.trConfList = function (p, a) { return trConfList_(p, a); };
ACTIONS.trConfirm = function (p, a) { return trConfirm_(p, a); };
ACTIONS.trConfDel = function (p, a) { return trConfDel_(p, a); };
READS.trConfList = 1;
WRITES.trConfirm = 1;   /* 走寫入路徑：排鎖、回條防重送、寫 log */
WRITES.trConfDel = 1;
/* 編輯器手動驗證（會在 tr_confirmed 寫 2 筆測試、取代 1 次、作廢，最後把測試列刪掉）＋建品項送出內容（不送後台） */
function trConfTestNow() {
  var lk = LockService.getScriptLock(); lk.waitLock(20000);
  var log = {};
  try {
    var z = '__翻譯測試__' + Date.now();
    log.w1 = trConfWrite_([trConfClean_({ lang: 'mo', kind: 'item', zh: z, tr: '測試甲', unit: 'g', src: 'trConfTestNow' }), trConfClean_({ lang: 'vn', kind: 'term', zh: z, tr: 'thử', src: 'trConfTestNow' })], '測試');
    log.w2same = trConfWrite_([trConfClean_({ lang: 'mo', kind: 'item', zh: z, tr: '測試甲', unit: 'g' })], '測試');
    log.w3replace = trConfWrite_([trConfClean_({ lang: 'mo', kind: 'item', zh: z, tr: '測試乙', unit: 'g' })], '測試');
    var live = trConfList_({ fresh: 1 }).data.rows.filter(function (r) { return r[2] === z; });
    log.live = live.map(function (r) { return r[0] + ':' + r[3]; });
    var sh = trConfSheet_(false), rows = trConfRows_().filter(function (o) { return o.zh === z; });
    log.allRows = rows.map(function (o) { return o.lang + ':' + o.tr + ':' + o.status; });
    rows.map(function (o) { return o._row; }).sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
    CacheService.getScriptCache().remove(TR_CONF_KEY);
    log.after = trConfRows_().filter(function (o) { return o.zh === z; }).length;
  } finally { lk.releaseLock(); }
  /* 建品項：三地名稱有沒有放進對的語言列（假的 im，不送） */
  var im = { token: 't', groups: [{ id: 'g1', name: '冰箱區' }], zh: { i: 0 }, rows: ['42575226-23b1-4ac3-b123-3cba50e3cee8', '6643fe9e-27fb-436a-896d-a745a7bf91a3', TR_LANG_ID.vn, TR_LANG_ID.ph, 'b78ccd25', 'b87b9040', 'bc9d54e5', 'd1095ff2', 'e71fb8b1', TR_LANG_ID.mo].map(function (l, i) { return { i: i, lang: l, cate: { g1: 'c' + i } }; }) };
  var o = nrIngClean_({ name: '測試果泥', unit: 'g', zone: '冰箱區', tr: { mo: { name: '測試果蓉', unit: '' }, ph: { name: 'Test puree', unit: 'g.' }, vn: { name: '' } } }, ['冰箱區']);
  var cap = null, orig = rbReq_;
  rbReq_ = function (s, m, path, opt) { cap = opt.payload; return { code: 302, loc: '/IngredSources' }; };
  try { nrCreateIng_({ jar: {} }, im, o); } finally { rbReq_ = orig; }
  log.payload = [0, 2, 3, 9].map(function (i) { var p = 'IngredItems[' + i + '].'; return i + ':' + cap[p + 'Name'] + '/' + cap[p + 'Unit'] + '/' + (cap[p + 'Status'] || '') + '/' + cap[p + 'IngredCateItemId']; });
  Logger.log(JSON.stringify(log));
}
