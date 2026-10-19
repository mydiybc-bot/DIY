/* ================= 🌐 食譜翻譯：讀後台食譜原文＋三地食材名稱對照（2026-10-19） =================
   經營者 10/19：「把任何食譜自由轉換成澳門中文、菲律賓英文、越南文的食譜；三地對同一個食譜有不同用語、名稱（特別是食材、耗材、包裝盒、器具和模具）」。
   第一段只產出翻譯稿給人看，**不寫進後台**。翻譯本身在前端呼叫訓練站的 Claude 代理（/api/recipe/claude，食譜小幫手密碼）；
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
