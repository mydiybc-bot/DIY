/* ================= 📤 新食譜上傳（獨立頁 static/recipe-upload.html，2026-10-08） =================
   經營者 10/08：「建立一個新儀表板－新食譜上傳自己做食譜系統：上傳新甜點的食譜，完整建到食譜後台；狀態預設停用、台灣、中文、吳寶春自己做；
   類別不確定就讓使用者確認；有新的食材／耗材／器具／模具，先列出詳細資料請使用者確認，再一併建立；每一步加圖片／影片空白欄位可上傳
   （圖片不超過 300KB、影片不超過 5MB，超過要壓縮或退回）。」
   跟新品申請表的「🍰 匯入自己做食譜系統」用同一套後台連線（rb_backend.gs）與紀錄分頁 fact_recipe_import；差別：
     ‧ 不用申請單：基本資料（名稱、售價、成本、類別、分店類別…）由這一頁填；紀錄的 req_id 欄放上傳編號 UP-…
     ‧ 後台沒有的品項可以在這裡建（後台「食材管理 → 新增食材」，只填繁體中文那一列、啟用；編號由後台自己給）
     ‧ 每一步可以放使用者上傳的圖片／影片（一次請求帶一個檔，用到那一步時前端才送，伺服器不暫存檔案）
   動作（主廚／管理者密碼；掛在本檔尾端，不改 Code.gs）：
     nrPreview（只讀）：登入後台讀類別、分店類別、取放區、食材清單；比對每一步食材工具；列出後台沒有的品項（含同名不同單位、相近名稱）、同名食譜、公版圖片影片
     nrIngCreate（寫後台食材管理）：使用者確認過的新品項逐一建立；建之前先看後台有沒有同名同單位（有就不重建）；建完重讀清單確認
     nrImport（分段）：建食譜（停用、不公開）＋逐步建步驟（食材工具、容器、圖片／影片）；中斷可接續，已建的不重建
     nrList（只讀）：最近 15 次上傳（給「繼續上傳」用）
   ⚠ 只新增，不改、不刪後台任何既有資料。 */
var NR_IMG_MAX = 300 * 1024;          /* 經營者：圖片不超過 300KB（前端先壓到 300,000 bytes 以下，這裡是最後把關） */
var NR_VID_MAX = 5 * 1024 * 1024;     /* 經營者：影片不超過 5MB（後台自己的上限是 5,300,000 bytes） */
var NR_TYPES = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/gif': 'gif', 'video/mp4': 'mp4' };
var NR_STORES = ['自己做海外備用', '自己做', '吳寶春自己做', '體驗中心專用'];   /* 後台台灣的分店類別（10/08 讀後台新增食譜頁） */
var NR_ING_MAX = 40;                  /* 一次最多建幾個新品項 */
var NR_ING_FIELDS = ['IngredCateGroupId', 'IngredItems[0].LanguageId', 'IngredItems[0].Name', 'IngredItems[0].Unit', 'IngredItems[0].Price',
  'IngredItems[0].Cost', 'IngredItems[0].IngredCateItemId', 'IngredItems[0].Status', 'IngredItems[0].DefaultName'];

function nrNeed_(auth) { if (!canPush_(auth)) throw fail_('新食譜上傳要「主廚」或「管理者」密碼', 'perm'); }
function nrNum_(v, name, max) {
  var s = String(v == null ? '' : v).trim();
  if (s === '') return 0;
  var n = Number(s);
  if (!isFinite(n) || n < 0 || n > (max || 1e6)) throw fail_(name + '要填 0 以上的數字', 'input');
  return Math.round(n * 100) / 100;
}

/* ---------- 後台「新增食材」頁：取放區、繁體中文那一列、各語言列 ---------- */
function nrIngMeta_(html, langId) {
  var f = rbFormHtml_(html, /^\/IngredSources\/Create/i);
  if (!f) throw rbErr_('食譜後台「新增食材」頁找不到表單（頁面可能改版）', 'rbpage');
  var miss = NR_ING_FIELDS.filter(function (n) { return !rbHasField_(f, n); });
  if (miss.length) throw rbErr_('食譜後台「新增食材」頁少了欄位 ' + miss.join('、') + '（頁面可能改版，這次先不送）', 'rbpage');
  var gsel = (f.match(/<select\b[^>]*name="IngredCateGroupId"[\s\S]*?<\/select>/i) || [''])[0], groups = [], re = /<option\b([^>]*)>([^<]*)<\/option>/gi, m;
  while ((m = re.exec(gsel))) { var gid = rbAttr_('<x' + m[1] + '>', 'value'); if (gid) groups.push({ id: gid, name: rbTrim_(rbDecode_(m[2])) }); }
  var rows = [], ire = /<input\b[^>]*name="IngredItems\[(\d+)\]\.LanguageId"[^>]*>/gi;
  while ((m = ire.exec(f))) {
    var i = +m[1], sel = (f.match(new RegExp('<select\\b[^>]*name="IngredItems\\[' + i + '\\]\\.IngredCateItemId"[\\s\\S]*?<\\/select>', 'i')) || [''])[0], cmap = {}, om;
    var ore = /<option\b([^>]*)>/gi;
    while ((om = ore.exec(sel))) { var t = '<x' + om[1] + '>', g = rbAttr_(t, 'data-ingredcategroup'), v = rbAttr_(t, 'value'); if (g && v) cmap[g] = v; }
    rows.push({ i: i, lang: rbAttr_(m[0], 'value'), cate: cmap });
  }
  if (rows.length < 1 || !groups.length) throw rbErr_('食譜後台「新增食材」頁讀不到語言列或取放區（頁面可能改版）', 'rbpage');
  var zh = rows.filter(function (r) { return r.lang === langId; })[0];
  if (!zh) throw rbErr_('食譜後台「新增食材」頁找不到台灣「繁體中文」那一列', 'rbpage');
  var bad = groups.filter(function (g) { return !zh.cate[g.id]; });
  if (bad.length) throw rbErr_('食譜後台「新增食材」頁的取放區 ' + bad.map(function (g) { return g.name; }).join('、') + ' 在繁體中文列沒有對應類別（頁面可能改版）', 'rbpage');
  return { token: rbToken_(f), groups: groups, rows: rows, zh: zh };
}
function nrMetaAll_(sess) {
  var pg = rbAuthedGet_(sess, '/Recipes/Create');
  if (pg.code !== 200) throw rbErr_('食譜後台「新增食譜」頁打不開（HTTP ' + pg.code + '）', 'rbnet');
  var rm = rbCreateMeta_(pg.text);
  var ig = rbAuthedGet_(sess, '/IngredSources/Create');
  if (ig.code !== 200) throw rbErr_('食譜後台「新增食材」頁打不開（HTTP ' + ig.code + '）', 'rbnet');
  return { recipe: rm, ing: nrIngMeta_(ig.text, rm.langId) };
}

/* ---------- 比對：跟 🍰 匯入同一套規則（rbMatch_＋單位帶括號合回），另外回結構化的「後台沒有」清單 ---------- */
function nrMatchOne_(ings, g) {
  var r = rbMatch_(ings, { name: g[0], unit: g[2], cate: g[3] }), inUnit = false;
  if (!r.hit && r.miss === 'unit' && g[5]) {
    var want = rbNorm_(g[2] + '(' + g[5] + ')').replace(/\s+/g, '');
    var u2 = r.units.filter(function (u) { return rbNorm_(u).replace(/\s+/g, '') === want; })[0];
    if (u2) { r = rbMatch_(ings, { name: g[0], unit: u2, cate: g[3] }); inUnit = !!r.hit; }
  }
  r.inUnit = inUnit;
  return r;
}
function nrMatchStep_(ings, s) {
  var ok = [], miss = [], notes = [], items = [];
  (s.ing || []).forEach(function (g) {
    var label = g[0] + ' ' + g[1] + ' ' + g[2], r = nrMatchOne_(ings, g);
    if (r.hit) ok.push({ id: r.hit.id, amount: g[1], unit: r.hit.unit, cont: g[4] });
    else {
      miss.push(label + (r.miss === 'unit' ? '（後台這個品項的單位是 ' + r.units.join('／') + '）' : '（後台食材清單沒有）'));
      items.push({ name: g[0], qty: g[1], unit: g[2], zone: g[3], note: g[5] || '', kind: r.miss, units: r.units || [] });
    }
    if (g[5] && !r.inUnit) notes.push(label + '（' + g[5] + '）');
  });
  return { ok: ok, miss: miss, notes: notes, items: items };
}
/* 相近名稱：互相包含，或去掉括號後相同；最多 6 個（給「改用後台既有品項」選） */
function nrSimilar_(ings, name) {
  var n = rbNorm_(name).replace(/\s+/g, ''), core = n.replace(/\([^)]*\)/g, '');
  if (!n) return [];
  var out = [], seen = {};
  ings.forEach(function (x) {
    var k = x.kn.replace(/\s+/g, ''), kc = k.replace(/\([^)]*\)/g, '');
    if (k === n) return;
    var hit = (core.length >= 2 && (kc === core || (kc.length >= 2 && (kc.indexOf(core) >= 0 || core.indexOf(kc) >= 0))));
    if (!hit) return;
    var key = x.name + '|' + x.unit;
    if (seen[key]) return; seen[key] = 1;
    out.push({ name: x.name, unit: x.unit, cate: x.cate, d: Math.abs(kc.length - core.length) });
  });
  return out.sort(function (a, b) { return a.d - b.d; }).slice(0, 6).map(function (x) { return { name: x.name, unit: x.unit, cate: x.cate }; });
}

/* ---------- 預覽（只讀） ---------- */
function nrPreview_(p, auth) {
  nrNeed_(auth);
  var rc = rbCleanRecipe_(p.recipe), title = rbTrim_(p.title || rc.name).slice(0, 100);
  var out = { categories: [], stores: [], zones: [], steps: [], missing: [], dup: [], warn: [], title: title };
  var sess = rbSession_(), mt = nrMetaAll_(sess);
  out.categories = mt.recipe.cates;
  out.stores = mt.recipe.stores.map(function (s) { return s.name; });
  out.zones = mt.ing.groups.map(function (g) { return g.name; });
  if (out.stores.indexOf('吳寶春自己做') < 0) out.warn.push('後台「分店類別」找不到「吳寶春自己做」，請改選其他分店類別');
  var list = null;
  try { list = rbListRows_(rbAuthedGet_(sess, '/Recipes').text); }
  catch (e) { out.warn.push('讀不到後台食譜清單，沒有檢查同名食譜（' + errMsg_(e) + '）'); }
  if (list && title) out.dup = list.filter(function (r) { return rbTitleNorm_(r.title) === rbTitleNorm_(title); }).slice(0, 10);
  var sp = rbProbeIngs_(sess, list), pool = [];
  try { pool = rbMediaPool_(); } catch (e) { out.warn.push('讀不到「食譜素材索引」公版候選，沒有建議公版圖片／影片（' + errMsg_(e) + '）'); }
  var miss = {}, order = [];
  rc.steps.forEach(function (s, i) {
    var m = nrMatchStep_(sp.ings, s), keys = [];
    m.items.forEach(function (it) {
      var k = rbNorm_(it.name).replace(/\s+/g, '') + '|' + rbNorm_(it.unit).replace(/\s+/g, '');
      if (!miss[k]) {
        miss[k] = { key: k, name: it.name, unit: it.unit, zone: it.zone, note: it.note, kind: it.kind, units: it.units, steps: [], qty: [], similar: nrSimilar_(sp.ings, it.name) };
        order.push(k);
      }
      if (miss[k].steps.indexOf(i + 1) < 0) miss[k].steps.push(i + 1);
      miss[k].qty.push(it.qty);
      if (!miss[k].zone && it.zone) miss[k].zone = it.zone;
      keys.push(k);
    });
    out.steps.push({ n: i + 1, title: s.t, ing_n: s.ing.length, ok: m.ok.length, miss: m.miss, mk: keys, notes: m.notes, media: rbMediaFor_(pool, s.t) });
  });
  out.missing = order.map(function (k) { return miss[k]; });
  rbSaveSession_(sess);
  return { id: '', data: out };
}

/* ---------- 建新品項（後台食材管理 → 新增食材） ---------- */
function nrIngClean_(it, zones) {
  it = it || {};
  var o = { name: rbTrim_(it.name).slice(0, 60), unit: rbTrim_(it.unit).slice(0, 20), zone: rbTrim_(it.zone), note: rbTrim_(it.note).slice(0, 200),
    price: nrNum_(it.price, '「' + rbTrim_(it.name) + '」的售價'), cost: nrNum_(it.cost, '「' + rbTrim_(it.name) + '」的成本') };
  if (!o.name) throw fail_('新品項的名稱是空的', 'input');
  if (!o.unit) throw fail_('「' + o.name + '」的單位是空的', 'input');
  if (zones.indexOf(o.zone) < 0) throw fail_('「' + o.name + '」的取放區不對（要從清單選）', 'input');
  return o;
}
function nrFindIng_(ings, o) {
  var n = rbNorm_(o.name), u = rbNorm_(o.unit);
  return ings.filter(function (x) { return (x.kn === n || x.kz === n) && x.ku === u; })[0] || null;
}
function nrCreateIng_(sess, im, o) {
  var gid = im.groups.filter(function (g) { return g.name === o.zone; })[0].id;
  var payload = { '__RequestVerificationToken': im.token, IngredCateGroupId: gid };
  im.rows.forEach(function (r) {
    var pre = 'IngredItems[' + r.i + '].', mine = r.i === im.zh.i;
    payload[pre + 'LanguageId'] = r.lang;
    payload[pre + 'Name'] = mine ? o.name : '';
    payload[pre + 'Note'] = mine ? o.note : '';
    payload[pre + 'Unit'] = mine ? o.unit : '';
    payload[pre + 'Price'] = mine ? String(o.price) : '';
    payload[pre + 'Cost'] = mine ? String(o.cost) : '';
    payload[pre + 'IngredCateItemId'] = r.cate[gid] || '';   /* 後台頁面選取放區時，每個語言的類別都跟著選好（照頁面行為） */
    if (mine) { payload[pre + 'DefaultName'] = 'true'; payload[pre + 'Status'] = 'true'; }
  });
  var res = rbReq_(sess, 'post', '/IngredSources/Create', { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受新品項「' + o.name + '」：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台新增品項「' + o.name + '」回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
function nrIngCreate_(p, auth) {
  nrNeed_(auth);
  var t0 = Date.now(), raw = Array.isArray(p.items) ? p.items : [];
  if (!raw.length) throw fail_('沒有要建立的品項', 'input');
  if (raw.length > NR_ING_MAX) throw fail_('一次最多建 ' + NR_ING_MAX + ' 個品項', 'input');
  var cache = CacheService.getScriptCache();
  if (cache.get('nring')) throw fail_('另一個畫面正在建立新品項，請等 1 分鐘再按', 'busy');
  cache.put('nring', '1', 120);
  try {
    var sess = rbSession_(), mt = nrMetaAll_(sess), zones = mt.ing.groups.map(function (g) { return g.name; });
    var items = raw.map(function (it) { return nrIngClean_(it, zones); });
    var seen = {};
    items.forEach(function (o) { var k = rbNorm_(o.name) + '|' + rbNorm_(o.unit); if (seen[k]) throw fail_('「' + o.name + '（' + o.unit + '）」重複了', 'input'); seen[k] = 1; });
    var before = rbProbeIngs_(sess, null).ings, out = [], err = null;
    for (var i = 0; i < items.length; i++) {
      var o = items[i], ex = nrFindIng_(before, o);
      if (ex) { out.push({ name: o.name, unit: o.unit, zone: o.zone, st: 'exists', id: ex.id }); continue; }
      if (Date.now() - t0 > RB_BUDGET_MS) { out.push({ name: o.name, unit: o.unit, zone: o.zone, st: 'later' }); continue; }
      try {
        try { nrCreateIng_(sess, mt.ing, o); }
        catch (e) {
          if (!e || e.code !== 'rbrelogin') throw e;
          rbDropSession_(); sess.jar = rbFreshSession_().jar;
          var g2 = rbAuthedGet_(sess, '/IngredSources/Create'); mt.ing = nrIngMeta_(g2.text, mt.recipe.langId);
          nrCreateIng_(sess, mt.ing, o);
        }
        out.push({ name: o.name, unit: o.unit, zone: o.zone, st: 'sent' });
      } catch (e2) {
        out.push({ name: o.name, unit: o.unit, zone: o.zone, st: 'unsure', msg: errMsg_(e2) });   /* 連線斷掉＝不確定有沒有建成 → 下面重讀清單判斷 */
        err = e2;
        if (e2 && (e2.code === 'rbreject' || e2.code === 'rbpage' || e2.code === 'rbbad')) break;
      }
    }
    var after = rbProbeIngs_(sess, null).ings, made = 0;
    out.forEach(function (r) {
      if (r.st !== 'sent' && r.st !== 'unsure') return;
      var hit = nrFindIng_(after, r);
      if (hit) { r.st = 'created'; r.id = hit.id; r.cate = hit.cate; made++; delete r.msg; }
      else if (r.st === 'sent') { r.st = 'missing'; r.msg = '後台回覆成功，但重讀食材清單找不到（可能名稱或單位被後台改過）'; }
    });
    rbSaveSession_(sess);
    var left = items.length - out.filter(function (r) { return r.st === 'created' || r.st === 'exists'; }).length;
    if (err && !made && left === items.length) throw fail_('建立新品項沒有成功：' + errMsg_(err), (err && err.code) || 'rbnet', { items: out });
    return { id: '', data: { items: out, made: made, left: left },
      msg: ('新品項：建立 ' + made + '、已存在 ' + out.filter(function (r) { return r.st === 'exists'; }).length + '、未完成 ' + left + '｜' +
        out.map(function (r) { return r.name + '(' + r.unit + ')' + r.st; }).join('、')).slice(0, 480) };   /* 寫進 log 分頁（doPost 記） */
  } finally { cache.remove('nring'); }
}

/* ---------- 建食譜（類別可多選、分店類別照勾選）；停用、不公開 ---------- */
function nrCreateRecipe_(sess, r) {
  var pg = rbAuthedGet_(sess, '/Recipes/Create');
  if (pg.code !== 200) throw rbErr_('食譜後台「新增食譜」頁打不開（HTTP ' + pg.code + '）', 'rbnet');
  var meta = rbCreateMeta_(pg.text);
  var storeIds = (r.stores || []).map(function (n) { var o = meta.stores.filter(function (x) { return x.name === n; })[0]; return o ? o.id : ''; }).filter(String);
  var cats = (r.cats || []).filter(function (id) { return meta.cates.some(function (c) { return c.id === id; }); });
  var payload = {
    '__RequestVerificationToken': meta.token,
    Title: r.title, Code: '', GroupId: meta.groupId, LanguageId: meta.langId, Note: r.note,
    ItemList: cats.join(','), StoreList: storeIds.join(','),
    RecipeUp: '', RecipeDown: '', MenuUp: '', MenuDown: '',
    Price: String(r.price), Cost: String(r.cost), InUse: 'false',
    PrepHr: String(r.prepHr || ''), Size: r.size || '', Content: r.desc || '', Preserve: r.preserve || '',
    Portion: '', PortionType: '', Difficulty: '', PrepMin: '', BakingHr: '', BakingMin: '', RestingHr: '', RestingMin: '',
    Public: 'false'
  };
  var res = rbReq_(sess, 'post', '/Recipes/Create', { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  var id = rbLocId_(res.loc, /\/Recipes\/Edit\/([0-9a-f-]{36})/i);
  if ((res.code === 302 || res.code === 301) && id) return id;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受新增食譜：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台新增食譜回覆異常（HTTP ' + res.code + (res.loc ? '，轉到 ' + res.loc : '') + '）', 'rbunknown');
}

/* ---------- 使用者上傳的圖片／影片（前端送 base64；檢查格式、檔頭、大小） ---------- */
function nrMediaBlob_(md, j) {
  if (!md || md.i !== j) throw fail_('第 ' + (j + 1) + ' 步的圖片／影片沒有收到', 'media');
  var type = String(md.type || ''), ext = NR_TYPES[type];
  if (!ext) throw fail_('第 ' + (j + 1) + ' 步的檔案格式不支援（' + type + '）：圖片用 JPG／PNG／GIF，影片用 MP4', 'media');
  var bytes;
  try { bytes = Utilities.base64Decode(String(md.b64 || '')); } catch (e) { throw fail_('第 ' + (j + 1) + ' 步的檔案傳送不完整，請再試一次', 'media'); }
  var n = bytes.length, isImg = type.indexOf('image/') === 0, max = isImg ? NR_IMG_MAX : NR_VID_MAX;
  if (!n) throw fail_('第 ' + (j + 1) + ' 步的檔案是空的', 'media');
  if (n > max) throw fail_('第 ' + (j + 1) + ' 步的' + (isImg ? '圖片' : '影片') + '超過 ' + (isImg ? '300KB' : '5MB') + '（' + (isImg ? Math.round(n / 1024) + 'KB' : (Math.round(n / 104857.6) / 10) + 'MB') + '）', 'media');
  var b = function (i) { return bytes[i] & 255; };
  var okHead = ext === 'jpg' ? (b(0) === 0xFF && b(1) === 0xD8) : ext === 'png' ? (b(0) === 0x89 && b(1) === 0x50 && b(2) === 0x4E && b(3) === 0x47)
    : ext === 'gif' ? (b(0) === 0x47 && b(1) === 0x49 && b(2) === 0x46) : (n > 12 && b(4) === 0x66 && b(5) === 0x74 && b(6) === 0x79 && b(7) === 0x70);
  if (!okHead) throw fail_('第 ' + (j + 1) + ' 步的檔案內容不是 ' + ext.toUpperCase() + '（副檔名或格式不對）', 'media');
  var blob = Utilities.newBlob(bytes, type, 'step.' + ext);
  return { blob: blob, bytes: n, type: type };
}

/* ---------- 上傳（分段） ---------- */
function nrNewId_() { return 'UP-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4); }
function nrNewJob_(uid, p, auth) {
  var rc = rbCleanRecipe_(p.recipe), o = p.opts || {};
  var title = rbTrim_(o.title).slice(0, 100);
  if (!title) throw fail_('食譜名稱是空的', 'input');
  var cats = (Array.isArray(o.cats) ? o.cats : []).map(function (x) { return str_(x).trim(); }).filter(String);
  if (cats.length > 5) throw fail_('類別最多選 5 個', 'input');
  cats.forEach(function (c) { if (!/^[0-9a-f-]{36}$/i.test(c)) throw fail_('類別不對，請重新檢查', 'input'); });
  var stores = (Array.isArray(o.stores) ? o.stores : []).map(function (x) { return rbTrim_(x); }).filter(String);
  if (!stores.length) throw fail_('請至少勾一個分店類別', 'input');
  stores.forEach(function (s) { if (NR_STORES.indexOf(s) < 0) throw fail_('分店類別「' + s + '」不對', 'input'); });
  var price = nrNum_(o.price, '售價'), cost = nrNum_(o.cost, '成本');
  var hr = String(o.prepHr == null ? '' : o.prepHr).trim();
  if (hr && !/^\d{1,2}(\.\d{1,2})?$/.test(hr)) throw fail_('製作時間（小時）要填數字', 'input');
  var media = Array.isArray(o.media) ? o.media : [], mi = Array.isArray(o.mediaInfo) ? o.mediaInfo : [];
  var steps = rc.steps.map(function (s, i) {
    var u = str_(media[i]).trim(), info = null;
    if (u === 'up') {
      var x = mi[i] || {};
      info = { name: rbTrim_(x.name).slice(0, 80), size: Number(x.size) || 0, type: String(x.type || '').slice(0, 30) };
      if (!NR_TYPES[info.type]) throw fail_('第 ' + (i + 1) + ' 步的檔案格式不支援（' + info.type + '）', 'input');
    } else if (u && u.indexOf(RB_MEDIA_PREFIX) !== 0) throw fail_('第 ' + (i + 1) + ' 步的圖片／影片不是公版素材', 'input');
    return { t: s.t, h: s.h, ing: s.ing, m: u, mi: info };
  });
  var imp = 'IMP-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4);
  var job = { imp_id: imp, rid: uid, kind: 'upload', title: title, cats: cats, stores: stores, price: price, cost: cost, prepHr: hr,
    size: rbTrim_(o.size).slice(0, 100), preserve: String(o.preserve || '').slice(0, 2000), desc: String(o.desc || '').slice(0, 3000),
    file: rc.file, docName: rc.name, by: auth.role, steps: steps };
  rbLog_(uid, imp, 'start', '', '', 'Y', title + '｜' + steps.length + ' 步', auth, JSON.stringify(job));
  job.backend_id = '';
  return job;
}
function nrNote_(job, mm) {
  var L = ['【新食譜上傳】上傳編號 ' + job.rid + '｜匯入編號 ' + job.imp_id + '｜' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') + '｜' + (job.by || ''),
    '食譜檔：' + (job.file || '（未記錄）') + (job.docName ? '（' + job.docName + '）' : '')];
  var miss = [], notes = [];
  mm.forEach(function (m, i) {
    m.miss.forEach(function (x) { miss.push('第 ' + (i + 1) + ' 步：' + x); });
    m.notes.forEach(function (x) { notes.push('第 ' + (i + 1) + ' 步：' + x); });
  });
  if (miss.length) L.push('沒掛上的食材工具（請食譜負責夥伴補上）：', miss.join('\n'));
  if (notes.length) L.push('食譜檔的附註沒帶進後台（後台只記數量和單位）：', notes.join('\n'));
  L.push('上傳時是「停用」：檢查調整好再到這裡勾「啟用」。');
  return L.join('\n').slice(0, 4000);
}
function nrImport_(p, auth) {
  nrNeed_(auth);
  var t0 = Date.now(), cache = CacheService.getScriptCache(), uid = str_(p.up_id).trim(), rqk = '';
  if (!uid && p.rq) {   /* 第一次送出：同一個回條編號重送（回覆在路上掉了）→ 接著同一個上傳，不會多建一支 */
    rqk = 'nrrq:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(String(p.rq).slice(0, 80)).getBytes());
    uid = cache.get(rqk) || '';
  }
  if (uid && !/^UP-\d{8}-\d{6}-[0-9a-f]{4}$/.test(uid)) throw fail_('上傳編號不對', 'input');
  var lockK = 'nrimp:' + (uid || 'new:' + String(p.rq || ''));
  if (cache.get(lockK)) throw fail_('這支食譜正在上傳中（可能另一個畫面也按了），請等 1 分鐘再按「繼續上傳」', 'busy');
  cache.put(lockK, '1', 120);
  try {
    var job;
    if (uid) {
      var st = rbState_(uid), im = p.imp_id ? st.imps[str_(p.imp_id)] : (st.open || st.done);
      if (!im) {
        if (!p.up_id && p.recipe) { job = nrNewJob_(uid, p, auth); }   /* 回條有記、但 start 還沒寫成 → 重新開始 */
        else throw fail_('找不到上傳紀錄 ' + uid, 'notfound');
      } else {
        if (im.abandoned) throw fail_('這次上傳已作廢，請重新上傳', 'state');
        if (im.done) return { id: uid, data: { up_id: uid, imp_id: im.imp_id, backend_id: im.backend_id, link: rbLink_(im.backend_id), done: true, next: im.total, total: im.total, need: -1 }, msg: '已完成' };
        job = rbJob_(st, im);
      }
    } else {
      if (!p.recipe) throw fail_('沒有收到食譜內容', 'input');
      uid = nrNewId_();
      if (rqk) cache.put(rqk, uid, 21600);
      job = nrNewJob_(uid, p, auth);
    }
    var res = nrRun_(uid, job, auth, t0, p.media || null, (p.skip_media === 0 || p.skip_media > 0) ? Number(p.skip_media) : -1);
    return { id: uid, data: res, msg: res.done ? '上傳完成 ' + res.total + ' 步' : '上傳到第 ' + res.next + '／' + res.total + ' 步' };
  } finally { cache.remove(lockK); }
}
function nrRun_(uid, job, auth, t0, media, skip) {
  var sess = rbSession_(), N = job.steps.length;
  var relogin = function () { rbDropSession_(); sess = rbFreshSession_(); };
  var base = { up_id: uid, imp_id: job.imp_id, total: N };
  var ext = function (o) { for (var k in base) o[k] = base[k]; return o; };
  if (!job.backend_id) {
    var found = rbFindByNote_(sess, job.imp_id);   /* 回覆遺失保險：備註已有這個匯入編號＝建過了 */
    if (found.length) job.backend_id = found[0].id;
    else {
      var probe = rbProbeIngs_(sess, null);
      var mm = job.steps.map(function (s) { return nrMatchStep_(probe.ings, s); });
      var r = { title: job.title, note: nrNote_(job, mm), cats: job.cats, stores: job.stores, price: job.price, cost: job.cost,
        prepHr: job.prepHr, size: job.size, desc: job.desc, preserve: job.preserve };
      try {
        try { job.backend_id = nrCreateRecipe_(sess, r); }
        catch (e) { if (e && e.code === 'rbrelogin') { relogin(); job.backend_id = nrCreateRecipe_(sess, r); } else throw e; }
      } catch (e2) {
        rbLog_(uid, job.imp_id, 'fail', '', '', 'N', '建食譜：' + errMsg_(e2), auth);
        throw fail_('在後台建食譜沒有成功：' + errMsg_(e2), (e2 && e2.code) || 'rbnet', ext({ next: 0, need: -1 }));
      }
    }
    rbLog_(uid, job.imp_id, 'recipe', job.backend_id, '', 'Y', '後台食譜已建好（停用、不公開）', auth);
  }
  var bid = job.backend_id, idx = rbStepIndex_(sess, bid), k = idx.length;
  base.backend_id = bid; base.link = rbLink_(bid);
  var badAt = function (lst) {
    for (var i = 0; i < lst.length; i++) if (i >= N || rbTitleNorm_(lst[i].title) !== rbTitleNorm_(job.steps[i].t)) return i + 1;
    return 0;
  };
  var b0 = badAt(idx);
  if (b0) {
    rbLog_(uid, job.imp_id, 'fail', bid, '', 'N', '後台步驟和要上傳的不一致（第 ' + b0 + ' 步）', auth);
    throw fail_('後台這支食譜的第 ' + b0 + ' 步和要上傳的不一樣（可能有人在後台改過），先停下來。請食譜負責夥伴檢查後台', 'mismatch', ext({ next: k, need: -1 }));
  }
  var made = [], stopErr = null, need = -1;
  if (k < N && Date.now() - t0 < RB_BUDGET_MS) {
    var sp = rbStepPage_(sess, bid);
    for (var j = k; j < N; j++) {
      if (Date.now() - t0 > RB_BUDGET_MS) break;
      var s = job.steps[j], m = nrMatchStep_(sp.ings, s), md = null, mwarn = '';
      if (s.m === 'up') {
        if (media && media.i === j) {
          try { md = nrMediaBlob_(media, j); media = null; }
          catch (em) { stopErr = em; break; }   /* 格式／大小不對：這步先不建，記下前面建好的，回報原因（code=media，前端請使用者換檔或略過） */
        }
        else if (skip === j) mwarn = '這步原本要放上傳的' + ((s.mi && /^video/.test(s.mi.type)) ? '影片' : '圖片') + '《' + ((s.mi && s.mi.name) || '') + '》，上傳時略過了';
        else { need = j; break; }   /* 要等前端把這一步的檔案送來 */
      } else if (s.m) { try { md = rbFetchMedia_(s.m); } catch (e) { mwarn = errMsg_(e); } }
      var one = { title: s.t, html: rbSanitize_(s.h), ings: m.ok, media: md ? md.blob : null };
      try {
        try { rbCreateStep_(sess, bid, sp.token, one); }
        catch (e) { if (e && e.code === 'rbrelogin') { relogin(); sp = rbStepPage_(sess, bid); rbCreateStep_(sess, bid, sp.token, one); } else throw e; }
      } catch (e3) { stopErr = e3; break; }
      made.push({ n: j + 1, t: s.t, ing: m.ok.length, media: md ? (md.type.indexOf('video') === 0 ? '影片' : '圖片') : '', mwarn: mwarn, miss: m.miss, notes: m.notes });
    }
  }
  var idx2 = rbStepIndex_(sess, bid), k2 = idx2.length, b2 = badAt(idx2);
  if (k2 > k) rbLog_(uid, job.imp_id, 'steps', bid, (k + 1) + '-' + k2, 'Y', '建好第 ' + (k + 1) + '～' + k2 + ' 步' + (k2 < k + made.length ? '（核對時少了 ' + (k + made.length - k2) + ' 步）' : ''), auth, JSON.stringify(made));
  rbSaveSession_(sess);
  if (b2 || k2 < k + made.length) {
    rbLog_(uid, job.imp_id, 'fail', bid, '', 'N', '核對後台步驟清單不一致', auth);
    throw fail_('建完步驟後核對後台步驟清單不一致（第 ' + (b2 || (k2 + 1)) + ' 步），先停下來；請食譜負責夥伴檢查後台，或通知 Claude', 'mismatch', ext({ next: k2, need: -1 }));
  }
  if (stopErr && k2 < N) {
    rbLog_(uid, job.imp_id, 'fail', bid, String(k2 + 1), 'N', '第 ' + (k2 + 1) + ' 步：' + errMsg_(stopErr), auth);
    var isMd = stopErr && stopErr.code === 'media';
    throw fail_('第 ' + (k2 + 1) + ' 步沒有建成：' + errMsg_(stopErr), (stopErr && stopErr.code) || 'rbnet', ext({ next: k2, need: isMd ? k2 : -1, needInfo: isMd ? (job.steps[k2].mi || null) : null }));
  }
  var done = k2 >= N;
  if (done) rbLog_(uid, job.imp_id, 'done', bid, String(N), 'Y', '上傳完成：' + N + ' 步', auth);
  var nd = (!done && need >= 0 && need === k2) ? need : -1;
  return ext({ done: done, next: k2, need: nd, needInfo: nd >= 0 ? (job.steps[nd].mi || null) : null });
}

/* ---------- 最近的上傳（只讀；給「繼續上傳」） ---------- */
function nrList_(p, auth) {
  nrNeed_(auth);
  var sh = ss_().getSheetByName('fact_recipe_import'), out = {}, order = [];
  if (!sh || sh.getLastRow() < 2) return { id: '', data: [] };
  var H = TABS.fact_recipe_import, v = sh.getRange(2, 1, sh.getLastRow() - 1, RB_LOG_N).getValues();
  v.forEach(function (r) {
    var rid = str_(r[1]);
    if (rid.indexOf('UP-') !== 0) return;
    var o = {}; for (var j = 0; j < RB_LOG_N; j++) o[H[j]] = str_(r[j]);
    var u = out[rid];
    if (!u) { u = out[rid] = { up_id: rid, imp_id: '', title: '', total: 0, made: 0, backend_id: '', done: false, abandoned: false, ts: o.ts, by: o.by, fail: '' }; order.push(rid); }
    if (o.action === 'start') { u.imp_id = o.imp_id; var mm = String(o.msg).match(/^(.*)｜(\d+) 步$/); if (mm) { u.title = mm[1]; u.total = +mm[2]; } }
    if (o.action === 'recipe' && o.backend_id) u.backend_id = o.backend_id;
    if (o.action === 'steps') { var sm = String(o.step).match(/(\d+)$/); if (sm) u.made = Math.max(u.made, +sm[1]); u.fail = ''; }
    if (o.action === 'fail') u.fail = o.msg + '（' + o.ts + '）';
    if (o.action === 'done') { u.done = true; u.made = u.total; u.done_ts = o.ts; u.fail = ''; }
    if (o.action === 'abandon') u.abandoned = true;
  });
  var list = order.slice(-15).reverse().map(function (k) { var u = out[k]; u.link = rbLink_(u.backend_id); return u; });
  return { id: '', data: list };
}

/* ---------- 掛上 API（本檔排在 Code.js 後面；包一層＝呼叫時才找函式） ---------- */
ACTIONS.nrPreview = function (p, a) { return nrPreview_(p, a); };
ACTIONS.nrIngCreate = function (p, a) { return nrIngCreate_(p, a); };
ACTIONS.nrImport = function (p, a) { return nrImport_(p, a); };
ACTIONS.nrList = function (p, a) { return nrList_(p, a); };
READS.nrPreview = 1; READS.nrList = 1;                 /* 只讀：不排全系統的鎖 */
SELFLOCK.nrIngCreate = 1; SELFLOCK.nrImport = 1;       /* 要連後台、一次 30 多秒：自己用 CacheService 擋重複，不佔全系統的鎖 */
