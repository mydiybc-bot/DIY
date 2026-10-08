/* ================= 🍰 匯入自己做食譜系統：跟門市食譜後台（diybc.azurewebsites.net）講話（2026-10-06） =================
   經營者 10/06：新品申請核准後按「📥 匯入自己做食譜系統」→ 在後台建一支「停用」的正式食譜（基本資料＋步驟＋每步食材工具＋公版圖片／影片），
   由負責食譜的夥伴檢查調整後再啟用給客人用。
   本檔只管跟後台連線：登入、讀表單、送出新增、讀步驟清單、抓公版圖片影片；不碰試算表（流程、權限、紀錄在 Code.gs 的 rb*_）。
   ⚠ 只會「新增」食譜與步驟；不改、不刪後台任何既有資料（剛建好的那支也不回頭改）。
   ⚠ 後台帳密存指令碼屬性 RB_EMAIL／RB_PASSWORD（管理者在食譜系統「🔑 密碼管理」輸入，先試登入成功才存）；任何回覆、紀錄都不含密碼。
   ⚠ 登入被拒（帳密不對、帳號被鎖）→ 設 RB_BAD，之後不再自動重試（避免連錯把後台帳號鎖住），等管理者重設。
   ⚠ 送出（POST）不自動重試：回覆遺失時由上層讀步驟清單判斷有沒有建成，不重送（避免重複建）。 */
var RB_BASE = 'https://diybc.azurewebsites.net';
var RB_UA = 'Mozilla/5.0 (compatible; DIYBC-recipe-req)';
var RB_COOKIE_TTL = 20 * 60 * 1000;                                   /* 登入狀態暫存 20 分鐘（同訂位管線） */
var RB_MEDIA_PREFIX = 'https://diybcstorage.blob.core.windows.net/stepmedias/';   /* 只抓自家雲端空間的公版素材 */
var RB_MEDIA_MAX = 5300000;                                          /* 後台步驟圖片／影片上限（後台頁面自己的檢查值） */
var RB_CREATE_FIELDS = ['Title', 'GroupId', 'LanguageId', 'Note', 'ItemList', 'StoreList', 'Price', 'Cost', 'InUse', 'PrepHr', 'Size', 'Content', 'Preserve', 'Public'];
var RB_STEP_FIELDS = ['RecipeId', 'StepTitle', 'Content', 'StopClock', 'Timer', 'Image'];
/* 2026-10-08 正式環境 bug 修正：建食譜時「食譜上架／食譜下架」送空白 → 門市平板點這支食譜會被導回分類頁（打不開）。
   現有可播放的食譜幾乎都是 2023-01-01 00:00 → 2030-01-01 00:00；後台用格林威治時間比對，所以不要填「今天幾點」這種短期間。
   ⚠ 2030-01-01 到期前要把所有食譜的下架日整批延長（新品上傳、🍰 匯入都用這兩個常數） */
var RB_RECIPE_UP = '2023-01-01T00:00', RB_RECIPE_DOWN = '2030-01-01T00:00';

function rbErr_(msg, code) { var e = new Error(msg); e.code = code; return e; }
function rbChr_(n) {
  if (!(n > 0)) return '';
  if (n < 0x10000) return String.fromCharCode(n);
  n -= 0x10000; return String.fromCharCode(0xD800 + (n >> 10), 0xDC00 + (n & 0x3FF));
}
function rbDecode_(s) {
  return String(s == null ? '' : s)
    .replace(/&#x([0-9a-f]+);/gi, function (m, h) { return rbChr_(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (m, d) { return rbChr_(parseInt(d, 10)); })
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, '\u00a0').replace(/&amp;/g, '&');
}
function rbAttr_(tag, name) { var m = String(tag).match(new RegExp('\\s' + name + '="([^"]*)"', 'i')); return m ? rbDecode_(m[1]) : ''; }
function rbTrim_(s) { return String(s == null ? '' : s).replace(/[\u00a0\s]+/g, ' ').trim(); }   /* 存值用：只整理空白，全形括號等原字不動 */
function rbNorm_(s) { return rbTrim_(s).replace(/（/g, '(').replace(/）/g, ')'); }   /* 比對用：全形／半形括號視為相同 */
function rbText_(html) { return rbTrim_(rbDecode_(String(html || '').replace(/<[^>]+>/g, ' '))); }

/* ---- cookie ---- */
function rbJarAdd_(jar, resp) {
  var h = resp.getAllHeaders(), sc = h['Set-Cookie'] || h['set-cookie'];
  if (!sc) return;
  if (typeof sc === 'string') sc = [sc];
  sc.forEach(function (c) {
    var kv = String(c).split(';')[0], eq = kv.indexOf('=');
    if (eq <= 0) return;
    var k = kv.slice(0, eq).trim(), v = kv.slice(eq + 1).trim();
    if (v === '') delete jar[k]; else jar[k] = v;   /* 空值＝伺服器要清掉這個 cookie */
  });
}
function rbJarStr_(jar) { return Object.keys(jar).map(function (k) { return k + '=' + jar[k]; }).join('; '); }

/* ---- 一次請求：不自動跟轉址；讀取（GET）遇 502／503／504 或連線例外最多再試 2 次；送出（POST）一律不重試 ---- */
function rbReq_(sess, method, path, opt) {
  opt = opt || {};
  var url = /^https?:/i.test(path) ? path : RB_BASE + path;
  var o = { method: method, muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA } };
  if (sess && sess.jar) { var c = rbJarStr_(sess.jar); if (c) o.headers.Cookie = c; }
  if (opt.payload !== undefined) o.payload = opt.payload;
  var waits = (method === 'get' && opt.retry !== false) ? [2000, 5000] : [], last = '';
  for (var a = 0; a <= waits.length; a++) {
    try {
      var r = UrlFetchApp.fetch(url, o), code = r.getResponseCode();
      if (sess && sess.jar) rbJarAdd_(sess.jar, r);
      if ((code === 502 || code === 503 || code === 504) && a < waits.length) { last = 'HTTP ' + code; Utilities.sleep(waits[a]); continue; }
      var h = r.getAllHeaders();
      return { code: code, loc: String(h.Location || h.location || ''), text: opt.binary ? '' : r.getContentText(), resp: r };
    } catch (e) {
      last = String((e && e.message) || e);
      if (a < waits.length) { Utilities.sleep(waits[a]); continue; }
    }
  }
  throw rbErr_('連不到食譜後台（' + last + '）', 'rbnet');
}
function rbIsLogin_(r) { return (r.code === 302 || r.code === 301) && /\/Identity\/Account\/Login/i.test(r.loc); }
function rbGet_(sess, path) {
  var r = rbReq_(sess, 'get', path);
  if (rbIsLogin_(r)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  return r;
}

/* ---- 登入（欄位名照登入頁動態抓，同訂位管線 login_） ---- */
function rbLoginRaw_(email, pass) {
  var sess = { jar: {} }, path = '/Identity/Account/Login?ReturnUrl=%2FRecipes';
  var r1 = rbReq_(sess, 'get', path);
  if (r1.code !== 200) throw rbErr_('食譜後台登入頁打不開（HTTP ' + r1.code + '）', 'rbnet');
  var html = r1.text;
  var tokTag = (html.match(/<input\b[^>]*name="__RequestVerificationToken"[^>]*>/i) || [''])[0];
  var tok = rbAttr_(tokTag, 'value');
  if (!tok) throw rbErr_('食譜後台登入頁找不到驗證碼欄位（頁面可能改版）', 'rbpage');
  var pwName = rbAttr_((html.match(/<input\b[^>]*type="password"[^>]*>/i) || [''])[0], 'name') || 'Input.Password';
  var emName = rbAttr_((html.match(/<input\b[^>]*type="email"[^>]*>/i) || [''])[0], 'name') || 'Input.Email';
  var payload = { '__RequestVerificationToken': tok, 'Input.RememberMe': 'false' };
  payload[emName] = email; payload[pwName] = pass;
  var r2 = rbReq_(sess, 'post', path, { payload: payload });
  var authed = Object.keys(sess.jar).some(function (k) { return k.indexOf('Identity.Application') >= 0; });
  if ((r2.code === 302 || r2.code === 301) && authed) return sess;
  if (/lockout/i.test(r2.loc)) throw rbErr_('食譜後台帳號被暫時鎖住（登入錯太多次），請稍後再試或請管理者處理', 'rbbad');
  if (/2fa|TwoFactor/i.test(r2.loc)) throw rbErr_('這個食譜後台帳號要兩步驟驗證，系統沒辦法自動登入；請改用不需要兩步驟驗證的帳號', 'rbbad');
  if (r2.code === 200 && /type="password"/i.test(r2.text)) throw rbErr_('食譜後台拒絕登入：帳號或密碼不對', 'rbbad');
  throw rbErr_('食譜後台登入沒有成功（HTTP ' + r2.code + '）', 'rbnet');
}
function rbProps_() { return PropertiesService.getScriptProperties(); }
function rbSaveSession_(s) {
  var P = rbProps_();
  P.setProperty('RB_COOKIE', JSON.stringify(s.jar)); P.setProperty('RB_COOKIE_AT', String(Date.now())); P.setProperty('RB_OK_AT', String(Date.now()));
}
function rbDropSession_() { rbProps_().deleteProperty('RB_COOKIE'); }
function rbFreshSession_() {
  var P = rbProps_();
  if (P.getProperty('RB_BAD') === '1') throw rbErr_('食譜後台帳號密碼上次被拒絕，請管理者到「🔑 密碼管理」重新設定後再試', 'rbbad');   /* 2026-10-08：📤 上傳／覆蓋也會碰到，不只匯入 */
  var email = P.getProperty('RB_EMAIL'), pass = P.getProperty('RB_PASSWORD');
  if (!email || !pass) throw rbErr_('還沒有設定食譜後台帳號：請管理者到「🔑 密碼管理」→「🍰 自己做食譜系統登入帳號」設定', 'rbnocred');
  var s;
  try { s = rbLoginRaw_(email, pass); }
  catch (e) { if (e && e.code === 'rbbad') { P.setProperty('RB_BAD', '1'); rbDropSession_(); } throw e; }
  rbSaveSession_(s);
  return s;
}
function rbSession_() {
  var P = rbProps_(), c = P.getProperty('RB_COOKIE'), at = Number(P.getProperty('RB_COOKIE_AT') || 0);
  if (c && Date.now() - at < RB_COOKIE_TTL && P.getProperty('RB_BAD') !== '1') { try { return { jar: JSON.parse(c), cached: true }; } catch (e) { } }
  return rbFreshSession_();
}
/* 讀取類：被導回登入頁（逾時）→ 重新登入一次再讀 */
function rbAuthedGet_(sess, path) {
  try { return rbGet_(sess, path); }
  catch (e) {
    if (!e || e.code !== 'rbrelogin') throw e;
    rbDropSession_();
    var s2 = rbFreshSession_();
    sess.jar = s2.jar;
    return rbGet_(sess, path);
  }
}

/* ---- 讀表單 ---- */
function rbFormHtml_(html, actionRe) {
  var re = /<form\b[^>]*>/gi, m;
  while ((m = re.exec(html))) {
    if (actionRe.test(rbAttr_(m[0], 'action'))) {
      var end = html.indexOf('</form>', m.index);
      return html.slice(m.index, end < 0 ? html.length : end);
    }
  }
  return '';
}
function rbToken_(formHtml) { return rbAttr_((formHtml.match(/<input\b[^>]*name="__RequestVerificationToken"[^>]*>/i) || [''])[0], 'value'); }
function rbHasField_(formHtml, name) { return new RegExp('\\sname="' + name.replace(/[.\[\]]/g, '\\$&') + '"', 'i').test(formHtml); }
function rbOptions_(html, cls) {
  var out = [], re = /<option\b([^>]*)>([^<]*)<\/option>/gi, m;
  while ((m = re.exec(html))) {
    var tag = '<x' + m[1] + '>';
    if (rbAttr_(tag, 'class').split(/\s+/).indexOf(cls) < 0) continue;
    out.push({ id: rbAttr_(tag, 'value'), name: rbTrim_(rbDecode_(m[2])), group: rbAttr_(tag, 'data-groupid'), lang: rbAttr_(tag, 'data-langid') });
  }
  return out;
}
/* 「新增食譜」頁：驗證碼、地區／語言／類別／分店類別的選項（全部動態讀，不寫死編號） */
function rbCreateMeta_(html) {
  var f = rbFormHtml_(html, /^\/Recipes\/Create/i);
  if (!f) throw rbErr_('食譜後台「新增食譜」頁找不到表單（頁面可能改版）', 'rbpage');
  var miss = RB_CREATE_FIELDS.filter(function (n) { return !rbHasField_(f, n); });
  if (miss.length) throw rbErr_('食譜後台「新增食譜」頁少了欄位 ' + miss.join('、') + '（頁面可能改版，這次先不送）', 'rbpage');
  var groups = [], gsel = (f.match(/<select\b[^>]*id="GroupId"[\s\S]*?<\/select>/i) || [''])[0], re = /<option\b([^>]*)>([^<]*)<\/option>/gi, m;
  while ((m = re.exec(gsel))) groups.push({ id: rbAttr_('<x' + m[1] + '>', 'value'), name: rbTrim_(rbDecode_(m[2])) });
  var tw = groups.filter(function (g) { return g.name === '台灣'; })[0];
  if (!tw) throw rbErr_('食譜後台「地區」找不到「台灣」', 'rbpage');
  var lang = rbOptions_(f, 'langOption').filter(function (o) { return o.group === tw.id && o.name === '繁體中文'; })[0];
  if (!lang) throw rbErr_('食譜後台「語言」找不到台灣的「繁體中文」', 'rbpage');
  return {
    token: rbToken_(f), groupId: tw.id, langId: lang.id,
    cates: rbOptions_(f, 'cateOption').filter(function (o) { return o.lang === lang.id && o.name; }).map(function (o) { return { id: o.id, name: o.name }; }),
    stores: rbOptions_(f, 'storeOption').filter(function (o) { return o.group === tw.id; }).map(function (o) { return { id: o.id, name: o.name }; })
  };
}
/* 「新增步驟」頁的食材／器具清單（後台食材管理的全部品項） */
function rbIngList_(html) {
  var out = [], re = /<div\s+class="row allIngred"([^>]*)>\s*<div[^>]*>([^<]*)<\/div>\s*<div[^>]*>([^<]*)<\/div>\s*<div[^>]*>([^<]*)<\/div>\s*<div[^>]*>([^<]*)<\/div>/gi, m;
  while ((m = re.exec(html))) {
    var a = '<x' + m[1] + '>';
    var o = { id: rbAttr_(a, 'data-id'), unit: rbTrim_(rbAttr_(a, 'data-unit')), name: rbTrim_(rbAttr_(a, 'data-name')), zh: rbTrim_(rbAttr_(a, 'data-zh')),
      cate: rbTrim_(rbDecode_(m[4])), note: rbTrim_(rbDecode_(m[5])) };
    o.kn = rbNorm_(o.name); o.kz = rbNorm_(o.zh); o.ku = rbNorm_(o.unit); o.kc = rbNorm_(o.cate);   /* 比對用 */
    out.push(o);
  }
  return out;
}
function rbStepPage_(sess, rid) {
  var r = rbAuthedGet_(sess, '/Steps/Create/' + encodeURIComponent(rid));
  if (r.code !== 200) throw rbErr_('食譜後台「新增步驟」頁打不開（HTTP ' + r.code + '）', 'rbnet');
  var f = rbFormHtml_(r.text, /^\/Steps\/Create\//i);
  if (!f) throw rbErr_('食譜後台「新增步驟」頁找不到表單（頁面可能改版）', 'rbpage');
  var miss = RB_STEP_FIELDS.filter(function (n) { return !rbHasField_(f, n); });
  if (miss.length) throw rbErr_('食譜後台「新增步驟」頁少了欄位 ' + miss.join('、') + '（頁面可能改版，這次先不送）', 'rbpage');
  var list = rbIngList_(r.text);
  if (list.length < 50) throw rbErr_('食譜後台「新增步驟」頁讀不到食材清單（只讀到 ' + list.length + ' 項，頁面可能改版）', 'rbpage');
  return { token: rbToken_(f), ings: list };
}
/* 食材比對：品名（後台顯示名或中文名）相同 → 單位相同且分區相同 → 單位相同；單位對不到＝不掛（回報後台有哪些單位） */
function rbMatch_(list, it) {
  var n = rbNorm_(it.name), u = rbNorm_(it.unit), c = rbNorm_(it.cate);
  var same = list.filter(function (x) { return x.kn === n || x.kz === n; });
  if (!same.length) return { miss: 'name' };
  var a = same.filter(function (x) { return x.ku === u && (!c || x.kc === c); });
  if (a.length) return { hit: a[0] };
  var b = same.filter(function (x) { return x.ku === u; });
  if (b.length) return { hit: b[0] };
  var us = []; same.forEach(function (x) { if (us.indexOf(x.unit) < 0) us.push(x.unit); });
  return { miss: 'unit', units: us };
}
/* 步驟清單（公開頁 /steps/recipeindex）：[{sid, ord, title}]，照後台順序 */
function rbStepIndex_(sess, rid) {
  var r = rbReq_(sess, 'get', '/steps/recipeindex/' + encodeURIComponent(rid));
  if (r.code !== 200) throw rbErr_('讀不到後台這支食譜的步驟清單（HTTP ' + r.code + '）', 'rbnet');
  var out = [], re = /<tr\b[^>]*\sdata="([^"]*)"[^>]*>\s*<td[^>]*>([\s\S]*?)<\/td>\s*<td[^>]*>([\s\S]*?)<\/td>/gi, m;
  while ((m = re.exec(r.text))) out.push({ sid: rbDecode_(m[1]), ord: rbText_(m[2]), title: rbText_(m[3]) });
  return out;
}
/* 用「備註」裡的匯入編號找後台食譜（後台清單的 Note 搜尋是伺服器端篩選） */
function rbFindByNote_(sess, key) {
  var r = rbAuthedGet_(sess, '/Recipes?Note=' + encodeURIComponent(key));
  if (r.code !== 200) throw rbErr_('讀不到後台食譜清單（HTTP ' + r.code + '）', 'rbnet');
  return rbListRows_(r.text);
}
function rbListRows_(html) {
  var out = [], re = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi, m;
  while ((m = re.exec(html))) {
    var row = m[1], id = rbAttr_((row.match(/<input\b[^>]*class="recipe-check"[^>]*>/i) || [''])[0], 'value');
    if (!id) continue;
    var tds = row.match(/<td\b[^>]*>[\s\S]*?<\/td>/gi) || [];
    out.push({ id: id, title: tds[1] ? rbText_(tds[1]) : '', lang: tds[5] ? rbText_(tds[5]) : '', status: tds[8] ? rbText_(tds[8]) : '' });
  }
  return out;
}

/* ---- 送出 ---- */
function rbLocId_(loc, re) { var m = String(loc || '').match(re); return m ? m[1] : ''; }
function rbFormErrors_(html) {
  var out = [], re = /<(?:span|div|li)\b[^>]*class="[^"]*(?:field-validation-error|validation-summary-errors|text-danger)[^"]*"[^>]*>([\s\S]*?)<\/(?:span|div|li)>/gi, m;
  while ((m = re.exec(html))) { var t = rbText_(m[1]); if (t && out.indexOf(t) < 0) out.push(t); }
  return out.slice(0, 5).join('；');
}
/* 建食譜（停用、不公開）：回後台食譜編號。r＝{title, note, itemList（類別編號，可空）, storeNames（分店類別名稱）, price, cost, prepHr, size, desc, preserve}
   登入逾時（什麼都還沒建）→ 丟 rbrelogin 給上層重登再來 */
function rbCreateRecipe_(sess, r) {
  var pg = rbAuthedGet_(sess, '/Recipes/Create');
  if (pg.code !== 200) throw rbErr_('食譜後台「新增食譜」頁打不開（HTTP ' + pg.code + '）', 'rbnet');
  var meta = rbCreateMeta_(pg.text);
  var storeIds = (r.storeNames || []).map(function (n) { var o = meta.stores.filter(function (x) { return x.name === n; })[0]; return o ? o.id : ''; }).filter(String);
  var cate = (r.itemList && meta.cates.some(function (c) { return c.id === r.itemList; })) ? r.itemList : '';   /* 類別不在台灣繁中清單裡 → 留空，由食譜負責夥伴選 */
  var payload = {
    '__RequestVerificationToken': meta.token,
    Title: r.title, Code: '', GroupId: meta.groupId, LanguageId: meta.langId, Note: r.note,
    ItemList: cate, StoreList: storeIds.join(','),
    RecipeUp: RB_RECIPE_UP, RecipeDown: RB_RECIPE_DOWN, MenuUp: '', MenuDown: '',   /* 空白＝平板打不開（2026-10-08） */
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
/* 建一個步驟。s＝{title, html, ings:[{id, amount, unit, cont}], media:Blob|null}。
   回 true；登入逾時（沒建）→ rbrelogin；其他錯誤（含連線中斷＝不確定有沒有建）→ 丟出，由上層讀步驟清單判斷，不重送 */
function rbCreateStep_(sess, rid, token, s) {
  var payload = {
    '__RequestVerificationToken': token, RecipeId: rid, OptionId: '', OptionItemId: '',
    StepTitle: s.title, Content: s.html, StopClock: 'false', Timer: ''
  };
  (s.ings || []).forEach(function (g, i) {
    payload['Ingredients[' + i + '].IngredSourceId'] = g.id;
    payload['Ingredients[' + i + '].Amount'] = String(g.amount);
    payload['Ingredients[' + i + '].Unit'] = g.unit;
    payload['Ingredients[' + i + '].Container'] = g.cont || '';
  });
  if (s.media) payload.Image = s.media;   /* 有 Blob → 自動用 multipart/form-data 送 */
  var res = rbReq_(sess, 'post', '/Steps/Create/' + encodeURIComponent(rid), { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受步驟「' + s.title + '」：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台新增步驟「' + s.title + '」回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
/* 抓公版圖片／影片（只限自家雲端空間 stepmedias；超過 5MB 不放） */
function rbFetchMedia_(url) {
  url = String(url || '');
  if (url.indexOf(RB_MEDIA_PREFIX) !== 0 || /[\s"'<>]/.test(url)) throw rbErr_('圖片／影片網址不是自家雲端空間：' + url.slice(0, 80), 'rbmedia');
  var path = url.slice(RB_MEDIA_PREFIX.length).split('?')[0];
  var ext = (path.match(/\.([a-z0-9]+)$/i) || [])[1];
  ext = ext ? ext.toLowerCase() : '';
  var types = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', mp4: 'video/mp4', mov: 'video/quicktime' };
  if (!types[ext]) throw rbErr_('不支援的圖片／影片格式：' + ext, 'rbmedia');
  var r = rbReq_(null, 'get', url.split('?')[0], { binary: true });
  if (r.code !== 200) throw rbErr_('抓不到公版圖片／影片（HTTP ' + r.code + '）', 'rbmedia');
  var b = r.resp.getBlob(), n = b.getBytes().length;
  if (n > RB_MEDIA_MAX) throw rbErr_('公版圖片／影片超過 5MB（' + (Math.round(n / 104857.6) / 10) + 'MB），後台不收', 'rbmedia');
  b.setContentType(types[ext]);
  b.setName('step.' + ext);
  return { blob: b, bytes: n, type: types[ext] };
}
/* 步驟內容安全處理：拿掉 script／iframe 等可執行標籤、on* 事件屬性、javascript: 連結；樣式（紅字、置中、字級）保留 */
function rbSanitize_(html) {
  return String(html || '')
    .replace(/<\s*(script|style|iframe|object|embed|form|template)\b[\s\S]*?<\s*\/\s*\1\s*>/gi, '')
    .replace(/<\s*\/?\s*(script|style|iframe|object|embed|link|meta|base|form|input|button|textarea|select|template)\b[^>]*>/gi, '')
    .replace(/\son[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/(href|src)\s*=\s*("|')\s*(javascript|vbscript):[^"']*\2/gi, '$1=$2#$2')
    .replace(/href\s*=\s*("|')\s*data:[^"']*\1/gi, 'href=$1#$1');
}

/* ===== 2026-10-08 ♻️ 覆蓋上傳（📤 新食譜上傳）：讀／改既有的食譜與步驟 =====
   經營者 10/08：「主廚改完再次上傳時直接覆蓋同一支，如果中途斷線，應該要能再次上傳直到成功為止」。
   後台頁面（10/08 讀真頁面）：
     ‧ /Recipes/Edit/{id}：multipart 表單；類別／分店類別是畫面上的 span.item[data-item]，送出前頁面腳本才把它們寫進 ItemList／StoreList（逗號）；
       InUse／Public 是 checkbox＋同名 hidden false；封面圖 TitleImage 等是 hidden（沒附新檔就照原值送回）。
     ‧ /Steps/Edit/{stepId}：RecipeId、StepId、Order、StepMedia（目前的圖片影片路徑，沒附新檔要原樣送回）、Image（新檔）、StepTitle、Content、
       StopClock、Timer、deletedIngred（要拿掉的既有食材列編號，逗號）、Ingredients[i].*（新加的食材）。
     ‧ /Steps/DeletePost：id＝步驟編號（頁面的「刪除」鈕就是送這個，沒有驗證碼）。 */
/* 表單現值：[[name, value], …]（照頁面順序；file 欄不帶；checkbox／radio 只帶勾起來的；select 帶選中的；同名欄位照實保留多個） */
function rbFormFields_(f) {
  var out = [], re = /<(input|textarea|select)\b[^>]*>/gi, m;
  while ((m = re.exec(f))) {
    var tag = m[0], kind = m[1].toLowerCase(), name = rbAttr_(tag, 'name');
    if (kind === 'input') {
      if (!name) continue;
      var type = (rbAttr_(tag, 'type') || 'text').toLowerCase();
      if (/^(file|submit|button|image|reset)$/.test(type)) continue;
      if ((type === 'checkbox' || type === 'radio') && !/\schecked\b/i.test(tag)) continue;
      out.push([name, (type === 'checkbox' || type === 'radio') ? (rbAttr_(tag, 'value') || 'on') : rbAttr_(tag, 'value')]);
    } else if (kind === 'textarea') {
      var end = f.indexOf('</textarea>', re.lastIndex);
      if (name) out.push([name, rbDecode_(f.slice(re.lastIndex, end < 0 ? re.lastIndex : end).replace(/^\r?\n/, ''))]);
      if (end >= 0) re.lastIndex = end;
    } else {
      var send = f.indexOf('</select>', re.lastIndex), body = f.slice(re.lastIndex, send < 0 ? re.lastIndex : send), ore = /<option\b[^>]*>/gi, om, first = null, sel = null;
      while ((om = ore.exec(body))) { var v = rbAttr_(om[0], 'value'); if (first === null) first = v; if (sel === null && /\sselected\b/i.test(om[0])) sel = v; }
      if (name && (sel !== null || first !== null)) out.push([name, sel !== null ? sel : first]);
      if (send >= 0) re.lastIndex = send;
    }
  }
  return out;
}
function rbFieldGet_(fields, name) { for (var i = 0; i < fields.length; i++) if (fields[i][0] === name) return fields[i][1]; return null; }
function rbEncode_(pairs) { return pairs.map(function (p) { return encodeURIComponent(p[0]) + '=' + encodeURIComponent(p[1] == null ? '' : String(p[1])); }).join('&'); }
function rbGone_(r) { return r.code === 404 || ((r.code === 302 || r.code === 301) && !rbIsLogin_(r)); }
var RB_EDIT_FIELDS = ['RecipeId', 'Title', 'GroupId', 'LanguageId', 'Note', 'ItemList', 'StoreList', 'RecipeUp', 'RecipeDown', 'Price', 'Cost', 'InUse', 'PrepHr', 'Size', 'Content', 'Preserve', 'Public'];
var RB_STEP_EDIT_FIELDS = ['RecipeId', 'StepId', 'Order', 'StepMedia', 'Image', 'StepTitle', 'Content', 'StopClock', 'Timer', 'deletedIngred'];
/* 「編輯食譜」頁：表單現值＋目前的類別／分店類別 */
function rbRecipeEditPage_(sess, bid) {
  var r = rbAuthedGet_(sess, '/Recipes/Edit/' + encodeURIComponent(bid));
  if (rbGone_(r)) throw rbErr_('後台找不到這支食譜（可能已經被刪掉）', 'rbgone');
  if (r.code !== 200) throw rbErr_('食譜後台「編輯食譜」頁打不開（HTTP ' + r.code + '）', 'rbnet');
  var f = rbFormHtml_(r.text, /^\/Recipes\/Edit\//i);
  if (!f) throw rbErr_('食譜後台「編輯食譜」頁找不到表單（頁面可能改版）', 'rbpage');
  var miss = RB_EDIT_FIELDS.filter(function (n) { return !rbHasField_(f, n); });
  if (miss.length) throw rbErr_('食譜後台「編輯食譜」頁少了欄位 ' + miss.join('、') + '（頁面可能改版，這次先不送）', 'rbpage');
  var fields = rbFormFields_(f);
  if (rbFieldGet_(fields, 'RecipeId') !== bid) throw rbErr_('食譜後台「編輯食譜」頁的食譜編號對不上（頁面可能改版）', 'rbpage');
  var t = r.text, a = t.indexOf('id="cateG"'), b = t.indexOf('id="storeType"'), c = t.indexOf('name="RecipeUp"');
  if (a < 0 || b < a || c < b) throw rbErr_('食譜後台「編輯食譜」頁讀不到類別／分店類別（頁面可能改版）', 'rbpage');
  var items = function (from, to) {
    var seg = t.slice(from, to), out = [], re = /<span\s+class=['"]item['"]\s+data-item=['"]([^'"]+)['"]\s*>([^<]*)/gi, m;
    while ((m = re.exec(seg))) out.push({ id: m[1], name: rbTrim_(rbDecode_(m[2])) });
    return out;
  };
  return { token: rbToken_(f), fields: fields, cats: items(a, b), stores: items(b, c), inUse: rbFieldGet_(fields, 'InUse') === 'true', pub: rbFieldGet_(fields, 'Public') === 'true' };
}
/* 改基本資料：其他欄位（啟用、公開、封面圖、流水碼…）照頁面現值原樣送回。r＝{title, note, cats（類別編號）, storeIds, price, cost, prepHr, size, desc, preserve} */
function rbUpdateRecipe_(sess, bid, pg, r) {
  var set = { Title: r.title, Note: r.note, ItemList: (r.cats || []).join(','), StoreList: (r.storeIds || []).join(','), Price: String(r.price), Cost: String(r.cost),
    PrepHr: String(r.prepHr || ''), Size: r.size || '', Content: r.desc || '', Preserve: r.preserve || '' };
  var done = {}, pairs = pg.fields.map(function (p) {
    var k = p[0], v = p[1];
    if (Object.prototype.hasOwnProperty.call(set, k)) { if (done[k]) return null; done[k] = 1; v = set[k]; }
    else if (k === 'RecipeUp' && !v) v = RB_RECIPE_UP;          /* 空白＝平板打不開（2026-10-08）：補上跟新建一樣的期間 */
    else if (k === 'RecipeDown' && !v) v = RB_RECIPE_DOWN;
    return [k, v];
  }).filter(Boolean);
  var res = rbReq_(sess, 'post', '/Recipes/Edit/' + encodeURIComponent(bid), { payload: rbEncode_(pairs) });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受修改食譜：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台修改食譜回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
/* 「編輯步驟」頁：現值＋目前掛的食材列編號 */
function rbStepEditParse_(r, sid) {
  if (rbGone_(r)) throw rbErr_('後台找不到這個步驟（可能已經被刪掉）', 'rbgone');
  if (r.code !== 200) throw rbErr_('食譜後台「編輯步驟」頁打不開（HTTP ' + r.code + '）', 'rbnet');
  var f = rbFormHtml_(r.text, /^\/Steps\/Edit\//i);
  if (!f) throw rbErr_('食譜後台「編輯步驟」頁找不到表單（頁面可能改版）', 'rbpage');
  var miss = RB_STEP_EDIT_FIELDS.filter(function (n) { return !rbHasField_(f, n); });
  if (miss.length) throw rbErr_('食譜後台「編輯步驟」頁少了欄位 ' + miss.join('、') + '（頁面可能改版，這次先不送）', 'rbpage');
  var fields = rbFormFields_(f), g = function (n) { var v = rbFieldGet_(fields, n); return v == null ? '' : v; };
  if (g('StepId') !== sid) throw rbErr_('食譜後台「編輯步驟」頁的步驟編號對不上（頁面可能改版）', 'rbpage');
  var ings = [], re = /<a\b[^>]*class="[^"]*deleteExistIngred[^"]*"[^>]*>/gi, m;
  while ((m = re.exec(r.text))) { var id = rbAttr_(m[0], 'data-ingredid'); if (id) ings.push(id); }
  var media = g('StepMedia');
  return { token: rbToken_(f), rid: g('RecipeId'), sid: sid, order: g('Order'), media: media, title: g('StepTitle'), stop: g('StopClock') === 'true', timer: g('Timer'), ings: ings,
    mediaUrl: media ? RB_MEDIA_PREFIX + media.split('?')[0] : '' };
}
function rbStepEditPage_(sess, sid) { return rbStepEditParse_(rbAuthedGet_(sess, '/Steps/Edit/' + encodeURIComponent(sid)), sid); }
/* 一次讀很多個「編輯步驟」頁（UrlFetchApp.fetchAll，同時送）；逾時被導回登入頁 → 重登一次再讀 */
function rbStepEditPages_(sess, sids) {
  var go = function () {
    var reqs = sids.map(function (sid) { return { url: RB_BASE + '/Steps/Edit/' + encodeURIComponent(sid), method: 'get', muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA, Cookie: rbJarStr_(sess.jar) } }; });
    var out = [];
    for (var i = 0; i < reqs.length; i += 20) {
      var rs = UrlFetchApp.fetchAll(reqs.slice(i, i + 20));
      rs.forEach(function (x) { var h = x.getAllHeaders(); out.push({ code: x.getResponseCode(), loc: String(h.Location || h.location || ''), text: x.getContentText() }); });
    }
    return out;
  };
  var rs = go();
  if (rs.some(rbIsLogin_)) { rbDropSession_(); sess.jar = rbFreshSession_().jar; rs = go(); }
  return rs.map(function (r, i) { return rbStepEditParse_(r, sids[i]); });
}
/* 改一個步驟（就地改：標題、內容、食材全部換成新的；圖片影片：有新檔換新檔、keep＝留原本的、都沒有＝拿掉；計時器照原本的）。
   s＝{title, html, ings:[{id, amount, unit, cont}], media:Blob|null, keep}。重送一樣的內容結果相同（可安全重做） */
function rbEditStep_(sess, pg, s) {
  var payload = {
    '__RequestVerificationToken': pg.token, RecipeId: pg.rid, StepId: pg.sid, Order: pg.order,
    StepMedia: (s.media || s.keep) ? pg.media : '', StepTitle: s.title, Content: s.html,
    StopClock: pg.stop ? 'true' : 'false', Timer: pg.timer || '', deletedIngred: pg.ings.length ? pg.ings.join(',') + ',' : ''
  };
  (s.ings || []).forEach(function (g, i) {
    payload['Ingredients[' + i + '].IngredSourceId'] = g.id;
    payload['Ingredients[' + i + '].Amount'] = String(g.amount);
    payload['Ingredients[' + i + '].Unit'] = g.unit;
    payload['Ingredients[' + i + '].Container'] = g.cont || '';
  });
  if (s.media) payload.Image = s.media;
  var res = rbReq_(sess, 'post', '/Steps/Edit/' + encodeURIComponent(pg.sid), { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受修改步驟「' + s.title + '」：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台修改步驟「' + s.title + '」回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
/* 刪一個步驟（只給覆蓋上傳用：新檔步驟比後台少時，刪掉多出來的那幾步） */
function rbDeleteStep_(sess, sid) {
  var res = rbReq_(sess, 'post', '/Steps/DeletePost', { payload: { id: sid } });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code >= 200 && res.code < 400) return true;
  throw rbErr_('食譜後台刪除步驟回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
