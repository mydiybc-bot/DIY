/* ================= 🀄 平板外語食譜：原文改用「中文食譜」（2026-11-24 經營者） =================
   經營者 2026-11-24：「接下來，我想用中文食譜做基準，進行外語食譜的翻譯更新」→ 決定：
     ①品名譯文參考食譜系統「食材管理」的名稱對照表（/IngredSources/NameList）
     ②平板英文和其他語言全部改用中文直接翻（英文也是翻的，不再直接用後台英文食譜）
     ③92 支（台灣區、繁體中文、啟用中、客人用，不含裝飾蛋糕；同 i18nZCands_）
     ④自己人募集中／裝飾物加購／六吋透明盒包裝這 3 種台灣限定頁也翻 ⑤先救急（11/24 已做）⑥不等步驟簡化。
   2026-11-23 裁定：外語食譜不用 API → 這裡只有「讀原文、換圖影、詞典、檢查」；翻譯一律由 Claude Code 投遞（tools/i18n_claude，i18nPull／i18nPut）。
   做法（新舊並行，平板不開天窗）：
   - 候選中文食譜記在 i18n_zsrc（id, name, old 對應的舊英文原文, on, steps, pair 自動／手動, at）。
     on＝Y 的才收進原文 i18n_src（id＝中文食譜 id、JSON 帶 base:'zh'），才進待翻清單 → 分批打開（動作 i18nZSet）。
   - 中文原文的語言＝en＋日韓越印泰（i18nZLangs_）；詞典另一張 i18n_zitems（中文名當索引：同一個英文名 Whipping cream
     在後台同時是克林姆、打發克林姆，用英文當索引會翻錯）。
   - 有舊英文原文（old）的：中文版 6 種語言都能顯示之前，平板照舊列舊的那支；6 種都好了 → 平板改列中文版，
     舊的不再列、也不再翻（i18nZRetired_）。沒有舊英文版的：哪個語言放行就在那個語言出現（同一般規則）。
   - 有中文字的共用圖影（原文更新時就換好存進 i18n_src）：烤熟測試圖、影片重置／攪拌球／打發測試湯匙／量匙怎麼舀插圖 → 英文版；
     教學影片（中文）記 vtut，平板組內容時換成熊主廚各語言版（英文＝英文共用教學影片）。
   ⚠ 只讀後台；不寫後台。只寫自己的分頁（i18n_zsrc、i18n_zitems，其他照 i18n.gs）。 */
var I18N_ZSRC_HEAD = ['id', 'name', 'old', 'on', 'steps', 'pair', 'at'];
var I18N_ZITEM_HEAD = ['kind', 'zh', 'en', 'ja', 'ko', 'vi', 'id', 'th', 'how', 'at'];
var I18N_ZSKIP = /訓練&材料|English Menu|台灣.*備份|備用/;   /* 只掛這些類別的不是客人用的 */
var I18N_ZINCLUDE = { '12de8b02-128f-410c-90fd-4ca0e04f6d8c': '鳳梨酥' };   /* 類別不符也要收的（2026-11-24 經營者：「這個鳳梨酥要收進來」——中文版只掛「訓練&材料」，但平板一直有英文版 Pineapple Cake） */
var I18N_ZCHEF = 'https://cdn.jsdelivr.net/gh/mydiybc-bot/DIY@f9e57f9cf05bd011524e2b4f21c6b40f284cfab4/media/tablet-intro/chef-bear-';
var I18N_EN_TUT = 'https://diybcstorage.blob.core.windows.net/stepmedias/248973a1-3163-434a-83f0-1ced05c9d8fd/9866e7c2-138a-437e-88ec-67bfca5e0c8c.mp4?v=251031030624';   /* 英文共用教學影片（recipe-english-copy skill「二、圖片影片」） */
var I18N_ZMD5 = {   /* 主圖／主影片（認檔案指紋 Content-MD5；每支食譜網址不同、檔案同一個） */
  'zwU1ifRZOqYg5KT5Bifb/Q==': { tut: 1 },   /* 中文教學影片 */
  'O51ZCnNhWpVOX1alTp0kTQ==': { img: 'https://diybcstorage.blob.core.windows.net/stepmedias/941ffc04-31c8-4c12-a7aa-6cc2745be443/9866e7c2-138a-437e-88ec-67bfca5e0c8c.jpg?v=261117063321' }   /* 烤熟測試竹籤圖（中文 2026-11-17 版）→ 英文版 */
};
var I18N_ZIMG = [   /* 說明裡的插圖（認檔名）：圖上有中文 → 英文版（同 recipe-english-copy 工具 IMG 表） */
  ['影片重置(中文)', 'https://diybcstorage.blob.core.windows.net/stepmedias/20241113101118影片重置(英文).jpg'],
  ['20250109100107攪拌球', 'https://lh3.googleusercontent.com/0_v4aS4xYNj_jXeLSbTAoD82hgWnXZzY2NabqVBF_fSmJtZQxPEoMRBdmz0IsDaierenKz2iq3hJXVaA8iMH27FlcdeiawBTjYkOmGU5XO1M3vndixl78IwdFPJvpBDT5SY81RI6rQ=w400'],
  ['20250109100148打發測試湯匙', 'https://lh3.googleusercontent.com/KQ4tNH0bkhrTBa1XWEER4piEg6DUJ6_Uf-5U_TWzuBYjuvv276AK9gkG0OCyUDfjcNzjbHgvrrgmo6UsDITGURMYXJ6aQ9TEswmuZjF2rnMFcmNbaxp1Er0BATaBkH3iaMBQa0zdTw=w400'],
  ['量匙怎麼舀_中文', 'https://diybcstorage.blob.core.windows.net/stepmedias/20261113141101量匙怎麼舀_英文.webp']
];
/* 中文數字：譯文多出來的阿拉伯數字，原文要有對應的中文字（一、兩、半…；日韓泰常寫成「1本」「2枚目」） */
var I18N_ZNUMW = { '1': /[一單每半首]/g, '2': /[二兩雙半對]/g, '3': /三/g, '4': /四/g, '5': /五/g, '6': /六/g, '7': /七/g, '8': /八/g, '9': /九/g, '10': /十/g };

function i18nZLangs_() { return ['en'].concat(I18N_LANGS); }

/* ---------------- 候選與開關 ---------------- */
var I18N_ZM = null;   /* 同一次執行只讀一次；i18nZSet 寫完會清掉 */
function i18nZMap_() {   /* i18n_zsrc 打開（on＝Y）的 → {id: {name, old, on, man}}＝目前當中文原文的食譜（關掉的照一般規則：中文試看那支只當試看） */
  if (I18N_ZM) return I18N_ZM;
  var m = {}, sh = ss_().getSheetByName('i18n_zsrc');
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, I18N_ZSRC_HEAD.length).getDisplayValues().forEach(function (r) {
    if (r[0] && r[3] === 'Y') m[String(r[0]).trim().toLowerCase()] = { name: r[1], old: String(r[2] || '').trim().toLowerCase(), on: true, man: r[5] === '手動' };
  });
  return (I18N_ZM = m);
}
function i18nIsZ_(id, ZM) { return !!(ZM || i18nZMap_())[String(id || '').toLowerCase()]; }
/* 候選（純讀 recipe_lineage）：台灣區、繁體中文、啟用中、有客人用的類別、不含裝飾蛋糕（11/24 評估時的 92 支同一個篩法）。
   配對舊英文原文：英文名「英文｜中文名」的中文名＝這支的名稱（全形半形括號、空白不計）；同名好幾支（例：全素／可做全素）挑步數最接近的。 */
function i18nZCands_() {
  var sh = ss_().getSheetByName(LG_TAB);
  if (!sh || sh.getLastRow() < 2) throw fail_('食譜演化清單（recipe_lineage）是空的', 'nolist');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues(), H = v[0], ix = {};
  H.forEach(function (h, i) { ix[h] = i; });
  var rows = v.slice(1).map(function (r) {
    return { id: String(r[ix.recipe_id] || '').trim().toLowerCase(), name: rbTrim_(r[ix.name]), lang: r[ix.lang], region: r[ix.region], status: r[ix.status], cats: i18nCatSplit_(r[ix.cat]), upd: r[ix.upd], steps: Number(r[ix.steps]) || 0 };
  }).filter(function (r) { return r.id; });
  var zh = rows.filter(function (r) {
    if (I18N_ZINCLUDE[r.id] && r.status === 'Active') return true;
    return r.region === '台灣' && r.lang === '繁體中文' && r.status === 'Active' && r.cats.length
      && !r.cats.every(function (c) { return I18N_ZSKIP.test(c); })
      && !(r.cats.some(function (c) { return /裝飾蛋糕/.test(c); }) || /裝飾蛋糕/.test(r.name));
  });
  var bare = function (n) { return i18nCatNameKey_(String(n || '').replace(/[(（][^)）]*[)）]/g, '')); };   /* 去掉括號備註：(可做全素)、（製作時間約3小時…) */
  var by = {}, by2 = {}; zh.forEach(function (z) { var k = i18nCatNameKey_(z.name), k2 = bare(z.name); (by[k] = by[k] || []).push(z); (by2[k2] = by2[k2] || []).push(z); });
  var taken = {}, ens = rows.filter(function (r) { return r.region === '台灣' && r.lang === '英文' && r.status === 'Active'; });
  [false, true].forEach(function (loose) {   /* 先找名稱完全一樣的，配不到再用去掉括號備註的名稱配 */
  ens.forEach(function (e) {
    if (e.paired) return;
    var p = String(e.name).split('｜'); if (p.length < 2) return;
    var zn = p.slice(1).join('｜'), list = ((loose ? by2[bare(zn)] : by[i18nCatNameKey_(zn)]) || []).filter(function (z) { return !taken[z.id]; });
    if (!list.length) return;
    e.paired = 1;
    list.sort(function (a, b) { return Math.abs(a.steps - e.steps) - Math.abs(b.steps - e.steps); });
    list[0].old = e.id; list[0].oldName = e.name; taken[list[0].id] = 1;
  });
  });
  return zh;
}
/* 動作 i18nZSet（投遞通行碼 wkey）：
   op list（預設）＝同步候選到 i18n_zsrc（新的記 on＝N；自動配對的 old 跟著更新，手動的不動）後回清單；
   op on／off＋ids＝打開／關掉（打開才收原文、進待翻）；op old＋id＋old＝手動指定對應的舊英文原文（old 空白＝沒有）。 */
function i18nZSet_(p) {
  i18nWriteKey_(p);
  var op = String(p.op || 'list'), t = now_(), C = i18nZCands_(), cm = {};
  C.forEach(function (c) { cm[c.id] = c; });
  var ids = (Array.isArray(p.ids) ? p.ids : p.id ? [p.id] : []).map(function (x) { return String(x || '').trim().toLowerCase(); }).filter(String);
  if (op === 'on') ids.forEach(function (id) { if (!cm[id]) throw fail_('不是候選的中文食譜（台灣區、繁中、啟用中、客人用、非裝飾蛋糕）：' + id, 'bad'); });
  if ((op === 'on' || op === 'off' || op === 'old') && !ids.length) throw fail_('沒有指定食譜', 'bad');
  var out = i18nLock_(function () {
    var sh = i18nSheet_('i18n_zsrc', I18N_ZSRC_HEAD), n = sh.getLastRow() - 1, have = {}, order = [];
    if (n > 0) sh.getRange(2, 1, n, I18N_ZSRC_HEAD.length).getDisplayValues().forEach(function (r) {
      var o = {}; I18N_ZSRC_HEAD.forEach(function (h, j) { o[h] = r[j]; }); o.id = String(o.id).trim().toLowerCase();
      if (o.id && !have[o.id]) { have[o.id] = o; order.push(o.id); }
    });
    C.forEach(function (c) {
      var o = have[c.id];
      if (!o) { o = have[c.id] = { id: c.id, on: 'N', pair: c.old ? '自動' : '', old: c.old || '', at: t }; order.push(c.id); }
      o.name = c.name; o.steps = String(c.steps);
      if (o.pair !== '手動') { o.old = c.old || ''; o.pair = c.old ? '自動' : ''; }
    });
    ids.forEach(function (id) {
      var o = have[id]; if (!o) throw fail_('i18n_zsrc 沒有這支：' + id, 'bad');
      if (op === 'on' || op === 'off') { o.on = op === 'on' ? 'Y' : 'N'; o.at = t; }
      if (op === 'old') { o.old = String(p.old || '').trim().toLowerCase(); o.pair = '手動'; o.at = t; }
    });
    var vals = order.map(function (id) { var o = have[id]; return I18N_ZSRC_HEAD.map(function (h) { return o[h] == null ? '' : String(o[h]); }); });
    if (vals.length) sh.getRange(2, 1, vals.length, I18N_ZSRC_HEAD.length).setNumberFormat('@').setValues(vals);
    return order.map(function (id) { return have[id]; });
  });
  I18N_ZM = null;
  if (op === 'off') i18nLock_(function () {   /* 關掉 → 原文立刻下架（不等明天重抓；不然會被當成英文原文） */
    var sm = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows;
    ids.forEach(function (id) { if (sm.some(function (r) { return r.id === id && r.state !== 'gone'; }) && !i18nZh_()[id]) i18nUpsert_('i18n_src', I18N_SRC_HEAD, 'id', { id: id, state: 'gone', at: now_() }, null); });
  });
  if (op !== 'list') i18nLog_('Claude Code', 'zset-' + op, ids.join(','), '', 'ok', op === 'old' ? '對應舊英文原文 ' + (p.old || '（無）') : '');
  var en = {}; C.forEach(function (c) { if (c.old) en[c.old] = c.oldName; });
  var list = out.map(function (o) { var c = cm[o.id]; return { id: o.id, name: o.name, cand: !!c, on: o.on === 'Y', old: o.old, old_name: en[o.old] || '', pair: o.pair, steps: Number(o.steps) || 0, upd: c ? c.upd : '', cats: c ? c.cats : [] }; });
  return { data: { op: op, cands: C.length, on: list.filter(function (x) { return x.on && x.cand; }).length, paired: list.filter(function (x) { return x.cand && x.old; }).length, items: list } };
}
ACTIONS.i18nZSet = function (p, a) { return i18nZSet_(p, a); };
NOAUTH.i18nZSet = 1; SELFLOCK.i18nZSet = 1;

/* 原文清單裡的中文食譜（i18nSources_ 用）：候選 ∩ on＝Y */
function i18nZSources_() {
  var ZM = i18nZMap_(), on = Object.keys(ZM).filter(function (id) { return ZM[id].on; });
  if (!on.length) return [];
  return i18nZCands_().filter(function (c) { return ZM[c.id] && ZM[c.id].on; })
    .map(function (c) { return { id: c.id, name: c.name, upd: c.upd, base: 'zh', old: ZM[c.id].old || '' }; });
}
/* 舊英文原文退場：中文版 6 種語言都能顯示（放行過）→ {舊 id: 中文 id}。rc＝{key: i18n_rcp 列}；okIds＝原文 ok 的中文 id（沒給就不檢查） */
function i18nZRetired_(ZM, rc, okIds) {
  var out = {};
  Object.keys(ZM).forEach(function (id) {
    var z = ZM[id]; if (!z.on || !z.old || (okIds && !okIds[id])) return;
    if (i18nZLangs_().every(function (l) { return i18nShowable_(rc[id + '|' + l]); })) out[z.old] = id;
  });
  return out;
}

/* ---------------- 原文：換有中文字的圖影 ---------------- */
function i18nZMd5_(urls) {   /* 檔案指紋（Content-MD5）：只問教學影片、烤熟測試兩種頁的主圖影；快取 6 小時 */
  var c = CacheService.getScriptCache(), md5 = {}, ask = [];
  urls.forEach(function (u) { var h = c.get('i18n:md5:' + i18nHash_(u)); if (h != null) md5[u] = h; else if (ask.indexOf(u) < 0) ask.push(u); });
  for (var i = 0; i < ask.length; i += 20) {
    var part = ask.slice(i, i + 20), rs = [];
    try { rs = UrlFetchApp.fetchAll(part.map(function (u) { return { url: u, headers: { Range: 'bytes=0-0' }, muteHttpExceptions: true }; })); } catch (e) { rs = []; }
    part.forEach(function (u, j) {
      var r = rs[j]; if (!r) return;
      var code = r.getResponseCode(), h = r.getHeaders(), m = '';
      Object.keys(h).forEach(function (k) { var lk = k.toLowerCase(); if (lk === 'x-ms-blob-content-md5' || lk === 'content-md5') m = m || String(h[k]); });
      md5[u] = m;
      if (code === 200 || code === 206) c.put('i18n:md5:' + i18nHash_(u), m, 21600);
    });
  }
  return md5;
}
function i18nZMediaUrls_(src) {
  var u = [];
  src.slides.forEach(function (s, k) { if (k === 0 || /教學影片|烤熟測試/.test(s.t || '')) [s.video, s.img].forEach(function (x) { if (x && u.indexOf(x) < 0) u.push(x); }); });
  return u;
}
/* 換圖影（原文更新時做一次，存進 i18n_src）：回傳要提醒的事（例：第 1 頁教學影片不是認得的中文版） */
function i18nZPrep_(src, md5) {
  var warn = [];
  src.base = 'zh';
  src.slides.forEach(function (s, k) {
    var v = s.video && I18N_ZMD5[md5[s.video]];
    if (v && v.tut) s.vtut = 1;
    else if (k === 0 && s.video && /教學影片/.test(s.t || '')) warn.push('第1頁教學影片不是認得的中文版（指紋 ' + (md5[s.video] || '讀不到') + '），照原影片顯示');
    var g = s.img && I18N_ZMD5[md5[s.img]];
    if (g && g.img) { s.img_zh = s.img; s.img = g.img; }
    (s.imgs || []).forEach(function (u, j) {
      var d = u; try { d = decodeURIComponent(u); } catch (e) { }
      I18N_ZIMG.forEach(function (x) { if (d.indexOf(x[0]) >= 0) { if (!s.imgs_zh) s.imgs_zh = s.imgs.slice(); s.imgs[j] = x[1]; } });
    });
  });
  src.zwarn = warn;
  return warn;
}
/* 台灣限定頁的海報（整張中文）→ 各語言版（2026-11-24 經營者：「改成外語版海報」）：保留原圖照片插畫、只換字，
   做法與文字在 DIY tools/i18n_poster/（make.py＋text.json，改字重跑後放 static/img/i18n-poster/、改這裡的 v）；平板組內容時依語言換 */
var I18N_ZPOSTER = { base: 'https://diybc-dash.onrender.com/static/img/i18n-poster/', v: '20261124', list: [
  ['20240903150925裝飾物加購_食譜系統', 'addon'],   /* 裝飾物加購（價目表） */
  ['BcRMCCKYfxbRqBUEwernOZcosxfl3y3Kot3l9MduvN7CnDYcEZIcp4', 'member']   /* 自己人募集中（lh3 圖） */
] };
function i18nZPoster_(u, lang) {
  if (i18nZLangs_().indexOf(lang) < 0) return u;
  var d = u; try { d = decodeURIComponent(u); } catch (e) { }
  for (var i = 0; i < I18N_ZPOSTER.list.length; i++) if (d.indexOf(I18N_ZPOSTER.list[i][0]) >= 0) return I18N_ZPOSTER.base + I18N_ZPOSTER.list[i][1] + '_' + lang + '.jpg?v=' + I18N_ZPOSTER.v;
  return u;
}
function i18nZTut_(lang, orig) { return lang === 'zh' ? orig : lang === 'en' ? I18N_EN_TUT : I18N_ZCHEF + lang + '.mp4'; }

/* ---------------- 詞典（中文名當索引） ---------------- */
function i18nZDictRows_() {
  var m = i18nMeta_('i18n_zitems', I18N_ZITEM_HEAD), d = {};
  m.rows.forEach(function (r) { d[r.kind + '|' + String(r.zh).toLowerCase()] = r; });
  return d;
}
function i18nDictFor_(src) { return src && src.base === 'zh' ? i18nZDictRows_() : i18nDictRows_(); }
function i18nZKey_(s) { return rbTrim_(String(s || '')).replace(/^【[^】]*】\s*/, '').replace(/^[A-Z]{1,2}\d{2}\s+/, '').replace(/\s*[(（]已?下架[)）]/g, '').toLowerCase(); }
/* 名稱對照表（/IngredSources/NameList）：中文名 → {en: [後台英文名…], vi: 越南文}；給譯者參考（英文名可能錯，中文才是準） */
function i18nZNameMap_() {
  var d = trDict_({}).data, langs = d.langs || [], ie = -1, iv = -1, iz = -1, m = {};
  langs.forEach(function (l, i) { var s = String(l).trim(); if (ie < 0 && /^(英文|English)$/i.test(s)) ie = i; if (iv < 0 && /Tiếng Việt|越南文/.test(s) && !/中文/.test(s)) iv = i; if (iz < 0 && /^繁體中文/.test(s)) iz = i; });
  if (iz < 0) iz = 0;
  (d.rows || []).forEach(function (r) {
    var zh = i18nZKey_(r[3 + iz]); if (!zh) return;
    var o = m[zh] = m[zh] || { en: [], vi: '' };
    var en = ie >= 0 ? rbTrim_(r[3 + ie]).replace(/^【[^】]*】\s*/, '').replace(/^[A-Z]{1,2}\d{2}\s+/, '').replace(/\s*[(（]已?下架[)）]/g, '') : '';   /* 去掉品項編號（A07）、下架字樣，同中文 */
    if (en && o.en.indexOf(en) < 0) o.en.push(en);
    if (!o.vi && iv >= 0) o.vi = rbTrim_(r[3 + iv]).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]/g, '');
  });
  return m;
}
/* 給譯者的提示：後台英文名、越南文後台名、舊詞典（英文索引）的譯名——只在英文名只對到這一個中文品項時才給（克林姆／打發克林姆共用 Whipping cream 就不給）
   EZ＝i18nZEnZh_(NM)（先算好，整批用） */
function i18nZEnZh_(NM) {
  var enZh = {};   /* 後台英文名 → 有幾個不同中文品項用它 */
  Object.keys(NM).forEach(function (zh) { NM[zh].en.forEach(function (e) { var k = e.toLowerCase(); (enZh[k] = enZh[k] || {})[zh] = 1; }); });
  return enZh;
}
function i18nZHint_(zh, lang, NM, EZ, OD) {
  var o = {}, h = NM[i18nZKey_(zh)];
  if (!h) return o;
  if (h.en.length) o.backend_en = h.en.join(' / ');
  if (h.vi) o.backend_vi = h.vi;
  /* 2026-11-24 經營者：「其他的這一些名詞，也還是按照後臺的英文名進行翻譯」→ 品名一律照後台英文名（第一個），不管有沒有歧義；
     全部食譜上線後再逐一檢查中文名跟後台英文名配不配，不合適的提建議給經營者確認後再改 */
  if (h.en.length) {
    var c = lang === 'en' ? h.en[0] : i18nLook_(OD, 'item', h.en[0], lang);
    if (c) o.cand = c;
    if (h.en.length > 1 || Object.keys(EZ[h.en[0].toLowerCase()] || {}).length > 1) o.note = '後台英文名跟別的中文品項共用或不只一個（照後台英文名翻，之後統一檢討）';
  }
  return o;
}
/* 到店須知頁的在地品名（客人在那頁看到的名字，翻譯要一致）：中文名對；對不到用舊英文原文的名字對 */
function i18nZLocal_(src, lang, oldName) {
  var g = i18nGuideNames_(), k = i18nNameKey_(src.name), k2 = i18nNameKey_(String(src.name || '').replace(/\s*[(（][^)）]*[)）]\s*$/, ''));
  var hit = g['z:' + k] || g['z:' + k2] || (oldName ? g[i18nNameKey_(oldName)] : null) || {};
  return hit[lang] || '';
}

/* ---------------- 檢查（中文原文） ---------------- */
function i18nZNumDiff_(sh, th, okx) {   /* 同 i18nNumDiff_，但原文是中文：多出來的數字要原文有對應的中文數字；8分滿 → 8/10 */
  var a = i18nNums_(sh), b = i18nNums_(th), left = b.slice(), miss = [];
  a.forEach(function (x) { var j = left.indexOf(x); if (j < 0) miss.push(x); else left.splice(j, 1); });
  if (miss.length) return '原文的數字 ' + miss.join(',') + ' 譯文沒有（原文 ' + (a.join(',') || '無') + '／譯文 ' + (b.join(',') || '無') + '）';
  var pt = i18nPlain_(sh), extra = [], room = {};
  if (/分滿/.test(pt)) room['10'] = (pt.match(/分滿/g) || []).length + (pt.match(/十/g) || []).length;
  left.forEach(function (x) {
    if (okx && okx[x]) return;
    if (room[x] == null) room[x] = I18N_ZNUMW[x] ? (pt.match(I18N_ZNUMW[x]) || []).length : 0;
    if (room[x] > 0) room[x]--; else extra.push(x);
  });
  return extra.length ? '譯文多了數字 ' + extra.join(',') + '（原文沒有對應的字）' : '';
}
function i18nZHan_(s, lang) { return lang !== 'ja' && /[㐀-鿿]/.test(String(s || '')); }   /* 日文本來就有漢字，不檢查 */

/* ---------------- 給譯者的規則（中文原文） ---------------- */
var I18N_ZTERMS = [   /* 內文固定用語（2026-11-18～19 試做踩到的；11/24 補攪拌球、巧克力奶油、冷凍庫） */
  '自己做 / 自己做烘焙聚樂部 = "DIYBC" (Latin letters, never Chinese). 小教室 / 自己做小教室 / DIY小教室 = "DIY Tip" (exactly in English, e.g. "※DIY Tip:").',
  '自己人 / 自己人募集中 = the store\'s membership program ("DIYBC members"); 裝飾物加購 = add-on decorations you can buy; 透明盒 / 六吋透明盒包裝 = clear (transparent) cake box / packing in a 6-inch clear box.',
  'Ingredient, tool, mold, box and area names: ALWAYS use the glossary name exactly, in the steps too — the glossary follows the store\'s backend English names (2026-11-24 經營者：「其他的這一些名詞，也還是按照後臺的英文名進行翻譯」), even when the Chinese would suggest another word.',
  '克林姆 = the store\'s whipping cream: always use the glossary name (English "Whipping cream"; ja 生クリーム, ko 생크림, vi Kem tươi, id Whipping cream, th วิปปิ้งครีม). It is NOT custard — never write custard / カスタード / 커스터드 / kem custard / krim custard / คัสตาร์ด for 克林姆 (2026-11-24 經營者：「不是卡仕達醬，沿用後台的英文名稱」). 動植鮮 / 動植鮮奶油 = whipped topping cream (dairy + vegetable fat); 巧克力奶油 = chocolate buttercream.',
  '麵糊 = batter (the mixed liquid dough); never the same word as flour / cake flour (低筋麵粉). 蛋白霜 = meringue. 蛋黃液 = liquid egg yolk (an ingredient, not egg wash).',
  '攪拌機 = electric stand mixer (not a hand mixer); 攪拌球 = the wire whisk attachment of the mixer; 手持攪拌機 = hand mixer; 刮刀 = rubber spatula; 抹刀 = frosting (palette) spatula; 擠花袋 = piping bag; 花嘴 = piping tip; 剪花嘴 = cut-tip piping bag.',
  '冷凍庫 = freezer; 冷藏 / 冰箱 = refrigerator (chill); 冷凍 = freeze. 平匙 = a level spoonful; 半平匙 = half a spoonful. 打發 = whip; 拌勻 = mix well; 過篩 = sift; 取放區 / 取料區 = the pick-up area in the store; 計時器 = timer.',
  '先不… = do not … yet (e.g. 先不倒入全部食材 = do not pour in all the ingredients yet). 8分滿 = filled to 8/10 (keep the 8). 剩下 3 片 = the remaining 3 pieces.',
  'Piping shapes written as Chinese characters: 「米」字 = a ✳ star shape (lines radiating from the center; do not add a number); 「口」字 = a square outline □. Describe the shape and add the symbol; never leave the Chinese character.',
  '10元硬幣 = a Taiwan NT$10 coin (not US dollars). Keep prices with "$" or "NT$" (never đồng, 円, 원, บาท, rupiah).'
];
function i18nZSystem_(lang, src, D, local, tt) {
  var gl = [], seen = {};
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D, x.kind, x.en, lang); if (tr && !seen[x.en]) { seen[x.en] = 1; gl.push('- ' + x.en + ' → ' + tr); } });
  return ['You translate the in-store tablet recipe of 自己做烘焙聚樂部 (a DIY baking studio in Taiwan; customers bake by themselves, following the tablet page by page).',
    'Source: the store\'s official Traditional Chinese (Taiwan) recipe. Target: ' + I18N_LABEL[lang] + '.',
    'Target style:', I18N_RULES[lang].map(function (r) { return '* ' + r; }).join('\n'),
    'General rules:',
    '* Translate faithfully from the Chinese; do not add, drop, merge or reorder sentences. Keep numbering (1. 2. 3.) and line breaks.',
    '* "h" is HTML. Keep only these tags: <p>, <p class="c">, <br>, <b>, <u>, <i>, <font color="...">. Keep each tag around the corresponding translated words. Keep placeholders like [[IMG1]] exactly where they are.',
    '* Never change any number, weight, time, temperature, speed or count; write them with Arabic numerals exactly as in the source (Thai too: 0-9, never Thai digits). Chinese number words (一、兩、半) may be written naturally in the target language.',
    '* Keep emojis and symbols (●, ※, ⭐, ►, 【】, ［］). Chinese quotes 「」 may become the normal quotes of the target language.',
    '* Never leave any Chinese character in the translation' + (lang === 'ja' ? ' (Japanese kanji are fine for Japanese words, but do not copy Chinese-only wording)' : '') + '.',
    '* When the text mentions an ingredient, tool, mold, area or box from the glossary below, use exactly that translation.',
    '* "t" is the page title (short). Translate it the same way every time it repeats; similar titles with a number (抹面1 / 抹面2) share the same words.',
    local ? '* The dessert name in this language is: ' + local + ' (use it wherever the dessert name appears; return "" for "name").' : '* "name" is the Chinese dessert name: return ONLY its translation (no "｜", no Chinese; the server adds "｜Chinese name"). Translate notes like (葷) = contains meat-based gelatin, (全素) = vegan, (可做全素) = can be made vegan.',
    '* "pres" is the storage advice; translate it keeping line breaks (\\n). Never return it unchanged in Chinese. If "name" or "pres" is given as "", return "".',
    '* Return every k exactly as given.',
    'Fixed terms (Taiwan baking words in this store):', I18N_ZTERMS.map(function (r) { return '* ' + r; }).join('\n') + (lang === 'th' ? '\n* Thai: 麵糊 = ส่วนผสมแป้ง; never แป้งเค้ก or แป้ง alone for batter.' : ''),
    'Glossary (Chinese → ' + I18N_LABEL[lang] + '):', gl.join('\n') || '(none)',
    'Page titles (already translated — use exactly):', Object.keys(tt || {}).map(function (k) { return '- ' + k + ' → ' + tt[k]; }).join('\n') || '(none)'].join('\n');
}

/* ---------------- 詞典整批（第 2 段：Claude Code 依名稱對照表重建中文詞典） ----------------
   動作 i18nZDict（投遞通行碼 wkey）：
     op terms：讀全部候選（或 ids 指定的）中文食譜平板頁，列出用到的詞（品項／取放區／單位）＋出現支數＋提示（後台英文名、越南文後台名、
               舊詞典譯名（只在英文名只對到這個中文品項時））＋目前中文詞典的譯名；不寫任何東西。一次最多 4 分鐘，回 left 就帶 from 再叫。
     op put：rows [{kind, zh, en, ja, ko, vi, id, th}] 寫進 i18n_zitems（空格補上；已有的要 overwrite:1 才蓋），回寫入／略過數。 */
function i18nZDict_(p) {
  i18nWriteKey_(p);
  var op = String(p.op || 'terms'), t0 = Date.now();
  if (op === 'put') {
    var rows = p.rows || [], ow = !!Number(p.overwrite), res = { added: 0, updated: 0, kept: 0, bad: 0 };
    var L = i18nZLangs_();
    i18nLock_(function () {
      var m = i18nMeta_('i18n_zitems', I18N_ZITEM_HEAD), idx = {};
      m.rows.forEach(function (r) { idx[r.kind + '|' + String(r.zh).toLowerCase()] = r; });
      var app = [];
      rows.forEach(function (x) {
        var kind = String(x.kind || ''), zh = rbTrim_(x.zh);
        if (['item', 'area', 'unit'].indexOf(kind) < 0 || !zh) { res.bad++; return; }
        var k = kind + '|' + zh.toLowerCase(), r = idx[k], ch = false;
        if (!r) { r = { kind: kind, zh: zh, how: 'claude-zdict', at: now_() }; L.forEach(function (l) { if (rbTrim_(x[l])) r[l] = rbTrim_(x[l]); }); idx[k] = r; app.push(r); res.added++; return; }
        L.forEach(function (l) { var v = rbTrim_(x[l]); if (!v) return; if (String(r[l] || '').trim() && !ow) { if (String(r[l]).trim() !== v) res.kept++; return; } if (String(r[l] || '').trim() !== v) { r[l] = v; ch = true; } });
        if (ch) { r.how = 'claude-zdict'; r.at = now_(); m.sh.getRange(r._r, 1, 1, I18N_ZITEM_HEAD.length).setNumberFormat('@').setValues([I18N_ZITEM_HEAD.map(function (h) { return r[h] == null ? '' : String(r[h]); })]); res.updated++; }
      });
      if (app.length) m.sh.getRange(m.sh.getLastRow() + 1, 1, app.length, I18N_ZITEM_HEAD.length).setNumberFormat('@').setValues(app.map(function (r) { return I18N_ZITEM_HEAD.map(function (h) { return r[h] == null ? '' : String(r[h]); }); }));
    });
    i18nLog_('Claude Code', 'zdict-put', '', '', 'ok', '新增 ' + res.added + '、更新 ' + res.updated + '、沒蓋 ' + res.kept + '、不收 ' + res.bad);
    return { data: res };
  }
  /* terms */
  var C = i18nZCands_(), want = (p.ids || []).map(function (x) { return String(x).toLowerCase(); });
  if (want.length) C = C.filter(function (c) { return want.indexOf(c.id) >= 0; });
  var from = Number(p.from || 0), seen = {}, list = [], done = from, errs = [];
  for (var i = from; i < C.length; i += 5) {
    if (Date.now() - t0 > 240000) break;
    var part = C.slice(i, i + 5), pages = i18nFetchPages_(part.map(function (c) { return c.id; }));
    part.forEach(function (c, j) {
      try {
        if (pages[j].code !== 200) throw new Error('HTTP ' + pages[j].code);
        var src = i18nParse_(pages[j].text, c.id);
        src.slides.forEach(function (s) { s.ings.forEach(function (g) {
          [['item', i18nSplitName_(g.n).core], ['unit', g.u], ['area', g.an]].forEach(function (a) {
            var zh = rbTrim_(a[1]); if (!zh) return; var k = a[0] + '|' + zh.toLowerCase();
            if (!seen[k]) { seen[k] = { kind: a[0], zh: zh, n: 0, ids: {}, area: a[0] === 'area' ? g.a : '' }; list.push(seen[k]); }
            if (!seen[k].ids[c.id]) { seen[k].ids[c.id] = 1; seen[k].n++; }
          });
        }); });
      } catch (e) { errs.push(c.name + '：' + errMsg_(e)); }
    });
    done = i + part.length;
  }
  var NM = {}; try { NM = i18nZNameMap_(); } catch (e) { }
  var EZ = i18nZEnZh_(NM), OD = i18nDictRows_(), ZD = i18nZDictRows_(), areaEn = {};
  i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (r) { return r.state === 'ok' && !i18nIsZ_(r.id); }).slice(0, 60).forEach(function (r) {
    var s = null; try { s = i18nSrc_(r.id); } catch (e) { } if (!s) return;
    s.slides.forEach(function (x) { x.ings.forEach(function (g) { if (g.a && g.an && !areaEn[g.a]) areaEn[g.a] = g.an; }); });
  });
  var L = i18nZLangs_();
  var items = list.map(function (x) {
    var o = { kind: x.kind, zh: x.zh, n: x.n, cur: {}, hint: {} }, z = ZD[x.kind + '|' + x.zh.toLowerCase()];
    L.forEach(function (l) { if (z && String(z[l] || '').trim()) o.cur[l] = String(z[l]).trim(); });
    if (x.kind === 'item') {
      var h = i18nZHint_(x.zh, 'en', NM, EZ, OD);
      ['backend_en', 'backend_vi', 'note'].forEach(function (f) { if (h[f]) o[f] = h[f]; });
      if (h.cand) o.hint.en = h.cand;
      I18N_LANGS.forEach(function (l) { var hh = i18nZHint_(x.zh, l, NM, EZ, OD); if (hh.cand) o.hint[l] = hh.cand; });
    }
    if (x.kind === 'area' && areaEn[x.area]) { o.backend_en = areaEn[x.area]; o.hint.en = areaEn[x.area]; I18N_LANGS.forEach(function (l) { var c = i18nLook_(OD, 'area', areaEn[x.area], l); if (c) o.hint[l] = c; }); }
    return o;
  });
  return { data: { cands: C.length, from: from, done: done, left: C.length - done, errors: errs, items: items, sec: Math.round((Date.now() - t0) / 1000) } };
}
ACTIONS.i18nZDict = function (p, a) { return i18nZDict_(p, a); };
NOAUTH.i18nZDict = 1; SELFLOCK.i18nZDict = 1;

/* ---------------- 平板清單的一支（中文原文） ----------------
   有舊英文原文（old，且舊的還在平板上）：6 種語言都能顯示之前不列（平板照舊列舊的那支）→ 回 null。
   name 寫成「英文名｜中文名」（平板卡片：英文頁副標＝中文、其他語言副標＝整串，跟舊的一樣）；英文名用到店須知頁或英文譯文的品名。 */
function i18nZListItem_(s, o, rc, sh, z, okm, zhm, names) {   /* names：{id: 原文名稱}（i18n_src） */
  var all = true;
  i18nZLangs_().forEach(function (l) {
    var r = rc[s.id + '|' + l]; o.langs[l] = i18nState_(r, s.fp);
    if (o.langs[l] !== 'live' && i18nShowable_(r)) o.langs[l] = 'live';   /* 新譯文還沒好 → 照舊顯示上一版 */
    if (o.langs[l] !== 'live') { all = false; return; }
    o.ver[l] = i18nHash_([s.fp, s.upd, r.fp, r.appr_fp, r.appr_at].join('|')).slice(0, 12);
  });
  if (!all && z.old && okm[z.old]) {   /* 舊英文那支還在平板上 → 中文版先不列；但同時是中文試看的，「中文」那一格照常（2026-11-24 修：之前整支不見） */
    if (!zhm[s.id]) return null;
    return { id: s.id, name: s.name, cover: s.cover, langs: { zh: 'live' }, ver: { zh: i18nHash_([s.fp, s.upd, 'zh'].join('|')).slice(0, 12) }, local: { zh: zhm[s.id] } };
  }
  var oldName = '';
  if (z.old) oldName = names[z.old] || '';
  i18nZLangs_().forEach(function (l) {
    if (o.langs[l] !== 'live') return;
    var t = i18nZLocal_(s, l, oldName) || i18nTrName_(sh, rc[s.id + '|' + l], o.ver[l]);
    if (t) o.local[l] = t;
  });
  if (o.local.en) o.name = o.local.en + '｜' + s.name;
  if (zhm[s.id]) { o.langs.zh = 'live'; o.ver.zh = i18nHash_([s.fp, s.upd, 'zh'].join('|')).slice(0, 12); o.local.zh = zhm[s.id]; }   /* 同時是中文試看（2026-11-17）：多一格中文 */
  return o;
}
function i18nZOldName_(src) {   /* 對應的舊英文原文名稱（到店須知品名對不到中文名時用英文名對） */
  var old = src && src.old; if (!old) return '';
  var r = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (x) { return x.id === old; })[0];
  return r ? r.name : '';
}
