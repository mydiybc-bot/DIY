/* ===== guide_src.gs — 外國客人到店須知（QR Code 頁）的甜點資料來源，2026-11-26 =====
   經營者：「QR code 掃描後的資料還是停留在 26 支食譜，這不對…未來要能和每天的外語食譜更新一起進行更新檢查」；
   「88 支都要上，不要自行合併」；甜點介紹「一律用後台中文食譜」。
   動作 guideSrc（唯讀預覽通行碼 pkey；只讀、不寫任何表、不寫後台、不花 API）：
     ‧ 不給 ids → 回平板外語食譜「中文原文」打開且原文抓取正常的清單 [{id, name, upd, cover}]（便宜，用來分批）
     ‧ ids:[…]（一次最多 15 支）→ 每支回：後台「編輯食譜」頁的 Title、Price、Size、PrepHr、Content（產品特色，純文字保留換行）、
       Preserve、InUse、Public、類別、分店類別；平板各語言放行中的品名 names{en,ja,ko,vi,id,th}（譯文自己的品名，
       I18N_ZNAME 自訂名優先、(全素)／(可做全素) 照 I18N_ZVEG 補，跟平板同規則；但不讀到店須知名，避免兩邊互相引用卡住舊名）、各語言狀態 langs。
   呼叫端：DIY tools/i18n_claude/guide.py（Claude Code 週一到週五 10:00 外語食譜巡檢時一起跑）。 */
var GUIDE_MAX_IDS = 15;

function guideSrc_(p) {
  if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) { Utilities.sleep(1000); throw fail_('要通行碼', 'auth'); }
  var ZM = i18nZMap_(), sm = i18nMeta_('i18n_src', I18N_SRC_HEAD), src = {};
  sm.rows.forEach(function (s) { src[s.id] = s; });
  var list = Object.keys(ZM).filter(function (id) { return src[id] && src[id].state === 'ok'; })
    .map(function (id) { var s = src[id]; return { id: id, name: s.name, upd: s.upd, cover: s.cover }; });
  if (!p.ids) return { data: { at: now_(), count: list.length, items: list } };
  var ids = [].concat(p.ids).map(function (x) { return String(x || '').trim().toLowerCase(); }).filter(Boolean);
  if (ids.length > GUIDE_MAX_IDS) throw fail_('一次最多 ' + GUIDE_MAX_IDS + ' 支', 'toomany');
  var bad = ids.filter(function (id) { return !ZM[id] || !src[id] || src[id].state !== 'ok'; });
  if (bad.length) throw fail_('不是中文原文上線中的食譜：' + bad.join(','), 'notz');
  var rm = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), rc = {};
  rm.rows.forEach(function (r) { rc[r.key] = r; });
  var pages = guideEditPages_(ids), out = [];
  ids.forEach(function (id, i) {
    var s = src[id], o = { id: id, name: s.name, upd: s.upd, cover: s.cover, names: {}, langs: {} };
    try { o.edit = guideParseEdit_(pages[i], id); } catch (e) { o.err = errMsg_(e); }
    i18nZLangs_().forEach(function (l) {
      var r = rc[id + '|' + l], st = i18nState_(r, s.fp);
      if (st !== 'live' && i18nShowable_(r)) st = 'live';
      o.langs[l] = st || '';
      if (st !== 'live') return;
      var ver = i18nHash_([s.fp, s.upd, r.fp, r.appr_fp, r.appr_at].join('|')).slice(0, 12);
      o.names[l] = guideLocalName_(s.name, l, i18nTrName_(rm.sh, r, ver));
    });
    out.push(o);
  });
  return { data: { at: now_(), items: out } };
}
/* 平板同規則的在地品名，只是不讀到店須知名（i18nZLocal_ 會先查到店須知名） */
function guideLocalName_(nm, lang, trName) {
  nm = String(nm || '');
  var base = nm.replace(/\s*[(（][^)）]*[)）]/g, '').trim(), fix = I18N_ZNAME[base];
  var t = (fix && fix[lang]) || rbTrim_(trName || '');   /* 2026-11-28 修：固定品名沒列的語言改用譯文名（原本留空白 → 到店須知檢查擋下） */
  if (!t) return '';
  var veg = /[(（]\s*可做全素\s*[)）]/.test(nm) ? 'o' : /[(（]\s*全素\s*[)）]/.test(nm) ? 'v' : '';
  if (veg && I18N_ZVEG[veg][lang] && !/vegan|ヴィーガン|비건|chay|วีแกน/i.test(t)) t += I18N_ZVEG[veg][lang];
  return t;
}
function guideEditPages_(ids) {
  var sess = rbSession_();
  var mk = function () { var c = rbJarStr_(sess.jar); return ids.map(function (id) { return { url: RB_BASE + '/Recipes/Edit/' + encodeURIComponent(id), method: 'get', muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA, Cookie: c } }; }); };
  var rs = UrlFetchApp.fetchAll(mk());
  if (rs.some(function (r) { return r.getResponseCode() === 302 && /Identity\/Account\/Login/i.test(String(r.getHeaders().Location || r.getHeaders().location || '')); })) {
    rbDropSession_(); sess.jar = rbFreshSession_().jar; rs = UrlFetchApp.fetchAll(mk());
  }
  rbSaveSession_(sess);
  return rs.map(function (r) { return { code: r.getResponseCode(), text: r.getResponseCode() === 200 ? r.getContentText() : '' }; });
}
function guideParseEdit_(pg, id) {
  if (pg.code !== 200) throw rbErr_('後台「編輯食譜」頁回 HTTP ' + pg.code, 'rbnet');
  var f = rbFormHtml_(pg.text, /^\/Recipes\/Edit\//i);
  if (!f) throw rbErr_('「編輯食譜」頁找不到表單（頁面可能改版）', 'rbpage');
  var fields = rbFormFields_(f);
  if (String(rbFieldGet_(fields, 'RecipeId') || '').toLowerCase() !== id) throw rbErr_('「編輯食譜」頁的食譜編號對不上', 'rbpage');
  var g = function (n) { var v = rbFieldGet_(fields, n); return v == null ? '' : String(v); };
  var t = pg.text, a = t.indexOf('id="cateG"'), b = t.indexOf('id="storeType"'), c = t.indexOf('name="RecipeUp"');
  var items = function (from, to) {
    if (from < 0 || to < from) return [];
    var seg = t.slice(from, to), o = [], re = /<span\s+class=['"]item['"]\s+data-item=['"]([^'"]+)['"]\s*>([^<]*)/gi, m;
    while ((m = re.exec(seg))) o.push(rbTrim_(rbDecode_(m[2])));
    return o;
  };
  return { title: rbTrim_(g('Title')), price: g('Price').trim(), size: rbTrim_(g('Size')), prep: g('PrepHr').trim(),
    content: i18nTxt_(g('Content').replace(/\r?\n/g, '<br>')), preserve: i18nTxt_(g('Preserve').replace(/\r?\n/g, '<br>')),
    inUse: g('InUse') === 'true', pub: g('Public') === 'true', cats: items(a, b), stores: items(b, c) };
}
ACTIONS.guideSrc = function (p, a) { return guideSrc_(p, a); };
NOAUTH.guideSrc = 1; READS.guideSrc = 1;
