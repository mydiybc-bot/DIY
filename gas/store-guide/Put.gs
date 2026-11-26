/* ===== Put.gs — 2026-11-26 起的發佈口：平板外語食譜（中文原文）88 支 → 到店須知 =====
   POST {action:'put', key, items:[…], lines:{key:{zh,en,ja,ko,vi,id,th}}}（DIY tools/i18n_claude/guide.py 送）
   items 每款（v:2）：{id, en, zh, p, img, s, size_raw, prep, tags, cats, t:{zh:['',產品特色,尺寸], en:[品名,產品特色,尺寸], ja…}}
   檢查沒過 → 不發佈、保留上一版；過了 → 寫 items、lines，再從網址讀回比對（verify_）。 */
var V2_LANGS = ['en', 'ja', 'ko', 'vi', 'id', 'th'];

function doPost(e) {
  var p, res;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, error: '送來的資料不是 JSON' }); }
  try {
    if (p.action !== 'put') res = { ok: false, error: 'unknown action' };
    else if (!p.key || sha256_(String(p.key)) !== PUT_KEY_SHA256) { Utilities.sleep(1000); res = { ok: false, error: '通行碼不對' }; }
    else { var st = put_(p.items || [], p.lines || {}); res = { ok: st.result !== 'fail', status: st }; }
  } catch (err2) { res = { ok: false, error: String(err2 && err2.message || err2) }; }
  return out_(res);
}
function sha256_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}

function put_(items, lines) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { result: 'fail', errors: ['忙碌中，稍後再試'], warnings: [] };
  var t0 = Date.now(), pr = PropertiesService.getScriptProperties();
  var st = { at: now_(), by: '外語食譜巡檢', src: '平板外語食譜（中文原文）＋門市食譜後台', result: 'fail', errors: [], warnings: [] };
  try {
    var prev = readItems_();
    var diff = diffV2_(prev, items);
    st.count = items.length;
    st.added = diff.added; st.removed = diff.removed; st.price = diff.price; st.text = diff.text;
    st.auto = [];
    st.errors = checkV2_(items, prev);
    if (st.errors.length) throw new Error('檢查沒過，保留上一版');
    writeLines_(lines);
    writeItems_(items);
    st.warnings = verify_(items);
    st.result = st.warnings.length ? 'warn' : 'ok';
  } catch (err) {
    if (st.errors.indexOf(String(err.message || err)) < 0) st.errors.push(String(err.message || err));
    st.result = 'fail';
  } finally {
    st.sec = Math.round((Date.now() - t0) / 1000);
    var old = status_();
    st.last_ok = st.result === 'fail' ? (old.last_ok || '') : st.at;
    st.published_count = st.result === 'fail' ? (old.published_count || 0) : st.count;
    pr.setProperty('status', JSON.stringify(trimStatus_(st)));
    log_(st);
    lock.releaseLock();
  }
  return st;
}

function checkV2_(items, prev) {
  var errs = [], seen = {}, han = /[㐀-鿿]/;
  if (!items.length) return ['沒有任何甜點'];
  if (prev.length >= 40 && items.length < Math.ceil(prev.length * 0.7)) errs.push('款數從 ' + prev.length + ' 掉到 ' + items.length + '，疑似讀取不完整');
  items.forEach(function (it) {
    var who = (it.zh || it.id || '?');
    if (it.v !== 2) errs.push(who + '：資料格式不是 v2');
    if (!it.id || seen[it.id]) errs.push(who + '：編號空白或重複'); seen[it.id] = 1;
    if (!it.en || !it.zh) errs.push(who + '：英文或中文品名空白');
    if (it.p !== null && (typeof it.p !== 'number' || !(it.p >= 0))) errs.push(who + '：售價不對');
    if (!it.img) errs.push(who + '：沒有照片');
    var t = it.t || {};
    if (!t.zh || typeof t.zh[1] !== 'string') errs.push(who + '：中文產品特色缺');
    V2_LANGS.forEach(function (l) {
      var x = t[l] || [];
      if (!x[0]) errs.push(who + '：' + l + ' 品名空白');
      if ((t.zh && t.zh[1] || '').trim() && !String(x[1] || '').trim()) errs.push(who + '：' + l + ' 介紹空白');
      if (l !== 'ja' && han.test(String(x[1] || '') + String(x[2] || ''))) errs.push(who + '：' + l + ' 介紹還有中文字');
      var zl = String(t.zh && t.zh[1] || '').split('\n').length, xl = String(x[1] || '').split('\n').length;
      if (zl !== xl) errs.push(who + '：' + l + ' 介紹行數 ' + xl + '，中文 ' + zl);
    });
  });
  var reqs = items.filter(function (it) { return it.img; }).map(function (it) { return { url: it.img, headers: { Range: 'bytes=0-255' }, muteHttpExceptions: true }; });
  var res = reqs.length ? UrlFetchApp.fetchAll(reqs) : [];
  res.forEach(function (r, i) { var c = r.getResponseCode(); if (c !== 200 && c !== 206) errs.push(items[i].zh + '：照片打不開（' + c + '）'); });
  return errs;
}
function diffV2_(prev, items) {
  var k = function (x) { return x.id || x.en; }, pm = {}, nm = {}, d = { added: [], removed: [], price: [], text: [] };
  prev.forEach(function (x) { pm[k(x)] = x; });
  items.forEach(function (x) {
    nm[k(x)] = 1;
    var o = pm[k(x)];
    if (!o) { d.added.push(x.zh || x.en); return; }
    if (o.p !== x.p) d.price.push((x.zh || x.en) + ' ' + o.p + '→' + x.p);
    if (JSON.stringify(o.t) !== JSON.stringify(x.t)) d.text.push(x.zh || x.en);
  });
  prev.forEach(function (x) { if (!nm[k(x)]) d.removed.push(x.zh || x.en); });
  return d;
}

/* ---------- 逐行譯文 ---------- */
function readLines_() {
  var sh = ss_().getSheetByName(LINES_TAB), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, LINES_COLS.length).getValues().forEach(function (r) {
    if (!r[0]) return;
    var o = m[r[0]] = {};
    LINES_COLS.forEach(function (c, i) { if (i && c !== 'at' && r[i] !== '') o[c] = String(r[i]); });
  });
  return m;
}
function writeLines_(lines) {   /* 送來的覆蓋同代號；沒送來的舊行留著（別支甜點或改版前的行，日後可清） */
  var m = readLines_(), at = now_();
  Object.keys(lines || {}).forEach(function (k) { var o = lines[k] || {}; m[k] = m[k] || {}; LINES_COLS.forEach(function (c) { if (c !== 'key' && c !== 'at' && o[c]) m[k][c] = String(o[c]); }); m[k].at = at; });
  var sh = sheet_(LINES_TAB, LINES_COLS), rows = Object.keys(m).sort().map(function (k) { return LINES_COLS.map(function (c) { return c === 'key' ? k : (m[k][c] || ''); }); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, LINES_COLS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, LINES_COLS.length).setNumberFormat('@').setValues(rows);
}
