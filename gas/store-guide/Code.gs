/* 外國客人到店須知 API（2026-10-23 經營者交辦）
   每天 23:00 讀官網 English Menu（品項、售價、尺寸、說明、照片）→ 翻 7 語 → 自我檢查 → 通過才發佈 → 發佈後再從網址讀回比對。
   頁面：DIY static/store-guide.html（客人用，免密碼）；?hq=1 顯示總部看的更新狀態。
   翻譯：每款記下官網英文原文（src）；原文沒變就沿用（第一次用 Seed.gs 的人工翻譯），原文變了或新品才用 LanguageApp 機器翻譯，並列在狀態的「機器翻譯」清單。
   檢查沒過 → 不發佈、保留上一版，狀態顯示原因。 */

var MENU_URL = 'https://mydiybc.com/home/menu?id=6c236c58-cf38-49a4-8591-616e933de566';
var LANGS = ['en', 'zh', 'ja', 'ko', 'vi', 'id', 'th'];
var GT = { en: 'en', zh: 'zh-TW', ja: 'ja', ko: 'ko', vi: 'vi', id: 'id', th: 'th' };
var ITEMS_TAB = 'items', I18N_TAB = 'i18n', LOG_TAB = 'log';
var I18N_COLS = ['key', 'lang', 'how', 'src', 'name', 'main', 'fill', 'deco', 'note', 'at'];
var SYNC_GAP_MS = 10 * 60 * 1000;

/* ---------- 網頁 API ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = p.action || 'data', res;
  try {
    if (a === 'data') res = data_();
    else if (a === 'status') res = { ok: true, status: status_() };
    else if (a === 'sync') res = syncThrottled_();
    else if (a === 'ping') res = { ok: true, at: now_() };
    else res = { ok: false, error: 'unknown action' };
  } catch (err) { res = { ok: false, error: String(err && err.message || err) }; }
  return out_(res, p.callback);
}
function out_(res, cb) {
  var s = JSON.stringify(res);
  if (cb && /^[\w.$]+$/.test(cb)) return ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function data_() { return { ok: true, status: status_(), items: readItems_() }; }
function status_() {
  var s;
  try { s = JSON.parse(PropertiesService.getScriptProperties().getProperty('status') || '{}'); } catch (e) { s = {}; }
  try { s.schedule = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'sync'; }).length; } catch (e) { s.schedule = -1; }
  return s;
}

/* ---------- 編輯器用 ---------- */
function setup() {   /* 第一次授權＋建分頁＋設每天 23:00 排程（重跑不會重複建） */
  sheet_(ITEMS_TAB, ['key', 'json', 'at']);
  sheet_(I18N_TAB, I18N_COLS);
  sheet_(LOG_TAB, ['at', 'by', 'result', 'count', 'added', 'removed', 'price', 'text', 'auto', 'errors', 'sec']);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'sync') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sync').timeBased().atHour(23).nearMinute(0).everyDays(1).inTimezone('Asia/Taipei').create();
  Logger.log('setup 完成：排程 sync 每天 23 點；官網讀取測試 ' + UrlFetchApp.fetch(MENU_URL, { muteHttpExceptions: true }).getResponseCode()
    + '；翻譯測試 ' + LanguageApp.translate('Chocolate', 'en', 'ja'));
}
function sync() { return sync_('排程'); }
function syncNow() { return sync_('手動'); }

function syncThrottled_() {
  var pr = PropertiesService.getScriptProperties(), last = Number(pr.getProperty('sync_ms') || 0);
  if (Date.now() - last < SYNC_GAP_MS) return { ok: false, error: '剛更新過，' + Math.ceil((SYNC_GAP_MS - (Date.now() - last)) / 60000) + ' 分鐘後再試', status: status_() };
  var r = sync_('體驗中心按鈕');
  return { ok: r.result === 'ok', status: status_() };
}

/* ---------- 同步主流程 ---------- */
function sync_(by) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { result: 'busy' };
  var t0 = Date.now(), pr = PropertiesService.getScriptProperties();
  pr.setProperty('sync_ms', String(t0));
  var st = { at: now_(), by: by, result: 'fail', errors: [], warnings: [] };
  try {
    var resp = UrlFetchApp.fetch(MENU_URL, { muteHttpExceptions: true, followRedirects: true });
    if (resp.getResponseCode() !== 200) throw new Error('官網回應 ' + resp.getResponseCode());
    var raw = parseMenu_(resp.getContentText('UTF-8'));
    var prev = readItems_(), i18n = readI18n_(), autoKeys = [];
    var items = raw.map(function (r) {
      var src = JSON.stringify([r.en, r.main, r.fill, r.deco, r.note]), t = {};
      LANGS.forEach(function (l) {
        var k = r.en + '|' + l, row = i18n[k];
        if (!row || row.src !== src) {
          var seed = !row && typeof SEED !== 'undefined' && SEED[r.en] && SEED[r.en][l];
          row = seed ? { how: 'manual', v: seed } : { how: 'auto', v: translate_(r, l) };
          row.src = src; row.at = now_(); row.dirty = true;
          i18n[k] = row;
        }
        if (row.how === 'auto' && l !== 'en' && autoKeys.indexOf(r.en) < 0) autoKeys.push(r.en);
        t[l] = row.v;
      });
      return { id: r.id, en: r.en, zh: r.zh, p: r.price, img: r.img, s: r.size, size_raw: r.size_raw, tags: r.tags, f: r.flags, hallo_end: r.hallo_end, t: t };
    });
    var diff = diff_(prev, items);
    st.count = items.length;
    st.added = diff.added; st.removed = diff.removed; st.price = diff.price; st.text = diff.text;
    st.auto = autoKeys;
    st.errors = check_(items, prev);
    if (st.errors.length) throw new Error('檢查沒過，保留上一版');
    saveI18n_(i18n);
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

/* ---------- 讀官網 ---------- */
function parseMenu_(html) {
  var parts = html.split('<div class="modal fade" id="m-').slice(1), out = [];
  parts.forEach(function (p) {
    var id = p.slice(0, p.indexOf('"'));
    var name = text_(between_(p, '<div class="menuName"', '</div>').replace(/^[^>]*>/, ''));
    if (!name || name.indexOf('｜') < 0) return;
    var price = text_(between_(p, '<div class="menuPrice"', '</div>').replace(/^[^>]*>/, ''));
    var sub = text_(between_(p, '<div class="menuSubTitle">', '</div>'));
    var desc = between_(p, '<div class="menuContent">', '<div class="menuDetail">');
    desc = desc.slice(desc.indexOf('white-space:pre-line">') + 22);
    desc = decode_(desc.slice(0, desc.indexOf('</div>'))).replace(/\r/g, '');
    var imgs = (p.match(/<img class="menuImg" src="([^"]+)"/g) || []).map(function (m) { return m.replace(/^.*src="/, '').replace(/"$/, ''); })
      .filter(function (u) { return /\.(jpe?g|png|webp)(\?|$)/i.test(u); });
    var nm = name.split('｜'), en = nm[0].trim(), zhFull = nm.slice(1).join('｜').trim();
    var sfx = (zhFull.match(/[\(（]([^\)）]*)[\)）]\s*$/) || ['', ''])[1];
    var zh = zhFull.replace(/\s*[\(（][^\)）]*[\)）]\s*$/, '').trim();
    var d = parseDesc_(desc);
    var tags = [];
    if (/壽星/.test(sfx)) tags.push('bday');
    if (/限信義/.test(sfx)) tags.push('bread');
    if (/無法製作/.test(sfx)) tags.push('noXK');
    if (/made by 1 or 2 people/i.test(desc)) tags.push('pair');
    if (/葷/.test(sfx)) tags.push('meat');
    if (d.hallo_end) tags.push('hallo');
    var sizeRaw = sub.split('/')[0].trim();
    out.push({ id: id, en: en, zh: zh, price: Number(String(price).replace(/[^\d]/g, '')), img: imgs[0] || '',
      size: size_(sizeRaw), size_raw: sizeRaw, tags: tags, flags: d.flags + (/gluten-free/i.test(en) ? 'g' : ''),
      hallo_end: d.hallo_end, main: d.main, fill: d.fill, deco: d.deco, note: d.note });
  });
  return out;
}
function parseDesc_(txt) {
  var r = { main: [], fill: [], deco: [], note: [], flags: '', hallo_end: '' };
  var lines = txt.split('\n').map(function (s) { return s.trim(); }).filter(String);
  for (var i = 0; i < lines.length; i++) {
    var s = lines[i];
    if (/Birthday Special Offer|Important Notes|How to Redeem/i.test(s)) break;
    var h = s.match(/Halloween Limited\s*(\d{1,2})\/(\d{1,2})\s*[~～-]\s*(\d{1,2})\/(\d{1,2})/i);
    if (h) { r.hallo_end = dayAfter_(+h[3], +h[4]); continue; }
    if (s.indexOf('［') >= 0) {
      (s.match(/［([^］]+)］/g) || []).forEach(function (b) {
        b = b.slice(1, -1);
        if (/^Ovo-lacto/i.test(b)) r.flags += 'o';
        else if (/^Lacto/i.test(b)) r.flags += 'l';
        else if (/^Vegan/i.test(b)) r.flags += 'v';
        else if (/alcohol \(optional\)/i.test(b)) r.flags += 'a';
        else if (/alcohol/i.test(b)) r.flags += 'A';
        else if (/Allergen/i.test(b)) r.flags += 'n';
        else if (/Ready-made/i.test(b)) r.flags += 'r';
      });
      continue;
    }
    var m = s.match(/^●\s*([^:：]+)[:：]\s*(.*)$/);
    if (m) {
      var k = m[1].toLowerCase(), v = m[2].trim();
      if (!v) continue;
      if (/main/.test(k)) r.main.push(v);
      else if (/filling|top layer/.test(k)) r.fill.push(v);
      else if (/decoration/.test(k)) r.deco.push(v);
      else if (/size/.test(k)) r.note.push('Size: ' + v);
      continue;
    }
    if (/^【/.test(s) && !/Halloween/i.test(s)) { r.note.push(s.replace(/^【|】$/g, '')); continue; }
    if (/^\*/.test(s)) { r.note.push(s.replace(/^\*\s*/, '')); continue; }
  }
  var flags = r.flags.split('').filter(function (c, i, a) { return a.indexOf(c) === i; }).join('');
  return { main: r.main.join(' / '), fill: r.fill.join(' / '), deco: r.deco.join(' / '), note: r.note.join(' '), flags: flags, hallo_end: r.hallo_end };
}
function size_(s) {
  var m;
  if ((m = s.match(/(\d+)\s*吋\s*[×x]\s*(\d+)/))) return ['tart', +m[1], +m[2]];
  if ((m = s.match(/(\d+)\s*cm\s*[×x]\s*(\d+)\s*cm/i))) return ['cm', m[1] + ' × ' + m[2]];
  if ((m = s.match(/(\d+)\s*(?:吋|-?\s*inch)/i))) return ['inch', +m[1]];
  if ((m = s.match(/(\d+)\s*(?:pcs?\.?|個|片)/i))) return ['pcs', +m[1]];
  return null;
}
function dayAfter_(mo, d) {
  var y = new Date().getFullYear(), dt = new Date(y, mo - 1, d + 1);
  return Utilities.formatDate(dt, 'Asia/Taipei', 'yyyy-MM-dd');
}

/* ---------- 翻譯 ---------- */
function translate_(r, l) {
  if (l === 'en') return ['', r.main, r.fill, r.deco, r.note];
  var tr = function (s) { return s ? LanguageApp.translate(s, 'en', GT[l]) : ''; };
  return [l === 'zh' ? '' : tr(r.en), tr(r.main), tr(r.fill), tr(r.deco), tr(r.note)];
}

/* ---------- 檢查（沒過就不發佈） ---------- */
function check_(items, prev) {
  var errs = [], seen = {};
  if (!items.length) return ['官網讀不到任何甜點（網頁格式可能改了）'];
  if (prev.length && items.length < Math.ceil(prev.length / 2)) errs.push('款數從 ' + prev.length + ' 掉到 ' + items.length + '，疑似讀取不完整');
  items.forEach(function (it) {
    var who = it.en || it.id;
    if (seen[it.en]) errs.push(who + '：品名重複'); seen[it.en] = 1;
    if (!it.en || !it.zh) errs.push(who + '：英文或中文品名空白');
    if (!(it.p >= 0) || isNaN(it.p)) errs.push(who + '：售價讀不到');
    if (!it.img) errs.push(who + '：沒有照片');
    LANGS.forEach(function (l) {
      var t = it.t[l] || [];
      if (l !== 'en' && l !== 'zh' && !t[0]) errs.push(who + '：' + l + ' 品名空白');
      if (!t[1] && !t[3]) errs.push(who + '：' + l + ' 介紹空白');
    });
  });
  var reqs = items.filter(function (it) { return it.img; }).map(function (it) { return { url: it.img, headers: { Range: 'bytes=0-255' }, muteHttpExceptions: true }; });
  var res = reqs.length ? UrlFetchApp.fetchAll(reqs) : [];
  res.forEach(function (r, i) { var c = r.getResponseCode(); if (c !== 200 && c !== 206) errs.push(items[i].en + '：照片打不開（' + c + '）'); });
  return errs;
}
/* 發佈後從對外網址讀回來，逐款比對（＝客人打開看到的資料） */
function verify_(items) {
  var w = [], url = ScriptApp.getService().getUrl();
  var back = readItems_();
  if (JSON.stringify(back) !== JSON.stringify(items)) w.push('資料表讀回與寫入不一致');
  if (!url) { w.push('還沒部署網址，略過網址讀回'); return w; }
  try {
    var r = UrlFetchApp.fetch(url + '?action=data', { muteHttpExceptions: true, followRedirects: true });
    var j = JSON.parse(r.getContentText());
    if (!j.ok) w.push('網址讀回失敗：' + j.error);
    else if (JSON.stringify(j.items) !== JSON.stringify(items)) w.push('網址讀回與發佈內容不一致');
  } catch (e) { w.push('網址讀回失敗：' + e.message); }
  return w;
}
function diff_(prev, items) {
  var pm = {}, nm = {}, d = { added: [], removed: [], price: [], text: [] };
  prev.forEach(function (x) { pm[x.en] = x; });
  items.forEach(function (x) {
    nm[x.en] = 1;
    var o = pm[x.en];
    if (!o) { d.added.push(x.en); return; }
    if (o.p !== x.p) d.price.push(x.en + ' ' + o.p + '→' + x.p);
    if (JSON.stringify(o.t.en) !== JSON.stringify(x.t.en)) d.text.push(x.en);
  });
  prev.forEach(function (x) { if (!nm[x.en]) d.removed.push(x.en); });
  return d;
}

/* ---------- 資料表 ---------- */
function ss_() { return SpreadsheetApp.getActive(); }
function sheet_(name, hdr) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { sh = ss_().insertSheet(name); sh.getRange(1, 1, 1, hdr.length).setValues([hdr]); sh.setFrozenRows(1); }
  return sh;
}
function readItems_() {
  var sh = ss_().getSheetByName(ITEMS_TAB);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return JSON.parse(r[0]); });
}
function writeItems_(items) {
  var sh = sheet_(ITEMS_TAB, ['key', 'json', 'at']), at = now_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 3).clearContent();
  if (items.length) sh.getRange(2, 1, items.length, 3).setNumberFormat('@').setValues(items.map(function (it) { return [it.en, JSON.stringify(it), at]; }));
  SpreadsheetApp.flush();
}
function readI18n_() {
  var sh = ss_().getSheetByName(I18N_TAB), m = {};
  if (!sh || sh.getLastRow() < 2) return m;
  sh.getRange(2, 1, sh.getLastRow() - 1, I18N_COLS.length).getValues().forEach(function (r) {
    m[r[0] + '|' + r[1]] = { how: r[2], src: r[3], v: [r[4], r[5], r[6], r[7], r[8]].map(String), at: r[9] };
  });
  return m;
}
function saveI18n_(m) {
  var sh = sheet_(I18N_TAB, I18N_COLS);
  var rows = Object.keys(m).sort().map(function (k) {
    var r = m[k], p = k.split('|');
    return [p[0], p[1], r.how, r.src].concat(r.v).concat([r.at || '']);
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, I18N_COLS.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, I18N_COLS.length).setNumberFormat('@').setValues(rows);
}
function log_(st) {
  var sh = sheet_(LOG_TAB, ['at', 'by', 'result', 'count', 'added', 'removed', 'price', 'text', 'auto', 'errors', 'sec']);
  var j = function (a) { return (a || []).join('；').slice(0, 2000); };
  sh.appendRow([st.at, st.by, st.result, st.count || '', j(st.added), j(st.removed), j(st.price), j(st.text), j(st.auto), j(st.errors.concat(st.warnings || [])), st.sec]);
}
function trimStatus_(st) {
  ['added', 'removed', 'price', 'text', 'auto', 'errors', 'warnings'].forEach(function (k) { if (st[k] && st[k].length > 40) st[k] = st[k].slice(0, 40).concat(['…共 ' + st[k].length + ' 項']); });
  return st;
}

/* ---------- 小工具 ---------- */
function now_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss'); }
function between_(s, a, b) { var i = s.indexOf(a); if (i < 0) return ''; var j = s.indexOf(b, i + a.length); return j < 0 ? '' : s.slice(i + a.length, j); }
function text_(s) { return decode_(String(s).replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ').trim(); }
function decode_(s) {
  return String(s)
    .replace(/&#x([0-9a-f]+);/gi, function (_, h) { return String.fromCodePoint(parseInt(h, 16)); })
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCodePoint(+d); })
    .replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
}
