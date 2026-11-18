/* ================= 🧪 平板外語食譜：改用「中文食譜」當原文的試做（2026-11-18） =================
   經營者 11/18：「將這個外國人用的多語言食譜全部用自己做食譜系統的中文食譜 Active 的台灣區食譜當做基準（不含裝飾蛋糕），來進行翻譯」
   → 評估後選「先試 10 支」。
   這裡只做試做：讀後台中文食譜的平板頁 → 換掉有中文字的共用圖片影片 → 用正式流程同一個模型（Sonnet）翻：
   詞彙表、標題先翻、每 4 頁一段並行、硬檢查（頁數、數字、圖片位置）、AI 回讀、有問題修一輪 → 結果直接回傳給呼叫端（本機存檔比對）。
   ⚠ 不寫任何分頁（i18n_src／i18n_rcp／i18n_items／i18n_log 都不動）、不寫後台；平板與現有譯文完全不受影響。
   動作 i18nZTry（要投遞通行碼 wkey）：{zid 中文食譜 id, enid 對應的英文食譜 id（拿到店須知的在地品名、取放區英文名用；可空）, lang, dry:1 只回原文不翻}
   試做結束、正式改版時再決定要不要併進 i18n.gs；不要了就刪掉這個檔。 */
var I18NZ_CHEF = 'https://cdn.jsdelivr.net/gh/mydiybc-bot/DIY@f9e57f9cf05bd011524e2b4f21c6b40f284cfab4/media/tablet-intro/chef-bear-';
var I18NZ_MD5 = {   /* 主圖／主影片（認檔案指紋）：中文版 → 外國客人版 */
  'zwU1ifRZOqYg5KT5Bifb/Q==': { video: 1 },   /* 中文教學影片 → 熊主廚各語言版（英文版食譜目前也是這樣換） */
  'O51ZCnNhWpVOX1alTp0kTQ==': { img: 'https://diybcstorage.blob.core.windows.net/stepmedias/941ffc04-31c8-4c12-a7aa-6cc2745be443/9866e7c2-138a-437e-88ec-67bfca5e0c8c.jpg' }   /* 烤熟測試（圖上有中文）→ 英文版 */
};
var I18NZ_IMG = [   /* 內文插圖（認檔名）：中文版 → 英文版 */
  ['影片重置(中文)', 'https://diybcstorage.blob.core.windows.net/stepmedias/20241113101118影片重置(英文).jpg'],
  ['量匙怎麼舀_中文', 'https://diybcstorage.blob.core.windows.net/stepmedias/20261113141101量匙怎麼舀_英文.webp']
];
var I18NZ_TW = /自己人募集中|裝飾物加購|透明盒/;   /* 台灣限定的頁（英文版目前刪掉）：照翻，但標出來給經營者決定 */
var I18NZ_FIXED = { t: '量匙怎麼舀', zh: ['舀起後輕輕抖一抖，一平匙齊平、半平匙裝一半。', '左滑開始製作'] };   /* 對到 I18N_FIXED 第一組的固定譯文 */
var I18NZ_NUMW = { '1': /[一單每半首]/g, '2': /[二兩雙半對]/g, '3': /三/g, '4': /四/g, '5': /五/g, '6': /六/g, '7': /七/g, '8': /八/g, '9': /九/g, '10': /十/g };

function i18nzTry_(p) {
  i18nWriteKey_(p);
  var t0 = Date.now(), zid = String(p.zid || '').trim().toLowerCase(), enid = String(p.enid || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  if (!/^[0-9a-f-]{36}$/.test(zid)) throw fail_('zid 要給中文食譜的完整 id', 'bad');
  if (!p.dry && I18N_LANGS.indexOf(lang) < 0) throw fail_('lang 要是 ' + I18N_LANGS.join('／'), 'bad');
  var pg = i18nFetchPages_([zid])[0];
  if (pg.code !== 200) throw fail_('後台回 HTTP ' + pg.code, 'rbnet');
  var src = i18nParse_(pg.text, zid);
  var swap = i18nzSwap_(src, p.dry ? '' : lang);
  src.slides.forEach(function (s) { s.tw = I18NZ_TW.test(s.t || '') ? 1 : 0; });
  var en = enid ? i18nSrc_(enid) : null;
  var G = i18nzGloss_(src, en, p.dry ? 'ja' : lang);
  if (p.dry) return { data: { src: src, swap: swap, gloss: G.lines, gstat: G.stat, en_name: en ? en.name : '', sec: Math.round((Date.now() - t0) / 1000) } };
  var local = en ? (i18nGuideNames_()[i18nNameKey_(en.name)] || {})[lang] || '' : '';
  var r = i18nzWork_(src, lang, G, local, t0 + 330000);
  r.sec = Math.round((Date.now() - t0) / 1000);
  return { data: { src: src, swap: swap, gloss: G.lines, gstat: G.stat, local: local, tr: r } };
}

/* 有中文字的共用圖影換成外國客人版 */
function i18nzSwap_(src, lang) {
  var urls = [], md5 = {}, log = [];
  src.slides.forEach(function (s) { [s.video, s.img].forEach(function (u) { if (u && urls.indexOf(u) < 0) urls.push(u); }); });
  for (var i = 0; i < urls.length; i += 20) {
    var part = urls.slice(i, i + 20), rs = [];
    try { rs = UrlFetchApp.fetchAll(part.map(function (u) { return { url: u, headers: { Range: 'bytes=0-0' }, muteHttpExceptions: true }; })); } catch (e) { rs = []; }
    part.forEach(function (u, j) {
      var r = rs[j]; if (!r) return;
      var h = r.getHeaders(), m = '';
      Object.keys(h).forEach(function (k) { var lk = k.toLowerCase(); if (lk === 'x-ms-blob-content-md5' || lk === 'content-md5') m = m || String(h[k]); });
      md5[u] = m;
    });
  }
  src.slides.forEach(function (s, k) {
    var v = s.video && I18NZ_MD5[md5[s.video]];
    if (v && v.video && lang) { s.video_orig = s.video; s.video = I18NZ_CHEF + lang + '.mp4'; log.push((k + 1) + '：教學影片→熊主廚' + lang + '版'); }
    var g = s.img && I18NZ_MD5[md5[s.img]];
    if (g && g.img) { s.img_orig = s.img; s.img = g.img; log.push((k + 1) + '：烤熟測試圖→英文版'); }
    (s.imgs || []).forEach(function (u, j) {
      var d = u; try { d = decodeURIComponent(u); } catch (e) { }
      I18NZ_IMG.forEach(function (x) { if (d.indexOf(x[0]) >= 0) { s.imgs[j] = x[1]; log.push((k + 1) + '：插圖「' + x[0] + '」→英文版'); } });
    });
  });
  return log;
}

/* 詞彙表：中文品名 → 後台名稱對照表的英文名 → 現有詞典（i18n_items）的譯名；取放區用英文版同代碼的英文名查。只讀。 */
function i18nzZhKey_(s) { return rbTrim_(String(s || '')).replace(/^【[^】]*】\s*/, '').replace(/^[A-Z]{1,2}\d{2}\s+/, '').replace(/\s*[(（]已?下架[)）]/g, '').toLowerCase(); }
function i18nzZhEn_() {
  var d = trDict_({}).data, langs = d.langs || [], ie = -1, iz = -1, m = {};
  langs.forEach(function (l, i) { var s = String(l).trim(); if (ie < 0 && /^(英文|English)$/i.test(s)) ie = i; if (iz < 0 && /^繁體中文/.test(s)) iz = i; });
  if (iz < 0) iz = 0; if (ie < 0) return m;
  (d.rows || []).forEach(function (r) {
    var en = rbTrim_(r[3 + ie]).replace(/^【[^】]*】\s*/, ''), zh = i18nzZhKey_(r[3 + iz]); if (!en || !zh) return;
    var o = m[zh] = m[zh] || []; if (o.indexOf(en) < 0) o.push(en);
  });
  return m;
}
function i18nzGloss_(src, en, lang) {
  var ZE = {}; try { ZE = i18nzZhEn_(); } catch (e) { }
  var D = i18nDictRows_(), area = {}, lines = [], map = {}, seen = {}, stat = { items: 0, hit: 0, multi: 0, none: 0 };
  if (en) en.slides.forEach(function (s) { s.ings.forEach(function (g) { if (g.a && g.an) area[g.a] = g.an; }); });
  src.slides.forEach(function (s) {
    s.ings.forEach(function (g) {
      var core = i18nSplitName_(g.n).core, k = i18nzZhKey_(core);
      if (k && !seen['i|' + k]) {
        seen['i|' + k] = 1; stat.items++;
        var c = ZE[k] || [], tr = c.length === 1 ? i18nLook_(D, 'item', c[0], lang) : '';
        if (tr) { stat.hit++; lines.push('- ' + core + ' → ' + tr); map[core] = tr; } else if (c.length > 1) stat.multi++; else stat.none++;
      }
      if (g.a && g.an && !seen['a|' + g.a]) {
        seen['a|' + g.a] = 1;
        var atr = area[g.a] ? i18nLook_(D, 'area', area[g.a], lang) : '';
        if (atr) { lines.push('- ' + g.an + ' → ' + atr); map['@' + g.a] = atr; }
      }
    });
  });
  return { lines: lines, map: map, stat: stat };
}

function i18nzNumDiff_(sh, th, okx) {   /* okx：詞彙表譯名自帶的數字（例：8切り星口金）不算多 */   /* 同 i18nNumDiff_，但原文是中文：譯文多出來的數字要原文有對應的中文數字（一、兩、半…） */
  var a = i18nNums_(sh), b = i18nNums_(th), left = b.slice(), miss = [];
  a.forEach(function (x) { var j = left.indexOf(x); if (j < 0) miss.push(x); else left.splice(j, 1); });
  if (miss.length) return '原文的數字 ' + miss.join(',') + ' 譯文沒有（原文 ' + (a.join(',') || '無') + '／譯文 ' + (b.join(',') || '無') + '）';
  var pt = i18nPlain_(sh), extra = [], room = {};
  left.forEach(function (x) {
    if (okx && okx[x]) return;
    if (room[x] == null) room[x] = I18NZ_NUMW[x] ? (pt.match(I18NZ_NUMW[x]) || []).length : 0;
    if (room[x] > 0) room[x]--; else extra.push(x);
  });
  return extra.length ? '譯文多了數字 ' + extra.join(',') + '（原文沒有對應的字）' : '';
}
function i18nzCheck_(s, t, lang, okx) {
  var bad = [];
  if (!t) return ['少了這一頁'];
  if (s.h && !i18nPlain_(t.h) && !i18nPh_(t.h)) bad.push('內文是空的');
  var d = i18nzNumDiff_(s.h, t.h, okx); if (d) bad.push(d);
  if (i18nPh_(s.h) !== i18nPh_(t.h)) bad.push('圖片位置不一致');
  var tn = i18nNums_(t.t); if (!i18nNums_(s.t).every(function (x) { var j = tn.indexOf(x); if (j < 0) return false; tn.splice(j, 1); return true; })) bad.push('標題數字不見了');
  if (s.t && !String(t.t || '').trim()) bad.push('標題是空的');
  if (lang !== 'ja' && /[一-鿿]/.test(i18nPlain_(t.h) + String(t.t || ''))) bad.push('譯文還有中文字');   /* 日文本來就有漢字，不檢查 */
  return bad;
}
function i18nzFixed_(s, lang) {
  var f = I18N_FIXED[0], tr = f && f.tr[lang], h = String(s.h || '');
  if (!tr || s.t !== I18NZ_FIXED.t || h.indexOf(I18NZ_FIXED.zh[0]) < 0) return null;
  I18NZ_FIXED.zh.forEach(function (x, j) { h = h.split(x).join(tr.s[j]); });
  return { t: tr.t, h: h };
}

function i18nzSystem_(lang, src, G, local, tt) {
  return ['You translate the in-store tablet recipe of 自己做烘焙聚樂部 (a DIY baking studio in Taiwan; customers bake by themselves, following the tablet page by page).',
    'Source: the store\'s official Traditional Chinese (Taiwan) recipe. Target: ' + I18N_LABEL[lang] + '.',
    'Target style:', I18N_RULES[lang].map(function (r) { return '* ' + r; }).join('\n'),
    'General rules:',
    '* Translate faithfully from the Chinese; do not add, drop, merge or reorder sentences. Keep numbering (1. 2. 3.) and line breaks.',
    '* "h" is HTML. Keep only these tags: <p>, <p class="c">, <br>, <b>, <u>, <i>, <font color="...">. Keep each tag around the corresponding translated words. Keep placeholders like [[IMG1]] exactly where they are.',
    '* Never change any number, weight, time, temperature, speed or count; write them with Arabic numerals exactly as in the source. Chinese number words (一、兩、半) may be written naturally in the target language.',
    '* Keep emojis and symbols (●, ※, ⭐, ►, 【】, ［］). Chinese quotes 「」 may become the normal quotes of the target language. Translate 小教室 / DIY小教室 as a short local label that keeps the word "DIY".',
    '* Taiwan baking terms: 克林姆 = custard cream (made from custard powder; NOT whipped cream); 動植鮮 / 動植鮮奶油 = whipped topping cream (dairy + vegetable fat); 鮮奶油 / 動物鮮奶油 = whipping cream; 平匙 = a level spoonful; 打發 = whip; 拌勻 = mix well; 過篩 = sift; 刮刀 = rubber spatula; 抹刀 = frosting spatula; 擠花袋 = piping bag; 花嘴 = piping tip; 攪拌機 = electric mixer; 取放區 / 取料區 = the pick-up area in the store; 計時器 = timer; 自己人 = the store\'s membership program.',
    '* When the text mentions an ingredient, tool, mold, area or box from the glossary below, use exactly that translation.',
    '* "t" is the page title (short). Translate it the same way every time it repeats.',
    local ? '* The dessert name in this language is: ' + local + ' (use it wherever the dessert name appears; return "" for "name").' : '* "name" is the Chinese dessert name: translate it naturally into the target language (translate notes like (葷) = contains meat-based gelatin, (全素) = vegan).',
    '* "pres" is the storage advice; translate it keeping line breaks (\\n). If "name" or "pres" is given as "", return "".',
    '* Return every k exactly as given.',
    'Glossary (Chinese → ' + I18N_LABEL[lang] + '):', G.lines.join('\n') || '(none)',
    'Page titles (already translated — use exactly):', Object.keys(tt || {}).map(function (k) { return '- ' + k + ' → ' + tt[k]; }).join('\n') || '(none)'].join('\n');
}
function i18nzTitles_(src, titles, lang, cred, local, U) {
  if (!titles.length) return {};
  var sys = ['You translate the page titles of one DIY baking recipe (customers follow a tablet page by page) from Traditional Chinese (Taiwan) into ' + I18N_LABEL[lang] + '.',
    'Short titles. Keep numbers (1, 2, 2-1) exactly. The same Chinese wording must always get the same translation; similar titles with a number (抹面1 / 抹面2) must share the same words.',
    'Dessert: ' + src.name + (local ? ' (local name: ' + local + ')' : '') + '. Return every k exactly as given.'].join('\n');
  var schema = { type: 'object', properties: { items: { type: 'array', items: { type: 'object', properties: { k: { type: 'string' }, tr: { type: 'string' } }, required: ['k', 'tr'], additionalProperties: false } } }, required: ['items'], additionalProperties: false };
  var r = i18nzOne_(sys, JSON.stringify({ titles: titles.map(function (t, j) { return { k: String(j), zh: t }; }) }), schema, cred, U), m = {};
  (r.json.items || []).forEach(function (it) { var t = titles[Number(it.k)]; if (t && rbTrim_(it.tr)) m[t] = rbTrim_(it.tr); });
  return m;
}
function i18nzUse_(resp, U) { try { var j = JSON.parse(resp.getContentText()); if (j && j.usage) { U.inp += j.usage.input_tokens || 0; U.out += j.usage.output_tokens || 0; U.calls++; } } catch (e) { } }
function i18nzOne_(sys, user, schema, cred, U) {
  for (var n = 0; n < 3; n++) {
    var r, resp = null;
    try { resp = UrlFetchApp.fetch(I18N_AI_URL, i18nAiReq_(sys, user, schema, cred)); i18nzUse_(resp, U); r = i18nAiRes_(resp, cred); }
    catch (e) { r = { ok: false, err: '連不到翻譯服務：' + errMsg_(e) }; }
    if (r.ok) return r;
    if (r.retryModel) { cred.model = I18N_MODELS[1]; continue; }
    if (n < 2) Utilities.sleep(2000 * (n + 1)); else throw fail_(r.err, 'ai');
  }
}

function i18nzWork_(src, lang, G, local, dl) {
  var cred = i18nCred_(); cred.model = I18N_AUTO_MODEL;
  var U = { inp: 0, out: 0, calls: 0 }, out = { name: '', pres: '', slides: [] }, fixed = {}, todo = [];
  src.slides.forEach(function (s, i) { var f = i18nzFixed_(s, lang); if (f) { fixed[i] = f; out.slides[i] = f; } else todo.push(i); });
  var tt = {}, ut = [], seenT = {};
  todo.forEach(function (i) { var t = src.slides[i].t; if (t && !seenT[t]) { seenT[t] = 1; ut.push(t); } });
  tt = i18nzTitles_(src, ut, lang, cred, local, U);
  Object.keys(fixed).forEach(function (i) { tt[src.slides[i].t] = fixed[i].t; });
  var sys = i18nzSystem_(lang, src, G, local, tt), model = cred.model;
  var chunks = [], cur = [], n = 0;
  todo.forEach(function (i) {
    var len = (src.slides[i].t || '').length + (src.slides[i].h || '').length;
    if (cur.length && (cur.length >= I18N_CHUNK || n + len > I18N_CHUNK_CHARS)) { chunks.push(cur); cur = []; n = 0; }
    cur.push(i); n += len;
  });
  if (cur.length) chunks.push(cur);
  var needName = !local;
  var reqOf = function (ix, first, note) {
    var u = { name: first && needName ? src.name : '', pres: first ? src.pres : '', slides: ix.map(function (i) { return { k: i, t: src.slides[i].t, h: src.slides[i].h }; }) };
    return i18nAiReq_(sys + (note ? '\n' + note : ''), 'Translate this JSON. Return the same structure.\n' + JSON.stringify(u), I18N_SCHEMA, cred);
  };
  var take = function (r, ix, first) {
    if (!r || !r.ok) return; model = r.model || model;
    if (first) { if (r.json.name && needName) out.name = r.json.name; if (r.json.pres && src.pres) out.pres = r.json.pres; }
    (r.json.slides || []).forEach(function (s) { if (ix.indexOf(s.k) >= 0) out.slides[s.k] = { t: s.t, h: s.h }; });
  };
  var tlimit = dl - 110000;
  var run = function (jobs) {
    for (var w = 0; w < jobs.length; w += I18N_WAVE) {
      if (Date.now() > tlimit) return;
      var wave = jobs.slice(w, w + I18N_WAVE), rs = null;
      try { rs = UrlFetchApp.fetchAll(wave.map(function (j) { return reqOf(j.ix, j.first, j.note); })); } catch (e) { rs = null; }
      wave.forEach(function (j, q) {
        var r = null;
        if (rs) { i18nzUse_(rs[q], U); r = i18nAiRes_(rs[q], cred); }
        if (r && r.retryModel) { cred.model = I18N_MODELS[1]; r = null; }
        if (!r || !r.ok) { try { var one = UrlFetchApp.fetch(I18N_AI_URL, reqOf(j.ix, j.first, j.note)); i18nzUse_(one, U); r = i18nAiRes_(one, cred); } catch (e2) { r = { ok: false }; } }
        take(r, j.ix, j.first);
      });
    }
  };
  run(chunks.map(function (ix, k) { return { ix: ix, first: k === 0 }; }));
  var fixT = function (ix) { ix.forEach(function (i) { var s = src.slides[i]; if (out.slides[i] && tt[s.t]) out.slides[i].t = tt[s.t]; }); };
  fixT(todo);
  var probs = function (ix) { var p = {}; ix.forEach(function (i) { var b = i18nzCheck_(src.slides[i], out.slides[i], lang, okx); if (b.length) p[i] = b; }); return p; };
  var okx = {}; G.lines.forEach(function (l) { i18nNums_(String(l).split(' → ').slice(1).join(' → ')).forEach(function (q) { okx[q] = 1; }); });
  if (local) i18nNums_(local).forEach(function (q) { okx[q] = 1; });
  var P = probs(todo), first = Object.keys(P).length, redo = Object.keys(P).map(Number), stage = { hard_first: first };
  if ((needName && !out.name) || (src.pres && !out.pres)) redo.unshift(-1);
  if (redo.length && Date.now() < tlimit) {
    run(redo.map(function (i) { return i < 0 ? { ix: [], first: true } : { ix: [i], first: false, note: 'IMPORTANT: a previous translation of this page had problems: ' + P[i].join('; ') + '. Keep every number and [[IMG]] placeholder exactly.' }; }));
    fixT(todo); P = probs(todo);
  }
  var rv = { block: [], minor: 0, fixed: 0, first_block: 0 };
  if (!Object.keys(P).length && todo.every(function (i) { return !!out.slides[i]; })) {
    try {
      rv = i18nzReview_(src, out, todo, needName, !!src.pres, lang, G, cred, U); rv.first_block = rv.block.length; rv.first_list = rv.block.slice();
      if (rv.block.length && Date.now() < dl - 70000) {
        var bp = [], why = {};
        rv.block.forEach(function (x) { var k = Number(x.k); if (!(k >= 0 && k < src.slides.length) || todo.indexOf(k) < 0) return; if (!why[k]) { why[k] = []; bp.push(k); } why[k].push(x.what); });
        tlimit = dl - 40000;
        run(bp.map(function (i) { return { ix: [i], first: false, note: 'IMPORTANT: a reviewer found these problems on this page: ' + why[i].join('; ') + '. Fix them. Keep every number and [[IMG]] placeholder exactly.' }; }));
        fixT(bp); P = probs(todo);
        if (!Object.keys(P).length && Date.now() < dl - 30000) {
          var rv2 = i18nzReview_(src, out, bp, false, false, lang, G, cred, U), other = rv.block.filter(function (x) { return bp.indexOf(Number(x.k)) < 0; });
          rv.block = other.concat(rv2.block); rv.fixed = bp.length;
        }
      }
    } catch (e) { rv.err = errMsg_(e); }
  }
  if (local) out.name = local + '｜' + src.name; else if (out.name) out.name = out.name + '｜' + src.name;
  var done = src.slides.every(function (s, i) { return !!out.slides[i]; });
  var status = !done ? 'fail' : Object.keys(P).length || rv.block.length || rv.err ? 'review' : 'ok';
  return { status: status, model: model, name: out.name, pres: out.pres, slides: out.slides, titles: tt, fixed: Object.keys(fixed).map(Number),
    hard: P, stage: stage, review: rv, usage: U, translated: todo.length };
}

function i18nzReview_(src, out, pages, withName, withPres, lang, G, cred, U) {
  var sys = ['You are the independent second reviewer of a translated in-store tablet recipe of a DIY baking studio in Taiwan (customers bake by themselves, following the tablet page by page).',
    'You did not translate it. Compare each page\'s Traditional Chinese source (zh_t, zh_h) with the ' + I18N_LABEL[lang] + ' translation (tr_t, tr_h). This is the last check before customers see it.',
    'Report severity "block" only for problems that would make a customer do something wrong or not understand:',
    '* a sentence missing or added; wrong meaning (verbs like mix / whisk / fold / pipe / bake / cool / refrigerate / freeze / sift, ingredient, tool, container, direction; "先不…" = "not yet");',
    '* a number, weight, time, temperature, speed, count or unit that differs from the Chinese (Thai must use Arabic numerals 0-9);',
    '* a safety warning (hot, gloves, knife, freezer) missing; a Chinese or English sentence left untranslated, garbled text, broken HTML;',
    '* a price turned into a local currency (đồng, 円, 원, บาท, rupiah); "10元硬幣" means a Taiwan NT$10 coin;',
    '* an ingredient/tool name that contradicts the glossary or the Chinese (克林姆 custard cream is NOT whipped cream; 動植鮮 is whipped topping cream; 攪拌機 electric mixer; 柚子 pomelo is not yuzu; 麵糊 batter is not flour).',
    'Everything else (slightly stiff wording, could be more natural) is "minor". Do not report differences in style that keep the meaning.',
    'k is the page number given. Use k = -1 for problems in the dessert name or storage advice. Return an empty list if there are no problems. Write "what" briefly in Traditional Chinese.',
    'Glossary (Chinese → ' + I18N_LABEL[lang] + '):', G.lines.join('\n') || '(none)'].join('\n');
  var groups = [], cur = [], n = 0;
  pages.forEach(function (i) {
    var len = (src.slides[i].h || '').length + (out.slides[i] ? String(out.slides[i].h || '').length : 0);
    if (cur.length && (cur.length >= 8 || n + len > 12000)) { groups.push(cur); cur = []; n = 0; }
    cur.push(i); n += len;
  });
  if (cur.length) groups.push(cur);
  if (!groups.length && (withName || withPres)) groups.push([]);
  var reqs = groups.map(function (g, gi) {
    var u = { pages: g.map(function (i) { var t = out.slides[i] || {}; return { k: i, zh_t: src.slides[i].t, zh_h: src.slides[i].h, tr_t: t.t || '', tr_h: t.h || '' }; }) };
    if (gi === 0 && withName) { u.name_zh = src.name; u.name_tr = out.name; }
    if (gi === 0 && withPres) { u.pres_zh = src.pres; u.pres_tr = out.pres; }
    return i18nAiReq_(sys, 'Review this JSON.\n' + JSON.stringify(u), I18N_REVIEW_SCHEMA, cred);
  });
  var res = { block: [], minor: 0, minors: [] };
  for (var w = 0; w < reqs.length; w += I18N_WAVE) {
    var wave = reqs.slice(w, w + I18N_WAVE), rs = null;
    try { rs = UrlFetchApp.fetchAll(wave); } catch (e) { rs = null; }
    wave.forEach(function (rq, q) {
      var r = null;
      if (rs) { i18nzUse_(rs[q], U); r = i18nAiRes_(rs[q], cred); }
      if (!r || !r.ok) { try { var one = UrlFetchApp.fetch(I18N_AI_URL, rq); i18nzUse_(one, U); r = i18nAiRes_(one, cred); } catch (e2) { r = { ok: false, err: errMsg_(e2) }; } }
      if (!r.ok) throw fail_('回讀失敗：' + r.err, 'ai');
      (r.json.issues || []).forEach(function (x) {
        if (x.severity === 'block') res.block.push({ k: Number(x.k), what: String(x.what || '').slice(0, 200) });
        else { res.minor++; if (res.minors.length < 30) res.minors.push({ k: Number(x.k), what: String(x.what || '').slice(0, 160) }); }
      });
    });
  }
  return res;
}
ACTIONS.i18nZTry = function (p, a) { return i18nzTry_(p, a); };
NOAUTH.i18nZTry = 1; READS.i18nZTry = 1;   /* 通行碼在函式裡自己驗；不寫任何分頁 */

/* 動作 i18nZReview（要投遞通行碼）：同一位 AI 檢查員，拿中文原文檢查「任何一份」譯文（現行英文轉譯版、試做版都可），讓兩邊用同一把尺比。
   {zid, enid, lang, pages:[{zh_t, zh_h, tr_t, tr_h}], name_zh, name_tr, pres_zh, pres_tr} → {review, usage}；k＝pages 的順序。不寫任何分頁。 */
function i18nzReviewAny_(p) {
  i18nWriteKey_(p);
  var lang = String(p.lang || '').trim(), zid = String(p.zid || '').trim().toLowerCase(), enid = String(p.enid || '').trim().toLowerCase();
  if (I18N_LANGS.indexOf(lang) < 0) throw fail_('lang 不對', 'bad');
  var pg = i18nFetchPages_([zid])[0]; if (pg.code !== 200) throw fail_('後台回 HTTP ' + pg.code, 'rbnet');
  var G = i18nzGloss_(i18nParse_(pg.text, zid), enid ? i18nSrc_(enid) : null, lang);
  var fs = { name: String(p.name_zh || ''), pres: String(p.pres_zh || ''), slides: [] }, fo = { name: String(p.name_tr || ''), pres: String(p.pres_tr || ''), slides: [] }, ix = [];
  (p.pages || []).forEach(function (x, i) { fs.slides[i] = { t: String(x.zh_t || ''), h: String(x.zh_h || '') }; fo.slides[i] = { t: String(x.tr_t || ''), h: String(x.tr_h || '') }; ix.push(i); });
  var cred = i18nCred_(); cred.model = I18N_AUTO_MODEL;
  var U = { inp: 0, out: 0, calls: 0 }, rv = i18nzReview_(fs, fo, ix, !!fo.name, !!fo.pres, lang, G, cred, U);
  return { data: { review: rv, usage: U } };
}
ACTIONS.i18nZReview = function (p, a) { return i18nzReviewAny_(p, a); };
NOAUTH.i18nZReview = 1; READS.i18nZReview = 1;
