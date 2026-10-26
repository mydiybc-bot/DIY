/* ================= 🌏 平板外語食譜：英文食譜 → 日／韓／越／印尼／泰（2026-10-24） =================
   經營者 10/24：外國客人在店內平板看自己語言的食譜；「不要在自己做食譜後台新增任何語言的食譜」。
   → 翻譯稿存在本試算表，平板開我們自己的網頁（static/tablet-recipe.html，第 2 段）。後台只讀、不寫。

   原文：後台「台灣區、英文、啟用中」的食譜（食譜演化清單 recipe_lineage 篩出來，目前 39 支），
         讀平板頁 /steps/recipedetails/{id}（跟店內平板看到的一樣：一頁一頁、影片、圖片、食材分區、容器）。
   翻譯：GAS 呼叫訓練站 /api/recipe/translate（Claude；身分密碼由本程式從 dim_role 的「管理者」讀，不經人手）。
         食材／工具／取放區／單位走共用詞典 i18n_items（每語言只翻一次，所有食譜同一個譯名；越南文先套後台名稱對照表）；
         步驟內文每 4 頁一段並行翻，自動檢查（頁數、每頁數字、圖片位置、標題數字）不過 → 單頁重翻一次 → 還不過標「待檢查」，平板不顯示。
   分頁：i18n_src（原文快照）、i18n_rcp（每款×語言的翻譯稿）、i18n_items（詞典）、i18n_log（紀錄）。JSON 超過一格上限就切成好幾格（j1…j8）。
   動作（POST，都不用身分密碼）：
     i18nStatus  只讀：幾款幾語完成、待檢查清單、最近紀錄（不含食譜內容）
     i18nList    平板：食譜清單（要平板密碼 tpw＝權限中心 tablet_i18n）
     i18nGet     平板：一支食譜某語言的完整內容（要 tpw）
     i18nPreview 主廚／管理者（或唯讀預覽通行碼 pkey）：原文與譯文對照（核准前一頁一頁看；待核准、待檢查都看得到）
     i18nApprove 主廚／管理者：核准放行（undo:1 收回）。**沒核准的平板看不到**（經營者 10/24：「完成後還是要一一打開比對，確認沒有問題才可以放行」）；
                 原文改版重翻 → 核准自動失效（appr_fp 對不上），要重新看過
     i18nKick    做下一件待翻譯的事：沒開自動（指令碼屬性 I18N_AUTO≠1）時只做指定的 id＋lang；同一件 6 分鐘內不會兩個人同時做
   排程：i18nDaily 每天 06:30（i18nKick 第一次被呼叫時自動建立，不用到編輯器按）：更新原文（後台有改才重抓、內容真的變了才重翻）；
         I18N_AUTO＝1 時接著翻待翻譯的，做到 5 分鐘為止。
   ⚠ 只讀後台；不寫後台任何東西。 */

var I18N_LANGS = ['ja', 'ko', 'vi', 'id', 'th'];
var I18N_LABEL = { ja: 'Japanese 日本語', ko: 'Korean 한국어', vi: 'Vietnamese Tiếng Việt', id: 'Indonesian Bahasa Indonesia', th: 'Thai ภาษาไทย' };
var I18N_RULES = {
  ja: ['Natural Japanese for customers following a recipe at a baking studio. Use です/ます style, short and clear (e.g. 「泡立て器で20回混ぜます。」).',
    'Use katakana for Western ingredient and tool names where natural (ボウル, ゴムベラ, 絞り袋).'],
  ko: ['Natural Korean for customers following a recipe at a baking studio. Use polite imperative -세요 style, short and clear (e.g. "거품기로 20번 저어 주세요.").'],
  vi: ['Tiếng Việt ngắn gọn, dễ hiểu cho khách đang tự làm bánh theo công thức (câu mệnh lệnh: "Cho bột vào thố, khuấy 20 lần.").'],
  id: ['Bahasa Indonesia yang sederhana dan jelas untuk pelanggan yang membuat kue sendiri (kalimat perintah: "Aduk 20 kali dengan whisk.").'],
  th: ['Simple polite Thai for customers baking by themselves (imperative, short). Always use Arabic numerals (0-9), never Thai digits.']
};
var I18N_MODELS = ['claude-opus-5-5', 'claude-sonnet-5-5'];
var I18N_AI_URL = 'https://diybc-training.onrender.com/api/recipe/translate';
var I18N_AUTH_API = 'https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec';
var I18N_AUTH_DASH = 'tablet_i18n';
var I18N_GUIDE_API = 'https://script.google.com/macros/s/AKfycbzKmiZv95tW2b_1oAr-DiJf90cmpadfkb0MI2WBikqD_uVE6uUSUf_vTNa6KcbjMYc4BQ/exec?action=data';
var I18N_J = 8, I18N_CELL = 45000;   /* JSON 切 8 格、每格 45,000 字（一格上限 50,000） */
var I18N_SRC_HEAD = ['id', 'name', 'upd', 'fp', 'slides', 'cover', 'state', 'err', 'at'];
var I18N_RCP_HEAD = ['key', 'id', 'lang', 'fp', 'status', 'tries', 'model', 'check', 'sec', 'at', 'appr_by', 'appr_at', 'appr_fp'];
var I18N_ITEM_HEAD = ['kind', 'en', 'ja', 'ko', 'vi', 'id', 'th', 'how', 'at'];
var I18N_LOG_HEAD = ['at', 'by', 'act', 'id', 'lang', 'result', 'msg', 'sec'];
var I18N_CHUNK = 4, I18N_CHUNK_CHARS = 5000, I18N_WAVE = 6, I18N_BUDGET = 270000, I18N_TRIES = 3;

/* ---------------- 共用 ---------------- */
function i18nSheet_(name, head) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var all = name === 'i18n_src' || name === 'i18n_rcp' ? head.concat(i18nJCols_()) : head;
    sh.getRange(1, 1, 1, all.length).setValues([all]);
    sh.getRange(1, 1, 1000, all.length).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}
function i18nJCols_() { var a = []; for (var i = 1; i <= I18N_J; i++) a.push('j' + i); return a; }
/* 只讀前面的欄（不讀大 JSON） */
function i18nMeta_(name, head) {
  var sh = i18nSheet_(name, head), n = sh.getLastRow() - 1, out = [];
  if (n < 1) return { sh: sh, rows: out };
  sh.getRange(2, 1, n, head.length).getDisplayValues().forEach(function (r, i) {
    var o = { _r: i + 2 }; head.forEach(function (h, j) { o[h] = r[j]; }); out.push(o);
  });
  return { sh: sh, rows: out };
}
function i18nReadJ_(sh, row, head) {
  var v = sh.getRange(row, head.length + 1, 1, I18N_J).getValues()[0].join('');
  return v ? JSON.parse(v) : null;
}
function i18nCut_(s) {
  if (s.length > I18N_J * I18N_CELL) throw fail_('內容太長，存不下（' + s.length + ' 字）', 'big');
  var out = []; for (var i = 0; i < I18N_J; i++) out.push(s.slice(i * I18N_CELL, (i + 1) * I18N_CELL));
  return out;
}
/* 依 key 欄找列（沒有就加在最後），寫前面欄＋（可選）JSON；要在鎖裡呼叫 */
function i18nUpsert_(name, head, keyName, obj, json) {
  var m = i18nMeta_(name, head), hit = m.rows.filter(function (r) { return r[keyName] === obj[keyName]; })[0];
  var row = hit ? hit._r : m.sh.getLastRow() + 1;
  var vals = head.map(function (h) { return obj[h] == null ? (hit ? hit[h] : '') : String(obj[h]); });
  m.sh.getRange(row, 1, 1, head.length).setNumberFormat('@').setValues([vals]);
  if (json != null) m.sh.getRange(row, head.length + 1, 1, I18N_J).setNumberFormat('@').setValues([i18nCut_(JSON.stringify(json))]);
  return row;
}
var I18N_HELD = false;   /* 已經在全系統的鎖裡（一般寫入路徑，例：i18nApprove）→ 不要再拿一次 */
function i18nLock_(fn) {
  if (I18N_HELD) return fn();
  var lk = LockService.getScriptLock(); lk.waitLock(30000);
  /* 放開鎖之前一定要 flush：Google 試算表的寫入會先暫存，沒送出就放鎖，下一個程序讀到舊資料會把剛寫的蓋掉
     （2026-10-24 四支同時翻譯，詞典的韓／越／印／泰譯名被互相蓋掉） */
  I18N_HELD = true;   /* 鎖裡面再呼叫 i18nLock_（例：寫紀錄）就直接做，不會提早放鎖 */
  try { return fn(); } finally { I18N_HELD = false; try { SpreadsheetApp.flush(); } catch (e) { } lk.releaseLock(); }
}
function i18nLog_(by, act, id, lang, result, msg, sec) {
  try {
    i18nLock_(function () {
      var sh = i18nSheet_('i18n_log', I18N_LOG_HEAD);
      sh.appendRow([now_(), by || '', act, id || '', lang || '', result, String(msg || '').slice(0, 500), sec == null ? '' : sec]);
    });
  } catch (e) { /* 紀錄寫不進去不影響主流程 */ }
}
function i18nHash_(s) { return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)).slice(0, 22); }
function i18nProp_(k, v) { var P = PropertiesService.getScriptProperties(); if (v === undefined) return P.getProperty(k) || ''; P.setProperty(k, String(v)); return v; }

/* ---------------- 解析平板頁（純字串處理，node 也能跑測試） ---------------- */
/* 從 start（某個 <div 的位置）找配對的 </div>，回傳 [內容起點, 內容終點] */
function i18nDivSpan_(html, start) {
  var open = html.indexOf('>', start) + 1, re = /<div\b|<\/div>/gi, depth = 1, m;
  re.lastIndex = open;
  while ((m = re.exec(html))) {
    if (m[0].charAt(1) === '/') { depth--; if (!depth) return [open, m.index]; }
    else depth++;
  }
  return [open, html.length];
}
function i18nTxt_(s) {   /* 去標籤、保留換行（<br>→換行），每行整理空白 */
  return rbDecode_(String(s || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, ' ')).split('\n').map(rbTrim_).join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
function i18nParse_(html, id) {
  var o = { id: id, name: '', cover: '', pres: '', areaPre: '', areaHas: [], slides: [] };
  var t = (html.match(/<title>([^<]*)<\/title>/i) || ['', ''])[1];
  o.name = rbTrim_(rbDecode_(t).replace(/\s*-\s*自己做烘焙聚樂部\s*$/, ''));
  o.cover = ((html.match(/url\('(https:\/\/diybcstorage\.blob\.core\.windows\.net\/recipeimages\/[^']+)'\)/) || [])[1]) || '';
  o.areaPre = ((html.match(/var\s+fn\s*=\s*"([0-9a-f-]{36})"/i) || [])[1]) || '';
  var hr = /id="f_([0-9a-f-]{36})"/gi, hm; while ((hm = hr.exec(html))) o.areaHas.push(hm[1]);
  var inner = html.indexOf('class="carousel-inner"');
  if (inner < 0) throw fail_('平板頁找不到食譜頁面（後台可能改版）', 'rbpage');
  /* 步驟標題：在第一個「步驟清單」裡（每頁都有一份一樣的） */
  var titles = {}, lv = html.indexOf('class="listView"');
  if (lv >= 0) {
    var lvs = html.lastIndexOf('<div', lv), sp = i18nDivSpan_(html, lvs), block = html.slice(sp[0], sp[1]);
    var tr = /data-slide-to="(\d+)"\s*>\s*<span>[^<]*<\/span>([\s\S]*?)<\/div>/gi, tm;
    while ((tm = tr.exec(block))) titles[tm[1]] = i18nTxt_(tm[2]);
  }
  /* 每一頁 */
  var re = /<div class="carousel-item\b([^"]*)"[^>]*>/gi, m, starts = [];
  re.lastIndex = inner;
  while ((m = re.exec(html))) starts.push({ at: m.index, cls: m[1] });
  starts.forEach(function (s, idx) {
    var sp2 = i18nDivSpan_(html, s.at), ch = html.slice(sp2[0], sp2[1]);
    if (/hasOption/.test(s.cls)) {
      var sl = { k: o.slides.length, t: titles[String(idx)] || '', video: '', img: '', h: '', imgs: [], ings: [] };
      sl.video = ((ch.match(/<source\b[^>]*data-src="([^"]+)"/i) || [])[1]) || '';
      var it = (ch.match(/<img\b[^>]*class="titleImage"[^>]*>/i) || [''])[0]; sl.img = it ? rbAttr_(it, 'src') : '';
      var rc = ch.indexOf('class="recipeContent"');
      if (rc >= 0) { var rs = ch.lastIndexOf('<div', rc), sp3 = i18nDivSpan_(ch, rs), cl = i18nClean_(ch.slice(sp3[0], sp3[1])); sl.h = cl.h; sl.imgs = cl.imgs; }
      var ir = /<div class="ingred"([^>]*)>([\s\S]*?)<\/div>/gi, g;
      while ((g = ir.exec(ch))) {
        var tag = '<x' + g[1] + '>', parts = rbDecode_(g[2].replace(/<[^>]+>/g, ' ')).split(/\n/).map(rbTrim_).filter(Boolean);
        if (parts.length < 3) { var mm = parts.join(' ').match(/^(.+?)\s+([\d.\/]+)\s*(\S*)$/); parts = mm ? [mm[1], mm[2], mm[3]] : [parts.join(' '), '', '']; }
        sl.ings.push({ n: parts[0], q: parts[1], u: parts.slice(2).join(' '), c: rbAttr_(tag, 'data-container'), a: rbAttr_(tag, 'data-cateshort'), an: rbAttr_(tag, 'data-cate') });
      }
      o.slides.push(sl);
    } else if (/recipeimages\//.test(ch) && /position:\s*fixed;\s*bottom:\s*0/i.test(ch)) {
      var pi = ch.search(/white-space:\s*pre-wrap/i);
      if (pi >= 0) { var ps = ch.lastIndexOf('<div', pi), sp4 = i18nDivSpan_(ch, ps); o.pres = i18nTxt_(ch.slice(sp4[0], sp4[1]).replace(/\r?\n/g, '<br>')); }
    }
  });
  if (!o.slides.length) throw fail_('平板頁沒有任何步驟（後台可能改版）', 'rbpage');
  return o;
}
/* 內文 HTML → 只留 p／br／b／u／i／font color；圖片換成 [[IMGn]]；置中段落 p class="c" */
function i18nClean_(html) {
  var imgs = [], out = '', stack = [], blk = 0;
  var hex = function (c) {
    c = String(c || '').trim(); var m = c.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i);
    if (m) return '#' + [m[1], m[2], m[3]].map(function (x) { return ('0' + Number(x).toString(16)).slice(-2); }).join('');
    return /^#[0-9a-f]{3,6}$/i.test(c) ? c.toLowerCase() : /^[a-z]+$/i.test(c) ? c.toLowerCase() : '';
  };
  var css = function (st, k) { var m = String(st || '').match(new RegExp('(?:^|;)\\s*' + k + '\\s*:\\s*([^;]+)', 'i')); return m ? m[1].trim() : ''; };
  var re = /<(\/?)([a-z0-9]+)\b([^>]*)>|([^<]+)/gi, m;
  while ((m = re.exec(String(html || '')))) {
    if (m[4] != null) { out += m[4].replace(/[\s ]+/g, ' ').replace(/&nbsp;/g, ' '); continue; }
    var close = m[1] === '/', tag = m[2].toLowerCase(), at = m[3] || '', x = '<x ' + at + '>';
    if (tag === 'br') { out += '<br>'; continue; }
    if (tag === 'img') { imgs.push(rbAttr_(x, 'src')); out += '[[IMG' + imgs.length + ']]'; continue; }
    if (tag === 'p' || tag === 'div') {
      if (!close) { var c = /text-align:\s*center/i.test(rbAttr_(x, 'style')); stack.push({ t: 'blk' }); if (!blk) out += c ? '<p class="c">' : '<p>'; blk++; }
      else { while (stack.length && stack[stack.length - 1].t !== 'blk') out += stack.pop().end; if (stack.length) stack.pop(); if (blk) { blk--; out += blk ? '<br>' : '</p>'; } }
      continue;
    }
    var open = '', end = '';
    if (!close) {
      var st = rbAttr_(x, 'style');
      if (tag === 'b' || tag === 'strong' || /font-weight:\s*(bold|bolder|[6-9]00)/i.test(st)) { open += '<b>'; end = '</b>' + end; }
      if (tag === 'u' || /text-decoration[^;]*underline/i.test(st)) { open += '<u>'; end = '</u>' + end; }
      if (tag === 'i' || tag === 'em') { open += '<i>'; end = '</i>' + end; }
      var col = hex(tag === 'font' ? rbAttr_(x, 'color') : '') || hex(css(st, 'color'));
      if (col) { open += '<font color="' + col + '">'; end = '</font>' + end; }
      stack.push({ t: tag, end: end }); out += open;
    } else {
      for (var i = stack.length - 1; i >= 0; i--) { if (stack[i].t === 'blk') break; if (stack[i].t === tag) { while (stack.length > i) out += stack.pop().end; break; } }
    }
  }
  while (stack.length) { var s = stack.pop(); if (s.t !== 'blk') out += s.end; }
  while (blk-- > 0) out += '</p>';
  out = out.replace(/<(b|u|i)>\s*<\/\1>/g, '').replace(/<font color="[^"]*">\s*<\/font>/g, '')
    .replace(/<p( class="c")?>(\s|<br>)*<\/p>/g, '').replace(/<p( class="c")?>(\s*<br>)+/g, '<p$1>').replace(/(<br>\s*)+<\/p>/g, '</p>')
    .replace(/(<br>\s*){3,}/g, '<br><br>').replace(/^(\s*<br>)+|(<br>\s*)+$/g, '').replace(/ {2,}/g, ' ').trim();
  return { h: out, imgs: imgs };
}
/* 原文指紋：內容（不含 cover 版本號）一樣＝不重翻 */
function i18nFp_(src) {
  return i18nHash_(JSON.stringify([src.name, src.pres, src.slides.map(function (s) { return [s.t, s.h, s.video.replace(/\?.*$/, ''), s.img.replace(/\?.*$/, ''), s.imgs, s.ings.map(function (g) { return [g.n, g.q, g.u, g.c, g.a]; })]; })]));
}

/* ---------------- 自動檢查（純函式） ---------------- */
function i18nPlain_(h) { return i18nTxt_(String(h || '').replace(/\[\[IMG\d+\]\]/g, ' ')); }
function i18nNums_(h) {
  var t = i18nPlain_(h).replace(/[０-９]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xfee0); });
  return (t.match(/\d+(?:[.,]\d+)?/g) || []).map(function (x) { return x.replace(',', '.'); }).sort();
}
/* 數字比對：原文的數字一個都不能少、不能變；譯文多出來的數字，原文要有對應的英文數量詞（a／one／two／second／both／half…，日韓泰常寫成「1本」「2枚目」） */
var I18N_NUMW = { '1': /\b(a|an|one|once|single|first|each|half)\b/gi, '2': /\b(two|twice|second|both|pair|half|couple)\b/gi, '3': /\b(three|third|thrice)\b/gi,
  '4': /\b(four|fourth|quarter)\b/gi, '5': /\b(five|fifth)\b/gi, '6': /\b(six|sixth)\b/gi, '7': /\b(seven|seventh)\b/gi, '8': /\b(eight|eighth)\b/gi, '9': /\b(nine|ninth)\b/gi, '10': /\b(ten|tenth)\b/gi };
function i18nNumDiff_(sh, th, okx) {   /* okx：詞典譯名自帶的數字（例：4 吋刮板），出現在譯文不算多（只有投遞口會傳） */
  var a = i18nNums_(sh), b = i18nNums_(th), left = b.slice(), miss = [];
  a.forEach(function (x) { var j = left.indexOf(x); if (j < 0) miss.push(x); else left.splice(j, 1); });
  if (miss.length) return '原文的數字 ' + miss.join(',') + ' 譯文沒有（原文 ' + (a.join(',') || '無') + '／譯文 ' + (b.join(',') || '無') + '）';
  var pt = i18nPlain_(sh), extra = [], room = {};
  left.forEach(function (x) {
    if (okx && okx[x]) return;
    if (room[x] == null) room[x] = I18N_NUMW[x] ? (pt.match(I18N_NUMW[x]) || []).length : 0;
    if (room[x] > 0) room[x]--; else extra.push(x);
  });
  return extra.length ? '譯文多了數字 ' + extra.join(',') + '（原文沒有對應的字）' : '';
}
function i18nPh_(h) { return (String(h || '').match(/\[\[IMG\d+\]\]/g) || []).sort().join(','); }
/* 一頁譯文對原文：回傳問題清單（空＝通過） */
function i18nCheckSlide_(s, t, okx) {
  var bad = [];
  if (!t) return ['少了這一頁'];
  if (s.h && !i18nPlain_(t.h) && !i18nPh_(t.h)) bad.push('內文是空的');
  var d = i18nNumDiff_(s.h, t.h, okx);
  if (d) bad.push(d);
  if (i18nPh_(s.h) !== i18nPh_(t.h)) bad.push('圖片位置不一致');
  var tn = i18nNums_(t.t); if (!i18nNums_(s.t).every(function (x) { var j = tn.indexOf(x); if (j < 0) return false; tn.splice(j, 1); return true; })) bad.push('標題數字不見了');
  if (s.t && !String(t.t || '').trim()) bad.push('標題是空的');
  return bad;
}
/* 分段：每段最多 4 頁、5,000 字 */
function i18nChunks_(src) {
  var out = [], cur = [], n = 0;
  src.slides.forEach(function (s, i) {
    var len = (s.t || '').length + (s.h || '').length;
    if (cur.length && (cur.length >= I18N_CHUNK || n + len > I18N_CHUNK_CHARS)) { out.push(cur); cur = []; n = 0; }
    cur.push(i); n += len;
  });
  if (cur.length) out.push(cur);
  return out;
}

/* ---------------- 原文：清單＋抓頁 ---------------- */
function i18nSources_() {
  var sh = ss_().getSheetByName(LG_TAB);
  if (!sh || sh.getLastRow() < 2) throw fail_('食譜演化清單（recipe_lineage）是空的', 'nolist');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues(), H = v[0], ix = {};
  H.forEach(function (h, i) { ix[h] = i; });
  return v.slice(1).filter(function (r) { return r[ix.region] === '台灣' && r[ix.lang] === '英文' && r[ix.status] === 'Active'; })
    .map(function (r) { return { id: r[ix.recipe_id], name: r[ix.name], upd: r[ix.upd] }; });
}
function i18nFetchPages_(ids) {
  var sess = rbSession_();
  var mk = function () { var c = rbJarStr_(sess.jar); return ids.map(function (id) { return { url: RB_BASE + '/steps/recipedetails/' + id, method: 'get', muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA, Cookie: c } }; }); };
  var rs = UrlFetchApp.fetchAll(mk());
  if (rs.some(function (r) { return r.getResponseCode() === 302 && /Identity\/Account\/Login/i.test(String(r.getHeaders().Location || r.getHeaders().location || '')); })) {
    rbDropSession_(); sess.jar = rbFreshSession_().jar; rs = UrlFetchApp.fetchAll(mk());
  }
  rbSaveSession_(sess);
  return rs.map(function (r) { return { code: r.getResponseCode(), text: r.getResponseCode() === 200 ? r.getContentText() : '' }; });
}
/* 更新原文：後台更新日期變了（或還沒抓過、上次失敗）才重抓；回傳 {checked, fetched, changed, gone, errors} */
function i18nRefresh_(force, deadline) {
  var list = i18nSources_(), meta = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows, have = {}, res = { checked: list.length, fetched: 0, changed: 0, gone: 0, errors: [] };
  meta.forEach(function (r) { have[r.id] = r; });
  var todo = list.filter(function (s) { var h = have[s.id]; return force || !h || h.upd !== s.upd || h.state !== 'ok'; });
  for (var i = 0; i < todo.length; i += 5) {
    if (deadline && Date.now() > deadline) { res.errors.push('時間到，剩 ' + (todo.length - i) + ' 支下次再抓'); break; }
    var batch = todo.slice(i, i + 5), pages = i18nFetchPages_(batch.map(function (s) { return s.id; }));
    batch.forEach(function (s, j) {
      var pg = pages[j], row = { id: s.id, name: s.name, upd: s.upd, at: now_() };
      try {
        if (pg.code !== 200) throw fail_('後台回 HTTP ' + pg.code + '（可能已下架或上架日期不在期間內）', 'rbnet');
        var src = i18nParse_(pg.text, s.id), fp = i18nFp_(src);
        src.fp = fp; row.fp = fp; row.slides = src.slides.length; row.cover = src.cover; row.state = 'ok'; row.err = '';
        if (!have[s.id] || have[s.id].fp !== fp) res.changed++;
        i18nLock_(function () { i18nUpsert_('i18n_src', I18N_SRC_HEAD, 'id', row, src); });
        res.fetched++;
      } catch (e) {
        row.state = 'err'; row.err = errMsg_(e); res.errors.push(s.name + '：' + row.err);
        i18nLock_(function () { i18nUpsert_('i18n_src', I18N_SRC_HEAD, 'id', row, null); });
      }
    });
  }
  /* 清單裡沒有了（下架／停用）→ 標 gone，平板不列 */
  var on = {}; list.forEach(function (s) { on[s.id] = 1; });
  meta.filter(function (r) { return !on[r.id] && r.state !== 'gone'; }).forEach(function (r) {
    res.gone++; i18nLock_(function () { i18nUpsert_('i18n_src', I18N_SRC_HEAD, 'id', { id: r.id, state: 'gone', at: now_() }, null); });
  });
  i18nProp_('I18N_REFRESH_AT', now_());
  return res;
}
function i18nSrc_(id) {
  var m = i18nMeta_('i18n_src', I18N_SRC_HEAD), r = m.rows.filter(function (x) { return x.id === id; })[0];
  if (!r || r.state !== 'ok') return null;
  var o = i18nReadJ_(m.sh, r._r, I18N_SRC_HEAD); if (o) o.fp = r.fp;
  return o;
}

/* ---------------- 詞典：食材／工具／取放區／單位 ---------------- */
function i18nArea_(code, lang) {   /* 【Area A】→ 各語言 */
  var m = String(code).match(/^Area\s*([A-Z0-9]+)$/i); if (!m) return null;
  var x = m[1].toUpperCase();
  return { ja: x + 'エリア', ko: x + ' 구역', vi: 'Khu ' + x, id: 'Area ' + x, th: 'โซน ' + x }[lang];
}
function i18nSplitName_(n) { var m = String(n).match(/^【([^】]*)】\s*(.*)$/); return m ? { pre: m[1], core: m[2] } : { pre: '', core: String(n) }; }
function i18nTerms_(src) {   /* 這支食譜要查詞典的詞：[{kind, en}] */
  var seen = {}, out = [], add = function (kind, en) { en = rbTrim_(en); if (!en) return; var k = kind + '|' + en.toLowerCase(); if (seen[k]) return; seen[k] = 1; out.push({ kind: kind, en: en }); };
  src.slides.forEach(function (s) { s.ings.forEach(function (g) { add('item', i18nSplitName_(g.n).core); add('unit', g.u); add('area', g.an); }); });
  return out;
}
function i18nDictRows_() {
  var m = i18nMeta_('i18n_items', I18N_ITEM_HEAD), d = {};
  m.rows.forEach(function (r) { d[r.kind + '|' + String(r.en).toLowerCase()] = r; });
  return d;
}
/* 後台名稱對照表：英文 → {zh 台灣繁中原名（給 AI 當依據，英文名常有歧義：Egg yolk wash＝蛋黃液，不是刷蛋液）, vi 越南文} */
function i18nNameMap_() {
  var d = trDict_({}).data, langs = d.langs || [], ie = -1, iv = -1, iz = -1, map = {};
  langs.forEach(function (l, i) { var s = String(l).trim(); if (ie < 0 && /^(英文|English)$/i.test(s)) ie = i; if (iv < 0 && /Tiếng Việt|越南文/.test(s) && !/中文/.test(s)) iv = i; if (iz < 0 && /^繁體中文/.test(s)) iz = i; });
  if (iz < 0) iz = 0;
  if (ie < 0) return map;
  (d.rows || []).forEach(function (r) {
    var en = rbTrim_(r[3 + ie]).replace(/^【[^】]*】\s*/, ''), k = en.toLowerCase(); if (!en) return;
    var zh = rbTrim_(r[3 + iz]).replace(/^【[^】]*】\s*/, '').replace(/^[A-Z]{1,2}\d{2}\s+/, '').replace(/\s*[(（]已?下架[)）]/g, '');
    var o = map[k] = map[k] || { zhs: [], vi: '' };
    if (zh && o.zhs.indexOf(zh) < 0) o.zhs.push(zh);
    if (!o.vi && iv >= 0) o.vi = rbTrim_(r[3 + iv]);
  });
  /* 同一個英文名對到好幾個不同中文品項（例：Whipping cream）→ 不給中文依據，照英文意思翻（10/24 實測取第一個會翻成卡士達） */
  Object.keys(map).forEach(function (k) { map[k].zh = map[k].zhs.length === 1 ? map[k].zhs[0] : ''; });
  return map;
}
function i18nEnsureDict_(terms, lang, cred, t0, redo) {
  var D = i18nDictRows_(), miss = terms.filter(function (x) { var r = D[x.kind + '|' + x.en.toLowerCase()]; return redo || !r || !String(r[lang] || '').trim(); });
  if (!miss.length) return { added: 0 };
  var got = {}, nm = {};
  try { nm = i18nNameMap_(); } catch (e) { /* 對照表讀不到 → 只看英文 */ }
  if (lang === 'vi') miss.forEach(function (x) { var h = nm[x.en.toLowerCase()]; if (x.kind === 'item' && h && h.vi) got[x.kind + '|' + x.en.toLowerCase()] = { tr: h.vi, how: 'backend' }; });
  var rest = miss.filter(function (x) { return !got[x.kind + '|' + x.en.toLowerCase()]; });
  for (var i = 0; i < rest.length; i += 120) {
    var part = rest.slice(i, i + 120);
    var sys = ['You translate short labels used at a DIY baking studio in Taiwan into ' + I18N_LABEL[lang] + '.',
      'kind=item: ingredient, tool, mold, consumable or box name as printed on a tablet recipe list (short noun phrase, no explanation).',
      'kind=area: a storage area in the store where customers pick things up (e.g. "Pegboard wall area", "Refrigerated area").',
      'kind=unit: a quantity unit after a number (e.g. "g." "pc." "sheet" "cup"); keep metric symbols g / ml / kg as they are (drop the trailing dot is fine).',
      'zh is the original Taiwan Chinese name (the English is a translation of it and is sometimes ambiguous) — translate the MEANING of zh when given, e.g. 蛋黃液 "Egg yolk wash" is liquid egg yolk used as an ingredient, not an egg wash for brushing.',
      'Keep brand names (Oreo, Nutella…) as is. Return every k exactly as given.'].join('\n');
    var user = JSON.stringify({ items: part.map(function (x, j) { var h = nm[x.en.toLowerCase()]; return { k: String(j), kind: x.kind, en: x.en, zh: h && h.zh || '' }; }) });
    var schema = { type: 'object', properties: { items: { type: 'array', items: { type: 'object', properties: { k: { type: 'string' }, tr: { type: 'string' } }, required: ['k', 'tr'], additionalProperties: false } } }, required: ['items'], additionalProperties: false };
    var r = i18nAiOne_(sys, user, schema, cred);
    (r.json.items || []).forEach(function (it) { var x = part[Number(it.k)]; if (x && rbTrim_(it.tr)) got[x.kind + '|' + x.en.toLowerCase()] = { tr: rbTrim_(it.tr), how: 'ai' }; });
  }
  var n = 0;
  i18nLock_(function () {
    var m = i18nMeta_('i18n_items', I18N_ITEM_HEAD), idx = {};
    m.rows.forEach(function (r) { idx[r.kind + '|' + String(r.en).toLowerCase()] = r; });
    miss.forEach(function (x) {
      var k = x.kind + '|' + x.en.toLowerCase(), g = got[k]; if (!g) return;
      var r = idx[k], vals;
      if (r) { r[lang] = g.tr; r.how = (r.how ? r.how + ' ' : '') + lang + ':' + g.how; r.at = now_(); vals = I18N_ITEM_HEAD.map(function (h) { return r[h]; }); m.sh.getRange(r._r, 1, 1, vals.length).setNumberFormat('@').setValues([vals]); }
      else { var o = { kind: x.kind, en: x.en, how: lang + ':' + g.how, at: now_() }; o[lang] = g.tr; vals = I18N_ITEM_HEAD.map(function (h) { return o[h] || ''; }); m.sh.appendRow(vals); idx[k] = o; }
      n++;
    });
  });
  return { added: n, missing: miss.length - n };
}
function i18nLook_(D, kind, en, lang) { var r = D[kind + '|' + rbTrim_(en).toLowerCase()]; return r && String(r[lang] || '').trim() ? String(r[lang]).trim() : ''; }
function i18nItemName_(D, n, lang) {
  var sp = i18nSplitName_(n), core = i18nLook_(D, 'item', sp.core, lang) || sp.core;
  return sp.pre ? '【' + (i18nArea_(sp.pre, lang) || sp.pre) + '】' + core : core;
}

/* ---------------- 呼叫 Claude（經訓練站） ---------------- */
function i18nCred_() {
  var rows = authRows_('dim_role'), r = rows.filter(function (x) { return x.role === '管理者' && x.password; })[0] || rows.filter(function (x) { return x.role === '主廚' && x.password; })[0];
  if (!r) throw fail_('dim_role 找不到管理者或主廚的密碼，沒辦法呼叫翻譯', 'nocred');
  return { role: r.role, password: r.password, model: I18N_MODELS[0] };
}
function i18nAiReq_(sys, user, schema, cred) {
  var payload = { model: cred.model, max_tokens: 16000, system: sys, messages: [{ role: 'user', content: user }],
    output_config: { effort: 'low', format: { type: 'json_schema', schema: schema } } };
  return { url: I18N_AI_URL, method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ role: cred.role, password: cred.password, payload: payload }) };
}
/* 回傳 {ok, json, err, model, retryModel} */
function i18nAiRes_(resp, cred) {
  var code = resp.getResponseCode(), t = resp.getContentText(), j = null;
  try { j = JSON.parse(t); } catch (e) { return { ok: false, err: 'AI 回應看不懂（HTTP ' + code + '）' }; }
  if (code !== 200) {
    var m = JSON.stringify((j && j.error) || j).slice(0, 300);
    if ((code === 400 || code === 404) && /model/i.test(m) && cred.model === I18N_MODELS[0]) return { ok: false, err: m, retryModel: true };
    return { ok: false, err: 'AI 回應錯誤 ' + code + '：' + m };
  }
  if (j.stop_reason === 'max_tokens') return { ok: false, err: '這一段太長' };
  if (j.stop_reason === 'refusal') return { ok: false, err: 'AI 拒絕翻譯這一段' };
  var txt = (j.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  var a = txt.indexOf('{'), z = txt.lastIndexOf('}');
  try { return { ok: true, json: JSON.parse(txt.slice(a, z + 1)), model: j.model || cred.model }; } catch (e2) { return { ok: false, err: 'AI 回覆格式不對' }; }
}
function i18nAiOne_(sys, user, schema, cred) {
  for (var n = 0; n < 3; n++) {
    var r;
    try { r = i18nAiRes_(UrlFetchApp.fetch(I18N_AI_URL, i18nAiReq_(sys, user, schema, cred)), cred); }
    catch (e) { r = { ok: false, err: '連不到翻譯服務：' + errMsg_(e) }; }
    if (r.ok) return r;
    if (r.retryModel) { cred.model = I18N_MODELS[1]; continue; }
    if (n < 2) Utilities.sleep(2000 * (n + 1));
    else throw fail_(r.err, 'ai');
  }
}

/* ---------------- 翻一支食譜的一個語言 ---------------- */
function i18nGuideNames_() {   /* 到店須知頁的甜點在地名稱（客人在那頁選的名字，翻譯要一致） */
  var c = CacheService.getScriptCache(), hit = c.get('i18n:guide:v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var map = {};
  try {
    var d = JSON.parse(UrlFetchApp.fetch(I18N_GUIDE_API, { muteHttpExceptions: true, followRedirects: true }).getContentText());
    (d.items || []).forEach(function (it) { var k = i18nNameKey_(it.en); if (!k) return; map[k] = {}; I18N_LANGS.forEach(function (l) { var t = it.t && it.t[l]; if (t && t[0]) map[k][l] = t[0]; }); });
    c.put('i18n:guide:v1', JSON.stringify(map), 21600);
  } catch (e) { }
  return map;
}
function i18nNameKey_(s) { return String(s || '').split('｜')[0].replace(/[《》\s]/g, '').toLowerCase(); }

function i18nTitles_(src, lang, cred, local) {
  var u = [], seen = {}; src.slides.forEach(function (s) { if (s.t && !seen[s.t]) { seen[s.t] = 1; u.push(s.t); } });
  if (!u.length) return {};
  var sys = ['You translate the page titles of one DIY baking recipe (customers follow a tablet page by page) into ' + I18N_LABEL[lang] + '.',
    'Short titles. Keep numbers (1, 2, 2-1) exactly. The same English wording must always get the same translation; similar titles with a number (Frosting 1 / Frosting 2) must share the same words.',
    'Dessert: ' + src.name + (local ? ' (local name: ' + local + ')' : '') + '. Return every k exactly as given.'].join('\n');
  var schema = { type: 'object', properties: { items: { type: 'array', items: { type: 'object', properties: { k: { type: 'string' }, tr: { type: 'string' } }, required: ['k', 'tr'], additionalProperties: false } } }, required: ['items'], additionalProperties: false };
  var r = i18nAiOne_(sys, JSON.stringify({ titles: u.map(function (t, j) { return { k: String(j), en: t }; }) }), schema, cred), m = {};
  (r.json.items || []).forEach(function (it) { var t = u[Number(it.k)]; if (t && rbTrim_(it.tr)) m[t] = rbTrim_(it.tr); });
  return m;
}
function i18nSystem_(lang, src, D, local, tt) {
  var gl = [], seen = {};
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D, x.kind, x.en, lang); if (tr && !seen[x.en]) { seen[x.en] = 1; gl.push('- ' + x.en + ' → ' + tr); } });
  return ['You translate the in-store tablet recipe of 自己做烘焙聚樂部 (a DIY baking studio in Taiwan; customers bake by themselves, following the tablet page by page).',
    'Source: the store\'s official English recipe. Target: ' + I18N_LABEL[lang] + '.',
    'Target style:', I18N_RULES[lang].map(function (r) { return '* ' + r; }).join('\n'),
    'General rules:',
    '* Translate faithfully; do not add, drop, merge or reorder sentences. Keep numbering (1. 2. 3.) and line breaks.',
    '* "h" is HTML. Keep only these tags: <p>, <p class="c">, <br>, <b>, <u>, <i>, <font color="...">. Keep each tag around the corresponding translated words. Keep placeholders like [[IMG1]] exactly where they are.',
    '* Never change any number, weight, time, temperature, speed or count; write them with Arabic numerals exactly as in the source.',
    '* Keep emojis and symbols (●, ※, ⭐, ►, 【】, ［］). Translate "DIY Tip" as a short local label that keeps the word "DIY".',
    '* When the text mentions an ingredient, tool, mold, area or box from the glossary below, use exactly that translation.',
    '* "t" is the page title (short). Translate it the same way every time it repeats.',
    local ? '* The dessert name in this language is: ' + local + ' (use it for "name" and wherever the dessert name appears).' : '* "name" is "English name｜Chinese name (note)": translate the English part into the target language and keep "｜Chinese name" after it.',
    '* "pres" is the storage advice; translate it keeping line breaks (\\n). If "name" or "pres" is given as "", return "".',
    '* Return every k exactly as given.',
    'Glossary (English → ' + I18N_LABEL[lang] + '):', gl.join('\n') || '(none)',
    'Page titles (already translated — use exactly):', Object.keys(tt || {}).map(function (k) { return '- ' + k + ' → ' + tt[k]; }).join('\n') || '(none)'].join('\n');
}
var I18N_SCHEMA = { type: 'object', properties: { name: { type: 'string' }, pres: { type: 'string' },
  slides: { type: 'array', items: { type: 'object', properties: { k: { type: 'integer' }, t: { type: 'string' }, h: { type: 'string' } }, required: ['k', 't', 'h'], additionalProperties: false } } },
  required: ['name', 'pres', 'slides'], additionalProperties: false };

function i18nWork_(id, lang, by, redict) {
  var t0 = Date.now(), src = i18nSrc_(id);
  if (!src) throw fail_('還沒有這支食譜的原文（或上次讀取失敗）', 'nosrc');
  var cred = i18nCred_();
  var dict = i18nEnsureDict_(i18nTerms_(src), lang, cred, t0, redict), D = i18nDictRows_();
  var guide = i18nGuideNames_()[i18nNameKey_(src.name)] || {}, local = guide[lang] || '';
  var tt = i18nTitles_(src, lang, cred, local);   /* 標題先整批統一翻（同一個標題每頁同一個譯法） */
  var sys = i18nSystem_(lang, src, D, local, tt), chunks = i18nChunks_(src), out = { name: '', pres: '', slides: [] }, model = cred.model;
  var reqOf = function (ix, first, note) {
    var u = { name: first ? src.name : '', pres: first ? src.pres : '', slides: ix.map(function (i) { return { k: i, t: src.slides[i].t, h: src.slides[i].h }; }) };
    return i18nAiReq_(sys + (note ? '\n' + note : ''), 'Translate this JSON. Return the same structure.\n' + JSON.stringify(u), I18N_SCHEMA, cred);
  };
  var take = function (r, ix) {
    if (!r.ok) return false; model = r.model || model;
    if (r.json.name) out.name = r.json.name; if (r.json.pres) out.pres = r.json.pres;
    (r.json.slides || []).forEach(function (s) { if (ix.indexOf(s.k) >= 0) out.slides[s.k] = { t: s.t, h: s.h }; });
    return true;
  };
  var run = function (jobs) {   /* jobs: [{ix, first, note}] → 並行（每波 6 個）；整波逾時就改一個一個送 */
    for (var w = 0; w < jobs.length; w += I18N_WAVE) {
      if (Date.now() - t0 > I18N_BUDGET) return false;
      var wave = jobs.slice(w, w + I18N_WAVE), rs = null;
      try { rs = UrlFetchApp.fetchAll(wave.map(function (j) { return reqOf(j.ix, j.first, j.note); })); } catch (e) { rs = null; }
      wave.forEach(function (j, q) {
        var r;
        if (rs) r = i18nAiRes_(rs[q], cred);
        if (r && r.retryModel) { cred.model = I18N_MODELS[1]; r = null; }
        if (!r || !r.ok) { try { r = i18nAiRes_(UrlFetchApp.fetch(I18N_AI_URL, reqOf(j.ix, j.first, j.note)), cred); } catch (e2) { r = { ok: false, err: errMsg_(e2) }; } }
        take(r, j.ix);
      });
    }
    return true;
  };
  run(chunks.map(function (ix, n) { return { ix: ix, first: n === 0 }; }));
  /* 檢查 → 有問題的頁一頁一頁重翻一次（附上問題） */
  var probs = function () { var p = {}; src.slides.forEach(function (s, i) { var b = i18nCheckSlide_(s, out.slides[i]); if (b.length) p[i] = b; }); return p; };
  var P = probs(), redo = Object.keys(P).map(Number);
  if (!out.name || (src.pres && !out.pres)) redo.unshift(-1);
  if (redo.length && Date.now() - t0 < I18N_BUDGET) {
    run(redo.map(function (i) {
      if (i < 0) return { ix: [], first: true };
      return { ix: [i], first: false, note: 'IMPORTANT: a previous translation of this page had problems: ' + P[i].join('; ') + '. Keep every number and [[IMG]] placeholder exactly.' };
    }));
    P = probs();
  }
  if (local) out.name = local + '｜' + src.name.split('｜').slice(1).join('｜');
  src.slides.forEach(function (s, i) { if (out.slides[i] && tt[s.t]) out.slides[i].t = tt[s.t]; });
  var bad = Object.keys(P), sec = Math.round((Date.now() - t0) / 1000);
  var done = src.slides.every(function (s, i) { return !!out.slides[i]; }) && !!out.name;
  var status = !done ? 'fail' : bad.length ? 'review' : 'ok';
  var check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；');
  if (!done) check = '沒翻完（' + src.slides.filter(function (s, i) { return !out.slides[i]; }).length + ' 頁沒有回來）' + (check ? '；' + check : '');
  var key = id + '|' + lang, prev = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (r) { return r.key === key; })[0];
  var tries = prev && prev.fp === src.fp ? Number(prev.tries || 0) + 1 : 1;
  var json = { id: id, lang: lang, fp: src.fp, model: model, at: now_(), name: out.name, pres: out.pres, slides: out.slides, titles: tt };
  i18nLock_(function () { i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', { key: key, id: id, lang: lang, fp: src.fp, status: status, tries: tries, model: model, check: check, sec: sec, at: now_(), appr_by: '', appr_at: '', appr_fp: '' }, json); });
  i18nLog_(by, 'work', id, lang, status, (dict.added ? '詞典新增 ' + dict.added + '；' : '') + (check || '通過') , sec);
  return { id: id, lang: lang, name: src.name, status: status, check: check, sec: sec, slides: src.slides.length, chunks: chunks.length, dict: dict, model: model };
}

function i18nLive_(r, fp) { return !!r && r.fp === fp && r.status !== 'fail' && r.appr_fp === fp && !!r.appr_by; }   /* 平板看得到＝這版翻譯已核准 */
function i18nState_(r, fp) { if (!r || r.fp !== fp) return ''; if (i18nLive_(r, fp)) return 'live'; return r.status === 'ok' ? 'wait' : r.status; }   /* live 已放行／wait 自動檢查過、待核准／review 待檢查／fail */

/* 只重新跑自動檢查（不重翻、不花錢）：檢查規則改了以後用 */
function i18nRecheck_(id, lang, by) {
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  if (!r || r.fp !== src.fp || r.status === 'fail') throw fail_('這份翻譯不能重新檢查（沒翻好或原文改版）', 'notyet');
  var tr = i18nReadJ_(m.sh, r._r, I18N_RCP_HEAD), P = {};
  src.slides.forEach(function (s, i) { var b = i18nCheckSlide_(s, (tr.slides || [])[i]); if (b.length) P[i] = b; });
  var bad = Object.keys(P), check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；'), status = bad.length ? 'review' : 'ok';
  i18nLock_(function () { i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', { key: r.key, status: status, check: check }, null); });
  i18nLog_(by, 'recheck', id, lang, status, (r.status !== status ? r.status + '→' + status + '；' : '') + (check || '通過'));
  return { id: id, lang: lang, before: r.status, status: status, check: check };
}

/* ---------------- 待辦 ---------------- */
function i18nPending_() {
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (r) { return r.state === 'ok'; });
  var rc = {}; i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.forEach(function (r) { rc[r.key] = r; });
  var out = [];
  src.forEach(function (s) {
    I18N_LANGS.forEach(function (l) {
      var r = rc[s.id + '|' + l];
      if (!r || r.fp !== s.fp) out.push({ id: s.id, lang: l, name: s.name });
      else if (r.status === 'fail' && Number(r.tries || 0) < I18N_TRIES) out.push({ id: s.id, lang: l, name: s.name });
    });
  });
  return out;
}
function i18nClaim_(key) {
  return i18nLock_(function () { var c = CacheService.getScriptCache(), k = 'i18n:claim:' + key; if (c.get(k)) return false; c.put(k, '1', 330); return true; });
}
function i18nRelease_(key) { try { CacheService.getScriptCache().remove('i18n:claim:' + key); } catch (e) { } }
function i18nEnsureTrigger_() {
  if (i18nProp_('I18N_TRIG') === '1') return;
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'i18nDaily'; });
  if (!has) ScriptApp.newTrigger('i18nDaily').timeBased().atHour(6).nearMinute(30).everyDays(1).inTimezone(TZ).create();
  i18nProp_('I18N_TRIG', '1');
}

/* ---------------- 動作 ---------------- */
function i18nKick_(p) {
  var t0 = Date.now(); i18nEnsureTrigger_();
  var auto = i18nProp_('I18N_AUTO') === '1', id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  if (lang && I18N_LANGS.indexOf(lang) < 0) throw fail_('語言只能是 ' + I18N_LANGS.join('／'), 'bad');
  if (!auto && !(id && lang)) throw fail_('還沒開放全部自動翻譯：請指定 id 和 lang（一次一支一個語言）', 'off');
  /* 原文 20 小時沒更新（或從來沒抓過）→ 先更新 */
  var ref = null, last = i18nProp_('I18N_REFRESH_AT'), age = last ? Date.now() - new Date(last.replace(' ', 'T') + '+08:00').getTime() : Infinity;
  if (age > 20 * 3600e3 || (id && !i18nSrc_(id))) ref = i18nRefresh_(false, t0 + 150000);
  if (p.recheck && id && lang) return { data: i18nRecheck_(id, lang, p.by || 'recheck') };
  if (p.redo || p.redict) i18nBossOrKey_(p);   /* 重翻會再花錢 → 要主廚／管理者密碼（或唯讀預覽通行碼＝Claude 驗證用） */
  var pend = i18nPending_().filter(function (x) { return (!id || x.id === id) && (!lang || x.lang === lang); });
  if (p.redo && id && lang && !pend.length && i18nSrc_(id)) pend = [{ id: id, lang: lang }];   /* 指定重翻（例：翻譯規則改了）；重翻後核准會失效 */
  if (id && lang && !pend.length) {
    var cur = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (r) { return r.key === id + '|' + lang; })[0];
    return { data: { did: null, msg: cur ? '這支這個語言已經翻好了（' + cur.status + '，' + cur.at + '）；原文沒變就不重翻' : '找不到這支食譜（不是台灣區啟用中的英文食譜？）', refresh: ref, left: i18nPending_().length } };
  }
  for (var i = 0; i < pend.length; i++) {
    var x = pend[i], key = x.id + '|' + x.lang;
    if (!i18nClaim_(key)) continue;
    try { var r = i18nWork_(x.id, x.lang, p.by || 'kick', !!p.redict); r.refresh = ref; r.left = i18nPending_().length; return { id: key, data: r }; }
    catch (e) { i18nLog_(p.by || 'kick', 'work', x.id, x.lang, 'error', errMsg_(e)); throw e; }
    finally { i18nRelease_(key); }
  }
  return { data: { did: null, msg: pend.length ? '剩下的都有人在做' : '沒有待翻譯的', refresh: ref, left: pend.length } };
}
function i18nDaily() {
  var t0 = Date.now(), ref = null;
  try { ref = i18nRefresh_(false, t0 + 180000); } catch (e) { i18nLog_('排程', 'refresh', '', '', 'error', errMsg_(e)); return; }
  i18nLog_('排程', 'refresh', '', '', 'ok', '檢查 ' + ref.checked + '、重抓 ' + ref.fetched + '、內容有變 ' + ref.changed + '、下架 ' + ref.gone + (ref.errors.length ? '；' + ref.errors.join('；') : ''));
  if (i18nProp_('I18N_AUTO') !== '1') return;
  while (Date.now() - t0 < 300000) {
    var pend = i18nPending_(); if (!pend.length) return;
    var x = pend[0], key = x.id + '|' + x.lang; if (!i18nClaim_(key)) return;
    try { i18nWork_(x.id, x.lang, '排程'); } catch (e) { i18nLog_('排程', 'work', x.id, x.lang, 'error', errMsg_(e)); } finally { i18nRelease_(key); }
  }
}
function i18nStatus_() {
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows, rc = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows, cnt = {}, review = [], fail = [], wait = [];
  var nm = {}; src.forEach(function (s) { nm[s.id] = s.name; });
  I18N_LANGS.forEach(function (l) { cnt[l] = { live: 0, wait: 0, review: 0, fail: 0, old: 0 }; });
  var fp = {}; src.forEach(function (s) { fp[s.id] = s.fp; });
  rc.forEach(function (r) {
    if (!cnt[r.lang] || !fp[r.id]) return;
    var st = r.fp !== fp[r.id] ? 'old' : i18nState_(r, fp[r.id]); if (cnt[r.lang][st] != null) cnt[r.lang][st]++;
    if (st === 'wait') wait.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', at: r.at });
    if (st === 'review') review.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', check: r.check, at: r.at });
    if (st === 'fail') fail.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', check: r.check, tries: r.tries, at: r.at });
  });
  var lg = [], sh = ss_().getSheetByName('i18n_log');
  if (sh && sh.getLastRow() > 1) { var n = Math.min(20, sh.getLastRow() - 1); lg = sh.getRange(sh.getLastRow() - n + 1, 1, n, I18N_LOG_HEAD.length).getDisplayValues().reverse(); }
  var trig = -1; try { trig = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'i18nDaily'; }).length; } catch (e) { }
  var rcm = {}; rc.forEach(function (r) { rcm[r.key] = r; });
  var guide = {}; try { guide = i18nGuideNames_(); } catch (e) { }
  var matrix = src.filter(function (s) { return s.state === 'ok'; }).map(function (s) {
    var o = { id: s.id, name: s.name, cover: s.cover, slides: s.slides, upd: s.upd, local: guide[i18nNameKey_(s.name)] || {}, st: {}, appr: {} };
    I18N_LANGS.forEach(function (l) { var r = rcm[s.id + '|' + l]; o.st[l] = i18nState_(r, s.fp); if (r && o.st[l] === 'live') o.appr[l] = r.appr_by + ' ' + r.appr_at; });
    return o;
  });
  return { data: { langs: I18N_LANGS, matrix: matrix, sources: src.filter(function (s) { return s.state === 'ok'; }).length, src_err: src.filter(function (s) { return s.state === 'err'; }).map(function (s) { return { id: s.id, name: s.name, err: s.err }; }),
    count: cnt, pending: i18nPending_().length, wait: wait, review: review, fail: fail, auto: i18nProp_('I18N_AUTO') === '1', refresh_at: i18nProp_('I18N_REFRESH_AT'), schedule: trig, log: lg } };
}
/* 平板密碼：問權限中心 verify（dim_auth 的 tablet_i18n），通過快取 2 小時（只存雜湊） */
function i18nTabletAuth_(pw) {
  pw = String(pw || '').trim(); if (!pw) throw fail_('請輸入平板密碼', 'auth');
  var c = CacheService.getScriptCache(), k = 'i18n:au:' + i18nHash_(I18N_AUTH_DASH + '|' + pw);
  if (c.get(k)) return true;
  var t = UrlFetchApp.fetch(I18N_AUTH_API + '?action=verify&dashboard=' + I18N_AUTH_DASH + '&pwd=' + encodeURIComponent(pw) + '&callback=cb', { muteHttpExceptions: true, followRedirects: true }).getContentText();
  var m = t.match(/^\s*[A-Za-z_$][\w$]*\(([\s\S]*)\);?\s*$/), j = {};
  try { j = JSON.parse(m ? m[1] : t); } catch (e) { throw fail_('權限中心回應看不懂，請稍後再試', 'authnet'); }
  if (!j || !j.ok) { Utilities.sleep(800); throw fail_('平板密碼不對（' + (j && j.error || '') + '）', 'auth'); }
  c.put(k, '1', 7200);
  return true;
}
function i18nList_(p) {
  i18nTabletAuth_(p.tpw);
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD), rc = {}, guide = i18nGuideNames_();
  i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.forEach(function (r) { rc[r.key] = r; });
  var out = src.rows.filter(function (s) { return s.state === 'ok'; }).map(function (s) {
    var o = { id: s.id, name: s.name, cover: s.cover, langs: {}, ver: {}, local: guide[i18nNameKey_(s.name)] || {} };
    I18N_LANGS.forEach(function (l) { var r = rc[s.id + '|' + l]; o.langs[l] = i18nState_(r, s.fp); if (o.langs[l] === 'live') o.ver[l] = String(s.fp).slice(0, 8) + '|' + r.appr_at; });   /* ver：平板快取用，核准或原文改了就換 */
    return o;
  });
  return { data: { at: now_(), items: out } };
}
function i18nGet_(p) {
  i18nTabletAuth_(p.tpw);
  var id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  if (I18N_LANGS.indexOf(lang) < 0) throw fail_('語言不對', 'bad');
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  if (!r || r.fp !== src.fp) throw fail_('這支食譜這個語言還沒翻好', 'notyet');
  if (!i18nLive_(r, src.fp)) throw fail_('這支食譜這個語言還沒核准放行', 'notyet');
  return { data: i18nBuild_(src, i18nReadJ_(m.sh, r._r, I18N_RCP_HEAD), i18nDictRows_(), lang, r) };
}
function i18nBoss_(auth) { if (!auth || auth.kind !== 'role' || ['主廚', '管理者'].indexOf(auth.role) < 0) throw fail_('要主廚或管理者的密碼', 'perm'); }
/* 對照預覽：譯文合成＋原文英文（每頁並排看） */
var I18N_PREVIEW_HASH = 'Fdcs8oo7F0UkiLkKLN3Arb';   /* 唯讀預覽通行碼的雜湊（原文只在經營者本機 ~/.config/diybc/secrets.env 的 I18N_PREVIEW_KEY；只能看對照，不能核准） */
function i18nBossOrKey_(p) {
  if (p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH) return;
  var auth; try { auth = auth_(p); } catch (e) { Utilities.sleep(1000); throw e; }
  i18nBoss_(auth);
}
function i18nPreview_(p) {
  i18nBossOrKey_(p);
  var id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  if (!r || r.fp !== src.fp) throw fail_('這支食譜這個語言還沒翻好', 'notyet');
  var D = i18nDictRows_(), b = i18nBuild_(src, i18nReadJ_(m.sh, r._r, I18N_RCP_HEAD), D, lang, r);
  b.state = i18nState_(r, src.fp); b.appr_by = r.appr_by; b.appr_at = r.appr_at;
  b.src = src.slides.map(function (s) { var h = s.h; (s.imgs || []).forEach(function (u, j) { h = h.split('[[IMG' + (j + 1) + ']]').join('<img src="' + String(u).replace(/"/g, '&quot;') + '">'); }); return { t: s.t, h: h, ings: s.ings.map(function (g) { return { n: g.n, q: g.q, u: g.u }; }) }; });
  b.src_name = src.name; b.src_pres = src.pres;
  return { data: b };
}
/* 核准放行／收回 */
function i18nApprove_(p, auth) {
  i18nBoss_(auth); I18N_HELD = true;
  var id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim(), key = id + '|' + lang;
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var r = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (x) { return x.key === key; })[0];
  if (!r || r.fp !== src.fp) throw fail_('這支食譜這個語言還沒翻好（或原文改版了，要先重翻）', 'notyet');
  if (r.status === 'fail') throw fail_('這份翻譯沒翻完，不能核准', 'bad');
  var who = auth.role + (p.who ? '：' + String(p.who).slice(0, 20) : '');
  var o = p.undo ? { key: key, appr_by: '', appr_at: '', appr_fp: '' } : { key: key, appr_by: who, appr_at: now_(), appr_fp: src.fp };
  i18nLock_(function () { i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', o, null); });
  SpreadsheetApp.flush();   /* 一般寫入路徑放鎖前不會 flush，這裡自己送出 */
  i18nLog_(who, p.undo ? 'unapprove' : 'approve', id, lang, 'ok', r.status === 'review' ? '核准時仍有自動檢查提醒：' + r.check : '');
  return { id: key, msg: (p.undo ? '收回 ' : '核准 ') + key, data: { key: key, state: p.undo ? (r.status === 'ok' ? 'wait' : r.status) : 'live', appr_by: o.appr_by, appr_at: o.appr_at } };
}
/* 合成平板要的完整內容：版面照原文，文字換譯文，食材／取放區／單位換詞典 */
function i18nBuild_(src, tr, D, lang, row) {
  var areas = {}, NM = {};
  try { NM = i18nNameMap_(); } catch (e) { /* 對照表讀不到 → 只附英文 */ }
  var cut = function (n) { return rbTrim_(String(n || '').replace(/^【[^】]*】\s*/, '')); };
  var slides = src.slides.map(function (s, i) {
    var t = (tr.slides || [])[i] || {}, h = String(t.h || '');
    (s.imgs || []).forEach(function (u, j) { h = h.split('[[IMG' + (j + 1) + ']]').join('<img src="' + String(u).replace(/"/g, '&quot;') + '">'); });
    var ings = s.ings.map(function (g) {
      if (g.a && !areas[g.a]) areas[g.a] = { name: i18nLook_(D, 'area', g.an, lang) || g.an, en: g.an, img: src.areaPre && src.areaHas.indexOf(g.a) >= 0 ? '/img/MenuCateGroups/' + src.areaPre + '-' + g.a + '.jpg' : '' };
      var nmk = NM[cut(g.n).toLowerCase()];
      return { n: i18nItemName_(D, g.n, lang), en: g.n, zh: nmk && nmk.zh || '', q: g.q, u: i18nLook_(D, 'unit', g.u, lang) || g.u, c: g.c, a: g.a };
    });
    return { t: t.t || s.t, video: s.video, img: s.img, h: h, ings: ings };
  });
  return { id: src.id, lang: lang, name: tr.name || src.name, en: src.name, cover: src.cover, pres: tr.pres || src.pres, areas: areas, slides: slides,
    status: row.status, check: row.check, at: row.at, model: row.model };
}

/* 一次性：權限中心 dim_auth 補「tablet_i18n」列（只補沒有的；先整張複製成備份分頁；密碼空白＝還不能用，經營者在權限管理後台設定） */
var I18N_AUTHC_SS = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4';
function i18nAuthSetup_(p) {
  if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) throw fail_('要通行碼', 'auth');
  var ss = SpreadsheetApp.openById(I18N_AUTHC_SS), sh = ss.getSheetByName('dim_auth');
  if (!sh) throw fail_('找不到權限中心 dim_auth', 'bad');
  var last = sh.getLastRow(), ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  if (ids.indexOf(I18N_AUTH_DASH) >= 0) return { data: { msg: '已經有 tablet_i18n 列，不用補', rows: last - 1 } };
  var bk = 'dim_auth_備份_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm');
  if (!ss.getSheetByName(bk)) sh.copyTo(ss).setName(bk);
  sh.getRange(last + 1, 1, 1, 7).setNumberFormat('@').setValues([[I18N_AUTH_DASH, '平板外語食譜', '', '720', 'Y', now_(), '店內平板看日韓越印泰食譜（tablet-recipe.html）；720 小時＝平板輸入一次記 30 天。2026-10-24 補；密碼空白＝還不能用，請在權限管理後台設定']]);
  var after = sh.getLastRow() - 1, chk = sh.getRange(last + 1, 1, 1, 5).getValues()[0];
  return { data: { msg: '已補 tablet_i18n 列', before: last - 1, after: after, backup: bk, row: last + 1, check: [chk[0], chk[1], chk[2] ? '(有值)' : '(空白)', chk[3], chk[4]] } };
}
ACTIONS.i18nAuthSetup = function (p, a) { return i18nAuthSetup_(p, a); };
NOAUTH.i18nAuthSetup = 1; SELFLOCK.i18nAuthSetup = 1;
/* 詞典整理：同一個 kind＋英文名重複的列合併成一列（各語言取有值的，衝突取最早那列並列出），其餘刪掉；唯讀預覽通行碼才能跑 */
function i18nDictRepair_(p) {
  if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) throw fail_('要通行碼', 'auth');
  return { data: i18nLock_(function () {
    var m = i18nMeta_('i18n_items', I18N_ITEM_HEAD), first = {}, dels = [], conflicts = [], merged = 0, before = m.rows.length;
    m.rows.forEach(function (r) {
      var k = r.kind + '|' + String(r.en).toLowerCase(), f = first[k];
      if (!f) { first[k] = r; return; }
      I18N_LANGS.forEach(function (l) { var a = String(f[l] || '').trim(), b = String(r[l] || '').trim(); if (!a && b) f[l] = b; else if (a && b && a !== b) conflicts.push(r.en + '［' + l + '］' + a + '／' + b); });
      f.how = [f.how, r.how].filter(Boolean).join(' '); f._dirty = 1; dels.push(r._r); merged++;
    });
    Object.keys(first).forEach(function (k) { var f = first[k]; if (!f._dirty) return; m.sh.getRange(f._r, 1, 1, I18N_ITEM_HEAD.length).setNumberFormat('@').setValues([I18N_ITEM_HEAD.map(function (h) { return f[h] == null ? '' : String(f[h]); })]); });
    dels.sort(function (a, b) { return b - a; }).forEach(function (r) { m.sh.deleteRow(r); });
    var empty = {}; I18N_LANGS.forEach(function (l) { empty[l] = Object.keys(first).filter(function (k) { return !String(first[k][l] || '').trim(); }).length; });
    i18nLog_('Claude', 'dictRepair', '', '', 'ok', '合併重複 ' + merged + ' 列；' + before + '→' + (before - merged) + ' 列' + (conflicts.length ? '；衝突 ' + conflicts.join('、') : ''));
    return { before: before, after: before - merged, merged: merged, conflicts: conflicts, empty: empty };
  }) };
}
ACTIONS.i18nDictRepair = function (p, a) { return i18nDictRepair_(p, a); };
NOAUTH.i18nDictRepair = 1; SELFLOCK.i18nDictRepair = 1;
ACTIONS.i18nStatus = function (p, a) { return i18nStatus_(p, a); };
ACTIONS.i18nList = function (p, a) { return i18nList_(p, a); };
ACTIONS.i18nGet = function (p, a) { return i18nGet_(p, a); };
ACTIONS.i18nKick = function (p, a) { return i18nKick_(p, a); };
ACTIONS.i18nPreview = function (p, a) { return i18nPreview_(p, a); };
ACTIONS.i18nApprove = function (p, a) { return i18nApprove_(p, a); };
READS.i18nPreview = 1; NOAUTH.i18nPreview = 1;   /* 身分在函式裡驗（主廚／管理者密碼，或唯讀預覽通行碼） */   /* i18nApprove 走一般寫入路徑（有鎖、有回條、寫紀錄） */
NOAUTH.i18nStatus = 1; NOAUTH.i18nList = 1; NOAUTH.i18nGet = 1; NOAUTH.i18nKick = 1;   /* 平板密碼在函式裡自己驗；Kick 只會做待辦，做完就沒事可做 */
READS.i18nStatus = 1; READS.i18nList = 1; READS.i18nGet = 1;
SELFLOCK.i18nKick = 1;   /* 一次要 1～4 分鐘，不能佔住全系統的鎖；同一件不重複做靠 i18nClaim_ */

/* ================= 🤝 Claude Code 投遞口（2026-10-24 經營者：用 Claude Code 翻、不花 API；主廚不看，Claude 自己檢查就放行） =================
   流程：Claude Code 呼叫 i18nPull 拿一支食譜某語言的原文＋詞典 → 自己翻 → 自己逐頁回讀 → 呼叫 i18nPut 投遞。
   i18nPut 會照原本同一套「硬檢查」（頁數、每頁數字、圖片位置、標題數字、空白）再驗一次：
     過 ＋ approve:1 → 狀態 ok 並直接放行（appr_by＝「Claude」，平板看得到）；
     沒過 → 存成 review（待檢查），不放行，等經營者決定。
   通行碼：投遞專用（跟唯讀預覽通行碼不同；原文只在經營者本機 ~/.config/diybc/secrets.env 的 I18N_PUT_KEY，程式只放雜湊）。
   放行後仍可在核准頁抽查、按「收回」。原文改版 → 放行自動失效（appr_fp 對不上）。
   ⚠ 只寫自己的分頁（i18n_rcp、i18n_items、i18n_log）；不寫食譜後台。 */
var I18N_PUT_HASH = 'SqDmIh_9xBSvmNWjYG94Rb';
function i18nWriteKey_(p) {
  if (!(p.wkey && i18nHash_(String(p.wkey)) === I18N_PUT_HASH)) { Utilities.sleep(1000); throw fail_('要投遞通行碼', 'auth'); }
}
/* 拿原稿：不給 id＋lang → 回待翻清單；給了 → 回該份的原文、規則、詞典（已有的）、缺的詞（附後台中文原名當依據） */
function i18nPull_(p) {
  i18nWriteKey_(p);
  var id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  if (!(id && lang)) {
    var pend = i18nPending_(), rc = {}; i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.forEach(function (r) { rc[r.key] = r; });
    return { data: { pending: pend.length, items: pend.slice(0, 400).map(function (x) { var r = rc[x.id + '|' + x.lang]; return { id: x.id, lang: x.lang, name: x.name, why: !r ? '還沒翻' : r.status === 'fail' ? '失敗' : '原文改版' }; }), at: now_() } };
  }
  if (I18N_LANGS.indexOf(lang) < 0) throw fail_('語言只能是 ' + I18N_LANGS.join('／'), 'bad');
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜（不是台灣區啟用中的英文食譜？）', 'nosrc');
  var D = i18nDictRows_(), terms = i18nTerms_(src), have = [], miss = [], nm = {};
  try { nm = i18nNameMap_(); } catch (e) { }
  terms.forEach(function (x) {
    var tr = i18nLook_(D, x.kind, x.en, lang);
    if (tr) { have.push({ kind: x.kind, en: x.en, tr: tr }); return; }
    var h = nm[x.en.toLowerCase()] || {}, o = { kind: x.kind, en: x.en, zh: h.zh || '' };
    if (lang === 'vi' && x.kind === 'item' && h.vi) o.backend = h.vi;   /* 越南文：後台已有的名稱優先 */
    miss.push(o);
  });
  var guide = i18nGuideNames_()[i18nNameKey_(src.name)] || {}, local = guide[lang] || '';
  var cur = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (r) { return r.key === id + '|' + lang; })[0];
  return { data: { id: id, lang: lang, label: I18N_LABEL[lang], rules: I18N_RULES[lang], fp: src.fp, name: src.name, pres: src.pres, local: local,
    slides: src.slides.map(function (s, i) { return { k: i, t: s.t, h: s.h }; }),
    dict: have, missing: miss, brief: i18nSystem_(lang, src, D, local, {}),
    now: cur ? { status: cur.status, state: i18nState_(cur, src.fp), at: cur.at, model: cur.model } : null } };
}
/* 把 Claude 翻好的詞放進詞典：只補空格、不蓋掉已有的（除非 overwrite）；只收這支食譜用到的詞 */
function i18nPutDict_(list, lang, src, overwrite) {
  var allowed = {}; i18nTerms_(src).forEach(function (x) { allowed[x.kind + '|' + x.en.toLowerCase()] = x; });
  var got = {}, ignored = 0;
  (list || []).forEach(function (x) {
    var k = String(x.kind) + '|' + rbTrim_(x.en).toLowerCase(), tr = rbTrim_(x.tr);
    if (allowed[k] && tr) got[k] = { x: allowed[k], tr: tr }; else ignored++;
  });
  var added = 0, kept = 0;
  i18nLock_(function () {
    var m = i18nMeta_('i18n_items', I18N_ITEM_HEAD), idx = {};
    m.rows.forEach(function (r) { idx[r.kind + '|' + String(r.en).toLowerCase()] = r; });
    Object.keys(got).forEach(function (k) {
      var g = got[k], r = idx[k], vals;
      if (r) {
        if (String(r[lang] || '').trim() && !overwrite) { kept++; return; }
        r[lang] = g.tr; r.how = (r.how ? r.how + ' ' : '') + lang + ':claude'; r.at = now_();
        vals = I18N_ITEM_HEAD.map(function (h) { return r[h] == null ? '' : r[h]; }); m.sh.getRange(r._r, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
      } else {
        var o = { kind: g.x.kind, en: g.x.en, how: lang + ':claude', at: now_() }; o[lang] = g.tr;
        vals = I18N_ITEM_HEAD.map(function (h) { return o[h] || ''; }); m.sh.appendRow(vals); idx[k] = o;
      }
      added++;
    });
  });
  return { added: added, kept: kept, ignored: ignored };
}
/* 收翻譜：詞典 → 硬檢查 → 存 i18n_rcp →（過＋approve:1）放行 */
function i18nPut_(p) {
  i18nWriteKey_(p);
  var t0 = Date.now(), id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim(), by = String(p.by || 'Claude Code').slice(0, 40);
  if (I18N_LANGS.indexOf(lang) < 0) throw fail_('語言只能是 ' + I18N_LANGS.join('／'), 'bad');
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  if (p.fp && String(p.fp) !== String(src.fp)) throw fail_('原文在你拿稿之後改版了（指紋不同），請重新拿稿再翻', 'stale');
  var sl = p.slides; if (!sl || sl.length !== src.slides.length) throw fail_('頁數不對：原文 ' + src.slides.length + ' 頁，收到 ' + (sl ? sl.length : 0) + ' 頁', 'bad');
  var out = { name: String(p.name || '').trim(), pres: String(p.pres || '').trim(), slides: [] };
  sl.forEach(function (s) { var k = Number(s.k); if (!(k >= 0 && k < src.slides.length) || out.slides[k]) throw fail_('頁碼 k 不對或重複：' + s.k, 'bad'); out.slides[k] = { t: String(s.t || '').trim(), h: String(s.h || '') }; });
  var dict = i18nPutDict_(p.dict, lang, src, !!p.overwrite);
  var local = (i18nGuideNames_()[i18nNameKey_(src.name)] || {})[lang] || '';
  if (local) out.name = local + '｜' + src.name.split('｜').slice(1).join('｜');
  var tt = {}; src.slides.forEach(function (s, i) { if (s.t && out.slides[i] && out.slides[i].t) tt[s.t] = out.slides[i].t; });
  var D2 = i18nDictRows_(), okx = {};   /* 這支食譜用到的詞典譯名裡的數字（尺寸、號數）不算「多出來的數字」 */
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D2, x.kind, x.en, lang); if (tr) i18nNums_(tr).forEach(function (n) { okx[n] = 1; }); });
  if (local) i18nNums_(local).forEach(function (n) { okx[n] = 1; });   /* 官網在地品名自帶的數字（例：오레오 2단）也放過 */
  var P = {}; src.slides.forEach(function (s, i) { var b = i18nCheckSlide_(s, out.slides[i], okx); if (b.length) P[i] = b; });
  var bad = Object.keys(P), check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；');
  if (src.pres && !out.pres) check += (check ? '；' : '') + '保存方式是空的';
  var status = !out.name ? 'fail' : (check ? 'review' : 'ok');
  if (!out.name) check = '名稱是空的' + (check ? '；' + check : '');
  var key = id + '|' + lang, prev = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (r) { return r.key === key; })[0];
  var tries = prev && prev.fp === src.fp ? Number(prev.tries || 0) + 1 : 1, model = String(p.model || 'claude-code').slice(0, 60);
  var approved = status === 'ok' && !!p.approve, sec = Math.round((Date.now() - t0) / 1000);
  var json = { id: id, lang: lang, fp: src.fp, model: model, at: now_(), name: out.name, pres: out.pres, slides: out.slides, titles: tt };
  i18nLock_(function () {
    i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', { key: key, id: id, lang: lang, fp: src.fp, status: status, tries: tries, model: model, check: check, sec: sec, at: now_(),
      appr_by: approved ? 'Claude' : '', appr_at: approved ? now_() : '', appr_fp: approved ? src.fp : '' }, json);
  });
  i18nLog_(by, 'put', id, lang, status, (dict.added ? '詞典新增 ' + dict.added + '；' : '') + (approved ? '自動放行；' : '') + (check || '通過'), sec);
  return { id: key, msg: 'put ' + key + ' ' + status, data: { id: id, lang: lang, name: out.name, status: status, check: check, approved: approved, state: approved ? 'live' : (status === 'ok' ? 'wait' : status), slides: src.slides.length, dict: dict, left: i18nPending_().length } };
}
ACTIONS.i18nPull = function (p, a) { return i18nPull_(p, a); };
ACTIONS.i18nPut = function (p, a) { return i18nPut_(p, a); };
NOAUTH.i18nPull = 1; NOAUTH.i18nPut = 1; READS.i18nPull = 1; SELFLOCK.i18nPut = 1;   /* 通行碼在函式裡驗；Put 自己拿鎖 */

/* 編輯器手動驗證：更新原文＋翻第一支的日文 */
function i18nTestNow() {
  var r = i18nRefresh_(false, Date.now() + 200000); Logger.log(JSON.stringify(r));
  var s = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (x) { return x.state === 'ok'; })[0];
  if (s) Logger.log(JSON.stringify(i18nWork_(s.id, 'ja', '編輯器')));
}
