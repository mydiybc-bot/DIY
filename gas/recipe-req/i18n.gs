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
     i18nPick    主廚／管理者：勾選哪幾支給外國客人用（2026-10-26 經營者：「讓哪一些食譜可以被我新增到外國人可用的食譜中，讓我可以點選」）。
                 沒勾的不翻、平板不列；取消勾譯文與核准都保留，再勾回來（原文沒改）馬上恢復。分頁 i18n_pick（id, name, on Y/N, by, at）。
                 i18n_pick 還不存在＝還沒人勾過 → 全部當作有勾（上線前 39 支已全數放行，不能讓平板突然變空）；
                 第一次勾選時把當下所有原文先記成 Y 再改那一支。
                 2026-11-04 經營者：「不需要手動勾選，只要外語食譜收入這新甜點，就自動啟用」→ 不在表上＝有勾（自動翻、放行後上平板）；
                 只有表上記成 N（有人按過取消勾）才不翻、平板不列。
     i18nApprove 主廚／管理者：核准放行（undo:1 收回）。**沒核准的平板看不到**（經營者 10/24：「完成後還是要一一打開比對，確認沒有問題才可以放行」）；
                 原文改版重翻 → 核准自動失效（appr_fp 對不上），要重新看過
                 2026-11-05：只改食材份量／圖片／影片不算改版（指紋只看要翻的字）；原文改版後、新譯文放行前，平板照舊顯示放行過的上一版；
                 Claude 投遞口只翻改到的頁（i18nPull 回 keep，i18nPut partial:1 沿用沒變的頁）
     i18nKick    做下一件待翻譯的事：沒開自動（指令碼屬性 I18N_AUTO≠1）時只做指定的 id＋lang；同一件 6 分鐘內不會兩個人同時做
   排程：i18nDaily 每天 06:30（i18nKick 第一次被呼叫時自動建立，不用到編輯器按）：更新原文（後台有改才重抓、內容真的變了才重翻）；
         I18N_AUTO＝1 時接著翻待翻譯的。2026-11-05 起改成 5 個語言各自接力（i18nAuto_ja…，見 i18nAutoStart_），全部翻完為止；
         Sonnet 翻＋AI 回讀，過了自動放行（放行人「AI 自動檢查」），Mac 不用開。
   2026-11-24 原文改用中文食譜（經營者：「平板英文和其他語言食譜全部改用中文直接翻譯」）：見 i18n_zbase.gs。
         英文原文這條路照舊保留，中文版 6 種語言都放行後，那支舊的才退場（新舊並行、平板不開天窗）。
   ⚠ 只讀後台；不寫後台任何東西。 */

var I18N_LANGS = ['ja', 'ko', 'vi', 'id', 'th'];
var I18N_LABEL = { en: 'English', ja: 'Japanese 日本語', ko: 'Korean 한국어', vi: 'Vietnamese Tiếng Việt', id: 'Indonesian Bahasa Indonesia', th: 'Thai ภาษาไทย' };
var I18N_RULES = {
  en: ['Simple, clear English for customers (many are not native speakers) following a recipe at a baking studio. Short imperative sentences (e.g. "Whisk 20 times."). Use common words (bowl, rubber spatula, piping bag).'],   /* 中文原文才有英文（2026-11-24） */
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
var I18N_PICK_HEAD = ['id', 'name', 'on', 'by', 'at'];
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
function i18nParse_(html, id, seq) {   /* seq：標題照步驟清單順序對（自選版頁面多一張重複的拿托盤頁，用頁序對會錯一格，2026-11-26） */
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
  var tl = Object.keys(titles).sort(function (a, b) { return a - b; }).map(function (k) { return titles[k]; });
  /* 每一頁 */
  var re = /<div class="carousel-item\b([^"]*)"[^>]*>/gi, m, starts = [];
  re.lastIndex = inner;
  while ((m = re.exec(html))) starts.push({ at: m.index, cls: m[1] });
  starts.forEach(function (s, idx) {
    var sp2 = i18nDivSpan_(html, s.at), ch = html.slice(sp2[0], sp2[1]);
    if (/hasOption/.test(s.cls)) {
      var sl = { k: o.slides.length, t: (seq ? tl[o.slides.length] : titles[String(idx)]) || '', video: '', img: '', h: '', imgs: [], ings: [] };
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
/* 原文指紋：只看「要翻的字」（名稱、保存方式、每頁標題＋內文）一樣＝不重翻。
   2026-11-05 經營者：前一天改英文食譜的份量，隔天 155 組外語全部下架、外國客人不能用 →
   食材份量／名稱／單位、圖片、影片不算改版（食材走詞典、份量照原文數字、圖片影片照原文網址，平板組畫面時都拿最新原文） */
function i18nFp_(src) {
  return i18nHash_(JSON.stringify([src.name, src.pres, src.slides.map(function (s) { return [s.t, s.h]; })]));
}
function i18nSlideKey_(s) { return i18nHash_(JSON.stringify([s.t || '', s.h || ''])); }   /* 一頁的「要翻的字」：沒變的頁沿用舊譯文 */

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

/* ---------------- 固定譯文（2026-11-16 經營者：「要，做固定」） ----------------
   全部食譜共用、一字不差的頁（例：第 4 步「量匙怎麼舀」），各支各翻會出現幾十種寫法，還會翻錯（詰め方、盛り付け方）。
   英文標題相同、內文含每一句 en → 不送 AI、不給譯者，直接換成這裡的譯文（HTML 標籤、圖片位置照原文）；
   英文只要改了一個字就對不到 → 照舊交給 AI 翻（要更新這裡的譯文）。自動翻譯（i18nWork_）、Claude 投遞口（i18nPull_／i18nPut_）都套用。 */
var I18N_FIXED = [
  { t: 'How to Fill Your Spoon', en: ['Scoop and gently shake: level for 1 spoon, half full for ½ spoon.', 'Swipe left to start making.'],
    zt: '量匙怎麼舀', zh: ['舀起後輕輕抖一抖，一平匙齊平、半平匙裝一半。', '左滑開始製作'], tr: {   /* zt／zh：中文原文的同一頁（2026-11-24） */
    en: { t: 'How to Fill Your Spoon', s: ['Scoop and gently shake: level for 1 spoon, half full for ½ spoon.', 'Swipe left to start making.'] },
    ja: { t: '計量スプーンのすくい方', s: ['すくって軽く振り、1杯はすりきり、½杯は半分くらいまで入れます。', '左にスワイプして作り始めましょう。'] },
    ko: { t: '계량스푼 뜨는 방법', s: ['재료를 떠서 살살 흔들어 주세요. 1스푼은 가장자리에 맞춰 평평하게, ½스푼은 반 정도만 채우면 돼요.', '왼쪽으로 넘겨 만들기를 시작하세요.'] },
    vi: { t: 'Cách múc muỗng đong', s: ['Múc nguyên liệu rồi lắc nhẹ: 1 muỗng thì ngang bằng miệng muỗng, ½ muỗng thì khoảng một nửa muỗng là được.', 'Vuốt sang trái để bắt đầu làm bánh.'] },
    id: { t: 'Cara Mengisi Sendok Takar', s: ['Ambil bahan dengan sendok, lalu goyangkan perlahan: 1 sendok rata dengan bibir sendok, ½ sendok cukup sekitar setengahnya.', 'Geser ke kiri untuk mulai membuat.'] },
    th: { t: 'วิธีตักด้วยช้อนตวง', s: ['ตักวัตถุดิบแล้วเขย่าเบา ๆ: 1 ช้อนให้เสมอขอบช้อน ส่วน ½ ช้อนตักประมาณครึ่งช้อนก็พอ', 'ปัดไปทางซ้ายเพื่อเริ่มทำ'] } } }
];
function i18nFixedSlide_(s, lang, zh) {   /* 原文這一頁有固定譯文 → {t, h}；沒有 → null。zh＝中文原文（2026-11-24） */
  if (!s) return null;
  var h = String(s.h || ''), t = String(s.t || '');
  for (var i = 0; i < I18N_FIXED.length; i++) {
    var f = I18N_FIXED[i], tr = f.tr[lang], ft = zh ? f.zt : f.t, fs = zh ? f.zh : f.en;
    if (!tr || !ft || t !== ft || !fs.every(function (x) { return h.indexOf(x) >= 0; })) continue;
    fs.forEach(function (x, j) { h = h.split(x).join(tr.s[j]); });
    return { t: tr.t, h: h };
  }
  return null;
}
/* 一頁譯文對原文：回傳問題清單（空＝通過） */
function i18nCheckSlide_(s, t, okx, zl) {   /* zl：中文原文時給語言（數字照中文數字比、檢查有沒有留中文字；2026-11-24） */
  var bad = [];
  if (!t) return ['少了這一頁'];
  if (s.h && !i18nPlain_(t.h) && !i18nPh_(t.h)) bad.push('內文是空的');
  var d = zl ? i18nZNumDiff_(s.h, t.h, okx) : i18nNumDiff_(s.h, t.h, okx);
  if (zl && i18nZHan_(i18nPlain_(t.h) + ' ' + String(t.t || ''), zl)) bad.push('譯文還有中文字');
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
    .map(function (r) { return { id: r[ix.recipe_id], name: r[ix.name], upd: r[ix.upd] }; })
    .filter(function (r) { return !i18nZh_()[String(r.id).toLowerCase()]; })
    .concat(i18nZSources_())   /* 中文原文（2026-11-24，i18n_zbase.gs）：i18n_zsrc 打開的才收 */
    .reduce(function (a, r) { if (!a.some(function (x) { return x.id === r.id; })) a.push(r); return a; }, [])
    .concat(I18N_ZH.map(function (z) { return { id: String(z.id).toLowerCase(), name: z.name, upd: '中文試看' }; }))
    .reduce(function (a, r) { if (!a.some(function (x) { return x.id === r.id; })) a.push(r); return a; }, []);   /* 中文試看那支同時是中文原文 → 以中文原文為準（平板另外多一格中文） */
}
/* 中文試看（2026-11-17 經營者：「放到給外國人使用的外語食譜，開一個中文的類別，我可以在外語食譜中看這個食譜」）：
   這裡列的後台食譜（不論啟用或停用）直接用後台中文原文、不翻譯、不進待翻清單；平板語言選單多一顆「中文」只列這幾支。
   要加／拿掉就改這個清單再部署；拿掉後每天重抓會把它標成下架（gone）。伺服器快照不收 zh → 平板讀這幾支時自動改問 Google。 */
var I18N_ZH = [
  { id: '8fb61bba-1139-425b-91c8-b54ef5eae96b', name: '小黑炭Oreo蛋糕 ①原始版（門市現行）' },
  { id: '852d4cb4-ddd0-4346-9d6f-151a8d59dacb', name: '小黑炭Oreo蛋糕 ②簡化版試作' },
  { id: 'fea011e8-2a6e-466a-8a3b-089d64b5d655', name: '小黑炭Oreo蛋糕 ③簡化版試作2（影片合併）' },
  { id: '05584c26-e8c5-44e0-8b15-05aaa99bf086', name: '小黑炭Oreo蛋糕 ④簡化版試作3（抹面2頁）' }
];
function i18nZh_() { var m = {}; I18N_ZH.forEach(function (z) { m[String(z.id).toLowerCase()] = z.name || '中文試看'; }); return m; }
function i18nZhOnly_(id, ZM) { id = String(id || '').toLowerCase(); return !!i18nZh_()[id] && !(ZM || i18nZMap_())[id]; }   /* 只當中文試看、不翻譯（同時是中文原文的照樣翻，2026-11-24） */
function i18nFetchPages_(ids) {
  var sess = rbSession_();
  var mk = function () { var c = rbJarStr_(sess.jar); return ids.map(function (id) { return { url: RB_BASE + '/steps/recipedetails/' + id, method: 'get', muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA, Cookie: c } }; }); };
  var rs = UrlFetchApp.fetchAll(mk());
  if (rs.some(function (r) { return r.getResponseCode() === 302 && /Identity\/Account\/Login/i.test(String(r.getHeaders().Location || r.getHeaders().location || '')); })) {
    rbDropSession_(); sess.jar = rbFreshSession_().jar; rs = UrlFetchApp.fetchAll(mk());
  }
  rbSaveSession_(sess);
  return rs.map(function (r, i) {
    var code = r.getResponseCode(), loc = String(r.getHeaders().Location || r.getHeaders().location || '');
    if (code === 302 && I18N_PICK_RE.test(loc)) {   /* 自選版（2026-11-26）：後台先跳選擇頁 → 讀全部選項的完整食譜＋每個選項的步驟 */
      try { return i18nPickFetch_(ids[i], sess); } catch (e) { return { code: 0, err: '自選版：' + errMsg_(e), text: '' }; }
    }
    return { code: code, text: code === 200 ? r.getContentText() : '' };
  });
}

/* ---------------- 自選版（客人先選蛋糕體／夾餡，2026-11-26） ----------------
   經營者：「生日蛋糕只有海綿蛋糕有外語版，價購內餡的版本沒有」→「照方案 A 做」。
   後台：/steps/recipedetails/{id} → 302 到 /Steps/Picker/{id}（每組一個下拉選單）→ 按下一步開
         /steps/recipedetailscode/{id}?stepList=選項值,…（共用步驟＋選到的選項步驟）。stepList 帶全部選項＝全部步驟。
   做法：原文＝全部選項的完整食譜；每個選項各抓一次（這組只帶它、其他組全帶），跟完整的比對 → 這個選項自己的步驟（src.pick[].opts[].ks）。
   選項名稱走詞典（kind＝pick，5 支同一組選項只翻一次）；平板先出選擇頁，選完只顯示共用＋選到的步驟。 */
var I18N_PICK_RE = /\/Steps\/Picker\//i;
function i18nPickGroups_(html) {
  var out = [], re = /<label[^>]*class="control-label"[^>]*>([\s\S]*?)<\/label>\s*<select\b[^>]*>([\s\S]*?)<\/select>/gi, m;
  while ((m = re.exec(String(html || '')))) {
    var g = { label: i18nTxt_(m[1]), opts: [] }, orx = /<option\b[^>]*value="([^"]+)"[^>]*>([\s\S]*?)<\/option>/gi, o;
    while ((o = orx.exec(m[2]))) g.opts.push({ v: o[1], label: i18nTxt_(o[2]) });
    if (g.label && g.opts.length) out.push(g);
  }
  return out;
}
function i18nPickSig_(s) { return JSON.stringify([s.t, s.h, s.video, s.img, (s.ings || []).map(function (g) { return [g.n, g.q, g.u]; })]); }
function i18nPickFetch_(id, sess) {
  var H = { 'User-Agent': RB_UA, Cookie: rbJarStr_(sess.jar) }, opt = function (u) { return { url: u, method: 'get', muteHttpExceptions: true, followRedirects: false, headers: H }; };
  var pr = UrlFetchApp.fetch(RB_BASE + '/Steps/Picker/' + id, opt(''));
  if (pr.getResponseCode() !== 200) throw fail_('選擇頁回 HTTP ' + pr.getResponseCode(), 'rbnet');
  var G = i18nPickGroups_(pr.getContentText());
  if (!G.length) throw fail_('選擇頁讀不到選項（後台可能改版）', 'rbpage');
  var u = function (list) { return RB_BASE + '/steps/recipedetailscode/' + id + '?stepList=' + list.join(','); }, all = [], urls = [];
  G.forEach(function (g) { g.opts.forEach(function (o) { all.push(o.v); }); });
  urls.push(u(all));
  G.forEach(function (g, gi) { g.opts.forEach(function (o) { var l = []; G.forEach(function (h, hi) { if (hi === gi) l.push(o.v); else h.opts.forEach(function (x) { l.push(x.v); }); }); urls.push(u(l)); }); });
  var rs = UrlFetchApp.fetchAll(urls.map(opt));
  rs.forEach(function (r) { if (r.getResponseCode() !== 200) throw fail_('選項步驟頁回 HTTP ' + r.getResponseCode(), 'rbnet'); });
  var k = 1;
  G.forEach(function (g) { g.opts.forEach(function (o) { o.sigs = i18nParse_(rs[k++].getContentText(), id, true).slides.map(i18nPickSig_); }); });
  return { code: 200, text: rs[0].getContentText(), pick: G };
}
/* 完整原文 src＋每個選項抓到的頁 → src.pick＝[{label, opts:[{v, label, ks:[頁碼]}]}]；對不上就丟錯（這支原文標失敗，不上平板） */
function i18nPickApply_(src, G) {
  var full = src.slides.map(i18nPickSig_), used = {};
  var has = function (sigs) {   /* 選項那次的頁 → 在完整原文的哪幾頁（照順序對） */
    var set = {}, j = 0;
    sigs.forEach(function (s) { while (j < full.length && full[j] !== s) j++; if (j >= full.length) throw fail_('自選版：選項的步驟頁對不上完整食譜', 'rbpage'); set[j++] = 1; });
    return set;
  };
  src.pick = G.map(function (g) {
    var P = g.opts.map(function (o) { return has(o.sigs); });
    var opts = g.opts.map(function (o, i) {
      var ks = [];
      full.forEach(function (x, kk) { if (P[i][kk] && !P.some(function (q, j) { return j !== i && q[kk]; })) ks.push(kk); });
      ks.forEach(function (kk) { if (used[kk]) throw fail_('自選版：第 ' + (kk + 1) + ' 頁同時屬於兩個選項', 'rbpage'); used[kk] = 1; });
      return { v: o.v, label: o.label, ks: ks };
    });
    opts.forEach(function (o, i) {   /* 檢查：只選這個時的頁數＝全部 − 同組其他選項的頁 */
      var want = full.length - opts.reduce(function (n, x, j) { return n + (j === i ? 0 : x.ks.length); }, 0);
      if (g.opts[i].sigs.length !== want) throw fail_('自選版：選項「' + o.label + '」頁數對不上（' + g.opts[i].sigs.length + '／' + want + '）', 'rbpage');
    });
    if (opts.every(function (o) { return !o.ks.length; })) throw fail_('自選版：「' + g.label + '」每個選項都找不到自己的步驟', 'rbpage');
    return { label: g.label, opts: opts };
  });
  return src;
}
/* 更新原文：後台更新日期變了（或還沒抓過、上次失敗）才重抓；回傳 {checked, fetched, changed, gone, errors} */
function i18nRefresh_(force, deadline) {
  var list = i18nSources_(), meta = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows, have = {}, res = { checked: list.length, fetched: 0, changed: 0, gone: 0, errors: [] };
  meta.forEach(function (r) { have[r.id] = r; });
  var todo = list.filter(function (s) { var h = have[s.id]; return force || !h || h.upd !== s.upd || h.state !== 'ok'; });
  for (var i = 0; i < todo.length; i += 5) {
    if (deadline && Date.now() > deadline) { res.errors.push('時間到，剩 ' + (todo.length - i) + ' 支下次再抓'); break; }
    var batch = todo.slice(i, i + 5), pages = i18nFetchPages_(batch.map(function (s) { return s.id; }));
    var parsed = batch.map(function (s, j) { try { return pages[j].code === 200 ? (pages[j].pick ? i18nPickApply_(i18nParse_(pages[j].text, s.id, true), pages[j].pick) : i18nParse_(pages[j].text, s.id)) : null; } catch (e) { return e; } });   /* 自選版：附上每個選項的步驟（2026-11-26） */
    var zu = []; batch.forEach(function (s, j) { var o = parsed[j]; if (s.base === 'zh' && o && o.slides) zu = zu.concat(i18nZMediaUrls_(o)); });
    var md5 = zu.length ? i18nZMd5_(zu) : {};   /* 中文原文：教學影片、烤熟測試圖的指紋（換外國客人版用；2026-11-24） */
    batch.forEach(function (s, j) {
      var pg = pages[j], row = { id: s.id, name: s.name, upd: s.upd, at: now_() };
      try {
        if (pg.code !== 200) throw fail_(pg.err || ('後台回 HTTP ' + pg.code + '（可能已下架或上架日期不在期間內）'), 'rbnet');
        if (parsed[j] instanceof Error || !parsed[j]) throw parsed[j] || fail_('讀不到平板頁', 'rbpage');
        var src = parsed[j];
        if (s.base === 'zh') { src.old = s.old || ''; i18nZPrep_(src, md5).forEach(function (w) { res.errors.push(s.name + '：' + w); }); }
        var fp = i18nFp_(src);
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
function i18nArea_(code, lang) {   /* 【Area A】／【A區】（中文原文）→ 各語言 */
  var m = String(code).match(/^(?:Area\s*([A-Z0-9]+)|([A-Z0-9]+)\s*區)$/i); if (!m) return null;
  var x = (m[1] || m[2]).toUpperCase();
  return { en: 'Area ' + x, ja: x + 'エリア', ko: x + ' 구역', vi: 'Khu ' + x, id: 'Area ' + x, th: 'โซน ' + x }[lang];
}
function i18nSplitName_(n) { var m = String(n).match(/^【([^】]*)】\s*(.*)$/); return m ? { pre: m[1], core: m[2] } : { pre: '', core: String(n) }; }
function i18nTerms_(src) {   /* 這支食譜要查詞典的詞：[{kind, en}] */
  var seen = {}, out = [], add = function (kind, en) { en = rbTrim_(en); if (!en) return; var k = kind + '|' + en.toLowerCase(); if (seen[k]) return; seen[k] = 1; out.push({ kind: kind, en: en }); };
  src.slides.forEach(function (s) { s.ings.forEach(function (g) { add('item', i18nSplitName_(g.n).core); add('unit', g.u); add('area', g.an); }); });
  (src.pick || []).forEach(function (g) { add('pick', g.label); g.opts.forEach(function (o) { add('pick', o.label); }); });   /* 自選版選擇頁的組名、選項名（2026-11-26） */
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
  var c = CacheService.getScriptCache(), hit = c.get('i18n:guide:v2');
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var map = {};
  try {
    var d = JSON.parse(UrlFetchApp.fetch(I18N_GUIDE_API, { muteHttpExceptions: true, followRedirects: true }).getContentText());
    (d.items || []).forEach(function (it) {
      var k = i18nNameKey_(it.en); if (!k) return; var o = map[k] = {}; I18N_LANGS.forEach(function (l) { var t = it.t && it.t[l]; if (t && t[0]) o[l] = t[0]; });
      o.en = rbTrim_(it.en || '');   /* 英文品名（中文原文翻英文時用；舊英文原文的流程只看日韓越印泰，不受影響） */
      var zk = i18nNameKey_(it.zh); if (zk && !map['z:' + zk]) map['z:' + zk] = o;   /* 中文品名也能查（2026-11-24 中文原文） */
    });
    c.put('i18n:guide:v2', JSON.stringify(map), 21600);
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

/* 翻一支食譜的一個語言。
   opt.auto（Google 排程自動翻，2026-11-05 經營者：「改 Google，花 API」，Mac 不用開）：
     用 Sonnet（I18N_AUTO_MODEL）；只翻改到的頁（放行過的上一版裡字沒變的頁、名稱、保存方式直接沿用）；
     翻完 AI 回讀（i18nReview_），有嚴重問題修一輪再回讀；硬檢查＋回讀都過 → 自動放行（放行人「AI 自動檢查」）；
     沒過 → 存成待檢查，平板照舊顯示上一版。opt.deadline：這次執行最晚做到幾點（GAS 一次最多 6 分鐘）。
   沒有 opt.auto（核准頁「🤖 翻譯這一份」）＝照舊整份重翻、待核准。 */
var I18N_AUTO_MODEL = 'claude-sonnet-5-5';
function i18nWork_(id, lang, by, redict, opt) {
  opt = opt || {};
  var t0 = Date.now(), src = i18nSrc_(id), dl = opt.deadline || (t0 + I18N_BUDGET + 60000);
  if (!src) throw fail_('還沒有這支食譜的原文（或上次讀取失敗）', 'nosrc');
  if (i18nZh_()[String(id).toLowerCase()]) throw fail_('這支是中文試看用，不翻譯', 'bad');
  if (src.base === 'zh') throw fail_('中文原文只由 Claude Code 翻（2026-11-23 不用 API）', 'bad');
  var cred = i18nCred_(); if (opt.auto) cred.model = I18N_AUTO_MODEL;
  var dict = i18nEnsureDict_(i18nTerms_(src), lang, cred, t0, redict), D = i18nDictRows_();
  var guide = i18nGuideNames_()[i18nNameKey_(src.name)] || {}, local = guide[lang] || '';
  var key = id + '|' + lang, rm = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), prev = rm.rows.filter(function (r) { return r.key === key; })[0];
  var base = opt.auto && prev ? i18nLiveContent_(rm.sh, prev) : null, reuse = base ? i18nAlign_(base, src, true) : [];
  (opt.redo || []).forEach(function (k) { if (reuse[k]) reuse[k] = null; });   /* 測試用：指定的頁強制重翻 */
  var todo = []; src.slides.forEach(function (s, i) { if (!reuse[i]) todo.push(i); });
  var fixed = {}; todo = todo.filter(function (i) { var f = i18nFixedSlide_(src.slides[i], lang); if (f) { fixed[i] = f; return false; } return true; });   /* 固定譯文的頁不送 AI（2026-11-16） */
  var nameKeep = !!(base && base.name && base.snk === i18nHash_(String(src.name || ''))), presKeep = !!(base && base.spk === i18nHash_(String(src.pres || '')));
  var oldAt = base && base.slides && base.slides.length === src.slides.length ? base.slides : null;   /* 頁數沒變：同頁碼的舊譯文給 AI 當底 */
  /* 標題：上一版用過的照用，只翻新出現的 */
  var tt = {};
  if (base && base.titles) src.slides.forEach(function (s) { if (s.t && base.titles[s.t]) tt[s.t] = base.titles[s.t]; });
  var needT = todo.map(function (i) { return src.slides[i]; }).filter(function (s) { return s.t && !tt[s.t]; });
  if (needT.length) { var nt = i18nTitles_({ name: src.name, slides: needT }, lang, cred, local); Object.keys(nt).forEach(function (k) { tt[k] = nt[k]; }); }
  var sys = i18nSystem_(lang, src, D, local, tt), model = cred.model;
  var out = { name: nameKeep ? String(base.name || '') : '', pres: presKeep ? String(base.pres || '') : '', slides: [] };
  reuse.forEach(function (s, i) { if (s) out.slides[i] = { t: String(s.t || ''), h: String(s.h || '') }; });
  Object.keys(fixed).forEach(function (i) { out.slides[i] = fixed[i]; tt[src.slides[i].t] = fixed[i].t; });
  var chunks = [], cur = [], n = 0;   /* 只把要翻的頁分段：每段最多 4 頁、5,000 字 */
  todo.forEach(function (i) {
    var len = (src.slides[i].t || '').length + (src.slides[i].h || '').length;
    if (cur.length && (cur.length >= I18N_CHUNK || n + len > I18N_CHUNK_CHARS)) { chunks.push(cur); cur = []; n = 0; }
    cur.push(i); n += len;
  });
  if (cur.length) chunks.push(cur);
  var needName = !nameKeep && !local, needPres = !presKeep && !!src.pres;
  if (!chunks.length && (needName || needPres)) chunks.push([]);
  var reqOf = function (ix, first, note) {
    var u = { name: first && needName ? src.name : '', pres: first && needPres ? src.pres : '', slides: ix.map(function (i) {
      var o = { k: i, t: src.slides[i].t, h: src.slides[i].h }; if (oldAt && oldAt[i]) o.prev = { t: oldAt[i].t, h: oldAt[i].h }; return o; }) };
    var hint = oldAt ? '\nSome pages have "prev": the approved translation of an older version of that page. Use it as the base: keep its wording where the English did not change, and only change what the English changed. Do not return "prev".' : '';
    return i18nAiReq_(sys + hint + (note ? '\n' + note : ''), 'Translate this JSON. Return the same structure.\n' + JSON.stringify(u), I18N_SCHEMA, cred);
  };
  var take = function (r, ix) {
    if (!r.ok) return false; model = r.model || model;
    if (r.json.name && needName) out.name = r.json.name; if (r.json.pres && needPres) out.pres = r.json.pres;
    (r.json.slides || []).forEach(function (s) { if (ix.indexOf(s.k) >= 0) out.slides[s.k] = { t: s.t, h: s.h }; });
    return true;
  };
  var tlimit = opt.auto ? dl - 100000 : dl;   /* 自動模式留 100 秒給回讀 */
  var run = function (jobs) {   /* jobs: [{ix, first, note}] → 並行（每波 6 個）；整波逾時就改一個一個送 */
    for (var w = 0; w < jobs.length; w += I18N_WAVE) {
      if (Date.now() > tlimit) return false;
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
  var okx = {};   /* 詞典譯名、在地品名自帶的數字不算多（同投遞口） */
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D, x.kind, x.en, lang); if (tr) i18nNums_(tr).forEach(function (q) { okx[q] = 1; }); });
  if (local) i18nNums_(local).forEach(function (q) { okx[q] = 1; });
  var fixT = function (ix) { ix.forEach(function (i) { var s = src.slides[i]; if (out.slides[i] && tt[s.t]) out.slides[i].t = tt[s.t]; }); };
  fixT(todo);
  /* 檢查（只看這次翻的頁）→ 有問題的頁一頁一頁重翻一次（附上問題） */
  var probs = function () { var p = {}; todo.forEach(function (i) { var b = i18nCheckSlide_(src.slides[i], out.slides[i], okx); if (b.length) p[i] = b; }); return p; };
  var P = probs(), redo = Object.keys(P).map(Number);
  if ((needName && !out.name) || (needPres && !out.pres)) redo.unshift(-1);
  if (redo.length && Date.now() < tlimit) {
    run(redo.map(function (i) {
      if (i < 0) return { ix: [], first: true };
      return { ix: [i], first: false, note: 'IMPORTANT: a previous translation of this page had problems: ' + P[i].join('; ') + '. Keep every number and [[IMG]] placeholder exactly.' };
    }));
    fixT(todo); P = probs();
  }
  if (local) out.name = local + '｜' + src.name.split('｜').slice(1).join('｜');
  var bad = Object.keys(P);
  var done = src.slides.every(function (s, i) { return !!out.slides[i]; }) && !!out.name;
  var status = !done ? 'fail' : bad.length ? 'review' : 'ok';
  var check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；');
  if (!done) check = '沒翻完（' + src.slides.filter(function (s, i) { return !out.slides[i]; }).length + ' 頁沒有回來）' + (check ? '；' + check : '');
  /* AI 回讀（只看這次翻的頁＋新翻的名稱／保存方式）；有嚴重問題 → 那幾頁帶著問題重翻一次 → 再回讀一次 */
  var rv = null, rvNote = '';
  if (opt.auto && status === 'ok') {
    try {
      rv = i18nReview_(src, out, todo, needName, needPres, lang, D, local, cred);
      if (rv.block.length && Date.now() < dl - 70000) {
        var bp = [], why = {};
        rv.block.forEach(function (x) { var k = Number(x.k); if (!(k >= 0 && k < src.slides.length) || todo.indexOf(k) < 0) return; if (!why[k]) { why[k] = []; bp.push(k); } why[k].push(x.what); });
        run(bp.map(function (i) { return { ix: [i], first: false, note: 'IMPORTANT: a reviewer found these problems on this page: ' + why[i].join('; ') + '. Fix them. Keep every number and [[IMG]] placeholder exactly.' }; }));
        fixT(bp); P = probs();
        if (Object.keys(P).length) { status = 'review'; check = Object.keys(P).map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；'); }
        else { var rv2 = i18nReview_(src, out, bp, false, false, lang, D, local, cred), other = rv.block.filter(function (x) { return bp.indexOf(Number(x.k)) < 0; }); rv.block = other.concat(rv2.block); rv.fixed = bp.length; }
      }
      if (status === 'ok' && rv.block.length) { status = 'review'; check = '回讀：' + rv.block.map(function (x) { return (Number(x.k) >= 0 ? '第' + (Number(x.k) + 1) + '頁 ' : '名稱／保存 ') + x.what; }).join('；'); }
      rvNote = '回讀 ' + (rv.block.length ? '有 ' + rv.block.length + ' 個問題' : '通過') + (rv.fixed ? '（修過 ' + rv.fixed + ' 頁）' : '') + (rv.minor ? '、小建議 ' + rv.minor : '');
    } catch (e) { status = 'fail'; check = '回讀沒做完：' + errMsg_(e); }
  }
  var sec = Math.round((Date.now() - t0) / 1000);
  var tries = prev && prev.fp === src.fp ? Number(prev.tries || 0) + 1 : 1;
  var approved = !!opt.auto && status === 'ok';
  var json = i18nJson_(src, { id: id, lang: lang, model: model, name: out.name, pres: out.pres, slides: out.slides, titles: tt });
  var kept = false;
  i18nLock_(function () { var meta = { key: key, id: id, lang: lang, fp: src.fp, status: status, tries: tries, model: model, check: check, sec: sec, at: now_() }; i18nSaveRcp_(meta, json, approved, 'AI 自動檢查'); kept = /^L:/.test(meta.appr_fp); });
  i18nLog_(by, 'work', id, lang, status, (opt.auto ? '翻 ' + todo.length + '／' + src.slides.length + ' 頁；' : '') + (dict.added ? '詞典新增 ' + dict.added + '；' : '') + (rvNote ? rvNote + '；' : '') + (approved ? '自動放行；' : kept ? '平板照舊顯示上一版；' : '') + (check || '通過'), sec);
  return { id: id, lang: lang, name: src.name, status: status, check: check, sec: sec, slides: src.slides.length, translated: todo.length, chunks: chunks.length, dict: dict, model: model, approved: approved, review: rvNote };
}

/* AI 回讀（2026-11-05，取代 Claude Code 的獨立回讀員；規則同 tools/i18n_claude/REVIEW.md）：回 {block:[{k,what}], minor} ；k＝頁碼（從 0），-1＝名稱或保存方式 */
var I18N_REVIEW_SCHEMA = { type: 'object', properties: { issues: { type: 'array', items: { type: 'object', properties: { k: { type: 'integer' }, severity: { type: 'string', enum: ['block', 'minor'] }, what: { type: 'string' } }, required: ['k', 'severity', 'what'], additionalProperties: false } } }, required: ['issues'], additionalProperties: false };
function i18nReview_(src, out, pages, withName, withPres, lang, D, local, cred) {
  var gl = [], seen = {};
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D, x.kind, x.en, lang); if (tr && !seen[x.en]) { seen[x.en] = 1; gl.push('- ' + x.en + ' → ' + tr); } });
  var sys = ['You are the independent second reviewer of a translated in-store tablet recipe of a DIY baking studio in Taiwan (customers bake by themselves, following the tablet page by page).',
    'You did not translate it. Compare each page\'s English source (en_t, en_h) with the ' + I18N_LABEL[lang] + ' translation (tr_t, tr_h). This is the last check before customers see it.',
    'Report severity "block" only for problems that would make a customer do something wrong or not understand:',
    '* a sentence missing or added; wrong meaning (verbs like mix / whisk / fold / pipe / bake / cool / refrigerate / freeze / sift, ingredient, tool, container, direction);',
    '* a number, weight, time, temperature, speed, count or unit that differs from the English (Thai must use Arabic numerals 0-9);',
    '* a safety warning (hot, gloves, knife, freezer) missing; an English or Chinese sentence left untranslated, garbled text, broken HTML;',
    '* a "$" price turned into a local currency (đồng, 円, 원, บาท, rupiah); "10-dollar coin" means a Taiwan NT$10 coin, not US dollars;',
    '* an ingredient/tool name that contradicts the glossary or the English (whipping cream is NOT custard; electronic mixer = stand mixer, not hand mixer; pomelo is not yuzu; batter is not flour).',
    'Everything else (slightly stiff wording, could be more natural) is "minor". Do not report differences in style that keep the meaning.',
    'k is the page number given. Use k = -1 for problems in the dessert name or storage advice. Return an empty list if there are no problems. Write "what" briefly in Traditional Chinese.',
    'Glossary (English → ' + I18N_LABEL[lang] + '):', gl.join('\n') || '(none)'].join('\n');
  var groups = [], cur = [], n = 0;
  pages.forEach(function (i) {
    var len = (src.slides[i].h || '').length + (out.slides[i] ? out.slides[i].h.length : 0);
    if (cur.length && (cur.length >= 8 || n + len > 12000)) { groups.push(cur); cur = []; n = 0; }
    cur.push(i); n += len;
  });
  if (cur.length) groups.push(cur);
  if (!groups.length && (withName || withPres)) groups.push([]);
  var reqs = groups.map(function (g, gi) {
    var u = { pages: g.map(function (i) { var t = out.slides[i] || {}; return { k: i, en_t: src.slides[i].t, en_h: src.slides[i].h, tr_t: t.t || '', tr_h: t.h || '' }; }) };
    if (gi === 0 && withName) { u.name_en = src.name; u.name_tr = out.name; }
    if (gi === 0 && withPres) { u.pres_en = src.pres; u.pres_tr = out.pres; }
    return i18nAiReq_(sys, 'Review this JSON.\n' + JSON.stringify(u), I18N_REVIEW_SCHEMA, cred);
  });
  var res = { block: [], minor: 0 };
  for (var w = 0; w < reqs.length; w += I18N_WAVE) {
    var wave = reqs.slice(w, w + I18N_WAVE), rs = null;
    try { rs = UrlFetchApp.fetchAll(wave); } catch (e) { rs = null; }
    wave.forEach(function (rq, q) {
      var r = rs ? i18nAiRes_(rs[q], cred) : null;
      if (!r || !r.ok) { try { r = i18nAiRes_(UrlFetchApp.fetch(I18N_AI_URL, rq), cred); } catch (e2) { r = { ok: false, err: errMsg_(e2) }; } }
      if (!r.ok) throw fail_('回讀失敗：' + r.err, 'ai');
      (r.json.issues || []).forEach(function (x) { if (x.severity === 'block') res.block.push({ k: Number(x.k), what: String(x.what || '').slice(0, 200) }); else res.minor++; });
    });
  }
  return res;
}

function i18nLive_(r, fp) { return !!r && r.fp === fp && r.status !== 'fail' && r.appr_fp === fp && !!r.appr_by; }   /* 這版翻譯已核准（跟現在的原文同一版） */

/* ---------------- 上一版保底（2026-11-05 經營者：新譯文還沒好之前，平板先顯示上一版，頁面不提示；改完再換新版） ----------------
   平板看得到＝有一份放行過的譯文：
     ① appr_fp＝fp：這一列的譯文本身放行過（原文改版後仍照舊顯示，對不上的頁見 i18nAlign_）；
     ② appr_fp＝'L:'＋舊指紋：新譯文還沒放行（待檢查），放行過的上一版收在 JSON 的 live 裡。 */
function i18nShowable_(r) { return !!r && !!r.appr_by && (/^L:/.test(r.appr_fp) || (r.appr_fp === r.fp && r.status !== 'fail')); }
function i18nLiveContent_(sh, r, j) {   /* 放行過的那份內容 {fp, name, pres, slides, titles, sk, snk, spk} */
  if (!i18nShowable_(r)) return null;
  j = j || i18nReadJ_(sh, r._r, I18N_RCP_HEAD); if (!j) return null;
  if (/^L:/.test(r.appr_fp)) return j.live || null;
  return { fp: r.fp, name: j.name, pres: j.pres, slides: j.slides || [], titles: j.titles || {}, sk: j.sk || null, snk: j.snk || '', spk: j.spk || '' };
}
/* 譯文 JSON：附上每頁原文的鍵（sk）、名稱／保存方式的鍵，下次原文改版時才知道哪幾頁沒變 */
function i18nJson_(src, o) {
  o.fp = src.fp; o.at = now_(); o.sk = src.slides.map(i18nSlideKey_); o.snk = i18nHash_(String(src.name || '')); o.spk = i18nHash_(String(src.pres || ''));
  return o;
}
/* 存一份新譯文（要在鎖裡呼叫）：沒放行、但上一版放行過 → 上一版收進 live、放行欄照舊（平板繼續顯示上一版） */
function i18nSaveRcp_(meta, json, approved, by) {
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), prev = m.rows.filter(function (r) { return r.key === meta.key; })[0];
  if (approved) { meta.appr_by = by; meta.appr_at = now_(); meta.appr_fp = meta.fp; }
  else {
    var live = prev ? i18nLiveContent_(m.sh, prev) : null;
    if (live) { json.live = live; meta.appr_by = prev.appr_by; meta.appr_at = prev.appr_at; meta.appr_fp = 'L:' + live.fp; }
    else { meta.appr_by = ''; meta.appr_at = ''; meta.appr_fp = ''; }
  }
  return i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', meta, json);
}
/* 放行過的譯文 → 現在原文每一頁：同一頁字沒變的直接用；只有沒變的才算（exact）；
   不 exact 時，頁數沒變的照頁碼用上一版，對不上的頁用英文原文（極少見：頁數變了又改字） */
function i18nAlign_(c, src, exact) {
  var cs = c.slides || [];
  if (c.fp === src.fp && cs.length === src.slides.length) return cs.slice();
  var by = {}; (c.sk || []).forEach(function (k, j) { if (by[k] == null && cs[j]) by[k] = j; });
  var same = cs.length === src.slides.length;
  return src.slides.map(function (s, i) {
    var j = by[i18nSlideKey_(s)]; if (j != null) return cs[j];
    if (exact) return null;
    return same && cs[i] ? cs[i] : { t: s.t, h: s.h };
  });
}
function i18nState_(r, fp) { if (!r || r.fp !== fp) return ''; if (i18nLive_(r, fp)) return 'live'; return r.status === 'ok' ? 'wait' : r.status; }   /* live 已放行／wait 自動檢查過、待核准／review 待檢查／fail */

/* 只重新跑自動檢查（不重翻、不花錢）：檢查規則改了以後用 */
function i18nRecheck_(id, lang, by) {
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  if (!r || r.fp !== src.fp || r.status === 'fail') throw fail_('這份翻譯不能重新檢查（沒翻好或原文改版）', 'notyet');
  var tr = i18nReadJ_(m.sh, r._r, I18N_RCP_HEAD), P = {};
  src.slides.forEach(function (s, i) { var b = i18nCheckSlide_(s, (tr.slides || [])[i], null, src.base === 'zh' ? lang : ''); if (b.length) P[i] = b; });
  var bad = Object.keys(P), check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；'), status = bad.length ? 'review' : 'ok';
  i18nLock_(function () { i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', { key: r.key, status: status, check: check }, null); });
  i18nLog_(by, 'recheck', id, lang, status, (r.status !== status ? r.status + '→' + status + '；' : '') + (check || '通過'));
  return { id: id, lang: lang, before: r.status, status: status, check: check };
}

/* ---------------- 勾選：哪幾支給外國客人用（2026-10-26） ---------------- */
/* 回傳 null＝還沒人勾過（分頁不存在）→ 全部當作有勾；否則 {id: true/false}。只讀，不建分頁 */
function i18nPickMap_() {
  var sh = ss_().getSheetByName('i18n_pick'); if (!sh) return null;
  var m = {}, n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, I18N_PICK_HEAD.length).getDisplayValues().forEach(function (r) { if (r[0]) m[r[0]] = r[2] === 'Y'; });
  return m;
}
function i18nOn_(pm, id) { return !pm || pm[id] !== false; }   /* 2026-11-04：新收入的（不在表上）自動算有勾，只有取消勾（N）的才關 */
function i18nPick_(p, auth) {
  i18nBoss_(auth); I18N_HELD = true;   /* 走一般寫入路徑（已在全系統的鎖裡） */
  var ids = (Array.isArray(p.ids) ? p.ids : [p.id]).map(function (x) { return String(x || '').trim().toLowerCase(); }).filter(String);
  if (!ids.length) throw fail_('沒有指定食譜', 'bad');
  var on = !!Number(p.on) || p.on === true, who = auth.role + (p.who ? '：' + String(p.who).slice(0, 20) : ''), t = now_();
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (r) { return r.state === 'ok'; }), nm = {};
  src.forEach(function (r) { nm[r.id] = r.name; });
  ids.forEach(function (id) { if (!nm[id]) throw fail_('找不到這支食譜（' + id + '）：要是後台台灣區、啟用中的英文食譜', 'nosrc'); });
  if (!ss_().getSheetByName('i18n_pick')) {   /* 第一次勾選：先把當下所有原文記成「有勾」（＝上線前的現況），再改這次指定的 */
    var sh = i18nSheet_('i18n_pick', I18N_PICK_HEAD);
    if (src.length) sh.getRange(2, 1, src.length, I18N_PICK_HEAD.length).setNumberFormat('@').setValues(src.map(function (r) { return [r.id, r.name, 'Y', '首次建立（原本就在平板上）', t]; }));
    i18nLog_(who, 'pick-init', '', '', 'ok', '建立勾選表：' + src.length + ' 支全部記成有勾');
  }
  ids.forEach(function (id) { i18nUpsert_('i18n_pick', I18N_PICK_HEAD, 'id', { id: id, name: nm[id], on: on ? 'Y' : 'N', by: who, at: t }, null); });
  SpreadsheetApp.flush();
  ids.forEach(function (id) { i18nLog_(who, on ? 'pick' : 'unpick', id, '', 'ok', nm[id]); });
  var pm = i18nPickMap_();
  return { id: ids.join(','), msg: (on ? '勾選 ' : '取消勾選 ') + ids.join(','), data: { ids: ids, on: on, picked: src.filter(function (r) { return i18nOn_(pm, r.id); }).length, total: src.length } };   /* 2026-10-26 提速：不再算待翻清單（畫面沒用到） */
}

/* ---------------- 待辦 ---------------- */
function i18nPending_() {
  var pm = i18nPickMap_();   /* 沒勾的不翻（2026-10-26） */
  var ZM = i18nZMap_();   /* 中文原文（2026-11-24）：語言多一個 en；舊英文原文被中文版取代後不再翻 */
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (r) { return r.state === 'ok' && i18nOn_(pm, r.id) && !i18nZhOnly_(r.id, ZM); });   /* 中文試看的不翻（2026-11-17） */
  var rc = {}; i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.forEach(function (r) { rc[r.key] = r; });
  var ok = {}; src.forEach(function (s) { ok[s.id] = 1; });
  var ret = i18nZRetired_(ZM, rc, ok);
  var out = [];
  src.forEach(function (s) {
    if (ret[s.id]) return;
    var z = !!ZM[s.id];
    (z ? i18nZLangs_() : I18N_LANGS).forEach(function (l) {
      var r = rc[s.id + '|' + l], o = { id: s.id, lang: l, name: s.name, base: z ? 'zh' : 'en' };
      if (!r || r.fp !== s.fp) out.push(o);
      else if (r.status === 'fail' && Number(r.tries || 0) < I18N_TRIES) out.push(o);
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
  var pend = i18nPending_().filter(function (x) { return x.base !== 'zh' && (!id || x.id === id) && (!lang || x.lang === lang); });   /* 中文原文只由 Claude Code 翻 */
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
  try { var zp = i18nZPairDaily_(); if (zp.changed.length) i18nLog_('排程', 'zpair', zp.changed.join(','), '', 'ok', '中文原文自動配對舊英文原文：更新 ' + zp.changed.length + ' 支'); }   /* 2026-11-27：先更新配對，舊英文版被取代就不再列、不再翻 */
  catch (e) { i18nLog_('排程', 'zpair', '', '', 'error', errMsg_(e)); }
  try { ref = i18nRefresh_(true, t0 + 180000); }   /* 2026-11-11 改全部重抓：後台「新增步驟」不會改食譜更新日（11/10 一平匙說明漏了 22 支）；只抓 38 頁約 16 秒，內容沒變指紋不變、不會重翻 */ catch (e) { i18nLog_('排程', 'refresh', '', '', 'error', errMsg_(e)); return; }
  i18nLog_('排程', 'refresh', '', '', 'ok', '檢查 ' + ref.checked + '、重抓 ' + ref.fetched + '、內容有變 ' + ref.changed + '、下架 ' + ref.gone + (ref.errors.length ? '；' + ref.errors.join('；') : ''));
  if (i18nProp_('I18N_AUTO') !== '1') return;
  i18nAutoStart_();
}
/* ---------------- Google 排程自動翻（2026-11-05 經營者：「改 Google，花 API」，Mac 不用開） ----------------
   i18nDaily（每天 06:30）更新原文後 → i18nAutoStart_ 替 5 個語言各排一個「1 分鐘後跑一次」的排程（i18nAuto_ja…）；
   每個語言一次只做一組（詞典每語言共用），一次執行最多 6 分鐘：60 秒內才開新的一組，還有待翻就再排下一次，直到全部做完。
   同一語言一天最多跑 I18N_AUTO_MAX 次（保險絲：API 一直失敗時不會無限重排、無限花錢）。 */
var I18N_AUTO_MAX = 150;
function i18nDropTrig_(h) { ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === h) ScriptApp.deleteTrigger(t); }); }
function i18nAutoStart_() {
  var pend = i18nPending_(), n = 0;
  I18N_LANGS.forEach(function (l, i) {
    if (!pend.some(function (x) { return x.lang === l && x.base !== 'zh'; })) return;
    var h = 'i18nAuto_' + l; i18nDropTrig_(h); ScriptApp.newTrigger(h).timeBased().after(20000 + i * 5000).create(); n++;
  });
  i18nLog_('排程', 'auto-start', '', '', 'ok', '待翻 ' + pend.length + ' 組；開工 ' + n + ' 個語言');
  return { pending: pend.length, langs: n };
}
function i18nAutoLang_(lang) {
  var t0 = Date.now(), h = 'i18nAuto_' + lang, day = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd'), pk = 'I18N_RUNS_' + lang;
  i18nDropTrig_(h);
  if (i18nProp_('I18N_AUTO') !== '1') return;
  var runs = i18nProp_(pk).split('|'); runs = runs[0] === day ? Number(runs[1] || 0) + 1 : 1; i18nProp_(pk, day + '|' + runs);
  if (runs > I18N_AUTO_MAX) { i18nLog_('排程', 'auto', '', lang, 'error', '今天已跑 ' + I18N_AUTO_MAX + ' 次，先停（保險絲）'); return; }
  while (Date.now() - t0 < 60000) {
    var pend = i18nPending_().filter(function (x) { return x.lang === lang && x.base !== 'zh'; }), x = null;   /* 中文原文只由 Claude Code 翻 */
    for (var i = 0; i < pend.length; i++) { if (i18nClaim_(pend[i].id + '|' + lang)) { x = pend[i]; break; } }
    if (!x) break;
    try { i18nWork_(x.id, lang, '排程', false, { auto: true, deadline: t0 + 330000 }); }
    catch (e) { i18nLog_('排程', 'work', x.id, lang, 'error', errMsg_(e)); }
    finally { i18nRelease_(x.id + '|' + lang); }
  }
  if (i18nPending_().some(function (x) { return x.lang === lang && x.base !== 'zh'; })) ScriptApp.newTrigger(h).timeBased().after(60000).create();
}
function i18nAuto_ja() { i18nAutoLang_('ja'); }
function i18nAuto_ko() { i18nAutoLang_('ko'); }
function i18nAuto_vi() { i18nAutoLang_('vi'); }
function i18nAuto_id() { i18nAutoLang_('id'); }
function i18nAuto_th() { i18nAutoLang_('th'); }
/* 開關（投遞通行碼）：on:1 打開自動翻並立刻開工；on:0 關掉。回目前排程 */
function i18nAutoSet_(p) {
  i18nWriteKey_(p);
  if (p.on != null) i18nProp_('I18N_AUTO', Number(p.on) ? '1' : '0');
  i18nEnsureTrigger_();
  var ref = p.refresh ? i18nRefresh_(true, Date.now() + 180000) : null;   /* 立刻全部重抓原文（同每天 06:30） */
  if (ref) i18nLog_('Claude Code', 'refresh', '', '', 'ok', '檢查 ' + ref.checked + '、重抓 ' + ref.fetched + '、內容有變 ' + ref.changed + '、下架 ' + ref.gone + (ref.errors.length ? '；' + ref.errors.join('；') : ''));
  var st = Number(p.on) && p.start ? i18nAutoStart_() : null;
  if (p.test && p.test.id && p.test.lang) {   /* 實測一組：指定的頁強制重翻＋回讀（花一點 API）；同自動排程的流程與放行規則 */
    var tid = String(p.test.id).toLowerCase(), tl = String(p.test.lang), tk = tid + '|' + tl;
    if (!i18nClaim_(tk)) throw fail_('這組有人在做', 'busy');
    try { st = i18nWork_(tid, tl, 'Claude Code 實測', false, { auto: true, deadline: Date.now() + 300000, redo: (p.test.pages || []).map(Number) }); } finally { i18nRelease_(tk); }
  }
  var trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).filter(function (f) { return /^i18n/.test(f); });
  return { data: { auto: i18nProp_('I18N_AUTO') === '1', refresh: ref, start: st, triggers: trig, pending: i18nPending_().length } };
}
ACTIONS.i18nAutoSet = function (p, a) { return i18nAutoSet_(p, a); };
NOAUTH.i18nAutoSet = 1; SELFLOCK.i18nAutoSet = 1;
function i18nStatus_() {
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows, rc = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows, cnt = {}, review = [], fail = [], wait = [];
  var nm = {}; src.forEach(function (s) { nm[s.id] = s.name; });
  i18nZLangs_().forEach(function (l) { cnt[l] = { live: 0, wait: 0, review: 0, fail: 0, old: 0, prev: 0 }; });   /* prev：新譯文還沒好、平板照舊顯示上一版（2026-11-05）；en 只有中文原文才有（2026-11-24） */
  var fp = {}; src.forEach(function (s) { fp[s.id] = s.fp; });
  rc.forEach(function (r) {
    if (!cnt[r.lang] || !fp[r.id]) return;
    var st = r.fp !== fp[r.id] ? 'old' : i18nState_(r, fp[r.id]); if (cnt[r.lang][st] != null) cnt[r.lang][st]++;
    if (st !== 'live' && i18nShowable_(r)) cnt[r.lang].prev++;
    if (st === 'wait') wait.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', at: r.at });
    if (st === 'review') review.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', check: r.check, at: r.at });
    if (st === 'fail') fail.push({ id: r.id, lang: r.lang, name: nm[r.id] || '', check: r.check, tries: r.tries, at: r.at });
  });
  var lg = [], sh = ss_().getSheetByName('i18n_log');
  if (sh && sh.getLastRow() > 1) { var n = Math.min(20, sh.getLastRow() - 1); lg = sh.getRange(sh.getLastRow() - n + 1, 1, n, I18N_LOG_HEAD.length).getDisplayValues().reverse(); }
  var trig = -1; try { trig = ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'i18nDaily'; }).length; } catch (e) { }
  var rcm = {}; rc.forEach(function (r) { rcm[r.key] = r; });
  var guide = {}; try { guide = i18nGuideNames_(); } catch (e) { }
  var pm = i18nPickMap_(), ZM = i18nZMap_();
  src = src.filter(function (s) { return !i18nZhOnly_(s.id, ZM); });   /* 中文試看的不列在翻譯總表（2026-11-17） */
  var okm = {}; src.forEach(function (s) { if (s.state === 'ok') okm[s.id] = 1; });
  var ret = i18nZRetired_(ZM, rcm, okm), retBy = {}; Object.keys(ret).forEach(function (k) { retBy[ret[k]] = k; });
  var matrix = src.filter(function (s) { return s.state === 'ok'; }).map(function (s) {
    var z = ZM[s.id], o = { id: s.id, name: s.name, cover: i18nCover_(s), slides: s.slides, upd: s.upd, local: z ? {} : guide[i18nNameKey_(s.name)] || {}, st: {}, appr: {}, on: i18nOn_(pm, s.id) };
    if (z) { o.base = 'zh'; o.old = z.old; o.replaced = !!retBy[s.id]; } else if (ret[s.id]) o.retired = ret[s.id];   /* replaced：已取代舊英文原文；retired：被中文版取代（不再列、不再翻） */
    (z ? i18nZLangs_() : I18N_LANGS).forEach(function (l) { var r = rcm[s.id + '|' + l]; o.st[l] = i18nState_(r, s.fp); if (r && o.st[l] === 'live') o.appr[l] = r.appr_by + ' ' + r.appr_at; });
    return o;
  });
  return { data: { langs: I18N_LANGS, zlangs: i18nZLangs_(), matrix: matrix, sources: src.filter(function (s) { return s.state === 'ok'; }).length, src_err: src.filter(function (s) { return s.state === 'err'; }).map(function (s) { return { id: s.id, name: s.name, err: s.err }; }),
    count: cnt, pending: i18nPending_().length, wait: wait, review: review, fail: fail, auto: i18nProp_('I18N_AUTO') === '1', refresh_at: i18nProp_('I18N_REFRESH_AT'), schedule: trig, log: lg, pick_ready: true, pick_init: !!pm } };   /* pick_ready：前端認得這版後台有勾選功能 */
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
function i18nList_(p) { i18nTabletAuth_(p.tpw); return i18nListData_(); }
/* 平板清單本體（不驗密碼）：i18nList 與推到 Render 的快照共用（2026-11-09），兩邊的 ver 一定一樣 */
function i18nListData_() {
  var src = i18nMeta_('i18n_src', I18N_SRC_HEAD), rc = {}, guide = i18nGuideNames_(), rm = i18nMeta_('i18n_rcp', I18N_RCP_HEAD);
  rm.rows.forEach(function (r) { rc[r.key] = r; });
  var pm = i18nPickMap_();   /* 沒勾的平板不列 */
  var zhm = i18nZh_(), ZM = i18nZMap_();
  var okRows = src.rows.filter(function (s) { return s.state === 'ok' && i18nOn_(pm, s.id); }), okm = {};
  var names = {}; src.rows.forEach(function (s) { names[s.id] = s.name; });
  okRows.forEach(function (s) { okm[s.id] = 1; });
  var ret = i18nZRetired_(ZM, rc, okm);   /* 中文版 6 種語言都好了 → 舊英文原文那支不列（2026-11-24） */
  var out = okRows.filter(function (s) { return !ret[s.id]; }).map(function (s) {
    var g = guide[i18nNameKey_(s.name)] || {}, loc = {};
    var o = { id: s.id, name: s.name, cover: i18nCover_(s), langs: {}, ver: {}, local: loc };
    if (ZM[s.id]) return i18nZListItem_(s, o, rc, rm.sh, ZM[s.id], okm, zhm, names);   /* 中文原文（i18n_zbase.gs） */
    if (zhm[s.id]) { o.langs.zh = 'live'; o.ver.zh = i18nHash_([s.fp, s.upd, 'zh'].join('|')).slice(0, 12); return o; }   /* 中文試看：只有「中文」這一格（2026-11-17） */
    I18N_LANGS.forEach(function (l) {
      var r = rc[s.id + '|' + l]; o.langs[l] = i18nState_(r, s.fp);
      if (o.langs[l] !== 'live' && i18nShowable_(r)) o.langs[l] = 'live';   /* 新譯文還沒好 → 平板照舊顯示上一版（2026-11-05） */
      if (o.langs[l] === 'live') o.ver[l] = i18nHash_([s.fp, s.upd, r.fp, r.appr_fp, r.appr_at].join('|')).slice(0, 12);   /* ver：平板快取用；原文任何更動（含只改份量）、新譯文放行就換 */
      if (g[l]) loc[l] = g[l];
      else if (o.langs[l] === 'live') { var t = i18nTrName_(rm.sh, r, o.ver[l]); if (t) loc[l] = t; }   /* 到店須知頁沒有這支（例：鳳梨酥）→ 用譯文自己的品名，卡片才不會只剩英文（2026-10-29） */
    });
    /* 英文（2026-11-11 經營者：「外語食譜，我要加英文食譜，放在第一個語言的位置」）：直接用後台台灣區英文版原文、不翻譯，所以收入就算放行 */
    o.langs.en = 'live'; o.ver.en = i18nHash_([s.fp, s.upd, 'en'].join('|')).slice(0, 12);
    return o;
  }).filter(Boolean);
  try { var cm = i18nListCats_(out); if (cm) out.forEach(function (o) { o.cats = cm[o.id] || []; }); } catch (e) { /* 類別讀不到 → 不帶 cats，平板照舊不分類（不會變空白） */ }
  return { data: { at: now_(), items: out } };
}
/* ---------------- 平板清單的「類別」（2026-11-24 經營者：「應該要用類別分類來顯示，不然一次出現 92 個食譜在一個語言中，太多，不容易找到要使用的食譜」） ----------------
   清單每支多帶 cats＝後台類別（去掉備份／English／訓練這些不是給客人看的分類）；平板頁（tablet-recipe.html 的 GROUPS）再歸成給外國客人看的大類。
   來源：🧬 食譜演化表 recipe_lineage 的 cat 欄（每天 1～2 點從後台 /Recipes 清單更新），不另外讀後台。
   怎麼找（兩種原文都能用）：①這支自己就有類別（原文改成中文食譜本身時）②英文食譜「英文名｜中文名」→ 同名的台灣中文食譜（啟用中優先）
   ③同一款（kuan，名稱去掉版本字樣後相同）的台灣中文啟用中食譜。都找不到 → []（平板放「其他」）。 */
var I18N_CAT_SKIP = /備份$|^English (Menu|Recipe)|^訓練&材料$/;
function i18nCatSplit_(s) {   /* 後台類別用頓號串起來，但類別名稱本身也有頓號（例：吳寶春大獎麵包 (限信義、高雄店製作)）→ 括號還沒關就接回去 */
  var out = [], cur = '';
  String(s || '').split('、').forEach(function (p) {
    cur = cur ? cur + '、' + p : p;
    if ((cur.match(/[(（]/g) || []).length <= (cur.match(/[)）]/g) || []).length) { cur = cur.trim(); if (cur) out.push(cur); cur = ''; }
  });
  if (cur.trim()) out.push(cur.trim());
  return out;
}
function i18nCatClean_(s) { return i18nCatSplit_(s).filter(function (c) { return !I18N_CAT_SKIP.test(c); }); }
function i18nCatNameKey_(n) { return String(n || '').replace(/（/g, '(').replace(/）/g, ')').replace(/\s+/g, '').toLowerCase(); }
/* rows：recipe_lineage 的列（{recipe_id, name, lang, region, status, cat, kuan}）；items：[{id, name}] → {id: [類別…]}（純計算，可在本機測） */
function i18nCatsCompute_(rows, items) {
  var byId = {}, byName = {}, byKuan = {};
  rows.forEach(function (r) {
    var id = String(r.recipe_id || '').toLowerCase(); if (!id) return;
    var o = { cats: i18nCatClean_(r.cat), on: r.status === 'Active', kuan: String(r.kuan || '') };
    byId[id] = o;
    if (r.region !== '台灣' || r.lang !== '繁體中文' || !o.cats.length) return;
    var nk = i18nCatNameKey_(r.name);
    (byName[nk] = byName[nk] || []).push(o);
    if (o.on && o.kuan) (byKuan[o.kuan] = byKuan[o.kuan] || []).push(o);
  });
  var union = function (list) {   /* 啟用中的優先；同名好幾支就合併它們的類別 */
    var act = list.filter(function (o) { return o.on; }), use = act.length ? act : list, seen = {}, out = [];
    use.forEach(function (o) { o.cats.forEach(function (c) { if (!seen[c]) { seen[c] = 1; out.push(c); } }); });
    return out;
  };
  var m = {};
  items.forEach(function (it) {
    var id = String(it.id || '').toLowerCase(), me = byId[id], segs = String(it.name || '').split('｜');
    var zhName = segs.length > 1 ? segs.slice(1).join('｜') : segs[0];
    var cats = me && me.cats.length ? me.cats.slice() : [];
    if (!cats.length && byName[i18nCatNameKey_(zhName)]) cats = union(byName[i18nCatNameKey_(zhName)]);
    if (!cats.length && me && me.kuan && byKuan[me.kuan]) cats = union(byKuan[me.kuan]);
    m[id] = cats;
  });
  return m;
}
/* 讀食譜演化表算類別；結果快取 6 小時（食譜演化表每天重算一次，LG_DATE 換了快取自動作廢）。表不在 → null */
function i18nListCats_(items) {
  var c = CacheService.getScriptCache(), key = 'i18n:cats:' + i18nHash_(i18nProp_('LG_DATE') + '|' + items.map(function (x) { return x.id + '~' + x.name; }).join('|'));
  var hit = c.get(key); if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var sh = ss_().getSheetByName(LG_TAB); if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues(), H = v[0];
  var rows = v.slice(1).map(function (r) { var o = {}; H.forEach(function (h, i) { o[h] = r[i]; }); return o; });
  var m = i18nCatsCompute_(rows, items), s = JSON.stringify(m);
  if (s.length < 90000) c.put(key, s, 21600);
  return m;
}
/* 譯文裡的品名（「在地名｜中文名」取前段）；快取 6 小時，核准或原文改了 ver 會換 */
function i18nTrName_(sh, r, ver) {
  var c = CacheService.getScriptCache(), k = 'i18n:tn:' + r.key + '|' + ver, hit = c.get(k);
  if (hit != null) return hit;
  var nm = '';
  try { var lc = i18nLiveContent_(sh, r); nm = rbTrim_(String((lc && lc.name) || '').split('｜')[0]); } catch (e) { }   /* 2026-11-06 修：原本內層也叫 c，把快取蓋掉 → 平板「c.put is not a function」 */
  c.put(k, nm, 21600);
  return nm;
}
function i18nGet_(p) { if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) i18nTabletAuth_(p.tpw); return { data: i18nGetData_(p.id, p.lang) }; }
/* 一支×語言組好的平板內容（不驗密碼）：i18nGet 與快照共用。ctx（i18nGetCtx_）＝一次推很多支時先讀好的表，不重複讀 */
function i18nGetCtx_() {
  var NM = {}; try { NM = i18nNameMap_(); } catch (e) { /* 對照表讀不到 → 只附英文（同 i18nBuild_） */ }
  return { sm: i18nMeta_('i18n_src', I18N_SRC_HEAD), rm: i18nMeta_('i18n_rcp', I18N_RCP_HEAD), pm: i18nPickMap_(), D: i18nDictRows_(), ZD: i18nZDictRows_(), ZM: i18nZMap_(), NM: NM, src: {} };   /* ZD／ZM：中文原文的詞典與開關（2026-11-24） */
}
function i18nGetData_(id, lang, ctx) {
  id = String(id || '').trim().toLowerCase(); lang = String(lang || '').trim();
  if (I18N_LANGS.indexOf(lang) < 0 && lang !== 'en' && lang !== 'zh') throw fail_('語言不對', 'bad');
  var isZ = !!(ctx ? ctx.ZM : i18nZMap_())[id], isT = !!i18nZh_()[id];
  if (lang === 'zh' ? !isT : (isT && !isZ)) throw fail_('這支食譜沒有這個語言', 'notyet');   /* 中文試看的只給中文；其他食譜沒有中文（2026-11-17）；同時是中文原文的兩種都有（2026-11-24） */
  var src;
  if (ctx) {
    if (!(id in ctx.src)) {
      var sr = ctx.sm.rows.filter(function (x) { return x.id === id; })[0], so = null;
      if (sr && sr.state === 'ok') { so = i18nReadJ_(ctx.sm.sh, sr._r, I18N_SRC_HEAD); if (so) so.fp = sr.fp; }
      ctx.src[id] = so;
    }
    src = ctx.src[id];
  } else src = i18nSrc_(id);
  if (!src) throw fail_('找不到這支食譜', 'nosrc');
  if (!i18nOn_(ctx ? ctx.pm : i18nPickMap_(), id)) throw fail_('這支食譜目前沒有開放給外國客人', 'notyet');
  if (lang === 'zh') {   /* 中文＝後台中文原文；品名、取料區已是中文，不再附原文對照 */
    var zb = i18nBuild_(src, { name: src.name, pres: src.pres, slides: src.slides.map(function (s) { return { t: s.t, h: s.h }; }) }, {}, 'zh', { status: 'ok', check: '', at: '', model: '原文' }, {});
    zb.slides.forEach(function (s) { s.ings.forEach(function (g) { g.en = ''; g.zh = ''; }); });
    return zb;
  }
  var Z = src.base === 'zh', D = Z ? (ctx ? ctx.ZD : i18nZDictRows_()) : (ctx ? ctx.D : i18nDictRows_());
  if (lang === 'en' && !Z) return i18nBuild_(src, { name: src.name, pres: src.pres, slides: src.slides.map(function (s) { return { t: s.t, h: s.h }; }) }, D, 'en', { status: 'ok', check: '', at: '', model: '原文' }, ctx ? ctx.NM : null);   /* 英文＝後台原文（2026-11-11）；中文原文的英文是翻的（2026-11-24） */
  var m = ctx ? ctx.rm : i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  var c = i18nLiveContent_(m.sh, r);   /* 放行過的那份（原文改版、新譯文還沒好 → 上一版） */
  if (!c) throw fail_('這支食譜這個語言還沒翻好', 'notyet');
  return i18nBuild_(src, { name: c.name, pres: c.pres, slides: i18nAlign_(c, src) }, D, lang, r, ctx ? ctx.NM : null);
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
  var D = i18nDictFor_(src), b = i18nBuild_(src, i18nReadJ_(m.sh, r._r, I18N_RCP_HEAD), D, lang, r);
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
/* 2026-11-27 經營者：「要把圖片補上去」→ 後台沒設封面的食譜，平板借用同款的封面（不動後台）；鍵＝食譜 id */
var I18N_COVER = {
  '80b02265-e6c2-49fe-b6e0-ee6386ce1137': 'https://diybcstorage.blob.core.windows.net/recipeimages/c5951d5d-b7f5-4b1b-b2bc-dc304d5de0fd/c5951d5d-b7f5-4b1b-b2bc-dc304d5de0fd.jpg?v=261117024659'   /* 草莓生日蛋糕(限當日壽星)_自選版 ← 海綿蛋糕版 */
};
function i18nCover_(s) { return (s && s.cover) || I18N_COVER[String(s && s.id || '').toLowerCase()] || ''; }
/* 2026-11-27 經營者：「在日文步驟內容的下方加入中文步驟內容，讓我們的夥伴可以在協助客人的時候看得懂」→ 先試做老奶奶檸檬蛋糕日文；鍵＝食譜id|語言 */
var I18N_ZHSUB = { 'a0662719-fe28-4c37-a1b3-db63c7477a48|ja': 1 };
function i18nBuild_(src, tr, D, lang, row, NM0) {
  var areas = {}, NM = NM0 || {}, Z = src.base === 'zh';   /* Z：中文原文（2026-11-24）——詞典 D 是中文索引；原文對照附中文名＋英文譯名 */
  if (!NM0 && !Z) { try { NM = i18nNameMap_(); } catch (e) { /* 對照表讀不到 → 只附英文 */ } }
  var cut = function (n) { return rbTrim_(String(n || '').replace(/^【[^】]*】\s*/, '')); };
  var slides = src.slides.map(function (s, i) {
    var t = (tr.slides || [])[i] || {}, h = String(t.h || ''), im = lang === 'zh' && s.imgs_zh ? s.imgs_zh : (s.imgs || []);   /* 中文試看看中文原圖 */
    im.forEach(function (u, j) { if (Z) u = i18nZPoster_(u, lang); h = h.split('[[IMG' + (j + 1) + ']]').join('<img src="' + String(u).replace(/"/g, '&quot;') + '">'); });   /* 中文原文：台灣限定頁海報換各語言版 */
    h = h.replace(/\[\[IMG\d+\]\]/g, '');   /* 上一版譯文的圖片位置比現在原文多（保底顯示時）→ 拿掉，不顯示怪字 */
    var ings = s.ings.map(function (g) {
      if (g.a && !areas[g.a]) areas[g.a] = { name: i18nLook_(D, 'area', g.an, lang) || g.an, en: Z ? (lang === 'en' ? '' : i18nLook_(D, 'area', g.an, 'en')) : g.an, img: src.areaPre && src.areaHas.indexOf(g.a) >= 0 ? '/img/MenuCateGroups/' + src.areaPre + '-' + g.a + '.jpg' : '' };
      if (Z) return { n: i18nItemName_(D, g.n, lang), en: lang === 'en' ? '' : i18nLook_(D, 'item', i18nSplitName_(g.n).core, 'en'), zh: cut(g.n), q: g.q, u: i18nLook_(D, 'unit', g.u, lang) || g.u, c: g.c, a: g.a };
      var nmk = NM[cut(g.n).toLowerCase()];
      return { n: i18nItemName_(D, g.n, lang), en: g.n, zh: nmk && nmk.zh || '', q: g.q, u: i18nLook_(D, 'unit', g.u, lang) || g.u, c: g.c, a: g.a };
    });
    var o = { t: t.t || s.t, video: s.video, img: lang === 'zh' && s.img_zh ? s.img_zh : s.img, h: h, ings: ings };
    if (s.vtut) { o.video_orig = s.video; o.video = i18nZTut_(lang, s.video); }   /* 中文教學影片 → 熊主廚各語言版／英文共用版 */
    if (Z && lang !== 'zh' && I18N_ZHSUB[src.id + '|' + lang] && !s.ings.length) {   /* 步驟頁下方附中文（給夥伴看）；圖片不重複放 */
      o.zt = s.t; o.zh = String(s.h || '').replace(/\[\[IMG\d+\]\]/g, '').replace(/<img\b[^>]*>/gi, '');
    }
    return o;
  });
  if (!Z) i18nVSwap_(slides, lang);
  var b = { id: src.id, lang: lang, name: tr.name || src.name, en: src.name, cover: i18nCover_(src), pres: tr.pres || src.pres, areas: areas, slides: slides,
    status: row.status, check: row.check, at: row.at, model: row.model };
  if (src.pick) b.pick = src.pick.map(function (g) {   /* 自選版（2026-11-26）：平板先選；ks＝這個選項自己的頁（沒選到的不顯示） */
    var tl = function (x) { return lang === 'zh' ? x : (i18nLook_(D, 'pick', x, lang) || x); };
    return { label: tl(g.label), zh: g.label, opts: g.opts.map(function (o) { return { label: tl(o.label), zh: o.label, ks: o.ks }; }) };
  });
  return b;
}

/* 共用影片換成各語言版（2026-11-03）：外語食譜第一頁的「熊主廚」開場動畫只有英文 → 認檔案指紋（Azure Content-MD5，每支食譜網址不同但檔案同一支）換成該語言版。
   成品與製作工具：DIY repo media/tablet-intro/、tools/tablet_intro_video/；網址用 jsDelivr 固定 commit（支援影片分段讀取）。 */
var I18N_VSWAP = {
  'zxY7HTTs9Oub93tFKzn3aw==': { base: 'https://cdn.jsdelivr.net/gh/mydiybc-bot/DIY@f9e57f9cf05bd011524e2b4f21c6b40f284cfab4/media/tablet-intro/chef-bear-', langs: ['ja', 'ko', 'vi', 'id', 'th'] }
};
function i18nVSwap_(slides, lang) {
  var c = CacheService.getScriptCache(), need = [], md5 = {};
  slides.forEach(function (s) {
    if (!s.video) return;
    var k = 'i18n:vmd5:' + i18nHash_(s.video), hit = c.get(k);
    if (hit != null) md5[s.video] = hit; else if (need.indexOf(s.video) < 0) need.push(s.video);
  });
  if (need.length) {
    var rs = [];
    try { rs = UrlFetchApp.fetchAll(need.map(function (u) { return { url: u, headers: { Range: 'bytes=0-0' }, muteHttpExceptions: true }; })); } catch (e) { rs = []; }
    need.forEach(function (u, i) {
      var r = rs[i]; if (!r) return;
      var code = r.getResponseCode(), h = r.getHeaders(), m = '';
      Object.keys(h).forEach(function (k) { var lk = k.toLowerCase(); if (lk === 'x-ms-blob-content-md5' || lk === 'content-md5') m = m || String(h[k]); });
      md5[u] = m;
      if (code === 200 || code === 206) c.put('i18n:vmd5:' + i18nHash_(u), m, 21600);   /* 6 小時；讀不到就下次再問 */
    });
  }
  slides.forEach(function (s) {
    var v = s.video && I18N_VSWAP[md5[s.video]];
    if (v && v.langs.indexOf(lang) >= 0) { s.video_orig = s.video; s.video = v.base + lang + '.mp4'; }
  });
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
/* 2026-11-26 唯讀：放行中的譯文原樣（[[IMGn]] 佔位不展開），給 Claude Code 批次改詞（只讀、要唯讀預覽通行碼） */
function i18nRaw_(p) {
  if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) throw fail_('要通行碼', 'auth');
  var id = String(p.id || '').trim().toLowerCase(), lang = String(p.lang || '').trim();
  var m = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r = m.rows.filter(function (x) { return x.key === id + '|' + lang; })[0];
  var c = r ? i18nLiveContent_(m.sh, r) : null; if (!c) throw fail_('這支這個語言沒有放行中的譯文', 'notyet');
  return { data: { id: id, lang: lang, fp: c.fp, name: c.name, pres: c.pres, slides: c.slides, cur_fp: r.fp } };
}
ACTIONS.i18nRaw = function (p) { return i18nRaw_(p); };
READS.i18nRaw = 1; NOAUTH.i18nRaw = 1;   /* 身分在函式裡驗（唯讀預覽通行碼） */
ACTIONS.i18nApprove = function (p, a) { return i18nApprove_(p, a); };
ACTIONS.i18nPick = function (p, a) { return i18nPick_(p, a); };   /* 2026-10-26 勾選；跟 i18nApprove 一樣走一般寫入路徑（鎖、回條、身分密碼） */
READS.i18nPreview = 1; NOAUTH.i18nPreview = 1;   /* 身分在函式裡驗（主廚／管理者密碼，或唯讀預覽通行碼） */   /* i18nApprove 走一般寫入路徑（有鎖、有回條、寫紀錄） */
NOAUTH.i18nStatus = 1; NOAUTH.i18nList = 1; NOAUTH.i18nGet = 1; NOAUTH.i18nKick = 1;   /* 平板密碼在函式裡自己驗；Kick 只會做待辦，做完就沒事可做 */
READS.i18nStatus = 1; READS.i18nList = 1; READS.i18nGet = 1;
SELFLOCK.i18nKick = 1;   /* 一次要 1～4 分鐘，不能佔住全系統的鎖；同一件不重複做靠 i18nClaim_ */

/* ================= 🚀 平板快照 → 訓練站伺服器 Render（2026-11-09 經營者：「未來如果我希望可以讓 100 人同時使用」→ 方案 A） =================
   原本平板每開一支食譜，GAS 都要當場讀表、組內容（約 5 秒），Google 帳號同時最多跑 30 個 → 約 25～30 人同時點開就會擋。
   改成：GAS 把「平板清單＋每支×語言組好的內容」推到訓練站伺服器存起來（DIY i18n_snap.py，存在 Render 硬碟），
   平板先向伺服器讀（/api/i18n/list、/api/i18n/get，同一個網域），讀不到才退回這裡的 i18nList／i18nGet（照舊保留）。
   推送：i18nSnapPush_ 先問伺服器現有哪些版本（manifest），只推 ver 變了的（ver 同 i18nList）→ 最後推清單（清單上沒有的伺服器會刪）。
   排程：i18nSnapTick 每 15 分鐘；表的內容沒變（i18nSnapSig_）就直接結束（約 1 秒）；每天 05:00 後第一次全部重推（詞典、到店須知品名有改也會跟上）。
   開關／立刻推：動作 i18nSnapSet（投遞通行碼 wkey）。伺服器通行碼在指令碼屬性 I18N_SNAP_KEY（伺服器端只放 SHA-256）。 */
var I18N_SNAP_URL = 'https://diybc-training.onrender.com/api/i18n/snap';
function i18nSnapPost_(body) {
  var key = i18nProp_('I18N_SNAP_KEY'); if (!key) throw fail_('還沒設定伺服器通行碼 I18N_SNAP_KEY', 'off');
  body.key = key;
  var r = UrlFetchApp.fetch(I18N_SNAP_URL, { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true, followRedirects: true });
  var j; try { j = JSON.parse(r.getContentText()); } catch (e) { throw fail_('伺服器回應看不懂（HTTP ' + r.getResponseCode() + '）', 'net'); }
  if (!j || !j.ok) throw fail_('伺服器：' + ((j && (j.msg || j.error)) || ('HTTP ' + r.getResponseCode())), 'net');
  return j;
}
/* force＝不問伺服器、全部重推；deadline＝毫秒時間點，到了就先停（清單照推；沒推到的那幾支平板會退回 GAS，下次再補） */
/* done（2026-11-24）：全部重推時「今天已推過」的 {id|lang|ver: 1}，跳過並把這次推的加進去——一次推不完，下一次接著推，不從頭來
   （中文原文 93 支×6 語言，一次 4.5 分鐘推不完；原本每 15 分鐘從頭重推，會把 Google 每天的執行時間用光） */
function i18nSnapPush_(force, deadline, done) {
  var t0 = Date.now(); deadline = deadline || t0 + 270000; done = done || {};
  var L = i18nListData_().data, have = force ? {} : (i18nSnapPost_({ op: 'manifest' }).vers || {}), todo = [];
  L.items.forEach(function (it) { I18N_LANGS.concat(['en']).forEach(function (l) { var k = it.id + '|' + l; if (it.langs[l] === 'live' && it.ver[l] && have[k] !== it.ver[l] && !done[k + '|' + it.ver[l]]) todo.push({ id: it.id, lang: l, ver: it.ver[l] }); }); });
  var ctx = todo.length ? i18nGetCtx_() : null, batch = [], size = 0, pushed = 0, failed = [], cut = false;
  var flush = function () { if (!batch.length) return; i18nSnapPost_({ op: 'put', items: batch }); batch.forEach(function (b) { done[b.id + '|' + b.lang + '|' + b.ver] = 1; }); pushed += batch.length; batch = []; size = 0; };
  for (var i = 0; i < todo.length; i++) {
    if (Date.now() > deadline) { cut = true; break; }
    var x = todo[i], d;
    try { d = i18nGetData_(x.id, x.lang, ctx); } catch (e) { failed.push(x.id + '|' + x.lang + '：' + errMsg_(e)); continue; }
    batch.push({ id: x.id, lang: x.lang, ver: x.ver, data: d }); size += JSON.stringify(d).length;
    if (size > 800000 || batch.length >= 15) flush();
  }
  flush();
  var lr = i18nSnapPost_({ op: 'list', at: L.at, items: L.items });
  return { items: L.items.length, todo: todo.length, pushed: pushed, failed: failed, left: todo.length - pushed - failed.length, cut: cut, removed: lr.removed || 0, sec: Math.round((Date.now() - t0) / 1000) };
}
/* 表的內容指紋：原文、譯文狀態、勾選有任何變動才推 */
function i18nSnapSig_() {
  var a = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.map(function (r) { return [r.id, r.name, r.upd, r.fp, r.cover, r.state].join('~'); });
  var b = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.map(function (r) { return [r.key, r.fp, r.status, r.appr_by, r.appr_at, r.appr_fp].join('~'); });
  return i18nHash_(JSON.stringify([a, b, i18nPickMap_(), i18nZMap_()]));   /* 中文原文的開關、對應也算（2026-11-24） */
}
function i18nSnapTick() {
  if (!i18nProp_('I18N_SNAP_KEY')) return;
  var now = new Date(), today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd'), full = i18nProp_('I18N_SNAP_FULL') !== today && Number(Utilities.formatDate(now, TZ, 'H')) >= 5;
  var sig = i18nSnapSig_();
  if (!full && sig === i18nProp_('I18N_SNAP_SIG') && i18nProp_('I18N_SNAP_LEFT') !== '1') return;
  var r, ck = 'i18n:snapfull:' + today, cache = CacheService.getScriptCache(), done = {};
  if (full) { try { done = JSON.parse(cache.get(ck) || '{}'); } catch (e) { done = {}; } }   /* 今天全部重推已推過的（2026-11-24 接著推） */
  try { r = i18nSnapPush_(full, Date.now() + 270000, done); }
  catch (e) { if (full) try { cache.put(ck, JSON.stringify(done), 21600); } catch (e2) { } i18nProp_('I18N_SNAP_LEFT', '1'); i18nLog_('排程', 'snap', '', '', 'error', errMsg_(e)); return; }
  if (full) try { cache.put(ck, JSON.stringify(done), 21600); } catch (e) { }
  var left = r.left > 0 || r.failed.length > 0;
  i18nProp_('I18N_SNAP_SIG', sig); i18nProp_('I18N_SNAP_LEFT', left ? '1' : '0');
  if (full && !r.left) i18nProp_('I18N_SNAP_FULL', today);
  if (r.pushed || r.failed.length || r.removed || full) i18nLog_('排程', 'snap', '', '', r.failed.length ? 'error' : 'ok', (full ? '全部重推；' : '') + '清單 ' + r.items + ' 支、推 ' + r.pushed + '/' + r.todo + '、刪 ' + r.removed + (r.left ? '、剩 ' + r.left : '') + (r.failed.length ? '；失敗 ' + r.failed.slice(0, 3).join('；') : ''), r.sec);
}
/* 開關：key＝設定伺服器通行碼；on:1 建排程（每 15 分鐘）／on:0 拿掉；push:1 立刻推（force:1 全部重推）；sig 會更新 */
function i18nSnapSet_(p) {
  i18nWriteKey_(p);
  if (p.key) i18nProp_('I18N_SNAP_KEY', String(p.key));
  if (p.on != null) { i18nDropTrig_('i18nSnapTick'); if (Number(p.on)) ScriptApp.newTrigger('i18nSnapTick').timeBased().everyMinutes(15).create(); }
  var r = null;
  if (p.push) { var sig = i18nSnapSig_(); r = i18nSnapPush_(!!p.force, Date.now() + 300000); i18nProp_('I18N_SNAP_SIG', sig); i18nProp_('I18N_SNAP_LEFT', r.left || r.failed.length ? '1' : '0'); i18nLog_('Claude Code', 'snap', '', '', r.failed.length ? 'error' : 'ok', (p.force ? '全部重推；' : '') + '清單 ' + r.items + ' 支、推 ' + r.pushed + '/' + r.todo, r.sec); }
  var chk = null;
  if (p.check) {   /* 驗證：一次推（ctx）組出來的內容＝平板 i18nGet 單支組出來的，一字不差 */
    var ctx = i18nGetCtx_(), L = i18nListData_().data, n = 0, same = 0, diff = [];
    L.items.forEach(function (it) { I18N_LANGS.concat(['en']).forEach(function (l) { if (n >= Number(p.check) || it.langs[l] !== 'live') return; n++;
      var x = JSON.stringify(i18nGetData_(it.id, l, ctx)), y = JSON.stringify(i18nGetData_(it.id, l)); if (x === y) same++; else diff.push(it.id + '|' + l); }); });
    chk = { n: n, same: same, diff: diff };
  }
  var trig = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).filter(function (f) { return /^i18n/.test(f); });
  return { data: { has_key: !!i18nProp_('I18N_SNAP_KEY'), push: r, check: chk, triggers: trig } };
}
ACTIONS.i18nSnapSet = function (p, a) { return i18nSnapSet_(p, a); };
NOAUTH.i18nSnapSet = 1; SELFLOCK.i18nSnapSet = 1;

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
    /* fresh：今天 06:00 後還沒更新過原文 → 先更新（Claude 早上巡檢不用等 06:30 的排程，排程晚跑也不會漏） */
    var ref = null, last = i18nProp_('I18N_REFRESH_AT'), today6 = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + ' 06:00';
    if (p.fresh && (!last || last < today6 || p.fresh === 'force')) ref = i18nRefresh_(false, Date.now() + 140000);
    var pend = i18nPending_(), rc = {}; i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.forEach(function (r) { rc[r.key] = r; });
    return { data: { pending: pend.length, refresh: ref, refresh_at: i18nProp_('I18N_REFRESH_AT'), items: pend.slice(0, 700).map(function (x) { var r = rc[x.id + '|' + x.lang]; return { id: x.id, lang: x.lang, name: x.name, base: x.base, why: !r ? '還沒翻' : r.status === 'fail' ? '失敗' : '原文改版', showing: i18nShowable_(r) }; }), at: now_() } };
  }
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜（不是收進外語食譜的原文？）', 'nosrc');
  var Z = src.base === 'zh', LL = Z ? i18nZLangs_() : I18N_LANGS;   /* 中文原文（2026-11-24）：多一個 en；詞典用中文索引 */
  if (LL.indexOf(lang) < 0) throw fail_('語言只能是 ' + LL.join('／'), 'bad');
  var D = Z ? i18nZDictRows_() : i18nDictRows_(), terms = i18nTerms_(src), have = [], miss = [], nm = {};
  if (Z) {
    var ZN = {}; try { ZN = i18nZNameMap_(); } catch (e) { }
    var EZ = i18nZEnZh_(ZN), OD = i18nDictRows_();
    terms.forEach(function (x) {
      var tr = i18nLook_(D, x.kind, x.en, lang);
      if (tr) { have.push({ kind: x.kind, zh: x.en, tr: tr }); return; }
      var o = { kind: x.kind, zh: x.en }, h = x.kind === 'item' ? i18nZHint_(x.en, lang, ZN, EZ, OD) : {};
      if (h.backend_en) o.backend_en = h.backend_en;
      if (lang === 'vi' && h.backend_vi) o.backend = h.backend_vi;   /* 越南文：後台已有的名稱優先 */
      if (h.cand) o.cand = h.cand;   /* 2026-11-24：品名照後台英文名（cand＝後台英文名或它在舊詞典的譯名） */
      if (h.note) o.note = h.note;
      if (x.kind === 'pick') o.note = '平板選擇頁的文字（自選版：客人做之前先選蛋糕體口味、夾餡；組名或選項名），照中文逐字翻、要短';
      miss.push(o);
    });
  } else {
    try { nm = i18nNameMap_(); } catch (e) { }
    terms.forEach(function (x) {
      var tr = i18nLook_(D, x.kind, x.en, lang);
      if (tr) { have.push({ kind: x.kind, en: x.en, tr: tr }); return; }
      var h = nm[x.en.toLowerCase()] || {}, o = { kind: x.kind, en: x.en, zh: h.zh || '' };
      if (lang === 'vi' && x.kind === 'item' && h.vi) o.backend = h.vi;   /* 越南文：後台已有的名稱優先 */
      miss.push(o);
    });
  }
  var guide = i18nGuideNames_()[i18nNameKey_(src.name)] || {}, local = Z ? i18nZLocal_(src, lang, i18nZOldName_(src)) : guide[lang] || '';
  var rm = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), cur = rm.rows.filter(function (r) { return r.key === id + '|' + lang; })[0];
  /* 只翻改到的頁（2026-11-05）：放行過的上一版裡字沒變的頁直接沿用（keep），只回要翻的頁；同頁碼的舊譯文附在 prev 給譯者參考 */
  var base = cur ? i18nLiveContent_(rm.sh, cur) : null, reuse = base ? i18nAlign_(base, src, true) : [], keep = [], prev = {}, tt = {};
  src.slides.forEach(function (s, i) {
    if (reuse[i] || i18nFixedSlide_(s, lang, Z)) { keep.push(i); return; }   /* 固定譯文的頁不給譯者（投遞時伺服器自己填，2026-11-16） */
    if (base && base.slides && base.slides.length === src.slides.length && base.slides[i]) prev[i] = base.slides[i];
  });
  if (base) src.slides.forEach(function (s) { if (s.t && base.titles && base.titles[s.t]) tt[s.t] = base.titles[s.t]; });   /* 沿用頁的標題譯法，新頁同標題照用 */
  var nameKeep = !!(base && base.name && base.snk === i18nHash_(String(src.name || ''))), presKeep = !!(base && base.spk === i18nHash_(String(src.pres || '')));
  return { data: { id: id, lang: lang, base: Z ? 'zh' : 'en', label: I18N_LABEL[lang], rules: I18N_RULES[lang], fp: src.fp, name: src.name, pres: src.pres, local: local,
    slides: src.slides.map(function (s, i) { return { k: i, t: s.t, h: s.h }; }), keep: keep, prev: prev,
    name_keep: nameKeep, pres_keep: presKeep, old_name: base ? base.name : '', old_pres: base ? base.pres : '',
    dict: have, missing: miss, brief: Z ? i18nZSystem_(lang, src, D, local, tt) : i18nSystem_(lang, src, D, local, tt),
    now: cur ? { status: cur.status, state: i18nState_(cur, src.fp), at: cur.at, model: cur.model } : null } };
}
/* 把 Claude 翻好的詞放進詞典：只補空格、不蓋掉已有的（除非 overwrite）；只收這支食譜用到的詞 */
function i18nPutDict_(list, lang, src, overwrite) {
  var Z = src.base === 'zh', TAB = Z ? 'i18n_zitems' : 'i18n_items', HEAD = Z ? I18N_ZITEM_HEAD : I18N_ITEM_HEAD, KF = Z ? 'zh' : 'en';   /* 中文原文 → 中文索引的詞典（2026-11-24） */
  var allowed = {}; i18nTerms_(src).forEach(function (x) { allowed[x.kind + '|' + x.en.toLowerCase()] = x; });
  var got = {}, ignored = 0;
  (list || []).forEach(function (x) {
    var k = String(x.kind) + '|' + rbTrim_(x[KF] || '').toLowerCase(), tr = rbTrim_(x.tr);
    if (allowed[k] && tr) got[k] = { x: allowed[k], tr: tr }; else ignored++;
  });
  var added = 0, kept = 0;
  i18nLock_(function () {
    var m = i18nMeta_(TAB, HEAD), idx = {};
    m.rows.forEach(function (r) { idx[r.kind + '|' + String(r[KF]).toLowerCase()] = r; });
    Object.keys(got).forEach(function (k) {
      var g = got[k], r = idx[k], vals;
      if (r) {
        if (String(r[lang] || '').trim() && !overwrite) { kept++; return; }
        r[lang] = g.tr; r.how = (r.how ? r.how + ' ' : '') + lang + ':claude'; r.at = now_();
        vals = HEAD.map(function (h) { return r[h] == null ? '' : r[h]; }); m.sh.getRange(r._r, 1, 1, vals.length).setNumberFormat('@').setValues([vals]);
      } else {
        var o = { kind: g.x.kind, how: lang + ':claude', at: now_() }; o[KF] = g.x.en; o[lang] = g.tr;
        vals = HEAD.map(function (h) { return o[h] || ''; }); m.sh.appendRow(vals); idx[k] = o;
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
  var src = i18nSrc_(id); if (!src) throw fail_('找不到這支食譜', 'nosrc');
  var Z = src.base === 'zh', LL = Z ? i18nZLangs_() : I18N_LANGS;   /* 中文原文（2026-11-24） */
  if (LL.indexOf(lang) < 0) throw fail_('語言只能是 ' + LL.join('／'), 'bad');
  if (p.fp && String(p.fp) !== String(src.fp)) throw fail_('原文在你拿稿之後改版了（指紋不同），請重新拿稿再翻', 'stale');
  var sl = p.slides || [], partial = !!Number(p.partial) || p.partial === true;
  if (!partial && sl.length !== src.slides.length) throw fail_('頁數不對：原文 ' + src.slides.length + ' 頁，收到 ' + sl.length + ' 頁', 'bad');
  var out = { name: String(p.name || '').trim(), pres: String(p.pres || '').trim(), slides: [] };
  sl.forEach(function (s) { var k = Number(s.k); if (!(k >= 0 && k < src.slides.length) || out.slides[k]) throw fail_('頁碼 k 不對或重複：' + s.k, 'bad'); out.slides[k] = { t: String(s.t || '').trim(), h: String(s.h || '') }; });
  src.slides.forEach(function (s, i) { var f = i18nFixedSlide_(s, lang, Z); if (f) out.slides[i] = f; });   /* 固定譯文的頁一律用固定版（譯者送什麼都蓋掉，2026-11-16） */
  var reused = 0;
  if (partial) {   /* 只翻改到的頁（2026-11-05）：沒送的頁＝字沒變，從放行過的上一版沿用；名稱、保存方式沒變也沿用 */
    var rm0 = i18nMeta_('i18n_rcp', I18N_RCP_HEAD), r0 = rm0.rows.filter(function (r) { return r.key === id + '|' + lang; })[0];
    var base = r0 ? i18nLiveContent_(rm0.sh, r0) : null, reuse = base ? i18nAlign_(base, src, true) : [], lack = [];
    src.slides.forEach(function (s, i) { if (out.slides[i]) return; if (reuse[i]) { out.slides[i] = { t: String(reuse[i].t || ''), h: String(reuse[i].h || '') }; reused++; } else lack.push(i + 1); });
    if (lack.length) throw fail_('第 ' + lack.join('、') + ' 頁沒翻、上一版也沒有同樣的頁，請重新拿稿', 'bad');
    if (!out.name && base && base.snk === i18nHash_(String(src.name || ''))) out.name = String(base.name || '');
    if (!out.pres && base && base.spk === i18nHash_(String(src.pres || ''))) out.pres = String(base.pres || '');
  }
  var dict = i18nPutDict_(p.dict, lang, src, !!p.overwrite);
  var local = Z ? i18nZLocal_(src, lang, i18nZOldName_(src)) : (i18nGuideNames_()[i18nNameKey_(src.name)] || {})[lang] || '';
  if (Z) { var nm0 = rbTrim_((local || out.name).split('｜')[0]); out.name = nm0 ? nm0 + '｜' + src.name : ''; }   /* 中文原文：「譯名｜中文名」，中文名由伺服器補 */
  else if (local) out.name = local + '｜' + src.name.split('｜').slice(1).join('｜');
  var tt = {}; src.slides.forEach(function (s, i) { if (s.t && out.slides[i] && out.slides[i].t) tt[s.t] = out.slides[i].t; });
  var D2 = Z ? i18nZDictRows_() : i18nDictRows_(), okx = {};   /* 這支食譜用到的詞典譯名裡的數字（尺寸、號數）不算「多出來的數字」 */
  i18nTerms_(src).forEach(function (x) { var tr = i18nLook_(D2, x.kind, x.en, lang); if (tr) i18nNums_(tr).forEach(function (n) { okx[n] = 1; }); });
  if (local) i18nNums_(local).forEach(function (n) { okx[n] = 1; });   /* 官網在地品名自帶的數字（例：오레오 2단）也放過 */
  var P = {}; src.slides.forEach(function (s, i) { var b = i18nCheckSlide_(s, out.slides[i], okx, Z ? lang : ''); if (b.length) P[i] = b; });
  var bad = Object.keys(P), check = bad.map(function (i) { return '第' + (Number(i) + 1) + '頁：' + P[i].join('、'); }).join('；');
  if (src.pres && !out.pres) check += (check ? '；' : '') + '保存方式是空的';
  if (Z && src.pres && out.pres && (out.pres === src.pres || i18nZHan_(out.pres, lang))) check += (check ? '；' : '') + '保存方式沒翻（還是中文）';
  if (Z && out.name && i18nZHan_(out.name.split('｜')[0], lang)) check += (check ? '；' : '') + '品名譯名還有中文字';
  var pmiss = i18nTerms_(src).filter(function (x) { if (x.kind !== 'pick') return false; var tr = i18nLook_(D2, 'pick', x.en, lang); return !tr || (Z && i18nZHan_(tr, lang)); }).map(function (x) { return x.en; });
  if (pmiss.length) check += (check ? '；' : '') + '選擇頁沒翻：' + pmiss.join('、');   /* 自選版（2026-11-26） */
  var status = !out.name ? 'fail' : (check ? 'review' : 'ok');
  if (!out.name) check = '名稱是空的' + (check ? '；' + check : '');
  var key = id + '|' + lang, prev = i18nMeta_('i18n_rcp', I18N_RCP_HEAD).rows.filter(function (r) { return r.key === key; })[0];
  var tries = prev && prev.fp === src.fp ? Number(prev.tries || 0) + 1 : 1, model = String(p.model || 'claude-code').slice(0, 60);
  var approved = status === 'ok' && !!p.approve, sec = Math.round((Date.now() - t0) / 1000);
  var json = i18nJson_(src, { id: id, lang: lang, model: model, name: out.name, pres: out.pres, slides: out.slides, titles: tt });
  var kept = false;
  i18nLock_(function () {
    var meta = { key: key, id: id, lang: lang, fp: src.fp, status: status, tries: tries, model: model, check: check, sec: sec, at: now_() };
    i18nSaveRcp_(meta, json, approved, 'Claude'); kept = /^L:/.test(meta.appr_fp);
  });
  i18nLog_(by, 'put', id, lang, status, (reused ? '沿用 ' + reused + ' 頁；' : '') + (dict.added ? '詞典新增 ' + dict.added + '；' : '') + (approved ? '自動放行；' : kept ? '平板照舊顯示上一版；' : '') + (check || '通過'), sec);
  return { id: key, msg: 'put ' + key + ' ' + status, data: { id: id, lang: lang, name: out.name, status: status, check: check, approved: approved, state: approved ? 'live' : (status === 'ok' ? 'wait' : status), tablet: approved ? 'new' : kept ? 'old' : 'none', slides: src.slides.length, reused: reused, dict: dict, left: i18nPending_().length } };
}
/* 一次性搬家（2026-11-05 指紋改成只看要翻的字）：原文與現有譯文一起換新指紋，譯文補每頁的鍵（sk）。
   只動「譯文指紋＝原文舊指紋」的列（跟原文同一版的）；放行過的放行欄一起換。每次最多跑 4 分鐘，回 left>0 就再呼叫，重跑安全。 */
function i18nMigrateFp_(p) {
  i18nWriteKey_(p);
  var t0 = Date.now(), res = { recipes: 0, rows: 0, left: 0, same: 0 }, sm = i18nMeta_('i18n_src', I18N_SRC_HEAD);
  sm.rows.forEach(function (s) {
    if (s.state !== 'ok') return;
    if (Date.now() - t0 > 240000) { res.left++; return; }
    var src = i18nReadJ_(sm.sh, s._r, I18N_SRC_HEAD); if (!src) return;
    var old = s.fp, nfp = i18nFp_(src); if (old === nfp) { res.same++; return; }
    src.fp = nfp;
    var sk = src.slides.map(i18nSlideKey_), snk = i18nHash_(String(src.name || '')), spk = i18nHash_(String(src.pres || ''));
    i18nLock_(function () {
      var rm = i18nMeta_('i18n_rcp', I18N_RCP_HEAD);
      rm.rows.filter(function (r) { return r.id === s.id && r.fp === old; }).forEach(function (r) {
        var j = i18nReadJ_(rm.sh, r._r, I18N_RCP_HEAD); if (!j) return;
        j.fp = nfp; j.sk = sk; j.snk = snk; j.spk = spk;
        i18nUpsert_('i18n_rcp', I18N_RCP_HEAD, 'key', { key: r.key, fp: nfp, appr_fp: r.appr_fp === old ? nfp : r.appr_fp }, j);
        res.rows++;
      });
      i18nUpsert_('i18n_src', I18N_SRC_HEAD, 'id', { id: s.id, fp: nfp }, null);
    });
    res.recipes++;
  });
  i18nLog_('Claude Code', 'migrate-fp', '', '', 'ok', '換指紋 ' + res.recipes + ' 支、譯文 ' + res.rows + ' 列；已是新的 ' + res.same + '；剩 ' + res.left);
  return { data: res };
}
ACTIONS.i18nMigrateFp = function (p, a) { return i18nMigrateFp_(p, a); };
NOAUTH.i18nMigrateFp = 1; SELFLOCK.i18nMigrateFp = 1;
ACTIONS.i18nPull = function (p, a) { return i18nPull_(p, a); };
ACTIONS.i18nPut = function (p, a) { return i18nPut_(p, a); };
NOAUTH.i18nPull = 1; NOAUTH.i18nPut = 1; READS.i18nPull = 1; SELFLOCK.i18nPut = 1;   /* 通行碼在函式裡驗；Put 自己拿鎖 */

/* 編輯器手動驗證：更新原文＋翻第一支的日文 */
function i18nTestNow() {
  var r = i18nRefresh_(false, Date.now() + 200000); Logger.log(JSON.stringify(r));
  var s = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows.filter(function (x) { return x.state === 'ok'; })[0];
  if (s) Logger.log(JSON.stringify(i18nWork_(s.id, 'ja', '編輯器')));
}

/* 同步檢查（2026-11-11 經營者：「檢查外語食譜是不是和自己做食譜實體保持同步更新到今天」）：只讀、不寫任何表。
   當場讀後台 /Recipes 清單＋每支台灣區英文啟用中食譜的內容，跟 i18n_src（平板用的原文）比：清單有沒有多／少、更新日、內容指紋。
   另列中文版（同名、台灣、繁中、啟用）更新日比英文版晚的（英文版可能還沒跟著改）。要唯讀預覽通行碼。 */
function i18nSyncCheck_(p) {
  if (!(p.pkey && i18nHash_(String(p.pkey)) === I18N_PREVIEW_HASH)) { Utilities.sleep(1000); throw fail_('要通行碼', 'auth'); }
  var t0 = Date.now(), sess = rbSession_(), html = rbAuthedGet_(sess, '/Recipes').text; rbSaveSession_(sess);
  var rows = nmListRows_(html), meta = lgListMeta_(html);
  if (rows.length < 100) throw fail_('後台食譜清單只讀到 ' + rows.length + ' 支', 'rbpage');
  rows.forEach(function (r) { var m = meta[String(r.id).toLowerCase()] || {}; r.id = String(r.id).toLowerCase(); r.region = m.region || ''; });
  var live = rows.filter(function (r) { return r.region === '台灣' && r.lang === '英文' && r.status === 'Active'; });
  var sm = i18nMeta_('i18n_src', I18N_SRC_HEAD).rows, src = {}; sm.forEach(function (s) { src[s.id] = s; });
  var pm = i18nPickMap_(), onLive = {}, out = [];
  var ids = live.map(function (r) { return r.id; }), pages = [];
  for (var i = 0; i < ids.length; i += 5) pages = pages.concat(i18nFetchPages_(ids.slice(i, i + 5)));
  live.forEach(function (r, i) {
    onLive[r.id] = 1;
    var s = src[r.id], o = { id: r.id, name: r.name, upd_now: r.upd, upd_tab: s ? s.upd : '', state: s ? s.state : '（平板沒有）', on: i18nOn_(pm, r.id) };
    try {
      var pg = pages[i]; if (pg.code !== 200) throw new Error('HTTP ' + pg.code);
      var nw = i18nParse_(pg.text, r.id); o.same = !!s && i18nFp_(nw) === s.fp;
      if (s && !o.same) {   /* 差在哪：頁數、哪幾頁標題／內文不同（列前 6 處） */
        var old = i18nReadJ_(sm[0] ? i18nMeta_('i18n_src', I18N_SRC_HEAD).sh : null, s._r, I18N_SRC_HEAD) || { slides: [] };
        o.pages_now = nw.slides.length; o.pages_tab = old.slides.length; o.diff = [];
        if (nw.name !== old.name) o.diff.push('名稱'); if (nw.pres !== old.pres) o.diff.push('保存方式');
        var ok = {}; old.slides.forEach(function (x) { ok[i18nSlideKey_(x)] = 1; });
        nw.slides.forEach(function (x, j) { if (!ok[i18nSlideKey_(x)] && o.diff.length < 8) o.diff.push('第' + (j + 1) + '頁「' + x.t + '」：' + i18nPlain_(x.h).replace(/\s+/g, ' ').slice(0, 60)); });
      }
    } catch (e) { o.same = null; o.err = errMsg_(e); }
    out.push(o);
  });
  var extra = sm.filter(function (s) { return s.state === 'ok' && !onLive[s.id]; }).map(function (s) { return { id: s.id, name: s.name }; });
  var zh = {}; rows.forEach(function (r) { if (r.region === '台灣' && r.lang === '繁體中文' && r.status === 'Active') zh[i18nNameKey_(r.name)] = r; });
  var zhNewer = [];
  live.forEach(function (r) {
    var k = i18nNameKey_(String(r.name).split('｜').slice(1).join('｜')), z = k && zh[k];
    if (z && z.upd > r.upd) zhNewer.push({ en: r.name, en_upd: r.upd, zh: z.name, zh_upd: z.upd });
  });
  return { data: { at: now_(), backend_total: rows.length, live: live.length, src_ok: sm.filter(function (s) { return s.state === 'ok'; }).length, refresh_at: i18nProp_('I18N_REFRESH_AT'), items: out, extra: extra, zh_newer: zhNewer, sec: Math.round((Date.now() - t0) / 1000) } };
}
ACTIONS.i18nSyncCheck = function (p, a) { return i18nSyncCheck_(p, a); };
NOAUTH.i18nSyncCheck = 1; READS.i18nSyncCheck = 1;
