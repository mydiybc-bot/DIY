/* ===== nm_scan.gs — 公版圖片／影片資料自動掃描（排程，2026-10-14）=====
   經營者：「幫我做後端排程，自己掃」。取代 nr_mediapool.gs 的一次性常數：
   ① 每 10 分鐘觸發 nmTick：每天 03:00 後第一次執行時，讀門市食譜後台的食譜清單（/Recipes），跟 media_scan 分頁比對——
      新食譜、更新日期變了的、上次掃失敗的 → 標成「待掃」；之後每一次執行掃最多 4 分鐘（約 150～250 支），掃完為止。
   ② 掃一支＝登入後台讀 /Steps/RecipeDetails2/{id}：步驟標題在 .listView 的 .slideNum[data-slide-to]，每一頁的影片在 <video><source src>、圖片在 <img class="titleImage">。
   ③ 全部掃完 → nmAggregate_：讀雲端空間檔案清單（公開，?restype=container&comp=list）拿 MD5／大小，同內容合併；
      以「近 15 個月有更新」的食譜為母體，每個步驟標題算各版本用了幾支 → 寫 media_pool／media_titles／media_meta 分頁。
   ④ nrMediaFor_ 改讀分頁（暫存 10 分鐘）；分頁還沒有資料時退回常數 NR_MEDIA_POOL。
   需要新權限 script.scriptapp（建觸發器）：第一次由經營者在編輯器執行 nmSetup 授權。
   2026-10-18 經營者「用 2026 年的食譜交叉比對步驟內容」：每一步多存步驟文字 x（.recipeContent 純文字，最多 300 字）；
   nmAggregate_ 另寫 media_match 分頁（2026-01-01 以後更新的繁中食譜，每一步一列）給 nm_match.gs 的五燈號配對用。
   順手修：更新日在分頁裡會被轉成日期格子，原本 String(日期) 跟 '2025/07/18' 比永遠成立 → 「近 15 個月」其實沒在篩（1,551 支全算）；改用 nmYmd_ 比。 */
var NM_SCAN = 'media_scan', NM_POOL = 'media_pool', NM_TITLES = 'media_titles', NM_META = 'media_meta';
var NM_MATCH = 'media_match', NM_MATCH_FROM = '2026/01/01';   /* 五燈號配對的母體：這天以後更新的繁中食譜（經營者 10/18：用 2026 年的食譜） */
var NM_MATCH_HEAD = ['recipe_id', 'name', 'base', 'step', 'key', 'text', 'path', 'kind', 'mb', 'grp'];
var NM_BUDGET_MS = 240000;        /* 每次觸發最多掃 4 分鐘（GAS 上限 6 分鐘） */
var NM_RECENT_MONTHS = 15;        /* 「近期」＝更新日在 15 個月內 */
var NM_DAILY_HOUR = 1;            /* 每天這個小時以後做一次清單比對（2026-10-17 3→1：經營者要 🧬 食譜演化每天凌晨 1～2 點更新，掃完接著算 lineage.gs） */
var NM_BLOB_LIST = 'https://diybcstorage.blob.core.windows.net/stepmedias?restype=container&comp=list&maxresults=5000';
var NM_MAX_MB = 5;
var NM_SCAN_HEAD = ['recipe_id', 'name', 'lang', 'status', 'upd', 'scanned_at', 'err', 'steps_json'];

/* ---------- 觸發器 ---------- */
function nmSetup() {   /* 編輯器執行一次：建每 10 分鐘的觸發器（已有就不重建）；順便建分頁 */
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'nmTick'; });
  if (!has) ScriptApp.newTrigger('nmTick').timeBased().everyMinutes(10).create();
  nmSheet_(NM_SCAN, NM_SCAN_HEAD); nmSheet_(NM_META, ['key', 'value']);
  nmMetaSet_('setup_at', nmNow_());
  return has ? 'trigger already exists' : 'trigger created';
}
function nmTick() {   /* 觸發器進入點：有待掃的就掃；沒有就看今天是否該做清單比對；都沒有就結束 */
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  var t0 = Date.now();
  try { nmRun_(NM_BUDGET_MS, false); try { lgMaybe_(t0); } catch (e) {} } finally { lock.releaseLock(); }   /* 2026-10-17：掃完接著算 🧬 食譜演化（lineage.gs，一天一次） */
}
function nmRun_(budget, force) {
  var t0 = Date.now(), P = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), hour = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  var pending = Number(P.getProperty('NM_PENDING') || 0);
  if (P.getProperty('NM_TEXT_V') !== '1') {   /* 2026-10-18：一次性——2026 年的食譜重掃一次，補上步驟文字 */
    pending = nmMarkText_(); P.setProperty('NM_TEXT_V', '1'); P.setProperty('NM_PENDING', String(pending)); nmMetaSet_('pending', String(pending));
  }
  var listed = P.getProperty('NM_LIST_DATE') === today;
  if (!pending && !force && (listed || hour < NM_DAILY_HOUR)) return { idle: true };
  var sess = rbSession_();
  if (!listed || force) {
    var list = nmListRows_(rbAuthedGet_(sess, '/Recipes').text);
    if (list.length < 100) throw fail_('後台食譜清單只讀到 ' + list.length + ' 支，先不比對', 'rbpage');
    pending = nmRefresh_(list);
    P.setProperty('NM_LIST_DATE', today); P.setProperty('NM_PENDING', String(pending));
    nmMetaSet_('list_at', nmNow_()); nmMetaSet_('recipes_total', String(list.length)); nmMetaSet_('pending', String(pending));
  }
  var done = 0;
  if (pending > 0) {
    var r = nmScanBatch_(sess, Math.max(1000, budget - (Date.now() - t0)));
    done = r.done; pending = r.left;
    P.setProperty('NM_PENDING', String(pending)); nmMetaSet_('pending', String(pending)); nmMetaSet_('scan_at', nmNow_());
  }
  rbSaveSession_(sess);
  var agg = null;
  if (!pending && (done > 0 || force || !nmMetaGet_('aggregated_at'))) agg = nmAggregate_();
  return { idle: false, scanned: done, pending: pending, aggregated: agg };
}

/* ---------- 分頁工具 ---------- */
function nmSheet_(name, head) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function nmMetaGet_(k) {
  var sh = nmSheet_(NM_META, ['key', 'value']), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === k) { var x = v[i][1]; return x instanceof Date ? Utilities.formatDate(x, TZ, 'yyyy-MM-dd HH:mm:ss') : String(x); }   /* 舊值被轉成日期格子的，照台北時間排回字串 */
  return '';
}
function nmMetaSet_(k, val) {   /* 2026-10-18：存成文字格式（不讓試算表把時間、日期轉成日期格子） */
  var sh = nmSheet_(NM_META, ['key', 'value']), v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === k) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String(val)); return; }
  var r = sh.getLastRow() + 1; sh.getRange(r, 1, 1, 2).setNumberFormat('@').setValues([[k, String(val)]]);
}
/* 「款」：名稱去掉期別前綴（115 11月大菜單_）、英文段、括號、版本字樣後相同＝同一款（nmAggregate_、nm_match.gs 共用） */
function nmBase_(n) { return String(n || '').replace(/^.*｜/, '').replace(/^\d{3}\s*[^_]*_/, '').replace(/\|.*$/, '').replace(/（.*?）|\(.*?\)|【.*?】/g, '').replace(/[_\s]+(海綿蛋糕版|自選版|已解凍版本)$/, '').replace(/\s/g, '').trim(); }
/* 更新日：分頁裡可能是日期格子（Date）或字串 '2026/10/16' → 一律 'yyyy/MM/dd' */
function nmYmd_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy/MM/dd');
  return String(v == null ? '' : v).trim().replace(/-/g, '/').replace(/^(\d{4})\/(\d{1,2})\/(\d{1,2}).*$/, function (a, y, m, d) { return y + '/' + (m.length < 2 ? '0' + m : m) + '/' + (d.length < 2 ? '0' + d : d); });
}
/* 一次性：2026 年更新、步驟還沒有文字（steps_json 沒有 "x":）的食譜清掉掃描時間 → 重掃；回全部待掃數 */
function nmMarkText_() {
  var sh = nmSheet_(NM_SCAN, NM_SCAN_HEAD), lr = sh.getLastRow();
  if (lr < 2) return 0;
  var v = sh.getRange(2, 1, lr - 1, NM_SCAN_HEAD.length).getValues(), col = [], left = 0;
  v.forEach(function (r) {
    var redo = r[5] && nmYmd_(r[4]) >= NM_MATCH_FROM && String(r[7] || '').indexOf('"x":') < 0;
    col.push([redo ? '' : r[5]]);
    if (redo || !r[5]) left++;
  });
  sh.getRange(2, 6, col.length, 1).setValues(col);
  return left;
}
function nmNow_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

/* ---------- 後台食譜清單（含更新日期） ---------- */
function nmListRows_(html) {
  var out = [], re = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi, m;
  while ((m = re.exec(html))) {
    var row = m[1], id = rbAttr_((row.match(/<input\b[^>]*class="recipe-check"[^>]*>/i) || [''])[0], 'value');
    if (!id) continue;
    var tds = row.match(/<td\b[^>]*>[\s\S]*?<\/td>/gi) || [];
    out.push({ id: id, name: tds[1] ? rbText_(tds[1]) : '', lang: tds[5] ? rbText_(tds[5]) : '', status: tds[8] ? rbText_(tds[8]) : '', upd: tds[9] ? rbText_(tds[9]).replace(/\s+/g, '') : '' });
  }
  return out;
}
/* 清單 vs media_scan：新的加列、更新日變了或上次失敗的清掉 scanned_at；後台已不存在的刪列。回待掃數。 */
function nmRefresh_(list) {
  var sh = nmSheet_(NM_SCAN, NM_SCAN_HEAD), lr = sh.getLastRow(), rows = lr > 1 ? sh.getRange(2, 1, lr - 1, 7).getValues() : [], idx = {};
  rows.forEach(function (r, i) { idx[String(r[0])] = i; });
  var seen = {}, pending = 0, add = [];
  list.forEach(function (r) {
    seen[r.id] = 1;
    var i = idx[r.id];
    if (i == null) { add.push([r.id, r.name, r.lang, r.status, r.upd, '', '', '']); pending++; return; }
    var row = rows[i], changed = String(row[4]) !== r.upd || String(row[1]) !== r.name || String(row[3]) !== r.status;
    var retry = row[6] && !row[5];
    if (changed || row[6]) { sh.getRange(i + 2, 2, 1, 6).setValues([[r.name, r.lang, r.status, r.upd, '', '']]); pending++; }
    else if (!row[5] || retry) pending++;
  });
  if (add.length) sh.getRange(lr + 1, 1, add.length, NM_SCAN_HEAD.length).setValues(add);
  var del = [];
  rows.forEach(function (r, i) { if (!seen[String(r[0])]) del.push(i + 2); });
  del.sort(function (a, b) { return b - a; }).forEach(function (n) { sh.deleteRow(n); });
  return pending;
}

/* ---------- 掃一批 ---------- */
function nmScanBatch_(sess, budget) {
  var t0 = Date.now(), sh = nmSheet_(NM_SCAN, NM_SCAN_HEAD), lr = sh.getLastRow();
  if (lr < 2) return { done: 0, left: 0 };
  var rows = sh.getRange(2, 1, lr - 1, 7).getValues(), done = 0, left = 0;
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][5]) continue;            /* 已掃 */
    if (Date.now() - t0 > budget) { left++; continue; }
    var id = String(rows[i][0]), res;
    try { res = nmScanOne_(sess, id); sh.getRange(i + 2, 6, 1, 3).setValues([[nmNow_(), '', nmJson_(res)]]); }
    catch (e) {
      if (e && e.code === 'rbrelogin') { rbDropSession_(); sess.jar = rbFreshSession_().jar; try { res = nmScanOne_(sess, id); sh.getRange(i + 2, 6, 1, 3).setValues([[nmNow_(), '', nmJson_(res)]]); } catch (e2) { sh.getRange(i + 2, 6, 1, 2).setValues([[nmNow_(), errMsg_(e2).slice(0, 200)]]); } }
      else sh.getRange(i + 2, 6, 1, 2).setValues([[nmNow_(), errMsg_(e).slice(0, 200)]]);
    }
    done++;
  }
  return { done: done, left: left };
}
/* 2026-10-18：試算表一格最多 50,000 字——步驟很多的食譜先把步驟文字縮短，還太長就不存文字（標題和檔名照存） */
function nmJson_(res) {
  var j = JSON.stringify(res);
  if (j.length <= 45000) return j;
  j = JSON.stringify(res.map(function (s) { return { t: s.t, m: s.m, k: s.k, x: String(s.x || '').slice(0, 60) }; }));
  if (j.length <= 45000) return j;
  return JSON.stringify(res.map(function (s) { return { t: s.t, m: s.m, k: s.k, x: '' }; }));
}
/* 讀一支食譜的每一步：[{t:標題, m:素材路徑(stepmedias/ 後面), k:'v'|'i'|''}]；只算有 .recipeContent 的頁（客製選項頁不算） */
function nmScanOne_(sess, id) {
  var r = rbAuthedGet_(sess, '/Steps/RecipeDetails2/' + encodeURIComponent(id));
  if (r.code !== 200) throw rbErr_('後台食譜頁 HTTP ' + r.code, r.code === 404 ? 'rbgone' : 'rbnet');
  return nmParseDetails_(r.text);
}
function nmParseDetails_(html) {
  var titles = {}, lv = (html.match(/<div\b[^>]*class="listView"[\s\S]*?<\/div>\s*<\/div>\s*<\/div>/i) || [''])[0] || html;
  var re = /<div\b([^>]*\bslideNum\b[^>]*)>([\s\S]*?)<\/div>/gi, m;
  while ((m = re.exec(lv))) {
    var n = rbAttr_('<x' + m[1] + '>', 'data-slide-to'); if (!n || titles[n]) continue;
    titles[n] = rbTrim_(rbDecode_(m[2].replace(/<[^>]+>/g, ' '))).replace(/^\d+\s*/, '');
  }
  var parts = html.split(/<div\b[^>]*class="carousel-item\b/i), out = [];
  for (var i = 1; i < parts.length; i++) {
    var chunk = parts[i], idx = i - 1;
    if (!/class="recipeContent"/i.test(chunk)) continue;
    var v = (chunk.match(/<video\b[\s\S]*?<source\b[^>]*>/i) || [''])[0], src = v ? rbAttr_(v.match(/<source\b[^>]*>/i)[0], 'src') : '';
    var kind = src ? 'v' : '';
    if (!src) { var im = (chunk.match(/<img\b[^>]*class="titleImage"[^>]*>/i) || [''])[0]; src = im ? rbAttr_(im, 'src') : ''; kind = src ? 'i' : ''; }
    var ci = chunk.search(/<div\b[^>]*class="recipeContent"/i), x = ci >= 0 ? rbText_(nmInner_(chunk, ci)).slice(0, 200) : '';   /* 2026-10-18：步驟文字（最多 200 字） */
    out.push({ t: titles[String(idx)] || '', m: nmPath_(src), k: kind, x: x });
  }
  return out;
}
/* 從 start 位置那個 <div …> 起，找到配對的 </div>，回中間的 HTML（內容裡有巢狀 div 也不會切太早） */
function nmInner_(html, start) {
  var re = /<\/?div\b[^>]*>/gi, depth = 0, from = -1, m;
  re.lastIndex = start;
  while ((m = re.exec(html))) {
    if (m[0].charAt(1) !== '/') { if (depth === 0) from = re.lastIndex; depth++; }
    else if (depth > 0) { depth--; if (depth === 0) return html.slice(from, m.index); }
  }
  return from >= 0 ? html.slice(from) : '';
}
function nmPath_(u) { u = String(u || '').split('?')[0].split('#')[0]; var i = u.indexOf('stepmedias/'); return i >= 0 ? u.slice(i + 11) : u.replace(/^https?:\/\/[^/]+\//, ''); }

/* ---------- 雲端檔案清單：stepmedias/ 路徑 → {md5, bytes} ---------- */
function nmBlobMap_() {
  var map = {}, marker = '', pages = 0;
  while (pages++ < 12) {
    var url = NM_BLOB_LIST + (marker ? '&marker=' + encodeURIComponent(marker) : '');
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw fail_('雲端檔案清單讀不到（HTTP ' + r.getResponseCode() + '）', 'blob');
    var x = r.getContentText(), re = /<Blob>([\s\S]*?)<\/Blob>/g, m;
    while ((m = re.exec(x))) {
      var b = m[1], name = (b.match(/<Name>([^<]*)<\/Name>/) || ['', ''])[1];
      map[name] = { md5: (b.match(/<Content-MD5>([^<]*)<\/Content-MD5>/) || ['', ''])[1], bytes: Number((b.match(/<Content-Length>(\d+)<\/Content-Length>/) || ['', '0'])[1]) };
    }
    marker = (x.match(/<NextMarker>([^<]*)<\/NextMarker>/) || ['', ''])[1];
    if (!marker) break;
  }
  return map;
}

/* ---------- 彙整：近期食譜、同內容合併 → media_pool／media_titles／media_meta；2026-10-18 起另寫 media_match（五燈號配對用） ---------- */
function nmAggregate_() {
  var sh = nmSheet_(NM_SCAN, NM_SCAN_HEAD), lr = sh.getLastRow();
  if (lr < 2) return { pool: 0 };
  var rows = sh.getRange(2, 1, lr - 1, NM_SCAN_HEAD.length).getValues();
  var from = new Date(); from.setMonth(from.getMonth() - NM_RECENT_MONTHS);
  var fromKey = Utilities.formatDate(from, TZ, 'yyyy/MM/dd');
  var blob = nmBlobMap_();
  var gkey = function (path) { var b = blob[path]; return (b && b.md5) ? 'md5:' + b.md5 : 'p:' + path; };
  var base = nmBase_;
  var okFile = function (p) { var b = blob[p]; return !!(b && b.bytes <= NM_MAX_MB * 1048576 && /\.(mp4|mov|m4v|jpg|jpeg|png|gif)$/i.test(p)); };
  var G = {}, T = {}, nRecent = 0, nScanned = 0, M = [], mRec = 0, gPaths = {};
  rows.forEach(function (r) {
    if (!r[5] || r[6] || !r[7]) return;
    nScanned++;
    var upd = nmYmd_(r[4]);
    if (upd < fromKey) return;
    nRecent++;
    var steps; try { steps = JSON.parse(r[7]); } catch (e) { return; }
    var rb = base(r[1]), seenT = {};
    if (upd >= NM_MATCH_FROM && String(r[2]) === '繁體中文') {   /* 2026-10-18：五燈號配對母體（每一步一列，沒放圖片影片的步驟也要，用來判斷「通常不放」） */
      mRec++;
      steps.forEach(function (s, i) {
        var g = s.m ? gkey(s.m) : '';
        if (g) (gPaths[g] = gPaths[g] || {})[s.m] = 1;
        M.push([String(r[0]), String(r[1]), rb, i + 1, nrTitleKey_(s.t), nmxNorm_(s.x).slice(0, 200), s.m || '', s.m ? (s.k || '') : '', '', g]);
      });
    }
    steps.forEach(function (s) {
      if (!s.m) return;
      var t = nrTitleKey_(s.t); if (!t) return;
      var g = gkey(s.m), x = G[g] = G[g] || { k: s.k, paths: {}, t: {} };
      x.paths[s.m] = 1; (x.t[t] = x.t[t] || {})[rb] = 1;
      if (!seenT[t]) { seenT[t] = 1; (T[t] = T[t] || {})[rb] = 1; }
    });
  });
  var pool = [];
  Object.keys(G).forEach(function (g) {
    var x = G[g], t = {}, mx = 0;
    Object.keys(x.t).forEach(function (tt) { var c = Object.keys(x.t[tt]).length; t[tt] = c; if (c > mx) mx = c; });
    if (mx < 2) return;
    var best = '', bb = -1;
    Object.keys(x.paths).forEach(function (p) {
      var b = blob[p]; if (!b || b.bytes > NM_MAX_MB * 1048576 || !/\.(mp4|mov|m4v|jpg|jpeg|png|gif)$/i.test(p)) return;
      if (b.bytes > bb) { bb = b.bytes; best = p; }
    });
    if (!best) return;
    pool.push([best, x.k, Math.round(bb / 1048576 * 100) / 100, JSON.stringify(t), mx]);
  });
  pool.sort(function (a, b) { return b[4] - a[4]; });
  var titles = Object.keys(T).map(function (t) { return [t, Object.keys(T[t]).length]; }).filter(function (r) { return r[1] >= 3; }).sort(function (a, b) { return b[1] - a[1]; });
  var ss = ss_();
  var sp = ss.getSheetByName(NM_POOL) || ss.insertSheet(NM_POOL); sp.clear();
  sp.getRange(1, 1, 1, 5).setValues([['path', 'kind', 'mb', 'titles_json', 'max_recipes']]).setFontWeight('bold');
  if (pool.length) sp.getRange(2, 1, pool.length, 5).setValues(pool);
  var st = ss.getSheetByName(NM_TITLES) || ss.insertSheet(NM_TITLES); st.clear();
  st.getRange(1, 1, 1, 2).setValues([['title', 'recent_recipes']]).setFontWeight('bold');
  if (titles.length) st.getRange(2, 1, titles.length, 2).setValues(titles);
  /* media_match：每一步的「代表檔」＝同內容裡能用的那個（≤5MB、圖片影片副檔名；同 MD5 內容一樣，挑哪個都行，先挑這一步自己的） */
  var rep = {}, byMd5 = {};
  Object.keys(blob).forEach(function (p) { var h = blob[p].md5; if (h) (byMd5[h] = byMd5[h] || []).push(p); });
  Object.keys(gPaths).forEach(function (g) {
    var cand = Object.keys(gPaths[g]);
    if (g.indexOf('md5:') === 0) (byMd5[g.slice(4)] || []).forEach(function (p) { if (!gPaths[g][p]) cand.push(p); });
    for (var i = 0; i < cand.length; i++) if (okFile(cand[i])) { rep[g] = cand[i]; break; }
  });
  M.forEach(function (m) {
    if (!m[9]) return;
    var p = rep[m[9]] || '';
    m[6] = p; m[8] = p ? Math.round(blob[p].bytes / 1048576 * 100) / 100 : '';   /* 沒有能用的檔（都 >5MB）→ path 空白，只拿來算「同一版」 */
  });
  var sm = ss.getSheetByName(NM_MATCH) || ss.insertSheet(NM_MATCH); sm.clear();
  sm.getRange(1, 1, Math.max(M.length + 1, 2), NM_MATCH_HEAD.length).setNumberFormat('@');
  sm.getRange(1, 1, 1, NM_MATCH_HEAD.length).setValues([NM_MATCH_HEAD]).setFontWeight('bold'); sm.setFrozenRows(1);
  if (M.length) sm.getRange(2, 1, M.length, NM_MATCH_HEAD.length).setValues(M.map(function (m) { return m.map(function (v) { return String(v); }); }));
  nmMetaSet_('aggregated_at', nmNow_()); nmMetaSet_('recent_from', fromKey); nmMetaSet_('recipes_recent', String(nRecent)); nmMetaSet_('recipes_scanned', String(nScanned));
  nmMetaSet_('pool_groups', String(pool.length)); nmMetaSet_('titles', String(titles.length)); nmMetaSet_('blob_files', String(Object.keys(blob).length));
  nmMetaSet_('match_from', NM_MATCH_FROM); nmMetaSet_('match_recipes', String(mRec)); nmMetaSet_('match_steps', String(M.length));
  try { CacheService.getScriptCache().remove('nmpool:v1'); } catch (e) {}
  try { nmxDrop_(); } catch (e) {}
  return { pool: pool.length, titles: titles.length, recent: nRecent, scanned: nScanned, match_recipes: mRec, match_steps: M.length };
}

/* ---------- 給 nrMediaFor_ 用：讀分頁（暫存 10 分鐘），沒資料就回 null（退回常數） ---------- */
function nmPool_() {
  var cache = CacheService.getScriptCache(), key = 'nmpool:v1', s = cacheGetBig_(cache, key);
  if (s) { try { var o = JSON.parse(s); if (o && o.pool) return o; } catch (e) {} }
  try {
    var ss = ss_(), sp = ss.getSheetByName(NM_POOL), st = ss.getSheetByName(NM_TITLES);
    if (!sp || !st || sp.getLastRow() < 2 || st.getLastRow() < 2) return null;
    var pv = sp.getRange(2, 1, sp.getLastRow() - 1, 4).getValues(), tv = st.getRange(2, 1, st.getLastRow() - 1, 2).getValues();
    var pool = pv.map(function (r) { var t = {}; try { t = JSON.parse(r[3]); } catch (e) {} return { u: String(r[0]), k: String(r[1]), mb: Number(r[2]) || 0, t: t }; });
    var titles = {}; tv.forEach(function (r) { titles[String(r[0])] = Number(r[1]) || 0; });
    var o = { pool: pool, titles: titles, meta: { scanned: nmMetaGet_('aggregated_at'), recipes: Number(nmMetaGet_('recipes_recent')) || 0, recent_from: nmMetaGet_('recent_from') } };
    cachePutBig_(cache, key, JSON.stringify(o), 600);
    return o;
  } catch (e) { return null; }
}

/* ---------- 動作：狀態、手動推一次、單支試掃 ---------- */
function nmStatus_(p, auth) {
  var P = PropertiesService.getScriptProperties(), keys = ['setup_at', 'list_at', 'scan_at', 'aggregated_at', 'recent_from', 'recipes_total', 'recipes_scanned', 'recipes_recent', 'pending', 'pool_groups', 'titles', 'blob_files', 'match_from', 'match_recipes', 'match_steps'], meta = {};
  keys.forEach(function (k) { meta[k] = nmMetaGet_(k); });
  meta.trigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'nmTick'; });
  meta.list_date = P.getProperty('NM_LIST_DATE') || '';
  meta.source = nmPool_() ? 'sheet' : 'constant';
  return { id: '', data: meta };
}
function nmKick_(p, auth) {   /* 立刻跑一次（最多 4 分鐘）；同時只跑一個 */
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw fail_('掃描正在跑，等它結束再按', 'busy');
  try { var r = nmRun_(NM_BUDGET_MS, !!p.force); return { id: '', data: r, msg: '公版掃描：' + JSON.stringify(r).slice(0, 300) }; }
  finally { lock.releaseLock(); }
}
function nmProbe_(p, auth) {   /* 試掃一支，回解析結果（驗證解析器用） */
  var id = String(p.recipe_id || '').trim(); if (!/^[0-9a-f-]{36}$/.test(id)) throw fail_('recipe_id 格式不對', 'input');
  var sess = rbSession_(), r = nmScanOne_(sess, id); rbSaveSession_(sess);
  return { id: '', data: { steps: r, n: r.length } };
}
ACTIONS.nmStatus = function (p, a) { return nmStatus_(p, a); };
ACTIONS.nmKick = function (p, a) { return nmKick_(p, a); };
ACTIONS.nmProbe = function (p, a) { return nmProbe_(p, a); };
READS.nmStatus = 1; READS.nmProbe = 1;
SELFLOCK.nmKick = 1;   /* 自己用 LockService 擋重複，不佔全系統的鎖；不需要身分密碼（只掃後台、寫自己的分頁，重跑無害） */
