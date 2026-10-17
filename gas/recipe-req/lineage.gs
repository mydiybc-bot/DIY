/* ===== lineage.gs — 🧬 食譜演化（家族樹＋時間軸），2026-10-17 =====
   經營者：「後台 1000 多支食譜很多是一代一代演化來的，想在新品申請表加一個功能頁看演化」；要看「家族樹＋時間軸」，每天凌晨 1～2 點自動更新。
   資料來源（都不另外爬後台步驟）：
   ① nm_scan.gs 每天掃好的 media_scan 分頁（每支食譜每一步的標題＋圖片影片路徑 stepmedias/{上傳時的步驟id}/{上傳時的食譜id}.副檔名）
   ② 後台食譜清單 /Recipes（一次）：類別、店別、地區
   ③ 雲端空間檔案清單（公開）：每個檔的上傳日 → 推「這支食譜哪天出生」（自己上傳的第一個檔／封面）
   家族怎麼分：複製食譜時，步驟圖片影片會沿用原本那支的檔（路徑裡留著原食譜編號）→ 兩支食譜共用 3 個以上「不是到處都在用」的檔（用的食譜 ≤ 20 支），
   或共用 2 個且佔較少那支的 4 成以上 → 同一家族；名稱去掉期別／版本字樣後相同的也算同家族。
   誰是誰的上一代：家族內依出生日排，每支找「比它早出生、跟它共用最多檔」的那支當上一代（名稱相同加分、步驟標題相近加分）。
   只讀：不寫後台；只寫本試算表的 recipe_lineage／lineage_meta 分頁（每次整張重寫）。 */
var LG_TAB = 'recipe_lineage', LG_META = 'lineage_meta';
var LG_HEAD = ['recipe_id', 'name', 'lang', 'region', 'status', 'cat', 'store', 'upd', 'birth', 'birth_src', 'kuan', 'kuan_birth', 'kuan_size', 'ver', 'family_id', 'family_kuans', 'family_size', 'parent_kuan', 'gen', 'shared', 'steps', 'media_steps', 'cover', 'cover_kb', 'shared_steps'];   /* cover：封面圖在 recipeimages/ 下的路徑（2026-10-17 晚加，頁面每款挑一張當代表圖） */
var LG_HOUR = 1;              /* 每天 1 點以後、公版掃描做完，算一次 */
var LG_COMMON = 20;           /* 超過這麼多「款」都在用的檔（教學影片等）不拿來認親 */
var LG_MIN_SHARED = 3;
/* 2026-10-17 晚 經營者：「製作步驟有多個步驟相似」也算親戚——步驟標題（去掉數字、括號、空白）相同的步驟；
   超過 LG_TCOMMON 款都有的標題（教學影片、軟化奶油…）不算；兩款共用 ≥ LG_TMIN 個不常見的步驟標題，且佔較少那款的 LG_TRATIO 以上 → 同家族 */
var LG_TCOMMON = 15, LG_TMIN = 5, LG_TRATIO = 0.6;        /* 兩款共用幾個非通用檔以上算有親戚關係（或共用 2 個且佔較少那款的一半以上） */
var LG_COVER_LIST = 'https://diybcstorage.blob.core.windows.net/recipeimages?restype=container&comp=list&maxresults=5000';

/* ---------- 「款」：名稱去掉期別、版本、英文名、葷素標示後相同＝同一款（中英文版、備份、匙版、半成品版都算同一款） ---------- */
function lgBase_(n) {
  var segs = String(n || '').split(/[|｜]/), cjk = /[㐀-鿿]/;
  var s = segs.filter(function (x) { return cjk.test(x); })[0] || segs[0] || '';
  return s.replace(/^\**\s*(原版未拆|舊版|備份|半成品|分裝食材)[_\s]*/, '').replace(/^\d{3}\s*[^_]*_/, '')
    .replace(/（.*?）|\(.*?\)|【.*?】|《|》/g, '')
    .replace(/[_\s]*(海綿蛋糕版|自選版|已解凍版本|原版備用|半成品版|要更新)$/, '').replace(/\s*\d+(\.\d+)?$/, '').replace(/[\s*！!]/g, '').trim();
}
function lgTitleKey_(t) { return String(t == null ? '' : t).replace(/（.*?）|\(.*?\)/g, '').replace(/[\s\d０-９.,，、。:：!！?？~～\-－_＿#＃*]/g, '').trim(); }
/* 路徑 {步驟id}/{食譜id}.副檔名 → 上傳時的食譜id */
function lgOrigin_(path) { var m = String(path || '').match(/^[0-9a-f-]{36}\/([0-9a-f-]{36})\./i); return m ? m[1].toLowerCase() : ''; }

/* ---------- 核心：輸入食譜清單＋檔案日期 → 每支食譜屬於哪一款、哪個家族、款的上一代、第幾代（純計算，可在本機測） ----------
   recipes: [{id,name,lang,region,status,cat,store,upd,steps:[{t,m}]}]；dates: {stepmedias 路徑: 'yyyy/MM/dd'}；covers: {食譜id: 'yyyy/MM/dd'}
   兩層：①同一款的各版本放一起（時間軸） ②款與款之間：誰從誰複製改來（家族樹，第幾代以「款」算） */
function lgCompute_(recipes, dates, covers) {
  dates = dates || {}; covers = covers || {};
  var n = recipes.length, use = {};
  recipes.forEach(function (r) {
    r._paths = {}; r._media = 0;
    (r.steps || []).forEach(function (s) { if (s.m) { r._media++; r._paths[s.m] = 1; } });
    Object.keys(r._paths).forEach(function (p) { use[p] = (use[p] || 0) + 1; });
    r._base = lgBase_(r.name) || r.id;
    var b = covers[r.id] || '', src = b ? 'cover' : '';
    Object.keys(r._paths).forEach(function (p) { if (lgOrigin_(p) === r.id && dates[p] && (!b || dates[p] < b)) { b = dates[p]; src = 'media'; } });
    /* 雲端空間 2022/10 搬過家，更早的檔日期都是搬家日 → 比更新日還晚時用更新日 */
    if (b && r.upd && String(r.upd) < b) { b = ''; }
    r._birth = b || r.upd || ''; r._bsrc = b ? src : 'upd';
  });
  /* 只有英文名的（英文版食譜）→ 併到跟它共用最多檔的中文名那一款（共用 3 個以上） */
  var cjk = /[\u3400-\u9fff]/, pathRec = {};
  recipes.forEach(function (r, i) { Object.keys(r._paths).forEach(function (p) { if (use[p] >= 2 && use[p] <= LG_COMMON) (pathRec[p] = pathRec[p] || []).push(i); }); });
  recipes.forEach(function (r) {
    if (cjk.test(r._base)) return;
    var cnt = {};
    Object.keys(r._paths).forEach(function (p) { (pathRec[p] || []).forEach(function (j) { if (cjk.test(recipes[j]._base)) cnt[j] = (cnt[j] || 0) + 1; }); });
    var bj = -1, bc = 0; Object.keys(cnt).forEach(function (j) { if (cnt[j] > bc) { bc = cnt[j]; bj = Number(j); } });
    if (bc >= LG_MIN_SHARED) r._base = recipes[bj]._base;
  });
  /* 款 */
  var K = {}, kl = [];
  recipes.forEach(function (r, i) { var k = K[r._base]; if (!k) { k = K[r._base] = { key: r._base, mem: [], birth: '' }; kl.push(k); } k.mem.push(i); if (!k.birth || r._birth < k.birth) k.birth = r._birth; });
  /* 共用「不是到處都在用」的檔 → 款與款兩兩計數（同一個檔在兩款各算一次） */
  var byPath = {};
  recipes.forEach(function (r) { Object.keys(r._paths).forEach(function (p) { if (use[p] >= 2 && use[p] <= LG_COMMON) { var o = byPath[p] = byPath[p] || {}; o[r._base] = 1; } }); });
  var kd = {};   /* 款 → 用到的「非通用」檔數 */
  recipes.forEach(function (r) { var o = kd[r._base] = kd[r._base] || {}; Object.keys(r._paths).forEach(function (p) { if (use[p] <= LG_COMMON) o[p] = 1; }); });
  var pair = {};
  Object.keys(byPath).forEach(function (p) {
    var a = Object.keys(byPath[p]); if (a.length < 2 || a.length > LG_COMMON) return;
    for (var x = 0; x < a.length; x++) for (var y = x + 1; y < a.length; y++) { var k = a[x] < a[y] ? a[x] + '\u0001' + a[y] : a[y] + '\u0001' + a[x]; pair[k] = (pair[k] || 0) + 1; }
  });
  /* 步驟標題：每款用到哪些標題（各版本合起來），算「不常見標題」兩兩共用數 */
  var kt = {}, tuse = {};
  recipes.forEach(function (r) { var o = kt[r._base] = kt[r._base] || {}; (r.steps || []).forEach(function (s) { var t = lgTitleKey_(s.t); if (t.length >= 2) o[t] = 1; }); });
  Object.keys(kt).forEach(function (k) { Object.keys(kt[k]).forEach(function (t) { tuse[t] = (tuse[t] || 0) + 1; }); });
  var byT = {}, tdist = {};
  Object.keys(kt).forEach(function (k) { tdist[k] = 0; Object.keys(kt[k]).forEach(function (t) { if (tuse[t] <= LG_TCOMMON) { tdist[k]++; if (tuse[t] >= 2) (byT[t] = byT[t] || []).push(k); } }); });
  var tpair = {};
  Object.keys(byT).forEach(function (t) { var a = byT[t]; for (var x = 0; x < a.length; x++) for (var y = x + 1; y < a.length; y++) { var kk = a[x] < a[y] ? a[x] + '\u0001' + a[y] : a[y] + '\u0001' + a[x]; tpair[kk] = (tpair[kk] || 0) + 1; } });
  var up = {}; kl.forEach(function (k) { up[k.key] = k.key; });
  var find = function (x) { while (up[x] !== x) { up[x] = up[up[x]]; x = up[x]; } return x; };
  var link = {};
  Object.keys(pair).forEach(function (k) {
    var ab = k.split('\u0001'), a = ab[0], b = ab[1], c = pair[k], mn = Math.min(Object.keys(kd[a] || {}).length, Object.keys(kd[b] || {}).length) || 1;
    if (c >= LG_MIN_SHARED || (c >= 2 && c / mn >= 0.5)) {
      var ra = find(a), rb = find(b); if (ra !== rb) up[rb] = ra;
      (link[a] = link[a] || {})[b] = c; (link[b] = link[b] || {})[a] = c;
    }
  });
  var tlink = {};
  Object.keys(tpair).forEach(function (k) {
    var ab = k.split('\u0001'), a = ab[0], b = ab[1], c = tpair[k], mn = Math.min(tdist[a] || 0, tdist[b] || 0) || 1;
    (tlink[a] = tlink[a] || {})[b] = c; (tlink[b] = tlink[b] || {})[a] = c;
    if (c >= LG_TMIN && c / mn >= LG_TRATIO) { var ra = find(a), rb = find(b); if (ra !== rb) up[rb] = ra; }
  });
  var fam = {}; kl.forEach(function (k) { (fam[find(k.key)] = fam[find(k.key)] || []).push(k); });
  var out = new Array(n);
  Object.keys(fam).forEach(function (f) {
    var ks = fam[f].slice().sort(function (a, b) { return String(a.birth).localeCompare(String(b.birth)) || a.key.localeCompare(b.key); });
    var root = ks[0], fsize = 0; ks.forEach(function (k) { fsize += k.mem.length; });
    ks.forEach(function (k, at) {
      var par = null, best = -1, bm = 0, bt = 0;
      for (var q = 0; q < at; q++) {
        var cm = (link[k.key] || {})[ks[q].key] || 0, ct = (tlink[k.key] || {})[ks[q].key] || 0, sc = cm * 10 + ct * 2;   /* 共用圖片影片最重要，相同步驟其次 */
        if (sc > best || (sc === best && sc > 0)) { best = sc; par = ks[q]; bm = cm; bt = ct; }
      }
      if (at > 0 && best <= 0) { par = null; bm = 0; bt = 0; }   /* 跟更早的款沒有直接關係（經由後來的款才連進家族）→ 自己也是這個家族的一支源頭（第 1 代），不硬掛在別人底下 */
      k.parent = par; k.gen = par ? par.gen + 1 : 1; k.shared = bm; k.shared_steps = bt;
      var vers = k.mem.slice().sort(function (a, b) { return String(recipes[a]._birth).localeCompare(String(recipes[b]._birth)) || String(recipes[a].upd).localeCompare(String(recipes[b].upd)); });
      vers.forEach(function (i, vi) {
        var r = recipes[i];
        out[i] = { recipe_id: r.id, name: r.name, lang: r.lang || '', region: r.region || '', status: r.status || '', cat: r.cat || '', store: r.store || '', upd: r.upd || '',
          birth: r._birth, birth_src: r._bsrc, kuan: k.key, kuan_birth: k.birth, kuan_size: k.mem.length, ver: vi + 1,
          family_id: root.key, family_kuans: ks.length, family_size: fsize, parent_kuan: par ? par.key : '', gen: k.gen, shared: k.shared, shared_steps: k.shared_steps,
          steps: (r.steps || []).length, media_steps: r._media };
      });
    });
  });
  return out;
}

/* ---------- 讀資料、算、寫分頁 ---------- */
function lgBlobDates_(listUrl, sizes) {   /* sizes：有給就順便記每個檔幾 KB */
  var map = {}, marker = '', pages = 0;
  while (pages++ < 12) {
    var r = UrlFetchApp.fetch(listUrl + (marker ? '&marker=' + encodeURIComponent(marker) : ''), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw fail_('雲端檔案清單讀不到（HTTP ' + r.getResponseCode() + '）', 'blob');
    var x = r.getContentText(), re = /<Blob>([\s\S]*?)<\/Blob>/g, m;
    while ((m = re.exec(x))) {
      var b = m[1], name = (b.match(/<Name>([^<]*)<\/Name>/) || ['', ''])[1];
      var t = (b.match(/<Creation-Time>([^<]*)<\/Creation-Time>/) || b.match(/<Last-Modified>([^<]*)<\/Last-Modified>/) || ['', ''])[1];
      if (name && t) { var d = new Date(t); if (!isNaN(d)) map[name] = Utilities.formatDate(d, TZ, 'yyyy/MM/dd'); }
      if (sizes && name) sizes[name] = Math.round(Number((b.match(/<Content-Length>(\d+)<\/Content-Length>/) || ['', '0'])[1]) / 1024);
    }
    marker = (x.match(/<NextMarker>([^<]*)<\/NextMarker>/) || ['', ''])[1];
    if (!marker) break;
  }
  return map;
}
function lgListMeta_(html) {   /* 後台清單每列：地區、類別、店別（多個用頓號） */
  var out = {}, re = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi, m;
  while ((m = re.exec(html))) {
    var row = m[1], id = rbAttr_((row.match(/<input\b[^>]*class="recipe-check"[^>]*>/i) || [''])[0], 'value');
    if (!id) continue;
    var tds = row.match(/<td\b[^>]*>[\s\S]*?<\/td>/gi) || [];
    var spans = function (td) { return (String(td || '').match(/<span\b[^>]*>[\s\S]*?<\/span>/gi) || []).map(rbText_).filter(Boolean).join('、'); };
    out[id.toLowerCase()] = { region: tds[4] ? rbText_(tds[4]) : '', cat: spans(tds[6]), store: spans(tds[7]) };
  }
  return out;
}
function lgRun_() {
  var t0 = Date.now(), sh = ss_().getSheetByName(NM_SCAN);
  if (!sh || sh.getLastRow() < 2) throw fail_('公版掃描還沒有資料（media_scan 空的）', 'nodata');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, NM_SCAN_HEAD.length).getDisplayValues();   /* 顯示值：更新日是日期格子時 getValues 會變成 Date 物件 */
  var sess = rbSession_(), meta = lgListMeta_(rbAuthedGet_(sess, '/Recipes').text); rbSaveSession_(sess);
  if (Object.keys(meta).length < 100) throw fail_('後台食譜清單只讀到 ' + Object.keys(meta).length + ' 支，先不算', 'rbpage');
  var dates = lgBlobDates_(NM_BLOB_LIST), covKb = {}, covAll = lgBlobDates_(LG_COVER_LIST, covKb), covers = {};
  Object.keys(dates).forEach(function (k) { if (!/^[0-9a-f-]{36}\/[0-9a-f-]{36}\./i.test(k)) delete dates[k]; });
  var coverPath = {};
  Object.keys(covAll).forEach(function (k) {
    var id = k.split('/')[0].toLowerCase(), d = covAll[k]; if (!covers[id] || d < covers[id]) covers[id] = d;
    if (new RegExp('^' + id + '/' + id + '\\.(jpe?g|png|gif|webp)$', 'i').test(k)) coverPath[id] = k;   /* 封面＝{食譜id}/{食譜id}.jpg（wi/ 底下是說明內嵌圖，不算） */
  });
  var recipes = [];
  rows.forEach(function (r) {
    var id = String(r[0]).toLowerCase(); if (!id) return;
    var steps = []; try { steps = JSON.parse(r[7] || '[]') || []; } catch (e) {}
    var mm = meta[id] || {};
    recipes.push({ id: id, name: String(r[1]), lang: String(r[2]), status: String(r[3]), upd: String(r[4]), region: mm.region || '', cat: mm.cat || '', store: mm.store || '', steps: steps });
  });
  var out = lgCompute_(recipes, dates, covers);
  out.forEach(function (o) { o.cover = coverPath[o.recipe_id] || ''; o.cover_kb = o.cover ? (covKb[o.cover] || 0) : ''; });
  var vals = out.map(function (o) { return LG_HEAD.map(function (h) { return o[h] == null ? '' : String(o[h]); }); });
  var ss = ss_(), lt = ss.getSheetByName(LG_TAB) || ss.insertSheet(LG_TAB);
  lt.clear(); lt.getRange(1, 1, Math.max(vals.length + 1, 2), LG_HEAD.length).setNumberFormat('@');
  lt.getRange(1, 1, 1, LG_HEAD.length).setValues([LG_HEAD]).setFontWeight('bold'); lt.setFrozenRows(1);
  if (vals.length) lt.getRange(2, 1, vals.length, LG_HEAD.length).setValues(vals);
  var fams = {}, kuans = {}; out.forEach(function (o) { fams[o.family_id] = 1; kuans[o.kuan] = 1; });
  var maxGen = out.reduce(function (m, o) { return Math.max(m, o.gen); }, 0);
  var info = { at: nmNow_(), recipes: out.length, kuans: Object.keys(kuans).length, families: Object.keys(fams).length, max_gen: maxGen, multi: out.filter(function (o) { return o.family_size > 1; }).length, secs: Math.round((Date.now() - t0) / 1000) };
  lgMetaSet_(info);
  try { CacheService.getScriptCache().remove('lineage:v2'); } catch (e) {}
  return info;
}
function lgMetaSet_(o) {
  var ss = ss_(), sh = ss.getSheetByName(LG_META) || ss.insertSheet(LG_META);
  sh.clear(); var rows = Object.keys(o).map(function (k) { return [k, String(o[k])]; });
  sh.getRange(1, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
  PropertiesService.getScriptProperties().setProperty('LG_DATE', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'));
}
function lgMetaGet_() {
  var sh = ss_().getSheetByName(LG_META), o = {};
  if (sh && sh.getLastRow() > 0) sh.getRange(1, 1, sh.getLastRow(), 2).getDisplayValues().forEach(function (r) { o[r[0]] = r[1]; });
  return o;
}
/* nmTick 跑完之後叫這個：今天還沒算、已過 1 點、公版掃描今天的清單比對做過且掃完 → 算一次 */
function lgMaybe_(t0) {
  if (t0 && Date.now() - t0 > 90000) return null;   /* 這次公版掃描已經跑很久，留到下一個 10 分鐘再算（GAS 一次最多 6 分鐘） */
  var P = PropertiesService.getScriptProperties(), today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), hour = Number(Utilities.formatDate(new Date(), TZ, 'H'));
  if (P.getProperty('LG_DATE') === today || hour < LG_HOUR) return null;
  if (P.getProperty('NM_LIST_DATE') !== today || Number(P.getProperty('NM_PENDING') || 0) > 0) return null;
  try { return lgRun_(); } catch (e) { try { lgMetaSet_({ at: nmNow_(), error: errMsg_(e).slice(0, 300) }); } catch (e2) {} throw e; }
}
/* 編輯器手動：立刻算一次（不管今天算過沒） */
function lgRunNow() { var lock = LockService.getScriptLock(); lock.waitLock(30000); try { Logger.log(JSON.stringify(lgRun_())); } finally { lock.releaseLock(); } }

/* ---------- 讀取動作：整張表（壓成陣列；暫存 30 分鐘） ---------- */
function lineage_(p, auth) {
  var cache = CacheService.getScriptCache(), key = 'lineage:v2', s = cacheGetBig_(cache, key);
  if (s) { try { return { id: '', data: JSON.parse(s) }; } catch (e) {} }
  var sh = ss_().getSheetByName(LG_TAB);
  if (!sh || sh.getLastRow() < 2) return { id: '', data: { meta: lgMetaGet_(), head: LG_HEAD, rows: [] } };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, LG_HEAD.length).getDisplayValues();
  var o = { meta: lgMetaGet_(), head: LG_HEAD, rows: rows };
  cachePutBig_(cache, key, JSON.stringify(o), 1800);
  return { id: '', data: o };
}
ACTIONS.lineage = function (p, a) { return lineage_(p, a); };
READS.lineage = 1;

NOAUTH.lineage = 1;   /* 只讀、只有食譜名稱與日期（頁面本身有閘門） */
