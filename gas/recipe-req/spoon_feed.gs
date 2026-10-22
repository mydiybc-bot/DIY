/* ===== spoon_feed.gs — 🥄 專案推動「食譜用量快照」自動更新（2026-10-22） =====
   經營者：「專案推動打開要讀到最新食譜；加『🔄 更新食譜用量』按鈕（免密碼）＋每天早上自動更新，範圍台灣繁中啟用中食譜，更新前自動備份」。
   為什麼放在這個專案：只有這裡存有門市食譜後台的登入帳號（RB_EMAIL／RB_PASSWORD，管理者在「🔑 密碼管理」設定）。
   做法：登入後台 → /Recipes 清單（台灣、繁體中文、啟用）→ 每支讀「匯出食譜」頁 /steps/recipeprint/{id} →
     只取「平匙規格表裡的食材」（專案推動 spoon_set，裝飾乾料 kind=deco 不算）且單位是 g 的每一處 →
     跟專案推動 spoon_lines 比：完全一樣 → 不寫（不備份）；有變 → 先把分頁複製成 bk_spoon_lines_MMdd-HHmm，再整張換掉。
   first_spoon（第一版平匙數）照舊資料同一支食譜同食材同克數帶過去；每一處另記 use_step／purpose（使用目的：
     領料那一步往後第一個步驟內容提到這個食材的步驟，依標題歸類；專案推動 purpose.gs 有人工修正的仍以那份為準）。
   ⚠ 只讀後台，不改、不刪後台任何東西；只寫專案推動試算表的 spoon_lines 與備份分頁。
   入口：POST action=spoonRefresh（免身分密碼；同時間只跑一個）、排程 spoonDaily（每天 6 點，spoonFeedSetup 建）。 */
var SF_SHEET = '1VE1dsSPk0Bl2dyTWrNo3G_oasD4pCR3ihllNedcOOT4';   /* 專案推動_資料庫 */
var SF_LANG = '42575226-23b1-4ac3-b123-3cba50e3cee8';             /* 繁體中文 */
var SF_RUN = 'spoonfeed:run', SF_HOUR = 6;
var SF_PURPOSE = [
  [/甘納許/, '甘納許'], [/夾餡/, '夾餡'], [/奶凍/, '奶凍'], [/凍|寒天/, '果凍／茶凍'], [/乳酪麵糊/, '乳酪蛋糕體'],
  [/餡/, '內餡'], [/塔皮/, '塔皮'], [/派皮/, '派皮'], [/餅乾底/, '餅乾底'], [/餅糰|麵糰|麵團/, '餅乾麵糰'],
  [/裝飾片|巧克力片|耳朵|裝飾巧克力|裝飾餅乾/, '裝飾配件'], [/淋面/, '淋面'], [/裝飾|塗抹|撒/, '表面裝飾'],
  [/鮮奶油/, '調味鮮奶油'], [/液\d*$/, '調味液'], [/麵糊|蛋白霜|巧克力奶油|蛋糕/, '蛋糕體']
];

function sfPurpose_(title) { for (var i = 0; i < SF_PURPOSE.length; i++) if (SF_PURPOSE[i][0].test(title)) return SF_PURPOSE[i][1]; return ''; }

/* 「匯出食譜」頁 → [{t 步驟標題, c 步驟內容純文字, ing:[{n 品名, q 數量, u 單位}]}] */
function sfSteps_(html) {
  var i = html.indexOf('id="exportContent"'); if (i >= 0) html = html.slice(i);
  var re = /<div\b[^>]*background\s*:\s*yellow[^>]*>([\s\S]*?)<\/div>/gi, m, marks = [];
  while ((m = re.exec(html))) marks.push({ at: m.index, end: re.lastIndex, t: rbText_(m[1]).replace(/^步驟標題[：:]\s*/, '') });
  return marks.map(function (k, j) {
    var seg = html.slice(k.end, j + 1 < marks.length ? marks[j + 1].at : html.length);
    var ix = seg.search(/食材工具[：:]/), body = ix >= 0 ? seg.slice(0, ix) : seg;
    var ing = [], ir = /<div\b[^>]*class="ingred"[^>]*>([\s\S]*?)<\/div>/gi, g;
    while ((g = ir.exec(seg))) {
      var t = rbText_(g[1].replace(/<span\b[\s\S]*?<\/span>/gi, ' ')), mm = t.match(/^(.+?)\s+([\d.]+)\s*(\S+)$/);
      if (mm) ing.push({ n: mm[1].trim(), q: Number(mm[2]), u: mm[3] });
    }
    return { t: k.t, c: rbText_(body).replace(/^步驟內容[：:]\s*/, ''), ing: ing };
  });
}

/* 一批讀「匯出食譜」頁；被導回登入頁 → 重新登入整批再讀一次 */
function sfFetch_(sess, ids) {
  var mk = function () {
    var c = rbJarStr_(sess.jar);
    return ids.map(function (id) { return { url: RB_BASE + '/steps/recipeprint/' + id, method: 'get', muteHttpExceptions: true, followRedirects: false, headers: { 'User-Agent': RB_UA, Cookie: c } }; });
  };
  var rs = UrlFetchApp.fetchAll(mk());
  if (rs.some(function (r) { return r.getResponseCode() === 302; })) {
    rbDropSession_(); sess.jar = rbFreshSession_().jar;
    rs = UrlFetchApp.fetchAll(mk());
  }
  return rs.map(function (r, j) {
    if (r.getResponseCode() !== 200) throw rbErr_('讀不到食譜（' + ids[j] + '，HTTP ' + r.getResponseCode() + '）', 'rbnet');
    return r.getContentText();
  });
}

function sfTable_(ss, name) {
  var sh = ss.getSheetByName(name); if (!sh || sh.getLastRow() < 1) return { sh: sh, head: [], rows: [] };
  var v = sh.getDataRange().getDisplayValues(), head = v[0].map(String);
  return { sh: sh, head: head, rows: v.slice(1).map(function (r) { var o = {}; head.forEach(function (h, i) { o[h] = r[i]; }); return o; }) };
}

function spoonRefresh_(p, auth) {
  var cache = CacheService.getScriptCache();
  if (cache.get(SF_RUN)) throw fail_('食譜用量正在更新中，請 1～2 分鐘後再看', 'busy');
  cache.put(SF_RUN, '1', 300);
  try { return sfRun_(p || {}, (auth && auth.kind) || ''); }
  finally { cache.remove(SF_RUN); }
}

function sfRun_(p, by) {
  var t0 = Date.now(), ss = SpreadsheetApp.openById(SF_SHEET);
  /* ① 要算的食材＝平匙規格表（不含裝飾乾料） */
  var ings = {}, setT = sfTable_(ss, 'spoon_set');
  setT.rows.forEach(function (r) { if (r.ingredient && r.kind !== 'deco') ings[r.ingredient] = 1; });
  if (!Object.keys(ings).length) throw fail_('專案推動的平匙規格表是空的，不更新', 'bad');
  /* ② 後台清單：台灣、繁體中文、啟用 */
  var sess = rbSession_();
  var lr = rbAuthedGet_(sess, '/Recipes?LanguageId=' + SF_LANG + '&Status=true');
  if (lr.code !== 200) throw rbErr_('讀不到後台食譜清單（HTTP ' + lr.code + '）', 'rbnet');
  var recs = rbListRows_(lr.text).filter(function (r) { return r.lang === '繁體中文' && /^Active$/i.test(r.status); });
  if (recs.length < 50) throw fail_('後台清單只讀到 ' + recs.length + ' 支啟用繁中食譜，太少了，先不更新（可能是後台頁面改版）', 'rbpage');
  /* ③ 每支讀「匯出食譜」頁，20 支一批 */
  var lines = [], miss = {};
  for (var b = 0; b < recs.length; b += 20) {
    var part = recs.slice(b, b + 20), htmls = sfFetch_(sess, part.map(function (r) { return r.id; }));
    part.forEach(function (r, j) {
      var steps = sfSteps_(htmls[j]);
      steps.forEach(function (s, k) {
        s.ing.forEach(function (g) {
          if (g.u !== 'g') return;
          if (!ings[g.n]) { if (/粉|糖|鹽|巧克力/.test(g.n)) miss[g.n] = (miss[g.n] || 0) + 1; return; }
          var use = '';
          for (var x = k; x < steps.length && !use; x++) if (steps[x].c.indexOf(g.n) >= 0) use = steps[x].t;
          lines.push({ recipe: r.title, ingredient: g.n, grams: String(g.q), use_step: use || s.t, purpose: sfPurpose_(use || s.t) });
        });
      });
    });
  }
  rbSaveSession_(sess);
  if (!lines.length) throw fail_('一處都沒讀到，先不更新（可能是「匯出食譜」頁改版）', 'rbpage');
  var tRead = Math.round((Date.now() - t0) / 1000);
  /* Google 一次最多跑 6 分鐘：後台太慢、讀完已超過 4 分半 → 這次不寫（資料維持原樣），下次再更新 */
  if (tRead > 270 && !p.dry) throw fail_('門市食譜後台這次很慢（讀了 ' + tRead + ' 秒），為了安全這次不寫入，資料維持原樣；請過幾分鐘再按一次', 'slow');
  /* ④ 跟現在的快照比 */
  var cur = sfTable_(ss, 'spoon_lines'), key = function (o) { return o.recipe + '|' + o.ingredient + '|' + Number(o.grams); };
  var cnt = function (arr) { var o = {}; arr.forEach(function (x) { var k = key(x); o[k] = (o[k] || 0) + 1; }); return o; };
  var a = cnt(cur.rows), bb = cnt(lines), add = [], del = [];
  Object.keys(bb).forEach(function (k) { for (var i = (a[k] || 0); i < bb[k]; i++) add.push(k); });
  Object.keys(a).forEach(function (k) { for (var i = (bb[k] || 0); i < a[k]; i++) del.push(k); });
  var rset = function (arr) { var o = {}; arr.forEach(function (x) { o[x.recipe] = 1; }); return o; };
  var ra = rset(cur.rows), rb = rset(lines);
  var res = { before: cur.rows.length, after: lines.length, recipes_before: Object.keys(ra).length, recipes_after: Object.keys(rb).length, backend: recs.length,
    add_recipes: Object.keys(rb).filter(function (r) { return !ra[r]; }), del_recipes: Object.keys(ra).filter(function (r) { return !rb[r]; }),
    add: add.length, del: del.length, add_lines: add.slice(0, 60), del_lines: del.slice(0, 60),
    other_g: Object.keys(miss).sort(function (x, y) { return miss[y] - miss[x]; }).slice(0, 30).map(function (n) { return n + ' ' + miss[n]; }),
    ings: Object.keys(ings).length, by: by || 'button', t_read: tRead };
  var curHasUse = cur.head.indexOf('use_step') >= 0;
  if (p.dry) { res.dry = true; res.changed = !!(add.length || del.length || !curHasUse); res.took = Math.round((Date.now() - t0) / 1000); return { msg: '試讀（沒寫）', data: res }; }
  if (!add.length && !del.length && curHasUse && !p.force) { res.changed = false; res.took = Math.round((Date.now() - t0) / 1000); return { msg: '沒有變動', data: res }; }
  /* ⑤ 有變 → 先備份，再整張換掉（first_spoon 照舊資料帶過去） */
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var bk = 'bk_spoon_lines_' + Utilities.formatDate(new Date(), TZ, 'MMdd-HHmm');
    if (cur.sh) { var old = ss.getSheetByName(bk); if (old) bk += '_' + Utilities.formatDate(new Date(), TZ, 'ss'); cur.sh.copyTo(ss).setName(bk); }
    var fsQ = {}; cur.rows.forEach(function (r) { (fsQ[key(r)] = fsQ[key(r)] || []).push(r.first_spoon || ''); });
    var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), src = '門市食譜後台：台灣繁中啟用中 ' + recs.length + ' 支（自動讀取）';
    var head = ['recipe', 'ingredient', 'grams', 'first_spoon', 'snap_date', 'source', 'use_step', 'purpose'];
    cur.head.forEach(function (h) { if (h && head.indexOf(h) < 0) head.push(h); });
    var vals = [head].concat(lines.map(function (l) {
      var q = fsQ[key(l)], o = { recipe: l.recipe, ingredient: l.ingredient, grams: l.grams, first_spoon: q && q.length ? q.shift() : '', snap_date: stamp, source: src, use_step: l.use_step, purpose: l.purpose };
      return head.map(function (h) { return o[h] == null ? '' : String(o[h]); });
    }));
    var sh = cur.sh || ss.insertSheet('spoon_lines');
    /* 先蓋上新資料，再清掉多出來的舊列（不先清空：中途被中斷也不會留下空白分頁） */
    sh.getRange(1, 1, vals.length, head.length).setNumberFormat('@').setValues(vals);
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr > vals.length) sh.getRange(vals.length + 1, 1, lr - vals.length, Math.max(lc, head.length)).clearContent();
    if (lc > head.length) sh.getRange(1, head.length + 1, vals.length, lc - head.length).clearContent();
    SpreadsheetApp.flush();
    var back = sh.getRange('A:A').getValues().filter(function (r) { return r[0] !== ''; }).length - 1;
    if (back !== lines.length) throw fail_('寫入後讀回 ' + back + ' 列，跟要寫的 ' + lines.length + ' 列不一樣；備份在分頁 ' + bk, 'verify');
    res.changed = true; res.backup = bk; res.snap_date = stamp;
  } finally { lock.releaseLock(); }
  res.took = Math.round((Date.now() - t0) / 1000);
  return { msg: '更新 ' + res.before + '→' + res.after + ' 項', data: res };
}

/* 排程：每天 6 點 */
function spoonDaily() {
  try { var r = spoonRefresh_({}, { kind: 'trigger' }); log_('排程', 'spoonRefresh', '', true, r.msg); }
  catch (e) { log_('排程', 'spoonRefresh', '', false, errMsg_(e)); }
}
/* 編輯器執行一次：建每天 6 點的排程（已經有就不重建） */
function spoonFeedSetup() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'spoonDaily'; });
  if (!has) ScriptApp.newTrigger('spoonDaily').timeBased().atHour(SF_HOUR).everyDays(1).inTimezone(TZ).create();
  Logger.log('spoonDaily 排程：' + (has ? '已經有' : '已建立（每天 ' + SF_HOUR + ' 點）'));
}
/* 編輯器測試：只讀不寫（看會讀到幾支、幾處、跟現在差多少）；POST 帶 dry:1 也一樣 */
function spoonFeedDry() { Logger.log(JSON.stringify(spoonRefresh_({ dry: 1 }, { kind: 'editor' }).data)); }

ACTIONS.spoonRefresh = function (p, a) { return spoonRefresh_(p, a); };
NOAUTH.spoonRefresh = 1;     /* 免身分密碼：只讀後台、寫專案推動的快照（寫前自動備份，同時只跑一個） */
SELFLOCK.spoonRefresh = 1;   /* 讀 140 多支要 30～60 秒，不佔全系統的鎖（只在寫分頁時短暫拿鎖） */
