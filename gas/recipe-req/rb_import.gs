/* ================= 🍰 匯入自己做食譜系統：流程、權限、紀錄（2026-10-06） =================
   動作（doPost；Code.gs 的 ACTIONS 用包一層的方式掛上，不受檔案載入順序影響）：
   ‧ rbStatus（管理者，只讀）：後台帳號設定狀態（只回遮罩後的帳號，不回密碼）
   ‧ rbSetCred（管理者）：設定後台帳號；先用這組帳密試登入，成功才存（錯的不存，免得之後一直錯把帳號鎖住）
   ‧ rbPreview（主廚／管理者，只讀）：登入後台讀選項與食材清單，比對這次要匯入的步驟、食材、公版圖片影片；不寫任何東西
   ‧ rbImport（主廚／管理者）：分段匯入，每次最多約 35 秒（前端反覆呼叫直到完成）：
       ① 建食譜（停用、不公開；備註寫匯入編號＋沒掛上的食材）② 逐步建步驟＋食材工具＋公版圖片／影片 ③ 讀後台步驟清單核對
   紀錄：申請單試算表 fact_recipe_import（一次匯入一個匯入編號 IMP-…；start 列的 d1～d4 存這次要匯入的完整內容，中斷後可以接續）。
   規則：只限已核准（含已寫入採購、採購寫入失敗）的申請單；同一張單匯入完成後不能再匯入，管理者可「重新匯入」（會在後台再建一支新的）。
   回覆遺失／中斷：建食譜前先用匯入編號搜後台備註（找到就沿用，不重建）；建步驟前先讀後台步驟清單（已建的不重建）。 */
var RB_ST = { '已核准': 1, '已核准（採購寫入失敗）': 1, '已寫入採購': 1 };
var RB_MEDIA_SHEET = '1vUqRP8GhSlwSHnatZDUl51IsjI-Xzv7u5krdwVIriPI';   /* 食譜素材索引（2026-10-05 建） */
var RB_MEDIA_TAB = '公版候選';
var RB_BUDGET_MS = 35000;
var RB_MAX_STEPS = 150, RB_MAX_HTML = 30000, RB_MAX_ING = 80;
var RB_BRAND_STORES = { '自己做': ['自己做'], '吳寶春自己做': ['吳寶春自己做'], '自己做＆吳寶春自己做': ['自己做', '吳寶春自己做'] };
var RB_D_COLS = ['d1', 'd2', 'd3', 'd4'];

function rbMask_(email) { var s = String(email || ''), at = s.indexOf('@'); return at > 0 ? s.slice(0, Math.min(3, at)) + '***' + s.slice(at) : (s ? s.slice(0, 2) + '***' : ''); }
function rbTitleNorm_(s) { return rbNorm_(s).replace(/\s+/g, '').toLowerCase(); }
function rbLink_(bid) { return bid ? RB_BASE + '/Recipes/Edit/' + bid : ''; }

/* ---------- 後台帳號（只限管理者） ---------- */
function rbStatus_(p, auth) {
  needAdmin_(auth);
  var P = rbProps_(), em = P.getProperty('RB_EMAIL'), at = Number(P.getProperty('RB_OK_AT') || 0);
  return { id: '', data: { set: !!(em && P.getProperty('RB_PASSWORD')), email: rbMask_(em), bad: P.getProperty('RB_BAD') === '1',
    okAt: at ? Utilities.formatDate(new Date(at), TZ, 'yyyy-MM-dd HH:mm') : '' } };
}
function rbSetCred_(p, auth) {
  needAdmin_(auth);
  var email = String(p.rb_email || '').trim(), pass = String(p.rb_password || '');
  if (!/^[^\s@]+@[^\s@]+$/.test(email) || email.length > 120) throw fail_('後台帳號要填登入用的 Email', 'input');
  if (!pass || pass.length > 200) throw fail_('請填後台密碼', 'input');
  var sess;
  try { sess = rbLoginRaw_(email, pass); }
  catch (e) { throw fail_(errMsg_(e) + '（這組帳號密碼沒有存）', (e && e.code) || 'rbnet'); }
  var P = rbProps_();
  P.setProperty('RB_EMAIL', email); P.setProperty('RB_PASSWORD', pass); P.deleteProperty('RB_BAD');
  rbSaveSession_(sess);
  return { id: '', data: { ok: true, email: rbMask_(email) }, msg: '食譜後台帳號已更新（' + rbMask_(email) + '），試登入成功' };
}

/* ---------- 申請單檢查 ---------- */
function rbCheck_(auth, rid) {
  if (!canPush_(auth)) throw fail_('匯入自己做食譜系統要「主廚」或「管理者」密碼', 'perm');
  if (!rid) throw fail_('缺申請單號', 'input');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var st = str_(row.status);
  if (!RB_ST[st]) throw fail_('這張單還沒核准（目前「' + st + '」），核准後才能匯入自己做食譜系統', 'state');
  var head = {};
  try { head = (JSON.parse(readPayload_(t, row._row)) || {}).head || {}; } catch (e) { throw fail_('這張單的內容讀不出來，請通知 Claude', 'payload'); }
  return { row: row, status: st, head: head };
}
/* 前端送來的食譜（解析 .doc 後的結構）→ 檢查、截長度、內容安全處理 */
function rbCleanRecipe_(rc) {
  if (!rc || typeof rc !== 'object' || !Array.isArray(rc.steps)) throw fail_('沒有收到食譜步驟（請選這支甜點的食譜檔 .doc）', 'input');
  if (!rc.steps.length) throw fail_('食譜檔裡沒有步驟', 'input');
  if (rc.steps.length > RB_MAX_STEPS) throw fail_('步驟太多（' + rc.steps.length + ' 步，上限 ' + RB_MAX_STEPS + '）', 'input');
  var info = rc.info || {};
  var steps = rc.steps.map(function (s, i) {
    s = s || {};
    var t = rbTrim_(s.t).slice(0, 100);
    if (!t) throw fail_('第 ' + (i + 1) + ' 步沒有標題', 'input');
    var h = String(s.h || '');
    if (h.length > RB_MAX_HTML) throw fail_('第 ' + (i + 1) + ' 步內容太長（' + h.length + ' 字）', 'input');
    var ing = (Array.isArray(s.ing) ? s.ing : []).slice(0, RB_MAX_ING).map(function (g) {
      g = Array.isArray(g) ? g : [];
      var q = Number(g[1]);
      return [rbTrim_(g[0]).slice(0, 100), (isFinite(q) && q > 0 && q < 1e6) ? q : 0, rbTrim_(g[2]).slice(0, 20), rbTrim_(g[3]).slice(0, 30),
        (g[4] === 'container1' || g[4] === 'container2') ? g[4] : '', rbTrim_(g[5]).slice(0, 100)];
    }).filter(function (g) { return g[0] && g[1] > 0 && g[2]; });
    return { t: t, h: rbSanitize_(h), ing: ing };
  });
  return { name: rbTrim_(rc.name).slice(0, 100), file: rbTrim_(rc.file).slice(0, 200),
    info: { time: rbTrim_(info.time).slice(0, 30), size: rbTrim_(info.size).slice(0, 100), desc: String(info.desc || '').slice(0, 3000), preserve: String(info.preserve || '').slice(0, 2000) },
    steps: steps };
}
/* 後台基本資料預設值：名稱＝正式名稱（沒有用暫定名稱、再沒有用食譜檔品名）；售價＝定價；成本＝申請單成本；店別＝品牌別 */
function rbDefaults_(head, rc) {
  head = head || {};
  var info = rc.info || {};
  var title = rbTrim_(head.fname) || rbTrim_(head.name) || rbTrim_(rc.name);
  var hr = (String(info.time || '').match(/(\d+(?:\.\d+)?)/) || String(head.time || '').match(/(\d+(?:\.\d+)?)/) || [])[1] || '';
  return {
    title: title.slice(0, 100), price: Number(head.price) || 0, cost: Math.round((Number(head.cost) || 0) * 100) / 100,
    brand: rbTrim_(head.brand), stores: RB_BRAND_STORES[rbTrim_(head.brand)] || RB_BRAND_STORES['自己做＆吳寶春自己做'],
    prepHr: hr, size: rbTrim_(info.size || head.spec || '').slice(0, 100), preserve: String(info.preserve || ''), desc: String(info.desc || '')
  };
}
function rbMatchStep_(ings, s) {
  var ok = [], miss = [], notes = [];
  (s.ing || []).forEach(function (g) {
    var label = g[0] + ' ' + g[1] + ' ' + g[2];
    var r = rbMatch_(ings, { name: g[0], unit: g[2], cate: g[3] });
    if (r.hit) ok.push({ id: r.hit.id, amount: g[1], unit: r.hit.unit, cont: g[4] });
    else miss.push(label + (r.miss === 'unit' ? '（後台這個品項的單位是 ' + r.units.join('／') + '）' : '（後台食材清單沒有）'));
    if (g[5]) notes.push(label + '（' + g[5] + '）');
  });
  return { ok: ok, miss: miss, notes: notes };
}

/* ---------- 公版圖片／影片（食譜素材索引「公版候選」，選用打勾的） ---------- */
function rbMediaPool_() {
  var cache = CacheService.getScriptCache(), key = 'rbmedia:v1', s = cacheGetBig_(cache, key);
  if (s) return JSON.parse(s);
  var sh = SpreadsheetApp.openById(RB_MEDIA_SHEET).getSheetByName(RB_MEDIA_TAB);
  if (!sh) throw fail_('食譜素材索引沒有「' + RB_MEDIA_TAB + '」分頁');
  var v = sh.getDataRange().getValues(), H = v[0].map(function (x) { return String(x).trim(); }), c = function (n) { return H.indexOf(n); };
  var iSel = c('選用'), iT = c('步驟標題'), iAlt = c('其他步驟標題'), iType = c('類型'), iN = c('用到的食譜數'), iMb = c('檔案大小MB'), iId = c('素材編號'), iUrl = c('連結');
  if ([iSel, iT, iType, iUrl].some(function (x) { return x < 0; })) throw fail_('食譜素材索引「' + RB_MEDIA_TAB + '」表頭不對');
  var out = [];
  for (var r = 1; r < v.length; r++) {
    var row = v[r];
    if (row[iSel] !== true) continue;
    var type = String(row[iType]).trim();
    if (type !== '影片' && type !== '圖片') continue;   /* 「圖片（說明內嵌）」是寫在步驟文字裡的圖，不當步驟圖片 */
    var url = String(row[iUrl]).trim(), mb = Number(row[iMb]) || 0;
    if (url.indexOf(RB_MEDIA_PREFIX) !== 0 || mb > 5) continue;
    var alt = iAlt >= 0 ? String(row[iAlt]).split('／').map(rbTitleNorm_).filter(String) : [];
    out.push({ k: rbTitleNorm_(row[iT]), alt: alt, title: String(row[iT]).trim(), type: type, recipes: Number(row[iN]) || 0, mb: mb, id: iId >= 0 ? String(row[iId]) : '', url: url });
  }
  cachePutBig_(cache, key, JSON.stringify(out), 600);
  return out;
}
/* 步驟標題相同（或在該素材的「其他步驟標題」裡）的公版素材，被最多食譜用的排前面，最多 5 個 */
function rbMediaFor_(pool, title) {
  var k = rbTitleNorm_(title);
  return pool.filter(function (c) { return c.k === k || c.alt.indexOf(k) >= 0; })
    .sort(function (a, b) { return b.recipes - a.recipes; }).slice(0, 5)
    .map(function (c) { return { url: c.url, type: c.type, mb: c.mb, recipes: c.recipes, id: c.id, title: c.title }; });
}
/* 讀食材清單要借一支現有食譜的「新增步驟」頁：先用上次記住的，打不開就從後台清單挑一支啟用中的 */
function rbProbeIngs_(sess, list) {
  var P = rbProps_(), rid = P.getProperty('RB_PROBE_RID');
  if (rid) { try { return rbStepPage_(sess, rid); } catch (e) { if (e && (e.code === 'rbpage' || e.code === 'rbbad' || e.code === 'rbnocred')) throw e; } }
  if (!list) list = rbListRows_(rbAuthedGet_(sess, '/Recipes').text);
  var cand = list.filter(function (r) { return r.status === 'Active'; }).concat(list).slice(0, 5), last = null;
  for (var i = 0; i < cand.length; i++) {
    try { var sp = rbStepPage_(sess, cand[i].id); P.setProperty('RB_PROBE_RID', cand[i].id); return sp; }
    catch (e) { last = e; if (e && (e.code === 'rbpage' || e.code === 'rbbad' || e.code === 'rbnocred')) throw e; }
  }
  throw last || rbErr_('後台沒有可以讀食材清單的食譜', 'rbpage');
}

/* ---------- 預覽（只讀） ---------- */
function rbPreview_(p, auth) {
  var rid = str_(p.req_id).trim(), q = rbCheck_(auth, rid);
  var rc = rbCleanRecipe_(p.recipe), d = rbDefaults_(q.head, rc);
  var out = { defaults: d, categories: [], steps: [], warn: [], dup: [], prior: rbInfo_(rid), docName: rc.name, file: rc.file };
  var sess = rbSession_();
  var pg = rbAuthedGet_(sess, '/Recipes/Create');
  if (pg.code !== 200) throw rbErr_('食譜後台「新增食譜」頁打不開（HTTP ' + pg.code + '）', 'rbnet');
  var meta = rbCreateMeta_(pg.text);
  out.categories = meta.cates;
  var nost = d.stores.filter(function (n) { return !meta.stores.some(function (s) { return s.name === n; }); });
  if (nost.length) out.warn.push('後台「分店類別」找不到：' + nost.join('、') + '（匯入後請食譜負責夥伴自己選）');
  var list = null;
  try { list = rbListRows_(rbAuthedGet_(sess, '/Recipes').text); }
  catch (e) { out.warn.push('讀不到後台食譜清單，沒有檢查同名食譜（' + errMsg_(e) + '）'); }
  if (list) out.dup = list.filter(function (r) { return rbTitleNorm_(r.title) === rbTitleNorm_(d.title); }).slice(0, 10);
  var sp = rbProbeIngs_(sess, list);
  var pool = [];
  try { pool = rbMediaPool_(); } catch (e) { out.warn.push('讀不到「食譜素材索引」公版候選，沒有建議圖片／影片（' + errMsg_(e) + '）'); }
  rc.steps.forEach(function (s, i) {
    var m = rbMatchStep_(sp.ings, s);
    out.steps.push({ n: i + 1, title: s.t, ing_n: s.ing.length, ok: m.ok.length, miss: m.miss, notes: m.notes, media: rbMediaFor_(pool, s.t) });
  });
  if (rc.name && d.title && rbTitleNorm_(rc.name).indexOf(rbTitleNorm_(d.title)) < 0 && rbTitleNorm_(d.title).indexOf(rbTitleNorm_(rc.name)) < 0)
    out.warn.push('食譜檔的品名「' + rc.name + '」和申請單名稱「' + d.title + '」不一樣：「確認品項」頁等步驟文字照食譜檔，匯入後請食譜負責夥伴檢查');
  rbSaveSession_(sess);
  return { id: rid, data: out };
}

/* ---------- 紀錄分頁 fact_recipe_import ---------- */
var RB_LOG_N = 9;   /* ts～by 九欄；d1～d4 另外讀（start 列存整份內容，很大） */
function rbLog_(rid, imp, action, bid, step, ok, msg, auth, json) {
  var segs = [];
  if (json) {
    try { segs = split_(json); }
    catch (e) { throw fail_('這支食譜內容太大，存不進匯入紀錄（' + json.length + ' 字，上限 ' + (SEG_MAX * SEG_N) + ' 字）；請精簡步驟內容', 'toobig'); }
  }
  var lock = LockService.getScriptLock(), got = false;
  try {
    got = lock.tryLock(20000);
    var sh = ss_().getSheetByName('fact_recipe_import');
    if (!sh) { ensureTab_(ss_(), 'fact_recipe_import'); sh = ss_().getSheetByName('fact_recipe_import'); }
    var r = sh.getLastRow() + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
    var vals = [now_(), rid, imp, action, bid || '', step == null ? '' : String(step), ok, String(msg || '').slice(0, 500), (auth && (auth.role || auth.name)) || '']
      .concat(RB_D_COLS.map(function (c, i) { return segs[i] || ''; }));
    sh.getRange(r, 1, 1, vals.length).setNumberFormat('@').setValues([vals.map(function (v) { return safe_(String(v)); })]);
  } finally { if (got) lock.releaseLock(); }
}
function rbState_(rid) {
  var out = { imps: {}, order: [], open: null, done: null, sh: null };
  var sh = ss_().getSheetByName('fact_recipe_import');
  if (!sh) return out;
  out.sh = sh;
  var last = sh.getLastRow();
  if (last < 2) return out;
  var H = TABS.fact_recipe_import, v = sh.getRange(2, 1, last - 1, RB_LOG_N).getValues();
  v.forEach(function (r, i) {
    if (str_(r[1]) !== rid) return;
    var o = { _row: i + 2 };
    for (var j = 0; j < RB_LOG_N; j++) o[H[j]] = str_(r[j]);
    var im = out.imps[o.imp_id];
    if (!im) { im = out.imps[o.imp_id] = { imp_id: o.imp_id, start_row: 0, backend_id: '', done: false, abandoned: false, ts: o.ts, by: o.by, title: '', total: 0, rows: [] }; out.order.push(o.imp_id); }
    im.rows.push(o);
    if (o.action === 'start') { im.start_row = o._row; var mm = String(o.msg).match(/^(.*)｜(\d+) 步$/); if (mm) { im.title = mm[1]; im.total = +mm[2]; } }
    if (o.action === 'recipe' && o.backend_id) im.backend_id = o.backend_id;
    if (o.action === 'done') { im.done = true; im.done_ts = o.ts; }
    if (o.action === 'abandon') im.abandoned = true;
  });
  out.order.forEach(function (k) { var im = out.imps[k]; if (im.done) out.done = im; else if (!im.abandoned && im.start_row) out.open = im; });
  return out;
}
function rbReadD_(sh, row) { return sh.getRange(row, RB_LOG_N + 1, 1, RB_D_COLS.length).getValues()[0].map(str_).join(''); }
function rbJob_(st, im) {
  if (!im.start_row) throw fail_('匯入紀錄 ' + im.imp_id + ' 不完整', 'state');
  var job;
  try { job = JSON.parse(rbReadD_(st.sh, im.start_row)); } catch (e) { throw fail_('匯入紀錄 ' + im.imp_id + ' 讀不出來，請通知 Claude', 'state'); }
  job.backend_id = im.backend_id || '';
  return job;
}
/* 給前端看的匯入狀態（getReq／預覽都會帶）：最近一次匯入＋之前的 */
function rbInfo_(rid) {
  var st = rbState_(rid);
  if (!st.order.length) return null;
  var k = st.order[st.order.length - 1], im = st.imps[k];
  var made = 0, media = 0, miss = [], notes = [], mw = [], fail = '';
  im.rows.forEach(function (r) {
    if (r.action === 'steps') {
      var m = String(r.step).match(/(\d+)$/); if (m) made = Math.max(made, +m[1]);
      var d = null; try { d = JSON.parse(rbReadD_(st.sh, r._row)); } catch (e) { }
      (d || []).forEach(function (x) {
        if (x.media) media++;
        if (x.miss && x.miss.length) miss.push({ n: x.n, t: x.t, items: x.miss });
        if (x.notes && x.notes.length) notes.push({ n: x.n, items: x.notes });
        if (x.mwarn) mw.push('第 ' + x.n + ' 步：' + x.mwarn);
      });
    }
    if (r.action === 'fail') fail = r.msg + '（' + r.ts + '）';
    if (r.action === 'steps' || r.action === 'done') fail = '';
  });
  return {
    imp_id: im.imp_id, title: im.title, total: im.total, made: made, backend_id: im.backend_id, link: rbLink_(im.backend_id),
    done: im.done, done_ts: im.done_ts || '', ts: im.ts, by: im.by, abandoned: im.abandoned, media: media, miss: miss, notes: notes, mwarn: mw, fail: fail,
    earlier: st.order.slice(0, -1).map(function (x) { var e = st.imps[x]; return { imp_id: e.imp_id, backend_id: e.backend_id, link: rbLink_(e.backend_id), done: e.done, ts: e.ts }; })
  };
}

/* ---------- 匯入（分段） ---------- */
function rbNote_(rid, job, mm) {
  var L = ['【食譜系統匯入】新品申請 ' + rid + '｜匯入編號 ' + job.imp_id + '｜' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') + '｜' + (job.by || ''),
    '食譜檔：' + (job.file || '（未記錄）') + (job.docName ? '（' + job.docName + '）' : '')];
  var miss = [], notes = [];
  mm.forEach(function (m, i) {
    m.miss.forEach(function (x) { miss.push('第 ' + (i + 1) + ' 步：' + x); });
    m.notes.forEach(function (x) { notes.push('第 ' + (i + 1) + ' 步：' + x); });
  });
  if (miss.length) L.push('沒掛上的食材工具（請食譜負責夥伴補上）：', miss.join('\n'));
  if (notes.length) L.push('食譜檔的附註沒帶進後台（後台只記數量和單位）：', notes.join('\n'));
  L.push('匯入時是「停用」：檢查調整好再到這裡勾「啟用」。');
  return L.join('\n').slice(0, 4000);
}
function rbNewJob_(rid, q, p, auth) {
  var rc = rbCleanRecipe_(p.recipe), d = rbDefaults_(q.head, rc), o = p.opts || {};
  var title = rbTrim_(o.title || d.title).slice(0, 100);
  if (!title) throw fail_('食譜名稱是空的', 'input');
  var cat = str_(o.category_id).trim();
  if (cat && !/^[0-9a-f-]{36}$/i.test(cat)) throw fail_('類別不對，請重新檢查', 'input');
  var media = Array.isArray(o.media) ? o.media : [];
  var steps = rc.steps.map(function (s, i) {
    var u = str_(media[i]).trim();
    if (u && u.indexOf(RB_MEDIA_PREFIX) !== 0) throw fail_('第 ' + (i + 1) + ' 步的圖片／影片不是公版素材', 'input');
    return { t: s.t, h: s.h, ing: s.ing, m: u };
  });
  var imp = 'IMP-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMddHHmmss') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4);
  var job = { imp_id: imp, rid: rid, title: title, category_id: cat, price: d.price, cost: d.cost, stores: d.stores, prepHr: d.prepHr, size: d.size,
    preserve: d.preserve, desc: d.desc, file: rc.file, docName: rc.name, by: auth.role, steps: steps };
  rbLog_(rid, imp, 'start', '', '', 'Y', title + '｜' + steps.length + ' 步', auth, JSON.stringify(job));
  job.backend_id = '';
  return job;
}
function rbImport_(p, auth) {
  var t0 = Date.now(), rid = str_(p.req_id).trim(), q = rbCheck_(auth, rid);
  var cache = CacheService.getScriptCache(), ck = 'rbimp:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(rid).getBytes());
  if (cache.get(ck)) throw fail_('這張單正在匯入中（可能另一個畫面也按了），請等 1 分鐘再按「繼續匯入」', 'busy');
  cache.put(ck, '1', 120);
  try {
    var st = rbState_(rid), job, im;
    if (p.imp_id) {
      im = st.imps[str_(p.imp_id)];
      if (!im) throw fail_('找不到匯入編號 ' + p.imp_id, 'notfound');
      if (im.abandoned) throw fail_('這次匯入已作廢，請重新按「📥 匯入自己做食譜系統」', 'state');
      if (im.done) return { id: rid, data: { imp_id: im.imp_id, backend_id: im.backend_id, link: rbLink_(im.backend_id), done: true, next: im.total, total: im.total }, msg: '已完成' };
      job = rbJob_(st, im);
    } else {
      var open = st.open;
      if (open && !open.backend_id) {
        /* 上次停在建食譜之前，或剛好建好但回覆遺失 → 先用匯入編號在後台備註找 */
        var found = rbFindByNote_(rbSession_(), open.imp_id);
        if (found.length) { open.backend_id = found[0].id; rbLog_(rid, open.imp_id, 'recipe', found[0].id, '', 'Y', '找到上次建好的食譜（回覆遺失，沿用）', auth); }
        else { rbLog_(rid, open.imp_id, 'abandon', '', '', 'Y', '上次沒建成食譜，改用新的匯入', auth); open = null; }
      }
      if (open) throw fail_('上次的匯入還沒完成（' + open.imp_id + '），請按「▶ 繼續匯入」', 'resume', { imp_id: open.imp_id });
      if (st.done && !(auth.role === ADMIN && p.reimport === true))
        throw fail_('這張單已經匯入過了；要在後台再建一支新的，請用管理者按「重新匯入」', 'already', { backend_id: st.done.backend_id, link: rbLink_(st.done.backend_id) });
      job = rbNewJob_(rid, q, p, auth);
    }
    var res = rbRun_(rid, job, auth, t0);
    return { id: rid, data: res, msg: res.done ? '匯入完成 ' + res.total + ' 步' : '匯入到第 ' + res.next + '／' + res.total + ' 步' };
  } finally { cache.remove(ck); }
}
function rbRun_(rid, job, auth, t0) {
  var sess = rbSession_(), N = job.steps.length;
  var relogin = function () { rbDropSession_(); sess = rbFreshSession_(); };
  if (!job.backend_id) {
    var found = rbFindByNote_(sess, job.imp_id);   /* 回覆遺失保險：備註已有這個匯入編號＝建過了 */
    if (found.length) job.backend_id = found[0].id;
    else {
      var probe = rbProbeIngs_(sess, null);
      var mm = job.steps.map(function (s) { return rbMatchStep_(probe.ings, s); });
      var r = { title: job.title, note: rbNote_(rid, job, mm), itemList: job.category_id, storeNames: job.stores, price: job.price, cost: job.cost,
        prepHr: job.prepHr, size: job.size, desc: job.desc, preserve: job.preserve };
      try {
        try { job.backend_id = rbCreateRecipe_(sess, r); }
        catch (e) { if (e && e.code === 'rbrelogin') { relogin(); job.backend_id = rbCreateRecipe_(sess, r); } else throw e; }
      } catch (e2) {
        rbLog_(rid, job.imp_id, 'fail', '', '', 'N', '建食譜：' + errMsg_(e2), auth);
        throw fail_('在後台建食譜沒有成功：' + errMsg_(e2), (e2 && e2.code) || 'rbnet', { imp_id: job.imp_id, next: 0, total: N });
      }
    }
    rbLog_(rid, job.imp_id, 'recipe', job.backend_id, '', 'Y', '後台食譜已建好（停用、不公開）', auth);
  }
  var bid = job.backend_id, idx = rbStepIndex_(sess, bid), k = idx.length;
  var badAt = function (lst) {
    for (var i = 0; i < lst.length; i++) if (i >= N || rbTitleNorm_(lst[i].title) !== rbTitleNorm_(job.steps[i].t)) return i + 1;
    return 0;
  };
  var b0 = badAt(idx);
  if (b0) {
    rbLog_(rid, job.imp_id, 'fail', bid, '', 'N', '後台步驟和要匯入的不一致（第 ' + b0 + ' 步）', auth);
    throw fail_('後台這支食譜的第 ' + b0 + ' 步和要匯入的不一樣（可能有人在後台改過），先停下來。請食譜負責夥伴檢查後台；要重來請管理者按「重新匯入」', 'mismatch', { imp_id: job.imp_id, next: k, total: N, link: rbLink_(bid) });
  }
  var made = [], stopErr = null;
  if (k < N && Date.now() - t0 < RB_BUDGET_MS) {
    var sp = rbStepPage_(sess, bid);
    for (var j = k; j < N; j++) {
      if (Date.now() - t0 > RB_BUDGET_MS) break;
      var s = job.steps[j], m = rbMatchStep_(sp.ings, s), md = null, mwarn = '';
      if (s.m) { try { md = rbFetchMedia_(s.m); } catch (e) { mwarn = errMsg_(e); } }
      var one = { title: s.t, html: rbSanitize_(s.h), ings: m.ok, media: md ? md.blob : null };
      try {
        try { rbCreateStep_(sess, bid, sp.token, one); }
        catch (e) { if (e && e.code === 'rbrelogin') { relogin(); sp = rbStepPage_(sess, bid); rbCreateStep_(sess, bid, sp.token, one); } else throw e; }
      } catch (e3) { stopErr = e3; break; }
      made.push({ n: j + 1, t: s.t, ing: m.ok.length, media: md ? (md.type.indexOf('video') === 0 ? '影片' : '圖片') : '', mwarn: mwarn, miss: m.miss, notes: m.notes });
    }
  }
  var idx2 = rbStepIndex_(sess, bid), k2 = idx2.length, b2 = badAt(idx2);
  if (k2 > k) rbLog_(rid, job.imp_id, 'steps', bid, (k + 1) + '-' + k2, 'Y', '建好第 ' + (k + 1) + '～' + k2 + ' 步' + (k2 < k + made.length ? '（核對時少了 ' + (k + made.length - k2) + ' 步）' : ''), auth, JSON.stringify(made));
  rbSaveSession_(sess);
  if (b2 || k2 < k + made.length) {
    rbLog_(rid, job.imp_id, 'fail', bid, '', 'N', '核對後台步驟清單不一致', auth);
    throw fail_('建完步驟後核對後台步驟清單不一致（第 ' + (b2 || (k2 + 1)) + ' 步），先停下來；請食譜負責夥伴檢查後台，或通知 Claude', 'mismatch', { imp_id: job.imp_id, next: k2, total: N, link: rbLink_(bid) });
  }
  if (stopErr && k2 < N) {
    rbLog_(rid, job.imp_id, 'fail', bid, String(k2 + 1), 'N', '第 ' + (k2 + 1) + ' 步：' + errMsg_(stopErr), auth);
    throw fail_('第 ' + (k2 + 1) + ' 步沒有建成：' + errMsg_(stopErr), (stopErr && stopErr.code) || 'rbnet', { imp_id: job.imp_id, next: k2, total: N, link: rbLink_(bid) });
  }
  var done = k2 >= N;
  if (done) rbLog_(rid, job.imp_id, 'done', bid, String(N), 'Y', '匯入完成：' + N + ' 步', auth);
  return { imp_id: job.imp_id, backend_id: bid, link: rbLink_(bid), done: done, next: k2, total: N };
}

/* 授權用（在 Apps Script 編輯器執行一次，跳出 Google 授權：連線到外部服務）；只讀後台公開的登入頁，回 HTTP 狀態碼，順便確認 Google 連得到後台 */
function rbAuthorize() {
  var r = UrlFetchApp.fetch(RB_BASE + '/Identity/Account/Login', { muteHttpExceptions: true, followRedirects: false });
  Logger.log('食譜後台登入頁 HTTP ' + r.getResponseCode());
  return r.getResponseCode();
}
