/* ================= 🌐 食譜翻譯 → 📤 上傳到自己做食譜系統（2026-10-22） =================
   經營者 10/22：「食譜翻譯完成後，我要加一個按鍵，讓我可以直接上傳到自己做食譜系統的後台」。
   做法（跟門市夥伴平常做翻譯版一樣）：用後台「編輯食譜 → Copy 複製食譜」把原文複製到目標語言（後台自己複製步驟、食材、圖片影片，
   圖片影片指向原本的檔，不另外上傳），再把複製出來的那一支換成翻譯稿：
     ① 複製：POST /Recipes/Copy（Replaceable、RecipeId、LanguageId，欄位照「編輯食譜」頁的複製表單）。後台不回新編號 →
        複製前後各讀一次「名稱＝原文名稱、語言＝目標語言」的清單，多出來的那一支就是（多出不只一支 → 停下來，不猜）。
     ② 基本資料：名稱、尺寸、產品特色、保存方式換成翻譯稿；**停用、不公開**；店別換成目標地區的（澳門＝加盟店(澳門)、越南＝越南區、
        菲律賓＝代理店(The Philippines)；英文＝台灣，照原本的）；類別照複製來的；上架／下架期間空白或 0001 年 → 補 2023～2030（不然平板打不開）；
        備註寫上傳編號＋原文（原文的備註接在後面）。
     ③ 每一步就地改：標題、內容換成翻譯稿；圖片影片留原本的；食材列重新掛一次（同一個品項，數量照原文，單位換成目標語言的單位、
        容器照原文）——食材名稱後台會自己依食譜語言顯示。某一步的食材有一項對不到 → 那一步的食材列不動（照複製來的），回報給畫面。
     ④ 核對：後台步驟數、每步標題＝翻譯稿 → 完成。
   ⚠ 只新增一支食譜（停用、不公開）；原文食譜一格都不改。同一支原文同一個語言同時只能跑一個；做完再上傳一次＝再建一支新的（要畫面確認）。
   ⚠ 中途斷線：同一個上傳編號再送一次就接著做（複製那步用清單比對，不會複製兩支）。
   紀錄：fact_recipe_import（req_id＝TR-…；start 列 step 欄＝「語言:原文編號」，d1～d4＝整份翻譯稿）。 */
var TR_UP_LANG = { mo: 'ee184749-9860-48f4-a76f-14c80dd8fdb2', ph: '73743a68-2ef5-4e71-be40-3bebd8d97712', vn: '68d41688-4c57-44ed-b35d-0fcdbd1ee292',
  me: '6643fe9e-27fb-436a-896d-a745a7bf91a3', en: 'b78ccd25-866b-4c8d-8e71-aa854a4c48f6' };
var TR_UP_LABEL = { mo: '澳門中文', ph: '菲律賓英文', vn: '越南文', me: '澳門英文', en: '英文' };
var TR_UP_STORE = { mo: '加盟店(澳門)', me: '加盟店(澳門)', vn: '越南區', ph: '代理店(The Philippines)', en: '' };   /* 空白＝照複製來的 */
var TR_UP_RE = /^TR-\d{8}-\d{6}-[0-9a-f]{4}$/;

function trUpNeed_(auth) { if (!canPush_(auth)) throw fail_('上傳到自己做食譜系統要「主廚」或「管理者」密碼', 'perm'); }
function trUpId_() { return 'TR-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 4); }

/* 這支原文、這個語言以前上傳過的（做完的）＋沒做完的；回 [{tr_id, done, backend_id, ts, by}] */
function trUpPast_(src, lang) {
  var sh = ss_().getSheetByName('fact_recipe_import'), key = lang + ':' + src, ids = {}, out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, RB_LOG_N).getValues().forEach(function (r) {
    var rid = str_(r[1]); if (rid.indexOf('TR-') !== 0) return;
    var a = str_(r[3]), o = ids[rid];
    if (a === 'start' && str_(r[5]) === key && !o) { o = ids[rid] = { tr_id: rid, done: false, abandoned: false, backend_id: '', ts: str_(r[0]), by: str_(r[8]) }; out.push(o); }
    if (!o) return;
    if (str_(r[4])) o.backend_id = str_(r[4]);
    if (a === 'done') o.done = true;
    if (a === 'abandon') o.abandoned = true;
  });
  out.forEach(function (o) { o.link = rbLink_(o.backend_id); });
  return out;
}

/* 畫面送來的翻譯稿 → 工作內容（存在 start 列，中斷可接續） */
function trUpJob_(p, auth) {
  var src = str_(p.src).trim().toLowerCase(), lang = str_(p.lang).trim(), t = p.tr || {};
  if (!TR_ID.test(src)) throw fail_('原文食譜編號不對', 'input');
  if (!TR_UP_LANG[lang]) throw fail_('不支援的語言：' + lang, 'input');
  var title = rbTrim_(t.name).slice(0, 100);
  if (!title) throw fail_('翻譯稿的食譜名稱是空的', 'input');
  var st = Array.isArray(p.st) ? p.st.map(function (x) { return rbTrim_(x).slice(0, 200); }) : null;
  if (!Array.isArray(t.steps) || !t.steps.length) throw fail_('翻譯稿沒有步驟', 'input');
  if (t.steps.length > RB_MAX_STEPS) throw fail_('步驟太多（' + t.steps.length + ' 步）', 'input');
  if (!st || st.length !== t.steps.length) throw fail_('原文步驟數跟翻譯稿對不上，請重新翻譯', 'input');
  var steps = t.steps.map(function (s, i) {
    s = s || {};
    var tt = rbTrim_(s.t).slice(0, 100);
    if (!tt) throw fail_('翻譯稿第 ' + (i + 1) + ' 步沒有標題', 'input');
    var h = String(s.h || '');
    if (h.length > RB_MAX_HTML) throw fail_('翻譯稿第 ' + (i + 1) + ' 步內容太長（' + h.length + ' 字）', 'input');
    var ing = (Array.isArray(s.ing) ? s.ing : []).slice(0, RB_MAX_ING).map(function (g) {
      g = Array.isArray(g) ? g : [];
      var q = Number(g[1]);
      return [rbTrim_(g[0]).slice(0, 100), (isFinite(q) && q > 0 && q < 1e6) ? q : 0, rbTrim_(g[2]).slice(0, 40),
        (g[3] === 'container1' || g[3] === 'container2') ? g[3] : '', rbTrim_(g[4]).slice(0, 40)];
    });
    return { t: tt, h: rbSanitize_(h), ing: ing };
  });
  return { kind: 'translate', src: src, lang: lang, title: title, size: rbTrim_(t.size).slice(0, 100), desc: String(t.desc || '').slice(0, 3000),
    preserve: String(t.preserve || '').slice(0, 2000), st: st, steps: steps, by: (auth && (auth.role || auth.name)) || '' };
}

/* 讀後台清單（名稱包含＋語言）→ 食譜編號 */
function trUpListIds_(sess, title, langId, groupId) {   /* ⚠ 不帶地區＝後台只列台灣（10/22 實測） */
  var r = rbAuthedGet_(sess, '/Recipes?Title=' + encodeURIComponent(title) + '&LanguageId=' + encodeURIComponent(langId) + '&GroupId=' + encodeURIComponent(groupId || ''));
  if (r.code !== 200) throw rbErr_('讀不到後台食譜清單（HTTP ' + r.code + '）', 'rbnet');
  return rbListRows_(r.text).map(function (x) { return x.id; });
}
/* 「編輯食譜」頁的複製表單 */
function trUpCopyForm_(html) {
  var f = rbFormHtml_(html, /^\/Recipes\/Copy/i);
  if (!f) throw rbErr_('食譜後台「編輯食譜」頁找不到「複製食譜」表單（頁面可能改版）', 'rbpage');
  var fields = rbFormFields_(f), g = function (n) { var v = rbFieldGet_(fields, n); return v == null ? '' : v; };
  if (!g('RecipeId') || !g('Replaceable') || !rbHasField_(f, 'LanguageId')) throw rbErr_('「複製食譜」表單少了欄位（頁面可能改版，這次先不送）', 'rbpage');
  return { token: rbToken_(f) || rbToken_(html), rid: g('RecipeId'), rep: g('Replaceable') };
}
/* 改基本資料：照頁面現值送回，只換名稱等欄；強制停用、不公開 */
function trUpUpdateRecipe_(sess, bid, pg, job, storeIds, note, groupId) {
  var set = { Title: job.title, Note: note, ItemList: pg.cats.map(function (c) { return c.id; }).join(','),
    StoreList: (storeIds || pg.stores.map(function (s) { return s.id; })).join(','), Size: job.size, Content: job.desc, Preserve: job.preserve };
  if (groupId) set.GroupId = groupId;
  var done = {}, pairs = [];
  pg.fields.forEach(function (p) {
    var k = p[0], v = p[1];
    if ((k === 'InUse' || k === 'Public') && v !== 'false') return;   /* 勾起來的 checkbox 拿掉＝停用、不公開（hidden false 留著） */
    if (Object.prototype.hasOwnProperty.call(set, k)) { if (done[k]) return; done[k] = 1; v = set[k]; }
    else if (k === 'RecipeUp' && (!v || /^0001-/.test(v))) v = RB_RECIPE_UP;
    else if (k === 'RecipeDown' && (!v || /^0001-/.test(v))) v = RB_RECIPE_DOWN;
    pairs.push([k, v]);
  });
  ['InUse', 'Public'].forEach(function (k) { if (!pairs.some(function (x) { return x[0] === k; })) pairs.push([k, 'false']); });
  var res = rbReq_(sess, 'post', '/Recipes/Edit/' + encodeURIComponent(bid), { payload: rbEncode_(pairs) });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受修改食譜：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台修改食譜回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
/* 改一步：同 rbEditStep_，但 ings＝null 時食材列不動 */
function trUpEditStep_(sess, pg, s) {
  var payload = {
    '__RequestVerificationToken': pg.token, RecipeId: pg.rid, StepId: pg.sid, Order: pg.order,
    StepMedia: pg.media, StepTitle: s.title, Content: s.html,
    StopClock: pg.stop ? 'true' : 'false', Timer: pg.timer || '', deletedIngred: (s.ings && pg.ings.length) ? pg.ings.join(',') + ',' : ''
  };
  (s.ings || []).forEach(function (g, i) {
    payload['Ingredients[' + i + '].IngredSourceId'] = g.id;
    payload['Ingredients[' + i + '].Amount'] = String(g.amount);
    payload['Ingredients[' + i + '].Unit'] = g.unit;
    payload['Ingredients[' + i + '].Container'] = g.cont || '';
  });
  var res = rbReq_(sess, 'post', '/Steps/Edit/' + encodeURIComponent(pg.sid), { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受修改步驟「' + s.title + '」：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台修改步驟「' + s.title + '」回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
/* 一步的食材列：原文清單（名稱＋單位）→ 品項編號；單位換成目標語言的。有一項對不到 → null（不動） */
function trUpIngs_(srcList, dstUnit, s, have) {
  if (!s.ing.length) return { ings: have ? null : [], why: '' };
  if (s.ing.length !== have) return { ings: null, why: '後台這步有 ' + have + ' 項、原文 ' + s.ing.length + ' 項' };
  var out = [], bad = [];
  s.ing.forEach(function (g) {
    var r = g[1] > 0 ? nrMatchOne_(srcList, [g[0], g[1], g[2], '', g[3], '']) : { miss: 'qty' };
    if (!r.hit) { bad.push(g[0]); return; }
    var u = dstUnit[r.hit.id];
    if (!u || /missing/i.test(u)) u = g[4] || r.hit.unit;
    out.push({ id: r.hit.id, amount: g[1], unit: u, cont: g[3] });
  });
  return bad.length ? { ings: null, why: '對不到：' + bad.join('、') } : { ings: out, why: '' };
}

function trUpload_(p, auth) {
  trUpNeed_(auth);
  var t0 = Date.now(), cache = CacheService.getScriptCache(), id = str_(p.tr_id).trim(), rqk = '';
  if (!id && p.rq) { rqk = 'trrq:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(String(p.rq).slice(0, 80)).getBytes()); id = cache.get(rqk) || ''; }
  if (id && !TR_UP_RE.test(id)) throw fail_('上傳編號不對', 'input');
  var job, st = id ? rbState_(id) : null, im = st ? (st.open || st.done) : null;
  if (im) {
    if (im.abandoned) throw fail_('這次上傳已作廢，請重新上傳', 'state');
    job = rbJob_(st, im);
    if (im.done) return { id: id, data: { tr_id: id, backend_id: job.backend_id || trUpBid_(im), link: rbLink_(trUpBid_(im)), done: true, next: im.total, total: im.total }, msg: '已完成' };
  } else {
    if (id && !rqk) throw fail_('找不到上傳紀錄 ' + id, 'notfound');
    job = trUpJob_(p, auth);
    if (!p.again) {   /* 以前傳過（或還沒做完）→ 先問畫面 */
      var past = trUpPast_(job.src, job.lang).filter(function (x) { return !x.abandoned; });
      if (past.length) throw fail_('這支食譜的' + TR_UP_LABEL[job.lang] + '以前上傳過 ' + past.length + ' 次', 'again', { past: past });
    }
  }
  var lockK = 'trup:' + job.lang + ':' + job.src, holder = cache.get(lockK), me = (im && id) || ('new:' + Utilities.getUuid());
  if (holder && holder !== me) throw fail_('這支食譜的' + TR_UP_LABEL[job.lang] + '正在上傳中（可能另一個畫面也按了），請等 1 分鐘再按', 'busy');
  cache.put(lockK, me, 120);
  try {
    if (!im) {
      id = trUpId_();
      if (rqk) cache.put(rqk, id, 21600);
      cache.put(lockK, id, 120);
      rbLog_(id, id, 'start', '', job.lang + ':' + job.src, 'Y', job.title + '｜' + job.steps.length + ' 步', auth, JSON.stringify(job));
      st = rbState_(id); im = st.open;
      if (!im) throw fail_('上傳紀錄寫不進去，請再按一次', 'state');
    }
    var res = trUpRun_(id, job, im, auth, t0);
    return { id: id, data: res, msg: TR_UP_LABEL[job.lang] + (res.done ? ' 上傳完成 ' + res.total + ' 步' : ' 到第 ' + res.next + '／' + res.total + ' 步') };
  } finally { cache.remove(lockK); }
}
function trUpBid_(im) { var b = ''; (im.rows || []).forEach(function (r) { if (r.backend_id) b = r.backend_id; }); return b; }

function trUpRun_(id, job, im, auth, t0) {
  var sess = rbSession_(), N = job.steps.length, langId = TR_UP_LANG[job.lang], label = TR_UP_LABEL[job.lang];
  var relogin = function () { rbDropSession_(); sess = rbFreshSession_(); };
  var rows = im.rows || [], has = function (a) { return rows.some(function (r) { return r.action === a; }); };
  var bid = trUpBid_(im), k = 0, warn = [];
  rows.forEach(function (r) { if (r.action === 'steps') { var m = String(r.step).match(/(\d+)$/); if (m) k = Math.max(k, +m[1]); } });
  var base = { tr_id: id, total: N, lang: job.lang };
  var ext = function (o) { for (var x in base) o[x] = base[x]; o.backend_id = bid; o.link = rbLink_(bid); return o; };
  var stop = function (msg, code) { rbLog_(id, id, 'fail', bid, '', 'N', msg, auth); throw fail_(msg, code, ext({ next: k })); };

  /* ① 複製（還沒有後台編號） */
  if (!bid) {
    var pg0 = rbAuthedGet_(sess, '/Recipes/Edit/' + encodeURIComponent(job.src));
    if (rbGone_(pg0)) stop('後台找不到原文食譜（可能已經被刪掉）', 'rbgone');
    if (pg0.code !== 200) stop('原文食譜的「編輯食譜」頁打不開（HTTP ' + pg0.code + '）', 'rbnet');
    var cf = trUpCopyForm_(pg0.text), srcFields = rbFormFields_(rbFormHtml_(pg0.text, /^\/Recipes\/Edit\//i) || '');
    var srcTitle = rbFieldGet_(srcFields, 'Title') || '', srcGroup = rbFieldGet_(srcFields, 'GroupId') || '';   /* 複製出來的那支留在原文的地區（10/22 實測），之後在 ② 改成目標語言的地區 */
    if (cf.rid !== job.src || !srcTitle) stop('原文食譜的頁面認不得（頁面可能改版），先停下來', 'rbpage');
    var idx0 = rbStepIndex_(sess, job.src);
    var same = idx0.length === N && idx0.every(function (x, i) { return rbTitleNorm_(x.title) === rbTitleNorm_(job.st[i]); });
    if (!same) stop('後台原文在翻譯之後改過（現在 ' + idx0.length + ' 步、翻譯時 ' + N + ' 步，或步驟標題不同），請重新翻譯再上傳', 'srcchanged');
    var copyResp = '';
    var tried = rows.filter(function (r) { return r.action === 'copytry'; }).pop(), before;
    if (tried) { try { before = JSON.parse(rbReadD_(trUpSh_(), tried._row)); } catch (e) { before = null; } }
    if (!before) {
      before = trUpListIds_(sess, srcTitle, langId, srcGroup);
      rbLog_(id, id, 'copytry', '', '', 'Y', '複製前：' + label + ' 同名 ' + before.length + ' 支', auth, JSON.stringify(before));
      var cr = rbReq_(sess, 'post', '/Recipes/Copy', { payload: { '__RequestVerificationToken': cf.token, Replaceable: cf.rep, RecipeId: cf.rid, LanguageId: langId } });
      if (rbIsLogin_(cr)) { relogin(); pg0 = rbAuthedGet_(sess, '/Recipes/Edit/' + encodeURIComponent(job.src)); cf = trUpCopyForm_(pg0.text);
        cr = rbReq_(sess, 'post', '/Recipes/Copy', { payload: { '__RequestVerificationToken': cf.token, Replaceable: cf.rep, RecipeId: cf.rid, LanguageId: langId } }); }
      if (cr.code >= 400) stop('食譜後台沒有接受複製（HTTP ' + cr.code + '）', 'rbreject');
      copyResp = 'HTTP ' + cr.code + (cr.loc ? ' → ' + cr.loc : '');
    }
    var after = [], neu = [];
    for (var w = 0; w < 3; w++) {
      after = trUpListIds_(sess, srcTitle, langId, srcGroup);
      neu = after.filter(function (x) { return before.indexOf(x) < 0; });
      if (neu.length) break;
      Utilities.sleep(2000);
    }
    if (neu.length !== 1) {
      if (!neu.length) { rbLog_(id, id, 'abandon', '', '', 'N', '複製沒有成功（後台清單沒有多出新的一支；複製回覆 ' + (copyResp || '－') + '）', auth); throw fail_('食譜後台複製沒有成功（清單裡沒有多出新的' + label + '版），這次上傳作廢，可以再按一次', 'rbreject', ext({ next: 0, retry: true })); }
      stop('複製後後台多出 ' + neu.length + ' 支同名的' + label + '版（可能有人同時在複製），先停下來；請到後台確認後通知 Claude', 'mismatch');
    }
    bid = neu[0];
    var chk = rbRecipeEditPage_(sess, bid), cLang = rbFieldGet_(chk.fields, 'LanguageId');
    if (cLang !== langId) stop('複製出來的那支語言不對（' + cLang + '），先停下來', 'mismatch');
    rbLog_(id, id, 'copy', bid, '', 'Y', '已複製成' + label + '版：' + bid + '（' + copyResp + '）', auth);
  }
  var idx = rbStepIndex_(sess, bid);
  if (idx.length !== N) stop('複製出來的' + label + '版有 ' + idx.length + ' 步、翻譯稿 ' + N + ' 步，對不上，先停下來', 'mismatch');

  /* ② 基本資料 */
  if (!has('recipe')) {
    try {
      var pg = rbRecipeEditPage_(sess, bid), oldNote = rbFieldGet_(pg.fields, 'Note') || '', storeIds = null, cpage = rbAuthedGet_(sess, '/Recipes/Create').text;
      var lo = rbOptions_(cpage, 'langOption').filter(function (o) { return o.id === langId; })[0], groupId = lo ? lo.group : '';   /* 目標語言屬於哪個地區 */
      if (!groupId) throw rbErr_('後台「語言」清單找不到' + label + '屬於哪個地區（頁面可能改版）', 'rbpage');
      if (TR_UP_STORE[job.lang]) {
        var opt = rbOptions_(cpage, 'storeOption').filter(function (o) { return o.name === TR_UP_STORE[job.lang] && o.group === groupId; })[0];
        if (opt) storeIds = [opt.id]; else warn.push('後台找不到店別「' + TR_UP_STORE[job.lang] + '」，店別照複製來的');
      }
      var note = ('【食譜翻譯上傳】' + id + '｜' + label + '｜原文：' + job.src + '｜' + Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm') + '｜' + job.by
        + '\n請負責的夥伴檢查後再啟用' + (oldNote && oldNote.indexOf('【食譜翻譯上傳】') !== 0 ? '\n—— 原文的備註 ——\n' + oldNote : '')).slice(0, 4000);
      try { trUpUpdateRecipe_(sess, bid, pg, job, storeIds, note, groupId); }
      catch (e) { if (e && e.code === 'rbrelogin') { relogin(); pg = rbRecipeEditPage_(sess, bid); trUpUpdateRecipe_(sess, bid, pg, job, storeIds, note, groupId); } else throw e; }
      var pg2 = rbRecipeEditPage_(sess, bid);
      if (rbTitleNorm_(rbFieldGet_(pg2.fields, 'Title')) !== rbTitleNorm_(job.title) || pg2.inUse || pg2.pub
        || rbFieldGet_(pg2.fields, 'GroupId') !== groupId || rbFieldGet_(pg2.fields, 'LanguageId') !== langId) throw rbErr_('改完讀回來不對（名稱、地區、語言或啟用狀態）', 'mismatch');
    } catch (e2) { stop('更新' + label + '版的基本資料沒有成功：' + errMsg_(e2), (e2 && e2.code) || 'rbnet'); }
    rbLog_(id, id, 'recipe', bid, '', 'Y', label + '版基本資料已換成翻譯稿（停用、不公開）' + (warn.length ? '；' + warn.join('；') : ''), auth);
  }

  /* ③ 每一步 */
  var srcList = null, dstUnit = null, made = [], stopErr = null, j = k;
  var lists = function () {
    if (srcList) return;
    srcList = rbStepPage_(sess, job.src).ings;
    dstUnit = {}; rbStepPage_(sess, bid).ings.forEach(function (x) { if (x.unit && !dstUnit[x.id]) dstUnit[x.id] = x.unit; });
  };
  for (; j < N; j++) {
    if (Date.now() - t0 > RB_BUDGET_MS) break;
    var s = job.steps[j];
    try {
      lists();
      var pgS = rbStepEditPage_(sess, idx[j].sid);
      if (pgS.rid !== bid) throw rbErr_('後台第 ' + (j + 1) + ' 步不是這支食譜的（頁面可能改版）', 'mismatch');
      var tn = rbTitleNorm_(pgS.title);
      if (tn !== rbTitleNorm_(job.st[j]) && tn !== rbTitleNorm_(s.t)) throw rbErr_('後台第 ' + (j + 1) + ' 步的標題是「' + pgS.title + '」，跟原文不同（可能有人在後台改過）', 'mismatch');
      var m = trUpIngs_(srcList, dstUnit, s, pgS.ings.length), one = { title: s.t, html: s.h, ings: m.ings };
      try { trUpEditStep_(sess, pgS, one); }
      catch (e) { if (e && e.code === 'rbrelogin') { relogin(); pgS = rbStepEditPage_(sess, idx[j].sid); trUpEditStep_(sess, pgS, one); } else throw e; }
      made.push({ n: j + 1, t: s.t, ing: m.ings ? m.ings.length : -1, why: m.why });
    } catch (e3) { stopErr = e3; break; }
  }
  var k2 = k + made.length;
  if (made.length) rbLog_(id, id, 'steps', bid, (k + 1) + '-' + k2, 'Y', '第 ' + (k + 1) + '～' + k2 + ' 步已換成翻譯稿', auth, JSON.stringify(made));
  rbSaveSession_(sess);
  if (stopErr) { k = k2; stop('第 ' + (k2 + 1) + ' 步沒有換成：' + errMsg_(stopErr), (stopErr && stopErr.code) || 'rbnet'); }

  /* ④ 核對 */
  var done = false;
  if (k2 >= N) {
    var idxF = rbStepIndex_(sess, bid), bad = 0;
    if (idxF.length !== N) bad = -1;
    else for (var i = 0; i < N; i++) if (rbTitleNorm_(idxF[i].title) !== rbTitleNorm_(job.steps[i].t)) { bad = i + 1; break; }
    if (bad) { k = k2; stop('上傳完核對後台步驟' + (bad < 0 ? '數不對（' + idxF.length + ' 步）' : '標題不對（第 ' + bad + ' 步）') + '，請到後台看一下後通知 Claude', 'mismatch'); }
    rbLog_(id, id, 'done', bid, String(N), 'Y', label + '版上傳完成：' + N + ' 步', auth);
    done = true;
  }
  /* 這次＋之前做的步驟裡，食材列沒換的（給畫面提醒） */
  var keepIng = [];
  rows.concat(made.length ? [{ action: 'steps', _made: made }] : []).forEach(function (r) {
    if (r.action !== 'steps') return;
    var d = r._made;
    if (!d) { try { d = JSON.parse(rbReadD_(trUpSh_(), r._row)); } catch (e) { d = []; } }
    (d || []).forEach(function (x) { if (x.ing < 0 && x.why) keepIng.push('第 ' + x.n + ' 步（' + x.why + '）'); });
  });
  return ext({ done: done, next: k2, keepIng: keepIng, warn: warn });
}
function trUpSh_() { return ss_().getSheetByName('fact_recipe_import'); }

/* 某支原文以前上傳過的翻譯版（只讀；給畫面顯示「已上傳」） */
function trUpList_(p, auth) {
  var src = str_(p.src).trim().toLowerCase();
  if (!TR_ID.test(src)) throw fail_('原文食譜編號不對', 'input');
  var out = {};
  Object.keys(TR_UP_LANG).forEach(function (l) { var a = trUpPast_(src, l).filter(function (x) { return !x.abandoned; }); if (a.length) out[l] = a; });
  return { id: src, data: out };
}

ACTIONS.trUpload = function (p, a) { return trUpload_(p, a); };
ACTIONS.trUpList = function (p, a) { return trUpList_(p, a); };
READS.trUpList = 1;
SELFLOCK.trUpload = 1;   /* 要連後台、一次 30 多秒：自己用 CacheService 擋重複，不佔全系統的鎖 */

/* 編輯器手動驗證（會在後台新增 1 支停用的測試食譜，名稱開頭【測試請刪除】，測完請經營者刪）：
   原文＝TR_UP_TEST_SRC（停用的舊食譜），語言＝菲律賓英文；翻譯稿用假字，食材照原文各步驟的「編輯步驟」頁 */
var TR_UP_TEST_SRC = '71498aca-2470-4c31-a922-f8a0ed185a5f';   /* 台灣、停用的舊版「蜂蜜檸檬瑪德蓮」6 步 */
function trUpTestNow() {
  var sess = rbSession_(), src = TR_UP_TEST_SRC, idx = rbStepIndex_(sess, src), log = { src: src, steps: idx.length };
  var steps = idx.map(function (x, i) {
    var r = rbAuthedGet_(sess, '/Steps/Edit/' + encodeURIComponent(x.sid)), t = r.text, a = t.indexOf('id="IngredList"'), seg = a < 0 ? '' : t.slice(a, t.indexOf('</form>', a) < 0 ? t.length : t.indexOf('</form>', a));
    var ing = [], re = /<div style="background:aliceblue">([\s\S]*?)<\/div>/gi, m;
    while ((m = re.exec(seg))) { var sp = (m[1].match(/<span\b[^>]*>([\s\S]*?)<\/span>/gi) || []).map(function (y) { return rbText_(y); });
      var cont = sp[3] === '彩色塑膠碗' ? 'container1' : sp[3] === '白色瓷碗' ? 'container2' : ''; ing.push([sp[0], Number(sp[1]), sp[2], cont, '']); }
    return { t: 'TEST step ' + (i + 1) + ' (' + x.title + ')', h: '<p>Test content ' + (i + 1) + ' <b>50</b> g</p>', ing: ing };
  });
  log.ingPerStep = steps.map(function (s) { return s.ing.length; });
  var auth = { kind: 'role', role: ADMIN }, p = { src: src, lang: 'ph', again: 1, rq: 'test' + Date.now(), st: idx.map(function (x) { return x.title; }),
    tr: { name: '【測試請刪除】翻譯上傳測試 ' + Utilities.formatDate(new Date(), TZ, 'MMdd HHmm'), size: 'TEST size', desc: 'TEST description', preserve: 'TEST preserve', steps: steps } };
  var res, n = 0;
  do { res = trUpload_(n ? { tr_id: res.data.tr_id } : p, auth); n++; } while (!res.data.done && n < 6);
  log.res = res.data; log.calls = n;
  var bid = res.data.backend_id, pg = rbRecipeEditPage_(sess, bid), g = function (k) { return rbFieldGet_(pg.fields, k); };
  log.after = { title: g('Title'), group: g('GroupId'), lang: g('LanguageId'), rep: g('Replaceable'), inUse: pg.inUse, pub: pg.pub, stores: pg.stores, cats: pg.cats, up: g('RecipeUp'), down: g('RecipeDown') };
  var idx2 = rbStepIndex_(sess, bid);
  log.titles = idx2.map(function (x) { return x.title; });
  var s2 = rbAuthedGet_(sess, '/Steps/Edit/' + encodeURIComponent(idx2[Math.min(1, idx2.length - 1)].sid)).text, a2 = s2.indexOf('id="IngredList"');
  log.step2ings = (s2.slice(a2, a2 + 6000).match(/<div style="background:aliceblue">[\s\S]*?<\/div>/gi) || []).map(rbText_);
  log.srcStillInactive = !rbRecipeEditPage_(sess, src).inUse;
  Logger.log(JSON.stringify(log));
}
