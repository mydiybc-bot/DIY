/* ===== nm_match.gs — 公版圖片／影片配對「五燈號」（2026-10-18）=====
   經營者 10/18：「新食譜上傳的圖片影片自動選配，還是有大量步驟配不到」→「用 2026 年的食譜交叉比對步驟內容」；
   評估報告《公版圖片影片配對改善評估_20261018.md》（輸出/成品）第四、六節，經營者同意照做。
   資料：nm_scan.gs 的 nmAggregate_ 每天寫 media_match 分頁（2026-01-01 以後更新的繁中食譜，每一步一列：標題鍵、步驟文字、代表檔、同內容群組）。
   比法：
   ・標題鍵＝nrTitleKey_（去空白、去括號內容）；步驟文字＝去標點數字空白後，相鄰兩字的重疊比例（0～1）。
   ・「最像的食譜」：兩支共同的步驟標題，越少見的越算分（教學影片這種每支都有的幾乎不算），除以這支自己的總分。
   每一步依序判斷：
   🟢 auto  同標題、文字相似 ≥ 0.9 的步驟裡，最多款用的那一版 ≥ 2 款且占 ≥ 70%        （試算 90% 對）
   🟡 sib   整份最像的那一支（相似度 ≥ 0.5）有同標題、文字幾乎一樣（≥ 0.99）的步驟 → 照它配，請人確認 （試算 76% 對）
   ⚪ text  2026 年這個標題 ≥ 5 款、八成以上都沒放圖片影片 → 「這一步通常不放」          （98% 確實不放）
   🔵 maybe 同標題有文字相近（≥ 0.7）的、或常用版本（≥ 3 款、占 ≥ 30%）→ 最多 3 個候選給人選
   🔴 none  都沒有 → 要上傳／新拍
   只讀：不寫後台，不寫試算表（media_match 由 nm_scan.gs 寫）。驗證動作 nmEval：每一支當成剛上傳、排除同一款，回統計數字。 */
var NMX_CACHE = 'nmmatch:v1';
var NMX = { autoD: 0.9, autoMin: 2, autoShare: 0.7, sibD: 0.99, sibSim: 0.5, textMin: 5, textShare: 0.8, maybeD: 0.7, maybeShare: 0.3, maybeMin: 3, maxList: 3, textMax: 200 };
var _nmx = null;   /* 同一次執行內記住索引 */

/* ---------- 文字 ---------- */
function nmxNorm_(s) { return String(s == null ? '' : s).replace(/[\s\d０-９.,，、。:：;；!！?？~～\-－_＿#＃*()（）【】\[\]「」『』\/／+＋%％'"“”‘’…·•]/g, ''); }
function nmxBg_(x) { var S = {}, n = 0; x = String(x || ''); for (var i = 0; i < x.length - 1; i++) { var g = x.substr(i, 2); if (!S[g]) { S[g] = 1; n++; } } return { s: S, n: n }; }
function nmxDice_(a, b) {
  if (!a.n || !b.n) return 0;
  var A = a.n <= b.n ? a : b, B = A === a ? b : a, c = 0;
  for (var g in A.s) if (B.s[g]) c++;
  return 2 * c / (a.n + b.n);
}

/* ---------- 資料：media_match 分頁 → {R:[[id,name,base]], G:[[path,kind,mb]], S:[[ri,key,text,gi]]}（暫存 10 分鐘） ---------- */
function nmxDrop_() { _nmx = null; try { CacheService.getScriptCache().remove(NMX_CACHE); } catch (e) {} }
function nmxRaw_() {
  var cache = CacheService.getScriptCache(), s = cacheGetBig_(cache, NMX_CACHE);
  if (s) { try { var o0 = JSON.parse(s); if (o0 && o0.S) return o0; } catch (e) {} }
  var sh = ss_().getSheetByName(NM_MATCH);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, NM_MATCH_HEAD.length).getValues();
  var R = [], ri = {}, G = [], gi = {}, S = [];
  v.forEach(function (r) {
    var id = String(r[0] || ''); if (!id) return;
    if (ri[id] == null) { ri[id] = R.length; R.push([id, String(r[1]), String(r[2])]); }
    var grp = String(r[9] || ''), g = -1;
    if (grp) {
      if (gi[grp] == null) { gi[grp] = G.length; G.push([String(r[6] || ''), String(r[7] || ''), Number(r[8]) || 0]); }
      g = gi[grp];
      if (!G[g][0] && r[6]) G[g] = [String(r[6]), String(r[7] || ''), Number(r[8]) || 0];
    }
    S.push([ri[id], String(r[4] || ''), String(r[5] || ''), g]);
  });
  var o = { R: R, G: G, S: S, at: nmMetaGet_('aggregated_at'), from: NM_MATCH_FROM };
  cachePutBig_(cache, NMX_CACHE, JSON.stringify(o), 600);
  return o;
}
function nmxIndex_() {
  if (_nmx) return _nmx;
  var o = nmxRaw_(); if (!o || !o.S.length) return null;
  _nmx = nmxIndexOf_(o);
  return _nmx;
}
function nmxIndexOf_(o) {
  var X = { R: o.R, G: o.G, S: o.S, at: o.at, from: o.from, byKey: {}, rk: [], rsteps: [], df: {}, km: {}, bg: [], nb: 0 };
  var baseKeys = {};
  o.R.forEach(function () { X.rk.push({}); X.rsteps.push([]); });
  o.S.forEach(function (s, i) {
    var r = s[0], key = s[1], g = s[3], b = o.R[r][2];
    X.rsteps[r].push(i);
    if (!key) return;
    X.rk[r][key] = 1;
    (baseKeys[b] = baseKeys[b] || {})[key] = 1;
    var km = X.km[key] = X.km[key] || { m: {}, n: {} };
    if (g >= 0) { km.m[b] = 1; (X.byKey[key] = X.byKey[key] || []).push(i); }
    else km.n[b] = 1;
  });
  var bases = Object.keys(baseKeys); X.nb = bases.length;
  bases.forEach(function (b) { for (var k in baseKeys[b]) X.df[k] = (X.df[k] || 0) + 1; });
  X.rk = X.rk.map(function (m) { return Object.keys(m); });
  return X;
}
/* 2026-10-18 試驗用：直接從 media_scan 組一個「這天以後更新的繁中食譜」母體（跟 media_match 同形狀，R 多第 4 欄更新日），給 nmEval 比較不同年份範圍 */
function nmxBuild_(fromKey) {
  var sh = ss_().getSheetByName(NM_SCAN);
  if (!sh || sh.getLastRow() < 2) return null;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, NM_SCAN_HEAD.length).getValues(), blob = nmBlobMap_(), byMd5 = {};
  Object.keys(blob).forEach(function (p) { var h = blob[p].md5; if (h) (byMd5[h] = byMd5[h] || []).push(p); });
  var okFile = function (p) { var b = blob[p]; return !!(b && b.bytes <= NM_MAX_MB * 1048576 && /\.(mp4|mov|m4v|jpg|jpeg|png|gif)$/i.test(p)); };
  var R = [], G = [], gi = {}, S = [], noText = 0;
  rows.forEach(function (r) {
    if (!r[5] || r[6] || !r[7] || String(r[2]) !== '繁體中文') return;
    var upd = nmYmd_(r[4]); if (upd < fromKey) return;
    var steps; try { steps = JSON.parse(r[7]); } catch (e) { return; }
    var ri = R.length; R.push([String(r[0]), String(r[1]), nmBase_(r[1]), upd]);
    steps.forEach(function (s) {
      if (s.x == null) noText++;
      var g = -1;
      if (s.m) {
        var b = blob[s.m], grp = (b && b.md5) ? 'md5:' + b.md5 : 'p:' + s.m;
        if (gi[grp] == null) {
          var p = okFile(s.m) ? s.m : '';
          if (!p && b && b.md5) (byMd5[b.md5] || []).some(function (q) { if (okFile(q)) { p = q; return true; } return false; });
          gi[grp] = G.length; G.push([p, s.k || '', p ? Math.round(blob[p].bytes / 1048576 * 100) / 100 : 0]);
        }
        g = gi[grp];
      }
      S.push([ri, nrTitleKey_(s.t), nmxNorm_(s.x).slice(0, NMX.textMax), g]);
    });
  });
  return { R: R, G: G, S: S, at: nmNow_(), from: fromKey, noText: noText };
}
function nmxIdf_(X, k) { return Math.log(X.nb / (X.df[k] || 1)); }
function nmxBgOf_(X, i) { return X.bg[i] || (X.bg[i] = nmxBg_(X.S[i][2])); }

/* ---------- 核心：Q＝整份食譜的步驟 [{key, bg}]；exBase＝評估時排除的「款」（上傳時不排除）---------- */
function nmxCore_(X, Q, exBase) {
  var qk = {}, self = 0, ex = exBase != null;
  Q.forEach(function (q) { if (q.key) qk[q.key] = 1; });
  Object.keys(qk).forEach(function (k) { self += nmxIdf_(X, k); });
  var sim = [], best = -1, bestSim = -1;
  X.R.forEach(function (r, ri) {
    if (ex && r[2] === exBase) { sim.push(-1); return; }
    var sc = 0; X.rk[ri].forEach(function (k) { if (qk[k]) sc += nmxIdf_(X, k); });
    var v = self ? sc / self : 0; sim.push(v);
    if (v > bestSim) { bestSim = v; best = ri; }
  });
  return Q.map(function (q) {
    var C = [];
    if (q.key) (X.byKey[q.key] || []).forEach(function (si) {
      var s = X.S[si], b = X.R[s[0]][2];
      if (ex && b === exBase) return;
      C.push({ si: si, r: s[0], b: b, g: s[3], d: nmxDice_(q.bg, nmxBgOf_(X, si)), sim: sim[s[0]] });
    });
    /* 🟢 文字相近的步驟「投票」 */
    var vote = {}, tot9 = 0, vs;
    C.forEach(function (x) { if (x.d >= NMX.autoD) (vote[x.g] = vote[x.g] || {})[x.b] = 1; });
    vs = Object.keys(vote).map(function (g) { return { g: +g, n: Object.keys(vote[g]).length }; }).sort(function (a, b) { return b.n - a.n; });
    vs.forEach(function (v) { tot9 += v.n; });
    if (vs.length && vs[0].n >= NMX.autoMin && vs[0].n / tot9 >= NMX.autoShare && X.G[vs[0].g][0]) {
      var dd = 0; C.forEach(function (x) { if (x.g === vs[0].g && x.d > dd) dd = x.d; });
      return { tier: 'auto', g: vs[0].g, n: vs[0].n, tot: tot9, share: Math.round(vs[0].n / tot9 * 100), d: dd };
    }
    /* 🟡 最像的那一支 */
    var s1 = null;
    C.forEach(function (x) { if (!s1 && x.r === best && x.d >= NMX.sibD) s1 = x; });
    if (s1 && bestSim >= NMX.sibSim && X.G[s1.g][0]) return { tier: 'sib', g: s1.g, n: 1, d: s1.d, from: X.R[best][1], rsim: Math.round(bestSim * 100) };
    /* ⚪ 通常不放 */
    var km = X.km[q.key];
    if (km) {
      var nN = 0, nM = 0, b;
      for (b in km.n) if (!ex || b !== exBase) nN++;
      for (b in km.m) if (!ex || b !== exBase) nM++;
      if (nN >= NMX.textMin && nN / (nN + nM) >= NMX.textShare) return { tier: 'text', n: nN, share: Math.round(nN / (nN + nM) * 100) };
    }
    /* 🔵 候選 */
    var vb = {}, nbS = {}, nb, topN = 0, sc = {}, at = {}, bestD = 0;
    C.forEach(function (x) { (vb[x.g] = vb[x.g] || {})[x.b] = 1; nbS[x.b] = 1; });
    nb = Object.keys(nbS).length;
    Object.keys(vb).forEach(function (g) { var c = Object.keys(vb[g]).length; vb[g] = c; if (c > topN) topN = c; });
    C.forEach(function (x) {
      var v = x.d + 0.5 * x.sim + 0.5 * (vb[x.g] / (nb || 1));
      if (sc[x.g] == null || v > sc[x.g]) { sc[x.g] = v; at[x.g] = x; }
      if (x.d > bestD) bestD = x.d;
    });
    var L = Object.keys(sc).filter(function (g) { return X.G[g][0]; }).sort(function (a, b) { return sc[b] - sc[a]; }).slice(0, NMX.maxList);
    if (L.length && (bestD >= NMX.maybeD || (nb >= NMX.maybeMin && topN / nb >= NMX.maybeShare)))
      return { tier: 'maybe', nb: nb, list: L.map(function (g) { var x = at[g]; return { g: +g, n: vb[g], share: Math.round(vb[g] / (nb || 1) * 100), d: x.d, from: X.R[x.r][1] }; }) };
    /* 🔴 要上傳 */
    return { tier: 'none', n: nb };
  });
}

/* ---------- 給 nrPreview_（📤 新食譜上傳）／rbPreview_（🍰 匯入）用：steps＝rbCleanRecipe_ 的 [{t, h}] ----------
   回每一步 {tier:'auto'|'sib'|'text'|'maybe'|'none', list:[自動配的那一個], alt:[候選], n, share, from}；
   media_match 還沒有資料 → 退回舊規則 nrMediaFor_（只看標題，fallback:true） */
function nmMatchRecipe_(steps) {
  var X = null, fb = function () {
    return steps.map(function (s) {
      var mf = nrMediaFor_(s.t);
      return { tier: mf.tier, list: mf.tier === 'auto' ? mf.list : [], alt: mf.tier === 'maybe' ? mf.list.slice(0, 3) : [], n: mf.n, share: 0, from: '', fallback: true };
    });
  };
  try { X = nmxIndex_(); } catch (e) { X = null; }
  if (!X) return fb();
  try { return nmxRecipe_(X, steps); } catch (e) { return fb(); }   /* 配對出錯也不讓整個檢查失敗：退回舊規則 */
}
function nmxRecipe_(X, steps) {
  var Q = steps.map(function (s) { return { key: nrTitleKey_(s.t), bg: nmxBg_(nmxNorm_(rbText_(s.h)).slice(0, NMX.textMax)) }; });
  var item = function (g, o, t) {
    var G = X.G[g];
    return { url: RB_MEDIA_PREFIX + G[0], type: (G[1] === 'v' || /\.(mp4|mov|m4v)$/i.test(G[0])) ? '影片' : '圖片', mb: G[2], recipes: o.n || 0, share: o.share || 0,
      sim: o.d != null ? Math.round(o.d * 100) : '', from: o.from || '', rsim: o.rsim || '', id: '', title: t };
  };
  return nmxCore_(X, Q, null).map(function (r, i) {
    var t = String(steps[i].t || '').trim();
    if (r.tier === 'auto' || r.tier === 'sib') return { tier: r.tier, list: [item(r.g, r, t)], alt: [], n: r.tot || 0, share: r.share || 0, from: r.from || '' };
    if (r.tier === 'maybe') return { tier: 'maybe', list: [], alt: r.list.map(function (c) { return item(c.g, c, t); }), n: r.nb || 0, share: 0, from: '' };
    return { tier: r.tier, list: [], alt: [], n: r.n || 0, share: r.share || 0, from: '' };
  });
}
function nmMatchMeta_() {
  var X = null;
  try { X = nmxIndex_(); } catch (e) { X = null; }
  return X ? { source: 'match', scanned: X.at, recipes: X.R.length, steps: X.S.length, recent_from: String(X.from).replace(/\//g, '-') } : null;
}

/* ---------- 驗證：每一支 2026 食譜當成剛上傳（排除同一款），配出來跟後台實際放的比 ---------- */
function nmEval_(p, auth) {
  var t0 = Date.now(), X, tm = {}, from = String((p && p.from) || '').trim(), only = null, noText = 0;
  var X26 = null, map26 = null, born = String((p && p.born) || '').replace(/\D/g, '').slice(0, 8), bornOk = null;
  if (from) {   /* 2026-10-18 試驗：母體改成 from 以後更新的食譜，受測的仍是 2026 年的食譜（排除同一款） */
    var o = nmxBuild_(nmYmd_(from)); if (!o) throw fail_('media_scan 沒有資料', 'nodata');
    tm.build = Date.now() - t0; noText = o.noText;
    var t1 = Date.now(); X = nmxIndexOf_(o); tm.index = Date.now() - t1;
    only = function (r) { return r[3] >= NM_MATCH_FROM; };
    if (p.cascade) {   /* 2026 優先：2026 母體配出綠／黃就用它，沒有才用大母體的綠／黃 */
      var o26 = { R: [], G: o.G, S: [] }; map26 = {};
      o.R.forEach(function (r, i) { if (r[3] >= NM_MATCH_FROM) { map26[i] = o26.R.length; o26.R.push(r); } });
      o.S.forEach(function (st) { if (map26[st[0]] != null) o26.S.push([map26[st[0]], st[1], st[2], st[3]]); });
      X26 = nmxIndexOf_(o26);
    }
  } else X = nmxIndex_();
  if (born) {   /* 只測「這一款第一次出現」在 born 以後的食譜（全新甜點；用 🧬 recipe_lineage 的 kuan_birth） */
    var lt = ss_().getSheetByName(LG_TAB); bornOk = {};
    if (lt && lt.getLastRow() > 1) {
      var lv = lt.getRange(2, 1, lt.getLastRow() - 1, LG_HEAD.length).getDisplayValues(), ci = LG_HEAD.indexOf('kuan_birth');
      lv.forEach(function (r) { if (String(r[ci]).replace(/\D/g, '').slice(0, 8) >= born) bornOk[String(r[0]).toLowerCase()] = 1; });
    }
  }
  if (!X) throw fail_('media_match 分頁還沒有資料（等公版掃描跑完再驗）', 'nodata');
  var t2 = Date.now(), nT = 0;
  var gb = {};
  X.S.forEach(function (s) { if (s[3] >= 0) (gb[s[3]] = gb[s[3]] || {})[X.R[s[0]][2]] = 1; });
  var stat = {}, tot = { recipes: X.R.length, steps: 0, media: 0, none: 0 }, ex = [];
  var add = function (tier, k) { var o = stat[tier] = stat[tier] || {}; o[k] = (o[k] || 0) + 1; };
  var want = String((p && p.name) || '').trim();
  X.R.forEach(function (r, ri) {
    if (only && !only(r)) return;
    if (bornOk && !bornOk[String(r[0]).toLowerCase()]) return;
    nT++;
    var idx = X.rsteps[ri];
    var Q = idx.map(function (si) { return { key: X.S[si][1], bg: nmxBgOf_(X, si) }; });
    var res = nmxCore_(X, Q, r[2]);
    if (X26) {
      var r26 = nmxCore_(X26, Q, r[2]);
      res = res.map(function (q, j) { var a = r26[j]; return (a.tier === 'auto' || a.tier === 'sib') ? a : ((q.tier === 'auto' || q.tier === 'sib') ? q : a); });
    }
    var line = [];
    idx.forEach(function (si, j) {
      var truth = X.S[si][3], q = res[j], kind, k;
      tot.steps++; if (truth >= 0) tot.media++; else tot.none++;
      if (truth < 0) kind = '本來沒放';
      else { var others = 0; for (var b in gb[truth]) if (b !== r[2]) others++; kind = others ? '別款有同檔' : '這款專屬'; }
      if (q.tier === 'auto' || q.tier === 'sib') k = truth < 0 ? '本來沒放' : (q.g === truth ? '配對正確' : '配錯');
      else if (q.tier === 'maybe') k = truth < 0 ? '本來沒放' : (q.list.some(function (c) { return c.g === truth; }) ? '候選內有對的' : '候選都不對');
      else k = kind;
      add(q.tier, k);
      if (want && r[1].indexOf(want) >= 0) line.push(X.S[si][1] + ':' + q.tier + (k === '配對正確' ? '✓' : k === '配錯' ? '✗' : ''));
    });
    if (line.length) ex.push(r[1] + '｜' + line.join('、'));
  });
  tm.match = Date.now() - t2; tm.per_recipe = nT ? Math.round(tm.match / nT) : 0;
  tot.targets = nT; tot.pool_recipes = X.R.length; tot.pool_steps = X.S.length; tot.pool_no_text = noText;
  return { id: '', data: { stat: stat, tot: tot, at: X.at, from: X.from, rule: NMX, example: ex.slice(0, 3), ms: tm, secs: Math.round((Date.now() - t0) / 1000) } };
}
ACTIONS.nmEval = function (p, a) { return nmEval_(p, a); };
READS.nmEval = 1;
NOAUTH.nmEval = 1;   /* 只回統計數字（可帶 name 看某一支每一步的燈號；帶 from＝試別的母體年份範圍），不回圖片影片、不寫任何東西 */
