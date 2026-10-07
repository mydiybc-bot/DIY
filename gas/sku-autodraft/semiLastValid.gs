/*****************************************************************
 * semiLastValid.gs v1.0（2026-09-23 ce 批）— 最後有效日也算進店製半成品的原料
 * 專案：diybc-sku-autodraft（新增檔案，放在檔案清單最下面）
 * ⚠ 一定要排在 skuAutoDraft.gs 後面載入（本檔一載入就去抓 adLastValid_；排前面＝抓不到、包裝沒接上）。用 clasp push 時 .clasp.json 要設 filePushOrder [skuAutoDraft.js, semiLastValid.js]，否則 clasp 會照字母排序把本檔排到前面（2026-10-07 實際發生過）。
 *
 * 為什麼：店長清單會把「最後有效日＜今天−7 天」的品項當過季藏起來。鹹蛋黃這類只出現在
 *   店製配方（dim_semi）裡的原料，BOM 表看不到它被哪支甜點用 ⇒ 最後有效日停在舊日期 ⇒ 被藏起來、店長看不到要叫。
 * 做什麼：只包裝 adLastValid_（D 段最後有效日）：把傳進來的 BOM 先照 dim_semi 展開，再交給原函式。
 *   A 別名／B 草稿／C 清標記完全不動。原函式照常執行，本檔不複製它。
 * 展開規則與 diybc-purchase-agg 的 semiExpand.gs 同一份（下方「核心」逐字相同，兩邊要一起改）。
 * 關掉：SEMI_EXPAND_ON 改 false。
 *****************************************************************/
var SEMI_EXPAND_ON = true;
var SEMI_EXPAND_VER = 'semiLastValid v1.1 2026-10-07';   // v1.1：店製品項本身也算最後有效日（檔期結束會退場）
var SEMI_TAB_ = 'dim_semi';
var SEMI_SHARED_ = '＊共用';
var SEMI_MAX_DEPTH_ = 3;

var SEMI_ORIG_ADLV_ = (typeof adLastValid_ === 'function') ? adLastValid_ : null;
if (SEMI_ORIG_ADLV_) {
  adLastValid_ = function (ss, skuSh, sv, h, bv, bh, bName, exact, dry, today, out) {
    var bv2 = bv;
    if (SEMI_EXPAND_ON) {
      try {
        var rules = semiXRules_(ss);
        if (rules && rules.n) {
          var r = semiXExpandTable_(bh, bv.slice(1), rules);
          bv2 = [bv[0]].concat(r.rows, bv.slice(1));   /* v1.1（2026-10-07）：原 BOM 列也保留——店製品項本身（例：啾啾鳥材料包）也要算最後有效日；原本展開後父列被換掉，店製品項永遠是空白＝長期有效、檔期結束也不退場。這裡只看「哪支甜點用到」，列重複不影響結果 */
          out.push(' 🧁 最後有效日已含店製半成品原料＋店製品項本身：BOM 展開 ' + r.stat.parents + ' 列 → ' + r.stat.children + ' 列（配方 ' + Object.keys(rules.by).length + ' 項）');
        }
      } catch (e) { out.push(' ⚠ 店製半成品展開失敗，最後有效日照原 BOM 算：' + e); }
    }
    return SEMI_ORIG_ADLV_(ss, skuSh, sv, h, bv2, bh, bName, exact, dry, today, out);
  };
}

/** 確認包裝有接上 */
function aaaSemiCheck() {
  var msg = SEMI_EXPAND_VER + '｜開關 ' + SEMI_EXPAND_ON + '｜adLastValid_ 包裝：' + (SEMI_ORIG_ADLV_ && adLastValid_ !== SEMI_ORIG_ADLV_ ? '✅ 已接上' : '❌ 沒接上（本檔要排在主程式後面）');
  var rules = semiXRules_(SpreadsheetApp.openById(AD_SHEET_ID));
  msg += '｜dim_semi 配方 ' + Object.keys(rules.by).length + ' 項／' + rules.n + ' 列';
  Logger.log(msg); return msg;
}

/* ---------- 核心（diybc-sku-autodraft 的 semiLastValid.gs 有同一份，兩邊要一起改）---------- */
function semiXNorm_(s) { return String(s == null ? '' : s).replace(/[\s\u3000]+/g, ''); }
function semiXKey_(d, it) { return semiXNorm_(d) + '\u0001' + semiXNorm_(it); }

/** 讀 dim_semi → {n: 有效原料列數, by: {甜點|品項: {dessert,item,yq,yu,lines:[{mat,q,u}]}}, warn: []}；沒有分頁回 n=0 */
function semiXRules_(ss, rowsOverride) {
  var res = { n: 0, by: {}, warn: [] };
  var v;
  if (rowsOverride) { v = rowsOverride; }
  else {
    var sh = ss.getSheetByName(SEMI_TAB_);
    if (!sh) return res;
    v = sh.getDataRange().getValues();
  }
  if (!v || v.length < 2) return res;
  var h = v[0].map(function (x) { return semiXNorm_(x); });
  var c = {};
  ['甜點名稱', '店製品項', '製作單位', '一批產出量', '產出單位', '原料', '一批用量', '單位'].forEach(function (k) { c[k] = h.indexOf(k); });
  if (c['店製品項'] < 0 || c['原料'] < 0 || c['一批產出量'] < 0 || c['一批用量'] < 0) { res.warn.push('dim_semi 表頭不完整，本次不展開'); return res; }
  function g(r, k) { return c[k] < 0 ? '' : r[c[k]]; }
  function s(x) { return String(x == null ? '' : x).trim(); }
  for (var i = 1; i < v.length; i++) {
    var r = v[i], item = s(g(r, '店製品項')), mat = s(g(r, '原料'));
    if (!item || !mat) continue;
    if (s(g(r, '製作單位')) === '出貨中心') continue;
    var ds = s(g(r, '甜點名稱')) || SEMI_SHARED_;
    var yq = Number(g(r, '一批產出量')), q = Number(g(r, '一批用量')), yu = s(g(r, '產出單位')), u = s(g(r, '單位'));
    if (!(yq > 0) || !(q > 0)) { res.warn.push('第 ' + (i + 1) + ' 列「' + item + '／' + mat + '」一批產出量或一批用量不是正數，略過'); continue; }
    var key = semiXKey_(ds, item), o = res.by[key];
    if (!o) o = res.by[key] = { dessert: ds, item: item, yq: yq, yu: yu, lines: [] };
    else if (o.yq !== yq || o.yu !== yu) res.warn.push('「' + ds + '｜' + item + '」各列一批產出量／單位不一致，以第一列（' + o.yq + ' ' + o.yu + '）為準');
    o.lines.push({ mat: mat, q: q, u: u });
    res.n++;
  }
  return res;
}

function semiXFind_(rules, d, it) { return rules.by[semiXKey_(d, it)] || rules.by[semiXKey_(SEMI_SHARED_, it)] || null; }

/** 單一品項展開成原料清單；沒有配方／單位不符／循環 → null（呼叫端保留原列） */
function semiXExpandOne_(rules, d, it, q, u, depth, path, st) {
  var rec = semiXFind_(rules, d, it);
  if (!rec) return null;
  if (semiXNorm_(rec.yu) !== semiXNorm_(u)) { st.unitMis[d + '｜' + it + '（BOM「' + u + '」≠ 產出單位「' + rec.yu + '」）'] = 1; return null; }
  var nk = semiXNorm_(it);
  if (depth >= SEMI_MAX_DEPTH_ || path[nk]) { st.cycles[d + '｜' + it] = 1; return null; }
  var p2 = {}; Object.keys(path).forEach(function (k) { p2[k] = 1; }); p2[nk] = 1;
  var f = q / rec.yq, out = [];
  rec.lines.forEach(function (l) {
    var qq = l.q * f;
    /* 分裝包：品項＝原料本身（例：紅豆餡 1 包 → 紅豆餡 130 g），直接當原料，不再往下找 */
    if (semiXNorm_(l.mat) === nk) { out.push({ mat: l.mat, q: qq, u: l.u }); return; }
    var sub = semiXExpandOne_(rules, d, l.mat, qq, l.u, depth + 1, p2, st);
    if (sub) out = out.concat(sub); else out.push({ mat: l.mat, q: qq, u: l.u });
  });
  return out;
}

/** BOM 整表展開：headers＝表頭陣列、rows＝資料列（不含表頭）。回傳新陣列，不改原陣列。 */
function semiXExpandTable_(headers, rows, rules) {
  var st = { parents: 0, children: 0, desserts: {}, items: {}, unitMis: {}, cycles: {} };
  if (!rules || !rules.n) return { rows: rows, stat: st };
  var H = headers.map(function (x) { return semiXNorm_(x); });
  var iD = H.indexOf('甜點名稱'), iI = H.indexOf('食材/器具名稱'), iQ = H.indexOf('數量'), iU = H.indexOf('單位'), iC = H.indexOf('容器');
  if (iD < 0 || iI < 0 || iQ < 0) return { rows: rows, stat: st };
  var out = [];
  rows.forEach(function (r) {
    var d = String(r[iD] == null ? '' : r[iD]).trim(), it = String(r[iI] == null ? '' : r[iI]).trim();
    var q = Number(r[iQ]), u = iU >= 0 ? String(r[iU] == null ? '' : r[iU]).trim() : '';
    if (!d || !it || !(q > 0)) { out.push(r); return; }
    var kids = semiXExpandOne_(rules, d, it, q, u, 0, {}, st);
    if (!kids || !kids.length) { out.push(r); return; }
    st.parents++; st.desserts[d] = 1; st.items[it] = 1;
    kids.forEach(function (k, j) {
      var nr = r.slice();
      nr[iI] = k.mat; nr[iQ] = k.q;
      if (iU >= 0) nr[iU] = k.u;
      if (iC >= 0 && j > 0) nr[iC] = '';
      out.push(nr); st.children++;
    });
  });
  return { rows: out, stat: st };
}