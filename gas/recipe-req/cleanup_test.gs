/* ===== 2026-10-18 經營者：「把這些『測試』的檔期都刪掉，看起來乾淨點」（🔎 查申請表的檔期下拉有 13 個測試檔期）=====
   只在編輯器執行，不接 doGet／doPost：
   ① tcReport()：只列清單，不改任何東西（看執行記錄）
   ② tcDelete()：先把整份試算表複製一份備份到雲端硬碟，再刪；前後筆數寫進執行記錄與 log 分頁
   刪什麼：檔期名稱有「測試」的 fact_campaign 列；這些檔期底下的申請表（fact_recipe_req）；
          這些申請表在 fact_req_unit／fact_signoff／fact_req_fill／fact_push／fact_recipe_import／fact_req_comment／fact_notify 的紀錄；
          這些申請表寫入採購前留的 backup_ 分頁（整份備份裡都還有）。
   不刪：log 分頁（歷史紀錄）、dim_signer 的測試簽核甲／乙／丙（已停用、送簽看不到）。
   安全檢查（任一項不過就不刪，只回報）：測試申請表現在是「已寫入採購」（採購系統裡還有測試資料）；
          測試申請表在 fact_recipe_import 有建到門市食譜後台的編號（後台可能還有測試食譜） */
var TC_KEY = '測試';
var TC_CHILD = ['fact_req_unit', 'fact_signoff', 'fact_req_fill', 'fact_push', 'fact_recipe_import', 'fact_req_comment', 'fact_notify'];

function tcPlan_() {
  var ss = ss_(), ct = load_('fact_campaign'), camps = [], keep = [], cid = {};
  ct.rows.forEach(function (r) {
    var id = str_(r.campaign_id), nm = str_(r['檔期名稱']);
    if (!id) return;
    var o = { row: r._row, id: id, name: nm, from: str_(r['起']), status: str_(r.status) };
    if (nm.indexOf(TC_KEY) >= 0) { camps.push(o); cid[id] = 1; } else keep.push(o);
  });
  var rt = load_('fact_recipe_req', true), reqs = [], rid = {}, keepReqs = [], orphan = [];
  var allCamp = {}; ct.rows.forEach(function (r) { allCamp[str_(r.campaign_id)] = 1; });
  rt.rows.forEach(function (r) {
    var id = str_(r.req_id); if (!id) return;
    var o = { row: r._row, id: id, camp: str_(r.campaign_id), name: str_(r['商品正式名稱']).trim() || str_(r['商品暫定名稱']).trim(), status: str_(r.status) };
    if (cid[o.camp]) { reqs.push(o); rid[id] = 1; }
    else { keepReqs.push(o); if (!allCamp[o.camp]) orphan.push(o); }
  });
  var child = {}, warn = [];
  TC_CHILD.forEach(function (name) {
    var t; try { t = load_(name); } catch (e) { child[name] = { total: 0, del: [], miss: true }; return; }
    var del = t.rows.filter(function (r) { return rid[str_(r.req_id)]; });
    child[name] = { total: t.rows.length, del: del.map(function (r) { return r._row; }) };
    if (name === 'fact_recipe_import') del.forEach(function (r) { if (str_(r.backend_id)) warn.push('門市後台有建檔：' + str_(r.req_id) + ' → ' + str_(r.backend_id) + '（' + str_(r.action) + '）'); });
  });
  reqs.forEach(function (q) { if (q.status === PUSH_OK) warn.push('現在是「已寫入採購」：' + q.id + ' ' + q.name); });
  var tabs = ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) {
    if (n.indexOf('backup_') !== 0) return false;
    return Object.keys(rid).some(function (id) { return n.indexOf(id) >= 0; });
  });
  return { camps: camps, keep: keep, reqs: reqs, keepReqs: keepReqs, orphan: orphan, child: child, tabs: tabs, warn: warn, ct: ct, rt: rt };
}

function tcSummary_(p) {
  var c = {}; TC_CHILD.forEach(function (n) { c[n] = (p.child[n].miss ? '沒有這分頁' : p.child[n].del.length + '/' + p.child[n].total); });
  return {
    刪的檔期: p.camps.map(function (x) { return x.name + '｜' + x.from + '｜' + (x.status || '啟用') + '｜' + x.id; }),
    留下的檔期: p.keep.map(function (x) { return x.name + '｜' + x.from + '｜' + (x.status || '啟用') + '｜' + x.id; }),
    刪的申請表: p.reqs.map(function (x) { return x.id + '｜' + x.name + '｜' + x.status + '｜' + x.camp; }),
    留下的申請表: p.keepReqs.map(function (x) { return x.id + '｜' + x.name + '｜' + x.status + '｜' + x.camp; }),
    沒有檔期的申請表: p.orphan.length,
    相關紀錄_刪幾列_共幾列: c,
    backup分頁: p.tabs,
    警告: p.warn
  };
}

function tcReport() {
  var s = tcSummary_(tcPlan_());
  console.log(JSON.stringify(s, null, 1));
  return s;
}

/* 一張分頁刪多列：由下往上、連續的併成一次 */
function tcDelRows_(sh, rows) {
  rows = rows.slice().sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < rows.length) {
    var end = rows[i], start = end;
    while (i + 1 < rows.length && rows[i + 1] === start - 1) { i++; start = rows[i]; }
    sh.deleteRows(start, end - start + 1);
    i++;
  }
}

function tcCount_(name) { var sh = ss_().getSheetByName(name); return sh ? Math.max(sh.getLastRow() - 1, 0) : -1; }

function tcDelete() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var p = tcPlan_();
    if (p.warn.length) { console.log('不刪，有警告：' + JSON.stringify(p.warn)); return { ok: false, warn: p.warn }; }
    if (!p.camps.length) { console.log('沒有名稱含「測試」的檔期，不用刪'); return { ok: true, nothing: true }; }
    var names = ['fact_campaign', 'fact_recipe_req'].concat(TC_CHILD), before = {}, after = {};
    names.forEach(function (n) { before[n] = tcCount_(n); });
    var ss = ss_(), bkName = '食譜系統_申請單_備份_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm') + '_刪測試檔期前';
    var bk = ss.copy(bkName);   /* 整份備份（含全部分頁），放在擁有者雲端硬碟最上層 */
    SpreadsheetApp.flush();
    tcDelRows_(p.ct.sh, p.camps.map(function (x) { return x.row; }));
    tcDelRows_(p.rt.sh, p.reqs.map(function (x) { return x.row; }));
    TC_CHILD.forEach(function (n) { if (!p.child[n].miss && p.child[n].del.length) tcDelRows_(ss.getSheetByName(n), p.child[n].del); });
    p.tabs.forEach(function (n) { var sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
    SpreadsheetApp.flush();
    names.forEach(function (n) { after[n] = tcCount_(n); });
    campCacheClear_();
    try { CacheService.getScriptCache().remove('pub:newitems:v7'); } catch (e) { }
    var res = { ok: true, backup: bkName, backup_id: bk.getId(), before: before, after: after,
      camps: p.camps.length, reqs: p.reqs.length, tabs: p.tabs, left: tcSummary_(tcPlan_()) };
    log_('管理者', 'tcDelete', '', true, '刪測試檔期 ' + p.camps.length + ' 個、申請表 ' + p.reqs.length + ' 張；備份 ' + bk.getId());
    console.log(JSON.stringify(res, null, 1));
    return res;
  } finally { lock.releaseLock(); }
}

/* 刪完核對：留下來的每一列（真的檔期、申請表與相關紀錄、其他分頁）跟刪之前的整份備份逐格相同 */
var TC_BACKUP_ID = '1tTVxlmKD1serZ4zRBr-Q0rFRu_155VKTJOFvSGH1Yxk';   /* 2026-10-18 23:40 tcDelete 自動建立的整份備份 */
function tcVerify() {
  var bk = SpreadsheetApp.openById(TC_BACKUP_ID), cur = ss_(), out = {}, bad = 0;
  cur.getSheets().forEach(function (sh) {
    var n = sh.getName(), bs = bk.getSheetByName(n);
    if (!bs) { out[n] = '備份裡沒有這分頁'; bad++; return; }
    var a = sh.getDataRange().getValues(), b = bs.getDataRange().getValues();
    var key = function (r) { return JSON.stringify(r); }, bset = {};
    b.forEach(function (r) { var k = key(r); bset[k] = (bset[k] || 0) + 1; });
    var miss = a.filter(function (r) { var k = key(r); if (bset[k]) { bset[k]--; return false; } return true; }).length;
    out[n] = (miss ? '❌ ' + miss + ' 列跟備份不同' : '✔') + '｜現在 ' + (a.length - 1) + ' 列／備份 ' + (b.length - 1) + ' 列';
    if (miss) bad++;
  });
  var gone = bk.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return !cur.getSheetByName(n); });
  console.log(JSON.stringify({ bad: bad, 只在備份的分頁: gone, 各分頁: out }, null, 1));
  return bad;
}
