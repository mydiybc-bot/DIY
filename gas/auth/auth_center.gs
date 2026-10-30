/** ═══════════════════════════════════════════════════════════════════
 * 儀表板權限中心 auth_center.gs v1.3（2026-10-30 分組密碼）
 * v1.3：①dim_auth 加 H 欄 group：填了組別代號（grp_store／grp_center／grp_boss）的列，改用「那一組」的密碼驗證；
 *         組別列密碼空白或找不到 → 照舊用自己這列的密碼（不會突然全部進不去）；組別列 enabled＝N → 整組停用
 *       ②組別列本身不能直接當儀表板驗（verify 回「不允許」）③admin_list 多回 group／group_name，後台組員列顯示「用哪一組」
 *       ④一次性函式 setupGroups_v13（編輯器執行）：先備份 dim_auth → 加三組 → 分組 → 刪不再使用的列 → 建「密碼放哪裡」分頁；
 *         密碼只在程式裡複製，不顯示、不寫進紀錄 ⑤selfTestGroups_v13：用表上的密碼逐列驗一次，只記 ok／不 ok
 * v1.2：①dim_auth 讀取結果放 CacheService 2 分鐘（verify／checkAdmin 用；admin_set／admin_unlock 成功後清掉 → 後台改密碼、停用立刻生效；
 *         直接在 Sheet 手改最多 2 分鐘後生效）②登入紀錄照寫、不拿掉，但改成「回應組好之後最後一步」寫，且 try/catch 包住（寫失敗不影響登入）
 *       ③試算表只開一次（ss_ 共用）④回傳格式一字不變（含 ttl_hours）；ping 回 v1.2
 * v1.1：①seed 納入 pnl／monthly（財務頁 TTL 4h，取代 hub 寫死的 PWD_B）
 *       ②新增 admin_unlock（後台一鍵解除防暴力鎖定）③admin_list 回傳當前錯誤數
 *       ※ 若已用 v1 執行過 seedAuthSheet：不必重灌，手動在 dim_auth 加
 *         pnl／monthly 兩列即可（GAS 動態讀表，無需改碼重部署以外動作）。
 * 用途：10 個儀表板「每頁獨立密碼」集中管理。密碼存 Sheet（僅本人可開），
 *       不再寫死在 public repo 前端；改密碼／停用即時生效（新登入立刻套用）。
 *
 * ── 本人建置 SOP（約 15 分鐘，一次性）──────────────────────────────
 * 1. 新建 Google Sheet，命名「儀表板權限中心」，複製其 Sheet ID 貼到下方 SHEET_ID。
 * 2. script.google.com 新建「獨立」專案（非 Sheet 綁定，方便日後備份），
 *    命名「儀表板權限中心」，整檔貼上本內容（一次貼齊，勿逐字打）。
 * 3. 編輯器內選 seedAuthSheet → 執行一次（授權）→ 自動建 dim_auth / auth_log
 *    兩分頁並灌 10 列初始資料（密碼皆為 CHANGE_ME_xx 佔位）。
 * 4. 回 Sheet 把 C 欄密碼逐列改成正式密碼（__admin__ 列＝管理後台密碼，最重要）。
 * 5. 部署 → 新增部署 → 網頁應用程式：執行身分＝我(mydiybc@gmail.com)、
 *    誰可以存取＝所有人 → 複製部署 URL。
 * 6. 瀏覽器直開  部署URL?action=ping  應回 callback({"ok":true,"ver":"v1"})。
 * 7. 把 URL 填入 admin-auth.html 的 AUTH_API 後上傳 GitHub；同一 URL 交給
 *    Cowork（第二批交辦書 v3 的閘門要用）。
 * ※ 日後改本檔：照 GAS 部署鐵則「管理部署作業→編輯→新版本」保留 URL。
 *
 * ── dim_auth 欄位 ────────────────────────────────────────────────
 * A dashboard_id｜B name｜C password｜D ttl_hours｜E enabled(Y/N)｜F updated_at｜G note
 * __admin__ 列＝管理後台密碼；hub 列 TTL 由 hub 前端固定 4h（見 note）。
 *
 * ── 安全等級（誠實聲明）─────────────────────────────────────────
 * 密碼驗證在伺服器端、不進 repo、可集中改與停用；但各儀表板「資料端點」
 * 仍為公開（gviz／GAS 所有人），故本系統屬「防路人直連＋權限收放」等級，
 * 非資料層資安。防暴力：同一儀表板 10 分鐘內錯 15 次即暫鎖 10 分鐘。
 * ═══════════════════════════════════════════════════════════════════ */

var SHEET_ID = '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4';   // ★ 建置 SOP 第 1 步
var TAB_AUTH = 'dim_auth';
var TAB_LOG  = 'auth_log';
var FAIL_LIMIT = 15;        // 10 分鐘內同一儀表板密碼錯誤上限
var FAIL_WINDOW_SEC = 600;

var COL = { id:1, name:2, password:3, ttl:4, enabled:5, updated:6, note:7 };
var AUTH_CACHE_SEC = 120;            /* v1.2：dim_auth 讀取結果快取秒數（admin_set／admin_unlock 後清掉） */
var AUTH_CACHE_KEY = 'dim_auth:v2';   /* v1.3：多了 group 欄，換 key 免讀到舊格式 */
var GROUP_PREFIX = 'grp_';

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = p.callback || 'callback';
  var out;
  try {
    switch (p.action) {
      case 'verify':     out = verify_(p);    break;
      case 'admin_list': out = adminList_(p); break;
      case 'admin_set':  out = adminSet_(p);  break;
      case 'admin_unlock': out = adminUnlock_(p); break;
      case 'ping':       out = { ok: true, ver: 'v1.3' }; break;
      default:           out = { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  var body = cb + '(' + JSON.stringify(out) + ')';
  flushLog_();   /* v1.2：登入紀錄放最後一步寫（回應已經組好；寫失敗只是少一列紀錄，不影響登入結果） */
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ── 內部工具 ─────────────────────────────────────────────────── */
var _ss = null;
function ss_() { if (!_ss) _ss = SpreadsheetApp.openById(SHEET_ID); return _ss; }   /* v1.2：同一次請求只開一次 */

/* v1.2：dim_auth 快取（2 分鐘）。verify／checkAdmin 走這裡；adminList／adminSet 仍每次讀最新（後台要看到、要寫到的列號必須是現況）。
   注意：Date 經 JSON 來回會變成字串，所以快取回來的 updated_at 只給 new Date(…) 用（adminList 不走快取，不受影響） */
function rowsCached_() {
  var cache = CacheService.getScriptCache(), hit = null;
  try { hit = cache.get(AUTH_CACHE_KEY); } catch (_) {}
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var arr = rows_();
  try { cache.put(AUTH_CACHE_KEY, JSON.stringify(arr), AUTH_CACHE_SEC); } catch (_) {}
  return arr;
}
function authCacheClear_() { try { CacheService.getScriptCache().remove(AUTH_CACHE_KEY); } catch (_) {} }

function rows_() {
  var sh = ss_().getSheetByName(TAB_AUTH);
  if (!sh) throw new Error('找不到 ' + TAB_AUTH + ' 分頁，請先執行 seedAuthSheet');
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 8).getValues();
  var arr = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0]) continue;
    arr.push({
      rowIndex: i + 2,
      dashboard_id: String(v[0]).trim(),
      name: String(v[1] || ''),
      password: String(v[2] || ''),
      ttl_hours: Number(v[3]) || 12,
      enabled: String(v[4] || 'Y').toUpperCase(),
      updated_at: v[5] || '',
      note: String(v[6] || ''),
      group: String(v[7] || '').trim()
    });
  }
  return arr;
}

function findDash_(id, fresh) {   /* fresh＝略過快取直接讀 Sheet（要寫入前用，列號才是現況） */
  return findIn_(fresh ? rows_() : rowsCached_(), id);
}
function findIn_(arr, id) {
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].dashboard_id === id) return arr[i];
  }
  return null;
}

/* v1.3：這一列實際要比對的密碼與是否停用。有分組、而且組別列有密碼 → 用組別的；否則用自己的 */
function effective_(arr, d) {
  var g = d.group ? findIn_(arr, d.group) : null;
  if (g && g.enabled !== 'Y') return { password: d.password, disabled: true, group: g };
  if (g && g.password) return { password: g.password, disabled: false, group: g };
  return { password: d.password, disabled: false, group: null };
}

/* v1.2：log_ 先排隊，doGet 組好回應後 flushLog_ 一次寫（照寫、不拿掉——稽核密碼外流要靠它） */
var _logQ = [];
function log_(action, dash, result) {
  _logQ.push([new Date(), action, dash, result]);
}
function flushLog_() {
  if (!_logQ.length) return;
  var q = _logQ; _logQ = [];
  try {
    var sh = ss_().getSheetByName(TAB_LOG);
    if (!sh) return;
    for (var i = 0; i < q.length; i++) sh.appendRow(q[i]);
  } catch (_) {}
}

/* ── 各儀表板密碼驗證 ─────────────────────────────────────────── */
function verify_(p) {
  var id = String(p.dashboard || '').trim();
  var pwd = String(p.pwd || '');
  if (!id || !pwd) return { ok: false, error: '缺少參數' };
  if (id === '__admin__' || id.indexOf(GROUP_PREFIX) === 0) return { ok: false, error: '不允許' };

  var cache = CacheService.getScriptCache();
  var fk = 'fail_' + id;
  var fails = Number(cache.get(fk) || 0);
  if (fails >= FAIL_LIMIT) {
    log_('verify', id, 'locked');
    return { ok: false, locked: true, error: '嘗試次數過多，請 10 分鐘後再試' };
  }

  var arr = rowsCached_();
  var d = findIn_(arr, id);
  if (!d) { log_('verify', id, 'not_found'); return { ok: false, error: '查無此儀表板設定' }; }
  var ef = effective_(arr, d);
  if (d.enabled !== 'Y' || ef.disabled) { log_('verify', id, 'disabled'); return { ok: false, error: '此儀表板已停用' }; }

  if (!ef.password || ef.password !== pwd) {
    cache.put(fk, String(fails + 1), FAIL_WINDOW_SEC);
    log_('verify', id, 'wrong_pwd');
    return { ok: false, error: '密碼錯誤' };
  }

  cache.remove(fk);
  log_('verify', id, 'ok');
  return { ok: true, ttl_hours: d.ttl_hours };
}

/* ── 管理後台 ─────────────────────────────────────────────────── */
function checkAdmin_(p) {
  var a = findDash_('__admin__');
  return !!(a && String(p.admin_pwd || '') !== '' && String(p.admin_pwd) === a.password);
}

function adminList_(p) {
  if (!checkAdmin_(p)) { log_('admin_list', '-', 'deny'); return { ok: false, error: '管理密碼錯誤' }; }
  var list = [];
  var arr = rows_();
  var cache = CacheService.getScriptCache();
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].dashboard_id === '__admin__') continue;
    var gr = arr[i].group ? findIn_(arr, arr[i].group) : null;
    list.push({
      group: arr[i].group,
      group_name: gr ? gr.name : '',
      group_active: !!(gr && gr.password),
      fails: Number(cache.get('fail_' + arr[i].dashboard_id) || 0),
      dashboard_id: arr[i].dashboard_id,
      name: arr[i].name,
      password: arr[i].password,
      ttl_hours: arr[i].ttl_hours,
      enabled: arr[i].enabled,
      updated_at: arr[i].updated_at ? Utilities.formatDate(new Date(arr[i].updated_at), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
      note: arr[i].note
    });
  }
  log_('admin_list', '-', 'ok');
  return { ok: true, list: list };
}

function adminSet_(p) {
  if (!checkAdmin_(p)) { log_('admin_set', String(p.dashboard || '-'), 'deny'); return { ok: false, error: '管理密碼錯誤' }; }
  var id = String(p.dashboard || '').trim();
  var d = findDash_(id, true);   /* 要寫入：直接讀 Sheet 拿現況列號，不用快取 */
  if (!d) return { ok: false, error: '查無 ' + id };

  var sh = ss_().getSheetByName(TAB_AUTH);
  if (p.new_pwd !== undefined && String(p.new_pwd) !== '') {
    sh.getRange(d.rowIndex, COL.password).setNumberFormat('@').setValue(String(p.new_pwd));
  }
  if (p.ttl_hours !== undefined && Number(p.ttl_hours) > 0) {
    sh.getRange(d.rowIndex, COL.ttl).setValue(Number(p.ttl_hours));
  }
  if (p.enabled !== undefined && String(p.enabled) !== '') {
    sh.getRange(d.rowIndex, COL.enabled).setValue(String(p.enabled).toUpperCase() === 'Y' ? 'Y' : 'N');
  }
  sh.getRange(d.rowIndex, COL.updated).setValue(new Date());
  authCacheClear_();   /* 改密碼、TTL、停用 → 立刻生效（下一次 verify 重讀 Sheet） */
  log_('admin_set', id, 'ok');
  return { ok: true };
}

function adminUnlock_(p) {
  if (!checkAdmin_(p)) { log_('admin_unlock', String(p.dashboard || '-'), 'deny'); return { ok: false, error: '管理密碼錯誤' }; }
  var id = String(p.dashboard || '').trim();
  if (!id) return { ok: false, error: '缺少參數' };
  CacheService.getScriptCache().remove('fail_' + id);
  authCacheClear_();   /* 順手清 dim_auth 快取（直接在 Sheet 手改密碼後，按一次「解鎖」也能立刻生效） */
  log_('admin_unlock', id, 'ok');
  return { ok: true };
}

/* ── 一次性初始化（編輯器手動執行）────────────────────────────── */
function seedAuthSheet() {
  var ss = ss_();
  var sh = ss.getSheetByName(TAB_AUTH);
  if (!sh) sh = ss.insertSheet(TAB_AUTH);
  if (sh.getLastRow() > 1) throw new Error(TAB_AUTH + ' 已有資料，為防覆蓋不重灌；如需重建請先手動清空。');

  var rows = [
    ['dashboard_id', 'name', 'password', 'ttl_hours', 'enabled', 'updated_at', 'note'],
    ['__admin__',   '管理後台密碼', 'CHANGE_ME_00', 12, 'Y', new Date(), '★ admin-auth.html 登入用，務必先改'],
    ['hub',         '主入口 hub',   'CHANGE_ME_01', 4,  'Y', new Date(), 'TTL 由 hub 前端固定 4h，此欄不生效'],
    ['pnl',         'P&L 損益',     'CHANGE_ME_10', 4,  'Y', new Date(), '財務頁，TTL 建議短'],
    ['monthly',     '月會模式',     'CHANGE_ME_11', 4,  'Y', new Date(), '財務頁，TTL 建議短'],
    ['pos',         'POS 銷售',     'CHANGE_ME_02', 12, 'Y', new Date(), ''],
    ['zijiren',     '自己人會員',   'CHANGE_ME_03', 12, 'Y', new Date(), ''],
    ['schedule',    '排班分析',     'CHANGE_ME_04', 12, 'Y', new Date(), ''],
    ['reviews',     'Google 評論',  'CHANGE_ME_05', 12, 'Y', new Date(), ''],
    ['budget',      '2026 預算',    'CHANGE_ME_06', 12, 'Y', new Date(), ''],
    ['vendor',      '廠商單價',     'CHANGE_ME_07', 12, 'Y', new Date(), ''],
    ['purchase',    '採購系統',     'CHANGE_ME_08', 12, 'Y', new Date(), '店長常用，密碼另行告知各店'],
    ['reservation', '訂位分析',     'CHANGE_ME_09', 12, 'Y', new Date(), '']
  ];
  sh.getRange(1, 1, rows.length, 7).setNumberFormats(
    (function(){ var f=[]; for (var r=0;r<rows.length;r++){ f.push(['@','@','@','0','@','yyyy-mm-dd hh:mm','@']); } return f; })()
  );
  sh.getRange(1, 1, rows.length, 7).setValues(rows);
  sh.setFrozenRows(1);

  var lg = ss.getSheetByName(TAB_LOG);
  if (!lg) lg = ss.insertSheet(TAB_LOG);
  if (lg.getLastRow() === 0) lg.appendRow(['timestamp', 'action', 'dashboard', 'result']);

  Logger.log('seedAuthSheet 完成：dim_auth 10 列＋auth_log 已建立。請回 Sheet 改 C 欄正式密碼。');
}
/* v1.1：pnl／monthly 已納入本系統（第二批 v3.1 交辦書同步移除 hub 的 PWD_B）。 */
/* v1.2（2026-10-04）：dim_auth 快取 2 分鐘＋登入紀錄最後一步寫＋試算表只開一次；回傳格式不變。回滾＝管理部署作業 → 編輯 → 版本選回上一版。 */

/* ── v1.3 一次性：分組密碼（2026-10-30 經營者裁定「照建議改」）── 編輯器選 setupGroups_v13 → 執行 ──
   ①整張 dim_auth 複製成「dim_auth_備份_YYYYMMDD-HHmm」（舊的 dim_auth_備份_* 刪掉，只留最新一份）
   ②加三組：grp_store 門市（密碼＝原 hub 那組，門市照舊不用改）、grp_center 體驗中心（＝原 hub_hq 那組）、grp_boss 老闆（＝原 pnl 那組）
   ③H 欄寫每列屬於哪一組；刪掉已不使用的列（店長交辦 12、區主管交辦 4、舊預算頁）
   ④建「密碼放哪裡」分頁（只寫位置、不寫密碼）
   密碼全在程式裡複製，Logger 只印代號與「改了／沒改」。重跑安全：已有組別列就不重加、不覆蓋組別密碼。 */
var GROUPS_V13 = [
  ['grp_store',  '🟢 門市（店長＋夥伴）共用密碼', 'hub',    '門市照舊用這組；店長或夥伴離職時只改這一格'],
  ['grp_center', '🟡 體驗中心共用密碼',           'hub_hq', '體驗中心同事、區主管用；改這一格＝下面所有 🟡 頁一起換'],
  ['grp_boss',   '🔴 老闆專用密碼（財務）',       'pnl',    '只有老闆；損益、月會']
];
var MEMBERS_V13 = {
  grp_store:  ['hub', 'store_1', 'store_2', 'store_3', 'store_4', 'store_5', 'store_6', 'store_7', 'store_8', 'store_9', 'store_10', 'store_11', 'store_12',
               'purchase', 'reservation', 'pettycash', 'announce'],
  grp_center: ['hub_hq', 'task_hq', 'purchase_hq', 'pos', 'schedule', 'zijiren', 'reviews', 'recipe', 'project', 'supplier', 'survey', 'poll', 'announce_pub'],
  grp_boss:   ['pnl', 'monthly']
};
var DEAD_V13 = ['budget', 'task_zone_n1', 'task_zone_n2', 'task_zone_n3', 'task_zone_cs'];
for (var _n = 1; _n <= 12; _n++) DEAD_V13.push('task_store_' + _n);

function setupGroups_v13() {
  var ss = ss_(), sh = ss.getSheetByName(TAB_AUTH);
  var before = rows_();
  if (!before.length) throw new Error('dim_auth 是空的，停止');
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  ss.getSheets().forEach(function (t) { if (/^dim_auth_備份_/.test(t.getName())) ss.deleteSheet(t); });
  var bk = sh.copyTo(ss).setName('dim_auth_備份_' + ts);

  var byId = {}; before.forEach(function (r) { byId[r.dashboard_id] = r; });
  GROUPS_V13.forEach(function (g) {
    if (!byId[g[0]] && !(byId[g[2]] && byId[g[2]].password)) throw new Error('來源列 ' + g[2] + ' 沒有密碼，停止（什麼都沒改，備份分頁可刪）');
  });
  var groupOf = {};
  Object.keys(MEMBERS_V13).forEach(function (g) { MEMBERS_V13[g].forEach(function (id) { groupOf[id] = g; }); });

  var now = new Date(), out = [], used = {};
  function push(r, grp) { out.push([r.dashboard_id, r.name, r.password, r.ttl_hours, r.enabled, r.updated_at || now, r.note, grp || '']); used[r.dashboard_id] = 1; }
  if (byId['__admin__']) push(byId['__admin__'], '');
  GROUPS_V13.forEach(function (g) {
    var ex = byId[g[0]];
    if (ex) push(ex, '');
    else { out.push([g[0], g[1], byId[g[2]].password, 12, 'Y', now, g[3], '']); used[g[0]] = 1; }
  });
  GROUPS_V13.forEach(function (g) {
    MEMBERS_V13[g[0]].forEach(function (id) { if (byId[id]) push(byId[id], g[0]); });
  });
  var dropped = [];
  before.forEach(function (r) {
    if (used[r.dashboard_id]) return;
    if (DEAD_V13.indexOf(r.dashboard_id) >= 0) { dropped.push(r.dashboard_id); return; }
    push(r, '');   /* 沒列在任何一組的（平板外語食譜、總分類帳通行碼、以後新加的）照舊用自己的密碼 */
  });

  var n = out.length, lastOld = sh.getLastRow();
  sh.getRange(1, 1, 1, 8).setValues([['dashboard_id', 'name', 'password', 'ttl_hours', 'enabled', 'updated_at', 'note', 'group']]);
  if (lastOld > 1) sh.getRange(2, 1, lastOld - 1, 8).clearContent();
  sh.getRange(2, 3, n, 1).setNumberFormat('@');
  sh.getRange(2, 1, n, 8).setValues(out);
  authCacheClear_();

  /* 前後比對：每個留下來的代號，實際生效的密碼有沒有變（只印代號） */
  var after = rows_(), changed = [], same = [];
  before.forEach(function (r) {
    if (dropped.indexOf(r.dashboard_id) >= 0 || r.dashboard_id === '__admin__') return;
    var a = findIn_(after, r.dashboard_id);
    if (!a) throw new Error('少了 ' + r.dashboard_id + '，請從備份分頁還原');
    (effective_(after, a).password === r.password ? same : changed).push(r.dashboard_id);
  });
  writeWhereTab_();
  var msg = '完成。備份分頁：' + bk.getName() + '｜前 ' + before.length + ' 列 → 後 ' + after.length + ' 列（含 3 組別列）｜刪掉 ' + dropped.length + ' 列：' + dropped.join('、') +
            '｜密碼換成組別的：' + (changed.join('、') || '無') + '｜沒變：' + same.length + ' 列';
  Logger.log(msg);
  return msg;
}

/* 驗證：用表上的密碼，每一列都走一次真正的 verify 判斷（不寫登入紀錄、不顯示密碼） */
function selfTestGroups_v13() {
  authCacheClear_();
  var arr = rows_(), bad = [], ok = 0;
  arr.forEach(function (d) {
    var id = d.dashboard_id;
    if (id === '__admin__' || id.indexOf(GROUP_PREFIX) === 0 || d.enabled !== 'Y') return;
    var ef = effective_(arr, d);
    if (!ef.password) { bad.push(id + '（沒有密碼）'); return; }
    var r1 = verify_({ dashboard: id, pwd: ef.password });
    var r2 = verify_({ dashboard: id, pwd: ef.password + '#x' });
    if (r1.ok && !r2.ok) ok++; else bad.push(id);
    CacheService.getScriptCache().remove('fail_' + id);   /* 故意錯的那一次不算進防暴力次數 */
  });
  var g = verify_({ dashboard: 'grp_store', pwd: 'x' });
  _logQ = [];   /* 測試不留登入紀錄 */
  var msg = '自我測試：通過 ' + ok + ' 列｜不通過：' + (bad.join('、') || '無') + '｜組別列直接驗＝' + (g.error || '?');
  Logger.log(msg);
  return msg;
}

function writeWhereTab_() {
  var ss = ss_(), name = '密碼放哪裡', t = ss.getSheetByName(name) || ss.insertSheet(name, 0);
  var rows = [
    ['種類', '放在哪裡', '怎麼查／怎麼改', '備註'],
    ['儀表板密碼', '本試算表 dim_auth 分頁', '看最上面三組：🟢 門市、🟡 體驗中心、🔴 老闆。改一格＝那一組的頁全部換', 'H 欄 group 寫每頁屬於哪一組；要把某頁換組，改它的 H 欄即可（最多 2 分鐘生效）'],
    ['權限後台密碼', '本試算表 dim_auth 的 __admin__ 列', 'admin-auth.html 登入用', '只有老闆'],
    ['平板外語食譜', 'dim_auth 的 tablet_i18n 列', '門市平板輸入一次記 30 天', '是平板用的，不跟人分組；改了平板要重輸'],
    ['外部網站帳密（104、Shopline、WACA、肚肚、食譜後台、LINE Pay…）', 'Google 密碼管理工具', 'passwords.google.com（用 mydiybc@gmail.com 登入）', '登入網站時讓 Chrome 記住即可'],
    ['總分類帳通行碼', 'Render diybc-ledger 環境變數 UPDATE_TOKEN', 'Claude 管；本機 ~/.config/diybc/secrets.env 也有一份', 'dim_auth 的 ledger_update 列只是紀錄'],
    ['程式用的金鑰（Claude API 等）', '本機 ~/.config/diybc/secrets.env、Render 環境變數、各 GAS 指令碼屬性', 'Claude 管，老闆不用記', '這裡只寫位置、不寫密碼'],
    ['新品申請表裡的主廚／簽核人密碼', '食譜系統_申請單 試算表 dim_role、dim_signer', '新品申請表 🔑 密碼管理', '跟本表分開管'],
    ['忘記的時候', '用 Google 帳號打開本試算表', '或跟 Claude 說「X 的密碼在哪」', '首頁體驗中心畫面有「🔑 密碼總表」按鈕']
  ];
  t.clear();
  t.getRange(1, 1, rows.length, 4).setValues(rows);
  t.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#fde68a');
  t.setFrozenRows(1);
  t.setColumnWidths(1, 4, 280);
  t.getRange(1, 1, rows.length, 4).setWrap(true).setVerticalAlignment('top');
}
