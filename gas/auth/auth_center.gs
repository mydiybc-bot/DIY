/** ═══════════════════════════════════════════════════════════════════
 * 儀表板權限中心 auth_center.gs v2.0（2026-11-18 登入一次整站通行）
 * v2.0（經營者：「太多頁面要一直不停的切換密碼」→ 認身分、發通行證、大鑰匙開小門；程式在檔尾「v2.0 身分與通行證」）：
 *       ①新動作 login：只給密碼 → 認出是哪個身分（老闆／會計／採購經營者／新品申請表的主廚・簽核人／體驗中心／出貨中心／門市）
 *         → 回「通行證」token＋這個身分能開的頁（grants）＋記住多久（ttl_hours）。密碼本身都不用改。
 *       ②新動作 check：拿通行證回來問還有沒有效（密碼換過＝馬上失效 → 改密碼就能把人登出）。
 *       ③verify 多認兩種：pwd 是通行證（pass1.…）、或 pwd 是「更高身分」的密碼 → 也放行。
 *         原本的比對一字不變且先比，所以舊頁面、各後端（公告／專案推動／問卷／費用申請／平板）都不用改。
 *       ④權限後台 admin_*：多認通行證——老闆＝可看可改；會計＝只能看（老闆那一組密碼不給看）、可解鎖。原 __admin__ 密碼暫時照舊可用。
 *       ⑤記住多久＝三個組別列的 D 欄（grp_store／grp_center／grp_boss）；一次性 setupV20 會把它們設成 720／720／24 小時。
 *       ⑥通行證用的密鑰自動產生、放指令碼屬性 PASS_SECRET（不在程式、不在試算表）。
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
      case 'login':      out = login_(p);     break;   /* v2.0 */
      case 'check':      out = check_(p);     break;   /* v2.0 */
      case 'admin_list': out = adminList_(p); break;
      case 'admin_set':  out = adminSet_(p);  break;
      case 'admin_unlock': out = adminUnlock_(p); break;
      case 'ping':       out = { ok: true, ver: 'v2.0' }; break;
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
  if (fails >= FAIL_LIMIT && pwd.indexOf(PASS_PREFIX) !== 0) {   /* v2.0：已經登入的人（帶通行證）不受別人亂試密碼的鎖影響 */
    log_('verify', id, 'locked');
    return { ok: false, locked: true, error: '嘗試次數過多，請 10 分鐘後再試' };
  }

  var arr = rowsCached_();
  var d = findIn_(arr, id);
  if (!d) { log_('verify', id, 'not_found'); return { ok: false, error: '查無此儀表板設定' }; }
  var ef = effective_(arr, d);
  if (d.enabled !== 'Y' || ef.disabled) { log_('verify', id, 'disabled'); return { ok: false, error: '此儀表板已停用' }; }

  /* v2.0 ①通行證：不算錯誤次數（過期的通行證不該把整頁鎖住） */
  if (pwd.indexOf(PASS_PREFIX) === 0) {
    var ps = readPass_(arr, pwd);
    if (!ps.ok) { log_('verify', id, 'pass_bad'); return { ok: false, expired: true, error: '登入已失效，請重新登入' }; }
    if (grantsOf_(arr, ps.def).indexOf(id) < 0) { log_('verify', id, 'pass_deny:' + ps.def.id); return { ok: false, denied: true, error: '「' + ps.def.name + '」身分不能開這一頁' }; }
    log_('verify', id, 'ok:pass:' + ps.def.id);
    return { ok: true, ttl_hours: Math.max(1, Math.ceil((ps.exp - Date.now()) / 3600000)), ident: ps.def.id };
  }

  /* ②這一頁（或它那一組）的密碼：v1.3 原本的比對，先比、結果不變 */
  if (ef.password && ef.password === pwd) {
    cache.remove(fk);
    log_('verify', id, 'ok');
    return { ok: true, ttl_hours: d.ttl_hours };
  }

  /* v2.0 ③更高身分的密碼（大鑰匙開小門） */
  var who = identify_(arr, pwd);
  if (who && grantsOf_(arr, who.def).indexOf(id) >= 0) {
    cache.remove(fk);
    log_('verify', id, 'ok:' + who.def.id);
    return { ok: true, ttl_hours: identTtl_(arr, who.def), ident: who.def.id };
  }

  cache.put(fk, String(fails + 1), FAIL_WINDOW_SEC);
  log_('verify', id, 'wrong_pwd');
  return { ok: false, error: '密碼錯誤' };
}

/* ── 管理後台 ─────────────────────────────────────────────────── */
function checkAdmin_(p) {
  var a = findDash_('__admin__');
  return !!(a && String(p.admin_pwd || '') !== '' && String(p.admin_pwd) === a.password);
}

/* v2.0：後台權限＝'edit'（老闆的通行證，或原 __admin__ 密碼）／'view'（會計的通行證）／''（不行） */
function adminWho_(p) {
  var tok = String(p.token || '');
  if (tok.indexOf(PASS_PREFIX) === 0) {
    var ps = readPass_(rowsCached_(), tok);
    return ps.ok ? (ps.def.admin || '') : '';
  }
  return checkAdmin_(p) ? 'edit' : '';
}
var BOSS_ONLY_TEXT = '（只有老闆看得到）';

function adminList_(p) {
  var who = adminWho_(p);
  if (!who) { log_('admin_list', '-', 'deny'); return { ok: false, error: p.token ? '這個身分不能看權限後台，或登入已失效' : '管理密碼錯誤' }; }
  var list = [];
  var arr = rows_();
  var cache = CacheService.getScriptCache();
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].dashboard_id === '__admin__') continue;
    var gr = arr[i].group ? findIn_(arr, arr[i].group) : null;
    var hide = who !== 'edit' && (arr[i].dashboard_id === GRP_BOSS || arr[i].group === GRP_BOSS);   /* 會計：老闆那一組的密碼不給看 */
    list.push({
      boss_only: hide,
      group: arr[i].group,
      group_name: gr ? gr.name : '',
      group_active: !!(gr && gr.password),
      fails: Number(cache.get('fail_' + arr[i].dashboard_id) || 0),
      dashboard_id: arr[i].dashboard_id,
      name: arr[i].name,
      password: hide ? BOSS_ONLY_TEXT : arr[i].password,
      ttl_hours: arr[i].ttl_hours,
      enabled: arr[i].enabled,
      updated_at: arr[i].updated_at ? Utilities.formatDate(new Date(arr[i].updated_at), 'Asia/Taipei', 'yyyy-MM-dd HH:mm') : '',
      note: arr[i].note
    });
  }
  log_('admin_list', '-', 'ok:' + who);
  return { ok: true, list: list, mode: who };
}

function adminSet_(p) {
  var who = adminWho_(p);
  if (who !== 'edit') { log_('admin_set', String(p.dashboard || '-'), 'deny'); return { ok: false, error: who ? '只有老闆可以修改' : (p.token ? '登入已失效，請重新登入' : '管理密碼錯誤') }; }
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
  if (!adminWho_(p)) { log_('admin_unlock', String(p.dashboard || '-'), 'deny'); return { ok: false, error: p.token ? '登入已失效，請重新登入' : '管理密碼錯誤' }; }   /* v2.0：會計也可以解鎖 */
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

/* ── 2026-10-30 經營者：「不知道每個儀表板該用什麼密碼」→ 建「每頁密碼對照」分頁（編輯器執行 writeCheatSheet_v13）──
   每列一個儀表板；「用哪組」「密碼」兩欄是公式，跟著 dim_auth 的 H 欄與組別密碼自動變，不用再維護。只寫這個新分頁，不碰 dim_auth。 */
var CHEAT_V13 = [
  ['首頁（🏪 門市畫面）', 'hub', ''],
  ['首頁（🏢 體驗中心畫面）', 'hub_hq', ''],
  ['決策中心（選 1～12 號店）', 'store_1', '12 家店同一組'],
  ['決策中心（選 🏢 體驗中心）', 'task_hq', ''],
  ['決策中心「📋 月會」分頁', 'monthly', ''],
  ['門市排班工作台', 'store_1', '跟決策中心選店同一組'],
  ['採購系統', 'purchase', ''],
  ['採購系統 切「體驗中心／經營者」', 'purchase_hq', ''],
  ['訂位分析', 'reservation', ''],
  ['POS 銷售', 'pos', ''],
  ['排班分析', 'schedule', ''],
  ['Google 評論', 'reviews', ''],
  ['自己人會員', 'zijiren', ''],
  ['零用金報銷', 'pettycash', ''],
  ['損益（P&L）', 'pnl', ''],
  ['新品申請表／新食譜上傳／外國人到店服務管理', 'recipe', ''],
  ['專案推動', 'project', ''],
  ['海外代理商購買平台', 'supplier', ''],
  ['顧客問卷', 'survey', ''],
  ['門市問卷管理', 'poll', ''],
  ['公告發佈（舊管理者模式）', 'announce_pub', ''],
  ['平板外語食譜', 'tablet_i18n', '平板輸一次記 30 天'],
  ['權限後台（admin-auth）', '__admin__', '只有老闆'],
  ['口語應對訓練站', '', '員工自己的帳號，不在這張表']
];
function writeCheatSheet_v13() {
  var ss = ss_(), name = '每頁密碼對照', t = ss.getSheetByName(name) || ss.insertSheet(name, 0);
  t.clear();
  var rows = [['儀表板', '用哪組密碼（自動）', '密碼（自動帶出）', '備註', '代號']];
  CHEAT_V13.forEach(function (c, i) {
    var r = i + 2, id = 'E' + r;
    if (!c[1]) { rows.push([c[0], '—', '—', c[2], '']); return; }
    var g = 'IFERROR(VLOOKUP(' + id + ',dim_auth!A:H,8,FALSE),"")';
    rows.push([c[0],
      '=IF(' + g + '="","自己一組",IFERROR(VLOOKUP(' + g + ',dim_auth!A:B,2,FALSE),"?"))',
      '=IF(' + g + '="",IFERROR(VLOOKUP(' + id + ',dim_auth!A:C,3,FALSE),"查無"),IFERROR(VLOOKUP(' + g + ',dim_auth!A:C,3,FALSE),"查無"))',
      c[2], c[1]]);
  });
  t.getRange(1, 1, rows.length, 5).setValues(rows);
  t.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#fde68a');
  t.setFrozenRows(1);
  t.setColumnWidth(1, 330); t.setColumnWidth(2, 230); t.setColumnWidth(3, 180); t.setColumnWidth(4, 200); t.setColumnWidth(5, 110);
  t.getRange(2, 5, rows.length - 1, 1).setFontColor('#999999');
  ss.setActiveSheet(t);
  Logger.log('每頁密碼對照：' + (rows.length - 1) + ' 列（公式，會跟著 dim_auth 自動變）');
  return rows.length - 1;
}

/* ═══════════════════════════════════════════════════════════════════
 * v2.0 身分與通行證（2026-11-18 經營者：登入一次整站通行、大鑰匙開小門）
 *
 * 身分（由高到低；同一組密碼符合多個身分時取「最低」的，見 identify_）與能開的頁：
 *   🔴 老闆          ＝ grp_boss 的密碼          → 全部（含權限後台：可看可改）
 *   🟣 會計          ＝ apply_review 那一列的密碼 → 體驗中心全部＋會計後台審核＋權限後台（只能看；老闆那一組密碼不顯示）
 *   🟠 採購經營者    ＝ purchase_hq 那一列的密碼  → 體驗中心全部＋採購「體驗中心／經營者」＋出貨中心
 *   🟡 新品申請表身分＝ 食譜系統_申請單 dim_role 的身分密碼、dim_signer 的簽核人 PIN → 體驗中心全部
 *   🟡 體驗中心      ＝ grp_center 的密碼         → 🟡 體驗中心組＋🟢 門市組的頁
 *   🔵 出貨中心      ＝ purchase_dc 那一列的密碼  → 門市組的頁＋採購「出貨中心」
 *   🟢 門市          ＝ grp_store 的密碼          → 門市組的頁
 * 「哪一頁屬於哪一組」照舊看 dim_auth 的 H 欄，換組＝改 H 欄，不用改程式。
 * 會計／採購經營者／出貨中心那一列如果 H 欄填了組別（改用組別密碼）→ 就不算獨立身分。
 * 記住多久：老闆＝grp_boss 的 D 欄；門市、出貨中心＝grp_store 的 D 欄；其餘＝grp_center 的 D 欄（單位：小時）。
 *
 * 通行證 pass1.<內容>.<簽章>：內容＝身分＋到期時間＋當時密碼的指紋；簽章用只在指令碼屬性的密鑰（PASS_SECRET）。
 * 回來驗的時候：簽章不對／過期／那個身分的密碼已經換過 → 失效。能開哪些頁每次用試算表現況重算（換組、停用馬上生效）。
 * ═══════════════════════════════════════════════════════════════════ */
var PASS_PREFIX = 'pass1.';
var GRP_STORE = 'grp_store', GRP_CENTER = 'grp_center', GRP_BOSS = 'grp_boss';
var RECIPE_SHEET_ID = '10K3j3mzGz8uFlIUQI7PXz11I-3W8aTCVRnA_6jgUpIs';   /* 食譜系統_申請單（dim_role／dim_signer）；清空＝不認新品申請表身分 */
var RCP_CACHE_KEY = 'rcp_ident:v1';
var IDENT_DEF = [   /* 順序＝由高到低，identify_ 從最後面（最低）往前找 */
  { id: 'boss',  name: '🔴 老闆',           row: 'grp_boss',     ttlRow: 'grp_boss',   groups: '*',                      extra: [],                             admin: 'edit' },
  { id: 'acct',   name: '🟣 會計',           row: 'apply_review', ttlRow: 'grp_center', groups: ['grp_center', 'grp_store'], extra: ['apply_review'],               admin: 'view' },
  { id: 'buyer',  name: '🟠 採購經營者',     row: 'purchase_hq',  ttlRow: 'grp_center', groups: ['grp_center', 'grp_store'], extra: ['purchase_hq', 'purchase_dc'], admin: '' },
  { id: 'rcp',    name: '🟡 新品申請表身分', row: '',             ttlRow: 'grp_center', groups: ['grp_center', 'grp_store'], extra: [],                             admin: '' },
  { id: 'center', name: '🟡 體驗中心',       row: 'grp_center',   ttlRow: 'grp_center', groups: ['grp_center', 'grp_store'], extra: [],                             admin: '' },
  { id: 'dc',     name: '🔵 出貨中心',       row: 'purchase_dc',  ttlRow: 'grp_store',  groups: ['grp_store'],              extra: ['purchase_dc'],                admin: '' },
  { id: 'store',  name: '🟢 門市',           row: 'grp_store',    ttlRow: 'grp_store',  groups: ['grp_store'],              extra: [],                             admin: '' }
];

/* 這個身分現在的密碼（沒有＝這個身分不能用） */
function identPassword_(arr, def) {
  var r = def.row ? findIn_(arr, def.row) : null;
  if (!r || r.enabled !== 'Y' || !r.password) return '';
  if (def.row.indexOf(GROUP_PREFIX) === 0) return r.password;
  var ef = effective_(arr, r);
  if (ef.disabled || ef.group) return '';   /* 這一列已經併進某一組（用組別密碼）→ 不是獨立身分 */
  return r.password;
}

/* 這個身分能開哪些頁（dim_auth 的代號清單；用現況算） */
function grantsOf_(arr, def) {
  if (def.fixed) return def.fixed.slice();
  var out = [];
  for (var i = 0; i < arr.length; i++) {
    var r = arr[i], id = r.dashboard_id;
    if (id === '__admin__' || id.indexOf(GROUP_PREFIX) === 0 || r.enabled !== 'Y') continue;
    if (effective_(arr, r).disabled) continue;
    if (def.groups === '*' || (r.group && def.groups.indexOf(r.group) >= 0) || def.extra.indexOf(id) >= 0) out.push(id);
  }
  return out;
}

function identTtl_(arr, def) {
  var r = findIn_(arr, def.ttlRow);
  return r ? r.ttl_hours : 12;
}

/* 新品申請表的身分密碼／簽核人 PIN（快取 2 分鐘；讀不到回 null，不影響其他身分） */
function rcpRows_() {
  if (!RECIPE_SHEET_ID) return [];
  var cache = CacheService.getScriptCache(), hit = null;
  try { hit = cache.get(RCP_CACHE_KEY); } catch (_) {}
  if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  var out = [];
  try {
    var ss = SpreadsheetApp.openById(RECIPE_SHEET_ID);
    readTab_(ss, 'dim_role', function (o) {
      if (o.role && o.password) out.push({ k: 'role:' + o.role, label: o.role, pw: o.password });
    });
    readTab_(ss, 'dim_signer', function (o) {
      var off = ['N', 'NO', 'FALSE', '否', '0', '停用'].indexOf(String(o.enabled || '').toUpperCase()) >= 0;
      if (o.name && o.pin && !off) out.push({ k: 'signer:' + o.name, label: o.name, pw: o.pin });
    });
  } catch (_) { return null; }
  try { cache.put(RCP_CACHE_KEY, JSON.stringify(out), AUTH_CACHE_SEC); } catch (_) {}
  return out;
}
function readTab_(ss, name, fn) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return;
  var v = sh.getDataRange().getValues(), h = [];
  for (var j = 0; j < v[0].length; j++) h.push(String(v[0][j]).trim());
  for (var i = 1; i < v.length; i++) {
    var o = {};
    for (var k = 0; k < h.length; k++) o[h[k]] = String(v[i][k] == null ? '' : v[i][k]).trim();
    fn(o);
  }
}

/* 這組密碼是哪個身分（都不是回 null）。同一組密碼符合好幾個身分時取「最低」的那個：
   例：某位簽核人的 PIN 剛好跟門市密碼一樣 → 只算門市。大家都知道的密碼不能拿來升級（2026-11-18 在真實資料上跑自我測試時抓到）。 */
function identify_(arr, pwd) {
  if (!pwd) return null;
  for (var i = IDENT_DEF.length - 1; i >= 0; i--) {
    var def = IDENT_DEF[i];
    if (def.id === 'rcp') {
      var rr = rcpRows_() || [];
      for (var j = 0; j < rr.length; j++) {
        if (rr[j].pw === pwd) return { def: def, sub: rr[j].k, label: rr[j].label };
      }
      continue;
    }
    var pw = identPassword_(arr, def);
    if (pw && pw === pwd) return { def: def, sub: '', label: '' };
  }
  return null;
}

/* ── 通行證 ── */
var _sec = null;
function secret_() {
  if (_sec) return _sec;
  var P = PropertiesService.getScriptProperties(), s = P.getProperty('PASS_SECRET');
  if (!s) {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      s = P.getProperty('PASS_SECRET');
      if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); P.setProperty('PASS_SECRET', s); }
    } finally { lock.releaseLock(); }
  }
  _sec = s;
  return s;
}
function b64_(bytes) { return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, ''); }
function unb64_(s) { while (s.length % 4) s += '='; return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString(); }
function sign_(s) { return b64_(Utilities.computeHmacSha256Signature(String(s), secret_())).slice(0, 27); }
function fp_(pw) { return sign_('fp:' + pw).slice(0, 8); }   /* 密碼指紋：密碼一換，舊通行證就失效 */

function issue_(identId, sub, pw, exp) {
  var body = b64_(Utilities.newBlob(JSON.stringify({ i: identId, s: sub || '', x: exp, v: fp_(pw) })).getBytes());
  return PASS_PREFIX + body + '.' + sign_(body);
}

function readPass_(arr, tok) {
  try {
    var parts = String(tok).slice(PASS_PREFIX.length).split('.');
    if (parts.length !== 2 || !parts[0] || sign_(parts[0]) !== parts[1]) return { ok: false };
    var o = JSON.parse(unb64_(parts[0]));
    if (!o || !o.x || Date.now() >= Number(o.x)) return { ok: false };
    var def = null, pw = '', label = '', skip = false, i;
    if (o.i === 'page') {   /* 單頁密碼（不屬於任何身分的頁，例：平板外語食譜） */
      var d = findIn_(arr, String(o.s || ''));
      if (!d || d.enabled !== 'Y') return { ok: false };
      var ef = effective_(arr, d);
      if (ef.disabled) return { ok: false };
      pw = ef.password;
      def = { id: 'page', name: d.name, fixed: [d.dashboard_id], admin: '', ttlRow: d.dashboard_id };
    } else {
      for (i = 0; i < IDENT_DEF.length; i++) if (IDENT_DEF[i].id === o.i) def = IDENT_DEF[i];
      if (!def) return { ok: false };
      if (def.id === 'rcp') {
        var rr = rcpRows_();
        if (rr === null) skip = true;   /* 新品申請表試算表一時讀不到：簽章與效期都對 → 先放行，不把人登出 */
        else for (i = 0; i < rr.length; i++) if (rr[i].k === o.s) { pw = rr[i].pw; label = rr[i].label; }
      } else pw = identPassword_(arr, def);
    }
    if (!skip && (!pw || fp_(pw) !== o.v)) return { ok: false };
    return { ok: true, def: def, sub: String(o.s || ''), label: label, exp: Number(o.x) };
  } catch (_) { return { ok: false }; }
}

/* 這一頁要哪些身分才能開（給「權限不夠」的提示用） */
function needNames_(arr, id) {
  var names = [], hasCenter = false;
  for (var i = IDENT_DEF.length - 1; i >= 0; i--) {
    var def = IDENT_DEF[i];
    if (def.id === 'rcp') continue;
    if (!identPassword_(arr, def) || grantsOf_(arr, def).indexOf(id) < 0) continue;
    if (hasCenter && (def.id === 'acct' || def.id === 'buyer')) continue;
    if (def.id === 'center') hasCenter = true;
    names.push(def.name);
  }
  return names.join('、');
}

function passInfo_(arr, def, label, exp, tok, id, ttl) {
  var g = grantsOf_(arr, def), here = id ? g.indexOf(id) >= 0 : true;
  return {
    ok: true, ident: def.id, ident_name: def.name, label: label || '',
    ttl_hours: ttl || Math.max(1, Math.ceil((exp - Date.now()) / 3600000)), exp: exp,
    token: tok, grants: g, admin: def.admin || '',
    here: here, need: here ? '' : needNames_(arr, id)
  };
}

/* ── login：只給密碼 → 認身分、發通行證。dashboard＝在哪一頁登入（選填；用來記錄、算錯誤次數、判斷這一頁開不開得了） ── */
function login_(p) {
  var pwd = String(p.pwd || ''), id = String(p.dashboard || '').trim();
  if (!pwd) return { ok: false, error: '缺少參數' };
  if (pwd.indexOf(PASS_PREFIX) === 0) return check_({ token: pwd, dashboard: id });
  if (id === '__admin__' || id.indexOf(GROUP_PREFIX) === 0) id = '';

  var cache = CacheService.getScriptCache();
  var fk = 'fail_' + (id || '_login');
  var fails = Number(cache.get(fk) || 0);
  if (fails >= FAIL_LIMIT) {
    log_('login', id || '-', 'locked');
    return { ok: false, locked: true, error: '嘗試次數過多，請 10 分鐘後再試' };
  }

  var arr = rowsCached_();
  var who = identify_(arr, pwd), ttl, exp;
  /* own＝這組密碼剛好就是「這一頁自己（或它那一組）的密碼」。舊規則「這一頁的密碼一定開得了這一頁」要保住：
     ①不屬於任何身分的頁（平板外語食譜）②密碼撞到較低身分、認出來的身分開不了這一頁 → 發只開這一頁的通行證 */
  var own = null;
  if (id) {
    var d = findIn_(arr, id);
    if (d && d.enabled === 'Y') {
      var ef = effective_(arr, d);
      if (!ef.disabled && ef.password && ef.password === pwd) own = d;
    }
  }
  if (who && !(own && grantsOf_(arr, who.def).indexOf(id) < 0)) {
    ttl = identTtl_(arr, who.def); exp = Date.now() + ttl * 3600000;
    cache.remove(fk);
    log_('login', id || '-', 'ok:' + who.def.id);
    return passInfo_(arr, who.def, who.label, exp, issue_(who.def.id, who.sub, pwd, exp), id, ttl);
  }
  if (own) {
    ttl = own.ttl_hours; exp = Date.now() + ttl * 3600000;
    cache.remove(fk);
    log_('login', id, 'ok:page');
    return passInfo_(arr, { id: 'page', name: own.name, fixed: [id], admin: '' }, '', exp, issue_('page', id, pwd, exp), id, ttl);
  }

  cache.put(fk, String(fails + 1), FAIL_WINDOW_SEC);
  log_('login', id || '-', 'wrong_pwd');
  return { ok: false, error: '密碼錯誤' };
}

/* ── check：通行證還有沒有效（不寫登入紀錄、不算錯誤次數） ── */
function check_(p) {
  var tok = String(p.token || ''), id = String(p.dashboard || '').trim();
  if (tok.indexOf(PASS_PREFIX) !== 0) return { ok: false, error: '缺少參數' };
  var arr = rowsCached_(), ps = readPass_(arr, tok);
  if (!ps.ok) return { ok: false, expired: true, error: '登入已失效，請重新登入' };
  return passInfo_(arr, ps.def, ps.label, ps.exp, tok, id, 0);
}

/* ── v2.0 一次性（編輯器選 setupV20 → 執行）：先備份 dim_auth → 把三個組別列的「記住多久」設成 門市 720／體驗中心 720／老闆 24 小時
      → 產生通行證密鑰 → 印出每個身分「有沒有密碼、能開幾頁」（不印密碼）。重跑安全。 ── */
var TTL_V20 = { grp_store: 720, grp_center: 720, grp_boss: 24 };
function setupV20() {
  var ss = ss_(), sh = ss.getSheetByName(TAB_AUTH), arr = rows_();
  if (!arr.length) throw new Error('dim_auth 是空的，停止');
  var ts = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd-HHmm');
  ss.getSheets().forEach(function (t) { if (/^dim_auth_備份_/.test(t.getName())) ss.deleteSheet(t); });
  var bk = sh.copyTo(ss).setName('dim_auth_備份_' + ts);
  var done = [], miss = [];
  Object.keys(TTL_V20).forEach(function (g) {
    var r = findIn_(arr, g);
    if (!r) { miss.push(g); return; }
    sh.getRange(r.rowIndex, COL.ttl).setValue(TTL_V20[g]);
    sh.getRange(r.rowIndex, COL.updated).setValue(new Date());
    done.push(g + '＝' + TTL_V20[g] + ' 小時（原 ' + r.ttl_hours + '）');
  });
  authCacheClear_();
  try { CacheService.getScriptCache().remove(RCP_CACHE_KEY); } catch (_) {}
  secret_();
  var after = rows_(), rep = [];
  IDENT_DEF.forEach(function (def) {
    if (def.id === 'rcp') { var rr = rcpRows_(); rep.push(def.name + '：' + (rr === null ? '讀不到 食譜系統_申請單' : rr.length + ' 組密碼／PIN')); return; }
    rep.push(def.name + '：' + (identPassword_(after, def) ? '有密碼，能開 ' + grantsOf_(after, def).length + ' 頁' : '沒有獨立密碼（這個身分暫時不能用）'));
  });
  var msg = '完成。備份分頁：' + bk.getName() + '｜記住多久：' + (done.join('、') || '無') + (miss.length ? '｜找不到：' + miss.join('、') : '') +
            '｜前 ' + arr.length + ' 列 → 後 ' + after.length + ' 列｜' + rep.join('｜');
  Logger.log(msg);
  return msg;
}

/* 驗證：每個身分用表上的密碼登入一次、拿通行證逐頁驗；只印對不對，不印密碼、不留登入紀錄 */
function selfTestV20() {
  authCacheClear_();
  var arr = rows_(), bad = [], n = 0, lines = [], cache = CacheService.getScriptCache();
  var pages = arr.filter(function (d) { return d.dashboard_id !== '__admin__' && d.dashboard_id.indexOf(GROUP_PREFIX) !== 0 && d.enabled === 'Y' && !effective_(arr, d).disabled; });
  IDENT_DEF.forEach(function (def) {
    if (def.id === 'rcp') return;
    var pw = identPassword_(arr, def);
    if (!pw) { lines.push(def.name + '＝沒有密碼，略過'); return; }
    var lg = login_({ pwd: pw });
    if (!lg.ok) { bad.push(def.id + ' 登入失敗'); return; }
    var top = identify_(arr, pw).def;   /* 密碼跟更低身分相同時，會被認成較低的那個 */
    var g = grantsOf_(arr, top), miss = 0;
    pages.forEach(function (d) {
      var want = g.indexOf(d.dashboard_id) >= 0;
      var r1 = verify_({ dashboard: d.dashboard_id, pwd: lg.token });
      var r2 = verify_({ dashboard: d.dashboard_id, pwd: pw });
      var own = effective_(arr, d).password === pw;
      if (!!r1.ok !== want) { miss++; bad.push(def.id + '→' + d.dashboard_id + '（通行證）'); }
      if (!!r2.ok !== (want || own)) { miss++; bad.push(def.id + '→' + d.dashboard_id + '（密碼）'); }
      cache.remove('fail_' + d.dashboard_id);
      n += 2;
    });
    var fake = verify_({ dashboard: pages[0].dashboard_id, pwd: lg.token.slice(0, -2) + 'xx' });
    if (fake.ok) bad.push(def.id + ' 假通行證被放行');
    lines.push(def.name + (top.id !== def.id ? '（密碼與「' + top.name + '」相同，只算後者）' : '') + '＝能開 ' + g.length + '／' + pages.length + ' 頁' + (miss ? '，不符 ' + miss : ''));
  });
  /* 新品申請表裡，密碼／PIN 跟某個身分密碼一樣的（只印是誰、跟哪個身分一樣，不印密碼）→ 那一位只會拿到較低的身分 */
  var dup = [];
  (rcpRows_() || []).forEach(function (x) {
    var w = identify_(arr, x.pw);
    if (w && w.def.id !== 'rcp') dup.push('「' + x.label + '」＝' + w.def.name);
    else if (w && w.sub !== x.k) dup.push('「' + x.label + '」＝「' + w.label + '」');
  });
  lines.push('新品申請表密碼與別人相同：' + (dup.join('、') || '無'));
  cache.remove('fail__login');
  _logQ = [];
  var msg = 'v2.0 自我測試：驗了 ' + n + ' 次｜不符：' + (bad.join('、') || '無') + '｜' + lines.join('｜');
  Logger.log(msg);
  return msg;
}
