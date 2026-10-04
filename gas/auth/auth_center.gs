/** ═══════════════════════════════════════════════════════════════════
 * 儀表板權限中心 auth_center.gs v1.2（2026-10-04 效能第 2 批，保守版）
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
var AUTH_CACHE_KEY = 'dim_auth:v1';

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
      case 'ping':       out = { ok: true, ver: 'v1.2' }; break;
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
  var vals = sh.getRange(2, 1, last - 1, 7).getValues();
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
      note: String(v[6] || '')
    });
  }
  return arr;
}

function findDash_(id, fresh) {   /* fresh＝略過快取直接讀 Sheet（要寫入前用，列號才是現況） */
  var arr = fresh ? rows_() : rowsCached_();
  for (var i = 0; i < arr.length; i++) {
    if (arr[i].dashboard_id === id) return arr[i];
  }
  return null;
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
  if (id === '__admin__') return { ok: false, error: '不允許' };

  var cache = CacheService.getScriptCache();
  var fk = 'fail_' + id;
  var fails = Number(cache.get(fk) || 0);
  if (fails >= FAIL_LIMIT) {
    log_('verify', id, 'locked');
    return { ok: false, locked: true, error: '嘗試次數過多，請 10 分鐘後再試' };
  }

  var d = findDash_(id);
  if (!d) { log_('verify', id, 'not_found'); return { ok: false, error: '查無此儀表板設定' }; }
  if (d.enabled !== 'Y') { log_('verify', id, 'disabled'); return { ok: false, error: '此儀表板已停用' }; }

  if (d.password !== pwd) {
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
    list.push({
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
