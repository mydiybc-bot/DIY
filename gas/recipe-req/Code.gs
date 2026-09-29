/**
 * 食譜系統申請單 API（recipe-req-v1）
 * 專案：「食譜系統申請單 API」（mydiybc，clasp 建立；原始碼備份在 DIY repo gas/recipe-req/）
 * 資料：Google 試算表「食譜系統_申請單」（ID 存在指令碼屬性 SHEET_ID；只有 mydiybc 能開，不開連結檢視）
 * 用途：食譜系統 dashboard-recipe.html「新品申請表」的檔期、送件、載回（A 階段，2026-09-28）；多人填寫與簽核（B 階段，2026-09-29）；
 *       核准後寫入採購系統 BOM 本（C 階段，2026-09-29：產品名稱對照表／BOM表／dim_sku，只新增列、不改既有列）
 *
 * 規則
 *   - doGet（JSONP，公開、不含申請單內容）：ping／listCampaigns／signInfo（簽核連結用：只回狀態與簽核人名字）／
 *       newItemsPub（D 階段：採購系統「🆕 檔期新品」用，只回已核准新品的品項名、貼紙名稱、購買連結、供應商覆寫、提供方式，不含價格與用量）
 *   - doPost（JSON 字串，Content-Type text/plain，前端直接讀回 {ok, data|msg}）：
 *       每筆帶 role＋password（或 signer＋pin）；伺服器依 dim_role.fields 過濾可寫欄位；
 *       全部包 LockService 10 秒；成功、失敗都寫 log 分頁。
 *   - 申請單內容（.json v3 字串）存 payload_1～4，每格 ≤ 45,000 字（試算表單格上限 50,000 字）
 *   - 本專案開兩個試算表：「食譜系統_申請單」（讀寫）；BOM 本（只在核准時新增列、管理者還原時刪自己寫的列）。不碰儀表板專用檔、P&L、排班。
 *
 * 部署：網頁應用程式／執行身分＝我／存取＝所有人。之後更新一律「管理部署作業 → 編輯 → 新版本」（網址不變）。
 * 第一次使用：編輯器選 setup → 執行 → 授權（建立試算表、分頁、表頭、各角色初始密碼）。
 */

var VERSION = 'recipe-req-v5';   /* v4＝D 階段：公開查詢 newItemsPub；v5＝E 階段：廠商品名 vname */   /* v2＝B 階段：各單位局部填寫、送簽、簽核；v3＝C 階段：核准 → 寫入採購系統 BOM 本 */
var TZ = 'Asia/Taipei';
var SEG_MAX = 45000, SEG_N = 4;       /* payload 每格上限、格數 */
var LOG_KEEP = 5000;                   /* log 分頁保留筆數 */

/* 分頁與欄序固定（改欄序前先確認前端與 B／C 階段程式） */
var TABS = {
  fact_campaign: ['campaign_id', '檔期名稱', '起', '迄', '第一批配貨日', '品牌別', '自己人搶先體驗', '配合活動',
    '作業_第一批出貨_時間', '作業_第一批出貨_備註', '作業_貼紙_時間', '作業_貼紙_備註', '作業_POS_時間', '作業_POS_備註',
    '檔期備註', 'status', 'created_at', 'updated_at', 'updated_by'],
  fact_recipe_req: ['req_id', 'campaign_id', 'seq', '商品暫定名稱', '商品正式名稱', '定價', '成本', '利潤率', '規格', '葷素',
    '保存方式', '包裝方式', '製作時間', '預估銷售數', 'status', 'signers', 'payload_1', 'payload_2', 'payload_3', 'payload_4',
    'created_at', 'updated_at', 'updated_by'],
  fact_req_unit: ['req_id', '品項', '容器', '購買連結', '供應商覆寫', '出貨中心預估出貨量', '品項備註', 'updated_at', 'updated_by'],
  dim_role: ['role', 'password', 'fields'],
  dim_signer: ['name', 'role', 'pin', 'enabled'],
  fact_signoff: ['req_id', 'signer', 'decision', 'comment', 'ts'],
  log: ['ts', 'role', 'action', 'id', 'ok', 'msg'],
  fact_req_fill: ['req_id', 'role', 'updated_at'],   /* B 階段新增：每張單每個身分最後一次存檔時間（「已填」徽章）；舊試算表由 migrate_ 自動補建 */
  fact_push: ['ts', 'req_id', 'batch', 'action', 'table', 'rows', 'detail', 'backup', 'ok', 'msg']   /* C 階段新增：寫入／還原採購系統的逐表紀錄 */
};
/* 數字欄（其餘一律純文字，避免「01」「2026-10-01」被試算表自動轉成數字或日期） */
var NUM_COLS = { '定價': 1, '成本': 1, '利潤率': 1, '預估銷售數': 1, '出貨中心預估出貨量': 1 };

/* 各身分初始可寫欄位（寫進 dim_role.fields，之後直接在試算表改即可，不用改程式）
   語法：C:＝檔期欄、R:＝申請單欄、U:＝品項欄、P:＝申請單內容（P:* 才能整張送件）；結尾 * ＝萬用；開頭 ! ＝排除 */
var ROLE_SEED = [
  ['主廚', 'C:*, !C:作業_*, R:*, P:*'],
  ['出貨中心', 'C:作業_第一批出貨_*, U:容器, U:出貨中心預估出貨量, U:品項備註'],
  ['採購', 'U:購買連結, U:供應商覆寫, U:品項備註, P:lines.vendor, P:lines.pkg_cost, P:lines.pkg_spec, P:lines.pkg_type, P:lines.split, P:lines.first'],
  ['行銷設計', 'C:作業_貼紙_*, P:lines.sticker'],
  ['營運POS', 'C:作業_POS_*, P:head.cat, P:head.catOther']
];
/* 系統欄：一律由伺服器寫，前端送來的值忽略 */
var CAMP_SYS = { campaign_id: 1, created_at: 1, updated_at: 1, updated_by: 1 };
var REQ_SYS = { req_id: 1, campaign_id: 1, seq: 1, status: 1, signers: 1, payload_1: 1, payload_2: 1, payload_3: 1, payload_4: 1, created_at: 1, updated_at: 1, updated_by: 1 };
var UNIT_SYS = { req_id: 1, '品項': 1, updated_at: 1, updated_by: 1 };
/* 這些狀態的申請單不能再整張修改 */
var LOCKED = { '送簽中': 1, '已核准': 1, '已核准（採購寫入失敗）': 1, '已寫入採購': 1 };
var ADMIN = '管理者';   /* 🔑 密碼管理：簽核主管名單與 PIN、各身分密碼 */

/* ================= 第一次設定 ================= */
function setup() {
  var P = PropertiesService.getScriptProperties();
  var id = P.getProperty('SHEET_ID'), ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { throw new Error('指令碼屬性 SHEET_ID 開不到：' + id + '（' + e.message + '）'); }
  } else {
    ss = SpreadsheetApp.create('食譜系統_申請單');
    P.setProperty('SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(TZ);
  var created = [];
  Object.keys(TABS).forEach(function (name) { if (ensureTab_(ss, name)) created.push(name); });
  /* 新建試算表附帶的空白「工作表1」刪掉 */
  var extra = ss.getSheets().filter(function (sh) { return !TABS[sh.getName()] && sh.getLastRow() === 0; });
  extra.forEach(function (sh) { if (ss.getSheets().length > 1) ss.deleteSheet(sh); });
  var seeded = seedRoles_(ss);
  var msg = '食譜系統_申請單 ' + ss.getUrl() + '｜新建分頁：' + (created.join('、') || '無') +
    '｜身分密碼：' + (seeded ? '已產生 ' + seeded + ' 組（在 dim_role 分頁）' : '已存在，未變更');
  Logger.log(msg);
  return msg;
}

function ensureTab_(ss, name) {
  var H = TABS[name], sh = ss.getSheetByName(name), isNew = false;
  if (!sh) { sh = ss.insertSheet(name, ss.getSheets().length); isNew = true; }
  var w = Math.max(sh.getLastColumn(), H.length);
  var cur = sh.getRange(1, 1, 1, w).getValues()[0].map(function (x) { return String(x).trim(); });
  if (cur.every(function (x) { return x === ''; })) {
    sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#E0F2F1');
    sh.setFrozenRows(1);
  } else {
    for (var i = 0; i < H.length; i++) {
      if (cur[i] !== H[i]) throw new Error(name + ' 表頭第 ' + (i + 1) + ' 欄是「' + cur[i] + '」，應為「' + H[i] + '」。setup 不改現有資料，請人工確認。');
    }
  }
  if (isNew && sh.getMaxColumns() > H.length) sh.deleteColumns(H.length + 1, sh.getMaxColumns() - H.length);
  var n = sh.getMaxRows() - 1;
  if (n > 0) H.forEach(function (h, j) { if (!NUM_COLS[h]) sh.getRange(2, j + 1, n, 1).setNumberFormat('@'); });
  return isNew;
}

function seedRoles_(ss) {
  var sh = ss.getSheetByName('dim_role');
  if (sh.getLastRow() > 1) return 0;
  var rows = ROLE_SEED.map(function (r) { return [r[0], randPw_(), r[1]]; });
  sh.getRange(2, 1, rows.length, 3).setNumberFormat('@').setValues(rows);
  return rows.length;
}

function randPw_(len) {
  var n = len || 8, cs = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + ':' + Date.now() + ':' + Math.random());
  var lim = 256 - (256 % cs.length);
  for (var i = 0; i < b.length && s.length < n; i++) { var v = (b[i] + 256) % 256; if (v < lim) s += cs.charAt(v % cs.length); }
  while (s.length < n) s += cs.charAt(Math.floor(Math.random() * cs.length));
  return s;
}

/* 自動遷移（第一次有人呼叫 API 時跑一次，之後看指令碼屬性就跳過）
   ① 2026-09-28「管理者」身分（🔑 密碼管理用）：舊試算表沒有這一列 → 補上；密碼 10 碼隨機，只在 dim_role 分頁
   ② 2026-09-29 B 階段：補建 fact_req_fill 分頁（只新增分頁，不動既有分頁）
   ③ 2026-09-29 C 階段：補建 fact_push 分頁 */
function migrate_() {
  var P = PropertiesService.getScriptProperties();
  if (!P.getProperty('SHEET_ID')) return;
  if (P.getProperty('MIG_ADMIN') === '1' && P.getProperty('MIG_FILL') === '1' && P.getProperty('MIG_PUSH') === '1') return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    if (P.getProperty('MIG_ADMIN') !== '1') {
      var t = load_('dim_role');
      if (!t.rows.some(function (r) { return str_(r.role).trim() === ADMIN; })) put_(t, nextRow_(t), { role: ADMIN, password: randPw_(10), fields: '*' });
      P.setProperty('MIG_ADMIN', '1');
    }
    if (P.getProperty('MIG_FILL') !== '1') {
      ensureTab_(ss_(), 'fact_req_fill');
      P.setProperty('MIG_FILL', '1');
    }
    if (P.getProperty('MIG_PUSH') !== '1') {
      ensureTab_(ss_(), 'fact_push');
      P.setProperty('MIG_PUSH', '1');
    }
  } catch (e) { /* 下次再試 */ } finally { lock.releaseLock(); }
}

/* ================= 入口 ================= */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || 'ping'), res;
  migrate_();
  try {
    if (a === 'ping') res = { ok: true, data: { version: VERSION, now: now_(), ready: !!PropertiesService.getScriptProperties().getProperty('SHEET_ID') } };
    else if (a === 'listCampaigns') res = { ok: true, data: listCampaigns_() };
    else if (a === 'signInfo') res = { ok: true, data: signInfo_(p.req_id) };
    else if (a === 'newItemsPub') res = { ok: true, data: newItemsPub_() };
    else res = { ok: false, msg: '不支援的查詢：' + a + '（讀申請單內容要用 POST 並帶密碼）' };
  } catch (err) { res = { ok: false, msg: errMsg_(err) }; }
  res.act = a;   /* 回覆註明是哪個動作的結果（Google 回傳鏈偶爾會把請求導回預設 ping，前端靠這個判斷要不要重試） */
  return out_(res, p.callback);
}

var ACTIONS = { whoami: whoami_, saveCampaign: saveCampaign_, saveReq: saveReq_, saveUnit: saveUnit_, getCampaign: getCampaign_, getReq: getReq_,
  adminList: adminList_, adminSaveSigner: adminSaveSigner_, adminDeleteSigner: adminDeleteSigner_, adminSetPassword: adminSetPassword_,
  listSigners: listSigners_, patchReq: patchReq_, submitSign: submitSign_, withdrawSign: withdrawSign_, signView: signView_, sign: sign_,
  pushPreview: pushPreview_, pushPurchase: pushPurchase_, rollbackPurchase: rollbackPurchase_ };
var WRITES = { saveCampaign: 1, saveReq: 1, saveUnit: 1, adminSaveSigner: 1, adminDeleteSigner: 1, adminSetPassword: 1,
  patchReq: 1, submitSign: 1, withdrawSign: 1, sign: 1, pushPurchase: 1, rollbackPurchase: 1 };   /* 回覆裡都不含密碼／PIN，才能放進回條快取 */
var RQ_SEC = 21600;   /* 回條保留 6 小時：同一個回條編號重送 → 直接回上次結果，不重複寫 */

function doPost(e) {
  var p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, msg: '送來的資料不是 JSON' }); }
  var act = String(p.action || ''), who = p.signer ? ('簽核:' + String(p.signer)) : String(p.role || ''), res;
  migrate_();
  var fn = ACTIONS[act];
  if (!fn) { res = { ok: false, act: act, msg: '不支援的動作：' + act }; log_(who, act, '', false, res.msg); return out_(res); }
  /* 先驗身分（在鎖外面）：密碼錯 → 停 1 秒再回（拖慢亂猜，不佔鎖、不影響別人）；不做「錯太多次就鎖整個身分」——
     那會讓知道網址的人故意打錯把全部主廚鎖在外面（2026-09-28 測試時實際發生過） */
  var auth;
  try { auth = auth_(p); } catch (err) {
    Utilities.sleep(1000);
    res = { ok: false, act: act, code: (err && err.code) || 'auth', msg: errMsg_(err) };
    log_(who, act, guessId_(p), false, res.msg);
    return out_(res);
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { res = { ok: false, act: act, code: 'busy', msg: '系統忙碌中，請 10 秒後再按一次' }; log_(who, act, '', false, res.msg); return out_(res); }
  /* 回條 key＝動作＋回條編號（2026-09-29：只用回條編號時，不同動作碰巧用到同一個編號會拿到別的動作的舊結果） */
  var rqKey = (WRITES[act] && p.rq) ? 'rq:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(act + '|' + String(p.rq).slice(0, 120)).getBytes()) : '';
  try {
    var cached = rqKey ? CacheService.getScriptCache().get(rqKey) : null;
    if (cached) {   /* 同一次送出的重送（上次已寫入，只是回覆在路上掉了）→ 回上次結果 */
      res = JSON.parse(cached);
      log_(who, act, guessId_(p), true, '重送：沿用上次結果，未重複寫入');
    } else {
      var r = fn(p, auth) || {};
      res = { ok: true, act: act, data: r.data };
      if (rqKey) { try { CacheService.getScriptCache().put(rqKey, JSON.stringify(res), RQ_SEC); } catch (ce) { /* 回覆太大放不進快取 → 只是重送時不能沿用 */ } }
      log_(who, act, r.id || '', true, r.msg || '');
    }
  } catch (err) {
    res = { ok: false, act: act, msg: errMsg_(err) };
    if (err && err.code) res.code = err.code;
    if (err && err.extra) res.data = err.extra;
    log_(who, act, guessId_(p), false, res.msg);
  } finally { lock.releaseLock(); }
  return out_(res);
}

function out_(res, cb) {
  var s = JSON.stringify(res);
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$]{0,63}$/.test(String(cb))) {
    return ContentService.createTextOutput(cb + '(' + s + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function guessId_(p) {
  if (!p) return '';
  return String(p.req_id || (p.req && p.req.req_id) || (p.campaign && p.campaign.campaign_id) || p.campaign_id ||
    (p.entry && p.entry.name) || p.target_role || (p.action === 'adminDeleteSigner' ? p.name : '') || '');
}

/* ================= 身分 ================= */
function auth_(p) {
  if (p.signer) {   /* 簽核人用 PIN：同一位錯 PIN_MAX 次 → 暫停 15 分鐘（只影響那一個人，其他簽核人、各身分都照常） */
    var name = String(p.signer).trim(), pin = String(p.pin || '').trim();
    var cache = CacheService.getScriptCache(), ck = 'pinfail:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(name).getBytes()).slice(0, 200);
    var fails = parseInt(cache.get(ck) || '0', 10) || 0;
    if (fails >= PIN_MAX) throw fail_('「' + name + '」PIN 錯太多次，請 15 分鐘後再試（其他人不受影響）', 'pinlock');
    var s = load_('dim_signer').rows.filter(function (r) { return str_(r.name).trim() === name && signerOn_(r.enabled); })[0];
    if (!s || !pin || str_(s.pin).trim() !== pin) { cache.put(ck, String(fails + 1), PIN_LOCK_SEC); throw fail_('簽核人或 PIN 不對', 'auth'); }
    if (fails) cache.remove(ck);
    return { kind: 'signer', name: name, role: str_(s.role).trim(), rules: [] };
  }
  var role = String(p.role || '').trim(), pw = String(p.password || '').trim();
  if (!role) throw fail_('請選擇身分並輸入密碼', 'auth');
  var r = load_('dim_role').rows.filter(function (x) { return str_(x.role).trim() === role; })[0];
  if (!r || !pw || str_(r.password).trim() !== pw) throw fail_('密碼不對', 'auth');
  return { kind: 'role', role: role, name: '', rules: parseRules_(r.fields), fields: str_(r.fields) };
}
function signerOn_(v) { var s = str_(v).trim().toUpperCase(); return ['N', 'NO', 'FALSE', '否', '0', '停用'].indexOf(s) < 0; }

function parseRules_(s) {
  return String(s || '').split(/[,，、;；\s]+/).map(function (x) { return x.trim(); }).filter(Boolean).map(function (x) {
    var neg = x.charAt(0) === '!'; if (neg) x = x.slice(1);
    return { neg: neg, pat: x };
  });
}
function match_(pat, key) { if (pat === '*') return true; if (pat.slice(-1) === '*') return key.indexOf(pat.slice(0, -1)) === 0; return pat === key; }
/* key 形如 'C:檔期名稱'、'R:定價'、'U:容器'、'P:*' */
function can_(auth, key) {
  if (!auth || auth.kind !== 'role') return false;
  var yes = auth.rules.some(function (r) { return !r.neg && match_(r.pat, key); });
  if (!yes) return false;
  return !auth.rules.some(function (r) { return r.neg && match_(r.pat, key); });
}

function whoami_(p, auth) {
  return { id: '', data: { kind: auth.kind, role: auth.role, name: auth.name, fields: auth.fields || '' } };
}

/* ================= 🔑 密碼管理（只限「管理者」） ================= */
function needAdmin_(auth) { if (!auth || auth.kind !== 'role' || auth.role !== ADMIN) throw fail_('這一頁要「' + ADMIN + '」密碼', 'perm'); }
function adminList_(p, auth) {
  needAdmin_(auth);
  var roles = load_('dim_role').rows.filter(function (x) { return str_(x.role).trim(); })
    .map(function (x) { return { role: str_(x.role).trim(), password: str_(x.password), fields: str_(x.fields) }; });
  var signers = load_('dim_signer').rows.filter(function (x) { return str_(x.name).trim(); })
    .map(function (x) { return { name: str_(x.name).trim(), role: str_(x.role), pin: str_(x.pin), enabled: signerOn_(x.enabled) ? 'Y' : 'N' }; });
  return { id: '', data: { roles: roles, signers: signers }, msg: '身分 ' + roles.length + '、簽核主管 ' + signers.length };
}
function adminSaveSigner_(p, auth) {   /* 資料放 p.entry（p.signer 是簽核人 PIN 登入用，不能混用） */
  needAdmin_(auth);
  var g = p.entry || {}, orig = str_(g.orig_name).trim(), name = str_(g.name).trim(), title = str_(g.role).trim(), pin = str_(g.pin).trim();
  var en = (g.enabled === false || str_(g.enabled).toUpperCase() === 'N') ? 'N' : 'Y';
  if (!name) throw fail_('請填簽核主管姓名');
  if (/[,，、;；\s]/.test(name)) throw fail_('姓名不能有逗號、頓號、分號或空白');   /* signers 欄用這些符號分隔 */
  if (!/^\d{4,8}$/.test(pin)) throw fail_('PIN 要 4～8 位數字');
  var t = load_('dim_signer');
  var row = orig ? t.rows.filter(function (r) { return str_(r.name).trim() === orig; })[0] : null;
  if (orig && !row) throw fail_('找不到簽核主管「' + orig + '」', 'notfound');
  if (t.rows.some(function (r) { return r !== row && str_(r.name).trim() === name; })) throw fail_('已經有叫「' + name + '」的簽核主管', 'dup');
  put_(t, row ? row._row : nextRow_(t), { name: name, role: title, pin: pin, enabled: en });
  return {
    id: name, data: { name: name, role: title, enabled: en, created: !row },
    msg: (row ? '更新' : '新增') + '簽核主管 ' + name + (orig && orig !== name ? '（原名 ' + orig + '）' : '') + (en === 'N' ? '（停用）' : '')
  };
}
function adminDeleteSigner_(p, auth) {
  needAdmin_(auth);
  var name = str_(p.name).trim();
  if (!name) throw fail_('缺姓名');
  var t = load_('dim_signer'), row = t.rows.filter(function (r) { return str_(r.name).trim() === name; })[0];
  if (!row) throw fail_('找不到簽核主管「' + name + '」', 'notfound');
  var signed = load_('fact_signoff').rows.some(function (r) { return str_(r.signer).trim() === name; });
  var picked = load_('fact_recipe_req', true).rows.some(function (r) { return str_(r.signers).split(/[,，、;；\s]+/).indexOf(name) >= 0; });
  if (signed || picked) throw fail_('「' + name + '」已經' + (signed ? '簽核過' : '被指定簽核') + '，為了保留紀錄不能刪除，請改成「停用」', 'used');
  t.sh.deleteRow(row._row);
  return { id: name, data: { name: name }, msg: '刪除簽核主管 ' + name };
}
/* 新密碼由網頁產生後送來（重送時是同一組，不會變成兩個不同的密碼）；回覆不含密碼。
   注意：role／password 是管理者自己的登入，要改的身分與新密碼用 target_role／new_password */
function adminSetPassword_(p, auth) {
  needAdmin_(auth);
  var role = str_(p.target_role).trim(), pw = str_(p.new_password).trim();
  var t = load_('dim_role'), row = t.rows.filter(function (r) { return str_(r.role).trim() === role; })[0];
  if (!row) throw fail_('找不到身分「' + role + '」', 'notfound');
  if (pw.length < 6) throw fail_('密碼至少 6 碼');
  if (/\s/.test(pw)) throw fail_('密碼不能有空白');
  var obj = {};
  t.H.forEach(function (h) { obj[h] = row[h]; });
  obj.password = pw;
  put_(t, row._row, obj);
  return { id: role, data: { role: role }, msg: '更改密碼 ' + role };
}

/* ================= 檔期 ================= */
function listCampaigns_() {
  var c = load_('fact_campaign'), q = load_('fact_recipe_req', true), cnt = {};
  q.rows.forEach(function (r) { var k = str_(r.campaign_id); if (k) cnt[k] = (cnt[k] || 0) + 1; });
  return c.rows.filter(function (r) { return str_(r.campaign_id); }).map(function (r) {
    var o = plain_(c.H, r); o.req_count = cnt[o.campaign_id] || 0; return o;
  });
}

function saveCampaign_(p, auth) {
  var c = p.campaign || {}, t = load_('fact_campaign'), now = now_();
  var id = str_(c.campaign_id).trim();
  var row = id ? t.rows.filter(function (r) { return str_(r.campaign_id) === id; })[0] : null;
  if (id && !row) throw fail_('找不到檔期 ' + id, 'notfound');
  var obj = {}, skipped = [];
  if (row) t.H.forEach(function (h) { obj[h] = row[h]; });
  else {
    if (!can_(auth, 'C:檔期名稱')) throw fail_('「' + auth.role + '」不能新增檔期', 'perm');
    var name = str_(c['檔期名稱']).trim();
    if (!name) throw fail_('請填檔期名稱');
    var dup = t.rows.filter(function (r) { return str_(r['檔期名稱']).trim() === name; })[0];
    if (dup && !p.force) throw fail_('已經有同名的檔期「' + name + '」（' + str_(dup.campaign_id) + '）', 'dup', { campaign_id: str_(dup.campaign_id) });
    id = 'C-' + stamp_();
    var base = id, k = 2;
    while (t.rows.some(function (r) { return str_(r.campaign_id) === id; })) id = base + '-' + (k++);
    obj.campaign_id = id; obj.created_at = now; obj.status = '啟用';
  }
  Object.keys(c).forEach(function (h) {
    if (CAMP_SYS[h] || t.H.indexOf(h) < 0) return;
    if (!can_(auth, 'C:' + h)) { if (str_(c[h]) !== str_(obj[h])) skipped.push(h); return; }
    obj[h] = c[h];
  });
  ['起', '迄', '第一批配貨日'].forEach(function (h) { obj[h] = normDate_(obj[h]); });
  if (obj['起'] && obj['迄'] && /^\d{4}-\d{2}-\d{2}$/.test(obj['起']) && /^\d{4}-\d{2}-\d{2}$/.test(obj['迄']) && obj['迄'] < obj['起']) throw fail_('檔期「迄」早於「起」，請確認日期');
  obj.updated_at = now; obj.updated_by = auth.role;
  put_(t, row ? row._row : nextRow_(t), obj);
  return {
    id: id,
    data: { campaign: plain_(t.H, obj), created: !row, skipped: skipped },
    msg: (row ? '更新檔期' : '新增檔期') + ' ' + str_(obj['檔期名稱']) + (skipped.length ? '｜略過無權限欄位：' + skipped.join('、') : '')
  };
}

/* ================= 申請單 ================= */
function saveReq_(p, auth) {
  var q = p.req || {}, now = now_();
  var campId = str_(q.campaign_id).trim();
  if (!campId) throw fail_('請先選檔期');
  var ct = load_('fact_campaign'), camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === campId; })[0];
  if (!camp) throw fail_('找不到檔期 ' + campId + '（可能被刪除，請重新選）', 'notfound');
  var t = load_('fact_recipe_req', true);
  var rid = str_(q.req_id).trim();
  var row = rid ? t.rows.filter(function (r) { return str_(r.req_id) === rid; })[0] : null;
  var hasPayload = typeof p.payload === 'string' && p.payload.length > 0;
  if ((hasPayload || !row) && !can_(auth, 'P:*')) throw fail_('「' + auth.role + '」不能' + (row ? '整張修改申請單內容' : '新增申請單'), 'perm');
  var j = null;
  if (hasPayload) {
    try { j = JSON.parse(p.payload); } catch (e) { throw fail_('申請單內容不是正確的 JSON'); }
    if (!j || typeof j !== 'object') throw fail_('申請單內容格式不對');
  }
  /* 防重複：畫面上還沒有單號、但同一張表單（內容裡的 created_at 相同）24 小時內已經送成功過
     （第一次其實寫進去了、只是回覆沒回到畫面）→ 當成更新那一張，不再新增第二張 */
  var sameForm = false;
  if (!row && j && typeof j.created_at === 'string' && j.created_at) {
    var hit = findByCreated_(t, j.created_at);
    if (hit) { row = hit; sameForm = true; }
  }
  if (row) {
    var st = str_(row.status);
    if (LOCKED[st]) throw fail_('這張單目前是「' + st + '」，不能修改', 'locked', { status: st });
    var base = str_(p.base_updated_at);
    if (!p.force && !sameForm && str_(row.updated_at) !== base) {
      /* 內容跟系統上一模一樣（同一次送件重按）→ 視為已送成功，不算衝突 */
      if (hasPayload && readPayload_(t, row._row) === p.payload && str_(row.campaign_id) === campId) {
        return {
          id: str_(row.req_id),
          data: { req_id: str_(row.req_id), campaign_id: campId, campaign_name: str_(camp['檔期名稱']), seq: str_(row.seq), status: str_(row.status),
            created: false, moved: false, same: true, created_at: str_(row.created_at), updated_at: str_(row.updated_at), updated_by: str_(row.updated_by),
            chars: p.payload.length, segments: null, skipped: [] },
          msg: '內容與系統上相同（重按），未重複寫入'
        };
      }
      throw fail_('這張單在 ' + str_(row.updated_at) + ' 被「' + str_(row.updated_by) + '」改過；你畫面上的版本' + (base ? '是 ' + base + ' 載入的' : '不是從檔期載入的'), 'conflict',
        { updated_at: str_(row.updated_at), updated_by: str_(row.updated_by) });
    }
  }
  var obj = {}, skipped = [];
  if (row) t.H.forEach(function (h) { obj[h] = row[h]; });
  else {
    var nid = rid || ('R-' + stamp_()), b0 = nid, k = 2;
    while (t.rows.some(function (r) { return str_(r.req_id) === nid; })) nid = b0 + '-' + (k++);
    obj.req_id = nid; obj.created_at = now; obj.status = '草稿'; obj.signers = '';
  }
  /* 新單或換檔期 → 序號＝該檔期現有最大號＋1（兩位數，超過 99 自動變三位） */
  var moved = row && str_(row.campaign_id) !== campId;
  if (!row || moved) {
    var mx = 0;
    t.rows.forEach(function (r) { if (r !== row && str_(r.campaign_id) === campId) { var n = parseInt(str_(r.seq), 10); if (n > mx) mx = n; } });
    obj.seq = pad2_(mx + 1);
  }
  obj.campaign_id = campId;
  Object.keys(q).forEach(function (h) {
    if (REQ_SYS[h] || t.H.indexOf(h) < 0) return;
    if (!can_(auth, 'R:' + h)) { if (str_(q[h]) !== str_(obj[h])) skipped.push(h); return; }
    obj[h] = q[h];
  });
  var pl = '', segs = [];
  if (hasPayload) {
    pl = p.payload;
    if (j.req_id !== obj.req_id || !j.head || j.head.req_id !== obj.req_id) {   /* 新單由伺服器配號 → 寫回內容裡的單號 */
      j.req_id = obj.req_id; if (j.head && typeof j.head === 'object') j.head.req_id = obj.req_id;
      pl = JSON.stringify(j);
    }
    segs = split_(pl);
    for (var i = 0; i < SEG_N; i++) obj['payload_' + (i + 1)] = segs[i] || '';
  }
  obj.updated_at = now; obj.updated_by = auth.role;
  var rowNo = row ? row._row : nextRow_(t);
  putCols_(t, rowNo, obj, function (h) { return hasPayload || h.indexOf('payload_') !== 0; });
  markFill_(obj.req_id, auth.role);
  return {
    id: obj.req_id,
    data: {
      req_id: obj.req_id, campaign_id: campId, campaign_name: str_(camp['檔期名稱']), seq: str_(obj.seq), status: str_(obj.status),
      created: !row, moved: !!moved, same_form: sameForm, created_at: str_(obj.created_at), updated_at: now, updated_by: auth.role,
      chars: hasPayload ? pl.length : null, segments: hasPayload ? segs.length : null, skipped: skipped
    },
    msg: (sameForm ? '同一張表單重送（之前已送成功）→更新' : (row ? (moved ? '改檔期並更新' : '更新') : '新增')) + ' 第' + str_(obj.seq) + '支 ' + str_(obj['商品暫定名稱'] || obj['商品正式名稱']) +
      (hasPayload ? '｜內容 ' + pl.length + ' 字 ' + segs.length + ' 格' : '') + (skipped.length ? '｜略過無權限欄位：' + skipped.join('、') : '')
  };
}

function saveUnit_(p, auth) {
  var rid = str_(p.req_id).trim(), list = p.units || [];
  if (!rid) throw fail_('缺申請單號');
  var qt = load_('fact_recipe_req', true), qrow = reqRow_(qt, rid);
  if (!qrow) throw fail_('找不到申請單 ' + rid, 'notfound');
  if (LOCKED[str_(qrow.status)]) throw fail_('這張單目前是「' + str_(qrow.status) + '」，不能修改', 'locked', { status: str_(qrow.status) });
  var t = load_('fact_req_unit'), now = now_(), n = 0, skipped = {};
  list.forEach(function (u) {
    var item = str_(u && u['品項']).trim(); if (!item) return;
    var row = t.rows.filter(function (r) { return str_(r.req_id) === rid && str_(r['品項']).trim() === item; })[0];
    var obj = {}, changed = false;
    if (row) t.H.forEach(function (h) { obj[h] = row[h]; }); else { obj.req_id = rid; obj['品項'] = item; }
    Object.keys(u).forEach(function (h) {
      if (UNIT_SYS[h] || t.H.indexOf(h) < 0) return;
      if (!can_(auth, 'U:' + h)) { if (str_(u[h]) !== str_(obj[h])) skipped[h] = 1; return; }
      if (str_(u[h]) !== str_(obj[h])) { obj[h] = u[h]; changed = true; }
    });
    if (!changed) return;
    obj.updated_at = now; obj.updated_by = auth.role;
    var rowNo = row ? row._row : nextRow_(t);
    put_(t, rowNo, obj);
    if (!row) { obj._row = rowNo; t.rows.push(obj); }
    n++;
  });
  var sk = Object.keys(skipped);
  if (n) markFill_(rid, auth.role);
  return { id: rid, data: { saved: n, skipped: sk, updated_at: now }, msg: '品項 ' + n + ' 列' + (sk.length ? '｜略過無權限欄位：' + sk.join('、') : '') };
}

function getCampaign_(p, auth) {
  var cid = str_(p.campaign_id).trim();
  if (!cid) throw fail_('缺檔期');
  var ct = load_('fact_campaign'), camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === cid; })[0];
  if (!camp) throw fail_('找不到檔期 ' + cid, 'notfound');
  var lite = !!p.lite || auth.kind === 'signer';
  var t = load_('fact_recipe_req', true);
  var rows = t.rows.filter(function (r) { return str_(r.campaign_id) === cid; })
    .sort(function (a, b) { return (parseInt(str_(a.seq), 10) || 0) - (parseInt(str_(b.seq), 10) || 0); });
  var ids = {};
  var reqs = rows.map(function (r) {
    ids[str_(r.req_id)] = 1;
    var o = plain_(t.H, r);
    if (!lite) o.payload = readPayload_(t, r._row);
    return o;
  });
  return {
    id: cid,
    data: { campaign: plain_(ct.H, camp), reqs: reqs, units: rowsOf_('fact_req_unit', ids), signoffs: rowsOf_('fact_signoff', ids),
      fills: safeRows_('fact_req_fill', ids), me: meOf_(auth) },
    msg: reqs.length + ' 支' + (lite ? '（清單）' : '')
  };
}

function getReq_(p, auth) {
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var t = load_('fact_recipe_req', true), row = t.rows.filter(function (r) { return str_(r.req_id) === rid; })[0];
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  if (auth.kind === 'signer') {
    var names = str_(row.signers).split(/[,，、;；\s]+/).map(function (x) { return x.trim(); });
    if (names.indexOf(auth.name) < 0) throw fail_('你不在這張單的簽核人名單', 'perm');
  }
  var req = plain_(t.H, row);
  req.payload = readPayload_(t, row._row);
  var ct = load_('fact_campaign'), camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === str_(row.campaign_id); })[0];
  var ids = {}; ids[rid] = 1;
  return {
    id: rid,
    data: { req: req, campaign: camp ? plain_(ct.H, camp) : null, units: rowsOf_('fact_req_unit', ids), signoffs: rowsOf_('fact_signoff', ids),
      fills: safeRows_('fact_req_fill', ids), me: meOf_(auth), push: pushInfo_(rid) },
    msg: '內容 ' + req.payload.length + ' 字'
  };
}

/* ================= B 階段：各單位局部填寫＋送簽＋簽核（2026-09-29） ================= */
var FILL_ROLES = ['主廚', '出貨中心', '採購', '行銷設計', '營運POS'];
var PIN_MAX = 10, PIN_LOCK_SEC = 900;
var VOID = '作廢：';   /* 重新送簽／撤回時，上一輪的簽核在 decision 前面加這個字（紀錄保留） */
/* 衍生欄：能改「原包裝規格」就能一起更新由它算出來的「內容量」 */
var PATCH_DERIVED = { 'lines.pack': 'lines.pkg_spec', 'lines.vname': 'lines.vendor' };   /* E 階段：能改進貨廠商就能改廠商品名 */

function meOf_(auth) { return auth.kind === 'role' ? { kind: 'role', role: auth.role, fields: auth.fields || '' } : { kind: 'signer', name: auth.name }; }
function signersOf_(row) { return str_(row.signers).split(/[,，、;；\s]+/).map(function (x) { return x.trim(); }).filter(Boolean); }
function isLive_(d) { d = str_(d).trim(); return d === '同意' || d === '退回'; }
function reqRow_(t, rid) { return t.rows.filter(function (r) { return str_(r.req_id) === rid; })[0] || null; }
/* 讀分頁中屬於這些單號的列；分頁還不存在（還沒遷移）→ 回空陣列，不讓整個查詢失敗 */
function safeRows_(name, ids) { try { return rowsOf_(name, ids); } catch (e) { return []; } }
function sameVal_(a, b) { return JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b); }

/* 「已填」紀錄：每張單每個身分一列（最後一次存檔時間） */
function markFill_(rid, role) {
  if (!rid || FILL_ROLES.indexOf(role) < 0) return;
  try {
    var t = load_('fact_req_fill');
    var row = t.rows.filter(function (r) { return str_(r.req_id) === rid && str_(r.role) === role; })[0];
    put_(t, row ? row._row : nextRow_(t), { req_id: rid, role: role, updated_at: now_() });
  } catch (e) { /* 分頁還不存在就不記，不影響存檔 */ }
}

/* 送簽時勾選用：啟用中的簽核主管（只回姓名、職稱，不含 PIN） */
function listSigners_(p, auth) {
  if (auth.kind !== 'role') throw fail_('要用身分密碼', 'perm');
  var list = load_('dim_signer').rows.filter(function (r) { return str_(r.name).trim() && signerOn_(r.enabled); })
    .map(function (r) { return { name: str_(r.name).trim(), role: str_(r.role).trim() }; });
  return { id: '', data: list, msg: list.length + ' 位' };
}

/* 各單位只改自己負責的申請單內容欄位（P:head.xxx／P:lines.xxx），伺服器逐欄檢查權限；
   直接改在系統上最新的內容上（不需要整張重送，也不會蓋掉主廚或別的單位的欄位） */
function patchReq_(p, auth) {
  if (auth.kind !== 'role') throw fail_('要用身分密碼', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var st = str_(row.status);
  if (LOCKED[st]) throw fail_('這張單目前是「' + st + '」，不能修改', 'locked', { status: st });
  var j;
  try { j = JSON.parse(readPayload_(t, row._row)); } catch (e) { throw fail_('這張單的內容讀不出來，請主廚重新送件一次'); }
  if (!j || typeof j !== 'object') throw fail_('這張單的內容格式不對');
  var pt = p.patch || {}, skipped = [], changed = [];
  function ok(key) { var d = PATCH_DERIVED[key]; return can_(auth, 'P:' + key) || (!!d && can_(auth, 'P:' + d)); }
  var ph = pt.head;
  if (ph && typeof ph === 'object') {
    j.head = j.head || {};
    Object.keys(ph).forEach(function (k) {
      if (!ok('head.' + k)) { skipped.push('表頭.' + k); return; }
      if (!sameVal_(j.head[k], ph[k])) { j.head[k] = ph[k]; changed.push('表頭.' + k); }
    });
    if (changed.indexOf('表頭.cat') >= 0 || changed.indexOf('表頭.catOther') >= 0)
      j.head.cat_text = str_(j.head.cat) === '__other' ? str_(j.head.catOther).trim() : str_(j.head.cat).trim();
  }
  var lines = j.lines || [];
  (pt.lines || []).forEach(function (pl) {
    if (!pl || typeof pl !== 'object') return;
    var nm = str_(pl.name).trim(), idx = findLine_(lines, pl);
    if (idx < 0) { skipped.push('找不到品項「' + nm + '」（可能被主廚改名或刪掉）'); return; }
    var L = lines[idx], f = pl.fields || {};
    Object.keys(f).forEach(function (k) {
      if (!ok('lines.' + k)) { skipped.push(nm + '.' + k); return; }
      if (!sameVal_(L[k], f[k])) { L[k] = f[k]; changed.push(nm + '.' + k); }
    });
  });
  if (!changed.length) return { id: rid, data: { changed: [], skipped: skipped, updated_at: str_(row.updated_at), updated_by: str_(row.updated_by) },
    msg: '沒有變更' + (skipped.length ? '｜略過：' + skipped.join('、') : '') };
  var pl2 = JSON.stringify(j), segs = split_(pl2), now = now_(), obj = { updated_at: now, updated_by: auth.role };
  for (var i = 0; i < SEG_N; i++) obj['payload_' + (i + 1)] = segs[i] || '';
  putCols_(t, row._row, obj, function (h) { return h.indexOf('payload_') === 0 || h === 'updated_at' || h === 'updated_by'; });
  markFill_(rid, auth.role);
  return { id: rid, data: { changed: changed, skipped: skipped, updated_at: now, updated_by: auth.role, payload: pl2 },
    msg: '改 ' + changed.length + ' 欄' + (skipped.length ? '｜略過：' + skipped.join('、') : '') };
}
function findLine_(lines, pl) {
  var nm = str_(pl.name).trim(), s = parseInt(pl.seq, 10);
  if (s > 0 && lines[s - 1] && str_(lines[s - 1].name).trim() === nm) return s - 1;
  for (var i = 0; i < lines.length; i++) if (lines[i] && str_(lines[i].name).trim() === nm) return i;
  return -1;
}

/* 上一輪（decision＝同意／退回）的簽核全部標記作廢，紀錄保留 */
function voidSignoffs_(rid) {
  var t;
  try { t = load_('fact_signoff'); } catch (e) { return 0; }
  var n = 0;
  t.rows.forEach(function (r) {
    if (str_(r.req_id) !== rid || !isLive_(r.decision)) return;
    putCols_(t, r._row, { decision: VOID + str_(r.decision).trim() }, function (h) { return h === 'decision'; });
    n++;
  });
  return n;
}

/* 送簽（主廚）：草稿／退回 → 送簽中；寫入簽核人；上一輪簽核標記作廢。要是系統上最新版（base_updated_at） */
function submitSign_(p, auth) {
  if (!can_(auth, 'P:*')) throw fail_('只有主廚可以送簽', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var st = str_(row.status);
  if (st !== '草稿' && st !== '退回') throw fail_('這張單目前是「' + st + '」，不能送簽', 'locked', { status: st });
  if (!p.force && str_(row.updated_at) !== str_(p.base_updated_at))
    throw fail_('這張單在 ' + str_(row.updated_at) + ' 被「' + str_(row.updated_by) + '」改過，請先載回最新版再送簽', 'conflict',
      { updated_at: str_(row.updated_at), updated_by: str_(row.updated_by) });
  var seen = {}, names = [];
  (p.signers || []).forEach(function (x) { var n = str_(x).trim(); if (n && !seen[n]) { seen[n] = 1; names.push(n); } });
  if (!names.length) throw fail_('請至少勾一位簽核主管');
  var on = {};
  load_('dim_signer').rows.forEach(function (r) { if (signerOn_(r.enabled)) on[str_(r.name).trim()] = 1; });
  var bad = names.filter(function (n) { return !on[n]; });
  if (bad.length) throw fail_('這些人不在啟用中的簽核主管名單：' + bad.join('、'), 'badsigner');
  var voided = voidSignoffs_(rid), now = now_();
  putCols_(t, row._row, { status: '送簽中', signers: names.join('、'), updated_at: now, updated_by: auth.role },
    function (h) { return h === 'status' || h === 'signers' || h === 'updated_at' || h === 'updated_by'; });
  return { id: rid, data: { status: '送簽中', signers: names, updated_at: now, updated_by: auth.role, voided: voided },
    msg: '送簽給 ' + names.join('、') + (voided ? '｜上一輪簽核 ' + voided + ' 筆標記作廢' : '') };
}

/* 撤回送簽（主廚）：送簽中 → 草稿；已經簽的標記作廢 */
function withdrawSign_(p, auth) {
  if (!can_(auth, 'P:*')) throw fail_('只有主廚可以撤回送簽', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var st = str_(row.status);
  if (st !== '送簽中') throw fail_('這張單目前是「' + st + '」，沒有在送簽', 'locked', { status: st });
  var voided = voidSignoffs_(rid), now = now_();
  putCols_(t, row._row, { status: '草稿', updated_at: now, updated_by: auth.role },
    function (h) { return h === 'status' || h === 'updated_at' || h === 'updated_by'; });
  return { id: rid, data: { status: '草稿', updated_at: now, updated_by: auth.role, voided: voided },
    msg: '撤回送簽' + (voided ? '｜已簽的 ' + voided + ' 筆標記作廢' : '') };
}

/* 簽核頁讀資料（簽核人 PIN）：這張單的完整內容＋同檔期其他商品一行清單＋簽核紀錄 */
function signView_(p, auth) {
  if (auth.kind !== 'signer') throw fail_('簽核頁要用簽核人 PIN', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  if (signersOf_(row).indexOf(auth.name) < 0) throw fail_('你不在這張單的簽核人名單（可能已改派其他主管）', 'perm');
  var req = plain_(t.H, row);
  req.payload = readPayload_(t, row._row);
  var cid = str_(row.campaign_id), ct = load_('fact_campaign'), camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === cid; })[0];
  var others = t.rows.filter(function (r) { return str_(r.campaign_id) === cid; })
    .sort(function (a, b) { return (parseInt(str_(a.seq), 10) || 0) - (parseInt(str_(b.seq), 10) || 0); })
    .map(function (r) {
      return { req_id: str_(r.req_id), seq: str_(r.seq), name: str_(r['商品正式名稱']) || str_(r['商品暫定名稱']),
        price: r['定價'] === '' ? '' : r['定價'], cost: r['成本'] === '' ? '' : r['成本'], margin: r['利潤率'] === '' ? '' : r['利潤率'], status: str_(r.status) };
    });
  var ids = {}; ids[rid] = 1;
  return { id: rid, data: { req: req, campaign: camp ? plain_(ct.H, camp) : null, others: others, units: rowsOf_('fact_req_unit', ids),
    signoffs: rowsOf_('fact_signoff', ids), me: meOf_(auth) }, msg: '簽核頁 ' + rid };
}

/* 簽核（簽核人 PIN）：同意／退回（退回要寫原因）。任一退回 → 退回；全部同意 → 已核准。送簽中可以改自己的決定 */
function sign_(p, auth) {
  if (auth.kind !== 'signer') throw fail_('簽核要用簽核人 PIN', 'perm');
  var rid = str_(p.req_id).trim(), dec = str_(p.decision).trim(), cm = str_(p.comment).trim().slice(0, 1000);
  if (!rid) throw fail_('缺申請單號');
  if (dec !== '同意' && dec !== '退回') throw fail_('請按「同意」或「退回」');
  if (dec === '退回' && !cm) throw fail_('退回請寫原因，主廚才知道要改哪裡');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var st = str_(row.status), names = signersOf_(row);
  if (names.indexOf(auth.name) < 0) throw fail_('你不在這張單的簽核人名單', 'perm');
  if (st !== '送簽中') throw fail_('這張單目前是「' + st + '」，不用再簽', 'locked', { status: st });
  var so = load_('fact_signoff');
  var mine = so.rows.filter(function (r) { return str_(r.req_id) === rid && str_(r.signer).trim() === auth.name && isLive_(r.decision); })[0];
  var rec = { req_id: rid, signer: auth.name, decision: dec, comment: cm, ts: now_() };
  var rowNo = mine ? mine._row : nextRow_(so);
  put_(so, rowNo, rec);
  if (mine) { mine.decision = dec; mine.comment = cm; mine.ts = rec.ts; } else { rec._row = rowNo; so.rows.push(rec); }
  var live = so.rows.filter(function (r) { return str_(r.req_id) === rid && isLive_(r.decision); });
  var agreed = names.filter(function (n) { return live.some(function (r) { return str_(r.signer).trim() === n && str_(r.decision).trim() === '同意'; }); });
  var status = dec === '退回' ? '退回' : (agreed.length === names.length ? '已核准' : st);
  if (status !== st) putCols_(t, row._row, { status: status }, function (h) { return h === 'status'; });
  /* C 階段：全部同意 → 同一次請求寫入採購系統；寫入失敗不影響核准（狀態記「已核准（採購寫入失敗）」，主廚可重試） */
  var push = null;
  if (status === '已核准') {
    try { push = pushToPurchase_(rid, '簽核:' + auth.name); status = PUSH_OK; }
    catch (e) { push = { ok: false, msg: errMsg_(e) }; status = e && e.code === 'dup' ? PUSH_OK : PUSH_FAIL; }
    putCols_(t, row._row, { status: status }, function (h) { return h === 'status'; });
  }
  return { id: rid, data: { status: status, decision: dec, signers: names, agreed: agreed, signoffs: live.map(function (r) { return plain_(so.H, r); }), push: push },
    msg: auth.name + ' ' + dec + '（' + agreed.length + '/' + names.length + ' 同意）' + (status !== st ? '→' + status : '') + (push ? '｜' + (push.ok === false ? push.msg : pushMsg_(push)) : '') };
}

/* 公開（簽核連結第一頁用）：只回狀態、簽核人名字、誰已經簽了——不含商品名稱與內容（看內容要 PIN） */
function signInfo_(rid) {
  rid = str_(rid).trim();
  if (!/^R-[0-9A-Za-z-]{6,40}$/.test(rid)) throw fail_('簽核連結的單號不對，請跟主廚要新的連結');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到這張申請單（連結可能打錯）', 'notfound');
  var ids = {}; ids[rid] = 1;
  var done = {};
  safeRows_('fact_signoff', ids).forEach(function (r) { if (isLive_(r.decision)) done[str_(r.signer).trim()] = 1; });
  return { req_id: rid, status: str_(row.status), signers: signersOf_(row), decided: Object.keys(done) };
}

/* ================= C 階段：核准 → 寫入採購系統 BOM 本（2026-09-29） =================
   全部同意（sign_）的同一次請求內寫三張表，寫 BOM 本（對照表隔天 07:05 由 mirrorTabs_ 鏡像到儀表板專用檔；BOM表 採購系統直接讀 BOM 本）
     ① dim_sku：只新增 🆕 品項（品名、BOM別名都對不到、且 BOM 會用到的）；已有主檔的品項一格都不動。來源＝「自動草稿 日期（新品申請 R-…）」→ 採購系統主檔頁「🆕 待確認」
        dim_sku 兩個檔都寫（BOM 本＋儀表板專用檔，經營者 2026-09-29 裁定）：採購系統主檔頁新增品項的編號看專用檔算，只寫 BOM 本會在 07:05 前撞號蓋掉新品草稿；
        專用檔寫不到不擋（BOM 本為準，07:05 會同步），只提醒這段期間不要在主檔頁新增品項
     ② BOM表：用料（全部）逐列（同品項同單位合併：食材耗材加總、器具模具取最大，同「複製 BOM 表貼上列」）；這支甜點在 BOM 表已經有配方 → 整段不寫
     ③ 產品名稱對照表：一列（POS 名＝BOM 名＝商品正式名稱、起訖＝檔期起訖、主／次類別）；同一個 POS 名已經有 → 不寫。最後寫：它是備料啟動開關
   安全網：寫前在「食譜系統_申請單」建 backup_日期_單號 分頁（各表改前列數、指紋、目標範圍原內容、要寫的每一列）；
          每張表寫完立刻讀回逐格比對；任何一步失敗 → 把這次已寫的列刪掉（內容相同才刪）→ 狀態「已核准（採購寫入失敗）」可重試；
          fact_push 分頁逐表記錄；同一張單寫過且未還原 → 不重寫；rollbackReq(單號)／管理者「還原」只刪這張單寫的列 */
var PURCHASE_ID_DEFAULT = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';   /* BOM 本；指令碼屬性 PURCHASE_ID 有值時改用它（測試副本用） */
var DASH_ID_DEFAULT = '1FF7lW3JINR0-Id7MMYSRkoRzA1BYdqktO94NbzAFYG0';       /* 儀表板專用檔（只寫 dim_sku）；指令碼屬性 DASH_PURCHASE_ID 可覆蓋 */
var PT = { map: '產品名稱對照表', bom: 'BOM表', sku: 'dim_sku', sku2: 'dim_sku（儀表板專用檔）' };
var PT_COLS = { map: 8, bom: 5 };                         /* 讀寫範圍（欄數）；dim_sku 依表頭 */
var PT_HEAD = {                                            /* 表頭對不上就停，不寫 */
  map: { 1: 'POS 資料產品名稱', 2: '手動對應產品名稱', 5: '起始有效日', 6: '結束有效日', 7: '主類別', 8: '次類別' },
  bom: { 1: '甜點名稱', 2: '食材/器具名稱', 3: '數量', 4: '單位', 5: '容器' }
};
var MAP_WRITE_COLS = [1, 2, 5, 6, 7, 8];                   /* 對照表 C／D 欄（偵測新產品公式區）永遠不寫 */
var SKU_NEED = ['sku_id', '品名', '品類別', '廠商', '使用單位', '採購單位', '每採購單位內容量', '單價', '來源'];
var SKU_PREFIX = { '食材': 'F', '耗材': 'C', '器具': 'T', '模具': 'M' };   /* 與主檔慣例、semiApi 相同 */
var DRAFT_TAG = '自動草稿';                                /* 採購系統主檔頁以「來源」開頭這四個字列入「🆕 待確認」 */
var DEFAULT_MAIN = '限定甜點';
var PUSHABLE = { '已核准': 1, '已核准（採購寫入失敗）': 1 };
var PUSH_OK = '已寫入採購', PUSH_FAIL = '已核准（採購寫入失敗）';

var _pss = null;
function pss_() {
  if (_pss) return _pss;
  var id = PropertiesService.getScriptProperties().getProperty('PURCHASE_ID') || PURCHASE_ID_DEFAULT;
  try { _pss = SpreadsheetApp.openById(id); } catch (e) { throw fail_('開不到採購系統 BOM 本（' + errMsg_(e) + '）', 'purchase'); }
  return _pss;
}
var _dss = null;
function dss_() {
  if (_dss) return _dss;
  var id = PropertiesService.getScriptProperties().getProperty('DASH_PURCHASE_ID') || DASH_ID_DEFAULT;
  try { _dss = SpreadsheetApp.openById(id); } catch (e) { throw fail_('開不到儀表板專用檔（' + errMsg_(e) + '）', 'purchase'); }
  return _dss;
}
function ptName_(v) { return str_(v).trim(); }
function ptRead_(key) {
  var sh = key === 'sku2' ? dss_().getSheetByName('dim_sku') : pss_().getSheetByName(PT[key]);
  if (!sh) throw fail_((key === 'sku2' ? '儀表板專用檔' : 'BOM 本') + '找不到分頁「' + PT[key] + '」', 'purchase');
  var last = sh.getLastRow(), w = PT_COLS[key] || sh.getLastColumn();
  if (last < 1 || w < 1) throw fail_('BOM 本「' + PT[key] + '」是空的', 'purchase');
  var all = sh.getRange(1, 1, last, w).getValues(), H = all[0].map(function (x) { return String(x).trim(); });
  var need = PT_HEAD[key];
  if (need) Object.keys(need).forEach(function (c) {
    if (H[c - 1] !== need[c]) throw fail_('BOM 本「' + PT[key] + '」第 ' + c + ' 欄表頭是「' + H[c - 1] + '」，應為「' + need[c] + '」——表格改過版面，先停下不寫', 'purchase');
  });
  var col = {};
  H.forEach(function (h, j) { if (h && col[h] === undefined) col[h] = j; });
  if (key === 'sku' || key === 'sku2') SKU_NEED.forEach(function (h) { if (col[h] === undefined) throw fail_(PT[key] + ' 找不到欄位「' + h + '」——先停下不寫', 'purchase'); });
  var vals = all.slice(1), lastA = 1;
  for (var i = vals.length - 1; i >= 0; i--) if (ptName_(vals[i][0]) !== '') { lastA = i + 2; break; }
  return { key: key, sh: sh, H: H, col: col, w: w, vals: vals, lastA: lastA };
}
/* 指紋：只算會用到的欄（對照表不含 C／D 公式區），看「還原後是否回到寫入前」 */
function ptDigest_(tab) {
  var cols = tab.key === 'map' ? MAP_WRITE_COLS.map(function (c) { return c - 1; }) : null;
  var rows = tab.vals.slice(0, Math.max(0, tab.lastA - 1)).map(function (r) {
    return (cols ? cols.map(function (j) { return r[j]; }) : r).map(ptNorm_);
  });
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(rows));
  return d.slice(0, 8).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('') + '·' + rows.length + '列';
}
function ptNorm_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  if (typeof v === 'number') return Math.round(v * 1e6) / 1e6;
  return String(v === null || v === undefined ? '' : v);
}
function ptSame_(a, b) {
  var x = ptNorm_(a), y = ptNorm_(b);
  if (typeof x === 'number' || typeof y === 'number') { var nx = Number(x), ny = Number(y); if (String(x) !== '' && String(y) !== '' && isFinite(nx) && isFinite(ny)) return Math.abs(nx - ny) < 1e-9; }
  return String(x) === String(y);
}
function ymd_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}
function r4_(n) { return Math.round(n * 1e4) / 1e4; }

/* 規劃要寫什麼（不寫任何東西）：預覽、實際寫入共用 */
function planPush_(rid) {
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var pl;
  try { pl = JSON.parse(readPayload_(t, row._row) || '{}') || {}; } catch (e) { throw fail_('申請單內容讀不出來（JSON 壞了），無法寫入採購系統', 'payload'); }
  var h = pl.head || {};
  var name = ptName_(h.fname) || ptName_(h.name) || ptName_(row['商品正式名稱']) || ptName_(row['商品暫定名稱']);
  if (!name) throw fail_('申請單沒有商品名稱，無法寫入採購系統', 'noname');
  var cid = str_(row.campaign_id), ct = load_('fact_campaign');
  var camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === cid; })[0] || null;
  var ids = {}; ids[rid] = 1;
  var cont = {};
  safeRows_('fact_req_unit', ids).forEach(function (u) { var k = ptName_(u['品項']); if (k) cont[k] = ptName_(u['容器']); });
  var semi = {};
  (pl.semis || []).forEach(function (s) { [s.name, s.link].forEach(function (x) { var n = ptName_(x); if (n) semi[n] = 1; }); });
  var warn = [];
  var lines = (pl.lines || []).map(function (l) {
    return { name: ptName_(l.name), qty: Number(l.qty) || 0, unit: ptName_(l.unit), cat: ptName_(l.cat), vendor: ptName_(l.vendor),
      pkgCost: Number(l.pkg_cost) || 0, pkgSpec: ptName_(l.pkg_spec), pack: Number(l.pack) || 0, zone: ptName_(l.zone),
      shelf: ptName_(l.shelf).replace(/\s*[|｜]\s*/g, '｜'), vname: ptName_(l.vname) };
  }).filter(function (l) { return l.name; });
  var P = { map: ptRead_('map'), bom: ptRead_('bom'), sku: ptRead_('sku'), sku2: null };
  try { P.sku2 = ptRead_('sku2'); } catch (e) { P.sku2 = null; P.sku2err = errMsg_(e); }

  /* ① dim_sku：品名或 BOM別名（trim 後全等）對得到＝已有主檔 */
  var S = P.sku, cN = S.col['品名'], cA = S.col['BOM別名'], cU = S.col['使用單位'], known = {};
  S.vals.forEach(function (r) {
    var n = ptName_(r[cN]); if (n) known[n] = { name: n, unit: ptName_(r[cU]) };
    if (cA !== undefined) String(r[cA] || '').split(/[,，、;；]/).forEach(function (a) { a = a.trim(); if (a && !known[a]) known[a] = { name: ptName_(r[cN]), unit: ptName_(r[cU]), alias: true }; });
  });
  var maxNo = {}, width = {};
  [S, P.sku2].forEach(function (T) {   /* 編號看兩個檔的最大號（專用檔可能有採購系統剛建、BOM 本也有的品項） */
    if (!T) return;
    T.vals.forEach(function (r) {
      var m = ptName_(r[T.col['sku_id']]).match(/^([A-Z])(\d+)$/);
      if (m) { var n = parseInt(m[2], 10); if (!(maxNo[m[1]] >= n)) maxNo[m[1]] = n; width[m[1]] = Math.max(width[m[1]] || 3, m[2].length); }
    });
  });
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'), skuRows = [], seen = {}, existing = [], used = {};
  lines.forEach(function (l) { if (l.qty > 0) used[l.name] = 1; });   /* 只建 BOM 會用到的（數量 > 0），不然主檔會多出沒人用的孤兒 */
  lines.forEach(function (l) {
    if (seen[l.name] || !used[l.name]) return; seen[l.name] = 1;
    var k = known[l.name];
    if (k) {
      existing.push(l.name);
      if (l.unit && k.unit && l.unit !== k.unit) warn.push('「' + l.name + '」申請單寫 ' + l.unit + '、主檔使用單位是 ' + k.unit + '：採購系統要有單位換算（dim_unitconv）才算得到用量');
      return;
    }
    var cat = SKU_PREFIX[l.cat] ? l.cat : '食材', isSemi = !!semi[l.name];
    var pre = SKU_PREFIX[cat], no = (maxNo[pre] || 0) + 1; maxNo[pre] = no;
    var id = pre + ('000000' + no).slice(-(width[pre] || 3));
    var u = l.unit || 'g', buy = u, pack = 1, price = 0;
    if (l.pack > 1) { var cm = l.pkgSpec.match(/[包袋箱盒罐瓶桶組件套盤打]/); buy = cm ? cm[0] : '包'; pack = l.pack; price = l.pkgCost; }
    else if (l.pack === 1) price = l.pkgCost;
    var src = DRAFT_TAG + ' ' + today + '（新品申請 ' + rid + '：' + name + (isSemi ? '；店製半成品' : '') + (l.vname && l.vname !== l.name ? '；廠商品名 ' + l.vname : '')
      + ((l.pkgCost || l.pkgSpec) ? '；原包裝 ' + (l.pkgCost ? l.pkgCost + ' 元' : '—') + '／' + (l.pkgSpec || '—') : '') + '）';
    var o = { sku_id: id, '品名': l.name, '品類別': cat, '廠商': l.vendor || (isSemi ? '半成品' : ''), '使用單位': u, '採購單位': buy,
      '每採購單位內容量': pack, '單價': price, '預設分區': l.zone, '來源': src, '效期': l.shelf };
    var arr = S.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; });
    var arr2 = P.sku2 ? P.sku2.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; }) : null;
    skuRows.push({ sku_id: id, name: l.name, vname: l.vname, cat: cat, vendor: o['廠商'], unit: u, buy: buy, pack: pack, price: price, semi: isSemi, arr: arr, arr2: arr2 });
  });

  /* ② BOM表：同品項＋同單位合併（同前端「複製 BOM 表貼上列」）；只收數量 > 0 */
  var agg = {}, order = [];
  lines.forEach(function (l) {
    if (!(l.qty > 0)) return;
    var u = l.unit || 'g', key = l.name + '|' + u, tool = l.cat === '器具' || l.cat === '模具';
    if (!agg[key]) { agg[key] = { m: l.name, u: u, q: 0, tool: tool }; order.push(key); }
    agg[key].q = tool ? Math.max(agg[key].q, l.qty) : agg[key].q + l.qty;
  });
  var bomRows = order.map(function (k) { var a = agg[k]; return [name, a.m, r4_(a.q), a.u, cont[a.m] || '']; });
  var bomHave = P.bom.vals.filter(function (r) { return ptName_(r[0]) === name; }).length;
  var bom = { rows: bomRows, skip: bomHave ? 'BOM 表已經有「' + name + '」的配方 ' + bomHave + ' 列，沒有重寫（要改配方請在 BOM 本手動改）' : (bomRows.length ? '' : '用料（全部）沒有數量 > 0 的品項') };

  /* ③ 對照表：主類別＝對照表裡同一個次類別最常用的主類別（沒有就「限定甜點」）；次類別＝POS 分類（沒選或選「無」就用檔期名稱） */
  var catTxt = ptName_(h.cat_text || (h.cat === '__other' ? h.catOther : h.cat));
  var campName = camp ? ptName_(camp['檔期名稱']) : '';
  var sub = (catTxt && catTxt !== '無') ? catTxt : (campName || catTxt || '無');
  var cnt = {}, main = '', best = 0;
  P.map.vals.forEach(function (r) { if (ptName_(r[7]) === sub) { var g = ptName_(r[6]); if (g) cnt[g] = (cnt[g] || 0) + 1; } });
  Object.keys(cnt).forEach(function (g) { if (cnt[g] > best) { best = cnt[g]; main = g; } });
  main = main || DEFAULT_MAIN;
  var d1 = camp ? normDate_(camp['起']) : '', d2 = camp ? normDate_(camp['迄']) : '';
  if (!d1 || !d2) warn.push('檔期沒有填起訖日：對照表的有效日期留空，採購系統的新品備料不會啟動（到對照表補日期即可）');
  var mapHave = P.map.vals.filter(function (r) { return ptName_(r[0]) === name; })[0];
  var map = { row: [name, name, '', '', ymd_(d1), ymd_(d2), main, sub], show: [name, name, d1, d2, main, sub],
    skip: mapHave ? '對照表已經有「' + name + '」（' + normDate_(mapHave[4]) + '～' + normDate_(mapHave[5]) + '），沒有新增（重新上架請到對照表改日期）' : '' };
  if (skuRows.length && !P.sku2) warn.push('儀表板專用檔讀不到（' + P.sku2err + '）：主檔新品項只寫 BOM 本、明天 07:05 同步；這段期間請不要在採購系統主檔頁新增品項（會撞號）');
  if (!skuRows.length && bom.skip && map.skip) warn.push('三張表都不用寫（都已經有了）');
  return { rid: rid, name: name, campaign: campName, from: d1, to: d2, sku: { rows: skuRows, existing: existing }, bom: bom, map: map, warn: warn, P: P };
}
function planView_(pl) {
  return { req_id: pl.rid, name: pl.name, campaign: pl.campaign, from: pl.from, to: pl.to, warn: pl.warn, dash: !!pl.P.sku2,
    sku: pl.sku.rows.map(function (r) { return { sku_id: r.sku_id, name: r.name, vname: r.vname || '', cat: r.cat, vendor: r.vendor, unit: r.unit, buy: r.buy, pack: r.pack, price: r.price, semi: r.semi }; }),
    existing: pl.sku.existing, bom: pl.bom.rows, bomSkip: pl.bom.skip, map: pl.map.show, mapSkip: pl.map.skip };
}

/* 寫入紀錄：這張單最近一批是否「寫了、還沒還原」 */
function pushState_(rid) {
  var rows = safeRows_('fact_push', (function () { var o = {}; o[rid] = 1; return o; })());
  var last = null;
  rows.forEach(function (r) { if (r.ok === 'Y' && (r.action === 'write' || r.action === 'rollback')) last = r; });
  if (!last || last.action !== 'write') return { live: false, rows: rows };
  var b = last.batch, tables = {};
  rows.forEach(function (r) { if (r.batch === b && r.action === 'write' && r.ok === 'Y') { try { tables[r.table] = JSON.parse(r.detail); } catch (e) { } } });
  return { live: true, batch: b, ts: last.ts, backup: last.backup, tables: tables, rows: rows };
}
function pushLog_(rec) {
  try { var t = load_('fact_push'); put_(t, nextRow_(t), rec); } catch (e) { /* 記錄失敗不影響主流程 */ }
}
/* 在 A 欄最後一列之後寫；寫前再讀一次目標範圍（要是空的），寫完讀回逐格比對 */
function ptAppend_(tab, rows) {
  var n = rows.length, sh = tab.sh, w = tab.key === 'map' ? 8 : tab.w;
  var colA = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues(), start = 2;
  for (var i = colA.length - 1; i >= 0; i--) if (ptName_(colA[i][0]) !== '') { start = i + 2; break; }
  if (start + n - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + n - 1 - sh.getMaxRows() + 50);
  var before = sh.getRange(start, 1, n, w).getValues();
  var busy = before.some(function (r, i) { return tab.key === 'map' ? MAP_WRITE_COLS.some(function (c) { return ptNorm_(r[c - 1]) !== ''; }) : r.some(function (v) { return ptNorm_(v) !== ''; }); });
  if (busy) throw fail_('「' + PT[tab.key] + '」第 ' + start + ' 列之後不是空的（可能剛好有人在寫），這次先不寫', 'busy');
  var put = rows.map(function (r) { return r.map(function (v) { return typeof v === 'string' ? safe_(v) : v; }); });
  /* 整欄都是有字的文字 → 先設純文字格式（避免品名像「1/2」被試算表轉成日期）；空白欄不動格式（之後有人填數字才不會變文字） */
  var wcols = tab.key === 'map' ? MAP_WRITE_COLS.map(function (c) { return c - 1; }) : rows[0].map(function (x, j) { return j; }).filter(function (j) { return j < w; });
  wcols.forEach(function (j) { if (rows.every(function (r) { return typeof r[j] === 'string' && r[j] !== ''; })) sh.getRange(start, j + 1, n, 1).setNumberFormat('@'); });
  if (tab.key === 'map') {
    sh.getRange(start, 1, n, 2).setValues(put.map(function (r) { return r.slice(0, 2); }));
    sh.getRange(start, 5, n, 4).setValues(put.map(function (r) { return r.slice(4, 8); }));
  } else sh.getRange(start, 1, n, w).setValues(put.map(function (r) { return r.slice(0, w); }));
  SpreadsheetApp.flush();
  var back = sh.getRange(start, 1, n, w).getValues();
  back.forEach(function (r, i) {
    var cols = tab.key === 'map' ? MAP_WRITE_COLS.map(function (c) { return c - 1; }) : r.map(function (x, j) { return j; });
    cols.forEach(function (j) { if (!ptSame_(r[j], rows[i][j])) throw fail_('「' + PT[tab.key] + '」第 ' + (start + i) + ' 列第 ' + (j + 1) + ' 欄讀回不一致（寫「' + ptNorm_(rows[i][j]) + '」讀到「' + ptNorm_(r[j]) + '」）', 'verify'); });
  });
  return { table: tab.key, start: start, rows: rows.map(function (r) { return r.map(ptNorm_); }), before: before.map(function (r) { return r.map(ptNorm_); }) };
}
/* 刪掉寫過的列：只刪內容仍相符的（對照表比 A／B；BOM 比甜點＋品項＋數量＋單位；dim_sku 比 sku_id＋品名），由下往上刪 */
function ptRemove_(tab, rec) {
  var want = rec.rows || [], hit = [], used = {}, miss = [];
  function same(r, x) {
    if (tab.key === 'map') return ptName_(r[0]) === x[0] && ptName_(r[1]) === x[1];
    if (tab.key === 'bom') return ptName_(r[0]) === x[0] && ptName_(r[1]) === x[1] && ptSame_(r[2], x[2]) && ptName_(r[3]) === String(x[3]);
    var ci = tab.col['sku_id'], cn = tab.col['品名'];
    return ptName_(r[ci]) === String(x[ci]) && ptName_(r[cn]) === String(x[cn]);
  }
  want.forEach(function (x, k) {
    var pref = (rec.start || 2) + k, found = -1;
    if (pref >= 2 && tab.vals[pref - 2] && !used[pref] && same(tab.vals[pref - 2], x)) found = pref;
    else for (var i = tab.vals.length - 1; i >= 0; i--) { var rn = i + 2; if (!used[rn] && same(tab.vals[i], x)) { found = rn; break; } }
    if (found > 0) { used[found] = 1; hit.push(found); } else miss.push(x[tab.key === 'sku' ? tab.col['品名'] : 1] || x[0]);
  });
  hit.sort(function (a, b) { return b - a; });
  var i = 0;
  while (i < hit.length) {   /* 連續列一次刪 */
    var top = hit[i], n = 1;
    while (i + n < hit.length && hit[i + n] === top - n) n++;
    tab.sh.deleteRows(top - n + 1, n);
    i += n;
  }
  if (hit.length) SpreadsheetApp.flush();
  return { removed: hit.length, miss: miss };
}
function backupTab_(pl, batch) {
  var ss = ss_(), base = 'backup_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd') + '_' + pl.rid, nm = base, k = 2;
  while (ss.getSheetByName(nm)) nm = base + '_' + (k++);
  var sh = ss.insertSheet(nm, ss.getSheets().length);
  var out = [['批次', batch, '申請單', pl.rid, '甜點', pl.name],
    ['表', '分頁', '寫入前 A 欄最後一列', '寫入前指紋', '這次要寫的列數', '說明']];
  ['sku', 'bom', 'map', 'sku2'].forEach(function (key) {
    var tab = pl.P[key], n = (key === 'sku' || key === 'sku2') ? pl.sku.rows.length : key === 'bom' ? (pl.bom.skip ? 0 : pl.bom.rows.length) : (pl.map.skip ? 0 : 1);
    if (!tab) { out.push([key, PT[key], '', '讀不到', '0', pl.P.sku2err || '']); return; }
    tab.digest = ptDigest_(tab);
    out.push([key, PT[key], String(tab.lastA), tab.digest, String(n), key === 'bom' && pl.bom.skip ? pl.bom.skip : key === 'map' && pl.map.skip ? pl.map.skip : '']);
  });
  out.push(['', '', '', '', '', '']);
  out.push(['表', '要寫的內容（JSON，一列一筆）', '', '', '', '']);
  pl.sku.rows.forEach(function (r) { out.push(['sku', JSON.stringify(r.arr.map(ptNorm_)), '', '', '', '']); });
  if (!pl.bom.skip) pl.bom.rows.forEach(function (r) { out.push(['bom', JSON.stringify(r.map(ptNorm_)), '', '', '', '']); });
  if (!pl.map.skip) out.push(['map', JSON.stringify(pl.map.row.map(ptNorm_)), '', '', '', '']);
  var rg = sh.getRange(1, 1, out.length, 6);
  rg.setNumberFormat('@');
  rg.setValues(out.map(function (r) { return r.map(function (v) { return safe_(String(v)); }); }));
  return nm;
}

/* 寫入（呼叫端要已經拿到鎖）：成功回摘要；失敗先刪掉這次寫的列再丟錯 */
function pushToPurchase_(rid, who) {
  var st0 = pushState_(rid);
  if (st0.live) throw fail_('這張單在 ' + st0.ts + ' 已經寫入採購系統，不會重複寫（要重寫請先由管理者還原）', 'dup');
  var pl = planPush_(rid), batch = stamp_() + '-' + Utilities.getUuid().slice(0, 4), bk = backupTab_(pl, batch), done = [], now = now_(), soft = [];
  try {
    if (pl.sku.rows.length) done.push(ptAppend_(pl.P.sku, pl.sku.rows.map(function (r) { return r.arr; })));
    if (pl.sku.rows.length && pl.P.sku2) {   /* 專用檔：寫不到不擋（BOM 本為準，07:05 同步） */
      try { done.push(ptAppend_(pl.P.sku2, pl.sku.rows.map(function (r) { return r.arr2; }))); }
      catch (e2) { soft.push('儀表板專用檔主檔沒寫到（' + errMsg_(e2) + '）：明天 07:05 會同步；這段期間請不要在採購系統主檔頁新增品項（會撞號）'); }
    }
    if (!pl.bom.skip) done.push(ptAppend_(pl.P.bom, pl.bom.rows));
    if (!pl.map.skip) done.push(ptAppend_(pl.P.map, [pl.map.row]));
  } catch (e) {
    var undo = [];
    done.slice().reverse().forEach(function (d) {
      try { var tab = ptRead_(d.table), r = ptRemove_(tab, d); undo.push(PT[d.table] + ' 刪回 ' + r.removed + ' 列' + (r.miss.length ? '（' + r.miss.length + ' 列找不到：' + r.miss.join('、') + '）' : '')); }
      catch (e2) { undo.push(PT[d.table] + ' 刪回失敗：' + errMsg_(e2)); }
    });
    pushLog_({ ts: now, req_id: rid, batch: batch, action: 'fail', table: '', rows: '', detail: JSON.stringify(done.map(function (d) { return { table: d.table, start: d.start, n: d.rows.length }; })),
      backup: bk, ok: 'N', msg: (who || '') + '｜' + errMsg_(e) + (undo.length ? '｜' + undo.join('；') : '') });
    throw fail_('寫入採購系統失敗：' + errMsg_(e) + (undo.length ? '（已把這次寫的列刪回：' + undo.join('；') + '）' : '（沒有寫進任何一列）'), 'push');
  }
  var show = { sku: pl.sku.rows.map(function (r) { return [r.sku_id, r.name, r.cat, r.vendor]; }), bom: pl.bom.rows.map(function (r) { return r.map(ptNorm_); }), map: [pl.map.show] };
  show.sku2 = show.sku;
  done.forEach(function (d) {
    pushLog_({ ts: now, req_id: rid, batch: batch, action: 'write', table: d.table, rows: String(d.rows.length),
      detail: JSON.stringify({ start: d.start, rows: d.rows, show: show[d.table] || [] }), backup: bk, ok: 'Y', msg: (who || '') + '｜' + PT[d.table] + ' 第 ' + d.start + '～' + (d.start + d.rows.length - 1) + ' 列' });
  });
  if (!done.length) pushLog_({ ts: now, req_id: rid, batch: batch, action: 'write', table: '', rows: '0', detail: '{}', backup: bk, ok: 'Y', msg: (who || '') + '｜三張表都已經有了，沒有寫' });
  var v = planView_(pl);
  v.written = done.map(function (d) { return { table: d.table, name: PT[d.table], start: d.start, n: d.rows.length }; });
  v.backup = bk; v.batch = batch; v.warn = v.warn.concat(soft);
  if (soft.length) pushLog_({ ts: now, req_id: rid, batch: batch, action: 'warn', table: 'sku2', rows: '0', detail: '{}', backup: bk, ok: 'Y', msg: soft.join('｜') });
  return v;
}
function rollbackPush_(rid, who) {
  var st = pushState_(rid);
  if (!st.live) throw fail_('這張單沒有「已寫入、未還原」的採購系統紀錄', 'none');
  var res = [], now = now_(), bkDig = {};
  try {
    var bs = ss_().getSheetByName(st.backup);
    if (bs) bs.getRange(3, 1, 4, 4).getValues().forEach(function (r) { bkDig[String(r[0])] = String(r[3]); });
  } catch (e) { }
  ['map', 'bom', 'sku', 'sku2'].forEach(function (key) {
    var rec = st.tables[key];
    if (key === 'sku2' && !rec) return;   /* 當初沒寫到專用檔 */
    var tab;
    try { tab = ptRead_(key); } catch (e) {
      if (key !== 'sku2') throw e;
      res.push({ table: key, name: PT[key], removed: 0, miss: [], digest: '', backToBackup: null, error: errMsg_(e) });
      pushLog_({ ts: now, req_id: rid, batch: st.batch, action: 'rollback', table: key, rows: '0', detail: '{}', backup: st.backup, ok: 'N', msg: (who || '') + '｜專用檔沒刪到（' + errMsg_(e) + '），明天 07:05 會跟 BOM 本同步' });
      return;
    }
    var r = rec ? ptRemove_(tab, rec) : { removed: 0, miss: [] };
    var after = ptDigest_(ptRead_(key)), same = bkDig[key] ? after === bkDig[key] : null;
    res.push({ table: key, name: PT[key], removed: r.removed, miss: r.miss, digest: after, backToBackup: same });
    pushLog_({ ts: now, req_id: rid, batch: st.batch, action: 'rollback', table: key, rows: String(r.removed),
      detail: JSON.stringify({ miss: r.miss, digest: after, before: bkDig[key] || '' }), backup: st.backup, ok: 'Y',
      msg: (who || '') + '｜' + PT[key] + ' 刪 ' + r.removed + ' 列' + (r.miss.length ? '，' + r.miss.length + ' 列找不到' : '') + (same === true ? '｜已回到寫入前' : same === false ? '｜和寫入前不同（期間可能有人改過這張表）' : '') });
  });
  return { tables: res, backup: st.backup };
}
function setStatus_(rid, status) {
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (row) putCols_(t, row._row, { status: status }, function (h) { return h === 'status'; });
}
function reqStatus_(rid) {
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  return str_(row.status);
}
function canPush_(auth) { return auth && auth.kind === 'role' && (auth.role === ADMIN || can_(auth, 'P:*')); }

/* 動作：預覽（主廚／管理者，不寫）、重試寫入（主廚／管理者）、還原（管理者） */
function pushPreview_(p, auth) {
  if (!canPush_(auth)) throw fail_('只有主廚或管理者可以看', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var st = pushState_(rid);
  return { id: rid, data: { plan: planView_(planPush_(rid)), live: st.live, ts: st.ts || '', backup: st.backup || '' }, msg: '預覽（沒有寫入）' };
}
function pushPurchase_(p, auth) {
  if (!canPush_(auth)) throw fail_('只有主廚或管理者可以寫入採購系統', 'perm');
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var st = reqStatus_(rid);
  if (!PUSHABLE[st]) throw fail_('這張單目前是「' + st + '」，' + (st === PUSH_OK ? '已經寫入過了' : '要全部主管同意（已核准）才能寫入採購系統'), 'locked', { status: st });
  try {
    var v = pushToPurchase_(rid, auth.role);
    setStatus_(rid, PUSH_OK);
    return { id: rid, data: { status: PUSH_OK, push: v }, msg: '寫入採購系統：' + pushMsg_(v) };
  } catch (e) {
    if (e && e.code === 'dup') { setStatus_(rid, PUSH_OK); throw fail_(errMsg_(e), 'dup', { status: PUSH_OK }); }
    setStatus_(rid, PUSH_FAIL);
    throw fail_(errMsg_(e), (e && e.code) || 'push', { status: PUSH_FAIL });
  }
}
function rollbackPurchase_(p, auth) {
  needAdmin_(auth);
  var rid = str_(p.req_id).trim();
  if (!rid) throw fail_('缺申請單號');
  var r = rollbackPush_(rid, auth.role);
  var st = reqStatus_(rid);
  if (st === PUSH_OK) setStatus_(rid, '已核准');
  return { id: rid, data: { status: st === PUSH_OK ? '已核准' : st, rollback: r }, msg: '還原：' + r.tables.map(function (x) { return x.name + ' 刪 ' + x.removed; }).join('、') };
}
function pushMsg_(v) {
  var w = v.written || [];
  return w.length ? w.map(function (d) { return d.name + ' ' + d.n + ' 列'; }).join('、') : '三張表都已經有了，沒有寫';
}
/* 給經營者在編輯器用：把某張單寫進採購系統的列刪掉（例：rollbackReq('R-20260929-0151')） */
function rollbackReq(rid) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('系統忙碌中，稍後再試');
  try {
    var r = rollbackPush_(String(rid || '').trim(), '編輯器');
    if (reqStatus_(String(rid).trim()) === PUSH_OK) setStatus_(String(rid).trim(), '已核准');
    Logger.log(JSON.stringify(r));
    return r;
  } finally { lock.releaseLock(); }
}

function pushInfo_(rid) {
  var st = pushState_(rid), rows = st.rows || [];
  return { live: st.live, ts: st.ts || '', backup: st.backup || '',
    written: st.live ? Object.keys(st.tables).filter(function (k) { return PT[k]; }).map(function (k) { var d = st.tables[k] || {}; return { table: k, name: PT[k], start: d.start, n: (d.rows || []).length, show: d.show || [] }; }) : [],
    history: rows.map(function (r) { return { ts: r.ts, action: r.action, table: r.table, rows: r.rows, ok: r.ok, msg: r.msg }; }) };
}

/* ================= D 階段：採購系統「🆕 檔期新品」公開查詢（2026-09-29） =================
   只給已核准（含已寫入採購、採購寫入失敗）、檔期沒有封存的申請單；每支甜點只回品項名＋廠商品名＋貼紙名稱＋購買連結＋供應商覆寫＋提供方式＋區域＋容器＋是否新品項（E 階段加廠商品名、區域、容器）。
   不回定價、成本、用量、配方步驟、簽核人（BOM 表本來就公開、用量採購系統自己讀）。結果快取 2 分鐘，12 店同時開不會重讀試算表。 */
var PUB_ST = { '已核准': 1, '已核准（採購寫入失敗）': 1, '已寫入採購': 1 };
function newItemsPub_() {
  var cache = CacheService.getScriptCache(), ck = 'pub:newitems:v3', hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var t = load_('fact_recipe_req', true), ct = load_('fact_campaign'), camp = {};
  ct.rows.forEach(function (r) { camp[str_(r.campaign_id)] = r; });
  var ok = t.rows.filter(function (r) { return PUB_ST[str_(r.status)] && str_((camp[str_(r.campaign_id)] || {}).status) !== '封存'; }), ids = {};   /* 封存的檔期（測試資料）不出現 */
  ok.forEach(function (r) { ids[str_(r.req_id)] = 1; });
  var units = {};
  (ok.length ? safeRows_('fact_req_unit', ids) : []).forEach(function (u) { (units[u.req_id] = units[u.req_id] || {})[str_(u['品項']).trim()] = u; });
  var out = ok.map(function (r) {
    var pl = {};
    try { pl = JSON.parse(readPayload_(t, r._row) || '{}') || {}; } catch (e) { pl = {}; }
    var h = pl.head || {}, c = camp[str_(r.campaign_id)] || {}, U = units[str_(r.req_id)] || {}, seen = {};
    var lines = (pl.lines || []).map(function (l) {
      var nm = str_(l.name).trim(); if (!nm || seen[nm]) return null; seen[nm] = 1;
      var u = U[nm] || {};
      var o = { name: nm, is_new: !!l.is_new, vname: str_(l.vname).trim(), sticker: str_(l.sticker).trim(), supply: str_(l.supply).trim(),
        zone: str_(l.zone).trim(), container: str_(u['容器']).trim(),
        link: str_(u['購買連結']).trim(), vendor_override: str_(u['供應商覆寫']).trim(), note: str_(u['品項備註']).trim() };
      return (o.vname || o.sticker || o.link || o.vendor_override || o.supply || o.note || o.zone || o.container || o.is_new) ? o : null;
    }).filter(Boolean);
    return { req_id: str_(r.req_id), status: str_(r.status), dessert: str_(h.fname).trim() || str_(h.name).trim() || str_(r['商品正式名稱']).trim() || str_(r['商品暫定名稱']).trim(),
      campaign: str_(c['檔期名稱']), from: normDate_(c['起']), to: normDate_(c['迄']), lines: lines };
  }).filter(function (x) { return x.dessert; });
  var res = { at: now_(), items: out };
  try { var js = JSON.stringify(res); if (js.length < 90000) cache.put(ck, js, 120); } catch (e) { }
  return res;
}

/* ================= 試算表工具 ================= */
var _ss = null;
function ss_() {
  if (_ss) return _ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw fail_('申請單系統尚未設定（管理者要先在編輯器執行 setup）', 'setup');
  _ss = SpreadsheetApp.openById(id);
  return _ss;
}

/* 讀整張分頁成物件；skipPayload＝不讀 payload_1～4（快很多，列表／計數用） */
function load_(name, skipPayload) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw fail_('找不到分頁 ' + name + '（管理者要先執行 setup）', 'setup');
  var H = TABS[name], last = sh.getLastRow(), rows = [];
  if (last > 1) {
    var cols = H.map(function (h, j) { return (skipPayload && h.indexOf('payload_') === 0) ? -1 : j; });
    var runs = runs_(cols.filter(function (j) { return j >= 0; }));
    var parts = runs.map(function (rn) { return { a: rn[0], v: sh.getRange(2, rn[0] + 1, last - 1, rn[1] - rn[0] + 1).getValues() }; });
    for (var i = 0; i < last - 1; i++) {
      var o = { _row: i + 2 };
      parts.forEach(function (pt) { pt.v[i].forEach(function (val, k) { o[H[pt.a + k]] = val; }); });
      rows.push(o);
    }
  }
  return { sh: sh, H: H, rows: rows };
}
/* 連續欄號分段：[0,1,2,5,6] → [[0,2],[5,6]] */
function runs_(idx) {
  var out = [];
  idx.forEach(function (j) { var r = out[out.length - 1]; if (r && r[1] === j - 1) r[1] = j; else out.push([j, j]); });
  return out;
}
function rowsOf_(name, ids) {
  var t = load_(name);
  return t.rows.filter(function (r) { return ids[str_(r.req_id)]; }).map(function (r) { return plain_(t.H, r); });
}
/* 24 小時內新增的申請單裡，找內容 created_at 相同的那一張（只讀 payload_1 開頭；created_at 在 JSON 前 100 字內） */
function findByCreated_(t, created) {
  var lim = Utilities.formatDate(new Date(Date.now() - 24 * 3600e3), TZ, 'yyyy-MM-dd HH:mm:ss');
  var key = '"created_at":' + JSON.stringify(created), a = t.H.indexOf('payload_1') + 1;
  for (var i = t.rows.length - 1; i >= 0; i--) {
    var r = t.rows[i];
    if (str_(r.created_at) < lim) continue;
    if (str_(t.sh.getRange(r._row, a).getValue()).slice(0, 400).indexOf(key) >= 0) return r;
  }
  return null;
}
function readPayload_(t, rowNo) {
  var a = t.H.indexOf('payload_1');
  return t.sh.getRange(rowNo, a + 1, 1, SEG_N).getValues()[0].map(str_).join('');
}
function nextRow_(t) {
  var r = t.sh.getLastRow() + 1;
  if (r > t.sh.getMaxRows()) t.sh.insertRowsAfter(t.sh.getMaxRows(), 200);
  return r;
}
function put_(t, rowNo, obj) { putCols_(t, rowNo, obj, function () { return true; }); }
/* 只寫 pred 為真的欄（連續欄一次寫）；文字欄先設純文字格式再寫值 */
function putCols_(t, rowNo, obj, pred) {
  var H = t.H, idx = [];
  H.forEach(function (h, j) { if (pred(h)) idx.push(j); });
  runs_(idx).forEach(function (rn) {
    var hs = H.slice(rn[0], rn[1] + 1);
    var rg = t.sh.getRange(rowNo, rn[0] + 1, 1, hs.length);
    var txt = runs_(hs.map(function (h, k) { return NUM_COLS[h] ? -1 : k; }).filter(function (k) { return k >= 0; }));
    txt.forEach(function (tr) { t.sh.getRange(rowNo, rn[0] + tr[0] + 1, 1, tr[1] - tr[0] + 1).setNumberFormat('@'); });
    rg.setValues([hs.map(function (h) { return cell_(h, obj[h]); })]);
  });
}
function cell_(h, v) {
  if (v === null || v === undefined) return '';
  if (NUM_COLS[h]) { if (v === '') return ''; var n = Number(v); return isFinite(n) ? n : safe_(String(v)); }
  return safe_(str_(v));
}
/* 開頭是 = 或 + 的文字會被試算表當成公式執行（2026-09-28 實測：純文字格式擋不住，"=1+1" 讀回變 2）
   → 前面加 ' 強制存成文字（' 不會出現在儲存格值裡）；本來就以 ' 開頭的文字也要多加一個，否則那個 ' 會被吃掉 */
function safe_(s) { return /^[=+']/.test(s) ? "'" + s : s; }
/* 物件 → 乾淨的 JSON 值（不含 payload、不含 _row） */
function plain_(H, r) {
  var o = {};
  H.forEach(function (h) { if (h.indexOf('payload_') === 0) return; o[h] = NUM_COLS[h] ? (r[h] === '' || r[h] === undefined ? '' : r[h]) : str_(r[h]); });
  return o;
}
/* payload 切段：每段 ≤ 45,000 字；不切斷 emoji（UTF-16 代理對）；下一段不以 = + - @ ' 開頭（避免被試算表當公式或吃掉引號；寫入時 safe_ 另有一層保護） */
function split_(s) {
  var out = [], i = 0;
  while (i < s.length) {
    var end = Math.min(i + SEG_MAX, s.length);
    if (end < s.length) {
      while (end > i + 1) {
        var hi = s.charCodeAt(end - 1);
        if (hi >= 0xD800 && hi <= 0xDBFF) { end--; continue; }
        if ("=+-@'".indexOf(s.charAt(end)) >= 0) { end--; continue; }
        break;
      }
    }
    out.push(s.slice(i, end));
    i = end;
  }
  if (out.length > SEG_N) throw fail_('申請單內容太大（' + s.length + ' 字，上限 ' + (SEG_MAX * SEG_N) + ' 字），請精簡備註或製作步驟', 'toobig');
  return out;
}
function log_(who, act, id, ok, msg) {
  try {
    var sh = ss_().getSheetByName('log');
    if (!sh) return;
    var r = sh.getLastRow() + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 500);
    sh.getRange(r, 1, 1, 6).setNumberFormat('@').setValues([[now_(), safe_(String(who || '')), safe_(String(act || '')), safe_(String(id || '')), ok ? 'Y' : 'N', safe_(String(msg || '').slice(0, 500))]]);
    if (r > LOG_KEEP + 500) sh.deleteRows(2, r - 1 - LOG_KEEP);
  } catch (e) { /* log 失敗不影響主流程 */ }
}

/* ================= 小工具 ================= */
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function stamp_() { return Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmm'); }
function pad2_(n) { return n < 10 ? '0' + n : String(n); }
function str_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}
function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  var s = str_(v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}
function fail_(msg, code, extra) { var e = new Error(msg); if (code) e.code = code; if (extra) e.extra = extra; return e; }
function errMsg_(err) { return String((err && err.message) || err || '未知錯誤'); }
