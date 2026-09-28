/**
 * 食譜系統申請單 API（recipe-req-v1）
 * 專案：「食譜系統申請單 API」（mydiybc，clasp 建立；原始碼備份在 DIY repo gas/recipe-req/）
 * 資料：Google 試算表「食譜系統_申請單」（ID 存在指令碼屬性 SHEET_ID；只有 mydiybc 能開，不開連結檢視）
 * 用途：食譜系統 dashboard-recipe.html「新品申請表」的檔期、送件、載回（A 階段，2026-09-28）；多人填寫與簽核（B 階段）
 *
 * 規則
 *   - doGet（JSONP，公開、不含申請單內容）：ping／listCampaigns
 *   - doPost（JSON 字串，Content-Type text/plain，前端直接讀回 {ok, data|msg}）：
 *       每筆帶 role＋password（或 signer＋pin）；伺服器依 dim_role.fields 過濾可寫欄位；
 *       全部包 LockService 10 秒；成功、失敗都寫 log 分頁。
 *   - 申請單內容（.json v3 字串）存 payload_1～4，每格 ≤ 45,000 字（試算表單格上限 50,000 字）
 *   - 本專案只開「食譜系統_申請單」這一個試算表，不碰 BOM 本、儀表板專用檔、P&L、排班。
 *
 * 部署：網頁應用程式／執行身分＝我／存取＝所有人。之後更新一律「管理部署作業 → 編輯 → 新版本」（網址不變）。
 * 第一次使用：編輯器選 setup → 執行 → 授權（建立試算表、分頁、表頭、各角色初始密碼）。
 */

var VERSION = 'recipe-req-v1';
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
  log: ['ts', 'role', 'action', 'id', 'ok', 'msg']
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

function randPw_() {
  var cs = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789', s = '';
  var b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + ':' + Date.now() + ':' + Math.random());
  var lim = 256 - (256 % cs.length);
  for (var i = 0; i < b.length && s.length < 8; i++) { var v = (b[i] + 256) % 256; if (v < lim) s += cs.charAt(v % cs.length); }
  while (s.length < 8) s += cs.charAt(Math.floor(Math.random() * cs.length));
  return s;
}

/* ================= 入口 ================= */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || 'ping'), res;
  try {
    if (a === 'ping') res = { ok: true, data: { version: VERSION, now: now_(), ready: !!PropertiesService.getScriptProperties().getProperty('SHEET_ID') } };
    else if (a === 'listCampaigns') res = { ok: true, data: listCampaigns_() };
    else res = { ok: false, msg: '不支援的查詢：' + a + '（讀申請單內容要用 POST 並帶密碼）' };
  } catch (err) { res = { ok: false, msg: errMsg_(err) }; }
  res.act = a;   /* 回覆註明是哪個動作的結果（Google 回傳鏈偶爾會把請求導回預設 ping，前端靠這個判斷要不要重試） */
  return out_(res, p.callback);
}

var ACTIONS = { whoami: whoami_, saveCampaign: saveCampaign_, saveReq: saveReq_, saveUnit: saveUnit_, getCampaign: getCampaign_, getReq: getReq_ };
var WRITES = { saveCampaign: 1, saveReq: 1, saveUnit: 1 };
var RQ_SEC = 21600;   /* 回條保留 6 小時：同一個回條編號重送 → 直接回上次結果，不重複寫 */

function doPost(e) {
  var p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, msg: '送來的資料不是 JSON' }); }
  var act = String(p.action || ''), who = p.signer ? ('簽核:' + String(p.signer)) : String(p.role || ''), res;
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
  var rqKey = (WRITES[act] && p.rq) ? 'rq:' + Utilities.base64EncodeWebSafe(Utilities.newBlob(String(p.rq).slice(0, 120)).getBytes()) : '';
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
  return String(p.req_id || (p.req && p.req.req_id) || (p.campaign && p.campaign.campaign_id) || p.campaign_id || '');
}

/* ================= 身分 ================= */
function auth_(p) {
  if (p.signer) {   /* B 階段：簽核人 PIN 較短，屆時加「同一簽核人錯太多次暫停」（只影響那一個人） */
    var name = String(p.signer).trim(), pin = String(p.pin || '').trim();
    var s = load_('dim_signer').rows.filter(function (r) { return str_(r.name).trim() === name && signerOn_(r.enabled); })[0];
    if (!s || !pin || str_(s.pin).trim() !== pin) throw fail_('簽核人或 PIN 不對', 'auth');
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
  var qt = load_('fact_recipe_req', true);
  if (!qt.rows.some(function (r) { return str_(r.req_id) === rid; })) throw fail_('找不到申請單 ' + rid, 'notfound');
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
  return { id: rid, data: { saved: n, skipped: sk }, msg: '品項 ' + n + ' 列' + (sk.length ? '｜略過無權限欄位：' + sk.join('、') : '') };
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
    data: { campaign: plain_(ct.H, camp), reqs: reqs, units: rowsOf_('fact_req_unit', ids), signoffs: rowsOf_('fact_signoff', ids) },
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
    data: { req: req, campaign: camp ? plain_(ct.H, camp) : null, units: rowsOf_('fact_req_unit', ids), signoffs: rowsOf_('fact_signoff', ids) },
    msg: '內容 ' + req.payload.length + ' 字'
  };
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
