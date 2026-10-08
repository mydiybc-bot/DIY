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
 *       寫入動作（WRITES）包 LockService 10 秒、成功失敗都寫 log 分頁；只讀動作（v7 起）不排鎖、成功不寫 log（adminList 例外）。
 *   - 申請單內容（.json v3 字串）存 payload_1～4，每格 ≤ 45,000 字（試算表單格上限 50,000 字）
 *   - 本專案開兩個試算表：「食譜系統_申請單」（讀寫）；BOM 本（只在核准時新增列、管理者還原時刪自己寫的列）。不碰儀表板專用檔、P&L、排班。
 *
 * 部署：網頁應用程式／執行身分＝我／存取＝所有人。之後更新一律「管理部署作業 → 編輯 → 新版本」（網址不變）。
 * 第一次使用：編輯器選 setup → 執行 → 授權（建立試算表、分頁、表頭、各角色初始密碼）。
 */

var VERSION = 'recipe-req-v13';   /* v13＝2026-10-08 經營者：新品申請表下方只留「產出申請表連結（整檔、唯讀）」「簽核通過→上傳採購系統」「上傳自己做食譜系統」，簽核改由主廚用 104 表單自己送 →
   ①新動作 listReqs（主廚／管理者：曾經送件過的全部申請單清單，給「重新載入申請表」選）②pushPurchase 帶 approved104＝主廚確認 104 已簽核通過：草稿／退回／送簽中 先改成「已核准」再寫採購（log 記一筆 approve104）。
   v12.1＝v12.1＝2026-10-08 簽核連結錯誤訊息改「Email 或 104 公告裡的連結」。v12＝2026-10-08 經營者：「前一個人簽完、輪到我簽，我要怎麼知道？每天去看太煩」→ 輪到誰簽核就自動寄 Email 給誰（dim_signer.email）；
   有一支被退回、或核准 → 寄給主廚（dim_role「主廚」列的 email）；寄送紀錄 fact_notify；🔑 密碼管理可填 Email、寄測試信（adminSetRoleEmail／adminTestMail）。見下方「自動寄 Email」段
   v11＝2026-10-07 晚 經營者裁示：①整檔一次送簽、一次簽完、單支可退回（submitSignCamp／signOpenCamp／signCamp，連結 #sign=檔期編號&k=）②簽核人固定批次（dim_signer.batch；前一批全部同意才輪下一批）③新器具／模具「每店要幾個」→ 核准時寫成各店標配（dim_store_par，BOM 本＋專用檔，還原一起刪）④提供方式＝出貨中心出貨 → 主檔廠商寫「出貨中心」⑤newItemsPub 多回第一批配貨日、檔期作業時間 */
var TZ = 'Asia/Taipei';
var SEG_MAX = 45000, SEG_N = 4;       /* payload 每格上限、格數 */
var LOG_KEEP = 5000;                   /* log 分頁保留筆數 */

/* 分頁與欄序固定（改欄序前先確認前端與 B／C 階段程式） */
var TABS = {
  fact_campaign: ['campaign_id', '檔期名稱', '起', '迄', '第一批配貨日', '品牌別', '自己人搶先體驗', '配合活動',
    '作業_第一批出貨_時間', '作業_第一批出貨_備註', '作業_貼紙_時間', '作業_貼紙_備註', '作業_POS_時間', '作業_POS_備註',
    '檔期備註', 'status', 'created_at', 'updated_at', 'updated_by', '自己人搶先開賣日'],   /* 最後一欄 2026-10-05 新增（加在最後，舊欄位位置不動；migrate_ 自動補表頭） */
  fact_recipe_req: ['req_id', 'campaign_id', 'seq', '商品暫定名稱', '商品正式名稱', '定價', '成本', '利潤率', '規格', '葷素',
    '保存方式', '包裝方式', '製作時間', '預估銷售數', 'status', 'signers', 'payload_1', 'payload_2', 'payload_3', 'payload_4',
    'created_at', 'updated_at', 'updated_by'],
  fact_req_unit: ['req_id', '品項', '容器', '購買連結', '供應商覆寫', '出貨中心預估出貨量', '品項備註', 'updated_at', 'updated_by'],
  dim_role: ['role', 'password', 'fields', 'email'],   /* email：2026-10-08 新增（通知用，目前只有「主廚」那一列會收到：被退回、核准；可填多個，逗號分隔；migrate_ ⑪ 自動補表頭） */
  dim_signer: ['name', 'role', 'pin', 'enabled', 'batch', 'email'],   /* batch：2026-10-07 晚新增（簽核批次 1～9，空白＝第 1 批；migrate_ ⑩ 自動補表頭）；email：2026-10-08 新增（輪到他簽核時寄通知；migrate_ ⑪） */
  fact_signoff: ['req_id', 'signer', 'decision', 'comment', 'ts'],
  log: ['ts', 'role', 'action', 'id', 'ok', 'msg'],
  fact_req_fill: ['req_id', 'role', 'updated_at'],   /* B 階段新增：每張單每個身分最後一次存檔時間（「已填」徽章）；舊試算表由 migrate_ 自動補建 */
  fact_push: ['ts', 'req_id', 'batch', 'action', 'table', 'rows', 'detail', 'backup', 'ok', 'msg'],   /* C 階段新增：寫入／還原採購系統的逐表紀錄 */
  fact_recipe_import: ['ts', 'req_id', 'imp_id', 'action', 'backend_id', 'step', 'ok', 'msg', 'by', 'd1', 'd2', 'd3', 'd4'],   /* 2026-10-06 新增：匯入自己做食譜系統的紀錄（start 列 d1～d4＝這次要匯入的完整內容，中斷可接續） */
  fact_req_comment: ['ts', 'req_id', 'who', 'who_kind', 'text'],   /* 2026-10-07 新增：申請單「💬 意見交流」（簽核人、主廚、各單位留言；大家都看得到；送出後不能刪） */
  fact_notify: ['ts', 'kind', 'key', 'req_id', 'to_name', 'to_email', 'ok', 'msg']   /* 2026-10-08 新增：自動寄 Email 的紀錄（每張單每位收件人一列；查誰收到了沒、為什麼沒寄） */
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
var MIG_KEYS = ['MIG_ADMIN', 'MIG_FILL', 'MIG_PUSH', 'MIG_BOM', 'MIG_EARLY', 'MIG_RBIMP', 'MIG_CMT', 'MIG_BATCH', 'MIG_EMAIL'];
/* 2026-10-07 加速：每次請求的指令碼屬性只讀一次（原本 migrate_ 逐一讀 6 個、ss_ 再讀 1 個；一個約 50～100 毫秒） */
var _props = null;
function props_() { if (!_props) _props = PropertiesService.getScriptProperties().getProperties() || {}; return _props; }
function setProp_(k, v) { PropertiesService.getScriptProperties().setProperty(k, v); props_()[k] = v; }
function migrate_() {
  var A = props_();
  if (!A.SHEET_ID) return;
  if (MIG_KEYS.every(function (k) { return A[k] === '1'; }) && A.SIGN_SECRET && !A.SIGNER_SEED) return;
  var P = PropertiesService.getScriptProperties();
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
    if (P.getProperty('MIG_BOM') !== '1') {   /* ④ 2026-10-01 需求 7：BOM 管理異動紀錄分頁＋預設權限（主廚可改 BOM、營運POS 可改對照表；只加不減） */
      ensureTab_(ss_(), 'fact_bom_log');
      var tr = load_('dim_role'), add = { '主廚': 'B:bom.edit', '營運POS': 'B:map.edit' };
      tr.rows.forEach(function (r) {
        var nm = str_(r.role).trim(), f = str_(r.fields);
        if (add[nm] && f.indexOf(add[nm]) < 0) putCols_(tr, r._row, { fields: f + (f.trim() ? ', ' : '') + add[nm] }, function (h) { return h === 'fields'; });
      });
      P.setProperty('MIG_BOM', '1');
    }
    if (P.getProperty('MIG_EARLY') !== '1') {   /* ⑤ 2026-10-05：fact_campaign 最後補一欄「自己人搶先開賣日」（只加表頭，既有資料一格不動） */
      var cs = ss_().getSheetByName('fact_campaign'), CH = TABS.fact_campaign, cn = CH.length;
      if (cs.getMaxColumns() < cn) cs.insertColumnsAfter(cs.getMaxColumns(), cn - cs.getMaxColumns());
      var hc = cs.getRange(1, cn);
      if (String(hc.getValue()).trim() === '') hc.setValue(CH[cn - 1]).setFontWeight('bold').setBackground('#E0F2F1');
      if (String(hc.getValue()).trim() !== CH[cn - 1]) throw new Error('fact_campaign 第 ' + cn + ' 欄表頭不對');
      cs.getRange(2, cn, Math.max(1, cs.getMaxRows() - 1), 1).setNumberFormat('@');
      P.setProperty('MIG_EARLY', '1');
    }
    if (P.getProperty('MIG_RBIMP') !== '1') {   /* ⑥ 2026-10-06：補建 fact_recipe_import 分頁（只新增分頁，不動既有分頁） */
      ensureTab_(ss_(), 'fact_recipe_import');
      P.setProperty('MIG_RBIMP', '1');
    }
    if (P.getProperty('MIG_CMT') !== '1') {   /* ⑦ 2026-10-07：補建 fact_req_comment 分頁（意見交流；只新增分頁） */
      ensureTab_(ss_(), 'fact_req_comment');
      setProp_('MIG_CMT', '1');
    }
    if (P.getProperty('MIG_BATCH') !== '1') {   /* ⑩ 2026-10-07 晚：dim_signer 最後補一欄「batch」（簽核批次；只加表頭，既有資料一格不動；空白＝第 1 批） */
      var gs = ss_().getSheetByName('dim_signer'), GH = TABS.dim_signer, gn = GH.indexOf('batch') + 1;   /* 2026-10-08：用欄名找位置（之後又加了 email 欄，不能再用最後一欄） */
      if (gs.getMaxColumns() < gn) gs.insertColumnsAfter(gs.getMaxColumns(), gn - gs.getMaxColumns());
      var gc = gs.getRange(1, gn);
      if (String(gc.getValue()).trim() === '') gc.setValue(GH[gn - 1]).setFontWeight('bold').setBackground('#E0F2F1');
      if (String(gc.getValue()).trim() !== GH[gn - 1]) throw new Error('dim_signer 第 ' + gn + ' 欄表頭不對');
      authCacheClear_();
      setProp_('MIG_BATCH', '1');
    }
    if (P.getProperty('MIG_EMAIL') !== '1') {   /* ⑪ 2026-10-08：dim_role 第 4 欄、dim_signer 第 6 欄補表頭「email」（自動寄通知用；只加表頭，既有資料一格不動）＋補建 fact_notify 分頁 */
      ['dim_role', 'dim_signer'].forEach(function (nm) {
        var es = ss_().getSheetByName(nm), EH = TABS[nm], en = EH.indexOf('email') + 1;
        if (es.getMaxColumns() < en) es.insertColumnsAfter(es.getMaxColumns(), en - es.getMaxColumns());
        var ec = es.getRange(1, en);
        if (String(ec.getValue()).trim() === '') ec.setValue(EH[en - 1]).setFontWeight('bold').setBackground('#E0F2F1');
        if (String(ec.getValue()).trim() !== EH[en - 1]) throw new Error(nm + ' 第 ' + en + ' 欄表頭不對');
        es.getRange(2, en, Math.max(1, es.getMaxRows() - 1), 1).setNumberFormat('@');
      });
      ensureTab_(ss_(), 'fact_notify');
      authCacheClear_();
      setProp_('MIG_EMAIL', '1');
    }
    if (!P.getProperty('SIGN_SECRET')) setProp_('SIGN_SECRET', Utilities.getUuid() + Utilities.getUuid());   /* ⑧ 簽核連結通行碼的密鑰（只在伺服器；在鎖裡產生，不會兩個請求各產一組） */
    var seed = P.getProperty('SIGNER_SEED');
    if (seed) {   /* ⑨ 2026-10-07：簽核人名單初始值（管理者／Claude 放在指令碼屬性 SIGNER_SEED＝[{name,role}]；只補名單沒有的人，PIN 留空＝本人第一次簽核時設定；用完就刪，名字不進程式碼） */
      seedSigners_(seed);
      P.deleteProperty('SIGNER_SEED'); delete props_().SIGNER_SEED;
    }
  } catch (e) { /* 下次再試 */ } finally { lock.releaseLock(); }
}
function seedSigners_(json) {
  var list; try { list = JSON.parse(json); } catch (e) { return 0; }
  if (!Array.isArray(list)) return 0;
  var t = load_('dim_signer'), have = {}, n = 0;
  t.rows.forEach(function (r) { have[str_(r.name).trim()] = 1; });
  list.forEach(function (o) {
    var name = str_(o && o.name).trim(), role = str_(o && o.role).trim();
    if (!name || have[name] || /[,，、;；\s]/.test(name)) return;
    have[name] = 1;
    put_(t, nextRow_(t), { name: name, role: role, pin: '', enabled: 'Y' });
    t.rows.push({ _row: t.sh.getLastRow(), name: name });
    n++;
  });
  if (n) authCacheClear_();
  return n;
}

/* ================= 入口 ================= */
function doGet(e) {
  var p = (e && e.parameter) || {}, a = String(p.action || 'ping'), res;
  migrate_();
  try {
    if (a === 'ping') res = { ok: true, data: { version: VERSION, now: now_(), ready: !!props_().SHEET_ID, mig: MIG_KEYS.filter(function (k) { return props_()[k] === '1'; }).length + '/' + MIG_KEYS.length } };   /* mig：遷移完成幾項（上線後從外面確認，不含資料） */
    else if (a === 'listCampaigns') res = { ok: true, data: listCampaignsCached_() };
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
  pushPreview: pushPreview_, pushPurchase: pushPurchase_, rollbackPurchase: rollbackPurchase_,
  /* 2026-10-07：簽核連結（帶通行碼 k）免 PIN 看申請單、本人第一次自設 PIN、意見交流 */
  signOpen: signOpen_, signSetPin: signSetPin_, commentAdd: commentAdd_,
  /* 2026-10-07 晚：整檔一次送簽／一次簽完 */
  submitSignCamp: submitSignCamp_, signOpenCamp: signOpenCamp_, signCamp: signCamp_,
  /* 2026-10-08 v13：重新載入申請表（全部申請單清單） */
  listReqs: listReqs_,
  /* 2026-10-08：🔑 密碼管理填主廚通知 Email、寄測試信 */
  adminSetRoleEmail: adminSetRoleEmail_, adminTestMail: adminTestMail_,
  /* 2026-10-06 🍰 匯入自己做食譜系統：函式在 rb_import.gs，包一層（呼叫時才找函式），不受檔案載入順序影響 */
  rbStatus: function (p, a) { return rbStatus_(p, a); }, rbSetCred: function (p, a) { return rbSetCred_(p, a); },
  rbPreview: function (p, a) { return rbPreview_(p, a); }, rbImport: function (p, a) { return rbImport_(p, a); } };
var WRITES = { saveCampaign: 1, saveReq: 1, saveUnit: 1, adminSaveSigner: 1, adminDeleteSigner: 1, adminSetPassword: 1,
  patchReq: 1, submitSign: 1, withdrawSign: 1, sign: 1, pushPurchase: 1, rollbackPurchase: 1, rbSetCred: 1, signSetPin: 1, commentAdd: 1, submitSignCamp: 1, signCamp: 1,
  adminSetRoleEmail: 1, adminTestMail: 1 };   /* 回覆裡都不含密碼／PIN，才能放進回條快取（adminTestMail 放這裡＝重送不會多寄一封） */
/* 2026-10-07：用簽核連結的通行碼 k 驗身分的動作（不用密碼／PIN）：signOpen 只讀、signSetPin 只能替「還沒設 PIN」的簽核人設一次 */
var KEY_ACTS = { signOpen: 1, signSetPin: 1, signOpenCamp: 1 };   /* signOpenCamp＝整檔連結（k＝檔期的通行碼） */
/* 寫完要讓檔期清單暫存失效的動作（檔期內容或各檔期支數會變） */
var CAMP_TOUCH = { saveCampaign: 1, saveReq: 1, submitSign: 1, submitSignCamp: 1 };
/* 寫完要讓身分表暫存失效的動作 */
var AUTH_TOUCH = { adminSaveSigner: 1, adminDeleteSigner: 1, adminSetPassword: 1, signSetPin: 1, adminSetRoleEmail: 1 };
var RQ_SEC = 21600;   /* 回條保留 6 小時：同一個回條編號重送 → 直接回上次結果，不重複寫 */
/* v7（2026-10-04 效能第 2 批）：只讀動作**明確列在 READS**（都已逐一確認沒有任何寫入）→ 不排 LockService、成功不寫 log 分頁（失敗照寫，留除錯線索）。
   沒列在 READS 的動作一律走下面原本的加鎖路徑。與「不在 WRITES 就當只讀」的寫法行為等價（WRITES 在 BOM 段落尾端補了 bomSave 等 6 個，
   30 個動作＝12 讀＋18 寫），改白名單是為了日後新增動作忘了登記 WRITES 時預設仍加鎖（審查時改的）。
   adminList 例外：回覆含各身分密碼，成功也要留一筆稽核 */
var READS = { whoami: 1, getCampaign: 1, getReq: 1, listSigners: 1, signView: 1, pushPreview: 1, bomMeta: 1, bomGet: 1, mapGet: 1, bomLog: 1, bomMetaMap: 1, adminList: 1,
  rbStatus: 1, rbPreview: 1, signOpen: 1, signOpenCamp: 1, listReqs: 1 };   /* rbPreview 只讀（會登入後台、讀頁面，不寫試算表也不寫後台） */
var AUDIT_READS = { adminList: 1 };
/* 2026-10-06：自己管鎖的動作——匯入自己做食譜系統要連後台、一次 30 多秒，不能佔住全系統的鎖（別人送件會卡住）；
   同一張單不重複跑由 rbImport_ 自己用 CacheService 擋，寫紀錄分頁時才短暫拿鎖 */
var SELFLOCK = { rbImport: 1 };
/* 會動到 BOM 本「BOM表／產品名稱對照表」的動作：做完（成功或失敗都算，失敗可能已寫一半又寫回）就讓 bomMeta／mapGet 的暫存失效 */
var BOM_TOUCH = { bomSave: 1, bomDelete: 1, bomRename: 1, mapSave: 1, mapDelete: 1, bomUndo: 1, sign: 1, signCamp: 1, pushPurchase: 1, rollbackPurchase: 1 };

function doPost(e) {
  var p;
  try { p = JSON.parse((e && e.postData && e.postData.contents) || '{}') || {}; } catch (err) { return out_({ ok: false, msg: '送來的資料不是 JSON' }); }
  var act = String(p.action || ''), who = p.signer ? ('簽核:' + String(p.signer)) : String(p.role || ''), res;
  _nq = []; _nctx = null;   /* 2026-10-08：這次請求要寄的通知（寫完、放開鎖之後才寄） */
  migrate_();
  var fn = ACTIONS[act];
  if (!fn) { res = { ok: false, act: act, msg: '不支援的動作：' + act }; log_(who, act, '', false, res.msg); return out_(res); }
  /* 先驗身分（在鎖外面）：密碼錯 → 停 1 秒再回（拖慢亂猜，不佔鎖、不影響別人）；不做「錯太多次就鎖整個身分」——
     那會讓知道網址的人故意打錯把全部主廚鎖在外面（2026-09-28 測試時實際發生過） */
  var auth;
  try { auth = KEY_ACTS[act] ? keyAuth_(p) : auth_(p); } catch (err) {
    Utilities.sleep(1000);
    res = { ok: false, act: act, code: (err && err.code) || 'auth', msg: errMsg_(err) };
    log_(who, act, guessId_(p), false, res.msg);
    return out_(res);
  }
  if (READS[act]) {   /* v7：只讀動作（READS 白名單）→ 不排鎖、不用回條、成功不寫 log（回覆格式與寫入動作相同）；其餘動作走下面的鎖 */
    try {
      var rr = fn(p, auth) || {};
      res = { ok: true, act: act, data: rr.data };
      if (AUDIT_READS[act]) log_(who, act, rr.id || '', true, rr.msg || '');
    } catch (err2) {
      res = { ok: false, act: act, msg: errMsg_(err2) };
      if (err2 && err2.code) res.code = err2.code;
      if (err2 && err2.extra) res.data = err2.extra;
      log_(who, act, guessId_(p), false, res.msg);
    }
    return out_(res);
  }
  if (SELFLOCK[act]) {
    try {
      var rs = fn(p, auth) || {};
      res = { ok: true, act: act, data: rs.data };
      log_(who, act, rs.id || '', true, rs.msg || '');
    } catch (err3) {
      res = { ok: false, act: act, msg: errMsg_(err3) };
      if (err3 && err3.code) res.code = err3.code;
      if (err3 && err3.extra) res.data = err3.extra;
      log_(who, act, guessId_(p), false, res.msg);
    }
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
  } finally {
    lock.releaseLock();
    if (BOM_TOUCH[act]) bomCacheClear_();
    if (CAMP_TOUCH[act]) campCacheClear_();
    if (AUTH_TOUCH[act]) authCacheClear_();
  }
  /* 2026-10-08：送簽／簽核成功 → 放開鎖之後才寄 Email（寄信 1 封約 0.5～1 秒，不佔住全系統的鎖）；寄給誰、誰沒填 Email 附在回覆裡給畫面顯示 */
  if (_nq.length) {
    if (res && res.ok) {
      var nr = flushNotify_();
      if (nr) {
        res.data = res.data || {}; res.data.notify = nr;
        if (rqKey) { try { CacheService.getScriptCache().put(rqKey, JSON.stringify(res), RQ_SEC); } catch (ce2) { } }
      }
    }
    _nq = [];
  }
  return out_(res);
}

/* ===== 2026-10-07 加速：暫存（CacheService）=====
   ①身分表 dim_role／dim_signer：每個請求驗身分都要讀一次（約 0.2～0.5 秒）→ 暫存 90 秒；🔑 密碼管理改密碼、改簽核人、本人設 PIN 後立刻失效。
     直接在試算表手改 dim_role／dim_signer（例：可填欄位、密碼）最多 90 秒後生效。
   ②檔期清單 listCampaigns：暫存 2 分鐘；新增／改檔期、送件（支數會變）後立刻失效。 */
var AUTH_TTL = 90, CAMP_TTL = 120;
function authRows_(name) {
  var c = CacheService.getScriptCache(), k = 'auth:' + name + ':v1', hit = c.get(k);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var rows = load_(name).rows.map(function (r) { var o = {}; TABS[name].forEach(function (h) { o[h] = str_(r[h]); }); return o; });
  try { c.put(k, JSON.stringify(rows), AUTH_TTL); } catch (e) { }
  return rows;
}
function authCacheClear_() { try { CacheService.getScriptCache().removeAll(['auth:dim_role:v1', 'auth:dim_signer:v1']); } catch (e) { } }
function listCampaignsCached_() {
  var c = CacheService.getScriptCache(), k = 'camps:v1', hit = c.get(k);
  if (hit) { try { return JSON.parse(hit); } catch (e) { } }
  var list = listCampaigns_();
  try { c.put(k, JSON.stringify(list), CAMP_TTL); } catch (e) { }
  return list;
}
function campCacheClear_() { try { CacheService.getScriptCache().remove('camps:v1'); } catch (e) { } }

/* ===== 2026-10-07 簽核連結通行碼 k：HMAC(申請單號, 只在伺服器的密鑰) 前 16 碼。
   拿到連結（104 指定員工公告）的人點開就能看申請單與意見交流；留言、同意／退回仍要選自己名字＋PIN */
function signKey_(rid) {
  var sec = props_().SIGN_SECRET || PropertiesService.getScriptProperties().getProperty('SIGN_SECRET');   /* 剛上線那一刻：別的請求剛在鎖裡產生密鑰，這個請求的設定暫存還沒有 → 重讀一次 */
  if (!sec) throw fail_('簽核連結尚未啟用（系統設定中），請 1 分鐘後再試', 'key');
  var sig = Utilities.computeHmacSha256Signature(String(rid), sec);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '').slice(0, 16);
}
function keyAuth_(p) {
  var cid = str_(p.campaign_id).trim();
  if (cid && !str_(p.req_id).trim()) {   /* 2026-10-07 晚：整檔簽核連結（#sign=檔期編號&k=）——k 用「camp:」前綴，和單張的 k 分開 */
    var kc = str_(p.k).trim();
    if (!kc || kc !== signKey_('camp:' + cid)) throw fail_('簽核連結不完整或打錯了，請跟主廚要 Email 或 104 公告裡的連結', 'key');
    return { kind: 'key', name: '', role: '', rules: [], camp: cid };
  }
  var rid = str_(p.req_id).trim(), k = str_(p.k).trim();
  if (!rid || !k || k !== signKey_(rid)) throw fail_('簽核連結不完整或打錯了，請跟主廚要 Email 或 104 公告裡的連結', 'key');
  return { kind: 'key', name: '', role: '', rules: [] };
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
    var s = authRows_('dim_signer').filter(function (r) { return str_(r.name).trim() === name && signerOn_(r.enabled); })[0];
    if (s && !str_(s.pin).trim()) throw fail_('「' + name + '」還沒有設定 PIN：請點 Email 或 104 公告裡的簽核連結，選自己的名字設定 PIN', 'nopin');
    if (!s || !pin || str_(s.pin).trim() !== pin) { cache.put(ck, String(fails + 1), PIN_LOCK_SEC); throw fail_('簽核人或 PIN 不對', 'auth'); }
    if (fails) cache.remove(ck);
    return { kind: 'signer', name: name, role: str_(s.role).trim(), rules: [] };
  }
  var role = String(p.role || '').trim(), pw = String(p.password || '').trim();
  if (!role) throw fail_('請選擇身分並輸入密碼', 'auth');
  var r = authRows_('dim_role').filter(function (x) { return str_(x.role).trim() === role; })[0];
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
    .map(function (x) { return { role: str_(x.role).trim(), password: str_(x.password), fields: str_(x.fields), email: str_(x.email).trim() }; });
  var signers = load_('dim_signer').rows.filter(function (x) { return str_(x.name).trim(); })
    .map(function (x) { return { name: str_(x.name).trim(), role: str_(x.role), pin: str_(x.pin), enabled: signerOn_(x.enabled) ? 'Y' : 'N', batch: str_(x.batch).trim(), email: str_(x.email).trim() }; });
  return { id: '', data: { roles: roles, signers: signers }, msg: '身分 ' + roles.length + '、簽核主管 ' + signers.length };
}
function adminSaveSigner_(p, auth) {   /* 資料放 p.entry（p.signer 是簽核人 PIN 登入用，不能混用） */
  needAdmin_(auth);
  var g = p.entry || {}, orig = str_(g.orig_name).trim(), name = str_(g.name).trim(), title = str_(g.role).trim(), pin = str_(g.pin).trim();
  var en = (g.enabled === false || str_(g.enabled).toUpperCase() === 'N') ? 'N' : 'Y', bt = str_(g.batch).trim();
  if (!name) throw fail_('請填簽核主管姓名');
  if (bt && !/^[1-9]$/.test(bt)) throw fail_('批次請填 1～9（空白＝第 1 批）');
  var keepBatch = g.batch === undefined || g.batch === null;   /* 舊版網頁（沒有批次欄）存檔：批次照舊，不清掉 */
  var keepEmail = g.email === undefined || g.email === null, em = keepEmail ? '' : emailList_(g.email, true).join(', ');   /* 2026-10-08：同上，舊版網頁沒有 Email 欄 → 照舊 */
  if (/[,，、;；\s]/.test(name)) throw fail_('姓名不能有逗號、頓號、分號或空白');   /* signers 欄用這些符號分隔 */
  if (pin && !/^\d{4,8}$/.test(pin)) throw fail_('PIN 要 4～8 位數字（留空＝本人第一次簽核時自己設定）');
  var t = load_('dim_signer');
  var row = orig ? t.rows.filter(function (r) { return str_(r.name).trim() === orig; })[0] : null;
  if (orig && !row) throw fail_('找不到簽核主管「' + orig + '」', 'notfound');
  if (t.rows.some(function (r) { return r !== row && str_(r.name).trim() === name; })) throw fail_('已經有叫「' + name + '」的簽核主管', 'dup');
  if (keepBatch) bt = row ? str_(row.batch).trim() : '';
  if (keepEmail) em = row ? str_(row.email).trim() : '';
  put_(t, row ? row._row : nextRow_(t), { name: name, role: title, pin: pin, enabled: en, batch: bt, email: em });
  return {
    id: name, data: { name: name, role: title, enabled: en, batch: bt, email: em, created: !row },
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
/* 2026-10-08：各身分的通知 Email（目前只有「主廚」會收到：有一支被退回、或核准）；只寫 email 這一格 */
function adminSetRoleEmail_(p, auth) {
  needAdmin_(auth);
  var role = str_(p.target_role).trim(), em = emailList_(p.email, true).join(', ');
  var t = load_('dim_role'), row = t.rows.filter(function (r) { return str_(r.role).trim() === role; })[0];
  if (!row) throw fail_('找不到身分「' + role + '」', 'notfound');
  putCols_(t, row._row, { email: em }, function (h) { return h === 'email'; });
  return { id: role, data: { role: role, email: em }, msg: '通知 Email ' + role + (em ? '' : '（清空）') };
}
/* 2026-10-08：寄一封測試信（確認 Email 打對、信不會被擋）；在 WRITES 裡＝回覆掉了重送不會多寄 */
function adminTestMail_(p, auth) {
  needAdmin_(auth);
  var to = emailList_(p.email, true), nm = str_(p.name).trim().slice(0, 40);
  if (!to.length) throw fail_('請先填 Email');
  if (MailApp.getRemainingDailyQuota() < to.length) throw fail_('今天 Gmail 寄信額度用完了（一天 100 封），明天再試', 'quota');
  var txt = (nm ? nm + ' 你好：\n\n' : '') + '這是「自己做 食譜系統」的測試信。\n之後新品申請輪到你簽核時，系統會寄通知到這個信箱，信裡附簽核連結，點開就能簽。\n\n收到這封就代表設定成功，不用回信。\n自己做 食譜系統';
  MailApp.sendEmail({ to: to.join(','), subject: '【食譜系統】測試信：之後輪到你簽核時會寄到這個信箱', body: txt, htmlBody: mailHtml_(txt, '', ''), name: NOTIFY_FROM });
  return { id: nm || '測試信', data: { to: to.length }, msg: '測試信 → ' + (nm || '') + '（' + to.length + ' 個信箱）' };
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
  ['起', '迄', '第一批配貨日', '自己人搶先開賣日'].forEach(function (h) { obj[h] = normDate_(obj[h]); });
  if (obj['自己人搶先開賣日'] && !/^\d{4}-\d{2}-\d{2}$/.test(obj['自己人搶先開賣日'])) throw fail_('「自己人搶先開賣日」日期格式不對');
  if (obj['自己人搶先開賣日'] && /^\d{4}-\d{2}-\d{2}$/.test(obj['起']) && obj['自己人搶先開賣日'] > obj['起']) throw fail_('「自己人搶先開賣日」不能晚於檔期開始日');
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
      fills: safeRows_('fact_req_fill', ids), me: meOf_(auth),
      camp_key: (auth.kind === 'role' && (auth.role === ADMIN || can_(auth, 'P:*'))) ? signKey_('camp:' + cid) : '' },   /* 2026-10-07 晚：整檔簽核連結（只給主廚／管理者） */
    msg: reqs.length + ' 支' + (lite ? '（清單）' : '')
  };
}

/* 2026-10-08 v13：「重新載入申請表」——曾經送件過的全部申請單（不含內容；最新的在前），主廚／管理者 */
function listReqs_(p, auth) {
  if (!canPush_(auth)) throw fail_('只有主廚或管理者可以看全部申請單', 'perm');
  var ct = load_('fact_campaign'), cm = {};
  ct.rows.forEach(function (r) { var k = str_(r.campaign_id); if (k) cm[k] = { name: str_(r['檔期名稱']), from: str_(r['起']), state: str_(r.status) }; });
  var t = load_('fact_recipe_req', true);
  var list = t.rows.filter(function (r) { return str_(r.req_id); }).map(function (r) {
    var c = cm[str_(r.campaign_id)] || {};
    return { req_id: str_(r.req_id), campaign_id: str_(r.campaign_id), campaign: c.name || '', camp_from: c.from || '', camp_state: c.state || '', seq: str_(r.seq),
      name: str_(r['商品正式名稱']).trim() || str_(r['商品暫定名稱']).trim(), price: r['定價'] === undefined ? '' : r['定價'], status: str_(r.status),
      updated_at: str_(r.updated_at), updated_by: str_(r.updated_by) };
  }).sort(function (a, b) { return a.updated_at < b.updated_at ? 1 : a.updated_at > b.updated_at ? -1 : 0; });
  return { id: '', data: { reqs: list }, msg: list.length + ' 張' };
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
      fills: safeRows_('fact_req_fill', ids), me: meOf_(auth), push: pushInfo_(rid), rbimp: (function () { try { return rbInfo_(rid); } catch (e) { return null; } })(),
      comments: commentsOf_(rid), sign_key: (auth.kind === 'role' && (auth.role === ADMIN || can_(auth, 'P:*'))) ? signKey_(rid) : '' },   /* 2026-10-07：意見交流＋簽核連結通行碼（只給主廚／管理者，貼進 104 公告用） */
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
  return { id: '', data: peopleOf_(), msg: '' };
}
/* 啟用中的簽核人（只回姓名、職稱／部門、有沒有設 PIN，不含 PIN） */
function peopleOf_() {
  return authRows_('dim_signer').filter(function (r) { return str_(r.name).trim() && signerOn_(r.enabled); })
    .map(function (r) { return { name: str_(r.name).trim(), role: str_(r.role).trim(), hasPin: !!str_(r.pin).trim(), batch: batchOf_(r.batch), hasEmail: emailList_(r.email).length > 0 }; });
}
/* 2026-10-07 晚：簽核批次（1～9；空白或亂填＝1）。前一批（數字小的）全部同意，下一批才能簽 */
function batchOf_(v) { var n = parseInt(str_(v).trim(), 10); return n >= 1 && n <= 9 ? n : 1; }
function batchMap_() { var m = {}; authRows_('dim_signer').forEach(function (r) { var n = str_(r.name).trim(); if (n) m[n] = batchOf_(r.batch); }); return m; }

/* ===== 2026-10-07 💬 意見交流（fact_req_comment）＋簽核連結 =====
   ‧ 誰都看得到所有人的留言與簽核意見（申請表「送簽核」卡、簽核頁都會顯示）；留言送出後不能改、不能刪
   ‧ 留言身分：簽核人（姓名＋PIN）或各身分（主廚、出貨中心…＋密碼） */
var CMT_MAX = 1000;
function commentsOf_(rid) {
  var ids = {}; ids[rid] = 1;
  return safeRows_('fact_req_comment', ids).sort(function (a, b) { return str_(a.ts) < str_(b.ts) ? -1 : str_(a.ts) > str_(b.ts) ? 1 : 0; });
}
function commentAdd_(p, auth) {
  var rid = str_(p.req_id).trim(), text = str_(p.text).replace(/\r\n?/g, '\n').trim();
  if (!rid) throw fail_('缺申請單號');
  if (!text) throw fail_('請寫留言內容');
  if (text.length > CMT_MAX) throw fail_('留言最多 ' + CMT_MAX + ' 字（現在 ' + text.length + ' 字），請分兩則');
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到申請單 ' + rid, 'notfound');
  var who, kind;
  if (auth.kind === 'signer') { who = auth.name; kind = '簽核人' + (auth.role ? '｜' + auth.role : ''); }
  else if (auth.kind === 'role') { who = auth.role; kind = '身分'; }
  else throw fail_('留言要先選自己的名字並輸入 PIN', 'perm');
  var ct = load_('fact_req_comment'), rec = { ts: now_(), req_id: rid, who: who, who_kind: kind, text: text.slice(0, CMT_MAX) };
  put_(ct, nextRow_(ct), rec);
  return { id: rid, data: { comment: rec, comments: commentsOf_(rid) }, msg: who + ' 留言 ' + text.length + ' 字' };
}
/* 簽核連結（k）打開：整張申請單＋同檔期清單＋品項補充＋簽核紀錄＋意見交流＋可選的名字（不用 PIN） */
function signOpen_(p, auth) {
  var rid = str_(p.req_id).trim();
  var t = load_('fact_recipe_req', true), row = reqRow_(t, rid);
  if (!row) throw fail_('找不到這張申請單（可能已刪除）', 'notfound');
  var d = signData_(t, row, rid);
  d.people = peopleOf_();
  return { id: rid, data: d, msg: '簽核連結 ' + rid };
}
function signData_(t, row, rid) {
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
  return { req: req, campaign: camp ? plain_(ct.H, camp) : null, others: others, units: rowsOf_('fact_req_unit', ids),
    signoffs: rowsOf_('fact_signoff', ids), comments: commentsOf_(rid) };
}
/* 本人第一次簽核時自設 PIN（簽核連結 k）：只限「這張單被指定的簽核人」、名單上啟用中、PIN 還是空的人（別人拿到連結也不能替同事設）；
   設好之後要改只能找管理者（清空 PIN 讓本人重設） */
function signSetPin_(p, auth) {
  var rid = str_(p.req_id).trim(), name = str_(p.signer).trim(), pin = str_(p.new_pin).trim();
  if (!name) throw fail_('請先選你的名字');
  if (!/^\d{4,8}$/.test(pin)) throw fail_('PIN 要 4～8 位數字');
  var qt = load_('fact_recipe_req', true), qrow = reqRow_(qt, rid);
  if (auth && auth.camp) {   /* 2026-10-07 晚：整檔連結——檔期內任一張的簽核人都可以自設 */
    var hit = qt.rows.filter(function (r) { return str_(r.campaign_id) === auth.camp && signersOf_(r).indexOf(name) >= 0; })[0];
    qrow = hit || null; rid = hit ? str_(hit.req_id) : auth.camp;
  }
  if (!qrow || signersOf_(qrow).indexOf(name) < 0) throw fail_('「' + name + '」不是這張單的簽核人，不能在這裡設定 PIN；請主廚把你加進簽核人，或請管理者在「🔑 密碼管理」幫你設定', 'perm');
  var t = load_('dim_signer'), row = t.rows.filter(function (r) { return str_(r.name).trim() === name && signerOn_(r.enabled); })[0];
  if (!row) throw fail_('名單上沒有「' + name + '」（或已停用），請找管理者', 'notfound');
  if (str_(row.pin).trim()) throw fail_('「' + name + '」已經設定過 PIN；忘記請找管理者在「🔑 密碼管理」重設', 'haspin');
  putCols_(t, row._row, { pin: pin }, function (h) { return h === 'pin'; });
  return { id: name, data: { name: name, role: str_(row.role).trim() }, msg: name + ' 第一次簽核自設 PIN（' + rid + '）' };
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
  /* 2026-10-07 加速：帶了 req＋payload ＝「先存最新內容再送簽」在同一次請求做完（原本前端要先送件、再送簽，等兩趟 Google）。
     存檔規則與「📨 送件到檔期」完全相同（別人在你之後改過 → conflict，不覆蓋） */
  var saved = null;
  if (p.req && typeof p.payload === 'string' && p.payload) {
    var sq = p.req || {};
    if (str_(sq.req_id).trim() !== rid) throw fail_('送簽的單號和內容對不上，請重新整理', 'bad');
    saved = saveReq_({ req: sq, payload: p.payload, base_updated_at: p.base_updated_at, force: false }, auth).data;
    p.base_updated_at = saved.updated_at;
  }
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
  nqTurn_(row, names, {}, now, '', 'submit');   /* 2026-10-08：寄 Email 給第 1 批 */
  return { id: rid, data: { status: '送簽中', signers: names, updated_at: now, updated_by: auth.role, voided: voided, saved: saved, sign_key: signKey_(rid) },
    msg: (saved ? '存檔＋' : '') + '送簽給 ' + names.join('、') + (voided ? '｜上一輪簽核 ' + voided + ' 筆標記作廢' : '') };
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
  var d = signData_(t, row, rid);
  d.me = meOf_(auth);
  return { id: rid, data: d, msg: '簽核頁 ' + rid };
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
  /* 2026-10-07 晚：分批簽核——自己的批次比「還沒同意的人裡最小的批次」大 → 還沒輪到 */
  var bmap = batchMap_(), myB = bmap[auth.name] || 1, agreed0 = {};
  so.rows.forEach(function (r) { if (str_(r.req_id) === rid && str_(r.decision).trim() === '同意') agreed0[str_(r.signer).trim()] = 1; });
  var pend0 = names.filter(function (n) { return n !== auth.name && !agreed0[n]; });
  var turn0 = Math.min.apply(null, pend0.map(function (n) { return bmap[n] || 1; }).concat([myB]));
  if (myB > turn0) {
    var wait0 = pend0.filter(function (n) { return (bmap[n] || 1) < myB; });
    throw fail_('還沒輪到你：第 ' + turn0 + ' 批（' + wait0.join('、') + '）全部同意後才能簽', 'notyet', { turn: turn0, wait: wait0 });
  }
  var mine = so.rows.filter(function (r) { return str_(r.req_id) === rid && str_(r.signer).trim() === auth.name && isLive_(r.decision); })[0];
  var turnBefore = turnOf_(names, agreed0, bmap);   /* 2026-10-08：簽之前輪到第幾批（自動寄 Email 用；agreed0 含自己先前的同意） */
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
  /* 2026-10-08 自動寄 Email：前一批全部同意 → 下一批；退回、核准 → 主廚 */
  var agSet = {}; agreed.forEach(function (n) { agSet[n] = 1; });
  if (dec === '同意' && status === st && turnOf_(names, agSet, bmap) > turnBefore) nqTurn_(row, names, agSet, str_(row.updated_at), (_nctx && _nctx.camp) || '', 'next');
  if (status !== st) _nq.push({ kind: dec === '退回' ? 'reject' : 'done', rid: rid, cid: str_(row.campaign_id), dessert: dessertOf_(row), round: str_(row.updated_at),
    by: auth.name, comment: dec === '退回' ? cm : '', status: status, push: push && push.ok === false ? push.msg : '' });
  return { id: rid, data: { status: status, decision: dec, signers: names, agreed: agreed, signoffs: live.map(function (r) { return plain_(so.H, r); }), push: push },
    msg: auth.name + ' ' + dec + '（' + agreed.length + '/' + names.length + ' 同意）' + (status !== st ? '→' + status : '') + (push ? '｜' + (push.ok === false ? push.msg : pushMsg_(push)) : '') };
}

/* ===== 2026-10-07 晚 經營者：「一個檔期 4～6 支新甜點，整檔一次簽核，不要一支一支簽」=====
   ‧ submitSignCamp（主廚）：勾這個檔期要送的幾支（草稿／退回、而且是系統上最新版），同一組簽核人一起送簽；全部檢查通過才寫（一支不合格整批不送）
   ‧ signOpenCamp（整檔連結 k）：整個檔期每一支的申請單、品項補充、簽核紀錄、意見交流＋可選的名字（不用 PIN 就能看）
   ‧ signCamp（簽核人 PIN）：一次送多支的決定（全部同意，或某幾支退回）；每一支照原本 sign_ 的規則（分批、退回要原因、全部同意就核准並寫入採購系統）
     某一支不能簽（例：還沒輪到、已經不在送簽中）不影響其他支，結果逐支回報 */
function submitSignCamp_(p, auth) {
  if (!can_(auth, 'P:*')) throw fail_('只有主廚可以送簽', 'perm');
  var cid = str_(p.campaign_id).trim();
  if (!cid) throw fail_('請先選檔期');
  var items = (p.items || []).map(function (x) { return { rid: str_(x && x.req_id).trim(), base: str_(x && x.base_updated_at) }; }).filter(function (x) { return x.rid; });
  if (!items.length) throw fail_('請至少勾一支要送簽的甜點');
  var seen = {}, names = [];
  (p.signers || []).forEach(function (x) { var n = str_(x).trim(); if (n && !seen[n]) { seen[n] = 1; names.push(n); } });
  if (!names.length) throw fail_('請至少勾一位簽核主管');
  var on = {};
  load_('dim_signer').rows.forEach(function (r) { if (signerOn_(r.enabled)) on[str_(r.name).trim()] = 1; });
  var bad = names.filter(function (n) { return !on[n]; });
  if (bad.length) throw fail_('這些人不在啟用中的簽核主管名單：' + bad.join('、'), 'badsigner');
  /* 主廚畫面上正在改的那一支（save）：先存最新內容（規則同「📨 送件到檔期」，別人在你之後改過 → conflict 不覆蓋），再用存好的版本一起送簽 */
  var saved = null;
  if (p.save && p.save.req && typeof p.save.payload === 'string' && p.save.payload) {
    var srid = str_(p.save.req.req_id).trim(), sit = items.filter(function (x) { return x.rid === srid; })[0];
    if (!sit) throw fail_('要先存的那一支不在這次送簽的清單裡，請重新整理', 'bad');
    if (str_(p.save.req.campaign_id).trim() !== cid) throw fail_('畫面上這一支選的檔期和這次整檔送簽的檔期不同，請先按「📨 送件到檔期」', 'bad');
    saved = saveReq_({ req: p.save.req, payload: p.save.payload, base_updated_at: p.save.base_updated_at, force: false }, auth).data;
    sit.base = saved.updated_at;
  }
  var t = load_('fact_recipe_req', true), probs = [], rows = [], dup = {};
  items.forEach(function (x) {
    if (dup[x.rid]) return; dup[x.rid] = 1;
    var row = reqRow_(t, x.rid), nm = row ? (str_(row['商品正式名稱']) || str_(row['商品暫定名稱']) || x.rid) : x.rid;
    if (!row) { probs.push(x.rid + ' 找不到'); return; }
    if (str_(row.campaign_id) !== cid) { probs.push(nm + ' 不是這個檔期的單'); return; }
    var st = str_(row.status);
    if (st !== '草稿' && st !== '退回') { probs.push(nm + ' 目前是「' + st + '」，不能送簽'); return; }
    if (!p.force && str_(row.updated_at) !== x.base) { probs.push(nm + ' 在 ' + str_(row.updated_at) + ' 被「' + str_(row.updated_by) + '」改過，請先按「🔄 重新整理」'); return; }
    rows.push(row);
  });
  if (probs.length) throw fail_((saved ? '畫面上這一支已存檔，但' : '') + '整檔送簽沒有送出（一支都沒送）：' + probs.join('；'), 'conflict', { probs: probs, saved: saved });
  var now = now_(), voided = 0;
  rows.forEach(function (row) {
    voided += voidSignoffs_(str_(row.req_id));
    putCols_(t, row._row, { status: '送簽中', signers: names.join('、'), updated_at: now, updated_by: auth.role },
      function (h) { return h === 'status' || h === 'signers' || h === 'updated_at' || h === 'updated_by'; });
    nqTurn_(row, names, {}, now, cid, 'submit');   /* 2026-10-08：寄 Email 給第 1 批（同一個人只收一封，列出這次送的每一支） */
  });
  return { id: cid, data: { campaign_id: cid, req_ids: rows.map(function (r) { return str_(r.req_id); }), status: '送簽中', signers: names, updated_at: now, voided: voided,
      camp_key: signKey_('camp:' + cid), saved: saved },
    msg: '整檔送簽 ' + rows.length + ' 支給 ' + names.join('、') + (voided ? '｜上一輪簽核 ' + voided + ' 筆標記作廢' : '') };
}
function signOpenCamp_(p, auth) {
  var cid = (auth && auth.camp) || str_(p.campaign_id).trim();
  var ct = load_('fact_campaign'), camp = ct.rows.filter(function (r) { return str_(r.campaign_id) === cid; })[0];
  if (!camp) throw fail_('找不到這個檔期（可能已刪除）', 'notfound');
  var t = load_('fact_recipe_req', true);
  var rows = t.rows.filter(function (r) { return str_(r.campaign_id) === cid; })
    .sort(function (a, b) { return (parseInt(str_(a.seq), 10) || 0) - (parseInt(str_(b.seq), 10) || 0); });
  var ids = {}; rows.forEach(function (r) { ids[str_(r.req_id)] = 1; });
  var units = rows.length ? safeRows_('fact_req_unit', ids) : [], sos = rows.length ? safeRows_('fact_signoff', ids) : [], cms = rows.length ? safeRows_('fact_req_comment', ids) : [];
  var byTs = function (a, b) { return str_(a.ts) < str_(b.ts) ? -1 : str_(a.ts) > str_(b.ts) ? 1 : 0; };
  var reqs = rows.map(function (r) {
    var rid = str_(r.req_id), q = plain_(t.H, r); q.payload = readPayload_(t, r._row);
    return { req: q, units: units.filter(function (u) { return str_(u.req_id) === rid; }), signoffs: sos.filter(function (o) { return str_(o.req_id) === rid; }),
      comments: cms.filter(function (c) { return str_(c.req_id) === rid; }).sort(byTs) };
  });
  return { id: cid, data: { campaign: plain_(ct.H, camp), reqs: reqs, people: peopleOf_() }, msg: '整檔簽核連結 ' + cid + '（' + reqs.length + ' 支）' };
}
function signCamp_(p, auth) {
  if (auth.kind !== 'signer') throw fail_('簽核要用簽核人 PIN', 'perm');
  var cid = str_(p.campaign_id).trim(), list = (p.decisions || []).slice(0, 30), out = [];
  if (!list.length) throw fail_('沒有要簽的甜點');
  var t = load_('fact_recipe_req', true);
  _nctx = { camp: cid };   /* 2026-10-08：整檔簽核 → 通知下一批時給整檔連結、同一個人只寄一封 */
  list.forEach(function (d) {
    var rid = str_(d && d.req_id).trim(), row = reqRow_(t, rid);
    if (!row || (cid && str_(row.campaign_id) !== cid)) { out.push({ req_id: rid, ok: false, code: 'notfound', msg: '不是這個檔期的單' }); return; }
    try { var r = sign_({ req_id: rid, decision: d.decision, comment: d.comment }, auth) || {}; out.push({ req_id: rid, ok: true, data: r.data, msg: r.msg }); }
    catch (e) { out.push({ req_id: rid, ok: false, code: (e && e.code) || '', msg: errMsg_(e), data: e && e.extra }); }
  });
  _nctx = null;
  var okN = out.filter(function (x) { return x.ok; }).length;
  return { id: cid, data: { results: out }, msg: auth.name + ' 整檔簽核 ' + okN + '／' + out.length + ' 支' + (okN < out.length ? '｜沒成功：' + out.filter(function (x) { return !x.ok; }).map(function (x) { return x.req_id + ' ' + x.msg; }).join('；') : '') };
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

/* ================= 自動寄 Email（2026-10-08 經營者：「前一個人簽完、輪到我簽，我要怎麼知道？每天去看太煩」）=================
   ‧ 簽核人：dim_signer.email（🔑 密碼管理填）。主廚送簽（單張／整檔）→ 寄給第 1 批；前一批全部同意 → 寄給下一批（只寄還沒同意的人）。
     信裡附簽核連結（單張 #sign=R-…&k=、整檔 #sign=C-…&k=，再帶 &who=本人），點開就能看、簽
   ‧ 主廚：dim_role「主廚」那一列的 email（可填多個，逗號分隔）。有一支被退回（附誰、原因）、或核准（附有沒有寫進採購系統）→ 寄給主廚
   ‧ 送簽／簽核的寫入做完、放開鎖之後才寄（doPost 尾端 flushNotify_）：寄信 1 封約 0.5～1 秒，不佔住全系統的鎖；寄信失敗不影響簽核本身
   ‧ 同一個人同一次只收一封（整檔簽核時列出這次輪到他的每一支）；每張單每位收件人記一列在 fact_notify（查誰收到了沒）
   ‧ 不會重寄：同一次按鈕的重送沿用上次結果（回條 rq，不再執行）；再按一次同意、輪次沒變就不寄（sign_ 比較簽之前／之後輪到第幾批）
   ‧ 沒填 Email 的人不寄，回覆裡列出來（畫面提醒主廚改用 104 通知）；Gmail 一般帳號一天最多寄 100 封
   ‧ 需要 Gmail 寄信權限（appsscript.json script.send_mail）：第一次要經營者授權（編輯器執行 notifyAuthorize） */
var NOTIFY_URL = 'https://diybc-training.onrender.com/static/dashboard-recipe.html';
var NOTIFY_FROM = '自己做 食譜系統';
var CHEF_ROLE = '主廚';
var _nq = [], _nctx = null;   /* 這次請求要寄的通知；_nctx.camp＝整檔簽核中（給整檔連結） */
var EMAIL_RE = /^[^@\s,，;；<>"']+@[^@\s,，;；<>"']+\.[A-Za-z]{2,}$/;
/* 文字 → Email 陣列（逗號、分號、空白分隔，最多 3 個）；strict＝有打錯的就擋（管理者存檔時用），否則略過打錯的 */
function emailList_(v, strict) {
  var out = [], bad = [];
  str_(v).split(/[,，、;；\s]+/).map(function (x) { return x.trim(); }).filter(Boolean).forEach(function (x) {
    if (EMAIL_RE.test(x)) { if (out.indexOf(x) < 0) out.push(x); } else bad.push(x);
  });
  if (strict && bad.length) throw fail_('Email 格式不對：' + bad.join('、'));
  if (strict && out.length > 3) throw fail_('Email 最多填 3 個');
  return out.slice(0, 3);
}
function dessertOf_(row) { return str_(row['商品正式名稱']).trim() || str_(row['商品暫定名稱']).trim() || str_(row.req_id); }
/* 現在輪到第幾批＝還沒同意的人裡最小的批次；全部都同意＝0 */
function turnOf_(names, agreed, bmap) {
  var pend = names.filter(function (n) { return !agreed[n]; });
  return pend.length ? Math.min.apply(null, pend.map(function (n) { return bmap[n] || 1; })) : 0;
}
/* 排一封「輪到你簽核」：寄給現在輪到的那一批裡還沒同意的人 */
function nqTurn_(row, names, agreed, round, camp, why) {
  var bmap = batchMap_(), b = turnOf_(names, agreed, bmap);
  if (!b) return;
  var people = names.filter(function (n) { return !agreed[n] && (bmap[n] || 1) === b; });
  if (people.length) _nq.push({ kind: 'turn', why: why, rid: str_(row.req_id), cid: str_(row.campaign_id), dessert: dessertOf_(row), round: str_(round), batch: b, people: people, camp: camp || '' });
}
function h_(s) { return str_(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
/* 純文字信 → 簡單的 HTML（連結做成按鈕） */
function mailHtml_(txt, link, btn) {
  var body = h_(txt).replace(/\n/g, '<br>');
  if (link) body = body.replace(h_(link), '<a href="' + h_(link) + '" style="display:inline-block;margin:6px 0;padding:10px 18px;background:#00796B;color:#fff;border-radius:6px;text-decoration:none;font-weight:bold">' + h_(btn || '打開') + '</a>');
  return '<div style="font-family:-apple-system,BlinkMacSystemFont,\'PingFang TC\',\'Microsoft JhengHei\',sans-serif;font-size:15px;line-height:1.7;color:#222;max-width:560px">' + body + '</div>';
}
function signUrl_(id, isCamp, who) {
  return NOTIFY_URL + '#sign=' + encodeURIComponent(id) + '&k=' + encodeURIComponent(signKey_(isCamp ? 'camp:' + id : id)) + (who ? '&who=' + encodeURIComponent(who) : '');
}
function mailTurn_(g, camps) {
  var cn = camps[g.cid] || '', ds = g.items.map(function (e) { return e.dessert; }), first = g.items.every(function (e) { return e.why === 'submit'; });
  var link = g.camp ? signUrl_(g.camp, true, g.name) : signUrl_(g.rid, false, g.name);
  var txt = g.name + ' 你好：\n\n' + (first ? '主廚送出' + (cn ? '「' + cn + '」' : '') + '的新品申請，請你簽核' : (cn ? '「' + cn + '」' : '') + '新品申請的前一批簽核人已經全部同意，現在輪到你簽核') + '（你是第 ' + g.batch + ' 批）：\n'
    + ds.map(function (d) { return '・' + d; }).join('\n') + '\n\n點這裡看內容、簽核：\n' + link + '\n\n'
    + '第一次簽核要先設定 4～8 位數字的 PIN（自己設，記住就好）。\n這封是系統自動寄出的通知，不用回信。\n' + NOTIFY_FROM;
  return { subject: '【食譜系統】輪到你簽核：' + (cn ? cn + ' ' : '') + (ds.length > 1 ? ds.length + ' 支新品' : ds[0]), text: txt, html: mailHtml_(txt, link, '打開簽核頁') };
}
function mailChef_(g, camps) {
  var rj = g.items.filter(function (e) { return e.kind === 'reject'; }), dn = g.items.filter(function (e) { return e.kind === 'done'; });
  var cn = camps[(g.items[0] || {}).cid] || '', link = NOTIFY_URL + '#apply', lines = [];
  rj.forEach(function (e) { lines.push('✕ 退回：' + e.dessert + '（' + e.by + '：' + (e.comment || '沒有寫原因') + '）'); });
  dn.forEach(function (e) { lines.push('✓ 已核准：' + e.dessert + '（' + (e.status === PUSH_OK ? '已寫入採購系統' : '寫入採購系統失敗' + (e.push ? '：' + e.push : '') + '，請到新品申請表按「📦 寫入採購系統」重試') + '）'); });
  var head = [rj.length ? rj.length + ' 支被退回' : '', dn.length ? dn.length + ' 支已核准' : ''].filter(Boolean).join('、');
  var txt = '主廚你好：\n\n' + (cn ? '「' + cn + '」' : '') + '新品申請的簽核結果：\n' + lines.join('\n') + '\n\n'
    + (rj.length ? '被退回的請到新品申請表「📂 從檔期載入」修改後重新送簽。\n' : '') + '新品申請表：\n' + link + '\n\n這封是系統自動寄出的通知，不用回信。\n' + NOTIFY_FROM;
  return { subject: '【食譜系統】' + (cn ? cn + ' ' : '') + head + (dn.length === 1 && !rj.length ? '：' + dn[0].dessert : '') + (rj.length === 1 && !dn.length ? '：' + rj[0].dessert : ''), text: txt, html: mailHtml_(txt, link, '打開新品申請表') };
}
/* doPost 尾端呼叫：把 _nq 依收件人合併成一封封信寄出；回 { sent:[名字], noemail:[名字], fail:[說明] }（不含 Email 地址） */
function flushNotify_() {
  var q = _nq; _nq = [];
  if (!q.length) return null;
  var out = { sent: [], noemail: [], fail: [] }, logs = [];
  try {
    var em = {}, chef = [];
    authRows_('dim_signer').forEach(function (r) { var n = str_(r.name).trim(); if (n && signerOn_(r.enabled)) em[n] = emailList_(r.email); });
    authRows_('dim_role').forEach(function (r) { if (str_(r.role).trim() === CHEF_ROLE) chef = emailList_(r.email); });
    var camps = {}; load_('fact_campaign').rows.forEach(function (c) { camps[str_(c.campaign_id)] = str_(c['檔期名稱']).trim(); });
    var done = {};   /* 同一次請求裡同一件事只排一次 */
    var groups = {}, order = [];
    var add = function (gk, base, e, key) { if (!groups[gk]) { groups[gk] = base; base.items = []; base.keys = []; order.push(gk); } groups[gk].items.push(e); groups[gk].keys.push(key); };
    q.forEach(function (e) {
      if (e.kind === 'turn') e.people.forEach(function (n) {
        var key = 'turn|' + e.rid + '|' + e.round + '|' + e.batch + '|' + n;
        if (done[key]) return; done[key] = 1;
        add('T|' + n + '|' + (e.camp ? 'C:' + e.camp : 'R:' + e.rid), { kind: 'turn', name: n, camp: e.camp, cid: e.cid, rid: e.rid, batch: e.batch }, e, key);
      });
      else {
        var key2 = e.kind + '|' + e.rid + '|' + e.round + '|' + e.status;
        if (done[key2]) return; done[key2] = 1;
        add('CHEF', { kind: 'chef', name: CHEF_ROLE }, e, key2);
      }
    });
    var quota = order.length ? MailApp.getRemainingDailyQuota() : 0;
    order.forEach(function (gk) {
      var g = groups[gk], to = g.kind === 'turn' ? (em[g.name] || []) : chef, ts = now_();
      var rec = function (ok, msg) { g.items.forEach(function (e, i) { logs.push([ts, g.kind === 'turn' ? 'turn' : e.kind, g.keys[i], e.rid, g.name, to.join(', '), ok ? 'Y' : 'N', msg]); }); };
      if (!to.length) { if (out.noemail.indexOf(g.name) < 0) out.noemail.push(g.name); rec(false, '沒有 Email'); return; }
      if (quota < to.length) { out.fail.push(g.name + '：今天 Gmail 寄信額度用完了'); rec(false, '寄信額度用完'); return; }
      try {
        var m = g.kind === 'turn' ? mailTurn_(g, camps) : mailChef_(g, camps);
        MailApp.sendEmail({ to: to.join(','), subject: m.subject, body: m.text, htmlBody: m.html, name: NOTIFY_FROM });
        quota -= to.length;
        if (out.sent.indexOf(g.name) < 0) out.sent.push(g.name); rec(true, m.subject);
      } catch (e1) { out.fail.push(g.name + '：' + errMsg_(e1)); rec(false, errMsg_(e1)); }
    });
  } catch (e2) { out.fail.push('通知沒有寄出：' + errMsg_(e2)); }
  notifyLog_(logs);
  return out;
}
/* 寄送紀錄寫進 fact_notify（短暫拿鎖；拿不到就不記——不影響已寄出的信） */
function notifyLog_(rows) {
  if (!rows.length) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return;
  try {
    var sh = ss_().getSheetByName('fact_notify');
    if (!sh) return;
    var r = sh.getLastRow() + 1;
    if (r + rows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), Math.max(200, rows.length));
    sh.getRange(r, 1, rows.length, 8).setNumberFormat('@').setValues(rows.map(function (x) { return x.map(function (v) { return safe_(str_(v).slice(0, 500)); }); }));
  } catch (e) { } finally { lock.releaseLock(); }
}
/* 編輯器執行一次：讓 Google 跳出「用你的 Gmail 寄信」授權（只查今天還能寄幾封，不寄信） */
function notifyAuthorize() { var n = MailApp.getRemainingDailyQuota(); Logger.log('Gmail 寄信權限 OK，今天還能寄 ' + n + ' 封'); return n; }

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
var PT = { map: '產品名稱對照表', bom: 'BOM表', sku: 'dim_sku', sku2: 'dim_sku（儀表板專用檔）', conv: 'dim_unitconv', par: 'dim_store_par', par2: 'dim_store_par（儀表板專用檔）' };
var PAR_NEED = ['店號', 'sku_id', '標配數量'];   /* 2026-10-07 晚：新器具／模具每店標配（店別參數只寫這三欄，其餘留空＝用預設） */
var BRAND_STORES = { '自己做': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], '吳寶春自己做': [11, 12], '自己做＆吳寶春自己做': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12] };   /* 同採購系統 EventLayer EV_BRAND_STORES */
function brandStores_(b) { b = ptName_(b); return (BRAND_STORES[b] || BRAND_STORES['自己做＆吳寶春自己做']).slice(); }
var PT_COLS = { map: 8, bom: 5, conv: 5 };                /* 讀寫範圍（欄數）；dim_sku 依表頭。dim_unitconv 只寫 A–E（F1 是單位對照網址，不碰） */
var PT_HEAD = {                                            /* 表頭對不上就停，不寫 */
  map: { 1: 'POS 資料產品名稱', 2: '手動對應產品名稱', 5: '起始有效日', 6: '結束有效日', 7: '主類別', 8: '次類別' },
  bom: { 1: '甜點名稱', 2: '食材/器具名稱', 3: '數量', 4: '單位', 5: '容器' },
  conv: { 1: '適用品項', 2: '單位', 3: '基準單位', 4: '換算量', 5: '備註' }   /* 2026-10-07：新單位換算（主廚在申請表填「1 平匙≈幾 g」） */
};
var CONV_COUNT = { '個': 1, '支': 1, '張': 1, '片': 1, '顆': 1, '包': 1, '罐': 1, '條': 1, '組': 1, '份': 1, '雙': 1, '對': 1, '朵': 1, '段': 1, '捲': 1, '卷': 1, '杯': 1, '碗': 1, '台': 1, '盒': 1, '瓶': 1, '袋': 1, '塊': 1, '根': 1 };
var CONV_GENERIC = { '（通用）': 1, '(通用)': 1, '通用': 1 };
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
  var sh = key === 'sku2' ? dss_().getSheetByName('dim_sku') : key === 'par2' ? dss_().getSheetByName('dim_store_par') : pss_().getSheetByName(PT[key]);
  if (!sh) throw fail_((key === 'sku2' || key === 'par2' ? '儀表板專用檔' : 'BOM 本') + '找不到分頁「' + PT[key] + '」', 'purchase');
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
  if (key === 'par' || key === 'par2') PAR_NEED.forEach(function (h) { if (col[h] === undefined) throw fail_(PT[key] + ' 找不到欄位「' + h + '」——先停下不寫', 'purchase'); });
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
      shelf: ptName_(l.shelf).replace(/\s*[|｜]\s*/g, '｜'), vname: ptName_(l.vname), supply: ptName_(l.supply),
      perStore: Number(l.per_store) > 0 ? Math.round(Number(l.per_store)) : 0,   /* 2026-10-07 晚：新器具／模具每店要幾個 */
      conv: Number(l.conv) > 0 ? Number(l.conv) : 0, cbase: ptName_(l.conv_base) };   /* 2026-10-07：新單位換算（1 unit ≈ conv cbase） */
  }).filter(function (l) { return l.name; });
  var P = { map: ptRead_('map'), bom: ptRead_('bom'), sku: ptRead_('sku'), sku2: null, conv: null };
  try { P.sku2 = ptRead_('sku2'); } catch (e) { P.sku2 = null; P.sku2err = errMsg_(e); }
  /* 單位換算表讀不到（分頁改名、表頭不對）→ 不擋核准寫入，只提醒「新單位要請採購手動加」 */
  try { P.conv = ptRead_('conv'); } catch (e) { P.conv = null; P.converr = errMsg_(e); }
  var convHave = {}, convGen = {}, convRows = [], convSeen = {}, convNote = DRAFT_TAG + ' ' + Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd') + '（新品申請 ' + rid + '：' + name + '；請採購確認）';
  if (P.conv) P.conv.vals.forEach(function (r) { var it = ptName_(r[0]), u = ptName_(r[1]); if (!it || !u) return; convHave[it + '|' + u] = 1; if (CONV_GENERIC[it]) convGen[u] = 1; });
  function convNeeded(unit, base) {   /* 和採購系統引擎同口徑：同單位、g／ml 互通、兩邊都是計數單位 → 不用換算 */
    if (!unit || !base || unit === base) return false;
    if ((unit === 'g' || unit === 'ml') && (base === 'g' || base === 'ml')) return false;
    if (CONV_COUNT[unit] && CONV_COUNT[base]) return false;
    return true;
  }
  function convHas(names, unit) { if (convGen[unit]) return true; for (var i = 0; i < names.length; i++) if (names[i] && convHave[names[i] + '|' + unit]) return true; return false; }
  function convAdd(l, base) {
    var key = l.name + '|' + l.unit; if (convSeen[key]) return; convSeen[key] = 1;
    convRows.push([l.name, l.unit, base, r4_(l.conv), convNote]);
  }

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
      /* 2026-10-07：主廚在申請表填了「1 平匙≈幾 g」→ 核准時新增一列 dim_unitconv（備註「自動草稿…請採購確認」）；沒填才提醒 */
      if (convNeeded(l.unit, k.unit) && !convHas([l.name, k.name], l.unit)) {
        if (l.conv > 0 && (!l.cbase || l.cbase === k.unit) && P.conv) convAdd(l, k.unit);
        else warn.push('「' + l.name + '」申請單寫 ' + l.unit + '、主檔使用單位是 ' + k.unit + '：' + (l.conv > 0 && l.cbase && l.cbase !== k.unit ? '申請表填的換算是 1 ' + l.unit + '≈' + l.conv + ' ' + l.cbase + '，和主檔單位不同，' : '') + '採購系統要有單位換算（dim_unitconv）才算得到用量' + (P.conv ? '' : '（單位換算表讀不到：' + P.converr + '）'));
      }
      return;
    }
    var cat = SKU_PREFIX[l.cat] ? l.cat : '食材', isSemi = !!semi[l.name];
    var pre = SKU_PREFIX[cat], no = (maxNo[pre] || 0) + 1; maxNo[pre] = no;
    var id = pre + ('000000' + no).slice(-(width[pre] || 3));
    var u = l.unit || 'g', buy = u, pack = 1, price = 0;
    /* 2026-10-07：新品項用「平匙」這類單位、又填了換算 → 主檔使用單位用基準單位（g／ml），另加一列換算（採購系統：使用單位必須＝換算的基準單位） */
    if (l.conv > 0 && l.cbase && convNeeded(l.unit, l.cbase)) {
      if (P.conv) { u = l.cbase; buy = u; if (!convHas([l.name], l.unit)) convAdd(l, l.cbase); }
      else warn.push('新品項「' + l.name + '」用 ' + l.unit + ' 計量（1 ' + l.unit + '≈' + l.conv + ' ' + l.cbase + '），但單位換算表讀不到（' + P.converr + '）：請採購手動加換算');
    } else if (/勺|匙|適量|少許|些許|酌量|滴/.test(u)) warn.push('新品項「' + l.name + '」用「' + u + '」計量、沒有填換算：採購系統會把 1 ' + u + ' 當 1 個單位算，請採購補單位換算');
    if (l.pack > 1) { var cm = l.pkgSpec.match(/[包袋箱盒罐瓶桶組件套盤打]/); buy = cm ? cm[0] : '包'; pack = l.pack; price = l.pkgCost; }
    else if (l.pack === 1) price = l.pkgCost;
    var dcSup = l.supply === '出貨中心出貨';   /* 2026-10-07 晚：提供方式＝出貨中心出貨 → 各店向出貨中心叫（主檔廠商寫「出貨中心」）；出貨中心向誰進貨寫在來源 */
    var src = DRAFT_TAG + ' ' + today + '（新品申請 ' + rid + '：' + name + (isSemi ? '；店製半成品' : '') + (l.vname && l.vname !== l.name ? '；廠商品名 ' + l.vname : '')
      + (dcSup && l.vendor && l.vendor !== '出貨中心' ? '；出貨中心向 ' + l.vendor + ' 進貨' : '')
      + ((l.pkgCost || l.pkgSpec) ? '；原包裝 ' + (l.pkgCost ? l.pkgCost + ' 元' : '—') + '／' + (l.pkgSpec || '—') : '') + '）';
    var o = { sku_id: id, '品名': l.name, '品類別': cat, '廠商': dcSup ? '出貨中心' : (l.vendor || (isSemi ? '半成品' : '')), '使用單位': u, '採購單位': buy,
      '每採購單位內容量': pack, '單價': price, '預設分區': l.zone, '來源': src, '效期': l.shelf };
    var arr = S.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; });
    var arr2 = P.sku2 ? P.sku2.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; }) : null;
    skuRows.push({ sku_id: id, name: l.name, vname: l.vname, cat: cat, vendor: o['廠商'], unit: u, buy: buy, pack: pack, price: price, semi: isSemi, arr: arr, arr2: arr2 });
  });

  /* ④（寫在主檔之後）2026-10-07 晚：這次新建的器具／模具，申請表填了「每店要幾個」→ 各店 dim_store_par 一列（店號、sku_id、標配數量），
     店＝品牌別對應的門市（且店別參數裡有這家店）；BOM 本＋儀表板專用檔都寫（專用檔寫不到不擋，07:05 會同步）。沒填 → 提醒 */
  var parRows = [], parRows2 = [], parShow = [], lineBy = {};
  lines.forEach(function (l) { if (!lineBy[l.name]) lineBy[l.name] = l; });
  var newTools = skuRows.filter(function (r) { return r.cat === '器具' || r.cat === '模具'; });
  if (newTools.length) {
    try { P.par = ptRead_('par'); } catch (e) { P.par = null; P.parerr = errMsg_(e); }
    try { P.par2 = ptRead_('par2'); } catch (e) { P.par2 = null; P.par2err = errMsg_(e); }
    /* 店＝甜點品牌別 ∩ 檔期品牌別（檔期只在自己做、甜點預設兩個品牌 → 只放自己做 1～10 店；同採購系統檔期新品用檔期品牌別） */
    var bD = ptName_(h.brand), bC = camp ? ptName_(camp['品牌別']) : '', stores = brandStores_(bD || bC), haveSt = {};
    if (bD && bC && BRAND_STORES[bD] && BRAND_STORES[bC]) stores = stores.filter(function (n) { return BRAND_STORES[bC].indexOf(n) >= 0; });
    if (!stores.length) warn.push('甜點品牌別「' + bD + '」和檔期品牌別「' + bC + '」沒有共同的店：新器具／模具各店標配沒有寫，請總部到採購系統店別參數設定');
    if (P.par) P.par.vals.forEach(function (r) { var n = Number(r[P.par.col['店號']]); if (n > 0) haveSt[n] = 1; });
    if (Object.keys(haveSt).length) stores = stores.filter(function (n) { return haveSt[n]; });
    newTools.forEach(function (r) {
      var l = lineBy[r.name] || {};
      if (!(l.perStore > 0)) { warn.push('新' + r.cat + '「' + r.name + '」申請表沒填「每店要幾個」：各店標配是 0（採購系統不會提醒補貨），請總部到採購系統店別參數設定'); return; }
      if (!P.par) { warn.push('店別參數讀不到（' + P.parerr + '）：新' + r.cat + '「' + r.name + '」每店 ' + l.perStore + ' 個要請總部手動設定標配'); return; }
      stores.forEach(function (st) {
        var o = { '店號': st, 'sku_id': r.sku_id, '標配數量': l.perStore };
        parRows.push(P.par.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; }));
        if (P.par2) parRows2.push(P.par2.H.map(function (hh) { return o[hh] === undefined ? '' : o[hh]; }));
        parShow.push([st, r.sku_id, r.name, l.perStore]);
      });
    });
    if (parRows.length) warn.push('新器具／模具每店標配 ' + parShow.length + ' 列會加進店別參數：' + newTools.filter(function (r) { return (lineBy[r.name] || {}).perStore > 0; }).map(function (r) { return r.name + ' 每店 ' + lineBy[r.name].perStore + ' 個（' + stores.length + ' 店）'; }).join('、'));
    if (parRows.length && !P.par2) warn.push('儀表板專用檔的店別參數讀不到（' + P.par2err + '）：明天 07:05 會同步');
  }

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

  /* ③ 對照表
     2026-10-01 需求 7（POS 分類改選「主類別」，head.cat_kind==='main'）：G＝選的主類別（要在對照表現有主類別內，否則用預設）；
       H＝檔期名稱（「YYYY 名稱」統一一個空白）——採購系統 EventLayer 靠 H 認檔期啟動首批備料，H 一律寫檔期名，不寫主類別。
     舊單（cat＝次類別）：沿用舊規則＝主類別取對照表裡同一個次類別最常用的主類別（沒有就「限定甜點」）；次類別＝POS 分類（沒選或選「無」就用檔期名稱） */
  var catTxt = ptName_(h.cat_text || (h.cat === '__other' ? h.catOther : h.cat));
  var campName = camp ? ptName_(camp['檔期名稱']) : '';
  var sub, cnt = {}, main = '', best = 0, mainsAll = {};
  P.map.vals.forEach(function (r) { var g0 = ptName_(r[6]); if (g0) mainsAll[g0] = 1; });
  if (h.cat_kind === 'main' || (catTxt && mainsAll[catTxt])) {   /* 營運POS 局部修改時 cat_kind 不一定有 → 值是現有主類別就當主類別 */
    main = mainsAll[catTxt] ? catTxt : '';
    if (catTxt && !main) warn.push('POS 分類「' + catTxt + '」不在對照表現有主類別，改用「' + DEFAULT_MAIN + '」');
    var campSub = campName.replace(/^(20\d{2})\s*/, '$1 ');
    sub = campSub || '無';
    if (campSub && !/^20\d{2} \S/.test(campSub)) warn.push('檔期名稱「' + campName + '」不是「YYYY 名稱」格式（例：2026 萬聖節）：採購系統的檔期備料會認不出這個檔期');
  } else {
    sub = (catTxt && catTxt !== '無') ? catTxt : (campName || catTxt || '無');
    P.map.vals.forEach(function (r) { if (ptName_(r[7]) === sub) { var g = ptName_(r[6]); if (g) cnt[g] = (cnt[g] || 0) + 1; } });
    Object.keys(cnt).forEach(function (g) { if (cnt[g] > best) { best = cnt[g]; main = g; } });
  }
  main = main || DEFAULT_MAIN;
  var d1 = camp ? normDate_(camp['起']) : '', d2 = camp ? normDate_(camp['迄']) : '';
  var dE = camp ? normDate_(camp['自己人搶先開賣日']) : '';   /* 2026-10-05：有搶先開賣日就用它當對照表起始有效日（POS 那天起就會賣、採購系統以它提前 14 天備料） */
  if (dE && /^\d{4}-\d{2}-\d{2}$/.test(dE) && (!d1 || dE < d1)) { warn.push('這個檔期有「自己人搶先開賣日」' + dE + '：對照表的起始日用它（正式開賣 ' + (d1 || '未填') + '），採購系統會以 ' + dE + ' 當開賣日提前備料'); d1 = dE; }
  if (!d1 || !d2) warn.push('檔期沒有填起訖日：對照表的有效日期留空，採購系統的新品備料不會啟動（到對照表補日期即可）');
  var mapHave = P.map.vals.filter(function (r) { return ptName_(r[0]) === name; })[0];
  var map = { row: [name, name, '', '', ymd_(d1), ymd_(d2), main, sub], show: [name, name, d1, d2, main, sub],
    skip: mapHave ? '對照表已經有「' + name + '」（' + normDate_(mapHave[4]) + '～' + normDate_(mapHave[5]) + '），沒有新增（重新上架請到對照表改日期）' : '' };
  if (skuRows.length && !P.sku2) warn.push('儀表板專用檔讀不到（' + P.sku2err + '）：主檔新品項只寫 BOM 本、明天 07:05 同步；這段期間請不要在採購系統主檔頁新增品項（會撞號）');
  if (!skuRows.length && bom.skip && map.skip && !convRows.length) warn.push('三張表都不用寫（都已經有了）');
  if (convRows.length) warn.push('新單位換算 ' + convRows.length + ' 筆會加進 dim_unitconv（備註「' + DRAFT_TAG + '…請採購確認」）：' + convRows.map(function (c) { return c[0] + ' 1 ' + c[1] + '≈' + c[3] + ' ' + c[2]; }).join('、') + '，請採購到採購系統「單位對照」確認數字');
  return { rid: rid, name: name, campaign: campName, from: d1, to: d2, sku: { rows: skuRows, existing: existing }, bom: bom, map: map, conv: { rows: convRows },
    par: { rows: parRows, rows2: parRows2, show: parShow }, warn: warn, P: P };
}
function planView_(pl) {
  return { req_id: pl.rid, name: pl.name, campaign: pl.campaign, from: pl.from, to: pl.to, warn: pl.warn, dash: !!pl.P.sku2,
    sku: pl.sku.rows.map(function (r) { return { sku_id: r.sku_id, name: r.name, vname: r.vname || '', cat: r.cat, vendor: r.vendor, unit: r.unit, buy: r.buy, pack: r.pack, price: r.price, semi: r.semi }; }),
    existing: pl.sku.existing, bom: pl.bom.rows, bomSkip: pl.bom.skip, map: pl.map.show, mapSkip: pl.map.skip,
    conv: (pl.conv ? pl.conv.rows : []).map(function (r) { return r.slice(0, 4).map(ptNorm_); }),
    par: (pl.par ? pl.par.show : []).map(function (r) { return r.map(ptNorm_); }) };
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
    if (tab.key === 'conv') return ptName_(r[0]) === x[0] && ptName_(r[1]) === x[1] && ptName_(r[2]) === x[2] && ptSame_(r[3], x[3]);
    if (tab.key === 'par' || tab.key === 'par2') { var cs = tab.col['店號'], cq = tab.col['sku_id']; return ptSame_(r[cs], x[cs]) && ptName_(r[cq]) === String(x[cq]); }
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
  ['sku', 'bom', 'map', 'sku2', 'conv', 'par', 'par2'].forEach(function (key) {   /* 2026-10-07 加 conv（dim_unitconv）；晚上再加 par／par2（店別參數）：還原時讀第 3～9 列 */
    var tab = pl.P[key], n = (key === 'sku' || key === 'sku2') ? pl.sku.rows.length : key === 'bom' ? (pl.bom.skip ? 0 : pl.bom.rows.length) : key === 'conv' ? pl.conv.rows.length
      : key === 'par' ? pl.par.rows.length : key === 'par2' ? pl.par.rows2.length : (pl.map.skip ? 0 : 1);
    if (!tab) { out.push([key, PT[key], '', '讀不到', '0', (key === 'conv' ? pl.P.converr : key === 'par' ? (pl.P.parerr || (pl.par.rows.length ? '' : '這次沒有新器具／模具標配')) : key === 'par2' ? (pl.P.par2err || '') : pl.P.sku2err) || '']); return; }
    tab.digest = ptDigest_(tab);
    out.push([key, PT[key], String(tab.lastA), tab.digest, String(n), key === 'bom' && pl.bom.skip ? pl.bom.skip : key === 'map' && pl.map.skip ? pl.map.skip : '']);
  });
  out.push(['', '', '', '', '', '']);
  out.push(['表', '要寫的內容（JSON，一列一筆）', '', '', '', '']);
  pl.sku.rows.forEach(function (r) { out.push(['sku', JSON.stringify(r.arr.map(ptNorm_)), '', '', '', '']); });
  if (!pl.bom.skip) pl.bom.rows.forEach(function (r) { out.push(['bom', JSON.stringify(r.map(ptNorm_)), '', '', '', '']); });
  if (!pl.map.skip) out.push(['map', JSON.stringify(pl.map.row.map(ptNorm_)), '', '', '', '']);
  pl.conv.rows.forEach(function (r) { out.push(['conv', JSON.stringify(r.map(ptNorm_)), '', '', '', '']); });
  pl.par.rows.forEach(function (r) { out.push(['par', JSON.stringify(r.map(ptNorm_)), '', '', '', '']); });
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
    if (pl.par.rows.length && pl.P.par) done.push(ptAppend_(pl.P.par, pl.par.rows));   /* 2026-10-07 晚：新器具／模具每店標配 */
    if (pl.par.rows2.length && pl.P.par2) {
      try { done.push(ptAppend_(pl.P.par2, pl.par.rows2)); }
      catch (e3) { soft.push('儀表板專用檔的店別參數沒寫到（' + errMsg_(e3) + '）：明天 07:05 會同步'); }
    }
    if (!pl.bom.skip) done.push(ptAppend_(pl.P.bom, pl.bom.rows));
    if (pl.conv.rows.length && pl.P.conv) done.push(ptAppend_(pl.P.conv, pl.conv.rows));   /* 2026-10-07：新單位換算（對照表仍最後寫＝備料啟動開關） */
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
  var show = { sku: pl.sku.rows.map(function (r) { return [r.sku_id, r.name, r.cat, r.vendor]; }), bom: pl.bom.rows.map(function (r) { return r.map(ptNorm_); }), map: [pl.map.show],
    conv: pl.conv.rows.map(function (r) { return r.slice(0, 4).map(ptNorm_); }), par: pl.par.show.map(function (r) { return r.map(ptNorm_); }) };
  show.sku2 = show.sku; show.par2 = show.par;
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
    if (bs) bs.getRange(3, 1, 7, 4).getValues().forEach(function (r) { if (String(r[0])) bkDig[String(r[0])] = String(r[3]); });   /* 2026-10-07 起第 7 列是 conv、晚上起第 8、9 列是 par／par2；舊備份那幾列是空白 */
  } catch (e) { }
  ['map', 'conv', 'bom', 'par', 'par2', 'sku', 'sku2'].forEach(function (key) {   /* 店別參數比主檔先刪（列裡有新品項的 sku_id） */
    var rec = st.tables[key];
    if ((key === 'sku2' || key === 'conv' || key === 'par' || key === 'par2') && !rec) return;   /* 當初沒寫到專用檔／這張單沒有新單位換算／沒有新器具標配 */
    var tab;
    try { tab = ptRead_(key); } catch (e) {
      if (key !== 'sku2' && key !== 'par2') throw e;
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
  /* v13：主廚確認 104 表單已簽核通過 → 先改成「已核准」（簽核不在本系統做了） */
  if (!PUSHABLE[st] && st !== PUSH_OK && p.approved104) {
    setStatus_(rid, '已核准');
    log_(auth.role, 'approve104', rid, true, '主廚確認 104 簽核通過（原狀態：' + (st || '草稿') + '）');
    st = '已核准';
  }
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
  var cache = CacheService.getScriptCache(), ck = 'pub:newitems:v6', hit = cache.get(ck);
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
      campaign: str_(c['檔期名稱']), from: normDate_(c['起']), to: normDate_(c['迄']), early: normDate_(c['自己人搶先開賣日']), lines: lines,
      qty: Number(r['預估銷售數']) || 0, brand: str_(c['品牌別']),
      ship: normDate_(c['第一批配貨日']),   /* 2026-10-07 晚：採購系統整檔總表顯示第一批配貨日、檔期作業時間 */
      ops: [['第一批出貨', str_(c['作業_第一批出貨_時間']).trim()], ['貼紙', str_(c['作業_貼紙_時間']).trim()], ['POS 上架', str_(c['作業_POS_時間']).trim()]].filter(function (o) { return o[1]; }) };   /* 2026-10-01 需求 6：預估銷售數（整檔 12 店合計）＋品牌別，採購「檔期新品」顯示首批量的來源 */
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
    var rd = function (rn) { return { a: rn[0], v: sh.getRange(2, rn[0] + 1, last - 1, rn[1] - rn[0] + 1).getValues() }; };
    var parts;
    try { parts = runs_(cols.filter(function (j) { return j >= 0; })).map(rd); }
    catch (eR) {   /* 2026-10-08：分頁比程式少了最後幾欄（新欄位的遷移還沒做完）→ 只讀現有的欄，缺的欄當空白（不讓整個系統讀不到身分表） */
      var mc = sh.getMaxColumns();
      if (mc >= H.length) throw eR;
      parts = runs_(cols.filter(function (j) { return j >= 0 && j < mc; })).map(rd);
    }
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
  var cand = t.rows.filter(function (r) { return str_(r.created_at) >= lim; });
  if (!cand.length) return null;
  /* 2026-10-07 加速：24 小時內的列一次讀（原本每列各讀一次） */
  var lo = cand[0]._row, hi = cand[cand.length - 1]._row;
  cand.forEach(function (r) { if (r._row < lo) lo = r._row; if (r._row > hi) hi = r._row; });
  var col = t.sh.getRange(lo, a, hi - lo + 1, 1).getValues();
  for (var i = cand.length - 1; i >= 0; i--) {
    var r = cand[i];
    if (str_(col[r._row - lo][0]).slice(0, 400).indexOf(key) >= 0) return r;
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

/* ================= 📚 BOM 管理（2026-10-01 經營者「20261001 待處理事項」#7）=================
   目的：BOM表／產品名稱對照表 改在食譜系統「📚 BOM 管理」維護，人不再直接改 BOM 本（Google 試算表繼續當後端資料庫）。
   讀：bomMeta（清單、選項、權限）、bomGet（一支甜點：列＋指紋）、mapGet（對照表全部列＋每列指紋）、bomLog（異動紀錄）
   寫：bomSave（新增甜點／整支取代）、bomDelete、bomRename（連同對照表 B 欄）、mapSave（新增／修改 A,B,E,F,G,H；C/D 永不寫）、mapDelete、bomUndo
   安全：①寫前指紋比對（不同＝有人改過 → 不覆蓋，回最新內容）②只寫 BOM表 A–E、對照表 A,B,E–H，表頭不對就停（ptRead_）
         ③寫後讀回逐格比對，不一致自動寫回原值 ④每次寫入記「食譜系統_申請單」fact_bom_log（before/after JSON＝版本紀錄，可一鍵還原）
         ⑤06:00–07:45 採購系統在讀 BOM 重算（autodraft／rebuild／EventLayer）→ 不寫
   權限（dim_role.fields，管理者 * 全有）：B:bom.edit（新增／修改／匯入食譜）、B:bom.delete（刪除／改名）、
         B:map.edit（對照表新增、改 B／日期）、B:map.cat（改既有列的主／次類別＝會改寫該品名全部歷史分類）、B:map.delete、B:undo */
var BOM_MAX_ROWS = 300;
var BOM_BASE_UNITS = ['g', 'ml', 'kg', '個', '支', '張', '片', '顆', '包', '罐', '條', '組', '份', '雙', '對', '朵', '段', '捲', '卷', '杯', '碗', '台', '盒', '瓶', '袋', '塊', '根', '匙', '平匙', '滴'];
var BOM_CONTAINERS = ['', '彩色塑膠碗', '白色瓷碗', '擠花袋', '水杯', '量杯', '無', '工作盆'];
/* POS 儀表板白名單（計甜點數／來客數的主類別）：這些主類別的對照列必須對到 BOM 甜點（找不到＝暫時佔位，可空） */
var BOM_DESSERT_MAINS = { '乳酪&奶蓋&慕斯': 1, '巧克力': 1, '裝飾蛋糕': 1, '限定甜點': 1, '水果': 1, '生日蛋糕': 1, '慶祝蛋糕': 1, '其他': 1,
  '主題活動＆群友限定': 1, '雙層蛋糕': 1, '蛋糕': 1, '點心&餅乾': 1, '塔派': 1, '吳寶春大獎麵包': 1, '吳寶春麵包': 1 };
var BOM_QUIET = [6 * 60, 7 * 60 + 45];
var BOM_PERM_LABEL = { 'bom.edit': 'BOM 新增／修改', 'bom.delete': 'BOM 刪除／改名', 'map.edit': '對照表新增／修改', 'map.cat': '對照表改主／次類別',
  'map.delete': '對照表刪除', 'undo': '還原異動' };
TABS.fact_bom_log = ['ts', 'batch', 'role', 'action', 'table', 'key', 'source', 'before_json', 'after_json', 'digest_before', 'digest_after', 'rows_before', 'rows_after', 'ok', 'msg'];
NUM_COLS.rows_before = 1; NUM_COLS.rows_after = 1;

function bomPerm_(auth, key) { return !!(auth && auth.kind === 'role' && can_(auth, 'B:' + key)); }
function needB_(auth, key) { if (!bomPerm_(auth, key)) throw fail_('這個動作要有「' + (BOM_PERM_LABEL[key] || key) + '」權限（管理者可以；其他身分請管理者在 🔑 密碼管理開權限）', 'perm'); }
function bomQuiet_() {
  var hm = Utilities.formatDate(new Date(), TZ, 'HH:mm').split(':'), m = (+hm[0]) * 60 + (+hm[1]);
  if (m >= BOM_QUIET[0] && m < BOM_QUIET[1]) throw fail_('採購系統每天 06:00–07:45 在讀 BOM 表重算建議量，這段時間先不寫（避免讀到寫一半的配方）。請 07:45 以後再按儲存。', 'quiet');
}
function bomDigest_(rows) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(rows || []));
  return d.slice(0, 8).map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('') + '·' + (rows || []).length + '列';
}
function bomRowNorm_(r) { return [ptName_(r[0]), ptName_(r[1]), ptNorm_(r[2] === '' ? '' : Number(r[2])), ptName_(r[3]), ptName_(r[4])]; }
/* 一支甜點在 BOM表 的列（列號、正規化值、是否連續） */
function bomBlock_(tab, name) {
  var idx = [], rows = [];
  for (var i = 0; i < tab.lastA - 1; i++) if (ptName_(tab.vals[i][0]) === name) { idx.push(i + 2); rows.push(bomRowNorm_(tab.vals[i])); }
  var contiguous = idx.every(function (rn, k) { return k === 0 || rn === idx[k - 1] + 1; });
  return { idx: idx, rows: rows, contiguous: contiguous };
}
function bomDesserts_(tab) {
  var o = {}, order = [];
  for (var i = 0; i < tab.lastA - 1; i++) { var n = ptName_(tab.vals[i][0]); if (!n) continue; if (!o[n]) { o[n] = 0; order.push(n); } o[n]++; }
  return order.map(function (n) { return { name: n, n: o[n] }; });
}
function bomUnits_(tab) {
  var u = {};
  BOM_BASE_UNITS.forEach(function (x) { u[x] = 1; });
  tab.vals.forEach(function (r) { var x = ptName_(r[3]); if (x) u[x] = 1; });
  try {
    var sh = pss_().getSheetByName('dim_unitconv');
    if (sh && sh.getLastRow() > 1) sh.getRange(2, 2, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { [r[0], r[1]].forEach(function (x) { x = ptName_(x); if (x) u[x] = 1; }); });
  } catch (e) { /* 讀不到單位對照表 → 只用 BOM 現有單位＋基本單位 */ }
  return u;
}
function bomContainers_(tab) {
  var c = {};
  BOM_CONTAINERS.forEach(function (x) { c[x] = 1; });
  tab.vals.forEach(function (r) { c[ptName_(r[4])] = 1; });
  return c;
}
/* 驗證：甜點名、列數、品項、數量、單位、容器；錯誤一次全部列出（code invalid，extra＝錯誤清單） */
function bomCheckRows_(name, rows, tab) {
  var err = [];
  if (!name) err.push('甜點名稱必填');
  else if (name.length > 60 || /[\r\n\t]/.test(name)) err.push('甜點名稱太長或含換行');
  if (!Array.isArray(rows) || !rows.length) err.push('至少要有 1 列用料');
  else if (rows.length > BOM_MAX_ROWS) err.push('一支甜點最多 ' + BOM_MAX_ROWS + ' 列（現在 ' + rows.length + ' 列）');
  var U = bomUnits_(tab), C = bomContainers_(tab), out = [];
  (rows || []).slice(0, BOM_MAX_ROWS).forEach(function (r, i) {
    var k = '第 ' + (i + 1) + ' 列';
    var m = ptName_(r && r[0]), q = Number(r && r[1]), u = ptName_(r && r[2]), c = ptName_(r && r[3]);
    if (!m) err.push(k + '：品項必填');
    else if (m.length > 60 || /[\r\n\t]/.test(m)) err.push(k + '：品項名稱太長或含換行');
    if (r && (r[1] === '' || r[1] === null || r[1] === undefined)) err.push(k + '「' + m + '」：數量必填');
    else if (!isFinite(q) || q <= 0) err.push(k + '「' + m + '」：數量要是大於 0 的數字');
    else if (q > 1e6) err.push(k + '「' + m + '」：數量太大');
    if (!u) err.push(k + '「' + m + '」：單位必填');
    else if (!U[u]) err.push(k + '「' + m + '」：單位「' + u + '」不在允許清單（BOM 現有單位＋單位對照表）');
    if (!C[c]) err.push(k + '「' + m + '」：容器「' + c + '」不在允許清單');
    out.push([name, m, Math.round(q * 1e4) / 1e4, u, c]);
  });
  if (err.length) throw fail_('資料沒有通過檢查，沒有寫入：' + err.slice(0, 8).join('；') + (err.length > 8 ? '…等 ' + err.length + ' 項' : ''), 'invalid', { errors: err });
  return out;
}
/* 整支取代：同列數原地寫、變多在區塊尾插列、變少刪多的；0 列＝刪除；新甜點接在 A 欄最後一列之後。寫完讀回逐格比對 */
function bomWriteBlock_(tab, blk, rows) {
  var sh = tab.sh, w = 5, n = blk.idx.length, m = rows.length, start;
  var put = rows.map(function (r) { return [safe_(r[0]), safe_(r[1]), Number(r[2]), safe_(r[3]), safe_(r[4])]; });
  if (n && !blk.contiguous) throw fail_('這支甜點在 BOM 表的列不連續（第 ' + blk.idx.join('、') + ' 列），為了安全不自動改，請管理者先整理', 'layout');
  if (!n) {
    if (!m) return 0;
    start = tab.lastA + 1;
    if (start + m - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), start + m - 1 - sh.getMaxRows());   /* 只加剛好需要的列（BOM 本容量吃緊） */
    var before = sh.getRange(start, 1, m, w).getValues();
    if (before.some(function (r) { return r.some(function (v) { return ptNorm_(v) !== ''; }); })) throw fail_('BOM 表第 ' + start + ' 列之後不是空的（可能剛好有人在寫），這次先不寫', 'busy');
  } else {
    start = blk.idx[0];
    if (m > n) sh.insertRowsAfter(start + n - 1, m - n);
  }
  if (m) {
    [1, 2, 4].forEach(function (c) { sh.getRange(start, c, m, 1).setNumberFormat('@'); });
    sh.getRange(start, 1, m, w).setValues(put);
  }
  if (n && m < n) sh.deleteRows(start + m, n - m);
  SpreadsheetApp.flush();
  if (m) {
    var back = sh.getRange(start, 1, m, w).getValues();
    back.forEach(function (r, i) {
      for (var j = 0; j < w; j++) if (!ptSame_(r[j], rows[i][j])) throw fail_('BOM 表第 ' + (start + i) + ' 列第 ' + (j + 1) + ' 欄讀回不一致（寫「' + ptNorm_(rows[i][j]) + '」讀到「' + ptNorm_(r[j]) + '」）', 'verify');
    });
  }
  return start;
}
/* 失敗時把這支甜點寫回原內容（盡力而為） */
function bomRestore_(name, before) {
  try { var tab = ptRead_('bom'); bomWriteBlock_(tab, bomBlock_(tab, name), before.map(function (r) { return [r[0], r[1], Number(r[2]), r[3], r[4]]; })); return true; }
  catch (e) { return false; }
}
/* v6.1（2026-10-02 驗收發現）：原本「分鐘＋3 位亂數」同一分鐘內多次寫入會撞號，還原一批時會連別批一起處理 → 改「秒＋6 碼隨機」 */
function bomBatch_() { return 'B' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 6); }
function bomLogW_(rec) {
  try {
    var t = load_('fact_bom_log');
    ['before_json', 'after_json'].forEach(function (k) { if (str_(rec[k]).length > 45000) rec[k] = '（內容超過 45,000 字，未存；此批無法一鍵還原）'; });
    put_(t, nextRow_(t), rec);
  } catch (e) { /* 紀錄失敗不影響主流程 */ }
}
function mapRowVals_(r) { return [ptName_(r[0]), ptName_(r[1]), normDate_(r[4]), normDate_(r[5]), ptName_(r[6]), ptName_(r[7])]; }
function mapRowDigest_(v) { return bomDigest_([v]).split('·')[0]; }
function mapDesserts_(tab) { var o = {}; for (var i = 0; i < tab.lastA - 1; i++) { var b = ptName_(tab.vals[i][1]); if (b) (o[b] = o[b] || []).push(ptName_(tab.vals[i][0])); } return o; }

/* v7 讀表暫存（2026-10-04 效能第 2 批）：bomMeta／mapGet 每次都把 BOM表（6,000+ 列）＋對照表整張讀一遍 → 算好的結果放 CacheService 5 分鐘。
   鍵尾帶版本號 bm:ver；任何會動到這兩張表的動作（BOM_TOUCH）做完就換版本號＝舊暫存立刻失效（比直接刪鍵保險：
   有人正在讀舊表、寫入者同時寫完清掉、讀的人再把舊資料放回去——換版本號後那筆放回去的是舊鍵，沒人會再讀到）。
   有人直接在 Google 試算表手改 BOM 本（不經本 API）→ 最多 5 分鐘後才看到；寫入時指紋比對擋得住（conflict 一樣會換版本號，重新讀取就是最新）。
   請求帶 fresh:true 可略過暫存（前端「重新讀取」日後可用）。bomGet 不暫存（每次讀最新，存檔前的指紋以它為準）。 */
var BM_TTL = 300;
function bmVer_(cache) { return cache.get('bm:ver') || '0'; }
function bomCacheClear_() {
  try { var c = CacheService.getScriptCache(); c.put('bm:ver', String((parseInt(c.get('bm:ver') || '0', 10) || 0) + 1), 21600); } catch (e) { /* 清不掉就等 5 分鐘自然過期 */ }
}
function bmCached_(name, fresh, build) {
  var cache = CacheService.getScriptCache(), key = 'bm:' + name + ':' + bmVer_(cache);
  if (!fresh) { var hit = cacheGetBig_(cache, key); if (hit) { try { return JSON.parse(hit); } catch (e) { /* 壞掉就重算 */ } } }
  var v = build();
  try { cachePutBig_(cache, key, JSON.stringify(v), BM_TTL); } catch (e) { /* 放不進去就每次算 */ }
  return v;
}
/* bomMeta 裡只跟表格內容有關的部分（權限、身分另外加） */
function bomMetaCore_(fresh) {
  return bmCached_('meta', fresh, function () {
    var tab = ptRead_('bom'), map = ptRead_('map'), U = bomUnits_(tab), C = bomContainers_(tab), mains = {}, subs = {};
    for (var i = 0; i < map.lastA - 1; i++) { var g = ptName_(map.vals[i][6]), h = ptName_(map.vals[i][7]); if (g) mains[g] = 1; if (h) subs[h] = 1; }
    return { desserts: bomDesserts_(tab), units: Object.keys(U), containers: Object.keys(C), mains: Object.keys(mains), subs: Object.keys(subs), mapB: mapDesserts_(map) };
  });
}
function mapGetCore_(fresh) {
  return bmCached_('map', fresh, function () {
    var map = ptRead_('map'), rows = [];
    for (var i = 0; i < map.lastA - 1; i++) {
      var v = mapRowVals_(map.vals[i]); if (!v[0] && !v[1]) continue;
      rows.push({ row: i + 2, a: v[0], b: v[1], e: v[2], f: v[3], g: v[4], h: v[5], digest: mapRowDigest_(v) });
    }
    var desserts = bomDesserts_(ptRead_('bom')).map(function (d) { return d.name; });
    return { rows: rows, desserts: desserts };
  });
}
function bomMeta_(p, auth) {
  var c = bomMetaCore_(!!p.fresh);
  var perms = {}; Object.keys(BOM_PERM_LABEL).forEach(function (k) { perms[k] = bomPerm_(auth, k); });
  return { id: '', data: { desserts: c.desserts, units: c.units, containers: c.containers, mains: c.mains, subs: c.subs,
    dessertMains: Object.keys(BOM_DESSERT_MAINS), mapB: c.mapB, perms: perms, role: auth.role, quiet: ['06:00', '07:45'], maxRows: BOM_MAX_ROWS } };
}
/* v7 合併查詢：一次回 bomMeta＋mapGet（前端開 BOM 管理分頁兩支可併成一支；這批前端未改，先提供） */
function bomMetaMap_(p, auth) {
  return { id: '', data: { meta: bomMeta_(p, auth).data, map: mapGet_(p, auth).data } };
}
function bomGet_(p, auth) {
  var name = ptName_(p.dessert); if (!name) throw fail_('缺甜點名稱');
  var tab = ptRead_('bom'), blk = bomBlock_(tab, name);
  var map = ptRead_('map'), pos = mapDesserts_(map)[name] || [];
  return { id: name, data: { dessert: name, rows: blk.rows, start: blk.idx[0] || 0, contiguous: blk.contiguous, digest: bomDigest_(blk.rows), pos: pos, exists: blk.idx.length > 0 } };
}
function bomSave_(p, auth) {
  needB_(auth, 'bom.edit'); bomQuiet_();
  var name = ptName_(p.dessert), isNew = !!p.is_new, src = str_(p.source).trim().slice(0, 200) || '手動';
  var tab = ptRead_('bom'), blk = bomBlock_(tab, name);
  if (isNew && blk.idx.length) throw fail_('BOM 表已經有「' + name + '」（' + blk.idx.length + ' 列），請改用修改', 'dup', { digest: bomDigest_(blk.rows), rows: blk.rows });
  if (!isNew) {
    if (!blk.idx.length) throw fail_('BOM 表找不到「' + name + '」（可能剛被改名或刪除），請重新讀取', 'notfound');
    var cur = bomDigest_(blk.rows);
    if (cur !== str_(p.base_digest)) throw fail_('「' + name + '」剛被別人改過（版本不同），為了不蓋掉別人的修改，這次沒有存。請按「重新讀取」看最新內容再改。', 'conflict', { digest: cur, rows: blk.rows });
  }
  var rows = bomCheckRows_(name, p.rows, tab), after = rows.map(bomRowNorm_), before = blk.rows;
  if (!isNew && bomDigest_(after) === bomDigest_(before)) return { id: name, data: { unchanged: true, digest: bomDigest_(before), rows: before, start: blk.idx[0] }, msg: '沒有變更' };
  var batch = bomBatch_(), start;
  try { start = bomWriteBlock_(tab, blk, rows); }
  catch (e) {
    var restored = before.length ? bomRestore_(name, before) : bomRestore_(name, []);
    bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: isNew ? 'bom.create' : 'bom.replace', table: 'BOM表', key: name, source: src, before_json: JSON.stringify(before), after_json: JSON.stringify(after),
      digest_before: bomDigest_(before), digest_after: '', rows_before: before.length, rows_after: after.length, ok: 'N', msg: errMsg_(e) + (restored ? '（已寫回原內容）' : '（⚠️ 寫回原內容失敗，請看備份）') });
    throw e;
  }
  bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: isNew ? 'bom.create' : 'bom.replace', table: 'BOM表', key: name, source: src, before_json: JSON.stringify(before), after_json: JSON.stringify(after),
    digest_before: bomDigest_(before), digest_after: bomDigest_(after), rows_before: before.length, rows_after: after.length, ok: 'Y', msg: '第 ' + start + ' 列起' });
  return { id: name, data: { batch: batch, start: start, digest: bomDigest_(after), rows: after }, msg: (isNew ? '新增' : '更新') + ' BOM「' + name + '」' + after.length + ' 列（' + src + '）' };
}
function bomDelete_(p, auth) {
  needB_(auth, 'bom.delete'); bomQuiet_();
  var name = ptName_(p.dessert), tab = ptRead_('bom'), blk = bomBlock_(tab, name);
  if (!blk.idx.length) throw fail_('BOM 表找不到「' + name + '」', 'notfound');
  var cur = bomDigest_(blk.rows);
  if (cur !== str_(p.base_digest)) throw fail_('「' + name + '」剛被別人改過，這次沒有刪除，請重新讀取', 'conflict', { digest: cur, rows: blk.rows });
  var pos = mapDesserts_(ptRead_('map'))[name] || [];
  if (pos.length) throw fail_('產品名稱對照表還有 ' + pos.length + ' 個 POS 品名對到「' + name + '」（' + pos.slice(0, 5).join('、') + '）：刪掉配方後這些銷售會展不出用料。請先到對照表改 B 欄或設結束日，再刪除。', 'inuse', { pos: pos });
  var batch = bomBatch_();
  try { bomWriteBlock_(tab, blk, []); }
  catch (e) { bomRestore_(name, blk.rows); throw e; }
  var left = bomBlock_(ptRead_('bom'), name);
  if (left.idx.length) throw fail_('刪除後讀回還有 ' + left.idx.length + ' 列，請重新讀取確認', 'verify');
  bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: 'bom.delete', table: 'BOM表', key: name, source: str_(p.source).slice(0, 200) || '手動',
    before_json: JSON.stringify(blk.rows), after_json: '[]', digest_before: cur, digest_after: bomDigest_([]), rows_before: blk.rows.length, rows_after: 0, ok: 'Y', msg: '' });
  return { id: name, data: { batch: batch, removed: blk.rows.length }, msg: '刪除 BOM「' + name + '」' + blk.rows.length + ' 列' };
}
function bomRename_(p, auth) {
  needB_(auth, 'bom.delete'); bomQuiet_();
  var name = ptName_(p.dessert), to = ptName_(p.to);
  if (!to || to === name) throw fail_('請填新的甜點名稱');
  if (to.length > 60 || /[\r\n\t]/.test(to)) throw fail_('新名稱太長或含換行');
  var tab = ptRead_('bom'), blk = bomBlock_(tab, name);
  if (!blk.idx.length) throw fail_('BOM 表找不到「' + name + '」', 'notfound');
  if (bomBlock_(tab, to).idx.length) throw fail_('BOM 表已經有「' + to + '」，不能改成同名', 'dup');
  var cur = bomDigest_(blk.rows);
  if (cur !== str_(p.base_digest)) throw fail_('「' + name + '」剛被別人改過，這次沒有改名，請重新讀取', 'conflict', { digest: cur, rows: blk.rows });
  var rows = blk.rows.map(function (r) { return [to, r[1], Number(r[2]), r[3], r[4]]; }), batch = bomBatch_();
  try { bomWriteBlock_(tab, blk, rows); } catch (e) { bomRestore_(to, blk.rows); throw e; }
  bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: 'bom.rename', table: 'BOM表', key: name + ' → ' + to, source: '改名',
    before_json: JSON.stringify(blk.rows), after_json: JSON.stringify(rows.map(bomRowNorm_)), digest_before: cur, digest_after: bomDigest_(rows.map(bomRowNorm_)), rows_before: blk.rows.length, rows_after: rows.length, ok: 'Y', msg: '' });
  /* 對照表 B 欄跟著改（同一批次記錄，可一起還原） */
  var map = ptRead_('map'), changed = [];
  for (var i = 0; i < map.lastA - 1; i++) {
    if (ptName_(map.vals[i][1]) !== name) continue;
    var rn = i + 2, bv = mapRowVals_(map.vals[i]);
    map.sh.getRange(rn, 2).setNumberFormat('@').setValue(safe_(to));
    var av = bv.slice(); av[1] = to; changed.push(bv[0]);
    bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: 'map.edit', table: '產品名稱對照表', key: bv[0], source: '甜點改名連動', before_json: JSON.stringify(bv), after_json: JSON.stringify(av),
      digest_before: mapRowDigest_(bv), digest_after: mapRowDigest_(av), rows_before: 1, rows_after: 1, ok: 'Y', msg: '第 ' + rn + ' 列 B 欄' });
  }
  SpreadsheetApp.flush();
  return { id: name, data: { batch: batch, to: to, mapChanged: changed }, msg: '改名「' + name + '」→「' + to + '」（BOM ' + rows.length + ' 列；對照表 ' + changed.length + ' 列）' };
}

function mapGet_(p, auth) {
  var c = mapGetCore_(!!p.fresh);
  return { id: '', data: { rows: c.rows, desserts: c.desserts } };
}
function mapCheck_(m, tab, bomNames, auth, isNew, curG) {
  var err = [], a = ptName_(m.a), b = ptName_(m.b), e = normDate_(m.e), f = normDate_(m.f), g = ptName_(m.g), h = ptName_(m.h) || '無';
  if (!a) err.push('POS 資料產品名稱必填'); else if (a.length > 80 || /[\r\n\t]/.test(a)) err.push('POS 產品名稱太長或含換行');
  var re = /^\d{4}-\d{2}-\d{2}$/;
  if ((e && !f) || (!e && f)) err.push('起始有效日、結束有效日要同時填或同時空白');
  if (e && !re.test(e)) err.push('起始有效日格式要 YYYY-MM-DD'); if (f && !re.test(f)) err.push('結束有效日格式要 YYYY-MM-DD');
  if (e && f && re.test(e) && re.test(f) && e > f) err.push('起始有效日不能晚於結束有效日');
  var mains = {}, subs = {};
  for (var i = 0; i < tab.lastA - 1; i++) { mains[ptName_(tab.vals[i][6])] = 1; subs[ptName_(tab.vals[i][7])] = 1; }
  if (!g) err.push('主類別必填');
  else if (!mains[g] && !(auth.role === ADMIN && m.new_main === true)) err.push('主類別「' + g + '」不在現有清單（新增主類別只限管理者，且 POS 儀表板白名單要同步改）');
  if (h !== '無' && !subs[h] && !/^20\d{2} \S/.test(h)) err.push('次類別要是「無」或「YYYY 檔期名」（例：2026 萬聖節，年份後一個空白）');
  if (g === '限定甜點' && (h === '無' || (!/^20\d{2} \S/.test(h) && !subs[h]))) err.push('主類別「限定甜點」的次類別要是檔期名（例：2026 萬聖節），採購系統靠它啟動檔期備料');
  if (b && !bomNames[b]) err.push('手動對應產品名稱「' + b + '」在 BOM 表找不到（要和 BOM 表甜點名稱一模一樣；新甜點請先建 BOM）');
  if (!b && BOM_DESSERT_MAINS[g]) err.push('主類別「' + g + '」是甜點，手動對應產品名稱（BOM 甜點名稱）必填，否則採購系統展不出用料');
  if (err.length) throw fail_('資料沒有通過檢查，沒有寫入：' + err.join('；'), 'invalid', { errors: err });
  return { a: a, b: b, e: e, f: f, g: g, h: h };
}
function mapSave_(p, auth) {
  needB_(auth, 'map.edit'); bomQuiet_();
  var origA = ptName_(p.orig_a), isNew = !origA, tab = ptRead_('map');
  var bomNames = {}; bomDesserts_(ptRead_('bom')).forEach(function (d) { bomNames[d.name] = 1; });
  var cur = null, curRow = 0;
  if (!isNew) {
    for (var i = 0; i < tab.lastA - 1; i++) if (ptName_(tab.vals[i][0]) === origA) { cur = mapRowVals_(tab.vals[i]); curRow = i + 2; break; }
    if (!cur) throw fail_('對照表找不到「' + origA + '」（可能剛被改名或刪除），請重新讀取', 'notfound');
    if (mapRowDigest_(cur) !== str_(p.base)) throw fail_('「' + origA + '」剛被別人改過，這次沒有存，請重新讀取', 'conflict', { row: cur });
  }
  var v = mapCheck_(p.row || {}, tab, bomNames, auth, isNew);
  for (var j = 0; j < tab.lastA - 1; j++) if (ptName_(tab.vals[j][0]) === v.a && (isNew || j + 2 !== curRow)) throw fail_('對照表已經有「' + v.a + '」（第 ' + (j + 2) + ' 列），POS 品名不能重複', 'dup');
  if (!isNew && (cur[4] !== v.g || cur[5] !== v.h)) needB_(auth, 'map.cat');
  var after = [v.a, v.b, v.e, v.f, v.g, v.h], batch = bomBatch_(), rn;
  var rowArr = [v.a, v.b, '', '', ymd_(v.e), ymd_(v.f), v.g, v.h];
  if (isNew) {
    var res = ptAppend_(tab, [rowArr]); rn = res.start;
  } else {
    rn = curRow;
    if (!isNew && mapRowDigest_(after) === mapRowDigest_(cur)) return { id: v.a, data: { unchanged: true, row: rn, digest: mapRowDigest_(cur) }, msg: '沒有變更' };
    var sh = tab.sh;
    sh.getRange(rn, 1, 1, 2).setNumberFormat('@').setValues([[safe_(v.a), safe_(v.b)]]);
    sh.getRange(rn, 5, 1, 2).setValues([[ymd_(v.e), ymd_(v.f)]]);
    sh.getRange(rn, 7, 1, 2).setNumberFormat('@').setValues([[safe_(v.g), safe_(v.h)]]);
    SpreadsheetApp.flush();
    var back = mapRowVals_(sh.getRange(rn, 1, 1, 8).getValues()[0]);
    if (JSON.stringify(back) !== JSON.stringify(after)) {
      sh.getRange(rn, 1, 1, 2).setValues([[safe_(cur[0]), safe_(cur[1])]]); sh.getRange(rn, 5, 1, 2).setValues([[ymd_(cur[2]), ymd_(cur[3])]]); sh.getRange(rn, 7, 1, 2).setValues([[safe_(cur[4]), safe_(cur[5])]]);
      throw fail_('對照表第 ' + rn + ' 列讀回不一致（' + JSON.stringify(back) + '），已寫回原值', 'verify');
    }
  }
  bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: isNew ? 'map.add' : 'map.edit', table: '產品名稱對照表', key: v.a, source: str_(p.source).slice(0, 200) || '手動',
    before_json: isNew ? '' : JSON.stringify(cur), after_json: JSON.stringify(after), digest_before: isNew ? '' : mapRowDigest_(cur), digest_after: mapRowDigest_(after), rows_before: isNew ? 0 : 1, rows_after: 1, ok: 'Y', msg: '第 ' + rn + ' 列' });
  return { id: v.a, data: { batch: batch, row: rn, digest: mapRowDigest_(after), vals: after }, msg: (isNew ? '新增' : '更新') + '對照「' + v.a + '」→「' + (v.b || '（空）') + '」' };
}
function mapDelete_(p, auth) {
  needB_(auth, 'map.delete'); bomQuiet_();
  if (p.confirm_history !== true) throw fail_('刪除對照列會讓這個 POS 品名的全部歷史銷售變成「找不到」分類，請改設結束日；確定要刪請勾選確認', 'confirm');
  var a = ptName_(p.a), tab = ptRead_('map'), cur = null, rn = 0;
  for (var i = 0; i < tab.lastA - 1; i++) if (ptName_(tab.vals[i][0]) === a) { cur = mapRowVals_(tab.vals[i]); rn = i + 2; break; }
  if (!cur) throw fail_('對照表找不到「' + a + '」', 'notfound');
  if (mapRowDigest_(cur) !== str_(p.base)) throw fail_('「' + a + '」剛被別人改過，這次沒有刪除，請重新讀取', 'conflict', { row: cur });
  var batch = bomBatch_();
  tab.sh.deleteRow(rn); SpreadsheetApp.flush();
  bomLogW_({ ts: now_(), batch: batch, role: auth.role, action: 'map.delete', table: '產品名稱對照表', key: a, source: '手動', before_json: JSON.stringify(cur), after_json: '',
    digest_before: mapRowDigest_(cur), digest_after: '', rows_before: 1, rows_after: 0, ok: 'Y', msg: '原第 ' + rn + ' 列' });
  return { id: a, data: { batch: batch }, msg: '刪除對照「' + a + '」' };
}
function bomLog_(p, auth) {
  var t = load_('fact_bom_log'), key = ptName_(p.key), lim = Math.min(Number(p.limit) || 80, 300), out = [];
  for (var i = t.rows.length - 1; i >= 0 && out.length < lim; i--) {
    var r = t.rows[i];
    if (key && str_(r.key).indexOf(key) < 0) continue;
    out.push({ ts: str_(r.ts), batch: str_(r.batch), role: str_(r.role), action: str_(r.action), table: str_(r.table), key: str_(r.key), source: str_(r.source),
      rows_before: r.rows_before, rows_after: r.rows_after, ok: str_(r.ok), msg: str_(r.msg), undoable: str_(r.ok) === 'Y' && str_(r.action) !== 'undo' && str_(r.before_json).charAt(0) !== '（' });
  }
  return { id: key, data: { rows: out } };
}
/* 還原一批：目前內容要等於該批寫入後的內容（指紋相同）才還原，避免蓋掉之後別人的修改 */
function bomUndo_(p, auth) {
  needB_(auth, 'undo'); bomQuiet_();
  var batch = str_(p.batch).trim(), t = load_('fact_bom_log');
  var recs = t.rows.filter(function (r) { return str_(r.batch) === batch && str_(r.ok) === 'Y' && str_(r.action) !== 'undo'; });
  if (!recs.length) throw fail_('找不到可還原的異動批次 ' + batch, 'notfound');
  if (t.rows.some(function (r) { return str_(r.action) === 'undo' && str_(r.source) === batch && str_(r.ok) === 'Y'; })) throw fail_('這一批已經還原過了', 'done');
  var done = [], nb = bomBatch_();
  recs.slice().reverse().forEach(function (r) {
    var before = str_(r.before_json) ? JSON.parse(str_(r.before_json)) : null, after = str_(r.after_json) ? JSON.parse(str_(r.after_json)) : null;
    if (str_(r.table) === 'BOM表') {
      var key = str_(r.action) === 'bom.rename' ? str_(r.key).split(' → ')[1] : str_(r.key), orig = str_(r.action) === 'bom.rename' ? str_(r.key).split(' → ')[0] : key;
      var tab = ptRead_('bom'), blk = bomBlock_(tab, key), curD = bomDigest_(blk.rows);
      if (curD !== str_(r.digest_after)) throw fail_('「' + key + '」在這批之後又被改過（版本不同），不能自動還原；請在 BOM 管理手動改回', 'conflict');
      bomWriteBlock_(tab, blk, (before || []).map(function (x) { return [orig, x[1], Number(x[2]), x[3], x[4]]; }));
      done.push('BOM「' + key + '」→ ' + (before || []).length + ' 列');
    } else {
      var map = ptRead_('map'), a = after ? after[0] : str_(r.key), rn = 0, cur = null;
      for (var i = 0; i < map.lastA - 1; i++) if (ptName_(map.vals[i][0]) === a) { cur = mapRowVals_(map.vals[i]); rn = i + 2; break; }
      if (after && (!cur || mapRowDigest_(cur) !== str_(r.digest_after))) throw fail_('對照「' + a + '」在這批之後又被改過，不能自動還原', 'conflict');
      if (!after && cur) throw fail_('對照「' + a + '」已經又被建立，不能自動還原', 'conflict');
      if (before && after) {
        map.sh.getRange(rn, 1, 1, 2).setValues([[safe_(before[0]), safe_(before[1])]]); map.sh.getRange(rn, 5, 1, 2).setValues([[ymd_(before[2]), ymd_(before[3])]]); map.sh.getRange(rn, 7, 1, 2).setValues([[safe_(before[4]), safe_(before[5])]]);
      } else if (!before && after) { map.sh.deleteRow(rn); }
      else if (before && !after) { ptAppend_(map, [[before[0], before[1], '', '', ymd_(before[2]), ymd_(before[3]), before[4], before[5]]]); }
      SpreadsheetApp.flush();
      done.push('對照「' + a + '」');
    }
  });
  bomLogW_({ ts: now_(), batch: nb, role: auth.role, action: 'undo', table: '', key: done.join('；').slice(0, 300), source: batch, before_json: '', after_json: '', digest_before: '', digest_after: '', rows_before: '', rows_after: '', ok: 'Y', msg: '還原批次 ' + batch });
  return { id: batch, data: { batch: nb, done: done }, msg: '還原 ' + batch + '：' + done.join('；') };
}
ACTIONS.bomMeta = bomMeta_; ACTIONS.bomGet = bomGet_; ACTIONS.bomSave = bomSave_; ACTIONS.bomDelete = bomDelete_; ACTIONS.bomRename = bomRename_;
ACTIONS.mapGet = mapGet_; ACTIONS.mapSave = mapSave_; ACTIONS.mapDelete = mapDelete_; ACTIONS.bomLog = bomLog_; ACTIONS.bomUndo = bomUndo_;
ACTIONS.bomMetaMap = bomMetaMap_;   /* v7 */
WRITES.bomSave = 1; WRITES.bomDelete = 1; WRITES.bomRename = 1; WRITES.mapSave = 1; WRITES.mapDelete = 1; WRITES.bomUndo = 1;

/* ================= 小工具 ================= */
/* CacheService 一個鍵最多 100KB：超過就切塊（每塊 ≤30,000 個字＝UTF-8 最多 90KB；不切在 emoji 代理對中間），主鍵只放「#塊數」 */
var CACHE_CHUNK = 30000, CACHE_MAX_CHUNKS = 40;
function cachePutBig_(cache, key, s, ttl) {
  s = String(s);
  if (s.length <= CACHE_CHUNK) { cache.put(key, s, ttl); return; }
  var parts = {}, n = 0, i = 0;
  while (i < s.length) {
    var end = Math.min(i + CACHE_CHUNK, s.length), hi = s.charCodeAt(end - 1);
    if (end < s.length && hi >= 0xD800 && hi <= 0xDBFF) end--;
    parts[key + ':' + n] = s.slice(i, end); n++; i = end;
    if (n > CACHE_MAX_CHUNKS) return;   /* 太大（>1.2MB）就不暫存 */
  }
  cache.putAll(parts, ttl);
  cache.put(key, '#' + n, ttl);
}
function cacheGetBig_(cache, key) {
  var head = cache.get(key);
  if (head === null || head === undefined) return null;
  if (head.charAt(0) !== '#') return head;
  var n = parseInt(head.slice(1), 10), keys = [];
  for (var i = 0; i < n; i++) keys.push(key + ':' + i);
  var all = cache.getAll(keys), s = '';
  for (i = 0; i < n; i++) { var c = all[key + ':' + i]; if (c === null || c === undefined) return null; s += c; }
  return s;
}
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
