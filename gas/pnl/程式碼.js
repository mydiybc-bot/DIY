/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — Apps Script 後端
 * 版本:v2.5(2026-08-12  效能修正:fact_pnl 切塊快取 + 觸發器預熱)
 * ============================================================
 *
 * ★★★ 貼上前必做 ★★★
 *   1. 刪除 pnl_cache.gs 與 pnl_diag.gs 這兩個檔案。
 *      它們的函式已全部併入本檔,留著會造成同名函式重複定義,
 *      GAS 不會報錯但行為不可預測(後載入的覆蓋前面的)。
 *   2. 貼上本檔 → Ctrl+S
 *   3. 執行一次 installPnlCacheTrigger()（只需一次,見下方說明）
 *   4. 部署 → 管理部署 → 編輯(鉛筆) → 版本:新版本 → 部署
 *      ★ Ctrl+S 不會更新 Web App,一定要走「新版本」
 *
 * ============================================================
 * 【v2.5 改動】效能修正
 * ------------------------------------------------------------
 * 實測數據(perf_probe / testPnlCache,2026-08-12):
 *   openById()                                    204 ms
 *   getRange(2,1,12542,12).getValues()          1,509 ms
 *   走快取讀同一份資料                             107 ms   (21.5 倍)
 *   fact_pnl 序列化後 640 KB → CacheService 單鍵上限 100 KB → 必須切塊
 *
 * 問題:儀表板每次載入呼叫 get_fact_data 三次(本期/上期/去年同期),
 *       而 getFactData 是唯一沒有快取的 API,等於同一份資料被完整讀三遍。
 *
 * 對策(三點,都以「最壞情況不劣於改版前」為前提):
 *   (1) getFactData 改走 pnlReadFactRaw_():優先讀快取,沒有就照原本讀 Sheet。
 *   (2) ★ 使用者請求「只讀快取、絕不寫快取」。
 *       寫入由 warmPnlCache() 負責,掛在 5 分鐘時間觸發器背景執行。
 *       理由:寫快取要多握一份 640KB 字串 + 切塊,把這個成本留在使用者請求裡,
 *             會讓最慢的那一發更慢。背景執行沒有這個顧慮。
 *   (3) 只讀前 10 欄(getFactData 實際只用 row[0]~row[9]),
 *       不碰 load_date,避免 Date 物件經 JSON 來回失真。
 *
 * 資料新鮮度(三道保險,不依賴人工):
 *   · 快取鍵綁 lastRow → 筆數一變,舊快取立刻失效
 *   · admin_commit / admin_rollback 成功後,路由層自動清快取
 *   · TTL 1 小時封頂;手動可執行 clearCache()
 *
 * 【v2.4 歷史】新增 get_ledger action,查 fact_ledger 總分類帳明細
 * 【v2.3 歷史】新增 get_yms action,回傳可用月份清單(D1a/D1b)
 * 【v2.2 歷史】
 * 移除 2026-04 期間由自動化工具加入的 4 個危險函式:
 *   - runRebuildBCD()    ← 擅自改寫 dim_account
 *   - runImportAll()     ← 硬編碼的檔案匯入,會把 fact_pnl 變 7 欄
 *   - fixFactPnl()       ← 強制重設 fact_pnl 為 7 欄(元凶)
 *   - patchDimAccount()  ← 自作主張的補丁
 * 這些函式曾造成 fact_pnl 從 5896 筆縮減為 4190 筆的災難,
 * 已於 2026-04-23 藉由備份還原。移除後徹底消除誤執行風險。
 *
 * 【紅線】本版未更動任何寫入邏輯
 *   commitStagingToFact / rollbackFromBackup / writeFactBatch 等一律原樣,
 *   只在它們「回傳之後」補一次清快取,不介入其內部流程與筆數核對。
 * ============================================================
 */

// ============================================================
// 【常數區】— 如需修改系統行為,通常改這裡即可
// ============================================================

const SHEET_ID = '1khrFp_AYp3mEsL02cTRrYu1LNiRTP68aowM5_ClsE_c';
const SHEET_NAMES = {
  FACT: 'fact_pnl',
  STORE: 'dim_store',
  ACCOUNT: 'dim_account',
  LOG: 'import_log',
  CONFIG: 'config'
};

// 快取設定(減少重複讀 Sheet,加速 API 回應)
const CACHE_DURATION_SEC = 600;  // 10 分鐘 — 給 stores / accounts / config / yms


// ============================================================
// 【v2.5 新增】fact_pnl 切塊快取
// ------------------------------------------------------------
// CacheService 單鍵上限約 100KB,fact_pnl 序列化約 640KB,故切塊存放:
//   pnl_raw_v2_<lastRow>_0 .. _N   各 90KB
//   pnl_raw_v2_idx_<lastRow>       記錄塊數
// 讀取時任何一步不完整(過期、被驅逐、塊數不符、筆數不符)一律當作 miss,
// 回頭讀 Sheet。最壞情況等於改版前,不會給出錯誤資料。
// ============================================================

function pnlCacheCfg_() {
  return {
    PREFIX: 'pnl_raw_v2',
    CHUNK_BYTES: 90000,   // 單塊 90KB,留 10KB 餘裕
    MAX_CHUNKS: 30,       // 上限約 2.7MB,超過就不快取(寧可慢,不要存半套)
    TTL_SEC: 3600,        // 1 小時封頂;實際靠觸發器每 5 分鐘刷新
    COLS: 10,             // getFactData 只用 row[0]~row[9]
    MAX_ROWS: 60000       // 保險絲
  };
}

/**
 * 取得 fact_pnl 原始資料 —— ★ 只讀快取,絕不在使用者請求中寫入
 * 回傳值與 sheet.getRange(2,1,lastRow-1,10).getValues() 完全等價
 */
function pnlReadFactRaw_(sheet, lastRow) {
  var C = pnlCacheCfg_();
  var rowCount = lastRow - 1;
  if (rowCount < 1) return [];

  try {
    var cache = CacheService.getScriptCache();
    var stamp = String(lastRow);
    var n = Number(cache.get(C.PREFIX + '_idx_' + stamp) || 0);
    if (n > 0 && n <= C.MAX_CHUNKS) {
      var keys = [];
      for (var i = 0; i < n; i++) keys.push(C.PREFIX + '_' + stamp + '_' + i);
      var got = cache.getAll(keys);
      var buf = '', complete = true;
      for (var j = 0; j < n; j++) {
        var part = got[keys[j]];
        if (part === null || part === undefined) { complete = false; break; }
        buf += part;
      }
      if (complete) {
        var parsed = JSON.parse(buf);
        if (parsed && parsed.length === rowCount) return parsed;   // 筆數必須完全吻合
      }
    }
  } catch (e) {
    Logger.log('[pnl_cache] 快取讀取失敗,改讀 Sheet:' + e);
  }

  // 快取沒有 → 照原本方式讀(不寫快取,寫入是 warmPnlCache 的工作)
  return sheet.getRange(2, 1, rowCount, C.COLS).getValues();
}

/**
 * ★ 預熱快取 —— 掛時間觸發器用,或手動執行
 * 這是唯一會「寫入」快取的地方,刻意隔離在使用者請求之外。
 */
function warmPnlCache() {
  var C = pnlCacheCfg_();
  var t0 = Date.now();
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.FACT);
  if (!sheet) { Logger.log('[pnl_cache] 找不到 fact_pnl'); return 'fact_pnl not found'; }

  var lastRow = sheet.getLastRow();
  var rowCount = lastRow - 1;
  if (rowCount < 1) { Logger.log('[pnl_cache] fact_pnl 無資料'); return 'no data'; }
  if (rowCount > C.MAX_ROWS) {
    Logger.log('[pnl_cache] 筆數 ' + rowCount + ' 超過保險絲 ' + C.MAX_ROWS + ',不快取');
    return 'too many rows';
  }

  var raw = sheet.getRange(2, 1, rowCount, C.COLS).getValues();
  var s = JSON.stringify(raw);
  var n = Math.ceil(s.length / C.CHUNK_BYTES);
  if (n < 1 || n > C.MAX_CHUNKS) {
    Logger.log('[pnl_cache] 需 ' + n + ' 塊,超過上限 ' + C.MAX_CHUNKS + ',放棄快取');
    return 'too large';
  }

  var cache = CacheService.getScriptCache();
  var stamp = String(lastRow);
  var obj = {};
  for (var i = 0; i < n; i++) {
    obj[C.PREFIX + '_' + stamp + '_' + i] = s.substr(i * C.CHUNK_BYTES, C.CHUNK_BYTES);
  }
  cache.putAll(obj, C.TTL_SEC);
  // 索引鍵最後才寫:避免塊還沒寫完就被讀到「看起來完整」的索引
  cache.put(C.PREFIX + '_idx_' + stamp, String(n), C.TTL_SEC);

  var msg = '[pnl_cache] 預熱完成 ' + rowCount + ' 列 / ' + s.length + ' bytes / ' +
            n + ' 塊 / ' + (Date.now() - t0) + ' ms';
  Logger.log(msg);
  return msg;
}

/**
 * 安裝 5 分鐘預熱觸發器 —— 只需執行一次
 * 重複執行不會產生重複觸發器(會先清掉舊的)
 */
function installPnlCacheTrigger() {
  removePnlCacheTrigger();
  ScriptApp.newTrigger('warmPnlCache').timeBased().everyMinutes(5).create();
  warmPnlCache();   // 立刻先暖一次,不必等 5 分鐘
  Logger.log('✓ 已安裝 warmPnlCache 每 5 分鐘觸發器,並已先預熱一次');
  return '✓ 已安裝預熱觸發器';
}

/** 移除預熱觸發器 */
function removePnlCacheTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'warmPnlCache') {
      ScriptApp.deleteTrigger(ts[i]);
      removed++;
    }
  }
  Logger.log('已移除 ' + removed + ' 個 warmPnlCache 觸發器');
  return removed;
}

/** 清掉 fact_pnl 切塊快取(給 clearCache 與 commit 後呼叫) */
function pnlClearFactCache_() {
  try {
    var C = pnlCacheCfg_();
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.FACT);
    var stamp = String(sh ? sh.getLastRow() : 0);
    var keys = [C.PREFIX + '_idx_' + stamp];
    for (var i = 0; i < C.MAX_CHUNKS; i++) keys.push(C.PREFIX + '_' + stamp + '_' + i);
    CacheService.getScriptCache().removeAll(keys);
  } catch (e) {
    Logger.log('[pnl_cache] 清快取失敗(不影響資料):' + e);
  }
}


// ============================================================
// 【初始化】— 第一次部署時執行一次即可
// 建立 config 分頁並寫入預設密碼
// ============================================================

function initSystem() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    Logger.log('已建立 config 分頁');
  }

  const existingData = configSheet.getDataRange().getValues();
  if (existingData.length <= 1) {
    configSheet.clear();
    configSheet.getRange('A1:C1').setValues([['key', 'value', '說明']]);
    configSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');

    const defaultConfig = [
      ['dashboard_password', 'DIYBC2026', '儀表板登入密碼,在此修改即可'],
      ['system_title', '烘焙 DIY 損益分析系統', '儀表板標題'],
      ['company_name', '共好平台有限公司', '公司名稱'],
      ['primary_color', '#F5A623', '主色系(黃色系)'],
      ['accent_color', '#D4841C', '輔色'],
      ['last_updated', new Date().toISOString(), '系統最後更新時間(自動更新)']
    ];

    configSheet.getRange(2, 1, defaultConfig.length, 3).setValues(defaultConfig);
    configSheet.setColumnWidth(1, 200);
    configSheet.setColumnWidth(2, 300);
    configSheet.setColumnWidth(3, 400);
    configSheet.setFrozenRows(1);

    Logger.log('已寫入預設設定');
    Logger.log('預設密碼:DIYBC2026(請在 config 分頁修改)');
  }

  const cache = CacheService.getScriptCache();
  cache.removeAll(['fact_pnl_data', 'dim_store_data', 'dim_account_data', 'config_data']);

  return '✓ 系統初始化完成。預設密碼:DIYBC2026,請立即在 config 分頁修改。';
}


// ============================================================
// 【Web App 入口】— 處理 GET 請求,返回 HTML 頁面或 API 回應
// ============================================================

function doGet(e) {
  // 若有 action 參數,當作 API 呼叫
  if (e && e.parameter && e.parameter.action) {
    // JSONP support for CORS bypass
    var cb = e.parameter.callback;
    var result = handleApiRequest(e.parameter);
    if (cb) {
      var jsonText = result.getContent();
      return ContentService.createTextOutput(cb + "(" + jsonText + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return result;
  }

  // 否則返回測試頁
  return HtmlService.createTemplateFromFile('TestPage')
    .evaluate()
    .setTitle('烘焙 DIY 損益分析系統 — 後端測試')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ============================================================
// 【API 路由】— 依 action 參數呼叫對應函式
// ============================================================

function handleApiRequest(params) {
  const action = params.action;
  let result;

  try {
    switch (action) {
      // --- 讀取 API ---
      case 'verify_password':
        result = verifyPassword(params.password);
        break;

      case 'get_config':
        result = getPublicConfig();
        break;

      case 'get_stores':
        result = getStores();
        break;

      case 'get_accounts':
        result = getAccounts();
        break;

      case 'get_fact_summary':
        result = getFactSummary();
        break;

      case 'get_fact_data':
        result = getFactData(params);
        break;

      case 'ping':
        result = { ok: true, time: new Date().toISOString() };
        break;

      // --- 【v2.3】取得所有可用月份 ---
      case 'get_yms':
        result = getAvailableYms();
        break;

      // --- 【v2.4】總分類帳明細(對應 ledger.gs) ---
      case 'get_ledger':
        result = getLedger_(params);
        break;

      // --- 寫入 API(未來半自動匯入會用到)---
      case 'backup_fact_pnl':
        result = backupFactPnl();
        break;

      case 'clear_sheet':
        result = clearSheet(params.sheet, parseInt(params.keep_header || '1'));
        break;

      case 'write_dim_account':
        result = writeDimAccount(params.data);
        break;

      case 'write_fact_batch':
        result = writeFactBatch(params.data);
        break;

      case 'write_import_log':
        result = writeImportLog(params.data);
        break;

      case 'get_sheet_count':
        result = getSheetCount(params.sheet);
        break;

      // ============================================================
      // 【M3】管理員 API endpoints(對應 m3_admin.gs / m2_batch.gs)
      // ============================================================
      case 'admin_verify':
        result = verifyAdminPassword(params.password);
        break;
      case 'admin_scan_folder':
        result = scanDingxinFolder(params.folderId);
        break;
      case 'admin_parse_batch':
        result = processBatch(params.folderId, params.yearMonth);
        break;
      case 'admin_get_staging_report':
        result = getStagingReport();
        break;
              // 【2026-09-11】會計原始檔直匯(raw_import.gs)
      case 'admin_raw_import':
        if (!verifyAdminToken(params.adminToken)) {
          result = { ok: false, error: '管理員驗證失敗' };
          break;
        }
        result = rawImportRun(params.folderId, params.yms, String(params.dryRun) === 'true');
        break;
      // 【M4.1】dry-run 偵測重複(不寫入)
      case 'admin_check_duplicates':
        if (!verifyAdminToken(params.adminToken)) {
          result = { ok: false, error: '管理員驗證失敗' };
          break;
        }
        result = checkDuplicates();
        break;
      case 'admin_commit':
        // 【M4.1】強制保證 duplicateMode 一定有值,避免 undefined 直接寫入
        var commitOpts = {
          duplicateMode: (params && params.duplicateMode) || 'cancel'
        };
        result = commitStagingToFact(commitOpts);
        // 【v2.5】寫入完成後清快取,避免儀表板讀到舊資料。
        // 注意:只在 commitStagingToFact 回傳之後執行,不介入其內部流程。
        clearCache();
        break;
      case 'admin_rollback':
        result = rollbackFromBackup(params.backupName);
        clearCache();   // 【v2.5】同上
        break;
      case 'admin_list_backups':
        result = listBackupSheets();
        break;
      // 【M3.5】config 讀寫
      case 'admin_get_config':
        result = getConfigValue(params);
        break;
      case 'admin_set_config':
        result = setConfigValue(params);
        break;
      default:
        result = { error: '未知的 action:' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: err.toString(),
        stack: err.stack
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// 【密碼驗證】
// ============================================================

function verifyPassword(inputPassword) {
  if (!inputPassword) {
    return { ok: false, message: '請輸入密碼' };
  }

  const config = getConfigInternal();
  const storedPassword = config.dashboard_password;

  if (inputPassword === storedPassword) {
    return {
      ok: true,
      token: generateToken(),
      message: '登入成功'
    };
  } else {
    Logger.log('密碼錯誤嘗試:' + new Date().toISOString());
    return { ok: false, message: '密碼錯誤' };
  }
}

function generateToken() {
  return Utilities.getUuid();
}


// ============================================================
// 【設定讀取】
// ============================================================

function getConfigInternal() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('config_data');
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);

  if (!sheet) {
    throw new Error('config 分頁不存在,請先執行 initSystem()');
  }

  const data = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < data.length; i++) {
    const [key, value] = data[i];
    if (key) config[key] = value;
  }

  cache.put('config_data', JSON.stringify(config), 60);
  return config;
}

function getPublicConfig() {
  const config = getConfigInternal();
  const publicConfig = Object.assign({}, config);
  delete publicConfig.dashboard_password;
  return publicConfig;
}


// ============================================================
// 【店別主檔讀取】
// ============================================================

function getStores() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dim_store_data');
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.STORE);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return { stores: [] };

  const headers = data[0];
  const stores = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const store = {};
    headers.forEach((h, idx) => {
      store[h] = row[idx];
    });
    stores.push(store);
  }

  const result = { stores, count: stores.length };
  cache.put('dim_store_data', JSON.stringify(result), CACHE_DURATION_SEC);
  return result;
}


// ============================================================
// 【科目主檔讀取】
// ============================================================

function getAccounts() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('dim_account_data');
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  const data = sheet.getDataRange().getValues();

  if (data.length < 2) return { accounts: [] };

  const headers = data[0];
  const accounts = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const acc = {};
    headers.forEach((h, idx) => {
      acc[h] = row[idx];
    });
    accounts.push(acc);
  }

  const result = { accounts, count: accounts.length };
  cache.put('dim_account_data', JSON.stringify(result), CACHE_DURATION_SEC);
  return result;
}


// ============================================================
// 【fact_pnl 摘要】— 診斷用途,不讀全部資料
// ============================================================

function getFactSummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FACT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      total_rows: 0,
      message: 'fact_pnl 無資料'
    };
  }

  // 【v2.5】改走快取路徑(等價於原本的 getRange(...).getValues())
  const data = pnlReadFactRaw_(sheet, lastRow);

  const summary = {
    total_rows: data.length,
    by_year_month: {},
    by_store: {},
    by_category: {},
    amount_stats: {
      min: Infinity,
      max: -Infinity,
      negative_count: 0,
      zero_count: 0
    }
  };

  data.forEach(row => {
    const ym = row[0];         // year_month
    const storeId = row[4];    // store_id
    const cat = row[7];        // category_l1
    const amount = Number(row[9]) || 0;  // amount

    summary.by_year_month[ym] = (summary.by_year_month[ym] || 0) + 1;
    summary.by_store[storeId] = (summary.by_store[storeId] || 0) + 1;
    summary.by_category[cat] = (summary.by_category[cat] || 0) + 1;

    if (amount < summary.amount_stats.min) summary.amount_stats.min = amount;
    if (amount > summary.amount_stats.max) summary.amount_stats.max = amount;
    if (amount < 0) summary.amount_stats.negative_count++;
    if (amount === 0) summary.amount_stats.zero_count++;
  });

  return summary;
}


// ============================================================
// 【fact_pnl 資料查詢】— 依條件篩選
// 這是儀表板主要的資料來源
// ============================================================

function getFactData(params) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FACT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return { data: [] };

  // 【v2.5】原本每次都跑 sheet.getRange(2,1,lastRow-1,12).getValues()(實測 1,509 ms),
  //         且一次載入會呼叫三次。改走切塊快取:命中約 107 ms,沒命中則行為與原本相同。
  const raw = pnlReadFactRaw_(sheet, lastRow);

  const storeFilter = params.store_ids ? params.store_ids.split(',').map(s => parseInt(s)) : null;
  const yearMonthMin = params.ym_min ? parseInt(params.ym_min) : null;
  const yearMonthMax = params.ym_max ? parseInt(params.ym_max) : null;
  const categoryFilter = params.categories ? params.categories.split(',') : null;

  const result = [];

  raw.forEach(row => {
    const ym = row[0];
    const storeId = row[4];
    const cat = row[7];

    if (storeFilter && !storeFilter.includes(storeId)) return;
    if (yearMonthMin && ym < yearMonthMin) return;
    if (yearMonthMax && ym > yearMonthMax) return;
    if (categoryFilter && !categoryFilter.includes(cat)) return;

    result.push({
      year_month: row[0],
      year_ad: row[1],
      quarter: row[2],
      half_year: row[3],
      store_id: row[4],
      account_code: row[5],
      account_name: row[6],
      category_l1: row[7],
      category_l2: row[8],
      amount: row[9]
    });
  });

  return {
    data: result,
    count: result.length,
    filters: { storeFilter, yearMonthMin, yearMonthMax, categoryFilter }
  };
}


// ============================================================
// 【取得所有可用月份】— v2.3 W2 D1a + D1b 用
// 回傳 fact_pnl 中所有 distinct year_month,以及最新月份
// 用途:
//   - D1a:儀表板顯示「資料更新到 YYYY/MM」標籤
//   - D1b:月份下拉動態化(取代 hardcoded 113~115 全 36 個月)
// ============================================================

function getAvailableYms() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('available_yms');
  if (cached) return JSON.parse(cached);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.FACT);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { ok: true, yms: [], latest_ym: null, earliest_ym: null, count: 0 };
  }

  // 只讀第 1 欄 year_month(整份 12 欄不必讀,加快速度)
  const raw = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  const ymSet = new Set();
  raw.forEach(row => {
    const ym = row[0];
    if (ym !== null && ym !== '' && !isNaN(ym)) {
      ymSet.add(Number(ym));
    }
  });

  const yms = Array.from(ymSet).sort((a, b) => a - b);
  const result = {
    ok: true,
    yms: yms,
    latest_ym: yms.length ? yms[yms.length - 1] : null,
    earliest_ym: yms.length ? yms[0] : null,
    count: yms.length
  };

  cache.put('available_yms', JSON.stringify(result), CACHE_DURATION_SEC);
  return result;
}


// ============================================================
// 【總分類帳明細查詢】— v2.4
// 查 fact_ledger 該科目明細,供損益明細表點科目展開
// 註:getLedger_ 與 accountNameToLedgerCodes_ 實作在 ledger.gs
//    (GAS 同專案共用命名空間,此處只是路由呼叫,不需重複定義)
// ============================================================


// ============================================================
// 【寫入 API】— 未來半自動匯入會用到
// 這些是安全的工具 API,只處理指定分頁
// ============================================================

/**
 * 備份 fact_pnl 為 fact_pnl_backup_YYYYMMDD 分頁
 * 匯入新資料前強烈建議呼叫這個
 */
function backupFactPnl() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var src = ss.getSheetByName(SHEET_NAMES.FACT);
  if (!src) return {ok: false, message: 'fact_pnl not found'};

  var rows = src.getLastRow();
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
  var backupName = 'fact_pnl_backup_' + dateStr;

  // 若同名備份已存在,加上時間戳
  if (ss.getSheetByName(backupName)) {
    var timeStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'HHmmss');
    backupName = backupName + '_' + timeStr;
  }

  var newSheet = src.copyTo(ss);
  newSheet.setName(backupName);

  return {ok: true, rows: rows - 1, sheet: backupName};
}

/**
 * 清空指定分頁的資料(保留標題列)
 * 白名單限定:只允許 fact_pnl / dim_account / import_log
 */
function clearSheet(sheetName, keepHeader) {
  var SAFE = ['fact_pnl', 'dim_account', 'import_log'];
  if (SAFE.indexOf(sheetName) === -1) {
    return {ok: false, message: '不允許清空此分頁:' + sheetName};
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return {ok: false, message: sheetName + ' not found'};

  var last = sh.getLastRow();
  if (last <= keepHeader) return {ok: true, cleared: 0};

  sh.deleteRows(keepHeader + 1, last - keepHeader);
  pnlClearFactCache_();   // 【v2.5】資料變動 → 清快取
  return {ok: true, cleared: last - keepHeader};
}

/**
 * 寫入 dim_account(整份覆蓋)
 * data 為 JSON 字串,格式:[[header row], [data row 1], [data row 2], ...]
 */
function writeDimAccount(dataJson) {
  var rows = JSON.parse(dataJson);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  if (!sh) return {ok: false, message: 'dim_account not found'};

  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  CacheService.getScriptCache().remove('dim_account_data');   // 【v2.5】
  return {ok: true, written: rows.length};
}

/**
 * 批次寫入 fact_pnl(附加到最後一列)
 * data 為 JSON 字串,格式:[[row1], [row2], ...]
 */
function writeFactBatch(dataJson) {
  var rows = JSON.parse(dataJson);
  if (!rows.length) return {ok: true, written: 0};

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAMES.FACT);
  if (!sh) return {ok: false, message: 'fact_pnl not found'};

  var last = sh.getLastRow();
  sh.getRange(last + 1, 1, rows.length, rows[0].length).setValues(rows);

  pnlClearFactCache_();   // 【v2.5】資料變動 → 清快取
  return {ok: true, written: rows.length, totalRows: sh.getLastRow() - 1};
}

/**
 * 寫入 import_log(附加到最後一列)
 * data 為 JSON 字串,格式:[row] 或 [[row1],[row2]]
 */
function writeImportLog(dataJson) {
  var row = JSON.parse(dataJson);
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAMES.LOG);
  if (!sh) return {ok: false, message: 'import_log not found'};

  var last = sh.getLastRow();
  if (Array.isArray(row[0])) {
    sh.getRange(last + 1, 1, row.length, row[0].length).setValues(row);
  } else {
    sh.getRange(last + 1, 1, 1, row.length).setValues([row]);
  }

  return {ok: true, written: Array.isArray(row[0]) ? row.length : 1};
}

/**
 * 取得分頁筆數(診斷用)
 */
function getSheetCount(sheetName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return {ok: false, message: sheetName + ' not found'};

  return {
    ok: true,
    sheet: sheetName,
    rows: sh.getLastRow() - 1,
    cols: sh.getLastColumn()
  };
}


// ============================================================
// 【工具函式】— 手動觸發、測試、清除快取
// ============================================================

/**
 * 清除所有快取 — 手動執行
 * 當資料倉更新後,清快取讓 API 重新讀取
 * 【v2.5】一併清掉 fact_pnl 切塊快取
 */
function clearCache() {
  const cache = CacheService.getScriptCache();
  const keys = ['fact_data', 'stores', 'accounts', 'config', 'fact_summary',
                'fact_pnl_data', 'dim_store_data', 'dim_account_data', 'config_data',
                'available_yms'];
  cache.removeAll(keys);
  pnlClearFactCache_();
  Logger.log('✓ 快取已清除(含 fact_pnl 切塊),下次 API 呼叫會重新讀取資料');
  return '✓ 快取已清除';
}

/**
 * 【v2.5】效能診斷 — 在編輯器直接執行,看執行記錄
 * 把瀏覽器那層拿掉,直接量伺服器端:讀取耗時、回應大小、快取是否命中
 */
function diagPnl() {
  var out = [];
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAMES.FACT);
  var lastRow = sh.getLastRow();
  out.push('fact_pnl ' + (lastRow - 1) + ' 列');

  // 快取狀態
  var C = pnlCacheCfg_();
  var idx = CacheService.getScriptCache().get(C.PREFIX + '_idx_' + lastRow);
  out.push('快取狀態:' + (idx ? '已預熱(' + idx + ' 塊)' : '★ 未預熱 — 請執行 installPnlCacheTrigger()'));

  // 模擬儀表板預設查詢:近 12 個月 × 全部店別
  var p = { store_ids: '1,2,3,4,5,6,7,8,9,10,11,12', ym_min: '11408', ym_max: '11507' };
  var t = Date.now();
  var r = getFactData(p);
  var ms = Date.now() - t;
  var size = JSON.stringify(r).length;
  out.push('');
  out.push('getFactData(近12月/全店)  ' + ms + ' ms  → ' + r.data.length + ' 列');
  out.push('JSONP 回應大小 ' + size + ' bytes (' + (size / 1024 / 1024).toFixed(2) + ' MB)');
  if (size > 5 * 1024 * 1024) {
    out.push('★ 回應超過 5MB — 即使伺服器很快,瀏覽器端傳輸與解析仍會很慢,');
    out.push('  下一步要改的是回傳格式(省掉可由 account_code 推導的欄位),不是快取。');
  }
  Logger.log(out.join('\n'));
  return out.join('\n');
}

/**
 * 測試用:在 Apps Script 編輯器直接執行看結果
 */
function testGetStores() {
  const result = getStores();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function testGetFactSummary() {
  const result = getFactSummary();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function testVerifyPassword() {
  const result = verifyPassword('DIYBC2026');
  Logger.log(JSON.stringify(result));
  return result;
}