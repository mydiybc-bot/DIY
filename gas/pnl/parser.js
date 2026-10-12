/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — 鼎新解析模組 v3.1
 * 檔案:parser.gs
 * 版本:v3.1(2026-06-09)
 * ============================================================
 *
 * 【v3.1 變更摘要 (2026-06-09)】
 * Step 5.5 新增「自動補 0」邏輯:
 *   - 對照 dim_account 39 筆科目
 *   - 把這份檔案沒出現的科目自動補上 amount=0
 *   - 解決會計風格漂移問題(有時剃 0 元、有時保留導致跨期不一致)
 *   - 每店每月固定 39 筆完整科目,儀表板跨期比較不再有「忽有忽無」
 *
 * 注意事項:
 *   - 補的 0 元列 source_file 標記為原檔名(跟其他列一致)
 *   - load_date 跟其他列一致
 *   - 補的列同樣經過 dim_account 對應(account_code/L1/L2)
 *   - 補 0 邏輯只對「鼎新原始科目名」存在的 dim_account 項目生效
 *
 * 【v3 變更摘要】
 * 支援「整理後格式」:
 *   - A欄 = 科目名稱
 *   - B欄 = 金額
 *   - 第1行 = 標題列(項目 / 店名),跳過
 *   - 第2行起 = 資料
 *
 * 期間(year_month)和店號(store_id)改從「檔名」讀取,不再依賴 R3/R4 欄位。
 * 檔名格式:YYMM_SS_損益表(例:11504_01_損益表)
 *
 * 【本模組只做一件事】
 * 把會計整理好的 Google Sheets 翻譯成 fact_pnl 結構。
 * 純翻譯,不做任何寫入、不做任何修改、不做任何聰明判斷。
 * (v3.1 例外:補 0 是「補齊」不是「修改」,維持每月科目齊全)
 *
 * 【絕對不會做的事】
 * - 不寫入 fact_pnl
 * - 不修改 dim_account
 * - 不在錯誤時自作主張「修補」資料
 * ============================================================
 */


// ============================================================
// 【常數】
// ============================================================

// Google Sheets 的 mimeType
const MIMETYPE_GOOGLE_SHEET = 'application/vnd.google-apps.spreadsheet';

// 對帳容差(避免浮點誤差)
const RECONCILE_TOLERANCE = 1.0;


// ============================================================
// 【主入口】解析單個整理後格式的 Google Sheets 檔案
// ============================================================

/**
 * 解析一個整理後格式的 Google Sheets 檔案,回傳結構化 JSON
 *
 * @param {string} fileId  Drive 上的 Google Sheets 檔案 ID
 * @return {object} 解析結果(records / warnings / errors / summary)
 */
function parseDingxinFile(fileId) {
  const result = {
    ok: false,
    fileId: fileId,
    fileName: '',
    records: [],
    warnings: [],
    errors: [],
    summary: {}
  };

  try {
    // ----- Step 1:取得檔案 metadata,驗證是否為 Google Sheets -----
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    const mimeType = file.getMimeType();
    result.fileName = fileName;

    if (mimeType !== MIMETYPE_GOOGLE_SHEET) {
      result.errors.push(
        '檔案「' + fileName + '」不是 Google Sheets 格式(實際 mimeType: ' + mimeType + ')。\n' +
        '請確認 Drive 是否啟用「上傳的 Excel 檔自動轉換為 Google 試算表」設定。'
      );
      return result;
    }

    // ----- Step 2:從檔名解析 year_month 和 store_id -----
    const fnParse = parseFileName(fileName);
    if (!fnParse.ok) {
      result.errors.push('檔名「' + fileName + '」解析失敗:' + fnParse.reason
        + '\n請改為 YYMM_SS_損益表 格式(例:11504_01_損益表)');
      return result;
    }

    const contentYm = fnParse.ym;
    const contentStoreId = fnParse.storeId;

    // ----- Step 3:直接打開 Sheets -----
    const ss = SpreadsheetApp.openById(fileId);
    let ws = ss.getSheets()[0]; // 取第一個分頁
    result.warnings.push('使用第一個分頁:' + ws.getName());

    const data = ws.getDataRange().getValues();
    const lastRow = data.length;

    if (lastRow < 2) {
      result.errors.push('檔案內容過少(只有 ' + lastRow + ' 列),無法解析');
      return result;
    }

    // ----- Step 4:載入 dim_account 對照表 -----
    const dimAccount = loadDimAccountMap();
    if (Object.keys(dimAccount).length === 0) {
      result.errors.push('dim_account 對照表為空,無法解析');
      return result;
    }

    // ----- Step 5:逐列掃描資料區(從第2行起,跳過標題行)-----
    const records = [];
    const subtotals = {}; // 存合計列金額,用於對帳
    const seenAccountNames = {}; // 2026-06-09: 追蹤已出現的科目,用於 Step 5.5 自動補 0

    // 時間維度(預先計算,Step 5 跟 Step 5.5 共用)
    const roc = Math.floor(contentYm / 100);
    const month = contentYm % 100;
    const yearAd = roc + 1911;
    const quarter = Math.floor((month - 1) / 3) + 1;
    const halfYear = month <= 6 ? 1 : 2;
    const loadDateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

    for (let r = 2; r <= lastRow; r++) {
      const aVal = getCell(data, r, 1); // A欄:科目名稱
      const bVal = getCell(data, r, 2); // B欄:金額

      if (!aVal) continue; // 空列跳過

      const accountName = String(aVal).trim();
      if (!accountName) continue;

      // 合計/小計列:記錄金額供對帳用,不寫進 records
      if (accountName.indexOf('合計') >= 0 ||
          accountName.indexOf('毛利') >= 0 ||
          accountName.indexOf('淨利') >= 0 ||
          accountName.indexOf('損益') >= 0 ||
          accountName.indexOf('總計') >= 0) {
        const totalNum = toNumber(bVal);
        if (totalNum !== null) subtotals[normalizeKey(accountName)] = totalNum;
        continue;
      }

      // 細項科目列
      if (bVal === null || bVal === '') continue;
      const amount = toNumber(bVal);
      if (amount === null) {
        result.warnings.push('R' + r + ' 「' + accountName + '」金額非數字:「' + bVal + '」,跳過');
        continue;
      }

      // 查 dim_account
      const dim = dimAccount[accountName];
      if (!dim) {
        result.errors.push('R' + r + ' 科目「' + accountName
          + '」在 dim_account 找不到對應(請確認科目名稱或補上 dim_account)');
        continue;
      }

      records.push({
        year_month: contentYm,
        year_ad: yearAd,
        quarter: quarter,
        half_year: halfYear,
        store_id: contentStoreId,
        account_code: dim.account_code,
        account_name: accountName,
        category_l1: dim.category_l1,
        category_l2: dim.category_l2,
        amount: amount,
        source_file: fileName,
        load_date: loadDateStr
      });

      // 2026-06-09: 標記此科目已出現,用於 Step 5.5 判斷哪些要補 0
      seenAccountNames[accountName] = true;
    }

    // ============================================================
    // 【Step 5.5】2026-06-09 新增:自動補 0
    // 對照 dim_account 39 筆,把此檔沒出現的科目補上 amount=0
    // 目的:解決會計風格漂移(有時剃 0 元、有時保留)導致的跨期不一致
    // ============================================================
    let autoFilledCount = 0;
    Object.keys(dimAccount).forEach(function(accountName) {
      if (seenAccountNames[accountName]) return; // 此科目本檔已出現,不補
      const dim = dimAccount[accountName];
      records.push({
        year_month: contentYm,
        year_ad: yearAd,
        quarter: quarter,
        half_year: halfYear,
        store_id: contentStoreId,
        account_code: dim.account_code,
        account_name: accountName,
        category_l1: dim.category_l1,
        category_l2: dim.category_l2,
        amount: 0,
        source_file: fileName,
        load_date: loadDateStr
      });
      autoFilledCount++;
    });
    if (autoFilledCount > 0) {
      result.warnings.push('Step 5.5 自動補 0:此檔未出現 ' + autoFilledCount +
                           ' 個科目,已補 amount=0(維持每店每月科目齊全)');
    }

    // ----- Step 6:對帳檢查 -----
    const checks = reconcile(records, subtotals);
    const failedChecks = checks.filter(function(c) { return !c.pass; });
    failedChecks.forEach(function(c) {
      result.warnings.push('對帳差異:' + c.name
        + ' 細項加總=' + c.calc + ',表內合計=' + c.dingxin
        + '(差異=' + (c.calc - c.dingxin) + ')');
    });

    // ----- Step 7:組裝 summary -----
    const revenueSum   = sumByL1(records, '營業收入');
    const costSum      = sumByL1(records, '直接成本');
    const grossProfit  = revenueSum - costSum;
    const hrSum        = sumByL1(records, '人事成本');
    const siteSum      = sumByL1(records, '場地成本');
    const opsSum       = sumByL1(records, '營運費用');
    const mgmtSum      = sumByL1(records, '管理費用');
    const opProfit     = grossProfit - hrSum - siteSum - opsSum - mgmtSum;

    result.summary = {
      total_records:    records.length,
      revenue:          revenueSum,
      cost_of_goods:    costSum,
      gross_profit:     grossProfit,
      gross_margin_pct: revenueSum ? Math.round(grossProfit / revenueSum * 10000) / 100 : 0,
      operating_profit: opProfit,
      op_margin_pct:    revenueSum ? Math.round(opProfit / revenueSum * 10000) / 100 : 0,
      year_month:       contentYm,
      store_id:         contentStoreId,
      checks:           checks,
      autoFilledCount:  autoFilledCount  // 2026-06-09: 補 0 的科目數
    };

    result.records = records;
    result.ok = result.errors.length === 0;
    return result;

  } catch (e) {
    result.errors.push('解析過程發生例外:' + e.toString()
      + (e.stack ? '\nStack:' + e.stack : ''));
    return result;
  }
}


// ============================================================
// 【輔助函式】
// ============================================================

/**
 * 從檔名解析年月與店號
 * 格式:YYMM_SS_損益表 或 YYMM_SS_損益表.xlsx
 */
function parseFileName(fileName) {
  // Drive 轉檔後可能沒有副檔名,兩種都接受
  const m = fileName.match(/^(\d{5})_(\d{2})_損益表(\.xlsx)?$/);
  if (!m) {
    return { ok: false, reason: '不符 YYMM_SS_損益表 格式(例:11504_01_損益表)' };
  }
  return {
    ok: true,
    ym: parseInt(m[1], 10),
    storeId: parseInt(m[2], 10)
  };
}

/**
 * 從二維陣列取單一儲存格(1-indexed)
 */
function getCell(data, row, col) {
  if (row < 1 || row > data.length) return null;
  const r = data[row - 1];
  if (!r || col < 1 || col > r.length) return null;
  return r[col - 1];
}

/**
 * 把任意值轉成數字,失敗回傳 null
 */
function toNumber(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return isNaN(v) ? null : v;
  const cleaned = String(v).replace(/,/g, '').trim();
  if (cleaned === '') return null;
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

/**
 * 標準化 key:移除冒號、全形冒號、空白
 */
function normalizeKey(s) {
  return String(s).replace(/[\u003A\uFF1A\s]/g, '');
}

/**
 * 載入 dim_account,回傳 { 科目名: { account_code, category_l1, category_l2 } }
 */
function loadDimAccountMap() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return {};

  const headers = data[0];
  const idxCode = headers.indexOf('account_code');
  const idxName = headers.indexOf('鼎新原始科目名');
  const idxL1   = headers.indexOf('大類 L1');
  const idxL2   = headers.indexOf('子類 L2');
  if (idxCode < 0 || idxName < 0 || idxL1 < 0 || idxL2 < 0) return {};

  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = row[idxName];
    if (!name) continue;
    map[String(name).trim()] = {
      account_code: row[idxCode],
      category_l1:  row[idxL1],
      category_l2:  row[idxL2]
    };
  }
  return map;
}

/**
 * 加總某 L1 大類的金額
 */
function sumByL1(records, l1) {
  let s = 0;
  for (let i = 0; i < records.length; i++) {
    if (records[i].category_l1 === l1) s += records[i].amount;
  }
  return s;
}

/**
 * 對帳:細項加總 vs 表內合計列
 */
function reconcile(records, subtotals) {
  const checks = [];
  const revenue      = sumByL1(records, '營業收入');
  const cost         = sumByL1(records, '直接成本');
  const hr           = sumByL1(records, '人事成本');
  const site         = sumByL1(records, '場地成本');
  const ops          = sumByL1(records, '營運費用');
  const mgmt         = sumByL1(records, '管理費用');
  const expenseTotal = hr + site + ops + mgmt;

  function addCheck(name, calc, key) {
    const normKey = normalizeKey(key);
    if (subtotals.hasOwnProperty(normKey)) {
      const dx = subtotals[normKey];
      checks.push({
        name:   name,
        calc:   calc,
        dingxin: dx,
        pass:   Math.abs(calc - dx) < RECONCILE_TOLERANCE
      });
    }
  }

  addCheck('營業收入',     revenue,                   '營業收入合計');
  addCheck('直接成本',     cost,                      '進貨合計');
  addCheck('營業毛利',     revenue - cost,            '營業毛利');
  addCheck('管理及總務費用', expenseTotal,             '管理及總務費用合計');
  addCheck('營業淨利',     revenue - cost - expenseTotal, '營業淨利');

  return checks;
}


// ============================================================
// 【測試函式】— 在 Apps Script 編輯器手動執行
// ============================================================

/**
 * 測試:解析指定的 Google Sheets 檔案
 */
function testParseDingxin() {
  const TEST_FILE_ID = '1rIovvZ337esPntYPA1IWgH1A3uRAAOLcg4JD0PSUSuM';

  if (TEST_FILE_ID === '請填入您的檔案 ID') {
    Logger.log('⚠️ 請先把 TEST_FILE_ID 改成您的檔案 ID');
    return;
  }

  const result = parseDingxinFile(TEST_FILE_ID);

  Logger.log('========================================');
  Logger.log('  解析結果(parser v3.1)');
  Logger.log('========================================');
  Logger.log('檔案:' + result.fileName);
  Logger.log('狀態:' + (result.ok ? '✅ 成功' : '❌ 失敗'));

  if (result.summary && result.summary.total_records) {
    const s = result.summary;
    Logger.log('期間/店:' + s.year_month + ' / 店' + s.store_id);
    Logger.log('解析筆數:' + s.total_records + '(含自動補 0 ' + (s.autoFilledCount || 0) + ' 筆)');
    Logger.log('營業收入:' + s.revenue);
    Logger.log('直接成本:' + s.cost_of_goods);
    Logger.log('營業毛利:' + s.gross_profit + '(' + s.gross_margin_pct + '%)');
    Logger.log('營業淨利:' + s.operating_profit + '(' + s.op_margin_pct + '%)');
    Logger.log('');
    Logger.log('對帳結果:');
    s.checks.forEach(function(c) {
      Logger.log('  ' + (c.pass ? '✅' : '⚠️') + ' ' + c.name
        + ' | 細項=' + c.calc + ' | 表內=' + c.dingxin);
    });
  }

  if (result.warnings.length > 0) {
    Logger.log('');
    Logger.log('⚠️ 警示(' + result.warnings.length + '):');
    result.warnings.forEach(function(w) { Logger.log('  - ' + w); });
  }

  if (result.errors.length > 0) {
    Logger.log('');
    Logger.log('❌ 錯誤(' + result.errors.length + '):');
    result.errors.forEach(function(e) { Logger.log('  - ' + e); });
  }

  Logger.log('');
  Logger.log('解析出的細項(' + result.records.length + '筆):');
  result.records.forEach(function(rec) {
    const note = rec.amount === 0 ? ' [補 0]' : '';
    Logger.log('  ' + rec.account_code + ' ' + rec.account_name
      + ' | ' + rec.category_l1 + ' | ' + rec.amount + note);
  });

  return result;
}

/**
 * 測試:檔名解析
 */
function testParseFileName() {
  const cases = [
    '11504_01_損益表',
    '11504_01_損益表.xlsx',
    '11504_12_損益表',
    '11504_1_損益表',     // 錯誤:店號沒補零
    '11504_01_損益表.xls' // 錯誤:副檔名
  ];
  cases.forEach(function(c) {
    const r = parseFileName(c);
    Logger.log(c + ' → ' + (r.ok ? 'YM=' + r.ym + ' Store=' + r.storeId : '❌ ' + r.reason));
  });
}

/**
 * 測試:dim_account 載入
 */
function testLoadDimAccount() {
  const map = loadDimAccountMap();
  Logger.log('載入科目數:' + Object.keys(map).length);
  Logger.log('範例:銷貨收入 → ' + JSON.stringify(map['銷貨收入']));
}