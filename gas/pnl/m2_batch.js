/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — M2 批次協調器
 * 檔案:m2_batch.gs
 * 版本:M2-CODE-A.1.2(2026-07-13)
 * ============================================================
 *
 * 【v1.2 變更摘要 (2026-07-13)】
 * 針對 7/12 雙資料事故(audit: 記錄刪 234、實刪 78,deleteRows 回報成功但未真正刪除):
 * 1. commitStagingToFact overwrite 分支改「過濾重寫」:
 *    整表讀出 → 濾掉重複 key → 清空 → 重寫,徹底不依賴 deleteRows 與列號
 * 2. 重寫前先做 fact_pnl 全表備份(backupFactPnl),重寫中斷可 rollback
 * 3. 重寫後 SpreadsheetApp.flush() + 複驗:待刪 key 殘留必須 = 0,否則硬 abort、不寫新資料
 * 4. 安全網 3 升級:筆數核對失敗 → ok:false(前端顯示失敗),不再默默成功
 * 5. LockService 防並發 commit
 * _mergeRowsToDeleteRanges 保留(測試函式仍引用),但 commit 已不再使用。
 *
 * 【v1.1 變更摘要 (2026-06-09)】
 * 1. 解禁 M2_FIRST_TIME_VERIFICATION_STORES = []
 *    11/12 店已歷經 3 次驗證(5/14、5/20、6/9),不再警示
 * 2. commitStagingToFact overwrite 分支改用 deleteRows 批次刪除
 *    把連續列號合併成區段,一次 API 呼叫,從 212 次→1~3 次
 *    速度從 80 秒→1~2 秒,徹底解決雙資料事故(地雷 14)
 * 3. deleteRows 失敗時硬性 abort,不執行後續 setValues
 *    並把 ROLLBACK 訊息寫入 audit_log
 *
 * 【本檔本階段(CODE-A)目的】
 * 1. 掃描 Drive 資料夾,找出本月鼎新檔
 * 2. 批次解析所有檔案,結果寫入 fact_pnl_staging
 * 3. 執行跨檔對帳(含修補 parser.gs 缺漏的「行銷費用」L1)
 * 4. 產出驗證報告(物件)
 *
 * 【絕對不會做的事】(CODE-A 階段嚴格邊界)
 * - ❌ 不修改 parser.gs(M1 已通過,留作基準)
 * - ❌ 不修改 Code.gs
 * - ❌ 不執行 runRebuildBCD / runImportAll / fixFactPnl / patchDimAccount
 * - ❌ 不部署 Web App
 *
 * 【可以做的事】
 * - ✅ 建立 fact_pnl_staging 分頁(M2 解禁)
 * - ✅ 寫入 fact_pnl_staging
 * - ✅ 自包 reconcileV2(包含行銷費用,不動 parser.gs)
 * ============================================================
 */


// ============================================================
// 【常數區】
// ============================================================

const M2_STAGING_SHEET_NAME = 'fact_pnl_staging';

// 首次匯入特殊保護:這些店號需要單獨驗證
// 2026-06-09: 11、12 店已歷經 5/14 全量重灌、5/20 補件、6/9 修正版三次驗證,解禁
const M2_FIRST_TIME_VERIFICATION_STORES = [];

// 跨檔對帳軟性警示閾值
const M2_GROSS_MARGIN_LOW = 0.50;   // 毛利率 < 50% 警示
const M2_GROSS_MARGIN_HIGH = 0.90;  // 毛利率 > 90% 警示
const M2_MOM_CHANGE_THRESHOLD = 0.30;  // 單月營收變動 ±30% 警示


// ============================================================
// 【函式 1】scanDingxinFolder(folderId)
// 掃描 Drive 資料夾,過濾出符合鼎新檔規範的 Google Sheets
// ============================================================

/**
 * @param {string} folderId  Drive 資料夾 ID
 * @return {object} { ok, files: [{id, name, ymFromName, storeIdFromName}], errors }
 */
function scanDingxinFolder(folderId) {
  const result = { ok: false, files: [], errors: [], skipped: [] };

  try {
    const folder = DriveApp.getFolderById(folderId);
    const iter = folder.getFiles();

    while (iter.hasNext()) {
      const f = iter.next();
      const name = f.getName();
      const id = f.getId();
      const mime = f.getMimeType();

      // 必須是 Google Sheets(M1 v2 已驗證的前提)
      if (mime !== 'application/vnd.google-apps.spreadsheet') {
        result.skipped.push({ id: id, name: name, reason: '非 Google Sheets:' + mime });
        continue;
      }

      // 檔名解析(沿用 parser.gs 的規範:YYMM_SS_損益表 [.xlsx])
      const fnParse = parseFileName(name);
      if (!fnParse.ok) {
        result.skipped.push({ id: id, name: name, reason: fnParse.reason });
        continue;
      }

      result.files.push({
        id: id,
        name: name,
        ymFromName: fnParse.ym,
        storeIdFromName: fnParse.storeId
      });
    }

    result.ok = true;
    return result;

  } catch (e) {
    result.errors.push('掃描資料夾失敗:' + e.toString());
    return result;
  }
}


// ============================================================
// 【函式 2】processBatch(folderId, yearMonth)
// 批次協調器:掃描 → 逐檔解析 → 寫 staging → 跨檔對帳 → 組報告
// ============================================================

/**
 * @param {string} folderId  Drive 資料夾 ID
 * @param {number} [yearMonth]  選填,期望期間(用於檢查所有檔是否同月)
 * @return {object} 完整批次結果
 */
function processBatch(folderId, yearMonth) {
  const batchResult = {
    ok: false,
    folderId: folderId,
    expectedYm: yearMonth || null,
    scan: null,
    fileResults: [],
    allRecords: [],
    crossCheck: null,
    stagingWrite: null,
    errors: [],
    warnings: [],
    isFirstTimeVerification: false
  };

  // Step 1: 掃描資料夾
  const scan = scanDingxinFolder(folderId);
  batchResult.scan = scan;
  if (!scan.ok) {
    batchResult.errors.push('資料夾掃描失敗,無法繼續');
    return batchResult;
  }
  if (scan.files.length === 0) {
    batchResult.errors.push('資料夾內沒有符合規範的檔案,無法繼續');
    return batchResult;
  }

  // Step 2: 檢查是否含首次驗證店號
  // 2026-06-09: 常數已解禁,此分支實際不會觸發,程式碼保留供未來新店首次驗證
  const verificationStoresInBatch = scan.files
    .map(f => f.storeIdFromName)
    .filter(s => M2_FIRST_TIME_VERIFICATION_STORES.indexOf(s) >= 0);
  if (verificationStoresInBatch.length > 0) {
    batchResult.isFirstTimeVerification = true;
    batchResult.warnings.push(
      '⚠️ 本批次含首次驗證店號 ' + verificationStoresInBatch.join(', ') +
      '──部門編碼解析需經營者人工確認,通過前不得 commit'
    );
  }

  // Step 3: 逐檔解析
  scan.files.forEach(function(file) {
    const fileResult = parseDingxinFile(file.id);
    fileResult.expectedYmFromName = file.ymFromName;
    fileResult.expectedStoreFromName = file.storeIdFromName;
    batchResult.fileResults.push(fileResult);

    if (fileResult.ok) {
      // 用 reconcileV2 重新對帳(覆蓋 parser.gs 的 reconcile 結果)
      const v2Checks = reconcileV2(fileResult.records, fileResult.summary.checks);
      fileResult.summary.checksV2 = v2Checks;

      // 把這檔的記錄收進 allRecords
      fileResult.records.forEach(function(rec) { batchResult.allRecords.push(rec); });
    } else {
      batchResult.errors.push(
        '檔案 ' + file.name + ' 解析失敗:' + fileResult.errors.join('; ')
      );
    }
  });

  // 硬性擋下:任何一檔解析失敗 → 整批不能匯入
  if (batchResult.errors.length > 0) {
    batchResult.errors.push('整批擋下:有檔案解析失敗,不允許部分匯入');
    return batchResult;
  }

  // 硬性擋下:檢查 (year_month, store_id) 重複
  const seen = {};
  batchResult.fileResults.forEach(function(fr) {
    if (fr.summary && fr.summary.year_month && fr.summary.store_id !== undefined) {
      const key = fr.summary.year_month + '_' + fr.summary.store_id;
      if (seen[key]) {
        batchResult.errors.push('重複的 (期間, 店) 組合:' + key);
      }
      seen[key] = true;
    }
  });
  if (batchResult.errors.length > 0) {
    return batchResult;
  }

  // Step 4: 跨檔對帳
  batchResult.crossCheck = crossFileReconcile(batchResult.allRecords, batchResult.fileResults);
  batchResult.warnings = batchResult.warnings.concat(batchResult.crossCheck.warnings);

  // Step 5: 寫入 staging
  batchResult.stagingWrite = writeToStaging(batchResult.allRecords);
  if (!batchResult.stagingWrite.ok) {
    batchResult.errors.push('staging 寫入失敗:' + batchResult.stagingWrite.reason);
    return batchResult;
  }

  batchResult.ok = batchResult.errors.length === 0;
  return batchResult;
}


// ============================================================
// 【函式 3】writeToStaging(records)
// 寫入 fact_pnl_staging,如不存在則建立分頁
// ============================================================

function writeToStaging(records) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let staging = ss.getSheetByName(M2_STAGING_SHEET_NAME);

    const factPnl = ss.getSheetByName('fact_pnl');
    if (!factPnl) {
      return { ok: false, reason: '找不到 fact_pnl,無法取得標準表頭' };
    }
    const headers = factPnl.getRange(1, 1, 1, factPnl.getLastColumn()).getValues()[0];

    if (!staging) {
      staging = ss.insertSheet(M2_STAGING_SHEET_NAME);
      staging.getRange(1, 1, 1, headers.length).setValues([headers]);
      ss.setActiveSheet(staging);
      ss.moveActiveSheet(ss.getNumSheets());
    } else {
      const lastRow = staging.getLastRow();
      if (lastRow > 1) {
        staging.getRange(2, 1, lastRow - 1, staging.getLastColumn()).clearContent();
      }
      staging.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    if (records.length === 0) {
      return { ok: true, rowsWritten: 0, sheetName: M2_STAGING_SHEET_NAME };
    }

    const grid = records.map(function(rec) {
      return headers.map(function(h) {
        return rec[h] !== undefined ? rec[h] : '';
      });
    });

    staging.getRange(2, 1, grid.length, headers.length).setValues(grid);

    return {
      ok: true,
      rowsWritten: grid.length,
      sheetName: M2_STAGING_SHEET_NAME
    };

  } catch (e) {
    return { ok: false, reason: '例外:' + e.toString() };
  }
}


// ============================================================
// 【函式 4】reconcileV2(records, originalChecks)
// 修補版對帳(包含行銷費用,不動 parser.gs)
// ============================================================

function reconcileV2(records, originalChecks) {
  const dingxinByName = {};
  if (Array.isArray(originalChecks)) {
    originalChecks.forEach(function(c) {
      dingxinByName[c.name] = c.dingxin;
    });
  }

  const revenue   = sumByL1(records, '營業收入');
  const cost      = sumByL1(records, '直接成本');
  const hr        = sumByL1(records, '人事成本');
  const site      = sumByL1(records, '場地成本');
  const ops       = sumByL1(records, '營運費用');
  const mgmt      = sumByL1(records, '管理費用');
  const marketing = sumByL1(records, '行銷費用');
  const expenseTotal = hr + site + ops + mgmt + marketing;

  const TOL = 1.0;
  function makeCheck(name, calc) {
    const dx = dingxinByName[name];
    if (dx === undefined) {
      return { name: name, calc: calc, dingxin: null, pass: null, note: '鼎新無此合計欄' };
    }
    return {
      name: name,
      calc: calc,
      dingxin: dx,
      pass: Math.abs(calc - dx) < TOL
    };
  }

  return [
    makeCheck('營業收入', revenue),
    makeCheck('直接成本', cost),
    makeCheck('營業毛利', revenue - cost),
    makeCheck('管理及總務費用', expenseTotal),
    makeCheck('營業淨利', revenue - cost - expenseTotal)
  ];
}


// ============================================================
// 【函式 5】crossFileReconcile(allRecords, fileResults)
// 跨檔對帳(軟性警示)
// ============================================================

function crossFileReconcile(allRecords, fileResults) {
  const result = { warnings: [], storeMetrics: [], totals: {} };

  fileResults.forEach(function(fr) {
    if (!fr.ok || !fr.summary) return;
    const s = fr.summary;
    const margin = s.revenue > 0 ? (s.gross_profit / s.revenue) : 0;
    result.storeMetrics.push({
      store_id: s.store_id,
      year_month: s.year_month,
      revenue: s.revenue,
      gross_profit: s.gross_profit,
      gross_margin: margin,
      operating_profit: s.operating_profit
    });

    if (margin < M2_GROSS_MARGIN_LOW) {
      result.warnings.push(
        '⚠️ 店 ' + s.store_id + ' 毛利率異常偏低:' + (margin * 100).toFixed(2) + '%(< 50%)'
      );
    }
    if (margin > M2_GROSS_MARGIN_HIGH) {
      result.warnings.push(
        '⚠️ 店 ' + s.store_id + ' 毛利率異常偏高:' + (margin * 100).toFixed(2) + '%(> 90%)'
      );
    }
  });

  let totalRev = 0, totalCost = 0, totalOpProfit = 0;
  result.storeMetrics.forEach(function(m) {
    totalRev += m.revenue;
    totalCost += m.gross_profit;
    totalOpProfit += m.operating_profit;
  });
  result.totals = {
    storeCount: result.storeMetrics.length,
    totalRevenue: totalRev,
    totalGrossProfit: totalCost,
    totalOpProfit: totalOpProfit
  };

  return result;
}


// ============================================================
// 【函式 6】generateReport(batchResult)
// ============================================================

function generateReport(batchResult) {
  const report = {
    ok: batchResult.ok,
    folderId: batchResult.folderId,
    timestamp: new Date().toISOString(),
    summary: {
      totalFiles: batchResult.scan ? batchResult.scan.files.length : 0,
      successFiles: 0,
      failedFiles: 0,
      skippedFiles: batchResult.scan ? batchResult.scan.skipped.length : 0,
      totalRecords: batchResult.allRecords.length,
      stagingRowsWritten: batchResult.stagingWrite ? batchResult.stagingWrite.rowsWritten : 0
    },
    fileSummaries: [],
    crossCheckWarnings: batchResult.crossCheck ? batchResult.crossCheck.warnings : [],
    crossCheckTotals: batchResult.crossCheck ? batchResult.crossCheck.totals : {},
    isFirstTimeVerification: batchResult.isFirstTimeVerification,
    errors: batchResult.errors,
    warnings: batchResult.warnings
  };

  batchResult.fileResults.forEach(function(fr) {
    if (fr.ok) report.summary.successFiles++;
    else report.summary.failedFiles++;

    const fs = {
      fileName: fr.fileName,
      ok: fr.ok,
      year_month: fr.summary ? fr.summary.year_month : null,
      store_id: fr.summary ? fr.summary.store_id : null,
      records: fr.records ? fr.records.length : 0,
      revenue: fr.summary ? fr.summary.revenue : null,
      operating_profit: fr.summary ? fr.summary.operating_profit : null,
      checksV2: fr.summary ? fr.summary.checksV2 : null,
      errors: fr.errors,
      warnings: fr.warnings
    };
    report.fileSummaries.push(fs);
  });

  return report;
}


// ============================================================
// 【測試函式】testBatchProcess
// ============================================================

function testBatchProcess() {
  const TEST_FOLDER_ID = '1cVq5IISD36I0uC63GI9ZuVATNGVpRqJ1';

  if (TEST_FOLDER_ID === '請填入測試資料夾 ID') {
    Logger.log('⚠️ 請先把 TEST_FOLDER_ID 改成您的測試資料夾 ID');
    return;
  }

  Logger.log('============================================================');
  Logger.log('  M2 CODE-A 批次處理測試');
  Logger.log('  資料夾 ID: ' + TEST_FOLDER_ID);
  Logger.log('  時間: ' + new Date().toISOString());
  Logger.log('============================================================');

  const baselineBefore = checkFactPnlBaseline();
  Logger.log('');
  Logger.log('fact_pnl 基線(前): ' + baselineBefore);
  Logger.log('');

  const batchResult = processBatch(TEST_FOLDER_ID);
  const report = generateReport(batchResult);

  Logger.log('--- 批次結果 ---');
  Logger.log('ok: ' + batchResult.ok);
  Logger.log('總檔案: ' + report.summary.totalFiles);
  Logger.log('成功: ' + report.summary.successFiles);
  Logger.log('失敗: ' + report.summary.failedFiles);
  Logger.log('跳過(不符規範): ' + report.summary.skippedFiles);
  Logger.log('總解析筆數: ' + report.summary.totalRecords);
  Logger.log('staging 寫入筆數: ' + report.summary.stagingRowsWritten);
  Logger.log('首次驗證模式: ' + report.isFirstTimeVerification);
  Logger.log('');

  if (report.crossCheckTotals && report.crossCheckTotals.storeCount !== undefined) {
    Logger.log('--- 跨檔總計 ---');
    Logger.log('店家數: ' + report.crossCheckTotals.storeCount);
    Logger.log('總營收: ' + report.crossCheckTotals.totalRevenue);
    Logger.log('總毛利: ' + report.crossCheckTotals.totalGrossProfit);
    Logger.log('總營業淨利: ' + report.crossCheckTotals.totalOpProfit);
    Logger.log('');
  }

  Logger.log('--- 各檔摘要 ---');
  report.fileSummaries.forEach(function(fs) {
    Logger.log('  ' + (fs.ok ? '✅' : '❌') + ' ' + fs.fileName +
               ' | 期間 ' + fs.year_month + ' / 店 ' + fs.store_id +
               ' | ' + fs.records + ' 筆 | 營收 ' + fs.revenue);
    if (fs.checksV2) {
      fs.checksV2.forEach(function(c) {
        const flag = c.pass === true ? '✅' : (c.pass === false ? '❌' : '➖');
        Logger.log('      ' + flag + ' ' + c.name +
                   '(細項=' + c.calc + ' 鼎新=' + c.dingxin + ')');
      });
    }
  });
  Logger.log('');

  if (report.crossCheckWarnings.length > 0) {
    Logger.log('--- 跨檔警示 ---');
    report.crossCheckWarnings.forEach(function(w) { Logger.log('  ' + w); });
    Logger.log('');
  }

  if (report.errors.length > 0) {
    Logger.log('--- ❌ 錯誤 ---');
    report.errors.forEach(function(e) { Logger.log('  ' + e); });
    Logger.log('');
  }

  if (report.warnings.length > 0) {
    Logger.log('--- ⚠️ 警示 ---');
    report.warnings.forEach(function(w) { Logger.log('  ' + w); });
    Logger.log('');
  }

  const baselineAfter = checkFactPnlBaseline();
  Logger.log('============================================================');
  Logger.log('fact_pnl 基線(前): ' + baselineBefore);
  Logger.log('fact_pnl 基線(後): ' + baselineAfter);
  Logger.log('fact_pnl 是否變動: ' +
             (baselineBefore === baselineAfter ? '✅ 沒變' : '❌ 有變'));
  Logger.log('============================================================');

  return report;
}

// ============================================================
// 【CODE-B】fact_pnl 寫入機制(備份 / commit / rollback / 清理)
// ============================================================

const M2_BACKUP_SHEET_PREFIX = 'fact_pnl_backup_';
const M2_BACKUP_KEEP_COUNT = 10;

function backupFactPnl() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const src = ss.getSheetByName('fact_pnl');
    if (!src) return { ok: false, reason: '找不到 fact_pnl 分頁' };
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    const backupName = M2_BACKUP_SHEET_PREFIX + stamp;
    let finalName = backupName;
    let suffix = 2;
    while (ss.getSheetByName(finalName)) {
      finalName = backupName + '_' + suffix;
      suffix++;
      if (suffix > 99) return { ok: false, reason: '命名衝突異常' };
    }
    const copy = src.copyTo(ss);
    copy.setName(finalName);
    ss.setActiveSheet(copy);
    ss.moveActiveSheet(ss.getNumSheets());
    const rowsBackedUp = copy.getLastRow() - 1;
    Logger.log('✅ 備份完成:' + finalName + '(' + rowsBackedUp + ' 筆)');
    return { ok: true, backupSheetName: finalName, rowsBackedUp: rowsBackedUp };
  } catch (e) {
    return { ok: false, reason: '例外:' + e.toString() };
  }
}


// ============================================================
// 【內部輔助函式 2026-06-09 新增】
// 把已排序(大到小)的列號合併成連續區段,用於 deleteRows 批次刪除
// 例: [10, 9, 8, 5, 4, 2] → [{start: 8, count: 3}, {start: 4, count: 2}, {start: 2, count: 1}]
// 注意: 從大到小回傳,呼叫端要從第一個區段開始刪(才不會影響後面區段的列號)
// 【v1.2 註記】commitStagingToFact 已改過濾重寫,不再使用本函式;
// 保留供 testMergeRowsToDeleteRanges 測試引用,勿刪。
// ============================================================

function _mergeRowsToDeleteRanges(rowsDesc) {
  if (!rowsDesc || rowsDesc.length === 0) return [];
  const ranges = [];
  let i = 0;
  while (i < rowsDesc.length) {
    // 找出從 rowsDesc[i] 起連續往下的列號(因為已 sort 從大到小,連續=值減 1)
    let count = 1;
    while (i + count < rowsDesc.length &&
           rowsDesc[i + count] === rowsDesc[i] - count) {
      count++;
    }
    // 這個區段:start 是這串連續列號裡最小的(也就是 rowsDesc[i + count - 1])
    // count 是區段長度
    ranges.push({
      start: rowsDesc[i + count - 1],
      count: count
    });
    i += count;
  }
  return ranges;
}


// ============================================================
// 【函式 7】commitStagingToFact(options) — v1.2(2026-07-13)
//
// 【v1.2 變更摘要】針對 7/12 雙資料事故(audit: 記錄刪 234、實刪 78)
// 1. overwrite 不再用 deleteRows(回報成功但實際未刪,6/9、7/12 兩次同簽名)
//    改「過濾重寫」:整表讀出 → 濾掉重複 key → 清空 → 重寫,無列號位移問題
// 2. 重寫前先做 fact_pnl 全表備份(backupFactPnl),重寫中斷可 rollback
// 3. 刪除後 SpreadsheetApp.flush() + 複驗:重讀 fact_pnl 確認待刪 key 已 0 筆,否則 abort
// 4. 安全網 3 升級:countMismatch → ok:false(前端顯示失敗),不再默默成功
// 5. LockService 防並發
// ============================================================

function commitStagingToFact(options) {
  const opts = options || { duplicateMode: 'cancel' };
  const result = {
    ok: false, backup: null, duplicates: [],
    duplicateMode: opts.duplicateMode,
    rowsCommitted: 0, rowsSkipped: 0, rowsOverwritten: 0,
    factPnlBefore: 0, factPnlAfter: 0,
    errors: []
  };

  // 【v1.2】防並發:同時只能有一個 commit
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    result.errors.push('另一個 commit 正在執行中,請稍後再試');
    return result;
  }

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const factPnl = ss.getSheetByName('fact_pnl');
    const staging = ss.getSheetByName(M2_STAGING_SHEET_NAME);
    if (!factPnl) { result.errors.push('找不到 fact_pnl 分頁'); return result; }
    if (!staging) { result.errors.push('找不到 staging 分頁'); return result; }
    result.factPnlBefore = factPnl.getLastRow() - 1;

    const headers = factPnl.getRange(1, 1, 1, factPnl.getLastColumn()).getValues()[0];
    const idxYm = headers.indexOf('year_month');
    const idxStore = headers.indexOf('store_id');
    if (idxYm < 0 || idxStore < 0) { result.errors.push('缺欄'); return result; }

    const stagingLastRow = staging.getLastRow();
    if (stagingLastRow < 2) { result.errors.push('staging 沒有資料'); return result; }

    const stagingData = staging.getRange(2, 1, stagingLastRow - 1, headers.length).getValues();

    // 整表讀出 fact_pnl(v1.2:之後過濾重寫都用這份,不再依賴列號)
    const factLastRow = factPnl.getLastRow();
    let factData = [];
    if (factLastRow >= 2) {
      factData = factPnl.getRange(2, 1, factLastRow - 1, headers.length).getValues();
    }

    function keyOf(row) { return row[idxYm] + '_' + row[idxStore]; }

    // 統計 fact / staging 的 (ym, store) 組合
    const factKeyCount = {};
    factData.forEach(function(row) {
      const k = keyOf(row);
      factKeyCount[k] = (factKeyCount[k] || 0) + 1;
    });
    const stagingPairs = {};
    stagingData.forEach(function(row) {
      const k = keyOf(row);
      stagingPairs[k] = (stagingPairs[k] || 0) + 1;
    });
    Object.keys(stagingPairs).forEach(function(k) {
      if (factKeyCount[k]) {
        result.duplicates.push({ key: k, stagingRows: stagingPairs[k], factRows: factKeyCount[k] });
      }
    });

    if (result.duplicates.length > 0) {
      if (opts.duplicateMode === 'cancel') {
        result.errors.push('偵測到 ' + result.duplicates.length + ' 組重複,已取消寫入');
        return result;
      }

      if (opts.duplicateMode === 'overwrite') {
        const keysToDelete = {};
        result.duplicates.forEach(function(dup) { keysToDelete[dup.key] = true; });

        // 【安全網 1】被覆蓋列備份到專屬分頁(直接從 factData 取,免逐列 getRange)
        const rowsToBackup = factData.filter(function(row) { return keysToDelete[keyOf(row)]; });
        const replacedBackup = backupReplacedRows(rowsToBackup, headers);
        if (!replacedBackup.ok) {
          result.errors.push('安全網 1 失敗:' + replacedBackup.error + ',已取消覆蓋');
          return result;
        }
        result.replacedBackupSheet = replacedBackup.sheetName;

        // 【v1.2 安全網 0】重寫前全表備份(重寫若中斷可 rollbackFromBackup)
        const fullBackup = backupFactPnl();
        if (!fullBackup.ok) {
          result.errors.push('全表備份失敗:' + fullBackup.reason + ',已取消覆蓋');
          return result;
        }
        result.backup = fullBackup.backupSheetName;

        // 【v1.2 核心】過濾重寫:保留列 = 不在待刪 key 的列
        const keptRows = factData.filter(function(row) { return !keysToDelete[keyOf(row)]; });
        const expectedDeleteCount = factData.length - keptRows.length;

        factPnl.getRange(2, 1, factData.length, headers.length).clearContent();
        SpreadsheetApp.flush();
        if (keptRows.length > 0) {
          factPnl.getRange(2, 1, keptRows.length, headers.length).setValues(keptRows);
        }
        SpreadsheetApp.flush();

        // 【v1.2 複驗】重讀 fact_pnl,待刪 key 殘留必須 = 0
        const verifyLastRow = factPnl.getLastRow();
        let residue = 0;
        if (verifyLastRow >= 2) {
          const verifyData = factPnl.getRange(2, 1, verifyLastRow - 1, headers.length).getValues();
          verifyData.forEach(function(row) {
            if (keysToDelete[keyOf(row)]) residue++;
          });
          if (verifyLastRow - 1 !== keptRows.length) {
            result.errors.push('重寫後筆數異常:預期 ' + keptRows.length + ',實際 ' + (verifyLastRow - 1));
          }
        }
        if (residue > 0) {
          result.errors.push('複驗失敗:待覆蓋資料仍殘留 ' + residue + ' 列,已中止(未寫入新資料)。' +
                             '可用 rollbackFromBackup(\'' + result.backup + '\') 還原');
          try {
            writeAuditLog({
              timestamp: new Date(),
              action: 'COMMIT_ABORTED_RESIDUE',
              duplicateMode: opts.duplicateMode,
              factPnlBefore: result.factPnlBefore,
              factPnlAfter: factPnl.getLastRow() - 1,
              rowsCommitted: 0,
              rowsOverwritten: expectedDeleteCount - residue,
              rowsSkipped: 0,
              replacedBackupSheet: result.replacedBackupSheet || null,
              countMismatch: true,
              errors: result.errors.join(' / ')
            });
          } catch(_) {}
          return result;
        }
        if (result.errors.length > 0) return result;
        result.rowsOverwritten = expectedDeleteCount;
      }
    }

    // 寫入新資料
    let recordsToWrite;
    if (opts.duplicateMode === 'skip') {
      recordsToWrite = stagingData.filter(function(row) {
        return !factKeyCount[keyOf(row)];
      });
      result.rowsSkipped = stagingData.length - recordsToWrite.length;
    } else {
      recordsToWrite = stagingData;
    }
    if (recordsToWrite.length > 0) {
      const startRow = factPnl.getLastRow() + 1;
      factPnl.getRange(startRow, 1, recordsToWrite.length, headers.length).setValues(recordsToWrite);
      SpreadsheetApp.flush();
      result.rowsCommitted = recordsToWrite.length;
    }
    if (stagingLastRow > 1) {
      staging.getRange(2, 1, stagingLastRow - 1, staging.getLastColumn()).clearContent();
    }
    result.factPnlAfter = factPnl.getLastRow() - 1;

    // 【v1.2 安全網 3 升級】筆數核對失敗 → 硬失敗(ok:false),不再默默成功
    const expectedAfter = result.factPnlBefore - result.rowsOverwritten + result.rowsCommitted;
    if (result.factPnlAfter !== expectedAfter) {
      result.countMismatch = true;
      result.errors.push(
        '安全網 3:筆數核對失敗 ' +
        '(預期 ' + expectedAfter + ' = ' + result.factPnlBefore +
        ' - ' + result.rowsOverwritten + ' + ' + result.rowsCommitted +
        ',實際 ' + result.factPnlAfter + ')。commit 判定失敗,請查 audit_log,' +
        '必要時 rollbackFromBackup(\'' + (result.backup || '最新 fact_pnl_backup') + '\')'
      );
    }

    // 【安全網 2】寫入 audit_log
    writeAuditLog({
      timestamp: new Date(),
      action: 'COMMIT_' + (opts.duplicateMode || 'cancel').toUpperCase(),
      duplicateMode: opts.duplicateMode,
      factPnlBefore: result.factPnlBefore,
      factPnlAfter: result.factPnlAfter,
      rowsCommitted: result.rowsCommitted,
      rowsOverwritten: result.rowsOverwritten,
      rowsSkipped: result.rowsSkipped,
      replacedBackupSheet: result.replacedBackupSheet || null,
      countMismatch: result.countMismatch || false,
      errors: result.errors.length > 0 ? result.errors.join(' / ') : null
    });
    /* 清除 get_yms cache,讓儀表板 F5 後立刻看到新月份 */
    try { CacheService.getScriptCache().remove('get_yms'); } catch(_) {}

    result.ok = !result.countMismatch;   // v1.2:mismatch 即失敗
    return result;
  } catch (e) {
    result.errors.push('commit 例外:' + e.toString());
    return result;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

function rollbackFromBackup(backupSheetName) {
  const result = { ok: false, backupSheetName: backupSheetName, preRollbackBackup: null, rowsRestored: 0, factPnlBefore: 0, factPnlAfter: 0, errors: [] };
  try {
    if (!backupSheetName || backupSheetName.indexOf(M2_BACKUP_SHEET_PREFIX) !== 0) {
      result.errors.push('backupSheetName 必須以 ' + M2_BACKUP_SHEET_PREFIX + ' 開頭');
      return result;
    }
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const backup = ss.getSheetByName(backupSheetName);
    if (!backup) { result.errors.push('備份分頁不存在:' + backupSheetName); return result; }
    const factPnl = ss.getSheetByName('fact_pnl');
    if (!factPnl) { result.errors.push('找不到 fact_pnl 分頁'); return result; }
    result.factPnlBefore = factPnl.getLastRow() - 1;
    const factHeaders = factPnl.getRange(1, 1, 1, factPnl.getLastColumn()).getValues()[0];
    const backupHeaders = backup.getRange(1, 1, 1, backup.getLastColumn()).getValues()[0];
    if (factHeaders.length !== backupHeaders.length) {
      result.errors.push('表頭欄位數不符');
      return result;
    }
    for (let i = 0; i < factHeaders.length; i++) {
      if (String(factHeaders[i]) !== String(backupHeaders[i])) {
        result.errors.push('表頭第 ' + (i + 1) + ' 欄不符');
        return result;
      }
    }
    const preRollback = backupFactPnl();
    if (!preRollback.ok) { result.errors.push('還原前備份失敗'); return result; }
    result.preRollbackBackup = preRollback.backupSheetName;
    const factLastRow = factPnl.getLastRow();
    if (factLastRow > 1) {
      factPnl.getRange(2, 1, factLastRow - 1, factPnl.getLastColumn()).clearContent();
    }
    const backupLastRow = backup.getLastRow();
    if (backupLastRow > 1) {
      const backupData = backup.getRange(2, 1, backupLastRow - 1, backupHeaders.length).getValues();
      factPnl.getRange(2, 1, backupData.length, backupHeaders.length).setValues(backupData);
      result.rowsRestored = backupData.length;
    }
    result.factPnlAfter = factPnl.getLastRow() - 1;
    result.ok = true;
    Logger.log('✅ 已從 ' + backupSheetName + ' 還原 ' + result.rowsRestored + ' 筆');
    return result;
  } catch (e) {
    result.errors.push('rollback 例外:' + e.toString());
    return result;
  }
}

function cleanOldBackups(keepCount) {
  const k = (typeof keepCount === 'number' && keepCount > 0) ? keepCount : M2_BACKUP_KEEP_COUNT;
  const result = { ok: false, kept: [], deleted: [], errors: [] };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const allSheets = ss.getSheets();
    const backups = [];
    allSheets.forEach(function(sheet) {
      const name = sheet.getName();
      if (name.indexOf(M2_BACKUP_SHEET_PREFIX) === 0) backups.push(name);
    });
    backups.sort();
    const toDelete = backups.length > k ? backups.slice(0, backups.length - k) : [];
    const toKeep = backups.length > k ? backups.slice(backups.length - k) : backups;
    toDelete.forEach(function(name) {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        ss.deleteSheet(sheet);
        result.deleted.push(name);
      }
    });
    result.kept = toKeep;
    result.ok = true;
    return result;
  } catch (e) {
    result.errors.push('cleanOldBackups 例外:' + e.toString());
    return result;
  }
}

function testBackupFactPnl() {
  Logger.log('============================================================');
  Logger.log('  M2 CODE-B 備份測試');
  Logger.log('============================================================');
  const baselineBefore = checkFactPnlBaseline();
  Logger.log('fact_pnl 基線(前): ' + baselineBefore);
  const result = backupFactPnl();
  Logger.log('--- 備份結果 ---');
  Logger.log('ok: ' + result.ok);
  Logger.log('備份分頁名: ' + result.backupSheetName);
  Logger.log('備份筆數: ' + result.rowsBackedUp);
  if (result.reason) Logger.log('reason: ' + result.reason);
  const baselineAfter = checkFactPnlBaseline();
  Logger.log('fact_pnl 基線(後): ' + baselineAfter);
  Logger.log('fact_pnl 是否變動: ' + (baselineBefore === baselineAfter ? '✅ 沒變' : '❌ 有變'));
  return result;
}

function tempCommitNow() {
  const result = commitStagingToFact({ duplicateMode: 'cancel' });
  Logger.log(JSON.stringify(result));
}


// ============================================================
// 【2026-06-09 新增測試函式】
// 測試 _mergeRowsToDeleteRanges 邏輯正確性
// 不動 fact_pnl,純算式測試
// ============================================================

function testMergeRowsToDeleteRanges() {
  Logger.log('============================================================');
  Logger.log('  測試 _mergeRowsToDeleteRanges 列號合併邏輯');
  Logger.log('============================================================');

  const cases = [
    {
      name: '全連續(212 列連在一起)',
      input: (function(){ var a=[]; for(var i=10848; i>=10637; i--) a.push(i); return a; })(),
      expected: '1 個區段 start=10637 count=212'
    },
    {
      name: '兩個分開區段',
      input: [10, 9, 8, 5, 4, 2],
      expected: '3 個區段:[start=8 count=3][start=4 count=2][start=2 count=1]'
    },
    {
      name: '全分散',
      input: [100, 50, 10],
      expected: '3 個區段:[start=100 count=1][start=50 count=1][start=10 count=1]'
    },
    {
      name: '空陣列',
      input: [],
      expected: '0 個區段'
    },
    {
      name: '單一列',
      input: [42],
      expected: '1 個區段:[start=42 count=1]'
    }
  ];

  cases.forEach(function(c) {
    const ranges = _mergeRowsToDeleteRanges(c.input);
    Logger.log('');
    Logger.log('▶ ' + c.name);
    Logger.log('  input length: ' + c.input.length);
    Logger.log('  output ranges: ' + ranges.length + ' 個');
    ranges.forEach(function(rg, i) {
      Logger.log('    區段' + (i+1) + ': start=' + rg.start + ' count=' + rg.count);
    });
    // 驗證 total count
    const totalCount = ranges.reduce(function(s, rg) { return s + rg.count; }, 0);
    if (totalCount === c.input.length) {
      Logger.log('  ✅ 總列數吻合: ' + totalCount);
    } else {
      Logger.log('  ❌ 總列數錯誤:input=' + c.input.length + ' 但 ranges 加總=' + totalCount);
    }
  });

  Logger.log('');
  Logger.log('============================================================');
  Logger.log('  測試完成,以上輸出應該都顯示 ✅');
  Logger.log('============================================================');
}