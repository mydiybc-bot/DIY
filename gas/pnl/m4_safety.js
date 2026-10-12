/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — M4.1 安全網模組
 * 檔案:m4_safety.gs(新增)
 * 版本:M4.1.0(2026-04-27)
 * ============================================================
 *
 * 【本檔目的】
 * 實作經營者要求的 3 道安全網:
 *   1. 刪除前精準備份(被覆蓋的列備份到專屬分頁)
 *   2. audit_log 紀錄所有 commit 操作
 *   3. 寫入前後筆數核對(在 commitStagingToFact 結尾)
 *
 * 【dry-run 偵測】
 * 額外提供 checkDuplicates() 給 dashboard 在按確認匯入前
 * 預先偵測 fact_pnl 是否有重複的 (year_month, store_id) 組合
 * ============================================================
 */

const M4_REPLACED_BACKUP_PREFIX = 'fact_pnl_replaced_';
const M4_AUDIT_LOG_SHEET = 'audit_log';
const M4_REPLACED_BACKUP_KEEP_DAYS = 90;


// ============================================================
// 【函式】checkDuplicates — dry-run 偵測重複(不寫入)
// ============================================================
function checkDuplicates() {
  const result = { ok: false, duplicates: [], stagingTotal: 0, errors: [] };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const factPnl = ss.getSheetByName('fact_pnl');
    const staging = ss.getSheetByName(M2_STAGING_SHEET_NAME);
    if (!factPnl) { result.errors.push('找不到 fact_pnl'); return result; }
    if (!staging) { result.errors.push('找不到 staging'); return result; }
    const headers = factPnl.getRange(1, 1, 1, factPnl.getLastColumn()).getValues()[0];
    const idxYm = headers.indexOf('year_month');
    const idxStore = headers.indexOf('store_id');
    if (idxYm < 0 || idxStore < 0) { result.errors.push('缺欄'); return result; }
    const stagingLastRow = staging.getLastRow();
    if (stagingLastRow < 2) {
      result.errors.push('staging 沒有資料');
      return result;
    }
    const stagingData = staging.getRange(2, 1, stagingLastRow - 1, headers.length).getValues();
    result.stagingTotal = stagingData.length;
    const factLastRow = factPnl.getLastRow();
    const factPairs = {};
    if (factLastRow >= 2) {
      const factData = factPnl.getRange(2, 1, factLastRow - 1, headers.length).getValues();
      factData.forEach(function(row, i) {
        const key = row[idxYm] + '_' + row[idxStore];
        if (!factPairs[key]) factPairs[key] = [];
        factPairs[key].push(i + 2);
      });
    }
    const stagingPairs = {};
    stagingData.forEach(function(row) {
      const key = row[idxYm] + '_' + row[idxStore];
      stagingPairs[key] = (stagingPairs[key] || 0) + 1;
    });
    Object.keys(stagingPairs).forEach(function(key) {
      if (factPairs[key]) {
        result.duplicates.push({
          key: key,
          stagingRows: stagingPairs[key],
          factRows: factPairs[key].length
        });
      }
    });
    result.ok = true;
    return result;
  } catch (e) {
    result.errors.push('checkDuplicates 例外:' + e.toString());
    return result;
  }
}


// ============================================================
// 【函式】backupReplacedRows — 安全網 1
// 把要刪除的列備份到專屬分頁
// ============================================================
function backupReplacedRows(rowsData, headers) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const tz = Session.getScriptTimeZone() || 'Asia/Taipei';
    const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
    let sheetName = M4_REPLACED_BACKUP_PREFIX + stamp;
    let suffix = 2;
    while (ss.getSheetByName(sheetName)) {
      sheetName = M4_REPLACED_BACKUP_PREFIX + stamp + '_' + suffix;
      suffix++;
      if (suffix > 99) return { ok: false, error: '命名衝突異常' };
    }
    const sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rowsData.length > 0) {
      sheet.getRange(2, 1, rowsData.length, headers.length).setValues(rowsData);
    }
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(ss.getNumSheets());
    Logger.log('✅ 安全網 1:已備份 ' + rowsData.length + ' 筆到 ' + sheetName);
    return { ok: true, sheetName: sheetName, rowsBackedUp: rowsData.length };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}


// ============================================================
// 【函式】writeAuditLog — 安全網 2
// 把 commit 操作紀錄寫入 audit_log 分頁
// ============================================================
function writeAuditLog(record) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(M4_AUDIT_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(M4_AUDIT_LOG_SHEET);
      const headers = [
        'timestamp', 'action', 'duplicateMode',
        'factPnlBefore', 'factPnlAfter',
        'rowsCommitted', 'rowsOverwritten', 'rowsSkipped',
        'replacedBackupSheet', 'countMismatch', 'errors'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      record.timestamp,
      record.action,
      record.duplicateMode || '',
      record.factPnlBefore,
      record.factPnlAfter,
      record.rowsCommitted || 0,
      record.rowsOverwritten || 0,
      record.rowsSkipped || 0,
      record.replacedBackupSheet || '',
      record.countMismatch || false,
      record.errors || ''
    ]);
    return { ok: true };
  } catch (e) {
    Logger.log('writeAuditLog 例外:' + e.toString());
    return { ok: false, error: e.toString() };
  }
}


// ============================================================
// 【函式】cleanOldReplacedBackups
// 清理超過 90 天的 fact_pnl_replaced_* 分頁
// (由經營者手動執行,不自動 trigger)
// ============================================================
function cleanOldReplacedBackups() {
  const result = { ok: false, deleted: [], kept: [], errors: [] };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = ss.getSheets();
    const now = new Date();
    sheets.forEach(function(sheet) {
      const name = sheet.getName();
      if (name.indexOf(M4_REPLACED_BACKUP_PREFIX) !== 0) return;
      const m = name.match(/^fact_pnl_replaced_(\d{8})_/);
      if (!m) return;
      const dateStr = m[1];
      const yyyy = parseInt(dateStr.substring(0, 4), 10);
      const mm = parseInt(dateStr.substring(4, 6), 10) - 1;
      const dd = parseInt(dateStr.substring(6, 8), 10);
      const sheetDate = new Date(yyyy, mm, dd);
      const daysOld = (now - sheetDate) / (1000 * 60 * 60 * 24);
      if (daysOld > M4_REPLACED_BACKUP_KEEP_DAYS) {
        ss.deleteSheet(sheet);
        result.deleted.push(name);
      } else {
        result.kept.push(name);
      }
    });
    result.ok = true;
    return result;
  } catch (e) {
    result.errors.push(e.toString());
    return result;
  }
}


// ============================================================
// 【測試函式】只跑 dry-run,不寫入
// ============================================================
function testCheckDuplicates() {
  const r = checkDuplicates();
  Logger.log('=== checkDuplicates 結果 ===');
  Logger.log('ok: ' + r.ok);
  Logger.log('stagingTotal: ' + r.stagingTotal);
  Logger.log('duplicates count: ' + (r.duplicates ? r.duplicates.length : 0));
  if (r.duplicates && r.duplicates.length > 0) {
    r.duplicates.forEach(function(d) {
      Logger.log('  - ' + d.key + ': staging=' + d.stagingRows + ', fact=' + d.factRows);
    });
  }
  if (r.errors && r.errors.length > 0) {
    Logger.log('errors:');
    r.errors.forEach(function(e) { Logger.log('  - ' + e); });
  }
  return r;
}
