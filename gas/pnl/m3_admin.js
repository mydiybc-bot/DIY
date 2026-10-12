/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — M3 管理員後端
 * 檔案:m3_admin.gs(新增)
 * 版本:M3.1.0(2026-04-26)
 * ============================================================
 *
 * 【本檔目的】
 * 1. 管理員密碼驗證(從 config 分頁讀 admin_password)
 * 2. staging 報告產生(供儀表板顯示)
 * 3. 備份分頁清單(供還原下拉選單使用)
 *
 * 【絕對不會做的事】
 * - 不寫入 fact_pnl
 * - 不執行 commitStagingToFact / rollbackFromBackup
 * - 不修改既有檔案
 * ============================================================
 */


/**
 * 驗證管理員密碼:從 config 分頁讀 admin_password 比對
 * @param {string} input  使用者輸入的密碼
 * @return {object} { ok: boolean, reason?: string }
 */
function verifyAdminPassword(input) {
  if (!input || typeof input !== 'string') {
    return { ok: false, reason: 'empty_input' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('config');
    if (!sheet) return { ok: false, reason: 'config_not_found' };
    const data = sheet.getDataRange().getValues();
    // 預期 config 結構:第一欄是 key,第二欄是 value
    let stored = null;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'dashboard_password') {
        stored = String(data[i][1]).trim();
        break;
      }
    }
    if (stored === null) return { ok: false, reason: 'admin_password_not_set' };
    if (stored === '') return { ok: false, reason: 'admin_password_empty' };
    return { ok: stored === input.trim() };
  } catch (e) {
    return { ok: false, reason: 'exception:' + e.toString() };
  }
}


/**
 * 產生 staging 摘要報告
 * @return {object} { ok, totalRecords, byStore: [{ store_id, year_month, records, revenue, gross_profit, op_profit, gross_margin }],
 *                    duplicates: [{ key, factCount }],
 *                    factPnlBaseline }
 */
function getStagingReport() {
  const result = {
    ok: false,
    totalRecords: 0,
    byStore: [],
    duplicates: [],
    factPnlBaseline: 0,
    errors: []
  };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const staging = ss.getSheetByName('fact_pnl_staging');
    const factPnl = ss.getSheetByName('fact_pnl');

    if (!factPnl) { result.errors.push('找不到 fact_pnl'); return result; }
    result.factPnlBaseline = factPnl.getLastRow() - 1;

    if (!staging) { result.errors.push('staging 分頁不存在(請先解析)'); return result; }

    const lastRow = staging.getLastRow();
    if (lastRow < 2) { result.ok = true; return result; }  // 空 staging 也是合法

    const headers = staging.getRange(1, 1, 1, staging.getLastColumn()).getValues()[0];
    const idxYm = headers.indexOf('year_month');
    const idxStore = headers.indexOf('store_id');
    const idxL1 = headers.indexOf('category_l1');
    const idxAmount = headers.indexOf('amount');

    if (idxYm < 0 || idxStore < 0 || idxL1 < 0 || idxAmount < 0) {
      result.errors.push('staging 缺欄');
      return result;
    }

    const data = staging.getRange(2, 1, lastRow - 1, headers.length).getValues();
    result.totalRecords = data.length;

    // 依 (year_month, store_id) 分組
    const groups = {};
    data.forEach(function(row) {
      const key = row[idxYm] + '_' + row[idxStore];
      if (!groups[key]) {
        groups[key] = { year_month: row[idxYm], store_id: row[idxStore], records: 0, revenue: 0, cost: 0, hr: 0, site: 0, ops: 0, mgmt: 0, marketing: 0 };
      }
      const g = groups[key];
      g.records++;
      const amt = Number(row[idxAmount]) || 0;
      const l1 = String(row[idxL1] || '');
      if (l1 === '營業收入') g.revenue += amt;
      else if (l1 === '直接成本') g.cost += amt;
      else if (l1 === '人事成本') g.hr += amt;
      else if (l1 === '場地成本') g.site += amt;
      else if (l1 === '營運費用') g.ops += amt;
      else if (l1 === '管理費用') g.mgmt += amt;
      else if (l1 === '行銷費用') g.marketing += amt;
    });

    // 計算每店指標
    Object.keys(groups).forEach(function(k) {
      const g = groups[k];
      const grossProfit = g.revenue - g.cost;
      const expenseTotal = g.hr + g.site + g.ops + g.mgmt + g.marketing;
      const opProfit = grossProfit - expenseTotal;
      result.byStore.push({
        key: k,
        year_month: g.year_month,
        store_id: g.store_id,
        records: g.records,
        revenue: g.revenue,
        cost: g.cost,
        gross_profit: grossProfit,
        gross_margin: g.revenue > 0 ? (grossProfit / g.revenue) : 0,
        operating_profit: opProfit,
        op_margin: g.revenue > 0 ? (opProfit / g.revenue) : 0
      });
    });

    // 重複偵測:讀 fact_pnl 找重複 (ym, store) pair
    const factHeaders = factPnl.getRange(1, 1, 1, factPnl.getLastColumn()).getValues()[0];
    const fIdxYm = factHeaders.indexOf('year_month');
    const fIdxStore = factHeaders.indexOf('store_id');
    if (fIdxYm >= 0 && fIdxStore >= 0 && factPnl.getLastRow() >= 2) {
      const factData = factPnl.getRange(2, 1, factPnl.getLastRow() - 1, factHeaders.length).getValues();
      const factPairs = {};
      factData.forEach(function(row) {
        const key = row[fIdxYm] + '_' + row[fIdxStore];
        factPairs[key] = (factPairs[key] || 0) + 1;
      });
      result.byStore.forEach(function(s) {
        if (factPairs[s.key]) {
          result.duplicates.push({ key: s.key, factCount: factPairs[s.key], stagingRecords: s.records });
        }
      });
    }

    result.ok = true;
    return result;
  } catch (e) {
    result.errors.push('例外:' + e.toString());
    return result;
  }
}


/**
 * 列出所有 fact_pnl_backup_* 分頁,依時間倒序(最新在前)
 * @return {object} { ok, backups: [{ name, rows, createdAt }] }
 */
function listBackupSheets() {
  const result = { ok: false, backups: [], errors: [] };
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheets = ss.getSheets();
    const backups = [];
    sheets.forEach(function(s) {
      const name = s.getName();
      if (name.indexOf('fact_pnl_backup_') === 0) {
        // 從名稱抽出時間戳:fact_pnl_backup_YYYYMMDD_HHMMSS
        const m = name.match(/fact_pnl_backup_(\d{8})_(\d{6})/);
        let createdAt = null;
        if (m) {
          const d = m[1], t = m[2];
          createdAt = d.slice(0, 4) + '-' + d.slice(4, 6) + '-' + d.slice(6, 8) + ' ' +
                      t.slice(0, 2) + ':' + t.slice(2, 4) + ':' + t.slice(4, 6);
        }
        backups.push({
          name: name,
          rows: s.getLastRow() - 1,
          createdAt: createdAt
        });
      }
    });
    // 依名稱倒序(命名格式 = 時間序)
    backups.sort(function(a, b) { return a.name < b.name ? 1 : -1; });
    result.backups = backups;
    result.ok = true;
    return result;
  } catch (e) {
    result.errors.push('例外:' + e.toString());
    return result;
  }
}


// ============================================================
// 【M3.5 新增】config 讀寫 + adminToken 安全保護
// ============================================================
//
// 安全設計(三道防護):
//   1. adminToken 驗證:呼叫端必須在 params 帶 adminToken
//      (值 = config.dashboard_password,等同管理員密碼)
//   2. 寫入黑名單:不允許寫入 admin_password / dashboard_password / password
//   3. 寫入白名單:目前只允許 default_drive_folder_id
//
// ============================================================

const M35_WRITE_BLACKLIST = ['admin_password', 'dashboard_password', 'password'];
const M35_WRITE_WHITELIST = ['default_drive_folder_id'];


/**
 * 驗證 adminToken(用 dashboard_password 比對)
 * @param {string} token  呼叫端傳來的 token
 * @return {boolean}
 */
function verifyAdminToken(token) {
  if (!token || typeof token !== 'string') return false;
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('config');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === 'dashboard_password') {
        const stored = String(data[i][1]).trim();
        if (!stored) return false;
        return stored === token.trim();
      }
    }
    return false;
  } catch (e) { return false; }
}


/**
 * 從 config 分頁讀一個鍵的值
 * @param {object} params  { key: string, adminToken: string }
 * @return {object} { ok, value, error? }
 */
function getConfigValue(params) {
  const p = params || {};
  if (!verifyAdminToken(p.adminToken)) {
    return { ok: false, error: 'unauthorized' };
  }
  if (!p.key || typeof p.key !== 'string') {
    return { ok: false, error: 'key_missing' };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('config');
    if (!sheet) return { ok: false, error: 'config 分頁不存在' };
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === p.key) {
        return { ok: true, value: data[i][1] };
      }
    }
    return { ok: true, value: null };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}


/**
 * 設定 config 分頁的鍵值
 * @param {object} params  { key: string, value: any, adminToken: string }
 * @return {object} { ok, action?: 'updated'|'inserted', row?, error? }
 */
function setConfigValue(params) {
  const p = params || {};
  if (!verifyAdminToken(p.adminToken)) {
    return { ok: false, error: 'unauthorized' };
  }
  if (!p.key || typeof p.key !== 'string') {
    return { ok: false, error: 'key_missing' };
  }
  // 黑名單
  if (M35_WRITE_BLACKLIST.indexOf(p.key) >= 0) {
    return { ok: false, error: 'forbidden_key:' + p.key };
  }
  // 白名單
  if (M35_WRITE_WHITELIST.indexOf(p.key) < 0) {
    return { ok: false, error: 'not_in_whitelist:' + p.key };
  }
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('config');
    if (!sheet) return { ok: false, error: 'config 分頁不存在' };
    const data = sheet.getDataRange().getValues();
    // 找現有 key 更新
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === p.key) {
        sheet.getRange(i + 1, 2).setValue(p.value);
        return { ok: true, action: 'updated', row: i + 1 };
      }
    }
    // 找不到 → append 新列
    const newRow = data.length + 1;
    sheet.getRange(newRow, 1).setValue(p.key);
    sheet.getRange(newRow, 2).setValue(p.value);
    return { ok: true, action: 'inserted', row: newRow };
  } catch (e) {
    return { ok: false, error: e.toString() };
  }
}
