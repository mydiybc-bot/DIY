/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — M2 PRE 階段函式
 * 檔案:m2_pre.gs(新增)
 * 版本:M2-PRE.1.0(2026-04-26)
 * ============================================================
 *
 * 【本檔案只做 PRE 階段的事】
 * 1. 備份 dim_account 為 dim_account_backup_20260426
 * 2. 用 39 筆 v2 替換現有 30 筆 dim_account
 * 3. 查 fact_pnl 基線筆數
 * 4. 統一執行入口 pre1_runAll()
 *
 * 【絕對不會做的事】
 * - 不寫入 fact_pnl(只查筆數)
 * - 不修改 Code.gs / parser.gs
 * - 不執行災難函式
 * ============================================================
 */


// ============================================================
// 【39 筆 dim_account v2 完整資料】
// 直接寫成 JavaScript 二維陣列,避免 base64 / TSV 解析的複雜度
// ============================================================

const DIM_ACCOUNT_V2_DATA = [
  ['account_code', '鼎新原始科目名', '大類 L1', '子類 L2', '正負號', '費用性質', '店主管可控', '備註'],
  ['A001', '銷貨收入', '營業收入', '銷貨收入', '收入', '變動', '直接', '主要營收來源'],
  ['B001', '進貨食材', '直接成本', '食材成本', '支出', '變動', '直接', '烘焙原料採購'],
  ['B002', '進貨耗材', '直接成本', '耗材成本', '支出', '變動', '直接', '包材、一次性用品(可能有退貨負值)'],
  ['C001', '薪資-員工薪資', '人事成本', '薪資', '支出', '固定', '直接', '正職員工月薪'],
  ['C002', '薪資-獎金', '人事成本', '薪資', '支出', '半變動', '直接', '績效獎金、加給'],
  ['C003', '薪資-實作費用', '人事成本', '薪資', '支出', '變動', '直接', 'DIY 實作師資費'],
  ['C004', '薪資-加班費', '人事成本', '薪資', '支出', '變動', '直接', '加班時數計算'],
  ['C010', '薪資-資遣費', '人事成本', '薪資', '支出', '變動', '間接', '【v2 新增】離職資遣金'],
  ['C005', '人險-勞保', '人事成本', '勞健退', '支出', '固定', '間接', '法定,與薪資連動'],
  ['C006', '人險-健保', '人事成本', '勞健退', '支出', '固定', '間接', '法定,與薪資連動'],
  ['C007', '人險-勞退', '人事成本', '勞健退', '支出', '固定', '間接', '法定,與薪資連動'],
  ['C008', '職工福利', '人事成本', '職工相關', '支出', '半變動', '直接', '聚餐、禮品等(可能有沖銷負值)'],
  ['C009', '職業訓練/考照費', '人事成本', '職工相關', '支出', '半變動', '直接', '教育訓練投資'],
  ['D001', '租金費用', '場地成本', '租金', '支出', '固定', '不可控', '合約期內鎖定(可能有押金返還負值)'],
  ['D002', '水電費', '場地成本', '水電', '支出', '半變動', '直接', '含瓦斯,隨營業量變動'],
  ['D003', '清潔費', '場地成本', '清潔修繕', '支出', '固定', '間接', '清潔外包、耗材'],
  ['D004', '修繕費', '場地成本', '清潔修繕', '支出', '變動', '間接', '設備維修,偶發性'],
  ['E001', '耗材費用', '營運費用', '營運耗材', '支出', '變動', '直接', '非直接入菜的營運耗材'],
  ['E002', '運費', '營運費用', '物流通訊', '支出', '變動', '直接', '進貨運費(可能有沖銷負值)'],
  ['E003', '郵電費', '營運費用', '物流通訊', '支出', '固定', '間接', '電話網路月租'],
  ['E004', '手續費', '營運費用', '支付工具', '支出', '變動', '間接', '刷卡、行動支付手續費'],
  ['E005', '差旅交通費', '營運費用', '物流通訊', '支出', '變動', '直接', '店務差旅(可能有沖銷負值)'],
  ['E006', '文具用品', '營運費用', '營運耗材', '支出', '變動', '直接', '【v2 新增】辦公文具'],
  ['E007', '雜項購置', '營運費用', '營運耗材', '支出', '變動', '直接', '【v2 新增】小額資本支出,未達資本化門檻'],
  ['H001', '廣告費', '行銷費用', '廣告推廣', '支出', '變動', '直接', '【v2 新增】店舖行銷廣告投入'],
  ['F001', '管理費', '管理費用', '管理費', '支出', '固定', '不可控', '集團統籌管理費'],
  ['F002', '內部管理費', '管理費用', '管理費', '支出', '半變動', '不可控', '集團抽成,隨營收浮動'],
  ['F003', '勞務費', '管理費用', '管理費', '支出', '半變動', '間接', '外包服務費'],
  ['F004', '軟體維護費用', '管理費用', '管理費', '支出', '固定', '不可控', 'POS、ERP 等'],
  ['F005', '產險', '管理費用', '保險稅負', '支出', '固定', '不可控', '公司產險'],
  ['F006', '稅捐', '管理費用', '保險稅負', '支出', '固定', '不可控', '營業稅、房屋稅等'],
  ['F007', '雜費', '管理費用', '雜項', '支出', '變動', '間接', '無法分類之雜項支出'],
  ['F008', '折舊及折耗', '管理費用', '折舊攤銷', '支出', '固定', '不可控', '【v2 新增】固定資產折舊'],
  ['F009', '各項攤銷', '管理費用', '折舊攤銷', '支出', '固定', '不可控', '【v2 新增】裝修、商標等無形資產攤銷'],
  ['F010', '研發費用', '管理費用', '雜項', '支出', '變動', '不可控', '【v2 新增】新品研發投入'],
  ['F011', '其他費用', '管理費用', '雜項', '支出', '變動', '間接', '【v2 新增】無法歸類之其他費用'],
  ['G001', '其他收入', '營業外', '營業外收入', '收入', '變動', '不可控', '獎勵金、補助等'],
  ['G002', '租賃稅費', '營業外', '營業外支出', '支出', '固定', '不可控', '租賃相關稅負'],
  ['G003', '其他損失', '營業外', '營業外支出', '支出', '變動', '不可控', '【v2 新增】報廢、罰款等非預期損失']
];


// ============================================================
// 【函式 1】備份 dim_account
// ============================================================

function pre1_backupDimAccount() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const source = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  if (!source) {
    Logger.log('❌ 找不到 dim_account 分頁');
    return null;
  }

  const backupName = 'dim_account_backup_20260426';

  // 如果備份已存在,先刪除舊的
  const existing = ss.getSheetByName(backupName);
  if (existing) {
    ss.deleteSheet(existing);
    Logger.log('⚠️ 舊備份已存在,已刪除:' + backupName);
  }

  // 複製分頁
  const backup = source.copyTo(ss);
  backup.setName(backupName);

  // 把備份分頁移到最後
  ss.setActiveSheet(backup);
  ss.moveActiveSheet(ss.getNumSheets());

  Logger.log('✅ 備份完成:' + backupName);
  Logger.log('   原分頁筆數:' + (source.getLastRow() - 1));
  Logger.log('   備份筆數:' + (backup.getLastRow() - 1));
  return backupName;
}


// ============================================================
// 【函式 2】用 39 筆 v2 替換 dim_account
// ============================================================

function pre1_replaceDimAccount() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAMES.ACCOUNT);
  if (!sheet) {
    Logger.log('❌ 找不到 dim_account 分頁');
    return false;
  }

  // 安全檢查:確認備份分頁存在,沒備份不替換
  const backup = ss.getSheetByName('dim_account_backup_20260426');
  if (!backup) {
    Logger.log('❌ 沒找到備份分頁 dim_account_backup_20260426');
    Logger.log('   請先執行 pre1_backupDimAccount()');
    return false;
  }

  const beforeRows = sheet.getLastRow();
  Logger.log('替換前筆數(含表頭):' + beforeRows);

  // 清空現有內容(保留分頁本身)
  sheet.clear();

  // 寫入 39 筆 v2 資料
  const numRows = DIM_ACCOUNT_V2_DATA.length;
  const numCols = DIM_ACCOUNT_V2_DATA[0].length;
  sheet.getRange(1, 1, numRows, numCols).setValues(DIM_ACCOUNT_V2_DATA);

  const afterRows = sheet.getLastRow();
  Logger.log('替換後筆數(含表頭):' + afterRows);
  Logger.log('  資料筆數(扣表頭):' + (afterRows - 1));

  // 驗證
  const lastCode = sheet.getRange(numRows, 1).getValue();
  if (lastCode !== 'G003') {
    Logger.log('❌ 驗證失敗:最後一筆 account_code 應為 G003,實際為 ' + lastCode);
    return false;
  }

  Logger.log('✅ 替換完成,最後一筆驗證通過(G003)');
  return true;
}


// ============================================================
// 【函式 3】查 fact_pnl 基線筆數
// ============================================================

function checkFactPnlBaseline() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('fact_pnl');
  if (!sheet) {
    Logger.log('❌ 找不到 fact_pnl 分頁');
    return -1;
  }

  const lastRow = sheet.getLastRow();
  const dataRows = lastRow - 1;  // 扣表頭

  Logger.log('========================================');
  Logger.log('  fact_pnl 基線檢查');
  Logger.log('========================================');
  Logger.log('總列數(含表頭):' + lastRow);
  Logger.log('資料筆數(扣表頭):' + dataRows);
  Logger.log('========================================');

  return dataRows;
}


// ============================================================
// 【函式 4】PRE 階段一次性協調器
// ============================================================

function pre1_runAll() {
  Logger.log('============================================================');
  Logger.log('  M2 PRE 階段執行開始');
  Logger.log('  時間:' + new Date().toISOString());
  Logger.log('============================================================');
  Logger.log('');

  // Step 1: 查 fact_pnl 基線
  Logger.log('[Step 1/4] 查 fact_pnl 基線筆數');
  const baseline = checkFactPnlBaseline();
  if (baseline < 0) {
    Logger.log('❌ 基線檢查失敗,中止');
    return;
  }
  Logger.log('');

  // Step 2: 備份 dim_account
  Logger.log('[Step 2/4] 備份 dim_account');
  const backupName = pre1_backupDimAccount();
  if (!backupName) {
    Logger.log('❌ 備份失敗,中止');
    return;
  }
  Logger.log('');

  // Step 3: 替換 dim_account 為 39 筆
  Logger.log('[Step 3/4] 替換 dim_account 為 v2(39 筆)');
  const replaced = pre1_replaceDimAccount();
  if (!replaced) {
    Logger.log('❌ 替換失敗,中止');
    return;
  }
  Logger.log('');

  // Step 4: 重跑 testLoadDimAccount(parser.gs 裡的)驗證
  Logger.log('[Step 4/4] 跑 testLoadDimAccount 驗證');
  testLoadDimAccount();  // 這是 parser.gs 裡的函式
  Logger.log('');

  // Step 5: 再查一次 fact_pnl 確認沒被動
  Logger.log('[Step 5/5] 再查 fact_pnl 確認沒被動');
  const afterBaseline = checkFactPnlBaseline();
  Logger.log('');

  // 收尾
  Logger.log('============================================================');
  Logger.log('  PRE 階段完成');
  Logger.log('============================================================');
  Logger.log('fact_pnl 基線(前):' + baseline);
  Logger.log('fact_pnl 基線(後):' + afterBaseline);
  Logger.log('fact_pnl 是否變動:' + (baseline === afterBaseline ? '✅ 沒變' : '❌ 有變'));
  Logger.log('dim_account 備份分頁:' + backupName);
  Logger.log('dim_account 已替換為 v2(預期載入 39 筆)');
  Logger.log('============================================================');
}


// ============================================================
// 【M3 PRE】基線檢查 + staging 清空
// ============================================================

function m3_pre_baseline() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const factRows = ss.getSheetByName('fact_pnl').getLastRow() - 1;
  const dimRows = ss.getSheetByName('dim_account').getLastRow() - 1;
  const stagingSheet = ss.getSheetByName('fact_pnl_staging');
  const stagingRows = stagingSheet ? stagingSheet.getLastRow() - 1 : 0;
  const backups = [];
  ss.getSheets().forEach(function(s) {
    const n = s.getName();
    if (n.indexOf('fact_pnl_backup_') === 0 || n.indexOf('dim_account_backup_') === 0) {
      backups.push(n + ' (' + (s.getLastRow() - 1) + ' 筆)');
    }
  });
  Logger.log('=== M3 PRE 基線檢查 ===');
  Logger.log('fact_pnl 筆數: ' + factRows);
  Logger.log('dim_account 筆數: ' + dimRows + ' (預期 39)');
  Logger.log('fact_pnl_staging 筆數: ' + stagingRows + ' (預期 0 或 21)');
  Logger.log('備份分頁清單(' + backups.length + ' 個):');
  backups.forEach(function(b) { Logger.log('  - ' + b); });
  return { factRows: factRows, dimRows: dimRows, stagingRows: stagingRows, backups: backups };
}

function m3_pre_clearStaging() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('fact_pnl_staging');
  if (!sheet) { Logger.log('staging 分頁不存在,跳過'); return { skipped: true }; }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { Logger.log('staging 已空,跳過'); return { skipped: true }; }
  sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  Logger.log('✅ staging 已清空(保留表頭),原 ' + (lastRow - 1) + ' 筆');
  return { cleared: true, originalRows: lastRow - 1 };
}

function m3_pre_runAll() {
  Logger.log('========================================');
  Logger.log('  M3 PRE 階段一次性執行');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('--- Step 1: 基線記錄(清空前) ---');
  const before = m3_pre_baseline();
  Logger.log('');
  Logger.log('--- Step 2: 清空 staging ---');
  m3_pre_clearStaging();
  Logger.log('');
  Logger.log('--- Step 3: 基線再記錄(清空後) ---');
  const after = m3_pre_baseline();
  Logger.log('');
  Logger.log('========================================');
  Logger.log('  M3 PRE 摘要');
  Logger.log('========================================');
  Logger.log('fact_pnl(前/後): ' + before.factRows + ' / ' + after.factRows);
  Logger.log('fact_pnl 是否變動: ' + (before.factRows === after.factRows ? '✅ 沒變' : '❌ 有變'));
  Logger.log('dim_account: ' + after.dimRows);
  Logger.log('staging(清空後): ' + after.stagingRows);
  Logger.log('備份分頁總數: ' + after.backups.length);
}
