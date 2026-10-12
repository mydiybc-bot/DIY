// ============================================
// fact_pnl 重灌操作工具 — B方案(分批解析)
// 執行方式:改 aRUN 裡呼叫的函式 → 存檔 → 按執行
// ============================================
var WAREHOUSE_ID = '1khrFp_AYp3mEsL02cTRrYu1LNiRTP68aowM5_ClsE_c';
var FOLDER_ID = '1cVq5IISD36I0uC63GI9ZuVATNGVpRqJ1';
var STAGING_NAME = 'fact_pnl_staging';

// ===== 執行轉發器:改這裡的呼叫,存檔,按執行 =====
function aRUN() {
  step4_commit();   // <- 下一步要跑什麼就改這行
}

// [步驟 2] 掃描資料夾
function step2_scan() {
  Logger.log('=== 步驟 2:掃描資料夾 ===');
  var scan = scanDingxinFolder(FOLDER_ID);
  Logger.log('認到符合規範的檔:' + scan.files.length + ' 個');
  Logger.log('跳過(不符規範):' + scan.skipped.length + ' 個');
  if (scan.skipped.length > 0) {
    scan.skipped.forEach(function(s) { Logger.log('  跳過:' + s.name + ' -> ' + s.reason); });
  }
  var yms = {};
  scan.files.forEach(function(f) { yms[f.ymFromName] = (yms[f.ymFromName]||0)+1; });
  Object.keys(yms).sort().forEach(function(ym) { Logger.log('  ' + ym + ': ' + yms[ym] + ' 檔'); });
}

// [步驟 1] 備份 + 清空 fact_pnl(整個流程只跑一次)
function step1_backup_and_clear() {
  Logger.log('=== 步驟 1:備份 + 清空 fact_pnl ===');
  var bk = backupFactPnl();
  if (!bk.ok) { Logger.log('X 備份失敗:' + bk.reason + ' - 中止'); return; }
  Logger.log('OK 備份完成:' + bk.backupSheetName + '(' + bk.rowsBackedUp + ' 筆)');
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var fact = ss.getSheetByName('fact_pnl');
  var before = fact.getLastRow() - 1;
  if (fact.getLastRow() > 1) {
    fact.getRange(2, 1, fact.getLastRow() - 1, fact.getLastColumn()).clearContent();
  }
  Logger.log('OK fact_pnl 已清空:' + before + ' 筆 -> ' + (fact.getLastRow() - 1) + ' 筆');
}

// [步驟 3 核心] 分批解析 -> append 寫 staging
function step3b_parse_chunk(startIndex, count) {
  Logger.log('=== 步驟 3b:解析第 ' + startIndex + ' ~ ' + (startIndex+count-1) + ' 檔 ===');
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var fact = ss.getSheetByName('fact_pnl');
  var headers = fact.getRange(1, 1, 1, fact.getLastColumn()).getValues()[0];
  var scan = scanDingxinFolder(FOLDER_ID);
  if (!scan.ok) { Logger.log('X 掃描失敗'); return; }
  var allFiles = scan.files;
  Logger.log('資料夾總檔數:' + allFiles.length);
  var chunk = allFiles.slice(startIndex, startIndex + count);
  Logger.log('本批處理:' + chunk.length + ' 檔');
  if (chunk.length === 0) { Logger.log('W 本批無檔案,可能已超出範圍'); return; }
  var staging = ss.getSheetByName(STAGING_NAME);
  if (startIndex === 0) {
    if (!staging) { staging = ss.insertSheet(STAGING_NAME); }
    if (staging.getLastRow() > 0) { staging.clear(); }
    staging.getRange(1, 1, 1, headers.length).setValues([headers]);
    Logger.log('staging 已清空重建(第一批)');
  } else {
    if (!staging) { Logger.log('X staging 不存在,請先從 startIndex=0 開始'); return; }
    Logger.log('staging append 模式,目前已有 ' + (staging.getLastRow()-1) + ' 筆');
  }
  var allRecords = [];
  var okCount = 0, failCount = 0;
  var failFiles = [];
  chunk.forEach(function(file) {
    var res = parseDingxinFile(file.id);
    if (res.ok) {
      okCount++;
      res.records.forEach(function(rec) { allRecords.push(rec); });
    } else {
      failCount++;
      failFiles.push(file.name + ':' + res.errors.join(';'));
    }
  });
  Logger.log('解析成功:' + okCount + ' 檔,失敗:' + failCount + ' 檔');
  if (failFiles.length > 0) {
    Logger.log('--- X 失敗檔 ---');
    failFiles.forEach(function(f) { Logger.log('  ' + f); });
  }
  if (allRecords.length > 0) {
    var grid = allRecords.map(function(rec) {
      return headers.map(function(h) { return rec[h] !== undefined ? rec[h] : ''; });
    });
    var startRow = staging.getLastRow() + 1;
    staging.getRange(startRow, 1, grid.length, headers.length).setValues(grid);
    Logger.log('本批寫入 staging:' + grid.length + ' 筆');
  }
  Logger.log('OK staging 目前累積:' + (staging.getLastRow() - 1) + ' 筆');
}

// [步驟 3 分批捷徑] 每批 80 檔
function step3_batch1() { step3b_parse_chunk(0, 80); }
function step3_batch2() { step3b_parse_chunk(80, 80); }
function step3_batch3() { step3b_parse_chunk(160, 80); }
function step3_batch4() { step3b_parse_chunk(240, 80); }
function step3_batch5() { step3b_parse_chunk(320, 80); }

// [步驟 4] commit staging -> fact_pnl
function step4_commit() {
  Logger.log('=== 步驟 4:commit -> fact_pnl ===');
  var r = commitStagingToFact({ duplicateMode: 'cancel' });
  Logger.log('commit ok:' + r.ok);
  Logger.log('寫入前:' + r.factPnlBefore + ' 筆');
  Logger.log('本次寫入:' + r.rowsCommitted + ' 筆');
  Logger.log('寫入後:' + r.factPnlAfter + ' 筆');
  if (r.errors.length > 0) { r.errors.forEach(function(e) { Logger.log('  X ' + e); }); }
  if (r.countMismatch) Logger.log('W 筆數核對不符');
}
// [commit 前驗證] 檢查 staging 資料健康度
function verify_staging() {
  Logger.log('=== commit 前驗證:staging 健康檢查 ===');
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var staging = ss.getSheetByName(STAGING_NAME);
  var data = staging.getDataRange().getValues();
  var header = data[0];
  var idxYm = header.indexOf('year_month');
  var idxStore = header.indexOf('store_id');
  var idxCode = header.indexOf('account_code');
  var idxAmt = header.indexOf('amount');

  Logger.log('staging 總筆數:' + (data.length - 1));

  var ymStores = {};
  var ymRows = {};
  var badRows = 0;
  var codeSet = {};
  for (var i = 1; i < data.length; i++) {
    var ym = data[i][idxYm];
    var st = data[i][idxStore];
    var code = data[i][idxCode];
    var amt = data[i][idxAmt];
    if (!ym || !st || !code) { badRows++; continue; }
    if (!ymStores[ym]) ymStores[ym] = {};
    ymStores[ym][st] = true;
    ymRows[ym] = (ymRows[ym] || 0) + 1;
    codeSet[code] = (codeSet[code] || 0) + 1;
  }

  Logger.log('=== 各年月:店數 / 筆數 ===');
  Object.keys(ymStores).sort().forEach(function(ym) {
    Logger.log('  ' + ym + ': ' + Object.keys(ymStores[ym]).length + ' 店, ' + ymRows[ym] + ' 筆');
  });

  Logger.log('=== 異常列(缺 year_month/store_id/account_code):' + badRows + ' 列 ===');

  Logger.log('=== account_code 種類數:' + Object.keys(codeSet).length + ' ===');
  ['H001','F008','F009','G003','A001'].forEach(function(c) {
    Logger.log('  ' + c + ': ' + (codeSet[c] || 0) + ' 筆');
  });

  var allStores = {};
  for (var i = 1; i < data.length; i++) { allStores[data[i][idxStore]] = true; }
  Logger.log('=== 出現的 store_id:' + Object.keys(allStores).sort(function(a,b){return a-b;}).join(',') + ' ===');
}
// [commit 前驗證] 檢查 staging 資料健康度
function verify_staging() {
  Logger.log("=== commit 前驗證:staging 健康檢查 ===");
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var staging = ss.getSheetByName(STAGING_NAME);
  var data = staging.getDataRange().getValues();
  var header = data[0];
  var idxYm = header.indexOf("year_month");
  var idxStore = header.indexOf("store_id");
  var idxCode = header.indexOf("account_code");
  var idxAmt = header.indexOf("amount");
  Logger.log("staging 總筆數:" + (data.length - 1));
  var ymStores = {};
  var ymRows = {};
  var badRows = 0;
  var codeSet = {};
  for (var i = 1; i < data.length; i++) {
    var ym = data[i][idxYm];
    var st = data[i][idxStore];
    var code = data[i][idxCode];
    if (!ym || !st || !code) { badRows++; continue; }
    if (!ymStores[ym]) ymStores[ym] = {};
    ymStores[ym][st] = true;
    ymRows[ym] = (ymRows[ym] || 0) + 1;
    codeSet[code] = (codeSet[code] || 0) + 1;
  }
  Logger.log("=== 各年月:店數 / 筆數 ===");
  Object.keys(ymStores).sort().forEach(function(ym) {
    Logger.log("  " + ym + ": " + Object.keys(ymStores[ym]).length + " 店, " + ymRows[ym] + " 筆");
  });
  Logger.log("=== 異常列(缺欄位):" + badRows + " 列 ===");
  Logger.log("=== account_code 種類數:" + Object.keys(codeSet).length + " ===");
  ["H001","F008","F009","G003","A001"].forEach(function(c) {
    Logger.log("  " + c + ": " + (codeSet[c] || 0) + " 筆");
  });
  var allStores = {};
  for (var i = 1; i < data.length; i++) { allStores[data[i][idxStore]] = true; }
  Logger.log("=== store_id:" + Object.keys(allStores).sort(function(a,b){return a-b;}).join(",") + " ===");
}
// [commit 前驗證]
function verify_staging() {
  Logger.log("=== commit 前驗證:staging 健康檢查 ===");
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var staging = ss.getSheetByName(STAGING_NAME);
  var data = staging.getDataRange().getValues();
  var header = data[0];
  var idxYm = header.indexOf("year_month");
  var idxStore = header.indexOf("store_id");
  var idxCode = header.indexOf("account_code");
  Logger.log("staging 總筆數:" + (data.length - 1));
  var ymStores = {};
  var ymRows = {};
  var badRows = 0;
  var codeSet = {};
  for (var i = 1; i < data.length; i++) {
    var ym = data[i][idxYm];
    var st = data[i][idxStore];
    var code = data[i][idxCode];
    if (!ym || !st || !code) { badRows++; continue; }
    if (!ymStores[ym]) ymStores[ym] = {};
    ymStores[ym][st] = true;
    ymRows[ym] = (ymRows[ym] ? ymRows[ym]+1 : 1);
    codeSet[code] = (codeSet[code] ? codeSet[code]+1 : 1);
  }
  Logger.log("=== 各年月:店數 / 筆數 ===");
  Object.keys(ymStores).sort().forEach(function(ym) {
    Logger.log("  " + ym + ": " + Object.keys(ymStores[ym]).length + " 店, " + ymRows[ym] + " 筆");
  });
  Logger.log("=== 異常列:" + badRows + " ===");
  Logger.log("=== account_code 種類:" + Object.keys(codeSet).length + " ===");
  ["H001","F008","F009","G003","A001"].forEach(function(c) {
    Logger.log("  " + c + ": " + (codeSet[c] ? codeSet[c] : 0) + " 筆");
  });
  var allStores = {};
  for (var i = 1; i < data.length; i++) { allStores[data[i][idxStore]] = true; }
  Logger.log("=== store_id:" + Object.keys(allStores).sort(function(a,b){return a-b;}).join(",") + " ===");
}
// [commit 前驗證]
function verify_staging() {
  Logger.log("=== commit 前驗證:staging 健康檢查 ===");
  var ss = SpreadsheetApp.openById(WAREHOUSE_ID);
  var staging = ss.getSheetByName(STAGING_NAME);
  var data = staging.getDataRange().getValues();
  var header = data[0];
  var idxYm = header.indexOf("year_month");
  var idxStore = header.indexOf("store_id");
  var idxCode = header.indexOf("account_code");
  Logger.log("staging 總筆數:" + (data.length - 1));
  var ymStores = {};
  var ymRows = {};
  var badRows = 0;
  var codeSet = {};
  for (var i = 1; i < data.length; i++) {
    var ym = data[i][idxYm];
    var st = data[i][idxStore];
    var code = data[i][idxCode];
    if (!ym || !st || !code) { badRows++; continue; }
    if (!ymStores[ym]) ymStores[ym] = {};
    ymStores[ym][st] = true;
    ymRows[ym] = (ymRows[ym] ? ymRows[ym]+1 : 1);
    codeSet[code] = (codeSet[code] ? codeSet[code]+1 : 1);
  }
  Logger.log("=== 各年月:店數 / 筆數 ===");
  Object.keys(ymStores).sort().forEach(function(ym) {
    Logger.log("  " + ym + ": " + Object.keys(ymStores[ym]).length + " 店, " + ymRows[ym] + " 筆");
  });
  Logger.log("=== 異常列:" + badRows + " ===");
  Logger.log("=== account_code 種類:" + Object.keys(codeSet).length + " ===");
  ["H001","F008","F009","G003","A001"].forEach(function(c) {
    Logger.log("  " + c + ": " + (codeSet[c] ? codeSet[c] : 0) + " 筆");
  });
  var allStores = {};
  for (var i = 1; i < data.length; i++) { allStores[data[i][idxStore]] = true; }
  Logger.log("=== store_id:" + Object.keys(allStores).sort(function(a,b){return a-b;}).join(",") + " ===");
}