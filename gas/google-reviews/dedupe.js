function dedupe536() {
  throw new Error('dedupe536 已停用，請改用新流程');
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const SHEET_NAME = '工作表1';
  const ORIGINAL_END_ROW = 4929; // header + 4928 baseline
  
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  
  Logger.log('Total rows: ' + lastRow);
  Logger.log('Original baseline: rows 2-' + ORIGINAL_END_ROW);
  Logger.log('New batch: rows ' + (ORIGINAL_END_ROW + 1) + '-' + lastRow);
  
  // 讀原本 4928 筆的 reviewId(F 欄)建 Set
  const originalIds = new Set();
  const originalData = sheet.getRange(2, 6, ORIGINAL_END_ROW - 1, 1).getValues();
  originalData.forEach(row => {
    if (row[0]) originalIds.add(row[0]);
  });
  Logger.log('Original unique reviewIds: ' + originalIds.size);
  
  // 讀新批次的 reviewId
  const newBatchStart = ORIGINAL_END_ROW + 1;
  const newBatchSize = lastRow - ORIGINAL_END_ROW;
  if (newBatchSize <= 0) {
    Logger.log('No new batch to dedupe');
    return;
  }
  const newData = sheet.getRange(newBatchStart, 6, newBatchSize, 1).getValues();
  
  // 找重複的 row 編號(從新批次裡)
  const duplicateRows = [];
  const newOnly = [];
  newData.forEach((row, i) => {
    const rowNum = newBatchStart + i;
    if (originalIds.has(row[0])) {
      duplicateRows.push(rowNum);
    } else {
      newOnly.push(rowNum);
    }
  });
  Logger.log('Duplicates in new batch: ' + duplicateRows.length);
  Logger.log('Truly new rows: ' + newOnly.length);
  
  // 從下往上刪重複 row(避免 index shift)
  duplicateRows.sort((a, b) => b - a);
  duplicateRows.forEach(r => sheet.deleteRow(r));
  
  Logger.log('Deletion complete. Final lastRow: ' + sheet.getLastRow());
  
  return {
    originalCount: originalIds.size,
    newBatchTotal: newBatchSize,
    duplicates: duplicateRows.length,
    trulyNew: newOnly.length,
    finalLastRow: sheet.getLastRow()
  };
}

function diagnoseDupes() {
  Logger.log('=== diagnoseDupes START at ' + new Date() + ' ===');
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('工作表1');
  const lastRow = sheet.getLastRow();
  const totalRows = lastRow - 1; // exclude header

  Logger.log('=== diagnoseDupes 診斷報告 ===');
  Logger.log('總筆數(不含表頭): ' + totalRows);

  if (totalRows <= 0) {
    Logger.log('無資料');
    return;
  }

  // 讀 C 欄(評論時間) + F 欄(reviewId)
  const data = sheet.getRange(2, 3, totalRows, 4).getValues(); // C,D,E,F => index 0=C, 3=F

  let blankF = 0;
  const legacyRows = [];
  const gmbRows = [];

  data.forEach(function(row) {
    const reviewId = String(row[3] || '').trim();
    const timeVal = row[0]; // C 欄

    if (!reviewId) {
      blankF++;
      return;
    }

    if (reviewId.indexOf('LEGACY_') === 0) {
      legacyRows.push({ id: reviewId, time: timeVal });
    } else {
      gmbRows.push({ id: reviewId, time: timeVal });
    }
  });

  Logger.log('F 欄空白筆數: ' + blankF);
  Logger.log('');

  // helper: min/max date from array
  function minMax(arr) {
    if (!arr.length) return { min: 'N/A', max: 'N/A' };
    const times = arr.map(r => r.time).filter(t => t instanceof Date || (typeof t === 'string' && t.length > 0));
    if (!times.length) return { min: 'N/A', max: 'N/A' };
    const sorted = times.map(t => new Date(t)).sort((a,b) => a-b);
    return {
      min: Utilities.formatDate(sorted[0], 'Asia/Taipei', 'yyyy-MM-dd'),
      max: Utilities.formatDate(sorted[sorted.length-1], 'Asia/Taipei', 'yyyy-MM-dd')
    };
  }

  const legacyMM = minMax(legacyRows);
  const gmbMM = minMax(gmbRows);

  Logger.log('--- LEGACY 堆 ---');
  Logger.log('筆數: ' + legacyRows.length);
  Logger.log('最早: ' + legacyMM.min);
  Logger.log('最晚: ' + legacyMM.max);
  Logger.log('');
  Logger.log('--- GMB 堆 ---');
  Logger.log('筆數: ' + gmbRows.length);
  Logger.log('最早: ' + gmbMM.min);
  Logger.log('最晚: ' + gmbMM.max);
  Logger.log('');

  // 重複 reviewId 分析
  const idCount = {};
  data.forEach(function(row) {
    const id = String(row[3] || '').trim();
    if (!id) return;
    idCount[id] = (idCount[id] || 0) + 1;
  });

  let dup2 = 0, dup3 = 0, dup4plus = 0, totalDupGroups = 0;
  Object.keys(idCount).forEach(function(id) {
    const c = idCount[id];
    if (c >= 2) {
      totalDupGroups++;
      if (c === 2) dup2++;
      else if (c === 3) dup3++;
      else dup4plus++;
    }
  });

  Logger.log('--- 重複 reviewId 分析 ---');
  Logger.log('重複組數(總): ' + totalDupGroups);
  Logger.log('出現 2 次: ' + dup2 + ' 組');
  Logger.log('出現 3 次: ' + dup3 + ' 組');
  Logger.log('出現 ≥4 次: ' + dup4plus + ' 組');
  Logger.log('');

  // 時間範圍重疊?
  Logger.log('--- 時間範圍重疊判斷 ---');
  Logger.log('LEGACY 最晚: ' + legacyMM.max);
  Logger.log('GMB 最早: ' + gmbMM.min);
  if (legacyMM.max !== 'N/A' && gmbMM.min !== 'N/A') {
    const legacyEnd = new Date(legacyMM.max);
    const gmbStart = new Date(gmbMM.min);
    if (legacyEnd >= gmbStart) {
      Logger.log('⚠️ 有重疊! LEGACY 最晚(' + legacyMM.max + ') >= GMB 最早(' + gmbMM.min + ')');
    } else {
      Logger.log('✅ 無重疊。LEGACY 最晚(' + legacyMM.max + ') < GMB 最早(' + gmbMM.min + ')');
    }
  } else {
    Logger.log('無法判斷(有 N/A)');
  }

  Logger.log('=== 診斷完畢 ===');
}


function diagnoseDupes() {
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('工作表1');
  const lastRow = sheet.getLastRow();
  const totalRows = lastRow - 1;
  Logger.log('=== diagnoseDupes ===');
  Logger.log('總筆數(不含表頭): ' + totalRows);
  if (totalRows <= 0) { Logger.log('無資料'); return; }
  const data = sheet.getRange(2, 3, totalRows, 4).getValues();
  let blankF = 0;
  const legacyRows = [];
  const gmbRows = [];
  data.forEach(function(row) {
    const reviewId = String(row[3] || '').trim();
    const timeVal = row[0];
    if (!reviewId) { blankF++; return; }
    if (reviewId.indexOf('LEGACY_') === 0) { legacyRows.push({id: reviewId, time: timeVal}); }
    else { gmbRows.push({id: reviewId, time: timeVal}); }
  });
  Logger.log('F欄空白: ' + blankF);
  function minMax(arr) {
    if (!arr.length) return {min:'N/A',max:'N/A'};
    const sorted = arr.map(function(r){return new Date(r.time);}).filter(function(d){return !isNaN(d);}).sort(function(a,b){return a-b;});
    if (!sorted.length) return {min:'N/A',max:'N/A'};
    return {min:Utilities.formatDate(sorted[0],'Asia/Taipei','yyyy-MM-dd'),max:Utilities.formatDate(sorted[sorted.length-1],'Asia/Taipei','yyyy-MM-dd')};
  }
  const lMM = minMax(legacyRows);
  const gMM = minMax(gmbRows);
  Logger.log('LEGACY筆數: ' + legacyRows.length + ' 最早:' + lMM.min + ' 最晚:' + lMM.max);
  Logger.log('GMB筆數: ' + gmbRows.length + ' 最早:' + gMM.min + ' 最晚:' + gMM.max);
  const idCount = {};
  data.forEach(function(row) {
    const id = String(row[3] || '').trim();
    if (!id) return;
    idCount[id] = (idCount[id] || 0) + 1;
  });
  let dup2=0,dup3=0,dup4p=0,total=0;
  Object.keys(idCount).forEach(function(id){
    const c=idCount[id];
    if(c>=2){total++;if(c===2)dup2++;else if(c===3)dup3++;else dup4p++;}
  });
  Logger.log('重複組數: ' + total + ' (2次:' + dup2 + ' 3次:' + dup3 + ' >=4次:' + dup4p + ')');
  if (lMM.max!=='N/A' && gMM.min!=='N/A') {
    Logger.log('時間重疊: ' + (new Date(lMM.max)>=new Date(gMM.min) ? '有重疊 LEGACY最晚='+lMM.max+' GMB最早='+gMM.min : '無重疊 LEGACY最晚='+lMM.max+' GMB最早='+gMM.min));
  }
  Logger.log('=== END ===');
}

function dupeLocation() {
  Logger.log('=== dupeLocation START at ' + new Date() + ' ===');
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('工作表1');
  const lastRow = sheet.getLastRow();
  const totalRows = lastRow - 1;
  if (totalRows <= 0) { Logger.log('無資料'); return; }

  // Read F column (column 6), all data rows
  const fData = sheet.getRange(2, 6, totalRows, 1).getValues();

  // Build map: reviewId -> array of row indices and type
  const idMap = {};
  fData.forEach(function(row, i) {
    const id = String(row[0] || '').trim();
    if (!id) return;
    if (!idMap[id]) idMap[id] = [];
    idMap[id].push(i + 2); // actual row number (1-indexed, +1 for header)
  });

  // Find duplicates and classify
  let legacyDupes = 0;
  let gmbDupes = 0;

  Object.keys(idMap).forEach(function(id) {
    const rows = idMap[id];
    if (rows.length < 2) return;
    const isLegacy = id.indexOf('LEGACY_') === 0;
    if (isLegacy) {
      legacyDupes++;
    } else {
      gmbDupes++;
    }
  });

  Logger.log('重複組總數: ' + (legacyDupes + gmbDupes));
  Logger.log('LEGACY堆內重複組數: ' + legacyDupes);
  Logger.log('GMB堆內重複組數: ' + gmbDupes);
  Logger.log('=== dupeLocation END ===');
}

function createBackup() {
  Logger.log('=== createBackup START ===');
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('工作表1');
  const backupName = 'reviews_backup_before_gmb_purge_20260504';
  // Check if backup already exists
  const existing = ss.getSheetByName(backupName);
  if (existing) {
    Logger.log('備份已存在: ' + backupName);
    return;
  }
  const backup = sheet.copyTo(ss);
  backup.setName(backupName);
  Logger.log('備份建立完成: ' + backupName);
  Logger.log('備份列數: ' + backup.getLastRow());
  Logger.log('=== createBackup END ===');
}

function purgeAllGMB() {
  Logger.log('=== purgeAllGMB START at ' + new Date() + ' ===');
  const SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('工作表1');
  const lastRow = sheet.getLastRow();
  
  // Read F column all at once
  const fData = sheet.getRange(2, 6, lastRow - 1, 1).getValues();
  
  // Collect rows to delete (non-LEGACY_)
  const toDelete = [];
  fData.forEach(function(row, i) {
    const id = String(row[0] || '').trim();
    if (id.indexOf('LEGACY_') !== 0) {
      toDelete.push(i + 2); // actual row number
    }
  });
  
  const total = toDelete.length;
  Logger.log('待刪列數: ' + total);
  
  // Log first 50
  const first50 = toDelete.slice(0, 50);
  Logger.log('前50筆row: ' + first50.join(','));
  if (total > 100) {
    Logger.log('... (中間省略) ...');
    const last50 = toDelete.slice(-50);
    Logger.log('後50筆row: ' + last50.join(','));
  }
  
  // Delete from bottom to top to avoid index shift
  const sortedDesc = toDelete.slice().sort(function(a, b) { return b - a; });
  sortedDesc.forEach(function(rowNum) {
    sheet.deleteRow(rowNum);
  });
  
  Logger.log('刪除完成，總刪除筆數: ' + total);
  Logger.log('目前 lastRow: ' + sheet.getLastRow());
  Logger.log('=== purgeAllGMB END ===');
}


function listRecentGMBDupes() {
  Logger.log('=== listRecentGMBDupes START ===');
  const SHEET_ID   = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  const SHEET_NAME = '工作表1';
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  Logger.log('Total rows (incl header): ' + lastRow);

  // 讀 F欄(reviewId) + C欄(createTime), 從 row2 到 lastRow
  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, 6).getValues(); // A~F

  // 建 map: reviewId -> [{rowIndex, createTime}]
  const idMap = {};
  data.forEach(function(row, i) {
    const reviewId   = row[5]; // F欄
    const createTime = row[2]; // C欄
    if (!reviewId || String(reviewId).indexOf('LEGACY_') === 0) return; // 跳過空值和LEGACY
    const rowIndex = i + 2; // sheet 真實行號
    if (!idMap[reviewId]) idMap[reviewId] = [];
    idMap[reviewId].push({ rowIndex: rowIndex, createTime: createTime });
  });

  // 找出重複組 (>=2筆)
  const dupeGroups = [];
  Object.keys(idMap).forEach(function(rid) {
    if (idMap[rid].length >= 2) {
      dupeGroups.push({ reviewId: rid, entries: idMap[rid] });
    }
  });

  Logger.log('重複 GMB reviewId 組數: ' + dupeGroups.length);
  dupeGroups.forEach(function(g, i) {
    const rows = g.entries.map(function(e) { return 'row' + e.rowIndex + '(time:' + e.createTime + ')'; }).join(' | ');
    Logger.log('[' + (i+1) + '] reviewId前30字: ' + String(g.reviewId).substring(0,30) + ' => ' + rows);
  });

  Logger.log('=== listRecentGMBDupes END ===');
}

function removeRecentDupes() {
  Logger.log('=== removeRecentDupes START at ' + new Date() + ' ===');
  var SHEET_ID   = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  var SHEET_NAME = '工作表1';
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  Logger.log('執行前 lastRow: ' + lastRow);

  var numRows = lastRow - 1;
  var data = sheet.getRange(2, 1, numRows, 6).getValues();

  var idMap = {};
  data.forEach(function(row, i) {
    var reviewId = row[5];
    if (!reviewId || String(reviewId).indexOf('LEGACY_') === 0) return;
    var rowIndex = i + 2;
    if (!idMap[reviewId]) idMap[reviewId] = [];
    idMap[reviewId].push(rowIndex);
  });

  var toDelete = [];
  Object.keys(idMap).forEach(function(rid) {
    if (idMap[rid].length >= 2) {
      var sorted = idMap[rid].slice().sort(function(a, b) { return a - b; });
      for (var k = 1; k < sorted.length; k++) {
        toDelete.push({ rowIndex: sorted[k], reviewId: rid });
      }
    }
  });

  Logger.log('準備刪除筆數: ' + toDelete.length);

  if (toDelete.length === 0) {
    Logger.log('沒有重複,無需刪除。');
    Logger.log('=== removeRecentDupes END ===');
    return;
  }

  toDelete.sort(function(a, b) { return b.rowIndex - a.rowIndex; });

  toDelete.forEach(function(item) {
    Logger.log('刪除 row' + item.rowIndex + ' reviewId前30字: ' + String(item.reviewId).substring(0, 30));
    sheet.deleteRow(item.rowIndex);
  });

  var newLastRow = sheet.getLastRow();
  Logger.log('刪除完成,總刪除數: ' + toDelete.length);
  Logger.log('執行後 lastRow: ' + newLastRow);
  Logger.log('=== removeRecentDupes END ===');
}

function dailyDedupeGMB() {
  Logger.log('=== dailyDedupeGMB START at ' + new Date() + ' ===');
  var SHEET_ID   = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  var SHEET_NAME = '工作表1';
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  Logger.log('執行前 lastRow: ' + lastRow);

  var numRows = lastRow - 1;
  if (numRows <= 0) {
    Logger.log('無資料,結束。');
    Logger.log('=== dailyDedupeGMB END ===');
    return;
  }

  var data = sheet.getRange(2, 1, numRows, 6).getValues();

  var idMap = {};
  data.forEach(function(row, i) {
    var reviewId = row[5];
    if (!reviewId || String(reviewId).indexOf('LEGACY_') === 0) return;
    var rowIndex = i + 2;
    if (!idMap[reviewId]) idMap[reviewId] = [];
    idMap[reviewId].push(rowIndex);
  });

  var toDelete = [];
  Object.keys(idMap).forEach(function(rid) {
    if (idMap[rid].length >= 2) {
      var sorted = idMap[rid].slice().sort(function(a, b) { return a - b; });
      for (var k = 1; k < sorted.length; k++) {
        toDelete.push({ rowIndex: sorted[k], reviewId: rid });
      }
    }
  });

  Logger.log('發現重複筆數: ' + toDelete.length);

  if (toDelete.length === 0) {
    Logger.log('無重複,今日跳過。');
    Logger.log('=== dailyDedupeGMB END ===');
    return;
  }

  toDelete.sort(function(a, b) { return b.rowIndex - a.rowIndex; });

  toDelete.forEach(function(item) {
    Logger.log('刪除 row' + item.rowIndex + ' reviewId前30字: ' + String(item.reviewId).substring(0, 30));
    sheet.deleteRow(item.rowIndex);
  });

  var newLastRow = sheet.getLastRow();
  Logger.log('刪除完成,總刪除數: ' + toDelete.length);
  Logger.log('執行後 lastRow: ' + newLastRow);
  Logger.log('=== dailyDedupeGMB END ===');
}

function installDailyDedupeTrigger() {
  Logger.log('=== installDailyDedupeTrigger START ===');
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyDedupeGMB') {
      ScriptApp.deleteTrigger(t);
      Logger.log('已刪除舊 trigger: ' + t.getUniqueId());
    }
  });

  ScriptApp.newTrigger('dailyDedupeGMB')
    .timeBased()
    .atHour(6)
    .nearMinute(30)
    .inTimezone('Asia/Taipei')
    .everyDays(1)
    .create();

  Logger.log('Trigger 建立完成: dailyDedupeGMB @ 09:30 Asia/Taipei');

  var allTriggers = ScriptApp.getProjectTriggers();
  Logger.log('現存 trigger 數: ' + allTriggers.length);
  allTriggers.forEach(function(t, i) {
    Logger.log('[trigger' + (i+1) + '] fn=' + t.getHandlerFunction() + ' type=' + t.getEventType() + ' id=' + t.getUniqueId());
  });
  Logger.log('=== installDailyDedupeTrigger END ===');
}
