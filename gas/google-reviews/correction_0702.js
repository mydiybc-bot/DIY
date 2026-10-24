function correction0702() {
  var DRY_RUN = false; // 先跑一次 true 確認邏輯正確，再改 false 正式執行

  var SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  var SHEET_NAME = '工作表1';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  var report = [];
  function log(msg) { report.push(msg); Logger.log(msg); }

  log('=== correction0702 START (DRY_RUN=' + DRY_RUN + ') ' + new Date() + ' ===');

  var row1 = sheet.getRange(1,1,1,3).getValues()[0];
  if (String(row1[0]) !== '台中精明店' || String(row1[1]) !== '仙人掌*') {
    throw new Error('PRE-CHECK 失敗：第1列不是預期的「台中精明店|仙人掌*」，實際=' + JSON.stringify(row1) + '。已停止，未做任何修改。');
  }
  log('PRE-CHECK 1 通過：無標題列，第1列=' + JSON.stringify(row1));

  var totalRowsBefore = sheet.getLastRow();
  if (totalRowsBefore !== 9674) {
    throw new Error('PRE-CHECK 失敗：總列數=' + totalRowsBefore + '，預期9674。已停止，未做任何修改。');
  }
  log('PRE-CHECK 2 通過：總列數=9674');

  var backupName = 'backup-0702-6月校正前';
  if (DRY_RUN) {
    log('[DRY RUN] 將建立備份分頁：' + backupName + (ss.getSheetByName(backupName) ? '（已存在，正式執行會中止）' : '（名稱可用）'));
  } else {
    if (ss.getSheetByName(backupName)) {
      throw new Error('PRE-CHECK 失敗：備份分頁「' + backupName + '」已存在，停止避免覆蓋。');
    }
    var backupSheet = sheet.copyTo(ss);
    backupSheet.setName(backupName);
    log('備份完成：已建立分頁「' + backupName + '」');
  }

  var skipped = [];
  function checkName(row, expectedName) {
    var actual = String(sheet.getRange(row, 2).getValue());
    if (actual !== expectedName) {
      skipped.push('第' + row + '列預期B欄=「' + expectedName + '」，實際=「' + actual + '」，已跳過');
      return false;
    }
    return true;
  }

  var taskB = [
    [883, '晞晞', '2023-10'],
    [3153, '安', '2026-05'],
    [3155, '玥', '2026-05'],
    [3207, 'Abby', '2026-05'],
    [3696, 'Yi jou Chang', '2025-08'],
    [3936, '李素琴', '2026-05'],
    [6912, 'fy chong', '2026-05'],
    [7482, 'Stephanie Liao', '2026-05'],
    [8429, 'phoebe', '2026-05'],
    [9058, 'JinChu', '2026-05'],
    [9059, '魏妘如（小金牛）', '2026-05']
  ];
  var doneB = 0;
  taskB.forEach(function(item) {
    var row = item[0], name = item[1], newG = item[2];
    if (!checkName(row, name)) return;
    if (DRY_RUN) {
      var curG = sheet.getRange(row,7).getValue();
      log('[DRY RUN] 第' + row + '列（' + name + '）G欄 ' + curG + ' → ' + newG);
    } else {
      var cell = sheet.getRange(row, 7);
      cell.setNumberFormat('@');
      cell.setValue(newG);
    }
    doneB++;
  });
  log('任務B：' + doneB + ' / 11 筆' + (DRY_RUN ? '（預演）' : '已完成'));

  var taskC = [
    [312, 'Kevin Cheng', '台中精明店'],
    [313, '蔡昀臻', '台中精明店'],
    [479, 'Jia', '台中精明店'],
    [5978, '奇奇跟花花', '吳寶春(高雄SKM)']
  ];
  var doneC = 0;
  taskC.forEach(function(item) {
    var row = item[0], name = item[1], store = item[2];
    if (!checkName(row, name)) return;
    var actualStore = String(sheet.getRange(row, 1).getValue());
    if (actualStore !== store) {
      skipped.push('第' + row + '列預期A欄=「' + store + '」，實際=「' + actualStore + '」，已跳過');
      return;
    }
    if (DRY_RUN) {
      var curG = sheet.getRange(row,7).getValue();
      log('[DRY RUN] 第' + row + '列（' + name + '/' + store + '）G欄 ' + curG + ' → 2026-06');
    } else {
      var cell = sheet.getRange(row, 7);
      cell.setNumberFormat('@');
      cell.setValue('2026-06');
    }
    doneC++;
  });
  log('任務C：' + doneC + ' / 4 筆' + (DRY_RUN ? '（預演）' : '已完成'));

  var newRowNum = sheet.getLastRow() + 1;
  var newRow = [
    '台北士林店',
    'asper J',
    new Date(2026, 5, 9, 0, 0, 0),
    'FIVE',
    '環境好服務態度也很好👍 小瑾服務的很好 雖然人很多 但全程都有細心照顧到我們',
    '',
    '2026-06',
    5,
    '自己做',
    '北一區',
    '達標'
  ];
  if (DRY_RUN) {
    log('[DRY RUN] 將於第 ' + newRowNum + ' 列新增：台北士林店/asper J/2026-06-09/FIVE/...(共11欄)');
  } else {
    var gCell = sheet.getRange(newRowNum, 7);
    gCell.setNumberFormat('@');
    sheet.getRange(newRowNum, 1, 1, newRow.length).setValues([newRow]);
    log('任務D 完成：已於第 ' + newRowNum + ' 列新增 台北士林店/asper J');
  }

  var taskA = [
    [9662, 'Maggie Chuang'],
    [9661, '王小安'],
    [9660, 'Tsung Che Lee'],
    [9659, '希卡'],
    [9658, '曾梓芸'],
    [9656, '邱晏凌'],
    [9654, 'ning'],
    [9653, '鄭安辰'],
    [9652, '許琇惠'],
    [9651, 'Carol Chen'],
    [9650, 'Anna'],
    [9649, 'Tong En']
  ];
  var doneA = 0;
  taskA.forEach(function(item) {
    var row = item[0], name = item[1];
    if (!checkName(row, name)) return;
    if (DRY_RUN) {
      log('[DRY RUN] 將刪除第 ' + row + ' 列（' + name + '）');
    } else {
      sheet.deleteRow(row);
    }
    doneA++;
  });
  log('任務A：' + doneA + ' / 12 筆' + (DRY_RUN ? '（預演）' : '已完成'));

  if (skipped.length > 0) {
    log('跳過清單（共' + skipped.length + '筆）：');
    skipped.forEach(function(s) { log('  - ' + s); });
  } else {
    log('27筆姓名比對全數通過，無跳過');
  }

  if (!DRY_RUN) {
    var lastRowAfter = sheet.getLastRow();
    var allData = sheet.getRange(1, 1, lastRowAfter, 7).getValues();

    var countG202606 = 0;
    var countTaichung = 0;
    var countShilin = 0;
    var keyCount = {};

    allData.forEach(function(row) {
      var store = String(row[0]);
      var name = String(row[1]);
      var time = row[2] instanceof Date ? row[2].getTime() : String(row[2]);
      var g = String(row[6]);

      if (g === '2026-06') {
        countG202606++;
        if (store === '台中精明店') countTaichung++;
        if (store === '台北士林店') countShilin++;
      }
      var key = store + '|' + name + '|' + time;
      keyCount[key] = (keyCount[key] || 0) + 1;
    });

    var dupCount = 0;
    Object.keys(keyCount).forEach(function(k) {
      if (keyCount[k] >= 2) dupCount += (keyCount[k] - 1);
    });

    log('=== 完工自檢 ===');
    log('1. G欄=2026-06 總筆數：' + countG202606 + '（應=283）');
    log('2. 全表總列數：' + lastRowAfter + '（應=9663）');
    log('3. 台中精明店 G欄=2026-06 筆數：' + countTaichung + '（應=18）');
    log('4. 台北士林店 G欄=2026-06 筆數：' + countShilin + '（應=18）');
    log('5. 同店+同評論者+同C欄時間 完全重複筆數：' + dupCount + '（應=0）');
  }

  log('=== correction0702 END ' + new Date() + ' ===');

  return report.join('\n');
}

function verify0702() {
  var SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
  var SHEET_NAME = '工作表1';
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  var tz = ss.getSpreadsheetTimeZone();

  var lastRow = sheet.getLastRow();
  var allData = sheet.getRange(1, 1, lastRow, 7).getValues();

  function normG(v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM');
    return String(v);
  }

  var countG202606 = 0;
  var countTaichung = 0;
  var countShilin = 0;
  var keyCount = {};
  var typeSample = { dateType: 0, textType: 0 };

  allData.forEach(function(row) {
    var store = String(row[0]);
    var name = String(row[1]);
    var time = row[2] instanceof Date ? row[2].getTime() : String(row[2]);
    var g = normG(row[6]);
    if (row[6] instanceof Date) typeSample.dateType++; else typeSample.textType++;

    if (g === '2026-06') {
      countG202606++;
      if (store === '台中精明店') countTaichung++;
      if (store === '台北士林店') countShilin++;
    }
    var key = store + '|' + name + '|' + time;
    keyCount[key] = (keyCount[key] || 0) + 1;
  });

  var dupCount = 0;
  Object.keys(keyCount).forEach(function(k) {
    if (keyCount[k] >= 2) dupCount += (keyCount[k] - 1);
  });

  Logger.log('G欄型別統計：Date型別=' + typeSample.dateType + '筆，文字型別=' + typeSample.textType + '筆');
  Logger.log('1. G欄=2026-06 總筆數（正規化後）：' + countG202606 + '（應=283）');
  Logger.log('2. 全表總列數：' + lastRow + '（應=9663）');
  Logger.log('3. 台中精明店 G欄=2026-06 筆數：' + countTaichung + '（應=18）');
  Logger.log('4. 台北士林店 G欄=2026-06 筆數：' + countShilin + '（應=18）');
  Logger.log('5. 同店+同評論者+同C欄時間 完全重複筆數：' + dupCount + '（應=0）');
}
