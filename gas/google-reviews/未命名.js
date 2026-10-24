function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.token !== 'diybc_rebuild_2024') {
    return ContentService.createTextOutput('unauthorized');
  }
  var ss = SpreadsheetApp.openById('1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw');
  var sheet = ss.getSheetByName('工作表1');
  var rows = data.rows;
  if (!rows || rows.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ok: true, added: 0}));
  }
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  return ContentService.createTextOutput(JSON.stringify({ok: true, added: rows.length, newLastRow: lastRow + rows.length}));
}