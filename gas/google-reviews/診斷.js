function 診斷評論() {
  var sh = SpreadsheetApp.openById('1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw')
                         .getSheetByName('工作表1');
  var n = sh.getLastRow();
  var v = sh.getRange(2, 1, n - 1, 11).getValues();
  Logger.log('GAS 讀到列數: ' + v.length);

  var total = 0, jul = [];
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).indexOf('草悟道') < 0) continue;
    total++;
    var g = v[i][6], gt = Object.prototype.toString.call(g);
    var gs = (gt === '[object Date]')
      ? g.getFullYear() + '-' + ('0'+(g.getMonth()+1)).slice(-2) : String(g);
    if (gs.indexOf('2026-07') >= 0)
      jul.push('列'+(i+2)+' '+v[i][1]+' | G型別='+gt+' 值='+g+' | 星等='+v[i][7]);
  }
  Logger.log('草悟道 總筆數: ' + total);
  Logger.log('草悟道 2026-07: ' + jul.length + ' 筆');
  jul.forEach(function(x){ Logger.log(x); });
}