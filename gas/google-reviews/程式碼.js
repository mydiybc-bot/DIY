// ====================================================================
// 排班分析資料倉 程式碼.gs v3.2
// 更新日期:2026-05-16
// 版本變更:
//   v3.2: _getStoreDailyData 加員工列表(name/hours/shift/label),供 dashboard hover tooltip
//   v3.1: parser 雙層去重(ScriptProperties by file_name + fact_schedule 已存 file_name 檢查)
//         加 resetFactSchedule + nukeFactSchedule + verifyFactSchedule 工具函式
//   v3.0: 初版,parser + syncDailySales + buildWeekKPI + JSONP API
// ====================================================================
var ANALYSIS_SS_ID = '19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM';
var FOLDER_ID_     = '1yNM8RkYIfdvmETpK9mPbydcn2DXYQYaz';
var NOTIFY_EMAIL_  = 'mydiybc@gmail.com';

// ====================================================================
// Phase 1: 班表 Parser
// ====================================================================

function parsePendingFiles(){
  Logger.log('=== parsePendingFiles START ===');
  var folder = DriveApp.getFolderById(FOLDER_ID_);
  var sp     = PropertiesService.getScriptProperties();

  var existingFiles = _getExistingFileNames();
  Logger.log('fact_schedule already contains ' + Object.keys(existingFiles).length + ' distinct file_names');

  var files = folder.getFiles();
  var cnt = 0;
  while(files.hasNext()){ files.next(); cnt++; }
  Logger.log('total files in folder=' + cnt);

  var files2 = folder.getFiles();
  while(files2.hasNext()){
    var f    = files2.next();
    var mime = f.getMimeType();
    var fname = f.getName();
    var isXls   = (mime === 'application/vnd.ms-excel');
    var isSheet = (mime === 'application/vnd.google-apps.spreadsheet');
    Logger.log('file: ' + fname + ' mime=' + mime);
    if(!isXls && !isSheet){ Logger.log('SKIP: unsupported mime'); continue; }

    if(sp.getProperty('parsed_fname_' + fname) === 'true'){
      Logger.log('SKIP: already processed (by fname property)');
      continue;
    }

    if(existingFiles[fname]){
      Logger.log('SKIP: file_name already in fact_schedule (' + existingFiles[fname] + ' rows)');
      sp.setProperty('parsed_fname_' + fname, 'true');
      continue;
    }

    _processFile(f, isXls, sp, fname);
    existingFiles[fname] = 'just_written';
  }
  Logger.log('=== parsePendingFiles DONE ===');
}

function _getExistingFileNames(){
  var result = {};
  try{
    var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
    var sh = ss.getSheetByName('fact_schedule');
    if(!sh) return result;
    var lastRow = sh.getLastRow();
    if(lastRow < 2) return result;
    var data = sh.getRange(2, 16, lastRow-1, 1).getValues();
    for(var i=0; i<data.length; i++){
      var fn = String(data[i][0]||'').trim();
      if(fn) result[fn] = (result[fn]||0) + 1;
    }
  } catch(e){
    Logger.log('_getExistingFileNames err: ' + e.message);
  }
  return result;
}

function _processFile(f, isXls, sp, fname){
  Logger.log('_processFile: ' + fname);
  var ssId;
  if(isXls){
    var converted = Drive.Files.copy({mimeType:'application/vnd.google-apps.spreadsheet'}, f.getId());
    ssId = converted.id;
  } else {
    ssId = f.getId();
  }
  var ss = SpreadsheetApp.openById(ssId);
  var rows = _parse(ss, fname);
  Logger.log('_parse returned rows=' + rows.length);
  _appendRows(rows);
  _appendLog(fname, f.getId(), rows.length, 'OK', '');
  sp.setProperty('parsed_fname_' + fname, 'true');
}

function _parse(ss, fname){
  var m = fname.match(/(\d{1,2})\s*店.*?(\d{6})/);
  if(!m){
    _appendLog(fname, '', 0, 'SKIPPED_BAD_FILENAME', 'no store/date in filename');
    return [];
  }
  var storeNum = parseInt(m[1], 10);
  var yyyymm   = m[2];
  var year     = parseInt(yyyymm.substring(0,4), 10);
  var month    = parseInt(yyyymm.substring(4,6), 10);
  Logger.log('store=' + storeNum + ' year=' + year + ' month=' + month);

  var storeMeta = _getStoreMeta(storeNum);

  var sheets = ss.getSheets();
  Logger.log('sheet count=' + sheets.length);

  var schedSheet = null;
  for(var si=0; si<sheets.length; si++){
    if(sheets[si].getName() === '班表'){
      schedSheet = sheets[si];
      Logger.log('found 班表 at index ' + si);
      break;
    }
  }
  if(!schedSheet){
    Logger.log('no 班表 sheet, using sheets[0]: ' + sheets[0].getName());
    schedSheet = sheets[0];
  }

  var data = schedSheet.getDataRange().getValues();
  Logger.log('data rows=' + data.length + ' cols=' + (data[0]?data[0].length:0));
  if(data.length < 3){ Logger.log('too few rows, skip'); return []; }

  var dateRowIdx = -1;
  var colDateMap = {};
  for(var rIdx=0; rIdx<Math.min(5, data.length); rIdx++){
    var tryMap = {};
    var row = data[rIdx];
    for(var c=1; c<row.length; c++){
      var v = row[c];
      var ds = _toDateStr(v, year, month);
      if(ds) tryMap[c] = ds;
    }
    if(Object.keys(tryMap).length >= 20){
      dateRowIdx = rIdx;
      colDateMap = tryMap;
      Logger.log('date row found at index ' + rIdx + ', entries=' + Object.keys(tryMap).length);
      break;
    }
  }
  if(dateRowIdx === -1){
    Logger.log('ERROR: cannot find date row in first 5 rows');
    for(var dr=0; dr<Math.min(3, data.length); dr++){
      Logger.log('row ' + dr + ' sample: ' + JSON.stringify(data[dr].slice(0,5)));
    }
    return [];
  }

  var holidaySet = _getHolidaySet();

  var rows = [];
  for(var r=dateRowIdx+1; r<data.length; r++){
    var nameCell = String(data[r][0]).trim();
    if(!nameCell || nameCell === 'undefined') continue;

    var empId = '', empName = '', role = '';
    var idMatch = nameCell.match(/^(\d+)-(.+)$/);
    if(idMatch){
      empId = idMatch[1];
      var rest = idMatch[2].trim();
      var sp2 = rest.split(/\s+/);
      if(sp2.length >= 2){
        role = sp2[sp2.length-1];
        empName = sp2.slice(0, sp2.length-1).join('');
      } else {
        empName = rest;
      }
    } else {
      empName = nameCell;
    }
    if(!empName) continue;

    var declared = _findDeclaredHours(data[r]);

    var computedTotal = 0;
    var dailyRows = [];
    for(var c in colDateMap){
      var cInt = parseInt(c, 10);
      var shiftRaw = String(data[r][cInt]).trim();
      if(!shiftRaw || shiftRaw === 'undefined' || shiftRaw === '') continue;
      if(shiftRaw === '月休' || shiftRaw === '國定假日' || shiftRaw === '未到職') continue;

      var segs = shiftRaw.split(/[\n\r]+/).map(function(s){return s.trim();}).filter(function(s){return s;});
      var validSegs = [];
      var hTotal=0, hMorn=0, hAft=0, hEve=0;
      for(var si2=0; si2<segs.length; si2++){
        var parsed = _parseShiftSeg(segs[si2]);
        if(!parsed) continue;
        validSegs.push(parsed.label);
        hTotal += parsed.total;
        hMorn  += parsed.morning;
        hAft   += parsed.afternoon;
        hEve   += parsed.evening;
      }
      if(validSegs.length === 0) continue;

      computedTotal += hTotal;
      var dateStr = colDateMap[c];
      var dow = new Date(dateStr + 'T00:00:00').getDay();
      var isWeekend = (dow === 0 || dow === 6) ? 1 : 0;
      var isHoliday = holidaySet[dateStr] ? 1 : 0;

      dailyRows.push([
        dateStr, storeMeta.id, storeMeta.name, empId, empName, role,
        validSegs.join('|'), _round2(hTotal),
        _round2(hMorn), _round2(hAft), _round2(hEve),
        isWeekend, isHoliday, shiftRaw, '', fname
      ]);
    }

    var warning = 'OK';
    if(declared !== null && Math.abs(computedTotal - declared) > 0.5){
      warning = 'SUM_MISMATCH:declared=' + declared + ',computed=' + _round2(computedTotal);
    }
    for(var dri=0; dri<dailyRows.length; dri++){
      dailyRows[dri][14] = warning;
      rows.push(dailyRows[dri]);
    }
  }

  Logger.log('total rows extracted=' + rows.length);
  return rows;
}

function _toDateStr(v, fallbackYear, fallbackMonth){
  if(v === null || v === undefined || v === '') return null;
  if(v instanceof Date && !isNaN(v)){
    return _fmtDate(v.getFullYear(), v.getMonth()+1, v.getDate());
  }
  var s = String(v).trim();
  if(!s) return null;
  var m1 = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if(m1) return _fmtDate(parseInt(m1[1],10), parseInt(m1[2],10), parseInt(m1[3],10));
  var m2 = s.match(/^(\d{1,2})[\/\-](\d{1,2})$/);
  if(m2) return _fmtDate(fallbackYear, parseInt(m2[1],10), parseInt(m2[2],10));
  var n = parseFloat(s);
  if(!isNaN(n) && n>=1 && n<=31 && Math.floor(n)===n){
    return _fmtDate(fallbackYear, fallbackMonth, n);
  }
  return null;
}

function _fmtDate(y, mo, d){
  return y + '-' + (mo<10?'0'+mo:String(mo)) + '-' + (d<10?'0'+d:String(d));
}

function _parseShiftSeg(seg){
  var m = seg.match(/^(\d{3,4})\s*-\s*(\d{3,4})$/);
  if(!m) return null;
  var s = _toMinutes(m[1]);
  var e = _toMinutes(m[2]);
  if(s===null || e===null || e<=s) return null;
  return {
    label: m[1] + '-' + m[2],
    total:     (e - s) / 60,
    morning:   _overlap(s, e, 9*60, 13*60) / 60,
    afternoon: _overlap(s, e, 13*60, 18*60) / 60,
    evening:   _overlap(s, e, 18*60, 23*60) / 60
  };
}

function _toMinutes(hhmm){
  var s = String(hhmm);
  if(s.length === 3) s = '0' + s;
  if(s.length !== 4) return null;
  var h = parseInt(s.substring(0,2), 10);
  var m = parseInt(s.substring(2,4), 10);
  if(isNaN(h)||isNaN(m)) return null;
  return h*60 + m;
}

function _overlap(a1, a2, b1, b2){
  return Math.max(0, Math.min(a2,b2) - Math.max(a1,b1));
}

function _round2(n){ return Math.round(n*100)/100; }

function _findDeclaredHours(row){
  for(var c=row.length-1; c>=20; c--){
    var v = row[c];
    if(typeof v === 'number' && v>=10 && v<=400) return v;
    var s = String(v).trim();
    var n = parseFloat(s);
    if(!isNaN(n) && n>=10 && n<=400) return n;
  }
  return null;
}

function _getStoreMeta(storeNum){
  try{
    var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
    var sh = ss.getSheetByName('dim_store_zone');
    var data = sh.getDataRange().getValues();
    for(var r=1; r<data.length; r++){
      if(parseInt(data[r][0],10) === storeNum){
        return {id: data[r][0], name: data[r][1] || ('店'+storeNum), zone: data[r][2]||''};
      }
    }
  }catch(e){ Logger.log('storeMeta err: '+e.message); }
  return {id: storeNum, name: '店'+storeNum, zone: ''};
}

function _getHolidaySet(){
  var set = {};
  try{
    var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
    var sh = ss.getSheetByName('dim_calendar');
    if(!sh) return set;
    var data = sh.getDataRange().getValues();
    for(var r=1; r<data.length; r++){
      var d = data[r][0];
      var h = data[r][1];
      if(!d) continue;
      var ds = (d instanceof Date) ? _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate()) : String(d).trim();
      if(h === true || h === 1 || String(h).trim() === '1' || String(h).toLowerCase() === 'true'){
        set[ds] = true;
      }
    }
  }catch(e){ Logger.log('holidaySet err: '+e.message); }
  return set;
}

function _appendRows(rows){
  if(!rows || rows.length === 0) return;
  var ss  = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh  = ss.getSheetByName('fact_schedule');
  if(!sh){ Logger.log('ERROR: fact_schedule not found'); return; }
  sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
  Logger.log('wrote ' + rows.length + ' rows to fact_schedule');
}

function _appendLog(fname, fid, rowsWritten, status, note){
  try{
    var ss  = SpreadsheetApp.openById(ANALYSIS_SS_ID);
    var sh  = ss.getSheetByName('import_log');
    if(!sh) return;
    sh.appendRow([new Date(), fname, fid, rowsWritten, status, note]);
  } catch(e){
    Logger.log('_appendLog error: ' + e.message);
  }
}

// ====================================================================
// 工具函式:清空 / 驗證
// ====================================================================

function clearProcessed(){
  PropertiesService.getScriptProperties().deleteAllProperties();
  Logger.log('cleared all script properties');
}

function resetFactSchedule(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  if(!sh){ Logger.log('ERROR: fact_schedule not found'); return; }
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if(lastRow > 1){
    sh.getRange(2, 1, lastRow-1, lastCol).clearContent();
    Logger.log('cleared ' + (lastRow-1) + ' rows from fact_schedule');
  } else {
    Logger.log('fact_schedule already empty');
  }
  var sp = PropertiesService.getScriptProperties();
  var keys = sp.getKeys();
  var cleared = 0;
  for(var i=0; i<keys.length; i++){
    if(keys[i].indexOf('parsed_fname_') === 0 || keys[i].indexOf('processed_') === 0){
      sp.deleteProperty(keys[i]);
      cleared++;
    }
  }
  Logger.log('cleared ' + cleared + ' ScriptProperties keys');
}

function nukeFactSchedule(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  var lastRow = sh.getLastRow();
  Logger.log('清空前: ' + lastRow + ' 列');
  if(lastRow > 1){
    sh.deleteRows(2, lastRow - 1);
  }
  PropertiesService.getScriptProperties().deleteAllProperties();
  SpreadsheetApp.flush();
  Utilities.sleep(2000);
  Logger.log('清空後: ' + sh.getLastRow() + ' 列（應該是 1）');
}

function verifyFactSchedule(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var idxStore = header.indexOf('store_id');
  var idxDate  = header.indexOf('date');
  var idxFile  = header.indexOf('file_name');

  var storeMonth = {};
  var fileCount = {};

  for(var i = 1; i < data.length; i++){
    var sid = String(data[i][idxStore]);
    var d = data[i][idxDate];
    var ym = (d instanceof Date)
      ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM')
      : String(d).substring(0,7);
    var fname = (idxFile >= 0) ? String(data[i][idxFile]||'').trim() : '';

    var key = 'store=' + (sid.length<2?'0'+sid:sid) + ' | ' + ym;
    storeMonth[key] = (storeMonth[key] || 0) + 1;
    if(fname) fileCount[fname] = (fileCount[fname] || 0) + 1;
  }

  Logger.log('=== fact_schedule store × 年月 ===');
  Object.keys(storeMonth).sort().forEach(function(k){ Logger.log(k + ': ' + storeMonth[k] + ' 列'); });

  Logger.log('\n=== file_name 列數分佈（每檔應該 1 份 70-130 列）===');
  var fileList = Object.keys(fileCount).sort();
  fileList.forEach(function(k){
    var v = fileCount[k];
    var flag = (v > 150 || v < 30) ? ' ⚠️ 異常' : '';
    Logger.log(k + ': ' + v + ' 列' + flag);
  });

  Logger.log('\n總列數: ' + (data.length - 1));
  Logger.log('不同檔名數: ' + fileList.length);
}

// ====================================================================
// Phase 2: 日營收同步
// ====================================================================
var POS_SS_ID_   = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';
var POS_SHEET_   = 'POS資料';

var POS_COL_STORE_ID = 0;
var POS_COL_DATE     = 2;
var POS_COL_CAT1     = 4;
var POS_COL_CAT2     = 5;
var POS_COL_QTY      = 7;
var POS_COL_AMOUNT   = 8;

var EXCLUDE_CAT1_ = {'加價購': true, '找不到': true};
var EXCLUDE_CAT2_ = {'陪同入場費': true};

function syncDailySales(){
  Logger.log('=== syncDailySales START ===');
  var monthsBack = 3;
  var cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - monthsBack);
  cutoff.setDate(1);
  cutoff.setHours(0,0,0,0);
  Logger.log('cutoff date = ' + _fmtDate(cutoff.getFullYear(), cutoff.getMonth()+1, cutoff.getDate()));

  var posSS = SpreadsheetApp.openById(POS_SS_ID_);
  var posSh = posSS.getSheetByName(POS_SHEET_);
  if(!posSh){ Logger.log('ERROR: POS 分頁找不到'); return; }
  var lastRow = posSh.getLastRow();
  var lastCol = posSh.getLastColumn();
  Logger.log('POS total rows=' + lastRow + ' cols=' + lastCol);

  var BATCH = 50000;
  var agg = {};
  var skipped = 0;

  for(var startRow = 2; startRow <= lastRow; startRow += BATCH){
    var rowsToRead = Math.min(BATCH, lastRow - startRow + 1);
    var data = posSh.getRange(startRow, 1, rowsToRead, lastCol).getValues();
    Logger.log('batch start=' + startRow + ' read=' + data.length);

    for(var r=0; r<data.length; r++){
      var row = data[r];
      var dateV = row[POS_COL_DATE];
      if(!dateV) { skipped++; continue; }

      var d;
      if(dateV instanceof Date) d = dateV;
      else { d = new Date(String(dateV)); if(isNaN(d)){ skipped++; continue; } }
      if(d < cutoff){ skipped++; continue; }

      var dateStr = _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate());
      var storeId = parseInt(row[POS_COL_STORE_ID], 10);
      if(isNaN(storeId)){ skipped++; continue; }

      var amount = parseFloat(row[POS_COL_AMOUNT]) || 0;
      var cat1   = String(row[POS_COL_CAT1] || '').trim();
      var cat2   = String(row[POS_COL_CAT2] || '').trim();

      var key = dateStr + '|' + storeId;
      if(!agg[key]) agg[key] = {revenue:0, headcount:0, rows_total:0};
      agg[key].revenue += amount;
      agg[key].rows_total += 1;

      if(!EXCLUDE_CAT1_[cat1] && !EXCLUDE_CAT2_[cat2]){
        agg[key].headcount += 1;
      }
    }
  }
  Logger.log('aggregated keys=' + Object.keys(agg).length + ' skipped=' + skipped);

  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_daily_sales');
  if(!sh){ Logger.log('ERROR: fact_daily_sales 分頁找不到'); return; }

  var existing = sh.getDataRange().getValues();
  var keyToRow = {};
  for(var er=1; er<existing.length; er++){
    var ed = existing[er][0];
    var es = existing[er][1];
    if(!ed || !es) continue;
    var edStr = (ed instanceof Date) ? _fmtDate(ed.getFullYear(), ed.getMonth()+1, ed.getDate()) : String(ed).trim();
    keyToRow[edStr + '|' + parseInt(es,10)] = er + 1;
  }

  var storeNames = {};
  var dimSh = ss.getSheetByName('dim_store_zone');
  if(dimSh){
    var dimData = dimSh.getDataRange().getValues();
    for(var dr=1; dr<dimData.length; dr++){
      storeNames[parseInt(dimData[dr][0],10)] = dimData[dr][1] || '';
    }
  }

  var now = new Date();
  var updates = [];
  var appends = [];
  for(var k in agg){
    var parts = k.split('|');
    var dateStr2 = parts[0];
    var sid     = parseInt(parts[1], 10);
    var a = agg[k];
    var rph = a.headcount > 0 ? _round2(a.revenue / a.headcount) : 0;
    var rowVals = [dateStr2, sid, storeNames[sid]||'', _round2(a.revenue), a.headcount, a.rows_total, rph, now];
    if(keyToRow[k]){
      updates.push([keyToRow[k], rowVals]);
    } else {
      appends.push(rowVals);
    }
  }
  Logger.log('updates=' + updates.length + ' appends=' + appends.length);

  for(var u=0; u<updates.length; u++){
    sh.getRange(updates[u][0], 1, 1, 8).setValues([updates[u][1]]);
  }
  if(appends.length > 0){
    sh.getRange(sh.getLastRow()+1, 1, appends.length, 8).setValues(appends);
  }

  _appendLog('syncDailySales', '', updates.length+appends.length, 'OK',
    'cutoff=' + _fmtDate(cutoff.getFullYear(), cutoff.getMonth()+1, cutoff.getDate()) +
    ' updates=' + updates.length + ' appends=' + appends.length + ' skipped=' + skipped);
  Logger.log('=== syncDailySales DONE ===');
}

// ====================================================================
// Phase 3: 週度 KPI 聚合
// ====================================================================

var FLAG_WEEKDAY_HOURS_ = 16;
var FLAG_WEEKEND_HOURS_ = 24;
var FLAG_REV_PER_HOUR_  = 500;

var MORNING_START_ = 9 * 60;
var MORNING_END_   = 13 * 60;
var EVENING_START_ = 18 * 60;
var EVENING_END_   = 23 * 60;

function buildWeekKPI(){
  Logger.log('=== buildWeekKPI START ===');
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);

  var dimSh = ss.getSheetByName('dim_store_zone');
  if(!dimSh){ Logger.log('ERROR: dim_store_zone 找不到'); return; }
  var dimData = dimSh.getDataRange().getValues();
  var storeMeta = {};
  for(var i=1; i<dimData.length; i++){
    var sid = parseInt(dimData[i][0], 10);
    if(isNaN(sid)) continue;
    storeMeta[sid] = {
      name:    dimData[i][1] || '',
      zone:    dimData[i][2] || '',
      manager: dimData[i][2] || ''
    };
  }
  Logger.log('storeMeta loaded=' + Object.keys(storeMeta).length);

  var schSh = ss.getSheetByName('fact_schedule');
  if(!schSh){ Logger.log('ERROR: fact_schedule 找不到'); return; }
  var schData = schSh.getDataRange().getValues();
  var schHeader = schData[0];
  var idx = _buildHeaderIdx(schHeader);
  Logger.log('fact_schedule rows=' + (schData.length-1));

  var salSh = ss.getSheetByName('fact_daily_sales');
  if(!salSh){ Logger.log('ERROR: fact_daily_sales 找不到'); return; }
  var salData = salSh.getDataRange().getValues();
  var salMap = {};
  for(var s=1; s<salData.length; s++){
    var sd = salData[s][0];
    var ss_ = parseInt(salData[s][1], 10);
    var rev = parseFloat(salData[s][3]) || 0;
    if(!sd || isNaN(ss_)) continue;
    var sdStr = (sd instanceof Date) ? _fmtDate(sd.getFullYear(), sd.getMonth()+1, sd.getDate()) : String(sd).trim();
    salMap[sdStr + '|' + ss_] = rev;
  }
  Logger.log('salMap keys=' + Object.keys(salMap).length);

  var weekAgg = {};
  var dayShifts = {};

  for(var r=1; r<schData.length; r++){
    var row = schData[r];
    var dateV = row[idx['date']];
    if(!dateV) continue;
    var d = (dateV instanceof Date) ? dateV : new Date(String(dateV));
    if(isNaN(d)) continue;
    var dateStr = _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate());
    var sid2 = parseInt(row[idx['store_id']], 10);
    if(isNaN(sid2)) continue;
    var hours = parseFloat(row[idx['hours_total']]) || 0;
    var isWeekend = parseInt(row[idx['is_weekend']], 10) === 1;
    var isHoliday = parseInt(row[idx['is_holiday']], 10) === 1;
    var isOff = isWeekend || isHoliday;
    var segStr = String(row[idx['shift_segments']] || '').trim();

    var weekStart = _getMondayOfWeek(d);
    var weekKey = _fmtDate(weekStart.getFullYear(), weekStart.getMonth()+1, weekStart.getDate()) + '|' + sid2;

    if(!weekAgg[weekKey]){
      weekAgg[weekKey] = {
        weekStart: weekStart, store_id: sid2,
        total_hours:0, weekday_hours:0, weekend_hours:0,
        weekday_revenue:0, weekend_revenue:0,
        flag_weekday_over:false, flag_weekend_over:false,
        seenDates: {}
      };
    }
    var W = weekAgg[weekKey];
    W.total_hours += hours;
    if(isOff){
      W.weekend_hours += hours;
      if(hours > FLAG_WEEKEND_HOURS_) W.flag_weekend_over = true;
    } else {
      W.weekday_hours += hours;
      if(hours > FLAG_WEEKDAY_HOURS_) W.flag_weekday_over = true;
    }

    if(!W.seenDates[dateStr]) W.seenDates[dateStr] = isOff ? 'off' : 'wd';

    var dayKey = dateStr + '|' + sid2;
    if(!dayShifts[dayKey]) dayShifts[dayKey] = {segs:[], isOff: isOff};
    if(segStr){
      dayShifts[dayKey].segs.push(_parseSegments(segStr));
    }
  }

  var soloMorning = {};
  var soloEvening = {};
  for(var dk in dayShifts){
    var dParts = dk.split('|');
    var dStr = dParts[0];
    var sidD = parseInt(dParts[1], 10);
    var info = dayShifts[dk];
    if(!info.isOff){
      if(_countActiveEmployees(info.segs, MORNING_START_, MORNING_END_) === 1){
        var wkS = _getMondayOfWeek(_parseYMD(dStr));
        var wKey = _fmtDate(wkS.getFullYear(), wkS.getMonth()+1, wkS.getDate()) + '|' + sidD;
        soloMorning[wKey] = true;
      }
    }
    if(_countActiveEmployees(info.segs, EVENING_START_, EVENING_END_) === 1){
      var wkS2 = _getMondayOfWeek(_parseYMD(dStr));
      var wKey2 = _fmtDate(wkS2.getFullYear(), wkS2.getMonth()+1, wkS2.getDate()) + '|' + sidD;
      soloEvening[wKey2] = true;
    }
  }

  for(var wk in weekAgg){
    var W2 = weekAgg[wk];
    for(var ds in W2.seenDates){
      var rev2 = salMap[ds + '|' + W2.store_id] || 0;
      if(W2.seenDates[ds] === 'off') W2.weekend_revenue += rev2;
      else W2.weekday_revenue += rev2;
    }
  }

  var outSh = ss.getSheetByName('fact_week_kpi');
  if(!outSh){ Logger.log('ERROR: fact_week_kpi 找不到'); return; }
  if(outSh.getLastRow() > 1){
    outSh.getRange(2, 1, outSh.getLastRow()-1, 15).clearContent();
  }

  var now = new Date();
  var outRows = [];
  for(var wk2 in weekAgg){
    var W3 = weekAgg[wk2];
    var meta = storeMeta[W3.store_id] || {name:'', zone:'', manager:''};
    var totalRev = W3.weekday_revenue + W3.weekend_revenue;
    var rph = W3.total_hours > 0 ? _round2(totalRev / W3.total_hours) : 0;
    var flags = [];
    if(W3.flag_weekday_over) flags.push('weekday_over_16');
    if(W3.flag_weekend_over) flags.push('weekend_over_24');
    if(soloMorning[wk2])     flags.push('morning_solo');
    if(soloEvening[wk2])     flags.push('evening_solo');
    if(rph < FLAG_REV_PER_HOUR_ && W3.total_hours > 0) flags.push('low_rev_per_hour');

    var wsDate = W3.weekStart;
    var yw = _getYearWeek(wsDate);
    outRows.push([
      _fmtDate(wsDate.getFullYear(), wsDate.getMonth()+1, wsDate.getDate()),
      yw,
      W3.store_id, meta.name, meta.zone, meta.manager,
      _round2(W3.total_hours), _round2(W3.weekday_hours), _round2(W3.weekend_hours),
      _round2(totalRev), _round2(W3.weekday_revenue), _round2(W3.weekend_revenue),
      rph,
      flags.join(','),
      now
    ]);
  }

  outRows.sort(function(a,b){
    if(a[0] < b[0]) return -1;
    if(a[0] > b[0]) return 1;
    return a[2] - b[2];
  });

  if(outRows.length > 0){
    outSh.getRange(2, 1, outRows.length, 15).setValues(outRows);
  }
  Logger.log('written rows=' + outRows.length);

  _appendLog('buildWeekKPI', '', outRows.length, 'OK', 'weeks×stores=' + outRows.length);
  Logger.log('=== buildWeekKPI DONE ===');
}

// ---- buildWeekKPI helpers ----

function _buildHeaderIdx(headerRow){
  var idx = {};
  for(var i=0; i<headerRow.length; i++){
    idx[String(headerRow[i]).trim()] = i;
  }
  return idx;
}

function _getMondayOfWeek(d){
  var dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var day = dd.getDay();
  var diff = (day === 0) ? -6 : (1 - day);
  dd.setDate(dd.getDate() + diff);
  return dd;
}

function _getYearWeek(d){
  var target = new Date(d.valueOf());
  var dayNr = (d.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  var jan4 = new Date(target.getFullYear(), 0, 4);
  var dayDiff = (target - jan4) / 86400000;
  var weekNum = 1 + Math.ceil((dayDiff - 3) / 7);
  return target.getFullYear() + '-W' + (weekNum < 10 ? '0'+weekNum : weekNum);
}

function _parseYMD(s){
  var p = String(s).split('-');
  return new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10));
}

function _parseSegments(segStr){
  var out = [];
  var parts = segStr.split('|');
  for(var i=0; i<parts.length; i++){
    var seg = parts[i].trim();
    if(seg.length < 9) continue;
    var hyphen = seg.indexOf('-');
    if(hyphen < 0) continue;
    var a = seg.substring(0, hyphen);
    var b = seg.substring(hyphen+1);
    var aMin = _hhmmToMin(a);
    var bMin = _hhmmToMin(b);
    if(aMin >= 0 && bMin > aMin) out.push([aMin, bMin]);
  }
  return out;
}

function _hhmmToMin(s){
  s = String(s).replace(':','').trim();
  if(s.length < 3 || s.length > 4) return -1;
  if(s.length === 3) s = '0' + s;
  var h = parseInt(s.substring(0,2), 10);
  var m = parseInt(s.substring(2,4), 10);
  if(isNaN(h) || isNaN(m)) return -1;
  return h * 60 + m;
}

function _countActiveEmployees(employeeSegs, winStart, winEnd){
  var n = 0;
  for(var e=0; e<employeeSegs.length; e++){
    var segs = employeeSegs[e];
    var active = false;
    for(var s=0; s<segs.length; s++){
      if(segs[s][1] > winStart && segs[s][0] < winEnd){
        active = true; break;
      }
    }
    if(active) n++;
  }
  return n;
}

// ====================================================================
// Phase 4: 儀表板 JSONP API
// ====================================================================

function doGet(e){
  var action = (e && e.parameter && e.parameter.action) || '';
  var callback = (e && e.parameter && e.parameter.callback) || 'callback';
  var result;
  try {
    if(action === 'getWeekKpiData'){
      result = _getWeekKpiData();
    } else if(action === 'getStoreDailyData'){
      var sid = parseInt(e.parameter.store_id, 10);
      var ym  = e.parameter.month || '';
      result = _getStoreDailyData(sid, ym);
    } else if(action === 'ping'){
      result = {ok:true, time:new Date().toISOString(), version:'schedule-v1.2'};
    } else {
      result = {ok:false, error:'unknown action: ' + action};
    }
  } catch(err){
    result = {ok:false, error: String(err)};
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function _getWeekKpiData(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);

  var sh = ss.getSheetByName('fact_week_kpi');
  if(!sh) return {ok:false, error:'fact_week_kpi not found'};
  var data = sh.getDataRange().getValues();
  var rows = [];
  for(var r=1; r<data.length; r++){
    var row = data[r];
    if(!row[0]) continue;
    var d = row[0];
    var dateStr = (d instanceof Date) ? _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate()) : String(d).trim();
    rows.push({
      week_start: dateStr,
      year_week:  row[1],
      store_id:   row[2],
      store_name: row[3],
      zone:       row[4],
      manager:    row[5],
      total_hours:      parseFloat(row[6])  || 0,
      weekday_hours:    parseFloat(row[7])  || 0,
      weekend_hours:    parseFloat(row[8])  || 0,
      total_revenue:    parseFloat(row[9])  || 0,
      weekday_revenue:  parseFloat(row[10]) || 0,
      weekend_revenue:  parseFloat(row[11]) || 0,
      rev_per_hour:     parseFloat(row[12]) || 0,
      flags:            String(row[13] || '').split(',').filter(function(x){return x;})
    });
  }

  var salSh = ss.getSheetByName('fact_daily_sales');
  var daily = [];
  if(salSh){
    var salData = salSh.getDataRange().getValues();
    for(var s=1; s<salData.length; s++){
      var row2 = salData[s];
      if(!row2[0]) continue;
      var sd = row2[0];
      var sdStr = (sd instanceof Date) ? _fmtDate(sd.getFullYear(), sd.getMonth()+1, sd.getDate()) : String(sd).trim();
      daily.push({
        date:      sdStr,
        store_id:  parseInt(row2[1], 10),
        revenue:   parseFloat(row2[3]) || 0,
        headcount: parseInt(row2[4], 10) || 0,
        rev_per_head: parseFloat(row2[6]) || 0
      });
    }
  }

  var dimSh = ss.getSheetByName('dim_store_zone');
  var stores = [];
  if(dimSh){
    var dimData = dimSh.getDataRange().getValues();
    for(var i=1; i<dimData.length; i++){
      var sid = parseInt(dimData[i][0], 10);
      if(isNaN(sid)) continue;
      stores.push({
        store_id:   sid,
        store_name: dimData[i][1] || '',
        zone:       dimData[i][2] || '',
        brand:      dimData[i][3] || ''
      });
    }
  }

  return {
    ok: true,
    generated_at: new Date().toISOString(),
    week_kpi: rows,
    daily_sales: daily,
    stores: stores
  };
}

// ====================================================================
// Phase 4.5 v1.2: 單店日級分析 (含員工列表)
// ====================================================================

function _getStoreDailyData(storeId, ym) {
  if (isNaN(storeId)) return { ok: false, error: 'store_id required' };
  if (!ym || !/^\d{4}-\d{2}$/.test(ym)) return { ok: false, error: 'month required (YYYY-MM)' };

  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);

  var schSh = ss.getSheetByName('fact_schedule');
  if (!schSh) return { ok: false, error: 'fact_schedule not found' };
  var schData = schSh.getDataRange().getValues();
  var idx = _buildHeaderIdx(schData[0]);

  var daily = {};
  for (var r = 1; r < schData.length; r++) {
    var row = schData[r];
    var sid = parseInt(row[idx['store_id']], 10);
    if (sid !== storeId) continue;
    var dateV = row[idx['date']];
    if (!dateV) continue;
    var d = (dateV instanceof Date) ? dateV : new Date(String(dateV));
    if (isNaN(d)) continue;
    var dStr = _fmtDate(d.getFullYear(), d.getMonth() + 1, d.getDate());
    if (dStr.substring(0, 7) !== ym) continue;

    if (!daily[dStr]) {
      daily[dStr] = {
        date: dStr,
        hours: 0,
        is_weekend: parseInt(row[idx['is_weekend']], 10) === 1,
        is_holiday: parseInt(row[idx['is_holiday']], 10) === 1,
        employees: [],
        weekday_num: d.getDay()
      };
    }
    var empHours = parseFloat(row[idx['hours_total']]) || 0;
    daily[dStr].hours += empHours;

    // 收集員工資訊
    var empName = String(row[idx['emp_name']] || '').trim();
    var shiftSegs = String(row[idx['shift_segments']] || '').trim();
    if (empName) {
      var label = empName + '(' + _round2(empHours) + 'h) ' + shiftSegs;
      daily[dStr].employees.push({
        name: empName,
        hours: _round2(empHours),
        shift: shiftSegs,
        label: label
      });
    }
  }

  var salSh = ss.getSheetByName('fact_daily_sales');
  if (salSh) {
    var salData = salSh.getDataRange().getValues();
    for (var s = 1; s < salData.length; s++) {
      var rr = salData[s];
      var sid2 = parseInt(rr[1], 10);
      if (sid2 !== storeId) continue;
      var sd = rr[0];
      var sdStr = (sd instanceof Date) ? _fmtDate(sd.getFullYear(), sd.getMonth() + 1, sd.getDate()) : String(sd).trim();
      if (sdStr.substring(0, 7) !== ym) continue;
      if (!daily[sdStr]) {
        var dd = _parseYMD(sdStr);
        daily[sdStr] = {
          date: sdStr, hours: 0, is_weekend: (dd.getDay() === 0 || dd.getDay() === 6),
          is_holiday: false, employees: [], weekday_num: dd.getDay()
        };
      }
      daily[sdStr].revenue   = parseFloat(rr[3]) || 0;
      daily[sdStr].headcount = parseInt(rr[4], 10) || 0;
      daily[sdStr].rev_per_head = parseFloat(rr[6]) || 0;
    }
  }

  var rows = Object.keys(daily).sort().map(function (k) {
    var d = daily[k];
    if (d.hours > 0 && d.revenue > 0) {
      d.rev_per_hour = _round2(d.revenue / d.hours);
    } else {
      d.rev_per_hour = 0;
    }
    d.revenue = d.revenue || 0;
    d.headcount = d.headcount || 0;
    d.rev_per_head = d.rev_per_head || 0;
    d.employees = d.employees || [];
    // 員工依工時降序排
    d.employees.sort(function(a,b){ return b.hours - a.hours; });
    return d;
  });

  var validRows = rows.filter(function (r) { return r.hours > 0; });
  var monthAvgHours   = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.hours; }, 0) / validRows.length : 0;
  var monthAvgRevenue = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.revenue; }, 0) / validRows.length : 0;
  var monthAvgRph     = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.rev_per_hour; }, 0) / validRows.length : 0;

  var storeMeta = null;
  var dimSh = ss.getSheetByName('dim_store_zone');
  if (dimSh) {
    var dimData = dimSh.getDataRange().getValues();
    for (var i = 1; i < dimData.length; i++) {
      if (parseInt(dimData[i][0], 10) === storeId) {
        storeMeta = {
          store_id: storeId,
          store_name: dimData[i][1] || '',
          zone:       dimData[i][2] || '',
          brand:      dimData[i][3] || ''
        };
        break;
      }
    }
  }

  return {
    ok: true,
    generated_at: new Date().toISOString(),
    store: storeMeta,
    month: ym,
    daily: rows,
    baseline: {
      avg_hours:   _round2(monthAvgHours),
      avg_revenue: _round2(monthAvgRevenue),
      avg_rph:     _round2(monthAvgRph),
      valid_days:  validRows.length
    }
  };
}

// ====================================================================
// 健康檢查
// ====================================================================

function healthCheck(){
  Logger.log('=== healthCheck START ===');
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  var data = sh.getDataRange().getValues();
  var header = data[0];
  var idx = _buildHeaderIdx(header);

  var storeRows = {};
  var storeDates = {};

  for(var r=1; r<data.length; r++){
    var row = data[r];
    var sid = parseInt(row[idx['store_id']], 10);
    if(isNaN(sid)) continue;
    storeRows[sid] = (storeRows[sid]||0) + 1;
    var dateV = row[idx['date']];
    if(dateV){
      var d = (dateV instanceof Date) ? dateV : new Date(String(dateV));
      if(!isNaN(d)){
        var ds = _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate());
        if(!storeDates[sid]) storeDates[sid] = {};
        storeDates[sid][ds] = true;
      }
    }
  }
  Logger.log('===== 各店資料分佈 =====');
  for(var s=1; s<=12; s++){
    var rc = storeRows[s] || 0;
    var dc = storeDates[s] ? Object.keys(storeDates[s]).length : 0;
    Logger.log('store_id='+s+' rows='+rc+' distinct_dates='+dc);
  }
  Logger.log('total rows = ' + (data.length-1));
  Logger.log('=== healthCheck DONE ===');
}

// ====================================================================
// Google 評論補欄 enrichColumns v2
// 2026-06-30 修改：G 欄(年月)只在空白時才補入 createTime 年月，
//   避免覆蓋 Make Module 7 寫入的 updateTime 年月（編輯過的評論）。
//   H/I/J/K 欄維持每次都更新。
// ====================================================================

var REVIEW_SHEET_ID   = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
var REVIEW_SHEET_NAME = '工作表1';

var REVIEW_BRAND_MAP = {
  '台北南京店':          '自己做',
  '台北士林店':          '自己做',
  '新北新店店':          '自己做',
  '新北板橋店':          '自己做',
  '桃園中壢店':          '自己做',
  '桃園藝文店':          '自己做',
  '台中精明店':          '自己做',
  '台中草悟道店':        '自己做',
  '台南focus店':         '自己做',
  '新竹文化店':          '自己做',
  '吳寶春(台北信義A13)': '吳寶春自己做',
  '吳寶春(高雄SKM)':    '吳寶春自己做'
};

var REVIEW_REGION_MAP = {
  '台北南京店':          '北一區',
  '台北士林店':          '北一區',
  '新北新店店':          '北一區',
  '新北板橋店':          '北二區',
  '吳寶春(台北信義A13)': '北二區',
  '桃園中壢店':          '北三區',
  '桃園藝文店':          '北三區',
  '台中精明店':          '中南區',
  '台中草悟道店':        '中南區',
  '台南focus店':         '中南區',
  '新竹文化店':          '中南區',
  '吳寶春(高雄SKM)':    '中南區'
};

var REVIEW_STAR_MAP = { ONE:1, TWO:2, THREE:3, FOUR:4, FIVE:5 };

function enrichColumns() {
  var ss    = SpreadsheetApp.openById(REVIEW_SHEET_ID);
  var sheet = ss.getSheetByName(REVIEW_SHEET_NAME);
  var last  = sheet.getLastRow();
  if (last < 2) return;

  var data = sheet.getRange(2, 1, last - 1, 11).getValues();

  for (var i = 0; i < data.length; i++) {
    var row   = data[i];
    var store = row[0];  // A 門市
    var cDate = row[2];  // C 評論時間(createTime)
    var star  = row[3];  // D 星等(ONE/TWO…)
    var gVal  = row[6];  // G 年月(現有值)

    if (!store) continue; // 空列跳過

    // ── G 年月：有值就不蓋（保護 Make Module 7 寫入的 updateTime 年月）──
    if (!gVal || gVal === '') {
      if (cDate instanceof Date) {
        var y = cDate.getFullYear();
        var mo = String(cDate.getMonth() + 1);
        if (mo.length < 2) mo = '0' + mo;
        row[6] = y + '-' + mo;
      } else if (typeof cDate === 'string' && cDate.length >= 7) {
        row[6] = cDate.substring(0, 7);
      }
    }

    // ── H 星等數字：每次都更新 ──
    var starNum = REVIEW_STAR_MAP[String(star).trim()] || null;
    if (starNum !== null) row[7] = starNum;

    // ── I 品牌：每次都更新 ──
    row[8] = REVIEW_BRAND_MAP[store] || '';

    // ── J 區域：每次都更新 ──
    row[9] = REVIEW_REGION_MAP[store] || '';

    // ── K 負評標記：每次都更新（依 H 欄）──
    var h = row[7];
    if (h !== null && h !== '') {
      row[10] = (h >= 5) ? '達標' : '未達5星';
    }
  }

  sheet.getRange(2, 1, data.length, 11).setValues(data);
  Logger.log('enrichColumns v2 done: ' + data.length + ' rows');
}