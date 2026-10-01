// ====================================================================
// 排班分析資料倉 程式碼.gs v5.4
// 更新日期:2026-07-01
// 版本變更:
//   v5.5: 【dailyRebuild 加髒日期偵測】(2026-07-13)
//         - 先查「近 26 小時被重寫、但 sale_date 在 3 天視窗外」的日期，逐日補 syncDailySales(d,d)。
//         - 查詢用既有 _getBqAccessToken_()（bq_connector.gs），失敗 try/catch 吞掉照跑 3 天窗。
//         - buildWeekKPI 仍只跑一次；無髒日期時行為與 v5.4 完全相同。
//         - 背景:POS→BQ 每日自癒(dailySyncBigQuery_v2)會回寫視窗外舊日期，本函式負責下游跟進。
//   v5.4: 【修 baseline 月均被未來 0 營收天稀釋】(2026-07-01)
//         - _getStoreDailyData 的 validRows 篩選由 hours>0 改為 hours>0 && revenue>0。
//         - 起因:月份下拉開放未來已排班月份(如 7 月)後,整月 hours>0 但多數天 revenue=0,
//           三個月均(工時/營收/rev-h)分母含 0 營收天 → 月均營收/rev-h 被稀釋到接近 $0。
//         - 已完整結束月份的營業日都同時 hours>0 且 revenue>0 → 對舊月份數字零影響。
//         - 純後端 1 行改動,前端零改動(valid_days=0 顯示美化為前端可選項,另計)。
//         - ping version 同步升為 schedule-v5.4,供部署後驗證。
//   v5.3: 【processUploads 改增量，修每次上傳等 70+ 秒】
//         - _processUploads 的 syncDailySales() 補傳 90 天視窗參數
//         - v5.2 已修 dailyRebuild，但 processUploads 漏修；v5.3 補齊
//   v5.2: 【dailyRebuild 改增量，修 6 分鐘超時】
//         - syncDailySales(startDate, endDate) 加可選日期參數
//         - 範圍模式：queryDailyNet 傳日期、POS 掃描 skip 範圍外列、setValues 只跑 ~36 筆
//         - dailyRebuild_salesAndKpi → 滾動 3 天增量（sync 限範圍，build 仍全期）
//         - 新增 dailyRebuild_salesAndKpi_full 保留全期重算路徑
//   v5.1: 【修正日曆/明細表 6/4 營收顯示 $0】
//         - _getStoreDailyData 的 daily[sdStr].revenue 由讀 gross(rr[3], 第4欄)
//           改讀 net(rr[8], 第9欄)，與分店比較(getWeekKpiData/buildWeekKPI)口徑統一。
//         - 根因:v5.0 後 gross 來自 v_daily_net，當日 gross 可能尚未進 view(=0)，
//           但 net 已有值 → 日曆讀 gross 得 0、分店比較讀 net 有值，造成同日不一致。
//         - 純後端 1 行改動，前端零改動。rev_per_hour / baseline 自動跟隨 net。
//   v5.0: 【金額單一真相 — 全取自 BigQuery v_daily_net】
//         - gross / net 不再從 POS資料Sheet 算、不再讀 fact_discount 分頁，
//           改從 v_daily_net view 取（店×日 gross/discount/net）。
//         - headcount 仍從 POS資料Sheet 掃（view 無此粒度），口徑不變。
//         - 需搭配同專案 bq_connector.gs（queryDailyNet 已測通）。
//         - fact_daily_sales 寫入格式、欄位數(9欄)不變 → buildWeekKPI、前端零改動。
//         驗證：1店2026-04 view gross=393,258 net=369,921（=肚肚 POS 對帳）✓
//   v4.2: (前版) net = Σ(G×H) − fact_discount(C欄)，金額從 Sheet 算。已被 v5.0 取代金額來源。
// ====================================================================
var ANALYSIS_SS_ID = '19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM';
var FOLDER_ID_     = '1yNM8RkYIfdvmETpK9mPbydcn2DXYQYaz';
var NOTIFY_EMAIL_  = 'mydiybc@gmail.com';

// ====================================================================
// Phase 1: 班表 Parser
// ====================================================================

function parsePendingFiles(){
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('parsePendingFiles: 已有處理中,跳過本次');
    return;
  }
  try {
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

    var ck = _canonicalKey(fname);

    // ★ 2026-09-11 架構變更：清空舊資料從「上傳當下」移到「解析當下」。
    //   pending_ck_* 標記代表這個 canonical 剛被重新上傳,必須覆蓋既有資料。
    //   舊架構:_uploadSchedule 先 purge → 之後才 parse。中間若中斷(處理鎖擋住、
    //          店長關視窗、逾時),該店該月就停在「已清空、未寫入」的空白狀態。
    //   新架構:purge 與 write 綁在同一支函式、同一次執行內完成。
    //          上傳只是把檔案放進 Drive,解析沒跑就維持舊資料,絕不開天窗。
    var pendingKey = ck ? ('pending_ck_' + ck) : '';
    var isPending  = pendingKey ? (sp.getProperty(pendingKey) === '1') : false;

    if(!isPending){
      if(sp.getProperty('parsed_fname_' + fname) === 'true'){
        Logger.log('SKIP: already processed (by fname property)');
        continue;
      }

      if(existingFiles[fname]){
        Logger.log('SKIP: file_name already in fact_schedule (' + existingFiles[fname] + ' rows)');
        sp.setProperty('parsed_fname_' + fname, 'true');
        continue;
      }

      if(ck && existingFiles['__canonical__' + ck]){
        Logger.log('SKIP: canonical key ' + ck + ' already in fact_schedule');
        sp.setProperty('parsed_fname_' + fname, 'true');
        continue;
      }
    } else {
      // 重新上傳:先清掉同 canonical 的舊資料,緊接著寫新的
      var pParts = ck.split('-');
      var pStore = parseInt(pParts[0], 10);
      var pYm    = pParts[1];
      var pRows  = _purgeFactScheduleByCanonical(pStore, pYm);
      var pNotes = _purgeFactNoteByCanonical(pStore, pYm);
      Logger.log('RE-UPLOAD ' + ck + ': 先清 fact_schedule ' + pRows + ' 列 / fact_note ' + pNotes + ' 列');
      delete existingFiles[fname];
      delete existingFiles['__canonical__' + ck];
    }

    _processFile(f, isXls, sp, fname);
    // 寫入成功才清 pending 標記;若上一行拋錯,標記留著,下一輪(或排程)會自動重試
    if(pendingKey) sp.deleteProperty(pendingKey);
    existingFiles[fname] = 'just_written';
    if(ck) existingFiles['__canonical__' + ck] = 'just_written';
  }
  Logger.log('=== parsePendingFiles DONE ===');
  } finally {
    lock.releaseLock();
  }
}

function _canonicalKey(fname){
  if(!fname) return '';
  var m = String(fname).match(/(\d{1,2})\s*店.*?(\d{6})/);
  if(!m) return '';
  return parseInt(m[1], 10) + '-' + m[2];
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
      if(!fn) continue;
      result[fn] = (result[fn]||0) + 1;
      var ck = _canonicalKey(fn);
      if(ck) result['__canonical__' + ck] = (result['__canonical__' + ck]||0) + 1;
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
  var noteMap = _buildNoteMap(ss, year, month);

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
        isWeekend, isHoliday, shiftRaw, '', fname, _noteFor(noteMap, dateStr, empId, empName)
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

  // ★ 2026-09-10 新增：同步產出 fact_note（備註獨立表）。
  //   上方 rows(fact_schedule) 的任何計算完全不動,這裡只是「另外多寫一張表」。
  //   整段包 try/catch：備註出任何狀況都不得影響班表主流程。
  try {
    var _noteWarn = [];
    _LAST_NOTE_STAT_ = {person: 0, date: 0, blocked: false};
    var _noteRows = _collectFactNoteRows_(ss, fname, storeMeta, year, month, rows, _noteWarn);
    for (var _w = 0; _w < _noteWarn.length; _w++) Logger.log('fact_note: ' + _noteWarn[_w]);
    if (_noteRows === null){
      Logger.log('fact_note: ' + fname + ' 版面自檢未過,整檔不寫入 fact_note');
    } else {
      var _n = _appendNoteRows_(_noteRows);
      Logger.log('fact_note: ' + fname + ' 寫入 ' + _n + ' 列 (person ' + _LAST_NOTE_STAT_.person +
                 ' / date ' + _LAST_NOTE_STAT_.date + ')');
    }
  } catch(_e){
    _LAST_NOTE_STAT_ = {person: 0, date: 0, blocked: true};
    Logger.log('fact_note: ' + fname + ' 產生失敗(不影響 fact_schedule): ' + _e);
  }

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
  _ensureNoteHeader(sh);
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
  Logger.log('清空後: ' + sh.getLastRow() + ' 列(應該是 1)');
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

  Logger.log('\n=== file_name 列數分佈 ===');
  var fileList = Object.keys(fileCount).sort();
  fileList.forEach(function(k){
    var v = fileCount[k];
    var flag = (v > 150 || v < 30) ? ' 異常' : '';
    Logger.log(k + ': ' + v + ' 列' + flag);
  });

  Logger.log('\n總列數: ' + (data.length - 1));
  Logger.log('不同檔名數: ' + fileList.length);
}

// ====================================================================
// Phase 2: 日營收同步 (v5.0: gross/net 全取自 v_daily_net, headcount 仍掃 Sheet)
// ====================================================================
var POS_SS_ID_   = '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao';
var POS_SHEET_   = 'POS資料';

// POS資料 欄位 index：A分店代碼(0) B分店名稱(1) C建立日期(2) D商品名稱(3)
//   E主類別(4) F次類別(5) G商品單價(6) H數量(7) I實收總額(8) J來客數(9)
var POS_COL_STORE_ID = 0;
var POS_COL_DATE     = 2;
var POS_COL_CAT1     = 4;
var POS_COL_CAT2     = 5;
var POS_COL_PRICE    = 6;
var POS_COL_QTY      = 7;
var POS_COL_AMOUNT   = 8;

/**
 * ★ v5.6 新增：syncDailySales 全期模式掃 POS 的起始下限 [年, 月(0-based), 日]。
 * 原本硬寫在函式內（new Date(2025,0,1)）。
 *
 * ⚠️ 已知副作用（v5.6 之前一直存在）：POS資料 最早是 2024-12-20，但這個下限是 2025-01-01，
 *    所以 2024-12 那 12 天從來沒被算進 headcountAgg；而 queryDailyNet 是 BQ 全歷史、有 2024-12。
 *    在舊的 fallback 行為下，fact_daily_sales 的 2024-12 各列 headcount 與 rph 早就被寫成 0。
 *    v5.6 改為「缺 headcount 就跳過」後，這些列不會再被覆寫，但既有的 0 值不會自動修復。
 *    若要補回 2024-12，把下方改成 [2024, 11, 1] 再跑一次全期同步即可（需經營者裁示）。
 */
var SYNC_POS_CUTOFF_ = [2025, 0, 1];

/**
 * ★ v5.6 新增：唯讀 dry-run。
 * 印出「若現在執行全期重算，有多少列會因缺 headcount 而被跳過」。
 * 不寫入 fact_daily_sales、不呼叫 buildWeekKPI、不寫任何 log 分頁。
 */
/** ★ 臨時包裝：讓它排在函式清單最上方，避免下拉選錯。驗完即刪。 */
function aaaRunDryRun(){ return syncDailySalesDryRun(); }

function syncDailySalesDryRun(){
  Logger.log('════════ syncDailySales DRY-RUN（唯讀，不寫入任何分頁）════════');
  var cutoff = new Date(SYNC_POS_CUTOFF_[0], SYNC_POS_CUTOFF_[1], SYNC_POS_CUTOFF_[2]);
  cutoff.setHours(0,0,0,0);
  Logger.log('POS 掃描下限 SYNC_POS_CUTOFF_ = ' +
    _fmtDate(cutoff.getFullYear(), cutoff.getMonth()+1, cutoff.getDate()));

  var netRows = queryDailyNet();
  var viewMap = {};
  for(var i=0; i<netRows.length; i++){
    viewMap[netRows[i].sale_date + '|' + netRows[i].store_code] = true;
  }
  Logger.log('BQ viewMap keys = ' + Object.keys(viewMap).length);

  var posSh = SpreadsheetApp.openById(POS_SS_ID_).getSheetByName(POS_SHEET_);
  var lastRow = posSh.getLastRow(), lastCol = posSh.getLastColumn();
  var BATCH = 50000, headcountAgg = {};
  for(var startRow = 2; startRow <= lastRow; startRow += BATCH){
    var rowsToRead = Math.min(BATCH, lastRow - startRow + 1);
    var data = posSh.getRange(startRow, 1, rowsToRead, lastCol).getValues();
    for(var r=0; r<data.length; r++){
      var dateV = data[r][POS_COL_DATE]; if(!dateV) continue;
      var d = (dateV instanceof Date) ? dateV : new Date(String(dateV));
      if(isNaN(d) || d < cutoff) continue;
      var storeId = parseInt(data[r][POS_COL_STORE_ID], 10); if(isNaN(storeId)) continue;
      headcountAgg[_fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate()) + '|' + storeId] = true;
    }
  }
  Logger.log('POS headcountAgg keys = ' + Object.keys(headcountAgg).length);

  var skipped = [], byYm = {};
  for(var k in viewMap){
    if (headcountAgg[k]) continue;
    skipped.push(k);
    var ym = k.substring(0, 7);
    byYm[ym] = (byYm[ym] || 0) + 1;
  }

  Logger.log('');
  Logger.log('會被跳過的列數（BQ 有、POS 無 headcount）= ' + skipped.length);
  if (skipped.length === 0) {
    Logger.log('✅ 0 列。資料完整，全期重算不會動到任何既有 headcount。');
  } else {
    Logger.log('逐月分布：');
    Object.keys(byYm).sort().forEach(function(ym){
      Logger.log('  ' + ym + '：' + byYm[ym] + ' 個店日');
    });
    Logger.log('');
    Logger.log('⚠️ v5.6 起這些列會「跳過不寫」，既有值保留。');
    Logger.log('   （v5.6 之前會被覆寫成 headcount=0、rph=0）');
  }
  Logger.log('');
  Logger.log('════════ DRY-RUN 結束，未寫入任何資料 ════════');
  return { skipped: skipped.length, byYm: byYm };
}

function syncDailySales(startDate, endDate){
  Logger.log('=== syncDailySales START (v5.0 金額全取自 BigQuery v_daily_net) ===');
  var cutoff = new Date(SYNC_POS_CUTOFF_[0], SYNC_POS_CUTOFF_[1], SYNC_POS_CUTOFF_[2]);
  cutoff.setHours(0,0,0,0);
  // 增量模式：傳入 startDate/endDate 時只處理範圍內日期
  var sdDate = null, edDate = null;
  if (startDate && endDate) {
    sdDate = new Date(startDate); sdDate.setHours(0, 0, 0, 0);
    edDate = new Date(endDate);   edDate.setHours(23, 59, 59, 999);
    cutoff = sdDate; // 以 startDate 取代全期 cutoff，加速 POS 掃描
  }
  Logger.log('cutoff date = ' + _fmtDate(cutoff.getFullYear(), cutoff.getMonth()+1, cutoff.getDate()));

  // ★ v5.0：gross/net 改從 v_daily_net 取（單一真相），headcount 仍從 POS資料 Sheet 算
  var netRows = queryDailyNet(startDate, endDate);  // 全期；回 [{sale_date, store_code, gross, total_discount, net}]
  var viewMap = {};               // key = "YYYY-MM-DD|storeCode" -> {gross, net}
  for(var i=0; i<netRows.length; i++){
    var nr = netRows[i];
    viewMap[nr.sale_date + '|' + nr.store_code] = { gross: nr.gross, net: nr.net };
  }
  Logger.log('viewMap keys=' + Object.keys(viewMap).length);

  // 掃 POS資料 Sheet，只算 headcount（金額不再從這裡來）
  var posSS = SpreadsheetApp.openById(POS_SS_ID_);
  var posSh = posSS.getSheetByName(POS_SHEET_);
  if(!posSh){ Logger.log('ERROR: POS 分頁找不到'); return; }
  var lastRow = posSh.getLastRow();
  var lastCol = posSh.getLastColumn();
  Logger.log('POS total rows=' + lastRow + ' cols=' + lastCol);

  var BATCH = 50000;
  var headcountAgg = {};
  var skipped = 0;

  for(var startRow = 2; startRow <= lastRow; startRow += BATCH){
    var rowsToRead = Math.min(BATCH, lastRow - startRow + 1);
    var data = posSh.getRange(startRow, 1, rowsToRead, lastCol).getValues();
    for(var r=0; r<data.length; r++){
      var row = data[r];
      var dateV = row[POS_COL_DATE];
      if(!dateV) { skipped++; continue; }
      var d;
      if(dateV instanceof Date) d = dateV;
      else { d = new Date(String(dateV)); if(isNaN(d)){ skipped++; continue; } }
      if(d < cutoff){ skipped++; continue; }
        if (edDate && d > edDate) { skipped++; continue; }
      var dateStr = _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate());
      var storeId = parseInt(row[POS_COL_STORE_ID], 10);
      if(isNaN(storeId)){ skipped++; continue; }
      var cat1 = String(row[POS_COL_CAT1] || '').trim();
      var cat2 = String(row[POS_COL_CAT2] || '').trim();
      var key = dateStr + '|' + storeId;
      if(!headcountAgg[key]) headcountAgg[key] = {headcount:0, rows_total:0};
      headcountAgg[key].rows_total += 1;
      if(!EXCLUDE_CAT1_[cat1] && !EXCLUDE_CAT2_[cat2]){
        headcountAgg[key].headcount += 1;
      }
    }
  }
  Logger.log('headcount keys=' + Object.keys(headcountAgg).length + ' skipped=' + skipped);

  // 合併：金額用 viewMap，headcount 用 headcountAgg
  var allKeys = {};
  for(var k in viewMap) allKeys[k] = true;
  for(var k in headcountAgg) allKeys[k] = true;

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
  var skippedNoHc = 0;   // ★ v5.6：因缺 headcount 而跳過的 key 數
  for(var k in allKeys){
    var parts = k.split('|');
    var dateStr2 = parts[0];
    var sid = parseInt(parts[1], 10);
    var v = viewMap[k] || {gross:0, net:0};
    // ★ v5.6：headcount 缺值時「跳過該列」，不再 fallback 成 0。
    //   舊行為：BQ 有該日、POS 已無該日（滾動裁切或 cutoff 之外）→ hc 落到 {headcount:0}
    //           → 把 fact_daily_sales 既有的正確 headcount 與 rph 覆寫成 0，資料永久毀損。
    //   新行為：這種 key 直接不寫，既有列原封不動保留。
    //   注意：headcountAgg[k] 存在但 headcount===0 是合法值（該日確實無計人品項），照常寫入。
    if (!headcountAgg[k]) { skippedNoHc++; continue; }
    var hc = headcountAgg[k];
    var gross = _round2(v.gross);
    var netRev = _round2(v.net);
    var rph = hc.headcount > 0 ? _round2(gross / hc.headcount) : 0;
    var rowVals = [dateStr2, sid, storeNames[sid]||'', gross, hc.headcount, hc.rows_total, rph, now, netRev];
    if(keyToRow[k]){
      updates.push([keyToRow[k], rowVals]);
    } else {
      appends.push(rowVals);
    }
  }
  Logger.log('updates=' + updates.length + ' appends=' + appends.length +
             ' skippedNoHc=' + skippedNoHc + '（BQ 有、POS 無 headcount，已保留既有列不覆寫）');

  for(var u=0; u<updates.length; u++){
    sh.getRange(updates[u][0], 1, 1, 9).setValues([updates[u][1]]);
  }
  if(appends.length > 0){
    sh.getRange(sh.getLastRow()+1, 1, appends.length, 9).setValues(appends);
  }

  _appendLog('syncDailySales', '', updates.length+appends.length, 'OK',
    'v5.0 金額源=v_daily_net updates=' + updates.length + ' appends=' + appends.length +
    ' skipped=' + skipped + ' viewKeys=' + Object.keys(viewMap).length);
  Logger.log('=== syncDailySales DONE ===');
}

var EXCLUDE_CAT1_ = {'加價購': true, '找不到': true};
var EXCLUDE_CAT2_ = {'陪同入場費': true};

// ====================================================================
// Phase 3: 週度 KPI 聚合 (net 欄位口徑跟隨 Phase 2，無需改動)
// ====================================================================

var FLAG_WEEKDAY_HOURS_ = 16;
var FLAG_WEEKEND_HOURS_ = 24;
var FLAG_REV_PER_HOUR_  = 500;

var MORNING_START_ = 9 * 60;
var MORNING_END_   = 13 * 60;
var EVENING_START_ = 18 * 60;
var EVENING_END_   = 23 * 60;

function buildWeekKPI(){
  Logger.log('=== buildWeekKPI START (v5.0 net 跟隨 fact_daily_sales) ===');
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
    var netRev = parseFloat(salData[s][8]) || 0;
    if(!sd || isNaN(ss_)) continue;
    var sdStr = (sd instanceof Date) ? _fmtDate(sd.getFullYear(), sd.getMonth()+1, sd.getDate()) : String(sd).trim();
    salMap[sdStr + '|' + ss_] = {gross: rev, net: netRev};
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
        weekday_net:0, weekend_net:0,
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
      var entry = salMap[ds + '|' + W2.store_id] || {gross:0, net:0};
      if(W2.seenDates[ds] === 'off'){
        W2.weekend_revenue += entry.gross;
        W2.weekend_net += entry.net;
      } else {
        W2.weekday_revenue += entry.gross;
        W2.weekday_net += entry.net;
      }
    }
  }

  var outSh = ss.getSheetByName('fact_week_kpi');
  if(!outSh){ Logger.log('ERROR: fact_week_kpi 找不到'); return; }
  if(outSh.getLastRow() > 1){
    outSh.getRange(2, 1, outSh.getLastRow()-1, 17).clearContent();
  }

  var now = new Date();
  var outRows = [];
  for(var wk2 in weekAgg){
    var W3 = weekAgg[wk2];
    var meta = storeMeta[W3.store_id] || {name:'', zone:'', manager:''};
    var totalRev = W3.weekday_revenue + W3.weekend_revenue;
    var totalNet = W3.weekday_net + W3.weekend_net;
    var rph = W3.total_hours > 0 ? _round2(totalRev / W3.total_hours) : 0;
    var rphNet = W3.total_hours > 0 ? _round2(totalNet / W3.total_hours) : 0;
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
      now,
      _round2(totalNet),
      rphNet
    ]);
  }

  outRows.sort(function(a,b){
    if(a[0] < b[0]) return -1;
    if(a[0] > b[0]) return 1;
    return a[2] - b[2];
  });

  if(outRows.length > 0){
    outSh.getRange(2, 1, outRows.length, 17).setValues(outRows);
  }
  Logger.log('written rows=' + outRows.length);

  _appendLog('buildWeekKPI', '', outRows.length, 'OK', 'v5.0 weeks×stores=' + outRows.length);
  Logger.log('=== buildWeekKPI DONE ===');
}

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
      result = _getStoreDailyData(sid, ym, e.parameter.bonus === '1');   // ★ v5.8：bonus=1 才多查甜點券$100（日獎金用）
    } else if(action === 'processUploads'){
      result = _processUploads();
    } else if(action === 'getRecentImportLog'){
      var n = parseInt(e.parameter.n, 10) || 20;
      result = _getRecentImportLog(n);
    } else if(action === 'ping'){
      result = {ok:true, time:new Date().toISOString(), version:'schedule-v5.8-bonus'};
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

function doPost(e){
  var result;
  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action || '';

    if(action === 'uploadSchedule'){
      result = _uploadSchedule(payload);
    } else {
      result = {ok:false, error:'unknown POST action: ' + action};
    }
  } catch(err){
    result = {ok:false, error: 'doPost err: ' + String(err)};
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function _uploadSchedule(payload){
  var rawFname = String(payload.fileName || '').trim();
  if(!rawFname) return {ok:false, error:'fileName 缺失'};

  var nameOk = /(\d{1,2})\s*店.*?(\d{6})/.test(rawFname);
  if(!nameOk){
    return {ok:false, error:'檔名不符規範,需含「N 店班表 YYYYMM」(N=1-12,YYYYMM=6 位)'};
  }

  var m = rawFname.match(/(\d{1,2})\s*店.*?(\d{6})/);
  var storeNum = parseInt(m[1], 10);
  if(storeNum < 1 || storeNum > 12){
    return {ok:false, error:'店號超出 1-12 範圍'};
  }
  var yyyymm = m[2];
  var year   = parseInt(yyyymm.substring(0,4), 10);
  var month  = parseInt(yyyymm.substring(4,6), 10);
  if(year < 2024 || year > 2030 || month < 1 || month > 12){
    return {ok:false, error:'年月不合理 (年=' + year + ', 月=' + month + ')'};
  }

  var ext = '.xls';
  var lowName = rawFname.toLowerCase();
  if(lowName.indexOf('.xlsx') >= 0) ext = '.xlsx';
  var fname = storeNum + '店班表' + yyyymm + ext;

  var base64 = String(payload.base64Data || '');
  if(!base64) return {ok:false, error:'base64Data 缺失'};

  var mime = payload.mimeType || 'application/vnd.ms-excel';

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return {ok: false, error: '已有上傳處理中,請 30 秒後重試'};
  }
  try {
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mime, fname);
    var folder = DriveApp.getFolderById(FOLDER_ID_);

    var deletedCount = 0;
    var driveFiles = folder.getFiles();
    while(driveFiles.hasNext()){
      var existingFile = driveFiles.next();
      var existingName = existingFile.getName();
      if(_canonicalKey(existingName) === storeNum + '-' + yyyymm){
        existingFile.setTrashed(true);
        deletedCount++;
      }
    }

    var xlsxFile = folder.createFile(blob);
    var converted = Drive.Files.copy({name: fname, mimeType: 'application/vnd.google-apps.spreadsheet'}, xlsxFile.getId(), {convert: true});
    xlsxFile.setTrashed(true);
    var newFile = DriveApp.getFileById(converted.id);

    // ★ 2026-09-11：這裡「不再」清空 fact_schedule / fact_note。
    //   只立一個 pending 標記,清空與寫入都交給 parsePendingFiles 在同一次執行內完成。
    //   好處:上傳中斷 → 維持舊資料(儀表板照常),不會出現空白月份。
    var ckNew = storeNum + '-' + yyyymm;
    _clearCanonicalScriptProperties(storeNum, yyyymm);
    PropertiesService.getScriptProperties().setProperty('pending_ck_' + ckNew, '1');

    // 排一個 60 秒後的自動處理。店長此刻即可離開,不必等解析。
    // 已有待跑排程時不重複建立(parsePendingFiles 一次會掃完所有待處理檔)。
    var scheduled = _ensureAutoProcessTrigger_(AUTO_TRIGGER_DELAY_);

    _appendLog(rawFname, newFile.getId(), 0, 'UPLOADED',
      'canonical=' + fname + ' replaced=' + deletedCount + ' pending=' + ckNew +
      ' autoTrigger=' + (scheduled ? 'new' : 'existing') + ' size=' + bytes.length);

    return {
      ok: true,
      originalName: rawFname,
      fileName: fname,
      renamed: (rawFname !== fname),
      fileId: newFile.getId(),
      size: bytes.length,
      replacedOld: deletedCount,
      purgedRows: 0,              // 保留欄位供前端相容;新架構不在上傳階段清空
      pending: ckNew,
      autoScheduled: scheduled
    };
  } catch(err){
    return {ok:false, error:'寫 Drive 失敗: ' + err.message};
  } finally {
    lock.releaseLock();
  }
}

function _purgeFactScheduleByCanonical(storeNum, yyyymm){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  if(!sh) return 0;
  var lastRow = sh.getLastRow();
  if(lastRow < 2) return 0;

  // ★ 2026-09-11：改 memory 過濾法（讀 → 過濾 → clearContent → setValues → 驗殘留）。
  //   原本逐列 deleteRow,一列約 0.1~0.3 秒;purge 移到解析端後一次可能要清多家店,
  //   逐列刪會直接撞 GAS 6 分鐘上限。作法比照 _purgeFactNoteByCanonical。
  var W = Math.max(17, sh.getLastColumn());   // fact_schedule 固定 17 欄(含 Q personal_note)
  var height = lastRow - 1;
  var data = sh.getRange(2, 1, height, W).getValues();
  var targetKey = parseInt(storeNum, 10) + '-' + yyyymm;

  var keep = [];
  for(var i = 0; i < data.length; i++){
    var fn = String(data[i][15] || '').trim();      // P 欄(index 15) file_name
    if(_canonicalKey(fn) === targetKey) continue;
    keep.push(data[i]);
  }
  var removed = data.length - keep.length;
  if(removed === 0) return 0;

  sh.getRange(2, 1, height, W).clearContent();
  if(keep.length > 0) sh.getRange(2, 1, keep.length, W).setValues(keep);
  SpreadsheetApp.flush();

  // 驗殘留（讀回原高度,被清掉的尾段一律為 ''）
  var after = sh.getRange(2, 16, height, 1).getValues();
  var leftover = 0;
  for(var j = 0; j < after.length; j++){
    if(_canonicalKey(String(after[j][0] || '').trim()) === targetKey) leftover++;
  }
  if(leftover > 0){
    throw new Error('_purgeFactScheduleByCanonical 殘留 ' + leftover + ' 列 (' + targetKey + ')');
  }
  Logger.log('_purgeFactScheduleByCanonical: ' + targetKey + ' 刪除 ' + removed + ' 列,保留 ' + keep.length + ' 列');
  return removed;
}

// ════════════════════════════════════════════════════════════════
// ★ 2026-09-11 上傳後自動處理（店長不必等解析）
// ────────────────────────────────────────────────────────────────
// 背景:12 家店月底同時上傳時,_processUploads 的 5 分鐘處理鎖會擋掉後來者,
// 店長若在畫面前乾等只會看到失敗、然後重按,愈按愈糟。
//
// 作法:上傳成功即排一個 60 秒後的 one-time trigger,由 Google 自己執行解析。
// 店長關掉視窗完全不影響。跑完若還有 pending,自動再排下一輪(最多 10 輪)。
//
// 注意:one-time trigger 觸發後不會自動消失,會佔用專案 trigger 額度(上限 20),
// 所以每次執行的第一件事就是清掉所有同名 trigger。
// ════════════════════════════════════════════════════════════════
var AUTO_TRIGGER_FN_     = 'autoProcessAfterUpload_';
var AUTO_TRIGGER_DELAY_  = 60 * 1000;        // 上傳後 60 秒開始處理
var AUTO_TRIGGER_RETRY_  = 90 * 1000;        // 沒做完 → 90 秒後再來
var AUTO_TRIGGER_MAX_    = 10;               // 最多自動重排 10 輪
var AUTO_TRIGGER_ROUND_  = 'auto_trigger_round';

// 目前還有幾個 canonical 等著被解析
function _countPendingUploads_(){
  var keys = PropertiesService.getScriptProperties().getKeys();
  var n = 0;
  for(var i = 0; i < keys.length; i++){
    if(keys[i].indexOf('pending_ck_') === 0) n++;
  }
  return n;
}

// 清掉所有同名 one-time trigger（含剛觸發的自己），回傳清掉幾個
function _cleanupAutoTriggers_(){
  var n = 0;
  try {
    var ts = ScriptApp.getProjectTriggers();
    for(var i = 0; i < ts.length; i++){
      if(ts[i].getHandlerFunction() === AUTO_TRIGGER_FN_){ ScriptApp.deleteTrigger(ts[i]); n++; }
    }
  } catch(e){ Logger.log('_cleanupAutoTriggers_ err: ' + e); }
  return n;
}

// 確保有一個待跑的自動處理排程（已有就不重複建）
// 整段包 try/catch：排程建不起來(額度/授權)絕不能讓上傳失敗
function _ensureAutoProcessTrigger_(delayMs){
  try {
    var ts = ScriptApp.getProjectTriggers();
    for(var i = 0; i < ts.length; i++){
      if(ts[i].getHandlerFunction() === AUTO_TRIGGER_FN_){
        Logger.log('_ensureAutoProcessTrigger_: 已有待跑排程,不重複建立');
        return false;
      }
    }
    ScriptApp.newTrigger(AUTO_TRIGGER_FN_).timeBased().after(delayMs || AUTO_TRIGGER_DELAY_).create();
    Logger.log('_ensureAutoProcessTrigger_: 已排定 ' + Math.round((delayMs || AUTO_TRIGGER_DELAY_)/1000) + ' 秒後自動處理');
    return true;
  } catch(e){
    Logger.log('⚠️ _ensureAutoProcessTrigger_ 失敗(不影響上傳,仍可由前端或每日排程補跑): ' + e);
    return false;
  }
}

// trigger 實際執行的函式
function autoProcessAfterUpload_(){
  _cleanupAutoTriggers_();        // 第一件事:清掉自己,避免額度累積

  var sp = PropertiesService.getScriptProperties();

  if(_countPendingUploads_() === 0){
    Logger.log('autoProcessAfterUpload_: 無待處理上傳,結束');
    sp.deleteProperty(AUTO_TRIGGER_ROUND_);
    return;
  }

  var round = parseInt(sp.getProperty(AUTO_TRIGGER_ROUND_) || '0', 10) + 1;
  if(round > AUTO_TRIGGER_MAX_){
    Logger.log('⚠️ autoProcessAfterUpload_: 已自動重試 ' + AUTO_TRIGGER_MAX_ + ' 輪仍有待處理項,停止重排');
    _appendLog('autoProcessAfterUpload_', '', 0, 'FAIL',
      '自動重試達上限,仍有 ' + _countPendingUploads_() + ' 個 pending,請執行 aaaCheckPendingUploads 檢查');
    sp.deleteProperty(AUTO_TRIGGER_ROUND_);
    return;
  }
  sp.setProperty(AUTO_TRIGGER_ROUND_, String(round));

  var r = _processUploads();
  Logger.log('autoProcessAfterUpload_ 第 ' + round + ' 輪: ok=' + (r && r.ok) +
             (r && r.error ? (' / ' + r.error) : ''));

  var left = _countPendingUploads_();
  if(left > 0){
    Logger.log('尚有 ' + left + ' 個待處理,排下一輪');
    _ensureAutoProcessTrigger_(AUTO_TRIGGER_RETRY_);
  } else {
    sp.deleteProperty(AUTO_TRIGGER_ROUND_);
    Logger.log('autoProcessAfterUpload_: 全部處理完畢 ✅');
  }
}

// 唯讀檢查工具:看還有哪些上傳沒被解析、目前有哪些排程
function aaaCheckPendingUploads(){
  var sp = PropertiesService.getScriptProperties();
  var keys = sp.getKeys();
  var list = [];
  for(var i = 0; i < keys.length; i++){
    if(keys[i].indexOf('pending_ck_') === 0) list.push(keys[i].substring('pending_ck_'.length));
  }
  list.sort();
  Logger.log('=== 待解析的上傳 (pending) ===');
  Logger.log(list.length === 0 ? '無,全部都解析完了 ✅' : (list.length + ' 筆: ' + list.join('、')));
  Logger.log('自動重試輪次: ' + (sp.getProperty(AUTO_TRIGGER_ROUND_) || '0'));

  var names = [];
  try {
    var ts = ScriptApp.getProjectTriggers();
    for(var j = 0; j < ts.length; j++) names.push(ts[j].getHandlerFunction());
  } catch(e){ names.push('(讀取失敗: ' + e + ')'); }
  Logger.log('目前專案排程 (' + names.length + '/20): ' + (names.length ? names.join('、') : '無'));
}

// 手動補跑:待解析項若卡住,執行這支即可(等同按一次前端的「補跑」)
function aaaRunPendingNow(){
  var before = _countPendingUploads_();
  Logger.log('執行前待解析: ' + before + ' 筆');
  var r = _processUploads();
  Logger.log('結果: ok=' + (r && r.ok) + (r && r.error ? (' / ' + r.error) : ''));
  Logger.log('執行後待解析: ' + _countPendingUploads_() + ' 筆');
  return r;
}

function _clearCanonicalScriptProperties(storeNum, yyyymm){
  var sp = PropertiesService.getScriptProperties();
  var keys = sp.getKeys();
  var targetKey = parseInt(storeNum, 10) + '-' + yyyymm;
  var cleared = 0;
  for(var i=0; i<keys.length; i++){
    if(keys[i].indexOf('parsed_fname_') !== 0) continue;
    var fn = keys[i].substring('parsed_fname_'.length);
    if(_canonicalKey(fn) === targetKey){
      sp.deleteProperty(keys[i]);
      cleared++;
    }
  }
  return cleared;
}

function _processUploads(){
  var sp = PropertiesService.getScriptProperties();
  var lockUntil = parseInt(sp.getProperty('process_lock_until') || '0', 10);
  var now = Date.now();
  if(lockUntil > now){
    var remainSec = Math.ceil((lockUntil - now) / 1000);
    return {
      ok: false,
      error: '其他人正在處理中,預估 ' + remainSec + ' 秒後完成,請稍候再試'
    };
  }
  sp.setProperty('process_lock_until', String(now + 5 * 60 * 1000));

  var startTime = Date.now();
  var steps = [];

  var s1Start = Date.now();
  var s1ok = false, s1err = '';
  try {
    parsePendingFiles();
    s1ok = true;
  } catch(err){ s1err = String(err); }
  steps.push({
    step: 'parsePendingFiles',
    ok: s1ok,
    error: s1err,
    seconds: Math.round((Date.now() - s1Start) / 1000)
  });

  var s2Start = Date.now();
  var s2ok = false, s2err = '';
  try {
    // v5.3: 改用 90 天增量視窗（取代全期 BQ 掃描，避免每次上傳等 70+ 秒）
    var _endD = new Date();
    var _startD = new Date();
    _startD.setDate(_endD.getDate() - 90);
    var _fmt90 = function(d){ return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd'); };
    syncDailySales(_fmt90(_startD), _fmt90(_endD));
    s2ok = true;
  } catch(err){ s2err = String(err); }
  steps.push({
    step: 'syncDailySales',
    ok: s2ok,
    error: s2err,
    seconds: Math.round((Date.now() - s2Start) / 1000)
  });

  var s3Start = Date.now();
  var s3ok = false, s3err = '';
  try {
    buildWeekKPI();
    s3ok = true;
  } catch(err){ s3err = String(err); }
  steps.push({
    step: 'buildWeekKPI',
    ok: s3ok,
    error: s3err,
    seconds: Math.round((Date.now() - s3Start) / 1000)
  });

  sp.deleteProperty('process_lock_until');

  var totalSec = Math.round((Date.now() - startTime) / 1000);
  var allOk = s1ok && s2ok && s3ok;

  _appendLog('processUploads', '', 0, allOk ? 'OK' : 'PARTIAL',
    'total=' + totalSec + 's steps=' + steps.map(function(s){
      return s.step + '(' + (s.ok?'ok':'fail') + ',' + s.seconds + 's)';
    }).join('|'));

  return {
    ok: allOk,
    totalSeconds: totalSec,
    steps: steps
  };
}

function _getRecentImportLog(n){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('import_log');
  if(!sh) return {ok:false, error:'import_log not found'};
  var lastRow = sh.getLastRow();
  if(lastRow < 2) return {ok:true, logs:[]};

  var start = Math.max(2, lastRow - n + 1);
  var rows = sh.getRange(start, 1, lastRow - start + 1, 6).getValues();
  var logs = [];
  for(var i = rows.length - 1; i >= 0; i--){
    var r = rows[i];
    var ts = r[0];
    var tsStr = (ts instanceof Date)
      ? Utilities.formatDate(ts, 'Asia/Taipei', 'MM/dd HH:mm:ss')
      : String(ts);
    logs.push({
      timestamp:    tsStr,
      file_name:    String(r[1] || ''),
      file_id:      String(r[2] || ''),
      rows_written: parseInt(r[3], 10) || 0,
      status:       String(r[4] || ''),
      note:         String(r[5] || '')
    });
  }
  return {ok:true, logs:logs};
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
      flags:            String(row[13] || '').split(',').filter(function(x){return x;}),
      total_net_revenue: parseFloat(row[15]) || 0,
      rev_per_hour_net:  parseFloat(row[16]) || 0
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
        rev_per_head: parseFloat(row2[6]) || 0,
        net_revenue: parseFloat(row2[8]) || 0
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

  var dailyHours = _aggregateDailyHours(ss);

  return {
    ok: true,
    generated_at: new Date().toISOString(),
    week_kpi: rows,
    daily_sales: daily,
    daily_hours: dailyHours,
    stores: stores
  };
}

function _aggregateDailyHours(ss){
  var schSh = ss.getSheetByName('fact_schedule');
  if(!schSh) return [];
  var lastRow = schSh.getLastRow();
  if(lastRow < 2) return [];

  var schData = schSh.getDataRange().getValues();
  var idx = _buildHeaderIdx(schData[0]);

  if(idx['date'] === undefined || idx['store_id'] === undefined || idx['hours_total'] === undefined){
    Logger.log('_aggregateDailyHours: missing header columns');
    return [];
  }

  var agg = {};
  for(var r=1; r<schData.length; r++){
    var row = schData[r];
    var dateV = row[idx['date']];
    if(!dateV) continue;
    var d = (dateV instanceof Date) ? dateV : new Date(String(dateV));
    if(isNaN(d)) continue;
    var dateStr = _fmtDate(d.getFullYear(), d.getMonth()+1, d.getDate());
    var sid = parseInt(row[idx['store_id']], 10);
    if(isNaN(sid)) continue;
    var hours = parseFloat(row[idx['hours_total']]) || 0;
    var isWeekend = parseInt(row[idx['is_weekend']], 10) === 1;
    var isHoliday = parseInt(row[idx['is_holiday']], 10) === 1;
    var isOff = (isWeekend || isHoliday) ? 1 : 0;

    var k = dateStr + '|' + sid;
    if(!agg[k]){
      agg[k] = { date: dateStr, store_id: sid, hours: 0, is_off: isOff };
    }
    agg[k].hours += hours;
  }

  var out = [];
  Object.keys(agg).forEach(function(k){
    var v = agg[k];
    out.push({
      date: v.date,
      store_id: v.store_id,
      hours: _round2(v.hours),
      is_off: v.is_off
    });
  });

  out.sort(function(a,b){
    if(a.date < b.date) return -1;
    if(a.date > b.date) return 1;
    return a.store_id - b.store_id;
  });

  Logger.log('_aggregateDailyHours: ' + out.length + ' rows aggregated');
  return out;
}

function _getStoreDailyData(storeId, ym, withBonus) {
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

    var empName = String(row[idx['employee_name']] || '').trim();
    var shiftSegs = String(row[idx['shift_segments']] || '').trim();
    if (empName) {
      var label = empName + '(' + _round2(empHours) + 'h) ' + shiftSegs;
      daily[dStr].employees.push({
        name: empName,
        hours: _round2(empHours),
        shift: shiftSegs,
        label: label,
        note: String(row[idx['personal_note']] || '').trim(),
        // ★ v5.8（2026-10-01 日獎金）：只多回傳既有欄位，不改 fact_schedule
        emp_id: String(row[idx['employee_id']] || '').trim(),
        role: String(row[idx['role']] || '').trim(),
        cell: String(row[idx['cell_raw']] || '').trim()
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
      // ★ v5.1：營收改讀 net(第9欄, index 8)，與分店比較(buildWeekKPI)口徑統一。
      //   原讀 gross(rr[3], 第4欄)，v5.0 後 gross 來自 v_daily_net，當日 gross 可能尚未進 view 而為 0，
      //   造成日曆顯示 $0 但分店比較有值。改讀 net 後兩處同源、同數字。
      daily[sdStr].revenue   = parseFloat(rr[8]) || 0;
      daily[sdStr].headcount = parseInt(rr[4], 10) || 0;
      daily[sdStr].rev_per_head = parseFloat(rr[6]) || 0;
      daily[sdStr].net_revenue = parseFloat(rr[8]) || 0;
    }
  }

  // ★ 2026-09-10：讀 fact_note,建「整日備註」與「未排班備註」兩張 map。
  //   employees[].note 維持原狀繼續讀 fact_schedule Q 欄,這裡完全不碰。
  var dateNoteMap = {}, offNoteMap = {};
  var noteSh = ss.getSheetByName(FACT_NOTE_SHEET_);
  if (noteSh && noteSh.getLastRow() > 1) {
    var nData = noteSh.getRange(2, 1, noteSh.getLastRow() - 1, FACT_NOTE_HEADER_.length).getValues();
    for (var ni = 0; ni < nData.length; ni++) {
      var nr = nData[ni];
      if (parseInt(nr[1], 10) !== storeId) continue;
      var nd = _dateKey_(nr[0]);
      if (!nd || nd.substring(0, 7) !== ym) continue;
      var nNote = String(nr[6] || '').trim();
      if (!nNote) continue;
      var nScope = String(nr[3] || '').trim();
      if (nScope === 'date') {
        dateNoteMap[nd] = dateNoteMap[nd] ? (dateNoteMap[nd] + '；' + nNote) : nNote;
      } else if (nScope === 'person' && parseInt(nr[7], 10) === 0) {
        if (!offNoteMap[nd]) offNoteMap[nd] = [];
        offNoteMap[nd].push({ name: String(nr[5] || '').trim(), note: nNote });
      }
    }
  }
  // 診斷用：備註日期在 daily 找不到對應天(該日完全無班表也無營收)的清單。
  // 本次不新建 daily 列(避免動到有效天/月均口徑),僅回報給前端與 Chat 判讀。
  var noteOrphans = [];
  Object.keys(dateNoteMap).forEach(function (k) { if (!daily[k]) noteOrphans.push('date:' + k); });
  Object.keys(offNoteMap).forEach(function (k) { if (!daily[k]) noteOrphans.push('person:' + k); });
  noteOrphans.sort();

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
    d.employees.sort(function(a,b){ return b.hours - a.hours; });
    d.date_note = dateNoteMap[k] || '';                 // ★ 2026-09-10 整日備註
    d.off_notes = offNoteMap[k] || [];                  // ★ 2026-09-10 未排班備註
    return d;
  });

  // ★ 2026-07-01：有效天 = 有排班「且」有實際營收的天。
  //   未來已排班但尚未進帳的天(hours>0, revenue=0)不算進分母，
  //   否則月均日營收/月均 rev-h 會被 0 營收天稀釋到接近 $0。
  //   已完整結束月份的營業日都同時 hours>0 且 revenue>0，故對舊月份數字零影響。
  var validRows = rows.filter(function (r) { return r.hours > 0 && r.revenue > 0; });
  var monthAvgHours   = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.hours; }, 0) / validRows.length : 0;
  var monthAvgRevenue = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.revenue; }, 0) / validRows.length : 0;
  var monthAvgRph     = validRows.length > 0 ? validRows.reduce(function (s, r) { return s + r.rev_per_hour; }, 0) / validRows.length : 0;

  // ★ v5.8（2026-10-01 經營者待處理事項 #3）：日獎金營收＝POS 淨營收＋「甜點券$100」折抵額。
  //   折扣名稱用「完全相等」比對（不可含「自己人$100甜點券」）；discount 欄本來就是正數（折抵額），直接加回。
  //   只在 bonus=1 才查；查不到不影響其他欄位（bonus.ok=false＋錯誤訊息）。
  var bonusInfo = null;
  if (withBonus) {
    bonusInfo = { ok: false, project: BONUS_VOUCHER_NAME_, error: '' };
    try {
      var vm = _queryVoucher100_(storeId, ym);
      rows.forEach(function (r) {
        var v = vm[r.date] || { amt: 0, qty: 0 };
        r.voucher100 = v.amt;
        r.voucher100_qty = v.qty;
        r.bonus_rev = (r.revenue || 0) + v.amt;
      });
      bonusInfo.ok = true;
    } catch (err) {
      bonusInfo.error = String(err);
    }
  }

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
    bonus: bonusInfo,
    note_orphans: noteOrphans,
    baseline: {
      avg_hours:   _round2(monthAvgHours),
      avg_revenue: _round2(monthAvgRevenue),
      avg_rph:     _round2(monthAvgRph),
      valid_days:  validRows.length
    }
  };
}

// ★ v5.8：日獎金營收要加回的折扣項目（完全相等比對）
var BONUS_VOUCHER_NAME_ = '甜點券$100';

/** 單店單月每日「甜點券$100」折抵額與張數 → { 'YYYY-MM-DD': {amt, qty} } */
function _queryVoucher100_(storeId, ym) {
  if (!/^\d{4}-\d{2}$/.test(ym)) throw new Error('bad ym');
  var y = parseInt(ym.slice(0, 4), 10), m = parseInt(ym.slice(5, 7), 10);
  var last = new Date(y, m, 0).getDate();
  var from = ym + '-01', to = ym + '-' + (last < 10 ? '0' + last : last);
  var sql = 'SELECT FORMAT_DATE("%Y-%m-%d", sale_date) AS d, SUM(discount) AS amt, SUM(quantity) AS qty ' +
            'FROM `diybc-make-sync.diybc_pos.pos_discounts` ' +
            'WHERE store_code = ' + parseInt(storeId, 10) + ' AND project_name = @p ' +
            "AND sale_date BETWEEN '" + from + "' AND '" + to + "' GROUP BY d";
  var res = UrlFetchApp.fetch(
    'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries',
    {
      method: 'post', contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + _getBqAccessToken_() },
      payload: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 30000, parameterMode: 'NAMED',
        queryParameters: [{ name: 'p', parameterType: { type: 'STRING' }, parameterValue: { value: BONUS_VOUCHER_NAME_ } }] }),
      muteHttpExceptions: true
    });
  var data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('甜點券查詢錯誤: ' + JSON.stringify(data.error));
  var out = {};
  (data.rows || []).forEach(function (row) {
    out[row.f[0].v] = { amt: Number(row.f[1].v) || 0, qty: Number(row.f[2].v) || 0 };
  });
  return out;
}

// ====================================================================
// 健康檢查 / 對帳工具
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

function monthlyAutoSync(){
  Logger.log('=== monthlyAutoSync START ===');
  var errors = [];

  try {
    parsePendingFiles();
  } catch(e){
    Logger.log('parsePendingFiles err: ' + e);
    errors.push('parse:' + e.message);
  }
  Utilities.sleep(3000);

  try {
    syncDailySales();
  } catch(e){
    Logger.log('syncDailySales err: ' + e);
    errors.push('sync:' + e.message);
  }
  Utilities.sleep(3000);

  try {
    buildWeekKPI();
  } catch(e){
    Logger.log('buildWeekKPI err: ' + e);
    errors.push('week:' + e.message);
  }

  var status = errors.length === 0 ? 'OK' : 'PARTIAL_FAIL';
  var note   = errors.length === 0 ? 'auto-sync done' : errors.join(' | ');
  _appendLog('monthlyAutoSync', '', 0, status, note);
  Logger.log('=== monthlyAutoSync DONE === status=' + status);
}

function verifyCanonicalDedup(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  var lastRow = sh.getLastRow();
  if(lastRow < 2){ Logger.log('fact_schedule 空白'); return; }

  var fileNames = sh.getRange(2, 16, lastRow-1, 1).getValues();

  var canonical = {};
  for(var i=0; i<fileNames.length; i++){
    var fn = String(fileNames[i][0]||'').trim();
    if(!fn) continue;
    var ck = _canonicalKey(fn);
    if(!ck) continue;
    if(!canonical[ck]) canonical[ck] = {file_names: {}, total: 0};
    canonical[ck].file_names[fn] = (canonical[ck].file_names[fn] || 0) + 1;
    canonical[ck].total++;
  }

  Logger.log('=== canonical key 重複檢查 ===');
  var keys = Object.keys(canonical).sort();
  var suspicious = [];
  keys.forEach(function(ck){
    var info = canonical[ck];
    var variantCount = Object.keys(info.file_names).length;
    var status = variantCount > 1 ? ' 多寫法' : '';
    Logger.log('canonical=' + ck + ' total=' + info.total + ' variants=' + variantCount + status);
    Object.keys(info.file_names).forEach(function(fn){
      Logger.log('  - "' + fn + '": ' + info.file_names[fn] + ' 列');
    });
    if(variantCount > 1){
      suspicious.push({key: ck, total: info.total, variants: info.file_names});
    }
  });

  Logger.log('\n=== 結論 ===');
  if(suspicious.length === 0){
    Logger.log('沒有重複,每個 canonical key 都只對應一種 file_name 寫法');
  } else {
    Logger.log('發現 ' + suspicious.length + ' 個 canonical key 有多種寫法,可能有重複匯入');
  }
}

function cleanupDuplicatesByCanonical(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  var lastRow = sh.getLastRow();
  if(lastRow < 2){ Logger.log('fact_schedule 空白'); return; }

  var fileNames = sh.getRange(2, 16, lastRow-1, 1).getValues();

  var canonicalGroups = {};
  for(var i=0; i<fileNames.length; i++){
    var fn = String(fileNames[i][0]||'').trim();
    if(!fn) continue;
    var ck = _canonicalKey(fn);
    if(!ck) continue;
    if(!canonicalGroups[ck]) canonicalGroups[ck] = [];
    canonicalGroups[ck].push({rowIdx: i + 2, fname: fn});
  }

  var keepFnameByCk = {};
  Object.keys(canonicalGroups).forEach(function(ck){
    var rows = canonicalGroups[ck];
    var uniqueFnames = {};
    rows.forEach(function(r){ uniqueFnames[r.fname] = true; });
    var fnList = Object.keys(uniqueFnames);
    if(fnList.length <= 1){
      keepFnameByCk[ck] = fnList[0] || '';
      return;
    }
    var parts = ck.split('-');
    var standardForms = [
      parts[0] + '店班表' + parts[1] + '.xls',
      parts[0] + '店班表' + parts[1] + '.xlsx',
      parts[0] + '店班表' + parts[1]
    ];
    var picked = null;
    for(var s=0; s<standardForms.length; s++){
      if(fnList.indexOf(standardForms[s]) >= 0){
        picked = standardForms[s];
        break;
      }
    }
    if(!picked){
      var maxCount = 0;
      fnList.forEach(function(fn){
        var c = rows.filter(function(r){ return r.fname === fn; }).length;
        if(c > maxCount){ maxCount = c; picked = fn; }
      });
    }
    keepFnameByCk[ck] = picked;
  });

  var rowsToDelete = [];
  for(var i = fileNames.length - 1; i >= 0; i--){
    var fn = String(fileNames[i][0]||'').trim();
    var ck = _canonicalKey(fn);
    if(!ck) continue;
    var keepFname = keepFnameByCk[ck];
    if(keepFname && fn !== keepFname){
      rowsToDelete.push(i + 2);
    }
  }

  Logger.log('準備刪除 ' + rowsToDelete.length + ' 列重複資料');
  if(rowsToDelete.length === 0){
    Logger.log('沒有重複,無需清理');
    return;
  }
  rowsToDelete.forEach(function(rowIdx){
    sh.deleteRow(rowIdx);
  });
  SpreadsheetApp.flush();
  Logger.log('已刪除 ' + rowsToDelete.length + ' 列重複資料');
}

// ====================================================================
// v5.0 對帳驗證（唯讀，跑完可刪）
// ====================================================================
function verifyV5_排班(){
  var rows = queryDailyNet('2026-04-01', '2026-04-30');
  var net1 = 0, gross1 = 0;
  rows.forEach(function(r){
    if(r.store_code === 1){ net1 += r.net; gross1 += r.gross; }
  });
  Logger.log('1店2026-04 gross=' + gross1 + ' net=' + net1 + ' (應 gross=393258 net=369921)');
}

function 查789店6月源頭(){
  var rows = queryDailyNet('2026-06-01', '2026-06-06');  // 只 SELECT，唯讀
  rows.forEach(function(r){
    if(r.store_code === 7 || r.store_code === 8 || r.store_code === 9){
      Logger.log(r.sale_date + ' 店' + r.store_code + ' gross=' + r.gross + ' net=' + r.net);
    }
  });
  Logger.log('--- 以上是 BigQuery view 裡 7/8/9 店 6/1~6/6 的資料 ---');
}

/**
 * 每日自動:把 BQ 最新資料推進 fact_daily_sales + 重算 fact_week_kpi。
 * 排程 06~07 點(等「一鍵追加新品」05~06 點寫完 BQ 之後)。
 */
var DAILY_WINDOW_DAYS = 3; // 滾動視窗，吸收 POS 隔日補洞修正

// 每日 06~07 點 trigger 綁這個（增量版 + v5.5 髒日期偵測）
function dailyRebuild_salesAndKpi() {
  var tz = 'Asia/Taipei';
  var end   = new Date(Date.now() - 86400000);                             // 昨天
  var start = new Date(end.getTime() - (DAILY_WINDOW_DAYS - 1) * 86400000); // 昨天往前 N-1 天
  var s = Utilities.formatDate(start, tz, 'yyyy-MM-dd');
  var e = Utilities.formatDate(end,   tz, 'yyyy-MM-dd');
  Logger.log('=== dailyRebuild (incremental) ' + s + ' ~ ' + e + ' ===');

  // --- v5.5:髒日期偵測（近 26h 被重寫、但在 3 天視窗外）---
  var dirtyDates = [];
  try {
    dirtyDates = _findDirtyDatesOutsideWindow_();
    Logger.log(dirtyDates.length > 0
      ? '--- 髒日期偵測:' + dirtyDates.length + ' 個視窗外日期 → ' + dirtyDates.join(', ') + ' ---'
      : '--- 髒日期偵測:無 ---');
  } catch (err) {
    Logger.log('⚠️ 髒日期偵測失敗(不影響本次重算，照跑 3 天窗): ' + err);
    dirtyDates = [];
  }
  dirtyDates.forEach(function (d) {
    try {
      Logger.log('--- 補算髒日期 ' + d + ' ---');
      syncDailySales(d, d);
    } catch (err) {
      Logger.log('⚠️ 髒日期 ' + d + ' 補算失敗: ' + err);
    }
  });

  // --- 原有 3 天滾動視窗（不變）---
  var t1 = Date.now();
  syncDailySales(s, e);                    // ★ 只有 sync 吃範圍參數
  Logger.log('--- syncDailySales ' + ((Date.now()-t1)/1000) + 's done, buildWeekKPI ---');
  var t2 = Date.now();
  buildWeekKPI();                          // ★ 不傳參數，整表重寫；髒日期+3天窗共用這一次
  Logger.log('=== dailyRebuild DONE, buildWeekKPI ' + ((Date.now()-t2)/1000) + 's ===');
}

// 災難復原 / 手動全期重算用（保留舊行為）
function dailyRebuild_salesAndKpi_full() {
  syncDailySales();
  buildWeekKPI();
}

/**
 * v5.5:查近 26 小時被重寫、但 sale_date 在 3 天視窗外的日期。
 * 重用 bq_connector.gs 的 _getBqAccessToken_() 與 BQ_PROJECT_ID，不重複 JWT 邏輯。
 * SA 若無 pos_transactions 讀取權會 throw，由呼叫端 try/catch 吸收。
 */
function _findDirtyDatesOutsideWindow_() {
  var token = _getBqAccessToken_();
  var sql =
    'SELECT DISTINCT CAST(sale_date AS STRING) ' +
    'FROM `diybc-make-sync.diybc_pos.pos_transactions` ' +
    'WHERE imported_at > TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 26 HOUR) ' +
    "AND sale_date < DATE_SUB(CURRENT_DATE('Asia/Taipei'), INTERVAL 3 DAY)";

  var res = UrlFetchApp.fetch(
    'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 30000 }),
      muteHttpExceptions: true
    }
  );
  var data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('髒日期查詢錯誤: ' + JSON.stringify(data.error));

  var out = [];
  (data.rows || []).forEach(function (row) { out.push(row.f[0].v); });

  // ★ v5.8（2026-10-01）：折扣（pos_discounts）事後補資料也要觸發重算。
  //   2026-08-20 折扣整天漏匯、10/1 補回時，排班 8/20 營收沒有自動跟著改（原本只看 pos_transactions）。
  //   另包 try/catch：這段失敗不影響上面原本的偵測結果。
  try {
    var sql2 =
      'SELECT DISTINCT CAST(sale_date AS STRING) ' +
      'FROM `diybc-make-sync.diybc_pos.pos_discounts` ' +
      'WHERE imported_at > TIMESTAMP_SUB(CURRENT_TIMESTAMP(), INTERVAL 26 HOUR) ' +
      "AND sale_date < DATE_SUB(CURRENT_DATE('Asia/Taipei'), INTERVAL 3 DAY)";
    var res2 = UrlFetchApp.fetch(
      'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries',
      {
        method: 'post', contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({ query: sql2, useLegacySql: false, timeoutMs: 30000 }),
        muteHttpExceptions: true
      });
    var data2 = JSON.parse(res2.getContentText());
    if (data2.error) throw new Error(JSON.stringify(data2.error));
    (data2.rows || []).forEach(function (row) { var d = row.f[0].v; if (out.indexOf(d) < 0) out.push(d); });
  } catch (err2) {
    Logger.log('⚠️ 折扣補資料日期偵測失敗(不影響交易端偵測): ' + err2);
  }
  return out;
}
// ===== 一次性備份函式（止血作業用，執行後可刪） =====
function backup_factschedule_20260609() {
  var ss = SpreadsheetApp.openById('19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM');
  var src = ss.getSheetByName('fact_schedule');
  if (!src) { Logger.log('ERROR: fact_schedule not found'); return; }
    
  // 確認備份分頁不存在
  var existing = ss.getSheetByName('fact_schedule_backup_20260609');
  if (existing) {
    Logger.log('備份分頁已存在，跳過');
    Logger.log('備份列數: ' + (existing.getLastRow()));
    return;
  }
    
  // 執行備份
  var backup = src.copyTo(ss);
  backup.setName('fact_schedule_backup_20260609');
    
  var srcRows = src.getLastRow();
  var bakRows = backup.getLastRow();
  Logger.log('=== 備份完成 ===');
  Logger.log('原始 fact_schedule 列數 (含 header): ' + srcRows);
  Logger.log('備份分頁列數 (含 header): ' + bakRows);
  Logger.log('列數一致: ' + (srcRows === bakRows));
}
function get_backup_gid() {
  var ss = SpreadsheetApp.openById('19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM');
  var sh = ss.getSheetByName('fact_schedule_backup_20260609');
  if (!sh) { Logger.log('NOT FOUND'); return; }
  Logger.log('GID: ' + sh.getSheetId());
  Logger.log('LastRow: ' + sh.getLastRow());
}

// ===== 一次性 Dry-run（止血作業用，只 Logger，不寫入） =====
function dedup_dryrun_202606() {
  var ss = SpreadsheetApp.openById('19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM');
  var sh = ss.getSheetByName('fact_schedule');
  if (!sh) { Logger.log('ERROR: fact_schedule not found'); return; }

  var lastRow = sh.getLastRow();
  // 讀全部資料（含 header）
  var data = sh.getRange(1, 1, lastRow, 16).getValues();
  var header = data[0];

  // 欄位 index（0-based）
  var colDate = 0;   // A: date
  var colStore = 1;  // B: store_id
  var colEmpId = 3;  // D: employee_id
  var colEmpName = 4; // E: employee_name
  var colShift = 6;  // G: shift_segments

  // 目標：只處理 2026-06，store 3 & 4
  var TARGET_STORES = [3, 4];
  var YEAR_MONTH = '2026-06';

  // 統計結構
  var storeStats = {};
  TARGET_STORES.forEach(function(s) {
    storeStats[s] = { total: 0, toDelete: 0, toKeep: 0, samples: [] };
  });

  var seenKeys = {};       // key -> true（第一次出現）
  var toDeleteRows = [];   // {rowIndex(1-based), date, storeId, empId, empName}

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[colDate];
    var storeId = parseInt(row[colStore], 10);

    // 只處理 target stores
    if (TARGET_STORES.indexOf(storeId) === -1) continue;

    // 只處理 2026-06
    var dateStr = '';
    if (dateVal instanceof Date) {
      var y = dateVal.getFullYear();
      var m = dateVal.getMonth() + 1;
      dateStr = y + '-' + (m < 10 ? '0' + m : m);
    } else {
      dateStr = String(dateVal).substring(0, 7);
    }
    if (dateStr !== YEAR_MONTH) continue;

    storeStats[storeId].total++;

    // canonical dedup key: date + store + empId + shift
    var empId = row[colEmpId];
    var shift = row[colShift];
    var fullDateStr = dateVal instanceof Date
      ? (dateVal.getFullYear() + '-' + (dateVal.getMonth()+1 < 10 ? '0' : '') + (dateVal.getMonth()+1) + '-' + (dateVal.getDate() < 10 ? '0' : '') + dateVal.getDate())
      : String(dateVal).substring(0, 10);
    var key = storeId + '|' + fullDateStr + '|' + empId + '|' + shift;

    if (seenKeys[key]) {
      // 重複 -> 標記刪除
      storeStats[storeId].toDelete++;
      toDeleteRows.push({
        rowIndex: i + 1,  // 1-based sheet row
        date: fullDateStr,
        storeId: storeId,
        empId: String(empId),
        empName: String(row[colEmpName])
      });
    } else {
      seenKeys[key] = true;
      storeStats[storeId].toKeep++;
    }
  }

  // === 輸出 ===
  Logger.log('===== dedup_dryrun_202606 =====');
  Logger.log('fact_schedule 總列數(含 header): ' + lastRow);
  Logger.log('');

  TARGET_STORES.forEach(function(s) {
    var st = storeStats[s];
    Logger.log('--- Store ' + s + ' ---');
    Logger.log('  2026-06 總列數: ' + st.total);
    Logger.log('  預計刪除(重複): ' + st.toDelete);
    Logger.log('  保留(唯一):     ' + st.toKeep);
    Logger.log('  刪除後筆數:     ' + st.toKeep);
  });

  Logger.log('');
  Logger.log('--- 全部其他店(非 store 3/4) ---');
  Logger.log('  預計刪除: 0 (本次完全不碰)');

  Logger.log('');
  Logger.log('--- 總計 ---');
  var totalDelete = 0, totalKeep = 0;
  TARGET_STORES.forEach(function(s) {
    totalDelete += storeStats[s].toDelete;
    totalKeep += storeStats[s].toKeep;
  });
  Logger.log('  預計刪除總列數: ' + totalDelete);
  Logger.log('  保留總列數:     ' + totalKeep);

  Logger.log('');
  Logger.log('--- 抽樣 5 列待刪 row ---');
  var samples = toDeleteRows.slice(0, 5);
  samples.forEach(function(r, idx) {
    Logger.log('  [' + (idx+1) + '] rowIndex=' + r.rowIndex + ' | date=' + r.date + ' | store=' + r.storeId + ' | empId=' + r.empId + ' | empName=' + r.empName);
  });

  Logger.log('');
  Logger.log('===== DRY-RUN 完成，未寫入任何資料 =====');
}

// ===== 一次性 Dry-run v2（配對驗證 + 隨機抽樣，只 Logger，不寫入） =====
function dedup_dryrun_v2_202606() {
  var ss = SpreadsheetApp.openById('19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM');
  var sh = ss.getSheetByName('fact_schedule');
  if (!sh) { Logger.log('ERROR: fact_schedule not found'); return; }

  var lastRow = sh.getLastRow();
  var data = sh.getRange(1, 1, lastRow, 16).getValues();

  // 欄位 index（0-based，與 dry-run v1 完全相同）
  var colDate = 0;    // A
  var colStore = 1;   // B
  var colEmpId = 3;   // D
  var colEmpName = 4; // E
  var colShift = 6;   // G

  var TARGET_STORES = [3, 4];
  var YEAR_MONTH = '2026-06';

  // === Pass 1: 建立 seenKeys map，key -> {firstRowIndex, date, empId, shift, empName} ===
  var seenKeys = {};
  var toDeleteRows = [];  // {rowIndex, key, date, storeId, empId, empName, shift}

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[colDate];
    var storeId = parseInt(row[colStore], 10);

    if (TARGET_STORES.indexOf(storeId) === -1) continue;

    // 年月篩選
    var dateStr = '';
    if (dateVal instanceof Date) {
      var y = dateVal.getFullYear();
      var m = dateVal.getMonth() + 1;
      dateStr = y + '-' + (m < 10 ? '0' + m : m);
    } else {
      dateStr = String(dateVal).substring(0, 7);
    }
    if (dateStr !== YEAR_MONTH) continue;

    // fullDateStr
    var fullDateStr = '';
    if (dateVal instanceof Date) {
      var y2 = dateVal.getFullYear();
      var m2 = dateVal.getMonth() + 1;
      var d2 = dateVal.getDate();
      fullDateStr = y2 + '-' + (m2 < 10 ? '0' + m2 : m2) + '-' + (d2 < 10 ? '0' + d2 : d2);
    } else {
      fullDateStr = String(dateVal).substring(0, 10);
    }

    var empId = row[colEmpId];
    var shift = row[colShift];
    var empName = String(row[colEmpName]);

    // key = storeId|date|empId|shift (完全同 v1)
    var key = storeId + '|' + fullDateStr + '|' + empId + '|' + shift;

    if (seenKeys[key]) {
      // 重複 -> 標刪
      toDeleteRows.push({
        rowIndex: i + 1,
        key: key,
        date: fullDateStr,
        storeId: storeId,
        empId: String(empId),
        empName: empName,
        shift: String(shift),
        keepRowIndex: seenKeys[key].rowIndex  // 對應的保留 row
      });
    } else {
      seenKeys[key] = {
        rowIndex: i + 1,
        date: fullDateStr,
        empId: String(empId),
        empName: empName,
        shift: String(shift)
      };
    }
  }

  Logger.log('===== dedup_dryrun_v2_202606 =====');
  Logger.log('總待刪筆數: ' + toDeleteRows.length);
  Logger.log('');

  // === 件事 1：完整 key 公式 ===
  Logger.log('--- KEY 組成 ---');
  Logger.log('  key = storeId + "|" + fullDateStr + "|" + empId + "|" + shift_segments');
  Logger.log('  colDate=0(A), colStore=1(B), colEmpId=3(D), colShift=6(G)');
  Logger.log('');

  // === 件事 2+3：隨機 10 筆，每筆配對顯示保留 row ===
  Logger.log('--- 隨機 10 筆待刪 row，配對對應保留 row ---');

  // Fisher-Yates shuffle 取 10 筆
  var indices = [];
  for (var j = 0; j < toDeleteRows.length; j++) indices.push(j);
  // 用簡單的 pseudo-random（seed: 時間）
  var seed = new Date().getTime();
  function rand(n) { seed = (seed * 1664525 + 1013904223) & 0xffffffff; return Math.abs(seed) % n; }
  for (var k = indices.length - 1; k > 0; k--) {
    var r = rand(k + 1);
    var tmp = indices[k]; indices[k] = indices[r]; indices[r] = tmp;
  }
  var sample10 = indices.slice(0, 10).map(function(idx){ return toDeleteRows[idx]; });
  // 依 rowIndex 排序方便對照
  sample10.sort(function(a,b){ return a.rowIndex - b.rowIndex; });

  var allOk = true;
  sample10.forEach(function(del, n) {
    var keepInfo = seenKeys[del.key];
    // 交叉驗證：保留 row 的 date/shift 必須與待刪 row 完全相同
    var dateMatch = keepInfo && keepInfo.date === del.date;
    var shiftMatch = keepInfo && keepInfo.shift === del.shift;
    var empMatch = keepInfo && keepInfo.empId === del.empId;
    var allMatch = dateMatch && shiftMatch && empMatch;
    if (!allMatch) allOk = false;

    Logger.log('[' + (n+1) + '] 待刪 row ' + del.rowIndex +
      ' | ' + del.date + ' | store=' + del.storeId +
      ' | empId=' + del.empId + ' | empName=' + del.empName +
      ' | shift="' + del.shift + '"');
    if (keepInfo) {
      Logger.log('    對應保留 row ' + keepInfo.rowIndex +
        ' | ' + keepInfo.date + ' | empId=' + keepInfo.empId +
        ' | shift="' + keepInfo.shift + '"' +
        ' | 完全一致: ' + (allMatch ? 'YES' : 'BUG!'));
    } else {
      Logger.log('    ERROR: 找不到對應保留 row -> BUG!');
      allOk = false;
    }
  });

  Logger.log('');
  Logger.log('10 筆全部通過: ' + (allOk ? 'YES' : 'NO - 有 BUG'));
  Logger.log('===== DRY-RUN v2 完成，未寫入任何資料 =====');
}

 // ===== Step 3 Apply — memory 過濾法，不用 deleteRow =====
function dedup_apply_202606() {
  var ss = SpreadsheetApp.openById('19X3cqX70aWNTc6KFTP5jushG5S06xNicYV1ulDXl-hM');
  var sh = ss.getSheetByName('fact_schedule');
  if (!sh) { Logger.log('ERROR: fact_schedule not found'); return; }
  // 2026-06-25 封存防呆：fact_schedule 加 Q 欄(personal_note)後,此一次性 16 欄 dedup re-run 會錯位 Q,禁止執行。
  if (sh.getLastColumn() > 16) {
    throw new Error('[已封存] dedup_apply_202606 僅支援 16 欄,目前 ' + sh.getLastColumn() + ' 欄,禁止執行');
  }

  var colDate  = 0;  // A
  var colStore = 1;  // B
  var colEmpId = 3;  // D
  var colShift = 6;  // G
  var TARGET_STORES = [3, 4];
  var YEAR_MONTH = '2026-06';

  var lastRow = sh.getLastRow();
  var data = sh.getRange(1, 1, lastRow, 16).getValues();
  var header = data[0];
  var rows = data.slice(1); // data rows only

  // === Pass 1: 建立 seenKeys，找出要刪的 row index（0-based in rows[]）===
  var seenKeys = {};
  var deleteSet = {};  // index in rows[] to delete

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var dateVal  = row[colDate];
    var storeId  = parseInt(row[colStore], 10);
    var empId    = row[colEmpId];
    var shift    = row[colShift];

    if (TARGET_STORES.indexOf(storeId) === -1) continue;

    var dateStr = '';
    if (dateVal instanceof Date) {
      var y = dateVal.getFullYear();
      var m = dateVal.getMonth() + 1;
      var d = dateVal.getDate();
      dateStr = y + '-' + (m < 10 ? '0' : '') + m + '-' + (d < 10 ? '0' : '') + d;
    } else {
      dateStr = String(dateVal).substring(0, 10);
    }

    if (!dateStr.startsWith(YEAR_MONTH)) continue;

    var key = storeId + '|' + dateStr + '|' + empId + '|' + shift;

    if (seenKeys[key] === undefined) {
      seenKeys[key] = i;  // 保留第一次出現
    } else {
      deleteSet[i] = true;  // 後出現的標為刪除
    }
  }

  var toDeleteCount = Object.keys(deleteSet).length;
  var toKeepCount   = rows.length - toDeleteCount;

  Logger.log('===== dedup_apply_202606 APPLY =====');
  Logger.log('原始資料列數 (不含 header): ' + rows.length);
  Logger.log('即將刪除: ' + toDeleteCount + ' 列');
  Logger.log('保留: ' + toKeepCount + ' 列');

  // 安全確認：必須剛好 209
  if (toDeleteCount !== 209) {
    Logger.log('❌ ABORT: 預期刪除 209 列，實際計算 ' + toDeleteCount + ' 列，中止！');
    return;
  }
  Logger.log('✅ 刪除數確認 = 209，繼續執行...');

  // === Pass 2: 建立過濾後的 rows ===
  var cleanRows = [];
  for (var j = 0; j < rows.length; j++) {
    if (!deleteSet[j]) cleanRows.push(rows[j]);
  }

  Logger.log('過濾後列數: ' + cleanRows.length + '（預期 ' + (rows.length - 209) + '）');

  // === 執行寫入：clearContents from row 2，再 setValues ===
  // 清除 row 2 到 lastRow（保留 header row 1）
  sh.getRange(2, 1, lastRow - 1, 16).clearContent();

  // 寫回過濾後資料
  if (cleanRows.length > 0) {
    sh.getRange(2, 1, cleanRows.length, 16).setValues(cleanRows);
  }

  Logger.log('✅ setValues 完成');

  // === Finally: 自驗 ===
  var newLastRow = sh.getLastRow();
  Logger.log('--- 自驗 ---');
  Logger.log('新 fact_schedule 總列數 (含 header): ' + newLastRow + '（預期 3858）');

  // 重新讀取驗證
  var verifyData = sh.getRange(2, 1, newLastRow - 1, 16).getValues();
  var cnt3 = 0, cnt4 = 0, cnt1 = 0, cnt2 = 0, cnt5 = 0;
  for (var k = 0; k < verifyData.length; k++) {
    var s = parseInt(verifyData[k][colStore], 10);
    if (s === 3) cnt3++;
    else if (s === 4) cnt4++;
    else if (s === 1) cnt1++;
    else if (s === 2) cnt2++;
    else if (s === 5) cnt5++;
  }
  Logger.log('store 3 列數: ' + cnt3 + '（預期 ≥1，2026-06 清後 114 列應消失）');
  Logger.log('store 4 列數: ' + cnt4 + '（預期 ≥1，2026-06 清後 95 列應消失）');
  Logger.log('store 1 列數: ' + cnt1 + '（對照用，應與 dry-run 前一致）');
  Logger.log('store 2 列數: ' + cnt2 + '（對照用，應與 dry-run 前一致）');
  Logger.log('store 5 列數: ' + cnt5 + '（對照用，應與 dry-run 前一致）');

  // 驗證 store3 + store4 的 2026-06 列數
  var cnt3_202606 = 0, cnt4_202606 = 0;
  for (var m2 = 0; m2 < verifyData.length; m2++) {
    var s2 = parseInt(verifyData[m2][colStore], 10);
    var dv = verifyData[m2][colDate];
    var ds = '';
    if (dv instanceof Date) {
      var yy = dv.getFullYear();
      var mm = dv.getMonth() + 1;
      ds = yy + '-' + (mm < 10 ? '0' : '') + mm;
    } else {
      ds = String(dv).substring(0, 7);
    }
    if (ds === '2026-06') {
      if (s2 === 3) cnt3_202606++;
      if (s2 === 4) cnt4_202606++;
    }
  }
  Logger.log('store 3 × 2026-06 列數: ' + cnt3_202606 + '（預期 114）');
  Logger.log('store 4 × 2026-06 列數: ' + cnt4_202606 + '（預期 95）');

  if (cnt3_202606 === 114 && cnt4_202606 === 95) {
    Logger.log('✅ 自驗通過：store3=114, store4=95');
  } else {
    Logger.log('❌ 自驗異常：store3=' + cnt3_202606 + ', store4=' + cnt4_202606);
  }
  Logger.log('===== Apply 完成 =====');
}

function testLock() {
  var lock = LockService.getScriptLock();
  Logger.log('[Test 1] First tryLock...');
  var ok1 = lock.tryLock(3000);
  Logger.log('[Test 1] Got lock: ' + ok1);

  if (ok1) {
    Logger.log('[Test 2] Second tryLock (reentrance test)...');
    var ok2 = lock.tryLock(3000);
    Logger.log('[Test 2] Got lock again (reentrant?): ' + ok2);

    Logger.log('[Test 3] Release first time...');
    lock.releaseLock();
    Logger.log('[Test 3] Released');

    if (ok2) {
      Logger.log('[Test 4] Release second time...');
      lock.releaseLock();
      Logger.log('[Test 4] Released');
    }

    Logger.log('[Test 5] Try tryLock again after release...');
    var ok3 = lock.tryLock(3000);
    Logger.log('[Test 5] Got lock: ' + ok3);
    if (ok3) lock.releaseLock();
  }

  Logger.log('===== testLock done =====');
}

// ============================================================
// 備份:Lock 上線前原版 _uploadSchedule
// 建立時間:2026-06-09
// 用途:若 Lock 版本出問題,改名回滾
// ============================================================
function _uploadSchedule_v_pre_lock_20260609(payload){
  // ⚠️ 2026-09-11 起已不相容:本版仍是舊架構(上傳當下就 purge),
  //    而 parsePendingFiles 已改為「看 pending 標記、在解析當下才 purge」。
  //    若改名啟用本版,上傳後不會留下 pending 標記 → 檔案會被當成新檔正常解析(不會重複),
  //    但會退回「先清空後解析」的空窗風險。要 rollback 請連同 parsePendingFiles 一起回退。
  throw new Error('[已封存] _uploadSchedule_v_pre_lock_20260609 與 v5.7 非同步架構不相容,禁止直接啟用');
  var rawFname = String(payload.fileName || '').trim();
  if(!rawFname) return {ok:false, error:'fileName 缺失'};

  var nameOk = /(\d{1,2})\s*店.*?(\d{6})/.test(rawFname);
  if(!nameOk){
    return {ok:false, error:'檔名不符規範,需含「N 店班表 YYYYMM」(N=1-12,YYYYMM=6 位)'};
  }

  var m = rawFname.match(/(\d{1,2})\s*店.*?(\d{6})/);
  var storeNum = parseInt(m[1], 10);
  if(storeNum < 1 || storeNum > 12){
    return {ok:false, error:'店號超出 1-12 範圍'};
  }
  var yyyymm = m[2];
  var year   = parseInt(yyyymm.substring(0,4), 10);
  var month  = parseInt(yyyymm.substring(4,6), 10);
  if(year < 2024 || year > 2030 || month < 1 || month > 12){
    return {ok:false, error:'年月不合理 (年=' + year + ', 月=' + month + ')'};
  }

  var ext = '.xls';
  var lowName = rawFname.toLowerCase();
  if(lowName.indexOf('.xlsx') >= 0) ext = '.xlsx';
  var fname = storeNum + '店班表' + yyyymm + ext;

  var base64 = String(payload.base64Data || '');
  if(!base64) return {ok:false, error:'base64Data 缺失'};

  var mime = payload.mimeType || 'application/vnd.ms-excel';
  try {
    var bytes = Utilities.base64Decode(base64);
    var blob = Utilities.newBlob(bytes, mime, fname);
    var folder = DriveApp.getFolderById(FOLDER_ID_);

    var deletedCount = 0;
    var driveFiles = folder.getFiles();
    while(driveFiles.hasNext()){
      var existingFile = driveFiles.next();
      var existingName = existingFile.getName();
      if(_canonicalKey(existingName) === storeNum + '-' + yyyymm){
        existingFile.setTrashed(true);
        deletedCount++;
      }
    }

    var xlsxFile = folder.createFile(blob);
    var converted = Drive.Files.copy({name: fname, mimeType: 'application/vnd.google-apps.spreadsheet'}, xlsxFile.getId(), {convert: true});
    xlsxFile.setTrashed(true);
    var newFile = DriveApp.getFileById(converted.id);

    var purgedRows = _purgeFactScheduleByCanonical(storeNum, yyyymm);
    var purgedNotes = _purgeFactNoteByCanonical(storeNum, yyyymm);   // ★ 2026-09-10 purge 鏡像
    _clearCanonicalScriptProperties(storeNum, yyyymm);

    _appendLog(rawFname, newFile.getId(), 0, 'UPLOADED',
      'canonical=' + fname + ' replaced=' + deletedCount + ' purged=' + purgedRows +
      ' purgedNotes=' + purgedNotes + ' size=' + bytes.length);

    return {
      ok: true,
      originalName: rawFname,
      fileName: fname,
      renamed: (rawFname !== fname),
      fileId: newFile.getId(),
      size: bytes.length,
      replacedOld: deletedCount,
      purgedRows: purgedRows
    };
  } catch(err){
    return {ok:false, error:'寫 Drive 失敗: ' + err.message};
  }
}

// ============================================================
// 備份:Lock 上線前原版 parsePendingFiles
// 建立時間:2026-06-09
// 用途:若 Lock 版本出問題,改名回滾
// ============================================================
function parsePendingFiles_v_pre_lock_20260609(){
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

    var ck = _canonicalKey(fname);
    if(ck && existingFiles['__canonical__' + ck]){
      Logger.log('SKIP: canonical key ' + ck + ' already in fact_schedule');
      sp.setProperty('parsed_fname_' + fname, 'true');
      continue;
    }

    _processFile(f, isXls, sp, fname);
    existingFiles[fname] = 'just_written';
    if(ck) existingFiles['__canonical__' + ck] = 'just_written';
  }
  Logger.log('=== parsePendingFiles DONE ===');
}
// TEMP TEST - delete after verify
function _testBqNet20260622() {
  var rows = queryDailyNet('2026-06-22', '2026-06-22');
  rows.forEach(function(r) {
    Logger.log('store=' + r.store_code + ' net=' + r.net + ' gross=' + r.gross);
  });
}

// ====================================================================
// 個人備註 tooltip 功能（2026-06-25 新增）
// ====================================================================

// 清洗備註儲存格：拆行 → 去空白 → 去重複行(保序) → 以 ' / ' 接
function _cleanNote(raw){
  if(raw == null) return '';
  var lines = String(raw).split(/\r?\n/);
  var seen = {}, out = [];
  for(var i=0; i<lines.length; i++){
    var s = lines[i].trim();
    if(!s || seen[s]) continue;
    seen[s] = 1; out.push(s);
  }
  return out.join(' / ');
}

// 從 (noteMap, 日期, empId, empName) 取備註：先用 empId,沒中再用姓名後援
function _noteFor(noteMap, dateStr, empId, empName){
  var n = '';
  if(empId) n = noteMap[dateStr + '|id|' + empId] || '';
  if(!n && empName) n = noteMap[dateStr + '|nm|' + empName] || '';
  return n;
}

// 讀「個人備註」分頁 → 建 {日期|id|員工編號: 備註, 日期|nm|姓名: 備註}
// 日期表頭與 empId 拆法完全比照 _parse,確保 key 對得上。沒這分頁回空物件、不影響主流程。
function _buildNoteMap(ss, year, month){
  var map = {};
  var sheets = ss.getSheets();
  var noteSheet = null;
  for(var si=0; si<sheets.length; si++){
    if(sheets[si].getName() === '個人備註'){ noteSheet = sheets[si]; break; }
  }
  if(!noteSheet){ Logger.log('_buildNoteMap: 無「個人備註」分頁,略過'); return map; }

  var data = noteSheet.getDataRange().getValues();
  if(data.length < 3) return map;

  var dateRowIdx = -1, colDateMap = {};
  for(var rIdx=0; rIdx<Math.min(5, data.length); rIdx++){
    var tryMap = {};
    var row = data[rIdx];
    for(var c=1; c<row.length; c++){
      var ds = _toDateStr(row[c], year, month);
      if(ds) tryMap[c] = ds;
    }
    if(Object.keys(tryMap).length >= 20){ dateRowIdx = rIdx; colDateMap = tryMap; break; }
  }
  if(dateRowIdx === -1){ Logger.log('_buildNoteMap: 個人備註找不到日期表頭列,略過'); return map; }

  for(var r=dateRowIdx+1; r<data.length; r++){
    var nameCell = String(data[r][0]).trim();
    if(!nameCell || nameCell === 'undefined') continue;

    var empId = '', empName = '';
    var idMatch = nameCell.match(/^(\d+)-(.+)$/);
    if(idMatch){
      empId = idMatch[1];
      var rest = idMatch[2].trim();
      var sp2 = rest.split(/\s+/);
      empName = (sp2.length >= 2) ? sp2.slice(0, sp2.length-1).join('') : rest;
    } else {
      empName = nameCell;
    }

    for(var c2 in colDateMap){
      var note = _cleanNote(data[r][parseInt(c2,10)]);
      if(!note) continue;
      var dateStr = colDateMap[c2];
      if(empId)   map[dateStr + '|id|' + empId]   = note;
      if(empName) map[dateStr + '|nm|' + empName] = note;
    }
  }
  Logger.log('_buildNoteMap: 建立 ' + Object.keys(map).length + ' 筆備註鍵');
  return map;
}

// 確保 fact_schedule 第 Q 欄(17)表頭為 personal_note,並清掉 Q1:S1 舊廢資料(冪等)
function _ensureNoteHeader(sh){
  var q1 = sh.getRange(1, 17).getValue();
  if(q1 === 'personal_note') return;
  sh.getRange(1, 17).setValue('personal_note');  // Q1 表頭
  sh.getRange(1, 18, 1, 2).clearContent();        // 清 R1:S1 廢資料
  Logger.log('_ensureNoteHeader: 已補 Q1 表頭 + 清 R1:S1');
}

// ====================================================================
// 一次性回填 202603~202606 個人備註（2026-06-25）
// 讀 Drive 既有轉檔 Sheet 的「個人備註」分頁 → 只回寫 fact_schedule 的 Q 欄。
// 不動工時/營收/dedup。可重複執行(idempotent)。在編輯器選此函式按「執行」即可。
// ====================================================================
function backfillNotes_202603_202606(){
  var TARGET = {'202603':1,'202604':1,'202605':1,'202606':1};
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  if(!sh){ Logger.log('ERROR: fact_schedule 不存在'); return; }
  _ensureNoteHeader(sh);

  // 1) 掃 Drive 資料夾,對每個 canonical 落在目標月份的 Sheet 建 note 索引(含店號前綴防撞號)
  var idxNote = {};
  var scanned = 0;
  var folder = DriveApp.getFolderById(FOLDER_ID_);
  var it = folder.getFiles();
  while(it.hasNext()){
    var f = it.next();
    var name = f.getName();
    var ck = _canonicalKey(name);           // 例 12-202606
    if(!ck) continue;
    var parts = ck.split('-');
    var store = parseInt(parts[0],10);
    var yyyymm = parts[1];
    if(!TARGET[yyyymm]) continue;
    var mime = f.getMimeType();
    if(mime !== 'application/vnd.google-apps.spreadsheet') continue;  // 只讀已轉檔的 Google Sheet
    var year = parseInt(yyyymm.substring(0,4),10);
    var month = parseInt(yyyymm.substring(4,6),10);
    var per = _buildNoteMap(SpreadsheetApp.open(f), year, month);
    // 鍵: 日期|id|empId , 日期|nm|姓名
    var cnt = 0;
    for(var k in per){ idxNote[store + '||' + k] = per[k]; cnt++; }
    scanned++;
    Logger.log('掃描 ' + name + ' (' + ck + ') → ' + cnt + ' 鍵');
  }
  Logger.log('共掃描 ' + scanned + ' 個 Sheet,note 索引 ' + Object.keys(idxNote).length + ' 鍵');

  // 2) 讀 fact_schedule,逐列比對,只覆寫命中者的 Q,其餘維持原值
  var lastRow = sh.getLastRow();
  if(lastRow < 2){ Logger.log('fact_schedule 無資料'); return; }
  var data = sh.getDataRange().getValues();
  var idx = _buildHeaderIdx(data[0]);
  var qCol = idx['personal_note'];          // 0-based
  var n = lastRow - 1;
  var qVals = sh.getRange(2, qCol+1, n, 1).getValues();  // 現有 Q(保留)

  var hit = 0, overwrite = 0;
  for(var r=1; r<data.length; r++){
    var row = data[r];
    var store = parseInt(row[idx['store_id']],10);
    var dk = _dateKey_(row[idx['date']]);
    if(!dk) continue;
    var empId = String(row[idx['employee_id']]||'').trim();
    var empName = String(row[idx['employee_name']]||'').trim();
    var note = '';
    if(empId)   note = idxNote[store + '||' + dk + '|id|' + empId]   || '';
    if(!note && empName) note = idxNote[store + '||' + dk + '|nm|' + empName] || '';
    if(note){
      hit++;
      if(qVals[r-1][0] !== note){ qVals[r-1][0] = note; overwrite++; }
    }
  }
  // 3) 一次寫回 Q 欄
  sh.getRange(2, qCol+1, n, 1).setValues(qVals);
  Logger.log('=== 回填完成：命中 ' + hit + ' 列,實際覆寫 ' + overwrite + ' 列 ===');
}

// 日期值 → 'yyyy-MM-dd'(相容 Date 物件與字串)
function _dateKey_(v){
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return m[1] + '-' + m[2] + '-' + m[3];
  var d = new Date(s);
  if(!isNaN(d)) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '';
}