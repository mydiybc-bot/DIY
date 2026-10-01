// ====================================================================
// fact_note：備註獨立表（2026-09-10 上線，schedule-v5.6-note2）
// 目的：fact_schedule 一列＝一個人的一天班，員工當天沒排班就沒有那一列，
//       備註無處可掛而蒸發；「日期備註」分頁更是從未被讀取過。
//       本區塊「另外加一張表」，fact_schedule 的 schema 與既有讀寫一個字都不動。
// 紅線：不得修改 _toDateStr / _buildNoteMap / _cleanNote / _parse 的既有演算法。
// ====================================================================

var FACT_NOTE_SHEET_  = 'fact_note';
var FACT_NOTE_HEADER_ = ['date','store_id','store_name','scope','employee_id','employee_name','note','has_shift','source_file'];
var NOTE_SHEET_PERSON_ = '個人備註';
var NOTE_SHEET_DATE_   = '日期備註';
var NOTE_BACKFILL_BATCH_ = 12;   // 每次執行處理幾個檔（GAS 6 分鐘上限保護）
var NOTE_BACKFILL_START_ = '202603';
var NOTE_BACKFILL_PROP_  = 'fact_note_backfill_done';

// _parse 每次執行後把本檔的 fact_note 統計放這裡，供 backfill 取用
var _LAST_NOTE_STAT_ = {person:0, date:0, blocked:false};

// 唯讀核對工具：印出 fact_note 每店每月列數
function verifyFactNote(){
  var sh = _ensureFactNoteSheet_();
  if (sh.getLastRow() < 2){ Logger.log('fact_note 無資料'); return; }
  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, FACT_NOTE_HEADER_.length).getValues();
  var agg = {};
  for (var i = 0; i < n; i++){
    var dk = _dateKey_(vals[i][0]); if (!dk) continue;
    var key = parseInt(vals[i][1], 10) + '-' + dk.substring(0, 4) + dk.substring(5, 7);
    if (!agg[key]) agg[key] = {person: 0, date: 0, off: 0};
    if (String(vals[i][3]) === 'date') agg[key].date++;
    else {
      agg[key].person++;
      if (parseInt(vals[i][7], 10) === 0) agg[key].off++;
    }
  }
  var keys = Object.keys(agg).sort();
  Logger.log('=== fact_note 每店每月列數（共 ' + n + ' 列）===');
  for (var j = 0; j < keys.length; j++){
    Logger.log(keys[j] + ' → date ' + agg[keys[j]].date + ' / person ' + agg[keys[j]].person +
               '（其中 has_shift=0 為 ' + agg[keys[j]].off + '）');
  }
}


// --------------------------------------------------------------------
// 歷史回填（2026-03 起）
// 只寫 fact_note,一列 fact_schedule 都不碰（_parse 本身不寫 fact_schedule,
// 寫入是由 _processFile 呼叫 _appendRows 完成,本函式不呼叫它）。
// 冪等：每檔先 purge 再寫。
// 分批：GAS 單次 6 分鐘上限,每次處理 NOTE_BACKFILL_BATCH_ 檔,進度存 ScriptProperties,
//       全部跑完會自動清除進度並輸出三段 Log,下次呼叫即為完整重跑。
// --------------------------------------------------------------------
function backfillFactNote_all(){
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)){
    Logger.log('backfillFactNote_all: 已有處理中,跳過本次');
    return;
  }
  try {
    _ensureFactNoteSheet_();
    var sp = PropertiesService.getScriptProperties();
    var done = {};
    try { done = JSON.parse(sp.getProperty(NOTE_BACKFILL_PROP_) || '{}'); } catch(e){ done = {}; }

    // 1) 盤點 Drive
    var folder = DriveApp.getFolderById(FOLDER_ID_);
    var it = folder.getFiles();
    var targets = [], skipped = [];
    while (it.hasNext()){
      var f = it.next();
      var name = f.getName();
      var ck = _canonicalKey(name);
      if (!ck) continue;
      var parts = ck.split('-');
      if (parts[1] < NOTE_BACKFILL_START_) continue;
      if (f.getMimeType() !== 'application/vnd.google-apps.spreadsheet'){
        skipped.push('[跳過]   ' + ck + '  (' + name + ' 未轉檔) → 未回填');
        continue;
      }
      targets.push({f: f, ck: ck, name: name, store: parseInt(parts[0], 10), yyyymm: parts[1]});
    }
    targets.sort(function(a, b){ return a.ck < b.ck ? -1 : (a.ck > b.ck ? 1 : 0); });

    // 2) 分批處理
    var processed = 0, remain = 0;
    for (var i = 0; i < targets.length; i++){
      var t = targets[i];
      if (done[t.ck]) continue;
      if (processed >= NOTE_BACKFILL_BATCH_){ remain++; continue; }

      _purgeFactNoteByCanonical(t.store, t.yyyymm);
      _LAST_NOTE_STAT_ = {person: 0, date: 0, blocked: false};
      var before = _factNoteCount_();
      _parse(SpreadsheetApp.open(t.f), t.name);        // 只會寫 fact_note
      var wrote = _factNoteCount_() - before;

      if (_LAST_NOTE_STAT_.blocked){
        Logger.log('[版面異常] ' + t.ck + ' (' + t.name + ') → 整檔未寫入,請人工檢查該檔備註分頁');
      } else {
        Logger.log('[已掃描] ' + t.ck + ' (Google Sheet) → person ' + _LAST_NOTE_STAT_.person +
                   ' 列 / date ' + _LAST_NOTE_STAT_.date + ' 列 (實際寫入 ' + wrote + ' 列)');
        done[t.ck] = 1;
      }
      processed++;
    }
    sp.setProperty(NOTE_BACKFILL_PROP_, JSON.stringify(done));

    // 3) 跳過清單
    for (var s = 0; s < skipped.length; s++) Logger.log(skipped[s]);

    if (remain > 0){
      Logger.log('=== 本次處理 ' + processed + ' 檔,尚有 ' + remain + ' 檔未處理,請「再執行一次」 backfillFactNote_all() ===');
      return;
    }

    // 4) 缺口差集：fact_schedule 有資料 但 fact_note 未回填的 (店,月)
    var haveSched = _distinctStoreMonthFromFactSchedule_();
    var gaps = [];
    for (var k in haveSched){ if (!done[k]) gaps.push(k); }
    gaps.sort();
    if (gaps.length === 0){
      Logger.log('[缺口]   無。fact_schedule 的所有 (店,月) 都已回填 fact_note ✅');
    } else {
      Logger.log('[缺口]   fact_schedule 有資料但 fact_note 未回填的 (店,月) 共 ' + gaps.length + ' 組：' + gaps.join('、'));
    }
    Logger.log('=== backfillFactNote_all 全部完成,fact_note 現有 ' + _factNoteCount_() + ' 列 ===');
    sp.deleteProperty(NOTE_BACKFILL_PROP_);   // 清進度,下次呼叫即為完整重跑
  } finally {
    lock.releaseLock();
  }
}


// --------------------------------------------------------------------
// 交辦書 §7：動 fact_note 之前先備份 fact_schedule（冪等，已存在就跳過）
// 比照既有 backup_factschedule_20260609() 的作法
// --------------------------------------------------------------------
function backup_factschedule_20260910(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var src = ss.getSheetByName('fact_schedule');
  if (!src){ Logger.log('ERROR: fact_schedule not found'); return; }

  var NAME = 'fact_schedule_backup_20260910';
  var existing = ss.getSheetByName(NAME);
  if (existing){
    Logger.log('備份分頁已存在,跳過。備份列數(含 header): ' + existing.getLastRow());
    return;
  }
  var backup = src.copyTo(ss);
  backup.setName(NAME);

  var srcRows = src.getLastRow();
  var bakRows = backup.getLastRow();
  Logger.log('=== 備份完成 ' + NAME + ' ===');
  Logger.log('原始 fact_schedule 列數 (含 header): ' + srcRows);
  Logger.log('備份分頁列數 (含 header): ' + bakRows);
  Logger.log('列數一致: ' + (srcRows === bakRows));
}


// --------------------------------------------------------------------
// 嚴格日期判定
// 刻意「不」沿用 _toDateStr()：它有一段寬鬆 fallback ——
//   純數字 1~31 → 視為該月第 N 天 ——
// 而本區塊一律從 c=0 起掃，若某店 A 欄放序號/列號會被誤判成日期並悄悄錯位。
// _toDateStr 仍在服務班表主解析，維持原狀不動。
// --------------------------------------------------------------------
function _noteDateStr_(v){
  if (v instanceof Date && !isNaN(v)) {
    return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

// 該月天數（自檢閘門用）
function _daysInMonth_(year, month){
  return new Date(year, month, 0).getDate();
}

// --------------------------------------------------------------------
// 掃描日期表頭列 —— 一律 c=0 起掃到底，不假設 A 欄是不是標籤欄。
// 「個人備註」A 欄是員工標籤（掃不到日期，自動從 index 1 開始）；
// 「日期備註」A 欄直接就是日期（自動從 index 0 開始）。
// 兩種版面、以及未來任何店的變形，都靠日期判定自己決定，不寫死偏移。
// 回傳 {rowIdx, map:{colIdx: 'YYYY-MM-DD'}, cnt}；掃不到回 null。
// --------------------------------------------------------------------
function _scanNoteDateRow_(data){
  var best = null;
  var limit = Math.min(5, data.length);
  for (var r = 0; r < limit; r++){
    var map = {}, cnt = 0;
    var row = data[r] || [];
    for (var c = 0; c < row.length; c++){
      var ds = _noteDateStr_(row[c]);
      if (ds){ map[c] = ds; cnt++; }
    }
    if (cnt > 0 && (!best || cnt > best.cnt)) best = {rowIdx: r, map: map, cnt: cnt};
  }
  return best;
}

// 取分頁（找不到回 null）
function _sheetByName_(ss, name){
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++){
    if (sheets[i].getName() === name) return sheets[i];
  }
  return null;
}

// 從本次 _parse 產出的班表列，建 has_shift 索引（date|id|xxx 與 date|nm|xxx 雙鍵，比照 _noteFor）
// 注意：班表格子寫「月休 / 國定假日 / 未到職」時 _parse 會 skip、不產生列，
//       因此這三種情形一律 has_shift=0，正是我們要的（那天確實沒上班）。
function _buildShiftSet_(schedRows){
  var s = {};
  for (var i = 0; i < schedRows.length; i++){
    var d  = schedRows[i][0];
    var id = String(schedRows[i][3] || '').trim();
    var nm = String(schedRows[i][4] || '').trim();
    if (!d) continue;
    if (id) s[d + '|id|' + id] = 1;
    if (nm) s[d + '|nm|' + nm] = 1;
  }
  return s;
}

// --------------------------------------------------------------------
// scope=person：讀「個人備註」分頁
// --------------------------------------------------------------------
function _buildPersonNoteRows_(ss, fname, storeMeta, year, month, shiftSet, warn){
  var sh = _sheetByName_(ss, NOTE_SHEET_PERSON_);
  if (!sh){ warn.push('無「' + NOTE_SHEET_PERSON_ + '」分頁,略過'); return []; }

  var data = sh.getDataRange().getValues();
  if (data.length < 3) return [];

  var scan = _scanNoteDateRow_(data);
  if (!scan){
    warn.push('[版面異常] ' + fname + '／' + NOTE_SHEET_PERSON_ + ' 找不到日期列');
    return null;
  }
  var need = _daysInMonth_(year, month);
  var got  = Object.keys(scan.map).length;
  if (got !== need){
    warn.push('[版面異常] ' + fname + '／' + NOTE_SHEET_PERSON_ + ' 偵測到 ' + got + ' 個日期欄,預期 ' + need + ' 個');
    return null;
  }

  var rows = [];
  for (var r = scan.rowIdx + 1; r < data.length; r++){
    var nameCell = String(data[r][0]).trim();
    if (!nameCell || nameCell === 'undefined') continue;

    // 拆法完全比照 _parse / _buildNoteMap,確保 key 對得上
    var empId = '', empName = '';
    var idMatch = nameCell.match(/^(\d+)-(.+)$/);
    if (idMatch){
      empId = idMatch[1];
      var rest = idMatch[2].trim();
      var sp2 = rest.split(/\s+/);
      empName = (sp2.length >= 2) ? sp2.slice(0, sp2.length - 1).join('') : rest;
    } else {
      empName = nameCell;
    }
    if (!empName) continue;

    for (var c in scan.map){
      var note = _cleanNote(data[r][parseInt(c, 10)]);   // 沿用既有清洗規則,不另寫一套
      if (!note) continue;
      var dateStr = scan.map[c];
      var hit = false;
      if (empId   && shiftSet[dateStr + '|id|' + empId])   hit = true;
      if (!hit && empName && shiftSet[dateStr + '|nm|' + empName]) hit = true;
      rows.push([dateStr, storeMeta.id, storeMeta.name, 'person', empId, empName, note, (hit ? 1 : 0), fname]);
    }
  }
  return rows;
}

// --------------------------------------------------------------------
// scope=date：讀「日期備註」分頁
// 日期列之下的所有列都要吃(不假設只有一列),同一天多列用「；」串接
// --------------------------------------------------------------------
function _buildDateNoteRows_(ss, fname, storeMeta, year, month, warn){
  var sh = _sheetByName_(ss, NOTE_SHEET_DATE_);
  if (!sh){ warn.push('無「' + NOTE_SHEET_DATE_ + '」分頁,略過'); return []; }

  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  var scan = _scanNoteDateRow_(data);
  if (!scan){
    warn.push('[版面異常] ' + fname + '／' + NOTE_SHEET_DATE_ + ' 找不到日期列');
    return null;
  }
  var need = _daysInMonth_(year, month);
  var got  = Object.keys(scan.map).length;
  if (got !== need){
    warn.push('[版面異常] ' + fname + '／' + NOTE_SHEET_DATE_ + ' 偵測到 ' + got + ' 個日期欄,預期 ' + need + ' 個');
    return null;
  }

  var byDate = {};
  for (var r = scan.rowIdx + 1; r < data.length; r++){
    for (var c in scan.map){
      var note = _cleanNote(data[r][parseInt(c, 10)]);
      if (!note) continue;
      var d = scan.map[c];
      if (!byDate[d]) byDate[d] = [];
      if (byDate[d].indexOf(note) < 0) byDate[d].push(note);   // 同日多列去重保序
    }
  }

  var keys = Object.keys(byDate).sort();
  var rows = [];
  for (var k = 0; k < keys.length; k++){
    rows.push([keys[k], storeMeta.id, storeMeta.name, 'date', '', '', byDate[keys[k]].join('；'), '', fname]);
  }
  return rows;
}

// 統包：任一分頁版面自檢未過 → 回 null,整檔不寫入
function _collectFactNoteRows_(ss, fname, storeMeta, year, month, schedRows, warn){
  var shiftSet = _buildShiftSet_(schedRows);
  var p = _buildPersonNoteRows_(ss, fname, storeMeta, year, month, shiftSet, warn);
  var d = _buildDateNoteRows_(ss, fname, storeMeta, year, month, warn);
  if (p === null || d === null){
    _LAST_NOTE_STAT_ = {person: 0, date: 0, blocked: true};
    return null;
  }
  _LAST_NOTE_STAT_ = {person: p.length, date: d.length, blocked: false};
  return p.concat(d);
}

// --------------------------------------------------------------------
// fact_note 分頁維護
// --------------------------------------------------------------------
function _ensureFactNoteSheet_(){
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName(FACT_NOTE_SHEET_);
  if (!sh){
    sh = ss.insertSheet(FACT_NOTE_SHEET_);
    Logger.log('_ensureFactNoteSheet_: 已建立 ' + FACT_NOTE_SHEET_ + ' 分頁');
  }
  var W = FACT_NOTE_HEADER_.length;
  var first = sh.getRange(1, 1, 1, W).getValues()[0];
  var ok = true;
  for (var i = 0; i < W; i++){
    if (String(first[i]).trim() !== FACT_NOTE_HEADER_[i]) ok = false;
  }
  if (!ok){
    sh.getRange(1, 1, 1, W).setValues([FACT_NOTE_HEADER_]);
    sh.setFrozenRows(1);
    Logger.log('_ensureFactNoteSheet_: 已寫入/修正 header');
  }
  return sh;
}

function _appendNoteRows_(rows){
  if (!rows || rows.length === 0) return 0;
  var sh = _ensureFactNoteSheet_();
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, FACT_NOTE_HEADER_.length).setValues(rows);
  return rows.length;
}

function _factNoteCount_(){
  var sh = _ensureFactNoteSheet_();
  var n = sh.getLastRow() - 1;
  return n < 0 ? 0 : n;
}

// --------------------------------------------------------------------
// purge 鏡像：比照 _purgeFactScheduleByCanonical,但改用 memory 過濾法重寫
// （讀 → 過濾 → clearContent → setValues → 驗殘留為 0），不用逐列 deleteRow。
// 鎖：唯二呼叫點都在 _uploadSchedule 系列的 tryLock 區塊內,故此處不重複加鎖。
// --------------------------------------------------------------------
function _purgeFactNoteByCanonical(storeNum, yyyymm){
  var sh = _ensureFactNoteSheet_();
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  var W = FACT_NOTE_HEADER_.length;
  var height = lastRow - 1;
  var data = sh.getRange(2, 1, height, W).getValues();
  var targetKey = parseInt(storeNum, 10) + '-' + yyyymm;

  var keep = [];
  for (var i = 0; i < data.length; i++){
    var fn = String(data[i][8] || '').trim();          // I 欄 source_file
    if (_canonicalKey(fn) === targetKey) continue;
    keep.push(data[i]);
  }
  var removed = data.length - keep.length;
  if (removed === 0) return 0;

  sh.getRange(2, 1, height, W).clearContent();
  if (keep.length > 0) sh.getRange(2, 1, keep.length, W).setValues(keep);
  SpreadsheetApp.flush();

  // 驗殘留（讀回原高度,空列一律為 ''）
  var after = sh.getRange(2, 9, height, 1).getValues();
  var leftover = 0;
  for (var j = 0; j < after.length; j++){
    if (_canonicalKey(String(after[j][0] || '').trim()) === targetKey) leftover++;
  }
  if (leftover > 0){
    throw new Error('_purgeFactNoteByCanonical 殘留 ' + leftover + ' 列 (' + targetKey + ')');
  }
  Logger.log('_purgeFactNoteByCanonical: ' + targetKey + ' 刪除 ' + removed + ' 列,保留 ' + keep.length + ' 列');
  return removed;
}

// fact_schedule 的 distinct (store_id, 年月) → {'1-202609': 1, ...}
function _distinctStoreMonthFromFactSchedule_(){
  var out = {};
  var ss = SpreadsheetApp.openById(ANALYSIS_SS_ID);
  var sh = ss.getSheetByName('fact_schedule');
  if (!sh || sh.getLastRow() < 2) return out;
  var n = sh.getLastRow() - 1;
  var vals = sh.getRange(2, 1, n, 2).getValues();      // A=date, B=store_id
  for (var i = 0; i < n; i++){
    var dk = _dateKey_(vals[i][0]);
    if (!dk) continue;
    var sid = parseInt(vals[i][1], 10);
    if (isNaN(sid)) continue;
    out[sid + '-' + dk.substring(0, 4) + dk.substring(5, 7)] = 1;
  }
  return out;
}
