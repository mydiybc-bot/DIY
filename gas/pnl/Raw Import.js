/**
 * ============================================================
 * 烘焙 DIY 損益分析系統 — 會計原始檔直匯模組
 * 檔案：raw_import.gs
 * 版本：v1.0 (2026-09-11)
 * ============================================================
 *
 * 【這個檔在幹嘛】
 * 讓會計把「鼎新原始匯出檔」直接丟進收件匣資料夾，不必再人工轉成
 * YYMM_SS_損益表 標準格式。本模組負責：
 *   讀原始檔 → 動態解析 → 雙重對帳 → 產生 records → 寫 staging
 * 之後接既有的 commitStagingToFact（含三層安全網），一行都沒改。
 *
 * 【絕對不會做的事】
 * - 不寫 fact_pnl（只寫 staging，commit 仍由經營者親自按）
 * - 不改 dim_account
 * - 不動 parser.gs / m2_batch.gs / m4_safety.gs 任何既有函式
 * - 對帳不過不寫入，不自行調整數字湊平
 *
 * 【支援的兩種原始格式】
 * 形態 B  鼎新多月檔      檔名 11501-08損益表-1店   每科目佔 2 列
 * 形態 C  北3區手工模板    分頁名 9店-…／10店-…      每科目佔 1 列
 * 兩者共用同一套「動態表頭」演算法，差別只在列偏移量。
 *
 * 【為什麼一定要動態讀表頭】（2026-09-11 實測）
 *   1/2/6/7/8 店：3月在 I 欄、4月在 K 欄、合計在 O 欄
 *   3/4/5/11/12 店：3月在 J 欄、4月在 L 欄、合計在 P 欄
 *   北3區 9 店：12月在 AC 欄／10 店：12月在 AB 欄
 * 寫死欄位索引 = 整批數字錯位而且不會報錯。
 *
 * 【地雷 19 提醒】
 * Apps Script 編輯器實際執行的是「本檔第一個公開函式」。
 * 因此把唯讀的 aaaRawImportDryRun 排在最前面，誤按也不會寫任何東西。
 * ============================================================
 */


// ============================================================
// 【函式 0】aaaRawImportDryRun — 唯讀試跑（放第一個，誤按安全）
// ============================================================
/**
 * 在編輯器直接執行，只看 Log，不寫 staging、不寫 fact_pnl。
 * 執行前請先把收件匣資料夾 ID 填進下面這行。
 */
function aaaRawImportDryRun() {
  var FOLDER_ID = ''; // ← 收件匣資料夾 ID 填這裡
  if (!FOLDER_ID) {
    Logger.log('請先在 aaaRawImportDryRun() 內填入收件匣資料夾 ID');
    return;
  }
  var r = rawImportRun(FOLDER_ID, '', true);
  Logger.log('ok=' + r.ok);
  Logger.log('讀到檔案：' + r.filesRead.join(', '));
  Logger.log('店月數=' + r.cells.length + '　筆數=' + r.recordCount);
  r.cells.forEach(function (c) {
    Logger.log(
      (c.ok ? '✅' : '❌') + ' ' + c.year_month + ' ' + c.store_id + '店' +
      '　營收=' + c.revenue + '　本期損益(原檔)=' + c.bottomLine +
      '　計算=' + c.calcBottom + '　筆數=' + c.records +
      (c.ok ? '' : '　← ' + c.reason)
    );
  });
  if (r.errors.length) Logger.log('錯誤：\n' + r.errors.join('\n'));
  if (r.warnings.length) Logger.log('警告：\n' + r.warnings.join('\n'));
}


// ============================================================
// 【函式 1】rawImportRun — 主流程
// ============================================================
/**
 * @param {string} folderId  收件匣資料夾 ID
 * @param {string} ymsCsv    選填，只匯這些期間（民國整數，逗號分隔，例 '11507,11508'）
 *                           留空 = 全部有資料的月份
 * @param {boolean} dryRun   true = 只解析對帳、不寫 staging
 * @return {object}
 */
function rawImportRun(folderId, ymsCsv, dryRun) {
  var out = {
    ok: false,
    dryRun: !!dryRun,
    folderId: folderId,
    filesRead: [],
    skipped: [],
    cells: [],          // 店月矩陣
    recordCount: 0,
    stagingWrite: null,
    errors: [],
    warnings: []
  };

  try {
    if (!folderId) { out.errors.push('沒有指定收件匣資料夾 ID'); return out; }

    // ---- Step 1：載入 dim_account（唯一 schema 來源，每次現讀）----
    var dim = _riLoadDim_();
    if (!dim.ok) { out.errors.push(dim.reason); return out; }
    out.warnings.push('dim_account 讀到 ' + dim.names.length + ' 個科目');

    // ---- Step 2：掃資料夾 ----
    var scan = _riScanFolder_(folderId);
    if (!scan.ok) { out.errors.push(scan.reason); return out; }
    out.skipped = scan.skipped;
    if (!scan.files.length) { out.errors.push('收件匣沒有可解析的原始檔'); return out; }

    // ---- Step 3：逐檔解析 ----
    var ymFilter = _riParseYmFilter_(ymsCsv);
    var allRecords = [];

    for (var i = 0; i < scan.files.length; i++) {
      var f = scan.files[i];
      var pf = _riParseFile_(f, dim);

      if (!pf.ok) {
        out.errors.push('【' + f.name + '】' + pf.errors.join('；'));
        continue;
      }
      out.filesRead.push(f.name);
      pf.warnings.forEach(function (w) { out.warnings.push('【' + f.name + '】' + w); });

      for (var k = 0; k < pf.cells.length; k++) {
        var cell = pf.cells[k];
        if (ymFilter && ymFilter.indexOf(cell.year_month) < 0) continue;
        out.cells.push(cell);
        if (cell.ok) allRecords = allRecords.concat(cell._records);
      }
    }

    // ---- Step 4：硬性閘門 ----
    var badCells = out.cells.filter(function (c) { return !c.ok; });
    if (badCells.length) {
      out.errors.push(
        '對帳未通過 ' + badCells.length + ' 個店月，整批不寫入：' +
        badCells.map(function (c) { return c.year_month + '/' + c.store_id + '店(' + c.reason + ')'; }).join('、')
      );
      return out;
    }
    if (out.errors.length) return out;
    if (!out.cells.length) { out.errors.push('沒有任何符合條件的店月資料'); return out; }

    // 同店月重複（兩個原始檔涵蓋同一店月）
    var seen = {};
    for (var j = 0; j < out.cells.length; j++) {
      var key = out.cells[j].year_month + '_' + out.cells[j].store_id;
      if (seen[key]) { out.errors.push('重複的店月組合：' + key + '，請確認收件匣沒有放到重複檔'); }
      seen[key] = true;
    }
    if (out.errors.length) return out;

    out.recordCount = allRecords.length;

    // ---- Step 5：寫 staging（沿用 m2_batch.gs 既有函式，未修改）----
    if (dryRun) {
      out.ok = true;
      out.warnings.push('dry-run：未寫入 staging');
      return out;
    }

    out.stagingWrite = writeToStaging(allRecords);
    if (!out.stagingWrite || !out.stagingWrite.ok) {
      out.errors.push('staging 寫入失敗：' + ((out.stagingWrite && out.stagingWrite.reason) || '未知'));
      return out;
    }

    out.ok = true;
    return out;

  } catch (e) {
    out.errors.push('例外：' + e.toString());
    return out;
  } finally {
    // 回傳前把每格的明細records拔掉，避免 payload 過大
    out.cells.forEach(function (c) { delete c._records; });
  }
}


// ============================================================
// 【函式 2】_riScanFolder_ — 掃收件匣
// ============================================================
function _riScanFolder_(folderId) {
  var res = { ok: false, files: [], skipped: [], reason: '' };
  try {
    var folder = DriveApp.getFolderById(folderId);
    var iter = folder.getFiles();

    while (iter.hasNext()) {
      var f = iter.next();
      var name = f.getName();
      var mime = f.getMimeType();

      if (mime !== 'application/vnd.google-apps.spreadsheet') {
        res.skipped.push({ name: name, reason: '非 Google Sheets（請確認 Drive 自動轉檔設定已開啟）：' + mime });
        continue;
      }

      var kind = _riDetectKind_(name, f.getId());
      if (!kind.ok) { res.skipped.push({ name: name, reason: kind.reason }); continue; }

      res.files.push({ id: f.getId(), name: name, kind: kind });
    }
    res.ok = true;
    return res;
  } catch (e) {
    res.reason = '掃描收件匣失敗：' + e.toString();
    return res;
  }
}


// ============================================================
// 【函式 3】_riDetectKind_ — 判斷原始檔形態
// ============================================================
/**
 * 形態 B：檔名 (\d{3})(\d{2})(-(\d{2}))?損益表-(\d+)店
 * 形態 C：任一分頁名以「N店」開頭
 * 判斷不出來就 skip，絕不用猜的。
 */
function _riDetectKind_(name, fileId) {
  var base = String(name).replace(/\.(xlsx|xls)$/i, '');

  var mB = base.match(/^(\d{3})(\d{2})(?:-(\d{2}))?損益表-(\d+)店/);
  if (mB) {
    return { ok: true, mode: 'multi', roc: parseInt(mB[1], 10), storeId: parseInt(mB[4], 10) };
  }

  // 形態 C：看分頁名
  try {
    var sheets = SpreadsheetApp.openById(fileId).getSheets();
    var targets = [];
    for (var i = 0; i < sheets.length; i++) {
      var sn = sheets[i].getName();
      var mS = sn.match(/^(\d+)店/);
      if (mS) targets.push({ sheetName: sn, storeId: parseInt(mS[1], 10) });
    }
    if (targets.length) {
      var mY = base.match(/(20\d{2})/);
      if (!mY) return { ok: false, reason: '分頁看起來是北3區格式，但檔名找不到西元年（例 2026_北3區_…）' };
      return { ok: true, mode: 'beisan', roc: parseInt(mY[1], 10) - 1911, targets: targets };
    }
  } catch (e) {
    return { ok: false, reason: '開檔失敗：' + e.toString() };
  }

  return { ok: false, reason: '檔名與分頁都不符已知格式（形態 B：11501-08損益表-1店／形態 C：北3區模板）' };
}


// ============================================================
// 【函式 4】_riParseFile_ — 解析單一原始檔
// ============================================================
function _riParseFile_(file, dim) {
  var res = { ok: false, cells: [], errors: [], warnings: [], unknownAccounts: [] };
  try {
    var ss = SpreadsheetApp.openById(file.id);
    var kind = file.kind;

    var jobs = [];
    if (kind.mode === 'multi') {
      jobs.push({ sheet: ss.getSheets()[0], storeId: kind.storeId });
      res.warnings.push('使用第一個分頁：' + ss.getSheets()[0].getName());
    } else {
      kind.targets.forEach(function (t) {
        jobs.push({ sheet: ss.getSheetByName(t.sheetName), storeId: t.storeId });
      });
    }

    for (var i = 0; i < jobs.length; i++) {
      var grid = jobs[i].sheet.getDataRange().getValues();
      var p = _riParseGrid_(grid, dim);
      if (!p.ok) {
        res.errors.push(jobs[i].sheet.getName() + '：' + p.reason);
        if (p.unknownAccounts) {
          res.unknownAccounts = (res.unknownAccounts || []).concat(p.unknownAccounts);
        }
        continue;
      }
      p.warnings.forEach(function (w) { res.warnings.push(jobs[i].sheet.getName() + '：' + w); });

      var built = _riBuildCells_(p, jobs[i].storeId, kind.roc, dim, file.name);
      res.cells = res.cells.concat(built);
    }

    res.ok = res.errors.length === 0;
    return res;
  } catch (e) {
    res.errors.push('例外：' + e.toString());
    return res;
  }
}


// ============================================================
// 【函式 5】_riParseGrid_ — 核心：動態表頭 + 抽科目（純二維陣列運算）
// ============================================================
/**
 * 回傳 { ok, months:[], subj:{科目:{月:金額}}, totals:{科目:合計}, bottom:{月:本期損益}, warnings }
 * 形態 B 與形態 C 共用：差別只是月份表頭落在基準列的第幾列偏移，由掃描自動得出。
 */
function _riParseGrid_(grid, dim) {
  var res = { ok: false, reason: '', warnings: [], months: [], subj: {}, totals: {}, bottom: {} };

  // ---- 5-1 找基準表頭列（A 欄 = 「項目」）----
  var baseRow = -1;
  for (var r = 0; r < Math.min(grid.length, 20); r++) {
    if (String(grid[r][0] || '').trim() === '項目') { baseRow = r; break; }
  }
  if (baseRow < 0) { res.reason = '找不到表頭列（A 欄應有「項目」）'; return res; }

  // ---- 5-2 動態建立「月份 → (列偏移, 欄索引)」----
  var mmap = {}, total = null;
  for (var off = 0; off <= 2; off++) {
    var rr = baseRow + off;
    if (rr >= grid.length) break;
    for (var c = 0; c < grid[rr].length; c++) {
      var v = grid[rr][c];
      if (typeof v !== 'string') continue;
      var s = v.trim();
      var mm = s.match(/^(\d{1,2})月$/);
      if (mm) {
        var mo = parseInt(mm[1], 10);
        if (mo >= 1 && mo <= 12 && !mmap[mo]) mmap[mo] = { off: off, col: c };
      } else if (s === '合計' && !total) {
        total = { off: off, col: c };
      }
    }
  }

  var months = Object.keys(mmap).map(Number).sort(function (a, b) { return a - b; });
  if (!months.length) { res.reason = '表頭找不到任何「N月」欄位'; return res; }
  for (var i = 0; i < months.length; i++) {
    if (months[i] !== i + 1) { res.reason = '月份不連續：' + months.join(','); return res; }
  }
  res.months = months;
  if (!total) res.warnings.push('找不到「合計」欄，驗算 A 略過');

  // ---- 5-3 每個科目佔幾列：由「最大列偏移」推得 ----
  var maxOff = 0;
  months.forEach(function (mo) { if (mmap[mo].off > maxOff) maxOff = mmap[mo].off; });
  var rowSpan = maxOff + 1;   // 形態 B = 2，形態 C = 1
  res.warnings.push('每科目佔 ' + rowSpan + ' 列，月份欄位：' +
    months.map(function (mo) { return mo + '月→(+' + mmap[mo].off + ',c' + mmap[mo].col + ')'; }).join(' '));

  // ---- 5-4 掃科目列（D 欄）與本期損益列（A 欄）----
  var unknown = {};
  var subjRow = {};   // 科目 → 該科目所在列，供合計欄校正用
  for (var r2 = baseRow + rowSpan; r2 < grid.length; r2++) {
    var row = grid[r2];

    // 科目：D 欄（索引 3）
    var d = row.length > 3 ? row[3] : null;
    if (typeof d === 'string' && d.trim()) {
      var nm = d.trim();
      if (dim.map[nm]) {
        if (!res.subj[nm]) res.subj[nm] = {};
        if (subjRow[nm] === undefined) subjRow[nm] = r2;
        months.forEach(function (mo) {
          var cell = _riAt_(grid, r2 + mmap[mo].off, mmap[mo].col);
          res.subj[nm][mo] = (res.subj[nm][mo] || 0) + _riNum_(cell);
        });
      } else if (!_riIsAggregate_(nm)) {
        unknown[nm] = true;
      }
    }

    // 本期損益：A 欄（索引 0），取最後一個（鼎新檔頭尾各有一個，數值相同）
    var a = row.length > 0 ? row[0] : null;
    if (typeof a === 'string' && /^本期損益[：:]?$/.test(a.trim())) {
      var b = {};
      months.forEach(function (mo) {
        b[mo] = _riNum_(_riAt_(grid, r2 + mmap[mo].off, mmap[mo].col));
      });
      res.bottom = b;
    }
  }

  var unk = Object.keys(unknown);
  if (unk.length) {
    res.unknownAccounts = unk;   // 結構化回傳,前端才能列成清單並指路
    res.reason = '有 ' + unk.length + ' 個科目在「科目對照表」裡找不到：' + unk.join('、');
    return res;
  }
  if (!Object.keys(res.bottom).length) { res.reason = '找不到「本期損益」列，無法對帳'; return res; }
  if (!res.subj['銷貨收入']) { res.reason = '找不到「銷貨收入」科目'; return res; }

  // ---- 5-5 合計欄「用內容校正」，不盲信表頭位置 ----
  //
  // 【為什麼要這段】2026-09-11 實測：北3區手工模板 10 店分頁，
  // 「合計」表頭在第 30 欄，但實際合計公式在第 29 欄（表頭與資料錯開一格，
  // 原因是 10 店的 12 月欄位比 9 店少一欄，後面欄位跟著左移但表頭沒跟著改）。
  // 若盲信表頭，驗算 A 會拿「百分比欄」去跟金額比，得出一堆假錯誤。
  //
  // 做法：拿「銷貨收入」當基準，Σ各月 應該等於合計。
  // 在表頭位置 ±2 欄內找對得上的那一欄；找不到就停用驗算 A 並發警告，
  // 改以驗算 B（Σ金額×正負號 = 本期損益）為唯一硬性閘門。
  if (total) {
    var revSum = 0;
    months.forEach(function (mo) { revSum += res.subj['銷貨收入'][mo] || 0; });
    var revRow = subjRow['銷貨收入'];

    var picked = null;
    for (var dc = 0; dc <= 2 && picked === null; dc++) {
      var tries = (dc === 0) ? [total.col] : [total.col - dc, total.col + dc];
      for (var t = 0; t < tries.length; t++) {
        var cand = tries[t];
        if (cand < 0) continue;
        var val = _riNum_(_riAt_(grid, revRow + total.off, cand));
        if (Math.abs(val - revSum) <= 1 && revSum !== 0) { picked = cand; break; }
      }
    }

    if (picked === null) {
      res.warnings.push('⚠️ 合計欄與各月加總對不上（表頭位於第 ' + (total.col + 1) +
        ' 欄），可能是原檔表頭與資料錯位。驗算 A 已停用，改以驗算 B 為準');
      total = null;
    } else {
      if (picked !== total.col) {
        res.warnings.push('⚠️ 合計欄表頭位置與實際資料錯開 ' + (picked - total.col) +
          ' 欄，已依內容自動校正為第 ' + (picked + 1) + ' 欄（原檔版面問題，建議請製表者修正）');
      }
      total.col = picked;
      Object.keys(subjRow).forEach(function (nm2) {
        res.totals[nm2] = _riNum_(_riAt_(grid, subjRow[nm2] + total.off, total.col));
      });
    }
  }

  res.ok = true;
  return res;
}


// ============================================================
// 【函式 6】_riBuildCells_ — 對帳 + 產生 records
// ============================================================
function _riBuildCells_(p, storeId, roc, dim, sourceFile) {
  var cells = [];
  var loadDateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');

  // 驗算 A：每科目 Σ各月 = 原檔合計欄
  var totalMismatch = [];
  if (Object.keys(p.totals).length) {
    Object.keys(p.subj).forEach(function (nm) {
      var sum = 0;
      p.months.forEach(function (mo) { sum += p.subj[nm][mo] || 0; });
      var t = p.totals[nm];
      if (typeof t === 'number' && Math.abs(sum - t) > 1) {
        totalMismatch.push(nm + '(Σ月=' + sum + ' vs 合計=' + t + ')');
      }
    });
  }

  p.months.forEach(function (mo) {
    var ym = roc * 100 + mo;
    var revenue = p.subj['銷貨收入'] ? (p.subj['銷貨收入'][mo] || 0) : 0;
    var bottomLine = p.bottom[mo] || 0;

    // 空月（模板未來月份）→ 不產出、不報錯
    var anyValue = false;
    Object.keys(p.subj).forEach(function (nm) { if (p.subj[nm][mo]) anyValue = true; });
    if (!anyValue && !bottomLine) return;

    // 驗算 B：Σ(金額 × 正負號) = 原檔本期損益
    var calc = 0;
    dim.names.forEach(function (nm) {
      var amt = p.subj[nm] ? (p.subj[nm][mo] || 0) : 0;
      calc += amt * dim.sign[nm];
    });

    var ok = true, reason = '';
    if (Math.abs(calc - bottomLine) > 1) {
      ok = false;
      reason = '驗算B不符：計算 ' + Math.round(calc) + ' vs 原檔 ' + Math.round(bottomLine);
    } else if (totalMismatch.length) {
      ok = false;
      reason = '驗算A不符：' + totalMismatch.join('、');
    }

    // 產 records：以 dim_account 為固定 schema，缺的補 0
    var recs = [];
    if (ok) {
      dim.names.forEach(function (nm) {
        var d = dim.map[nm];
        var amt = p.subj[nm] ? (p.subj[nm][mo] || 0) : 0;
        recs.push({
          year_month: ym,
          year_ad: roc + 1911,
          quarter: Math.floor((mo - 1) / 3) + 1,
          half_year: mo <= 6 ? 1 : 2,
          store_id: storeId,
          account_code: d.account_code,
          account_name: nm,
          category_l1: d.category_l1,
          category_l2: d.category_l2,
          amount: amt,
          source_file: sourceFile,
          load_date: loadDateStr
        });
      });
    }

    cells.push({
      year_month: ym,
      store_id: storeId,
      revenue: Math.round(revenue),
      bottomLine: Math.round(bottomLine),
      calcBottom: Math.round(calc),
      records: recs.length,
      ok: ok,
      reason: reason,
      source_file: sourceFile,
      _records: recs
    });
  });

  return cells;
}


// ============================================================
// 【函式 7】_riLoadDim_ — 載入 dim_account（含正負號）
// ============================================================
/**
 * 正負號來源優先序：
 *   1. dim_account 若有「正負號」或「收入/支出」欄 → 用它
 *   2. 否則 account_code 為 A001 / G001 → +1，其餘 −1
 * 第 2 條與前端 _SIGN_PLUS_CODES 一致，且 2026-09-11 以 12 店 × 8 個月
 * 實測驗證過（96 個店月全數 0 元誤差）。
 */
function _riLoadDim_() {
  var out = { ok: false, reason: '', map: {}, names: [], sign: {} };
  try {
    var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.ACCOUNT);
    if (!sh) { out.reason = '找不到 dim_account 分頁'; return out; }

    var data = sh.getDataRange().getValues();
    if (data.length < 2) { out.reason = 'dim_account 是空的'; return out; }

    var h = data[0].map(function (x) { return String(x).trim(); });
    var iCode = h.indexOf('account_code');
    var iName = h.indexOf('鼎新原始科目名');
    var iL1 = h.indexOf('大類 L1');
    var iL2 = h.indexOf('子類 L2');
    if (iCode < 0 || iName < 0 || iL1 < 0 || iL2 < 0) {
      out.reason = 'dim_account 欄位表頭不符預期（需要 account_code / 鼎新原始科目名 / 大類 L1 / 子類 L2）';
      return out;
    }

    var iSign = -1;
    for (var c = 0; c < h.length; c++) {
      if (h[c].indexOf('正負') >= 0 || h[c].indexOf('收入/支出') >= 0) { iSign = c; break; }
    }

    for (var i = 1; i < data.length; i++) {
      var nm = String(data[i][iName] || '').trim();
      if (!nm) continue;
      var code = String(data[i][iCode] || '').trim();

      var sg;
      if (iSign >= 0) {
        var sv = String(data[i][iSign] || '').trim();
        if (sv === '收入' || sv === '+' || sv === '1' || sv === '+1') sg = 1;
        else if (sv === '支出' || sv === '-' || sv === '-1') sg = -1;
      }
      if (sg === undefined) sg = (code === 'A001' || code === 'G001') ? 1 : -1;

      out.map[nm] = { account_code: code, category_l1: data[i][iL1], category_l2: data[i][iL2] };
      out.sign[nm] = sg;
      out.names.push(nm);
    }

    if (!out.names.length) { out.reason = 'dim_account 沒有讀到任何科目'; return out; }
    out.ok = true;
    return out;
  } catch (e) {
    out.reason = '讀 dim_account 失敗：' + e.toString();
    return out;
  }
}


// ============================================================
// 【輔助函式】
// ============================================================
function _riAt_(grid, r, c) {
  if (r < 0 || r >= grid.length) return null;
  var row = grid[r];
  if (!row || c < 0 || c >= row.length) return null;
  return row[c];
}

/** 只認數字；'#DIV/0!'、空白、文字一律當 0（百分比欄不會被選到，這裡是保險） */
function _riNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var s = String(v).replace(/,/g, '').trim();
  if (!s || s.charAt(0) === '#') return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : n;
}

/** 判斷是否為匯總列名（合計／毛利／淨利／損益／總計等），這些不是葉節點科目 */
function _riIsAggregate_(s) {
  return s.indexOf('合計') >= 0 || s.indexOf('毛利') >= 0 || s.indexOf('淨利') >= 0 ||
         s.indexOf('損益') >= 0 || s.indexOf('總計') >= 0 || s.indexOf('每股盈餘') >= 0;
}

/** '11507,11508' → [11507,11508]；空字串 → null（代表不篩選） */
function _riParseYmFilter_(csv) {
  if (!csv) return null;
  var arr = String(csv).split(',').map(function (s) { return parseInt(String(s).trim(), 10); })
    .filter(function (n) { return !isNaN(n); });
  return arr.length ? arr : null;
}

// ============================================================
// 【函式 8】doPost — 拖拉上傳入口（2026-09-11 新增）
// ============================================================
/**
 * 為什麼要 doPost：
 *   儀表板原本用 JSONP（GET），網址長度有限，塞不下 Excel 檔。
 *   POST 才能把檔案送上來。本函式定義在這個檔案裡，
 *   所以「程式碼.gs」完全不用改，doGet 路由原封不動。
 *
 * 前端送：{ action:'raw_upload', adminToken:'…', dryRun:true/false,
 *          files:[{ name:'11501-08損益表-1店.xlsx', b64:'…' }, …] }
 */
function doPost(e) {
  var out = { ok: false, errors: [] };
  try {
    var body = JSON.parse(e.postData.contents || '{}');

    if (body.action === 'ping') {
      return _riJson_({ ok: true, pong: true, ts: new Date().toISOString() });
    }

    if (body.action === 'snapshot') {
      if (!verifyAdminToken(body.adminToken)) {
        return _riJson_({ ok: false, errors: ['管理員驗證失敗，請重新輸入管理密碼'] });
      }
      return _riJson_(riSnapshot());
    }

    if (body.action === 'commit_month') {
      if (!verifyAdminToken(body.adminToken)) {
        return _riJson_({ ok: false, errors: ['管理員驗證失敗，請重新輸入管理密碼'] });
      }
      return _riJson_(riCommitMonth(parseInt(body.ym, 10)));
    }

    if (body.action === 'ledger_upload') {
      if (!verifyAdminToken(body.adminToken)) {
        return _riJson_({ ok: false, errors: ['管理員驗證失敗，請重新輸入管理密碼'] });
      }
      return _riJson_(ledgerUploadToDrive(body.files, body.folderId));
    }

    if (body.action !== 'raw_upload') {
      out.errors.push('未知的 action：' + body.action);
      return _riJson_(out);
    }

    if (!verifyAdminToken(body.adminToken)) {
      out.errors.push('管理員驗證失敗，請重新輸入管理密碼');
      return _riJson_(out);
    }

    return _riJson_(rawImportFromUploads(body.files, body.yms, !!body.dryRun));

  } catch (err) {
    out.errors.push('doPost 例外：' + err.toString());
    return _riJson_(out);
  }
}

function _riJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// 【函式 9】rawImportFromUploads — 直接吃上傳檔（不需要收件匣資料夾）
// ============================================================
/**
 * 流程：base64 → Drive 暫存檔（轉成 Google Sheets）→ 解析對帳 → 立刻刪除暫存檔
 * 暫存檔一定會被刪：放在 finally，連例外中斷也會清掉。
 *
 * @param {Array}   files   [{name, b64}]
 * @param {string}  ymsCsv  選填期間篩選
 * @param {boolean} dryRun  true = 只對帳不寫 staging
 */
function rawImportFromUploads(files, ymsCsv, dryRun) {
  var out = {
    ok: false, dryRun: !!dryRun, filesRead: [], skipped: [], cells: [],
    recordCount: 0, stagingWrite: null, errors: [], warnings: [], timing: [],
    unknownAccounts: []
  };
  var tempIds = [];
  var t0 = new Date().getTime();

  try {
    if (!files || !files.length) { out.errors.push('沒有收到任何檔案'); return out; }
    if (files.length > 20) {
      out.errors.push('一次最多 20 個檔（收到 ' + files.length + ' 個）。請分兩批上傳。');
      return out;
    }

    // ---- dim_account：唯一 schema 來源，每次現讀 ----
    var dim = _riLoadDim_();
    if (!dim.ok) { out.errors.push(dim.reason); return out; }
    out.warnings.push('dim_account 讀到 ' + dim.names.length + ' 個科目');

    var ymFilter = _riParseYmFilter_(ymsCsv);
    var allRecords = [];

    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var tf = new Date().getTime();
      var conv = _riUploadToSheet_(f.name, f.b64);
      if (!conv.ok) { out.skipped.push({ name: f.name, reason: conv.reason }); continue; }
      tempIds.push(conv.id);
      var tConv = new Date().getTime() - tf;

      var kind = _riDetectKind_(f.name, conv.id);
      if (!kind.ok) { out.skipped.push({ name: f.name, reason: kind.reason }); continue; }

      var pf = _riParseFile_({ id: conv.id, name: f.name, kind: kind }, dim);
      if (!pf.ok) {
        out.errors.push('【' + f.name + '】' + pf.errors.join('；'));
        if (pf.unknownAccounts && pf.unknownAccounts.length) {
          out.unknownAccounts = out.unknownAccounts.concat(pf.unknownAccounts);
        }
        continue;
      }

      out.filesRead.push(f.name);
      out.timing.push(f.name + '：轉檔 ' + tConv + 'ms／解析 ' +
        (new Date().getTime() - tf - tConv) + 'ms');
      pf.warnings.forEach(function (w) { out.warnings.push('【' + f.name + '】' + w); });

      for (var k = 0; k < pf.cells.length; k++) {
        var cell = pf.cells[k];
        if (ymFilter && ymFilter.indexOf(cell.year_month) < 0) continue;
        out.cells.push(cell);
        if (cell.ok) allRecords = allRecords.concat(cell._records);
      }
    }

    // ---- 硬性閘門（與資料夾版完全相同）----
    var bad = out.cells.filter(function (c) { return !c.ok; });
    if (bad.length) {
      out.errors.push('對帳未通過 ' + bad.length + ' 個店月，整批不寫入：' +
        bad.map(function (c) { return c.year_month + '/' + c.store_id + '店(' + c.reason + ')'; }).join('、'));
      return out;
    }
    if (out.errors.length) return out;
    if (!out.cells.length) { out.errors.push('沒有解析出任何店月資料'); return out; }

    var seen = {};
    for (var j = 0; j < out.cells.length; j++) {
      var key = out.cells[j].year_month + '_' + out.cells[j].store_id;
      if (seen[key]) out.errors.push('重複的店月組合：' + key + '，請確認沒有上傳到重複的檔');
      seen[key] = true;
    }
    if (out.errors.length) return out;

    /* 未知科目去重,讓前端好呈現 */
    out.unknownAccounts = out.unknownAccounts.filter(function (v, i, a) { return a.indexOf(v) === i; });

    out.recordCount = allRecords.length;

    /* 【2026-10-12】逐月跟 fact_pnl 比對：完全相同的月份寫入時會跳過；有更動的列出差額 */
    try { out.diffByYm = _riDiffVsFact_(allRecords); }
    catch (eDiff) { out.warnings.push('跟現有資料比對失敗：' + eDiff); }

    if (dryRun) {
      out.ok = true;
      out.warnings.push('試算模式：未寫入 staging');
      return out;
    }

    /* 【2026-09-13 架構修正】
       之前是全部 3,744 筆一次進 staging → 一次 commit，實測 commit 跑超過 5 分鐘撞牆。
       現在改成：全部寫進「暫存區」_raw_import_hold，之後由 commit_month 逐月
       取出 → 寫 staging → commitStagingToFact（既有三層安全網原封不動）。
       每月約 468 筆、十幾秒，永遠不會撞六分鐘上限。 */
    var tw = new Date().getTime();
    out.holdWrite = _riHoldWrite_(allRecords);
    out.timing.push('寫暫存區：' + (new Date().getTime() - tw) + 'ms');
    if (!out.holdWrite.ok) {
      out.errors.push('暫存區寫入失敗：' + out.holdWrite.reason);
      return out;
    }
    out.holdYms = _riHoldSummary_().yms;

    out.ok = true;
    return out;

  } catch (e) {
    out.errors.push('例外：' + e.toString());
    return out;

  } finally {
    out.timing.push('總計：' + (new Date().getTime() - t0) + 'ms');
    // ★ 暫存檔一律刪除，例外中斷也會走到這裡
    for (var t = 0; t < tempIds.length; t++) {
      try { DriveApp.getFileById(tempIds[t]).setTrashed(true); } catch (ignore) {}
    }
    out.cells.forEach(function (c) { delete c._records; });
  }
}


// ============================================================
// 【函式 10】_riUploadToSheet_ — base64 → Drive 暫存 Google Sheets
// ============================================================
/**
 * Drive 進階服務在不同專案可能是 v2（Files.insert）或 v3（Files.create），
 * 這裡兩種都支援，執行時自動判斷，避免換版就壞。
 */
function _riUploadToSheet_(name, b64) {
  try {
    if (!b64) return { ok: false, reason: '檔案內容是空的' };

    var bytes = Utilities.base64Decode(b64);
    var blob = Utilities.newBlob(
      bytes,
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      name
    );
    var tmpName = '_tmp_rawimport_' + Utilities.getUuid().slice(0, 8);

    var file = null;
    if (typeof Drive !== 'undefined' && Drive.Files) {
      if (typeof Drive.Files.insert === 'function') {
        file = Drive.Files.insert({ title: tmpName, mimeType: MimeType.GOOGLE_SHEETS }, blob);
      } else if (typeof Drive.Files.create === 'function') {
        file = Drive.Files.create({ name: tmpName, mimeType: MimeType.GOOGLE_SHEETS }, blob);
      }
    }
    if (!file) {
      return { ok: false, reason: 'Drive 進階服務不可用，無法轉檔（請在 Apps Script「服務」確認已加入 Drive）' };
    }

    return { ok: true, id: file.id };

  } catch (e) {
    return { ok: false, reason: '轉檔失敗（檔案可能毀損或不是 Excel）：' + e.toString() };
  }
}


// ============================================================
// 【函式 11】ledgerUploadToDrive — 總分類帳原始檔上傳（2026-09-11 新增）
// ============================================================
/**
 * 這支跟損益表那條線「完全無關」，不寫 fact_pnl、不寫 fact_ledger、不碰 staging。
 * 它只做一件事：把鼎新總分類帳檔案放進 diybc-ledger 引擎在看的那個 Drive 資料夾。
 * 放完之後，還是要由經營者按儀表板上既有的「⬆️ 更新總分類帳」才會真的寫進 fact_ledger
 * （那條線本來就有逐子科目對帳閘門）。
 *
 * 順手擋掉 ledger-update skill 記錄的兩個地雷：
 *   地雷 5：店號靠檔名抓。檔名不含 -N店 或年度 → 引擎會靜默忽略。這裡先擋下。
 *   地雷 6：每店每年只能有一個檔。同店同年的舊檔自動丟垃圾桶（可從垃圾桶救回）。
 *
 * ⚠️ 地雷 4（全量重建）提醒：引擎每次會清空 fact_ledger 重建，
 *    所以資料夾必須保有「全部店、全部年度」的檔，少放一家那家就會被洗掉。
 *    本函式只新增/取代，不會刪掉其他店的檔。
 */
function ledgerUploadToDrive(files, folderId) {
  var out = { ok: false, uploaded: [], replaced: [], rejected: [], errors: [], folderFiles: 0 };

  try {
    if (!folderId) { out.errors.push('沒有指定總分類帳資料夾 ID'); return out; }
    if (!files || !files.length) { out.errors.push('沒有收到任何檔案'); return out; }
    if (files.length > 20) {
      out.errors.push('一次最多 20 個檔（收到 ' + files.length + ' 個）');
      return out;
    }

    var folder = DriveApp.getFolderById(folderId);

    // ---- 先全部驗檔名，有任何一個不合規就整批不上傳 ----
    var parsed = [];
    for (var i = 0; i < files.length; i++) {
      var k = _riLedgerKey_(files[i].name);
      if (!k.ok) { out.rejected.push({ name: files[i].name, reason: k.reason }); continue; }
      parsed.push({ file: files[i], key: k });
    }
    if (out.rejected.length) {
      out.errors.push('有 ' + out.rejected.length + ' 個檔名不合規，整批未上傳。' +
        '檔名必須同時含「-N店」與年度（例：11501-08總分類帳-1店.xlsx）');
      return out;
    }

    // ---- 同店同年重複 ----
    var seen = {};
    for (var j = 0; j < parsed.length; j++) {
      var kk = parsed[j].key.storeId + '_' + parsed[j].key.roc;
      if (seen[kk]) {
        out.errors.push('這批裡有兩個檔都是 ' + parsed[j].key.storeId + ' 店 ' +
          parsed[j].key.roc + ' 年，請只留一個');
        return out;
      }
      seen[kk] = true;
    }

    // ---- 逐檔上傳 ----
    for (var n = 0; n < parsed.length; n++) {
      var f = parsed[n].file, key = parsed[n].key;

      // 同店同年的舊檔先丟垃圾桶（地雷 6）
      var it = folder.getFiles();
      while (it.hasNext()) {
        var old = it.next();
        var ok2 = _riLedgerKey_(old.getName());
        if (ok2.ok && ok2.storeId === key.storeId && ok2.roc === key.roc) {
          try { old.setTrashed(true); out.replaced.push(old.getName()); } catch (ignore) {}
        }
      }

      var blob = Utilities.newBlob(
        Utilities.base64Decode(f.b64),
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        f.name
      );
      folder.createFile(blob);
      out.uploaded.push(f.name);
    }

    // ---- 回報資料夾現況，讓經營者一眼看出有沒有缺店（地雷 4）----
    var inv = {}, cnt = 0, it2 = folder.getFiles();
    while (it2.hasNext()) {
      var ff = it2.next(); cnt++;
      var kx = _riLedgerKey_(ff.getName());
      if (kx.ok) inv[kx.storeId + '店-' + kx.roc + '年'] = true;
    }
    out.folderFiles = cnt;
    out.inventory = Object.keys(inv).sort();

    out.ok = true;
    return out;

  } catch (e) {
    out.errors.push('例外：' + e.toString());
    return out;
  }
}

/** 從總分類帳檔名抓 (店號, 民國年)，規則沿用 ledger_parser._store_year */
function _riLedgerKey_(name) {
  var base = String(name).replace(/\.(xlsx|xls)$/i, '');
  var mS = base.match(/-(\d+)店/);
  if (!mS) return { ok: false, reason: '檔名缺少「-N店」（例：…-1店.xlsx）' };

  var roc = null;
  var mY1 = base.match(/(\d{3})年度/);
  var mY2 = base.match(/^(\d{3})(\d{2})/);
  if (mY1) roc = parseInt(mY1[1], 10);
  else if (mY2) roc = parseInt(mY2[1], 10);
  if (!roc) return { ok: false, reason: '檔名缺少年度（需為「114年度」或開頭「11501-08」這種寫法）' };

  return { ok: true, storeId: parseInt(mS[1], 10), roc: roc };
}


// ============================================================
// 【暫存區 _raw_import_hold】上傳後的全部 records 先放這裡，逐月取用
// ============================================================
var RI_HOLD_SHEET = '_raw_import_hold';

/** 取 fact_pnl 的表頭當欄序（與 writeToStaging 同一套規則） */
function _riFactHeaders_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var fact = ss.getSheetByName(SHEET_NAMES.FACT);
  if (!fact) return null;
  return fact.getRange(1, 1, 1, fact.getLastColumn()).getValues()[0];
}

/** 整份覆蓋暫存區 */
function _riHoldWrite_(records) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var headers = _riFactHeaders_();
    if (!headers) return { ok: false, reason: '找不到 fact_pnl，無法取得標準表頭' };

    var sh = ss.getSheetByName(RI_HOLD_SHEET);
    if (!sh) {
      sh = ss.insertSheet(RI_HOLD_SHEET);
      ss.setActiveSheet(sh); ss.moveActiveSheet(ss.getNumSheets());
    } else {
      sh.clearContents();
    }
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (!records.length) return { ok: true, rowsWritten: 0 };

    var grid = records.map(function (rec) {
      return headers.map(function (h) { return rec[h] !== undefined ? rec[h] : ''; });
    });
    sh.getRange(2, 1, grid.length, headers.length).setValues(grid);
    SpreadsheetApp.flush();
    return { ok: true, rowsWritten: grid.length };
  } catch (e) {
    return { ok: false, reason: '例外：' + e.toString() };
  }
}

/** 讀暫存區 → [{欄名:值}] */
function _riHoldRead_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(RI_HOLD_SHEET);
  if (!sh || sh.getLastRow() < 2) return { headers: [], rows: [] };
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) continue;
    var o = {};
    headers.forEach(function (h, c) { o[h] = data[i][c]; });
    rows.push(o);
  }
  return { headers: headers, rows: rows };
}

/** 暫存區摘要：還有哪些期間沒 commit */
function _riHoldSummary_() {
  var hold = _riHoldRead_();
  var by = {};
  hold.rows.forEach(function (r) {
    var ym = Number(r.year_month);
    if (!by[ym]) by[ym] = { ym: ym, records: 0, stores: {} };
    by[ym].records++;
    by[ym].stores[r.store_id] = true;
  });
  var yms = Object.keys(by).map(Number).sort(function (a, b) { return a - b; }).map(function (ym) {
    return { ym: ym, records: by[ym].records, stores: Object.keys(by[ym].stores).length };
  });
  return { totalRows: hold.rows.length, yms: yms };
}

/** 某期間 commit 成功後，從暫存區移除（用「讀→過濾→整份重寫」，不用 deleteRows，避免靜默失敗） */
function _riHoldRemoveYm_(ym) {
  var hold = _riHoldRead_();
  var keep = hold.rows.filter(function (r) { return Number(r.year_month) !== ym; });
  return _riHoldWrite_(keep);
}


// ============================================================
// 【函式 12】riCommitMonth — 逐月 commit（前端迴圈呼叫）
// ============================================================
/**
 * 一次只處理一個期間（12 店 × 39 = 468 筆）：
 *   暫存區取該月 → writeToStaging → commitStagingToFact(覆蓋) → 成功才從暫存區移除
 * commitStagingToFact 是既有函式，未修改：備份、稽核、筆數核對、LockService 全部照舊。
 */
function riCommitMonth(ym) {
  var out = { ok: false, ym: ym, errors: [], timing: [] };
  var t0 = new Date().getTime();

  /* 【2026-09-13】前端會在「回應沒送到」時自動重試，所以同一個月可能同時被叫兩次。
     用「文件鎖」把整段串起來（commitStagingToFact 內部用的是「指令碼鎖」，兩者獨立，不會互卡）。
     等不到鎖就回錯，讓前端稍後再試。 */
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(150000)) {
    out.errors.push('另一個寫入正在進行，請稍後再試');
    return out;
  }

  try {
    if (!ym || isNaN(ym)) { out.errors.push('期間參數不正確'); return out; }

    var hold = _riHoldRead_();
    var subset = hold.rows.filter(function (r) { return Number(r.year_month) === ym; });
    if (!subset.length) {
      /* 冪等：這個月已經不在暫存區 ＝ 上一次其實成功了，只是回應沒送到。回 ok 讓前端打勾。 */
      out.ok = true;
      out.alreadyDone = true;
      out.remaining = _riHoldSummary_().yms;
      return out;
    }
    out.records = subset.length;
    out.stores = Object.keys(subset.reduce(function (a, r) { a[r.store_id] = 1; return a; }, {})).length;

    /* 【2026-10-12】跟 fact_pnl 完全相同 → 不覆蓋，直接從暫存區移除（紅線：相同不覆蓋） */
    var dv = _riDiffVsFact_(subset)[0];
    out.diff = { status: dv.status, newStores: dv.newStores, changedStores: dv.changedStores, diffCount: dv.diffCount };
    if (dv.status === 'same') {
      _riHoldRemoveYm_(ym);
      out.ok = true;
      out.skippedIdentical = true;
      out.remaining = _riHoldSummary_().yms;
      return out;
    }

    var t1 = new Date().getTime();
    var sw = writeToStaging(subset);
    out.timing.push('寫 staging ' + (new Date().getTime() - t1) + 'ms');
    if (!sw || !sw.ok) { out.errors.push('staging 寫入失敗：' + (sw && sw.reason)); return out; }

    var t2 = new Date().getTime();
    var cr = commitStagingToFact({ duplicateMode: 'overwrite' });
    out.timing.push('commit ' + (new Date().getTime() - t2) + 'ms');

    out.factPnlBefore = cr.factPnlBefore;
    out.factPnlAfter = cr.factPnlAfter;
    out.rowsCommitted = cr.rowsCommitted;
    out.rowsOverwritten = cr.rowsOverwritten;
    out.replacedBackupSheet = cr.replacedBackupSheet || (cr.backup && cr.backup.sheetName) || null;
    out.countMismatch = !!cr.countMismatch;

    if (!cr.ok) {
      out.errors.push('commit 未成功：' + ((cr.errors && cr.errors.join('；')) || '未知'));
      return out;   // ★ 不從暫存區移除，下次可以重跑這個月
    }

    _riHoldRemoveYm_(ym);
    try { if (typeof pnlClearFactCache_ === 'function') pnlClearFactCache_(); } catch (ignore) {}

    out.remaining = _riHoldSummary_().yms;
    out.ok = true;
    return out;

  } catch (e) {
    out.errors.push('例外：' + e.toString());
    return out;
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
    out.timing.push('總計 ' + (new Date().getTime() - t0) + 'ms');
  }
}


// ============================================================
// 【函式 13】riSnapshot — 現況快照（每一段都 try/catch，任何一段壞了其他照樣回）
// ============================================================
function riSnapshot() {
  var out = { ok: true, hold: null, staging: null, factPnl: null, audit: null, backups: [], errors: [] };
  var ss = SpreadsheetApp.openById(SHEET_ID);

  try { out.hold = _riHoldSummary_(); }
  catch (e) { out.errors.push('暫存區：' + e); }

  try {
    var st = ss.getSheetByName(M2_STAGING_SHEET_NAME);
    out.staging = { rows: st ? Math.max(0, st.getLastRow() - 1) : 0 };
  } catch (e) { out.errors.push('staging：' + e); }

  try {
    var fp = ss.getSheetByName(SHEET_NAMES.FACT);
    out.factPnl = { rows: fp ? Math.max(0, fp.getLastRow() - 1) : 0 };
  } catch (e) { out.errors.push('fact_pnl：' + e); }

  try {
    var au = ss.getSheetByName('audit_log');
    if (au && au.getLastRow() >= 2) {
      var lr = au.getLastRow(), lc = au.getLastColumn();
      out.audit = {
        header: au.getRange(1, 1, 1, lc).getValues()[0],
        last: au.getRange(lr, 1, 1, lc).getValues()[0].map(function (v) {
          return (v instanceof Date) ? Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss') : v;
        })
      };
    }
  } catch (e) { out.errors.push('audit_log：' + e); }

  try {
    out.backups = ss.getSheets().map(function (x) { return x.getName(); })
      .filter(function (n) { return n.indexOf('fact_pnl_replaced_') === 0 || n.indexOf('fact_pnl_backup') === 0; })
      .sort().slice(-3);
  } catch (e) { out.errors.push('備份清單：' + e); }

  return out;
}

// ============================================================
// 【函式 14】_riDiffVsFact_ — 這批 records 跟 fact_pnl 逐「期間 × 店 × 科目」比對（2026-10-12）
// ============================================================
/**
 * 只比這批有出現的（期間, 店）；檔案裡沒有的店（例：沒上傳北3區就沒有 9、10 店）不比、也不會被覆蓋。
 * 回傳每個期間一筆：
 *   status  'new'＝這批的店在 fact_pnl 都還沒有資料；'same'＝全部一模一樣；'changed'＝其他
 *   newStores / sameStores / changedStores、diffCount、diffs[{store_id, account_code, account_name, old, new, delta}]
 * 唯讀，不寫任何資料。
 */
function _riDiffVsFact_(records) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var fact = ss.getSheetByName(SHEET_NAMES.FACT);
  if (!fact) throw new Error('找不到 fact_pnl');
  var data = fact.getDataRange().getValues();
  var h = data[0];
  var iYm = h.indexOf('year_month'), iSt = h.indexOf('store_id'),
      iCode = h.indexOf('account_code'), iName = h.indexOf('account_name'), iAmt = h.indexOf('amount');
  if (iYm < 0 || iSt < 0 || iCode < 0 || iAmt < 0) throw new Error('fact_pnl 缺欄');

  var want = {}, newMap = {}, names = {};
  records.forEach(function (r) {
    var k = Number(r.year_month) + '_' + Number(r.store_id);
    want[k] = true;
    if (!newMap[k]) newMap[k] = {};
    newMap[k][r.account_code] = (newMap[k][r.account_code] || 0) + (Number(r.amount) || 0);
    names[r.account_code] = r.account_name;
  });

  var oldMap = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var k = Number(row[iYm]) + '_' + Number(row[iSt]);
    if (!want[k]) continue;
    if (!oldMap[k]) oldMap[k] = {};
    var c = row[iCode];
    oldMap[k][c] = (oldMap[k][c] || 0) + (Number(row[iAmt]) || 0);
    if (!names[c] && iName >= 0) names[c] = row[iName];
  }

  var byYm = {};
  Object.keys(want).forEach(function (k) {
    var parts = k.split('_'), ym = Number(parts[0]), st = Number(parts[1]);
    var e = byYm[ym] || (byYm[ym] = { ym: ym, stores: 0, newStores: [], sameStores: [], changedStores: [], diffCount: 0, diffs: [] });
    e.stores++;
    if (!oldMap[k]) { e.newStores.push(st); return; }
    var codes = {};
    Object.keys(oldMap[k]).forEach(function (c) { codes[c] = true; });
    Object.keys(newMap[k]).forEach(function (c) { codes[c] = true; });
    var n = 0;
    Object.keys(codes).sort().forEach(function (c) {
      var o = oldMap[k][c] || 0, nv = newMap[k][c] || 0;
      if (Math.abs(o - nv) > 0.5) {
        n++; e.diffCount++;
        e.diffs.push({ store_id: st, account_code: c, account_name: names[c] || '', old: o, 'new': nv, delta: nv - o });
      }
    });
    (n ? e.changedStores : e.sameStores).push(st);
  });

  return Object.keys(byYm).map(Number).sort(function (a, b) { return a - b; }).map(function (ym) {
    var e = byYm[ym];
    [e.newStores, e.sameStores, e.changedStores].forEach(function (l) { l.sort(function (a, b) { return a - b; }); });
    e.diffs.sort(function (a, b) { return a.store_id - b.store_id || (a.account_code < b.account_code ? -1 : 1); });
    e.status = (e.newStores.length === e.stores) ? 'new'
             : (e.newStores.length || e.changedStores.length) ? 'changed' : 'same';
    return e;
  });
}

/** 驗證 _riDiffVsFact_（唯讀）：拿 fact_pnl 現有 11509 當「相同」、改一格當「有更動」、換成不存在的期間當「新月份」 */
function aaaTestDiffVsFact() {
  var data = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAMES.FACT).getDataRange().getValues();
  var h = data[0], recs = [];
  for (var i = 1; i < data.length; i++) {
    var o = {}; h.forEach(function (k, c) { o[k] = data[i][c]; });
    if (Number(o.year_month) === 11509) recs.push(o);
  }
  var same = _riDiffVsFact_(recs)[0];
  var mod = recs.map(function (r) { return JSON.parse(JSON.stringify(r)); });
  mod[0].amount = Number(mod[0].amount) + 100;
  var ch = _riDiffVsFact_(mod)[0];
  var nw = _riDiffVsFact_(recs.map(function (r) { var x = JSON.parse(JSON.stringify(r)); x.year_month = 19912; return x; }))[0];
  var res = { n: recs.length, same: [same.status, same.stores, same.diffCount], changed: [ch.status, ch.changedStores, ch.diffCount, ch.diffs[0]], fresh: [nw.status, nw.newStores.length] };
  Logger.log(JSON.stringify(res));
  return res;
}
