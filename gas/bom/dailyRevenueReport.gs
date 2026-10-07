/**
 * dailyRevenueReport.gs — 每日營收早報（2026-10-07 新增）
 * 專案：「一鍵追加新品」（綁 BOM 本，script ID 1PkavxV6r3GZbEbS_B5ZWGL1ovm30T55XhllAuI1aglBe_RPERdP-DJnJ）
 * 經營者 2026-10-07 要求：每天早上 6 點，表列昨天各店營收＋合計、本月累計＋合計，並比較去年同期
 *   （去年同日、去年同星期都列）；資料要正確 → POS 每日流程整條提早，06:00 前跑完對肚肚修正。
 *
 * 每日排程（2026-10-07 起；台北時間，GAS 誤差 ±15 分）：
 *   02:30 dudooPOS_importYesterday  肚肚 → POS資料 Sheet（原 04:50；避開每月 1 日 03 點 autoUpdateFormulaRanges）
 *   04:15 dailySyncBigQuery_v2      Sheet → BigQuery（原 05~06）
 *   05:00 dudooCheck_daily          BigQuery vs 肚肚業績概況，自動修正（原 07:30）
 *   06:00 dailyRevenueReport        讀 BigQuery v_daily_net（＝POS 儀表板營收口徑）寄營收表（本檔）
 *   （排班 dailyRebuild_salesAndKpi 在排班專案，06~07 不變；Make 折扣 01:00 不變）
 *
 * 入口：
 *   dailyRevenueReport()：06:00 觸發。先確認今天的肚肚對帳已跑完；沒跑完就 15 分鐘後再試（最多 4 次），
 *     最後一次照寄並在信首標紅字。一天只寄一次。
 *   previewRevenueReport()：只寫執行記錄，不寄信。
 *   sendRevenueReportTest()：立刻寄一封【測試】（不記已寄、不檢查對帳）。
 *   zzSetupMorningSchedule_20261007()：一次性改排程（找不到原觸發器就整個停下、不改）。
 *   zzRestoreMorningSchedule_old()：回滾成原時間，並刪掉營收早報觸發器。
 */

var REVRPT = {
  TZ: 'Asia/Taipei',
  TO: '（通知信箱，見線上版）',
  STORES: [
    [1, '台中精明'], [2, '台中草悟道'], [3, '台北南京'], [4, '台北士林'], [5, '台南Focus'], [6, '新竹文化'],
    [7, '新北板橋'], [8, '新北新店'], [9, '桃園中壢'], [10, '桃園藝文'], [11, '吳寶春信義A13'], [12, '吳寶春高雄SKM']
  ],
  MAX_TRY: 4,
  RETRY_MIN: 15,
  PROP_SENT: 'REVRPT_SENT_DATE',
  PROP_TRY: 'REVRPT_TRY',
  RETRY_HANDLER: 'dailyRevenueReport',
  WD: ['日', '一', '二', '三', '四', '五', '六']
};

/* ═══════════ 日期小工具（字串 yyyy-MM-dd） ═══════════ */

function revAddDays_(ds, n) {
  var p = ds.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.toISOString().substring(0, 10);
}
function revWd_(ds) { var p = ds.split('-'); return REVRPT.WD[new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay()]; }
function revLastYear_(ds) {            // 去年同日；2/29 → 2/28
  var y = +ds.substring(0, 4) - 1, md = ds.substring(5);
  if (md === '02-29') md = '02-28';
  return y + '-' + md;
}
function revShort_(ds) { return ds.substring(5).replace('-', '/') + '（' + revWd_(ds) + '）'; }

/* ═══════════ 取數 ═══════════ */

/** 回 {y, m1, ly, lyM1, lyWd, rows:{store|date: net}} */
function revFetch_(today) {
  var y = revAddDays_(today, -1), m1 = y.substring(0, 8) + '01';
  var ly = revLastYear_(y), lyM1 = ly.substring(0, 8) + '01', lyWd = revAddDays_(y, -364);
  var rows = audit_bqAll_(
    'SELECT store_code, CAST(sale_date AS STRING), CAST(SUM(net) AS INT64) FROM ' + DGUARD.BQ + "v_daily_net` " +
    "WHERE store_code BETWEEN 1 AND 12 AND ((sale_date BETWEEN '" + m1 + "' AND '" + y + "') " +
    "OR (sale_date BETWEEN '" + lyM1 + "' AND '" + ly + "') OR sale_date = '" + lyWd + "') GROUP BY 1, 2");
  var map = {};
  rows.forEach(function (r) { map[parseInt(r[0], 10) + '|' + r[1]] = Number(r[2]) || 0; });
  return { y: y, m1: m1, ly: ly, lyM1: lyM1, lyWd: lyWd, map: map };
}

function revSum_(map, st, from, to) {
  var s = 0, d = from;
  while (d <= to) { s += map[st + '|' + d] || 0; d = revAddDays_(d, 1); }
  return s;
}

/** 整理成表格資料 */
function revBuild_(f) {
  var lines = [], tot = { y: 0, ly: 0, lyWd: 0, mtd: 0, lyMtd: 0 }, missing = [];
  REVRPT.STORES.forEach(function (s) {
    var st = s[0];
    var r = {
      name: s[1],
      y: f.map[st + '|' + f.y] || 0,
      hasY: (st + '|' + f.y) in f.map,
      ly: f.map[st + '|' + f.ly] || 0,
      lyWd: f.map[st + '|' + f.lyWd] || 0,
      mtd: revSum_(f.map, st, f.m1, f.y),
      lyMtd: revSum_(f.map, st, f.lyM1, f.ly)
    };
    if (!r.hasY) missing.push(s[1]);
    ['y', 'ly', 'lyWd', 'mtd', 'lyMtd'].forEach(function (k) { tot[k] += r[k]; });
    lines.push(r);
  });
  return { rows: lines, tot: tot, missing: missing };
}

/* ═══════════ 版面 ═══════════ */

function revN_(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function revPct_(a, b) {
  if (!b) return { t: '—', c: '#888' };
  var p = (a - b) / b * 100;
  return { t: (p > 0 ? '+' : '') + p.toFixed(1) + '%', c: p >= 0 ? '#0a7d32' : '#c62828' };
}

function revHtml_(f, b, warn) {
  var th = 'style="padding:6px 8px;border-bottom:2px solid #333;text-align:right;white-space:nowrap;font-size:13px"';
  var thL = 'style="padding:6px 8px;border-bottom:2px solid #333;text-align:left;white-space:nowrap;font-size:13px"';
  var td = function (v, bold, color) {
    return '<td style="padding:5px 8px;border-bottom:1px solid #ddd;text-align:right;white-space:nowrap;font-size:13px' +
      (bold ? ';font-weight:bold' : '') + (color ? ';color:' + color : '') + '">' + v + '</td>';
  };
  var tdL = function (v, bold) {
    return '<td style="padding:5px 8px;border-bottom:1px solid #ddd;white-space:nowrap;font-size:13px' + (bold ? ';font-weight:bold' : '') + '">' + v + '</td>';
  };
  var pctTd = function (a, c, bold) { var p = revPct_(a, c); return td(p.t, bold, p.c); };
  var tbl = 'style="border-collapse:collapse;margin:6px 0 18px"';

  var h = '<div style="font-family:-apple-system,\'PingFang TC\',\'Microsoft JhengHei\',sans-serif;color:#222">';
  if (warn) h += '<p style="color:#c62828;font-weight:bold">🔴 ' + warn + '</p>';

  // 表 1：昨日
  h += '<h3 style="margin:8px 0 2px">昨日營收　' + revShort_(f.y) + '</h3>';
  h += '<div style="font-size:12px;color:#666">去年同日 ' + revShort_(f.ly) + '／去年同星期 ' + revShort_(f.lyWd) + '</div>';
  h += '<table ' + tbl + '><tr><th ' + thL + '>門市</th><th ' + th + '>昨日</th><th ' + th + '>去年同日</th><th ' + th +
    '>增減</th><th ' + th + '>去年同星期</th><th ' + th + '>增減</th></tr>';
  b.rows.forEach(function (r) {
    h += '<tr>' + tdL(r.name) + td(r.hasY ? revN_(r.y) : '無資料', false, r.hasY ? '' : '#c62828') + td(revN_(r.ly)) +
      pctTd(r.y, r.ly) + td(revN_(r.lyWd)) + pctTd(r.y, r.lyWd) + '</tr>';
  });
  h += '<tr style="background:#f3f3f3">' + tdL('合計', true) + td(revN_(b.tot.y), true) + td(revN_(b.tot.ly), true) +
    pctTd(b.tot.y, b.tot.ly, true) + td(revN_(b.tot.lyWd), true) + pctTd(b.tot.y, b.tot.lyWd, true) + '</tr></table>';

  // 表 2：本月累計
  var mLabel = (+f.y.substring(5, 7)) + '/1～' + (+f.y.substring(5, 7)) + '/' + (+f.y.substring(8));
  h += '<h3 style="margin:8px 0 2px">本月累計　' + mLabel + '</h3>';
  h += '<div style="font-size:12px;color:#666">去年同期 ' + f.lyM1.replace(/-/g, '/') + '～' + f.ly.replace(/-/g, '/') + '</div>';
  h += '<table ' + tbl + '><tr><th ' + thL + '>門市</th><th ' + th + '>本月累計</th><th ' + th + '>去年同期</th><th ' + th + '>增減</th></tr>';
  b.rows.forEach(function (r) {
    h += '<tr>' + tdL(r.name) + td(revN_(r.mtd)) + td(revN_(r.lyMtd)) + pctTd(r.mtd, r.lyMtd) + '</tr>';
  });
  h += '<tr style="background:#f3f3f3">' + tdL('合計', true) + td(revN_(b.tot.mtd), true) + td(revN_(b.tot.lyMtd), true) +
    pctTd(b.tot.mtd, b.tot.lyMtd, true) + '</tr></table>';

  h += '<div style="font-size:12px;color:#666;line-height:1.6">營收＝淨營收（已扣折扣），與 POS 儀表板、肚肚業績概況「營業額」同口徑。<br>' +
    '去年同星期＝往前推 364 天（同一個星期幾）。去年為 0 的店顯示「—」。<br>' +
    '<a href="https://diybc-training.onrender.com/static/pos-dashboard.html">開 POS 儀表板</a></div></div>';
  return h;
}

function revText_(f, b, warn) {
  var out = [];
  if (warn) out.push('🔴 ' + warn, '');
  out.push('昨日營收 ' + revShort_(f.y) + '（去年同日 ' + revShort_(f.ly) + '／去年同星期 ' + revShort_(f.lyWd) + '）');
  b.rows.forEach(function (r) {
    out.push(r.name + '　' + (r.hasY ? revN_(r.y) : '無資料') + '　同日 ' + revN_(r.ly) + ' ' + revPct_(r.y, r.ly).t +
      '　同星期 ' + revN_(r.lyWd) + ' ' + revPct_(r.y, r.lyWd).t);
  });
  out.push('合計　' + revN_(b.tot.y) + '　同日 ' + revN_(b.tot.ly) + ' ' + revPct_(b.tot.y, b.tot.ly).t +
    '　同星期 ' + revN_(b.tot.lyWd) + ' ' + revPct_(b.tot.y, b.tot.lyWd).t, '');
  out.push('本月累計（去年同期 ' + f.lyM1 + '～' + f.ly + '）');
  b.rows.forEach(function (r) { out.push(r.name + '　' + revN_(r.mtd) + '　去年 ' + revN_(r.lyMtd) + ' ' + revPct_(r.mtd, r.lyMtd).t); });
  out.push('合計　' + revN_(b.tot.mtd) + '　去年 ' + revN_(b.tot.lyMtd) + ' ' + revPct_(b.tot.mtd, b.tot.lyMtd).t);
  return out.join('\n');
}

function revSubject_(f, b, test) {
  return (test ? '【測試】' : '') + '【每日營收】' + revShort_(f.y) + ' 合計 ' + revN_(b.tot.y) +
    '（去年同星期 ' + revPct_(b.tot.y, b.tot.lyWd).t + '）｜本月累計 ' + revN_(b.tot.mtd) +
    '（去年同期 ' + revPct_(b.tot.mtd, b.tot.lyMtd).t + '）';
}

/* ═══════════ 就緒檢查 ═══════════ */

/** 今天的肚肚對帳（dudooCheck_daily）跑完了嗎？回 {ok, why} */
function revCheckReady_(today) {
  var sh = SpreadsheetApp.openById(RECON_POS_SS_ID).getSheetByName(DGUARD.RESULT_TAB);
  if (!sh) return { ok: false, why: '找不到「' + DGUARD.RESULT_TAB + '」分頁' };
  var v = sh.getRange(2, 1, 1, 3).getDisplayValues()[0];
  if (String(v[0]).indexOf(today) !== 0) return { ok: false, why: '今天的肚肚對帳還沒跑（最後一次 ' + (v[0] || '無') + '）' };
  if (String(v[2]).indexOf('❌') >= 0) return { ok: false, why: '今天的肚肚對帳執行失敗' };
  return { ok: true, why: String(v[2]) };
}

function revClearRetryTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === REVRPT.RETRY_HANDLER && t.getEventType() === ScriptApp.EventType.CLOCK &&
        PropertiesService.getScriptProperties().getProperty('REVRPT_RETRY_ID') === t.getUniqueId()) {
      ScriptApp.deleteTrigger(t);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty('REVRPT_RETRY_ID');
}

/* ═══════════ 入口 ═══════════ */

function dailyRevenueReport() {
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), REVRPT.TZ, 'yyyy-MM-dd');
  revClearRetryTriggers_();
  if (props.getProperty(REVRPT.PROP_SENT) === today) { Logger.log('[營收早報] 今天已寄過'); return; }

  var tries = {};
  try { tries = JSON.parse(props.getProperty(REVRPT.PROP_TRY) || '{}'); } catch (e) { tries = {}; }
  var n = (tries[today] || 0) + 1;
  props.setProperty(REVRPT.PROP_TRY, JSON.stringify({ [today]: n }));

  try {
    var ready = revCheckReady_(today);
    var f = revFetch_(today), b = revBuild_(f);
    var notReady = !ready.ok || b.missing.length > 0;
    if (notReady && n < REVRPT.MAX_TRY) {
      var t = ScriptApp.newTrigger(REVRPT.RETRY_HANDLER).timeBased().after(REVRPT.RETRY_MIN * 60000).create();
      props.setProperty('REVRPT_RETRY_ID', t.getUniqueId());
      Logger.log('[營收早報] 第 ' + n + ' 次：資料未就緒（' + ready.why + (b.missing.length ? '；無資料：' + b.missing.join('、') : '') +
        '），' + REVRPT.RETRY_MIN + ' 分鐘後再試');
      return;
    }
    var warn = '';
    if (!ready.ok) warn = ready.why + '，數字可能還會被自動修正。';
    if (b.missing.length) warn += '昨天沒有資料的店：' + b.missing.join('、') + '（公休或資料未進來）。';
    MailApp.sendEmail({ to: REVRPT.TO, subject: revSubject_(f, b, false), body: revText_(f, b, warn), htmlBody: revHtml_(f, b, warn) });
    props.setProperty(REVRPT.PROP_SENT, today);
    Logger.log('[營收早報] ✅ 已寄出（第 ' + n + ' 次）：' + revSubject_(f, b, false));
  } catch (e) {
    Logger.log('[營收早報] ❌ ' + e + '\n' + (e && e.stack));
    if (n < REVRPT.MAX_TRY) {
      var t2 = ScriptApp.newTrigger(REVRPT.RETRY_HANDLER).timeBased().after(REVRPT.RETRY_MIN * 60000).create();
      props.setProperty('REVRPT_RETRY_ID', t2.getUniqueId());
    } else {
      MailApp.sendEmail(REVRPT.TO, '【每日營收】❌ 今天的營收表產生失敗（請把這封信轉給 Claude）', '錯誤：' + e + '\n\n' + ((e && e.stack) || ''));
    }
  }
}

function previewRevenueReport() {
  var today = Utilities.formatDate(new Date(), REVRPT.TZ, 'yyyy-MM-dd');
  var ready = revCheckReady_(today);
  var f = revFetch_(today), b = revBuild_(f);
  Logger.log('就緒：' + JSON.stringify(ready));
  Logger.log(revSubject_(f, b, true) + '\n\n' + revText_(f, b, ready.ok ? '' : ready.why));
  return { ready: ready, tot: b.tot, missing: b.missing };
}

function sendRevenueReportTest() {
  var today = Utilities.formatDate(new Date(), REVRPT.TZ, 'yyyy-MM-dd');
  var f = revFetch_(today), b = revBuild_(f);
  MailApp.sendEmail({ to: REVRPT.TO, subject: revSubject_(f, b, true), body: revText_(f, b, ''), htmlBody: revHtml_(f, b, '') });
  Logger.log('已寄測試信：' + revSubject_(f, b, true));
}

/* ═══════════ 排程設定（一次性） ═══════════ */

function zzListTriggers() {
  var out = ScriptApp.getProjectTriggers().map(function (t) {
    return t.getHandlerFunction() + '｜' + t.getEventType() + '｜' + t.getUniqueId();
  });
  Logger.log(out.join('\n'));
  return out;
}

/** 把 POS 每日流程提早，並建 06:00 營收早報。原觸發器缺任何一個就整個停下、不改任何東西。 */
function zzSetupMorningSchedule_20261007() {
  var plan = [
    ['dudooPOS_importYesterday', 2, 30],
    ['dailySyncBigQuery_v2', 4, 15],
    ['dudooCheck_daily', 5, 0]
  ];
  var all = ScriptApp.getProjectTriggers();
  Logger.log('改前：\n' + zzListTriggers().join('\n'));
  var found = {};
  plan.forEach(function (p) {
    found[p[0]] = all.filter(function (t) { return t.getHandlerFunction() === p[0] && t.getEventType() === ScriptApp.EventType.CLOCK; });
    if (found[p[0]].length !== 1) throw new Error(p[0] + ' 的時間觸發器有 ' + found[p[0]].length + ' 個（預期 1 個），停下不改');
  });
  plan.forEach(function (p) {
    ScriptApp.deleteTrigger(found[p[0]][0]);
    ScriptApp.newTrigger(p[0]).timeBased().everyDays(1).atHour(p[1]).nearMinute(p[2]).inTimezone(REVRPT.TZ).create();
  });
  all.forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyRevenueReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyRevenueReport').timeBased().everyDays(1).atHour(6).nearMinute(0).inTimezone(REVRPT.TZ).create();
  Logger.log('改後：\n' + zzListTriggers().join('\n'));
  return '完成';
}

/** 回滾：恢復 04:50／05~06／07:30，刪營收早報觸發器。 */
function zzRestoreMorningSchedule_old() {
  var names = ['dudooPOS_importYesterday', 'dailySyncBigQuery_v2', 'dudooCheck_daily', 'dailyRevenueReport'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (names.indexOf(t.getHandlerFunction()) >= 0 && t.getEventType() === ScriptApp.EventType.CLOCK) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dudooPOS_importYesterday').timeBased().everyDays(1).atHour(4).nearMinute(50).inTimezone(REVRPT.TZ).create();
  ScriptApp.newTrigger('dailySyncBigQuery_v2').timeBased().everyDays(1).atHour(5).inTimezone(REVRPT.TZ).create();
  ScriptApp.newTrigger('dudooCheck_daily').timeBased().everyDays(1).atHour(7).nearMinute(30).inTimezone(REVRPT.TZ).create();
  Logger.log('已回滾：\n' + zzListTriggers().join('\n'));
}
