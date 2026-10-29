/**
 * 專案推動 API 的本機模擬後端（2026-10-18）：用 Node 直接跑 gas/project-hub/Code.gs＋seed.gs，
 * 試算表、密碼驗證、寄信全部用假的，讓 static/project-hub.html 在本機完整點測，不碰線上資料。
 *
 * 啟動（.claude/launch.json 的 project-hub-mock 也是這條）：
 *   node tools/project-hub-mock/mock.js            → http://127.0.0.1:8766
 * 頁面：http://127.0.0.1:8766/static/project-hub.html?api=/api
 *   密碼固定 test；閘門用 auth-center 的緊急繞過（localStorage dash_auth_project）。
 * 看假信：GET /mock/mails；看假試算表：GET /mock/sheets；直接跑 GAS 函式：GET /mock/run?fn=dailyReminder
 * 跟 GAS 行為不同的地方：沒有 1,000 列上限、LockService 永遠拿得到鎖。
 */
'use strict';
var http = require('http'), fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto'), url = require('url');
var PROPS = { SCAN_TOKEN: 'scan-test' };   /* 指令碼屬性（正式用 zz 暫時檔設 SCAN_TOKEN） */
var ROOT = path.resolve(__dirname, '..', '..'), GAS = path.join(ROOT, 'gas', 'project-hub');
var PORT = Number(process.env.PORT || 8766), PW = 'test';

/* ---------- 假試算表 ---------- */
function Sheet(name) { this.name = name; this.rows = []; this.frozen = 0; }
Sheet.prototype.getName = function () { return this.name; };
Sheet.prototype.setName = function (n) { this.name = n; return this; };
Sheet.prototype.getLastRow = function () { var n = this.rows.length; while (n > 0 && this.rows[n - 1].every(function (v) { return v === '' || v == null; })) n--; return n; };
Sheet.prototype.getLastColumn = function () { var m = 0; this.rows.forEach(function (r) { var n = r.length; while (n > 0 && (r[n - 1] === '' || r[n - 1] == null)) n--; if (n > m) m = n; }); return m; };
Sheet.prototype.setFrozenRows = function (n) { this.frozen = n; return this; };
Sheet.prototype.copyTo = function (ss) { var s = new Sheet(this.name + ' 的副本'); s.rows = this.rows.map(function (r) { return r.slice(); }); ss.sheets.push(s); return s; };
Sheet.prototype.cell_ = function (r, c) { while (this.rows.length < r) this.rows.push([]); var row = this.rows[r - 1]; while (row.length < c) row.push(''); return row; };
Sheet.prototype.getRange = function (r, c, nr, nc) {
  var sh = this; nr = nr || 1; nc = nc || 1;
  return {
    setNumberFormat: function () { return this; }, setFontWeight: function () { return this; },
    setValues: function (vals) { vals.forEach(function (vr, i) { var row = sh.cell_(r + i, c + nc - 1); vr.forEach(function (v, j) { row[c - 1 + j] = v == null ? '' : String(v); }); }); return this; },
    setValue: function (v) { sh.cell_(r, c)[c - 1] = v == null ? '' : String(v); return this; },
    getValues: function () { return this.getDisplayValues(); },
    getDisplayValues: function () { var out = []; for (var i = 0; i < nr; i++) { var row = sh.rows[r - 1 + i] || [], o = []; for (var j = 0; j < nc; j++) o.push(row[c - 1 + j] == null ? '' : String(row[c - 1 + j])); out.push(o); } return out; }
  };
};
function SS(name) { this.name = name; this.sheets = []; }
SS.prototype.getName = function () { return this.name; };
SS.prototype.getSheets = function () { return this.sheets; };
SS.prototype.getSheetByName = function (n) { return this.sheets.filter(function (s) { return s.name === n; })[0] || null; };
SS.prototype.insertSheet = function (n) { var s = new Sheet(n); this.sheets.push(s); return s; };
SS.prototype.deleteSheet = function (s) { this.sheets = this.sheets.filter(function (x) { return x !== s; }); };
var MAIN = new SS('專案推動_資料庫（假）'), AUTHC = new SS('權限中心（假）');
(function () { var d = AUTHC.insertSheet('dim_auth'); d.getRange(1, 1, 2, 7).setValues([['dashboard', 'name', 'pwd', 'ttl', 'enabled', 'updated', 'note'], ['project', '專案推動', 'x', '12', 'Y', '', '']]); })();

/* ---------- 其他假 Google 服務 ---------- */
var MAILS = [], LOGS = [], TRIGGERS = [];
function fmt(d, tz, f) {
  var p = function (n, w) { return ('0000' + n).slice(-(w || 2)); };
  var map = { yyyy: d.getFullYear(), yy: p(d.getFullYear() % 100), MM: p(d.getMonth() + 1), dd: p(d.getDate()), HH: p(d.getHours()), mm: p(d.getMinutes()), ss: p(d.getSeconds()) };
  return f.replace(/yyyy|yy|MM|dd|HH|mm|ss/g, function (k) { return map[k]; });
}
var G = {
  SpreadsheetApp: { getActiveSpreadsheet: function () { return MAIN; }, openById: function (id) { return AUTHC; } },
  Utilities: {
    formatDate: fmt, Charset: { UTF_8: 'utf8' }, DigestAlgorithm: { SHA_256: 'sha256' },
    computeDigest: function (alg, s) { return Array.from(crypto.createHash('sha256').update(String(s), 'utf8').digest()); },
    base64EncodeWebSafe: function (bytes) { return Buffer.from(bytes).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'); }
  },
  ContentService: { MimeType: { JAVASCRIPT: 'js', JSON: 'json' }, createTextOutput: function (s) { return { text: s, mime: 'json', setMimeType: function (m) { this.mime = m; return this; } }; } },
  LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, releaseLock: function () {} }; } },
  CacheService: (function () { var m = {}; return { getScriptCache: function () { return { get: function (k) { return m[k] || null; }, put: function (k, v) { m[k] = v; } }; } }; })(),
  UrlFetchApp: { fetch: function (u) { var q = url.parse(u, true).query, ok = q.action === 'ping' || q.pwd === PW; var body = JSON.stringify(ok ? { ok: true } : { ok: false, error: '密碼錯誤' }); return { getContentText: function () { return (q.callback || 'cb') + '(' + body + ');'; } }; } },
  MailApp: { sendEmail: function (o) { MAILS.push(Object.assign({ at: new Date().toISOString() }, o)); }, getRemainingDailyQuota: function () { return 100; } },
  ScriptApp: {
    getProjectTriggers: function () { return TRIGGERS.slice(); }, deleteTrigger: function (t) { TRIGGERS = TRIGGERS.filter(function (x) { return x !== t; }); },
    newTrigger: function (fn) { var t = { fn: fn, getHandlerFunction: function () { return fn; } }, b = { timeBased: function () { return b; }, everyDays: function () { return b; }, atHour: function (h) { t.hour = h; return b; }, nearMinute: function () { return b; }, inTimezone: function () { return b; }, create: function () { TRIGGERS.push(t); return t; } }; return b; }
  },
  PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return PROPS[k] || null; }, setProperty: function (k, v) { PROPS[k] = String(v); } }; } },
  Logger: { log: function (s) { LOGS.push(String(s)); console.log('[GAS]', s); } },
  Session: { getEffectiveUser: function () { return { getEmail: function () { return 'mock@example.com'; } }; } }
};
var ctx = vm.createContext(Object.assign({ console: console }, G));
['Code.gs', 'seed.gs', 'purpose.gs'].concat(fs.readdirSync(GAS).filter(function (f) { return /^zz_.*\.gs$/.test(f); })).forEach(function (f) {   /* zz_*.gs＝要在編輯器跑一次的暫時檔，本機先跑 /mock/run?fn= 驗證 */ vm.runInContext(fs.readFileSync(path.join(GAS, f), 'utf8'), ctx, { filename: f }); });
vm.runInContext('function __call(fn, arg){ return this[fn](arg); }', ctx);
/* 2026-10-29：MOCK_SPOON＝線上 GET ?action=spoon 的 JSON（或同格式），先放進假試算表，用真資料點測平匙換算／執行進度 */
if (process.env.MOCK_SPOON) {
  var fx = JSON.parse(fs.readFileSync(process.env.MOCK_SPOON, 'utf8')); fx = fx.data || fx;
  var TT = vm.runInContext('T', ctx);
  [['spoon_lines', fx.lines], ['spoon_set', fx.set], ['spoon_spec_log', fx.spec_log], ['spoon_evo', fx.evo], ['spoon_final', fx.final]].forEach(function (x) {
    if (!x[1] || !x[1].length) return;
    var cols = TT[x[0]].cols, sh = MAIN.insertSheet(x[0]);
    sh.rows = [cols.slice()].concat(x[1].map(function (o) { return cols.map(function (c) { return o[c] == null ? '' : String(o[c]); }); }));
  });
  console.log('MOCK_SPOON 已載入：' + process.env.MOCK_SPOON);
}

/* ---------- HTTP ---------- */
var MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json' };
function send(res, code, body, type) { res.writeHead(code, { 'Content-Type': type || 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'Access-Control-Allow-Origin': '*' }); res.end(body); }
http.createServer(function (req, res) {
  var u = url.parse(req.url, true);
  if (u.pathname === '/api') {
    if (req.method === 'POST') {
      var buf = ''; req.on('data', function (c) { buf += c; }); req.on('end', function () {
        var out = ctx.doPost({ postData: { contents: buf } }); send(res, 200, out.text);
      }); return;
    }
    var o = ctx.doGet({ parameter: u.query }); return send(res, 200, o.text, o.mime === 'js' ? 'text/javascript; charset=utf-8' : undefined);
  }
  if (u.pathname === '/mock/mails') return send(res, 200, JSON.stringify(MAILS, null, 1));
  if (u.pathname === '/mock/logs') return send(res, 200, JSON.stringify(LOGS, null, 1));
  if (u.pathname === '/mock/triggers') return send(res, 200, JSON.stringify(TRIGGERS.map(function (t) { return { fn: t.fn, hour: t.hour }; })));
  if (u.pathname === '/mock/sheets') { var d = {}; MAIN.sheets.forEach(function (s) { d[s.name] = s.rows; }); return send(res, 200, JSON.stringify(d, null, 1)); }
  if (u.pathname === '/mock/run') { try { var r = ctx.__call(u.query.fn); return send(res, 200, JSON.stringify({ ok: true, result: r === undefined ? null : r, logs: LOGS.slice(-5) })); } catch (e) { return send(res, 500, JSON.stringify({ ok: false, error: String(e && e.stack || e) })); } }
  var fp = path.normalize(path.join(ROOT, decodeURIComponent(u.pathname)));
  if (fp.indexOf(path.join(ROOT, 'static')) !== 0 || !fs.existsSync(fp) || fs.statSync(fp).isDirectory()) return send(res, 404, 'not found', 'text/plain');
  send(res, 200, fs.readFileSync(fp), MIME[path.extname(fp)] || 'application/octet-stream');
}).listen(PORT, '127.0.0.1', function () { console.log('project-hub mock on http://127.0.0.1:' + PORT + '/static/project-hub.html?api=/api  （密碼 ' + PW + '）'); });
