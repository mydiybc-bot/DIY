/**
 * 顧客問卷 API 的本機模擬後端（2026-10-28）：用 Node 直接跑 gas/survey/Code.gs，
 * 試算表、快取、鎖全部用假的，讓 static/survey.html 與 static/dashboard-survey.html 在本機完整點測，不碰線上資料。
 *
 * 啟動（.claude/launch.json 的 survey-mock 也是這條）：
 *   node tools/survey-mock/mock.js            → http://127.0.0.1:8768
 * 頁面：http://127.0.0.1:8768/static/survey.html?api=/api
 *       http://127.0.0.1:8768/static/dashboard-survey.html?api=/api（閘門用 auth-center 的緊急繞過 localStorage dash_auth_survey）
 * 看假試算表：GET /mock/sheets；直接跑 GAS 函式：GET /mock/run?fn=authorize；灌假資料：GET /mock/seed?n=300
 * 跟 GAS 行為不同的地方：沒有列數上限、LockService 永遠拿得到鎖、CacheService 100KB 上限沒有模擬（Code.gs 自己有判斷）。
 */
'use strict';
var http = require('http'), fs = require('fs'), path = require('path'), vm = require('vm'), crypto = require('crypto'), url = require('url');
var ROOT = path.resolve(__dirname, '..', '..'), GAS = path.join(ROOT, 'gas', 'survey');
var PORT = Number(process.env.PORT || 8768);

/* ---------- 假試算表（值保留原型別：Date 進去就是 Date，跟 GAS getValues 一樣） ---------- */
function Sheet(name, ss) { this.name = name; this.ss = ss; this.rows = []; this.frozen = 0; }
Sheet.prototype.getName = function () { return this.name; };
Sheet.prototype.setName = function (n) { this.name = n; return this; };
Sheet.prototype.getLastRow = function () { var n = this.rows.length; while (n > 0 && this.rows[n - 1].every(function (v) { return v === '' || v == null; })) n--; return n; };
Sheet.prototype.getLastColumn = function () { var m = 0; this.rows.forEach(function (r) { var n = r.length; while (n > 0 && (r[n - 1] === '' || r[n - 1] == null)) n--; if (n > m) m = n; }); return m; };
Sheet.prototype.setFrozenRows = function (n) { this.frozen = n; return this; };
Sheet.prototype.copyTo = function (ss) { var s = new Sheet(this.name + ' 的副本', ss); s.rows = this.rows.map(function (r) { return r.slice(); }); ss.sheets.push(s); return s; };
Sheet.prototype.appendRow = function (vals) { var row = vals.map(function (v) { return v == null ? '' : v; }); this.rows.push(row); return this; };
Sheet.prototype.cell_ = function (r, c) { while (this.rows.length < r) this.rows.push([]); var row = this.rows[r - 1]; while (row.length < c) row.push(''); return row; };
Sheet.prototype.getDataRange = function () { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); };
Sheet.prototype.getRange = function (r, c, nr, nc) {
  var sh = this; nr = nr || 1; nc = nc || 1;
  return {
    setNumberFormat: function () { return this; }, setFontWeight: function () { return this; },
    setValues: function (vals) { vals.forEach(function (vr, i) { var row = sh.cell_(r + i, c + nc - 1); vr.forEach(function (v, j) { row[c - 1 + j] = v == null ? '' : v; }); }); return this; },
    setValue: function (v) { sh.cell_(r, c)[c - 1] = v == null ? '' : v; return this; },
    getValues: function () { var out = []; for (var i = 0; i < nr; i++) { var row = sh.rows[r - 1 + i] || [], o = []; for (var j = 0; j < nc; j++) { var x = row[c - 1 + j]; o.push(x == null ? '' : x); } out.push(o); } return out; },
    getDisplayValues: function () { return this.getValues().map(function (r) { return r.map(String); }); }
  };
};
function SS(name, id) { this.name = name; this.id = id; this.sheets = [new Sheet('工作表1', this)]; }
SS.prototype.getName = function () { return this.name; };
SS.prototype.getId = function () { return this.id; };
SS.prototype.getUrl = function () { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; };
SS.prototype.setSpreadsheetTimeZone = function () { return this; };
SS.prototype.getSheets = function () { return this.sheets; };
SS.prototype.getSheetByName = function (n) { return this.sheets.filter(function (s) { return s.name === n; })[0] || null; };
SS.prototype.insertSheet = function (n) { var s = new Sheet(n, this); this.sheets.push(s); return s; };
SS.prototype.deleteSheet = function (s) { this.sheets = this.sheets.filter(function (x) { return x !== s; }); };
var FILES = {}, AUTHC = new SS('權限中心（假）', '1_aXfjX-JKo8rr9EJRjiRxhSR0zdMSJfL3I4f0FjIxj4');   /* id 用真的權限中心 id，Code.gs 的 authSetup_ 才找得到 */
FILES[AUTHC.id] = AUTHC;
(function () { var d = AUTHC.insertSheet('dim_auth'); d.getRange(1, 1, 2, 7).setValues([['dashboard', 'name', 'pwd', 'ttl', 'enabled', 'updated', 'note'], ['hub_hq', '經營儀表板（總部）', 'hqpw', '12', 'Y', '', '']]); })();

/* ---------- 其他假 Google 服務 ---------- */
var LOGS = [], PROPS = {}, CACHE = {};
function fmt(d, tz, f) {
  if (!(d instanceof Date)) d = new Date(d);
  var p = function (n, w) { return ('0000' + n).slice(-(w || 2)); };
  var map = { yyyy: d.getFullYear(), yy: p(d.getFullYear() % 100), MM: p(d.getMonth() + 1), dd: p(d.getDate()), HH: p(d.getHours()), mm: p(d.getMinutes()), ss: p(d.getSeconds()) };
  return f.replace(/yyyy|yy|MM|dd|HH|mm|ss/g, function (k) { return map[k]; });
}
var G = {
  SpreadsheetApp: {
    create: function (name) { var id = 'SS' + crypto.randomBytes(6).toString('hex'); var ss = new SS(name, id); FILES[id] = ss; return ss; },
    openById: function (id) { if (!FILES[id]) throw new Error('找不到試算表 ' + id); return FILES[id]; },
    flush: function () {}
  },
  PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return PROPS[k] == null ? null : PROPS[k]; }, setProperty: function (k, v) { PROPS[k] = String(v); }, getKeys: function () { return Object.keys(PROPS); }, deleteProperty: function (k) { delete PROPS[k]; } }; } },
  CacheService: { getScriptCache: function () { return { get: function (k) { var e = CACHE[k]; if (!e) return null; if (e.exp < Date.now()) { delete CACHE[k]; return null; } return e.v; }, put: function (k, v, sec) { CACHE[k] = { v: String(v), exp: Date.now() + (sec || 600) * 1000 }; }, remove: function (k) { delete CACHE[k]; } }; } },
  Utilities: { formatDate: fmt, getUuid: function () { return crypto.randomUUID(); } },
  ContentService: { MimeType: { JAVASCRIPT: 'js', JSON: 'json' }, createTextOutput: function (s) { return { text: s, mime: 'json', setMimeType: function (m) { this.mime = m; return this; } }; } },
  LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, waitLock: function () { return true; }, releaseLock: function () {} }; } },
  Logger: { log: function (s) { LOGS.push(String(s)); console.log('[GAS]', s); } }
};
var ctx = vm.createContext(Object.assign({ console: console }, G));
['Code.gs'].forEach(function (f) { vm.runInContext(fs.readFileSync(path.join(GAS, f), 'utf8'), ctx, { filename: f }); });
vm.runInContext('function __call(fn, arg){ return this[fn](arg); }', ctx);

/* 灌假資料：n 筆，隨機店、隨機答案，時間散在最近 120 天（src=mock，儀表板算進去） */
function seed(n) {
  var form = JSON.parse(ctx.doGet({ parameter: { action: 'form' } }).text), qs = form.questions, made = 0;
  for (var i = 0; i < n; i++) {
    var a = {}; qs.forEach(function (q) {
      if (q.type === 'text') { if (Math.random() < 0.3) a[q.qid] = ['老師很親切', '影片有點快', '環境很舒服，下次再來', '希望多一點巧克力口味', '桌子有點擠'][i % 5]; return; }
      if (q.type === 'multi') { a[q.qid] = [q.options[i % q.options.length]]; if (Math.random() < 0.2) a[q.qid].push(q.options[(i + 2) % q.options.length]); return; }
      var k = q.options.length; var pick = q.options[Math.floor(Math.random() * k)];
      if (/^(是|否)$/.test(q.options[0] || '') || /^(會|不會)$/.test(q.options[0] || '')) pick = Math.random() < 0.88 ? q.options[0] : q.options[1];
      a[q.qid] = pick;
    });
    var r = JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ action: 'submit', id: 'mock-' + crypto.randomBytes(6).toString('hex'), a: a, src: 'mock', sec: 60 + Math.floor(Math.random() * 200), ua: 'mock' }) } }).text);
    if (r.ok) { made++; var sh = FILES[PROPS.DB_SS].getSheetByName('fact_response'); var row = sh.rows[sh.rows.length - 1]; row[0] = new Date(Date.now() - Math.floor(Math.random() * 120) * 864e5 - Math.floor(Math.random() * 864e5)); }
  }
  CACHE = {};
  return made;
}

/* ---------- HTTP ---------- */
var MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.png': 'image/png', '.svg': 'image/svg+xml', '.json': 'application/json' };
function send(res, code, body, type) { res.writeHead(code, { 'Content-Type': type || 'application/json; charset=utf-8', 'Cache-Control': 'no-store', 'Access-Control-Allow-Origin': '*' }); res.end(body); }
http.createServer(function (req, res) {
  var u = url.parse(req.url, true);
  if (u.pathname === '/api' || /^\/api\/gas\//.test(u.pathname)) {
    if (req.method === 'POST') {
      var buf = ''; req.on('data', function (c) { buf += c; }); req.on('end', function () {
        var out = ctx.doPost({ postData: { contents: buf } }); send(res, 200, out.text);
      }); return;
    }
    var o = ctx.doGet({ parameter: u.query }); return send(res, 200, o.text, o.mime === 'js' ? 'text/javascript; charset=utf-8' : undefined);
  }
  if (u.pathname === '/mock/logs') return send(res, 200, JSON.stringify(LOGS, null, 1));
  if (u.pathname === '/mock/props') return send(res, 200, JSON.stringify(PROPS, null, 1));
  if (u.pathname === '/mock/sheets') { var d = {}; Object.keys(FILES).forEach(function (id) { FILES[id].sheets.forEach(function (s) { d[FILES[id].name + '/' + s.name] = s.rows; }); }); return send(res, 200, JSON.stringify(d, null, 1)); }
  if (u.pathname === '/mock/seed') { try { return send(res, 200, JSON.stringify({ ok: true, made: seed(Number(u.query.n) || 200) })); } catch (e) { return send(res, 500, JSON.stringify({ ok: false, error: String(e && e.stack || e) })); } }
  if (u.pathname === '/mock/run') { try { var r = ctx.__call(u.query.fn); return send(res, 200, JSON.stringify({ ok: true, result: r === undefined ? null : r, logs: LOGS.slice(-5) })); } catch (e) { return send(res, 500, JSON.stringify({ ok: false, error: String(e && e.stack || e) })); } }
  var fp = path.normalize(path.join(ROOT, decodeURIComponent(u.pathname)));
  if (fp.indexOf(path.join(ROOT, 'static')) !== 0 || !fs.existsSync(fp) || fs.statSync(fp).isDirectory()) return send(res, 404, 'not found', 'text/plain');
  send(res, 200, fs.readFileSync(fp), MIME[path.extname(fp)] || 'application/octet-stream');
}).listen(PORT, '127.0.0.1', function () { console.log('survey mock on http://127.0.0.1:' + PORT + '/static/survey.html?api=/api'); });
