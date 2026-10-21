/**
 * 出貨中心訂單 API（diybc-dcorder-api）的本機模擬後端（2026-10-21，新竹托運單上線前測試用）：
 * 用 Node 直接跑 gas/dcorder/dcOrderApi.gs，試算表用假的（存在記憶體，關掉就沒了）；
 * 新竹物流照樣真的連（指令碼屬性沒設＝新竹公開測試帳號 test／test1，標籤有「不得寄件」浮水印，不會真的寄）。
 *
 * 啟動（.claude/launch.json 的 dcorder-mock 也是這條）：
 *   node tools/dcorder-mock/mock.js            → http://127.0.0.1:8767
 * 頁面：http://127.0.0.1:8767/static/dashboard-purchase.html
 *   開頁後在 Console 把 API 指到這裡：DCO_API=location.origin+"/exec";dcoRefresh();render();
 *   閘門用 auth-center 的緊急繞過（localStorage dash_auth_purchase）。
 * 一開始會放 4 張假訂單（3、7、12 號店；7 號店那張備貨中）。看假試算表：GET /mock/sheets
 * 跟 GAS 行為不同的地方：CacheService 不存、LockService 永遠拿得到鎖。
 */
'use strict';
var http = require('http'), fs = require('fs'), path = require('path'), vm = require('vm'), cp = require('child_process'), url = require('url');
var ROOT = path.resolve(__dirname, '..', '..'), PORT = Number(process.env.PORT || 8767);

/* ---------- 假試算表 ---------- */
function Sheet(name) { this.name = name; this.d = []; }
Sheet.prototype = {
  getName: function () { return this.name; }, setName: function (n) { this.name = n; return this; }, setFrozenRows: function () { },
  getLastRow: function () { return this.d.length; },
  getLastColumn: function () { return this.d.reduce(function (m, r) { return Math.max(m, r.length); }, 0); },
  getDataRange: function () { return this.getRange(1, 1, Math.max(1, this.d.length), Math.max(1, this.getLastColumn())); },
  appendRow: function (a) { this.d.push(a.slice()); }, deleteRow: function (r) { this.d.splice(r - 1, 1); }, deleteRows: function (r, n) { this.d.splice(r - 1, n); },
  getRange: function (r, c, nr, nc) {
    if (typeof r === 'string') return { setNumberFormat: function () { } };
    var sh = this; nr = nr || 1; nc = nc || 1;
    return {
      getValues: function () { var o = []; for (var i = 0; i < nr; i++) { var row = []; for (var j = 0; j < nc; j++) { var v = (sh.d[r - 1 + i] || [])[c - 1 + j]; row.push(v === undefined ? '' : v); } o.push(row); } return o; },
      getValue: function () { return this.getValues()[0][0]; },
      setValues: function (v) { for (var i = 0; i < nr; i++) { sh.d[r - 1 + i] = sh.d[r - 1 + i] || []; for (var j = 0; j < nc; j++) sh.d[r - 1 + i][c - 1 + j] = v[i][j]; } },
      setValue: function (v) { this.setValues([[v]]); }, setNumberFormat: function () { },
      createTextFinder: function (t) {
        var done = false;
        return { matchEntireCell: function () { return this; }, findNext: function () { if (done) return null; done = true; for (var i = 0; i < nr; i++) if (String((sh.d[r - 1 + i] || [])[c - 1]) === t) return { getRow: function () { return r + i; } }; return null; } };
      }
    };
  }
};
var SS = {
  sheets: [new Sheet('Sheet1')],
  getSheetByName: function (n) { return this.sheets.filter(function (s) { return s.name === n; })[0] || null; },
  insertSheet: function (n) { var s = new Sheet(n); this.sheets.push(s); return s; },
  getSheets: function () { return this.sheets; }, getId: function () { return 'MOCK'; }, getUrl: function () { return 'mock'; }, setSpreadsheetTimeZone: function () { }
};
var props = {};
function fmtDate(d, tz, f) {
  var t = new Date(d.getTime() + 8 * 3600e3), p = function (n) { return String(n).padStart(2, '0'); };
  return f.replace('yyyy', t.getUTCFullYear()).replace('yy', String(t.getUTCFullYear()).slice(2)).replace('MM', p(t.getUTCMonth() + 1)).replace('dd', p(t.getUTCDate()))
    .replace('HH', p(t.getUTCHours())).replace('mm', p(t.getUTCMinutes())).replace('ss', p(t.getUTCSeconds()));
}
var ctx = {
  console: console, JSON: JSON, Math: Math, Date: Date, String: String, Number: Number, Array: Array, Object: Object, RegExp: RegExp, parseInt: parseInt, isFinite: isFinite, Error: Error,
  SpreadsheetApp: { openById: function () { return SS; }, create: function () { return SS; } },
  PropertiesService: { getScriptProperties: function () { return { getProperty: function (k) { return props[k] || null; }, setProperty: function (k, v) { props[k] = v; } }; } },
  CacheService: { getScriptCache: function () { return { get: function () { return null; }, getAll: function () { return {}; }, putAll: function () { } }; } },
  LockService: { getScriptLock: function () { return { tryLock: function () { return true; }, releaseLock: function () { } }; } },
  Logger: { log: console.log },
  ContentService: { MimeType: { JSON: 1, JAVASCRIPT: 2 }, createTextOutput: function (s) { return { s: s, setMimeType: function () { return this; } }; } },
  Utilities: { formatDate: fmtDate, base64Encode: function (b) { return Buffer.from(b.map(function (x) { return x & 255; })).toString('base64'); } },
  UrlFetchApp: {   /* 真的連新竹（curl，同步） */
    fetch: function (u, o) {
      o = o || {}; var args = ['-s', '-m', '60', '-w', '\n%{http_code}'];
      if (o.method === 'post') { args.push('-H', 'Content-Type: ' + o.contentType); Object.keys(o.headers || {}).forEach(function (k) { args.push('-H', k + ': ' + o.headers[k]); }); args.push('--data-binary', '@-'); }
      args.push(u);
      var out = cp.execFileSync('curl', args, { input: o.payload || '', maxBuffer: 5e7 }).toString('utf8'), i = out.lastIndexOf('\n');
      var body = out.slice(0, i), code = +out.slice(i + 1);
      return { getResponseCode: function () { return code; }, getContentText: function () { return body; } };
    }
  }
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(ROOT, 'gas', 'dcorder', 'dcOrderApi.gs'), 'utf8'), ctx);
ctx.setup();
var TOKEN = vm.runInContext('DCO_TOKEN', ctx);
function post(b) { b.token = TOKEN; return JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(b) } }).s); }
/* 假訂單 */
post({ action: 'create', store: '3', cid: 'm1', items: [{ sku_id: 'F001', name: '高筋麵粉', unit: '包', qty: 4, price: 120 }, { sku_id: 'C001', name: '烘焙紙', unit: '捲', qty: 2, price: 85 }] });
var o2 = post({ action: 'create', store: '7', cid: 'm2', items: [{ sku_id: 'F002', name: '細砂糖', unit: '包', qty: 3, price: 60 }] });
post({ action: 'status', id: o2.id, status: '備貨中' });
post({ action: 'create', store: '12', cid: 'm3', items: [{ sku_id: 'F003', name: '無鹽奶油', unit: '塊', qty: 6, price: 210 }] });
post({ action: 'create', store: '9', cid: 'm4', items: [{ sku_id: 'F004', name: '可可粉', unit: '包', qty: 1, price: 330 }] });

var TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.json': 'application/json; charset=utf-8', '.css': 'text/css; charset=utf-8', '.png': 'image/png', '.jpg': 'image/jpeg', '.svg': 'image/svg+xml' };
http.createServer(function (req, res) {
  var u = url.parse(req.url, true);
  if (u.pathname === '/exec') {
    if (req.method === 'POST') {
      var body = '';
      req.on('data', function (c) { body += c; });
      req.on('end', function () {
        var out; try { out = ctx.doPost({ postData: { contents: body } }).s; } catch (e) { out = JSON.stringify({ ok: false, error: String(e && e.message || e) }); }
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' }); res.end(out);
      });
      return;
    }
    var o; try { o = ctx.doGet({ parameter: u.query }); } catch (e) { o = { s: JSON.stringify({ ok: false, error: String(e) }) }; }
    res.writeHead(200, { 'Content-Type': (u.query.callback ? 'text/javascript' : 'application/json') + '; charset=utf-8' }); res.end(o.s);
    return;
  }
  if (u.pathname === '/mock/sheets') {
    var dump = {}; SS.sheets.forEach(function (s) { dump[s.name] = s.d.map(function (r) { return r.map(function (v) { v = v instanceof Date ? v.toISOString() : v; return typeof v === 'string' && v.length > 80 ? v.slice(0, 40) + '…(' + v.length + ' 字)' : v; }); }); });
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' }); res.end(JSON.stringify(dump, null, 1)); return;
  }
  var f = path.join(ROOT, path.normalize(decodeURIComponent(u.pathname)).replace(/^(\.\.[\/\\])+/, ''));
  if (f.indexOf(ROOT) !== 0 || !fs.existsSync(f) || fs.statSync(f).isDirectory()) { res.writeHead(404); res.end('not found'); return; }
  res.writeHead(200, { 'Content-Type': TYPES[path.extname(f)] || 'application/octet-stream' }); fs.createReadStream(f).pipe(res);
}).listen(PORT, '127.0.0.1', function () { console.log('dcorder mock → http://127.0.0.1:' + PORT + '/static/dashboard-purchase.html'); });
