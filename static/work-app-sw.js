/* 2026-11-26 手機版工作清單 App 的 service worker（範圍只有 /static/work-app*，不影響其他儀表板）
   頁面與圖示：先連網、失敗才用快取（沒網路時還打得開，但資料要連網才更新）。API（script.google.com）不經過這裡。 */
var CACHE = 'work-app-v2';
var FILES = ['work-app.html', 'work-app.manifest.json', 'work-app/icon-192.png', 'work-app/icon-512.png', 'work-app/apple-touch-icon.png'];
self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(FILES); }).catch(function () {}).then(function () { return self.skipWaiting(); }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k.indexOf('work-app-') === 0 && k !== CACHE; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(fetch(req).then(function (res) {
    if (res && res.ok) { var copy = res.clone(); caches.open(CACHE).then(function (c) { c.put(req.url.split('#')[0].split('?')[0], copy); }); }
    return res;
  }).catch(function () { return caches.match(req.url.split('#')[0].split('?')[0]).then(function (r) { return r || caches.match('work-app.html'); }); }));
});
