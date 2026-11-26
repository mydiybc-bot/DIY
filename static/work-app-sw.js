/* 2026-11-26 自己做 App 的 service worker（範圍只有 /static/work-app*，不影響其他儀表板）
   頁面與圖示：先連網、失敗才用快取（沒網路時還打得開，但資料要連網才更新）。API（script.google.com）不經過這裡。
   2026-11-26 晚 推播：收到推播 → 跳通知＋桌面圖示數字（payload.n）；點通知 → 打開那件工作。 */
var CACHE = 'work-app-v7';
var FILES = ['work-app.html', 'work-app.manifest.json', 'work-app/icon-192.png', 'work-app/icon-512.png', 'work-app/apple-touch-icon.png', 'work-app/logo.png'];
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
self.addEventListener('push', function (e) {
  var d = {};
  try { d = e.data ? e.data.json() : {}; } catch (_) { d = { title: '自己做', body: e.data ? e.data.text() : '' }; }
  var icon = new URL('work-app/icon-192.png', self.location).href;
  var jobs = [self.registration.showNotification(d.title || '自己做', {
    body: d.body || '', icon: icon, badge: icon, tag: d.tag || 'ziji', renotify: true,
    data: { url: d.url || new URL('work-app.html', self.location).href }
  })];
  try {
    if (typeof d.n === 'number' && self.navigator && self.navigator.setAppBadge) jobs.push(d.n ? self.navigator.setAppBadge(d.n) : self.navigator.clearAppBadge());
  } catch (_) {}
  e.waitUntil(Promise.all(jobs).catch(function () {}));
});
self.addEventListener('notificationclick', function (e) {
  e.notification.close();
  var u = (e.notification.data && e.notification.data.url) || new URL('work-app.html', self.location).href;
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) {
      var c = cs[i];
      if (c.url.indexOf('work-app') >= 0 && 'focus' in c) {
        return c.focus().then(function (cc) { return (cc || c).navigate ? (cc || c).navigate(u).catch(function () { return self.clients.openWindow(u); }) : self.clients.openWindow(u); });
      }
    }
    return self.clients.openWindow(u);
  }));
});
