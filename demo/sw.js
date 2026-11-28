/* Mượn thiết bị dạy học – service worker: mở nhanh, luôn lấy bản mới nhất khi có mạng */
var C = 'mtb-v1';
self.addEventListener('install', function (e) {
  self.skipWaiting();
  e.waitUntil(caches.open(C).then(function (c) { return c.addAll(['./', 'icon-192.png', 'manifest.webmanifest']); }).catch(function () { }));
});
self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (ks) { return Promise.all(ks.filter(function (k) { return k !== C; }).map(function (k) { return caches.delete(k); })); }).then(function () { return self.clients.claim(); }));
});
self.addEventListener('fetch', function (e) {
  var r = e.request;
  if (r.method !== 'GET') return;
  var u = new URL(r.url);
  var same = u.origin === location.origin, cdn = /cdn\.jsdelivr\.net$/.test(u.hostname);
  if (!same && !cdn) return; // không can thiệp máy chủ Google
  if (same && (r.mode === 'navigate' || /\/$|\.html$/.test(u.pathname))) {
    // trang chính: mạng trước, mất mạng thì dùng bản đã lưu
    e.respondWith(fetch(r).then(function (res) {
      var cl = res.clone(); caches.open(C).then(function (c) { c.put(r, cl); }); return res;
    }).catch(function () { return caches.match(r).then(function (m) { return m || caches.match('./'); }); }));
    return;
  }
  e.respondWith(caches.match(r).then(function (m) {
    return m || fetch(r).then(function (res) { if (res.ok) { var cl = res.clone(); caches.open(C).then(function (c) { c.put(r, cl); }); } return res; });
  }));
});
