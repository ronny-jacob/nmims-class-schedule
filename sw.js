var CACHE = 'nmims-cc-v2';
var STATIC_ASSETS = [
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable.png',
  './apple-touch-icon.png',
  './sw.js',
];

self.addEventListener('install', function(e) {
  self.skipWaiting();
});

self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(keys.map(function(k) {
        if (k !== CACHE) return caches.delete(k);
      }));
    }).then(function() {
      return self.clients.claim();
    })
  );
});

self.addEventListener('fetch', function(e) {
  if (e.request.method !== 'GET') return;
  var url = new URL(e.request.url);
  if (url.origin !== location.origin) return;

  // Network-first for the app shell and data — the embedded DATA in index.html
  // is rebuilt every deploy, so a stale copy would hide placement/timetable
  // updates from returning visitors for up to 24h.
  if (
    url.pathname === '/' ||
    url.pathname.endsWith('/index.html') ||
    url.pathname.endsWith('/data.json')
  ) {
    e.respondWith(
      fetch(e.request).then(function(res) {
        var copy = res.clone();
        caches.open(CACHE).then(function(c) { c.put(e.request, copy); });
        return res;
      }).catch(function() {
        return caches.match(e.request);
      })
    );
    return;
  }

  // Cache-first for static assets (icons, manifest, sw.js itself).
  e.respondWith(
    caches.match(e.request).then(function(cached) {
      if (cached) return cached;
      return fetch(e.request).then(function(res) {
        if (res && res.status === 200 && res.type === 'basic') {
          var copy = res.clone();
          caches.open(CACHE).then(function(c) { c.put(e.request, copy); });
        }
        return res;
      }).catch(function() {
        return caches.match('./manifest.json');
      });
    })
  );
});
