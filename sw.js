const CACHE_NAME = 'controle-entrega-v2';
const PRECACHE = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png'];

self.addEventListener('install', function(e) {
  e.waitUntil( caches.open(CACHE_NAME).then(function(c) { return c.addAll(PRECACHE); }).then(function() { return self.skipWaiting(); }) );
});

self.addEventListener('activate', function(e) {
  e.waitUntil( caches.keys().then(function(keys) {
    return Promise.all(keys.filter(function(k) { return k !== CACHE_NAME; }).map(function(k) { return caches.delete(k); }));
  }).then(function() { return self.clients.claim(); }) );
});

self.addEventListener('fetch', function(e) {
  var req = e.request;
  if (req.method !== 'GET') { return; }
  var url = new URL(req.url);
  if (url.hostname === 'script.google.com' || url.hostname.endsWith('.googleusercontent.com')) { return; }
  e.respondWith(
    fetch(req).then(function(res) {
      if (res && res.ok && url.origin === self.location.origin) {
        var clone = res.clone();
        caches.open(CACHE_NAME).then(function(c) { c.put(req, clone); });
      }
      return res;
    }).catch(function() {
      return caches.match(req).then(function(cached) { return cached || caches.match('./index.html'); });
    })
  );
});

