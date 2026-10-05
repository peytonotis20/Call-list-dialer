// Offline support: the app page loads fresh when online (so updates show up right away)
// and falls back to the cached copy when offline. Other files are served from cache.
const CACHE = 'call-list-dialer-v2';
const FILES = ['./', 'index.html', 'xlsx.full.min.js', 'manifest.webmanifest', 'icon-180.png', 'icon-192.png', 'icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(FILES)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET' || new URL(e.request.url).origin !== location.origin) return;
  const fromNet = () => fetch(e.request).then(res => {
    if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
    return res;
  });
  if (e.request.mode === 'navigate') {
    e.respondWith(fromNet().catch(() => caches.match(e.request).then(hit => hit || caches.match('index.html'))));
  } else {
    e.respondWith(caches.match(e.request).then(hit => hit || fromNet()));
  }
});
