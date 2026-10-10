const CACHE = 'vegcount-v18';
const ASSETS = ['./', './index.html', './manifest.json', './xlsx-report.js',
  './icon-192.png', './icon-512.png', './apple-touch-icon.png'];
// CDN libraries (Firebase, ExcelJS) are cached on first use so the app also opens offline
const CDN = ['www.gstatic.com', 'cdnjs.cloudflare.com'];

self.addEventListener('install', e =>
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS))));

self.addEventListener('activate', e =>
  e.waitUntil(caches.keys().then(keys =>
    Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))));

self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;
  const u = new URL(e.request.url);
  if (u.origin === location.origin) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
  } else if (CDN.includes(u.hostname)) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    })));
  }                                     // everything else (Firestore, Apps Script) goes straight to the network
});
