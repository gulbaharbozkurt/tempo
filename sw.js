const C = 'tempo-v9';
const FILES = ['./', 'index.html', 'app.js', 'xlsx.full.min.js', 'manifest.webmanifest', 'icon.svg', 'icon-192.png', 'icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(C).then(c => c.addAll(FILES))); self.skipWaiting(); });
self.addEventListener('activate', e => e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== C).map(k => caches.delete(k))))));
// Önce ağ (güncelleme gelsin), yoksa önbellek (internetsiz çalışsın)
self.addEventListener('fetch', e => e.respondWith(
  fetch(e.request, { cache: 'no-cache' }).then(r => { const cp = r.clone(); caches.open(C).then(c => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request))));
