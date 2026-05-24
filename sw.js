const CACHE = 'subqueue-v4';
const ASSETS = ['./index.html', './manifest.json'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // Never intercept cross-origin requests (e.g. Anthropic API calls)
  if (!e.request.url.startsWith(self.location.origin)) return;

  // Network-first for HTML so updates are always picked up immediately
  if (e.request.mode === 'navigate' || e.request.url.endsWith('index.html')) {
    e.respondWith(
      fetch(e.request).then(r => {
        const clone = r.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return r;
      }).catch(() => caches.match('./index.html'))
    );
    return;
  }

  // Cache-first for same-origin assets (icons, manifest)
  e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
});
