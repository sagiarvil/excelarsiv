// Service Worker — MANDATE v8.0 Mobile-First PWA
const CACHE_NAME = 'excelarsiv-pwa-v8';
const STATIC_ASSETS = [
  '/',
  '/manifest.webmanifest',
  '/images/brand/excelarsiv-header-logo.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;

  // HTML: Network-First with Cache fallback
  if (req.headers.get('accept')?.includes('text/html')) {
    event.respondWith(
      fetch(req).catch(() => caches.match(req).then((res) => res || caches.match('/')))
    );
    return;
  }

  // Assets: Cache-First
  event.respondWith(
    caches.match(req).then((cached) => cached || fetch(req).then((networkRes) => {
      if (networkRes && networkRes.status === 200 && (req.url.startsWith('https:') || req.url.startsWith('http:'))) {
        const cloned = networkRes.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(req, cloned));
      }
      return networkRes;
    }))
  );
});
