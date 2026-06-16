/* sw.js — 서비스워커: 정적 자산 precache + cache-first (오프라인 동작) */
const CACHE = 'sample-scanner-v1';
const ASSETS = [
  './',
  './index.html',
  './style.css',
  './db.js',
  './app.js',
  './manifest.webmanifest',
  './lib/xlsx.full.min.js',
  './lib/html5-qrcode.min.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-512-maskable.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (e) => {
  const { request } = e;
  if (request.method !== 'GET') return;
  e.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request)
        .then((res) => {
          // 동일 출처 자산은 런타임 캐시에 보관
          if (res && res.status === 200 && new URL(request.url).origin === self.location.origin) {
            const copy = res.clone();
            caches.open(CACHE).then((c) => c.put(request, copy));
          }
          return res;
        })
        .catch(() => cached);
    })
  );
});
