/* 일본여행 번역기 — 서비스워커 (오프라인 지원) */
const CACHE = 'japan-travel-v1';
const ASSETS = [
  '/japan-travel/',
  '/japan-travel/index.html',
  '/japan-travel/manifest.webmanifest',
  '/japan-travel/icon-192.png',
  '/japan-travel/icon-512.png',
  '/japan-travel/icon-maskable-512.png',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);
  // 번역 API 등 외부 요청은 그대로 네트워크로 (캐시하지 않음)
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (!url.pathname.startsWith('/japan-travel')) return;

  // 앱 셸: 캐시 우선, 없으면 네트워크 후 캐시 저장, 최후엔 index.html 폴백
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy));
        }
        return res;
      }).catch(() => {
        if (e.request.mode === 'navigate') return caches.match('/japan-travel/index.html');
        return Response.error();
      });
    })
  );
});
