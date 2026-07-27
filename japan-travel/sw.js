/* 일본여행 번역기 — 서비스워커 (오프라인 지원)
   서브도메인 루트(japan.joocnj.com/)와 하위 경로(/japan-travel/) 어디에
   마운트되어도 동작하도록 모든 경로를 SW 위치 기준 상대경로로 계산한다. */
const CACHE = 'japan-travel-v3';
const BASE = new URL('./', self.location).pathname; // '/' 또는 '/japan-travel/'
const ASSETS = [
  '', 'index.html', 'manifest.webmanifest',
  'icon-192.png', 'icon-512.png', 'icon-maskable-512.png',
].map(p => BASE + p);

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
  // 번역 API 등 외부 요청과 스코프 밖 요청은 그대로 네트워크로
  if (e.request.method !== 'GET' || url.origin !== self.location.origin) return;
  if (!url.pathname.startsWith(BASE)) return;

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
        if (e.request.mode === 'navigate') return caches.match(BASE + 'index.html');
        return Response.error();
      });
    })
  );
});
