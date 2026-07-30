/* 일본여행 번역기 — 서비스워커 (오프라인 지원)
   서브도메인 루트(japan.joocnj.com/)와 하위 경로(/japan-travel/) 어디에
   마운트되어도 동작하도록 모든 경로를 SW 위치 기준 상대경로로 계산한다.
   전략: HTML(내비게이션)은 네트워크 우선 → 항상 최신 버전 표시, 오프라인일 때만 캐시.
         아이콘·매니페스트 등 정적 자산은 캐시 우선. */
const CACHE = 'japan-travel-v7';
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

  const isNav = e.request.mode === 'navigate'
    || url.pathname === BASE || url.pathname === BASE + 'index.html';

  if (isNav) {
    // HTML: 네트워크 우선 — 배포 즉시 최신 버전 반영, 오프라인이면 캐시 폴백
    e.respondWith(
      fetch(e.request).then(res => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(BASE + 'index.html', copy));
        }
        return res;
      }).catch(() =>
        caches.match(BASE + 'index.html').then(c => c || Response.error())
      )
    );
    return;
  }

  // 정적 자산: 캐시 우선, 없으면 네트워크 후 저장
  e.respondWith(
    caches.match(e.request, { ignoreSearch: true }).then(cached => {
      if (cached) return cached;
      return fetch(e.request).then(res => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(e.request, copy));
        }
        return res;
      }).catch(() => Response.error());
    })
  );
});
