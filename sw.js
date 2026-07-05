// MET2026 서비스 워커
// 항상 최신 버전을 우선 사용(네트워크 우선) - 신청서 로직이 자주 바뀌므로 캐시로 인한 오래된 화면 노출 방지
// 오프라인일 때만 마지막으로 접속 성공했던 화면을 보여줌

const CACHE_NAME = 'met2026-v2'; // v1→v2: 기존 캐시 전체 폐기(스테일 admin.html 등 정리)
const CORE_ASSETS = [
  './index.html',
  './buscheck.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  // GET 요청만 처리 (POST/GAS 통신은 그대로 네트워크로)
  if (event.request.method !== 'GET') return;

  // 관리자 페이지(admin.html 등)는 서비스워커가 아예 건드리지 않음
  // → 관리자 화면은 항상 서버의 최신 코드를 그대로 받아야 하고, 캐시로 인해
  //   수정한 내용이 반영 안 되는 문제를 막기 위함
  if (event.request.url.indexOf('admin') !== -1) return;

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
