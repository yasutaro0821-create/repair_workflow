const CACHE_NAME = 'mtinn-repair-v1';
const ASSETS = [
  './',
  './index.html',
  './css/style.css',
  './js/app.js',
  './js/camera.js',
  './js/voice.js',
  './icons/icon-192.png',
  './icons/icon-512.png'
];

// インストール時にApp Shellをキャッシュ
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// 古いキャッシュを削除
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// ネットワーク優先、フォールバックでキャッシュ
self.addEventListener('fetch', (event) => {
  // POSTリクエスト（API送信）はキャッシュしない
  if (event.request.method === 'POST') {
    return;
  }

  event.respondWith(
    fetch(event.request)
      .then((response) => {
        // 成功したレスポンスをキャッシュに保存
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      })
      .catch(() => caches.match(event.request))
  );
});
